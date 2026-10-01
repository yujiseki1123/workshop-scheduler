/**
 * スケジュール管理ブック(Excel)を読み取る層(app/services/excel_source.py の JavaScript 版)。
 *
 * Excel が唯一の正(Single Source of Truth)。このアプリは読むだけで書き換えない。
 * ブックの形式は tools/make_schedule_workbook.py で作るもの。
 *
 * 読むシート:
 *   WS          : 行の一覧(WS名・種別・並び順)
 *   タスク一覧  : 1行 = 1タスク。WS名・タスク名などは値、開始日・終了日・担当者・警告は数式の計算結果
 *                 (WSのタスクは「タスク生成」シートから値として貼り付けて作る)
 *   個別タスク  : WSに属さない作業(社内MTGなど)。WS名が空欄なら「その他」
 *   メンバー    : 担当者の並び順
 *   タスクマスタ: 参照用
 *
 * 入力は SheetJS(XLSX.read)で読んだブック。数式は Excel が保存した「計算結果」を読む
 * (Python 版の openpyxl data_only と同じ)。日付はシリアル値のまま受け取って自分で変換する
 * (タイムゾーンによるずれを避けるため。XLSX.read は cellDates を付けずに呼ぶこと)。
 * 日付は "YYYY-MM-DD" の文字列で扱う。
 *
 * Python 版と同じ入力から同じ結果(/api/schedule の JSON)を返すことを目標にしている。
 * 片方を直したらもう片方も直すこと。
 *
 * ブラウザでは window.ExcelSource、Node では module.exports で使える。
 */
(function (root, factory) {
  if (typeof module === "object" && module.exports) module.exports = factory();
  else root.ExcelSource = factory();
})(typeof self !== "undefined" ? self : this, function () {
  "use strict";

  const SHEET_WS = "WS";
  const SHEET_TASKS = "タスク一覧";
  const SHEET_OTHER = "個別タスク";
  const SHEET_MEMBERS = "メンバー";
  const SHEET_MASTERS = "タスクマスタ";
  const SHEET_OLD_INPUT = "日程入力"; // 旧形式(タスク一覧を数式で自動生成していた頃)のブックにだけある
  const UNDECIDED = "未定";
  const CATEGORY_WORKSHOP = "ワークショップ";
  const CATEGORY_OTHER = "その他";
  const ASSIGNEE_SEPARATOR = "、";
  const STATUSES = ["未着手", "着手中", "完了"];
  const DEFAULT_STATUS = "未着手";
  const MAX_ASSIGNEES = 10;

  /** ブックを読めない・形式が違うときのエラー(画面にそのまま表示できる文言)。 */
  class ExcelSourceError extends Error {}
  class ExcelNotCalculatedError extends ExcelSourceError {}

  // -------------------------------------------------------------------------
  // セル値の変換(Python 版の _text / _int / _date / _status と同じ規則)
  // -------------------------------------------------------------------------

  function text(v) {
    if (v === null || v === undefined) return null;
    if (typeof v === "boolean") return v ? "True" : "False";
    const s = String(v).trim();
    return s || null;
  }

  function int(v) {
    if (v === null || v === undefined || typeof v === "boolean") return null;
    if (typeof v === "number") return Number.isFinite(v) ? Math.trunc(v) : null;
    const s = String(v).normalize("NFKC").trim();
    return /^[+-]?\d+$/.test(s) ? parseInt(s, 10) : null;
  }

  const pad = (n, w = 2) => String(n).padStart(w, "0");
  const DAY_MS = 86400000;
  const EXCEL_EPOCH = Date.UTC(1899, 11, 30);

  function isoFromUTC(ms) {
    const d = new Date(ms);
    return `${pad(d.getUTCFullYear(), 4)}-${pad(d.getUTCMonth() + 1)}-${pad(d.getUTCDate())}`;
  }

  /** 日付セル → "YYYY-MM-DD"。「未定」や空欄、読めない値は null。 */
  function toDate(v) {
    if (v === null || v === undefined || typeof v === "boolean") return null;
    if (v instanceof Date) {
      return `${pad(v.getFullYear(), 4)}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    }
    if (typeof v === "number") {
      // Excel のシリアル値(時刻部分は切り捨て)
      return Number.isFinite(v) ? isoFromUTC(EXCEL_EPOCH + Math.floor(v) * DAY_MS) : null;
    }
    const s = String(v).trim();
    if (!s || s === UNDECIDED) return null;
    const m = s.match(/^(\d{4})([/.-])(\d{1,2})\2(\d{1,2})$/);
    if (!m) return null;
    const [y, mo, d] = [Number(m[1]), Number(m[3]), Number(m[4])];
    const ms = Date.UTC(y, mo - 1, d);
    const back = new Date(ms);
    if (back.getUTCFullYear() !== y || back.getUTCMonth() !== mo - 1 || back.getUTCDate() !== d) return null;
    return isoFromUTC(ms);
  }

  function addDays(iso, days) {
    const [y, m, d] = iso.split("-").map(Number);
    return isoFromUTC(Date.UTC(y, m - 1, d) + days * DAY_MS);
  }

  function diffDays(fromIso, toIso) {
    const ms = (iso) => {
      const [y, m, d] = iso.split("-").map(Number);
      return Date.UTC(y, m - 1, d);
    };
    return Math.round((ms(toIso) - ms(fromIso)) / DAY_MS);
  }

  function status(v) {
    const s = text(v);
    return STATUSES.includes(s) ? s : DEFAULT_STATUS;
  }

  // -------------------------------------------------------------------------
  // シートの読み取り
  // -------------------------------------------------------------------------

  function colName(c) {
    let s = "";
    for (let n = c + 1; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return s;
  }

  function lastCell(ref) {
    // "A1:Y601" → {r: 600, c: 24}(0始まり)
    const end = String(ref || "A1").split(":").pop();
    const m = end.match(/^([A-Z]+)(\d+)$/);
    if (!m) return { r: 0, c: 0 };
    let c = 0;
    for (const ch of m[1]) c = c * 26 + (ch.charCodeAt(0) - 64);
    return { r: Number(m[2]) - 1, c: c - 1 };
  }

  function cellValue(ws, r, c) {
    const cell = ws[colName(c) + (r + 1)];
    if (!cell || cell.v === undefined || cell.t === "z") return null;
    if (cell.t === "e") return cell.w || "#ERROR"; // エラー値(#N/A など)は文字列として扱う
    return cell.v;
  }

  /** シートの行を [行番号, {見出し: 値}] で返す。見出しは改行の前(「WS名\n(自動)」→「WS名」)で照合する。 */
  function rows(ws) {
    if (!ws || !ws["!ref"]) return [];
    const end = lastCell(ws["!ref"]);
    const names = [];
    for (let c = 0; c <= end.c; c++) names.push((text(cellValue(ws, 0, c)) || "").split("\n")[0]);
    const out = [];
    for (let r = 1; r <= end.r; r++) {
      const obj = {};
      names.forEach((n, c) => {
        if (n) obj[n] = cellValue(ws, r, c);
      });
      out.push([r + 1, obj]);
    }
    return out;
  }

  const get = (r, key) => (key in r ? r[key] : null);

  function require(wb, name) {
    if (!wb.SheetNames.includes(name)) {
      throw new ExcelSourceError(`シート「${name}」がありません(スケジュール管理ブックの形式か確認してください)`);
    }
    return wb.Sheets[name];
  }

  // -------------------------------------------------------------------------
  // 読み取り
  // -------------------------------------------------------------------------

  /**
   * SheetJS のブック → {projects, tasks, members, task_masters, warnings, path, modified_at}
   * path / modified_at は画面の「データ: ファイル名(更新 …)」表示用。
   */
  function parseWorkbook(wb, { path = "", modifiedAt = "" } = {}) {
    const book = { projects: [], tasks: [], members: [], task_masters: [], warnings: [], path, modified_at: modifiedAt };
    readMembers(wb, book);
    readMasters(wb, book);
    readProjects(wb, book);
    readTasks(wb, book);
    readOtherTasks(wb, book);

    const order = new Map(book.projects.map((p, i) => [p.project_id, i]));
    book.tasks.sort(
      (a, b) =>
        order.get(a.project_id) - order.get(b.project_id) || a.task_order - b.task_order || a._row - b._row
    );
    book.tasks.forEach((t, i) => {
      t.task_id = i + 1;
      delete t._row;
    });
    return book;
  }

  function readMembers(wb, book) {
    if (!wb.SheetNames.includes(SHEET_MEMBERS)) return;
    for (const [, r] of rows(wb.Sheets[SHEET_MEMBERS])) {
      const name =
        text(get(r, "表示名")) || [text(get(r, "姓")), text(get(r, "名"))].filter(Boolean).join(" ");
      if (!name) continue;
      book.members.push({
        member_id: book.members.length + 1,
        display_name: name,
        email: text(get(r, "メールアドレス")),
        is_active: text(get(r, "状態")) !== "無効",
      });
    }
  }

  function readMasters(wb, book) {
    if (!wb.SheetNames.includes(SHEET_MASTERS)) return;
    for (const [, r] of rows(wb.Sheets[SHEET_MASTERS])) {
      const name = text(get(r, "タスク名"));
      if (!name) continue;
      book.task_masters.push({
        task_master_id: book.task_masters.length + 1,
        task_name: name,
        default_work_days: int(get(r, "標準作業日数")),
        default_task_order: int(get(r, "標準並び順")),
        requires_date_decision: text(get(r, "日付決定有無")) === "Yes",
        auto_generate: text(get(r, "WS自動生成")) !== "対象外",
      });
    }
  }

  function readProjects(wb, book) {
    const list = [];
    for (const [rowNo, r] of rows(require(wb, SHEET_WS))) {
      const name = text(get(r, "WS名"));
      if (!name) continue;
      if (list.some((p) => p.project_name === name)) {
        book.warnings.push(`[WS] ${rowNo}行目: WS名「${name}」が重複しています(2件目以降は無視)`);
        continue;
      }
      const sortOrder = int(get(r, "並び順"));
      list.push({
        project_id: list.length + 1,
        project_name: name,
        project_category: text(get(r, "種別")) || CATEGORY_WORKSHOP,
        sort_order: sortOrder !== null ? sortOrder : 9999,
        start_date: toDate(get(r, "開始日")),
        status: text(get(r, "ステータス")),
      });
    }
    list.sort((a, b) => a.sort_order - b.sort_order || a.project_id - b.project_id);
    book.projects = list;
  }

  /** 担当者名 → {member_id, display_name}。メンバーシートに無い名前も表示はする。 */
  function memberRefs(book, names) {
    const byName = new Map();
    book.members.forEach((m) => byName.set(m.display_name, m));
    return names.map((n) => {
      let m = byName.get(n);
      if (!m) {
        m = { member_id: 100000 + byName.size, display_name: n };
        byName.set(n, m);
        book.members.push({ ...m, email: null, is_active: true, unregistered: true });
      }
      return { member_id: m.member_id, display_name: m.display_name };
    });
  }

  function makeTask(book, project, rowNo, name, order, days, start, end, names, extra) {
    return {
      _row: rowNo,
      project_id: project.project_id,
      task_name: name,
      task_order: order !== null ? order : 9999,
      work_days: days,
      start_date: start,
      end_date: end,
      is_undecided: start === null,
      assignees: memberRefs(book, names),
      ...extra,
    };
  }

  function readTasks(wb, book) {
    const projects = new Map(book.projects.map((p) => [p.project_name, p]));
    const ws = require(wb, SHEET_TASKS);
    const workshops = book.projects.filter((p) => p.project_category === CATEGORY_WORKSHOP);
    const hasMasters = book.task_masters.some((m) => m.auto_generate);
    let found = 0;
    let calculated = 0;
    const noFormula = []; // 計算の数式が入っていない行(途中に挿入した行など)
    for (const [rowNo, r] of rows(ws)) {
      const wsName = text(get(r, "WS名"));
      if (!wsName) continue;
      found += 1;
      if (text(get(r, "開始日の決め方")) === null) noFormula.push(rowNo);
      else calculated += 1;
      const project = projects.get(wsName);
      if (!project) {
        book.warnings.push(`[タスク一覧] ${rowNo}行目: WS「${wsName}」がWSシートにありません`);
        continue;
      }
      const start = toDate(get(r, "開始日"));
      const days = int(get(r, "作業日数")) || 1;
      const end = toDate(get(r, "終了日")) || (start ? addDays(start, days - 1) : null);
      const names = (text(get(r, "担当者")) || "")
        .split(ASSIGNEE_SEPARATOR)
        .map((n) => n.trim())
        .filter(Boolean);
      book.tasks.push(
        makeTask(book, project, rowNo, text(get(r, "タスク名")), int(get(r, "並び順")), days, start, end, names, {
          requires_date_decision: text(get(r, "日付決定有無")) === "Yes",
          start_mode: text(get(r, "開始日の決め方")),
          status: status(get(r, "ステータス")),
          warning: text(get(r, "警告")),
          memo: text(get(r, "メモ")),
          source: "タスク一覧",
        })
      );
    }
    const notCalculated = () =>
      new ExcelNotCalculatedError(
        "タスク一覧に計算結果がありません。Excel(ブラウザ版可)でファイルを一度開いて保存してから再読み込みしてください。"
      );
    // タスクの行はあるのに自動計算の列がすべて空 = 数式の計算結果がファイルに保存されていない
    if (found && calculated === 0) throw notCalculated();
    // 旧形式: ワークショップがあるのにタスク一覧が空 = 計算結果が保存されていない
    if (wb.SheetNames.includes(SHEET_OLD_INPUT) && workshops.length && hasMasters && found === 0) throw notCalculated();
    for (const rowNo of noFormula) {
      book.warnings.push(
        `[タスク一覧] ${rowNo}行目: 自動計算の数式が入っていません(途中に挿入した行は、一番下の空き行に追加し直してください)`
      );
    }
    // タスクをまだ生成していないワークショップ
    const withTasks = new Set(book.tasks.map((t) => t.project_id));
    for (const p of workshops) {
      if (!withTasks.has(p.project_id)) {
        book.warnings.push(
          `[WS] 「${p.project_name}」のタスクがまだありません(Excel の「タスク生成」シートで作ってタスク一覧に貼り付けてください)`
        );
      }
    }
  }

  /**
   * 個別タスクの WS名が空欄のときに入れる「その他」の行。
   * WSシートに種別「その他」の行があればそれ(複数あれば並び順が最初のもの)、無ければ
   * 「その他」という名前の行、それも無ければ「その他」の行を自動で作る(画面では一番下に出る)。
   */
  function otherProject(book) {
    const byCategory = book.projects.find((p) => p.project_category === CATEGORY_OTHER);
    if (byCategory) return byCategory;
    const byName = book.projects.find((p) => p.project_name === CATEGORY_OTHER);
    if (byName) return byName;
    const project = {
      project_id: Math.max(0, ...book.projects.map((p) => p.project_id)) + 1,
      project_name: CATEGORY_OTHER,
      project_category: CATEGORY_OTHER,
      sort_order: 9999,
      start_date: null,
      status: null,
    };
    book.projects.push(project);
    return project;
  }

  function readOtherTasks(wb, book) {
    if (!wb.SheetNames.includes(SHEET_OTHER)) return;
    const projects = new Map(book.projects.map((p) => [p.project_name, p]));
    for (const [rowNo, r] of rows(wb.Sheets[SHEET_OTHER])) {
      const wsName = text(get(r, "WS名"));
      const name = text(get(r, "タスク名"));
      if (!wsName && !name) continue;
      const where = `[個別タスク] ${rowNo}行目`;
      if (!name) {
        book.warnings.push(`${where}: タスク名が空欄です`);
        continue;
      }
      // WS名が空欄 = WSに属さない作業 →「その他」
      const project = !wsName ? otherProject(book) : projects.get(wsName);
      if (!project) {
        book.warnings.push(`${where}: WS「${wsName}」がWSシートにありません`);
        continue;
      }
      const start = toDate(get(r, "開始日"));
      let end = toDate(get(r, "終了日"));
      let days = int(get(r, "作業日数")) || 1;
      let warning = null;
      if (start && end) {
        days = diffDays(start, end) + 1;
        if (end < start) warning = "終了日が開始日より前です";
      } else if (start) {
        end = addDays(start, days - 1);
      } else {
        end = null; // 開始日が無ければ未定
      }
      const names = [];
      for (let i = 1; i <= MAX_ASSIGNEES; i++) {
        const n = text(get(r, `担当者${i}`));
        if (n) names.push(n);
      }
      book.tasks.push(
        makeTask(book, project, 10000 + rowNo, name, null, days, start, end, names, {
          requires_date_decision: false,
          start_mode: "手動",
          status: status(get(r, "ステータス")),
          warning,
          memo: text(get(r, "メモ")),
          source: "個別タスク",
        })
      );
    }
  }

  /** 画面の描画に必要な全データ(Python 版 /api/schedule と同じ形)。 */
  function scheduleData(book) {
    return {
      projects: book.projects.map((p) => ({
        project_id: p.project_id,
        project_name: p.project_name,
        project_category: p.project_category,
      })),
      tasks: book.tasks,
      warnings: book.warnings,
      source: { path: book.path, modified_at: book.modified_at },
    };
  }

  return { parseWorkbook, scheduleData, ExcelSourceError, ExcelNotCalculatedError, _internal: { toDate, int, text } };
});
