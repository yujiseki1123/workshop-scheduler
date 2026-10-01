/**
 * スケジュール画面の描画ロジック。
 *
 * データは OneDrive の Excel を onedrive.js で取得し、excel_source.js で
 * {projects: [...], tasks: [...]}(Flask 版の /api/schedule と同じ形)に変換したもの。
 * 週割り・グリッド化・現在日線の位置計算など「見た目の組み立て」はこのファイルで行う。
 *
 * 表示モード(縦軸の切り方)は2種類:
 *   - "project"  : 行=WS(プロジェクト)(デフォルト)。
 *   - "assignee" : 担当者 → WS の2階層。1つのタスクに複数担当者がいる場合は、
 *                  該当する担当者すべてのグループにバーを複製して表示する
 *                  (担当者ごとの持ちタスク量を横断的に見えるようにするため)。
 * 同じ行の中で期間が重なるタスクは段(レーン)を分けて並べる。段は最大 MAX_LANES(5)段。
 * 開始日が未定のタスク(start_date=null)はバーを出さず、WS名の下に「未定 N件」を表示する
 * (ホバーでタスク名の一覧)。
 * どちらのモードでも取得済みのAPIレスポンス(projects/tasks)をそのまま使い、
 * 再フェッチはしない(表示切替はクライアント側の並び替えに過ぎないため)。
 */

// 表示サイズ(PC / スマホ)ごとの寸法。applyDeviceLayout() で切り替える(CSS変数も合わせて変える)
const LAYOUTS = {
  pc: { dayWidth: 56, assigneeCol: 110, wsCol: 140, rowHeight: 44, laneHeight: 26, laneGap: 4, lanePad: 5, barFont: 12 },
  mobile: { dayWidth: 34, assigneeCol: 64, wsCol: 88, rowHeight: 36, laneHeight: 22, laneGap: 3, lanePad: 4, barFont: 11 },
};
let DAY_WIDTH = LAYOUTS.pc.dayWidth; // CSSの --day-width と合わせる
const WEEKDAY_LABELS = ["月", "火", "水", "木", "金", "土", "日"]; // 月曜始まり
const PROJECT_COLORS = ["#c9d9f2", "#cfc3e0", "#f0cbd8", "#c9ecd7", "#f5e3b3"];
const ASSIGNEE_ROW_COLOR = "#e8e8ec"; // 担当者別表示の行ラベル背景(プロジェクト色と混同しないよう単色)
const UNASSIGNED_KEY = "__unassigned__";
// プロジェクト種別「その他」(app/services/excel_source.py の CATEGORY_OTHER と合わせる)。
// WSに属さない作業の受け皿で、担当者別表示では全担当者に常に行を表示する。
const CATEGORY_OTHER = "その他";
const OTHER_COLOR = "#e2e2e2"; // 「その他」はWSの色と区別するためグレー
// タスクのステータスごとのバーの色(未着手=黄 / 着手中=青 / 完了=グレー)
const STATUS_STYLES = {
  未着手: { background: "#ffe38a", color: "#3d3200" },
  着手中: { background: "#8fbcf0", color: "#0b2545" },
  完了: { background: "#d6d6d6", color: "#6b6b6b" },
};
const DEFAULT_STATUS = "未着手";
// 日付決定有無=Yes のタスク(リハーサル・WS実施など)はステータスにかかわらず赤で表示する
const DECISION_STYLE = { background: "#f28b82", color: "#5c0000" };
let BAR_FONT = "12px -apple-system, sans-serif";
// 行ラベル列の幅。プロジェクト別=「WS」、担当者別=「担当者 | WS」
let ASSIGNEE_COL_WIDTH = LAYOUTS.pc.assigneeCol;
let WS_COL_WIDTH = LAYOUTS.pc.wsCol;

function labelColumns(mode) {
  return mode === "assignee" ? [ASSIGNEE_COL_WIDTH, WS_COL_WIDTH] : [WS_COL_WIDTH];
}

function labelWidthFor(mode) {
  return labelColumns(mode).reduce((a, b) => a + b, 0);
}

/**
 * 与えられたテキストが指定フォントで実際に何pxになるかを計測する。
 * 固定のしきい値ではなく実測することで、担当者名の長さが変わっても
 * 「入りきらない時だけ "*" にする」判定を正しく行える。
 */
function measureTextWidth(text, font) {
  const canvas =
    measureTextWidth._canvas || (measureTextWidth._canvas = document.createElement("canvas"));
  const ctx = canvas.getContext("2d");
  ctx.font = font;
  return ctx.measureText(text).width;
}

function parseISODate(str) {
  // "YYYY-MM-DD" をローカルタイムのDateとして解釈する(タイムゾーンずれ防止のため
  // new Date("YYYY-MM-DD") ではなく年月日を分解して組み立てる)。
  const [y, m, d] = str.split("-").map(Number);
  return new Date(y, m - 1, d);
}

function addDays(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function diffDays(from, to) {
  const ms = to.setHours(0, 0, 0, 0) - new Date(from).setHours(0, 0, 0, 0);
  return Math.round(ms / 86400000);
}

function mondayOfWeek(date) {
  const day = date.getDay(); // 0=日,1=月,...6=土
  const offset = day === 0 ? -6 : 1 - day;
  return addDays(date, offset);
}

function isoWeekNumber(date) {
  // ISO 8601 週番号(画像の "Week 38" のような表示に対応)
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  const dayNum = d.getUTCDay() || 7;
  d.setUTCDate(d.getUTCDate() + 4 - dayNum);
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil(((d - yearStart) / 86400000 + 1) / 7);
}

function formatMonthDay(date) {
  return `${date.getMonth() + 1}/${date.getDate()}`;
}

/**
 * OneDrive の Excel を取得して画面用のデータにする(Flask 版の /api/schedule と同じ形)。
 * Excel が無い・形式が違う・計算結果が無い場合は、理由を書いたエラーを投げる。
 */
async function fetchSchedule() {
  const file = await OneDriveSource.loadWorkbook();
  const book = ExcelSource.parseWorkbook(file.workbook, { path: file.name, modifiedAt: file.modifiedAt });
  const data = ExcelSource.scheduleData(book);
  data.source.modified_by = file.modifiedBy;
  return data;
}

/**
 * 表示する期間。既定は「今週の weeksBehind 週前 〜 weeksAhead 週先」だが、
 * それより前・後に予定のあるタスクがあれば、そのタスクの週まで広げる
 * (タスクのバーだけが日付の列の外にはみ出さないようにするため)。
 */
function buildTimeline(today, weeksBehind, weeksAhead, tasks = []) {
  let start = addDays(mondayOfWeek(today), -weeksBehind * 7);
  let end = addDays(start, (weeksBehind + weeksAhead) * 7); // 表示期間の翌日
  tasks.forEach((t) => {
    if (!t.start_date || !t.end_date) return; // 未定のタスクはバーを出さない
    const s = mondayOfWeek(parseISODate(t.start_date));
    const e = addDays(mondayOfWeek(parseISODate(t.end_date)), 7);
    if (s < start) start = s;
    if (e > end) end = e;
  });
  const totalDays = diffDays(start, end);
  const days = [];
  for (let i = 0; i < totalDays; i++) {
    days.push(addDays(start, i));
  }
  return { start, totalDays, days };
}

function buildHeader(days, cornerLabel, labelWidth) {
  const headerRow = document.createElement("div");
  headerRow.className = "header-row";

  const corner = document.createElement("div");
  corner.className = "corner-cell";
  corner.style.width = corner.style.minWidth = `${labelWidth}px`;
  headerRow.appendChild(corner);

  // 週ごとにグループ化して "Week NN" ラベルを出す
  let i = 0;
  while (i < days.length) {
    const weekBlock = document.createElement("div");
    weekBlock.className = "week-block";
    weekBlock.style.width = `${DAY_WIDTH * 7}px`;
    weekBlock.textContent = `Week ${isoWeekNumber(days[i])}`;
    headerRow.appendChild(weekBlock);
    i += 7;
  }

  const dayRow = document.createElement("div");
  dayRow.className = "day-header-row";
  const dayCorner = document.createElement("div");
  dayCorner.className = "corner-cell";
  dayCorner.style.width = dayCorner.style.minWidth = `${labelWidth}px`;
  dayCorner.classList.add("corner-columns");
  // 行ラベルの各列の見出し(例: 「WS」「タスク」)
  (cornerLabel || []).forEach(([text, width]) => {
    const c = document.createElement("div");
    c.className = "corner-col";
    c.style.width = `${width}px`;
    c.textContent = text;
    dayCorner.appendChild(c);
  });
  dayRow.appendChild(dayCorner);

  days.forEach((d) => {
    const cell = document.createElement("div");
    const weekdayIdx = (d.getDay() + 6) % 7; // 月=0 ... 日=6
    cell.className = "day-cell" + (weekdayIdx >= 5 ? " weekend" : "");
    cell.innerHTML = `<span class="date-num">${formatMonthDay(d)}</span><span class="weekday">${WEEKDAY_LABELS[weekdayIdx]}</span>`;
    dayRow.appendChild(cell);
  });

  const wrapper = document.createElement("div");
  wrapper.appendChild(headerRow);
  wrapper.appendChild(dayRow);
  return wrapper;
}

function tooltipText(task, projectName) {
  const names = task.assignees.map((a) => a.display_name).join("、") || "(未割当)";
  return (
    `${projectName} / ${task.task_name}\n` +
    `期間: ${task.start_date} 〜 ${task.end_date} (${task.work_days}日間)\n` +
    `担当者: ${names}\n` +
    `ステータス: ${task.status || DEFAULT_STATUS}` +
    (task.warning ? `\n⚠ ${task.warning}` : "") +
    (task.memo ? `\nメモ: ${task.memo}` : "")
  );
}

function showTooltip(evt, text) {
  const tooltip = document.getElementById("tooltip");
  tooltip.textContent = text;
  tooltip.hidden = false;
  positionTooltip(evt);
}

function positionTooltip(evt) {
  const tooltip = document.getElementById("tooltip");
  const offset = 14;
  let left = evt.clientX + offset;
  let top = evt.clientY + offset;
  tooltip.style.left = `${left}px`;
  tooltip.style.top = `${top}px`;
  tooltip.style.whiteSpace = "pre-line";
}

function hideTooltip() {
  if (currentDevice === "mobile") return; // スマホはタップで開閉する
  document.getElementById("tooltip").hidden = true;
}

/** スマホ表示: タップした要素の詳細を画面下に出す(ホバーが無いため)。もう一度どこかをタップすると閉じる。 */
function showSheet(evt, text) {
  evt.stopPropagation();
  const tooltip = document.getElementById("tooltip");
  tooltip.textContent = text + "\n\n(タップで閉じる)";
  tooltip.style.left = tooltip.style.top = "";
  tooltip.style.whiteSpace = "pre-line";
  tooltip.hidden = false;
}

/** ホバー(PC)とタップ(スマホ)の両方で詳細を出せるようにする。 */
function attachDetails(el, text) {
  el.addEventListener("mouseenter", (evt) => {
    if (currentDevice !== "mobile") showTooltip(evt, text);
  });
  el.addEventListener("mousemove", (evt) => {
    if (currentDevice !== "mobile") positionTooltip(evt);
  });
  el.addEventListener("mouseleave", hideTooltip);
  el.addEventListener("click", (evt) => {
    if (currentDevice === "mobile") showSheet(evt, text);
  });
}

/**
 * 担当者別表示の行一覧を組み立てる。
 *
 * タスクに登場する担当者を最初に見つかった順ではなく member_id 昇順で
 * 安定ソートし、切替のたびに行の並びがガタつかないようにする。
 * 担当者未設定のタスクが1件でもあれば "(未割当)" 行を末尾に追加する。
 */
function buildAssigneeRows(data) {
  const nameById = new Map();
  let hasUnassigned = false;

  data.tasks.forEach((task) => {
    if (task.assignees.length === 0) {
      hasUnassigned = true;
      return;
    }
    task.assignees.forEach((a) => {
      if (!nameById.has(a.member_id)) {
        nameById.set(a.member_id, a.display_name);
      }
    });
  });

  const rows = Array.from(nameById.entries())
    .sort((a, b) => a[0] - b[0])
    .map(([memberId, name]) => ({
      key: memberId,
      label: name,
      color: ASSIGNEE_ROW_COLOR,
    }));

  if (hasUnassigned) {
    rows.push({ key: UNASSIGNED_KEY, label: "(未割当)", color: ASSIGNEE_ROW_COLOR });
  }

  return rows;
}

function isOtherProject(p) {
  return p.project_category === CATEGORY_OTHER;
}

/**
 * プロジェクトの表示順。基本は API の並び(sort_order順)のままとし、
 * 種別「その他」は sort_order の値にかかわらず常に末尾へ回す。
 */
function orderedProjects(data) {
  return [
    ...data.projects.filter((p) => !isOtherProject(p)),
    ...data.projects.filter(isOtherProject),
  ];
}

/** プロジェクトID → 色。WSは PROJECT_COLORS を順に、「その他」はグレー。 */
function buildProjectColors(data) {
  const colors = new Map();
  let wsIdx = 0;
  orderedProjects(data).forEach((p) => {
    colors.set(
      p.project_id,
      isOtherProject(p) ? OTHER_COLOR : PROJECT_COLORS[wsIdx++ % PROJECT_COLORS.length]
    );
  });
  return colors;
}

// 同じ行内でタスク期間が重なった場合のレーン(段)の設定。
// 重なりが無い行は CSSの --task-row-height の1段のまま、重なる行だけ段数ぶん高さを広げる。
const MAX_LANES = 5; // 1行あたりの最大段数
let LANE_HEIGHT = LAYOUTS.pc.laneHeight;
let LANE_GAP = LAYOUTS.pc.laneGap;
let LANE_TOP_PAD = LAYOUTS.pc.lanePad;

/**
 * 1行分のタスクバー群に、重ならないよう貪欲法でレーン番号を割り当てる。
 *
 * 例えば担当者別表示では、同じ担当者が別プロジェクトの案件を同時期に
 * 掛け持ちしているとバーの期間が重なり得る(プロジェクト別表示でも
 * 理論上は起こり得る)。重なったバーを単純に同じ段へ重ねて描画すると
 * 判読不能になるため、開始位置が早い順に「空いている一番上の段」へ
 * 詰めていき、重なるバーだけ段を分けて表示する。
 * 段は最大 MAX_LANES まで。それ以上重なった場合(想定外)は一番早く空く段に
 * 重ねて置き、点線枠(.overflow)で重なっていることが分かるようにする。
 */
function assignLanes(bars) {
  const sorted = [...bars].sort((a, b) => a.left - b.left);
  const laneRightEdge = []; // 各レーンの現在の右端x(px)

  const placed = sorted.map((bar) => {
    let lane = laneRightEdge.findIndex((edge) => edge <= bar.left);
    let overflow = false;
    if (lane === -1 && laneRightEdge.length < MAX_LANES) {
      lane = laneRightEdge.length;
      // 表示期間より前に始まるタスクは left が負になるため、初期値は 0 ではなく -Infinity
      laneRightEdge.push(-Infinity);
    } else if (lane === -1) {
      // 最大段数まで埋まっている → 一番早く空く段に重ねて置く(点線枠で重なりを示す)
      lane = laneRightEdge.indexOf(Math.min(...laneRightEdge));
      overflow = true;
    }
    // はみ出したバーは段の右端を更新しない(後続の本来重ならないバーまで
    // はみ出し扱いにならないようにするため)
    if (!overflow) laneRightEdge[lane] = bar.left + bar.width;
    return { ...bar, lane, overflow };
  });

  return { bars: placed, laneCount: Math.max(laneRightEdge.length, 1) };
}

/**
 * 1本のタスクバー要素を生成する(ラベルの省略判定・ツールチップ設定込み)。
 */
function buildBarElement(bar, laneCount) {
  const el = document.createElement("div");
  el.className = "task-bar" + (bar.overflow ? " overflow" : "");
  el.style.left = `${bar.left}px`;
  el.style.width = `${bar.width}px`;
  const st = bar.decision ? DECISION_STYLE : STATUS_STYLES[bar.status];
  el.style.background = st.background;
  el.style.color = st.color;
  el.style.borderLeft = `5px solid ${bar.accent}`;
  el.dataset.status = bar.status;
  if (bar.decision) el.classList.add("decision");

  if (laneCount > 1) {
    el.style.top = `${LANE_TOP_PAD + bar.lane * (LANE_HEIGHT + LANE_GAP)}px`;
    el.style.height = `${LANE_HEIGHT}px`;
  }

  const textWidth = measureTextWidth(bar.label, BAR_FONT);
  if (bar.width < textWidth + 8) {
    // ラベルが入りきらない → "*" だけ表示し、詳細はホバー(スマホはタップ)で見せる
    el.classList.add("compact");
    el.textContent = "*";
  } else {
    el.textContent = bar.label;
  }

  attachDetails(el, bar.text);
  return el;
}

/**
 * タスクバーを並べたトラック(横軸部分)を生成し、行に付与する。
 * 期間が重なるバーがある場合は段(レーン)を分け、行の高さを広げる。
 */
function appendTrack(row, barList) {
  const track = document.createElement("div");
  track.className = "task-track";
  row.appendChild(track);

  const { bars, laneCount } = assignLanes(barList);
  if (laneCount > 1) {
    // 重なりがある行だけ、段数に応じて行の高さを広げる
    row.style.height = `${LANE_TOP_PAD * 2 + laneCount * LANE_HEIGHT + (laneCount - 1) * LANE_GAP}px`;
  }
  bars.forEach((bar) => track.appendChild(buildBarElement(bar, laneCount)));
}

function taskLabel(task) {
  const names = task.assignees.map((a) => a.display_name);
  return names.length ? `${task.task_name} (${names.join("、")})` : task.task_name;
}

/** 開始日未定タスクの件数バッジ。ホバーでタスク名の一覧を出す。 */
function buildUndecidedBadge(tasks) {
  const badge = document.createElement("span");
  badge.className = "undecided-badge";
  badge.textContent = `未定 ${tasks.length}件`;
  const text =
    "開始日が未定のタスク\n" +
    tasks
      .map((t) => {
        const names = t.assignees.map((a) => a.display_name).join("、");
        const mark = t.requires_date_decision ? "【日付決定待ち】" : "";
        return `・${mark}${t.task_name} (${t.work_days}日間)${names ? " " + names : ""}`;
      })
      .join("\n");
  attachDetails(badge, text);
  return badge;
}

/** 左側に固定表示されるラベルセルを作る。 */
function buildLabelCell(className, text, width, left, background) {
  const cell = document.createElement("div");
  cell.className = className;
  cell.style.width = cell.style.minWidth = `${width}px`;
  cell.style.left = `${left}px`;
  if (background) cell.style.background = background;
  cell.textContent = text;
  cell.title = text;
  return cell;
}

/**
 * ラベルセル(縦に結合された見た目)+子要素の縦並び、という1階層分のグループを作る。
 * 例: 「WS A」セルの右に、WS A のタスク行を縦に並べる。
 */
function buildGroup(level, labelCell, children) {
  const group = document.createElement("div");
  group.className = `row-group level-${level}`;
  group.appendChild(labelCell);
  const box = document.createElement("div");
  box.className = "group-children";
  children.forEach((c) => box.appendChild(c));
  group.appendChild(box);
  return group;
}

/**
 * mode に応じて行とタスクバーを組み立てる。
 *
 * バーのラベル: プロジェクト別=「タスク名 (担当者名)」/ 担当者別=「タスク名」
 * (担当者別ではWS名・担当者名が行ラベルに出ているため)。詳細はホバーで見せる。
 * バーの背景色はどちらのモードでも「所属プロジェクトの色」で統一している。
 */
function buildBody(data, timelineStart, todayOffset, mode) {
  const body = document.createElement("div");
  body.className = "body-area";

  const projectColor = buildProjectColors(data);
  const projectName = new Map(data.projects.map((p) => [p.project_id, p.project_name]));
  const wsColLeft = mode === "assignee" ? ASSIGNEE_COL_WIDTH : 0;

  const toBar = (task) => {
    const startOffset = diffDays(timelineStart, parseISODate(task.start_date));
    return {
      left: startOffset * DAY_WIDTH,
      width: Math.max(task.work_days * DAY_WIDTH - 4, 8), // 隣接バーとの間に少し隙間を作る
      label: mode === "assignee" ? task.task_name : taskLabel(task),
      text: tooltipText(task, projectName.get(task.project_id) || ""),
      status: STATUS_STYLES[task.status] ? task.status : DEFAULT_STATUS,
      decision: Boolean(task.requires_date_decision),
      // 左端の太線はWSの色(ステータス色のバーでもどのWSか分かるように)
      accent: projectColor.get(task.project_id) || "#ccc",
    };
  };

  /** WS 1行分(左にWS名、右にタスクバー。重なる分は最大5段)を作る。 */
  const buildWsRow = (project, tasks) => {
    const color = projectColor.get(project.project_id);
    const row = document.createElement("div");
    row.className = "task-row";
    const label = buildLabelCell("group-label ws-label", project.project_name, WS_COL_WIDTH, wsColLeft, color);
    const undecided = tasks.filter((t) => t.is_undecided);
    if (undecided.length) label.appendChild(buildUndecidedBadge(undecided));
    row.appendChild(label);
    appendTrack(row, tasks.filter((t) => !t.is_undecided).map(toBar));
    return row;
  };

  const tasksByProject = (tasks) => {
    const m = new Map();
    tasks.forEach((t) => {
      if (!m.has(t.project_id)) m.set(t.project_id, []);
      m.get(t.project_id).push(t);
    });
    return m;
  };

  if (mode === "assignee") {
    // 担当者キー → タスク配列。複数担当者のタスクは全員に複製し、未割当は "(未割当)" へ。
    const rowsMeta = buildAssigneeRows(data);
    const tasksByAssignee = new Map(rowsMeta.map((r) => [r.key, []]));
    data.tasks.forEach((task) => {
      const keys = task.assignees.length ? task.assignees.map((a) => a.member_id) : [UNASSIGNED_KEY];
      keys.forEach((key) => tasksByAssignee.get(key)?.push(task));
    });

    rowsMeta.filter((r) => isSelected("assignee", r.key)).forEach((r) => {
      const byProject = tasksByProject(tasksByAssignee.get(r.key));
      // WSの並びはプロジェクト別表示と同じ(「その他」は末尾)。
      // タスクがあるWSのみ出すが、「その他」は全担当者に常に表示する(未割当は除く)。
      const wsRows = orderedProjects(data)
        .filter((p) => byProject.has(p.project_id) || (isOtherProject(p) && r.key !== UNASSIGNED_KEY))
        .map((p) => buildWsRow(p, byProject.get(p.project_id) || []));
      body.appendChild(
        buildGroup(1, buildLabelCell("group-label", r.label, ASSIGNEE_COL_WIDTH, 0, r.color), wsRows)
      );
    });
  } else {
    const byProject = tasksByProject(data.tasks);
    orderedProjects(data).filter((p) => isSelected("project", p.project_id)).forEach((p) => {
      const row = buildWsRow(p, byProject.get(p.project_id) || []);
      row.classList.add("top-row");
      body.appendChild(row);
    });
  }

  // 現在日の縦線。日付セルの左端ではなく中央に来るよう半日分ずらす
  // (線幅2pxの半分も引いて、線の中心がセル中央と一致するようにする)。
  const line = document.createElement("div");
  line.className = "current-date-line";
  line.style.left = `${labelWidthFor(mode) + todayOffset * DAY_WIDTH + DAY_WIDTH / 2 - 1}px`;
  body.appendChild(line);

  return body;
}

// 直近のfetch結果とタイムライン計算をキャッシュしておき、表示切替時に
// 再フェッチ・再計算をしなくて済むようにする(モード切替はクライアント側の
// 並び替えだけなので、サーバーに問い合わせる必要がない)。
let scheduleCache = null; // { data, start, days, todayOffset }
let currentMode = "project"; // "project" | "assignee"
let currentDevice = "pc"; // "pc" | "mobile"(表示サイズ。切替ボタンで変更し、ブラウザに記憶する)

// 表示の絞り込み。モードごとに「選択中のキーの集合」を持つ。null = すべて表示。
// プロジェクト別 = WS(project_id)、担当者別 = 担当者(member_id / 未割当)。
const filters = { project: null, assignee: null };

function isSelected(mode, key) {
  return filters[mode] === null || filters[mode].has(key);
}

/** 現在のモードで絞り込める選択肢(キーと表示名)。 */
function filterOptions(data, mode) {
  if (mode === "assignee") {
    return buildAssigneeRows(data).map((r) => ({ key: r.key, label: r.label }));
  }
  return orderedProjects(data).map((p) => ({ key: p.project_id, label: p.project_name }));
}

/** 絞り込み(複数選択のチェックボックス)とステータスの凡例を描く。 */
function renderFilterBar() {
  const bar = document.getElementById("filter-bar");
  if (!bar || !scheduleCache) return;
  const mode = currentMode;
  const options = filterOptions(scheduleCache.data, mode);
  const selected = filters[mode];
  const title = mode === "assignee" ? "担当者" : "WS";
  const count = selected === null ? options.length : options.filter((o) => selected.has(o.key)).length;

  const details = document.createElement("details");
  details.className = "filter";
  // 絞り込みを操作している間は開いたまま。表示モードを切り替えたら閉じる
  details.open = bar.dataset.mode === mode && (bar.querySelector("details")?.open ?? false);
  bar.dataset.mode = mode;
  const summary = document.createElement("summary");
  summary.textContent =
    selected === null || count === options.length
      ? `${title}で絞り込み(すべて表示)`
      : `${title}で絞り込み(${count} / ${options.length} 件を表示)`;
  details.appendChild(summary);

  const panel = document.createElement("div");
  panel.className = "filter-panel";
  const actions = document.createElement("div");
  actions.className = "filter-actions";
  const allBtn = document.createElement("button");
  allBtn.type = "button";
  allBtn.textContent = "すべて選択";
  allBtn.addEventListener("click", () => {
    filters[mode] = null;
    renderGrid({ preserveScroll: true });
  });
  const noneBtn = document.createElement("button");
  noneBtn.type = "button";
  noneBtn.textContent = "すべて解除";
  noneBtn.addEventListener("click", () => {
    filters[mode] = new Set();
    renderGrid({ preserveScroll: true });
  });
  actions.append(allBtn, noneBtn);
  panel.appendChild(actions);

  options.forEach((o) => {
    const label = document.createElement("label");
    const cb = document.createElement("input");
    cb.type = "checkbox";
    cb.checked = isSelected(mode, o.key);
    cb.addEventListener("change", () => {
      const next = filters[mode] === null ? new Set(options.map((x) => x.key)) : new Set(filters[mode]);
      if (cb.checked) next.add(o.key);
      else next.delete(o.key);
      filters[mode] = next.size === options.length ? null : next;
      renderGrid({ preserveScroll: true });
    });
    label.append(cb, document.createTextNode(o.label));
    panel.appendChild(label);
  });
  details.appendChild(panel);

  const legend = document.createElement("div");
  legend.className = "status-legend";
  Object.entries({ ...STATUS_STYLES, 日付決定タスク: DECISION_STYLE }).forEach(([name, st]) => {
    const item = document.createElement("span");
    item.className = "legend-item";
    const swatch = document.createElement("span");
    swatch.className = "legend-swatch";
    swatch.style.background = st.background;
    item.append(swatch, document.createTextNode(name));
    legend.appendChild(item);
  });

  bar.replaceChildren(details, legend);
}

/**
 * #schedule-root の中身(グリッド全体)を現在のキャッシュ+modeで再構築する。
 * preserveScroll=true のときは呼び出し前の横スクロール位置を維持する
 * (表示切替のたびに現在日位置へ戻ってしまうと使いづらいため)。
 * false のとき(初回描画)は現在日付近が見えるようスクロール位置を調整する。
 */
function renderGrid({ preserveScroll = false } = {}) {
  if (!scheduleCache) return;
  const root = document.getElementById("schedule-root");
  const { data, start, days, todayOffset } = scheduleCache;

  const scrollLeftBefore = root.scrollLeft;

  const grid = document.createElement("div");
  grid.className = "schedule-grid";
  const labelWidth = labelWidthFor(currentMode);
  grid.style.width = `${labelWidth + days.length * DAY_WIDTH}px`;
  const cornerTitles = currentMode === "assignee" ? ["担当者", "WS"] : ["WS"];
  const cornerLabel = cornerTitles.map((t, i) => [t, labelColumns(currentMode)[i]]);
  grid.appendChild(buildHeader(days, cornerLabel, labelWidth));
  grid.appendChild(buildBody(data, start, todayOffset, currentMode));
  renderFilterBar();

  root.innerHTML = "";
  root.appendChild(grid);

  if (preserveScroll) {
    root.scrollLeft = scrollLeftBefore;
  } else {
    root.scrollLeft = Math.max(0, (todayOffset - 1) * DAY_WIDTH);
  }
}

const DEVICE_KEY = "workshopScheduler.device";

/** 最初の表示サイズ: 前回選んだもの。無ければ画面幅で決める(幅 768px 以下はスマホ)。 */
function initialDevice() {
  try {
    const saved = localStorage.getItem(DEVICE_KEY);
    if (saved === "pc" || saved === "mobile") return saved;
  } catch (_) {
    /* 保存できない環境 */
  }
  return window.matchMedia("(max-width: 768px)").matches ? "mobile" : "pc";
}

/** 表示サイズの寸法を JS の定数と CSS 変数の両方に反映する。 */
function applyDeviceLayout(device) {
  currentDevice = device;
  const L = LAYOUTS[device];
  DAY_WIDTH = L.dayWidth;
  ASSIGNEE_COL_WIDTH = L.assigneeCol;
  WS_COL_WIDTH = L.wsCol;
  LANE_HEIGHT = L.laneHeight;
  LANE_GAP = L.laneGap;
  LANE_TOP_PAD = L.lanePad;
  BAR_FONT = `${L.barFont}px -apple-system, sans-serif`;
  const root = document.documentElement.style;
  root.setProperty("--day-width", `${L.dayWidth}px`);
  root.setProperty("--task-row-height", `${L.rowHeight}px`);
  root.setProperty("--bar-font-size", `${L.barFont}px`);
  document.body.classList.toggle("mobile", device === "mobile");
  document.getElementById("tooltip").hidden = true;
  const btn = document.getElementById("device-toggle-btn");
  if (btn) {
    btn.textContent = device === "mobile" ? "PC表示" : "スマホ表示";
    btn.setAttribute("aria-pressed", device === "mobile" ? "true" : "false");
  }
}

function initDeviceToggle() {
  applyDeviceLayout(initialDevice());
  document.getElementById("device-toggle-btn").addEventListener("click", () => {
    const next = currentDevice === "mobile" ? "pc" : "mobile";
    try {
      localStorage.setItem(DEVICE_KEY, next);
    } catch (_) {
      /* 保存できなくても切替はする */
    }
    applyDeviceLayout(next);
    // 寸法が変わるので描き直し、今日の付近へスクロールし直す
    renderGrid({ preserveScroll: false });
  });
  // スマホ表示: 詳細の外をタップしたら閉じる
  document.addEventListener("click", () => {
    if (currentDevice === "mobile") document.getElementById("tooltip").hidden = true;
  });
}

/** 右上の「Excelを開く」: 共有リンクを新しいタブで開く(ブラウザ版 Excel / スマホは Excel アプリ)。 */
function initExcelLink() {
  const link = document.getElementById("excel-link");
  const { shareUrl } = OneDriveSource.settings();
  if (link && shareUrl) {
    link.href = shareUrl;
    link.hidden = false;
  }
}

function updateToggleButtonLabel(btn) {
  const isAssignee = currentMode === "assignee";
  btn.textContent = isAssignee ? "表示切替: プロジェクト別" : "表示切替: 担当者別";
  btn.setAttribute("aria-pressed", isAssignee ? "true" : "false");
}

function initViewToggle() {
  const btn = document.getElementById("view-toggle-btn");
  if (!btn) return;
  updateToggleButtonLabel(btn);
  btn.addEventListener("click", () => {
    if (!scheduleCache) return; // データ取得前は何もしない
    currentMode = currentMode === "project" ? "assignee" : "project";
    updateToggleButtonLabel(btn);
    renderGrid({ preserveScroll: true });
  });
}

/** 読み込んだExcelファイル名・更新日時と、ブック全体の注意(読み飛ばした行・日付の警告)を表示する。 */
function renderSourceInfo(data) {
  const info = document.getElementById("source-info");
  if (info && data.source) {
    const file = data.source.path.split(/[\\/]/).pop();
    const by = data.source.modified_by ? ` ${data.source.modified_by}` : "";
    info.textContent = `データ: ${file}(更新 ${data.source.modified_at.replace("T", " ")}${by})`;
    if (data.projects.length === 0) {
      info.textContent += " / WSがまだ登録されていません。Excelの「WS」シートにワークショップを追加してください。";
    }
  }
  const box = document.getElementById("warnings");
  if (!box) return;
  const taskWarnings = data.tasks
    .filter((t) => t.warning)
    .map((t) => {
      const p = data.projects.find((x) => x.project_id === t.project_id);
      return `${p ? p.project_name : ""}「${t.task_name}」: ${t.warning}`;
    });
  const all = [...(data.warnings || []), ...taskWarnings];
  box.hidden = all.length === 0;
  box.innerHTML = "";
  if (all.length) {
    const title = document.createElement("strong");
    title.textContent = `注意 ${all.length}件`;
    const list = document.createElement("ul");
    all.forEach((w) => {
      const li = document.createElement("li");
      li.textContent = w;
      list.appendChild(li);
    });
    box.append(title, list);
  }
}

function showMessage(text, { withReset = false } = {}) {
  const p = document.createElement("p");
  p.className = "loading";
  p.textContent = text;
  const nodes = [p];
  if (withReset) nodes.push(buildResetButton());
  document.getElementById("schedule-root").replaceChildren(...nodes);
}

/** 保存した設定(クライアント ID・共有リンク)とサインイン情報をこのブラウザから消す。 */
function resetSettings() {
  OneDriveSource.saveSettings({ clientId: "", shareUrl: "" });
  try {
    // MSAL がこのブラウザに保存したサインイン情報(msal. で始まるキー)
    Object.keys(localStorage).filter((k) => k.startsWith("msal.") || k.includes("login.windows.net") || k.includes("login.microsoftonline.com")).forEach((k) => localStorage.removeItem(k));
    sessionStorage.clear();
  } catch (_) {
    /* 消せない環境でも続ける */
  }
}

function buildResetButton() {
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "header-btn reset-btn";
  btn.textContent = "設定をやり直す(保存した設定を消す)";
  btn.addEventListener("click", () => {
    resetSettings();
    window.location.replace(window.location.pathname);
  });
  return btn;
}

// localhost 以外の書き方(127.0.0.1 / 0.0.0.0 / [::])で開いたときは localhost に移す。
// サインイン後の戻り先(リダイレクト URI)は localhost で登録しているうえ、[::] や 0.0.0.0 は
// ブラウザが「安全な接続」と見なさず、サインインに必要な暗号機能(crypto)が使えないため。
const LOCAL_ALIASES = ["127.0.0.1", "0.0.0.0", "[::]", "[::1]", "::", "::1"];

function redirectToLocalhostIfNeeded() {
  if (!LOCAL_ALIASES.includes(window.location.hostname)) return false;
  const url = new URL(window.location.href);
  url.hostname = "localhost";
  window.location.replace(url.href);
  return true;
}

/** preserveScroll=true は「再読み込み」ボタン(表示モード・絞り込み・横スクロール位置を保つ)。 */
async function renderSchedule({ preserveScroll = false } = {}) {
  const cfg = window.APP_CONFIG || {};
  const now = new Date();
  const today = new Date(now.getFullYear(), now.getMonth(), now.getDate());
  const weeksAhead = cfg.weeksAhead ?? 10;
  const weeksBehind = cfg.weeksBehind ?? 1;

  if (!preserveScroll) showMessage("OneDrive から Excel を読み込み中...");
  let data;
  try {
    data = await fetchSchedule();
  } catch (err) {
    showMessage(`読み込みに失敗しました: ${err.message}`, { withReset: !OneDriveSource.settings().fromConfig });
    return;
  }

  renderSourceInfo(data);

  const { start, days } = buildTimeline(today, weeksBehind, weeksAhead, data.tasks);
  const todayOffset = diffDays(start, today);

  scheduleCache = { data, start, days, todayOffset };
  renderGrid({ preserveScroll });
}

/** 設定(クライアント ID・共有リンク)の入力欄。config.js が空のときだけ出す。 */
function renderSettingsForm() {
  const current = OneDriveSource.settings();
  const form = document.createElement("form");
  form.className = "auth-panel";
  form.innerHTML = `
    <p>最初に設定を入力してください(このブラウザにだけ保存されます。公開時は js/config.js に書きます)。</p>
    <label>アプリケーション (クライアント) ID<input name="clientId" type="text" required></label>
    <label>Excel の共有リンク<input name="shareUrl" type="text" required></label>
    <button type="submit" class="header-btn">保存</button>`;
  form.clientId.value = current.clientId;
  form.shareUrl.value = current.shareUrl;
  form.addEventListener("submit", (e) => {
    e.preventDefault();
    OneDriveSource.saveSettings({ clientId: form.clientId.value, shareUrl: form.shareUrl.value });
    window.location.reload();
  });
  document.getElementById("schedule-root").replaceChildren(form);
}

function renderSignIn() {
  const panel = document.createElement("div");
  panel.className = "auth-panel";
  const p = document.createElement("p");
  p.textContent =
    "スケジュールを見るには、Excel を共有されている Microsoft アカウント(個人用)でサインインしてください。";
  const btn = document.createElement("button");
  btn.type = "button";
  btn.className = "header-btn primary";
  btn.textContent = "Microsoft アカウントでサインイン";
  btn.addEventListener("click", () => OneDriveSource.signIn().catch((e) => showMessage(e.message)));
  panel.append(p, btn);
  if (!OneDriveSource.settings().fromConfig) {
    const edit = document.createElement("button");
    edit.type = "button";
    edit.className = "link-btn";
    edit.textContent = "設定を変更";
    edit.addEventListener("click", renderSettingsForm);
    panel.append(edit);
  }
  document.getElementById("schedule-root").replaceChildren(panel);
}

/** サインインの状態をヘッダーに出す。 */
function updateAccountBar(account) {
  document.getElementById("account-name").textContent = account ? account.username : "";
  document.getElementById("reload-btn").hidden = !account;
  document.getElementById("signout-btn").hidden = !account;
}

function initAccountButtons() {
  document.getElementById("reload-btn").addEventListener("click", () => {
    renderSchedule({ preserveScroll: Boolean(scheduleCache) });
  });
  document.getElementById("signout-btn").addEventListener("click", () => OneDriveSource.signOut());
}

/** 起動: 設定 → サインイン → Excel の読み込み、の順に必要なものを確認する。 */
async function start() {
  if (redirectToLocalhostIfNeeded()) return;
  const params = new URLSearchParams(window.location.search);
  if (params.has("reset")) {
    // ?reset 付きで開くと保存した設定を消してやり直す
    resetSettings();
    window.location.replace(window.location.pathname);
    return;
  }
  if (!window.isSecureContext) {
    showMessage(
      "このアドレスではサインインできません(https か http://localhost で開く必要があります)。" +
        "手元で試すときは http://localhost:8000/ を開いてください。"
    );
    return;
  }
  const { clientId, shareUrl } = OneDriveSource.settings();
  if (!clientId || !shareUrl) {
    renderSettingsForm();
    return;
  }
  let account;
  try {
    account = await OneDriveSource.init();
  } catch (e) {
    showMessage(`サインインの処理に失敗しました: ${e.message}`, { withReset: true });
    return;
  }
  updateAccountBar(account);
  if (!account) {
    renderSignIn();
    return;
  }
  renderSchedule();
}

// 絞り込みパネルの外をクリックしたら閉じる
document.addEventListener("click", (evt) => {
  const open = document.querySelector("#filter-bar details[open]");
  // パネル内の操作で描き直された要素(もうページに無い)のクリックは「外」とみなさない
  if (open && evt.target.isConnected && !open.contains(evt.target)) open.open = false;
});

document.addEventListener("DOMContentLoaded", () => {
  initDeviceToggle();
  initViewToggle();
  initAccountButtons();
  initExcelLink();
  start();
});
