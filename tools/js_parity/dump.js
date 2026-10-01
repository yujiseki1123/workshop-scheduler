// docs/js/excel_source.js でブックを読み、結果を JSON で出す(compare.py から呼ぶ)
const fs = require("fs");
const path = require("path");
const XLSX = require("xlsx");
const ES = require(path.join(__dirname, "..", "..", "docs", "js", "excel_source.js"));

const file = process.argv[2];
try {
  const book = ES.parseWorkbook(XLSX.read(fs.readFileSync(file)), { path: file });
  const data = ES.scheduleData(book);
  data.source = null;
  console.log(JSON.stringify({ data, members: book.members, task_masters: book.task_masters }));
} catch (e) {
  console.log(JSON.stringify({ error: `${e.constructor.name}: ${e.message}` }));
}
