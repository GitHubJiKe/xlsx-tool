const xlsx = require("xlsx");
const path = require("path");
const fs = require("fs");
/**
 * @description 根据类型解析excel文件
 * @author peter.yuan
 * @param {*} worksheet
 * @param {*} type
 * @returns
 */
function worksheetTo(worksheet, type) {
  let res = null;
  switch (type) {
    case "csv":
      res = xlsx.stream.to_csv(worksheet);
      break;
    case "json":
      res = xlsx.utils.sheet_to_json(worksheet);
      break;
    case "html":
      res = xlsx.utils.sheet_to_html(worksheet);
      break;
    case "txt":
      res = xlsx.utils.sheet_to_txt(worksheet);
      break;

    default:
      res = xlsx.utils.sheet_to_json(worksheet);
      break;
  }
  return res;
}
/**
 * @description 获取workbook
 * @author peter.yuan
 * @param {*} fileFullName
 * @returns
 */
function excelToWorkbook(fileFullName) {
  const workbook = xlsx.readFile(fileFullName);
  return workbook;
}

function parseExcels(type) {
  const excelDir = path.resolve(__dirname, "../excels");
  const fileNames = fs.readdirSync(excelDir);
  const fileFullNames = fileNames.map(name => path.resolve(excelDir, name));
  const workbooks = fileFullNames.map(fullName => excelToWorkbook(fullName));
  let parseType = type || "json";
  let fn = type_fn_Map[parseType];
  workbooks.forEach(wb => {
    wb.SheetNames.forEach(sn => {
      const sheetRes = worksheetTo(wb.Sheets[sn], parseType);
      const dirPath = path.resolve(__dirname, `../${parseType}s`);
      createTargetDir(dirPath);
      fn({ sheetRes, dirPath, sn });
    });
  });
}

function createTargetDir(dirPath) {
  if (!fs.existsSync(dirPath)) {
    fs.mkdirSync(dirPath);
  }
}

const type_fn_Map = {
  json: parseToJson,
  csv: parseToCsv,
  html: parseToHtml,
  txt: parseToTxt
};

function parseToCsv({ sheetRes, dirPath, sn }) {
  sheetRes.pipe(fs.createWriteStream(path.resolve(dirPath, `${sn}.csv`)));
}

function parseToJson({ sheetRes, dirPath, sn }) {
  fs.writeFileSync(
    path.resolve(dirPath, `${sn}.json`),
    JSON.stringify(sheetRes)
  );
}

function parseToHtml({ sheetRes, dirPath, sn }) {
  fs.writeFileSync(path.resolve(dirPath, `${sn}.html`), sheetRes);
}
/**
 * @description 会乱码
 * @author peter.yuan
 * @param {*} { sheetRes, dirPath, sn }
 */
function parseToTxt({ sheetRes, dirPath, sn }) {
  fs.writeFileSync(path.resolve(dirPath, `${sn}.txt`), sheetRes.toString());
}


exports.parseExcels = parseExcels;
