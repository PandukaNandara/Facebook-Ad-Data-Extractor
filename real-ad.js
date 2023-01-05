import excel from "exceljs";
import minifyJson from "./minifyJson.js";
import { readFileSync } from "fs";
const Workbook = excel.Workbook;
const today = new Date().toISOString().replace(/T.*/,'').split('-').reverse().join('-');
const fileName = `${today}.json`;

let bufferData = readFileSync(fileName);
let stData = bufferData.toString();
let jsonObj = JSON.parse(minifyJson(stData));

/** @type{Array} */
const dataObj = jsonObj ?? [];

let hasSetupColumns = false;

const name = "./" + fileName.split(".")[0] + ".xlsx";

/**
 *
 * @param {xl.Workbook} wb
 */
function getRowCount(ws) {
  const rows = ws.getColumn(1);
  const rowsCount = rows["_worksheet"]["_rows"].length;
  return rowsCount;
}
async function useWorkBook(wb) {
  let ws = wb.getWorksheet(1);
  if (!ws) {
    ws = wb.addWorksheet("Sheet1");
  }
  let nextY = getRowCount(ws);
  const columns = new Set();
  const allDataArray = [];
  for (let y = 0; y < dataObj.length; y++) {
    console.log(y + 1);
    const dataRow = dataObj[y] ?? {};
    const { id, creative, created_time, insights } = dataRow ?? {};
    const { body, image_url, object_type, video_id = '' } = creative ?? {};
    const { data = [] } = insights ?? {};
    const insightData = {};
    for (let i = 0; i < data.length; i++) {
      const el = data[i];
      for (const key in el) {
        if (
          key === "age" ||
          key === "gender" ||
          key === "date_start" ||
          key === "date_stop"
        )
          continue;
        if (Object.hasOwnProperty.call(el, key)) {
          const element = el[key];
          insightData[`${el.age}_${el.gender}_${key}`] = element;
        }
      }
    }
    const allData = {
      id,
      created_time,
      body,
      image_url,
      object_type,
      video_id,
      ...insightData,
    };
    Object.keys(allData).forEach((key) => columns.add(key));
    allDataArray.push(allData);
  }
  const _col = [];
  columns.forEach((key) =>
    _col.push({
      key,
      header: key,
    })
  );
  ws.columns = _col;
  for (const allData of allDataArray) {
    const rows = ws.getRows(0, getRowCount(ws));
    const id = allData.id;
    const alreadyExistedRow = rows.find(
      (row) => row.getCell("id").value === id
    );
    if (alreadyExistedRow?.number) {
      const row = ws.getRow(alreadyExistedRow.number);
      for (const key in allData) {
        const element = allData[key];
        console.log(key, alreadyExistedRow.number, element);
        row.getCell(key).value = element;
      }
      row.commit();
    } else {
      ws.insertRow(nextY++ + 2, allData);
    }
  }
  // for (const key in allData) {
  //   const data = allData[key];

  //   ws.cell(realY, realX).string(String(data));
  //   realX++;
  // }
  //   wb.write(name);

  await wb.xlsx.writeFile(name);
}

let wb = new Workbook();
wb.xlsx
  .readFile(name)
  .then(async (w) => {
    console.log("Pass 1");
    await useWorkBook(w);
  })
  .catch(() => {
    useWorkBook(wb);
  });
