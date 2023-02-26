import excel from "exceljs";
import minifyJson from "./minifyJson.js";
import { fstat, readFileSync, unlinkSync } from "fs";
import translate from '@saipulanuar/google-translate-api';
import * as openai from "openai";


const {OpenAIApi}  = openai;

async function translateText(text, targetLanguage) {
  const translation = await translate(text, {to: 'en' });
  return translation;
}

// Set up OpenAI client
const openaiApiKey = "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";

const api = new OpenAIApi( new openai.Configuration({
  apiKey: openaiApiKey
}));
async function categorizeText(text) {
  const prompt = `Please categorize the following product description: \n\n"${text}"\n\nCategory:`;
  
  const completions = await api.createCompletion({
    model: "text-davinci-002",
    prompt,
  });
  const category = completions.data.choices[0].text.trim();
  return category;
}

const Workbook = excel.Workbook;
const today = new Date()
  .toISOString()
  .replace(/T.*/, "")
  .split("-")
  .reverse()
  .join("-");
const fileName = `${today}-data.json`;

let bufferData = readFileSync(fileName);
let stData = bufferData.toString();
let jsonObj = JSON.parse(minifyJson(stData));

/** @type{Array} */
const dataObj = jsonObj ?? [];

let hasSetupColumns = false;

const name = "./" + fileName.split(".")[0] + ".xlsx";

try {
  unlinkSync(name);
} catch (e) {}
/**
 *
 * @param {typeof Workbook} wb
 */
let count = 0;
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
  const allDataArray = [];
  for (let y = 0; y < dataObj.length; y++) {
    const dataRow = dataObj[y] ?? {};
    const { id, creative, created_time, insights } = dataRow ?? {};
    const { body, image_url, object_type, video_id = "" } = creative ?? {};
    const { data = [] } = insights ?? {};
    if (body?.trim()?.length) {
      const b = body?.trim();
      count++;
      const translatedText = await translateText(b, "en");
      const category = await categorizeText(translatedText.text);
      console.log(`Original text: ${b}`);
      console.log(`Translated text: ${translatedText}`);
      console.log(`Category: ${category}`);
      console.log("---");
    }
    if (count > 10) return;
    for (let i = 0; i < data.length; i++) {
      const el = data[i];
      const { impressions, spend, clicks, frequency, reach, gender, age } = el;
      const allData = {
        id: `${id}_${age}`,
        created_time: new Date(created_time),
        body,
        image_url: image_url
          ? {
              text: image_url,
              hyperlink: image_url,
            }
          : undefined,
        object_type,
        video_id,
        age,
        gender,
        impressions: Number(impressions),
        spend: Number(spend),
        clicks: Number(clicks),
        frequency: Number(frequency),
        reach: Number(reach),
      };
      allDataArray.push(allData);
    }
  }
  const columns = [
    { key: "id", header: "id" },
    { key: "created_time", header: "created_time" },
    { key: "body", header: "body" },
    { key: "image_url", header: "image_url" },
    { key: "object_type", header: "object_type" },
    { key: "video_id", header: "video_id" },
    { key: "age", header: "age" },
    { key: "gender", header: "gender" },
    { key: "impressions", header: "impressions" },
    { key: "spend", header: "spend" },
    { key: "clicks", header: "clicks" },
    { key: "frequency", header: "frequency" },
    { key: "reach", header: "reach" },
  ];
  ws.columns = columns;
  let i = 0;
  for (const data of allDataArray) {
    console.log("Processing " + ++i);
    const rows = ws.getRows(0, getRowCount(ws));
    const id = data.id;
    const alreadyExistedRow = rows.find(
      (row) => row.getCell("id").value === id
    );
    // if (alreadyExistedRow?.number) {
    //   const row = ws.getRow(alreadyExistedRow.number);
    //   for (const key in data) {
    //     const element = data[key];
    //     row.getCell(key).value = element;
    //   }
    //   row.commit();
    // } else {
    ws.insertRow(nextY++ + 2, data);
    // }
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
