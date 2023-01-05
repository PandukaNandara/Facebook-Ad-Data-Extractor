
import { Workbook } from 'excel4node';
import { readFileSync } from 'fs';

const today = new Date().toISOString().replace(/T.*/,'').split('-').reverse().join('-');
const fileName = `${today}.json`;

let bufferData = readFileSync(fileName);
let stData = bufferData.toString()
let jsonObj = JSON.parse(stData)

/** @type{Array} */
const dataObj = jsonObj ?? [];

var wb = new Workbook(fileName);

var ws = wb.addWorksheet('Sheet 1');

const dataCols = [
    'post_impressions_unique',
    'post_impressions',
    'post_impressions_fan',
    'post_reactions_like_total',
    'post_reactions_love_total',
    'post_reactions_wow_total',
    'post_reactions_haha_total',
    'post_reactions_sorry_total',
    'post_reactions_anger_total',
    'post_video_complete_views_30s_autoplayed',
    'post_video_complete_views_30s_clicked_to_play',
    'post_video_complete_views_30s_organic',
    'post_video_complete_views_30s_paid',
    'post_video_complete_views_30s_unique',
    'post_video_avg_time_watched',
    'post_video_avg_time_watched',
];

const singleCols = [
    'id',
    'created_time',
    "full_picture",
    'message',
]

const columns = [
    ...singleCols,
    ...dataCols,
];



for (let i = 0; i < columns.length; i++) {
    const col = columns[i];
    ws.cell(1, i + 1)
        .string(col);
}

for (let y = 0; y < dataObj.length; y++) {
    const dataRow = dataObj[y];
    const realY = y + 2;
    for (let i = 0; i < singleCols.length; i++) {
        const data = dataRow[singleCols[i]];
        ws.cell(realY, i + 1)
            .string(data);
    }


    for (let x = 0; x < dataCols.length; x++) {
        const realColX = singleCols.length + x + 1;
        /** @type{Array} */
        const data = dataRow.insights.data;
        const index = data.findIndex((e) => e.name === dataCols[x]);
        if (index >= 0) {
            ws.cell(realY, realColX)
                .string(String(data[index].values[0].value));
        }
    }
}


const name = './' + fileName.split('.')[0] + '.xlsx';
wb.write(name);