import got from "got";
import fs from "fs";

// const adAccountIds = [
//   559998375454784,
//   483895050016166,
//   261805952717165,
//   1415231415625987,
//   5557570121017546,
//   2418878611687099,
//   806766016963986,
//   717633782800376,
//   108493532938188,
//   463105855833576,
//   668993584166544,
//   1053176932748765,
//   723426075590170,
//   3020876284885641,
//   611399683956532,
//   157407392391237,
//   875996373357310,
//   525433929217184,
//   2106427609565457,
//   1099885367534804,
//   502016591834937,
//   494060919375105,
//   497085225650510,
// ];
const adAccountIds = [
  // Thishan
  917375782495203,
  515625550049534,
  402697164195732,
  312900672574164,
  261805952717165,
  545893443486305,
  414429877110593,
  668993584166544,
  530759158500361,
  463105855833576,
  866865451186600,
  525433929217184,
  882141853058650,
  1085783205518087,
  1099885367534804,
  611399683956532,
  1415231415625987,
  1435258090286361,
  800010257981902,
  689218002445826,
  511280281119686,
  494060919375105,
  497085225650510,
];
const accessToken = process.argv[2];
const today = new Date()
  .toISOString()
  .replace(/T.*/, "")
  .split("-")
  .reverse()
  .join("-");
const fileName = `${today}-data.json`;

if (!accessToken) {
  throw new Error(`Access token not found`);
}
const allAds = [];
for (let i = 0; i < adAccountIds.length; i++) {
  const adAccountId = adAccountIds[i];
  console.log(`Processing ${adAccountId}`);
  const adAccount = await got
    .extend({
      responseType: "json",
    })
    .get(
      `https://graph.facebook.com/v15.0/act_${adAccountId}/ads?limit=100&fields=creative%7Badlabels%2Cbody%2Cid%2Cimage_url%2Cvideo_id%2Ccall_to_action_type%2Cobject_type%7D%2Cname%2Cid%2Ccreated_time%2Cinsights.breakdowns(gender%2Cage).date_preset(maximum)%7Bimpressions%2Cspend%2Cclicks%2Cconversions%2Cfrequency%2Creach%2Cvideo_30_sec_watched_actions%2Cvideo_avg_time_watched_actions%7D&access_token=${accessToken}`
    );
  const { data } = adAccount.body;
  allAds.push(...data);
}

console.log(`Contains ${allAds.length} ads`);
var json = JSON.stringify(allAds);
fs.writeFile(fileName, json, () => console.log("Done! " + fileName));
