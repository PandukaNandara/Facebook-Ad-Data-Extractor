import excel from "exceljs";
import minifyJson from "./minifyJson.js";
import { fstat, readFileSync, unlinkSync } from "fs";
import translate from "@saipulanuar/google-translate-api";
import * as openai from "openai";

const categories = {
  "23851572016560183": "Health and Wellness",
  "23851847017600183": "Health and Wellness",
  "23851646142950183": "0",
  "23851639339640183": "0",
  "23851629936170183": "Utilities and Services.",
  "23850630196420183": "0",
  "23850048946170040": "Clothing and Accessories",
  "23851318996570634": "0",
  "23850104004280634": "0",
  "23850610537850634": "Health and Wellness",
  "23850440236790634": "0",
  "23850104032450634": "Education",
  "23850078997150634": "Education",
  "23851634155150542": "Education",
  "23852889792760542": "Education",
  "23852889594490542": "Business and Industrial",
  "23852887983990542": "Utilities and Services.",
  "23851556196580542": "Home and Garden",
  "23851606010210542": "Education",
  "23851606004980542": "Education",
  "23851624946030542": "Education",
  "23851605657000542": "Business and Industrial",
  "23851536430330542": "Education",
  "23851598013600542": "Education",
  "23851609041580542": "Education",
  "23851605770300542": "Education",
  "23851597994050542": "Business and Industrial",
  "23851586318330542": "Other",
  "23851586318800542": "Other",
  "23851586346860542": "Other",
  "23851575549450542": "Business and Industrial",
  "23851575562200542": "Education",
  "23851542475680542": "Business and Industrial",
  "23851525595960542": "Business and Industrial",
  "23851525621450542": "Business and Industrial",
  "23851526997830542": "Business and Industrial",
  "23851513135770542": "Business and Industrial",
  "23851513146580542": "Education",
  "23851513153730542": "Education",
  "23851458795380542": "Business and Industrial",
  "23851468676860542": "Education",
  "23851445610490542": "Education",
  "23851403760480542": "Education",
  "23851403565630542": "Business and Industrial",
  "23851286108260529": "Business and Industrial",
  "23850577664340297": "Business and Industrial",
  "23850571771610297": "Education",
  "23850559731540297": "Education",
  "23850494830530297": "Utilities and Services.",
  "23850483171510297": "Utilities and Services.",
  "23850429749000297": "Consumer Electronics",
  "23850403450200297": "Education",
  "23850265207670297": "Business and Industrial",
  "23851880332710119": "Health and Wellness",
  "23851703050560119": "Other",
  "23851806808190119": "Other",
  "23851840306130119": "0",
  "23851807293420119": "Business and Industrial",
  "23851806792300119": "Other",
  "23851743694430119": "Business and Industrial",
  "23851743650800119": "Business and Industrial",
  "23851646757920119": "Business and Industrial",
  "23851634358590119": "Utilities and Services.",
  "23851418287060119": "Pets and Animals",
  "23851333410020119": "Pets and Animals",
  "23851333381940119": "Pets and Animals",
  "23851333331690119": "Pets and Animals",
  "23851333245900119": "Pets and Animals",
  "23851314382380119": "Pets and Animals",
  "23851234180970119": "Clothing and Accessories",
  "23851204525440119": "Clothing and Accessories",
  "23851204597100119": "Clothing and Accessories",
  "23851194981440119": "Clothing and Accessories",
  "23851186613770119": "Clothing and Accessories",
  "23850782979990300": "0",
  "23850929296720300": "Business and Industrial",
  "23850585370540300": "Business and Industrial",
  "23850561583520300": "Utilities and Services.",
  "23850550723830300": "Business and Industrial",
  "23850520771190300": "Business and Industrial",
  "23850422181450300": "Utilities and Services.",
  "23850423015940300": "Utilities and Services.",
  "23850421241530300": "Utilities and Services.",
  "23851271169040625": "Home and Garden",
  "23851301744480625": "Home and Garden",
  "23851314188290625": "Home and Garden",
  "23851253736260625": "Home and Garden",
  "23843375252130625": "Home and Garden",
  "23843375252120625": "Home and Garden",
  "23842773509640625": "Home and Garden",
  "23842755580580625": "Home and Garden",
  "23842690307960625": "Home and Garden",
  "23842673259820625": "Home and Garden",
  "23842685373890625": "Home and Garden",
  "23842730033240625": "Home and Garden",
  "23851437830520625": "Home and Garden",
  "23851271396710625": "Home and Garden",
  "23851255459850625": "Home and Garden",
  "23851253776280625": "0",
  "23851255510780625": "Home and Garden",
  "23848643780980625": "Home and Garden",
  "23848518940870625": "Home and Garden",
  "23843089873820625": "Entertainment and Media",
  "23842746013520625": "Home and Garden",
  "23842675518190625": "Home and Garden",
  "23842673354690625": "0",
  "23842746891700625": "0",
  "23842680450400625": "Home and Garden",
  "23848405679480625": "Home and Garden",
  "23842706906060625": "Home and Garden",
  "23842647709610625": "Home and Garden",
  "23842523317720625": "Home and Garden",
  "23842510632800625": "Home and Garden",
  "23852080368030403": "Education",
  "23851998651750403": "0",
  "23852143567080403": "Education",
  "23852184078560403": "Education",
  "23852218868950403": "Education",
  "23852218931470403": "Education",
  "23852143476950403": "Education",
  "23852179172720403": "0",
  "23852143220550403": "Education",
  "23852125510120403": "Education",
  "23852143257150403": "Education",
  "23852133820520403": "Education",
  "23852098353000403": "Education",
  "23851931955220403": "Utilities and Services",
  "23850486319600785": "0",
  "23850646796940785": "0",
  "23850646810880785": "0",
  "23850436594030785": "0",
  "23850640005280785": "Beauty and Personal Care",
  "23850640432680785": "Utilities and Services.",
  "23850624298250785": "Health and Wellness",
  "23850622520120785": "Utilities and Services.",
  "23850428201260785": "0",
  "23850503545400785": "Education",
  "23850490166080785": "Education",
  "23850486419430785": "0",
  "23850485018040785": "Business and Industrial",
  "23850462783650785": "Finance",
  "23850425131480785": "Finance",
  "23850425121780785": "Education",
  "23850413273140785": "Health and Wellness",
  "23850380950720785": "Health and Wellness",
  "23850384458410785": "Health and Wellness",
  "23850384285830785": "Education",
  "23850384459150785": "Health and Wellness",
  "23850370309880785": "0",
  "23850351787590785": "Education",
  "23850357671910785": "Education",
  "23852792184620245": "Retail",
  "23852722888730245": "Retail",
  "23852698311930245": "Retail",
  "23852495712840245": "Retail",
  "23852386020760245": "0",
  "23852160224580245": "Retail",
  "23851260836700245": "Retail",
  "23851788049770245": "Retail",
  "23851470318830245": "Retail",
  "23851328632270245": "Retail",
  "23851260877780245": "Retail",
  "23851178527450245": "Retail",
  "23851085312630245": "Retail",
  "23851036201340245": "Retail",
  "23851015346920245": "Retail",
  "23851004109700245": "Retail",
  "23850963500860245": "Retail",
  "23850947578910245": "Retail",
  "23850955195880245": "Retail",
  "23850634846910245": "0",
  "23852663130890768": "Entertainment and Media",
  "23852660611520768": "Entertainment and Media",
  "23852643320620768": "Education",
  "23852650888980768": "Business and Industrial",
  "23852650080700768": "Entertainment and Media",
  "23852626454390768": "0",
  "23852624192380768": "Entertainment and Media",
  "23851740442680210": "0",
  "23851680581450210": "Education",
  "23851687229390210": "Education",
  "23851699837800210": "Education",
  "23851619646720210": "Education",
  "23851618311380210": "Education",
  "23851615304000210": "Education",
  "23851610213220210": "Education",
  "23851581401640210": "Education",
  "23851533951490210": "Education",
  "23851534028490210": "Education",
  "23851542305370210": "Education",
  "23851546457860210": "Education",
  "23846802611730694": "Retail",
  "23851563620840541": "0",
  "23851532128280541": "Education",
  "23851477369440541": "Education",
  "23851404013990541": "0",
  "23851416364290541": "Education",
  "23851410034930541": "Education",
  "23852618368350550": "0",
  "23852610299170550": "Education",
  "23852599903180550": "Education",
  "23852602789600550": "Education",
  "23852598437310550": "0",
  "23852569466570550": "0",
  "23852569459170550": "0",
  "23852569446750550": "0",
  "23852564808060550": "Education",
  "23852564815180550": "Education",
  "23852564167800550": "0",
  "23852563979010550": "0",
  "23852563096850550": "0",
  "23852556419160550": "Retail",
  "23852555213550550": "Education",
  "23852555188270550": "Education",
  "23852555188140550": "Education",
  "23852548447530550": "Business and Industrial",
  "23852536094830550": "0",
  "23852526324950550": "Education",
  "23852526412620550": "Education",
  "23852469432590550": "Education",
  "23852525739110550": "Education",
  "23852339699930550": "0",
  "23852339697840550": "0",
  "23852339700000550": "0",
  "23852524962750550": "Education",
  "23852518583140550": "Education",
  "23852491843520550": "Education",
  "23852491846060550": "Education",
  "23852524903420550": "Education",
  "23852513659970550": "Education",
  "23852513663940550": "Education",
  "23852495079940550": "Education",
  "23852489348680550": "0",
  "23852489348670550": "0",
  "23852489348710550": "0",
  "23852469559230550": "Education",
  "23852439501030550": "Utilities and Services",
  "23852423981910550": "0",
  "23852423883560550": "0",
  "23852337481440550": "0",
  "23852337681490550": "Home and Garden",
  "23852337797790550": "Beauty and Personal Care",
  "23852328245930550": "Education",
  "23853049790980743": "Retail",
  "23853104743000743": "Health and Wellness",
  "23853035633300743": "Health and Wellness",
  "23853035634150743": "0",
  "23853035633570743": "Health and Wellness",
  "23853129102770743": "0",
  "23853102640540743": "Health and Wellness",
  "23853104717000743": "Health and Wellness",
  "23853097355260743": "Clothing and Accessories",
  "23853037331360743": "Entertainment and Media",
  "23853058560600743": "Education",
  "23853058695680743": "Education",
  "23853057770760743": "0",
  "23853055312930743": "0",
  "23853055286510743": "0",
  "23853055298660743": "0",
  "23853036679010743": "Entertainment and Media",
  "23853037239280743": "0",
  "23853036534080743": "Entertainment and Media",
  "23853026407270743": "Entertainment and Media",
  "23853026036220743": "Entertainment and Media",
  "23853019275280743": "Education",
  "23851068832120701": "0",
  "23850433361270701": "0",
  "23850957802560701": "Education",
  "23851180023380701": "0",
  "23851073428900701": "Education",
  "23851047697730701": "Education",
  "23851035988590701": "Education",
  "23851008751870701": "Education",
  "23850992975930701": "0",
  "23850989096450701": "Technology",
  "23850964304810701": "Education",
  "23850922185650701": "Education",
  "23850942147540701": "Education",
  "23850886379200701": "Education",
  "23850883519240701": "Education",
  "23850872117550701": "Education",
  "23850776343310701": "Finance",
  "23850745106030701": "Health and Wellness",
  "23850718279660701": "Education",
  "23850718231050701": "Education",
  "23850718235300701": "Education",
  "23850718207620701": "Education",
  "23850688462310701": "Education",
  "23850672335880701": "Education",
  "23850674935610701": "Education",
  "23850676990660701": "0",
  "23850674834710701": "Education",
  "23850667635690701": "Education",
  "23850667555710701": "Education",
  "23850667589410701": "Education",
  "23850591358370701": "Education",
  "23850589698350701": "Education",
  "23850589495980701": "Education",
  "23850577415610701": "Education",
  "23850580023480701": "Education",
  "23850580032650701": "Education",
  "23850521141600701": "0",
  "23850504789830701": "Education",
  "23850463279980701": "Education",
  "23850430801480701": "0",
  "23851847309780132": "Other",
  "23851830661350132": "Other",
  "23851796003010132": "Other",
  "23851795552580132": "Other",
  "23851777180150132": "Other",
  "23851776786220132": "Sports and Fitness",
  "23851764098440132": "Business and Industrial",
  "23853383428530399": "0",
  "23853383428500399": "0",
  "23853529977050399": "Education",
  "23853525828230399": "Education",
  "23853401412280399": "0",
  "23853457821760399": "Beauty and Personal Care",
  "23853426173970399": "Beauty and Personal Care",
  "23853485526800399": "Beauty and Personal Care",
  "23853480689850399": "Business and Industrial",
  "23853482664170399": "Beauty and Personal Care",
  "23853457812670399": "Beauty and Personal Care",
  "23853423372510399": "Beauty and Personal Care",
  "23853406800890399": "Business and Industrial",
  "23853406974020399": "Business and Industrial",
  "23853406845920399": "Education",
  "23853400919000399": "0",
  "23853401650130399": "Beauty and Personal Care",
  "23852802305160399": "Business and Industrial",
  "23853371559760399": "Education",
  "23853372293590399": "Education",
  "23853372139510399": "Education",
  "23853350359010399": "0",
  "23853311045960399": "Business and Industrial",
  "23853102987270399": "Education",
  "23853249688560399": "Travel and Hospitality",
  "23853064408650399": "0",
  "23853157421160399": "0",
  "23853157406430399": "0",
  "23853157416150399": "0",
  "23853160642640399": "0",
  "23853124858920399": "Education",
  "23853124669200399": "Beauty and Personal Care",
  "23853124858910399": "Education",
  "23853124858950399": "Education",
  "23853124809710399": "Beauty and Personal Care",
  "23853124799660399": "Beauty and Personal Care",
  "23853120485930399": "Health and Wellness",
  "23853100319390399": "Beauty and Personal Care",
  "23853095385860399": "Beauty and Personal Care",
  "23853063812890399": "0",
  "23853084930870399": "Beauty and Personal Care",
  "23853065393260399": "Home and Garden",
  "23853078112830399": "0",
  "23853063882260399": "Beauty and Personal Care",
  "23853071894970399": "Beauty and Personal Care",
  "23853064385520399": "Education",
  "23853063950870399": "Beauty and Personal Care",
  "23853058052740399": "0",
  "23852959908830399": "Entertainment and Media",
  "23852952456220399": "Business and Industrial",
  "23852952366870399": "Business and Industrial",
  "23852952365880399": "Business and Industrial",
  "23852952454900399": "Business and Industrial",
  "23852802321510399": "0",
  "23852802321060399": "0",
  "23852802319400399": "0",
  "23852939023230399": "Other",
  "123": "Business and Industrial",
  "23852905117260399": "Finance",
  "23852880756480399": "0",
  "23852848319090399": "Food and Beverage",
  "23852841715580399": "Business and Industrial",
  "23852833849460399": "Education",
  "23852833849400399": "Education",
  "23852833849450399": "Education",
  "23852844820960399": "Business and Industrial",
  "23852833808170399": "Education",
  "23852833808190399": "Education",
  "23852833808160399": "0",
  "23852833772830399": "Business and Industrial",
  "23852833767190399": "Business and Industrial",
  "23852827422460399": "Other",
  "23852803395510399": "Education",
  "23851760779350102": "Home and Garden",
  "23851880247460102": "Education",
  "23851828179510102": "Education",
  "23851770786420102": "Health and Wellness",
  "23851853469160102": "Education",
  "23851914637680102": "0",
  "23851914490410102": "0",
  "23851914490390102": "0",
  "23851905947800102": "Business and Industrial",
  "23851901681830102": "0",
  "23851901587410102": "0",
  "23851901743480102": "0",
  "23851895533930102": "Business and Industrial",
  "23851812589930102": "Health and Wellness",
  "23851807252260102": "Health and Wellness",
  "23851812680490102": "Health and Wellness",
  "23851867583620102": "Business and Industrial",
  "23851867532800102": "Business and Industrial",
  "23851797668470102": "Business and Industrial",
  "23851853428020102": "0",
  "23851853443770102": "0",
  "23851853428290102": "0",
  "23851750189930102": "0",
  "23851828688500102": "Business and Industrial",
  "23851812159400102": "Health and Wellness",
  "23851813050980102": "Business and Industrial",
  "23851806483510102": "Business and Industrial",
  "23851798371220102": "Business and Industrial",
  "23851784384740102": "Health and Wellness",
  "23851776338430102": "Health and Wellness",
  "23851783046840102": "Education",
  "23851771571150102": "0",
  "23851771482840102": "0",
  "23851763668210102": "Other",
  "23851758873480102": "Other",
  "23851756642510102": "Other",
  "23851750030050102": "0",
  "23851719806250102": "Education",
  "23851735855330102": "Education",
};

const { OpenAIApi } = openai;

async function translateText(text, targetLanguage) {
  const translation = await translate(text, { to: "en" });
  return translation;
}

// Set up OpenAI client
const openaiApiKey = "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";

const api = new OpenAIApi(
  new openai.Configuration({
    apiKey: openaiApiKey,
  })
);
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
    // if (body?.trim()?.length) {
    //   const b = body?.trim();
    //   count++;
    //   const translatedText = await translateText(b, "en");
    //   const category = await categorizeText(translatedText.text);
    //   console.log(`Original text: ${b}`);
    //   console.log(`Translated text: ${translatedText}`);
    //   console.log(`Category: ${category}`);
    //   console.log("---");
    // }
    // if (count > 10) return;
    for (let i = 0; i < data.length; i++) {
      const el = data[i];
      const { impressions, spend, clicks, frequency, reach, gender, age } = el;
      const allData = {
        id: `${id}_${age}`,
        category: categories[id],
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
    { key: "category", header: "category" },
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
