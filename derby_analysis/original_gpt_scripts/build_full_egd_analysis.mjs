import fs from "node:fs/promises";
import path from "node:path";
import crypto from "node:crypto";
import { SpreadsheetFile, Workbook } from "@oai/artifact-tool";

const root = String.raw`C:\Codex\Renshe_Derby`;
const transcriptDir = path.join(root, "transcript");
const outputDir = path.join(root, "outputs", "01a1117c-031d-7c03-8f22-fad260b4270e");
const outputPath = path.join(outputDir, "Derby_Full_EGD_PO_Brand_Preference_Content_Analysis.xlsx");
const previewDir = path.join(root, ".qa_previews_full_egd");

const coverage = JSON.parse(await fs.readFile(path.join(root, "egd_section_coverage.json"), "utf8"));
const prior = JSON.parse(await fs.readFile(path.join(root, "prior_analysis_extract.json"), "utf8"));
const transcriptNames = (await fs.readdir(transcriptDir)).filter(n => /\.(xlsx|docx)$/i.test(n)).sort();

const findName = (test) => {
  const matches = transcriptNames.filter(test);
  if (matches.length !== 1) throw new Error(`Expected exactly one matching file; found ${matches.length}: ${matches.join(", ")}`);
  return matches[0];
};

const currentNames = {
  DM3: findName(n => /^Checked_Dm3_EGD_/i.test(n)),
  DM6: findName(n => /^Checked_DM6_EGDs_/i.test(n)),
  RM5: findName(n => /^Checked_RM5_EGD_/i.test(n)),
  RM2: findName(n => /^RM2_EGD_/i.test(n)),
  POD3: findName(n => /^Checked_PO-D3_/i.test(n)),
  R5PO: findName(n => /^Checked_R5_PO_/i.test(n)),
  R2PO: findName(n => /^R2_PO_/i.test(n)),
  POD06: findName(n => /^PO_D06_/i.test(n)),
};

const groupProfiles = [
  ["DM3", "Dhaka Metro", "Explorer", "19–24", "SEC D–E", 6,
    "Students and young workers. Daily life centers on study/work, friends, mobile use and careful allocation of limited personal money.",
    "Success means earning capacity, parental security and continued progress. Family expectations create pressure.",
    "Travel, Bangla/Hindi entertainment, short video, fashion content and peer observation shape interests and choices.",
    "Discounts and price-quality balance matter. Online trust is mixed; purchases are adjusted within allowance or daily cash.",
    "Smoking is woven into breaks, friendship and daily completion. Derby is the affordable mid-strength regular choice.",
    "Odor, fast burn and hot/soft filter weaken an otherwise strong value proposition."],
  ["DM6", "Dhaka Metro", "Explorer", "25–29", "SEC D–E", 7,
    "Early-career workers, small-business participants and job aspirants. Work pressure and long hours shape routines.",
    "More income sources, respectable employment and business progress are central goals. Status is linked with recognition and achievement.",
    "Mobile use, theatre, entertainment and accessible mid-priced brands/products matter more than exclusive luxury alone.",
    "Budgeting and saving are explicit. Availability during travel or in rural areas is important.",
    "Cigarettes mark breaks, relief and social connection. Derby wins through satisfaction, price, peer use and availability.",
    "Filter-end softness/heat and odor are clear gaps; equal pricing opens premium trade-up."],
  ["RM5", "Rajshahi Metro", "Explorer", "19–24", "SEC D–E", 8,
    "Mostly students with study, tea-stall, campus and friend-circle routines. Financial dependence on family is still visible.",
    "Employment, income growth, side income, family repayment and social advancement define success; job scarcity creates anxiety.",
    "Life-related films, games, music and social content support identification and escape. Shopping is mainly need-led.",
    "Smoking spend is relatively protected even when other social expenditure is reduced.",
    "Tea, adda and cigarettes form repeated daily rituals. Derby is the strongest satisfying option inside the student budget.",
    "Persistent odor is the dominant social barrier; pack cannot compensate for product experience."],
  ["RM2", "Rajshahi Metro", "Image Seeker", "25–29", "SEC D–E", 7,
    "Young working men with strong friendship networks and routines built around work, tea, travel and night-time socializing.",
    "Brands are understood as markers of progress, respect and status. Economic realism remains central to everyday choice.",
    "Brand symbolism, peer observation and interpersonal recommendation matter more than passive poster claims.",
    "Pack versus single-stick purchase follows location, convenience and expected consumption; Derby's ubiquity reduces search effort.",
    "Cigarettes provide completion, focus, relief, company and social bonding. Derby is habitual, affordable, strong and smoky.",
    "Odor, soft filter and low-class/rural imagery create tension with status aspiration."],
];

const themes = [
  ["F01", "Family duty and upward mobility", [1,1,1,1], "Success is framed through earning, parental security, respectable work and continued progress.", "Position improvement as practical progress and dependability, not distant luxury."],
  ["F02", "Financial pressure shapes daily trade-offs", [1,1,1,1], "Allowances, savings, household contribution and unstable income make value calculations routine.", "Protect affordability and communicate product value in concrete everyday terms."],
  ["F03", "Friendship and adda organize life", [1,1,1,1], "Friends influence leisure, discovery, smoking initiation, brand trial and consumption frequency.", "Use social proof and small-group trial mechanics."],
  ["F04", "Mobile and entertainment are discovery spaces", [1,1,1,1], "Short video, drama, films, games, travel and fashion content shape attention and aspiration.", "Use culturally familiar content environments, within applicable marketing restrictions."],
  ["F05", "Shopping is pragmatic and value-led", [1,1,1,1], "Need, discount, accessibility and trust often matter more than formal brand loyalty in general shopping.", "Make quality proof visible and easy to verify at point of purchase."],
  ["F06", "Brands signal progress and social standing", [1,1,1,1], "Premium products can change how participants expect others to see and treat them.", "Improve Derby's social acceptability without losing its everyday accessibility."],
  ["F07", "Smoking is embedded in routine", [1,1,1,1], "Morning, meals, tea, work breaks, travel, waiting and night-time moments trigger use.", "Understand Derby as a ritual product, not an occasional indulgence."],
  ["F08", "Smoking provides relief, focus and completion", [1,1,1,1], "Participants describe calm, recharge, decision support, concentration and a sense of completion.", "Any product change must preserve the felt satisfaction attached to these moments."],
  ["F09", "Harm and social cost are recognized", [1,1,1,1], "Health concerns, odor and discomfort to non-smokers are understood, but habit often overrides them.", "Reduce avoidable social and sensory costs; do not claim health benefits."],
  ["F10", "Strong regular identity rejects lightness", [1,1,1,1], "Red/regular/strong cigarettes are associated with satisfaction; white/light options often feel insufficient.", "Keep Derby's strong regular character as a product guardrail."],
  ["F11", "Derby owns affordable satisfaction", [1,1,1,1], "Price, strength, smoke and frequent-use economics combine into one value equation.", "Defend value-per-stick rather than competing on low price alone."],
  ["F12", "Peers, availability and habit sustain Derby", [1,1,1,1], "Friends and local availability create entry; repeated physical fit creates inertia.", "Maintain distribution and use credible peer/retailer product experience."],
  ["F13", "Odor is the largest product vulnerability", [1,1,1,1], "Lingering smell limits comfort, acceptance by others and use in visible settings.", "Prioritize odor reduction while preserving strength, smoke and taste."],
  ["F14", "Filter and burn execution signal quality", [1,1,1,1], "Hot/soft filters, peeling cover, quick burn, paper and ash stability recur as quality cues.", "Improve end-of-stick filter integrity and burn consistency."],
  ["F15", "Price parity encourages trade-up", [1,1,1,1], "When the economic gap narrows, Star and premium brands gain on quality and status.", "Track relative price gaps and add a clear quality reason to remain with Derby."],
  ["F16", "Everyday loyalty has a status gap", [1,0,1,1], "Derby suits self/everyday use, while premium brands can be preferred for campus, guests or senior company.", "Address odor and visible quality cues to improve occasion stretch."],
  ["F17", "Product proof converts better than pack alone", [1,1,1,1], "Pack can attract attention, but peers, retailers and actual product performance determine repeat use.", "Lead sampling with familiar satisfaction, better filter and less odor."],
  ["F18", "Derby communication is under-recognized", [1,0,1,1], "Competitor or retailer activity is recalled more readily than Derby communication in three groups.", "Build retailer visibility around demonstrable product improvements."],
];

const contextEvidence = [
  ["DM3","Daily life","A","row 203","কলেজে যাই… ক্লাস… বন্ধুবান্ধবের সাথে আড্ডা… লাইব্রেরিতে গিয়ে কিছু সময়।","College, classes, friends and library time organize the day.","Study and peer interaction are the main everyday context."],
  ["DM3","Dreams and duty","B/C","rows 490–491","মাকে… চাওয়া মাত্রই দিতে পারবো… বাবা-মা বিনা টেনশনে থাকবে।","Success means being able to provide immediately and remove parents' financial worry.","Upward mobility is measured through family security."],
  ["DM3","Media and inspiration","C","row 786","ফ্যাশনের কন্টেন্ট… যেটা ভালো লাগে… বাজারে গেলে নিজের মতো করে নিই।","Fashion content supplies pieces of inspiration that are recombined personally.","Digital content influences style without creating simple imitation."],
  ["DM3","Shopping/value","F/B","rows 676, 691","ডিসকাউন্ট… ১০০০ টাকার জিনিস ৪০০–৫০০… অনলাইনের উপর বিশ্বাস উঠে গেছে।","Discounts attract, while a poor online purchase damaged trust.","Value and trust jointly shape purchase."],
  ["DM3","Personal budget","E","rows 2409–2427","মাসে ৩০০০–৪০০০… একদিকে কম খরচ করলে অন্য প্রয়োজন মিলে যায়।","A limited monthly allowance is reallocated between needs.","Daily consumption competes inside a fixed personal budget."],
  ["DM6","Work context","F","row 168","লাইব্রেরিতে লাইব্রেরিয়ান… স্টুডেন্টরা জব প্রিপারেশন নেয়।","Works as a librarian where graduates prepare for jobs.","Employment aspiration is close to daily experience."],
  ["DM6","Dreams","A/F","rows 424, 450","আর্নিং সোর্স বাড়াবো… বাংলাদেশ ব্যাংকের এডি… সমাজের চোখে সম্মান।","Increase income sources and reach a respected professional role.","Income and recognition are linked aspirations."],
  ["DM6","Brand meaning","G/B","rows 1360–1369","ব্র্যান্ড ভ্যালু… মানুষ ভাবে ছেলেটার জিনিস… ভালো লাগে।","Recognized brands can make others view the user differently.","Status is a real component of brand choice."],
  ["DM6","Accessible value","F","row 1179","প্রাইস রেঞ্জ হাই না… মিডিয়াম… সবার কাছে বেয়ার করা যায়।","A medium price that many can bear is viewed positively.","Accessibility is valued beyond the cigarette category."],
  ["DM6","Smoking moment","E","row 2281","রাতে… রিল্যাক্স… দিনে… কাজের প্রেসার কমানোর কারণে।","Night smoking is for relaxation; daytime smoking reduces work pressure.","Consumption serves different functional-emotional roles by moment."],
  ["DM6","Consumption planning","F","row 2152","সিগারেট খাওয়ার জন্য… নির্দিষ্ট সময়… বের করতে হবে।","Time is deliberately made for a cigarette break.","Smoking is scheduled into the workday."],
  ["RM5","Student routine","E","rows 233–247","ক্লাসের গ্যাপে চা আর বিড়ি… বন্ধুদের সাথে আড্ডা… কার্ড খেলি।","Tea, cigarettes, friends and cards repeat through the student day.","Smoking is integrated into campus and social routine."],
  ["RM5","Family and work","A","row 540","বাবা-মা কষ্ট করে চালায়… রেজাল্ট ভালো… চাকরি হয়।","Parental sacrifice motivates study and job achievement.","Family duty anchors aspiration."],
  ["RM5","Success/status","G","row 626","সোশ্যাল স্ট্যাটাস আপ… প্রয়োজনীয় পণ্য কিনতে টাকার কথা ভাবতে হবে না।","Progress means higher social status and buying necessities without price worry.","Economic freedom and social advancement are paired."],
  ["RM5","Challenges","B/A","rows 643, 651","চাকরির প্রতিদ্বন্দ্বিতা বাড়ছে… বাবা-মা বলে কিছু একটা কর।","Job competition and family pressure create anxiety.","Financial uncertainty intensifies value sensitivity."],
  ["RM5","Shopping","B/C","rows 720, 737","নেসেসারি হলে কিনি… ব্র্যান্ডে কখনো যাওয়া হয় নাই।","Purchases are need-led; branded stores may not be part of experience.","Practical access matters more than formal retail prestige."],
  ["RM5","Spending priority","D","rows 2004–2007","ক্যাম্পাসের এক্সট্রা খরচ কমে… সিগারেট কমে না।","Other campus spending falls, but cigarette consumption does not.","Cigarette spend is protected within a constrained budget."],
  ["RM5","Social cost","A","row 1225","মুখ থেকে গন্ধ… ভালো লাগতেছে না… দূরে গিয়ে স্মোক।","Non-smokers object to odor, forcing distance or masking.","Odor reduces social acceptability."],
  ["RM5","Strength need","F","row 1416","লাইট সিগারেট… কাজ করলো না… কড়া সিগারেট হলে মনে হয় নিকোটিন আছে।","Light cigarettes feel ineffective; strong cigarettes feel functional.","Strong delivery is a category-level need."],
  ["RM2","Brand and status","Group","paragraphs 923–970","ব্র্যান্ড… আমাকে অন্যভাবে দেখবে… নিজের স্ট্যাটাস বাড়ে।","Brands are expected to change how others see the user and raise status.","Image-seeking aspiration sits behind product symbolism."],
  ["RM2","Routine","Group","paragraphs 1305–1311","ঘুমাতে যাওয়ার আগে এবং ঘুম থেকে উঠার পর… না খেলে মিসিং লাগে।","The cigarette before sleep and after waking feels essential.","Smoking brackets the day as a ritual."],
  ["RM2","Work and stress","Group","paragraphs 1353–1388","কাজে মনোযোগ… চার্জ চলে আসে… ডিসিশন নিতে পারি।","Smoking is described as improving focus, restoring energy and aiding decisions.","Participants attach performance and coping benefits to the ritual."],
  ["RM2","Social role","Group","paragraphs 1395–1397","স্মোকিং ছাড়া আড্ডা জমে না… আড্ডার মেইন অংশ।","Adda does not feel complete without smoking.","Cigarettes function as social infrastructure."],
  ["RM2","Emotional role","Group","paragraph 1410","মানসিক শান্তি দেয়… না থাকলে বিরক্ত এবং একা লাগে।","It provides mental calm; absence creates irritation and loneliness.","Habit combines emotional comfort and companionship."],
  ["RM2","Brand image","Group","paragraphs 2083–2090","নিম্নবিত্ত… মিডেল ক্লাসে বেশি খায়… জনপ্রিয়তা বেশি।","Derby carries lower-/middle-class imagery but very high popularity.","Reach and familiarity coexist with a status ceiling."],
  ["RM2","Marketing credibility","Group","paragraphs 2469–2472","প্রোডাক্ট ভালো হলে বাজারে চলবে… ডার্বি ছাড়া তৃপ্তি আসে না।","Participants say product quality, not activity alone, explains market strength.","Communication must be backed by recognizable product delivery."],
  ["RM2","Trust and influence","Group","paragraphs 2595–2630","অথেন্টিক বোঝা যায়… ইনফ্লুয়েন্সার দরকার… শাকিব খান/কাবিলা।","Certification signals authenticity; familiar entertainers are suggested as endorsers.","Trust and cultural familiarity can support communication."],
];

const sourceNameByGroup = {DM3: currentNames.DM3, DM6: currentNames.DM6, RM5: currentNames.RM5, RM2: currentNames.RM2};
const oldEvidence = prior.evidence.map(r => [r[0], r[1], r[2], r[3], r[4], r[5], r[6]]);
const evidence = [...contextEvidence, ...oldEvidence].map(r => [...r, sourceNameByGroup[r[0]]]);

const poRows = prior.po.map(r => {
  const id = String(r[0]);
  const source = id === "PO-D3" ? currentNames.POD3 : id === "R5" ? currentNames.R5PO : id === "R2-PO" ? currentNames.R2PO : currentNames.POD06;
  return [...r, "Cigarette Brand Preference module only", source];
});

function family(section) {
  const s = section.toLowerCase();
  if (/intro|daily|dream|perfect|lifestyle|phych|challenge|entertain|media|shopping|spending|collecting|celebrity/.test(s)) return "Life, aspiration and media";
  if (/category|smoker|consum|word game|purchase$/.test(s)) return "Category and consumption";
  if (/brand|alternative|reject|selection|unique|persona|pack|stick|strength|weak|price|quality|appeal|stimulus|name/.test(s)) return "Brand and product";
  return "Awareness, activation and future";
}

const sectionCoverage = coverage.map(r => [r.group, family(r.section), r.section, r.location, r.records ?? "n.a.", r.included, r.source]);

const wb = Workbook.create();
const colors = {navy:"#17324D", blue:"#2F75B5", red:"#C6342E", gold:"#D7A23A", teal:"#2A7F8E", green:"#548235", paleBlue:"#EAF1F8", paleGold:"#FFF2CC", paleRed:"#FCE4D6", paleGreen:"#E2F0D9", grey:"#F2F2F2", mid:"#D9E1F2", white:"#FFFFFF", text:"#222222"};
const font = "Arial";

function colName(n) {
  let s = "";
  while (n > 0) { n--; s = String.fromCharCode(65 + n % 26) + s; n = Math.floor(n / 26); }
  return s;
}

function setup(sheet, title, subtitle, cols) {
  sheet.showGridLines = false;
  sheet.getRange(`A2:${colName(cols)}2`).format.borders = {bottom:{style:"medium",color:colors.navy}};
  sheet.getRange("A2").values = [[title]];
  sheet.getRange("A2").format.font = {name:font,size:16,bold:true,color:colors.navy};
  sheet.getRange("A3").values = [[subtitle]];
  sheet.getRange(`A3:${colName(cols)}3`).format.font = {name:font,size:9,italic:true,color:"#566573"};
  sheet.getRange(`A3:${colName(cols)}3`).format.rowHeight = 30;
}

function table(sheet, startRow, headers, rows, widths, bodyHeight=46, headerFill=colors.navy) {
  const endCol = colName(headers.length);
  const endRow = startRow + rows.length;
  sheet.getRange(`A${startRow}:${endCol}${endRow}`).values = [headers, ...rows];
  const hr = sheet.getRange(`A${startRow}:${endCol}${startRow}`);
  hr.format = {fill:headerFill,font:{name:font,size:10,bold:true,color:colors.white},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#FFFFFF"}};
  hr.format.rowHeight = 34;
  const br = sheet.getRange(`A${startRow}:${endCol}${endRow}`);
  br.format.font = {name:font,size:9,color:colors.text};
  br.format.verticalAlignment = "top";
  br.format.wrapText = true;
  br.format.borders = {preset:"all",style:"thin",color:"#D9D9D9"};
  if (rows.length) sheet.getRange(`A${startRow+1}:${endCol}${endRow}`).format.rowHeight = bodyHeight;
  widths.forEach((w,i) => sheet.getRange(`${colName(i+1)}1:${colName(i+1)}${Math.max(endRow,30)}`).format.columnWidth = w);
  if (rows.length) sheet.freezePanes.freezeRows(startRow);
  return endRow;
}

// Summary
{
  const s = wb.worksheets.add("Summary");
  setup(s, "Derby full EGD content analysis", "Full EGD discussions are primary evidence. PO evidence is restricted to the Cigarette Brand Preference module.", 8);
  s.getRange("A5:H5").values = [["Evidence base","Value proposition","Main vulnerability","Conversion condition","Status tension","Full EGD coverage","PO boundary","Method"]];
  s.getRange("A5:H5").format = {fill:colors.navy,font:{name:font,size:10,bold:true,color:colors.white},horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#FFFFFF"}};
  s.getRange("A6:H6").values = [["4 complete EGDs\n~28 participants","Affordable, strong and smoky everyday satisfaction","Persistent odor, then filter and burn execution","Same price, Derby-like feel, better filter and less odor","Everyday fit is stronger than visible-status fit","88 source sections or paragraph bands included","4 PO interviews; brand preference section only","Qualitative group-level thematic analysis"]];
  s.getRange("A6:H6").format = {font:{name:font,size:10,bold:true,color:colors.text},fill:colors.paleBlue,horizontalAlignment:"center",verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#D9D9D9"}};
  s.getRange("A6:H6").format.rowHeight = 78;
  const main = [
    ["1","Life context","Participants are managing study, work, family expectation and limited cash while trying to progress socially and economically."],
    ["2","Role of smoking","Smoking is attached to morning/night routines, tea, meals, work breaks, stress relief, focus, waiting and adda."],
    ["3","Why Derby wins","Derby combines frequent-use affordability with strong regular delivery, visible smoke, availability and learned physical fit."],
    ["4","Why Derby is vulnerable","Odor reduces social acceptability. Filter softness/heat, paper, burn and ash issues signal lower execution quality."],
    ["5","Competitive risk","At price parity, Star and premium brands gain. Some smokers already trade up for guests, campus or visible occasions."],
    ["6","How trial converts","Peers and retailers are more credible than poster copy. Product must retain Derby-like satisfaction while improving filter and odor."],
  ];
  s.getRange("A9:H9").values = [["No.","Area","Finding","","","","",""]];
  s.getRange("A9:H9").format = {fill:colors.red,font:{name:font,size:10,bold:true,color:colors.white},borders:{preset:"all",style:"thin",color:"#FFFFFF"}};
  main.forEach((r,i) => { const rr=10+i; s.getRange(`A${rr}`).values=[[r[0]]]; s.getRange(`B${rr}`).values=[[r[1]]]; s.getRange(`C${rr}:H${rr}`).values=[[r[2],null,null,null,null,null]]; });
  s.getRange("A10:H15").format = {font:{name:font,size:10,color:colors.text},verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#D9D9D9"}};
  s.getRange("A10:A15").format = {fill:colors.paleRed,font:{name:font,size:11,bold:true,color:colors.red},horizontalAlignment:"center",verticalAlignment:"center"};
  s.getRange("A10:H15").format.rowHeight = 46;
  s.getRange("A18:H18").values = [["Priority","Recommended action","Reason","","","","",""]];
  s.getRange("A18:H18").format = {fill:colors.teal,font:{name:font,size:10,bold:true,color:colors.white},borders:{preset:"all",style:"thin",color:"#FFFFFF"}};
  const actions = [
    ["P1","Reduce lingering odor without reducing strength, smoke volume or taste.","Odor is the most consistent cross-group product and social barrier."],
    ["P2","Improve filter-end integrity and burn/paper/ash consistency.","These are repeated, tangible quality signals."],
    ["P3","Defend relative value against Star and premium brands before price convergence.","Equal pricing weakens practical loyalty."],
    ["P4","Use peer/retailer-led product trial around a clear proof.","Recommendation and experience are more credible than passive claims."],
  ];
  actions.forEach((r,i) => { const rr=19+i; s.getRange(`A${rr}`).values=[[r[0]]]; s.getRange(`B${rr}:E${rr}`).values=[[r[1],null,null,null]]; s.getRange(`F${rr}:H${rr}`).values=[[r[2],null,null]]; });
  s.getRange("A19:H22").format = {font:{name:font,size:10,color:colors.text},verticalAlignment:"center",wrapText:true,borders:{preset:"all",style:"thin",color:"#D9D9D9"}};
  s.getRange("A19:A22").format = {fill:colors.paleGreen,font:{name:font,size:10,bold:true,color:colors.green},horizontalAlignment:"center",verticalAlignment:"center"};
  s.getRange("A19:H22").format.rowHeight = 52;
  [10,18,18,18,18,18,18,18].forEach((w,i)=>s.getRange(`${colName(i+1)}1:${colName(i+1)}30`).format.columnWidth=w);
}

// Themes
{
  const s = wb.worksheets.add("Full EGD Themes");
  setup(s, "Themes across the complete EGDs", "Context, aspiration, media, shopping, category, consumption, brand, product and activation sections are included.", 10);
  const rows = themes.map((t,i) => [t[0],t[1],t[3],t[2][0],t[2][1],t[2][2],t[2][3],null,null,t[4]]);
  const end = table(s,5,["Code","Theme","Evidence summary","DM3","DM6","RM5","RM2","Groups","% groups","Implication"],rows,[9,34,58,8,8,8,8,10,11,58],60);
  for (let r=6;r<=end;r++) {
    s.getRange(`H${r}`).formulas = [[`=SUM(D${r}:G${r})`]];
    s.getRange(`I${r}`).formulas = [[`=H${r}/4`]];
  }
  s.getRange(`D6:I${end}`).format.horizontalAlignment = "center";
  s.getRange(`I6:I${end}`).format.numberFormat = "0%";
  s.getRange(`D6:G${end}`).conditionalFormats.add("cellIs",{operator:"equal",formula:1,format:{fill:colors.paleGreen,font:{bold:true,color:colors.green}}});
  s.getRange(`D6:G${end}`).conditionalFormats.add("cellIs",{operator:"equal",formula:0,format:{fill:colors.grey,font:{color:"#777777"}}});
}

// Group profiles
{
  const s = wb.worksheets.add("EGD Group Profiles");
  setup(s, "EGD group profiles", "Each row summarizes the complete discussion before interpreting Derby preference.", 12);
  table(s,5,["Group","City","Segment","Age","SEC","Participants","Life context","Dreams and pressures","Media and culture","Shopping and money","Smoking and Derby role","Main tension"],groupProfiles,[9,15,14,10,10,11,36,38,38,38,40,40],112);
}

// Section coverage
{
  const s = wb.worksheets.add("EGD Section Coverage");
  setup(s, "Full EGD coverage register", "All tagged sections from the three checked EGD spreadsheets and all paragraph bands from the RM2 EGD DOCX are included.", 7);
  table(s,5,["Group","Section family","Source section","Source location","Transcript records","Included","Current source file"],sectionCoverage,[9,28,34,20,16,11,70],34);
}

// Evidence log
{
  const s = wb.worksheets.add("EGD Evidence");
  setup(s, "EGD evidence log", "Participant evidence across the full EGD scope. Moderator questions are not treated as findings.", 8);
  table(s,5,["Group","Theme","Speaker","Source row or paragraph","Participant verbatim (Bangla)","Concise English translation","Analytic meaning","Current source file"],evidence,[9,19,12,20,54,50,50,70],72);
}

// PO comparison
{
  const s = wb.worksheets.add("PO Brand Preference");
  setup(s, "PO comparison: Cigarette Brand Preference only", "Lifestyle, psychography, media and unrelated PO sections are excluded from this sheet and from PO conclusions.", 8);
  table(s,5,["PO ID","Profile","Brand journey and entry","Derby equity","Price, fallback and trial","How it supports the EGD conclusion","PO scope used","Current source file"],poRows,[12,29,45,43,45,45,31,70],96,colors.teal);
}

// Implications
{
  const s = wb.worksheets.add("Implications");
  setup(s, "Brand implications", "Actions integrate full-EGD context with the restricted PO brand-preference comparison.", 6);
  const rows = [
    [1,"What must Derby retain?","Affordable strong regular satisfaction, smoke volume, familiarity and easy availability.","All scoped PO cases reinforce value plus satisfactory delivery.","Set non-negotiable sensory guardrails before reformulation.","A lighter or less recognizable Derby could lose its core users."],
    [2,"What should be fixed first?","Odor, then filter-end integrity and burn/paper/ash consistency.","POs also use filter, paper and burn as quality cues.","Prototype and blind-test low-odor, stronger-filter executions.","Social rejection and visible quality defects sustain trade-up."],
    [3,"How strong is loyalty?","Habit is strong, but price parity and visible occasions can unlock trade-up.","PO price tolerance varies, confirming conditional loyalty.","Model relative price gaps and test parity with Star/premium choices.","A small architecture change may cause disproportionate switching."],
    [4,"How should trial work?","Peer/retailer proof, Derby-like feel, better filter and less odor.","POs also prefer recommendation and one- or two-stick trial.","Use guided sampling and repeat exposure, not awareness alone.","Attractive pack without product superiority will not retain."],
    [5,"How can status improve?","Reduce odor and strengthen quality cues while remaining accessible.","POs show pack cues help only when product fit is maintained.","Lift social acceptability without over-premiumizing price or tone.","Luxury cues alone could alienate the everyday core."],
  ];
  table(s,5,["Priority","Business question","Full EGD answer","PO triangulation","Recommended action","Risk if ignored"],rows,[10,31,51,48,51,48],86,colors.red);
}

// Source audit and method
{
  const s = wb.worksheets.add("Source Audit");
  setup(s, "Source audit and scope", "Current filenames and file hashes are recorded so later renaming or replacement can be detected.", 8);
  const sourceRows = [];
  for (const name of transcriptNames) {
    const buf = await fs.readFile(path.join(transcriptDir,name));
    const stat = await fs.stat(path.join(transcriptDir,name));
    const isEgd = /EGD/i.test(name);
    const sha = crypto.createHash("sha256").update(buf).digest("hex").toUpperCase();
    sourceRows.push([isEgd?"Primary":"Supporting",isEgd?"EGD":"PO",name,stat.size,sha,isEgd?"Complete transcript":"Cigarette Brand Preference section only",isEgd?"Full EGD synthesis and evidence":"PO triangulation only","Included"]);
  }
  const end = table(s,5,["Evidence tier","Type","Current filename","Bytes","SHA-256","Scope used","Analytical role","Status"],sourceRows,[16,12,76,14,68,34,32,13],52);
  s.getRange(`D6:D${end}`).format.numberFormat = "#,##0";
  const methodRow = end + 3;
  s.getRange(`A${methodRow}:H${methodRow}`).values = [["Method note","The analysis uses deductive categories from the discussion guides and inductive themes repeated across transcripts.",null,null,null,null,null,null]];
  s.getRange(`A${methodRow}:H${methodRow}`).format = {fill:colors.paleGold,font:{name:font,size:10,italic:true,color:colors.text},wrapText:true,borders:{preset:"all",style:"thin",color:"#D9D9D9"}};
  s.getRange(`A${methodRow}:H${methodRow}`).format.rowHeight = 44;
  s.getRange(`A${methodRow+1}:H${methodRow+1}`).values = [["Counting rule","Theme prevalence is counted once per EGD group. It is not a participant percentage or population estimate.",null,null,null,null,null,null]];
  s.getRange(`A${methodRow+1}:H${methodRow+1}`).format = {fill:colors.paleBlue,font:{name:font,size:10,italic:true,color:colors.text},wrapText:true,borders:{preset:"all",style:"thin",color:"#D9D9D9"}};
  s.getRange(`A${methodRow+1}:H${methodRow+1}`).format.rowHeight = 44;
}

// Final calculation, inspection, render and export.
wb.recalculate();
const summaryCheck = await wb.inspect({kind:"table",sheetId:"Summary",range:"A1:H24",include:"values,formulas",tableMaxRows:24,tableMaxCols:8,maxChars:7000});
console.log(summaryCheck.ndjson);
const errorCheck = await wb.inspect({kind:"match",searchTerm:"#REF!|#DIV/0!|#VALUE!|#NAME\\?|#N/A|#NUM!|#NULL!|#SPILL!|#CALC!",options:{useRegex:true,maxResults:300},summary:"final formula error scan"});
console.log(errorCheck.ndjson);

await fs.mkdir(previewDir,{recursive:true});
const previewRanges = {
  "Summary":"A1:H24", "Full EGD Themes":"A1:J24", "EGD Group Profiles":"A1:L10",
  "EGD Section Coverage":"A1:G28", "EGD Evidence":"A1:H22", "PO Brand Preference":"A1:H10",
  "Implications":"A1:F11", "Source Audit":"A1:H18"
};
for (const [sheetName, range] of Object.entries(previewRanges)) {
  const img = await wb.render({sheetName,range,scale:1,format:"png"});
  await fs.writeFile(path.join(previewDir,`${sheetName.replaceAll(" ","_")}.png`),new Uint8Array(await img.arrayBuffer()));
}

await fs.mkdir(outputDir,{recursive:true});
const out = await SpreadsheetFile.exportXlsx(wb);
await out.save(outputPath);
console.log(JSON.stringify({outputPath,transcripts:transcriptNames.length,egdFiles:4,poFiles:4,coverageRows:sectionCoverage.length,evidenceRows:evidence.length,themeRows:themes.length},null,2));
