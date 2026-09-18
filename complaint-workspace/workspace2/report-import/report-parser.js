import { MsgReader } from './vendor/msgreader.esm.js';
let checkCancelled = () => {};
export function setExtractionCancellationCheck(check) { checkCancelled=check; }
function assertExtractionActive() { checkCancelled(); }
function yieldToBrowser() { return new Promise(resolve=>setTimeout(resolve,0)); }
const GENERIC_TEST_RULES = [
  {name:"Visual inspection",re:/\bvisual\s+inspection\b/i,purpose:"Appearance / physical defect review",method:"Inspect customer, retain or reference material and available photographs."},
  {name:"IPW review",re:/\b(?:batch|manufacturing|production)\s+(?:record|documentation)\s+review\b|\breview\s+of\s+(?:the\s+)?(?:batch|manufacturing|production)\s+(?:record|documentation)/i,purpose:"Manufacturing history review",method:"Review production, release and in-process records for deviations or relevant trends."},
  {name:"Peel strength test",re:/\b(?:peel[- ]test|peel\s+strength\s+(?:test|measurement))\b/i,purpose:"Backing adhesion / peel strength",method:"Measure peel or adhesive strength under the report-defined conditions."},
  {name:"Adhesive strength test",re:/\badhesive\s+strength\s+(?:test|measurement)s?\b/i,purpose:"Backing adhesion",method:"Measure adhesive strength and compare with the applicable internal reference."},
  {name:"Illuminated inspection table",re:/\billuminated\s+inspection\s+table\b/i,purpose:"Surface and structure inspection",method:"Inspect membrane on an illuminated inspection table."},
  {name:"Cold light inspection",re:/\bcold\s+light\s+inspection\b/i,purpose:"Surface damage inspection",method:"Inspect membrane using cold-light illumination."},
  {name:"Phenol red drop test",re:/\bphenol\s+red\s+drop\s+test\b/i,purpose:"Surface wetting / flow behavior",method:"Apply phenol-red solution and assess wetting or spreading behavior."},
  {name:"Functional hCG assay",re:/\b(?:functional\s+)?hcg\s+(?:test|assay)\b/i,purpose:"Functional lateral-flow performance",method:"Run the report-defined hCG lateral-flow assay and compare signal behavior with reference."}
];

function splitUniqueValues(value) {
  return String(value||"").split(/\s*;\s*|\n+/).map(item=>item.trim()).filter(Boolean);
}

function mergeSummaryValue(current,incoming) {
  const values=[...splitUniqueValues(current),...splitUniqueValues(incoming)];
  return [...new Set(values.map(value=>value.trim()).filter(Boolean))].join("; ");
}

function normalizeId(v) { return String(v||"").trim().toLowerCase(); }

function normalizeComplaintId(v) {
  const text=normalizeId(v).replace(/[–—]/g,"-");
  const match=text.match(/^(?:comp(?:laint)?[-\s]*)?0*(\d+)$/i);
  return match?String(Number(match[1])):text;
}

function productFamily(material="") {
  const code = material.toUpperCase().replace(/\s/g,"");
  if (code.startsWith("1UN14AR")) return "CN140ub";
  if (code.startsWith("1UN14ER")) return "CN140";
  if (code.startsWith("1UN95")) return "CN95";
  if (code.startsWith("1UN11")) return "CN110";
  if (code.startsWith("1UN18")) return "CN180";
  return "";
}

function normalizedProductFamily(material="", statedFamily="") {
  const materialCodes=String(material||"").toUpperCase().match(/1UN(?:14AR|14ER|95|11|18)[A-Z0-9-]*/g)||[];
  const derived=[...new Set(materialCodes.map(code=>productFamily(code)).filter(Boolean))];
  if (derived.length) return derived.join("; ");
  const stated=String(statedFamily||"").match(/CN140ub|CN140|CN180|CN110|CN95/gi)||[];
  return [...new Set(stated.map(value=>value.toLowerCase()==="cn140ub"?"CN140ub":value.toUpperCase()))].join("; ");
}

function membraneType(material="", description="") {
  return productFamily(material)||productFamilyFromText(description)||"";
}

function complaintClassification(problem="", customerFailure="") {
  const text=cleanBlock(`${problem} ${customerFailure}`).toLowerCase().replace(/[_/]+/g," ");
  const matches=[];
  const rules=[
    ["Poor absorption","Flow / Wetting",/\b(?:poor|low|reduced)?\s*absorption\b|\babsorption\s+(?:issue|problem)\b/],
    ["Abnormal wicking / flow","Flow / Wetting",/\b(?:wicking|slow running|running speed (?:is )?slow|slow flow|fast flow|flow time|much slower|fail(?:ed)? to run)\b/],
    ["Wetting issue","Flow / Wetting",/\b(?:wetting|hydrophilic|hydrophobic|dry spot|white circle)\w*\b/],
    ["Uneven line","Line / Printing",/\buneven\b[^.]{0,45}\b(?:print(?:ing|ed)?|line)s?\b|\b(?:print(?:ing|ed)?|line)s?\b[^.]{0,45}\buneven\b/],
    ["Discontinuous line","Line / Printing",/\b(?:discontinuous|interrupted|broken)\b[^.]{0,45}\b(?:print(?:ing|ed)?|line)s?\b|\b(?:print(?:ing|ed)?|line)s?\b[^.]{0,45}\b(?:discontinuous|interrupted|broken)\b/],
    ["Wide / spreading line","Line / Printing",/\b(?:wide|wider|spreading|diffusion|fringing)\b[^.]{0,75}\b(?:print(?:ing|ed)?|line)s?|\b(?:print(?:ing|ed)?|line)s?\b[^.]{0,75}\b(?:wide|wider|spreading|diffus)/],
    ["Ghost line","Signal / Assay",/\bghost\s+line\b/],
    ["Shadow line","Signal / Assay",/\bshadow(?:s)?\s+(?:on\s+)?(?:stripe|line)s?\b/],
    ["Printing issue","Line / Printing",/\bprint(?:ing|ed)?\s+(?:issue|problem)s?\b/],
    ["Abnormal signal","Signal / Assay",/\b(?:abnormal|low|weak|strong)\s+signal\b|\bsignal\s+(?:low|high|weak|strong|issue)\b/],
    ["Low sensitivity","Signal / Assay",/\blow\s+sensitivity\b/],
    ["False positive","Signal / Assay",/\bfalse\s+positive\b/],
    ["Color / visual issue","Appearance / Surface",/\b(?:coloring|colouring|discolou?r|color (?:is )?(?:abnormal|not consistent)|colour (?:is )?(?:abnormal|not consistent)|color variation|colour variation|blue lines?|visual issue|stain|spot)\w*\b/],
    ["Surface roughness","Appearance / Surface",/\b(?:rough|roughness)\b[^.]{0,35}\b(?:membrane|surface)?\b/],
    ["Imprint","Appearance / Surface",/\bimprints?\b/],
    ["Particles","Appearance / Surface",/\b(?:particle|dust|debris)\w*\b/],
    ["Dirt","Appearance / Surface",/\b(?:dirt|contaminat|foreign material)\w*\b/],
    ["Scratch","Mechanical",/\bscratch(?:es|ed)?\b/],
    ["Crack","Mechanical",/\bcracks?\b|裂痕|裂纹/],
    ["Rupture","Mechanical",/\b(?:rupture|broken membrane)\w*\b/],
    ["Edge damage / telescoping","Mechanical",/\b(?:edge damage|telescop|mechanical damage)\w*\b/],
    ["Membrane sticking","Mechanical",/\bmembrane\s+(?:sticks?|sticking|stuck)\s+(?:together)?\b/],
    ["Backing delamination","Mechanical",/\b(?:membrane|backing)\b[^.]{0,45}\b(?:separat(?:ed|ion)|delaminat(?:ed|ion))\b|\b(?:separat(?:ed|ion)|delaminat(?:ed|ion))\b[^.]{0,45}\b(?:membrane|backing)\b/],
    ["Thickness out of specification","Specification / Conversion",/\bthickness\b[^.]{0,35}\b(?:spec|specification|out)\b/],
    ["Length out of specification","Specification / Conversion",/\b(?:length|short roll)\b[^.]{0,35}\b(?:short|spec|specification|out)?\b/],
    ["Width out of specification","Specification / Conversion",/\bwidth\b[^.]{0,35}\b(?:spec|specification|out)\b/],
    ["Low peel strength","Specification / Conversion",/\b(?:low peel strength|peel strength|backing adhesion|label lifting)\b/],
    ["Bag unsealed","Packaging / Storage",/\b(?:bag unsealed|unsealed bag|open bag|packaging integrity|bags?\s+(?:are\s+)?sealed\s+loosely|loosely\s+sealed\s+bags?|bags?\b[^.]{0,24}\bunsealed)\b/],
    ["Odor / storage concern","Packaging / Storage",/\b(?:odor|odour|storage concern)\b/],
    ["Inter-roll / intra-lot variation","Variation",/\b(?:difference between rolls|roll-to-roll|intra-lot|zone difference|variation between rolls)\b/]
  ];
  for (const [symptom,type,re] of rules) if (re.test(text)) matches.push({symptom,type});
  const symptoms=[...new Set(matches.map(x=>x.symptom))];
  const types=[...new Set(matches.map(x=>x.type))];
  const fallback=cleanBlock(problem).replace(/^(?:performance|quality|product|functionality|functional)(?:\s+\w+)?\s+(?:issue|problem)\s*:?\s*/i,"");
  const shortFallback=fallback && fallback.length<=100 && !/[.!?]\s+.+[.!?]/.test(fallback)
    && !/(?:\bComp\s*-|\b1UN\w+|批次|投诉)/i.test(fallback)?fallback:"";
  const standardizedSymptoms=symptoms.length?symptoms.join("; "):(problemFromChineseText(fallback)||shortFallback||"Review required");
  const problemTypes=types.length?types.join("; "):"Review required";
  const lfaRelevance=types.some(x=>["Flow / Wetting","Line / Printing","Signal / Assay"].includes(x))
    ?"LFA performance"
    :types.some(x=>["Appearance / Surface","Mechanical","Variation"].includes(x))
      ?"Potential LFA impact"
      :types.some(x=>["Specification / Conversion","Packaging / Storage"].includes(x))
        ?"Non-LFA":"Review required";
  return {standardizedSymptoms,problemTypes,lfaRelevance};
}

function parseFlexibleDate(value="") {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  const calendarDate=(year,month,day)=>{
    const date=new Date(Date.UTC(year,month,day));
    return date.getUTCFullYear()===year && date.getUTCMonth()===month && date.getUTCDate()===day?date:null;
  };
  const text=String(value).trim();
  let m=text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (m) return calendarDate(Number(m[1]),Number(m[2])-1,Number(m[3]));
  m=text.match(/^(\d{1,2})[.\/-](\d{1,2})[.\/-](\d{2,4})$/);
  if (m) {
    const year=Number(m[3])+(m[3].length===2?2000:0);
    return calendarDate(year,Number(m[2])-1,Number(m[1]));
  }
  m=text.match(/^(\d{1,2})[.\/-]([A-Za-z]{3,9})[.\/-](\d{2,4})$/);
  if (m) {
    const month=new Date(`${m[2]} 1, 2000`).getMonth();
    const year=Number(m[3])+(m[3].length===2?2000:0);
    if (month>=0) return calendarDate(year,month,Number(m[1]));
  }
  m=text.match(/^(\d{1,2})\s+([A-Za-z]{3,9})\s+(\d{2,4})$/);
  if (m) {
    const month=new Date(`${m[2]} 1, 2000`).getMonth();
    const year=Number(m[3])+(m[3].length===2?2000:0);
    if (month>=0) return calendarDate(year,month,Number(m[1]));
  }
  return null;
}

function daysBetweenDates(start,end) {
  const a=parseFlexibleDate(start), b=parseFlexibleDate(end);
  return a&&b?Math.round((b-a)/86400000):"";
}

function isoDate(value="") {
  const date=value instanceof Date?value:parseFlexibleDate(value);
  if (!date || Number.isNaN(date.getTime())) return String(value||"").trim();
  return `${date.getUTCFullYear()}-${String(date.getUTCMonth()+1).padStart(2,"0")}-${String(date.getUTCDate()).padStart(2,"0")}`;
}

function normalizeExtractedDate(value="") {
  const text=String(value).trim();
  let m=text.match(/^(\d(?:\s*\d)?)\s+([A-Za-z]{3,9})\s+(\d(?:\s*\d){1,3})$/);
  if (m) return `${m[1].replace(/\s/g,"")} ${m[2]} ${m[3].replace(/\s/g,"")}`;
  m=text.match(/^(\d(?:\s*\d)?)\s*([.\/-])\s*([A-Za-z0-9](?:\s*[A-Za-z0-9]){1,8})\s*([.\/-])\s*(\d(?:\s*\d){1,3})$/);
  if (m) return `${m[1].replace(/\s/g,"")}${m[2]}${m[3].replace(/\s/g,"")}${m[4]}${m[5].replace(/\s/g,"")}`;
  return text.replace(/\s*([.\/-])\s*/g,"$1");
}

function firstMatch(text, regexes) {
  for (const re of regexes) {
    const m = text.match(re);
    if (m?.[1]) return m[1].replace(/\s+/g," ").trim().replace(/^[:\s]+|[:\s]+$/g,"");
  }
  return "";
}

function cleanBlock(value="") {
  return String(value)
    .replace(/--- Page \d+ ---/g, " ")
    .replace(/\s+/g, " ")
    .replace(/^[:\s“”"']+|[:\s“”"']+$/g, "")
    .trim();
}

function repairCommonPdfSpacing(value="") {
  return String(value)
    .replace(/\bL\s+ength\b/gi,"Length")
    .replace(/\bC\s+racks?\b/gi,match=>match.toLowerCase().includes("s")?"Cracks":"Crack")
    .replace(/\bDu\s+ring\b/gi,"During")
    .replace(/\bo\s+bserved\b/gi,"observed")
    .replace(/\bH\s+angzhou\b/gi,"Hangzhou")
    .replace(/\bHa\s+ngzhou\b/gi,"Hangzhou")
    .replace(/\bC\s+ustomer\b/gi,"Customer")
    .replace(/\bR\s+eport\b/gi,"Report")
    .replace(/\bS\s+ignal\b/gi,"Signal")
    .replace(/\bO\s+ne\b/gi,"One")
    .replace(/\bP\s+articles\b/gi,"Particles")
    .replace(/\bB\s+ags\b/gi,"Bags")
    .replace(/\bAc\s+on\b/gi,"Acon")
    .replace(/\bUni\s+s\s+art\b/gi,"Unisart")
    .replace(/\bK\s+angpu\b/gi,"Kangpu")
    .replace(/\bL\s+ine\b/gi,"Line")
    .replace(/\branqe\b/gi,"range");
}

function sanitizeSummaryText(value="") {
  return repairCommonPdfSpacing(String(value||""))
    .replace(/\bSartori(?:us|ous)\b/gi,"supplier")
    .replace(/\s{2,}/g," ")
    .trim();
}

function normalizeProblemText(value="") {
  let text=sanitizeSummaryText(cleanBlock(value))
    .replace(/^Customer\s+statement\s*:\s*/i,"")
    .replace(/[“”"]/g,"")
    .trim();
  text=text.split(/\s+The\s+customer\s+claimed\b/i)[0].trim()||text;
  text=text.split(/\s+(?:Customer\s+CRN|Product\s+Code|SSB\s+batch|Please\s+use\s+the\s+reference|A\s+report\s+will\s+be\s+available|Due\s+date\s+complaint\s+report)\s*:?/i)[0].trim()||text;
  return text.replace(/^[„“”"']+|[„“”"']+$/g,"").replace(/\s+([,.;:])/g,"$1").replace(/\s{2,}/g," ");
}

function problemFromChineseText(value="") {
  const matches=[];
  const rules=[["Imprints",/压痕/],["Scratches",/划痕/],["Cracks",/裂痕|裂纹/],["Bag unsealed",/未封口/],["Width out of specification",/宽度变窄/],["Thickness issue",/厚度/]];
  for (const [label,re] of rules) if (re.test(value)) matches.push(label);
  return [...new Set(matches)].join(" / ");
}

function productFamilyFromText(value="") {
  const text=String(value);
  if (/CN\s*140\s*(?:ub|unbacked)/i.test(text)) return "CN140ub";
  if (/CN\s*140/i.test(text)) return "CN140";
  if (/CN\s*180/i.test(text)) return "CN180";
  if (/CN\s*110/i.test(text)) return "CN110";
  if (/CN\s*95/i.test(text)) return "CN95";
  return "";
}

function customerFromFilename(filename="") {
  const tag=String(filename).match(/\(([^()]+)\)(?=\.[^.]+$)/)?.[1]||"";
  if (!tag) return "";
  const label=(tag.split("_").pop()||tag).replace(/[-_]+/g," ").trim();
  return /^[A-Z0-9]{2,}$/.test(label)?label:label.replace(/\b\w/g,char=>char.toUpperCase());
}

function canonicalCustomerName(value="") {
  return repairCommonPdfSpacing(cleanBlock(value))
    .replace(/^the\s+customer\s+/i,"")
    .replace(/^(?:customer|company|organization|account)\s*(?:company|name)?\s*[:#-]?\s*/i,"")
    .replace(/\s*\(\s*/g," (").replace(/\s*\)\s*/g,") ")
    .replace(/\bCo\.?\s*,?\s*Ltd\.?\b/gi,"Co., Ltd.")
    .replace(/\s+([,.;])/g,"$1")
    .replace(/,\s*/g,", ")
    .replace(/\.{2,}$/,".")
    .replace(/\s{2,}/g," ")
    .trim();
}

const CUSTOMER_LOCATIONS = [
  ["Hangzhou","CN",["hangzhou","hz"]],["Xiamen","CN",["xiamen","xm"]],
  ["Jiangsu","CN",["jiangsu","js"]],["Jinan","CN",["jinan","jn"]],
  ["Anhui","CN",["anhui","ah"]],["Tangshan","CN",["tangshan","ts"]],
  ["Nanjing","CN",["nanjing","nj"]],["Guangzhou","CN",["guangzhou","gz"]],
  ["Shenzhen","CN",["shenzhen","sz"]],["Chengdu","CN",["chengdu","cd","sichuan","sc"]],
  ["Chongqing","CN",["chongqing","cq"]],["Beijing","CN",["beijing","bj"]],
  ["Shanghai","CN",["shanghai","sh"]],["Suzhou","CN",["suzhou"]],
  ["Wuhan","CN",["wuhan"]],["Qingdao","CN",["qingdao"]],
  ["Tokyo","JP",["tokyo"]],["Osaka","JP",["osaka"]],["Seoul","KR",["seoul"]],
  ["Mumbai","IN",["mumbai"]],["Delhi","IN",["delhi","new delhi"]],
  ["Athens","GR",["athens"]],["Thessaloniki","GR",["thessaloniki"]],
  ["Manchester","UK",["manchester"]],["Birmingham","UK",["birmingham"]],
  ["Bristol","UK",["bristol"]],["Oxford","UK",["oxford"]],
  ["Toronto","CA",["toronto"]],["Vancouver","CA",["vancouver"]],
  ["Montreal","CA",["montreal","montréal"]],["Ottawa","CA",["ottawa"]],["Calgary","CA",["calgary"]]
];

function countryCodeFromCustomerText(value="",sourceFile="") {
  const text=`${value} ${sourceFile}`.toLowerCase();
  const rules=[
    ["US",/\b(?:usa|u\.s\.a|united states|america)\b/],["JP",/\b(?:japan|japanese|jap)\b/],
    ["KR",/\b(?:south korea|korea|korean)\b/],["IN",/\b(?:india|indian|ind)\b/],
    ["DE",/\b(?:germany|german)\b/],["UK",/\b(?:united kingdom|uk|britain|british)\b/],
    ["GR",/\b(?:greece|greek|gre)\b/],["CA",/\b(?:canada|canadian|can)\b/],
    ["FR",/\b(?:france|french)\b/],["CN",/\b(?:china|chinese|prc)\b/]
  ];
  return rules.find(([,pattern])=>pattern.test(text))?.[0]||"";
}

function compactCustomerName(value="",sourceFile="") {
  let original=canonicalCustomerName(value);
  const originalIsLocation=CUSTOMER_LOCATIONS.some(([name])=>name.toLowerCase()===original.toLowerCase());
  if (!original || originalIsLocation || /^(?:complaint|final)\s+report\b/i.test(original)) {
    original=canonicalCustomerName(customerFromFilename(sourceFile))||original;
  }
  if (!original) return "";
  const existingParts=original.split(",").map(part=>part.trim()).filter(Boolean);
  let existingCountry="", existingCity="";
  if (existingParts.length>=2 && /^(?:CN|US|JP|KR|IN|DE|UK|GR|CA|FR|GB)$/i.test(existingParts.at(-1))) {
    existingCountry=existingParts.pop().toUpperCase().replace(/^GB$/,"UK");
    existingCity=CUSTOMER_LOCATIONS.some(([name])=>name.toLowerCase()===String(existingParts.at(-1)||"").toLowerCase())
      ?existingParts.pop():"";
    const existingKey=existingParts.join(" ").replace(/\s{2,}/g," ").trim();
    if (CUSTOMER_LOCATIONS.some(([name])=>name.toLowerCase()===existingKey.toLowerCase())) {
      original=canonicalCustomerName(customerFromFilename(sourceFile))||existingKey;
    } else original=existingKey;
  }
  const combined=`${original} ${sourceFile}`.toLowerCase();
  let city=existingCity, country=existingCountry||countryCodeFromCustomerText(original,sourceFile);
  for (const [name,code,aliases] of CUSTOMER_LOCATIONS) {
    if (aliases.some(alias=>new RegExp(`(?:^|[^a-z])${alias.replace(/\s+/g,"\\s+")}(?:[^a-z]|$)`,"i").test(combined))) {
      city=name;
      country=country||code;
      break;
    }
  }
  const sourceTag=String(sourceFile).match(/(?:^|[(_-])(USA|US|JAP|JP|KOREA|KR|IND|IN|DE|UK|GRE|GR|CAN|CA)(?=[_)-])/i)?.[1]?.toUpperCase()||"";
  if (!country && sourceTag) country={USA:"US",US:"US",JAP:"JP",JP:"JP",KOREA:"KR",KR:"KR",IND:"IN",IN:"IN",DE:"DE",UK:"UK",GRE:"GR",GR:"GR",CAN:"CA",CA:"CA"}[sourceTag]||sourceTag;
  if (!country && city) country=CUSTOMER_LOCATIONS.find(([name])=>name===city)?.[1]||"";
  let key=original
    .replace(/\([^)]*\)/g," ")
    .replace(/\b(?:hangzhou|xiamen|jiangsu|jinan|anhui|tangshan|nanjing|guangzhou|shenzhen|chengdu|chongqing|beijing|shanghai|suzhou|wuhan|qingdao|sichuan|tokyo|osaka|seoul|mumbai|new delhi|delhi|athens|thessaloniki|manchester|birmingham|bristol|oxford|toronto|vancouver|montreal|montréal|ottawa|calgary)\b/gi," ")
    .replace(/\b(?:china|chinese|prc|usa|united states|america|japan|japanese|korea|korean|india|indian|germany|german|united kingdom|uk|britain|british|greece|greek|canada|canadian)\b/gi," ")
    .replace(/^(?:shandong|anhui|jiangsu)\s+/i,"")
    .replace(/\b(?:biopharm(?:aceuticals?)?|biotech(?:nology)?|biotechnol(?:ogy|gogy)|technology|medical|diagnostics?|diagnostic|healthcare|products?|laborator(?:y|ies))\b.*$/i,"")
    .replace(/\b(?:co\.?\s*,?\s*ltd\.?|company|corporation|corp\.?|inc(?:ms)?\.?|llc|gmbh|limited)\b.*$/i,"")
    .replace(/(?:\s*,\s*){2,}/g,", ")
    .replace(/^[,\s-]+|[,\s-]+$/g,"")
    .replace(/\s{2,}/g," ")
    .trim();
  if (!key) key=original.split(/[,()]/)[0].trim();
  if (/^(?:complaint|final)\s+report\b/i.test(key)) key=customerFromFilename(sourceFile)||"";
  return [key,city,country].filter(Boolean).join(", ");
}

function plausibleCustomerName(value="") {
  const text=canonicalCustomerName(value);
  return Boolean(text && text.length>=2 && text.length<=140 && !/@|Biotech\s+GmbH|Complaint\s+Report|QN\s+no|thank\s+you/i.test(text));
}

function externalCustomerCandidate(value="") {
  return canonicalCustomerName(String(value||"")
    .split(/\s+(?=(?:Sartorius|Supplier)\b)/i)[0]
    .replace(/\s{2,}/g," ")
    .trim());
}

function sectionMatch(text, headingPattern, nextHeadingPattern, maxLength=2400) {
  const re = new RegExp(
    `${headingPattern}\\s*([\\s\\S]{1,${maxLength}}?)(?=${nextHeadingPattern}|$)`,
    "i"
  );
  return cleanBlock(text.match(re)?.[1] || "");
}

function customerCompanyFromHeader(text, filename="", sourceType="") {
  const explicit = externalCustomerCandidate(firstMatch(text, [
    /(?:Customer\s+(?:company|organization|account)|Company\s+name)\s*[:#]?\s*([^\n]{2,160})/i,
    /(?:new\s+complaint\s+report\s+from|please\s+forward[^\n]{0,80}?\s+to)\s*[:#]?\s*([^\n.]{2,160})/i
  ]));
  if (plausibleCustomerName(explicit)) return canonicalCustomerName(explicit);
  if (sourceType==="msg") return customerFromFilename(filename);
  const lines = text.split(/\n/).map(x=>x.trim()).filter(Boolean);
  const emailIndex = lines.findIndex(line => /@[A-Z0-9.-]+\.[A-Z]{2,}/i.test(line));
  if (emailIndex >= 0) {
    for (let i=emailIndex+1; i<Math.min(emailIndex+4, lines.length); i++) {
      const line=lines[i];
      if (/^(?:\d|Report\s+date|Complaint\s+information)/i.test(line)) continue;
      if (/@/.test(line)) continue;
      const candidate=externalCustomerCandidate(line.replace(/\b([A-Za-z]{5,})\s+([a-z])\b/g,"$1$2").replace(/\s{2,}.*$/, "").trim());
      if (!/^(?:From:|Address:|Subject:|Dear\b|We\s+acknowledge)/i.test(candidate) && plausibleCustomerName(candidate)) return canonicalCustomerName(candidate);
    }
  }
  const fallback=externalCustomerCandidate(firstMatch(text, [
    /\n([^\n]*(?:Co\.?|Ltd\.?|Inc\.?|LLC|GmbH|Corporation|Company|Biopharm|Biotechnology)[^\n]*)/i
  ]));
  if (sourceType!=="msg" && !/Biotech\s+GmbH/i.test(fallback)) return canonicalCustomerName(fallback);
  return customerFromFilename(filename);
}

function coordinatorFromText(text) {
  const raw=firstMatch(text, [
    /Written\s+by\s*:\s*(?:Reviewed\s+by\s*:)?\s*\n\s*([A-Z][A-Za-zÀ-ÿ'’-]+(?:\s+[A-Z][A-Za-zÀ-ÿ'’-]+){1,3})/i,
    /\n\s*([A-Z][A-Za-zÀ-ÿ'’-]+(?:\s+[A-Z][A-Za-zÀ-ÿ'’-]+){1,3})\s*\n\s*Quality\s+Professional/i
  ]);
  const parts=raw.split(/\s+/).filter(Boolean);
  return parts.length>=4 ? parts.slice(0,2).join(" ") : raw;
}

function validateProblemAgainstRootCause(problem, customerFailure, rootCauseConclusion) {
  if (!rootCauseConclusion) return "Not checked - root-cause conclusion not extracted";
  const stop = new Set([
    "about","after","against","although","been","being","complaint","conclusion","could","failure",
    "from","have","identified","issue","most","process","product","related","report","root","sample",
    "that","their","there","these","this","those","through","were","which","within","with"
  ]);
  const tokens = value => new Set((String(value).toLowerCase().match(/[a-z]{4,}/g)||[])
    .filter(x=>!stop.has(x))
    .map(x=>x.endsWith("ing")?x.slice(0,-3):x.endsWith("ed")?x.slice(0,-2):x.endsWith("s")?x.slice(0,-1):x));
  const problemTokens=tokens(`${problem} ${customerFailure}`);
  const rootTokens=tokens(rootCauseConclusion);
  const overlap=[...problemTokens].filter(x=>rootTokens.has(x));
  if (overlap.length>=2) return `Consistent with root-cause conclusion (${overlap.slice(0,4).join(", ")})`;
  if (overlap.length===1) return `Broad match - review wording (${overlap[0]})`;
  return "Potential mismatch - review problem description against root-cause conclusion";
}

function enrichProblemDescription(formalProblem, customerFailure, rootCauseConclusion) {
  const formal=cleanBlock(formalProblem);
  let detail=cleanBlock(customerFailure).replace(/[.!?]+$/g,"");
  const isBroad=/^(?:performance|quality|product|functionality|functional)(?:\s+\w+)?\s+(?:issue|problem)$/i.test(formal);
  if (!isBroad) return formal || detail;
  if (/absorption(?:-related)?\s+(?:issue|problem)/i.test(rootCauseConclusion) && !/absorption/i.test(detail)) {
    detail=detail?`Absorption issue; ${detail}`:"Absorption issue";
  }
  if (!detail && /line\s+(?:behavior|quality|printing)/i.test(rootCauseConclusion)) detail="Line quality / printing issue";
  if (!detail) return formal;
  return `${formal}: ${detail.charAt(0).toLowerCase()}${detail.slice(1)}`;
}

function parseMrFr(text) {
  text=String(text).normalize('NFKC').replace(/[–—]/g,'-').replace(/Master\s+roll\s*(?:\(MR\))?\s*[-–—]\s*Final\s+roll\s*(?:\(FR\))?/gi, "MR-FR");
  text=text.replace(/\bMR\s*(\d+)\s*[-–—]\s*FR\s*(\d+(?:\s*[,;]\s*\d+)*)/gi, (_,mr,frs)=>frs.split(/[,;]/).map(fr=>`MR-FR${mr}-${fr.trim()}`).join("; "));
  const pairs = [];
  const re = /\bMR\s*-\s*FR\s*:?\s*([0-9]+(?:\s*\/\s*[0-9]+)?)\s*-\s*([0-9]+)\b/gi;
  let m;
  while ((m = re.exec(text))) {
    const key = `${m[1]}-${m[2]}`;
    if (!pairs.some(p => p.key === key)) pairs.push({ key, master:m[1], final:m[2] });
  }
  const listRe = /MR\s*-\s*FR\s*:?\s*((?:\d+(?:\s*\/\s*\d+)?\s*-\s*\d+)(?:\s*[,;]\s*\d+(?:\s*\/\s*\d+)?\s*-\s*\d+)+)/gi;
  while ((m = listRe.exec(text))) {
    for (const part of m[1].split(/[,;]/)) {
      const p = part.trim().match(/^(\d+(?:\s*\/\s*\d+)?)\s*-\s*(\d+)$/);
      if (p) {
        const key = `${p[1]}-${p[2]}`;
        if (!pairs.some(x => x.key === key)) pairs.push({key, master:p[1], final:p[2]});
      }
    }
  }
  const masters = [...new Set(pairs.map(p=>p.master))];
  const finals = [...new Set(pairs.map(p=>p.final))].sort((a,b)=>Number(a)-Number(b));
  const areas = pairs.map(p=>`MR-FR${p.master}-${p.final}`);

  if (!masters.length) {
    const stripped = text.replace(/MR\s*-\s*FR\s*\d+(?:\s*\/\s*\d+)?\s*-\s*\d+/gi,"");
    const rollLists = [...stripped.matchAll(/\bMR(?:s|[\u00B4'\u2019]s)?\s*:?\s*((?:\d+\s*(?:,|and|&)\s*)+\d+)/gi)];
    for (const match of rollLists) {
      masters.push(...match[1].match(/\d+/g) || []);
    }
    if (!masters.length) {
      const retain = [...stripped.matchAll(/\bMR\s*([0-9]+)\b/gi)].map(x=>x[1]);
      masters.push(...retain);
    }
    masters.splice(0, masters.length, ...new Set(masters));
  }
  return { masters, finals, areas };
}

function selectedCheckbox(text, label, options) {
  const labelPattern=label.replace(/[.*+?^${}()|[\]\\]/g,"\\$&").replace(/\s+/g,"\\s+");
  const labelMatch=text.match(new RegExp(labelPattern,"i"));
  if (!labelMatch || labelMatch.index===undefined) return "";
  const win = text.slice(labelMatch.index, labelMatch.index + 350);
  for (const option of options) {
    const escaped = option.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
    if (new RegExp(`(?:☒|☑|■|\\bX\\b)\\s*${escaped}`, "i").test(win)) return option;
  }
  return "";
}

function extractAssays(text) {
  const assays = [];
  const hasFlowMeasurement = /Capillary\s+Flow\s+Time\s*#|could\s+not\s+be\s+measured\s+against\s+specification/i.test(text);
  const hasInProcessReview = /Review\s+in[- ]process\s+data\s+of\s+capillary\s+flow\s+time/i.test(text);
  if (hasFlowMeasurement) assays.push("Capillary flow time");
  else if (hasInProcessReview) assays.push("IPW review");
  if (/Phenol\s+red\s+(?:buffer\s+)?line\s+test/i.test(text)) assays.push("Phenol red line test");
  if (/Protein\s+(?:binding\s+(?:assay|capacity)|lines?)/i.test(text)) {
    assays.push("Protein binding assay");
  }
  for (const rule of GENERIC_TEST_RULES) if (rule.re.test(text)) assays.push(rule.name);
  return [...new Set(assays)].join("; ");
}

function sourcePageFor(text, phrase) {
  const i=text.toLowerCase().indexOf(phrase.toLowerCase());
  if (i<0) return "";
  const before=text.slice(0,i);
  const pages=[...before.matchAll(/--- Page (\d+) ---/g)];
  return pages.length ? `p.${pages.at(-1)[1]}` : "";
}

function sourcePageRange(text, phrases) {
  const pages=phrases.map(x=>sourcePageFor(text,x)).filter(Boolean).map(x=>Number(x.replace("p.","")));
  if (!pages.length) return "";
  const first=Math.min(...pages), last=Math.max(...pages);
  return first===last?`p.${first}`:`p.${first}-${last}`;
}

function sourceContextForRegex(text,re,maxLength=360) {
  const flags=re.flags.includes("g")?re.flags:re.flags+"g";
  const match=new RegExp(re.source,flags).exec(text);
  if (!match || match.index===undefined) return "";
  const pageEnd=text.indexOf("--- Page ",match.index+1);
  const end=Math.min(match.index+maxLength,pageEnd>match.index?pageEnd:text.length);
  return cleanBlock(text.slice(match.index,end));
}

function genericEvidenceOutcome(result="") {
  const text=String(result).toLowerCase();
  const negatedIssue=/\b(?:no|not|without)\b[^.]{0,45}\b(?:defect|damage|issue|deviation|discrepanc|irregularit|scratch|crack|imprint)/i.test(result);
  if (/within\s+(?:the\s+)?(?:internal\s+)?specification|no\s+(?:visible\s+)?(?:defect|issue|deviation|discrepanc|irregularit)/i.test(result)) return {outcome:"Pass",withinSpec:"Yes",issueObserved:"No"};
  if (!negatedIssue && /below\s+(?:the\s+)?(?:internal\s+)?(?:threshold|specification)|out\s+of\s+specification|reveals?\s+(?:low|a\s+defect)|(?:defect|damage|scratch|crack|imprint)\s+(?:was|were)\s+(?:observed|identified)/i.test(result)) return {outcome:"Issue observed",withinSpec:"No",issueObserved:"Yes"};
  if (text) return {outcome:"Review required",withinSpec:"",issueObserved:""};
  return {outcome:"Recorded",withinSpec:"",issueObserved:""};
}

function extractTestEvidence(text) {
  const tests=[];
  const rollInfo=parseMrFr(text);
  const retainIds=rollInfo.masters.map(x=>`MR${x}`).join("; ");
  const returnedIds=firstMatch(text, [
    /returned\s+\d+\s+(?:pieces|samples)[\s\S]{0,180}?\(([#\d\s,;and]+)(?:identification|customer|\))/i
  ]).replace(/\s+/g," ").replace(/\s*,\s*/g,"; ").replace(/\s+and\s+/gi,"; ");
  if (/capillary\s+flow\s+time/i.test(text)) {
    const result=firstMatch(text, [
      /(The\s+capillary\s+flow\s+time[\s\S]{5,420}?)(?=As\s+the\s+samples|Complaint\s+number|--- Page|$)/i,
      /(in-process[- ]data[\s\S]{5,260}?within\s+specification)/i
    ]);
    tests.push({
      name:/Capillary\s+Flow\s+Time\s*#|could\s+not\s+be\s+measured\s+against\s+specification/i.test(text)
        ?"Capillary flow time":"IPW review",
      purpose:"Flow performance / specification check",
      sampleSource:"Retain sample",
      sampleId:retainIds,
      method:"Measure capillary flow time over 40 mm using the report-defined liquid/conditions and compare with specification/reference.",
      result:cleanBlock(result),
      outcome:/within\s+specification/i.test(result)?"Pass":"Review required",
      withinSpec:/within\s+specification/i.test(result)?"Yes":"",
      issueObserved:"No",
      sourcePage:sourcePageFor(text,"capillary flow time"),
      conditions:/65\s+to\s+115\s+sec\s*\/\s*40\s*mm/i.test(text)
        ?"Specification 65-115 sec/40 mm; customer samples were post-use and not measurable against specification":""
    });
  }
  if (/Phenol\s+red\s+buffer\s+line\s+test/i.test(text)) {
    const result=firstMatch(text, [
      /(The\s+printed\s+buffer\s+lines[\s\S]{5,360}?)(?=Protein\s+binding|Complaint\s+number|--- Page|$)/i,
      /(Phenol\s+red\s+buffer\s+lines[\s\S]{5,240}?(?:irregularities|disruptions|reference))/i
    ]);
    const issue=/irregularit|interrupt|disrupt/i.test(result) && !/without\s+(?:any\s+)?(?:irregularit|interrupt|disrupt)/i.test(result);
    tests.push({
      name:"Phenol red line test", purpose:"Line quality / wetting", sampleSource:"Retain sample",
      sampleId:retainIds,
      method:"Print a phenol-red buffer line on the membrane and compare continuity, width, shape and wetting with reference.",
      result:cleanBlock(result),
      outcome:issue?"Minor irregularity / review":"Pass", withinSpec:issue?"":"Yes",
      issueObserved:issue?"Partial":"No", sourcePage:sourcePageFor(text,"Phenol red buffer line test"),
      conditions:/1\s+or\s+2\s*[μµu]l\s*\/\s*cm/i.test(text)?"Phenol red buffer printed at 1 or 2 µl/cm":""
    });
  }
  if (/Protein\s+(?:binding\s+(?:assay|capacity)|lines?)/i.test(text)) {
    const normal=firstMatch(text, [
      /(The\s+protein\s+lines[\s\S]{5,260}?(?:reference\s+membrane|without\s+any\s+disruptions))/i
    ]);
    const issues=firstMatch(text, [
      /(The\s+following\s+issues\s+were\s+identified[\s\S]{5,720}?)(?=No\s+irregularities|1\.4\.|Manufacturing\s+documentation|--- Page|$)/i
    ]);
    const result=cleanBlock([normal,issues].filter(Boolean).join(" "));
    tests.push({
      name:"Protein binding assay", purpose:"Binding capacity / line morphology",
      sampleSource:returnedIds?"Customer return + retain":"Retain sample",
      sampleId:[returnedIds?`Return pieces ${returnedIds}`:"",retainIds?`Retain ${retainIds}`:""].filter(Boolean).join("; "),
      method:"Print protein line(s), stain with SyproRuby, and evaluate line morphology and binding/signal versus reference.",
      result,
      outcome:issues?"Issue reproduced / mixed":"Pass", withinSpec:issues?"Functionality retained":"Yes",
      issueObserved:issues?"Yes":"No", sourcePage:sourcePageRange(text,["Protein binding","The following issues were identified"]),
      conditions:/0\.5\s+to\s+4\s+mg\s*\/\s*ml/i.test(text)?"Protein 0.5-4 mg/ml; 1 µl/cm; SyproRuby staining and image analysis":""
    });
  }
  const existing=new Set(tests.map(test=>test.name));
  for (const rule of GENERIC_TEST_RULES) {
    if (existing.has(rule.name) || !rule.re.test(text)) continue;
    const result=sourceContextForRegex(text,rule.re);
    const assessment=genericEvidenceOutcome(result);
    const phrase=text.match(rule.re)?.[0]||rule.name;
    tests.push({
      name:rule.name,purpose:rule.purpose,sampleSource:"See report",sampleId:retainIds,
      method:rule.method,result,...assessment,sourcePage:sourcePageFor(text,phrase),conditions:""
    });
  }
  return tests;
}

function complaintIdsFromFilename(filename="") {
  return [...String(filename).matchAll(/Comp\s*-\s*(\d{4,10})/gi)]
    .map(match=>`Comp-${String(Number(match[1])).padStart(7,"0")}`);
}

function isBlankComplaintTemplate(text="") {
  const source=String(text||"");
  return /Customer\s+Complaint\s+Information\s+Form/i.test(source)
    && /Click\s+here\s+to\s+enter\s+text/i.test(source)
    && !/\b(?:Comp\s*-\s*\d{6,10}|13\d{8}|1UN(?:14|95|11|18)[A-Z0-9]+)\b/i.test(source);
}

function receivedSampleDescription(text) {
  // Read the complete value inside this form field, never the adjacent received date.
  const field = text.match(/Number\s+of\s+samples\s+received\s*[:#]?\s*([\s\S]{0,400}?)(?=\s*Date\s+(?:samples|complaint)|\s*Issue\s+description|\s*Complaint\s+status|$)/i);
  if (field) return cleanBlock(field[1]).replace(/\s+/g," ").trim();
  const lines = text.split(/\r?\n/);
  for (const line of lines) {
    if (/\bdate\s+(?:samples\s+)?received/i.test(line)) continue;
    const match = line.match(/^\s*Samples\s+received\s*[:#]?\s*(.+)$/i);
    if (match) return cleanBlock(match[1]).trim();
  }
  return "";
}

function sampleReceivedValue(description) {
  const value=String(description||"").trim();
  // Preserve reported quantities and forms; no roll unit is inferred from implicated rolls.
  return /^\d+(?:[.,]\d+)?$/.test(value)?`${value} (unit not stated)`:value;
}

function parseRecord(text, filename, sourceType) {
  const rawSourceText=text;
  text=repairCommonPdfSpacing(text);
  const searchableText = `${text}\nFile name: ${filename}`;
  let complaintNo = firstMatch(searchableText, [
    /\b(Comp\s*-\s*\d{6,10})\b/i,
    /\b(13\d{8})\b/,
    /(?:Complaint|Notification)\s*(?:number|no\.?|#)\s*[:#]?\s*([A-Z0-9][A-Z0-9\-\/]{4,})/i
  ]);
  complaintNo = complaintNo.replace(/\s+/g, "").replace(/^comp/i, "Comp");
  let material = firstMatch(text, [
    /Product\s+code\s*[:#]?\s*((?:1UN)[A-Z0-9\s]{7,35}?)(?=\s+Lot\s+number|\s+SSB\s+batch|\n)/i,
    /Product\s+code\s*[:#]?\s*([A-Z0-9]+)/i,
    /\b(1UN(?:95|14|11|18)[A-Z0-9]+)\b/i
  ]);
  material = material.replace(/\s+/g, "");
  const productDescription = firstMatch(text, [
    /Product\s+description\s*[:#]?\s*([\s\S]{3,220}?)(?=\s+Total\s+units\s+implicated|\s+Number\s+of\s+samples|\n)/i
  ]);
  const lot = firstMatch(text, sourceType==="msg" ? [
    /Subject\s*:[^\n]*?(?:batch|lot|批次)\s*([0-9]{7,9})/i,
    /Batch\s+number\s+([0-9]{7,9})\s+must\s+be\s+the\s+correct\s+one/i,
    /SSB\s+(?:batch|lot)\s*(?:No\.?|number)?\s*[:#]?\s*([0-9]{7,9})/i,
    /批次\s*([0-9]{7,9})/i,
    /Lot\s+number\s*[:#]?\s*([0-9]{7,9})/i,
    /\blot\s+(?:number\s*)?[:#]?\s*([0-9]{7,9})\b/i
  ] : [
    /Lot\s+number\s*[:#]?\s*([0-9]{7,9})/i,
    /\blot\s+(?:number\s*)?[:#]?\s*([0-9]{7,9})\b/i
  ]);
  let reportDate = firstMatch(text, [
    /(?:Report\s+date|Date\s+of\s+(?:the\s+)?report|Report\s+(?:issued|created)\s+(?:on)?)\s*[:#]?\s*(\d(?:\s*\d)?\s+[A-Za-z]{3,9}\s+\d(?:\s*\d){1,3})/i,
    /(?:Report\s+date|Date\s+of\s+(?:the\s+)?report|Report\s+(?:issued|created)\s+(?:on)?)\s*[:#]?\s*([0-9]{1,2}\s*[-./]\s*[A-Za-z]{3,9}\s*[-./]\s*[0-9]{2,4})/i,
    /(?:Report\s+date|Date\s+of\s+(?:the\s+)?report|Report\s+(?:issued|created)\s+(?:on)?)\s*[:#]?\s*([0-9]{1,4}\s*[-./]\s*[0-9]{1,2}\s*[-./]\s*[0-9]{1,4})/i
  ]);
  reportDate = normalizeExtractedDate(reportDate);
  const customerCompany = canonicalCustomerName(customerCompanyFromHeader(text,filename,sourceType));
  const rollsImplicated = firstMatch(text, [
    /Total\s+units\s+implicated\s*[:#]?\s*(\d+)\s*rolls?/i,
    /(?:Number|Total)\s+of\s+rolls?\s+implicated\s*[:#]?\s*(\d+)/i
  ]);
  let sampleDetails = receivedSampleDescription(text);
  const samplesReceived = sampleReceivedValue(sampleDetails);
  const receivedIdentifiers=firstMatch(text, [
    /returned\s+\d+\s+(?:pieces|samples)[\s\S]{0,180}?\(([#\d\s,;and]+)(?:identification|customer|\))/i
  ]).replace(/\s+/g," ");
  if (receivedIdentifiers && !sampleDetails.includes(receivedIdentifiers)) {
    sampleDetails=cleanBlock(`${sampleDetails} (${receivedIdentifiers})`);
  }
  let formalProblem = normalizeProblemText(firstMatch(text, [
    /(?:Issue|Problem|Complaint)\s+description\s*[:#]?\s*([\s\S]{5,600}?)(?=\s+(?:(?:[A-Za-z]+\s+)?Criticality|Complaint\s+status|Date\s+Complaint|Could\s+the\s+failure|Root\s+cause|Investigation|Final\s+assessment|Conclusion|Figure|Fig\.)\b|$)/i,
    /Customer\s+(?:statement|complaint)\s*[:#]?\s*[“"]?([\s\S]{5,600}?)(?=[”"]?\s*(?:Criticality|Complaint\s+status|Figure|Fig\.|Investigation|Conclusion|$))/i,
    /(?:Issue|Problem|Complaint)\s+description\s*[:#]?\s*(.{5,500})/i,
    /Customer\s+(?:statement|complaint)\s*[:#]?\s*[“"]?(.{5,500})/i,
    /Subject\s*:\s*(.{1,240})/i
  ]).replace(/^["“]|["”]$/g,""));
  if ((!formalProblem || /投诉批次|异常沟通|Final\s+Report/i.test(formalProblem)) && problemFromChineseText(text)) {
    formalProblem=problemFromChineseText(text);
  }
  const status = sourceType === "msg"
    ? "Ongoing – email only"
    : selectedCheckbox(text, "Complaint status", ["Confirmed","Not confirmed","Not conclusive"]);
  const criticality = selectedCheckbox(text, "Criticality", ["Critical","Major","Minor","Track&Trend"]);
  const reproduced = selectedCheckbox(text, "Could the failure be reproduced", ["Yes","No"]);
  const rootCause = selectedCheckbox(text, "Root cause identified within", ["Yes","No"])
    || selectedCheckbox(text, "root cause related", ["Yes","No"]);
  const assaysApplied = extractAssays(text);
  const testEvidence = extractTestEvidence(text);
  const mrfr = parseMrFr(text);
  const zones = firstMatch(text, [
    /Zone\s*\(s\)\s*[:#]?\s*([^\n]{1,120})/i,
    /(?:Membrane|Roll)\s+zones?\s*[:#]?\s*([^\n]{1,120})/i
  ]);
  const mrfrCombined = mrfr.areas.length
    ? mrfr.areas.join("; ")
    : mrfr.masters.map(x=>`MR${x}`).join("; ");
  const sourceGroup = sourceType === "msg" ? "Ongoing - Email" : "Final Reports";
  const customerReportedFailure = normalizeProblemText(firstMatch(text, [
    /received\s+a\s+complaint\s+with\s+the\s+following\s+statement\s*:\s*[“"]?([\s\S]{3,500}?)(?=\s*(?:Picture\s*s|Pictures|Figure|1\.2\.|Criticality|--- Page))/i,
    /Customer\s+statement[\s\S]{0,420}?[“"]([^”"]{3,500})[”"]/i,
    /Customer\s+statement\s*:\s*([\s\S]{3,700}?)(?=\s+(?:The\s+customer\s+claimed|Root\s+cause|(?:Company|Manufacturer)\s+Criticality|Complaint\s+status|1\.2\.|Criticality|--- Page))/i,
    /Issue\s+description\s*:\s*[“"]?([\s\S]{3,700}?)[”"]?\s*(?=Customer\s+CRN|Product\s+Code|SSB\s+batch|Due\s+date|Please\s+use)/i
  ]).replace(/\s*-\s*/g,"-").replace(/[“”"]+/g,""));
  const rootCauseConclusion = sanitizeSummaryText(sectionMatch(
    text,
    "Conclusion\\s+of\\s+the\\s+root\\s+cause\\s+analysis\\s*",
    "(?:4\\.\\s*Correction|4\\.\\s*Conclusion|5\\.\\s*Conclusion|Corrective\\s*/\\s*Preventive|--- Page)",
    2600
  ));
  const problem = enrichProblemDescription(formalProblem, customerReportedFailure, rootCauseConclusion);
  const problemValidation = validateProblemAgainstRootCause(problem, customerReportedFailure, rootCauseConclusion);
  const finalAssessment = sanitizeSummaryText(sectionMatch(
    text,
    "(?:^|\\n)\\s*(?:4|5)\\.\\s*Conclusion\\s*",
    "(?:Best\\s+regards|Written\\s+by|Complaint\\s+number|--- Page)",
    2200
  ));
  const similarEvents = selectedCheckbox(text, "Similar events reported", ["Yes","No"]);
  const containmentNecessary = selectedCheckbox(text, "Containment action necessary", ["Yes","No"]);
  const correctiveActionNecessary = selectedCheckbox(text, "Corrective / Preventive action necessary", ["Yes","No"])
    || selectedCheckbox(text, "Corrective / Preventive actions necessary", ["Yes","No"]);
  const finalScope = sanitizeSummaryText(firstMatch(text, [
    /(The\s+scope\s+of\s+failure\s+is\s+concluded[\s\S]{3,240}?\.)(?=\s|$)/i,
    /(Complaint\s+(?:not\s+confirmed|confirmed)[\s\S]{3,180}?(?:monitoring|claimed\s+units)\.?)/i
  ]));
  const coordinator = coordinatorFromText(text);
  const reportVersion = firstMatch(text, [/Report\s+version\s*[:#]?\s*([A-Z0-9.-]+)/i]);
  let complaintRegisteredDate = firstMatch(text, [
    /Date\s+complaint\s+registered\s*[:#]?\s*(\d(?:\s*\d)?\s*[-./]\s*[A-Za-z0-9](?:\s*[A-Za-z0-9]){1,8}\s*[-./]\s*\d(?:\s*\d){1,3})/i,
    /Date\s+complaint\s+registered\s*[:#]?\s*([0-9]{1,2}\s*[-./]\s*[A-Za-z0-9]{2,9}\s*[-./]\s*[0-9]{2,4})/i,
    /Date\s+complaint\s+registered\s*[:#]?\s*([0-9]{1,2}\s+[A-Za-z]{3,9}\s+[0-9]{2,4})/i
  ]);
  complaintRegisteredDate=normalizeExtractedDate(complaintRegisteredDate);
  const samplesReceivedDate = firstMatch(text, [/Date\s+samples\s+received\s*[:#]?\s*([A-Z0-9][A-Z0-9 .\/-]{1,30})/i]);
  const classification = complaintClassification(problem,customerReportedFailure);
  const daysToReport = daysBetweenDates(complaintRegisteredDate,reportDate);

  const warnings = [];
  if (!complaintNo) warnings.push("Complaint number not extracted");
  if (!material) warnings.push("Material number not extracted");
  if (!lot) warnings.push("Lot not extracted");
  if (!problem) warnings.push("Problem not extracted");
  if (!complaintRegisteredDate) warnings.push("Complaint registered date not extracted");
  if (!reportDate) warnings.push("Report date not extracted");
  if (!customerCompany) warnings.push("Customer not extracted");
  if (!customerReportedFailure) warnings.push("Customer-reported failure not extracted");
  if (classification.problemTypes==="Review required") warnings.push("Standardized symptom classification requires review");
  if (problemValidation.startsWith("Potential mismatch")) warnings.push(problemValidation);
  if (sourceType === "pdf" && text.replace(/\s/g, "").length < 80) {
    warnings.push("This PDF has little or no selectable text; OCR may be required");
  }
  if (sourceType === "pdf" && !status) warnings.push("Complaint status not confidently extracted");
  const filenameComplaints=complaintIdsFromFilename(filename);
  if (filenameComplaints.length===1 && complaintNo && normalizeComplaintId(filenameComplaints[0])!==normalizeComplaintId(complaintNo)) {
    warnings.push(`Filename says ${filenameComplaints[0]}, but report content says ${complaintNo}`);
  }

  return {
    sourceFile: filename, sourceType, sourceGroup,
    complaintNo, reportDate, customerCompany, rollsImplicated, samplesReceived,
    sampleDetails, materialNo: material, productDescription,
    productFamily: productFamily(material)||productFamilyFromText(`${productDescription} ${text.slice(0,1200)}`), lot,
    membraneType:membraneType(material,productDescription)||productFamilyFromText(`${productDescription} ${text.slice(0,1200)}`), zones, mrfrCombined,
    standardizedSymptoms:classification.standardizedSymptoms,
    problemTypes:classification.problemTypes, lfaRelevance:classification.lfaRelevance,
    formalProblem:sanitizeSummaryText(formalProblem), problem:sanitizeSummaryText(problem), assaysApplied, resultStatus:status, criticality,
    masterRolls: mrfr.masters.join("; "),
    finalRolls: mrfr.finals.join("; "),
    mrfrAreas: mrfr.areas.join("; "),
    failureReproduced: reproduced,
    rootCauseRelated: rootCause,
    customerReportedFailure:sanitizeSummaryText(customerReportedFailure), coordinator, reportVersion,
    complaintRegisteredDate, samplesReceivedDate, daysToReport,
    similarEvents, containmentNecessary, correctiveActionNecessary,
    rootCauseConclusion, problemValidation, finalAssessment, finalScope,
    testEvidence:testEvidence.map(test=>Object.fromEntries(Object.entries(test).map(([key,value])=>[key,typeof value==="string"?sanitizeSummaryText(value):value]))),
    warnings: warnings.join("; "),
    rawText:sanitizeSummaryText(rawSourceText).slice(0,30000)
  };
}

async function pdfText(arrayBuffer) {
  const pdfjsLib=globalThis.pdfjsLib;
  if (!pdfjsLib) throw new Error("Local PDF reader is unavailable.");
  pdfjsLib.GlobalWorkerOptions.workerSrc=new URL("./vendor/pdf.worker.min.js",import.meta.url).href;
  const loadingTask = pdfjsLib.getDocument({data:arrayBuffer, isEvalSupported:false});
  try {
    const pdf = await loadingTask.promise;
    const pages = [];
    for (let n=1; n<=pdf.numPages; n++) {
      assertExtractionActive();
      const page = await pdf.getPage(n);
      const content = await page.getTextContent();
      const lines = [];
      let line = [];
      let lastY = null;
      for (const item of content.items) {
        const y = Number(item.transform?.[5]);
        const changedLine = lastY !== null && Number.isFinite(y) && Math.abs(y - lastY) > 2;
        if (line.length && changedLine) {
          lines.push(line.join(" ").trim());
          line = [];
        }
        if (item.str) line.push(item.str);
        if (item.hasEOL && line.length) {
          lines.push(line.join(" ").trim());
          line = [];
        }
        if (Number.isFinite(y)) lastY = y;
      }
      if (line.length) lines.push(line.join(" ").trim());
      pages.push(`--- Page ${n} ---\n${lines.filter(Boolean).join("\n")}`);
      page.cleanup();
      await yieldToBrowser();
    }
    return pages.join("\n");
  } finally {
    // Release the worker and document buffers even after cancellation or a bad page.
    await loadingTask.destroy();
  }
}

async function msgText(arrayBuffer) {
  const reader = new MsgReader(arrayBuffer);
  const info = reader.getFileData();
  const candidates = [
    `Subject: ${info.subject || ""}`,
    `From: ${info.senderName || info.senderEmail || ""}`,
    info.body || "",
    info.bodyHtml || "",
  ];
  return candidates.join("\n").replace(/<[^>]+>/g," ");
}

function markerIndexes(marker="") {
  const output=new Set();
  for (const part of String(marker).split(/[+,;&]/).map(value=>value.trim()).filter(Boolean)) {
    const range=part.match(/^(\d+)\s*-\s*(\d+)$/);
    if (range) {
      const start=Number(range[1]), end=Number(range[2]);
      for (let value=Math.min(start,end);value<=Math.max(start,end);value++) output.add(value);
    } else if (/^\d+$/.test(part)) output.add(Number(part));
  }
  return [...output];
}

function assignMarkedValue(target,marker,value) {
  for (const index of markerIndexes(marker)) if (value) target.set(index,cleanBlock(value));
}

function markedValuesFromArea(area,valuePattern) {
  const values=new Map();
  let match;
  const re=new RegExp(`${valuePattern}\\s*\\(([0-9+,&;\\s-]+)\\)`,"gi");
  while ((match=re.exec(area))) assignMarkedValue(values,match[2],match[1]);
  return values;
}

function splitNumberedPdfCases(text,base) {
  const early=repairCommonPdfSpacing(text.slice(0,12000));
  const ids=new Map();
  let match;
  const idRe=/\b(13\d{8}|Comp\s*-\s*\d{6,10})\s*\((\d+)\)/gi;
  while ((match=idRe.exec(early))) {
    const id=match[1].replace(/\s+/g,"").replace(/^comp/i,"Comp");
    if (!ids.has(Number(match[2]))) ids.set(Number(match[2]),id);
  }
  if (ids.size<2) return [base];

  const materialByCase=new Map();
  const materialRe=/\b(1UN(?:14AR|14ER|95|11|18)[A-Z0-9\s]{5,28}?)\s*\(([0-9+,&;\s-]+)\)/gi;
  while ((match=materialRe.exec(early))) {
    const code=match[1].replace(/\s+/g,"");
    if (/^1UN(?:14AR|14ER|95|11|18)[A-Z0-9]{5,}$/i.test(code)) assignMarkedValue(materialByCase,match[2],code);
  }
  const lotByCase=new Map();
  const lotRe=/\b(2[0-9]{6})\s*\(([0-9+,&;\s-]+)\)/g;
  while ((match=lotRe.exec(early))) assignMarkedValue(lotByCase,match[2],match[1]);
  const rollsByCase=new Map();
  const rollsRe=/\b([0-9]+)\s*rolls?\s*\(([0-9+,&;\s-]+)\)/gi;
  const implicatedArea=early.match(/Total\s+units\s+implicated\s*([\s\S]*?)(?=Number\s+of\s+samples\s+received|Date\s+complaint)/i)?.[1]||"";
  while ((match=rollsRe.exec(implicatedArea))) assignMarkedValue(rollsByCase,match[2],match[1]);
  const samplesByCase=new Map();
  const samplesArea=receivedSampleDescription(early);
  const sampleRe=/([^()]+?)\s*\(([0-9+,&;\s-]+)\)/g;
  while ((match=sampleRe.exec(samplesArea))) assignMarkedValue(samplesByCase,match[2],sampleReceivedValue(match[1]));
  const dateByCase=new Map();
  const dateRe=/(\d{1,2}\s*[-./]\s*[A-Za-z]{3,9}\s*[-./]\s*\d{2,4})\s*\(([0-9+,&;\s-]+)\)/gi;
  while ((match=dateRe.exec(early))) assignMarkedValue(dateByCase,match[2],normalizeExtractedDate(match[1]));

  const problemArea=early.match(/Issue\s+description[\s\S]{0,120}?Customer\s+statement\s*:\s*([\s\S]{1,900}?)(?=The\s+customer\s+claimed|Root\s+cause)/i)?.[1]||"";
  const problemByCase=new Map();
  const problemRe=/(?:^|\n)\s*([^()\n]{3,180}?)\s*\(([0-9+,&;\s-]+)\)\s*(?=\n|$)/g;
  while ((match=problemRe.exec(problemArea))) {
    const value=normalizeProblemText(match[1]);
    if (value && !/^(?:issue\s+description|customer\s+statement)$/i.test(value)) assignMarkedValue(problemByCase,match[2],value);
  }

  const confirmedArea=firstMatch(early,[/Confirmed\s+for\s+lot\s+([\s\S]{1,180}?)(?=Not\s+confirmed|Not\s+conclusive|II\.|$)/i]);
  const notConfirmedArea=firstMatch(early,[/Not\s+confirmed\s+for\s+lot\s+([\s\S]{1,220}?)(?=Not\s+conclusive|II\.|$)/i]);
  const confirmedLots=new Set(confirmedArea.match(/\b\d{7,9}\b/g)||[]);
  const notConfirmedLots=new Set(notConfirmedArea.match(/\b\d{7,9}\b/g)||[]);

  return [...ids.entries()].sort((a,b)=>a[0]-b[0]).map(([index,complaintNo])=>{
    const materialNo=materialByCase.get(index)||base.materialNo;
    const lot=lotByCase.get(index)||base.lot;
    const complaintRegisteredDate=normalizeExtractedDate(dateByCase.get(index)||base.complaintRegisteredDate);
    const problem=problemByCase.get(index)||base.problem;
    let resultStatus=base.resultStatus;
    if (confirmedLots.has(lot)) resultStatus="Confirmed";
    else if (notConfirmedLots.has(lot)) resultStatus="Not confirmed";
    const classification=complaintClassification(problem,problem);
    const notice=`Multi-complaint report: case ${index} of ${ids.size}; review case-specific MR-FR and shared test evidence`;
    return {
      ...structuredClone(base),complaintNo,materialNo,lot,problem,formalProblem:problem,
      customerReportedFailure:problem,rollsImplicated:rollsByCase.get(index)||base.rollsImplicated,
      samplesReceived:samplesByCase.get(index)||base.samplesReceived,
      complaintRegisteredDate,daysToReport:daysBetweenDates(complaintRegisteredDate,base.reportDate),resultStatus,
      productFamily:productFamily(materialNo),membraneType:membraneType(materialNo,base.productDescription),
      standardizedSymptoms:classification.standardizedSymptoms,problemTypes:classification.problemTypes,lfaRelevance:classification.lfaRelevance,
      warnings:mergeSummaryValue(base.warnings,notice)
    };
  });
}

function splitAcknowledgementEmailCases(text,base,filename) {
  const starts=[...text.matchAll(/We\s+acknowledge\s+the\s+receipt\s+of\s+your\s+reported\s+complaint\s*\((Comp\s*-\s*\d{6,10})\)/gi)];
  if (starts.length<2) return [base];
  return starts.map((start,index)=>{
    const end=starts[index+1]?.index ?? text.search(/Please\s+use\s+the\s+reference/i);
    const block=text.slice(start.index,end>start.index?end:text.length);
    const record=parseRecord(block,filename,"msg");
    record.complaintNo=start[1].replace(/\s+/g,"").replace(/^comp/i,"Comp");
    record.customerCompany=customerFromFilename(filename)||base.customerCompany;
    record.warnings=mergeSummaryValue(record.warnings,`Multi-complaint email: extracted complaint ${index+1} of ${starts.length}`);
    return record;
  });
}

async function extractMany(name, buffer) {
  const ext = name.toLowerCase().split(".").pop();
  if (ext === "pdf") {
    const text=await pdfText(buffer);
    if (isBlankComplaintTemplate(text)) return [];
    return splitNumberedPdfCases(text,parseRecord(text,name,"pdf"));
  }
  if (ext === "msg") {
    const text=await msgText(buffer);
    return splitAcknowledgementEmailCases(text,parseRecord(text,name,"msg"),name);
  }
  throw new Error(`Unsupported file: ${name}`);
}

async function extractOne(name, buffer) {
  return (await extractMany(name,buffer))[0];
}


export { productFamily, normalizedProductFamily, membraneType, complaintClassification, parseFlexibleDate, daysBetweenDates, isoDate, normalizeExtractedDate, firstMatch, cleanBlock, repairCommonPdfSpacing, sanitizeSummaryText, normalizeProblemText, problemFromChineseText, productFamilyFromText, customerFromFilename, canonicalCustomerName, countryCodeFromCustomerText, compactCustomerName, plausibleCustomerName, externalCustomerCandidate, sectionMatch, customerCompanyFromHeader, coordinatorFromText, validateProblemAgainstRootCause, enrichProblemDescription, parseMrFr, selectedCheckbox, extractAssays, sourcePageFor, sourcePageRange, sourceContextForRegex, genericEvidenceOutcome, extractTestEvidence, complaintIdsFromFilename, isBlankComplaintTemplate, receivedSampleDescription, sampleReceivedValue, parseRecord, pdfText, msgText, markerIndexes, assignMarkedValue, markedValuesFromArea, splitNumberedPdfCases, splitAcknowledgementEmailCases, extractMany, extractOne };

export { CUSTOMER_LOCATIONS };
