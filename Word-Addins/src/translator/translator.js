/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */
// npm cache --force clean  
// npm install --force
// const translate = require("@vitalets/google-translate-api");
Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("insert").onclick = writeData;
    Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, async (eventArgs) => {
      await translate();
    });
  }
});

const API_KEY = "9f7725a9-cb3d-4814-bfb2-a2a033e19d60";
const select = document.getElementById("selectCountry");
const textArea = document.getElementById("textTranSlated");

function writeData(){
  Office.context.document.setSelectedDataAsync(textArea.value, function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
      write(asyncResult.error.message);
    }
  });
}

const countryListAlpha2 = {
  auto: "Automatic",
  af: "Afrikaans",
  sq: "Albanian",
  am: "Amharic",
  ar: "Arabic",
  hy: "Armenian",
  az: "Azerbaijani",
  eu: "Basque",
  be: "Belarusian",
  bn: "Bengali",
  bs: "Bosnian",
  bg: "Bulgarian",
  ca: "Catalan",
  ceb: "Cebuano",
  ny: "Chichewa",
  "zh-CN": "Chinese (Simplified)",
  "zh-TW": "Chinese (Traditional)",
  co: "Corsican",
  hr: "Croatian",
  cs: "Czech",
  da: "Danish",
  nl: "Dutch",
  en: "English",
  eo: "Esperanto",
  et: "Estonian",
  tl: "Filipino",
  fi: "Finnish",
  fr: "French",
  fy: "Frisian",
  gl: "Galician",
  ka: "Georgian",
  de: "German",
  el: "Greek",
  gu: "Gujarati",
  ht: "Haitian Creole",
  ha: "Hausa",
  haw: "Hawaiian",
  he: "Hebrew",
  iw: "Hebrew",
  hi: "Hindi",
  hmn: "Hmong",
  hu: "Hungarian",
  is: "Icelandic",
  ig: "Igbo",
  id: "Indonesian",
  ga: "Irish",
  it: "Italian",
  ja: "Japanese",
  jw: "Javanese",
  kn: "Kannada",
  kk: "Kazakh",
  km: "Khmer",
  ko: "Korean",
  ku: "Kurdish (Kurmanji)",
  ky: "Kyrgyz",
  lo: "Lao",
  la: "Latin",
  lv: "Latvian",
  lt: "Lithuanian",
  lb: "Luxembourgish",
  mk: "Macedonian",
  mg: "Malagasy",
  ms: "Malay",
  ml: "Malayalam",
  mt: "Maltese",
  mi: "Maori",
  mr: "Marathi",
  mn: "Mongolian",
  my: "Myanmar (Burmese)",
  ne: "Nepali",
  no: "Norwegian",
  ps: "Pashto",
  fa: "Persian",
  pl: "Polish",
  pt: "Portuguese",
  pa: "Punjabi",
  ro: "Romanian",
  ru: "Russian",
  sm: "Samoan",
  gd: "Scots Gaelic",
  sr: "Serbian",
  st: "Sesotho",
  sn: "Shona",
  sd: "Sindhi",
  si: "Sinhala",
  sk: "Slovak",
  sl: "Slovenian",
  so: "Somali",
  es: "Spanish",
  su: "Sundanese",
  sw: "Swahili",
  sv: "Swedish",
  tg: "Tajik",
  ta: "Tamil",
  te: "Telugu",
  th: "Thai",
  tr: "Turkish",
  uk: "Ukrainian",
  ur: "Urdu",
  uz: "Uzbek",
  vi: "Vietnamese",
  cy: "Welsh",
  xh: "Xhosa",
  yi: "Yiddish",
  yo: "Yoruba",
  zu: "Zulu",
};;

function loadCountry(){
  Object.entries(countryListAlpha2).forEach(([key,value]) => {
       var option = document.createElement("option");
       key.toLowerCase() === 'vi' ? option.selected ="selected" : value;
       option.value = key.toLowerCase();
       option.text = value;
       select.add(option);
  });
}
loadCountry();

async function getSelectionText(){
   const result=  await Word.run(async (context)=>{
    let paragraph = context.document.getSelection();
    paragraph.load('text');
    await context.sync(); 
    return paragraph.text;
  });
  return result;
}

async function checkSelectedText(){
  let text = await getSelectionText();
  if(text ===''){
    console.log('empty')
  }else{
    return text;
  }
}
async function autoDetect(textSelection){

  const objectDectect=  fetch(`https://api-translate.systran.net/compatmode/google/language/translate/v2/detect?key=${API_KEY}&q=${textSelection}`)
  .then(kq=>kq.json()).then(kq=> {return kq.data.detections[0].language})
  .catch(err=> console.log(err))
 return objectDectect;
}
async function translate(){
  let optionTarget = select.value;
  let textSelection = await checkSelectedText();
  let languageCode = await autoDetect(textSelection);
  if(textSelection && languageCode){
    const langpair = `${languageCode}|${optionTarget}`;
   fetch(`https://api.mymemory.translated.net/get?q=${textSelection}&langpair=${langpair}`)
   .then(kq=>kq.json())
   .then(kq=> textArea.innerHTML = kq.responseData.translatedText)
  }
}