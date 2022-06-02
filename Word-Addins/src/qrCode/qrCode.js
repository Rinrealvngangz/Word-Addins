const QRCode = require('qrcode')
let colorHexEnd  = "#ffffffff";
let level = "L";
let version = "2";

Office.onReady(async(info) => {
  document.getElementById("btnQrCode").onclick = run;
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged,async (eventArgs)=>{
    await run();
 });
});


 async function run() {
  return Word.run(async (context) => {
   const range = context.document.getSelection();
   const data = range.load("text");
    await context.sync();
    const textArea = document.getElementById("textInputArea");
    if(data.text){
      document.getElementById("textInputArea").value = data.text;
    }
    if(textArea.value !== ""){  
      getValueLevel();
      version = document.getElementById("capacity").value;
      const options ={
        ColorDot:"#000000ff",
        ColorBackground:colorHexEnd,
        level:level,
        version:version
      }
     QRImageUrl(textArea.value,options);
    }
  });
}
function message(text){
  document.getElementById("test").innerText = text;
}

function QRCanvas(text){
  var canvas = document.getElementById('canvas')
  QRCode.toCanvas(canvas, text, function (error) {
    if (error) console.error(error)
    console.log('success!');
})
}

function QRImageUrl(text,options){
  var opts = {
    errorCorrectionLevel: options.level,
    type: 'image/jpeg',
    version:options.version,
    quality: 0.3,
    margin: 1,
    color: {
      dark:options.ColorDot,
      light:options.ColorBackground
    }
  }
  
  QRCode.toDataURL(text, opts, function (err, url) {
    if (err) {
      console.log(`ERR:${err}`);
      message(err);
    }else{
      message("");
      var img = document.getElementById('image');
      img.src = url
    }
   
  })
}


$("#mycolorEnd").colorpicker({
  defaultPalette: 'web',
  history: false
});

$("#mycolorEnd").on("change.color", function(event, color){
  colorHexEnd =color;
});

function getValueLevel() {
  var levelValues = document.getElementsByName('errorLevel');
  for(i = 0; i < levelValues.length; i++) {
      if(levelValues[i].checked)
     level = levelValues[i].value;
  }
}