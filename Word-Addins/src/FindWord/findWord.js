let colorHexHightlight = "pink";
let colorHexText ="purple";
Office.onReady(async(info) => {
    document.getElementById("btnFind").onclick = findWord;
    await getCheckOption();
 });
 let textGlobal =""
async function findWord(){
await Word.run(async (context) => {
    const text = document.getElementById("find").value;
    textGlobal = text;
    const searchResults = context.document.body.search(text, {matchCase:false});
    // Queue a command to load the font property values.
    searchResults.load('font');
    await context.sync();
    const totalSearch = searchResults.items.length;
    document.getElementById("totalResult").innerText =`Result for ${text}: ${totalSearch}`;
    if(totalSearch>0){
        let listFindText="";
        for (let i = 0; i < searchResults.items.length; i++) {
            let elementP = `<p data-value=text-${i}>${i+1}:${text}</p>`;
            listFindText += elementP;
        }
        document.getElementById("results").innerHTML = listFindText;
        createHandleEvent(text);
        document.getElementById("optionSearch").style.display ="block";
    }else{
        document.getElementById("results").innerHTML = "";
        document.getElementById("optionSearch").style.display ="none";
    }
  
    await context.sync();
});
}
  function createHandleEvent(text){
  const p_array = document.getElementsByTagName("p");
   p_array.forEach((p)=>{
    p.addEventListener("click",async function() {
        const indexSearch = p.dataset.value.split('-')[1];
       await navigatePositionText(text,indexSearch);
      })
   })
}

async function navigatePositionText(text,indexSearch){
     
        await Word.run(async (context) => {
            const searchResults = context.document.body.search(text, {matchWildcards: true});
            searchResults.load('font');
            await context.sync();
            searchResults.items[indexSearch].select("Select");
            await context.sync();
        });
    }

async function startOptionSearch(text,value){
        await Word.run(async (context) => {
            const searchResults = context.document.body.search(text, {matchWildcards: true});
            searchResults.load('font');
            await context.sync();
            for (let i = 0; i < searchResults.items.length; i++) {
                optionSearch(searchResults,value,i);
            }
            await context.sync();
        });
}

function optionSearch(searchResults,value,i){
    if(value === "color"){
        searchResults.items[i].font.color = colorHexText;
    }if(value === "hightlight"){
        searchResults.items[i].font.highlightColor = colorHexHightlight;  
    }if(value === "bold"){
        searchResults.items[i].font.bold = true;
    }
}    

 async function getCheckOption(){
    let allCheckBox = document.querySelectorAll('.colorSearch')
    allCheckBox.forEach((checkbox) => { 
    checkbox.addEventListener('change',async (event) => {
      if (event.target.checked) {
        const val =event.target.value;
        await startOptionSearch(textGlobal,val)
      }
    })
  })
}

$("#mycolorEnd").colorpicker({
    defaultPalette: 'web',
    history: false
  });
  
  $("#mycolorEnd").on("change.color", function(event, color){
    colorHexText =color;
  });

  $("#mycolorStart").colorpicker({
    defaultPalette: 'web',
    history: false
  });
  
  $("#mycolorStart").on("change.color", function(event, color){
    colorHexHightlight =color;
  });

