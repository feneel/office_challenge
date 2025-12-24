/// <reference types="office-js" />

console.log(" Executing taskpane!!!")

function setStatus(msg: string){
    const element = document.getElementById("status-text")
    if (element) element.textContent = msg
}


document.addEventListener("DOMContentLoaded", ()=>{
    setStatus("Ready. Click Connect!!")
})