/// <reference types="office-js" />

console.log(" Executing taskpane!!!")

function setStatus(msg: string) {
    const element = document.getElementById("status-text")
    if (element) element.textContent = msg
}


function wireClick() {
    const btn = document.getElementById("run-btn") as HTMLButtonElement | null;
    if (!btn) return;

    btn.onclick = async () => {
        setStatus("Clicked. Checking environment...")
        ""
        if (typeof Office === "undefined" || typeof Word === "undefined") {
            setStatus("Office/ Word API not available!!")
            return;
        }

        const canUseTracking = Office.context.requirements.isSetSupported("WordApi", "1.5")
        setStatus(`Connecting to Word.. Tracking: ${canUseTracking ? "Yes" : "No"}`)

        try {
            setStatus("Connecting to Word....")


            await Word.run(async (context) => {

                if (canUseTracking) {
                    context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll
                }

                const body = context.document.body
                body.load("text")
                await context.sync();


                const text = (body.text || "").replace(/\s+/g, " ")
                const preview = text.slice(0, 120)


                setStatus(`Connected!! Preview: "${preview}${text.length > 120 ? "_" : "}"}"`)



            })


        }

        catch (e) {
            console.error("Word.run failed! Check console!")
        }


    }
}


document.addEventListener("DOMContentLoaded", () => {
    wireClick();
    setStatus("Ready. Click Connect!!")
})

if (typeof Office != "undefined") {
    Office.onReady(() => {
        wireClick()
    })
}