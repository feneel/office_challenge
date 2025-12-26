/// <reference types="office-js" />

let isRunning = false;

function setStatus(msg: string) {
  const element = document.getElementById("status-text");
  if (element) element.textContent = msg;
}

async function addConfidentialHeader(context: Word.RequestContext): Promise<boolean> {
  const headerText = "CONFIDENTIAL DOCUMENT";

  const sections = context.document.sections;
  sections.load("items");
  await context.sync();

  const headers: any[] = [];
  for (const section of sections.items) {
    const header = section.getHeader(Word.HeaderFooterType.primary);
    header.load("text");
    headers.push(header);
  }
  await context.sync();

  let changed = false;

  for (const header of headers) {
    const current = (header.text || "").replace(/\s+/g, " ").trim();
    if (current.includes(headerText)) continue;

    const p = header.insertParagraph(headerText, Word.InsertLocation.start);
    p.font.bold = true;
    p.alignment = Word.Alignment.centered;

    changed = true;
  }

  await context.sync();
  return changed;
}

function wireClick() {
  const btn = document.getElementById("run-btn") as HTMLButtonElement | null;
  if (!btn) return;

  btn.onclick = async () => {
    if (isRunning) return;
    isRunning = true;
    btn.disabled = true;

    try {
      setStatus("Clicked. Checking environment…");

      if (typeof Office === "undefined" || typeof Word === "undefined") {
        setStatus("Office/Word API not available (not running inside Word).");
        return;
      }

      const canUseTracking = Office.context.requirements.isSetSupported("WordApi", "1.5");
      setStatus(`Connecting to Word… Tracking: ${canUseTracking ? "Yes" : "No"}`);

      await Word.run(async (context) => {
        if (canUseTracking) {
          context.document.changeTrackingMode = Word.ChangeTrackingMode.trackAll;
          await context.sync(); // ensure tracking applies before edits
        }

        const headerAdded = await addConfidentialHeader(context);

        const body = context.document.body;
        body.load("text");
        await context.sync();

        const text = (body.text || "").replace(/\s+/g, " ").trim();
        const preview = text.slice(0, 120);

        setStatus(
          `Done. Track Changes: ${canUseTracking ? "ON" : "OFF"} | Header: ${
            headerAdded ? "added" : "already present"
          } | Preview: "${preview}${text.length > 120 ? "…" : ""}"`
        );
      });
    } catch (e) {
      console.error(e);
      setStatus(`Word.run failed: ${String(e)}`);
    } finally {
      btn.disabled = false;
      isRunning = false;
    }
  };
}

document.addEventListener("DOMContentLoaded", () => {
  wireClick();
  setStatus("Ready. Click Connect!!");
});
