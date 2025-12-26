/// <reference types="office-js" />

let isRunning = false;

function setStatus(msg: string) {
  const element = document.getElementById("status-text");
  if (element) element.textContent = msg;
}



type RedactionCounts = {
  emails: number,
  phones: number,
  ssns: number,
  total: number
}


const REDACTION_MARK = "🀫🀫🀫🀫🀫🀫▍"

function extractUniqueMatches(text: string, re: RegExp): string[] {

  const set = new Set<string>()
  let m: RegExpExecArray | null


  while ((m = re.exec(text)) !== null) {
    const hit = (m[0] || "").trim()

    if (!hit) continue;


    if (hit.includes("🀫") || hit.toLowerCase().includes("redacted")) continue

    set.add(hit)

  }

  return Array.from(set)


}


async function replaceAllOccurrences(
  context: Word.RequestContext,
  body: Word.Body,
  literal: string,
  replacement: string
): Promise<number> {
  const results = body.search(literal, {
    matchCase: false,
    matchWholeWord: false,
    matchWildcards: false,
  ignorePunct: true,
  ignoreSpace: true,
  } as any);

  results.load("items");
  await context.sync();

  const count = results.items.length;
  if (count === 0) return 0;

  const canClearLinks = Office.context.requirements.isSetSupported("WordApi", "1.4");
  if (canClearLinks) {
    for (const r of results.items) r.load("hyperlink");
    await context.sync();
  }

  for (const r of results.items) {
    if (canClearLinks && r.hyperlink) r.hyperlink = "";
    r.insertText(replacement, Word.InsertLocation.replace);
  }


  return count;
}


async function redactSensitiveInfo(context: Word.RequestContext): Promise<RedactionCounts> {

  const body = context.document.body

  body.load("text")

  await context.sync()


  const text = body.text || ""

  const emailRe = /\b[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}\b/gi;

  const phoneRe =   /\b(?:\+?1[\s\u00A0().\-‐-‒–—]*)?(?:\(?\d{3}\)?[\s\u00A0().\-‐-‒–—]*)\d{3}[\s\u00A0().\-‐-‒–—]*\d{4}\b/g;

  const ssnRe = /\b\d{3}[- ]\d{2}[- ]\d{4}\b/g;


  const emails = extractUniqueMatches(text, emailRe)
  const phones = extractUniqueMatches(text, phoneRe)
  const ssns = extractUniqueMatches(text, ssnRe)


    setStatus(`Found matches: emails=${emails.length}, phones=${phones.length}, ssns=${ssns.length}`);


  const counts: RedactionCounts = { emails: 0, phones: 0, ssns: 0, total: 0 }

  for (const e of emails) counts.emails += await replaceAllOccurrences(context, body, e, REDACTION_MARK)
  for (const p of phones) counts.phones += await replaceAllOccurrences(context, body, p, REDACTION_MARK)
  for (const s of ssns) counts.ssns += await replaceAllOccurrences(context, body, s, REDACTION_MARK)


  counts.total = counts.emails + counts.phones + counts.ssns


  await context.sync()
  return counts

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

        const redactions= await redactSensitiveInfo(context)


        const body = context.document.body;
        body.load("text");
        await context.sync();

        const text = (body.text || "").replace(/\s+/g, " ").trim();
        const preview = text.slice(0, 120);

        setStatus(
          `Done. Track Changes: ${canUseTracking ? "ON" : "OFF"} | Header: ${headerAdded ? "added" : "already present"
          } | 
          `+ 
          `Redacted: ${redactions.total} (Emails ${redactions.emails}, Phones: ${redactions.phones}, SSNs: ${redactions.ssns})`+
          
          `Preview: "${preview}${text.length > 120 ? "…" : ""}"`
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
