/**
 * Outlook forwarded-message parsing for the staff-forward intake path on
 * connected Microsoft 365 mailboxes (the Gmail counterpart is forwardParse.ts).
 *
 * Microsoft Graph hands the pipeline the message body as HTML, and Outlook
 * marks a forward structurally rather than with a text marker line:
 *
 *  - Outlook on the web, new Outlook for Windows/Mac and Outlook mobile wrap
 *    the forwarded header block in `<div id="divRplyFwdMsg">` (the visible
 *    separator is an `<hr>`, the "underscore rule" in a text rendering);
 *  - classic Outlook for Windows/Mac opens it with a top-bordered
 *    `<div style="border:none;border-top:solid #E1E1E1 1.0pt;...">` (#B5C4DF in
 *    older versions);
 *  - newer Outlook for Mac uses `<div id="mail-editor-reference-message-container">`.
 *
 * The header block's labels follow the forwarder's Outlook language, so the
 * From / Sent / Subject labels are recognized in English, French, Dutch,
 * German, Spanish and Portuguese.
 *
 * Deterministic and bounded, like forwardParse.ts: only those anchors count,
 * and only when the text right after the anchor starts with a From header line
 * followed by a Sent/Date or Subject header in the same contiguous block.
 * Only the FIRST recognized block is used (the most recent forward); when its
 * From line carries no address (classic Outlook can show a directory display
 * name alone), the result is null and the message routes normally. The address
 * is parsed, never a display name, and the caller uses its DOMAIN only, gated
 * on the outer message being verified and sent by partner-level staff (see
 * the staff-forward step in inboundEmailService).
 *
 * Recognition is structural but still textual: a staff member who pastes the
 * same HTML gets the same routing. That is accepted for the same reason as on
 * the Gmail path: only authenticated partner-level staff with access to the
 * target org can trigger it, and it only chooses which customer org the new
 * ticket is filed under.
 */
import { htmlToText } from './htmlToText';
import { addrFromLine } from './forwardParse';

/** Structural anchors Outlook writes in front of a forwarded header block. */
const ANCHORS: readonly RegExp[] = [
  /<[a-z][^<>]*\bid\s*=\s*["']?(?:x_)?divRplyFwdMsg\b/gi,
  /<[a-z][^<>]*\bid\s*=\s*["']?(?:x_)?mail-editor-reference-message-container\b/gi,
  /<div\b[^<>]*style\s*=\s*["'][^"'<>]*border\s*:\s*none\s*;\s*border-top\s*:\s*solid\s+#(?:E1E1E1|B5C4DF)\s+1(?:\.0)?pt/gi,
];

/** How many anchors are tried, and how much HTML after each is flattened. */
const MAX_ANCHORS = 5;
const MAX_BLOCK_HTML = 20_000;
/** How many lines a header block may span (From, Sent, To, Cc, Subject, ...). */
const MAX_HEADER_LINES = 12;

// Labels as Outlook writes them, per language. A French label is followed by a
// (non-breaking) space before the colon: "De :", "Envoyé :", "Objet :".
const FROM_LABEL = /^(?:From|De|Van|Von)\s?:/;
const DATE_LABEL = /^(?:Sent|Date|Envoyé|Verzonden|Gesendet|Datum|Enviado|Fecha|Enviada)\s?:/;
const SUBJECT_LABEL = /^(?:Subject|Objet|Onderwerp|Betreff|Asunto|Assunto)\s?:/;
const HEADER_LINE = /^[\p{L}][\p{L} .-]{0,30}\s?:/u;

/** Positions of every anchor, in document order, at most MAX_ANCHORS. */
function anchorPositions(html: string): number[] {
  const found = new Set<number>();
  for (const re of ANCHORS) {
    re.lastIndex = 0;
    let m: RegExpExecArray | null;
    while ((m = re.exec(html)) !== null) {
      found.add(m.index);
      if (found.size > MAX_ANCHORS * ANCHORS.length) break;
    }
  }
  return [...found].sort((a, b) => a - b).slice(0, MAX_ANCHORS);
}

type BlockResult = { recognized: false } | { recognized: true; from: string | null };

/** Read the header block that starts the flattened text after an anchor. */
function readHeaderBlock(text: string): BlockResult {
  const lines = text
    .replace(/\r\n/g, '\n')
    .replace(/[  ]/g, ' ')
    .split('\n')
    .map((l) => l.trim());
  let i = 0;
  while (i < lines.length && !lines[i]) i++;
  const first = lines[i];
  if (first === undefined || !FROM_LABEL.test(first)) return { recognized: false };

  let corroborated = false;
  for (let j = i + 1; j < Math.min(lines.length, i + MAX_HEADER_LINES); j++) {
    const inner = lines[j] ?? '';
    if (!inner || !HEADER_LINE.test(inner)) break;
    if (DATE_LABEL.test(inner) || SUBJECT_LABEL.test(inner)) corroborated = true;
  }
  if (!corroborated) return { recognized: false };
  return { recognized: true, from: addrFromLine(first.replace(FROM_LABEL, '')) };
}

/**
 * Extract the original sender's email address from an Outlook-forwarded HTML
 * body, or null if there is no recognized forward block (see the module comment).
 */
export function extractOutlookForwardedSender(html: string | null | undefined): string | null {
  if (!html) return null;
  for (const pos of anchorPositions(html)) {
    const text = htmlToText(html.slice(pos, pos + MAX_BLOCK_HTML));
    const block = readHeaderBlock(text);
    if (block.recognized) return block.from;
  }
  return null;
}
