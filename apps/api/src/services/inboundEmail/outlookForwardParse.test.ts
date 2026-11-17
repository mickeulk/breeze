import { describe, it, expect } from 'vitest';
import { extractOutlookForwardedSender } from './outlookForwardParse';

// Shapes modelled on what Outlook writes when a user forwards a message.

/** Outlook on the web / new Outlook: <hr> + <div id="divRplyFwdMsg">. */
function owaForward(headerRows: string[], note = 'Can you look at this?'): string {
  return [
    '<html><body>',
    `<div dir="ltr">${note}</div>`,
    '<div id="appendonsend"></div>',
    '<hr style="display:inline-block;width:98%" tabindex="-1">',
    '<div id="divRplyFwdMsg" dir="ltr"><font face="Calibri, sans-serif" style="font-size:11pt" color="#000000">',
    headerRows.join('<br>'),
    '</font><div>&nbsp;</div></div>',
    '<div>Our printer stopped working.</div>',
    '</body></html>',
  ].join('');
}

/** Classic Outlook desktop: top-bordered div + MsoNormal paragraph. */
function classicForward(headerRows: string[], color = 'E1E1E1'): string {
  return [
    '<div class=WordSection1><p class=MsoNormal>FYI<o:p></o:p></p><div>',
    `<div style='border:none;border-top:solid #${color} 1.0pt;padding:3.0pt 0cm 0cm 0cm'>`,
    `<p class=MsoNormal>${headerRows.join('<br>')}<o:p></o:p></p>`,
    '</div></div><p class=MsoNormal>Our printer stopped working.<o:p></o:p></p></div>',
  ].join('');
}

const EN = [
  '<b>From:</b> Jane Client &lt;jane@client.example&gt;',
  '<b>Sent:</b> Monday, November 17, 2026 9:12 AM',
  '<b>To:</b> Tech &lt;tech@msp.example&gt;',
  '<b>Subject:</b> Printer down',
];
const FR = [
  '<b>De&nbsp;:</b> Jean Dupont &lt;jean@client.example&gt;',
  '<b>Envoyé&nbsp;:</b> lundi 17 novembre 2026 09:12',
  '<b>À&nbsp;:</b> Support &lt;support@msp.example&gt;',
  '<b>Objet&nbsp;:</b> Imprimante',
];

describe('extractOutlookForwardedSender', () => {
  it('parses an Outlook on the web forward (English)', () => {
    expect(extractOutlookForwardedSender(owaForward(EN))).toBe('jane@client.example');
  });

  it('parses an Outlook on the web forward (French labels, non-breaking spaces)', () => {
    expect(extractOutlookForwardedSender(owaForward(FR))).toBe('jean@client.example');
  });

  it('parses a classic Outlook forward with a mailto link', () => {
    const rows = [
      '<b><span lang=FR>De&nbsp;:</span></b><span lang=FR> Jean Dupont &lt;<a href="mailto:jean@client.example">jean@client.example</a>&gt;',
      '<b>Envoyé&nbsp;:</b> lundi 17 novembre 2026 09:12',
      '<b>À&nbsp;:</b> support@msp.example',
      '<b>Objet&nbsp;:</b> Imprimante</span>',
    ];
    expect(extractOutlookForwardedSender(classicForward(rows))).toBe('jean@client.example');
  });

  it('parses the older classic Outlook border colour and the [mailto:] form', () => {
    const rows = [
      '<b>From:</b> Jane Client [mailto:jane@client.example]',
      '<b>Sent:</b> Monday, November 17, 2026 9:12 AM',
      '<b>Subject:</b> Printer down',
    ];
    expect(extractOutlookForwardedSender(classicForward(rows, 'B5C4DF'))).toBe('jane@client.example');
  });

  it('parses Dutch and German header labels', () => {
    const nl = ['<b>Van:</b> Jan &lt;jan@klant.example&gt;', '<b>Verzonden:</b> maandag 17 november 2026', '<b>Onderwerp:</b> Printer'];
    const de = ['<b>Von:</b> Jan &lt;jan@kunde.example&gt;', '<b>Gesendet:</b> Montag, 17. November 2026', '<b>Betreff:</b> Drucker'];
    expect(extractOutlookForwardedSender(owaForward(nl))).toBe('jan@klant.example');
    expect(extractOutlookForwardedSender(owaForward(de))).toBe('jan@kunde.example');
  });

  it('parses the new Outlook for Mac reference container', () => {
    const html = '<div>FYI</div><div id="mail-editor-reference-message-container"><div>'
      + EN.join('<br>') + '</div><div>Body</div></div>';
    expect(extractOutlookForwardedSender(html)).toBe('jane@client.example');
  });

  it('picks the real mailbox past a quoted display name holding an address', () => {
    const rows = ['<b>From:</b> "attacker@evil.example" &lt;jane@client.example&gt;', ...EN.slice(1)];
    expect(extractOutlookForwardedSender(owaForward(rows))).toBe('jane@client.example');
  });

  it('uses only the first (most recent) forward block', () => {
    const inner = owaForward(['<b>From:</b> Old &lt;old@other.example&gt;', ...EN.slice(1)], 'older');
    const html = owaForward(EN) + inner.replace('divRplyFwdMsg', 'x_divRplyFwdMsg');
    expect(extractOutlookForwardedSender(html)).toBe('jane@client.example');
  });

  it('returns null when the first block has a display name but no address (no deeper fallback)', () => {
    const first = owaForward(['<b>From:</b> Jane Client', ...EN.slice(1)]);
    const deeper = owaForward(['<b>From:</b> Bob &lt;bob@other.example&gt;', ...EN.slice(1)]);
    expect(extractOutlookForwardedSender(first + deeper)).toBeNull();
  });

  it('requires a Sent/Date or Subject header next to the From', () => {
    expect(extractOutlookForwardedSender(owaForward(['<b>From:</b> Jane &lt;jane@client.example&gt;', '<b>To:</b> tech@msp.example']))).toBeNull();
  });

  it('requires the block to start with the From header', () => {
    expect(extractOutlookForwardedSender(owaForward([EN[1]!, EN[0]!, EN[3]!]))).toBeNull();
  });

  it('ignores header-looking lines without an Outlook anchor', () => {
    const html = `<div>From: Jane &lt;jane@client.example&gt;<br>Sent: Monday<br>Subject: Printer</div>`;
    expect(extractOutlookForwardedSender(html)).toBeNull();
  });

  it('ignores a top border that is not Outlook\'s forward separator', () => {
    const html = classicForward(EN).replace('#E1E1E1', '#123456');
    expect(extractOutlookForwardedSender(html)).toBeNull();
  });

  it('returns null for empty input and plain mail', () => {
    expect(extractOutlookForwardedSender(undefined)).toBeNull();
    expect(extractOutlookForwardedSender('')).toBeNull();
    expect(extractOutlookForwardedSender('<p>Hello, my printer is down.</p>')).toBeNull();
  });
});
