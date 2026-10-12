const fs = require('fs');
const path = require('path');
const {
  Document, Packer, Paragraph, TextRun, ImageRun, Table, TableRow, TableCell,
  WidthType, ShadingType, BorderStyle, AlignmentType, VerticalAlign, HeightRule,
  Header, Footer, PageNumber, TabStopType, LevelFormat, HeadingLevel, PageBreak,
} = require('docx');

const [, , LOGO, OUT] = process.argv;
const logo = fs.readFileSync(LOGO);

const NAVY = '1A2B4C', CREAM = 'F2EBDD', ORANGE = 'E8821E', ORANGE_INK = 'A8520A';
const MUTED = '4F5D78', BORDER = 'DCD1BC', CHIP = 'E8DFCB', STRIPE = 'F4EEE2';
const RED = 'B3161F', FLAGWHITE = 'F7F4EC', BLUE = '1F3F8A', HINT = '8A8F99';
const HEAD = 'Arial Black', BODY = 'Arial';

const PAGE_W = 11906, PAGE_H = 16838;          // A4
const MARGIN = 907;                            // 16 mm
const CONTENT_W = PAGE_W - 2 * MARGIN;         // 10092

const NONE = { style: BorderStyle.NIL, size: 0, color: 'auto' };
const noBorders = { top: NONE, bottom: NONE, left: NONE, right: NONE };
const shade = (fill) => ({ fill, type: ShadingType.CLEAR, color: 'auto' });

// ---------- helpers ----------
const run = (text, o = {}) => new TextRun({ text, font: BODY, size: 20, color: NAVY, ...o });

function heading(num, title, pageBreakBefore = false) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    pageBreakBefore,
    spacing: { before: 320, after: 120 },
    keepNext: true,
    children: [
      new TextRun({ text: num + '  ', font: HEAD, size: 26, color: ORANGE }),
      new TextRun({ text: title.toUpperCase(), font: HEAD, size: 34, color: NAVY }),
    ],
  });
}

function para(children, o = {}) {
  return new Paragraph({ spacing: { after: 100, line: 276 }, children, ...o });
}

function bullet(children) {
  return new Paragraph({ numbering: { reference: 'diamond', level: 0 }, spacing: { after: 70, line: 276 }, children });
}
const lead = (label, text) => [run(label, { bold: true }), run(text)];

function note(text) {
  return para([run(text, { size: 17, color: MUTED })], { spacing: { before: 100, after: 60 } });
}

function cell(text, width, o = {}) {
  const { header, fill, align, bold } = o;
  return new TableCell({
    width: { size: width, type: WidthType.DXA },
    shading: shade(header ? NAVY : (fill || 'FFFFFF')),
    margins: { top: 70, bottom: 70, left: 140, right: 140 },
    verticalAlign: VerticalAlign.CENTER,
    borders: header ? noBorders : { top: NONE, left: NONE, right: NONE, bottom: { style: BorderStyle.SINGLE, size: 4, color: BORDER } },
    children: [new Paragraph({
      alignment: align || AlignmentType.LEFT,
      children: [header
        ? run(text.toUpperCase(), { bold: true, size: 16, color: CREAM, characterSpacing: 20 })
        : run(text, { bold: !!bold, size: 19 })],
    })],
  });
}

function table(widths, head, rows, aligns = []) {
  return new Table({
    width: { size: widths.reduce((a, b) => a + b, 0), type: WidthType.DXA },
    columnWidths: widths,
    rows: [
      new TableRow({ tableHeader: true, children: head.map((h, i) => cell(h, widths[i], { header: true, align: aligns[i] })) }),
      ...rows.map((r, ri) => new TableRow({
        cantSplit: true,
        children: r.map((t, i) => cell(t, widths[i], { fill: ri % 2 ? STRIPE : 'FFFFFF', align: aligns[i], bold: aligns[i] === AlignmentType.RIGHT })),
      })),
    ],
  });
}

// Dashed box with grey hint text for the TO to replace
function placeholder(lines) {
  const dash = { style: BorderStyle.DASHED, size: 8, color: ORANGE };
  return [new Table({
    width: { size: CONTENT_W, type: WidthType.DXA },
    columnWidths: [CONTENT_W],
    rows: [new TableRow({ children: [new TableCell({
      width: { size: CONTENT_W, type: WidthType.DXA },
      borders: { top: dash, bottom: dash, left: dash, right: dash },
      margins: { top: 140, bottom: 140, left: 220, right: 220 },
      children: lines.map((l) => new Paragraph({ spacing: { after: 40, line: 288 }, children: [run(l, { italics: true, color: HINT })] })),
    })] })],
  })];
}

// ---------- cover ----------
const coverRow = (height, fill, children, vAlign = VerticalAlign.TOP, margins = {}) => new TableRow({
  height: { value: height, rule: HeightRule.EXACT },
  children: [new TableCell({
    width: { size: PAGE_W, type: WidthType.DXA },
    shading: shade(fill), borders: noBorders, verticalAlign: vAlign,
    margins: { top: 0, bottom: 0, left: 0, right: 0, ...margins },
    children: children.length ? children : [new Paragraph({ spacing: { after: 0 }, children: [run('', { size: 2 })] })],
  })],
});

const COVER_PAD = { left: 1134, right: 1134 };
const kickerTabs = [{ type: TabStopType.RIGHT, position: PAGE_W - 2 * 1134 }];
const cover = new Table({
  width: { size: PAGE_W, type: WidthType.DXA },
  columnWidths: [PAGE_W],
  rows: [
    coverRow(1300, NAVY, [new Paragraph({
      indent: COVER_PAD, tabStops: kickerTabs, spacing: { before: 1000, after: 0 },
      children: [
        run('DUTCH BLOOD BOWL ASSOCIATION', { bold: true, size: 16, color: ORANGE, characterSpacing: 40 }),
        run('\tSEASON 2027', { bold: true, size: 16, color: ORANGE, characterSpacing: 40 }),
      ],
    })]),
    coverRow(12900, NAVY, [
      new Paragraph({ indent: COVER_PAD, alignment: AlignmentType.CENTER, spacing: { after: 500 },
        children: [new ImageRun({ type: 'png', data: logo, transformation: { width: 330, height: 330 }, altText: { title: 'DBBA logo', description: 'Dutch Blood Bowl Association logo', name: 'logo' } })] }),
      new Paragraph({ indent: COVER_PAD, alignment: AlignmentType.CENTER, spacing: { after: 0 },
        children: [new TextRun({ text: 'RECOMMENDED', font: HEAD, size: 104, color: CREAM })] }),
      new Paragraph({ indent: COVER_PAD, alignment: AlignmentType.CENTER, spacing: { after: 400 },
        children: [new TextRun({ text: 'RULESPACK', font: HEAD, size: 104, color: ORANGE })] }),
      new Paragraph({ indent: COVER_PAD, alignment: AlignmentType.CENTER,
        children: [run('Blood Bowl Sevens tournament rules for DBBA qualifiers', { size: 26, bold: true, color: 'D9D2C3' })] }),
    ], VerticalAlign.CENTER),
    coverRow(180, RED, []),
    coverRow(180, FLAGWHITE, []),
    coverRow(180, BLUE, []),
    coverRow(1100, ORANGE, [new Paragraph({
      indent: COVER_PAD, tabStops: kickerTabs, spacing: { after: 0 },
      children: [
        run('BB2025 · SPIKE! #22', { bold: true, size: 19, color: NAVY, characterSpacing: 30 }),
        run('\tVERSION 1.0 · OCTOBER 2026', { bold: true, size: 19, color: NAVY, characterSpacing: 30 }),
      ],
    })], VerticalAlign.CENTER),
  ],
});

// ---------- content ----------
const C = AlignmentType.CENTER, R = AlignmentType.RIGHT, L = AlignmentType.LEFT;

const tiers = [
  ['Amazon', 'Bretonnian', 'Black Orc', 'Gnome'],
  ['Chaos Dwarf', 'Dwarf', 'Chaos Chosen', 'Goblin'],
  ['Dark Elf', 'Imperial Nobility', 'Chaos Renegade', 'Ogre'],
  ['Elven Union', 'Lizardmen', 'Halfling', 'Snotling'],
  ['Human', 'Nurgle', 'Khorne', ''],
  ['High Elf', 'Orc', '', ''],
  ['Necromantic Horror', 'Shambling Undead', '', ''],
  ['Norse', 'Tomb Kings', '', ''],
  ['Old World Alliance', 'Vampire', '', ''],
  ['Skaven', '', '', ''],
  ['Underworld Denizens', '', '', ''],
  ['Wood Elf', '', '', ''],
];
const q = CONTENT_W / 4;

const body = [
  new Paragraph({
    shading: shade(CHIP),
    border: { left: { style: BorderStyle.SINGLE, size: 36, color: ORANGE, space: 8 } },
    indent: { left: 160, right: 0 },
    spacing: { before: 120, after: 120, line: 300 },
    children: [run('These are the DBBA’s recommended rules for Blood Bowl Sevens tournaments. Tournament organisers add their own schedule, prizes, venue and entry details in the sections marked “To be completed by the TO”.', { size: 21 })],
  }),

  heading('01', 'Ruleset'),
  bullet([run('Blood Bowl 2025 rules, with the Sevens rules from Blood Bowl Spike! Magazine #22, and official errata.')]),
  bullet([run('The tournament is NAF sanctioned.')]),

  heading('02', 'Format'),
  bullet(lead('Resurrection: ', 'every match starts with the team as registered. No SPP or injuries carry over between matches.')),
  bullet(lead('Swiss pairing: ', 'winners play winners and losers play losers. Rematches are avoided where possible.')),

  heading('03', 'Team building'),
  bullet(lead('Team value: ', 'up to 600,000 gold pieces.')),
  bullet(lead('Permitted teams: ', 'the teams listed in the tier table.')),
  bullet(lead('Star Players: ', 'not allowed.')),
  bullet(lead('Inducements: ', 'bought at team creation from your 600,000 gold, and kept for the whole tournament. Tourplay shows which inducements are available to your team.')),

  heading('04', 'Improvement pack'),
  para([run('On top of the 600,000 gold, each team adds extra skills based on its tier.')]),
  table([CONTENT_W - 3 * 2400, 2400, 2400, 2400], ['Tier', 'Primary', 'Secondary', 'Max. Elite'], [
    ['Tier 1', '2', '–', '1'],
    ['Tier 2', '3', '–', '1'],
    ['Tier 3', '2', '1', '1'],
    ['Tier 4', '2', '2', '2'],
  ], [L, C, C, C]),
  new Paragraph({ spacing: { after: 60 }, children: [] }),
  bullet([run('No player may receive more than one extra skill.')]),
  bullet([run('A Primary skill may always be taken instead of a Secondary skill.')]),
  bullet(lead('Banned skills: ', 'Leader and Pro, as per Spike! #22.')),

  new Paragraph({ children: [new PageBreak()] }),
  heading('05', 'Team tiers'),
  table([q, q, q, q], ['Tier 1', 'Tier 2', 'Tier 3', 'Tier 4'], tiers),

  heading('06', 'Scoring'),
  table([CONTENT_W - 2200, 2200], ['Per match', 'Points'], [
    ['Win', '20'],
    ['Draw', '10'],
    ['Loss', '0'],
    ['Touchdown (max. 5 per match)', '1 each'],
    ['Casualty (max. 5 per match)', '1 each'],
    ['Perfect Defence (no TDs conceded)', '3'],
  ], [L, R]),
  note('Casualties count when you cause them on your opponent’s players: blocks, fouls, crowd pushes, Secret Weapons, and thrown or kicked players.'),

  heading('07', 'Tiebreakers'),
  para([run('When coaches finish on equal points, these decide the ranking, in order:')]),
  ...['Score in direct matches', 'Touchdown difference in direct matches', 'Touchdown difference',
    'Casualty difference in direct matches', 'Casualty difference', 'Team value (lowest wins)']
    .map((t) => new Paragraph({ numbering: { reference: 'tiebreak', level: 0 }, spacing: { after: 60 }, children: [run(t)] })),
  note('This is the default tiebreaker order in Tourplay.'),

  // ---------- TO sections ----------
  new Paragraph({ children: [new PageBreak()] }),
  new Paragraph({
    shading: shade(NAVY), spacing: { after: 120 }, indent: { left: 0 },
    border: { top: { style: BorderStyle.SINGLE, size: 1, color: NAVY, space: 6 }, bottom: { style: BorderStyle.SINGLE, size: 1, color: NAVY, space: 6 } },
    children: [
      new TextRun({ text: '  TO BE COMPLETED BY THE TO', font: HEAD, size: 26, color: ORANGE }),
    ],
  }),
  para([run('Replace the grey text in each box with your own tournament details, then remove this note.', { color: MUTED })]),

  heading('08', 'Tournament details'),
  ...placeholder([
    'Tournament name: …',
    'Date: …',
    'Venue and address: …',
    'Entry fee: … (including the €1 Dutch Super Sevens Showdown contribution)',
    'Tickets and registration: … (e.g. buy a ticket at … and register your team on Tourplay)',
    'Contact: … (e-mail, Discord, phone)',
  ]),

  heading('09', 'Schedule'),
  ...placeholder([
    '00:00 – Registration',
    '00:00 – Round 1',
    '00:00 – Lunch break',
    '00:00 – Round 2',
    '00:00 – Round 3',
    '00:00 – Round 4',
    '00:00 – Awards and closing',
  ]),

  heading('10', 'Prizes and awards'),
  ...placeholder([
    'List the prizes and awards for your tournament, for example:',
    '1st place, 2nd place, 3rd place',
    'Most casualties, most touchdowns, Best Painted, Wooden Spoon',
  ]),
];

const doc = new Document({
  creator: 'Dutch Blood Bowl Association',
  title: 'DBBA Recommended Rulespack',
  description: 'Recommended Blood Bowl Sevens tournament rules for DBBA qualifiers',
  styles: {
    default: { document: { run: { font: BODY, size: 20, color: NAVY } } },
    paragraphStyles: [{
      id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
      run: { font: HEAD, size: 34, color: NAVY },
      paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 0 },
    }],
  },
  numbering: {
    config: [
      { reference: 'diamond', levels: [{ level: 0, format: LevelFormat.BULLET, text: '■', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 360, hanging: 300 } }, run: { color: ORANGE, size: 16 } } }] },
      { reference: 'tiebreak', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1', alignment: AlignmentType.LEFT,
        style: { paragraph: { indent: { left: 400, hanging: 400 } }, run: { color: ORANGE, bold: true } } }] },
    ],
  },
  sections: [
    {
      properties: { page: { size: { width: PAGE_W, height: PAGE_H }, margin: { top: 0, bottom: 0, left: 0, right: 0, header: 0, footer: 0 } } },
      children: [cover],
    },
    {
      properties: {
        page: {
          size: { width: PAGE_W, height: PAGE_H },
          margin: { top: 1300, bottom: 1100, left: MARGIN, right: MARGIN, header: 450, footer: 450 },
          pageNumbers: { start: 1 },
        },
      },
      headers: {
        default: new Header({ children: [new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
          border: { bottom: { style: BorderStyle.SINGLE, size: 12, color: NAVY, space: 4 } },
          children: [
            new TextRun({ text: 'DBBA RECOMMENDED RULESPACK', font: HEAD, size: 22, color: NAVY }),
            new TextRun({ text: '\t', font: BODY }),
            new ImageRun({ type: 'png', data: logo, transformation: { width: 34, height: 34 }, altText: { title: 'DBBA logo', description: 'DBBA logo', name: 'logo-small' } }),
          ],
        })] }),
      },
      footers: {
        default: new Footer({ children: [new Paragraph({
          tabStops: [{ type: TabStopType.RIGHT, position: CONTENT_W }],
          children: [
            run('DBBA RECOMMENDED RULESPACK · V1.0', { size: 14, bold: true, color: MUTED, characterSpacing: 10 }),
            new TextRun({ children: ['\t', PageNumber.CURRENT], font: BODY, size: 14, bold: true, color: MUTED }),
          ],
        })] }),
      },
      children: body,
    },
  ],
});

Packer.toBuffer(doc).then((buf) => { fs.writeFileSync(OUT, buf); console.log('wrote', OUT); });
