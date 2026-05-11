const fs = require("fs");
const path = require("path");
const { execFileSync } = require("child_process");
const {
  Document, Packer, Paragraph, TextRun, Header, Footer,
  AlignmentType, HeadingLevel, PageNumber, PageBreak, LineRuleType,
  TableOfContents, TabStopType, TabStopPosition, LeaderType, ImageRun, Bookmark, SimpleField,
  Table, TableRow, TableCell, WidthType, BorderStyle,
  Math: DocxMath, MathRun, MathSubScript, MathRadical
} = require("docx");

const CM = (cm) => Math.round(cm * 567);
const PT = (pt) => pt * 2;
const SONGTI = "宋体";
const HEITI = "黑体";
const KAITI = "楷体";
const TNR = "Times New Roman";
const TOC_TITLE_FONT = HEITI;
const MIXED_FONT = { ascii: TNR, hAnsi: TNR, eastAsia: SONGTI, cs: TNR };
const TOC1_FONT = { ascii: TNR, hAnsi: TNR, eastAsia: HEITI, cs: TNR };

// ===== 占位/通用常量 =====
const UNIVERSITY_NAME = "天津理工大学";
const COLLEGE_NAME = "【学院名称】";
const MAJOR_NAME = "【专业名称】";
const GRADE_YEAR = "【年级】";
const STUDENT_ID = "【学号】";
const STUDENT_NAME = "【姓名】";
const ADVISOR_NAME = "【指导教师】";
const THESIS_YEAR = "【年份】";
const THESIS_MONTH = "【月份】";

const THESIS_TITLE_CN = "【中文论文标题】";
const THESIS_TITLE_EN = [
  "【English Thesis Title Part 1】",
  "【English Thesis Title Part 2】"
];
const ZH_ABSTRACT_TEXT = "【此处填写中文摘要】";
const EN_ABSTRACT_TEXT = "【Here goes the English Abstract】";
const ZH_KEYWORDS = "【关键词1】 【关键词2】 【关键词3】";
const EN_KEYWORDS = "【Keyword1】; 【Keyword2】; 【Keyword3】";

// 如果需要自动生成图表，配置脚本路径和资源
const VSDX_DIR = "F:/论文/vsdx"; // 请根据需要修改
const FLOWCHART_SCRIPT = path.join(VSDX_DIR, "generate_flowcharts.py");
const FIGURE_ASSET_CONFIG = {
  // 示例:
  // "1.1": { file: path.join(VSDX_DIR, "图1.1.png"), width: 380, height: 882 },
};

function ensureFigureAssets() {
  if (fs.existsSync(FLOWCHART_SCRIPT)) {
    execFileSync("python", [FLOWCHART_SCRIPT], { stdio: "inherit" });
  }
}

function loadFigureAssets() {
  return Object.fromEntries(
    Object.entries(FIGURE_ASSET_CONFIG).map(([number, config]) => ([
      number,
      { ...config, data: fs.readFileSync(config.file) }
    ]))
  );
}

// 确保资源存在并加载（视需求取消注释）
// ensureFigureAssets();
const FIGURE_IMAGES = loadFigureAssets();

const NO_BORDER = { style: BorderStyle.NONE, size: 0, color: "FFFFFF" };
const THIN_BORDER = { style: BorderStyle.SINGLE, size: 4, color: "000000" };
const THICK_BORDER = { style: BorderStyle.SINGLE, size: 20, color: "000000" };
const TABLE_BORDER = {
  top: NO_BORDER,
  bottom: NO_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
  insideHorizontal: NO_BORDER,
  insideVertical: NO_BORDER,
};
const HEADER_BORDER = {
  top: THICK_BORDER,
  bottom: THIN_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
};
const BODY_BORDER = {
  top: NO_BORDER,
  bottom: NO_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
};
const LAST_ROW_BORDER = {
  top: NO_BORDER,
  bottom: THICK_BORDER,
  left: NO_BORDER,
  right: NO_BORDER,
};

// 中文字号
const YI_HAO = 52;    // 一号 26pt
const XIAO_YI = 48;   // 小一 24pt
const ER_HAO = 44;    // 二号 22pt
const XIAO_ER = 36;   // 小二 18pt
const SAN_HAO = 32;   // 三号 16pt
const XIAO_SAN = 30;  // 小三 15pt
const SI_HAO = 28;    // 四号 14pt
const XIAO_SI = 24;   // 小四 12pt
const WU_HAO = 21;    // 五号 10.5pt
const XIAO_WU = 18;   // 小五 9pt
const CITATION_SIZE = XIAO_SI;

function p(text, opts = {}) {
  const {
    size = XIAO_SI, bold = false, font = SONGTI,
    alignment = AlignmentType.JUSTIFIED,
    spacing = { line: 300, before: 0, after: 0 },
    indent, firstLine = 480
  } = opts;
  const pOpts = { alignment, spacing, children: [] };
  if (alignment === AlignmentType.JUSTIFIED || alignment === AlignmentType.LEFT) {
    pOpts.indent = { firstLine, ...(indent || {}) };
  }
  if (indent) pOpts.indent = { ...pOpts.indent, ...indent };
  pOpts.children = [new TextRun({ text, font, size, bold })];
  return new Paragraph(pOpts);
}
function pRuns(children, opts = {}) {
  const {
    alignment = AlignmentType.JUSTIFIED,
    spacing = { line: 300, before: 0, after: 0 },
    indent, firstLine = 480
  } = opts;
  const pOpts = { alignment, spacing, children };
  if (alignment === AlignmentType.JUSTIFIED || alignment === AlignmentType.LEFT) {
    pOpts.indent = { firstLine, ...(indent || {}) };
  }
  if (indent) pOpts.indent = { ...pOpts.indent, ...indent };
  return new Paragraph(pOpts);
}
function citationField(id) {
  return new SimpleField(`REF ref_${id} \\h`, String(id));
}
function citationMarkerRun(text) {
  return new TextRun({ text, font: TNR, size: CITATION_SIZE, superScript: true });
}
function citationChildren(ids, separator = ",") {
  const refs = Array.isArray(ids) ? ids : [ids];
  const children = [citationMarkerRun("[")];
  refs.forEach((id, index) => {
    children.push(citationField(id));
    if (index < refs.length - 1) {
      children.push(citationMarkerRun(separator));
    }
  });
  children.push(citationMarkerRun("]"));
  return children;
}

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    alignment: AlignmentType.CENTER,
    spacing: { before: 480, after: 480, line: 360 },
    children: [new TextRun({ text, font: HEITI, size: SAN_HAO, bold: true })]
  });
}
function h2(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_2,
    alignment: AlignmentType.CENTER,
    spacing: { before: 360, after: 360, line: 360 },
    children: [new TextRun({ text, font: HEITI, size: XIAO_SAN, bold: true })]
  });
}
function h3(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_3,
    spacing: { before: 240, after: 240, line: 360 },
    children: [new TextRun({ text, font: HEITI, size: SI_HAO, bold: true })]
  });
}
function sectionLead(text) {
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 120, after: 120, line: 300 },
    indent: { left: 480 },
    children: [new TextRun({ text, font: HEITI, size: XIAO_SI })]
  });
}
function paragraphTitle(title, body) {
  return new Paragraph({
    alignment: AlignmentType.JUSTIFIED,
    spacing: { line: 300, before: 0, after: 0 },
    indent: { firstLine: 480 },
    children: [
      new TextRun({ text: `${title} `, font: SONGTI, size: XIAO_SI, bold: true }),
      new TextRun({ text: body, font: SONGTI, size: XIAO_SI })
    ]
  });
}
function bodySection(title, body) {
  return [sectionLead(title), p(body)];
}
function equationLine(children, number) {
  return new Paragraph({
    spacing: { before: 120, after: 120, line: 300 },
    tabStops: [
      { type: TabStopType.CENTER, position: TabStopPosition.MAX / 2 },
      { type: TabStopType.RIGHT, position: TabStopPosition.MAX },
    ],
    children: [
      new TextRun({ text: "\t", font: TNR, size: XIAO_SI }),
      new DocxMath({ children }),
      new TextRun({ text: `\t(${number})`, font: TNR, size: XIAO_SI }),
    ],
  });
}
function equationDisplay(children) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 120, after: 120, line: 300 },
    children: [new DocxMath({ children })],
  });
}
function m(text) {
  return new MathRun(text);
}
function abstractTitle(text, opts = {}) {
  const { font = HEITI, size = ER_HAO, bold = false } = opts;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: 300 },
    children: [new TextRun({ text, font, size, bold })]
  });
}
function abstractHeading(text, opts = {}) {
  const { font = HEITI, size = SAN_HAO, bold = false } = opts;
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 0, after: 0, line: 300 },
    children: [new TextRun({ text, font, size, bold })]
  });
}
function abstractParagraph(text, opts = {}) {
  const { font = KAITI, size = SI_HAO, alignment = AlignmentType.JUSTIFIED } = opts;
  return new Paragraph({
    alignment,
    spacing: { before: 0, after: 0, line: 400, lineRule: LineRuleType.EXACT },
    indent: { firstLine: 480 },
    children: [new TextRun({ text, font, size })]
  });
}
function keywordParagraph(label, text, opts = {}) {
  const {
    labelFont = HEITI,
    labelSize = SAN_HAO,
    labelBold = false,
    valueFont = KAITI,
    valueSize = SI_HAO,
    indent = { firstLine: 480 },
    colon = "："
  } = opts;
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 0, line: 400, lineRule: LineRuleType.EXACT },
    ...(indent ? { indent } : {}),
    children: [
      new TextRun({ text: `${label}${colon}`, font: labelFont, size: labelSize, bold: labelBold }),
      new TextRun({ text: ` ${text}`, font: valueFont, size: valueSize })
    ]
  });
}
function blankLine(size = XIAO_SI, font = SONGTI) {
  return new Paragraph({
    children: [new TextRun({ text: "\u200B", size, font })]
  });
}
function empty() {
  return new Paragraph({ children: [new TextRun("")] });
}
function pb() {
  return new Paragraph({ children: [new PageBreak()] });
}

function coverLine(label, value) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 80, after: 80, line: 360 },
    children: [
      new TextRun({ text: label + "    ", font: SONGTI, size: SAN_HAO }),
      new TextRun({ text: value, font: SONGTI, size: SAN_HAO })
    ]
  });
}

function figure(number, title, opts = {}) {
  const asset = FIGURE_IMAGES[number];
  if (!asset) {
    throw new Error(`未找到图 ${number} 的流程图资源`);
  }
  const { width = asset.width, height = asset.height } = opts;
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 120, line: 360 },
      children: [
        new ImageRun({
          type: "png",
          data: asset.data,
          transformation: { width, height }
        })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 0, line: 360 },
      children: [
        new TextRun({ text: "图", font: KAITI, size: WU_HAO }),
        new TextRun({ text: `${number} `, font: TNR, size: WU_HAO }),
        new TextRun({ text: title, font: KAITI, size: WU_HAO })
      ]
    }),
    empty(),
  ];
}

function referenceParagraph(text) {
  const match = text.match(/^\[(\d+)\]\s*(.+)$/);
  const number = match ? match[1] : "";
  const body = match ? match[2] : text;
  return new Paragraph({
    alignment: AlignmentType.LEFT,
    spacing: { before: 0, after: 0, line: 300 },
    indent: { left: 480, hanging: 480 },
    children: [
      new TextRun({ text: "[", font: TNR, size: XIAO_SI }),
      new Bookmark({
        id: `ref_${number}`,
        children: [new TextRun({ text: number, font: TNR, size: XIAO_SI })],
      }),
      new TextRun({ text: "] ", font: TNR, size: XIAO_SI }),
      new TextRun({ text: body, font: MIXED_FONT, size: XIAO_SI }),
    ]
  });
}

const REFERENCES = [
  // "[1] 示例参考文献1",
  // "[2] 示例参考文献2",
];

function tableText(text, opts = {}) {
  const { alignment = AlignmentType.CENTER, bold = false, font = SONGTI } = opts;
  return new Paragraph({
    alignment,
    spacing: { before: 80, after: 80, line: 300 },
    children: [new TextRun({ text, font, size: WU_HAO, bold })]
  });
}

function tableCaption(number, title) {
  return new Paragraph({
    alignment: AlignmentType.CENTER,
    spacing: { before: 240, after: 120, line: 360 },
    children: [
      new TextRun({ text: "表", font: KAITI, size: WU_HAO }),
      new TextRun({ text: `${number} `, font: TNR, size: WU_HAO }),
      new TextRun({ text: title, font: KAITI, size: WU_HAO })
    ]
  });
}

function tableCell(text, opts = {}) {
  const { alignment = AlignmentType.CENTER, bold = false, width, borders } = opts;
  return new TableCell({
    width,
    ...(borders ? { borders } : {}),
    children: [tableText(text, { alignment, bold })]
  });
}

function placeholderTable(number, title, headers, rows, opts = {}) {
  const { columnWidths = [], columnAlignments = [] } = opts;
  const getWidth = (index) =>
    columnWidths[index] ? { size: columnWidths[index], type: WidthType.DXA } : undefined;
  const getAlignment = (index) =>
    columnAlignments[index] || AlignmentType.CENTER;

  return [
    tableCaption(number, title),
    new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      alignment: AlignmentType.CENTER,
      layout: "fixed",
      borders: TABLE_BORDER,
      rows: [
        new TableRow({
          tableHeader: true,
          children: headers.map((header, index) =>
            tableCell(header, { bold: true, width: getWidth(index), borders: HEADER_BORDER })
          )
        }),
        ...rows.map((row, rowIndex) =>
          new TableRow({
            children: headers.map((_, index) =>
              tableCell(row[index] || "待补充", {
                alignment: getAlignment(index),
                width: getWidth(index),
                borders: rowIndex === rows.length - 1 ? LAST_ROW_BORDER : BODY_BORDER
              })
            )
          })
        ),
      ]
    }),
    empty(),
  ];
}

// ===== 封面（已按需求注释，文档从摘要开始） =====
/*
const cover = [
  empty(), empty(), empty(),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 200, line: 360 },
    children: [new TextRun({ text: UNIVERSITY_NAME, font: HEITI, size: YI_HAO, bold: true })] }),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 500, line: 360 },
    children: [new TextRun({ text: "本科毕业设计说明书（毕业论文）", font: HEITI, size: ER_HAO, bold: true })] }),
  empty(), empty(),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 300, line: 360 },
    children: [new TextRun({ text: THESIS_TITLE_CN, font: HEITI, size: ER_HAO, bold: true })] }),
  empty(), empty(), empty(),
  coverLine("学    院", COLLEGE_NAME),
  coverLine("专    业", MAJOR_NAME),
  coverLine("年    级", GRADE_YEAR),
  coverLine("学    号", STUDENT_ID),
  coverLine("姓    名", STUDENT_NAME),
  coverLine("指导教师", ADVISOR_NAME),
  empty(), empty(),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 300 },
    children: [new TextRun({ text: `${THESIS_YEAR}年${THESIS_MONTH}月`, font: SONGTI, size: SAN_HAO })] }),
];
*/
const cover = [];

// ===== 声明（已按需求注释，文档从摘要开始） =====
/*
const declaration = [
  pb(),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 480, after: 480, line: 360 },
    children: [new TextRun({ text: "独创性声明", font: HEITI, size: SAN_HAO, bold: true })] }),
  p(`本人声明所呈交的毕业设计说明书（毕业论文）是本人在指导教师指导下进行的研究工作和取得的研究成果。据我查证，除了文中特别加以标注和致谢的地方外，本论文中不包含其他人已经发表或撰写过的研究成果，也不包含为获得${UNIVERSITY_NAME}或其他教育机构的学位或证书而使用过的材料。与我一同工作的同志对本研究所做的任何贡献均已在论文中作了明确的说明并表示了谢意。`),
  empty(),
  new Paragraph({ alignment: AlignmentType.RIGHT, spacing: { before: 300 },
    children: [new TextRun({ text: "毕业设计说明书（毕业论文）作者签名：              ", font: SONGTI, size: XIAO_SI })] }),
  new Paragraph({ alignment: AlignmentType.RIGHT,
    children: [new TextRun({ text: "签字日期：    年    月    日", font: SONGTI, size: XIAO_SI })] }),
];
*/
const declaration = [];

// ===== 中文摘要 =====
const zhAbstract = [
  abstractTitle(THESIS_TITLE_CN),
  blankLine(ER_HAO),
  abstractHeading("摘　要"),
  abstractParagraph(ZH_ABSTRACT_TEXT),
  empty(),
  keywordParagraph("关键词", ZH_KEYWORDS),
  empty(),
  empty(),
];

// ===== 英文摘要 =====
const enAbstract = [
  abstractTitle(THESIS_TITLE_EN[0], { font: TNR }),
  abstractTitle(THESIS_TITLE_EN[1], { font: TNR }),
  blankLine(ER_HAO, TNR),
  abstractHeading("ABSTRACT", { font: TNR, bold: true }),
  abstractParagraph(EN_ABSTRACT_TEXT, { font: TNR }),
  empty(),
  keywordParagraph("Key Words", EN_KEYWORDS, {
    labelFont: TNR,
    labelBold: true,
    valueFont: TNR,
    indent: null,
    colon: ":",
  }),
];

// ===== 目录 =====
const toc = [
  pb(),
  new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 480, after: 480, line: 300 },
    children: [new TextRun({ text: "目    录", font: HEITI, size: SAN_HAO, bold: true })] }),
  blankLine(XIAO_WU),
  blankLine(XIAO_WU),
  new TableOfContents("", { hyperlink: true, headingStyleRange: "1-3" }),
];

// ===== 第1章 示例章节 =====
const ch1 = [
  pb(),
  h1("第一章  【第一章标题】"),
  h2("1.1 【节标题】"),
  p("【此处填写段落内容】"),
  h3("1.1.1 【小节标题】"),
  p("【此处填写小节内容】"),
  // placeholderTable("1.1", "示例表格", ["列1", "列2"], [["数据1", "数据2"]], { columnWidths: [CM(5), CM(5)] }),
  // figure("1.1", "示例图表", { width: 400, height: 300 }),
];

// ===== 参考文献 =====
const referencesSection = [
  pb(),
  h1("参考文献"),
  empty(),
  ...REFERENCES.map(referenceParagraph),
];

// ===== 致谢 =====
const thanks = [
  pb(),
  h1("致    谢"),
  p("本论文的完成得益于各方的关心与支持。在此，我谨向所有给予我指导和帮助的人表达最诚挚的感谢。", { firstLine: 0, alignment: AlignmentType.LEFT }),
  p("【致谢内容】"),
];

const JSZip = require("jszip");

async function patchLang(buf) {
  const zip = await JSZip.loadAsync(buf);
  const citationRunPropsXml = `<w:rPr><w:rFonts w:ascii="Times New Roman" w:hAnsi="Times New Roman" w:eastAsia="宋体"/><w:sz w:val="${CITATION_SIZE}"/><w:szCs w:val="${CITATION_SIZE}"/><w:vertAlign w:val="superscript"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr>`;

  // 为 styles.xml 的默认 rPr 添加 w:lang eastAsia
  const stylesXml = await zip.file("word/styles.xml").async("string");
  const patched = stylesXml
    .replace(/<w:docDefaults>/, '$&<!-- patched -->')
    .replace(
      /(<w:rPrDefault>\s*<w:rPr>)([\s\S]*?)(<\/w:rPr>)/,
      (m, open, inner, close) => {
        if (inner.includes("w:lang")) return m;
        return open + inner + '<w:lang w:val="en-US" w:eastAsia="zh-CN"/>' + close;
      }
    )
    .replace(/<w:tab w:val="right"/g, '<w:tab w:val="right" w:leader="dot"')
    .replace(
      /(<w:style\b[^>]*w:styleId="TOC\d"[^>]*>[\s\S]*?<w:rFonts\s)([^/]*)(\/?>)/g,
      (m, before, attrs, end) => {
        attrs = attrs.replace(/w:ascii="[^"]*"/, 'w:ascii="Times New Roman"')
                     .replace(/w:hAnsi="[^"]*"/, 'w:hAnsi="Times New Roman"');
        return before + attrs + end;
      }
    );
  zip.file("word/styles.xml", patched);

  // 为 document.xml 中所有 <w:rPr> 补充 w:lang eastAsia
  let docXml = await zip.file("word/document.xml").async("string");

  // 将 fldSimple REF 转为复杂域，添加 CHARFORMAT 开关与上标格式
  const BS = String.fromCharCode(92); // single backslash
  docXml = docXml.replace(
    /<w:fldSimple\b[^>]*w:instr="REF (ref_\d+)[^"]*"[^>]*>\s*<w:r\b[^>]*>[\s\S]*?<\/w:r>\s*<\/w:fldSimple>/g,
    (m, bookmarkId) => {
      const num = bookmarkId.replace('ref_', '');
      const rpr = citationRunPropsXml;
      const instr = ` REF ${bookmarkId} ${BS}h ${BS}* CHARFORMAT `;
      return [
        `<w:r>${rpr}<w:fldChar w:fldCharType="begin"/></w:r>`,
        `<w:r>${rpr}<w:instrText xml:space="preserve">${instr}</w:instrText></w:r>`,
        `<w:r>${rpr}<w:fldChar w:fldCharType="separate"/></w:r>`,
        `<w:r>${rpr}<w:t>${num}</w:t></w:r>`,
        `<w:r>${rpr}<w:fldChar w:fldCharType="end"/></w:r>`,
      ].join('');
    }
  );
  docXml = docXml.replace(
    /<w:rPr>([^]*?)<\/w:rPr>/g,
    (m, inner) => {
      if (inner.includes("w:lang") || inner.includes("w:instrText") || inner.includes("w:fldChar")) return m;
      return '<w:rPr>' + inner + '<w:lang w:val="en-US" w:eastAsia="zh-CN"/></w:rPr>';
    }
  );
  zip.file("word/document.xml", docXml);

  return zip.generateAsync({ type: "nodebuffer", compression: "DEFLATE" });
}

// ===== 构建文档 =====
async function buildThesis(options) {
  const {
    cover = [],
    declaration = [],
    zhAbstract = [],
    enAbstract = [],
    toc = [],
    chapters = [],
    referencesSection = [],
    thanks = [],
    universityName = UNIVERSITY_NAME,
    thesisYear = THESIS_YEAR,
    outputPath = process.env.THESIS_OUTPUT || "f:/论文/论文初稿_通用模板.docx"
  } = options;

  const mainHeader = new Header({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text: `${universityName}${thesisYear}届本科毕业设计说明书（毕业论文）`, font: SONGTI, size: XIAO_WU })]
    })]
  });

  const mainFooter = new Footer({
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ children: [PageNumber.CURRENT], font: TNR, size: XIAO_WU })]
    })]
  });

  const doc = new Document({
    features: { updateFields: true },
    styles: {
      default: {
        document: { run: { font: SONGTI, size: XIAO_SI } }
      },
      paragraphStyles: [
        {
          id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: SAN_HAO, bold: true, font: HEITI },
          paragraph: { spacing: { before: 480, after: 480, line: 360 }, alignment: AlignmentType.CENTER, outlineLevel: 0 }
        },
        {
          id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: XIAO_SAN, bold: true, font: HEITI },
          paragraph: { spacing: { before: 360, after: 360, line: 360 }, alignment: AlignmentType.CENTER, outlineLevel: 1 }
        },
        {
          id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: SI_HAO, bold: true, font: HEITI },
          paragraph: { spacing: { before: 240, after: 240, line: 360 }, outlineLevel: 2 }
        },
        {
          id: "TOC1", name: "toc 1", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: SI_HAO, font: HEITI },
          paragraph: { spacing: { before: 0, after: 0, line: 300 }, rightTabStop: TabStopPosition.MAX }
        },
        {
          id: "TOC2", name: "toc 2", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: XIAO_SI, font: SONGTI },
          paragraph: { spacing: { before: 0, after: 0, line: 300 }, rightTabStop: TabStopPosition.MAX, indent: { left: 420 } }
        },
        {
          id: "TOC3", name: "toc 3", basedOn: "Normal", next: "Normal", quickFormat: true,
          run: { size: XIAO_SI, font: SONGTI },
          paragraph: { spacing: { before: 0, after: 0, line: 300 }, rightTabStop: TabStopPosition.MAX, indent: { left: 840 } }
        }
      ]
    },
    sections: [
      {
        properties: {
          page: {
            margin: { top: CM(2.5), bottom: CM(2.5), left: CM(2.5), right: CM(2.5) },
          }
        },
        children: [
          // 从摘要开始生成，已去除封面与原创性声明
          // ...cover,
          // ...declaration,
          ...zhAbstract,
          ...enAbstract,
          ...toc,
        ]
      },
      {
        properties: {
          page: {
            margin: { top: CM(2.5), bottom: CM(2.5), left: CM(2.5), right: CM(2.5) },
            pageNumbers: { start: 1 },
          }
        },
        headers: {
          default: mainHeader
        },
        footers: {
          default: mainFooter
        },
        children: [
          ...chapters,
          ...referencesSection,
          ...thanks,
        ]
      }
    ]
  });

  const buffer = await Packer.toBuffer(doc);
  const final = await patchLang(buffer);
  fs.writeFileSync(outputPath, final);
  console.log("论文生成完成！文件路径: " + outputPath);
}

// 导出模块供外部调用
module.exports = {
  UNIVERSITY_NAME, COLLEGE_NAME, MAJOR_NAME, GRADE_YEAR, STUDENT_ID, STUDENT_NAME, ADVISOR_NAME, THESIS_YEAR, THESIS_MONTH,
  THESIS_TITLE_CN, THESIS_TITLE_EN, ZH_ABSTRACT_TEXT, EN_ABSTRACT_TEXT, ZH_KEYWORDS, EN_KEYWORDS,
  CM, PT, SONGTI, HEITI, KAITI, TNR, MIXED_FONT,
  p, pRuns, h1, h2, h3, sectionLead, paragraphTitle, bodySection,
  equationLine, equationDisplay, m, abstractTitle, abstractHeading, abstractParagraph, keywordParagraph,
  blankLine, empty, pb, coverLine, figure, referenceParagraph, tableText, tableCaption, tableCell, placeholderTable,
  buildThesis, patchLang
};

// 如果直接运行该脚本
if (require.main === module) {
  const keepAlive = setInterval(() => {}, 1000);
  buildThesis({
    // 从摘要开始生成，已去除封面与原创性声明
    // cover, declaration,
    zhAbstract, enAbstract, toc,
    chapters: ch1,
    referencesSection, thanks
  }).catch(console.error).finally(() => clearInterval(keepAlive));
}
