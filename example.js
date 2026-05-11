const fs = require("fs");
const path = require("path");
const { TableOfContents, AlignmentType, TextRun, MathSuperScript, Paragraph, ImageRun } = require("docx");

// 引入通用模板 js 中的工具函数与常量
const {
  p, h1, h2, h3, pb, empty, blankLine,
  placeholderTable, equationLine, m,
  abstractTitle, abstractHeading, abstractParagraph, keywordParagraph,
  referenceParagraph,
  buildThesis,
  CM, HEITI, ER_HAO, YI_HAO, SAN_HAO, SONGTI, XIAO_SI, TNR,
  // THESIS_YEAR, THESIS_MONTH
} = require("./generate_thesis");

// ===== 1. 定义论文基本信息 =====
const myTitleCN = "基于Node.js的自动化论文排版系统设计与实现";
const myTitleEN = [
  "Design and Implementation of Automated Thesis Typesetting",
  "System Based on Node.js"
];
const myUniversity = "某某大学";
const kitaKitaGifPath = path.join(__dirname, "图", "喜多喜多.gif");

function kitaKitaFigure() {
  return [
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 240, after: 120, line: 360 },
      children: [
        new ImageRun({
          type: "gif",
          data: fs.readFileSync(kitaKitaGifPath),
          transformation: { width: 262, height: 258 }
        })
      ]
    }),
    new Paragraph({
      alignment: AlignmentType.CENTER,
      spacing: { before: 0, after: 240, line: 360 },
      children: [new TextRun({ text: "图1.1 喜多喜多示例图", font: SONGTI, size: XIAO_SI })]
    }),
  ];
}

// ===== 3. 构造中英文摘要 =====
const myZhAbstract = [
  abstractTitle(myTitleCN),
  blankLine(ER_HAO),
  abstractHeading("摘　要"),
  abstractParagraph("随着信息技术的不断发展，文档自动化排版技术在学术界和工业界都发挥着越来越重要的作用。本文基于Node.js环境与docx第三方库，设计并实现了一套可复用的学位论文格式生成系统。本系统抽象出了论文排版中的标题、段落、图表、公式以及交叉引用等常见组件，使得使用者只需关注文本内容即可生成符合严格格式规范的Word文档。"),
  empty(),
  keywordParagraph("关键词", "Node.js 自动化排版 docx 模板生成"),
  empty(), empty()
];

const myEnAbstract = [
  abstractTitle(myTitleEN[0], { font: TNR }),
  abstractTitle(myTitleEN[1], { font: TNR }),
  blankLine(ER_HAO, TNR),
  abstractHeading("ABSTRACT", { font: TNR, bold: true }),
  abstractParagraph("With the continuous development of information technology, document automation typesetting plays an increasingly important role... This paper designs and implements a reusable formatting system based on Node.js and the docx library.", { font: TNR }),
  empty(),
  keywordParagraph("Key Words", "Node.js; Automated Typesetting; docx; Template", { labelFont: TNR, labelBold: true, valueFont: TNR, indent: null, colon: ":" }),
];

// ===== 4. 构造目录 =====
const myToc = [
  pb(),
  p("目    录", { alignment: AlignmentType.CENTER, font: HEITI, size: SAN_HAO, bold: true, spacing: { before: 480, after: 480, line: 300 } }),
  blankLine(XIAO_SI),
  new TableOfContents("", { hyperlink: true, headingStyleRange: "1-3" }),
];

// ===== 5. 构造正文（含表格和公式） =====
const myCh1 = [
  pb(),
  h1("第一章  绪论"),
  h2("1.1 研究背景"),
  p("毕业论文排版往往消耗学生大量的时间，且经常因为格式不符合标准而被退回修改。传统的手动排版方式不仅效率低下，而且容易出错。"),
  h3("1.1.1 自动化排版的现状"),
  p("目前市面上已经存在部分基于LaTeX的排版工具，但在部分高校中，仍强制要求提交Microsoft Word格式的文档。因此，基于Node.js直接生成Word文档成为了一种可行的替代方案。"),

  h2("1.2 相关技术与公式示例"),
  p("在科学研究中，经常需要书写复杂的数学公式。基于本系统提供的辅助函数，可以轻松实现带编号的公式行："),
  // 公式示例： E = mc^2
  equationLine([m("E"), m("="), m("m"), m("c"), new MathSuperScript({ children: [m("2")] })], "1-1"),

  h2("1.3 图片插入示例"),
  p("本节演示如何在论文正文中插入位于图文件夹下的GIF图片。示例图片见图1.1。"),
  ...kitaKitaFigure(),

  h2("1.4 系统需求与表格示例"),
  p("本系统支持动态生成标准的三线表。以下为简化的系统需求分析列表，见表1.1。"),
  // 表格示例
  ...placeholderTable("1.1", "系统功能需求表", ["需求编号", "需求名称", "优先级"], [
    ["REQ-01", "生成标准段落", "高"],
    ["REQ-02", "渲染数学公式", "中"],
    ["REQ-03", "自动化生成三线表", "高"]
  ], { columnWidths: [CM(4), CM(6), CM(4)] }),
];

// ===== 6. 构造参考文献 =====
const myReferences = [
  pb(),
  h1("参考文献"),
  empty(),
  referenceParagraph("[1] 张三. Node.js高级编程[M]. 北京: 电子工业出版社, 2021: 45-50."),
  referenceParagraph("[2] 李四. 自动化文档排版技术研究[J]. 计算机应用, 2022, 42(3): 123-128.")
];

// ===== 7. 构造致谢 =====
const myThanks = [
  pb(),
  h1("致    谢"),
  p("大学四年时光转瞬即逝，在此我由衷地感谢我的指导老师、我的家人以及所有关心和帮助过我的人。"),
];

// ===== 8. 执行生成 =====
async function generateExample() {
  const outputPath = "./example_output.docx";

  console.log("正在构建示例论文...");

  await buildThesis({
    universityName: myUniversity, // 覆盖默认的常量
    thesisYear: "2026",
    // 从摘要开始生成，已去除封面与原创性声明
    // cover: myCover,
    // declaration: [],
    zhAbstract: myZhAbstract,
    enAbstract: myEnAbstract,
    toc: myToc,
    chapters: myCh1, // 传入所有章节的合并数组
    referencesSection: myReferences,
    thanks: myThanks,
    outputPath: outputPath
  });

  console.log("示例生成结束！");
}

generateExample().catch(console.error);
