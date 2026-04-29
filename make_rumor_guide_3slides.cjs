const path = require('path');
const pptxgen = require('pptxgenjs');

const pptx = new pptxgen();
pptx.layout = 'LAYOUT_WIDE';
pptx.author = 'OpenAI Codex';
pptx.company = 'OpenAI';
pptx.subject = 'Rumor Defense Game Guide';
pptx.title = '谣言辟谣攻略（三页版）';
pptx.lang = 'zh-CN';
pptx.theme = {
  headFontFace: 'Microsoft YaHei',
  bodyFontFace: 'Microsoft YaHei',
  lang: 'zh-CN',
};

const COLORS = {
  bg: 'F5F4F0',
  paper: 'FCFBF8',
  ink: '162033',
  muted: '667085',
  line: 'D7DCE3',
  accent: '284B63',
  green: '2A7B5F',
  yellow: 'C58A18',
  red: 'B84B4B',
  white: 'FFFFFF',
};

const IMG = {
  intro: path.join(__dirname, '介绍.png'),
  cards: path.join(__dirname, '卡牌页.png'),
  error: path.join(__dirname, '错误解析页.png'),
};

function addPageChrome(slide, pageNo, section, en) {
  slide.background = { color: COLORS.bg };

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.38,
    y: 0.38,
    w: 12.54,
    h: 6.74,
    line: { color: COLORS.line, width: 1 },
    fill: { color: COLORS.paper },
  });

  slide.addText(section, {
    x: 0.8,
    y: 0.55,
    w: 3,
    h: 0.2,
    fontFace: 'Aptos',
    fontSize: 10,
    color: COLORS.muted,
    charSpacing: 1.2,
    margin: 0,
  });

  slide.addText(en, {
    x: 10.0,
    y: 0.55,
    w: 1.8,
    h: 0.2,
    fontFace: 'Aptos',
    fontSize: 9,
    color: COLORS.muted,
    align: 'right',
    margin: 0,
  });

  slide.addText(String(pageNo).padStart(2, '0'), {
    x: 11.95,
    y: 0.48,
    w: 0.45,
    h: 0.28,
    fontFace: 'Aptos',
    fontSize: 16,
    bold: true,
    color: COLORS.accent,
    align: 'right',
    margin: 0,
  });
}

function addTitle(slide, title, subtitle) {
  slide.addText(title, {
    x: 0.8,
    y: 0.95,
    w: 6.0,
    h: 0.8,
    fontSize: 28,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addText(subtitle, {
    x: 0.82,
    y: 1.78,
    w: 6.0,
    h: 0.24,
    fontFace: 'Aptos',
    fontSize: 10,
    color: COLORS.muted,
    charSpacing: 0.7,
    margin: 0,
  });
}

function addMetric(slide, x, y, title, body) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w: 2.08,
    h: 1.18,
    rectRadius: 0.08,
    line: { color: COLORS.line, width: 1 },
    fill: { color: COLORS.white },
  });
  slide.addText(title, {
    x: x + 0.18,
    y: y + 0.16,
    w: 1.72,
    h: 0.2,
    fontSize: 10,
    color: COLORS.accent,
    bold: true,
    margin: 0,
  });
  slide.addText(body, {
    x: x + 0.18,
    y: y + 0.48,
    w: 1.68,
    h: 0.52,
    fontSize: 14,
    color: COLORS.ink,
    margin: 0,
  });
}

function addCard(slide, x, y, label, desc, accent, fill) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w: 1.8,
    h: 1.42,
    rectRadius: 0.08,
    line: { color: COLORS.line, width: 1 },
    fill: { color: fill },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: x + 0.16,
    y: y + 0.16,
    w: 0.18,
    h: 0.86,
    line: { color: accent, width: 0 },
    fill: { color: accent },
  });
  slide.addText(label, {
    x: x + 0.48,
    y: y + 0.2,
    w: 1.06,
    h: 0.24,
    fontSize: 15,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addText(desc, {
    x: x + 0.48,
    y: y + 0.56,
    w: 1.1,
    h: 0.56,
    fontSize: 10,
    color: COLORS.muted,
    margin: 0,
  });
}

function addRule(slide, x, y, n, title, desc) {
  slide.addText(String(n).padStart(2, '0'), {
    x,
    y,
    w: 0.42,
    h: 0.24,
    fontFace: 'Aptos',
    fontSize: 16,
    bold: true,
    color: COLORS.accent,
    margin: 0,
  });
  slide.addText(title, {
    x: x + 0.54,
    y,
    w: 2.4,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addText(desc, {
    x: x + 0.54,
    y: y + 0.22,
    w: 2.42,
    h: 0.36,
    fontSize: 9.5,
    color: COLORS.muted,
    margin: 0,
  });
}

function addMiniChip(slide, x, y, text, fill, color, w = 0.96) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h: 0.28,
    rectRadius: 0.08,
    line: { color: fill, width: 1 },
    fill: { color: fill },
  });
  slide.addText(text, {
    x: x + 0.05,
    y: y + 0.06,
    w: w - 0.1,
    h: 0.14,
    fontSize: 8,
    bold: true,
    color,
    align: 'center',
    margin: 0,
  });
}

function addPill(slide, x, y, text, fill, color, w = 1.08) {
  slide.addShape(pptx.ShapeType.roundRect, {
    x,
    y,
    w,
    h: 0.34,
    rectRadius: 0.1,
    line: { color: fill, width: 1 },
    fill: { color: fill },
  });
  slide.addText(text, {
    x: x + 0.08,
    y: y + 0.08,
    w: w - 0.16,
    h: 0.16,
    fontSize: 9,
    bold: true,
    color,
    align: 'center',
    margin: 0,
  });
}

function addLabelValue(slide, x, y, label, value) {
  slide.addText(label, {
    x,
    y,
    w: 1.0,
    h: 0.16,
    fontFace: 'Aptos',
    fontSize: 8.5,
    color: COLORS.muted,
    margin: 0,
  });
  slide.addText(value, {
    x,
    y: y + 0.16,
    w: 1.8,
    h: 0.18,
    fontSize: 10.5,
    color: COLORS.ink,
    bold: true,
    margin: 0,
  });
}

function coverSlide() {
  const slide = pptx.addSlide();
  addPageChrome(slide, 1, '攻略总览', 'OVERVIEW');

  slide.addShape(pptx.ShapeType.rect, {
    x: 0.8,
    y: 1.04,
    w: 0.22,
    h: 1.15,
    line: { color: COLORS.accent, width: 0 },
    fill: { color: COLORS.accent },
  });

  slide.addText('《相亲相爱一家人：\n辟谣防卫战》', {
    x: 1.18,
    y: 1.02,
    w: 4.6,
    h: 1.45,
    fontSize: 25,
    bold: true,
    color: COLORS.ink,
    margin: 0,
    valign: 'mid',
  });
  slide.addText('三页通关攻略', {
    x: 1.2,
    y: 2.36,
    w: 2.2,
    h: 0.24,
    fontSize: 12,
    color: COLORS.accent,
    bold: true,
    margin: 0,
  });
  slide.addText('用更短的阅读成本，快速理解胜利条件、选牌逻辑与误导性辟谣识别方式。', {
    x: 1.2,
    y: 2.74,
    w: 4.3,
    h: 0.58,
    fontSize: 13,
    color: COLORS.muted,
    margin: 0,
  });

  addMetric(slide, 1.18, 3.52, '胜利条件', '谣言强度归零\n家庭和谐度不归零');
  addMetric(slide, 3.4, 3.52, '方案 B', '选对只回血\n选错仅扣 10HP');
  addMetric(slide, 1.18, 4.88, '通关节奏', '优先绿标\n5-6 回合理想');
  addMetric(slide, 3.4, 4.88, '核心原则', '先判断标签\n再看文案措辞');

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 6.28,
    y: 1.02,
    w: 5.98,
    h: 5.65,
    rectRadius: 0.06,
    line: { color: COLORS.line, width: 1 },
    fill: { color: 'E9EEF2' },
    shadow: { type: 'outer', color: '8A94A6', blur: 2, angle: 45, offset: 2, opacity: 0.15 },
  });
  slide.addImage({
    path: IMG.intro,
    x: 6.46,
    y: 1.18,
    w: 5.62,
    h: 4.8,
    sizing: { type: 'cover', x: 6.46, y: 1.18, w: 5.62, h: 4.8 },
  });
  slide.addShape(pptx.ShapeType.rect, {
    x: 6.46,
    y: 5.62,
    w: 5.62,
    h: 0.36,
    line: { color: COLORS.paper, width: 0 },
    fill: { color: COLORS.paper, transparency: 10 },
  });
  addLabelValue(slide, 6.7, 6.0, '版本', 'v2.1 / 方案 B');
  addLabelValue(slide, 8.42, 6.0, '玩法关键词', '判断 / 共情 / 纠偏');
  addLabelValue(slide, 10.2, 6.0, '适用场景', '展览、演示、教学');
}

function gameplaySlide() {
  const slide = pptx.addSlide();
  addPageChrome(slide, 2, '机制与策略', 'PLAYBOOK');
  addTitle(slide, '核心玩法与出牌优先级', 'CARD LOGIC / LABEL JUDGMENT / QUICK WIN FLOW');

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.8,
    y: 2.32,
    w: 4.46,
    h: 3.98,
    rectRadius: 0.06,
    line: { color: COLORS.line, width: 1 },
    fill: { color: COLORS.white },
  });
  slide.addText('卡牌类型', {
    x: 1.05,
    y: 2.58,
    w: 1.1,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.accent,
    margin: 0,
  });
  addCard(slide, 1.05, 2.98, '论据卡', '蓝色卡\n降谣言 + 小幅回血', '4D7EA8', 'EEF5F8');
  addCard(slide, 2.97, 2.98, '情绪卡', '黄色卡\n降谣言 + 较大回血', 'C58A18', 'FBF4E8');
  addCard(slide, 1.05, 4.58, '特殊牌', '绿色卡\n效果因牌而异', '2A7B5F', 'EEF7F3');
  slide.addText('操作原则', {
    x: 2.98,
    y: 4.62,
    w: 1.2,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.accent,
    margin: 0,
  });
  addRule(slide, 2.98, 5.0, 1, '先看标签颜色', '绿色优先，黄色谨慎，红色不选。');
  addRule(slide, 2.98, 5.62, 2, '再看表述方式', '模糊转折和暗示性语言通常是陷阱。');

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 5.52,
    y: 2.32,
    w: 6.72,
    h: 3.98,
    rectRadius: 0.06,
    line: { color: COLORS.line, width: 1 },
    fill: { color: 'EAF0F4' },
  });
  slide.addImage({
    path: IMG.cards,
    x: 5.68,
    y: 2.48,
    w: 6.4,
    h: 3.2,
    sizing: { type: 'cover', x: 5.68, y: 2.48, w: 6.4, h: 3.2 },
  });
  addMiniChip(slide, 5.88, 5.78, '正确辟谣', 'DDEFE4', COLORS.green);
  addMiniChip(slide, 6.94, 5.78, '误导性辟谣', 'F4E9C8', COLORS.yellow, 1.14);
  addMiniChip(slide, 8.16, 5.78, '明显错误', 'F4DCDC', COLORS.red);
  slide.addText('推荐流程：先锁定绿色标签的正确辟谣，再结合角色人格与当前和谐度选择论据牌或情绪牌。', {
    x: 9.28,
    y: 5.79,
    w: 2.38,
    h: 0.36,
    fontSize: 9.2,
    color: COLORS.muted,
    margin: 0,
  });
}

function cheatSheetSlide() {
  const slide = pptx.addSlide();
  addPageChrome(slide, 3, '速查与识别', 'CHEAT SHEET');
  addTitle(slide, '常见谣言、识别陷阱与人物应对', 'RUMOR TYPES / MISLEADING COPY / PERSONA RESPONSE');

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 0.8,
    y: 2.24,
    w: 3.2,
    h: 4.2,
    rectRadius: 0.06,
    line: { color: COLORS.line, width: 1 },
    fill: { color: COLORS.white },
  });
  slide.addText('高频谣言类型', {
    x: 1.04,
    y: 2.5,
    w: 1.4,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.accent,
    margin: 0,
  });
  addPill(slide, 1.04, 2.94, '健康养生', 'EAF1F5', COLORS.accent);
  addPill(slide, 2.2, 2.94, '食品安全', 'F5EFE3', '8A5A00');
  addPill(slide, 1.04, 3.36, '社会民生', 'EFEDE7', COLORS.ink);
  addPill(slide, 2.2, 3.36, '阴谋论', 'F7E7E7', COLORS.red);
  addPill(slide, 1.62, 3.78, '科技类', 'E9F0EC', COLORS.green);
  slide.addText('从“微波炉致癌”“隔夜菜致癌”到“5G致病”“地震云预警”，共同特点都是借伪科学、权威误读或情绪传播制造可信错觉。', {
    x: 1.04,
    y: 4.34,
    w: 2.54,
    h: 1.06,
    fontSize: 10.5,
    color: COLORS.muted,
    margin: 0,
  });
  slide.addText('优先思路', {
    x: 1.04,
    y: 5.62,
    w: 1.0,
    h: 0.18,
    fontSize: 10,
    bold: true,
    color: COLORS.ink,
    margin: 0,
  });
  slide.addText([
    { text: '看是否有明确证据与标准', options: { bullet: true, breakLine: true } },
    { text: '看是否在偷换概念或扩大风险', options: { bullet: true, breakLine: true } },
    { text: '看是否用“可能”“也许”制造恐慌', options: { bullet: true } },
  ], {
    x: 1.04,
    y: 5.88,
    w: 2.5,
    h: 0.72,
    fontSize: 9.2,
    color: COLORS.muted,
    margin: 0,
  });

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 4.18,
    y: 2.24,
    w: 4.0,
    h: 4.2,
    rectRadius: 0.06,
    line: { color: COLORS.line, width: 1 },
    fill: { color: COLORS.white },
  });
  slide.addText('误导性辟谣的 4 个典型陷阱', {
    x: 4.42,
    y: 2.5,
    w: 2.7,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.accent,
    margin: 0,
  });
  addRule(slide, 4.42, 2.96, 1, '“虽然……但……”', '表面承认事实，实则保留恐慌暗示。');
  addRule(slide, 4.42, 3.7, 2, '“没有证据，也不能否定”', '借不确定性维持谣言生命力。');
  addRule(slide, 4.42, 4.44, 3, '以偏概全', '把局部问题扩大为普遍风险。');
  addRule(slide, 4.42, 5.18, 4, '伪专业背书', '用“长期累积”“内部人士”制造假权威。');

  slide.addShape(pptx.ShapeType.roundRect, {
    x: 8.36,
    y: 2.24,
    w: 3.88,
    h: 4.2,
    rectRadius: 0.06,
    line: { color: COLORS.line, width: 1 },
    fill: { color: 'EAF0F4' },
  });
  slide.addImage({
    path: IMG.error,
    x: 8.52,
    y: 2.4,
    w: 3.56,
    h: 1.88,
    sizing: { type: 'cover', x: 8.52, y: 2.4, w: 3.56, h: 1.88 },
  });
  slide.addText('人物应对速记', {
    x: 8.6,
    y: 4.56,
    w: 1.6,
    h: 0.2,
    fontSize: 12,
    bold: true,
    color: COLORS.accent,
    margin: 0,
  });
  slide.addText([
    { text: 'ESTJ 固执大爷：优先论据卡，看权威和数据。', options: { bullet: true, breakLine: true } },
    { text: 'ESFJ 焦虑老妈：先共情，再用情绪卡慢慢纠偏。', options: { bullet: true, breakLine: true } },
    { text: 'INTJ 阴谋论者：最难说服，以减少伤害为主。', options: { bullet: true, breakLine: true } },
    { text: 'ESFP 爱孙奶奶：情绪牌更有效，亲情线索最好用。', options: { bullet: true } },
  ], {
    x: 8.56,
    y: 4.92,
    w: 3.1,
    h: 1.18,
    fontSize: 9.2,
    color: COLORS.muted,
    margin: 0,
  });
}

coverSlide();
gameplaySlide();
cheatSheetSlide();

const outPath = path.join(__dirname, '谣言辟谣攻略_三页版.pptx');

pptx.writeFile({ fileName: outPath }).then(() => {
  console.log(outPath);
}).catch((error) => {
  console.error(error);
  process.exit(1);
});
