const pptxgen = require('pptxgenjs');
const path = require('path');

const pres = new pptxgen();
pres.layout = 'LAYOUT_16x9';

// 复古胶片相机主题配色：黑机身 + 银灰 + 暖橙点缀
const theme = {
  primary: "1a1a1a",    // 黑机身
  secondary: "4a4a4a",  // 深灰
  accent: "c8702a",     // 暖橙点缀
  light: "d4d0c8",      // 暖灰/银
  bg: "f5f0e8"          // 米白背景
};

// 图片路径
const imgDir = path.join(__dirname, 'imgs');
const imgPath = (name) => path.join(imgDir, `${name}.jpeg`);

// 全局字体设置
const CN_FONT = "Microsoft YaHei";
const EN_FONT = "Arial";

// 页码追踪
let pageNum = 0;

// 通用页码徽章函数
function addPageBadge(slide, pres, theme) {
  pageNum++;
  if (pageNum <= 1) return; // 封面不加页码
  slide.addShape(pres.shapes.OVAL, {
    x: 9.3, y: 5.1, w: 0.4, h: 0.4,
    fill: { color: theme.accent }
  });
  slide.addText(String(pageNum), {
    x: 9.3, y: 5.1, w: 0.4, h: 0.4,
    fontSize: 12, fontFace: EN_FONT,
    color: "FFFFFF", bold: true,
    align: "center", valign: "middle"
  });
}

// ===================== P1 封面 =====================
function createSlide01(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.primary };

  // 右侧图片区域 (60%)
  try {
    slide.addImage({
      path: imgPath('front-3d'),
      x: 4.0, y: 0, w: 6.0, h: 5.625,
      sizing: { type: 'cover', w: 6.0, h: 5.625 }
    });
  } catch (e) {
    // 图片失败用色块替代
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 4.0, y: 0, w: 6.0, h: 5.625,
      fill: { color: theme.secondary }
    });
  }

  // 左侧深色遮罩
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0, y: 0, w: 4.2, h: 5.625,
    fill: { color: theme.primary }
  });

  // 标题
  slide.addText("Bromo", {
    x: 0.5, y: 1.8, w: 3.5, h: 0.9,
    fontSize: 60, fontFace: EN_FONT,
    color: theme.light, bold: true,
    align: "left", valign: "bottom"
  });

  slide.addText("复古数码相机", {
    x: 0.5, y: 2.7, w: 3.5, h: 0.5,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.light, bold: false,
    align: "left", valign: "top"
  });

  // 副标题
  slide.addText("以经典胶片机的质感与节奏，让人重新认真拍照", {
    x: 0.5, y: 3.3, w: 3.5, h: 0.6,
    fontSize: 14, fontFace: CN_FONT,
    color: "999999", bold: false,
    align: "left", valign: "top"
  });

  // 关键信息条（胶囊式）
  const capsules = ["135 经典旁轴", "数字胶卷", "氙气闪光灯", "¥1299"];
  let capX = 0.5;
  capsules.forEach((text, i) => {
    const w = text.length <= 4 ? 1.1 : (text.length <= 6 ? 1.5 : 2.0);
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: capX, y: 4.2, w: w, h: 0.35,
      fill: { color: theme.accent },
      rectRadius: 0.175
    });
    slide.addText(text, {
      x: capX, y: 4.2, w: w, h: 0.35,
      fontSize: 11, fontFace: CN_FONT,
      color: "FFFFFF", bold: true,
      align: "center", valign: "middle"
    });
    capX += w + 0.15;
  });

  // 页脚元信息
  slide.addText("产品定义评审 · 内部资料 · 产品团队", {
    x: 0.5, y: 5.0, w: 3.5, h: 0.3,
    fontSize: 10, fontFace: CN_FONT,
    color: "666666", bold: false,
    align: "left", valign: "middle"
  });

  pageNum++;
}

// ===================== P2 目录 =====================
function createSlide02(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  // 左侧大标题 "Bromo"
  slide.addText("Bromo", {
    x: 0.6, y: 1.5, w: 3.0, h: 1.0,
    fontSize: 72, fontFace: EN_FONT,
    color: "d4d0c8", bold: true,
    align: "left", valign: "middle"
  });

  // 右侧目录标题
  slide.addText("目录", {
    x: 4.5, y: 0.6, w: 4.5, h: 0.6,
    fontSize: 36, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  // 目录条目
  const items = [
    { num: "01", title: "核心定义与目标用户", page: "P3–P4" },
    { num: "02", title: "四大核心价值", page: "P5" },
    { num: "03", title: "与 AiKen 的对比", page: "P6" },
    { num: "04", title: "硬件定义 · 质感来自每一个结构件", page: "P7–P14" },
    { num: "05", title: "软件沿用说明", page: "P15" },
    { num: "06", title: "价值主张与立项请求", page: "P16–P18" }
  ];

  items.forEach((item, i) => {
    const y = 1.4 + i * 0.65;
    // 编号
    slide.addText(item.num, {
      x: 4.5, y: y, w: 0.6, h: 0.4,
      fontSize: 22, fontFace: EN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "middle"
    });
    // 标题
    slide.addText(item.title, {
      x: 5.2, y: y, w: 3.5, h: 0.4,
      fontSize: 16, fontFace: CN_FONT,
      color: theme.primary, bold: false,
      align: "left", valign: "middle"
    });
    // 页码
    slide.addText(item.page, {
      x: 8.5, y: y, w: 1.0, h: 0.4,
      fontSize: 12, fontFace: EN_FONT,
      color: "999999", bold: false,
      align: "right", valign: "middle"
    });
  });

  // 角注
  slide.addText("附：研发关键参数速查（P17）", {
    x: 4.5, y: 5.0, w: 4.0, h: 0.3,
    fontSize: 10, fontFace: CN_FONT,
    color: "999999", bold: false,
    align: "left", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P3 核心定义 =====================
function createSlide03(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("Bromo 是什么", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.6,
    fontSize: 36, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  // 大字定义
  slide.addText("「一台以经典胶片机的质感与节奏，让人重新认真拍照的复古数码相机。」", {
    x: 0.5, y: 1.2, w: 9.0, h: 1.2,
    fontSize: 22, fontFace: CN_FONT,
    color: theme.secondary, bold: false,
    align: "left", valign: "top"
  });

  // 四个定位关键词（2x2网格）
  const keywords = [
    { title: "135 经典旁轴", desc: "金属机身，巴掌大小依旧便携" },
    { title: "数字胶卷", desc: "36 张一卷，多卷连续，事后统一冲扫" },
    { title: "弹出式氙气闪光灯", desc: "独立供电，弹起即蓄电" },
    { title: "¥1299", desc: "AiKen（¥699）的质感升级款" }
  ];

  keywords.forEach((kw, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 4.6;
    const y = 2.6 + row * 1.3;

    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: y, w: 4.4, h: 1.1,
      fill: { color: "FFFFFF" },
      rectRadius: 0.08
    });

    slide.addText(kw.title, {
      x: x + 0.15, y: y + 0.1, w: 4.1, h: 0.35,
      fontSize: 16, fontFace: CN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "middle"
    });

    slide.addText(kw.desc, {
      x: x + 0.15, y: y + 0.45, w: 4.1, h: 0.5,
      fontSize: 12, fontFace: CN_FONT,
      color: theme.secondary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 底部收束句
  slide.addText("AiKen 定义了\"便捷的复古数码\"，Bromo 定义\"有仪式感的复古数码\"。", {
    x: 0.5, y: 5.0, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P4 目标用户与场景 =====================
function createSlide04(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("谁会买 Bromo", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.6,
    fontSize: 36, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  // 场景引入（斜体）
  slide.addText("它不是参数最强的相机，但会是口袋里最有性格的一台。", {
    x: 0.5, y: 1.0, w: 9.0, h: 0.4,
    fontSize: 14, fontFace: CN_FONT,
    color: "888888", bold: false, italic: true,
    align: "left", valign: "middle"
  });

  // 三栏卡片
  const cards = [
    { title: "胶片爱好者", tag: "存量换机", line1: "懂胶片的好，也懂胶片的贵与慢。", line2: "Bromo 给他们熟悉的全部仪式，免去每卷冲扫的费用与两周的等待——胶片的灵魂，数码的身体。" },
    { title: "复古数码玩家", tag: "增量主力", line1: "要颜值、要个性、要出片能发。", line2: "Bromo 是\"能天天带出门的胶片机\"：1 秒开机、0 基础上手、多张拼贴直出即可分享。" },
    { title: "礼品场景", tag: "情绪溢价", line1: "送礼的人要\"送得出手\"，收礼的人要\"拿得出手\"。", line2: "¥1299 位于黄金礼品价位带：开箱有质感、上手零门槛、每一卷都有故事。" }
  ];

  cards.forEach((card, i) => {
    const x = 0.5 + i * 3.1;
    const y = 1.6;

    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: y, w: 2.9, h: 2.8,
      fill: { color: "FFFFFF" },
      rectRadius: 0.08
    });

    slide.addText(card.title, {
      x: x + 0.15, y: y + 0.15, w: 2.6, h: 0.35,
      fontSize: 16, fontFace: CN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "middle"
    });

    slide.addText(card.tag, {
      x: x + 0.15, y: y + 0.5, w: 2.6, h: 0.25,
      fontSize: 11, fontFace: CN_FONT,
      color: "999999", bold: false,
      align: "left", valign: "middle"
    });

    slide.addText(card.line1, {
      x: x + 0.15, y: y + 0.85, w: 2.6, h: 0.5,
      fontSize: 12, fontFace: CN_FONT,
      color: theme.primary, bold: true,
      align: "left", valign: "top"
    });

    slide.addText(card.line2, {
      x: x + 0.15, y: y + 1.4, w: 2.6, h: 1.2,
      fontSize: 11, fontFace: CN_FONT,
      color: theme.secondary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 底部高频场景条
  const scenes = ["周末 citywalk", "旅行随拍", "朋友聚会", "纪念日记录"];
  let sceneX = 1.5;
  scenes.forEach((scene) => {
    const w = scene.length <= 6 ? 1.3 : 1.8;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: sceneX, y: 4.7, w: w, h: 0.3,
      fill: { color: theme.light },
      rectRadius: 0.15
    });
    slide.addText(scene, {
      x: sceneX, y: 4.7, w: w, h: 0.3,
      fontSize: 10, fontFace: CN_FONT,
      color: theme.primary, bold: false,
      align: "center", valign: "middle"
    });
    sceneX += w + 0.2;
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P5 四大核心价值 =====================
function createSlide05(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("四大核心价值", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.6,
    fontSize: 36, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  const values = [
    { num: "01", title: "经典质感", desc: "经典旁轴形态、金属机身，巴掌大小依旧便携。拨杆过片、弹出式氙气闪光灯、可装数字胶卷——每一步操作都有机械的分量与声响。" },
    { num: "02", title: "完整的胶片体验", desc: "36 张一卷，多卷连续拍摄、事后统一冲扫。保留胶卷时代的完整仪式：过片、光学取景、计数、等待揭晓。" },
    { num: "03", title: "大底直出氛围感", desc: "更大传感器、自动对焦、氙气补光，底子更扎实、对焦更可靠。承袭 AiKen 全部风格滤镜，颗粒、偏色、漏光的\"不完美之美\"依然被鼓励。" },
    { num: "04", title: "数码内核的便捷", desc: "1 秒开机拿起就拍：半按对焦、自动曝光、自动 EV 补偿，0 基础上手。数字胶卷 Type-C 直插手机冲扫，多张拼贴直出即可分享。" }
  ];

  values.forEach((v, i) => {
    const col = i % 2;
    const row = Math.floor(i / 2);
    const x = 0.5 + col * 4.6;
    const y = 1.3 + row * 2.0;

    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: y, w: 4.4, h: 1.7,
      fill: { color: "FFFFFF" },
      rectRadius: 0.08
    });

    slide.addText(v.num, {
      x: x + 0.15, y: y + 0.1, w: 0.6, h: 0.35,
      fontSize: 28, fontFace: EN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "middle"
    });

    slide.addText(v.title, {
      x: x + 0.8, y: y + 0.1, w: 3.4, h: 0.35,
      fontSize: 16, fontFace: CN_FONT,
      color: theme.primary, bold: true,
      align: "left", valign: "middle"
    });

    slide.addText(v.desc, {
      x: x + 0.15, y: y + 0.55, w: 4.1, h: 1.0,
      fontSize: 11, fontFace: CN_FONT,
      color: theme.secondary, bold: false,
      align: "left", valign: "top"
    });
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P6 与 AiKen 对比表格 =====================
function createSlide06(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("与 AiKen 的对比——从品类入门到质感升级", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.6,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  // 表格数据
  const rows = [
    ["定位", "便捷、品类入门", "质感升级"],
    ["价位", "¥699", "¥1299"],
    ["形态", "110 轻便随行", "135 经典旁轴"],
    ["影像", "1/3 英寸，等效 35mm，泛焦", "更大底、等效 35mm、AF"],
    ["存储", "内置存储，24 张 JPEG", "支持 TF 卡的数字胶卷仓"],
    ["连续拍摄", "满卷后必须传输", "多卷连续，统一显影"],
    ["闪光灯", "不支持", "弹出式氙气闪光灯，电源独立"]
  ];

  const startY = 1.2;
  const rowH = 0.55;
  const colW = [1.8, 3.5, 3.5];
  const startX = [0.5, 2.5, 6.2];

  // 表头
  const headers = ["", "AiKen", "Bromo"];
  headers.forEach((h, i) => {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: startX[i], y: startY, w: colW[i], h: rowH,
      fill: { color: theme.primary }
    });
    slide.addText(h, {
      x: startX[i], y: startY, w: colW[i], h: rowH,
      fontSize: 12, fontFace: CN_FONT,
      color: "FFFFFF", bold: true,
      align: i === 0 ? "left" : "center", valign: "middle"
    });
  });

  // 数据行
  rows.forEach((row, rIdx) => {
    const y = startY + (rIdx + 1) * rowH;
    const bgColor = rIdx % 2 === 0 ? "FFFFFF" : "f8f6f2";

    row.forEach((cell, cIdx) => {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: startX[cIdx], y: y, w: colW[cIdx], h: rowH,
        fill: { color: bgColor }
      });

      const isBromoCol = cIdx === 2;
      slide.addText(cell, {
        x: startX[cIdx] + 0.1, y: y, w: colW[cIdx] - 0.2, h: rowH,
        fontSize: 11, fontFace: CN_FONT,
        color: theme.primary, bold: isBromoCol,
        align: cIdx === 0 ? "left" : "center", valign: "middle"
      });
    });
  });

  // 结论条
  const conclusionY = startY + (rows.length + 1) * rowH + 0.2;
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: conclusionY, w: 9.0, h: 0.6,
    fill: { color: theme.accent },
    rectRadius: 0.05
  });
  slide.addText("AiKen 已验证\"复古数码\"的真实需求；Bromo 以完整胶片仪式感承接升级需求，客单价 ¥699 → ¥1299（+86%），把品类从入门款抬升到质感款。", {
    x: 0.7, y: conclusionY, w: 8.6, h: 0.6,
    fontSize: 12, fontFace: CN_FONT,
    color: "FFFFFF", bold: true,
    align: "left", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P7 分节页：硬件定义 =====================
function createSlide07(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.primary };

  // 左侧草图
  try {
    slide.addImage({
      path: imgPath('sketch01'),
      x: 0, y: 0, w: 4.5, h: 5.625,
      sizing: { type: 'cover', w: 4.5, h: 5.625 }
    });
  } catch (e) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 0, y: 0, w: 4.5, h: 5.625,
      fill: { color: theme.secondary }
    });
  }

  // 右侧内容区
  slide.addText("04", {
    x: 5.0, y: 1.5, w: 1.0, h: 0.8,
    fontSize: 72, fontFace: EN_FONT,
    color: theme.accent, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("硬件定义", {
    x: 5.0, y: 2.3, w: 4.5, h: 0.5,
    fontSize: 36, fontFace: CN_FONT,
    color: theme.light, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("质感来自每一个结构件", {
    x: 5.0, y: 2.8, w: 4.5, h: 0.4,
    fontSize: 18, fontFace: CN_FONT,
    color: "999999", bold: false,
    align: "left", valign: "middle"
  });

  slide.addText("软件沿袭 AiKen 为主，硬件承载 Bromo 的全部差异化。本章先定形态，再逐一拆解六大核心部件。", {
    x: 5.0, y: 3.4, w: 4.5, h: 0.8,
    fontSize: 12, fontFace: CN_FONT,
    color: "888888", bold: false,
    align: "left", valign: "top"
  });

  // 部件清单
  const parts = "外观形态 ｜ 过片拨杆 ｜ 快门 ｜ 风格拨轮 ｜ 数字胶卷 ｜ 氙气闪光灯 ｜ 旁轴取景";
  slide.addText(parts, {
    x: 5.0, y: 4.4, w: 4.5, h: 0.3,
    fontSize: 10, fontFace: CN_FONT,
    color: theme.accent, bold: false,
    align: "left", valign: "middle"
  });

  // 角注
  slide.addText("注：影像部分（传感器 / 对焦模组规格）不在本次定义范围，另行评审。", {
    x: 5.0, y: 5.0, w: 4.5, h: 0.3,
    fontSize: 9, fontFace: CN_FONT,
    color: "666666", bold: false,
    align: "left", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P8 外观形态 =====================
function createSlide08(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("外观形态——先在物理上成立，再谈数码", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.5,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("挂绳在颈间、机身在掌心，它应该像一件老物件，而不是一台电子产品。", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.35,
    fontSize: 12, fontFace: CN_FONT,
    color: "888888", bold: false, italic: true,
    align: "left", valign: "middle"
  });

  // 三图并列
  const images = [
    { name: 'leica-m3', label: 'Leica M 的设计语言', desc: '经典旁轴的比例与顶板布局，辨识度与"正统感"的来源' },
    { name: 'c35', label: 'Konica C35 的尺寸', desc: '巴掌大小、单手可握：复古不等于笨重' },
    { name: 'front-3d', label: 'Bromo 正面 3D 方案', desc: '布局参考 Yashica，功能件排布直觉化' }
  ];

  images.forEach((img, i) => {
    const x = 0.5 + i * 3.1;
    const y = 1.4;

    try {
      slide.addImage({
        path: imgPath(img.name),
        x: x, y: y, w: 2.9, h: 2.2,
        sizing: { type: 'cover', w: 2.9, h: 2.2 }
      });
    } catch (e) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x, y: y, w: 2.9, h: 2.2,
        fill: { color: theme.light }
      });
    }

    slide.addText(img.label, {
      x: x, y: y + 2.3, w: 2.9, h: 0.3,
      fontSize: 13, fontFace: CN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "middle"
    });

    slide.addText(img.desc, {
      x: x, y: y + 2.6, w: 2.9, h: 0.5,
      fontSize: 10, fontFace: CN_FONT,
      color: theme.secondary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 价值锚点
  slide.addText("经典相机的形制是几十年验证过的\"手感标准\"，Bromo 不发明形态，只继承最优解。", {
    x: 0.5, y: 4.9, w: 9.0, h: 0.4,
    fontSize: 12, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P9 开关 + 过片拨杆 =====================
function createSlide09(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("开关 + 过片拨杆——胶片机的心跳", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.5,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("拍完一张，拇指拨动过片杆，\"咔哒\"一声上弦——这一声，是普通数码相机给不了的。", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.35,
    fontSize: 12, fontFace: CN_FONT,
    color: "888888", bold: false, italic: true,
    align: "left", valign: "middle"
  });

  // 左侧文案
  const points = [
    "开关与过片拨杆一体：On/Off 切换",
    "真实机械阻尼与声音反馈，还原胶片机过片手感",
    "与快门释放 1:1 互锁：已过片、未拍摄时拨杆物理锁止、不可再拨——还原真实胶片机的防呆逻辑",
    "满卷离合脱开：满卷时离合脱开，拨杆行程中阻尼变小，\"脱开卸力\"的手感即是换卷提示"
  ];

  points.forEach((pt, i) => {
    const y = 1.5 + i * 0.6;
    slide.addText("•", {
      x: 0.5, y: y, w: 0.3, h: 0.4,
      fontSize: 14, fontFace: CN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "top"
    });
    slide.addText(pt, {
      x: 0.8, y: y, w: 4.5, h: 0.5,
      fontSize: 12, fontFace: CN_FONT,
      color: theme.primary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 右侧图片
  try {
    slide.addImage({
      path: imgPath('advance-lever'),
      x: 5.5, y: 1.5, w: 4.0, h: 2.5,
      sizing: { type: 'cover', w: 4.0, h: 2.5 }
    });
  } catch (e) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y: 1.5, w: 4.0, h: 2.5,
      fill: { color: theme.light }
    });
  }

  // 价值锚点
  slide.addText("把\"过片\"从多余动作变成产品签名——用户为这一声\"咔哒\"买单。", {
    x: 0.5, y: 4.9, w: 9.0, h: 0.4,
    fontSize: 12, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P10 快门 =====================
function createSlide10(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("快门——半按，是对画面的郑重其事", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.5,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("半按对焦、全按拍摄。两段行程，两份专注。", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.35,
    fontSize: 12, fontFace: CN_FONT,
    color: "888888", bold: false, italic: true,
    align: "left", valign: "middle"
  });

  // 左侧文案
  const points = [
    "机械连杆 + 双段式微动开关：半按对焦 + 测光，全按释放拍摄",
    "未过片时无半按行程，全按亦不触发\"快门声\"、不触发拍摄（与过片互锁互为闭环）",
    "取景器内对焦确认灯 + 声音反馈：对焦状态无需离开取景器即可确认",
    "快门键 M3x0.5 锥形螺纹：兼容标准快门线等延展配件"
  ];

  points.forEach((pt, i) => {
    const y = 1.5 + i * 0.6;
    slide.addText("•", {
      x: 0.5, y: y, w: 0.3, h: 0.4,
      fontSize: 14, fontFace: CN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "top"
    });
    slide.addText(pt, {
      x: 0.8, y: y, w: 4.5, h: 0.5,
      fontSize: 12, fontFace: CN_FONT,
      color: theme.primary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 右侧双图
  try {
    slide.addImage({
      path: imgPath('half-press'),
      x: 5.5, y: 1.5, w: 1.9, h: 1.5,
      sizing: { type: 'cover', w: 1.9, h: 1.5 }
    });
  } catch (e) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y: 1.5, w: 1.9, h: 1.5,
      fill: { color: theme.light }
    });
  }

  try {
    slide.addImage({
      path: imgPath('focus-confirm'),
      x: 7.6, y: 1.5, w: 1.9, h: 1.5,
      sizing: { type: 'cover', w: 1.9, h: 1.5 }
    });
  } catch (e) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 7.6, y: 1.5, w: 1.9, h: 1.5,
      fill: { color: theme.light }
    });
  }

  // 图注
  slide.addText("半按对焦及测光", {
    x: 5.5, y: 3.05, w: 1.9, h: 0.25,
    fontSize: 9, fontFace: CN_FONT,
    color: "999999", bold: false,
    align: "center", valign: "middle"
  });
  slide.addText("取景器内对焦确认灯及声音反馈", {
    x: 7.6, y: 3.05, w: 1.9, h: 0.25,
    fontSize: 9, fontFace: CN_FONT,
    color: "999999", bold: false,
    align: "center", valign: "middle"
  });

  // 价值锚点
  slide.addText("\"半按\"是摄影最经典的肌肉记忆——Bromo 完整保留它，而不是用触屏替代。", {
    x: 0.5, y: 4.9, w: 9.0, h: 0.4,
    fontSize: 12, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P11 换卷弹钮 + 风格拨轮 =====================
function createSlide11(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("换卷弹钮 + 风格拨轮——一个零件，三种身份", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.5,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("换一卷胶卷，换一种心情。弹起、旋转、按下——三个动作完成换卷与选风格。", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.35,
    fontSize: 12, fontFace: CN_FONT,
    color: "888888", bold: false, italic: true,
    align: "left", valign: "middle"
  });

  // 左侧文案
  const points = [
    "一个零件，三种身份：换卷弹钮 = 风格拨轮 = 卷状态件",
    "轴向两态：弹起高位（可旋转）↔ 按下低位（沉入顶面，不可旋转）",
    "旋转三档：暖调 / 冷调 / 灰调，弹起态可自由切换",
    "状态可视：沉入后红刻线与顶面图标依然可视——整卷拍摄期间，顶面始终读得出\"这卷是什么调\"",
    "几何锁定：按下后滚花侧壁全部低于表面、手指无法施力——零额外锁止机构"
  ];

  points.forEach((pt, i) => {
    const y = 1.5 + i * 0.55;
    slide.addText("•", {
      x: 0.5, y: y, w: 0.3, h: 0.4,
      fontSize: 14, fontFace: CN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "top"
    });
    slide.addText(pt, {
      x: 0.8, y: y, w: 4.5, h: 0.5,
      fontSize: 11, fontFace: CN_FONT,
      color: theme.primary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 右侧两态对比示意图（用形状绘制）
  const drawX = 5.5;
  const drawY = 1.5;

  // 弹起态
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: drawX, y: drawY, w: 1.8, h: 1.8,
    fill: { color: theme.light },
    rectRadius: 0.1
  });
  slide.addShape(pres.shapes.OVAL, {
    x: drawX + 0.5, y: drawY + 0.4, w: 0.8, h: 0.8,
    fill: { color: theme.accent }
  });
  slide.addText("弹起态", {
    x: drawX, y: drawY + 1.28, w: 1.8, h: 0.25,
    fontSize: 10, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });
  slide.addText("可旋转 · 高位 3.5–4mm", {
    x: drawX, y: drawY + 1.55, w: 1.8, h: 0.2,
    fontSize: 9, fontFace: CN_FONT,
    color: "999999", bold: false,
    align: "center", valign: "middle"
  });

  // 按下态
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: drawX + 2.2, y: drawY, w: 1.8, h: 1.8,
    fill: { color: theme.light },
    rectRadius: 0.1
  });
  slide.addShape(pres.shapes.OVAL, {
    x: drawX + 2.7, y: drawY + 0.6, w: 0.8, h: 0.4,
    fill: { color: theme.secondary }
  });
  slide.addText("按下态", {
    x: drawX + 2.2, y: drawY + 1.28, w: 1.8, h: 0.25,
    fontSize: 10, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });
  slide.addText("不可旋转 · 沉入约 1mm", {
    x: drawX + 2.2, y: drawY + 1.55, w: 1.8, h: 0.2,
    fontSize: 9, fontFace: CN_FONT,
    color: "999999", bold: false,
    align: "center", valign: "middle"
  });

  // 三档色环示意
  slide.addText("三档色环", {
    x: drawX, y: drawY + 2.0, w: 4.0, h: 0.25,
    fontSize: 10, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });
  const colors = [
    { name: "暖调", color: "c8702a" },
    { name: "灰调", color: "999999" },
    { name: "冷调", color: "4a7a9a" }
  ];
  colors.forEach((c, i) => {
    const cx = drawX + i * 1.3;
    slide.addShape(pres.shapes.OVAL, {
      x: cx, y: drawY + 2.3, w: 0.3, h: 0.3,
      fill: { color: c.color }
    });
    slide.addText(c.name, {
      x: cx + 0.35, y: drawY + 2.3, w: 0.8, h: 0.3,
      fontSize: 9, fontFace: CN_FONT,
      color: theme.secondary, bold: false,
      align: "left", valign: "middle"
    });
  });

  // 价值锚点
  slide.addText("工业设计的\"少即是多\"——一个零件讲完换卷、选风格、读状态三件事，成本与装配效率兼得。", {
    x: 0.5, y: 4.9, w: 9.0, h: 0.4,
    fontSize: 12, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P12 数字胶卷 =====================
function createSlide12(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("数字胶卷——36 张的节制，揭晓的期待", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.5,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("拍满 36 张才能\"冲扫\"。节制带来专注，等待带来惊喜。", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.35,
    fontSize: 12, fontFace: CN_FONT,
    color: "888888", bold: false, italic: true,
    align: "left", valign: "middle"
  });

  // 左侧文案
  const points = [
    "仿胶卷形态读卡器，可插入 TF 卡——\"装卷\"动作完整还原",
    "触点或 Type-C 公头与相机连接；支持 Type-C 直插手机，App 经 USB 导图",
    "相机可对 TF 卡格式化；数据加密，U 盘模式无法直接复制图片",
    "U 盘模式可用空间与相机存储为分离空间",
    "体验节奏：36 张一卷、多卷连续、事后统一冲扫"
  ];

  points.forEach((pt, i) => {
    const y = 1.5 + i * 0.55;
    slide.addText("•", {
      x: 0.5, y: y, w: 0.3, h: 0.4,
      fontSize: 14, fontFace: CN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "top"
    });
    slide.addText(pt, {
      x: 0.8, y: y, w: 4.5, h: 0.5,
      fontSize: 11, fontFace: CN_FONT,
      color: theme.primary, bold: i === 2, // 加密条加粗
      align: "left", valign: "top"
    });
  });

  // 右侧三图（一大两小）
  try {
    slide.addImage({
      path: imgPath('digital-film'),
      x: 5.5, y: 1.5, w: 4.0, h: 2.0,
      sizing: { type: 'cover', w: 4.0, h: 2.0 }
    });
  } catch (e) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y: 1.5, w: 4.0, h: 2.0,
      fill: { color: theme.light }
    });
  }

  try {
    slide.addImage({
      path: imgPath('film-usb'),
      x: 5.5, y: 3.6, w: 1.9, h: 1.0,
      sizing: { type: 'cover', w: 1.9, h: 1.0 }
    });
  } catch (e) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y: 3.6, w: 1.9, h: 1.0,
      fill: { color: theme.light }
    });
  }

  try {
    slide.addImage({
      path: imgPath('storage-bay'),
      x: 7.6, y: 3.6, w: 1.9, h: 1.0,
      sizing: { type: 'cover', w: 1.9, h: 1.0 }
    });
  } catch (e) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 7.6, y: 3.6, w: 1.9, h: 1.0,
      fill: { color: theme.light }
    });
  }

  // 价值锚点
  slide.addText("把\"冲扫的仪式\"与\"即时的速度\"都交给用户选择——而且都免费。", {
    x: 0.5, y: 4.9, w: 9.0, h: 0.4,
    fontSize: 12, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P13 氙气闪光灯模块 =====================
function createSlide13(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("弹出式氙气闪光灯——夜晚的开场仪式", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.5,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("指尖一拨，\"叮\"的一声弹起。夜晚的故事，从这一声开始。", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.35,
    fontSize: 12, fontFace: CN_FONT,
    color: "888888", bold: false, italic: true,
    align: "left", valign: "middle"
  });

  // 左侧文案
  const points = [
    "按压弹出结构或弹出开关，收纳时与顶面齐平",
    "弹出后电容蓄电，蓄电完成指示灯亮起——等待本身即是仪式",
    "独立 2×AA 电池仓，专为闪光灯供电：闪光性能不向整机续航妥协",
    "氙气光源：真闪光灯的亮度与氛围感，LED 补光无法替代"
  ];

  points.forEach((pt, i) => {
    const y = 1.5 + i * 0.6;
    slide.addText("•", {
      x: 0.5, y: y, w: 0.3, h: 0.4,
      fontSize: 14, fontFace: CN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "top"
    });
    slide.addText(pt, {
      x: 0.8, y: y, w: 4.5, h: 0.5,
      fontSize: 12, fontFace: CN_FONT,
      color: theme.primary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 右侧三图：上排两图 + 下排一宽图，图注与机制要点一一对应
  const imgs = [
    { name: 'popup-flash', label: '弹出式闪光灯' },
    { name: 'charge-confirm', label: '弹出后蓄电及完成指示' },
    { name: 'aa-battery', label: '独立 2×AA 供电' }
  ];

  imgs.forEach((img, i) => {
    const x = i < 2 ? 5.5 + i * 2.1 : 5.5;
    const y = i < 2 ? 1.5 : 3.1;
    const w = i < 2 ? 1.9 : 4.0;
    const h = i < 2 ? 1.2 : 1.3;
    const labelY = i < 2 ? 2.75 : 4.45;

    try {
      slide.addImage({
        path: imgPath(img.name),
        x: x, y: y, w: w, h: h,
        sizing: { type: 'cover', w: w, h: h }
      });
    } catch (e) {
      slide.addShape(pres.shapes.RECTANGLE, {
        x: x, y: y, w: w, h: h,
        fill: { color: theme.light }
      });
    }

    slide.addText(img.label, {
      x: x, y: labelY, w: w, h: 0.22,
      fontSize: 9, fontFace: CN_FONT,
      color: "999999", bold: false,
      align: "center", valign: "middle"
    });
  });

  // 价值锚点
  slide.addText("从\"不支持\"到\"弹出式氙气\"——这是 Bromo 与 AiKen 差距最外显的一环。", {
    x: 0.5, y: 4.9, w: 9.0, h: 0.4,
    fontSize: 12, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P14 旁轴取景 =====================
function createSlide14(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("旁轴取景——一只眼睛的专注", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.5,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("贴着眼眶构图，而不是隔着一块屏幕。", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.35,
    fontSize: 12, fontFace: CN_FONT,
    color: "888888", bold: false, italic: true,
    align: "left", valign: "middle"
  });

  // 左侧文案
  const points = [
    "旁轴光学取景器；与摄像头尽量保持 X 轴无偏移、仅 Y 轴偏移",
    "框线系统：成像区域框 / 视差补偿框 / 对焦构图框",
    "取景器内 LED 双色灯：曝光不足橙色闪烁 ｜ 曝光正常绿色常亮",
    "仅中央区域对焦——与\"半按对焦 + 确认灯\"形成一致的胶片机操作习惯",
    "视差补偿框为非必须项，但应尽力达成"
  ];

  points.forEach((pt, i) => {
    const y = 1.5 + i * 0.55;
    slide.addText("•", {
      x: 0.5, y: y, w: 0.3, h: 0.4,
      fontSize: 14, fontFace: CN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "top"
    });
    slide.addText(pt, {
      x: 0.8, y: y, w: 4.5, h: 0.5,
      fontSize: 11, fontFace: CN_FONT,
      color: theme.primary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 右侧图片
  try {
    slide.addImage({
      path: imgPath('viewfinder'),
      x: 5.5, y: 1.5, w: 4.0, h: 2.5,
      sizing: { type: 'cover', w: 4.0, h: 2.5 }
    });
  } catch (e) {
    slide.addShape(pres.shapes.RECTANGLE, {
      x: 5.5, y: 1.5, w: 4.0, h: 2.5,
      fill: { color: theme.light }
    });
  }

  // 价值锚点
  slide.addText("光学取景无延迟、无耗电、有\"相机感\"——这是复古数码的取景答案。", {
    x: 0.5, y: 4.9, w: 9.0, h: 0.4,
    fontSize: 12, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P15 软件沿用说明 =====================
function createSlide15(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("软件沿用说明——工程重心清晰", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.5,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("软件不新增故事，硬件才是本次立项的主角。", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.35,
    fontSize: 12, fontFace: CN_FONT,
    color: "888888", bold: false, italic: true,
    align: "left", valign: "middle"
  });

  // 左栏：沿用 AiKen
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 1.5, w: 4.4, h: 2.8,
    fill: { color: "FFFFFF" },
    rectRadius: 0.08
  });

  slide.addText("沿用 AiKen（主体，复用）", {
    x: 0.7, y: 1.6, w: 4.0, h: 0.35,
    fontSize: 16, fontFace: CN_FONT,
    color: theme.accent, bold: true,
    align: "left", valign: "middle"
  });

  const reuseItems = [
    "风格滤镜引擎（颗粒 / 偏色 / 漏光等全部风格）",
    "App 导图与拼贴分享流程"
  ];
  reuseItems.forEach((item, i) => {
    slide.addText("• " + item, {
      x: 0.7, y: 2.05 + i * 0.4, w: 4.0, h: 0.35,
      fontSize: 12, fontFace: CN_FONT,
      color: theme.primary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 右栏：Bromo 增量适配
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 5.1, y: 1.5, w: 4.4, h: 2.8,
    fill: { color: "FFFFFF" },
    rectRadius: 0.08
  });

  slide.addText("Bromo 增量适配（有限，可控）", {
    x: 5.3, y: 1.6, w: 4.0, h: 0.35,
    fontSize: 16, fontFace: CN_FONT,
    color: theme.secondary, bold: true,
    align: "left", valign: "middle"
  });

  const newItems = [
    "36 张计数与多卷管理",
    "数字胶卷仓识别与数据加密",
    "风格拨轮三档联动",
    "闪光灯弹出与蓄电状态"
  ];
  newItems.forEach((item, i) => {
    slide.addText("• " + item, {
      x: 5.3, y: 2.05 + i * 0.4, w: 4.0, h: 0.35,
      fontSize: 12, fontFace: CN_FONT,
      color: theme.primary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 结论条
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 4.5, w: 9.0, h: 0.6,
    fill: { color: theme.accent },
    rectRadius: 0.05
  });
  slide.addText("软件复用度高，开发风险与周期可控；Bromo 的工程投入集中于结构件与光机电配合。", {
    x: 0.7, y: 4.5, w: 8.6, h: 0.6,
    fontSize: 12, fontFace: CN_FONT,
    color: "FFFFFF", bold: true,
    align: "left", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P16 价值主张 =====================
function createSlide16(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("¥1299，买的是什么", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.5,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("同价位买得到的\"复古外观\"很多；买得到的\"机械仪式感\"，只有这一台。", {
    x: 0.5, y: 0.9, w: 9.0, h: 0.35,
    fontSize: 12, fontFace: CN_FONT,
    color: "888888", bold: false, italic: true,
    align: "left", valign: "middle"
  });

  // 三栏卡片
  const cards = [
    { title: "同价位横评", desc: "¥1000–1500 价位段的复古数码普遍停留在\"外观复古\"；真拨杆过片、快门互锁、换卷机制、弹出氙闪的\"操作复古\"，为 Bromo 独有。" },
    { title: "对比胶片", desc: "一卷胶片加冲扫，动辄几十上百元、以周计等待；Bromo 保留全部仪式，耗材成本为零，冲扫即时可选——仪式感\"无限续杯\"。" },
    { title: "对比 AiKen", desc: "+¥600，换得更大底与 AF、氙气闪光、完整胶片仪式、135 金属旁轴形态；客单价提升 86%，产品线完成从入门到质感的卡位。" }
  ];

  cards.forEach((card, i) => {
    const x = 0.5 + i * 3.1;
    const y = 1.5;

    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: x, y: y, w: 2.9, h: 2.5,
      fill: { color: "FFFFFF" },
      rectRadius: 0.08
    });

    slide.addText(card.title, {
      x: x + 0.15, y: y + 0.15, w: 2.6, h: 0.35,
      fontSize: 16, fontFace: CN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "middle"
    });

    slide.addText(card.desc, {
      x: x + 0.15, y: y + 0.6, w: 2.6, h: 1.7,
      fontSize: 11, fontFace: CN_FONT,
      color: theme.secondary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 底部收束条
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: 0.5, y: 4.2, w: 9.0, h: 0.8,
    fill: { color: theme.primary },
    rectRadius: 0.05
  });
  slide.addText("¥1299 = 一台旁轴相机 + 无限卷数字胶卷 + AiKen 全部风格滤镜", {
    x: 0.7, y: 4.2, w: 8.6, h: 0.8,
    fontSize: 16, fontFace: CN_FONT,
    color: theme.light, bold: true,
    align: "center", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P17 附录：研发关键参数速查 =====================
function createSlide17(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("附录 · 研发关键参数速查", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.5,
    fontSize: 28, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "left", valign: "middle"
  });

  slide.addText("本页供研发职能速查，汇报时可快速带过或跳过。", {
    x: 0.5, y: 0.85, w: 9.0, h: 0.25,
    fontSize: 10, fontFace: CN_FONT,
    color: "999999", bold: false,
    align: "left", valign: "middle"
  });

  // 分组表格
  const params = [
    ["快门", "高位约 3.5–4mm ｜ 键帽 M3x0.5 锥形螺纹（兼容标准快门线）｜ 机械连杆 + 双段式微动开关"],
    ["过片拨杆", "与快门释放 1:1 互锁（物理锁止）｜ 满卷离合脱开、行程阻尼变小"],
    ["换卷弹钮/风格拨轮", "弹起高位 3.5–4mm（可旋转）／ 按下沉入顶面约 1mm（不可旋转）｜ 三档 120° 等分（暖/冷/灰）｜ 几何锁定：滚花侧壁低于表面"],
    ["旁轴取景", "与摄像头 X 轴无偏移、仅 Y 轴偏移 ｜ 视差补偿框：约 1m 内近摄，非必须、尽力达成 ｜ 仅中央区域对焦 ｜ LED：橙闪 = 欠曝 / 绿常亮 = 正常"],
    ["数字胶卷", "触点或 Type-C 公头与相机连接 ｜ TF 卡可格式化、数据加密（U 盘模式不可直拷）｜ U 盘模式空间与相机存储分离"],
    ["氙气闪光灯", "独立 2×AA 电池仓 ｜ 弹出后电容蓄电、完成指示灯 ｜ 按压弹出结构或弹出开关"],
    ["外观形态", "设计参考：Leica M ｜ 尺寸参考：Konica C35 ｜ 布局参考：Yashica"]
  ];

  let y = 1.2;
  params.forEach(([module, detail]) => {
    slide.addText(module, {
      x: 0.5, y: y, w: 2.0, h: 0.35,
      fontSize: 11, fontFace: CN_FONT,
      color: theme.accent, bold: true,
      align: "left", valign: "top"
    });
    slide.addText(detail, {
      x: 2.6, y: y, w: 7.0, h: 0.5,
      fontSize: 10, fontFace: CN_FONT,
      color: theme.secondary, bold: false,
      align: "left", valign: "top"
    });
    y += 0.55;
  });

  addPageBadge(slide, pres, theme);
}

// ===================== P18 结尾：立项请求 =====================
function createSlide18(pres, theme) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  slide.addText("请求立项评审通过", {
    x: 0.5, y: 0.4, w: 9.0, h: 0.5,
    fontSize: 36, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });

  // 左栏：已收敛的定义
  slide.addText("已收敛的定义", {
    x: 0.5, y: 1.2, w: 4.0, h: 0.35,
    fontSize: 16, fontFace: CN_FONT,
    color: theme.accent, bold: true,
    align: "left", valign: "middle"
  });

  const doneItems = [
    "一句话定义与目标用户画像收敛",
    "外观形态 + 六大核心部件已定义到机制层级",
    "软件高复用，工程重心集中于结构件，风险可控",
    "¥1299 价值主张与产品线卡位清晰"
  ];
  doneItems.forEach((item, i) => {
    slide.addText("• " + item, {
      x: 0.5, y: 1.6 + i * 0.4, w: 4.0, h: 0.35,
      fontSize: 12, fontFace: CN_FONT,
      color: theme.primary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 右栏：下一步
  slide.addText("下一步（评审通过后）", {
    x: 5.5, y: 1.2, w: 4.0, h: 0.35,
    fontSize: 16, fontFace: CN_FONT,
    color: theme.accent, bold: true,
    align: "left", valign: "middle"
  });

  const nextItems = [
    "结构件详细设计与手板验证",
    "光机电联调",
    "EVT 样机"
  ];
  nextItems.forEach((item, i) => {
    slide.addText("• " + item, {
      x: 5.5, y: 1.6 + i * 0.4, w: 4.0, h: 0.35,
      fontSize: 12, fontFace: CN_FONT,
      color: theme.primary, bold: false,
      align: "left", valign: "top"
    });
  });

  // 收束语
  slide.addText("让认真拍照这件事，重新值得被认真对待。", {
    x: 0.5, y: 3.8, w: 9.0, h: 0.6,
    fontSize: 22, fontFace: CN_FONT,
    color: theme.primary, bold: true,
    align: "center", valign: "middle"
  });

  slide.addText("Bromo · 产品定义评审", {
    x: 0.5, y: 4.5, w: 9.0, h: 0.3,
    fontSize: 12, fontFace: CN_FONT,
    color: "999999", bold: false,
    align: "center", valign: "middle"
  });

  addPageBadge(slide, pres, theme);
}

// ===================== 生成所有幻灯片 =====================
createSlide01(pres, theme);
createSlide02(pres, theme);
createSlide03(pres, theme);
createSlide04(pres, theme);
createSlide05(pres, theme);
createSlide06(pres, theme);
createSlide07(pres, theme);
createSlide08(pres, theme);
createSlide09(pres, theme);
createSlide10(pres, theme);
createSlide11(pres, theme);
createSlide12(pres, theme);
createSlide13(pres, theme);
createSlide14(pres, theme);
createSlide15(pres, theme);
createSlide16(pres, theme);
createSlide17(pres, theme);
createSlide18(pres, theme);

// 输出
const outputPath = path.join(__dirname, '..', 'Bromo产品定义介绍.pptx');
pres.writeFile({ fileName: outputPath });
console.log('PPT generated:', outputPath);
