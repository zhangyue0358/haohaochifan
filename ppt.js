const pptxgen = require("pptxgenjs");
const React = require("react");
const ReactDOMServer = require("react-dom/server");
const sharp = require("sharp");
const {
  FaLeaf, FaHeartbeat, FaUsers, FaChartLine, FaShieldAlt,
  FaStar, FaCheckCircle, FaTimesCircle, FaArrowRight, FaMobileAlt,
  FaHospital, FaBaby, FaUtensils, FaStore, FaBullseye,
  FaLock, FaRocket, FaMoneyBillWave
} = require("react-icons/fa");

async function icon(IC, color="#FFFFFF", size=256) {
  const svg = ReactDOMServer.renderToStaticMarkup(React.createElement(IC, { color, size: String(size) }));
  const buf = await sharp(Buffer.from(svg)).png().toBuffer();
  return "image/png;base64," + buf.toString("base64");
}

const G1 = "1A7A4C"; // dark green
const G2 = "2BA76A"; // brand green
const G3 = "6DCCA0"; // mid green
const GL = "E8F7F0"; // light green bg
const WH = "FFFFFF";
const BK = "1A1A18";
const GY = "555555";
const LG = "F5F5F0";

const shadow = () => ({ type:"outer", blur:6, offset:2, angle:135, color:"000000", opacity:0.12 });

async function run() {
  let pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.title = "好好吃饭 · 融资计划书";

  // ─────────────────────────────────────────
  // S1: 封面
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: G1 };

    // 左侧装饰色块
    s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:3.2, h:5.625, fill:{color:"164E37"}, line:{color:"164E37"} });

    // 左侧大字装饰
    s.addText("🥗", { x:0.3, y:1.2, w:2.6, h:2.2, fontSize:96, align:"center", valign:"middle" });

    // 右侧内容
    s.addText("好好吃饭", { x:3.6, y:0.6, w:6, h:1.2, fontSize:52, bold:true, color:WH, fontFace:"Arial Black", align:"left" });
    s.addText("天天向上", { x:3.6, y:1.7, w:5, h:0.7, fontSize:28, color:G3, fontFace:"Arial", align:"left", charSpacing:8 });
    s.addText("不是外卖App，是健康行为改变平台", { x:3.6, y:2.5, w:6.1, h:0.5, fontSize:16, color:"C2EDD8", fontFace:"Arial", align:"left" });

    // 分割线
    s.addShape(pres.shapes.LINE, { x:3.6, y:3.1, w:5.8, h:0, line:{color:G3, width:1} });

    // 融资信息
    s.addText([
      { text:"天使轮融资：", options:{ color:G3, fontSize:15 } },
      { text:"500万人民币", options:{ color:WH, fontSize:15, bold:true } },
      { text:"  ·  出让", options:{ color:G3, fontSize:15 } },
      { text:" 10%", options:{ color:WH, fontSize:15, bold:true } },
      { text:" 股权", options:{ color:G3, fontSize:15 } },
    ], { x:3.6, y:3.3, w:6, h:0.5 });

    s.addText([
      { text:"📱 Demo：", options:{ color:G3, fontSize:13 } },
      { text:"haohaochifan-1.netlify.app", options:{ color:WH, fontSize:13, underline:true, hyperlink:{url:"https://haohaochifan-1.netlify.app"} } },
    ], { x:3.6, y:3.85, w:6.1, h:0.45 });

    // 联系人
    s.addText("张越  |  18810409001  |  北京", { x:3.6, y:4.4, w:5, h:0.4, fontSize:12, color:"8ECFB0", fontFace:"Arial" });

    // 右下角幻灯片编号
    s.addText("01 / 12", { x:9, y:5.1, w:0.8, h:0.35, fontSize:10, color:"6DCCA0", align:"right" });
  }

  // ─────────────────────────────────────────
  // S2: 一句话 + 4大数据
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: WH };

    s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:1.1, fill:{color:G1}, line:{color:G1} });
    s.addText("执行摘要", { x:0.5, y:0.2, w:4, h:0.65, fontSize:22, bold:true, color:WH, fontFace:"Arial Black" });
    s.addText("02 / 12", { x:9, y:0.3, w:0.8, h:0.4, fontSize:10, color:G3, align:"right" });

    s.addText("我们不屏蔽任何食物，只做精准标识——让用户在点餐时做出更知情的选择，帮助慢病患者、孕产妇和健康人群养成更好的饮食习惯", {
      x:0.5, y:1.25, w:9, h:0.8, fontSize:15, color:GY, fontFace:"Arial", align:"center"
    });

    // 4个数字卡片
    const cards = [
      { num:"4亿+", lbl:"国内慢病患者", sub:"高血压1.3亿" },
      { num:"1000万", lbl:"年孕产妇数量", sub:"月子中心切入口" },
      { num:"1.2万亿", lbl:"外卖市场规模", sub:"2024年数据" },
      { num:"0", lbl:"同类竞品", sub:"市场空白明确" },
    ];
    cards.forEach((c, i) => {
      const x = 0.4 + i * 2.33;
      s.addShape(pres.shapes.RECTANGLE, { x, y:2.2, w:2.1, h:2.2, fill:{color:G1}, line:{color:G1}, shadow:shadow() });
      s.addText(c.num, { x, y:2.25, w:2.1, h:1.0, fontSize:c.num.length>4?26:36, bold:true, color:WH, fontFace:"Arial Black", align:"center", valign:"middle" });
      s.addText(c.lbl, { x, y:3.2, w:2.1, h:0.5, fontSize:13, bold:true, color:G3, fontFace:"Arial", align:"center" });
      s.addText(c.sub, { x, y:3.7, w:2.1, h:0.45, fontSize:10, color:"8ECFB0", fontFace:"Arial", align:"center" });
    });

    s.addText("核心差异：健康档案驱动的完整餐饮闭环，越用越精准，数据不可迁移", {
      x:0.5, y:4.65, w:9, h:0.45, fontSize:12, color:G2, fontFace:"Arial", align:"center", italic:true
    });
  }

  // ─────────────────────────────────────────
  // S3: 用户痛点
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: WH };
    s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:1.1, fill:{color:G1}, line:{color:G1} });
    s.addText("用户痛点", { x:0.5, y:0.22, w:4, h:0.65, fontSize:22, bold:true, color:WH, fontFace:"Arial Black" });
    s.addText("03 / 12", { x:9, y:0.3, w:0.8, h:0.4, fontSize:10, color:G3, align:"right" });

    // 左侧核心洞察
    s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.25, w:3.8, h:3.8, fill:{color:G1}, line:{color:G1} });
    s.addText("核心洞察", { x:0.5, y:1.4, w:3.6, h:0.5, fontSize:14, bold:true, color:G3, fontFace:"Arial", align:"center" });
    s.addText("问题不在于\n「没有健康食物」", { x:0.5, y:1.95, w:3.6, h:1.0, fontSize:18, bold:true, color:WH, fontFace:"Arial Black", align:"center" });
    s.addText("在于「信息不对称」\n用户不知道哪道菜今天适不适合自己吃", {
      x:0.55, y:3.05, w:3.5, h:1.2, fontSize:12, color:"C2EDD8", fontFace:"Arial", align:"center"
    });

    // 右侧5个痛点
    const pains = [
      "🤔  不知道自己该吃什么，不会选",
      "😟  外卖平台看不出哪道菜对自己健康",
      "🏥  孕妇/高血压/痛风患者完全得不到针对性引导",
      "😤  知道要吃健康，但好习惯难坚持",
      "😰  食材质量和配送安全没有保障",
    ];
    pains.forEach((p, i) => {
      s.addShape(pres.shapes.RECTANGLE, { x:4.5, y:1.2 + i*0.82, w:5.1, h:0.68, fill:{color:i%2===0?GL:WH}, line:{color:"E0E0E0",width:0.5} });
      s.addText(p, { x:4.6, y:1.25 + i*0.82, w:4.9, h:0.58, fontSize:13, color:BK, fontFace:"Arial", valign:"middle" });
    });
  }

  // ─────────────────────────────────────────
  // S4: 市场规模
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: WH };
    s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:1.1, fill:{color:G1}, line:{color:G1} });
    s.addText("市场规模", { x:0.5, y:0.22, w:4, h:0.65, fontSize:22, bold:true, color:WH, fontFace:"Arial Black" });
    s.addText("04 / 12", { x:9, y:0.3, w:0.8, h:0.4, fontSize:10, color:G3, align:"right" });

    // TAM/SAM/SOM 漏斗
    const markets = [
      { label:"TAM · 总市场", num:"1.2万亿", desc:"中国在线外卖市场（2024年）", color:"1A3D2B", w:9.2 },
      { label:"SAM · 目标市场", num:"~800亿", desc:"有健康需求的外卖用户群", color:G1, w:7 },
      { label:"SOM · 3年可获市场", num:"~15亿", desc:"北上深切入，100万付费用户", color:G2, w:5 },
    ];
    markets.forEach((m, i) => {
      const x = (10 - m.w) / 2;
      s.addShape(pres.shapes.RECTANGLE, { x, y:1.3 + i*1.35, w:m.w, h:1.1, fill:{color:m.color}, line:{color:m.color} });
      s.addText(m.label, { x:x+0.3, y:1.35 + i*1.35, w:3, h:0.5, fontSize:13, bold:true, color:G3, fontFace:"Arial", valign:"middle" });
      s.addText(m.num, { x:x + m.w/2 - 1.5, y:1.35 + i*1.35, w:3, h:0.5, fontSize:22, bold:true, color:WH, fontFace:"Arial Black", align:"center" });
      s.addText(m.desc, { x:x + m.w - 3.2, y:1.35 + i*1.35, w:3, h:0.5, fontSize:11, color:"C2EDD8", fontFace:"Arial", align:"right", valign:"middle" });
      s.addText("", { x:x + m.w/2 - 0.3, y:2.35 + i*1.35, w:0.6, h:0.2, fontSize:14, color:G2, align:"center" });
    });

    // 底部关键数据
    s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:5.0, w:9.2, h:0.45, fill:{color:GL}, line:{color:"C2EDD8"} });
    s.addText("外卖用户年增速 CAGR 15%  ·  健康食品消费年增速 22%  ·  慢病患者人数持续增长", {
      x:0.5, y:5.05, w:9, h:0.35, fontSize:11, color:G1, fontFace:"Arial", align:"center", bold:true
    });
  }

  // ─────────────────────────────────────────
  // S5: 产品核心
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: WH };
    s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:1.1, fill:{color:G1}, line:{color:G1} });
    s.addText("产品核心", { x:0.5, y:0.22, w:4, h:0.65, fontSize:22, bold:true, color:WH, fontFace:"Arial Black" });
    s.addText("05 / 12", { x:9, y:0.3, w:0.8, h:0.4, fontSize:10, color:G3, align:"right" });

    // 核心理念大卡
    s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.2, w:9.2, h:1.1, fill:{color:G1}, line:{color:G1} });
    s.addText([
      { text:"核心理念：", options:{ color:G3, fontSize:14, bold:false } },
      { text:"标识而非限制", options:{ color:WH, fontSize:16, bold:true } },
      { text:"  ——  选择权永远在用户手里", options:{ color:"C2EDD8", fontSize:13 } },
    ], { x:0.6, y:1.32, w:8.8, h:0.45 });
    s.addText("「今天钠快够了」「尿酸高时少吃」「✓ 挺适合你」— 朋友式提示，不屏蔽，不强制", {
      x:0.6, y:1.82, w:8.8, h:0.38, fontSize:12, color:"8ECFB0", fontFace:"Arial", italic:true
    });

    // 六大功能模块
    const feats = [
      { icon:"🧬", title:"健康档案", desc:"AI解读+营养师审核\n个性化营养方案" },
      { icon:"🏷️", title:"菜品标识", desc:"100分制评分\n4维度实时计算" },
      { icon:"🍽️", title:"智能推荐", desc:"今日已摄入动态调整\n越用越懂你" },
      { icon:"👨‍👩‍👧", title:"家庭账号", desc:"6人独立档案\n一键全家下单" },
      { icon:"🤰", title:"特殊人群", desc:"孕妇按孕周定制\n慢病多病种联合" },
      { icon:"📊", title:"饮食记录", desc:"外卖自动同步\n周报营养趋势" },
    ];
    feats.forEach((f, i) => {
      const col = i % 3, row = Math.floor(i / 3);
      const x = 0.4 + col * 3.1, y = 2.55 + row * 1.4;
      s.addShape(pres.shapes.RECTANGLE, { x, y, w:2.8, h:1.2, fill:{color:row===0?GL:WH}, line:{color:"D0ECD8"}, shadow:shadow() });
      s.addText(f.icon, { x, y:y+0.05, w:0.8, h:0.8, fontSize:22, align:"center", valign:"middle" });
      s.addText(f.title, { x:x+0.7, y:y+0.08, w:2, h:0.4, fontSize:13, bold:true, color:G1, fontFace:"Arial" });
      s.addText(f.desc, { x:x+0.7, y:y+0.5, w:2, h:0.6, fontSize:10, color:GY, fontFace:"Arial" });
    });
  }

  // ─────────────────────────────────────────
  // S6: Demo展示
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: G1 };

    s.addText("App Demo 在线体验", { x:0.5, y:0.3, w:9, h:0.8, fontSize:28, bold:true, color:WH, fontFace:"Arial Black", align:"center" });
    s.addText("涵盖注册、首页、点外卖、健康评分说明、孕妇/慢病专属、家庭账号、下单配送全链路", {
      x:0.5, y:1.1, w:9, h:0.5, fontSize:13, color:G3, fontFace:"Arial", align:"center"
    });

    // 模拟手机屏幕框
    s.addShape(pres.shapes.RECTANGLE, { x:0.8, y:1.75, w:2.2, h:3.5, fill:{color:"164E37"}, line:{color:G3,width:2}, shadow:shadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:2.0, y:0.9, w:2.2, h:3.5, fill:{color:"1D5C3A"}, line:{color:G3,width:2}, shadow:shadow() });
    s.addShape(pres.shapes.RECTANGLE, { x:3.2, y:1.75, w:2.2, h:3.5, fill:{color:"164E37"}, line:{color:G3,width:2}, shadow:shadow() });

    // 标注文字
    const labels = [
      { x:0.7, y:5.3, t:"注册流程\n4步极简" },
      { x:1.9, y:4.45, t:"首页\n今日核心卡" },
      { x:3.1, y:5.3, t:"点外卖\n健康评分" },
    ];
    labels.forEach(l => s.addText(l.t, { x:l.x, y:l.y, w:2.4, h:0.6, fontSize:10, color:G3, fontFace:"Arial", align:"center" }));

    // 右侧亮点说明
    const pts = [
      "🏷️ 每道菜实时计算个性化健康评分",
      "❓ 点「?」查看评分4维度详细说明",
      "🤰 孕妇孕周定制 + 慢病多病种管理",
      "👨‍👩‍👧 家庭账号各自独立健康档案",
      "📦 下单→配送→健康数据自动同步",
    ];
    pts.forEach((p, i) => {
      s.addShape(pres.shapes.RECTANGLE, { x:5.8, y:1.75 + i*0.68, w:3.8, h:0.56, fill:{color:"164E37"}, line:{color:G3,width:0.5} });
      s.addText(p, { x:5.9, y:1.8 + i*0.68, w:3.6, h:0.45, fontSize:12, color:WH, fontFace:"Arial", valign:"middle" });
    });

    // 扫码区
    s.addShape(pres.shapes.RECTANGLE, { x:5.8, y:5.2, w:3.8, h:0.55, fill:{color:G2}, line:{color:G2} });
    s.addText([
      { text:"📱 立即体验：", options:{ color:WH, fontSize:12, bold:false } },
      { text:"haohaochifan-1.netlify.app", options:{ color:WH, fontSize:12, bold:true, underline:true, hyperlink:{url:"https://haohaochifan-1.netlify.app"} } },
    ], { x:5.85, y:5.22, w:3.7, h:0.38, valign:"middle" });
  }

  // ─────────────────────────────────────────
  // S7: 目标用户
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: WH };
    s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:1.1, fill:{color:G1}, line:{color:G1} });
    s.addText("目标用户", { x:0.5, y:0.22, w:4, h:0.65, fontSize:22, bold:true, color:WH, fontFace:"Arial Black" });
    s.addText("07 / 12", { x:9, y:0.3, w:0.8, h:0.4, fontSize:10, color:G3, align:"right" });

    const users = [
      { icon:"🤰", label:"P1 首选切入", title:"孕妇 / 产妇", num:"年约 1000万", why:"孕期按周营养定制，月子中心合作渠道", color:G1 },
      { icon:"🏥", label:"P1 首选切入", title:"慢病患者", num:"4亿+人群", why:"多病种联合管理，医院营养科渠道", color:G1 },
      { icon:"💼", label:"P2 规模化", title:"健康白领", num:"1.5亿", why:"企业健康福利合作切入，信任专业背书", color:"2C7A55" },
      { icon:"👴", label:"P2 规模化", title:"中老年养生", num:"2.7亿", why:"子女代为下单，家庭账号功能驱动", color:"2C7A55" },
    ];
    users.forEach((u, i) => {
      const x = 0.4 + i * 2.33;
      s.addShape(pres.shapes.RECTANGLE, { x, y:1.2, w:2.1, h:0.35, fill:{color:u.color}, line:{color:u.color} });
      s.addText(u.label, { x, y:1.22, w:2.1, h:0.28, fontSize:9, bold:true, color:WH, fontFace:"Arial", align:"center" });
      s.addShape(pres.shapes.RECTANGLE, { x, y:1.55, w:2.1, h:2.8, fill:{color:GL}, line:{color:"D0ECD8"} });
      s.addText(u.icon, { x, y:1.6, w:2.1, h:0.7, fontSize:28, align:"center" });
      s.addText(u.title, { x:x+0.05, y:2.35, w:2.0, h:0.45, fontSize:14, bold:true, color:G1, fontFace:"Arial Black", align:"center" });
      s.addText(u.num, { x:x+0.05, y:2.8, w:2.0, h:0.4, fontSize:12, bold:true, color:G2, fontFace:"Arial", align:"center" });
      s.addText(u.why, { x:x+0.1, y:3.2, w:1.9, h:1.1, fontSize:10, color:GY, fontFace:"Arial", align:"center" });
    });

    s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:4.45, w:9.2, h:0.75, fill:{color:G1}, line:{color:G1} });
    s.addText([
      { text:"冷启动策略：", options:{ color:G3, fontSize:13, bold:true } },
      { text:"月子中心→医院营养科→企业健康福利，精准渠道逐步扩量，前100-500用户服务到极致", options:{ color:WH, fontSize:13 } },
    ], { x:0.6, y:4.52, w:8.9, h:0.55, valign:"middle" });
  }

  // ─────────────────────────────────────────
  // S8: 竞争分析
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: WH };
    s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:1.1, fill:{color:G1}, line:{color:G1} });
    s.addText("竞争分析", { x:0.5, y:0.22, w:4, h:0.65, fontSize:22, bold:true, color:WH, fontFace:"Arial Black" });
    s.addText("08 / 12", { x:9, y:0.3, w:0.8, h:0.4, fontSize:10, color:G3, align:"right" });

    const headers = ["功能维度", "饿了么/美团", "薄荷健康/Keep", "好好吃饭 ✦"];
    const rows = [
      ["健康档案接入", "✗", "△ 有限", "✓"],
      ["个性化菜品标识", "✗", "✗", "✓"],
      ["营养师人工审核", "✗", "✗", "✓"],
      ["慢病/孕妇专属管理", "✗", "△ 有限", "✓"],
      ["外卖+健康完整闭环", "✗", "✗", "✓"],
      ["数据护城河", "△ 订单历史", "△ 饮食记录", "✓ 健康档案+饮食"],
    ];
    const colW = [2.8, 2.0, 2.2, 2.2];
    const startX = [0.4, 3.2, 5.2, 7.4];

    // 表头
    headers.forEach((h, i) => {
      s.addShape(pres.shapes.RECTANGLE, { x:startX[i], y:1.2, w:colW[i]-0.1, h:0.55, fill:{color:i===3?G2:G1}, line:{color:i===3?G2:G1} });
      s.addText(h, { x:startX[i]+0.05, y:1.22, w:colW[i]-0.15, h:0.5, fontSize:11, bold:true, color:WH, fontFace:"Arial", align:"center", valign:"middle" });
    });

    // 数据行
    rows.forEach((row, ri) => {
      row.forEach((cell, ci) => {
        const bg = ri%2===0 ? (ci===3?GL:LG) : WH;
        s.addShape(pres.shapes.RECTANGLE, { x:startX[ci], y:1.78 + ri*0.55, w:colW[ci]-0.1, h:0.52, fill:{color:bg}, line:{color:"E0E0E0",width:0.5} });
        const isCheck = cell === "✓", isCross = cell === "✗";
        s.addText(cell, {
          x:startX[ci]+0.05, y:1.8 + ri*0.55, w:colW[ci]-0.15, h:0.45,
          fontSize:ci===0?11:13, bold:isCheck||isCross,
          color:isCheck?"1D8348":isCross?"C0392B":ci===3?G1:GY,
          fontFace:"Arial", align:ci===0?"left":"center", valign:"middle"
        });
      });
    });

    s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:5.05, w:9.2, h:0.4, fill:{color:GL}, line:{color:G3} });
    s.addText("护城河：数据飞轮（365天饮食+健康数据不可迁移）· 营养师网络 · 特殊人群壁垒", {
      x:0.5, y:5.1, w:9, h:0.3, fontSize:11, color:G1, fontFace:"Arial", align:"center", bold:true
    });
  }

  // ─────────────────────────────────────────
  // S9: 商业模式
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: WH };
    s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:1.1, fill:{color:G1}, line:{color:G1} });
    s.addText("商业模式", { x:0.5, y:0.22, w:4, h:0.65, fontSize:22, bold:true, color:WH, fontFace:"Arial Black" });
    s.addText("09 / 12", { x:9, y:0.3, w:0.8, h:0.4, fontSize:10, color:G3, align:"right" });

    const biz = [
      { icon:"👤", title:"用户会员订阅", price:"19元/月 · 199元/年", target:"100万付费用户", rev:"~2亿/年", color:G1 },
      { icon:"🏪", title:"商家平台服务费", price:"GMV抽佣2-5%+年费", target:"2000家商家", rev:"~9000万/年", color:"2C7A55" },
      { icon:"🏬", title:"品牌加盟+食材供应", price:"加盟费10-30万", target:"200家加盟店", rev:"~3000万/年", color:"1D5C3A" },
      { icon:"🏢", title:"企业健康福利B2B", price:"5-20元/人/月", target:"1000家企业", rev:"~2000万/年", color:"164E37" },
    ];
    biz.forEach((b, i) => {
      const x = 0.4 + i * 2.33;
      s.addShape(pres.shapes.RECTANGLE, { x, y:1.2, w:2.1, h:3.2, fill:{color:b.color}, line:{color:b.color}, shadow:shadow() });
      s.addText(b.icon, { x, y:1.3, w:2.1, h:0.65, fontSize:28, align:"center" });
      s.addText(b.title, { x:x+0.05, y:1.95, w:2.0, h:0.6, fontSize:11, bold:true, color:WH, fontFace:"Arial", align:"center" });
      s.addText(b.price, { x:x+0.08, y:2.55, w:1.94, h:0.45, fontSize:10, color:"8ECFB0", fontFace:"Arial", align:"center" });
      s.addText(b.target, { x:x+0.08, y:3.0, w:1.94, h:0.35, fontSize:9, color:"C2EDD8", fontFace:"Arial", align:"center" });
      s.addShape(pres.shapes.RECTANGLE, { x:x+0.1, y:3.35, w:1.9, h:0.65, fill:{color:G2}, line:{color:G2} });
      s.addText(b.rev, { x:x+0.1, y:3.38, w:1.9, h:0.55, fontSize:13, bold:true, color:WH, fontFace:"Arial Black", align:"center", valign:"middle" });
    });

    s.addText("冷启动关键：前100名种子用户服务到极致，用真实健康改善数据做口碑传播内容", {
      x:0.4, y:4.6, w:9.2, h:0.45, fontSize:12, color:G2, fontFace:"Arial", align:"center", italic:true
    });
    s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:5.05, w:9.2, h:0.4, fill:{color:G1}, line:{color:G1} });
    s.addText("第2年预计实现盈利  ·  第3年目标总收入 6,000万  ·  第4年目标 2亿", {
      x:0.5, y:5.1, w:9, h:0.3, fontSize:11, bold:true, color:WH, fontFace:"Arial", align:"center"
    });
  }

  // ─────────────────────────────────────────
  // S10: 财务预测
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: WH };
    s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:1.1, fill:{color:G1}, line:{color:G1} });
    s.addText("财务预测", { x:0.5, y:0.22, w:4, h:0.65, fontSize:22, bold:true, color:WH, fontFace:"Arial Black" });
    s.addText("10 / 12", { x:9, y:0.3, w:0.8, h:0.4, fontSize:10, color:G3, align:"right" });

    // 收入增长柱状图
    s.addChart(pres.charts.BAR, [
      { name:"总收入（万元）", labels:["第1年（种子期）","第2年（起步期）","第3年（增长期）","第4年（规模期）"], values:[150,1500,6000,20000] }
    ], {
      x:0.4, y:1.2, w:5.8, h:3.5, barDir:"col",
      chartColors:[G2],
      chartArea:{ fill:{color:WH}, roundedCorners:true },
      catAxisLabelColor:"64748B",
      valAxisLabelColor:"64748B",
      valGridLine:{ color:"E2E8F0", size:0.5 },
      catGridLine:{ style:"none" },
      showValue:true, dataLabelPosition:"outEnd", dataLabelColor:G1, dataLabelFontSize:11,
      showLegend:false,
      showTitle:false,
    });

    // 右侧关键指标
    const metrics = [
      { year:"第1年", rev:"150万", profit:"亏损 230万", note:"种子期，验证模型" },
      { year:"第2年", rev:"1,500万", profit:"+盈利 600万", note:"起步期，规模化" },
      { year:"第3年", rev:"6,000万", profit:"+盈利 3,200万", note:"增长期，扩张" },
      { year:"第4年", rev:"2亿", profit:"+盈利 1.2亿", note:"规模期，IPO准备" },
    ];
    metrics.forEach((m, i) => {
      const y = 1.2 + i * 1.0;
      s.addShape(pres.shapes.RECTANGLE, { x:6.5, y, w:3.1, h:0.85, fill:{color:i===0?LG:GL}, line:{color:"D0ECD8"} });
      s.addText(m.year, { x:6.6, y:y+0.05, w:0.8, h:0.35, fontSize:11, bold:true, color:G1, fontFace:"Arial" });
      s.addText(m.rev, { x:7.4, y:y+0.05, w:1.0, h:0.35, fontSize:12, bold:true, color:G2, fontFace:"Arial Black" });
      s.addText(m.profit, { x:6.6, y:y+0.42, w:2.9, h:0.32, fontSize:10, color:m.profit.startsWith("+")?"1D8348":"C0392B", fontFace:"Arial", bold:true });
    });

    s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:5.05, w:9.2, h:0.4, fill:{color:GL}, line:{color:G3} });
    s.addText("第2年底实现盈利  ·  关键指标：月留存率 > 70%  ·  付费转化率 > 15%", {
      x:0.5, y:5.1, w:9, h:0.3, fontSize:11, color:G1, fontFace:"Arial", align:"center", bold:true
    });
  }

  // ─────────────────────────────────────────
  // S11: 融资计划 & 路线图
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: WH };
    s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:1.1, fill:{color:G1}, line:{color:G1} });
    s.addText("融资计划 & 执行路线图", { x:0.5, y:0.22, w:6, h:0.65, fontSize:22, bold:true, color:WH, fontFace:"Arial Black" });
    s.addText("11 / 12", { x:9, y:0.3, w:0.8, h:0.4, fontSize:10, color:G3, align:"right" });

    // 融资核心信息
    s.addShape(pres.shapes.RECTANGLE, { x:0.4, y:1.2, w:9.2, h:0.75, fill:{color:G2}, line:{color:G2} });
    s.addText([
      { text:"天使轮  ", options:{ color:WH, fontSize:13 } },
      { text:"500万人民币", options:{ color:WH, fontSize:20, bold:true } },
      { text:"  ·  出让  ", options:{ color:WH, fontSize:13 } },
      { text:"10%", options:{ color:WH, fontSize:20, bold:true } },
      { text:" 股权  ·  估值  ", options:{ color:WH, fontSize:13 } },
      { text:"5,000万元", options:{ color:WH, fontSize:20, bold:true } },
    ], { x:0.6, y:1.28, w:8.8, h:0.58, align:"center", valign:"middle" });

    // 资金用途 4块
    const funds = [
      { pct:"30%", amt:"150万", use:"产品与技术团队", color:"1D5C3A" },
      { pct:"20%", amt:"100万", use:"线下渠道拓展", color:G1 },
      { pct:"20%", amt:"100万", use:"食材供应链合作", color:"2C7A55" },
      { pct:"30%", amt:"150万", use:"营养师+市场推广", color:G2 },
    ];
    funds.forEach((f, i) => {
      const x = 0.4 + i * 2.33;
      s.addShape(pres.shapes.RECTANGLE, { x, y:2.1, w:2.1, h:1.3, fill:{color:f.color}, line:{color:f.color}, shadow:shadow() });
      s.addText(f.pct, { x, y:2.15, w:2.1, h:0.5, fontSize:26, bold:true, color:WH, fontFace:"Arial Black", align:"center" });
      s.addText(f.amt, { x, y:2.65, w:2.1, h:0.35, fontSize:14, bold:true, color:G3, fontFace:"Arial", align:"center" });
      s.addText(f.use, { x:x+0.05, y:3.0, w:2.0, h:0.35, fontSize:10, color:WH, fontFace:"Arial", align:"center" });
    });

    // 时间路线
    const phases = [
      { phase:"0-3个月", key:"搭班底·跑MVP", color:G1 },
      { phase:"3-6个月", key:"月子中心·种子用户500+", color:"2C7A55" },
      { phase:"6-12个月", key:"医院渠道·付费破5千", color:G2 },
      { phase:"12个月+", key:"A轮·付费用户5万", color:"1A5C35" },
    ];
    s.addShape(pres.shapes.LINE, { x:0.75, y:4.1, w:8.5, h:0, line:{color:G3,width:2} });
    phases.forEach((p, i) => {
      const x = 0.5 + i * 2.25;
      s.addShape(pres.shapes.OVAL, { x:x+0.6, y:3.85, w:0.5, h:0.5, fill:{color:p.color}, line:{color:p.color} });
      s.addText(p.phase, { x, y:4.42, w:2.1, h:0.35, fontSize:9, bold:true, color:G1, fontFace:"Arial", align:"center" });
      s.addText(p.key, { x:x-0.05, y:4.78, w:2.2, h:0.55, fontSize:9, color:GY, fontFace:"Arial", align:"center" });
    });
  }

  // ─────────────────────────────────────────
  // S12: 结尾联系
  // ─────────────────────────────────────────
  {
    let s = pres.addSlide();
    s.background = { color: G1 };

    s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:10, h:5.625, fill:{color:"164E37"}, line:{color:"164E37"} });
    s.addShape(pres.shapes.RECTANGLE, { x:6.8, y:0, w:3.2, h:5.625, fill:{color:G1}, line:{color:G1} });

    s.addText("🥗", { x:7.0, y:0.5, w:2.8, h:2.2, fontSize:96, align:"center" });

    s.addText("好好吃饭", { x:0.6, y:0.8, w:6, h:1.1, fontSize:52, bold:true, color:WH, fontFace:"Arial Black" });
    s.addText("天天向上", { x:0.6, y:1.8, w:5, h:0.7, fontSize:28, color:G3, fontFace:"Arial", charSpacing:8 });
    s.addText("不是外卖App，是健康行为改变平台", { x:0.6, y:2.55, w:6, h:0.5, fontSize:15, color:"8ECFB0", fontFace:"Arial" });

    s.addShape(pres.shapes.LINE, { x:0.6, y:3.2, w:5.8, h:0, line:{color:G3,width:1} });

    const contacts = [
      "👤  张越",
      "📞  18810409001 / 13288199969",
      "📧  zhangyue0358@126.com",
      "📱  Demo：haohaochifan-1.netlify.app",
      "📍  北京",
    ];
    contacts.forEach((c, i) => {
      s.addText(c, { x:0.6, y:3.35 + i*0.42, w:5.8, h:0.38, fontSize:c.includes("Demo")? 13:12, color:i===3?"6DCCA0":WH, fontFace:"Arial", bold:i===3 });
    });

    s.addText("感谢您的时间\n期待与您进一步交流", {
      x:6.8, y:3.2, w:2.8, h:1.6, fontSize:14, color:G3, fontFace:"Arial", align:"center", italic:true
    });
    s.addText("12 / 12", { x:9, y:5.1, w:0.8, h:0.35, fontSize:10, color:G3, align:"right" });
  }

  await pres.writeFile({ fileName:"/mnt/user-data/outputs/好好吃饭_融资路演PPT.pptx" });
  console.log("PPT 生成完成！");
}

run().catch(e => { console.error(e); process.exit(1); });
