const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  AlignmentType, HeadingLevel, BorderStyle, WidthType, ShadingType,
  LevelFormat, PageBreak, ExternalHyperlink, VerticalAlign,
  Header, Footer, PageNumber
} = require('docx');
const fs = require('fs');

const GREEN = '2BA76A';
const DARK_GREEN = '1A7A4C';
const LIGHT_GREEN = 'E8F7F0';
const WARN = 'FEF3E2';
const GRAY = 'F5F5F0';
const WHITE = 'FFFFFF';

const border = { style: BorderStyle.SINGLE, size: 1, color: 'D0D0D0' };
const borders = { top: border, bottom: border, left: border, right: border };
const noBorder = { style: BorderStyle.NONE, size: 0, color: 'FFFFFF' };
const noBorders = { top: noBorder, bottom: noBorder, left: noBorder, right: noBorder };

const cm = (n) => Math.round(n * 567); // cm to DXA (1cm = 567 DXA approx)

function h1(text) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    children: [new TextRun({ text, bold: true, color: DARK_GREEN, size: 36, font: 'Arial' })],
    spacing: { before: 400, after: 200 },
    border: { bottom: { style: BorderStyle.SINGLE, size: 8, color: GREEN, space: 4 } }
  });
}

function h2(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, color: DARK_GREEN, size: 28, font: 'Arial' })],
    spacing: { before: 320, after: 120 }
  });
}

function h3(text) {
  return new Paragraph({
    children: [new TextRun({ text, bold: true, color: '333333', size: 24, font: 'Arial' })],
    spacing: { before: 200, after: 80 }
  });
}

function p(text, opts = {}) {
  return new Paragraph({
    children: [new TextRun({ text, size: 22, font: 'Arial', color: '333333', ...opts })],
    spacing: { before: 60, after: 60 },
    alignment: opts.center ? AlignmentType.CENTER : AlignmentType.LEFT
  });
}

function bullet(text, bold_part = '') {
  const runs = [];
  if (bold_part && text.includes(bold_part)) {
    const idx = text.indexOf(bold_part);
    if (idx > 0) runs.push(new TextRun({ text: text.substring(0, idx), size: 22, font: 'Arial', color: '333333' }));
    runs.push(new TextRun({ text: bold_part, size: 22, font: 'Arial', bold: true, color: DARK_GREEN }));
    const rest = text.substring(idx + bold_part.length);
    if (rest) runs.push(new TextRun({ text: rest, size: 22, font: 'Arial', color: '333333' }));
  } else {
    runs.push(new TextRun({ text, size: 22, font: 'Arial', color: '333333' }));
  }
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    children: runs,
    spacing: { before: 40, after: 40 }
  });
}

function divider() {
  return new Paragraph({
    border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E0E0E0', space: 1 } },
    spacing: { before: 200, after: 200 }
  });
}

function pageBreak() {
  return new Paragraph({ children: [new PageBreak()] });
}

function greenBox(title, content) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [9360],
    rows: [
      new TableRow({ children: [
        new TableCell({
          borders,
          width: { size: 9360, type: WidthType.DXA },
          shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR },
          margins: { top: 120, bottom: 120, left: 180, right: 180 },
          children: [
            new Paragraph({ children: [new TextRun({ text: title, bold: true, size: 24, color: DARK_GREEN, font: 'Arial' })], spacing: { before: 40, after: 60 } }),
            new Paragraph({ children: [new TextRun({ text: content, size: 21, color: '333333', font: 'Arial' })], spacing: { before: 0, after: 40 } })
          ]
        })
      ]})
    ]
  });
}

function twoCol(left_title, left_content, right_title, right_content, leftColor = LIGHT_GREEN, rightColor = WARN) {
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: [4620, 4740],
    rows: [
      new TableRow({ children: [
        new TableCell({
          borders,
          width: { size: 4620, type: WidthType.DXA },
          shading: { fill: leftColor, type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 160, right: 160 },
          children: [
            new Paragraph({ children: [new TextRun({ text: left_title, bold: true, size: 22, color: DARK_GREEN, font: 'Arial' })], spacing: { before: 40, after: 60 } }),
            new Paragraph({ children: [new TextRun({ text: left_content, size: 20, color: '333333', font: 'Arial' })], spacing: { before: 0, after: 40 } })
          ]
        }),
        new TableCell({
          borders,
          width: { size: 4740, type: WidthType.DXA },
          shading: { fill: rightColor, type: ShadingType.CLEAR },
          margins: { top: 100, bottom: 100, left: 160, right: 160 },
          children: [
            new Paragraph({ children: [new TextRun({ text: right_title, bold: true, size: 22, color: '92500E', font: 'Arial' })], spacing: { before: 40, after: 60 } }),
            new Paragraph({ children: [new TextRun({ text: right_content, size: 20, color: '333333', font: 'Arial' })], spacing: { before: 0, after: 40 } })
          ]
        })
      ]})
    ]
  });
}

function compTable(headers, rows) {
  const colW = Math.floor(9360 / headers.length);
  const colWidths = headers.map(() => colW);
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: colWidths,
    rows: [
      new TableRow({
        tableHeader: true,
        children: headers.map((h, i) => new TableCell({
          borders,
          width: { size: colW, type: WidthType.DXA },
          shading: { fill: i === 0 ? DARK_GREEN : GREEN, type: ShadingType.CLEAR },
          margins: { top: 80, bottom: 80, left: 120, right: 120 },
          children: [new Paragraph({ children: [new TextRun({ text: h, bold: true, size: 20, color: WHITE, font: 'Arial' })], alignment: AlignmentType.CENTER })]
        }))
      }),
      ...rows.map((row, ri) => new TableRow({
        children: row.map((cell, ci) => new TableCell({
          borders,
          width: { size: colW, type: WidthType.DXA },
          shading: { fill: ri % 2 === 0 ? 'FAFAFA' : WHITE, type: ShadingType.CLEAR },
          margins: { top: 70, bottom: 70, left: 120, right: 120 },
          children: [new Paragraph({
            children: [new TextRun({
              text: cell,
              size: 19,
              font: 'Arial',
              color: cell === '✓' ? DARK_GREEN : cell === '✗' ? 'CC0000' : cell === '△' ? '854F0B' : '333333',
              bold: cell === '✓' || cell === '✗'
            })],
            alignment: ci > 0 ? AlignmentType.CENTER : AlignmentType.LEFT
          })]
        }))
      }))
    ]
  });
}

function statBox(stats) {
  // stats: [{num, label}]
  const colW = Math.floor(9360 / stats.length);
  return new Table({
    width: { size: 9360, type: WidthType.DXA },
    columnWidths: stats.map(() => colW),
    rows: [
      new TableRow({ children: stats.map(s => new TableCell({
        borders: noBorders,
        width: { size: colW, type: WidthType.DXA },
        shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR },
        margins: { top: 140, bottom: 140, left: 80, right: 80 },
        verticalAlign: VerticalAlign.CENTER,
        children: [
          new Paragraph({ children: [new TextRun({ text: s.num, bold: true, size: 52, color: DARK_GREEN, font: 'Arial' })], alignment: AlignmentType.CENTER, spacing: { before: 0, after: 40 } }),
          new Paragraph({ children: [new TextRun({ text: s.label, size: 18, color: '555555', font: 'Arial' })], alignment: AlignmentType.CENTER })
        ]
      }))}),
    ]
  });
}

// ─── Main Document ───────────────────────────────────
const doc = new Document({
  numbering: {
    config: [
      { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 240 } } } }] },
      { reference: 'nums', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 480, hanging: 240 } } } }] }
    ]
  },
  styles: {
    default: { document: { run: { font: 'Arial', size: 22 } } },
    paragraphStyles: [
      { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 36, bold: true, font: 'Arial', color: DARK_GREEN },
        paragraph: { spacing: { before: 400, after: 200 }, outlineLevel: 0 } },
      { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
        run: { size: 28, bold: true, font: 'Arial', color: DARK_GREEN },
        paragraph: { spacing: { before: 320, after: 120 }, outlineLevel: 1 } },
    ]
  },
  sections: [{
    properties: {
      page: {
        size: { width: 11906, height: 16838 }, // A4
        margin: { top: 1080, right: 1080, bottom: 1080, left: 1080 }
      }
    },
    headers: {
      default: new Header({
        children: [new Paragraph({
          children: [
            new TextRun({ text: '好好吃饭  天天向上', size: 18, color: GREEN, font: 'Arial', bold: true }),
            new TextRun({ text: '   |   项目融资计划书 v2.0', size: 18, color: '999999', font: 'Arial' })
          ],
          border: { bottom: { style: BorderStyle.SINGLE, size: 4, color: 'E0E0E0', space: 4 } }
        })]
      })
    },
    footers: {
      default: new Footer({
        children: [new Paragraph({
          children: [
            new TextRun({ text: 'Demo体验链接：haohaochifan-1.netlify.app   |   联系：张越  18810409001', size: 16, color: '999999', font: 'Arial' }),
          ],
          alignment: AlignmentType.CENTER,
          border: { top: { style: BorderStyle.SINGLE, size: 4, color: 'E0E0E0', space: 4 } }
        })]
      })
    },
    children: [

      // ══════════════ 封面 ══════════════
      new Paragraph({ spacing: { before: 1200 } }),
      new Paragraph({
        children: [new TextRun({ text: '好好吃饭', bold: true, size: 72, color: DARK_GREEN, font: 'Arial' })],
        alignment: AlignmentType.CENTER
      }),
      new Paragraph({
        children: [new TextRun({ text: '天天向上', size: 40, color: GREEN, font: 'Arial' })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 240 }
      }),
      new Paragraph({
        children: [new TextRun({ text: '不是外卖App，是健康行为改变平台', size: 28, color: '555555', font: 'Arial' })],
        alignment: AlignmentType.CENTER,
        spacing: { after: 480 }
      }),
      new Table({
        width: { size: 6000, type: WidthType.DXA },
        columnWidths: [6000],
        rows: [new TableRow({ children: [new TableCell({
          borders,
          width: { size: 6000, type: WidthType.DXA },
          shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR },
          margins: { top: 160, bottom: 160, left: 240, right: 240 },
          children: [
            new Paragraph({ children: [new TextRun({ text: '融资需求：天使轮 500万人民币，出让 10% 股权', size: 24, bold: true, color: DARK_GREEN, font: 'Arial' })], alignment: AlignmentType.CENTER, spacing: { before: 40, after: 40 } }),
            new Paragraph({ children: [new TextRun({ text: '项目估值：5,000万元（Pre-Money）', size: 22, color: '333333', font: 'Arial' })], alignment: AlignmentType.CENTER }),
          ]
        })]})],
        margins: { left: cm(2), right: cm(2) }
      }),
      new Paragraph({ spacing: { before: 480 } }),
      new Paragraph({
        children: [
          new TextRun({ text: '📱 App Demo 在线体验：', size: 22, font: 'Arial', color: '555555' }),
          new ExternalHyperlink({
            link: 'https://haohaochifan-1.netlify.app',
            children: [new TextRun({ text: 'haohaochifan-1.netlify.app', size: 22, font: 'Arial', color: GREEN, underline: {} })]
          })
        ],
        alignment: AlignmentType.CENTER,
        spacing: { before: 80, after: 40 }
      }),
      new Paragraph({
        children: [new TextRun({ text: '联系人：张越   |   18810409001   |   北京', size: 20, color: '888888', font: 'Arial' })],
        alignment: AlignmentType.CENTER,
        spacing: { before: 40 }
      }),
      pageBreak(),

      // ══════════════ 执行摘要 ══════════════
      h1('01  执行摘要'),
      p('好好吃饭是一个以健康档案为核心驱动的餐饮行为改变平台。我们不屏蔽任何食物，只做精准标识——让用户在点餐时看到「今天钠快够了」「✓ 挺适合你」这样的朋友式提示，帮助4亿慢病患者、1000万年孕产妇及更广泛的健康人群，在不改变点餐习惯的前提下，逐步养成更好的饮食行为。'),
      new Paragraph({ spacing: { before: 160 } }),
      statBox([
        { num: '4亿+', label: '国内慢病患者' },
        { num: '1000万', label: '年孕产妇数量' },
        { num: '7000亿+', label: '外卖市场规模' },
        { num: '0', label: '竞品有同类方案' }
      ]),
      new Paragraph({ spacing: { before: 160 } }),
      h3('核心差异：健康档案驱动的餐饮闭环'),
      p('市场上现有的饿了么、美团外卖只解决「吃什么」，薄荷健康、Keep只解决「记什么」，没有任何一家把「健康档案→AI分析→菜品标识→用户点餐→数据沉淀」这条链路打通。好好吃饭的护城河是用户的365天饮食数据——越用越精准，不可迁移。'),
      new Paragraph({ spacing: { before: 80 } }),
      twoCol(
        '✅ 我们做的',
        '基于用户健康档案，对每道菜进行个性化标识，只做提示不做限制，帮用户在知情状态下做选择，推动行为习惯改变。',
        '❌ 我们不做的',
        '不成为外卖平台（轻资产接入），不替用户决策（选择权永远在用户），不追求DAU而是追求用户真实健康改善。'
      ),
      pageBreak(),

      // ══════════════ 问题与市场 ══════════════
      h1('02  问题与市场机会'),
      h2('2.1 用户的真实痛点'),
      p('中国有4亿慢病患者，每年约1000万孕产妇，以及数以亿计想吃健康但不知道怎么选的普通人。他们面临的不是「买不到健康食物」的问题，而是：'),
      new Paragraph({ spacing: { before: 80 } }),
      bullet('不知道自己应该吃什么', '不知道自己应该吃什么'),
      bullet('外卖平台无法识别哪些菜适合自己的身体状况'),
      bullet('特殊人群（孕妇、高血压、痛风患者）完全得不到针对性引导'),
      bullet('知道要吃健康，但习惯难以坚持'),
      bullet('食材质量和配送安全没有保障'),
      new Paragraph({ spacing: { before: 80 } }),
      greenBox(
        '关键洞察：问题不在于「没有健康食物」，在于「信息不对称」',
        '用户不缺健康餐厅，缺的是一个能根据自己健康状况告诉他「这道菜今天适不适合吃」的系统。好好吃饭要做的，就是把营养师的专业判断，用一个友善的标识系统，嵌入用户的每一次点餐行为。'
      ),
      new Paragraph({ spacing: { before: 160 } }),
      h2('2.2 市场规模'),
      p('好好吃饭所在的市场，可以从三个维度来看：'),
      new Paragraph({ spacing: { before: 80 } }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2200, 2800, 4360],
        rows: [
          new TableRow({ tableHeader: true, children: [
            new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, shading: { fill: DARK_GREEN, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '市场层级', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, width: { size: 2800, type: WidthType.DXA }, shading: { fill: DARK_GREEN, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '规模', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, shading: { fill: DARK_GREEN, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '说明', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, shading: { fill: 'FAFAFA', type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'TAM · 总市场', size: 20, font: 'Arial', color: '333333' })] })] }),
            new TableCell({ borders, width: { size: 2800, type: WidthType.DXA }, shading: { fill: 'FAFAFA', type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '中国外卖市场 ~1.2万亿', size: 20, font: 'Arial', bold: true, color: DARK_GREEN })] })] }),
            new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, shading: { fill: 'FAFAFA', type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '2024年中国在线外卖市场规模，用户超6亿，CAGR约15%', size: 20, font: 'Arial', color: '333333' })] })] }),
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'SAM · 目标市场', size: 20, font: 'Arial', color: '333333' })] })] }),
            new TableCell({ borders, width: { size: 2800, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '健康饮食用户 ~800亿', size: 20, font: 'Arial', bold: true, color: DARK_GREEN })] })] }),
            new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '慢病患者、孕产妇、中老年、注重健康的年轻白领', size: 20, font: 'Arial', color: '333333' })] })] }),
          ]}),
          new TableRow({ children: [
            new TableCell({ borders, width: { size: 2200, type: WidthType.DXA }, shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: 'SOM · 可获市场', size: 20, font: 'Arial', bold: true, color: DARK_GREEN })] })] }),
            new TableCell({ borders, width: { size: 2800, type: WidthType.DXA }, shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '3年目标 ~15亿', size: 20, font: 'Arial', bold: true, color: DARK_GREEN })] })] }),
            new TableCell({ borders, width: { size: 4360, type: WidthType.DXA }, shading: { fill: LIGHT_GREEN, type: ShadingType.CLEAR }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '北京/上海/深圳切入，3年内覆盖100万付费用户', size: 20, font: 'Arial', color: DARK_GREEN })] })] }),
          ]}),
        ]
      }),
      pageBreak(),

      // ══════════════ 产品与解决方案 ══════════════
      h1('03  产品与解决方案'),
      h2('3.1 产品定位'),
      p('好好吃饭不是外卖App，是一个以「健康档案」为核心的饮食行为改变平台。外卖是载体，健康改变是目的。'),
      new Paragraph({ spacing: { before: 80 } }),
      greenBox(
        '核心产品理念：标识而非限制，选择权永远在用户手里',
        '当用户点餐时，我们不会屏蔽任何菜品，只会用朋友的口吻做提示——「今天钠快够了」「尿酸高时少吃」「✓ 挺适合你」。知道了，然后选，这是产品真正帮助用户建立长期习惯的方式。'
      ),
      new Paragraph({ spacing: { before: 160 } }),
      h2('3.2 核心功能模块'),
      new Paragraph({ spacing: { before: 80 } }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1800, 3780, 3780],
        rows: [
          new TableRow({ tableHeader: true, children: [
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 1800, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '模块', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 3780, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '功能说明', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 3780, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '核心价值', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
          ]}),
          ...[
            ['🧬 健康档案', '用户上传体检报告、病历，AI解读+营养师人工复核（5%高风险人群人工审核），生成个性化营养方案', '数据护城河，越用越精准，不可迁移'],
            ['🏷️ 菜品标识', '每道菜基于用户健康档案实时计算100分制健康评分，分四维度（营养匹配/禁忌/烹饪/今日状态），点击?查看详细说明', '差异化核心，同一道菜对不同人分数不同'],
            ['🍽️ 智能推荐', '结合今日已摄入数据动态调整推荐，今天缺什么补什么，早晚餐联动', '越用越懂你的AI引擎'],
            ['👨‍👩‍👧 家庭账号', '最多6位成员各自独立健康档案，AI分别生成推荐，支持一键全家下单', '提升客单价，增强家庭黏性'],
            ['🤰 特殊人群', '孕妇按孕周定制（8/16/20/28/36周），慢病患者支持高血压/糖尿病/痛风多病种联合管理', '高价值用户深度运营'],
            ['📊 饮食记录', '外卖自动同步饮食记录，手动添加兜底，周报展示营养趋势', '习惯养成的数据支撑'],
          ].map(([m, f, v], i) => new TableRow({ children: [
            new TableCell({ borders, shading: { fill: i%2===0 ? 'FAFAFA' : WHITE, type: ShadingType.CLEAR }, width: { size: 1800, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: m, size: 19, font: 'Arial', bold: true, color: DARK_GREEN })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? 'FAFAFA' : WHITE, type: ShadingType.CLEAR }, width: { size: 3780, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: f, size: 19, font: 'Arial', color: '333333' })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? 'FAFAFA' : WHITE, type: ShadingType.CLEAR }, width: { size: 3780, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: v, size: 19, font: 'Arial', color: DARK_GREEN })] })] }),
          ]}))
        ]
      }),
      new Paragraph({ spacing: { before: 160 } }),
      h2('3.3 App Demo'),
      p('产品交互原型已完成，涵盖注册流程、首页、点外卖、健康评分说明、特殊人群（孕妇/慢病）、家庭账号、下单配送等全链路，可在线实时体验：'),
      new Paragraph({
        children: [
          new TextRun({ text: '🔗 ', size: 24, font: 'Arial' }),
          new ExternalHyperlink({
            link: 'https://haohaochifan-1.netlify.app',
            children: [new TextRun({ text: 'haohaochifan-1.netlify.app', size: 24, font: 'Arial', color: GREEN, underline: {}, bold: true })]
          }),
          new TextRun({ text: '  （手机或电脑浏览器直接打开）', size: 22, font: 'Arial', color: '888888' })
        ],
        spacing: { before: 80, after: 80 }
      }),
      pageBreak(),

      // ══════════════ 目标用户 ══════════════
      h1('04  目标用户'),
      p('好好吃饭的用户分三个优先级切入，从需求最强烈、付费意愿最高的人群开始：'),
      new Paragraph({ spacing: { before: 80 } }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1600, 1800, 2400, 3560],
        rows: [
          new TableRow({ tableHeader: true, children: [
            new TableCell({ borders, shading: { fill: DARK_GREEN, type: ShadingType.CLEAR }, width: { size: 1600, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '优先级', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: DARK_GREEN, type: ShadingType.CLEAR }, width: { size: 1800, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '人群', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: DARK_GREEN, type: ShadingType.CLEAR }, width: { size: 2400, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '规模', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: DARK_GREEN, type: ShadingType.CLEAR }, width: { size: 3560, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '核心需求 & 获客渠道', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
          ]}),
          ...[
            ['P1 首批切入', '孕妇 / 产妇', '年约1000万人', '孕期按周营养指导，月子调理餐。渠道：月子中心合作、孕产APP投放'],
            ['P1 首批切入', '慢病患者', '4亿+（高血压1.3亿）', '多病种联合管理，禁忌食材精准标识。渠道：医院营养科、社区健康中心'],
            ['P2 规模化', '健康白领', '一线城市约1.5亿', '想吃健康但没时间研究，信任专业背书。渠道：企业健康福利合作'],
            ['P2 规模化', '中老年养生', '60岁以上约2.7亿', '子女代为下单，家庭账号切入。渠道：子女营销，家庭账号功能'],
            ['P3 长期', '术后调养/特殊饮食需求', '细分长尾人群', '高度定制化，高付费意愿。渠道：医院出院营养指导合作'],
          ].map(([p, u, s, n], i) => new TableRow({ children: [
            new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT_GREEN : WHITE, type: ShadingType.CLEAR }, width: { size: 1600, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: p, size: 18, font: 'Arial', color: DARK_GREEN, bold: true })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT_GREEN : WHITE, type: ShadingType.CLEAR }, width: { size: 1800, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: u, size: 18, font: 'Arial', bold: true, color: '333333' })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT_GREEN : WHITE, type: ShadingType.CLEAR }, width: { size: 2400, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: s, size: 18, font: 'Arial', color: '333333' })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT_GREEN : WHITE, type: ShadingType.CLEAR }, width: { size: 3560, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: n, size: 18, font: 'Arial', color: '333333' })] })] }),
          ]}))
        ]
      }),
      pageBreak(),

      // ══════════════ 竞争分析 ══════════════
      h1('05  竞争分析'),
      h2('5.1 竞争格局'),
      p('现有市场上没有任何一家把「健康档案→餐饮推荐→行为改变」完整打通。竞争对手各自只解决了一部分问题：'),
      new Paragraph({ spacing: { before: 80 } }),
      compTable(
        ['维度', '饿了么/美团', '薄荷健康/Keep', '好好吃饭'],
        [
          ['核心定位', '外卖配送平台', '饮食记录/卡路里计算', '健康行为改变平台'],
          ['健康档案接入', '✗', '△ 部分', '✓ 核心驱动'],
          ['个性化菜品标识', '✗', '✗', '✓ 实时计算'],
          ['营养师审核', '✗', '✗', '✓ 5%高风险人工'],
          ['慢病/孕妇专属', '✗', '△ 有限支持', '✓ 多病种联合管理'],
          ['外卖+健康闭环', '✗（只有外卖）', '✗（只有记录）', '✓ 完整闭环'],
          ['家庭账号', '△ 基础', '✗', '✓ 各自独立档案'],
          ['数据护城河', '△ 订单历史', '△ 饮食记录', '✓ 健康档案+饮食数据'],
        ]
      ),
      new Paragraph({ spacing: { before: 160 } }),
      h2('5.2 我们的护城河'),
      new Paragraph({ spacing: { before: 80 } }),
      bullet('数据飞轮：用户越用，健康数据越丰富，推荐越精准，用户越难离开。365天饮食+健康记录不可迁移。', '数据飞轮'),
      bullet('营养师网络：与注册营养师建立合作，形成专业背书，是竞品无法快速复制的信任资产。', '营养师网络'),
      bullet('特殊人群壁垒：孕妇、慢病患者的使用场景具有强需求、长周期、高付费意愿，且愿意口碑传播。', '特殊人群壁垒'),
      bullet('商家生态：好好吃饭为商家带来的是有健康需求的高质量精准用户，转化率和复购率优于泛流量平台。', '商家生态'),
      pageBreak(),

      // ══════════════ 商业模式 ══════════════
      h1('06  商业模式'),
      h2('6.1 收入来源'),
      new Paragraph({ spacing: { before: 80 } }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2200, 2400, 2400, 2360],
        rows: [
          new TableRow({ tableHeader: true, children: [
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 2200, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '收入来源', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 2400, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '定价逻辑', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 2400, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '启动时间', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 2360, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '3年目标规模', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
          ]}),
          ...[
            ['用户会员订阅', '19元/月 · 199元/年，特殊人群（孕妇/慢病）专属档案功能', '第1年', '100万付费用户 → 2亿/年'],
            ['商家平台服务费', '按GMV抽佣2-5% + 健康认证年费5000-2万/年', '第1年', '接入500家商家 → 5000万/年'],
            ['品牌加盟费+食材供应', '加盟费10-30万 + 食材供应链服务费', '第2年', '200家加盟店 → 3000万/年'],
            ['企业健康福利合作', '企业团购年费，5-20元/人/月', '第2年', '1000家企业 → 2000万/年'],
            ['竞价广告/精准推荐', '健康相关品牌定向广告，按CPM/CPC计费', '第3年', '500万/年'],
          ].map(([t, d, s, g], i) => new TableRow({ children: [
            new TableCell({ borders, shading: { fill: i%2===0 ? 'FAFAFA' : WHITE, type: ShadingType.CLEAR }, width: { size: 2200, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: t, size: 19, font: 'Arial', bold: true, color: DARK_GREEN })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? 'FAFAFA' : WHITE, type: ShadingType.CLEAR }, width: { size: 2400, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: d, size: 19, font: 'Arial', color: '333333' })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? 'FAFAFA' : WHITE, type: ShadingType.CLEAR }, width: { size: 2400, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: s, size: 19, font: 'Arial', color: '555555' })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? 'FAFAFA' : WHITE, type: ShadingType.CLEAR }, width: { size: 2360, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: g, size: 19, font: 'Arial', color: DARK_GREEN, bold: true })] })] }),
          ]}))
        ]
      }),
      new Paragraph({ spacing: { before: 160 } }),
      h2('6.2 冷启动破局策略'),
      bullet('第0-3个月：不追用户量，只服务100个种子用户，做到极致，用真实身体改善数据做内容素材', '不追用户量'),
      bullet('第3-6个月：聚焦孕妇群体，与3-5家月子中心深度合作，用口碑驱动裂变', '聚焦孕妇群体'),
      bullet('第6-12个月：切入慢病渠道，与医院营养科合作，医生推荐比任何广告都有效', '切入慢病渠道'),
      bullet('第12个月+：开放企业健康福利B2B，一次签约带来几百个高质量用户', '开放企业健康福利'),
      pageBreak(),

      // ══════════════ 增长策略 ══════════════
      h1('07  市场推广与增长策略'),
      h2('7.1 获客漏斗'),
      p('好好吃饭的增长逻辑是「精准渠道获客 → 极致首周体验 → 数据证明价值 → 口碑自然裂变」，不依赖烧钱买流量。'),
      new Paragraph({ spacing: { before: 80 } }),
      twoCol(
        '📍 线下精准渠道（初期）',
        '月子中心：全国约1万家，孕妇是最理想种子用户\n\n医院营养科：直接面对慢病患者，医生背书效果强\n\n社区卫生服务中心：触达中老年慢病管理人群',
        '📱 线上内容渠道（规模化）',
        '真实用户健康改善数据内容（血压/血糖/体重变化）\n\n孕妈社区/宝妈圈口碑传播\n\n企业健康福利合作：一签带百人'
      ),
      new Paragraph({ spacing: { before: 160 } }),
      h2('7.2 留存与习惯养成'),
      bullet('第一周引导：注册后7天有引导任务，每完成一步给予正向反馈，提升首月留存'),
      bullet('营养改善可视化：每周推送「你的蛋白质摄入比上周提升18%」，让用户看见真实进步'),
      bullet('连续打卡机制：12天打卡视觉化呈现，社交分享功能强化习惯'),
      bullet('家庭账号黏性：一个家庭成员开始用，自然带动其他成员加入'),
      pageBreak(),

      // ══════════════ 财务预测 ══════════════
      h1('08  财务预测'),
      p('以下为保守预测，基于月子中心/医院渠道切入，前6个月以服务质量为主，不追用户规模。'),
      new Paragraph({ spacing: { before: 120 } }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [2100, 1815, 1815, 1815, 1815],
        rows: [
          new TableRow({ tableHeader: true, children: [
            new TableCell({ borders, shading: { fill: DARK_GREEN, type: ShadingType.CLEAR }, width: { size: 2100, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '指标', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            ...['第1年（种子期）', '第2年（起步期）', '第3年（增长期）', '第4年（规模期）'].map(t =>
              new TableCell({ borders, shading: { fill: DARK_GREEN, type: ShadingType.CLEAR }, width: { size: 1815, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: t, bold: true, size: 18, color: WHITE, font: 'Arial' })], alignment: AlignmentType.CENTER })] })
            )
          ]}),
          ...[
            ['付费用户数', '5,000', '50,000', '300,000', '1,000,000'],
            ['合作商家数', '30家', '200家', '500家', '2,000家'],
            ['总收入', '150万', '1,500万', '6,000万', '2亿'],
            ['  └ 会员订阅', '60万', '600万', '2,400万', '8,000万'],
            ['  └ 商家服务费', '60万', '700万', '2,800万', '9,000万'],
            ['  └ 其他', '30万', '200万', '800万', '3,000万'],
            ['运营成本', '380万', '900万', '2,800万', '8,000万'],
            ['净利润/亏损', '-230万', '+600万', '+3,200万', '+1.2亿'],
          ].map((row, i) => new TableRow({ children: [
            new TableCell({ borders, shading: { fill: row[0].startsWith('  ') ? WHITE : (i%2===0 ? 'FAFAFA' : WHITE), type: ShadingType.CLEAR }, width: { size: 2100, type: WidthType.DXA }, margins: { top: 65, bottom: 65, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: row[0], size: 18, font: 'Arial', color: row[0].startsWith('净利润') ? (row[4].startsWith('+') ? DARK_GREEN : 'CC0000') : '333333', bold: !row[0].startsWith('  ') })] })] }),
            ...row.slice(1).map(v => new TableCell({ borders, shading: { fill: row[0].startsWith('  ') ? WHITE : (i%2===0 ? 'FAFAFA' : WHITE), type: ShadingType.CLEAR }, width: { size: 1815, type: WidthType.DXA }, margins: { top: 65, bottom: 65, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: v, size: 18, font: 'Arial', color: v.startsWith('+') ? DARK_GREEN : v.startsWith('-') ? 'CC0000' : '333333', bold: v.startsWith('+') || v.startsWith('-') })], alignment: AlignmentType.CENTER })] }))
          ]}))
        ]
      }),
      new Paragraph({ spacing: { before: 80 } }),
      p('注：第1年底实现现金流盈亏平衡的关键指标是付费用户月留存率 > 70%，付费转化率 > 15%。', { color: '888888', italics: true }),
      pageBreak(),

      // ══════════════ 融资计划 ══════════════
      h1('09  融资计划与资金用途'),
      new Paragraph({ spacing: { before: 80 } }),
      greenBox(
        '本轮融资：天使轮 500万人民币，出让 10% 股权（估值5000万元）',
        '资金将用于：搭建核心团队、打磨产品MVP、拓展首批合作商家和渠道，目标在12个月内验证核心商业逻辑，为A轮做好准备。'
      ),
      new Paragraph({ spacing: { before: 160 } }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [3000, 1600, 4760],
        rows: [
          new TableRow({ tableHeader: true, children: [
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 3000, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '用途', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 1600, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '金额', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 4760, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '具体说明', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
          ]}),
          ...[
            ['产品与技术团队', '150万 (30%)', '招募产品经理×1、前后端工程师×2、AI算法工程师×1，完成MVP开发'],
            ['营养师团队与内容', '80万 (16%)', '签约5名注册营养师，建立营养数据库，完成健康档案审核体系'],
            ['线下渠道拓展', '100万 (20%)', '北京/上海月子中心合作BD，医院营养科渠道建立，种子用户获取'],
            ['食材供应链合作', '100万 (20%)', '与美菜/宋小菜等B2B平台对接，为合作商家建立健康食材采购渠道'],
            ['市场推广', '70万 (14%)', '内容营销（真实用户健康改善案例），品牌建设'],
          ].map(([u, a, d], i) => new TableRow({ children: [
            new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT_GREEN : WHITE, type: ShadingType.CLEAR }, width: { size: 3000, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: u, size: 19, font: 'Arial', bold: true, color: DARK_GREEN })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT_GREEN : WHITE, type: ShadingType.CLEAR }, width: { size: 1600, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: a, size: 19, font: 'Arial', bold: true, color: DARK_GREEN })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT_GREEN : WHITE, type: ShadingType.CLEAR }, width: { size: 4760, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: d, size: 19, font: 'Arial', color: '333333' })] })] }),
          ]}))
        ]
      }),
      new Paragraph({ spacing: { before: 160 } }),
      h2('A轮里程碑目标（融资后12-18个月）'),
      bullet('付费用户突破 50,000 人，月留存率 > 70%'),
      bullet('合作商家达到 200 家，月GMV > 1,000万'),
      bullet('孕妇/慢病用户的健康改善数据完整，可作为A轮核心论据'),
      bullet('收入跑到年化 1,500万，具备A轮 3,000-5,000万融资条件'),
      pageBreak(),

      // ══════════════ 执行路线图 ══════════════
      h1('10  执行路线图'),
      new Paragraph({ spacing: { before: 80 } }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1800, 7560],
        rows: [
          ...[
            ['第0-3个月\n（搭班底）', '• 组建核心团队（技术+产品+营养师）\n• 完成App MVP开发（4步注册、健康档案、菜品标识、点餐下单）\n• 北京签约第一批合作商家 10-20家\n• 服务100个种子用户，极致体验，收集真实反馈'],
            ['第3-6个月\n（跑模型）', '• 与3-5家月子中心深度合作，孕妇群体种子用户500+\n• AI推荐模型迭代，用真实数据优化健康评分准确性\n• 营养师审核闭环跑通，5%高风险人群人工审核标准化\n• 收集第一批「血压/血糖改善」的真实用户故事'],
            ['第6-12个月\n（找增长）', '• 切入医院营养科渠道，慢病用户破5,000\n• 家庭账号功能上线，客单价提升\n• 企业健康福利B2B试点，3-5家企业签约\n• 完成A轮融资材料准备，核心数据指标达标'],
            ['第12-24个月\n（规模化）', '• 付费用户破50,000，扩张至上海/深圳\n• 开放加盟，先2-3家直营跑通后开放品牌加盟\n• 供应链升级，「好好吃饭认证食材」体系建立\n• 完成A轮融资'],
          ].map(([t, d], i) => new TableRow({ children: [
            new TableCell({ borders, shading: { fill: i%2===0 ? DARK_GREEN : GREEN, type: ShadingType.CLEAR }, width: { size: 1800, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 120, right: 120 }, verticalAlign: VerticalAlign.CENTER, children: [new Paragraph({ children: [new TextRun({ text: t, size: 19, font: 'Arial', bold: true, color: WHITE })], alignment: AlignmentType.CENTER })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? 'FAFAFA' : WHITE, type: ShadingType.CLEAR }, width: { size: 7560, type: WidthType.DXA }, margins: { top: 100, bottom: 100, left: 160, right: 160 }, children: d.split('\n').map(line => new Paragraph({ children: [new TextRun({ text: line, size: 19, font: 'Arial', color: '333333' })], spacing: { before: 30, after: 30 } })) }),
          ]}))
        ]
      }),
      pageBreak(),

      // ══════════════ 团队 ══════════════
      h1('11  创始团队'),
      new Paragraph({ spacing: { before: 80 } }),
      greenBox(
        '我们需要的团队：找到相信这件事的人',
        '好好吃饭现由创始人张越主导。本轮融资完成后，将围绕产品/技术/运营三个方向补充核心合伙人。以下是我们正在寻找的人。'
      ),
      new Paragraph({ spacing: { before: 160 } }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [1600, 2400, 5360],
        rows: [
          new TableRow({ tableHeader: true, children: [
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 1600, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '角色', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 2400, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '现状', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
            new TableCell({ borders, shading: { fill: GREEN, type: ShadingType.CLEAR }, width: { size: 5360, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: '寻找的人', bold: true, size: 20, color: WHITE, font: 'Arial' })] })] }),
          ]}),
          ...[
            ['创始人/CEO', '张越（在位）', '产品和商业方向负责人，主导产品设计和商业拓展，北京'],
            ['CTO/技术合伙人', '招募中', '有App从0到1经验，熟悉AI推荐系统，有医疗健康领域背景优先'],
            ['营养师顾问', '洽谈中2名', '注册营养师，有临床经验，认可通过科技改变大众健康饮食习惯'],
            ['运营合伙人', '招募中', '有医疗/健康/餐饮渠道资源，擅长B2B合作，能推动月子中心/医院渠道落地'],
          ].map(([r, s, d], i) => new TableRow({ children: [
            new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT_GREEN : WHITE, type: ShadingType.CLEAR }, width: { size: 1600, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: r, size: 19, font: 'Arial', bold: true, color: DARK_GREEN })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT_GREEN : WHITE, type: ShadingType.CLEAR }, width: { size: 2400, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: s, size: 19, font: 'Arial', color: '333333' })] })] }),
            new TableCell({ borders, shading: { fill: i%2===0 ? LIGHT_GREEN : WHITE, type: ShadingType.CLEAR }, width: { size: 5360, type: WidthType.DXA }, margins: { top: 70, bottom: 70, left: 120, right: 120 }, children: [new Paragraph({ children: [new TextRun({ text: d, size: 19, font: 'Arial', color: '333333' })] })] }),
          ]}))
        ]
      }),
      pageBreak(),

      // ══════════════ 结尾 ══════════════
      h1('12  联系我们'),
      new Paragraph({ spacing: { before: 160 } }),
      new Table({
        width: { size: 9360, type: WidthType.DXA },
        columnWidths: [9360],
        rows: [new TableRow({ children: [new TableCell({
          borders,
          width: { size: 9360, type: WidthType.DXA },
          shading: { fill: DARK_GREEN, type: ShadingType.CLEAR },
          margins: { top: 240, bottom: 240, left: 360, right: 360 },
          children: [
            new Paragraph({ children: [new TextRun({ text: '好好吃饭  天天向上', bold: true, size: 40, color: WHITE, font: 'Arial' })], alignment: AlignmentType.CENTER, spacing: { before: 0, after: 120 } }),
            new Paragraph({ children: [new TextRun({ text: '不是外卖App，是健康行为改变平台', size: 22, color: 'C2EDD8', font: 'Arial' })], alignment: AlignmentType.CENTER, spacing: { before: 0, after: 200 } }),
            new Paragraph({ children: [new TextRun({ text: '📱 App Demo：haohaochifan-1.netlify.app', size: 22, color: WHITE, font: 'Arial' })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }),
            new Paragraph({ children: [new TextRun({ text: '👤 张越  |  📞 18810409001  |  📧 zhangyue0358@126.com', size: 22, color: WHITE, font: 'Arial' })], alignment: AlignmentType.CENTER, spacing: { after: 80 } }),
            new Paragraph({ children: [new TextRun({ text: '💬 微信：18810409001  |  📍 北京', size: 22, color: WHITE, font: 'Arial' })], alignment: AlignmentType.CENTER, spacing: { after: 0 } }),
          ]
        })]})],
      }),
      new Paragraph({ spacing: { before: 240 } }),
      p('感谢您阅读好好吃饭的融资计划书。我们相信，帮助中国4亿慢病患者吃得更科学、更放心，是一件值得用10年时间去做的事。如果您认可这个方向，我们期待与您的进一步交流。', { color: '555555' }),

    ]
  }]
});

Packer.toBuffer(doc).then(buf => {
  fs.writeFileSync('/mnt/user-data/outputs/好好吃饭_融资计划书_v2.docx', buf);
  console.log('Done! Size:', buf.length, 'bytes');
}).catch(e => { console.error(e); process.exit(1); });
