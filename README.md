# 好好吃饭 · 项目文件包

> 好好吃饭 天天向上 · 不是外卖App，是健康行为改变平台

---

## 📁 目录说明

```
haohao-chifan/
├── output/               # 直接使用的成品文件
│   ├── App_Demo_v2.html  # App 交互 Demo（双击浏览器打开）
│   ├── BP_v2.docx        # 融资计划书（Word格式）
│   └── PPT.pptx          # 融资路演PPT（11张）
│
├── source/               # 可重新生成 BP 和 PPT 的源代码
│   ├── bp.js             # 生成融资计划书 Word 文档
│   └── ppt.js            # 生成融资路演 PPT
│
└── README.md             # 本说明文件
```

---

## 🚀 直接使用（无需安装）

### App Demo
双击 `output/App_Demo_v2.html`，用浏览器打开即可。

也可以上传到 Netlify 在线分享：
1. 打开 https://drop.netlify.com
2. 把 `App_Demo_v2.html` 改名为 `index.html`
3. 拖进去即可获得分享链接

### 融资计划书 & PPT
直接用 Word / WPS / PowerPoint 打开 output/ 里的文件。

---

## 🔧 重新生成文件（需要 Node.js 环境）

如果需要修改内容后重新生成，按以下步骤操作：

### 第一步：安装 Node.js
- 下载地址：https://nodejs.org（推荐 LTS 版本）

### 第二步：安装依赖
打开终端（Windows 用 PowerShell / CMD，Mac 用 Terminal），在本文件夹内运行：

```bash
npm install -g docx pptxgenjs react react-dom sharp react-icons
```

### 第三步：修改内容
用任意文本编辑器（VSCode、记事本等）打开 source/ 里的 .js 文件修改内容。

### 第四步：重新生成

生成融资计划书：
```bash
node source/bp.js
# 输出文件：output/BP_v2.docx
```

生成路演 PPT：
```bash
node source/ppt.js
# 输出文件：output/PPT.pptx（路径在代码末尾 writeFile 里）
```

---

## 📱 在线 Demo 链接

https://haohaochifan-1.netlify.app

---

## 📞 联系方式

- 姓名：张越
- 电话：18810409001
- 邮箱：zhangyue0358@126.com
- 微信：18810409001
- 地址：北京
