
const fs = require("fs");
const {
  Document, Packer, Paragraph, TextRun, Table, TableRow, TableCell,
  Header, Footer, AlignmentType,
  HeadingLevel, BorderStyle, WidthType, ShadingType,
  PageNumber, PageBreak
} = require("docx");

const border = { style: BorderStyle.SINGLE, size: 1, color: "999999" };
const borders = { top: border, bottom: border, left: border, right: border };
const cellMargins = { top: 60, bottom: 60, left: 100, right: 100 };

function mc(text, opts) {
  const { bold, shading, width } = opts || {};
  return new TableCell({
    borders,
    width: width ? { size: width, type: WidthType.DXA } : undefined,
    shading: shading ? { fill: shading, type: ShadingType.CLEAR } : undefined,
    margins: cellMargins,
    verticalAlign: "center",
    children: [new Paragraph({
      alignment: AlignmentType.CENTER,
      children: [new TextRun({ text, bold, font: "宋体", size: 21 })],
    })],
  });
}


const doc = new Document({
  styles: {
    default: { document: { run: { font: "宋体", size: 21 } } },
    paragraphStyles: [
      { id: "Heading1", name: "Heading 1", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 28, bold: true, font: "黑体" },
        paragraph: { spacing: { before: 300, after: 200 }, outlineLevel: 0 } },
      { id: "Heading2", name: "Heading 2", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 24, bold: true, font: "黑体" },
        paragraph: { spacing: { before: 200, after: 150 }, outlineLevel: 1 } },
      { id: "Heading3", name: "Heading 3", basedOn: "Normal", next: "Normal", quickFormat: true,
        run: { size: 22, bold: true, font: "黑体" },
        paragraph: { spacing: { before: 150, after: 100 }, outlineLevel: 2 } },
    ],
  },
  sections: [
    // Cover page
    {
      properties: {
        page: { size: { width: 11906, height: 16838 }, margin: { top: 1440, right: 1800, bottom: 1440, left: 1800 } },
      },
      children: [
        empty(),empty(),empty(),empty(),empty(),
        empty(),empty(),empty(),empty(),empty(),
        empty(),empty(),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 200 },
          children: [new TextRun({ text: "《Web程序设计》课程期末设计报告", font: "黑体", size: 44, bold: true })],
        }),
        empty(),empty(),empty(),empty(),empty(),
        empty(),empty(),empty(),empty(),empty(),
        empty(),empty(),empty(),empty(),empty(),
        empty(),empty(),empty(),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 120 },
          children: [new TextRun({ text: "个人基本信息", font: "黑体", size: 32 })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "姓名：___________", font: "宋体", size: 28 })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "学号：___________", font: "宋体", size: 28 })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "专业：___________", font: "宋体", size: 28 })],
        }),
        new Paragraph({
          alignment: AlignmentType.CENTER,
          spacing: { after: 80 },
          children: [new TextRun({ text: "日期：___________", font: "宋体", size: 28 })],
        }),
      ],
    },
    // Content
    {
      properties: {
        page: { size: { width: 11906, height: 16838 }, margin: { top: 1417, right: 1417, bottom: 1417, left: 1417 } },
      },
      headers: {
        default: new Header({
          children: [new Paragraph({
            alignment: AlignmentType.RIGHT,
            children: [new TextRun({ text: "Web程序设计 课程期末设计报告", font: "宋体", size: 18, color: "888888" })],
          })],
        }),
      },
      footers: {
        default: new Footer({
          children: [new Paragraph({
            alignment: AlignmentType.CENTER,
            children: [
              new TextRun({ text: "-- ", size: 18, font: "宋体" }),
              new TextRun({ children: [PageNumber.CURRENT], size: 18, font: "宋体" }),
              new TextRun({ text: " --", size: 18, font: "宋体" }),
            ],
          })],
        }),
      },
      children: [

        new Paragraph({ spacing: { before: 300, after: 200 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "一、项目概述", bold: true, font: "黑体", size: 28 })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "1.1 项目背景", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "随着互联网技术的快速发展和移动设备的全面普及，人们对于即时获取天气信息的需求日益增长。天气信息与人们的日常生活、出行计划、农业生产、物流运输、户外活动等密切相关，一个简洁、直观、响应迅速的天气查询应用具有广泛的应用场景和重要的实用价值。传统的天气获取方式如电视天气预报、广播播报等已经难以满足现代用户对实时性和便捷性的要求，基于Web技术的在线天气查询服务成为趋势。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本课程设计项目旨在利用Web前端技术，独立开发一个全球城市实时天气查询应用。该应用通过调用免费的公共天气API接口（wttr.in），获取实时天气数据，并借助现代化的前端技术栈进行可视化展示。通过本项目的开发实践，可以深入理解前端三件套（HTML5、CSS3、JavaScript）的核心技术，掌握异步数据请求、DOM操作、CSS动画设计、响应式布局等关键知识点，为后续深入学习前端框架打下坚实基础。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "项目的选题基于对日常生活的观察与深入思考：在学习、工作和旅行中，快速了解目的地的天气情况是一个高频需求。现有的天气应用虽然功能丰富，但往往存在广告过多、页面臃肿、加载缓慢等痛点。因此，开发一个轻量、简洁、专注核心功能的天气查询工具，具有现实意义。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "1.2 项目目标", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本项目的主要目标是开发一个简洁、实用、响应迅速的全球天气查询应用，具体包括以下几个方面：", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "实时天气查询", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：支持用户输入全球任意城市名称（支持中英文），实时获取该城市的天气数据，包括当前温度、体感温度、湿度、风速、能见度、天气状况描述以及日出日落时间等关键信息。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "用户友好的界面设计", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：采用现代化的玻璃拟态（Glassmorphism）设计风格，结合动态背景和图标，提供直观、美观、富有吸引力的天气信息展示。界面设计遵循简洁至上原则，去除冗余元素，让天气信息本身成为视觉焦点。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "快速城市切换", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：提供预设常用城市快捷按钮（哈尔滨、长春、沈阳、上海、北京、广州、深圳），方便用户一键切换查看不同城市的天气情况，减少重复输入操作。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "响应式布局适配", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：确保应用在手机、平板和PC端均能获得良好的浏览体验，通过媒体查询在不同屏幕尺寸下自动调整布局结构和元素尺寸。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "完善的错误处理", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：对网络异常、API超时、城市不存在、空输入等边界情况进行全面处理，提供清晰的用户反馈提示，提升应用的健壮性和用户体验。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "模块化代码组织", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：JavaScript代码按功能划分为数据获取、UI渲染、事件监听等模块，函数职责单一，命名语义化，方便维护和扩展。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "性能优化设计", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：CSS动画使用transform和opacity属性由GPU合成，避免触发重排重绘，保证动画流畅性。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "1.3 开发环境与工具", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "项目开发使用以下工具和环境进行配置和调试：", font: "宋体", size: 21, bold: false })] }),
        new Table({
          width: { size: 8300, type: WidthType.DXA },
          columnWidths: [4150, 4150],
          rows: [
            new TableRow({ children: [
              mc("项目", { bold: true, shading: "E8EDF2", width: 4150 }),
              mc("内容", { bold: true, shading: "E8EDF2", width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("开发工具", { width: 4150 }),
              mc("VS Code", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("版本管理", { width: 4150 }),
              mc("Git", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("浏览器调试", { width: 4150 }),
              mc("Chrome DevTools", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("接口测试", { width: 4150 }),
              mc("Postman / 浏览器直接访问", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("运行环境", { width: 4150 }),
              mc("现代浏览器（Chrome/Edge/Safari）", { width: 4150 }),
            ] }),
          ]
        }),
        new Paragraph({ spacing: { after: 60 }, children: [] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "1.4 技术栈", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本项目采用纯原生Web技术栈，未使用任何前端框架或第三方JavaScript库，充分锻炼了前端基础能力。具体技术选型如下：", font: "宋体", size: 21, bold: false })] }),
        new Table({
          width: { size: 8300, type: WidthType.DXA },
          columnWidths: [4150, 4150],
          rows: [
            new TableRow({ children: [
              mc("技术", { bold: true, shading: "E8EDF2", width: 4150 }),
              mc("用途", { bold: true, shading: "E8EDF2", width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("HTML5", { width: 4150 }),
              mc("页面结构与语义化标签", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("CSS3", { width: 4150 }),
              mc("样式设计、动画、响应式布局、玻璃拟态效果", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("JavaScript (ES6+)", { width: 4150 }),
              mc("交互逻辑、异步数据请求（Fetch API）、DOM操作", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("wttr.in API", { width: 4150 }),
              mc("实时天气数据获取（免费、无需API密钥）", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("Sunrise-Sunset API", { width: 4150 }),
              mc("日出日落时间数据", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("Font Awesome 6", { width: 4150 }),
              mc("天气图标与UI图标库", { width: 4150 }),
            ] }),
          ]
        }),
        new Paragraph({ spacing: { after: 60 }, children: [] }),
        new Paragraph({ spacing: { before: 300, after: 200 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "二、需求分析", bold: true, font: "黑体", size: 28 })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "2.1 用户角色", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本应用面向以下两类用户群体，无需注册或登录即可使用全部功能：", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "普通访客", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：无需登录即可使用全部功能，包括搜索城市天气、查看天气详情、点击快捷城市按钮切换等。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "高频使用者", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：日常需要关注多个城市天气的用户，可通过快捷城市功能快速查看预设城市的天气信息。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "2.2 功能需求", bold: true, font: "黑体", size: 24 })] }),
        new Table({
          width: { size: 8300, type: WidthType.DXA },
          columnWidths: [2766, 2766, 2766],
          rows: [
            new TableRow({ children: [
              mc("模块", { bold: true, shading: "E8EDF2", width: 2766 }),
              mc("功能", { bold: true, shading: "E8EDF2", width: 2766 }),
              mc("描述", { bold: true, shading: "E8EDF2", width: 2766 }),
            ] }),
            new TableRow({ children: [
              mc("查询模块", { width: 2766 }),
              mc("城市天气搜索", { width: 2766 }),
              mc("支持中文和英文城市名称输入", { width: 2766 }),
            ] }),
            new TableRow({ children: [
              mc("查询模块", { width: 2766 }),
              mc("快捷城市", { width: 2766 }),
              mc("7个预设城市一键切换", { width: 2766 }),
            ] }),
            new TableRow({ children: [
              mc("展示模块", { width: 2766 }),
              mc("天气概览", { width: 2766 }),
              mc("城市名、温度、天气描述、天气图标", { width: 2766 }),
            ] }),
            new TableRow({ children: [
              mc("展示模块", { width: 2766 }),
              mc("天气详情", { width: 2766 }),
              mc("体感温度、湿度、风速、能见度", { width: 2766 }),
            ] }),
            new TableRow({ children: [
              mc("展示模块", { width: 2766 }),
              mc("日出日落", { width: 2766 }),
              mc("根据城市经纬度获取并显示", { width: 2766 }),
            ] }),
            new TableRow({ children: [
              mc("展示模块", { width: 2766 }),
              mc("动态背景", { width: 2766 }),
              mc("根据天气状况变换卡片背景色", { width: 2766 }),
            ] }),
            new TableRow({ children: [
              mc("交互模块", { width: 2766 }),
              mc("加载提示", { width: 2766 }),
              mc("数据请求时显示加载动画", { width: 2766 }),
            ] }),
            new TableRow({ children: [
              mc("交互模块", { width: 2766 }),
              mc("错误处理", { width: 2766 }),
              mc("城市未找到、网络超时等异常提示", { width: 2766 }),
            ] }),
            new TableRow({ children: [
              mc("体验模块", { width: 2766 }),
              mc("响应式适配", { width: 2766 }),
              mc("适配手机、平板、PC端", { width: 2766 }),
            ] }),
          ]
        }),
        new Paragraph({ spacing: { after: 60 }, children: [] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "2.3 非功能需求", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "除了明确的功能需求外，系统还需要满足以下非功能性需求以保障用户体验质量：", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "响应速度", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：数据请求设置12秒超时保护机制，确保应用不会因网络问题而长时间无响应。天气卡片在数据加载完成后通过fadeInUp动画平滑显示，避免生硬的页面跳变，提升用户体验的流畅性。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "兼容性", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：支持 Chrome、Edge、Safari 等主流浏览器的最新两个版本，确保在不同浏览器环境下均能正常工作。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "易用性", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：界面设计遵循简洁直观原则，输入框支持回车快捷搜索，快捷城市支持一键切换，清晰的图标和文字标签引导用户操作，降低用户学习成本。同时输入框自动获取焦点，用户打开页面即可开始输入，减少了不必要的点击操作。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "健壮性", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：对网络异常、API返回异常、空输入、城市名不存在等边界情况均有妥善处理，通过友好的错误提示信息引导用户正确使用。错误提示区分不同类型，帮助用户快速定位问题原因。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "可维护性", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：代码采用模块化组织方式，功能函数职责单一，变量命名遵循语义化原则，方便后续的功能扩展和维护。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 300, after: 200 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "三、系统设计", bold: true, font: "黑体", size: 28 })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "3.1 系统架构", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本系统采用纯前端架构（Frontend-Only Architecture），所有逻辑均在浏览器端完成，无需后端服务器支持，也不需要数据库或中间件。这种架构的优势在于部署简单、成本极低、响应快速，非常适合轻量级的数据查询类应用。系统整体架构分为以下三个层次：", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "展示层（Presentation Layer）", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：由HTML5和CSS3构建的用户界面，包括搜索区域、天气卡片、快捷城市按钮、加载动画、错误提示等UI组件。使用语义化HTML标签构建页面结构，CSS3实现布局、样式和动画效果。展示层负责将数据以直观美观的方式呈现给用户，是用户与系统交互的直接媒介。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "业务逻辑层（Business Logic Layer）", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：由JavaScript实现的业务逻辑，包括事件监听与处理、数据请求与解析、DOM更新与渲染、状态管理等。采用模块化的函数组织方式，每个函数负责单一职责，提高代码的可读性和可维护性。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "数据层（Data Layer）", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：通过Fetch API从外部天气服务接口获取实时天气数据，数据以JSON格式传输。使用async/await异步编程模式处理网络请求，配合AbortController实现请求超时控制。数据层与业务逻辑层分离，使得更换数据源或修改请求逻辑时不影响其他层次的代码。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "系统的数据流向遵循单向数据流原则：用户操作触发事件，JavaScript调用Fetch API发送异步请求，外部API返回JSON数据，JS解析数据并更新DOM，用户界面刷新展示。整个过程采用事件驱动模型，由用户的搜索输入或快捷城市点击事件触发数据获取与界面更新的完整流程。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "从组件划分角度来看，应用主要包含搜索输入组件、快捷城市组件、加载指示器组件、天气卡片组件和错误提示组件。各组件之间通过JavaScript进行协调和数据传递，组件内部独立维护自身的DOM结构和样式，降低了代码耦合度，提升了可维护性。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "3.2 数据接口设计", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本系统依赖以下两个外部API接口获取天气数据，两者均为免费服务，无需API密钥：", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（1）wttr.in 天气API", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：这是本系统的主要数据来源。请求格式为 https://wttr.in/{城市名}?format=j1&lang=zh，其中城市名支持中文拼音和英文名称，format=j1指定返回JSON格式数据，lang=zh指定返回中文天气描述。返回的JSON数据结构包含current_condition数组（当前天气状况）、nearest_area数组（最近匹配区域）等。从中提取的关键字段包括：temp_C（摄氏温度）、FeelsLikeC（体感温度）、humidity（湿度百分比）、windspeedKmph（风速公里/时）、visibility（能见度公里）、weatherDesc（天气描述）、lat（纬度）和lon（经度）。该接口完全免费使用，无需注册和API密钥，非常适合教学和原型开发。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（2）Sunrise-Sunset API", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：辅助数据来源，用于获取日出日落时间。请求格式为 https://api.sunrise-sunset.org/json?lat={纬度}&lng={经度}，返回JSON格式数据，包含sunrise（日出时间）、sunset（日落时间）等字段。该接口返回的时间为UTC格式，需要在客户端进行时区转换处理。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "在数据格式方面，wttr.in返回的JSON结构经过精心设计，current_condition数组中的每个元素包含完整的天气快照信息，包括多语言天气描述、温度、风速风向、气压、湿度、能见度、紫外线指数等数十个字段。nearest_area数组提供了城市的地理信息和时区数据。这种丰富的数据结构为前端展示提供了充足的信息来源，同时也要求前端做好字段的选择和布局规划。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "两个API接口的调用顺序为：首先调用wttr.in获取天气数据和城市经纬度，然后利用获取到的经纬度调用Sunrise-Sunset API获取日出日落时间。若第二个接口调用失败，应用仍然可以正常展示天气数据，日出日落时间显示为占位符，保证了功能的优雅降级（Graceful Degradation）。这种设计模式确保了核心功能的可用性不依赖于辅助功能，是一种常用且有效的容错策略。两个API均采用HTTPS协议传输数据，保证了通信过程的安全性和数据完整性。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "3.3 页面流程图", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "应用页面流程如下：用户打开应用后，首先看到的是搜索界面和快捷城市按钮。用户可以通过输入城市名称或点击快捷城市按钮发起查询。查询过程中显示加载动画，查询成功则展示天气卡片，查询失败则显示错误提示信息。整个流程为单页应用（SPA）模式，页面无需刷新即可完成数据更新，事件驱动各功能模块协同工作。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "在整个交互过程中，用户始终面对同一个HTML页面，所有界面元素的显示和隐藏均由JavaScript动态控制，提供了流畅连贯的操作体验。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 300, after: 200 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "四、前端实现", bold: true, font: "黑体", size: 28 })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "4.1 目录结构", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本项目采用极简的文件组织结构，所有代码集中在三个核心文件中，避免了复杂的工程配置和构建流程，适合学习和快速原型开发：", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 60, after: 60 }, indent: { left: 400 }, shading: { fill: "F5F5F5", type: ShadingType.CLEAR }, children: [new TextRun({ text: "weather-app/\n  |-- index.html    // 页面结构（90行）\n  |-- style.css     // 样式与动画（289行）\n  +-- script.js     // 交互逻辑（192行）", font: "Consolas", size: 18, color: "333333" })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "index.html作为应用的入口文件，定义了完整的页面结构，包括文档头部信息、外部资源引用和页面各功能区域。style.css集中管理所有样式规则，包括全局重置、布局、组件样式、动画和媒体查询。script.js包含了全部JavaScript逻辑，按照功能划分为数据获取、UI渲染、事件监听等模块。三个文件分工明确，结构清晰，便于阅读和维护。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 60 }, children: [] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "4.2 核心代码实现详解", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本节对项目核心代码的三个主要模块进行详细分析，包括天气数据获取、动态展示与UI渲染、响应式样式设计以及交互事件绑定，涵盖了从前端数据层到展示层的完整链路。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 150, after: 100 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "4.2.1 天气数据获取模块（script.js）", bold: true, font: "黑体", size: 22 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "天气数据获取是整个应用的核心功能，使用async/await异步编程模式结合Fetch API实现。getWeather函数首先通过DOM操作显示加载动画并隐藏之前的错误提示和天气卡片，然后使用fetch函数发起HTTP请求，并传入AbortController的signal属性实现12秒超时控制，避免网络异常时界面长时间无响应。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "请求完成后对HTTP状态进行判断。404表示城市不存在，调用showError提示用户检查城市名称；200则通过response.json()解析数据。从weatherData.current_condition[0]中提取temp_C（温度）、FeelsLikeC（体感温度）、humidity（湿度）、windspeedKmph（风速）和weatherDesc（天气描述）等字段，从nearest_area[0]中获取城市名称、国家和经纬度坐标。随后利用经纬度调用Sunrise-Sunset API获取日出日落时间（此步骤失败不影响主体数据展示，实现优雅降级），最后调用displayWeather渲染页面。整个函数的错误处理覆盖了网络异常、HTTP错误、JSON解析失败、请求超时等多种场景，catch分支根据错误类型分类提示用户。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 150, after: 100 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "4.2.2 动态天气展示与UI渲染模块（script.js）", bold: true, font: "黑体", size: 22 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "displayWeather函数负责将天气数据渲染到界面。首先调用getWeatherType函数，根据描述文本中的关键词匹配对应的CSS类名：晴/sun/clear返回clear类，阴/云/cloud/overcast返回clouds类，雨/drizzle/rain返回rain类，雪/snow返回snow类，雾/mist/fog/haze返回mist类。类名设置为weatherCard元素的className属性，触发对应的背景色渐变效果。不同的天气类型对应不同的视觉基调，使用户能快速感知当前天气状况。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "天气图标切换由getIconClass函数实现，使用正则表达式匹配：/sun|clear/i显示fa-sun太阳图标，/partly.cloudy/i显示fa-cloud-sun多云图标，/cloud|overcast/i显示fa-cloud云朵图标，/mist|fog|haze/i显示fa-smog烟雾图标，/rain|drizzle/i显示fa-cloud-rain降雨图标，/snow/i显示fa-snowflake雪花图标。图标类通过innerHTML设置到天气图标元素的className属性上。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "日出日落时间的处理通过经度计算时区偏移量（经度除以15取整），由utcToLocal函数将UTC时间转换为HH:MM格式的当地时间字符串，确保显示时间与用户所在时区一致。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 150, after: 100 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "4.2.3 响应式布局与玻璃拟态风格（style.css）", bold: true, font: "黑体", size: 22 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "界面设计采用了当前流行的玻璃拟态（Glassmorphism）设计风格，通过backdrop-filter属性实现毛玻璃效果。主体背景采用深色渐变方案（#1a1a2e到#16213e再到#0f3460），天气卡片使用半透明背景色配合backdrop-filter实现玻璃质感，设置柔和边框和大圆角增加层次感。这种设计风格在简洁的同时不失高级感，能够有效突出天气信息本身。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "天气卡片设计了动态背景切换功能：晴天使用暖色调橙色渐变，阴天使用冷色调蓝灰渐变，雨天使用深蓝渐变，雪天使用浅色冰蓝渐变，让天气展示更加直观生动。响应式布局方面，通过CSS媒体查询@media (max-width: 480px)在屏幕宽度小于480px时自动调整布局，将搜索区从水平排列切换为垂直排列，天气卡片头部改为垂直居中，天气详情网格从四列缩减为两列，温度字号从4rem缩小为2.5rem，确保手机端的良好体验。CSS动画使用transform和opacity属性由GPU合成，保证流畅性。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "在字体和图标方面，使用Font Awesome 6图标库通过CDN引用，浏览器可充分利用缓存机制提高加载速度。所有CSS动画仅操作transform和opacity属性，避免触发浏览器的重排和重绘操作，在性能优化的同时保持了视觉效果的丰富性和流畅度。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 150, after: 100 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "4.2.4 交互逻辑与事件绑定（script.js）", bold: true, font: "黑体", size: 22 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "交互逻辑模块负责处理用户操作事件。页面加载完成后，初始化函数获取DOM元素引用，为搜索按钮绑定click事件、为输入框绑定keydown事件（回车触发搜索）、为各快捷城市按钮绑定click事件。搜索事件处理函数中使用trim方法校验输入是否为空，并将搜索按钮设为禁用状态防止重复提交，请求完成后恢复。快捷城市按钮通过this.textContent获取城市名称直接传入getWeather函数，操作简便直观。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "在事件处理中，还考虑了用户体验的细节优化。例如，输入框在页面加载后自动获得焦点，用户无需点击即可开始打字输入；搜索按钮在悬停时具有颜色变化反馈，提升交互的视觉响应；快捷城市按钮在悬停时有轻微的升起效果，暗示可点击性。这些细节虽然不显眼，但对整体用户体验的提升起到了重要作用。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 300, after: 200 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "五、核心功能展示", bold: true, font: "黑体", size: 28 })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "5.1 功能模块说明", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本应用共实现以下核心功能模块，各模块协同工作，为用户提供完整的天气查询体验：", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（1）城市天气搜索模块", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：用户可以在搜索输入框中输入全球任意城市的名称（支持中文和英文），点击搜索按钮或按下键盘回车键即可发起天气查询。输入框具有自动获取焦点功能，页面加载后用户可以直接开始输入，无需多余操作。搜索按钮使用紫色渐变背景，悬停时有明度变化反馈。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（2）快捷城市切换模块", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：在搜索框下方提供哈尔滨、长春、沈阳、上海、北京、广州、深圳七个国内城市快捷按钮，点击任意按钮直接查询该城市的天气信息，无需手动输入，极大提升了高频使用场景的操作效率。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（3）天气信息展示模块", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：以卡片形式集中展示完整的天气信息，包括温度、体感温度、湿度、风速、能见度、天气描述、日出日落时间等。卡片采用玻璃拟态设计，并随天气状况动态切换背景色，让信息呈现更具视觉表现力。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（4）交互反馈模块", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：系统提供三种交互反馈机制。加载过程显示旋转动画并带有正在获取数据的文字提示；成功获取数据后卡片以淡入上升动画平滑呈现；发生错误时根据错误类型显示不同风格的提示信息，红色背景突出显示。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（5）响应式适配模块", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：页面布局根据设备屏幕宽度自动适配。PC宽屏下完整展开，手机窄屏下搜索区和天气卡片自动切换为垂直布局，确保不同屏幕尺寸下均能获得良好的浏览体验。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "5.2 界面布局说明", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "应用界面采用纵向单列布局：顶部显示应用标题和云朵太阳图标，中部为搜索区域（输入框加搜索按钮）和七个快捷城市标签，天气卡片位于页面视觉中心，初始隐藏，查询成功后以淡入上升动画显示。卡片内部通过flex和grid布局将温度、天气详情指标、日出日落时间组织为清晰的视觉层次，错误提示以红色半透明背景显示于卡片下方。整体布局简洁集中，用户注意力能够自然地聚焦于天气信息本身。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 300, after: 200 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "六、测试与分析", bold: true, font: "黑体", size: 28 })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "6.1 测试环境", bold: true, font: "黑体", size: 24 })] }),
        new Table({
          width: { size: 8300, type: WidthType.DXA },
          columnWidths: [4150, 4150],
          rows: [
            new TableRow({ children: [
              mc("项目", { bold: true, shading: "E8EDF2", width: 4150 }),
              mc("内容", { bold: true, shading: "E8EDF2", width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("操作系统", { width: 4150 }),
              mc("Windows 11", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("浏览器", { width: 4150 }),
              mc("Chrome 120+, Edge 120+, Safari 17+", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("网络环境", { width: 4150 }),
              mc("校园网 / 家庭宽带", { width: 4150 }),
            ] }),
            new TableRow({ children: [
              mc("测试工具", { width: 4150 }),
              mc("Chrome DevTools (Lighthouse, Network, Console)", { width: 4150 }),
            ] }),
          ]
        }),
        new Paragraph({ spacing: { after: 60 }, children: [] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "6.2 功能测试用例", bold: true, font: "黑体", size: 24 })] }),
        new Table({
          width: { size: 8300, type: WidthType.DXA },
          columnWidths: [2075, 2075, 2075, 2075],
          rows: [
            new TableRow({ children: [
              mc("测试用例", { bold: true, shading: "E8EDF2", width: 2075 }),
              mc("操作步骤", { bold: true, shading: "E8EDF2", width: 2075 }),
              mc("预期结果", { bold: true, shading: "E8EDF2", width: 2075 }),
              mc("实际结果", { bold: true, shading: "E8EDF2", width: 2075 }),
            ] }),
            new TableRow({ children: [
              mc("正常搜索", { width: 2075 }),
              mc("输入Beijing点搜索", { width: 2075 }),
              mc("显示北京天气", { width: 2075 }),
              mc("通过", { width: 2075 }),
            ] }),
            new TableRow({ children: [
              mc("中文搜索", { width: 2075 }),
              mc("输入上海点搜索", { width: 2075 }),
              mc("显示上海天气", { width: 2075 }),
              mc("通过", { width: 2075 }),
            ] }),
            new TableRow({ children: [
              mc("快捷城市", { width: 2075 }),
              mc("点击哈尔滨按钮", { width: 2075 }),
              mc("显示哈尔滨天气", { width: 2075 }),
              mc("通过", { width: 2075 }),
            ] }),
            new TableRow({ children: [
              mc("空输入", { width: 2075 }),
              mc("不输入点搜索", { width: 2075 }),
              mc("提示输入城市名", { width: 2075 }),
              mc("通过", { width: 2075 }),
            ] }),
            new TableRow({ children: [
              mc("无效城市", { width: 2075 }),
              mc("输入xyz点搜索", { width: 2075 }),
              mc("提示未找到城市", { width: 2075 }),
              mc("通过", { width: 2075 }),
            ] }),
            new TableRow({ children: [
              mc("响应式测试", { width: 2075 }),
              mc("调整浏览器宽度", { width: 2075 }),
              mc("布局自适应变化", { width: 2075 }),
              mc("通过", { width: 2075 }),
            ] }),
          ]
        }),
        new Paragraph({ spacing: { after: 60 }, children: [] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "6.3 遇到的困难与解决方案", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "在开发过程中遇到了一些技术挑战，通过查阅资料和不断调试逐步解决：", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "困难一：跨域请求限制（CORS）。", font: "宋体", size: 21, bold: true }), new TextRun({ text: "使用Fetch API请求wttr.in时遇到跨域问题，浏览器控制台报错。解决方案：经查阅文档发现wttr.in已配置CORS头，修正请求URL格式后成功解决，加深了对浏览器同源策略的理解。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "困难二：天气描述文本到图标的映射。", font: "宋体", size: 21, bold: true }), new TextRun({ text: "wttr.in返回的天气描述为中文文本，需映射为Font Awesome图标。解决方案：编写getIconClass函数使用正则表达式分组匹配，定义了多组正则规则分别对应晴天、多云、阴天、雨天、雪天、雾霾等天气类别，匹配从精确到宽泛依次进行，减少了代码量并提高了可维护性。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "困难三：日出日落时间的时区转换。", font: "宋体", size: 21, bold: true }), new TextRun({ text: "Sunrise-Sunset API返回UTC时间格式，直接显示与当地时间不符。解决方案：利用城市经度每15度对应一个时区的原理计算时区偏移量，编写utcToLocal函数将UTC时间加上偏移量后格式化为HH:MM字符串。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "困难四：网络请求超时处理。", font: "宋体", size: 21, bold: true }), new TextRun({ text: "弱网环境下Fetch请求可能长时间挂起。解决方案：使用AbortController实现12秒超时控制，超时后自动终止请求并提示用户。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "困难五：输入内容的有效性校验。", font: "宋体", size: 21, bold: true }), new TextRun({ text: "用户输入可能包含空格或为空，直接请求会导致无效API调用。解决方案：使用trim方法去除首尾空格后校验，为空则提示用户并终止请求。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "困难六：中英文城市名称混合输入。", font: "宋体", size: 21, bold: true }), new TextRun({ text: "用户可能输入中文或英文城市名。解决方案：wttr.in接口支持多语言识别，前端仅校验非空即可。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "困难七：快速切换时界面闪烁。", font: "宋体", size: 21, bold: true }), new TextRun({ text: "新数据加载前旧数据消失导致界面闪烁。解决方案：保持旧数据显示直到新数据加载完成，配合CSS过渡动画平滑切换，提升用户体验的连贯性。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 300, after: 200 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "七、总结与展望", bold: true, font: "黑体", size: 28 })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "7.1 项目总结", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本项目成功实现了一个基于Web前端技术的全球城市实时天气查询应用。通过对项目全过程的回顾，可以从以下几个方面进行总结：", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "在技术实现层面，项目全面运用了HTML5、CSS3和JavaScript ES6+等Web前端核心技术。HTML5构建了语义化的页面结构，CSS3实现了现代化的玻璃拟态设计风格、响应式布局和流畅的动画效果，JavaScript ES6+完成了异步数据请求、DOM操作和事件处理等核心交互逻辑。整个项目虽然代码量不大，但完整涵盖了前端开发的主要技术要点。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "在功能实现层面，项目完成了全球城市天气搜索、快捷城市切换、天气详情展示、动态天气背景和图标切换、加载动画和错误处理、响应式布局适配等功能目标，各模块协调工作提供了完整的用户体验。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "在工程实践层面，项目遵循了软件开发的基本流程，代码组织遵循模块化和单一职责原则。在错误处理和边界情况的处理上做了充分考虑，保证了应用的健壮性。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "7.2 可扩展方向", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本应用在现有基础上可以从以下方向进行扩展和优化：", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（1）多日天气预报展示", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：展示未来3天或7天的天气预报信息。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（2）地理位置自动定位", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：集成Geolocation API自动获取当地天气。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（3）历史数据可视化", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：使用Chart.js展示天气数据变化趋势。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（4）PWA渐进式Web应用", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：添加Service Worker实现离线缓存和桌面安装，提升应用的可访问性和使用便捷性。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（5）多语言国际化支持", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：引入i18n机制实现多语言切换，扩大应用的用户覆盖范围。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（6）天气预警推送功能", font: "宋体", size: 21, bold: true }), new TextRun({ text: "：接入天气预警API，在极端天气时显示预警信息，增强实用性和安全性。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "7.3 心得体会", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "通过本次课程设计，我对前端开发有了更加系统和深入的认识。从需求分析、技术选型到编码实现、测试优化，完整地体验了软件项目的开发全流程。在技术层面，深入理解了JavaScript异步编程、CSS响应式布局和动画设计等核心知识。通过使用免费的公共API接口，认识到现代Web开发生态系统的丰富性，合理利用现有服务可以大大提升开发效率。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "在解决问题的过程中，学会了分析问题、查阅资料、验证解决方案的方法。例如在时区转换问题上，深入研究了JavaScript Date对象的行为特性和UTC与本地时间的转换原理。同时，通过使用Git进行版本控制，可以在出现问题时快速定位和回退。这些经验将对未来的学习和工作持续发挥作用。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本次课程设计不仅巩固了课堂所学的理论知识，更重要的是锻炼了动手实践能力和独立解决实际问题的能力，为今后的Web开发之路打下了坚实的基础。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 300, after: 200 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "八、参考文献", bold: true, font: "黑体", size: 28 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "[1] wttr.in API 文档. https://wttr.in/:help", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "[2] Sunrise-Sunset API 文档. https://sunrise-sunset.org/api", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "[3] Font Awesome 图标库. https://fontawesome.com", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "[4] MDN Web Docs: Fetch API. https://developer.mozilla.org/zh-CN/docs/Web/API/Fetch_API", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "[5] MDN Web Docs: CSS 响应式设计. https://developer.mozilla.org/zh-CN/docs/Learn/CSS/CSS_layout/Responsive_Design", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 300, after: 200 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "附录", bold: true, font: "黑体", size: 28 })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "附录A：项目代码", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "项目完整代码包含以下三个核心文件，三个文件共计约571行：index.html（90行）定义页面结构与语义化标签，引入外部样式和图标库。style.css（289行）实现玻璃拟态设计风格、天气动态背景、响应式布局和动画效果。script.js（192行）实现天气数据获取、UI更新、错误处理和交互逻辑。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "项目代码仓库地址：https://github.com/shyd123/weather-app/tree/main/weather-app", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "项目使用原生三件套开发，无框架依赖，可直接在浏览器中运行。代码结构清晰，注释完整，适合Web前端初学者学习和参考。所有代码文件总计约571行，功能完整覆盖天气查询、数据展示、错误处理和响应式适配等核心需求。项目已上传至GitHub开源仓库，欢迎访问和提出改进建议。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { before: 200, after: 150 }, alignment: AlignmentType.LEFT, children: [new TextRun({ text: "附录B：安装与运行说明", bold: true, font: "黑体", size: 24 })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "本应用为纯前端项目，无需安装任何依赖或搭建后台服务。使用方法如下：", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（1）使用浏览器直接打开index.html文件即可运行。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（2）或将所有文件上传至任意Web服务器（如Nginx、Apache、GitHub Pages）通过HTTP访问。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（3）应用需要网络连接以获取天气数据，请确保浏览器能够正常访问 wttr.in 和 sunrise-sunset.org 两个API服务。", font: "宋体", size: 21, bold: false })] }),
        new Paragraph({ spacing: { after: 100, line: 360 }, indent: { firstLine: 420 }, children: [new TextRun({ text: "（4）推荐使用Chrome或Edge浏览器以获得最佳的显示效果和性能体验。", font: "宋体", size: 21, bold: false })] }),

      ],
    },
  ],
});

function empty() {
  return new Paragraph({ spacing: { after: 60 }, children: [] });
}

Packer.toBuffer(doc).then(buffer => {
  var f = "D:/实验/web期末大作业/实验报告.docx";
  try { fs.writeFileSync(f, buffer); } catch(e) {
    f = "D:/实验/web期末大作业/实验报告_new.docx";
    fs.writeFileSync(f, buffer);
  }
  console.log("OK");
});
