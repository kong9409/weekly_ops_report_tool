# 运营项目周报生成工具 v3

## 功能

- 只需上传与《绩效复盘表格优化-白板6.15.xlsx》结构一致的绩效表。
- 自动识别最新月份、最新周和上一周。
- 长短周不一致时，总量指标按日均环比。
- 自动生成主要指标波动和推荐分析维度。
- 差距原因、关键动作、下周计划均来自《运营项目周报.xlsx》。
- 每项支持下拉选择和补充说明。
- 支持下载 Word、TXT，以及打印为 PDF。
- Excel 数据只在浏览器本地处理，不上传服务器。

## 本地运行

需要 Node.js 18 或更高版本。

```bash
npm start
```

浏览器打开 `http://localhost:3000`。

也可以直接双击 `index.html` 使用；项目已内置 Excel 解析组件。

## 部署到 GitHub 和 Zeabur

1. 将本文件夹全部文件上传到 GitHub 仓库根目录。
2. 在 Zeabur 新建 Service，并选择该 GitHub 仓库。
3. Framework 选择 Node.js 或 Auto Detect。
4. Start Command 使用 `npm start`。
5. 部署完成后打开 Zeabur 分配的域名即可使用。

项目不需要数据库、环境变量或额外依赖。
