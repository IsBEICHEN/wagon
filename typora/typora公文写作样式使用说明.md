# Typora 公文写作 Word 导出样式使用说明

## 文件用途

- `typora公文写作_word导出_reference.docx`：Typora 导出 Word 时使用的参考样式文件，核心有效。
- `typora-gongwen.css`：Typora 编辑/预览样式，主要影响编辑器显示和 HTML/PDF，不保证控制 Word 导出。

## Typora 中使用 reference.docx

1. 打开 Typora 的“偏好设置”。
2. 进入“导出”或“Export”设置。
3. 找到 Word / DOCX 导出配置。
4. 将“Reference Docx”设置为本目录下的 `typora公文写作_word导出_reference.docx`。
5. 导出 Markdown 为 Word。

## Markdown 写作建议

```markdown
# 山东济阳农村商业银行股份有限公司工作机项目

## 需求调研报告

### 一、调研摘要

正文段落直接书写。导出后正文为三号仿宋_GB2312、首行缩进约两个汉字、固定行距约 28 磅。

![架构图](images/architecture.png)
```

## 样式口径

- 页面：A4。
- 页边距：上 37mm、下 35mm、左 28mm、右 26mm。
- 正文：三号仿宋_GB2312，固定行距 28 磅，首行缩进约 2 字符。
- Markdown `#`：二号黑体，居中。
- Markdown `##`：小二号黑体，居中。
- Markdown `###`：三号黑体，左对齐。
- Markdown `####` 及以下：三号黑体，左对齐。
- 表格：小四号仿宋_GB2312，网格线。
- 图片：建议使用标准 Markdown 图片语法，单独成段，居中显示，不使用 float/position。
- 所有文字颜色：黑色。

## 注意

Typora 的 DOCX 导出通常由 Pandoc 完成，Word 样式主要取决于 `reference.docx`。如果导出后字体不一致，请先确认本机是否安装 `仿宋_GB2312` 和 `黑体`。

如果图片导出后压到文字下面，通常是图片被导成 Word 浮动对象或“衬于文字下方”。处理建议：

1. Markdown 中使用标准 `![](path)`，不要用 `<img style="float:...">` 或 `position:absolute`。
2. 图片单独占一段，图片前后各留一个空行。
3. 在 Word 中选中异常图片，将“布局选项”改为“嵌入型”或“上下型”。
4. 使用 `typora-gongwen.css` 后重新导出，CSS 已强制图片不浮动、不绝对定位。
