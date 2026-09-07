# dsh-law-thesis-review

一个 DeepSeek Harness **技能插件**：批改学生法学论文（国际经济法等），并输出一份带 **修订 + 批注** 的 Word 审查稿。

## 功能

- 读取学生上传的 Word（.docx）论文（支持单篇或文件夹）。
- 对照打包的常见问题清单（硕士/本科自查清单合并版 + 国际经济法/法学论文高发问题）逐项检查。
- 对可直接确定的错漏以 **Word 修订模式**（w:ins / w:del）修改。
- 对学生需自行处理的问题（部分整体偏题、缺依据/缺澄清等）写 **Word 批注**。
- 生成简短、客观、以宏观点评为主的总体评语与待修改清单。

## 安装

~~~bash
# 本地路径
dsh plugin --profile web add link:D:/Coding/Creating/dsh-law-thesis-review
# 或 npm 包（发布后）
dsh plugin --profile web add @signalight/dsh-law-thesis-review
~~~

重启 web 版后，技能自动生效，用“批改学生论文 / 审阅这篇论文 / 以修订+批注检查论文”等触发。

## 使用

技能底层使用自带脚本：

~~~bash
# 1) 提取正文 + 段索引
python skills/iet-law-thesis-review/scripts/review_docx.py extract 论文.docx --out ./work

# 2) 按 spec 应用修订 + 批注
python skills/iet-law-thesis-review/scripts/review_docx.py review 论文.docx --spec spec.json --out 论文_审查稿.docx [--author "教师名"]
~~~

spec 示例见 `skills/iet-law-thesis-review/examples/spec-template.json`。

## 说明

- **自包含**：`apply_edits.py`、`models.py`、`add_comment.py` 已 vendor，**无需**额外安装 word-docx 插件。
- **清单打包版**：使用 `skills/iet-law-thesis-review/references/checklist.md`，不依赖本机 Obsidian 路径。
- 输出保留原标题与学生信息；批注/修订作者默认取本机 Word/WPS 用户名，也可用 `--author` 指定。
- Word 中打开 `_审查稿.docx` 可看到修订与批注气泡。

## License

MIT

