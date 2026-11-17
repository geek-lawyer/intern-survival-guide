# 实习生生存指南（公开阅读版）

一份面向实习生和带教者的上手手册，共 43 页、5 篇、26 章，涵盖工作底线、协作方法、成长方向、管理者责任、AI 使用和可复用的 SOP 思维。

- 版本：2026 年 7 月版
- 作者、主编与主要撰写：胡泽贤（George Hu）

## 阅读与下载

成品（DOCX / PDF）不放在本仓库，统一发布在作者网站：

- 中文：<https://brbrhu.com/intern-survival-guide/>
- English：<https://brbrhu.com/en/intern-survival-guide/>
- Français：<https://brbrhu.com/fr/intern-survival-guide/>

## 仓库结构

- `src/build_public_guide.py`：中文公开阅读版的 DOCX 构建脚本
- `src/build_public_guide_en.py`：英文公开版构建脚本，复用中文脚本中的版式与样式函数
- `assets/`：构建所需的公众号二维码
- `scripts/tone_audit.py`：检查篇幅、标题、章首与提示框密度、模板句式、重复正文和禁用词的零依赖文风审计
- `EDITORIAL_STYLE.md`：保持全书文风与删改尺度的编辑口径

## 如何构建

需要 Python 3.11 或更高版本。

```bash
python3 -m pip install -r requirements.txt
python3 src/build_public_guide.py      # 中文
python3 src/build_public_guide_en.py   # English
```

生成文件位于 `dist/`（已被 `.gitignore` 忽略）。

## 文风审计

```bash
python3 scripts/tone_audit.py --self-test   # 确认审计本身能抓到故意植入的问题
python3 scripts/tone_audit.py               # 审计 src/build_public_guide.py
```

审计阈值只留少量余量。新增内容原则上先替换较弱内容；确有理由调整阈值时，请同时说明保留了什么、删掉了什么。

## 勘误与反馈

发现错字、事实错误或表述问题，请在本仓库提交 GitHub Issue，注明页码或章节、原文和建议的修改。

## 许可

Copyright © 2026 胡泽贤（George Hu）

- **手册文字**（`src/` 中作为字符串写入的正文、`EDITORIAL_STYLE.md`、本说明）：[CC BY-NC-SA 4.0](LICENSE)。注明书名、作者、版本与出处后，可以复制、转载、翻译和改编，改编后的作品须以相同许可发布。商业出版、付费培训等商业使用请先联系作者取得书面许可。
- **程序代码**（`src/` 中的版式与构建逻辑、`scripts/tone_audit.py`）：[MIT](LICENSE-CODE)。
- **不在许可范围内**：`assets/official_account_qr.jpg`（公众号二维码），以及手册延伸阅读中链接的第三方文章和历届实习生的原文，权利归各自的作者。

---

## English

**The Intern Survival Guide** (public edition) is a 43-page handbook in 5 parts and 26 chapters for interns and the people who mentor them: working baselines, collaboration, growth, manager responsibilities, AI use and reusable SOP thinking. July 2026 edition. Author: 胡泽贤 (George Hu).

- Read and download: <https://brbrhu.com/en/intern-survival-guide/> (Chinese: <https://brbrhu.com/intern-survival-guide/>, French: <https://brbrhu.com/fr/intern-survival-guide/>)
- This repository holds only the build scripts, assets and the tone audit. Built DOCX/PDF files are published on the website, not here.
- Build: `pip install -r requirements.txt`, then `python3 src/build_public_guide.py` (Chinese) and `python3 src/build_public_guide_en.py` (English). Output goes to `dist/`.
- Tone audit: `python3 scripts/tone_audit.py --self-test`, then `python3 scripts/tone_audit.py`.
- Errata: please open a GitHub issue with the page or chapter, the original text and your suggested fix.
- Licence: handbook text is [CC BY-NC-SA 4.0](LICENSE); program code is [MIT](LICENSE-CODE). The QR image and linked third-party articles are not covered. Commercial use of the text needs the author's written permission.
