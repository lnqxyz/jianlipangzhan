# 0.3.1 重跑产物（只出表 9 / 表 10）

- date: 2026-11-16；台账：三份现行台账（路径见 jianlipangzhan/SKILL.md 第 2 步）
- extract.json：三份台账全量提取（`_验收/e2e/031b` 快照同一份）
- out/：表 9 防排水安装、表 10 超前水平钻孔/加深炮孔 两张 fin docx
- raw/：fill → postprocess → fix_merge 中间产物（selfcheck 全 PASS）
- table-9-extract.json / table-10-extract.json：两表 extract 节选（含 source 溯源）
- verify-9.json / verify-10.json：产物核对（sha256、留白口径、编号连续）

## 复现步骤

```bash
cd "/Users/lnq/Documents/智能体 skill/监理旁站记录/_验收/重跑/0.3.1"
python3 -B jianlipangzhan/scripts/fill_docs.py extract.json jianlipangzhan/assets/template-20rows.docx raw --no 9
python3 -B jianlipangzhan/scripts/fill_docs.py extract.json jianlipangzhan/assets/template-20rows.docx raw --no 10
python3 -B jianlipangzhan/scripts/postprocess.py raw/raw-9.docx raw/pp-9.docx --records 1 --table-no 9
python3 -B jianlipangzhan/scripts/postprocess.py raw/raw-10.docx raw/pp-10.docx --records 1 --table-no 10
python3 -B jianlipangzhan/scripts/fix_merge.py raw/pp-9.docx raw/fin-9.docx
python3 -B jianlipangzhan/scripts/fix_merge.py raw/pp-10.docx raw/fin-10.docx
```

## 与 0.3.0 的差异

- 表 9：document.xml 与 0.3.0 完全一致（source 溯源补 c10 只进证据，不进 docx）。
- 表 10：仅「外插角 27°。」→「外插角 27°、28°。」一处文案差异（0.3.1 修复点）。
- 其余 10 张：未重出；031b 快照 E2E 中 document.xml 与 样例输出/ 逐一比对 10/10 一致。
