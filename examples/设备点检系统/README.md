# 示例：设备点检管理系统（虚构）

这套文件演示三份草稿的写法与格式，全部为虚构内容，不对应任何真实项目或公司。

- `申请表信息.yaml` — `fill_application_form.py` 的输入，"主要功能"约 900 字，符合 500–1300 字要求
- `操作手册.md` — `build_manual_docx.py` 的输入，按真实菜单分章，截图用 `【截图预留】` 占位
- `代码文件顺序.txt` — `build_code_docx.py` 的输入格式

试跑（不含代码文档，因为没有真实源码）：

```bash
cd examples/设备点检系统
uv run --with python-docx --with pyyaml python3 ../../scripts/fill_application_form.py 申请表信息.yaml -o out/
uv run --with python-docx python3 ../../scripts/build_manual_docx.py 操作手册.md --name 设备点检管理系统 --version V1.0 --owner 示例智能科技有限公司 -o out/
```
