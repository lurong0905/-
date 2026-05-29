# 发货资料生成 Skill

这个 Codex Skill 用于根据合同资料、发货信息搜集表、客户模板和仓库发货清单生成发货资料，当前重点覆盖中核体系的发货通知、唛头和装箱清单。

## 主要能力

- 创建并预填 `发货信息搜集表.xlsx`
- 按中核模板生成 `发货通知`
- 按箱号生成一箱一页的 `唛头`
- 按规则生成 `装箱清单`
- 内置中核电厂默认收货地址
- 支持仓库清单中的合并单元格、续行箱号、箱号范围

## 中核规则摘要

- 正式生成前必须确认 `计划发运时间`、`装卸货方式`、`储存级别`、`储存条件`
- `预计到达时间` 缺失时应由用户确认
- `商务合同处合同经理` 取 `合同固定信息.合同管理联系人`
- `立项单号` 和 `采购申请处室 项目负责人` 默认填 `/`
- 唛头 `总件数` 取本批发货清单展开后的最大箱号数字
- 唛头 `物项名称` 只保留中文物项名称；同箱多种物项可用 `等`
- 中核装箱清单只在单箱不同物品类别达到 3 类及以上时生成

## 目录结构

```text
generate-shipping-notice/
  SKILL.md
  README.md
  agents/
    openai.yaml
  references/
    customer-template-rules.md
    shipping-info-table.md
  scripts/
    create_shipping_info_table.py
    generate_cnnp_shipping_notice.py
    generate_cnnp_shipping_marks.py
    generate_cnnp_packing_lists.py
```

## 使用方式

在合同文件夹内准备：

- 合同 PDF 或其他合同资料
- 仓库发货清单 Excel
- 客户模板文件夹

先创建信息搜集表：

```bash
python scripts/create_shipping_info_table.py <合同文件夹>
```

生成中核发货通知：

```bash
python scripts/generate_cnnp_shipping_notice.py <合同文件夹> --template <中核发货通知模板.docx>
```

生成中核唛头：

```bash
python scripts/generate_cnnp_shipping_marks.py <合同文件夹> --template <中核唛头模板.docx>
```

生成中核装箱清单：

```bash
python scripts/generate_cnnp_packing_lists.py <合同文件夹> --template <中核装箱清单模板.docx>
```

## 注意事项

本仓库只保存可复用的 Skill 规则和脚本，不保存合同原件、发货输出文件、客户模板原件或业务过程中的临时文件。
