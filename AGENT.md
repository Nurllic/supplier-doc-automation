# QA 质量协议自动生成系统 — Agent Harness

## 项目概述

自动化生成供应商质量保证协议。输入明细表 + 质量标准文件 → 输出 Word 质量保证协议。

**项目路径**: `.`
**核心脚本**: `agreement-automation\batch_generate_v3.py`（多物料）、`agreement-automation\generate_qa_agreement_v3.py`（单物料）
**模板**: `agreement-automation\质量保证协议模板.docx`
**数据**: `data\{供应商名}\materials\新增明细表_vN.docx` + `data\{供应商名}\materials\质量标准_vN\`
**联系人表**: `agreement-automation\平台联系人的对应关系.xlsx`
**输出**: `data\{供应商名}\agreements\`
**增量记录**: `data\{供应商名}\history.json`

---

## 1. Agent Loop（思考→行动→观察循环）

每次启动遵循以下循环：

1. **读取状态** → 检查 `state/` 目录下的 checkpoint.json 确定当前阶段
2. **检查 AI-Agent 邮箱** → 查找主题含【企业质量协议签订】的新邮件
3. **下载附件** → 新增明细表.docx + 质量标准.zip/.rar
4. **解析输入** → 确定 incoming_data 中的文件路径、供应商名、平台、OA流水号
5. **执行阶段** → 调用对应 tool/script
6. **验证输出** → 检查生成文件完整性
7. **保存检查点** → 更新 checkpoint.json
8. **通过AI-Agent邮箱发送** → 将协议发送给平台负责人
9. **汇报** → 输出执行摘要

### 执行阶段定义

| 阶段 | 操作 | 输出 |
|------|------|------|
| 阶段 | 操作 | 输出 |
|------|------|------|
| `CHECK_EMAIL` | 检查AI-Agent邮箱，查找【企业质量协议签订】邮件 | 邮件ID/附件路径 |
| `EXTRACT_ATTACHMENTS` | 下载附件，解压质量标准.zip/.rar | docx文件列表 |
| `WRITE_TO_DINGTALK` | 将邮件信息写入钉钉AI表格（状态：待确认供应商） | 表格记录ID |
| `SEND_CARD` | 发送群卡片消息通知相关人员确认供应商 | 群消息结果 |
| `WAIT_CONFIRM` | 等待人工在钉钉AI表格确认供应商名称 | — |
| `READ_FROM_DINGTALK` | 读取钉钉AI表格 → 获取供应商名称/平台 | 供应商信息 |
| `CHECK_HISTORY` | 检查 data/{供应商}/history.json 是否存在 | 增量/全新标记 |
| `PARSE_INPUT` | 读取新增明细表.docx，解析物料列表 | materials.json |
| `PARSE_STANDARDS` | 逐一解析各物料的企业质量标准文件 | standards.json |
| `MERGE_MATERIALS` | 增量模式：合并历史物料+新增物料，去重 | 完整物料列表 |
| `GENERATE` | 调用 batch_generate_v3.py 生成协议 | *.docx |
| `VALIDATE` | 验证文件完整性、物料数量匹配 | 验证报告 |
| `SAVE_HISTORY` | 更新 history.json，协议存入 data/{供应商}/agreements/ | history.json |
| `SEND_EMAIL` | 通过AI-Agent邮箱发送给平台负责人，附协议文件 | 发送状态 |
| `UPDATE_DINGTALK` | 更新钉钉AI表格状态为"协议已发送" | 表格状态 |
| `SEND_NOTIFICATION` | 发送群通知：协议已发送 | 群消息结果 |
| `COMPLETE` | 更新状态，输出最终摘要 | 最终结果 |

---

## 2. Tool Layer（工具层）

### 可用工具

| 工具 | 用途 | 命令 |
|------|------|------|
| `parse_raw_material` | 解析明细表 docx → 物料列表 | `python -c "from generate_qa_agreement_v3 import parse_raw_material; materials = parse_raw_material(path)"` |
| `parse_quality_standard` | 解析单一质量标准 docx → 指标列表 | `python -c "from generate_qa_agreement_v3 import parse_quality_standard; inds = parse_quality_standard(path)"` |
| `batch_generate` | 批处理生成质量协议 | `python batch_generate_v3.py` |
| `get_supplier_from_dingtalk` | 从钉钉AI表格读取供应商名称/平台 | 待对接钉钉API |
| `get_platform_contact` | 查询平台审核人信息 | `python -c "from generate_qa_agreement_v3 import get_platform_contact; ..."` |
| `extract_spec` | 从规格/参数字段提取纯规格 | 内置于 generate_qa_agreement_v3.py |
| `replace_placeholder` | 替换 Word 模板占位符 | 内置于脚本 |

### 文件系统操作

- **读**: `python-docx` 解析 docx, `pandas` 读 xlsx, `open()` 读 json/txt
- **写**: `python-docx` 生成 docx, `json.dump` 写状态文件
- **删除**: `os.remove()` 清理旧输出

---

## 3. Memory / State（状态管理）

### 检查点文件: `state/checkpoint.json`（运行时自动创建）

```json
{
  "project": "qa-agreement-gen",
  "version": "3.0",
  "phase": "GENERATE",
  "timestamp": "2026-07-25T10:30:00",
  "input": {
    "raw_material_file": "email-database/QA/新增明细表.docx",
    "supplier_name": "供应商A公司",
    "platform": "平台D",
    "oa_serial": "QA-20260707-001"
  },
  "materials": {
    "total": 16,
    "with_std": 10,
    "without_std": 6,
    "indicators_total": 83
  },
  "output": {
    "path": "output/QA-20260707-001-【生物】-供应商A公司-16.docx",
    "file_size_bytes": 33350,
    "validated": true
  },
  "errors": [],
  "session_id": null
}
```

### 物料-质量标准映射: `state/material_std_map.json`（运行时自动创建）

维护条目编码与质量标准文件的映射关系。当新增物料时需更新此文件。

---

## 4. Safety / Permissions（安全与权限）

### 文件保护

- **模板文件** (`质量保证协议模板.docx`): 只读，不允许修改
- **联系人表** (`平台联系人的对应关系.xlsx`): 只读
- **质量标准文件** (`email-database/QA/*.docx`): 只读
- **输出目录** (`output/`): 可写，每次生成前清理同名文件

### 参数校验

| 参数 | 校验规则 |
|------|---------|
| `supplier_name` | 非空字符串，不含路径遍历字符 (`../`) |
| `platform` | 必须在联系人表中存在 |
| `oa_serial` | 非空字符串，sanitize 后写入文件名 |
| `raw_material_file` | 文件必须存在且为有效 docx |
| `std_files` | 必须存在且大小 > 0 |

### 拒绝操作

- 不允许修改模板文件内容
- 不允许删除质量标准文件
- 不允许输出到 output/ 以外的目录

---

## 5. Lifecycle / Recovery（生命周期与恢复）

### 续跑流程

如果 Agent 中断重启：

1. 检查 `state/checkpoint.json` 是否存在
2. 如果存在，读取 `phase` 字段
3. 根据 phase 跳过已完成的阶段：
   - `PARSE_INPUT` 完成 → 用 cached materials.json
   - `PARSE_STANDARDS` 完成 → 用 cached standards.json
   - `GENERATE` 完成但 `VALIDATE` 失败 → 重新生成
   - `COMPLETE` → 直接返回结果
4. 如果不存在 checkpoint，从头开始

### 错误恢复

- **文件解析失败**: 跳过该物料，记录到 errors 列表，继续处理其余物料
- **生成失败**: 清除不完整输出文件，重试最多 3 次
- **磁盘空间不足**: 暂停并报错

---

## 6. Client/Runtime Integration（客户端集成）

### CLI 模式

```bash
cd .\agreement-automation
python batch_generate_v3.py
```

### 参数化调用

```python
from batch_generate_v3 import batch_generate

result = batch_generate(
    raw_material_docx="email-database/QA/新增明细表.docx",
    supplier_name="供应商A公司",
    platform_name="平台D",
    oa_serial="QA-20260707-001",
)
```

### 环境要求

- Python 3.10+
- `python-docx`, `pandas`, `openpyxl`, `lxml`
- Hermes Agent（用于 skill 调用）

### 条目编码 → 供应商货号映射（自动从 供应商货号/标识号 列提取）

供应商货号列（质量标准表中的第一列）取自明细表的"供应商货号/标识号"字段中的"供应商货号：XXX"部分，而非"条目编码"列。

---

## 数据文件清单

| 文件 | 说明 |
|------|------|
| `data/{供应商}/materials/新增明细表_vN.docx` | 明细表（按版本） |
| `email-database/QA/*原料企业质量标准.docx` 等 | 各物料质量标准（共 16 份原料企业标准 + 16 份企业标准-{申请人}，无 0 字节） |
| `data/_party_c_std_map.json` | 示例丙公司条目编码→质量标准映射 |
| `data/_supplier_mapping.json` | 供应商名称归一化映射 |
| `agreement-automation/质量保证协议模板.docx` | Word 模板 |
| `agreement-automation/平台联系人的对应关系.xlsx` | 平台→审核人/电话/邮箱映射 |
| `agreement-automation/batch_generate_v3.py` | 多物料批量生成脚本（核心） |
| `agreement-automation/generate_qa_agreement_v3.py` | 单物料生成脚本 |
