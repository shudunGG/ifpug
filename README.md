# COSMIC Function Point Calculator

该项目提供了一个按照 COSMIC 方法度量软件功能点的命令行工具。通过编写一份结构化的配置文件（YAML/JSON），即可对系统的功能过程及对应的数据移动进行记录，并自动生成包含明细的 Excel 报表。

## 功能特性

- 支持 Entry (E)、Exit (X)、Read (R)、Write (W) 四类数据移动的建模与统计。
- 按功能过程输出详细的 CFP 汇总信息。
- 输出数据移动明细，包含触发方、业务对象、代码定位、备注等字段。
- 生成 Excel 工作簿，包含 Summary、Functional Processes、Data Movements 三个工作表，方便审阅与归档。

## 安装依赖

该工具完全基于 Python 标准库实现，不需要额外依赖。建议使用虚拟环境运行：

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt  # 文件中仅用于提示，无需额外安装
```

## 准备配置文件

可以参考仓库中的 [`example_measurement.yaml`](example_measurement.yaml) 文件。其结构包括系统信息、对象列表以及功能过程定义：

```yaml
system:
  name: Cosmic Demo System
  boundary: External users and partner APIs
  description: Sample measurement definition illustrating COSMIC counting.
  persistence_resources:
    - PostgreSQL order database
    - Redis cache (write-through to PostgreSQL)
  external_actors:
    - Customer portal
    - Partner API Gateway
objects_of_interest:
  - Order
  - Invoice
  - Payment
functional_processes:
  - name: Submit Order
    trigger: Customer submits order form
    object_of_interest: Order
    description: Validates and persists a newly submitted order.
    data_movements:
      - type: E
        description: Receive order payload from customer portal.
        code_reference: services/order_service.py:create_order
      - type: R
        description: Fetch customer profile for validation.
        object_of_interest: Customer
        code_reference: repositories/customer_repository.py:get_customer
      - type: W
        description: Persist order into PostgreSQL database.
        code_reference: repositories/order_repository.py:save_order
      - type: X
        description: Return order confirmation with order identifier.
        code_reference: services/order_service.py:create_order
```

在 `data_movements` 中可以为每一次有业务意义的数据移动记录类型、业务描述、涉及的业务对象以及代码定位，程序会按照 COSMIC 计量规则自动统计功能点。

## 生成 Excel 报表

```bash
python cosmic_cli.py example_measurement.yaml --output cosmic_report.xlsx
```

执行完成后会在当前目录生成 `cosmic_report.xlsx`，其中包含：

- **Summary**：系统级总 CFP。
- **Functional Processes**：每个功能过程的 E/X/R/W 数量与总 CFP。
- **Data Movements**：所有数据移动的详细清单。

## 底层模块

如果需要在其他 Python 程序中集成，可直接使用 `cosmic` 包：

```python
from cosmic import load_measurement, ExcelExporter

measurement = load_measurement("example_measurement.yaml")
exporter = ExcelExporter(measurement)
exporter.export("cosmic_report.xlsx")
```

该模块也提供了 `CosmicCalculator` 类，可在生成报表之前获取每个功能过程的统计结果。
