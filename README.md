# 行业数据库自动化更新 Pipeline 架构

## 1. 设计初衷

针对日益庞大的行业数据库（1000+ 个 Sheet）更新任务，传统的“硬编码 + 复制粘贴”方式带来了极大的维护成本、冗余代码和执行瓶颈。

新架构采用 **Data-Driven Configuration (配置驱动)** 与 **Pipeline (流水线)** 设计模式，彻底分离“业务逻辑（做什么）”与“底层能力（怎么做）”。

## 2. 架构分层

新架构按照职责划分为以下几个层次：

- **`configs/` (配置层)**：存放每个行业数据库独立的 YAML 配置文件。业务分析师或开发可以通过修改 YAML，实现指标的增删改查、公式的应用和后处理逻辑的调用。
- **`core_engine/` (能力层)**：
  - `data_reader.py`: 提供全局缓存能力的数据读取器，解决重复读取同一源文件导致的严重 I/O 瓶颈。
  - `data_processor.py`: 内存级/Pandas 数据处理基座。负责在数据写入 Excel 前进行“洗切”（如排序、聚合、季度/YTD统计），供各个行业插件内部调用。
  - `transformers.py`: 表现级/Excel 处理库。存放跨行业的通用 Excel 动作（如直接向单元格注入原生 Excel 公式 `"=B2/C2-1"`、调整格式等），暴露给 YAML 作为通用的 action 路由调用。
  - `action_dispatcher.py`: 核心执行引擎，负责解析 YAML 配置，路由并分发动作。
- **`plugins/` (特异逻辑层)**：存放特定行业独有的复杂数据处理逻辑。例如，航空库特有的季度统计、YTD 特殊差值计算、清洗早期数据等。这些逻辑被封装为插件函数，只在 YAML 中被调用。
- **`run_xxx.py` (调度层)**：负责读取配置并启动引擎（未来可无缝接入 Dagster）。

## 3. 工作流程 (以航空月报为例)

1. **读取配置**：引擎读取 `configs/db_aviation.yaml`。
2. **预加载与缓存**：引擎通过 `DataReader` 一次性将所需的源文件数据预加载到内存中。
3. **打开目标文件**：目标 Excel 模板只被打开一次。
4. **分发执行 (Action Dispatching)**：
   - 对于通用公式（如同比），引擎调用 `transformers.py` 中的通用方法。
   - 对于航空特有逻辑（如 `aviation_write_airline_sheet`），引擎从 `plugins.aviation_plugin` 路由调用对应的特异方法。
5. **保存输出**：所有 Sheet 处理完毕后，目标 Excel 在内存中一次性保存至输出目录。

## 4. 如何新增一个数据库或 Sheet？

1. 如果只需使用已有能力，**完全不需要修改任何 Python 代码**，只需在对应的 `yaml` 文件中复制一份配置，修改 `sheet_name` 和 `indicators` 即可。
2. 如果遇到了新的奇葩处理逻辑，则在 `plugins/` 下新建一个插件文件，编写处理函数，并在 YAML 的 `actions` 中通过 `type: "your_new_action"` 调用。
