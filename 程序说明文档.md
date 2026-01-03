# 中证指数爬虫程序说明文档

## 1. 程序概述

本项目包含两个中证指数网站数据爬取程序，分别采用不同的技术方案和实现思路，适用于不同的场景需求。

| 程序名称 | 核心技术 | 适用场景 | 主要特色 |
|---------|---------|---------|---------|
| 1_爬取中证全指数所有_demo1.py | API请求 + pandas | 直接API获取数据场景 | 高效、简洁、模块化 |
| 2_爬取中证全指数所有_demo1.py | Playwright浏览器模拟 | 动态网站爬取场景 | 动态适配、多格式保存 |

## 2. 程序详细特点

### 2.1 1_爬取中证全指数所有_demo1.py

#### 2.1.1 核心功能
- 直接调用API接口获取中证指数数据
- 使用pandas高效处理数据
- 智能Excel导出，自动计算列宽
- 支持自定义导出选项

#### 2.1.2 技术实现
- **数据获取**：使用`requests`库发送POST请求，直接获取API返回的Excel数据
- **数据处理**：使用`pandas`库读取Excel数据，进行日期转换和字符串处理
- **Excel导出**：使用`openpyxl`库实现智能列宽计算和左对齐设置

#### 2.1.3 代码结构
```python
# 数据爬取函数
def crawl_csindex_data() -> pd.DataFrame:
    # API请求和数据处理
    
# Excel导出函数
def export_to_excel(dataframe: pd.DataFrame, filename: str = None) -> str:
    # 智能列宽计算和格式设置
    
# 主函数
def index_csindex_all(export_excel: bool = True) -> pd.DataFrame:
    # 调用爬取和导出函数
```

#### 2.1.4 特色亮点
- **高效数据获取**：直接调用API，避免浏览器渲染开销
- **模块化设计**：函数职责分离，易于扩展和维护
- **智能Excel导出**：
  - 考虑中文字符宽度，自动计算列宽
  - 首行保持默认对齐，数据行左对齐
  - 支持自定义文件名
- **简洁的代码风格**：153行代码，逻辑清晰，易于理解

### 2.2 2_爬取中证全指数所有_demo1.py

#### 2.2.1 核心功能
- 模拟浏览器行为，处理动态加载的网站
- 实现多种分页爬取策略
- 保存原始数据和结构化数据
- 集成智能Excel导出功能

#### 2.2.2 技术实现
- **动态网站爬取**：使用`playwright`库模拟Chromium浏览器
- **数据提取**：从HTML中提取表格数据，手动结构化
- **分页处理**：实现5种分页策略，包括点击下一页、修改页码、滚动加载等
- **数据保存**：同时保存原始数据和结构化数据到JSON文件
- **Excel导出**：集成了1中的智能Excel导出功能

#### 2.2.3 代码结构
```python
# Excel导出功能（集成自1_程序）
def export_to_excel(data: pd.DataFrame | list | None = None, filename: str = None) -> str:
    # 智能列宽计算和格式设置
    
# 主爬取函数
def crawl_all_csindex_data() -> pd.DataFrame:
    # 浏览器初始化
    # 页面访问和数据提取
    # 分页处理
    # 数据保存和导出
```

#### 2.2.4 特色亮点
- **动态网站适配**：适合处理JavaScript动态加载的网站
- **多种分页策略**：5种分页方式，提高爬取成功率
- **多格式数据保存**：
  - 原始数据保存为JSON
  - 结构化数据保存为JSON
  - 支持从文件读取数据导出Excel
- **集成智能Excel导出**：
  - 自动计算列宽（考虑中文字符）
  - 数据行左对齐
  - 美观的输出格式
- **详细的爬取日志**：便于监控爬取过程和调试

## 3. 适用场景对比

| 场景类型 | 推荐程序 | 原因 |
|---------|---------|------|
| API接口稳定可访问 | 1_程序 | 高效、简洁、资源消耗低 |
| 网站动态加载数据 | 2_程序 | 能处理JavaScript渲染的内容 |
| 需要多种数据格式保存 | 2_程序 | 同时保存原始数据和结构化数据 |
| 简单快速爬取 | 1_程序 | 代码简洁，易于部署和运行 |
| 复杂网站结构 | 2_程序 | 多种分页策略，适应性强 |

## 4. 使用方法

### 4.1 1_爬取中证全指数所有_demo1.py

#### 4.1.1 基本使用
```python
# 直接调用主函数，默认导出Excel
import 1_爬取中证全指数所有_demo1 as csindex

df = csindex.index_csindex_all()
```

#### 4.1.2 自定义选项
```python
# 只爬取数据，不导出Excel
df = csindex.index_csindex_all(export_excel=False)

# 单独调用爬取和导出函数
df = csindex.crawl_csindex_data()
csindex.export_to_excel(df, filename="自定义文件名.xlsx")
```

### 4.2 2_爬取中证全指数所有_demo1.py

#### 4.2.1 基本使用
```python
# 直接运行程序，会自动爬取并导出数据
import 2_爬取中证全指数所有_demo1 as csindex_crawler

csindex_crawler.crawl_all_csindex_data()
```

#### 4.2.2 单独使用导出功能
```python
# 从JSON文件读取数据导出Excel
filename = csindex_crawler.export_to_excel()

# 传入数据列表导出Excel
filename = csindex_crawler.export_to_excel(data_list)

# 传入DataFrame导出Excel
filename = csindex_crawler.export_to_excel(dataframe)
```

## 5. 依赖说明

### 5.1 1_爬取中证全指数所有_demo1.py
```
pandas>=1.0.0
requests>=2.0.0
openpyxl>=3.0.0
```

### 5.2 2_爬取中证全指数所有_demo1.py
```
playwright>=1.0.0
pandas>=1.0.0
openpyxl>=3.0.0
json
```

## 6. 运行环境

- Python 3.8+
- Windows/macOS/Linux
- 对于2_程序，需要安装Playwright浏览器驱动：
  ```bash
  playwright install chromium
  ```

## 7. 输出文件说明

### 7.1 1_爬取中证全指数所有_demo1.py
```
中证指数列表_YYYY-MM-DD.xlsx  # 智能导出的Excel文件
```

### 7.2 2_爬取中证全指数所有_demo1.py
```
csindex_raw_data.json         # 原始爬取数据
csindex_structured_data.json  # 结构化处理后的数据
中证指数有限公司_指数列表_YYYY-MM-DD.xlsx  # 智能导出的Excel文件
```

## 8. 性能对比

| 性能指标 | 1_程序 | 2_程序 | 差异原因 |
|---------|-------|-------|---------|
| 爬取速度 | 快 | 慢 | 1_程序直接调用API，2_程序需要浏览器渲染 |
| 资源消耗 | 低 | 高 | 2_程序需要加载浏览器进程 |
| 代码复杂度 | 低 | 高 | 2_程序需要处理动态网站和多种分页策略 |
| 维护成本 | 低 | 中 | 1_程序模块化设计，2_程序逻辑更复杂 |

## 9. 总结

### 9.1 1_爬取中证全指数所有_demo1.py 优势
- 高效简洁，资源消耗低
- 模块化设计，易于维护和扩展
- 智能Excel导出，用户体验好
- 适合API稳定的场景

### 9.2 2_爬取中证全指数所有_demo1.py 优势
- 动态网站适配能力强
- 多种分页策略，爬取成功率高
- 多格式数据保存，便于后续分析
- 适合复杂网站结构

### 9.3 选择建议
- 如果目标网站提供稳定的API，优先选择1_程序
- 如果需要处理动态加载的网站，选择2_程序
- 如果需要多种数据格式保存，选择2_程序
- 如果追求简单快速部署，选择1_程序

两个程序各有特色，用户可以根据实际需求选择合适的程序，或结合使用以发挥各自优势。