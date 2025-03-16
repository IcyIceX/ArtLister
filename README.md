# ArtLister

## 简介

美术清单助手是一款简单的，针对影视/广告/游戏等行业剧本的美术分析软件。
它支持绝大多数文本文件格式(CSV, JSON, XML, TXT, XLSX, DOC, DOCX, PDF)，通过LLM对传入的文件进行美术信息提取，并格式化输出JSON和XLSX表格文档，方便从业人员用习惯的软件进一步处理信息。

## 功能特点

- 支持多种文件格式输入
- 使用DeepSeek API进行智能分析
- 自动生成标准化的美术清单
- 输出JSON和Excel格式文件
- 简洁直观的图形界面

## 安装说明

1. 克隆项目
```bash
git clone https://github.com/IcyIceX/ArtLister.git
cd ArtLister
```

2. 安装依赖
```bash
pip install -r requirements.txt
```

3. 配置DeepSeek API
- 在GUI界面中输入您的API密钥
- 或在config.yaml中配置默认API密钥

## 使用方法

1. 运行程序
```bash
python app.py
```

2. 在界面中：
- 输入DeepSeek API密钥
- 选择要分析的剧本文件
- 点击"开始处理"
- 等待处理完成，查看生成的清单文件

## 输出说明

程序会在output目录下生成两个文件：
- art_list.json：包含完整的场景道具清单
- art_list.xlsx：便于查看的Excel格式清单

## 依赖说明

- Python 3.6+
- FreeSimpleGUI >= 4.60.0
- OpenAI >= 1.0.0（用于DeepSeek API）
- 其他依赖见requirements.txt

## 许可证

MIT License

## 作者

Chopin

## 版本历史

- v1.0-dev：首个开发版本
  - 基础文件格式支持
  - DeepSeek API集成
  - GUI界面
  - JSON/Excel输出
