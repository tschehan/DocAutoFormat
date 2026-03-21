# DocAutoFormat

这是一个能自动修改word文档格式的程序，只需传入一个docx文档，以及一段修改要求的文字，即可完成格式修改。  
本项目基于 Python-docx 和本地部署的 AI，旨在面向隐私性或保密性要求较高的文档场景，更高效、更灵活地实现格式与内容修改需求。  
限于开发者精力有限，目前版本仅为demo版，即0.1版本。目前能实现的修改需求仅限于调整【字体、字号、加粗、行距、对齐方式】，能区分的文章结构仅限于【总标题、副标题、一级标题、二级标题、三级标题、正文】，能区分的字符类型仅限于【中文字符、西文字符】。  
感谢ChatGPT（codex）大人在项目中发挥的中流砥柱作用。

### 运行所需配置

本项目默认使用本地模型 `qwen3:8b`。硬件门槛主要来源于本地模型 `qwen3:8b` 的推理内存。推荐使用独立GPU运行，且GPU专用内存最好在6GB及以上；若无独立GPU也可以使用CPU运行，但是需要系统内存剩余至少约8GB。

### 运行方式

1. 安装 Python（建议 3.10 及以上）
2. 安装 Ollama
3. 打开终端，进入项目目录
4. 创建虚拟环境

```powershell
py -3 -m venv .venv
.\.venv\Scripts\Activate.ps1
```

5. 安装依赖

```powershell
pip install -r requirements.txt
```

6. 执行以下命令以下载模型

```powershell
ollama pull qwen3:8b
```

7. 启动图形界面

```powershell
py -3 ui.py
```

也可以直接双击 [run_ui.bat](c:/Users/75061/Desktop/tree/have-a-try/DocAutoFormat_0.1/run_ui.bat) 运行。

注：如果处理结果处出现“错误信息：Structure Ollama request failed.”，基本是因为Ollama未在后台启动。一般来说，Ollama在Windows系统会自启动，所以可以重启，也可以手动打开Ollama。





