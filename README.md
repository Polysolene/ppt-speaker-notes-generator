# PPT 智能备注生成器 (PPT Speaker Notes Generator)

> 专为“全图片格式PPT”设计的智能备注生成工具，利用豆包视觉API自动生成专业、连贯的演讲备注。  
> AI-powered speaker notes generator for image-only PPTs, using Doubao vision API.

---

## 功能特点 | Features

- 🎨 **完美适配图片版PPT**：无需本地Office渲染，直接解析每页图片生成备注  
  Perfect for image-only PPTs: No local Office rendering, parse slide images directly.
- 🧠 **上下文记忆**：自动追踪PPT逻辑主线，让每页备注更连贯、有深度  
  Contextual memory: Tracks PPT logic flow for coherent, insightful notes.
- ⚡ **断点续跑**：支持从指定页码开始，意外中断后可安全恢复进度  
  Resume from breakpoint: Restart from any slide, safe progress recovery.
- 📝 **追加/覆盖模式**：可选择在原有备注基础上追加，或完全覆盖  
  Append/overwrite mode: Choose to append to existing notes or overwrite.
- 🔒 **安全配置**：API密钥通过环境变量管理，避免硬编码泄露  
  Secure config: API keys managed via environment variables, no hardcoding.

---

## 前置要求 | Prerequisites

1.  **PPT格式**：将你的PPT另存为“图片格式PPT”（每页仅包含一张图片）  
   PPT format: Save as "image-only PPT" (one image per slide).
2.  **Python环境**：Python 3.8+  
   Python environment: Python 3.8 or later.
3.  **豆包API**：已开通火山方舟豆包视觉模型（`doubao-1-5-vision-pro-32k-250115`）并获取API Key  
   Doubao API: Access to `doubao-1-5-vision-pro-32k-250115` on Volcano Engine, with API Key.

---

## 快速开始 | Quick Start

### 1. 安装依赖 | Install Dependencies

```bash
pip install openai python-pptx pillow python-dotenv
```

### 2. 配置API密钥 | Configure API Key

在项目根目录创建 `.env` 文件（或设置系统环境变量）：

```env
DOUBAO_API_KEY=你的豆包API密钥
```

> ⚠️ 注意：`.env` 文件切勿提交到代码仓库，已在 `.gitignore` 中自动忽略。  
> ⚠️ Note: Never commit `.env` to repo; it's auto-ignored by `.gitignore`.

### 3. 运行脚本 | Run the Script

1.  **修改核心配置项**（代码中大写变量名）：  
    Modify core configuration items (uppercase variables in code):
    - `PPTX_PATH`：填写你的图片版PPT文件路径（如 `BNCT_PPT.pptx`）  
      Fill in your image-only PPT file path (e.g., `BNCT_PPT.pptx`).
    - `START_SLIDE`：设置起始处理页码（默认1）  
      Set the starting slide number (default 1).
    - `API_SLEEP_TIME`：调整API调用间隔（免费版建议保留1秒）  
      Adjust API call interval (recommend 1s for free tier).
    - `append_mode`：设置备注为追加/覆盖模式（True/False）  
      Set notes to append/overwrite mode (True/False).

2.  **自定义Prompt（可选）**：  
    Customize Prompt (optional):
    - 找到代码中 `generate_speaker_note` 函数内的 `prompt` 字符串；  
      Locate the `prompt` string in the `generate_speaker_note` function.
    - 根据需求修改备注生成规则（如调整字数、分析深度、输出格式等）；  
      Modify note generation rules (e.g., adjust word count, analysis depth, output format).
    - 保持 `<notes>` 和 `<context_update>` 标签不变，仅修改标签内的提示逻辑。  
      Keep `<notes>` and `<context_update>` tags unchanged, only modify the prompt logic inside.

3.  **运行脚本**：  
    Run the script:
    - 将你的图片版PPT文件放入项目目录；  
      Place your image-only PPT in the project folder.
    - 执行以下命令启动生成：  
      Execute the following command to start generation:

```bash
python ppt_notes_generator.py
```

---

## 配置项说明 | Configuration

| 配置项 | 说明 | 默认值 |
|--------|------|--------|
| `DOUBAO_MODEL` | 选择豆包视觉模型 | `doubao-1-5-vision-pro-32k-250115` |
| `START_SLIDE` | 从第几页开始处理 | `1` |
| `API_SLEEP_TIME` | API调用间隔（秒） | `1` |
| `append_mode` | 是否在原有备注基础上追加 | `True` |

---

## 技术原理 | How It Works

1.  **图片提取**：从PPT中提取每页图片  
    Image extraction: Extract images from each slide.
2.  **视觉解析**：调用豆包视觉API分析图片内容  
    Visual analysis: Call Doubao Vision API to analyze image content.
3.  **备注生成**：结合上下文记忆，生成专业演讲备注  
    Note generation: Generate professional notes with contextual memory.
4.  **写入PPT**：将备注写入对应幻灯片的备注页  
    Write to PPT: Insert notes into corresponding slide's notes page.
5.  **进度保存**：自动保存进度，支持断点续跑  
    Progress saving: Auto-save progress, support breakpoint resumption.

---

## 注意事项 | Notes

- 仅支持“全图片格式PPT”，即每页仅包含一张图片  
  Only supports "image-only PPTs" (one image per slide).
- 首次运行会从第1页开始，后续可通过修改 `START_SLIDE` 实现断点续跑  
  First run starts from slide 1; resume by modifying `START_SLIDE`.
- API调用有频率限制，免费版建议保持 `API_SLEEP_TIME=1`  
  API rate limits apply; free tier recommends `API_SLEEP_TIME=1`.
- 生成的备注会自动翻译为中文，非中文内容会被翻译  
  Notes are auto-translated to Chinese; non-Chinese content is translated.

---

## 许可证 | License

本项目采用 **MIT License**，详见 [LICENSE](LICENSE) 文件。  
This project is licensed under the **MIT License** - see [LICENSE](LICENSE) for details.

---

## 贡献 | Contributing

欢迎提交 Issue 和 Pull Request 来改进这个工具！  
Contributions are welcome! Submit issues and PRs to improve this tool.
