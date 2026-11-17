# 小学数学AI学情分析——多模型网页版 V2.0

**无需安装 Ollama 即可运行统计分析；任选“零配置复制提示词”或“API直连”实现大模型应用。**

## 1. 本地运行（Windows / PyCharm / python_intro）

在 PyCharm 中打开本文件夹，使用课程 `python_intro` Python 3.11 环境，在终端执行：

```bat
python -m pip install -r requirements.txt
python -m streamlit run app.py
```

浏览器访问 `http://localhost:8501`。

## 2. 部署为别人能访问的网址

1. 登录 GitHub，新建私有或公开仓库，把 `app.py`、`demo_data.json`、`requirements.txt`、`.streamlit/config.toml`、示例Excel上传到仓库根目录（或直接上传整个文件夹内容）。
2. 打开 https://share.streamlit.io/ ，点击 `Create app` / `Deploy`，选 GitHub 仓库、分支和入口 `app.py`。
3. 部署成功后得到 `https://xxxx.streamlit.app/` 的可访问网址。网址由部署平台分配，**本压缩包本身不是已上线的网站**。
4. 访问者无需安装 Python，但网站运行须有网络连接；首次启动可能需要一些时间。

## 3. 模型选择

- **零配置**：选“零配置：复制到模型网页”，复制生成的汇总数据提示词，粘贴到通义千问、DeepSeek、Kimi 或 ChatGPT 官方网页；回复粘贴回系统即可下载 Word。此方式**不由网站直接调用模型**，无需 API Key。
- **API直连**：选“API直连：在网站内生成”，填平台 API Key，选择模型 ID，点击生成。支持兼容 OpenAI Chat Completions 的平台。**网页聊天账号 ≠ API密钥**；接口调用可能收费。
- 默认接口示例（不同地域/账号可能要求改用控制台专属域名）：
  - 千问（阿里云百炼）：`https://dashscope.aliyuncs.com/compatible-mode/v1`、`qwen-plus`。
  - DeepSeek：`https://api.deepseek.com`、`deepseek-chat`（实际可用 ID 以控制台为准）。
  - Kimi（月之暗面）：`https://api.moonshot.cn/v1`、`moonshot-v1-8k`（实际可用 ID 以控制台为准）。
  - 其他平台：自行填写兼容协议 URL 和模型 ID。

## 4. 数据格式及隐私

默认40名匿名学生的**教学模拟成绩**，不得当作真实教学案例或实测提效数据。上传Excel必须有 `学生编号` 以及以下五列：
`同分母分数加减法`(20), `异分母分数加减法`(25), `分数混合运算`(20), `简便运算`(15), `解决实际问题`(20)。

**重要**：上传Excel时，文件会传给运行网页的服务器，不能保证完全离线。请不要上传真实姓名等敏感信息，实际课堂使用应先经过学校批准并匿名化；调用模型的提示词仅发送班级汇总统计，不发送学生级明细。API Key仅作为当前运行时输入使用，不能公开展示，不要写进GitHub代码。

## 5. 验证与限制

- Python统计不依赖模型，可立即演示。
- API能否成功、费用和模型名称取决于对应平台账户、地域和网络。
- 模型数学计算及题目可能出错；必须人工校验后用于教学。
- Streamlit Community Cloud不是长期隐私数据管理平台；演示建议仅使用模拟数据。
