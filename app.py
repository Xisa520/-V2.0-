"""五年级数学学情分析｜在线模型可选｜模拟数据默认｜Streamlit Cloud 可部署"""
import hashlib
import io
import json
import urllib.error
import urllib.request
from pathlib import Path

import pandas as pd
import streamlit as st
from docx import Document
from docx.shared import Cm, Pt

ROOT = Path(__file__).resolve().parent
TOPICS = ['同分母分数加减法', '异分母分数加减法', '分数混合运算', '简便运算', '解决实际问题']
MAXIMA = dict(zip(TOPICS, [20, 25, 20, 15, 20]))
PROVIDERS = {
    '阿里云百炼 · 千问': ('https://dashscope.aliyuncs.com/compatible-mode/v1', 'qwen-plus', 'https://bailian.console.aliyun.com/'),
    'DeepSeek 官方': ('https://api.deepseek.com', 'deepseek-chat', 'https://platform.deepseek.com/'),
    'Kimi · Moonshot': ('https://api.moonshot.cn/v1', 'moonshot-v1-8k', 'https://platform.moonshot.cn/'),
    '其他兼容 OpenAI 接口': ('', '', ''),
}

st.set_page_config(page_title='小学数学AI学情分析 | MBA课程项目', page_icon='📚', layout='wide')
st.markdown('''<style>
.stApp{background:#ffffff;color:#17324D} section[data-testid="stSidebar"]{background:#eff5fc}
h1,h2,h3{color:#123C6B} [data-testid="stMetric"]{background:#f1f6fc;padding:15px;border:1px solid #e0ebf7;border-radius:12px}
[data-testid="stTabs"] [role="tab"]{font-weight:600}
</style>''', unsafe_allow_html=True)

@st.cache_data
def demo_data():
    j = json.loads((ROOT / 'demo_data.json').read_text(encoding='utf8'))
    return pd.DataFrame(j['rows'], columns=j['headers'])


def read_excel(file):
    file.seek(0)
    book = pd.ExcelFile(file)
    sheet = '模拟成绩' if '模拟成绩' in book.sheet_names else book.sheet_names[0]
    return pd.read_excel(book, sheet_name=sheet)


def analyze(frame):
    df = frame.copy()
    missing = [k for k in ['学生编号'] + TOPICS if k not in df.columns]
    if missing: raise ValueError('缺少必需列：' + '、'.join(missing))
    if not 1 <= len(df) <= 5000: raise ValueError('学生人数须在1—5000之间。')
    if df['学生编号'].isna().any() or df['学生编号'].astype(str).duplicated().any():
        raise ValueError('学生编号不可缺失或重复。')
    for k, maximum in MAXIMA.items():
        df[k] = pd.to_numeric(df[k], errors='coerce')
        if df[k].isna().any() or (~df[k].between(0, maximum)).any():
            raise ValueError(f'{k}每名学生的得分必须为0—{maximum}之间的数值。')
    df['总分'] = df[TOPICS].sum(axis=1)
    rates = pd.Series({k: df[k].sum()/(len(df)*m) for k, m in MAXIMA.items()}).sort_values()
    dist = pd.cut(df['总分'], [-0.001, 59.999, 69.999, 79.999, 89.999, 100.001], labels=['0—59','60—69','70—79','80—89','90—100']).value_counts().reindex(['0—59','60—69','70—79','80—89','90—100'], fill_value=0)
    return df, rates, dist


def summary_stats(df, rates, dist):
    return {
        '数据性质': '教学模拟数据' if not st.session_state.get('uploaded_real') else '用户自行上传的匿名成绩数据',
        '学生人数': len(df), '试卷总分': 100,
        '平均分': round(float(df['总分'].mean()), 2),
        '中位数': round(float(df['总分'].median()), 2),
        '最高分': round(float(df['总分'].max()), 2),
        '最低分': round(float(df['总分'].min()), 2),
        '及格率': f"{(df['总分'] >= 60).mean():.1%}",
        '各分数段人数': {str(k): int(v) for k, v in dist.items()},
        '知识点得分率': {str(k): f'{v:.1%}' for k, v in rates.items()},
        '薄弱知识点前两位': list(rates.index[:2]),
    }


def prompt_for(kind, stats, extra=''):
    data = json.dumps(stats, ensure_ascii=False, indent=2)
    safety = ('以下是小学五年级数学课堂统计数据。不得编造具体错题、学生表现、实际教学效果和干预耗时；'
              '知识点低得分只表示需要进一步核实的薄弱领域，错误原因只能表述为待验证的假设。'
              '严格区分事实与建议。用通俗易懂的简体中文输出，适合教师直接使用。\n')
    if kind == '学情诊断报告':
        task = ('请写一份800字以内的学情诊断报告，包含：1.班级总体表现（准确引用数据）；'
                '2.薄弱知识点及可能原因（说明证据限制）；3.基础层和提升层的教学策略；'
                '4.下一次测验应如何验证；5.教师人工审核清单。'
                '可提供适合PPT展示的3句结论。')
    else:
        task = ('请根据“薄弱知识点前两位”设计小学五年级《分数的加法和减法》练习：'
                '基础巩固4题、提升拓展3题；按题号输出【题目】【知识点】【参考答案】【逐步解题过程】【难度】；'
                '分数使用1/2这样的格式，应用题条件必须充分、量纲正确。'
                '出题之后再独立验算每一道题，标记可能存在歧义的题。'
                '教师必须在打印前复核全部答案，不得声称模型输出已被人工审核。')
    return safety + '\n任务：' + task + ('\n附加教学要求：' + extra if extra.strip() else '') + '\n\n汇总统计（不包含学生级记录）：\n' + data


def chat_api(provider, base_url, model, key, prompt):
    """OpenAI-compatible /chat/completions; sends aggregate-only prompt."""
    from urllib.parse import urlsplit
    url = base_url.rstrip('/') + '/chat/completions'
    parts = urlsplit(url)
    if parts.scheme != 'https' or not parts.hostname or parts.username or parts.password or parts.port not in (None, 443):
        raise ValueError('接口必须是有效的 HTTPS 地址，且不能在URL中包含账号、密码或非常规端口。')
    if not key.strip(): raise ValueError('请填写该平台提供的 API Key。')
    if not model.strip(): raise ValueError('请填写正确的模型ID。')
    body = {
        'model': model.strip(),
        'messages': [
            {'role':'system','content':'你是一位严谨的小学数学教研员。不可编造不存在的证据；练习答案需人工复核。'},
            {'role':'user','content':prompt},
        ],
        'stream': False,
        'temperature': 0.3,
    }
    req = urllib.request.Request(url, data=json.dumps(body, ensure_ascii=False).encode('utf8'),
        headers={'Authorization': 'Bearer ' + key.strip(), 'Content-Type':'application/json'}, method='POST')
    try:
        with urllib.request.urlopen(req, timeout=90) as r:
            payload = json.loads(r.read(3_000_000).decode('utf8'))
    except urllib.error.HTTPError as exc:
        detail = exc.read(1500).decode('utf8', errors='replace')
        raise RuntimeError(f'模型接口 HTTP {exc.code}：{detail[:350]}（请核对API Key、余额、模型ID和地域）') from None
    except urllib.error.URLError as exc:
        raise RuntimeError(f'网络连接失败：{str(exc.reason)[:250]}') from None
    try:
        content = payload['choices'][0]['message']['content']
        if isinstance(content, list):
            content = '\n'.join(x.get('text','') for x in content if isinstance(x, dict))
        if not isinstance(content, str) or not content.strip(): raise ValueError('empty response')
        return content
    except (KeyError, IndexError, TypeError, ValueError):
        raise RuntimeError('模型响应缺少有效内容，请检查该模型是否兼容 Chat Completions 接口。') from None


def word_bytes(title, content):
    d = Document(); sec = d.sections[0]; sec.top_margin = Cm(2.2); sec.bottom_margin = Cm(2)
    normal = d.styles['Normal']; normal.font.name='Microsoft YaHei'; normal.font.size=Pt(10.5)
    d.add_heading(title, 0)
    d.add_paragraph('提示：本系统默认使用教学模拟数据；如调用大模型，输出须经教师逐项校验。')
    for line in content.splitlines():
        if line.strip(): d.add_paragraph(line.strip())
    buf=io.BytesIO(); d.save(buf); return buf.getvalue()


st.title('📚 小学数学AI学情分析与分层作业设计')
st.caption('Python统计 + 大模型智能诊断 + 分层练习 + Word导出｜MBA课程结课演示')
with st.sidebar:
    st.header('🧭 使用设置')
    input_type=st.radio('数据来源', ['内置40人模拟成绩', '上传匿名Excel成绩表'])
    file=st.file_uploader('选择Excel文件（.xlsx）', type=['xlsx']) if input_type.startswith('上传') else None
    st.session_state['uploaded_real'] = input_type.startswith('上传')
    st.info('请仅上传匿名编号。网页部署时，文件会传到运行该网站的服务器；切勿上传姓名、联系方式等敏感信息。')
    st.markdown('**大模型使用方式**')
    llm_mode = st.radio('选择方式', ['零配置：复制到模型网页', 'API直连：在网站内生成'], index=0)
    if llm_mode.startswith('API'):
        provider=st.selectbox('模型平台', list(PROVIDERS))
        initial_url, initial_model, provider_site=PROVIDERS[provider]
        base_url=st.text_input('API Base URL', value=initial_url, help='只填写到 /v1 或 /compatible-mode/v1，程序自动附加 /chat/completions。')
        model=st.text_input('模型ID', value=initial_model, help='模型ID以你自己平台控制台可用的实际ID为准。')
        api_key=st.text_input('API Key（仅当前会话使用）', type='password')
        st.caption('API通常需要另行开通、可能计费。密钥不写入项目源码、不会被本应用持久化保存。')
        if provider_site: st.link_button('查看模型平台', provider_site)
    else:
        provider=base_url=model=api_key=''
    st.divider()
    st.caption('默认模拟数据，不能宣称为真实学生实验成果。')

if input_type.startswith('上传') and file is None:
    st.warning('请上传匿名Excel数据；也可改选“内置40人模拟成绩”直接展示。')
    st.stop()
try:
    raw=read_excel(file) if file is not None else demo_data()
    df, rates, dist = analyze(raw)
except Exception as exc:
    st.error('数据读取或校验失败：' + str(exc)); st.stop()

stats=summary_stats(df, rates, dist)
fingerprint=hashlib.sha256(json.dumps(stats,ensure_ascii=False,sort_keys=True).encode('utf8')).hexdigest()[:14]
if st.session_state.get('dataset_id') != fingerprint:
    st.session_state['dataset_id']=fingerprint
    st.session_state.pop('llm_results', None)
if 'llm_results' not in st.session_state: st.session_state['llm_results']={}

c1,c2,c3,c4=st.columns(4)
c1.metric('学生人数', len(df)); c2.metric('平均分', f'{df["总分"].mean():.2f}')
c3.metric('最高分 / 最低分', f'{df["总分"].max():.0f} / {df["总分"].min():.0f}')
c4.metric('及格率', f'{(df["总分"] >= 60).mean():.1%}')
if file is None: st.info('当前使用教学模拟数据。所有统计数值由Python根据40条成绩记录实时计算。')

t1,t2,t3,t4=st.tabs(['📊 成绩可视化','🧠 AI学情诊断','✏️ AI分层练习','📂 数据与说明'])
with t1:
    col1,col2=st.columns(2)
    with col1:
        st.subheader('成绩分布'); st.bar_chart(dist.rename('人数'))
    with col2:
        st.subheader('知识点得分率（%）'); st.bar_chart((rates*100).rename('得分率'))
    st.subheader('薄弱知识点（按得分率升序）')
    st.dataframe(pd.DataFrame({'知识点':rates.index, '满分':[MAXIMA[k] for k in rates.index], '得分率':[f'{x:.1%}' for x in rates.values]}), hide_index=True, use_container_width=True)
    st.caption('仅凭各知识点总得分，无法判断每一名学生的具体错误类型；需结合错题和课堂表现人工核实。')

def render_llm_task(task, container):
    with container:
        st.subheader(task)
        extra=st.text_area('附加要求（可选）', key=f'extra_{task}', placeholder='例如：重点关注分数混合运算与解决实际问题；采用人教版小学五年级难度。')
        prompt=prompt_for(task, stats, extra)
        if llm_mode.startswith('零配置'):
            st.success('无需在本网站配置API Key。复制下面的提示词，粘贴到你平时使用的千问、DeepSeek、Kimi或ChatGPT网页即可。')
            st.code(prompt, language=None)
            st.download_button('下载AI提示词 TXT', data=prompt, file_name=('学情诊断' if task=='学情诊断报告' else '分层作业')+'_AI提示词.txt', mime='text/plain', key='prompt_'+task)
            st.caption('在外部模型网页得到的结果可粘贴回来，在本系统导出Word。注意：这不是网站自动调用模型。')
            pasted=st.text_area('粘贴大模型生成的结果（可选）', height=200, key='paste_'+task)
            if pasted.strip():
                st.download_button('将粘贴结果导出Word', data=word_bytes(task+'（需教师审核）', pasted), file_name=task+'.docx', mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document', key='wordpaste_'+task)
        else:
            if st.button('开始调用大模型生成', type='primary', key='run_'+task):
                try:
                    with st.spinner(f'正在请求 {provider}，请勿关闭页面...'):
                        result=chat_api(provider,base_url,model,api_key,prompt)
                    st.session_state['llm_results'][task]=(fingerprint, model, result)
                except Exception as exc:
                    st.error('调用失败：'+str(exc))
            saved=st.session_state['llm_results'].get(task)
            if saved and saved[0]==fingerprint:
                st.success(f'已收到模型回复：{saved[1]}。请检查统计引用及数学题答案。')
                st.markdown(saved[2])
                st.download_button('下载模型输出Word', data=word_bytes(task+'（AI草稿）',saved[2]), file_name=task+'_AI草稿.docx', mime='application/vnd.openxmlformats-officedocument.wordprocessingml.document', key='wordapi_'+task)
            else:
                st.caption('未调用模型。不会把模板结果冒充AI实际生成。')

render_llm_task('学情诊断报告', t2)
render_llm_task('分层练习题', t3)
with t4:
    st.subheader('当前数据（匿名编号）')
    st.dataframe(df, hide_index=True, use_container_width=True)
    st.download_button('下载当前成绩与计算结果 CSV', data=df.to_csv(index=False).encode('utf-8-sig'), file_name='学情统计结果.csv', mime='text/csv')
    st.subheader('模型与数据说明')
    st.markdown('''- **统计计算**：由Python和pandas完成，与模型无关。\n- **智能报告/练习题**：只有真正调用API成功，或从模型网页取得回答，才称为大模型输出。\n- **数据保护**：本程序给模型的提示词仅包含班级汇总统计，不包含学生编号及学生个人成绩。\n- **教学应用**：生成题目必须由教师验算，尤其注意分数运算、应用题条件和答案唯一性。\n- **模型费用**：网页版聊天免费使用权限通常不等于API免费额度；请以模型供应商实际政策为准。\n- **部署**：将本文件夹推送到GitHub，Streamlit Community Cloud主文件选 `app.py`。''')
