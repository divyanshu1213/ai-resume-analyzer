import streamlit as st
from google import genai
import PyPDF2
import docx
import io
import time

# ── PAGE CONFIG ────────────────────────────────────────────
st.set_page_config(
    page_title="ResumeAI — Smart Resume Analyzer",
    page_icon="🎯",
    layout="wide"
)

# ── THEME TOGGLE ───────────────────────────────────────────
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = True

if "analysis_done" not in st.session_state:
    st.session_state.analysis_done = False

if "result" not in st.session_state:
    st.session_state.result = ""

# ── DYNAMIC STYLES ─────────────────────────────────────────
def get_styles(dark):
    bg = "#0D1117" if dark else "#F8FAFF"
    bg2 = "#161B22" if dark else "#FFFFFF"
    bg3 = "#1C2333" if dark else "#EEF2FF"
    text = "#E8EEF8" if dark else "#111827"
    muted = "#7A8BA8" if dark else "#6B7280"
    border = "#1A2A45" if dark else "#E5E7EB"
    card = "#0D1830" if dark else "#FFFFFF"

    return f"""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Space+Grotesk:wght@300;400;500;600;700&family=JetBrains+Mono:wght@400;700&display=swap');

    * {{ font-family: 'Space Grotesk', sans-serif !important; }}

    .stApp {{
        background-color: {bg};
        color: {text};
    }}

    /* Hero Section */
    .hero {{
        background: {'linear-gradient(135deg, #0D1117 0%, #0D1830 50%, #0D1117 100%)' if dark else 'linear-gradient(135deg, #EEF2FF 0%, #E0E7FF 50%, #EEF2FF 100%)'};
        border-radius: 20px;
        padding: 60px 40px;
        text-align: center;
        margin-bottom: 32px;
        border: 1px solid {border};
        position: relative;
        overflow: hidden;
    }}

    .hero::before {{
        content: '';
        position: absolute;
        top: 0; left: 0; right: 0; bottom: 0;
        background-image: linear-gradient({border} 1px, transparent 1px),
                          linear-gradient(90deg, {border} 1px, transparent 1px);
        background-size: 40px 40px;
        opacity: 0.3;
    }}

    .hero-title {{
        font-size: 3.5rem;
        font-weight: 800;
        background: linear-gradient(90deg, #1E6FFF, #00C8FF);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        margin-bottom: 16px;
        position: relative;
        z-index: 1;
    }}

    .hero-subtitle {{
        font-size: 1.15rem;
        color: {muted};
        margin-bottom: 32px;
        position: relative;
        z-index: 1;
    }}

    .hero-stats {{
        display: flex;
        justify-content: center;
        gap: 48px;
        position: relative;
        z-index: 1;
    }}

    .stat {{
        text-align: center;
    }}

    .stat-num {{
        font-size: 2rem;
        font-weight: 800;
        background: linear-gradient(90deg, #1E6FFF, #00C8FF);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
    }}

    .stat-label {{
        font-size: 0.8rem;
        color: {muted};
        font-family: 'JetBrains Mono', monospace !important;
    }}

    /* Cards */
    .card {{
        background: {card};
        border: 1px solid {border};
        border-radius: 16px;
        padding: 28px;
        margin-bottom: 20px;
        transition: all 0.3s ease;
    }}

    .card:hover {{
        border-color: #1E6FFF;
        box-shadow: 0 0 24px rgba(30,111,255,0.1);
    }}

    .card-title {{
        font-size: 1rem;
        font-weight: 700;
        color: {'#00C8FF' if dark else '#1E6FFF'};
        font-family: 'JetBrains Mono', monospace !important;
        letter-spacing: 0.1em;
        text-transform: uppercase;
        margin-bottom: 16px;
        display: flex;
        align-items: center;
        gap: 8px;
    }}

    /* Score Circle */
    .score-container {{
        display: flex;
        align-items: center;
        justify-content: center;
        gap: 48px;
        padding: 32px;
        background: {bg3};
        border-radius: 16px;
        margin-bottom: 24px;
        border: 1px solid {border};
    }}

    .score-circle {{
        width: 140px;
        height: 140px;
        border-radius: 50%;
        background: conic-gradient(#1E6FFF var(--score), {border} 0);
        display: flex;
        align-items: center;
        justify-content: center;
        position: relative;
    }}

    .score-inner {{
        width: 110px;
        height: 110px;
        border-radius: 50%;
        background: {bg2};
        display: flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
    }}

    .score-num {{
        font-size: 2rem;
        font-weight: 800;
        background: linear-gradient(90deg, #1E6FFF, #00C8FF);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        background-clip: text;
        line-height: 1;
    }}

    .score-label {{
        font-size: 0.65rem;
        color: {muted};
        font-family: 'JetBrains Mono', monospace !important;
    }}

    /* Result sections */
    .result-section {{
        background: {bg3};
        border-radius: 12px;
        padding: 20px 24px;
        margin-bottom: 16px;
        border-left: 3px solid #1E6FFF;
        animation: slideIn 0.5s ease;
    }}

    @keyframes slideIn {{
        from {{ opacity: 0; transform: translateX(-20px); }}
        to {{ opacity: 1; transform: translateX(0); }}
    }}

    .result-section h2 {{
        color: {'#00C8FF' if dark else '#1E6FFF'} !important;
        font-size: 1rem !important;
        font-weight: 700 !important;
        margin-bottom: 12px !important;
    }}

    /* Streamlit overrides */
    .stTextInput input, .stTextArea textarea {{
        background-color: {bg2} !important;
        color: {text} !important;
        border: 1px solid {border} !important;
        border-radius: 10px !important;
        font-family: 'Space Grotesk', sans-serif !important;
    }}

    .stTextInput input:focus, .stTextArea textarea:focus {{
        border-color: #1E6FFF !important;
        box-shadow: 0 0 0 2px rgba(30,111,255,0.2) !important;
    }}

    .stFileUploader {{
        background: {bg2} !important;
        border: 2px dashed {border} !important;
        border-radius: 12px !important;
        padding: 20px !important;
    }}

    .stFileUploader:hover {{
        border-color: #1E6FFF !important;
    }}

    .stButton > button {{
        background: linear-gradient(90deg, #1E6FFF, #0052CC) !important;
        color: white !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 14px 28px !important;
        font-weight: 700 !important;
        font-size: 1rem !important;
        width: 100% !important;
        transition: all 0.3s ease !important;
        font-family: 'Space Grotesk', sans-serif !important;
    }}

    .stButton > button:hover {{
        transform: translateY(-2px) !important;
        box-shadow: 0 8px 24px rgba(30,111,255,0.4) !important;
    }}

    .stRadio > div {{
        background: {bg2};
        border-radius: 10px;
        padding: 12px;
        border: 1px solid {border};
    }}

    .stProgress .st-bo {{
        background: linear-gradient(90deg, #1E6FFF, #00C8FF) !important;
    }}

    /* Toggle button */
    .toggle-btn {{
        background: {bg2};
        border: 1px solid {border};
        border-radius: 20px;
        padding: 6px 16px;
        color: {text};
        cursor: pointer;
        font-size: 0.85rem;
    }}

    /* Badge */
    .badge {{
        display: inline-block;
        background: rgba(30,111,255,0.15);
        color: #1E6FFF;
        border: 1px solid rgba(30,111,255,0.3);
        border-radius: 20px;
        padding: 3px 12px;
        font-size: 0.75rem;
        font-family: 'JetBrains Mono', monospace !important;
        margin: 3px;
    }}

    /* Step indicators */
    .step {{
        display: flex;
        align-items: center;
        gap: 12px;
        padding: 12px 0;
        border-bottom: 1px solid {border};
    }}

    .step:last-child {{ border-bottom: none; }}

    .step-num {{
        width: 28px;
        height: 28px;
        border-radius: 50%;
        background: linear-gradient(135deg, #1E6FFF, #00C8FF);
        color: white;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 0.8rem;
        font-weight: 700;
        flex-shrink: 0;
    }}

    .step-text {{
        font-size: 0.9rem;
        color: {muted};
    }}

    div[data-testid="stMarkdownContainer"] p {{
        color: {text} !important;
    }}

    .footer {{
        text-align: center;
        padding: 32px 0 16px;
        color: {muted};
        font-size: 0.85rem;
        border-top: 1px solid {border};
        margin-top: 48px;
    }}

    /* Hide streamlit branding */
    #MainMenu, footer, header {{ visibility: hidden; }}
    </style>
    """

st.markdown(get_styles(st.session_state.dark_mode), unsafe_allow_html=True)

dark = st.session_state.dark_mode
text_color = "#E8EEF8" if dark else "#111827"
muted_color = "#7A8BA8" if dark else "#6B7280"

# ── TOP BAR ────────────────────────────────────────────────
col1, col2 = st.columns([6, 1])
with col1:
    st.markdown(f"<span style='font-family:JetBrains Mono,monospace;color:#1E6FFF;font-size:1.1rem;font-weight:700;'>&lt;ResumeAI/&gt;</span>", unsafe_allow_html=True)
with col2:
    theme_label = "☀️ Light" if dark else "🌙 Dark"
    if st.button(theme_label):
        st.session_state.dark_mode = not st.session_state.dark_mode
        st.rerun()

# ── HERO ───────────────────────────────────────────────────
st.markdown(f"""
<div class="hero">
    <div class="hero-title">AI Resume Analyzer</div>
    <div class="hero-subtitle">
        Get instant, detailed feedback on your resume powered by Google Gemini AI.<br/>
        Improve your ATS score and land more interviews.
    </div>
    <div class="hero-stats">
        <div class="stat">
            <div class="stat-num">98%</div>
            <div class="stat-label">// accuracy</div>
        </div>
        <div class="stat">
            <div class="stat-num">&lt;10s</div>
            <div class="stat-label">// analysis time</div>
        </div>
        <div class="stat">
            <div class="stat-num">ATS</div>
            <div class="stat-label">// optimized</div>
        </div>
    </div>
</div>
""", unsafe_allow_html=True)

# ── MAIN LAYOUT ────────────────────────────────────────────
left_col, right_col = st.columns([1.2, 1])

with left_col:
    # API Key
    st.markdown(f"""
    <div class="card">
        <div class="card-title">⚙️ // Setup</div>
    </div>
    """, unsafe_allow_html=True)

    api_key = st.text_input(
        "Gemini API Key",
        type="password",
        placeholder="AIzaSy...",
        label_visibility="collapsed"
    )
    st.markdown(f"<p style='font-size:0.78rem;color:{muted_color};margin-top:-8px;'>🔑 Get your free key at <a href='https://aistudio.google.com' target='_blank' style='color:#1E6FFF;'>aistudio.google.com</a></p>", unsafe_allow_html=True)

    st.markdown("<br/>", unsafe_allow_html=True)

    # Resume Upload
    st.markdown(f"""
    <div class="card-title" style="color:{'#00C8FF' if dark else '#1E6FFF'};">📋 // Upload Resume</div>
    """, unsafe_allow_html=True)

    input_method = st.radio(
        "Input method",
        ["📁 Upload PDF/Word", "✏️ Paste Text"],
        label_visibility="collapsed"
    )

    resume_text = ""

    if "PDF" in input_method:
        uploaded_file = st.file_uploader(
            "Drop your resume here",
            type=["pdf", "docx"],
            label_visibility="collapsed"
        )
        if uploaded_file:
            if uploaded_file.type == "application/pdf":
                pdf_reader = PyPDF2.PdfReader(io.BytesIO(uploaded_file.read()))
                for page in pdf_reader.pages:
                    resume_text += page.extract_text()
                st.success("✅ PDF loaded successfully!")
            elif "wordprocessingml" in uploaded_file.type:
                doc = docx.Document(io.BytesIO(uploaded_file.read()))
                for para in doc.paragraphs:
                    resume_text += para.text + "\n"
                st.success("✅ Word document loaded!")
    else:
        resume_text = st.text_area(
            "Resume text",
            height=200,
            placeholder="Paste your complete resume here...",
            label_visibility="collapsed"
        )

    st.markdown("<br/>", unsafe_allow_html=True)

    # Job Description
    st.markdown(f"""
    <div class="card-title" style="color:{'#00C8FF' if dark else '#1E6FFF'};">💼 // Job Description <span style="font-size:0.7rem;opacity:0.6;">(optional)</span></div>
    """, unsafe_allow_html=True)

    job_desc = st.text_area(
        "Job description",
        height=150,
        placeholder="Paste job description for targeted analysis...",
        label_visibility="collapsed"
    )

    st.markdown("<br/>", unsafe_allow_html=True)

    # Analyze Button
    analyze = st.button("🚀 Analyze My Resume", use_container_width=True)

with right_col:
    # How it works
    st.markdown(f"""
    <div class="card">
        <div class="card-title">📖 // How It Works</div>
        <div class="step">
            <div class="step-num">1</div>
            <div class="step-text">Enter your Gemini API key</div>
        </div>
        <div class="step">
            <div class="step-num">2</div>
            <div class="step-text">Upload your resume (PDF or Word)</div>
        </div>
        <div class="step">
            <div class="step-num">3</div>
            <div class="step-text">Optionally add the job description</div>
        </div>
        <div class="step">
            <div class="step-num">4</div>
            <div class="step-text">Click Analyze and get instant AI feedback</div>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # What you get
    st.markdown(f"""
    <div class="card">
        <div class="card-title">🎯 // What You Get</div>
        <div style="display:flex;flex-wrap:wrap;gap:6px;margin-top:8px;">
            <span class="badge">ATS Score</span>
            <span class="badge">Strengths</span>
            <span class="badge">Weaknesses</span>
            <span class="badge">Missing Keywords</span>
            <span class="badge">Suggestions</span>
            <span class="badge">Summary</span>
            <span class="badge">Download Report</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    # Tips
    st.markdown(f"""
    <div class="card">
        <div class="card-title">💡 // Pro Tips</div>
        <p style="font-size:0.85rem;color:{muted_color};line-height:1.8;margin:0;">
        ✦ Add a job description for targeted feedback<br/>
        ✦ Use the suggestions to tailor each application<br/>
        ✦ Aim for ATS score above 75/100<br/>
        ✦ Download your report to track improvements
        </p>
    </div>
    """, unsafe_allow_html=True)

# ── ANALYSIS ───────────────────────────────────────────────
if analyze:
    if not api_key:
        st.error("❌ Please enter your Gemini API key!")
    elif not resume_text.strip():
        st.error("❌ Please upload or paste your resume!")
    else:
        st.markdown("---")
        st.markdown(f"<h2 style='color:{'#00C8FF' if dark else '#1E6FFF'};text-align:center;'>🎯 Analysis Results</h2>", unsafe_allow_html=True)

        # Progress bar animation
        progress_bar = st.progress(0)
        status = st.empty()

        stages = [
            (15, "📄 Reading resume..."),
            (35, "🔍 Identifying keywords..."),
            (55, "🤖 Running AI analysis..."),
            (75, "📊 Calculating ATS score..."),
            (90, "💡 Generating suggestions..."),
            (100, "✅ Analysis complete!")
        ]

        try:
            client = genai.Client(api_key=api_key)

            for progress, message in stages[:-2]:
                progress_bar.progress(progress)
                status.markdown(f"<p style='text-align:center;color:{muted_color};'>{message}</p>", unsafe_allow_html=True)
                time.sleep(0.4)

            # Build prompt
            if job_desc.strip():
                prompt = f"""
You are an expert resume reviewer and career coach with 10+ years of experience in tech hiring.
Analyze this resume against the provided job description.

RESUME:
{resume_text}

JOB DESCRIPTION:
{job_desc}

Provide analysis in EXACTLY this format:

## 📊 Overall ATS Score: [X/100]

## ✅ Strengths
- [strength 1]
- [strength 2]
- [strength 3]

## ⚠️ Weaknesses
- [weakness 1]
- [weakness 2]
- [weakness 3]

## 🎯 Missing Keywords
- [keyword 1]
- [keyword 2]
- [keyword 3]

## 💡 Suggestions to Improve
1. [suggestion 1]
2. [suggestion 2]
3. [suggestion 3]
4. [suggestion 4]
5. [suggestion 5]

## 📝 Summary
[2-3 sentence overall assessment]
"""
            else:
                prompt = f"""
You are an expert resume reviewer and career coach with 10+ years of experience in tech hiring.
Analyze this resume in detail.

RESUME:
{resume_text}

Provide analysis in EXACTLY this format:

## 📊 Overall ATS Score: [X/100]

## ✅ Strengths
- [strength 1]
- [strength 2]
- [strength 3]

## ⚠️ Weaknesses
- [weakness 1]
- [weakness 2]
- [weakness 3]

## 🔍 Missing Elements
- [missing 1]
- [missing 2]
- [missing 3]

## 💡 Suggestions to Improve
1. [suggestion 1]
2. [suggestion 2]
3. [suggestion 3]
4. [suggestion 4]
5. [suggestion 5]

## 📝 Summary
[2-3 sentence overall assessment]
"""

            response = client.models.generate_content(
                model="gemini-2.5-flash",
                contents=prompt
            )
            result = response.text

            for progress, message in stages[-2:]:
                progress_bar.progress(progress)
                status.markdown(f"<p style='text-align:center;color:{muted_color};'>{message}</p>", unsafe_allow_html=True)
                time.sleep(0.3)

            time.sleep(0.5)
            progress_bar.empty()
            status.empty()

            st.session_state.result = result
            st.session_state.analysis_done = True

        except Exception as e:
            progress_bar.empty()
            status.empty()
            st.error(f"❌ Error: {str(e)}")
            st.info("💡 Make sure your API key is correct!")

# ── SHOW RESULTS ───────────────────────────────────────────
if st.session_state.analysis_done and st.session_state.result:
    result = st.session_state.result

    # Display formatted results
    st.markdown(result)

    # Download
    col1, col2, col3 = st.columns([1,2,1])
    with col2:
        st.download_button(
            label="⬇️ Download Full Report",
            data=result,
            file_name="resume_analysis.txt",
            mime="text/plain",
            use_container_width=True
        )

# ── FOOTER ─────────────────────────────────────────────────
st.markdown(f"""
<div class="footer">
    Built by <a href="https://divyanshu1213.github.io" style="color:#1E6FFF;text-decoration:none;font-weight:600;">Divyanshu Raj</a>
    &nbsp;•&nbsp; Powered by Google Gemini AI
    &nbsp;•&nbsp; <a href="https://github.com/divyanshu1213/ai-resume-analyzer" style="color:#1E6FFF;text-decoration:none;">GitHub</a>
</div>
""", unsafe_allow_html=True)