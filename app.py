"""
MRU Secure Quiz App – Faculty creates quiz, shares link, all student
results land in ONE Excel sheet that faculty can download / email.
"""

import streamlit as st
import streamlit.components.v1 as components
import re
import json
import time
import random
import hashlib
import smtplib
import io
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
from datetime import datetime, timedelta
from io import BytesIO

import pandas as pd

try:
    from docx import Document
    HAS_DOCX = True
except ImportError:
    HAS_DOCX = False

st.set_page_config(
    page_title="MRU Secure Quiz",
    page_icon="📝",
    layout="centered",
    initial_sidebar_state="expanded",
)

# ═══════════════════════════════════════════════════════════════════════════
# SHARED RESULTS STORE (lives for the lifetime of the Streamlit process)
# ═══════════════════════════════════════════════════════════════════════════
@st.cache_resource
def get_results_store():
    """Shared list of all submissions across all student sessions."""
    return []


@st.cache_resource
def get_quiz_store():
    """Shared dict of published quizzes {quiz_id: quiz_data}."""
    return {}


# ═══════════════════════════════════════════════════════════════════════════
# PARSING HELPERS
# ═══════════════════════════════════════════════════════════════════════════
def extract_text(uploaded_file) -> str:
    name = uploaded_file.name.lower()
    raw = uploaded_file.read()
    if name.endswith((".txt", ".md")):
        return raw.decode("utf-8", errors="ignore")
    if name.endswith(".docx"):
        if not HAS_DOCX:
            st.error("python-docx not installed.")
            return ""
        doc = Document(BytesIO(raw))
        return "\n".join(p.text for p in doc.paragraphs)
    st.error("Use .txt / .md / .docx only.")
    return ""


def parse_questions(text: str) -> list:
    if not text or not text.strip():
        return []
    text = text.replace("\r\n", "\n").replace("\r", "\n")
    text = re.sub(r"\n{3,}", "\n\n", text).strip()
    lines = [l.strip() for l in text.split("\n")]
    questions, current_q, current_opts = [], None, {}
    q_start = re.compile(r"^(?:Q\s*)?(\d+)\s*[\.\)\:\-]\s*(.+)$", re.IGNORECASE)
    opt_pat = re.compile(r"^\(?([a-dA-D])\)?[\.\)\:\-\s]\s*(.+)$")

    def save():
        nonlocal current_q, current_opts
        if current_q and len(current_opts) >= 2:
            questions.append({"text": current_q, "options": dict(current_opts)})
        current_q, current_opts = None, {}

    for line in lines:
        if not line:
            continue
        mq = q_start.match(line)
        if mq:
            save()
            current_q = mq.group(2).strip()
            current_opts = {}
            continue
        mo = opt_pat.match(line)
        if mo and current_q is not None:
            current_opts[mo.group(1).lower()] = mo.group(2).strip()
            continue
        if current_q is not None and not current_opts:
            current_q += " " + line
        elif current_q is None:
            current_q = line
            current_opts = {}
    save()
    return questions


def parse_answers(text: str) -> dict:
    answers = {}
    if not text:
        return answers
    for line in text.replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        line = line.strip()
        if not line:
            continue
        m = re.match(
            r"^(?:Q\s*|Answer\s*)?(\d+)\s*[\.\)\:\-=]?\s*([a-dA-D])\b",
            line, re.IGNORECASE,
        )
        if m:
            answers[int(m.group(1))] = m.group(2).lower()
    return answers


def make_quiz_id(title: str) -> str:
    raw = f"{title}-{datetime.now().isoformat()}-{random.randint(1000,9999)}"
    return hashlib.md5(raw.encode()).hexdigest()[:10]


# ═══════════════════════════════════════════════════════════════════════════
# EXCEL + EMAIL
# ═══════════════════════════════════════════════════════════════════════════
def results_to_excel_bytes(results: list) -> bytes:
    """Build one Excel workbook from all submissions."""
    if not results:
        df = pd.DataFrame(columns=["Name", "Roll No", "Class", "Section", "Score", "Total", "Percentage", "Violations", "Submitted At"])
    else:
        rows = []
        for r in results:
            rows.append({
                "Name": r.get("name", ""),
                "Roll No": r.get("roll", ""),
                "Class": r.get("class", ""),
                "Section": r.get("section", ""),
                "Score": r.get("score", 0),
                "Total": r.get("total", 0),
                "Percentage": round(r.get("score", 0) / max(r.get("total", 1), 1) * 100, 1),
                "Violations": r.get("violations", 0),
                "Submitted At": r.get("submitted_at", ""),
            })
        df = pd.DataFrame(rows)
        # Sort by Class, Section, Roll
        df = df.sort_values(by=["Class", "Section", "Roll No"], ignore_index=True)

    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Results")
    return buf.getvalue()


def send_excel_email(to_email: str, quiz_title: str, excel_bytes: bytes, n_students: int) -> bool:
    """Email the full class Excel to faculty via SMTP secrets."""
    try:
        smtp_cfg = st.secrets.get("smtp", {})
    except Exception:
        smtp_cfg = {}

    host = smtp_cfg.get("host", "")
    port = int(smtp_cfg.get("port", 587))
    user = smtp_cfg.get("user", "")
    password = smtp_cfg.get("password", "")
    from_email = smtp_cfg.get("from_email", user)

    if not (host and user and password):
        return False

    msg = MIMEMultipart()
    msg["From"] = from_email
    msg["To"] = to_email
    msg["Subject"] = f"{quiz_title} – Full Class Results ({n_students} students)"

    body = (
        f"Dear Faculty,\n\n"
        f"Please find attached the complete results sheet for:\n"
        f"Quiz : {quiz_title}\n"
        f"Students submitted : {n_students}\n"
        f"Generated : {datetime.now().strftime('%d %b %Y, %I:%M:%S %p')}\n\n"
        f"Columns: Name | Roll No | Class | Section | Score | Total | Percentage | Violations | Submitted At\n\n"
        f"Regards,\nMRU Secure Quiz App"
    )
    msg.attach(MIMEText(body, "plain"))

    part = MIMEBase("application", "vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    part.set_payload(excel_bytes)
    encoders.encode_base64(part)
    part.add_header("Content-Disposition", "attachment", filename=f"{quiz_title[:40]}_results.xlsx")
    msg.attach(part)

    try:
        with smtplib.SMTP(host, port) as server:
            server.starttls()
            server.login(user, password)
            server.send_message(msg)
        return True
    except Exception as e:
        st.error(f"Email failed: {e}")
        return False


# ═══════════════════════════════════════════════════════════════════════════
# FACULTY PANEL
# ═══════════════════════════════════════════════════════════════════════════
def faculty_panel():
    st.header("👨‍🏫 Faculty Dashboard")

    tab1, tab2 = st.tabs(["📤 Create & Publish Quiz", "📊 Class Results (Excel)"])

    # ── Create Quiz ──────────────────────────────────────────────────────
    with tab1:
        st.caption("Upload questions + answers → get a shareable student link")

        with st.form("create_quiz"):
            title = st.text_input("Quiz Title *", value="MRU Assessment")
            email = st.text_input("Your Email (results Excel will be sent here) *",
                                  placeholder="faculty@college.edu")
            duration = st.number_input("Duration (minutes)", 5, 720, 60, 5)

            q_file = st.file_uploader("Questions (.txt / .docx)", type=["txt", "md", "docx"])
            a_file = st.file_uploader("Answer Key (.txt / .docx)", type=["txt", "md", "docx"])

            st.markdown("""
**Question format**
```
Q1. What is ...?
a) option
b) option
c) option
d) option
```
**Answer format**
```
1. b
2. a
```
""")
            published = st.form_submit_button("🚀 Publish Quiz & Get Student Link", type="primary")

        if published:
            if not email or "@" not in email:
                st.error("Valid faculty email required.")
                return
            if not q_file or not a_file:
                st.error("Upload both questions and answers.")
                return

            q_text = extract_text(q_file)
            a_text = extract_text(a_file)
            questions = parse_questions(q_text)
            answers = parse_answers(a_text)

            if not questions:
                st.error("Could not parse questions. Check format.")
                with st.expander("File preview"):
                    st.code(q_text[:1500])
                return
            if not answers:
                st.error("Could not parse answers. Check format.")
                return

            quiz_id = make_quiz_id(title)
            quiz_data = {
                "id": quiz_id,
                "title": title.strip() or "MRU Quiz",
                "faculty_email": email.strip(),
                "duration_min": int(duration),
                "questions": questions,
                "answers": answers,
                "created_at": datetime.now().isoformat(),
            }
            store = get_quiz_store()
            store[quiz_id] = quiz_data

            # Build student link
            try:
                base = st.get_option("server.baseUrlPath") or ""
            except Exception:
                base = ""

            # Use query params – works on Streamlit Cloud
            student_link = f"?mode=student&quiz={quiz_id}"

            st.success(f"✅ Quiz published!  **{len(questions)} questions** · **{duration} min**")
            st.markdown("### 🔗 Share this link with students")
            st.code(student_link, language=None)
            st.info(
                "Students open your Streamlit app URL and append the link above, "
                "OR select **Student** in the sidebar and enter the Quiz ID: "
                f"**`{quiz_id}`**"
            )
            st.session_state["last_quiz_id"] = quiz_id
            st.balloons()

        # Show last published ID if any
        if "last_quiz_id" in st.session_state:
            st.markdown(f"**Last Quiz ID:** `{st.session_state['last_quiz_id']}`")

        # List active quizzes
        store = get_quiz_store()
        if store:
            st.markdown("---")
            st.subheader("Active Quizzes")
            for qid, q in store.items():
                st.write(f"• **{q['title']}** — ID: `{qid}` — {len(q['questions'])} Qs — {q['duration_min']} min")

    # ── Results ──────────────────────────────────────────────────────────
    with tab2:
        st.caption("All student submissions appear here in one sheet")

        results = get_results_store()
        quiz_store = get_quiz_store()

        # Filter by quiz if multiple
        quiz_ids = list({r.get("quiz_id") for r in results if r.get("quiz_id")})
        if quiz_ids:
            selected = st.selectbox("Filter by Quiz", ["All"] + quiz_ids)
        else:
            selected = "All"

        filtered = results if selected == "All" else [r for r in results if r.get("quiz_id") == selected]

        st.metric("Total Submissions", len(filtered))

        if filtered:
            excel_bytes = results_to_excel_bytes(filtered)
            st.dataframe(
                pd.DataFrame([{
                    "Name": r["name"], "Roll": r["roll"], "Class": r["class"],
                    "Section": r["section"], "Score": f"{r['score']}/{r['total']}",
                    "Submitted": r["submitted_at"],
                } for r in filtered]),
                use_container_width=True,
            )

            col1, col2 = st.columns(2)
            with col1:
                st.download_button(
                    "⬇️ Download Full Excel Sheet",
                    data=excel_bytes,
                    file_name=f"class_results_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )
            with col2:
                # Find faculty email from first result's quiz
                fac_email = ""
                if filtered:
                    qid = filtered[0].get("quiz_id")
                    if qid and qid in quiz_store:
                        fac_email = quiz_store[qid].get("faculty_email", "")
                if not fac_email and quiz_store:
                    fac_email = list(quiz_store.values())[0].get("faculty_email", "")

                if st.button("📧 Email Full Excel to Faculty"):
                    title = "Class Results"
                    if filtered and filtered[0].get("quiz_id") in quiz_store:
                        title = quiz_store[filtered[0]["quiz_id"]]["title"]
                    if fac_email:
                        ok = send_excel_email(fac_email, title, excel_bytes, len(filtered))
                        if ok:
                            st.success(f"Excel emailed to {fac_email}")
                        else:
                            st.warning(
                                "SMTP not configured. Download the Excel and send it manually.\n"
                                "Add [smtp] in Streamlit Secrets to enable auto-email."
                            )
                    else:
                        st.error("No faculty email found.")
        else:
            st.info("No submissions yet. Share the student link and wait for students to finish.")


# ═══════════════════════════════════════════════════════════════════════════
# STUDENT PANEL (with shuffle + security + Class/Section)
# ═══════════════════════════════════════════════════════════════════════════
def student_panel():
    st.header("📝 Student Exam")

    # Resolve quiz
    quiz_store = get_quiz_store()
    params = st.query_params
    quiz_id = params.get("quiz", "")

    if not quiz_id:
        quiz_id = st.text_input("Enter Quiz ID (given by faculty)", value=st.session_state.get("input_quiz_id", ""))
        st.session_state["input_quiz_id"] = quiz_id

    if not quiz_id or quiz_id not in quiz_store:
        if quiz_store:
            st.warning("Enter a valid Quiz ID, or open the full link shared by faculty.")
            st.write("Active quiz IDs:", ", ".join(f"`{k}`" for k in quiz_store.keys()))
        else:
            st.warning("No quiz published yet. Ask faculty to create one first.")
        return

    quiz = quiz_store[quiz_id]
    questions = quiz["questions"]
    answers = quiz["answers"]
    total = len(questions)
    duration = quiz["duration_min"]
    title = quiz["title"]
    faculty_email = quiz["faculty_email"]

    # One-attempt: block if this roll already submitted
    results = get_results_store()

    # ── Start gate ───────────────────────────────────────────────────────
    if "exam_started" not in st.session_state or not st.session_state.exam_started:
        st.subheader(title)
        st.markdown(f"**{total} questions** · **{duration} minutes** · One attempt only")
        st.markdown("""
**Rules**
- Questions & options are **shuffled** for every student
- Stay on this page – do not switch tabs
- Timer starts when you click Start
- Result goes into the class Excel sheet for the faculty
""")
        with st.form("start_form"):
            name = st.text_input("Full Name *")
            roll = st.text_input("Roll Number *")
            cls = st.text_input("Class *", placeholder="e.g. BCA 2nd Year")
            section = st.text_input("Section *", placeholder="e.g. A")
            start = st.form_submit_button("🚀 Start Exam", type="primary")

            if start:
                if not all([name.strip(), roll.strip(), cls.strip(), section.strip()]):
                    st.error("All fields are required.")
                else:
                    # Block duplicate roll for this quiz
                    already = any(
                        r["roll"].lower() == roll.strip().lower() and r.get("quiz_id") == quiz_id
                        for r in results
                    )
                    if already:
                        st.error("This Roll Number has already submitted this quiz.")
                    else:
                        # Shuffle questions & options uniquely per student
                        seed = int(hashlib.md5(f"{roll.strip()}{quiz_id}".encode()).hexdigest()[:8], 16)
                        rng = random.Random(seed)
                        order = list(range(total))
                        rng.shuffle(order)

                        shuffled_qs = []
                        shuffled_ans = {}  # new index -> correct letter after option shuffle
                        for new_i, old_i in enumerate(order):
                            q = questions[old_i]
                            opts = list(q["options"].items())  # [(letter, text), ...]
                            rng.shuffle(opts)
                            # Remap letters a,b,c,d
                            letters = ["a", "b", "c", "d"][:len(opts)]
                            new_opts = {letters[j]: opts[j][1] for j in range(len(opts))}
                            # Find which new letter is the correct answer
                            old_correct = answers.get(old_i + 1, "")
                            new_correct = ""
                            for j, (old_let, text) in enumerate(opts):
                                if old_let == old_correct:
                                    new_correct = letters[j]
                                    break
                            shuffled_qs.append({"text": q["text"], "options": new_opts})
                            shuffled_ans[new_i + 1] = new_correct

                        st.session_state.exam_started = True
                        st.session_state.student_name = name.strip()
                        st.session_state.student_roll = roll.strip()
                        st.session_state.student_class = cls.strip()
                        st.session_state.student_section = section.strip()
                        st.session_state.quiz_id = quiz_id
                        st.session_state.shuffled_qs = shuffled_qs
                        st.session_state.shuffled_ans = shuffled_ans
                        st.session_state.start_time = datetime.now()
                        st.session_state.end_time = datetime.now() + timedelta(minutes=duration)
                        st.session_state.submitted = False
                        st.session_state.violations = 0
                        st.rerun()
        return

    # ── Already submitted in this session ────────────────────────────────
    if st.session_state.get("submitted"):
        st.success("You have already submitted. Your result is in the class sheet.")
        st.metric("Your Score", f"{st.session_state.get('score', 0)} / {total}")
        return

    # ── Live exam ────────────────────────────────────────────────────────
    now = datetime.now()
    remaining = st.session_state.end_time - now

    if remaining.total_seconds() <= 0:
        st.warning("⏰ Time over – submitting automatically…")
        _submit_student(auto=True)
        return

    # Status bar
    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.metric("Clock", now.strftime("%I:%M:%S %p"))
    with c2:
        mins, secs = divmod(int(remaining.total_seconds()), 60)
        hrs, mins = divmod(mins, 60)
        st.metric("Time Left", f"{hrs:02d}:{mins:02d}:{secs:02d}" if hrs else f"{mins:02d}:{secs:02d}")
    with c3:
        st.metric("Roll", st.session_state.student_roll)
    with c4:
        st.metric("Class", f"{st.session_state.student_class}-{st.session_state.student_section}")

    progress = 1 - remaining.total_seconds() / (duration * 60)
    st.progress(min(max(progress, 0), 1))

    # Security reminder + lightweight JS focus detector
    components.html("""
    <script>
    (function(){
      let v = 0;
      document.addEventListener('visibilitychange', function(){
        if(document.hidden){
          v++;
          // Store in parent if possible
          try { window.parent.postMessage({type:'violation', count:v}, '*'); } catch(e){}
        }
      });
    })();
    </script>
    <div style="color:#f59e0b;font-family:sans-serif;font-size:13px;padding:6px 0">
      ⚠️ Stay on this tab. Switching away is recorded.
    </div>
    """, height=40)

    st.subheader(title)
    st.caption(f"{st.session_state.student_name} · {st.session_state.student_roll} · "
               f"{st.session_state.student_class} / {st.session_state.student_section}")

    shuffled_qs = st.session_state.shuffled_qs
    responses = {}

    for idx, q in enumerate(shuffled_qs, start=1):
        st.markdown(f"**Q{idx}. {q['text']}**")
        ordered = sorted(q["options"].keys())
        choice = st.radio(
            f"q{idx}",
            options=ordered,
            format_func=lambda k, o=q["options"]: f"{k.upper()}) {o[k]}",
            key=f"ans_{quiz_id}_{idx}",
            label_visibility="collapsed",
            index=None,
        )
        responses[idx] = choice
        st.divider()

    st.session_state.current_responses = responses

    if st.button("✅ Submit Quiz", type="primary", use_container_width=True):
        unanswered = [i for i, v in responses.items() if v is None]
        if unanswered:
            st.error(f"Answer all questions. Missing: {unanswered}")
        else:
            _submit_student(auto=False)

    # Soft refresh to keep timer alive
    time.sleep(15)
    st.rerun()


def _submit_student(auto: bool = False):
    quiz_id = st.session_state.quiz_id
    quiz = get_quiz_store()[quiz_id]
    shuffled_ans = st.session_state.shuffled_ans
    responses = st.session_state.get("current_responses", {})
    total = len(st.session_state.shuffled_qs)

    score = 0
    for i in range(1, total + 1):
        given = (responses.get(i) or "").lower()
        if given and given == shuffled_ans.get(i, ""):
            score += 1

    record = {
        "quiz_id": quiz_id,
        "quiz_title": quiz["title"],
        "name": st.session_state.student_name,
        "roll": st.session_state.student_roll,
        "class": st.session_state.student_class,
        "section": st.session_state.student_section,
        "score": score,
        "total": total,
        "violations": st.session_state.get("violations", 0),
        "submitted_at": datetime.now().strftime("%d %b %Y %I:%M:%S %p"),
        "auto": auto,
    }

    # Append to shared store
    store = get_results_store()
    # Prevent duplicate
    if not any(r["roll"].lower() == record["roll"].lower() and r["quiz_id"] == quiz_id for r in store):
        store.append(record)

    st.session_state.submitted = True
    st.session_state.score = score

    # Try to email updated full sheet to faculty
    excel_bytes = results_to_excel_bytes([r for r in store if r["quiz_id"] == quiz_id])
    sent = send_excel_email(
        quiz["faculty_email"],
        quiz["title"],
        excel_bytes,
        len([r for r in store if r["quiz_id"] == quiz_id]),
    )

    st.success(f"{'⏰ Time over – ' if auto else ''}Submitted! Score: **{score} / {total}**")
    if sent:
        st.info(f"Updated class Excel emailed to {quiz['faculty_email']}")
    else:
        st.info("Your result is saved in the class sheet. Faculty can download it from the Results tab.")
    st.balloons()
    time.sleep(2)
    st.rerun()


# ═══════════════════════════════════════════════════════════════════════════
# SIDEBAR + ROUTING
# ═══════════════════════════════════════════════════════════════════════════
params = st.query_params
mode_param = params.get("mode", "")

with st.sidebar:
    st.title("📝 MRU Secure Quiz")
    st.markdown("---")

    if mode_param == "student":
        default_role = "Student"
    else:
        default_role = "Faculty"

    role = st.radio("I am a…", ["Faculty", "Student"],
                    index=0 if default_role == "Faculty" else 1)
    st.markdown("---")
    st.caption("Server time (12-hour)")
    st.code(datetime.now().strftime("%d %b %Y  %I:%M:%S %p"))

    results = get_results_store()
    st.metric("Total submissions so far", len(results))

    st.markdown("---")
    st.caption(
        "Faculty: publish quiz → share link\n\n"
        "Students: open link → take exam\n\n"
        "All results → one Excel sheet"
    )

if role == "Faculty":
    faculty_panel()
else:
    student_panel()
