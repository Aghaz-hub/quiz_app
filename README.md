# MRU Secure Quiz App

Faculty creates a quiz → shares a link → **all student results appear in one Excel sheet**.

## Features

| Feature | Detail |
|---------|--------|
| Faculty upload | Questions + answers (Word / Notepad) |
| Shareable link | Student opens link with Quiz ID |
| Shuffle | Questions **and** options reshuffled uniquely per student (by Roll No) |
| Student fields | Name, Roll No, Class, Section |
| Timer | Live 12-hour clock + countdown; auto-submit when time ends |
| One attempt | Same Roll No cannot submit twice for the same quiz |
| **One Excel sheet** | All students of the class/section in a single spreadsheet |
| Email Excel | Faculty can download or email the full class sheet (SMTP) |

## Quick start

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Deploy on Streamlit Cloud

1. Push this folder to GitHub
2. [share.streamlit.io](https://share.streamlit.io) → New app → select `app.py`
3. (Optional) Add SMTP secrets for automatic email of the Excel:

```toml
[smtp]
host = "smtp.gmail.com"
port = 587
user = "your@gmail.com"
password = "app-password"
from_email = "your@gmail.com"
```

## How to use

### Faculty
1. Sidebar → **Faculty**
2. Tab **Create & Publish Quiz** → upload files, set email & duration → Publish
3. Copy the **Quiz ID** / student link and share with the class
4. Tab **Class Results (Excel)** → see all submissions live
5. Click **Download Full Excel Sheet** or **Email Full Excel to Faculty**

### Student
1. Open the link shared by faculty (or choose Student + enter Quiz ID)
2. Fill Name, Roll, Class, Section → Start Exam
3. Questions are shuffled for you
4. Submit before time ends
5. Your row is added to the class Excel sheet automatically

## Excel columns

Name | Roll No | Class | Section | Score | Total | Percentage | Violations | Submitted At

## Notes

- Results are held in memory for the lifetime of the Streamlit app process. On Streamlit Cloud a restart clears them — download the Excel regularly.
- For permanent storage, connect a database (Supabase / Google Sheets) later.
- Strong tab-lock security works best in the earlier standalone HTML quizzes; this app prioritises the single class Excel sheet.

