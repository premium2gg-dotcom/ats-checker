from flask import Flask, request, render_template_string
import re
from pypdf import PdfReader
from docx import Document

app = Flask(__name__)

STOP_WORDS = {
    "في", "من", "على", "إلى", "عن", "مع", "هذا", "هذه", "تم", "عبر", "أن", "أو",
    "و", "أجلي", "كذلك", "التي", "الذي", "الذين", "عنها", "منها", "إليها", "لذلك",
    "and", "the", "to", "of", "in", "for", "with", "a", "an", "on", "at", "by", "is", "are"
}

HTML_TEMPLATE = """
<!DOCTYPE html>
<html lang="ar" dir="rtl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>فاحص ATS - التقرير الشامل</title>
    <style>
        * { box-sizing: border-box; margin: 0; padding: 0; font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif; }
        body { background: #f0f2f5; color: #1c1e21; padding: 20px; display: flex; justify-content: center; align-items: flex-start; min-height: 100vh; }
        .card { background: #ffffff; border-radius: 12px; box-shadow: 0 8px 24px rgba(0,0,0,0.08); width: 100%; max-width: 700px; padding: 24px; margin-bottom: 30px; }
        h1 { font-size: 1.6rem; color: #0d6efd; text-align: center; margin-bottom: 8px; }
        p.subtitle { text-align: center; color: #6c757d; font-size: 0.9rem; margin-bottom: 20px; }
        label { font-weight: bold; font-size: 0.9rem; color: #495057; display: block; margin-bottom: 6px; margin-top: 10px; }
        textarea { width: 100%; height: 90px; border: 2px solid #e0e0e0; border-radius: 8px; padding: 10px; font-size: 0.9rem; outline: none; transition: 0.3s; resize: vertical; }
        textarea:focus { border-color: #0d6efd; }
        input[type="file"] { width: 100%; padding: 10px; border: 2px dashed #0d6efd; border-radius: 8px; background: #f8f9fa; cursor: pointer; }
        .or-divider { text-align: center; margin: 10px 0; color: #888; font-size: 0.85rem; font-weight: bold; }
        button.btn-submit { width: 100%; background: #0d6efd; color: white; border: none; padding: 14px; border-radius: 8px; font-size: 1rem; font-weight: bold; cursor: pointer; margin-top: 15px; transition: 0.2s; }
        button.btn-submit:hover { background: #0b5ed7; }
        
        .report-card { margin-top: 24px; padding: 20px; border-radius: 10px; background: #ffffff; border: 1px solid #dee2e6; }
        .progress-bar-bg { background: #e9ecef; border-radius: 20px; height: 22px; width: 100%; margin: 12px 0 20px 0; overflow: hidden; }
        .progress-bar-fill { height: 100%; text-align: center; color: white; font-size: 0.8rem; font-weight: bold; line-height: 22px; transition: width 0.5s ease-in-out; }
        
        .tag { display: inline-block; background: #e7f1ff; color: #0d6efd; padding: 5px 10px; border-radius: 6px; font-size: 0.85rem; margin: 3px; font-weight: 500; }
        .tag-missing { background: #f8d7da; color: #842029; }
        
        .advice-box { background: #fff3cd; border-right: 4px solid #ffc107; padding: 12px 16px; border-radius: 6px; margin-top: 15px; font-size: 0.9rem; color: #664d03; line-height: 1.5; }
        .print-btn { background: #198754; color: white; border: none; padding: 10px 18px; border-radius: 6px; font-weight: bold; cursor: pointer; margin-top: 15px; float: left; }
        .print-btn:hover { background: #146c43; }
        
        @media print {
            form, .or-divider, button.btn-submit, .print-btn { display: none !important; }
            body { background: white; padding: 0; }
            .card { box-shadow: none; border: none; }
        }
    </style>
</head>
<body>
    <div class="card">
        <h1>التقرير الشامل لفحص ATS 📊</h1>
        <p class="subtitle">حلل ملائمتك للوظيفة واحصل على التقرير والتوصيات فوراً</p>
        
        <form method="post" enctype="multipart/form-data">
            <label>1. الوصف الوظيفي (Job Description):</label>
            <textarea name="jobDescription" placeholder="انسخ الوصف الوظيفي هنا..." required></textarea>
            
            <label>2. رفع ملف السيرة الذاتية (PDF أو Word):</label>
            <input type="file" name="resumeFile" accept=".pdf,.docx">
            
            <div class="or-divider">─── أَوْ ───</div>
            
            <label>لصق نص السيرة الذاتية يدويًا:</label>
            <textarea name="resumeText" placeholder="انسخ نص السيرة الذاتية هنا في حال عدم رفع ملف..."></textarea>
            
            <button type="submit" class="btn-submit">توليد التقرير الشامل 🚀</button>
        </form>

        {% if error %}
        <div style="background: #f8d7da; color: #842029; padding: 10px; border-radius: 6px; margin-top: 15px;">{{ error }}</div>
        {% endif %}

        {% if result %}
        <div class="report-card">
            <button onclick="window.print()" class="print-btn">🖨️ حفظ التقرير PDF</button>
            <h2 style="font-size: 1.2rem; color: #212529; margin-bottom: 8px;">نتيجة التقييم:</h2>
            
            <div class="progress-bar-bg">
                <div class="progress-bar-fill" style="width: {{ result.score }}%; background-color: {% if result.score >= 70 %}#198754{% elif result.score >= 40 %}#ffc107{% else %}#dc3545{% endif %};">
                    {{ result.score }}%
                </div>
            </div>

            <p style="margin-bottom: 8px;"><strong>الكلمات المفتاحية المطابقة ({{ result.matched|length }}):</strong></p>
            <div style="margin-bottom: 15px;">
                {% for word in result.matched %}
                    <span class="tag">{{ word }}</span>
                {% else %}
                    <span style="color: #6c757d; font-size: 0.85rem;">لا توجد كلمات مطابقة.</span>
                {% endfor %}
            </div>

            <p style="margin-bottom: 8px;"><strong>الكلمات المفتاحية المفقودة ({{ result.missing|length }}):</strong></p>
            <div>
                {% for word in result.missing %}
                    <span class="tag tag-missing">{{ word }}</span>
                {% else %}
                    <span style="color: #198754; font-size: 0.85rem;">ممتاز! سيرتك تغطي كافة المهارات.</span>
                {% endfor %}
            </div>

            <div class="advice-box">
                <strong>💡 التوجيهات والتوصيات:</strong><br>
                {{ result.advice }}
            </div>
        </div>
        {% endif %}
    </div>
</body>
</html>
"""

def extract_text_from_pdf(file_stream):
    reader = PdfReader(file_stream)
    text = ""
    for page in reader.pages:
        extracted = page.extract_text()
        if extracted:
            text += extracted + " "
    return text

def extract_text_from_docx(file_stream):
    doc = Document(file_stream)
    text = " ".join([paragraph.text for paragraph in doc.paragraphs])
    return text

def extract_keywords(text):
    words = re.findall(r'\b\w+\b', text.lower())
    return set(w for w in words if len(w) > 2 and w not in STOP_WORDS)

def analyze_and_report(job_text, resume_text):
    job_keywords = extract_keywords(job_text)
    resume_keywords = extract_keywords(resume_text)
    
    if not job_keywords:
        return {"score": 0, "matched": [], "missing": [], "advice": "الوصف الوظيفي قصير جداً ولا يحوي كلمات مفتاحية كافية."}
        
    matched = list(job_keywords.intersection(resume_keywords))
    missing = list(job_keywords.difference(resume_keywords))
    score = int((len(matched) / len(job_keywords)) * 100)
    
    if score >= 80:
        advice = "أداء ممتاز جداً! سيرتك الذاتية مهيأة تماماً لتجاوز أنظمة ATS لهذه الوظيفة."
    elif score >= 50:
        advice = f"تطابق جيد، ولكن يُنصح بدمج الكلمات المفقودة مثل ({', '.join(missing[:4])}) ضمن قسم المهارات أو الخبرات لتجاوز نسبة 80%."
    else:
        advice = f"نسبة القبول منخفضة. قم بتعديل سيرتك الذاتية لتشمل الكلمات المفقودة الموضحة أعلاه، مع التركيز على صياغة الخبرات بما يتوافق مع متطلبات الوصف الوظيفي."

    return {
        "score": score,
        "matched": matched[:12],
        "missing": missing[:12],
        "advice": advice
    }

@app.route("/", methods=["GET", "POST"])
def home():
    result = None
    error = None
    if request.method == "POST":
        job_text = request.form.get("jobDescription", "")
        resume_text = request.form.get("resumeText", "")
        uploaded_file = request.files.get("resumeFile")

        if uploaded_file and uploaded_file.filename != '':
            filename = uploaded_file.filename.lower()
            try:
                if filename.endswith('.pdf'):
                    resume_text = extract_text_from_pdf(uploaded_file)
                elif filename.endswith('.docx'):
                    resume_text = extract_text_from_docx(uploaded_file)
                else:
                    error = "نوع الملف غير مدعوم! يرجى رفع PDF أو Word (.docx) فقط."
            except Exception as e:
                error = f"حدث خطأ أثناء قراءة الملف: {str(e)}"

        if not error:
            if job_text and resume_text.strip():
                result = analyze_and_report(job_text, resume_text)
            else:
                error = "يرجى رفع ملف السيرة الذاتية أو كتابة نصها، مع إدخال الوصف الوظيفي."

    return render_template_string(HTML_TEMPLATE, result=result, error=error)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8080)
  
