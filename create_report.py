from pathlib import Path
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.style import WD_STYLE_TYPE

OUT = Path('/mnt/data/CreditWise_AI/docs/CreditWise_AI_End_Term_Project_Report.docx')

def set_cell_shading(cell, fill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:fill'), fill)
    tcPr.append(shd)

def set_cell_border(cell, **kwargs):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    tcBorders = tcPr.first_child_found_in('w:tcBorders')
    if tcBorders is None:
        tcBorders = OxmlElement('w:tcBorders')
        tcPr.append(tcBorders)
    for edge in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
        if edge in kwargs:
            tag = 'w:{}'.format(edge)
            element = tcBorders.find(qn(tag))
            if element is None:
                element = OxmlElement(tag)
                tcBorders.append(element)
            for key in ['val', 'sz', 'space', 'color']:
                if key in kwargs[edge]:
                    element.set(qn('w:' + key), str(kwargs[edge][key]))

def style_table(table, header_fill='1D4ED8'):
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.style = 'Table Grid'
    for i, row in enumerate(table.rows):
        for cell in row.cells:
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            for p in cell.paragraphs:
                for run in p.runs:
                    run.font.size = Pt(9)
            set_cell_border(cell, top={'val':'single','sz':4,'color':'CBD5E1'}, bottom={'val':'single','sz':4,'color':'CBD5E1'}, left={'val':'single','sz':4,'color':'CBD5E1'}, right={'val':'single','sz':4,'color':'CBD5E1'})
        if i == 0:
            for cell in row.cells:
                set_cell_shading(cell, header_fill)
                for p in cell.paragraphs:
                    for r in p.runs:
                        r.bold = True
                        r.font.color.rgb = RGBColor(255,255,255)


def add_bullet(doc, text, level=0):
    p = doc.add_paragraph(style='List Bullet' if level == 0 else 'List Bullet 2')
    p.add_run(text)
    return p

def add_numbered(doc, text):
    p = doc.add_paragraph(style='List Number')
    p.add_run(text)
    return p

def add_heading(doc, text, level=1):
    p = doc.add_heading(text, level=level)
    return p


doc = Document()
sec = doc.sections[0]
sec.top_margin = Inches(0.65)
sec.bottom_margin = Inches(0.65)
sec.left_margin = Inches(0.7)
sec.right_margin = Inches(0.7)

styles = doc.styles
styles['Normal'].font.name = 'Aptos'
styles['Normal'].font.size = Pt(10.5)
styles['Title'].font.name = 'Aptos Display'
styles['Title'].font.size = Pt(28)
styles['Title'].font.bold = True
styles['Heading 1'].font.name = 'Aptos Display'
styles['Heading 1'].font.size = Pt(19)
styles['Heading 1'].font.color.rgb = RGBColor(15, 23, 42)
styles['Heading 2'].font.name = 'Aptos Display'
styles['Heading 2'].font.size = Pt(14)
styles['Heading 2'].font.color.rgb = RGBColor(29, 78, 216)

# Cover
t = doc.add_paragraph()
t.alignment = WD_ALIGN_PARAGRAPH.CENTER
t.space_after = Pt(6)
r = t.add_run('CREDITWISE AI')
r.bold = True
r.font.size = Pt(30)
r.font.color.rgb = RGBColor(29, 78, 216)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('Explainable AI Credit-Risk Pre-Screener')
r.bold = True
r.font.size = Pt(18)
r.font.color.rgb = RGBColor(15, 23, 42)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('End-Term Project Report | App / Chatbot Application')
r.font.size = Pt(12)
r.font.color.rgb = RGBColor(71, 85, 105)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run('\nSelected use case from the supplied menu: Loan eligibility / credit-risk pre-screener').italic = True

# Executive summary box
box = doc.add_table(rows=1, cols=1)
box.alignment = WD_TABLE_ALIGNMENT.CENTER
cell = box.cell(0,0)
set_cell_shading(cell, 'EFF6FF')
cell.text = ''
pp = cell.paragraphs[0]
run = pp.add_run(
    'Project concept: CreditWise AI is a hybrid decision-support application. A transparent logistic-regression model estimates an applicant\'s approval probability, a separate rule layer catches obvious risk signals, and Gemini optionally converts the result into a plain-English explanation. Gemini cannot override the model or make a final lending decision.'
)
run.font.size = Pt(11)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
p.add_run('\nPrepared for academic evaluation').font.size = Pt(10)

doc.add_page_break()

# 1
add_heading(doc, '1. Why this use case was selected', 1)
doc.add_paragraph(
    'The supplied use-case menu contains 20 possible application ideas. CreditWise AI selects the loan eligibility / credit-risk pre-screener because it fits an MBA Finance + Big Data Analytics profile especially well. The problem is clearly business-facing, structured applicant data is easy to demonstrate with synthetic records, machine learning is meaningful, explainability can be shown visually, and a generative AI layer can be added without allowing hallucinations to control the core decision.'
)
for x in [
    'Strong finance relevance: credit screening, affordability, risk segmentation and lending workflow.',
    'Strong analytics relevance: structured features, model training, holdout validation, ROC-AUC and what-if analysis.',
    'Strong AI relevance: Gemini is used as an explanation / copilot layer rather than a black-box decision maker.',
    'Strong demonstration value: the examiner can change inputs live and immediately see the probability change.',
    'Strong scalability story: the same architecture could later connect to verified bureau, income, KYC and transaction data in a real organization.'
]: add_bullet(doc, x)

add_heading(doc, 'Project objective', 2)
doc.add_paragraph('Build a shareable, working application that screens a synthetic loan applicant, explains the major risk drivers, provides a manual-review recommendation, and demonstrates portfolio analytics plus what-if simulation.')

# 2
add_heading(doc, '2. Problem statement and users', 1)
doc.add_paragraph('Manual loan pre-screening can be slow and difficult to explain consistently. A lender or fintech needs a quick first-pass tool that can rank applications into an operational queue while keeping a human reviewer in control.')
t = doc.add_table(rows=1, cols=3)
headers=['Stakeholder','Need','How CreditWise helps']
for i,h in enumerate(headers): t.cell(0,i).text=h
rows=[
    ('End user - credit analyst','Fast first-pass risk signal','Gets probability, risk bucket, flags and explanation in one screen.'),
    ('Business / lender','Lower manual effort on low-risk triage','Standardized scoring plus rule checks create a consistent workflow.'),
    ('Risk / compliance reviewer','Auditability and explainability','Rules are transparent and the model is separated from the generative explanation layer.'),
    ('Student / project evaluator','Visible AI + analytics functionality','Live scoring, charts, what-if analysis and optional Gemini Copilot are easy to demo.')
]
for row in rows:
    cells=t.add_row().cells
    for i,v in enumerate(row): cells[i].text=v
style_table(t)

# 3
add_heading(doc, '3. Business model and value proposition', 1)
doc.add_paragraph('The prototype is a B2B decision-support component for a lender, fintech, NBFC or embedded-finance provider. The immediate customer would be the lending organization; the analyst or underwriting team is the primary operational user.')
add_bullet(doc, 'Core value: faster, more consistent pre-screening with a human-readable reason trail.')
add_bullet(doc, 'Adoption path: start as an internal analyst tool, then connect it to an application-origination workflow through an API.')
add_bullet(doc, 'Monetization in a real venture: B2B SaaS subscription / usage-based pricing for decisioning volume, plus implementation and analytics services.')

# 4
add_heading(doc, '4. End-to-end application architecture', 1)
arch = doc.add_table(rows=1, cols=5)
for i,h in enumerate(['Layer','Technology','Role','Why used','Failure behavior']): arch.cell(0,i).text=h
arch_rows=[
    ('UI','Streamlit','Forms, tabs, charts, chat','Fast Python UI suitable for an academic prototype','App remains usable if Gemini is offline'),
    ('Data','Synthetic CSV','Training and demo portfolio','Reproducible and privacy-safe','No real customer data required'),
    ('ML','scikit-learn Logistic Regression','Approval probability','Interpretable and lightweight','Rule layer still operates'),
    ('Rules','Python business checks','DTI / credit / payment flags','Transparent guardrails','Flags shown even if model output is ambiguous'),
    ('GenAI','Google Gemini API','Explanation + Copilot','Natural-language interaction','Local deterministic fallback used if API fails')
]
for row in arch_rows:
    cells=arch.add_row().cells
    for i,v in enumerate(row): cells[i].text=v
style_table(arch)

doc.add_paragraph('Design principle: use AI where language generation helps, but do not let a generative model overwrite the numeric screening result. This makes the app easier to defend during a viva because the decision path is explicit.')

# 5
add_heading(doc, '5. Data and model methodology', 1)
doc.add_paragraph('The project includes 2,400 synthetic loan applications generated with a fixed random seed so the demonstration is reproducible. No real customer information is used.')

add_heading(doc, 'Input features', 2)
features = ['Age','Annual income','Employment history','Requested loan','Loan term','Credit score','Debt-to-income ratio','Existing loans','Missed payments in last 12 months','Savings balance','Employment type','Housing status','Dependents']
for f in features: add_bullet(doc, f)

add_heading(doc, 'Modeling approach', 2)
for x in [
    'Split data into 75% training and 25% holdout test data using stratification and random_state=42.',
    'Impute numeric values using the median and categorical values using the most frequent category.',
    'Standardize numeric features and one-hot encode categorical features.',
    'Fit a logistic regression classifier with max_iter=1500.',
    'Report probability of the positive class as an estimated probability of approval.',
    'Convert the probability into three demonstration buckets: Low Risk / Pre-Qualify at >=70%, Medium Risk / Refer at 50%-69.9%, and High Risk / Refer below 50%.',
    'Run a separate rule layer to flag very high DTI, low credit score, multiple missed payments, large loan-to-income or a low savings buffer.'
]: add_numbered(doc, x)

metrics = doc.add_table(rows=1, cols=3)
for i,h in enumerate(['Metric','Holdout result','Interpretation']): metrics.cell(0,i).text=h
for row in [
    ('Accuracy','72.2%','Share of holdout applications classified correctly at a 0.50 probability threshold.'),
    ('ROC-AUC','0.781','Shows useful ranking ability on the synthetic holdout set.'),
    ('Synthetic approval rate','67.2%','Observed positive-class rate in the generated portfolio.'),
    ('Holdout size','600','25% of the 2,400-row synthetic dataset.')
]:
    cells=metrics.add_row().cells
    for i,v in enumerate(row): cells[i].text=v
style_table(metrics)

doc.add_paragraph('Important: these metrics do not demonstrate real-world credit performance. The dataset is synthetic and intentionally created for a classroom prototype.')

# 6
add_heading(doc, '6. Application screens and functionality', 1)
for heading, body in [
    ('Screen Applicant','Form inputs, quick demo profiles, model probability, risk classification, rule flags, recommended workflow action, explanation and local feature contribution chart.'),
    ('What-If Simulator','Lets the user change credit score, DTI, income, loan amount, missed payments and savings, then compares the new probability with the baseline.'),
    ('Portfolio Analytics','Shows sample size, observed approval rate, average income, average DTI, credit-score distribution, DTI vs loan amount, confusion matrix, accuracy and ROC-AUC.'),
    ('AI Copilot','Scoped conversational assistant for explaining the model, data fields, what-if simulator and API behavior. It refuses or redirects questions outside the prototype scope.'),
    ('About','Explains the use case, architecture, technology stack, limitations, sources and the viva-ready summary.')
]:
    add_heading(doc, heading, 2)
    doc.add_paragraph(body)

# 7 demo profiles
add_heading(doc, '7. Demo scenarios for the video recording', 1)
demo = doc.add_table(rows=1, cols=5)
for i,h in enumerate(['Scenario','Credit score','DTI','Missed payments','Expected probability']): demo.cell(0,i).text=h
for row in [
    ('Strong profile', '790','22%','0','87.7%'),
    ('Borderline profile','670','34%','0','55.3%'),
    ('High-risk profile','545','58%','4','1.1%')
]:
    cells=demo.add_row().cells
    for i,v in enumerate(row): cells[i].text=v
style_table(demo)

doc.add_paragraph('Recommended video flow: start with the Strong profile, run the pre-screen, move to the What-If Simulator, change the credit score and DTI, then open Portfolio Analytics and finish with the AI Copilot.')

# 8 SWOT
add_heading(doc, '8. SWOT analysis', 1)
swot = doc.add_table(rows=1, cols=2)
for i,h in enumerate(['Area','Assessment']): swot.cell(0,i).text=h
swot_rows=[
    ('Strengths','Transparent model + rules; explainable output; live what-if simulation; Gemini does not control the final screen; simple deployment.'),
    ('Weaknesses','Synthetic data; simplified lending logic; no real credit-bureau/KYC integration; model fairness and calibration are not production validated.'),
    ('Opportunities','Connect to verified bureau/income data; add document extraction; add drift monitoring; extend to credit cards, SME lending and BNPL; support multilingual analyst workflows.'),
    ('Threats','Regulatory requirements, biased models, API/model changes, data privacy incidents, enterprise competitors, and over-reliance on automated outputs.')
]
for row in swot_rows:
    cells=swot.add_row().cells
    for i,v in enumerate(row): cells[i].text=v
style_table(swot)

# 9 competitor
add_heading(doc, '9. Competitor comparison', 1)
doc.add_paragraph('Two real commercial competitors in the broader credit decisioning category are TransUnion and Experian. Their enterprise platforms use credit data, analytics and decisioning workflows for lending decisions. CreditWise is intentionally much smaller and is a classroom prototype rather than an enterprise credit bureau or decision platform. (Official product references are listed in the sources section.)')
comp = doc.add_table(rows=1, cols=4)
for i,h in enumerate(['Dimension','CreditWise AI','TransUnion','Experian']): comp.cell(0,i).text=h
for row in [
    ('Primary role','Classroom credit pre-screening prototype','Commercial credit data / risk / decisioning ecosystem','Commercial credit decisioning / origination ecosystem'),
    ('Data','Synthetic demonstration data','Large proprietary and bureau-linked data ecosystem','Large proprietary and bureau-linked data ecosystem'),
    ('Explainability','Explicit rule flags + ML contribution chart + optional Gemini narrative','Enterprise analytics / decisioning capabilities','Enterprise decisioning and analytics capabilities'),
    ('Differentiator','Low-cost, transparent teaching prototype with live what-if simulation','Enterprise data depth and scale','Enterprise decisioning depth and integration')
]:
    cells=comp.add_row().cells
    for i,v in enumerate(row): cells[i].text=v
style_table(comp)

# 10 Technical eval A
add_heading(doc, '10. Evaluation questions - Section A: Business & Strategic Framing', 1)
A = [
('Q1. What problem does the app solve, and for whom? Who is the paying customer vs. end user?','It speeds up first-pass loan screening for lenders. The paying customer would be a lender, NBFC, fintech or similar institution; the main end user is a credit analyst or underwriter.'),
('Q2. Strengths: what does the app do better than a human doing the same task, or a competing tool?','It produces a standardized score and explanation quickly, lets the user test what-if scenarios instantly, and keeps the rule layer visible. It is not claiming to outperform enterprise systems.'),
('Q3. Weaknesses: where does your app break, hallucinate, or give unreliable output? Show one example.','The ML model can misclassify because it is trained on synthetic data. Gemini can also generate an imperfect explanation. The app therefore treats the model/rules as the scoring source and uses Gemini only as an explanation/copilot layer with a fallback.'),
('Q4. Opportunities: how would this scale?','Connect verified data sources, add multilingual support, integrate document extraction, expose an API, add monitoring and retraining, and extend the same architecture to other lending products.'),
('Q5. Threats: what would kill the product?','Regulatory non-compliance, poor model fairness, privacy breaches, API/model changes, enterprise competitors and bad operational decisions caused by over-trust in automation.'),
('Q6. Who are 2 real competitors and how is yours different?','TransUnion and Experian are real commercial decisioning/credit-risk providers. CreditWise is differentiated by its small scope, transparency, synthetic data, student-friendly architecture and live what-if demonstration.'),
('Q7. What is the monetization/adoption path?','A real product could start as an internal SaaS screening layer and charge by application volume or seats, with additional implementation and analytics services.')]
for q,a in A:
    p=doc.add_paragraph(); r=p.add_run(q); r.bold=True; doc.add_paragraph(a)

# 11 Technical B
add_heading(doc, '11. Evaluation questions - Section B: AI / Technical Understanding', 1)
B = [
('Q1. Which model/API are you using and why?','The numeric model is logistic regression because it is fast, interpretable and easy to validate. Gemini is optional and is used for natural-language explanation/Copilot. The project uses the current Google GenAI Python SDK pattern documented by Google and defaults to Gemini 3.8 Flash, with environment-based model selection.'),
('Q2. Walk through prompt/system design.','The Gemini prompt states the role, scope, input data, model probability, rule flags and recommendation. It explicitly forbids final lending decisions, invented facts, use of sensitive attributes and instructions to manipulate an application.'),
('Q3. What happens outside intended scope?','The Copilot is scoped to explaining the app, score, data and what-if analysis. Out-of-scope questions are redirected. The numeric screening path does not depend on Gemini.'),
('Q4. How are you handling privacy?','The project uses synthetic data. If a real API key is configured, only the input summary needed for the explanation is sent to Gemini. The key is stored in deployment secrets or environment variables, never in source code.'),
('Q5. What happens if the API is down or returns garbage?','The app catches the Gemini error and falls back to a deterministic local explanation. The model score and rule flags remain available.'),
('Q6. If you used RAG/fine-tuning, what data did you use?','No RAG or fine-tuning is used. This keeps the system simple, reproducible and easy to explain.')]
for q,a in B:
    p=doc.add_paragraph(); r=p.add_run(q); r.bold=True; doc.add_paragraph(a)

# 12 Critical thinking C
add_heading(doc, '12. Evaluation questions - Section C: Critical Thinking', 1)
C = [
('Q1. Give one example where your bot could confidently give a wrong answer.','Gemini could describe a high-probability applicant as “safe” even though the formal lender policy might contain a rule not represented in this prototype. That is exactly why Gemini cannot override the numeric/rule layer.'),
('Q2. What would you NOT trust this app to do unsupervised?','I would not trust it to approve/decline real customers, set final loan terms, or make adverse-action decisions. Those functions need validated data, policy controls, compliance review and human oversight.'),
('Q3. Who is accountable for a bad decision?','The app is a decision-support prototype. In production, accountability would sit with the lending organization operating the decision process; model vendors and engineering teams also have responsibilities for the components they provide. This app is not presented as a substitute for governance.'),
('Q4. What is the biggest underlying AI limitation?','Statistical models only learn the relationships represented in training data. Distribution shift, missing variables, noisy labels and hidden policy rules can make a model wrong even when the UI looks confident.')]
for q,a in C:
    p=doc.add_paragraph(); r=p.add_run(q); r.bold=True; doc.add_paragraph(a)

# 13 Execution D
add_heading(doc, '13. Evaluation question - Section D: Execution', 1)
doc.add_paragraph('Edge case tested: an applicant with a strong income and credit score but an extremely high debt-to-income ratio. The change made was to keep a separate rule layer so that a single high-risk affordability signal is visible even if the overall model probability remains moderate. This demonstrates that model output and business guardrails are intentionally separate.')

# 14 E
add_heading(doc, '14. Evaluation questions - Section E: App-Specific Workflow', 1)
E = [
('Q1. Walk through input-to-output flow.','User enters applicant data -> inputs are validated by widget ranges/types -> data is transformed by the preprocessing pipeline -> logistic regression produces an approval probability -> rule layer produces flags -> workflow recommendation is generated -> Gemini optionally creates a narrative explanation -> the app shows charts and what-if analysis.'),
('Q2. How do you validate or sanity-check user input?','Streamlit numeric controls enforce ranges and types such as credit score 300-900, DTI 0-80%, non-negative income/savings and sensible loan-term selections. In a production version, additional cross-field validation and document verification would be required.'),
('Q3. If AI disagrees with a rule-of-thumb, how would a user catch that?','The rule flags are displayed separately from the Gemini narrative. Therefore a user can see that the AI wording differs from the deterministic rule layer and should rely on the formal policy, not the narrative.'),
('Q4. How is the output explained?','The user gets a probability, risk bucket, rule flags, workflow recommendation, natural-language explanation and approximate local feature contributions.'),
('Q5. What happens if the user refreshes?','Current session results are kept in Streamlit session state. A browser refresh can reset transient state depending on deployment behavior; a production version would persist applications to a database and assign application IDs.'),
('Q6. What changes at 100 users/day?','The app could remain lightweight at that volume, but production would need concurrency testing, persistent storage, API-rate management, caching, authentication, monitoring and a queue if model/API latency increases.'),
('Q7. Show similar inputs that produce different outputs.','The What-If Simulator is designed for this demonstration. For example, holding other fields roughly constant and moving the credit score or DTI can move the approval probability across the Low/Medium/High bands. The change is expected because those variables are meaningful model drivers.')]
for q,a in E:
    p=doc.add_paragraph(); r=p.add_run(q); r.bold=True; doc.add_paragraph(a)

# 15 demo narration
add_heading(doc, '15. Video recording script (approximately 3-4 minutes)', 1)
script = [
    ('0:00-0:25','Introduce CreditWise AI as an explainable credit-risk pre-screener built from the supplied use-case menu.'),
    ('0:25-1:20','Open Screen Applicant. Select Demo - Strong profile. Click Run AI pre-screen. Explain the probability, risk bucket, flags and recommendation.'),
    ('1:20-2:05','Move to What-If Simulator. Lower the credit score and raise DTI. Explain that the probability changes because the numeric model sees a different risk profile.'),
    ('2:05-2:40','Open Portfolio Analytics. Show the synthetic dataset size, approval rate, ROC-AUC, accuracy and confusion matrix.'),
    ('2:40-3:20','Open AI Copilot and ask why DTI matters. Mention that Gemini explains the app but cannot override the score.'),
    ('3:20-3:45','Finish with the About tab and state the key limitation: synthetic data means this is an academic prototype, not a real lending system.')]
tt=doc.add_table(rows=1, cols=2)
for i,h in enumerate(['Time','What to say / do']): tt.cell(0,i).text=h
for row in script:
    cells=tt.add_row().cells; cells[0].text=row[0]; cells[1].text=row[1]
style_table(tt)

# 16 deployment
add_heading(doc, '16. Deployment and submission checklist', 1)
for x in [
    'Upload the project folder to GitHub.',
    'Deploy app.py on Streamlit Community Cloud or another Python-compatible host.',
    'Add GEMINI_API_KEY to the platform secrets if the generative layer is required for the demo.',
    'Open the generated streamlit.app URL in an incognito browser and test the full flow before recording.',
    'Submit the live link, the project report, and the video recording as requested by the faculty.'
]: add_bullet(doc, x)

doc.add_paragraph('The provided package contains app.py, requirements.txt, sample data, Gemini secrets template, README, and run scripts. No secret key is included.')

# 17 references
add_heading(doc, '17. References', 1)
refs = [
    'Google AI for Developers - Gemini API getting started: https://ai.google.dev/gemini-api/docs/get-started',
    'Google AI for Developers - Generate Content API: https://ai.google.dev/gemini-api/docs/generate-content',
    'Google AI for Developers - Models: https://ai.google.dev/gemini-api/docs/models',
    'TransUnion - Credit Solutions: https://www.transunion.com/solution/credit-solutions',
    'TransUnion - TruVision Credit Risk: https://www.transunion.com/solution/truvision/credit-risk',
    'Experian - Credit decisioning software: https://www.experian.com/business/solutions/credit-decisioning/credit-decision-tools',
    'Experian - Loan origination software: https://www.experian.com/business/solutions/cloud-applications-platforms/loan-origination-software'
]
for r in refs: add_bullet(doc, r)

# final note
p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = p.add_run('End of report')
r.bold = True
r.font.color.rgb = RGBColor(29,78,216)

doc.save(OUT)
print(OUT)
