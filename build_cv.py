from docx import Document
from docx.shared import Inches, Pt, RGBColor, Mm
from docx.enum.text import WD_TAB_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

out = r'C:\Users\david\Documents\GitHub\ResumeSkills\David_Dionisio_Goncalves_CV.docx'
doc = Document()
sec=doc.sections[0]; sec.page_width=Mm(210); sec.page_height=Mm(297)
sec.top_margin=Inches(.58); sec.bottom_margin=Inches(.58); sec.left_margin=Inches(.72); sec.right_margin=Inches(.72)
NAVY=RGBColor(26,48,70); BLUE=RGBColor(53,93,125); GRAY=RGBColor(91,103,114); INK=RGBColor(35,40,45)
normal=doc.styles['Normal']; normal.font.name='Arial'; normal.font.size=Pt(10.2); normal.font.color.rgb=INK
normal.paragraph_format.line_spacing=1.05; normal.paragraph_format.space_after=Pt(3)

def fmt(run,size=10.2,bold=False,color=INK):
    run.font.name='Arial'; run._element.get_or_add_rPr().rFonts.set(qn('w:ascii'),'Arial'); run._element.get_or_add_rPr().rFonts.set(qn('w:hAnsi'),'Arial')
    run.font.size=Pt(size); run.bold=bold; run.font.color.rgb=color

def rule(p,color='355D7D'):
    pPr=p._p.get_or_add_pPr(); borders=OxmlElement('w:pBdr'); bottom=OxmlElement('w:bottom')
    bottom.set(qn('w:val'),'single'); bottom.set(qn('w:sz'),'6'); bottom.set(qn('w:space'),'3'); bottom.set(qn('w:color'),color); borders.append(bottom); pPr.append(borders)

def section(text):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(10); p.paragraph_format.space_after=Pt(4); p.paragraph_format.keep_with_next=True
    fmt(p.add_run(text.upper()),12.5,True,NAVY); rule(p)

def role(title,org,dates):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(3); p.paragraph_format.space_after=Pt(0); p.paragraph_format.keep_with_next=True
    p.paragraph_format.tab_stops.add_tab_stop(Inches(6.83),WD_TAB_ALIGNMENT.RIGHT)
    fmt(p.add_run(title),10.8,True,INK); fmt(p.add_run('\t'+dates),9.6,True,BLUE)
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); p.paragraph_format.keep_with_next=True
    fmt(p.add_run(org),9.8,False,GRAY)

def bullet(text):
    p=doc.add_paragraph(); p.paragraph_format.left_indent=Inches(.2); p.paragraph_format.first_line_indent=Inches(-.14); p.paragraph_format.space_after=Pt(2); p.paragraph_format.keep_together=True
    fmt(p.add_run('•  '+text),10.1)

# Contact block is plain document body text for reliable ATS parsing.
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(1); fmt(p.add_run('DAVID DIONÍSIO GONÇALVES'),20,True,NAVY)
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(4); fmt(p.add_run('INFORMATION MANAGEMENT STUDENT  |  DATA ANALYSIS'),9.4,True,BLUE)
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); fmt(p.add_run('Lisbon, Portugal  |  +351 934 942 942  |  daviddioniogoncalves@gmail.com'),9.5,False,INK)
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); fmt(p.add_run('linkedin.com/in/david-dionísio-gonçalves/'),9.5,False,GRAY)

section('Professional Summary')
p=doc.add_paragraph('I’m an Information Management student at NOVA IMS, building on a Science and Technology background. I’m curious about how organisations use information and technology to solve practical problems, and I enjoy learning by doing, exploring new tools, and working with others. Through my studies, club activities, research projects, and hackathons, I’ve developed an adaptable and analytical approach. I’m now seeking a summer internship where I can contribute, keep learning, and build practical experience in information management.')
p.paragraph_format.space_after=Pt(1)
section('Education')
role('Bachelor’s in Information Management','NOVA IMS','2025–Present')
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); fmt(p.add_run('Current grade: 16.00/20'),9.8,False,GRAY)
role('Science and Technology','Colégio Sagrado Coração de Maria','2017–2025')
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); fmt(p.add_run('Final grade: 17.78/20'),9.8,False,GRAY)
section('Experience and Projects')
role('Research and Development Associate','NOVAe R&D, NOVA Entrepreneurship Club','2026–Present')
bullet('Assist the club in developing and updating its official website.')
bullet('Set up cloud-based databases to organise information for the student entrepreneurship club.')

role('Student Research Project: The Impact of Diet on Chicken Growth and Health','Colégio Sagrado Coração de Maria','2024–2025')
bullet('Observed and documented chicken behaviour and growth over a two-month period; collected, processed, and analysed data.')
bullet('Formulated a hypothesis and evaluated results to develop evidence-based conclusions.')
role('Hackathons','Various health and AI hackathons','2025–Present')
bullet('Took part in collaborative hackathon projects focused on health and AI.')
role('Participant and Speaker','Sagrado MUN','2022–2025')
bullet('Represented a country in simulated United Nations councils, discussing international affairs and developing public speaking, critical thinking, and collaboration skills.')
section('Skills and Tools')
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); fmt(p.add_run('Data and databases: '),10.1,True,NAVY); fmt(p.add_run('SQL, SAS, R, cloud-based databases, data collection and analysis'),10.1)
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); fmt(p.add_run('Programming and web: '),10.1,True,NAVY); fmt(p.add_run('Python, VS Code, GitHub, website development'),10.1)
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); fmt(p.add_run('Microsoft 365: '),10.1,True,NAVY); fmt(p.add_run('Excel, Word, PowerPoint, Outlook'),10.1)
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); fmt(p.add_run('AI-assisted development: '),10.1,True,NAVY); fmt(p.add_run('Codex, Cursor, Antigravity, Claude Code'),10.1)
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); fmt(p.add_run('Creative tools: '),10.1,True,NAVY); fmt(p.add_run('Canva, CapCut, Blender, video and image editing'),10.1)
section('Languages')
p=doc.add_paragraph('Portuguese (C2)  |  English (C1)  |  French (A2)'); p.paragraph_format.space_after=Pt(0)
for para in doc.paragraphs: para.paragraph_format.widow_control=True
doc.core_properties.title='David Dionísio Gonçalves CV'; doc.core_properties.subject='Entry-level internship CV'; doc.core_properties.author='David Dionísio Gonçalves'
doc.save(out); print(out)



















