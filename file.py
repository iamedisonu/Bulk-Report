from pathlib import Path
from docxtpl import DocxTemplate
import re

base_dir = Path(__file__).parent
word_template_path = base_dir / "Isomo_Report_Card _Template.docx"
doc = DocxTemplate(word_template_path)

context = {
"Name": "Edison Uwamungu",
"Philosophy": "B+",
"Research": "A-",
"Climate": "A",
"African" : "A",
"Lead" : "A-",
"Worldviews": "A",
"Jesus": "A",
"GPA" :3.83,

}
doc.render(context)
doc.save(base_dir /"Edison Uwamungu.docx")