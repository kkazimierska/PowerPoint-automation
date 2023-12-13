from pptx import Presentation
from datetime import datetime
import pandas as pd

team_competency = pd.read_csv('team_competencies_v1.xlsx', sep = ";")

# Documentacja: https://python-pptx.readthedocs.io/_/downloads/en/stable/pdf/

# Open pptx and explore
prs = Presentation('Facebookv1.pptx')

slide5 = prs.slides[5]

# # Explore text in all shapes
i = 0
for shape in slide5.shapes:
    print(i)
    if not shape.has_text_frame:
        continue
    print(shape.text)
    i +=1

slide5.shapes[4].text = str(team_competency['Name'][0])
slide5.shapes[5].text = str(team_competency['ART.'][0])
slide5.shapes[6].text = str(team_competency['Title'][0])
slide5.shapes[7].text = str(team_competency['Competencies'][0])


ts = datetime.now()

prs.save(f"Draft_readed{ts.microsecond}.pptx")