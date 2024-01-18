from pptx import Presentation
from datetime import datetime
import settings
from update_facebook import get_data
from pptx.enum.shapes import MSO_SHAPE_TYPE
import pptx

## Get data
team_competency = get_data(settings.INPUT_EXCEL)


## Open pptx and explore
prs = Presentation('Facebookv1.pptx')

slide5 = prs.slides[5]

## Explore text and types in all shapes
i = 0
for shape in slide5.shapes:
    print(i)
    print(shape.shape_type)
    if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        print(i)
    if not shape.has_text_frame:
        continue
    print(shape.text)
    i +=1
Slide5 = {
    0: "empty",
    1: "Education",
    2: "Professional Interest",
    3: "Delivery Org 1",
    4: "Position",
    5: "Experience",
    6: "_professional_competencies",
    7: "Name+Initials"
}

## Replace the picture from shape = 1 to "images/dev.jpg"
img_path = "images/dev.jpg"
img2 = pptx.parts.image.Image.from_file(img_path)

# what image you're actually changing...
img_shape = slide5.shapes[1]

# get part and rId from shape we need to change
slide_part, rId = img_shape.part, img_shape._element.blip_rId
image_part = slide_part.related_part(rId)

# overwrite old blob info with new blob info
image_part.blob = img2._blob

## Replace the text values with the ones from excel and bold them
slide5.shapes[6].text = str(team_competency['Education'][1])
slide5.shapes[7].text = str(team_competency['Professional Interest'][1])
slide5.shapes[8].text = str(team_competency['Delivery Org 1'][1])
slide5.shapes[9].text = str(team_competency['Position'][1])
slide5.shapes[9].text_frame.paragraphs[0].font.bold = True
slide5.shapes[10].text = str(team_competency['Experience'][1])
slide5.shapes[11].text = str(team_competency['_professional_competencies'][1])
slide5.shapes[12].text = str(f"{team_competency['Name'][1]} ({team_competency['Initials'][1]})")
slide5.shapes[12].text_frame.paragraphs[0].font.bold = True



ts = datetime.now()

prs.save(f"Draft_readed{ts.microsecond}.pptx")