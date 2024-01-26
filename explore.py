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