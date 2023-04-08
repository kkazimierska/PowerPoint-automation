import collections 
import collections.abc
from pptx import Presentation
from pptx.util import Inches

# Create a new presentation
presentation = Presentation()

# Title slide
title_slide_layout = presentation.slide_layouts[0]
slide = presentation.slides.add_slide(title_slide_layout)
title = slide.shapes.title
subtitle = slide.placeholders[1]
title.text = "Upgrade Process for Odoo Database"
subtitle.text = "Using the Python-pptx Library"

# Slide 1: Introduction
slide1_layout = presentation.slide_layouts[1]
slide1 = presentation.slides.add_slide(slide1_layout)
slide1.shapes.title.text = "Introduction"
slide1.placeholders[1].text = (
    "- What is Odoo?\n"
    "- Why upgrade the Odoo database?\n"
    "- What is Python-pptx?"
)

# Slide 2: Prerequisites
slide2_layout = presentation.slide_layouts[1]
slide2 = presentation.slides.add_slide(slide2_layout)
slide2.shapes.title.text = "Prerequisites"
slide2.placeholders[1].text = (
    "- Python and Odoo installation\n"
    "- Python-pptx library\n"
    "- Backup your Odoo database\n"
    "- Latest Odoo version"
)

# Slide 3: Python-pptx Overview
slide3_layout = presentation.slide_layouts[1]
slide3 = presentation.slides.add_slide(slide3_layout)
slide3.shapes.title.text = "Python-pptx Overview"
slide3.placeholders[1].text = (
    "- Installing the library\n"
    "- Basic syntax and usage\n"
    "- Creating a PowerPoint presentation with Python"
)

# Slide 4: Odoo Database Upgrade Steps
slide4_layout = presentation.slide_layouts[1]
slide4 = presentation.slides.add_slide(slide4_layout)
slide4.shapes.title.text = "Odoo Database Upgrade Steps"
slide4.placeholders[1].text = (
    "- Exporting the Odoo database\n"
    "- Uploading the database to the Odoo Upgrade Platform\n"
    "- Testing the upgraded database\n"
    "- Migrating the database to the production environment"
)

# Save the presentation
presentation.save("Odoo_Database_Upgrade_Process.pptx")
