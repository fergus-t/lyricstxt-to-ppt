from pathlib import Path
from pptx import Presentation  
from pptx.util import Inches, Pt 
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_ANCHOR, MSO_AUTO_SIZE


INPUT_TXT_FILE = r"lyrics.txt"
input_txt_file_path = Path(INPUT_TXT_FILE)

# read txt line by line
lyrics: list = input_txt_file_path.read_text(encoding="utf8").split("\n")


# delete blank lines
lyrics = list(line for line in lyrics if line != "")

# create ppt file
root = Presentation()
root.slide_width = Inches(13.33)
root.slide_height = Inches(7.5)


first_slide_layout = root.slide_layouts[6]  

# write slides line by line
for line in lyrics:
    slide = root.slides.add_slide(first_slide_layout)

    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(0, 0, 0)

    # creating textFrames 
    left = top = Inches(1)
    width = root.slide_width - Inches(2)
    height = root.slide_height - Inches(2)
    txBox = slide.shapes.add_textbox(left, top, 
                                 width, height) 
    tf = txBox.text_frame 
    p = tf.add_paragraph()
    p.text = line
    p.font.size = Pt(55)
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    tf.vertical_anchor =  MSO_ANCHOR.MIDDLE
    tf.word_wrap = True

    # tf.fit_text()
    

# output the ppt file
root.save("output.pptx")