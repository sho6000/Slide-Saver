import os
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN, MSO_ANCHOR
from datetime import datetime
import streamlit as st

def generate_presentation(lyrics):
    # Font customization options
    st.subheader("Font Customization")
    font_name = st.selectbox("Font:", ["Arial", "Times New Roman", "Special Elite", "Helvetica", "Georgia", "Tahoma"], help="default is arial")
    font_size = st.slider("Font Size:", min_value=1, max_value=72, value=12)
    font_color = st.color_picker("Font Color:")

    #convert hex color string to RGB int
    font_color_rgb = tuple(int(font_color[i:i+2], 16) for i in (1, 3, 5))

    # Text formatting options
    st.subheader("Text Formatting")
    bold = st.checkbox("Bold")
    italics = st.checkbox("Italics")
    underline = st.checkbox("Underline")

    # Line spacing customization with predefined options
    st.subheader("Line Spacing")
    line_spacing_options = [1.0, 1.5, 2.0, 2.5, 3.0]
    line_spacing = st.selectbox("Select Line Spacing:", line_spacing_options)

    # Background color or custom background image
    st.subheader("Background")
    background_options = ["Black", "White", "Custom Image"]
    background_choice = st.radio("Select Background:", background_options)

    # Custom background image upload
    custom_background = None
    if background_choice == "Custom Image":
        custom_background = st.file_uploader("Upload Custom Background Image", type=["jpg", "jpeg", "png"], help="Custom Image can be 1024x768px OR 4:3")

    # Generate presentation button
    if st.button("Generate Presentation"):
        # Create a blank PowerPoint presentation
        presentation = Presentation()

        # Split input text into stanzas (assuming stanzas are separated by empty lines)
        stanzas = lyrics.split("\n\n")

        for stanza_text in stanzas:
            # Skip empty stanzas
            if not stanza_text.strip():
                continue

            # Add a blank slide
            slide = presentation.slides.add_slide(presentation.slide_layouts[6])

            # Set background color or custom background image
            if background_choice == "Black":
                background = slide.background
                fill = background.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(0, 0, 0)  # Black background
            elif background_choice == "White":
                background = slide.background
                fill = background.fill
                fill.solid()
                fill.fore_color.rgb = RGBColor(255, 255, 255)  # White background
            elif background_choice == "Custom Image" and custom_background is not None:
                # Use the uploaded custom background image
                slide.shapes.add_picture(custom_background, Inches(0), Inches(0), Inches(10), Inches(7.5))

            # Add text to the slide
            text_box = slide.shapes.add_textbox(Inches(1), Inches(1), Inches(8), Inches(5))
            text_frame = text_box.text_frame

            # Split the stanza into lines
            lines = stanza_text.split('\n')

            for line in lines:
                # Skip empty lines
                if not line.strip():
                    continue

                # Add a paragraph for each line
                p = text_frame.add_paragraph()
                p.text = line

                # Apply font customization to the entire stanza
                run = p.runs[0]
                run.font.name = font_name
                run.font.size = Pt(font_size)
                run.font.color.rgb = RGBColor(*font_color_rgb)
                run.font.embed_url = 'C:/Users/ASUS/Desktop/py_proje/Slide Saver/font/SpecialElite-Regular.ttf'

                # Apply line spacing
                text_frame.space_after = Pt(font_size * line_spacing - font_size)

                # Set vertical alignment to middle
                text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

                # Apply paragraph alignment (always centered)
                p.alignment = PP_ALIGN.CENTER

                # Apply text formatting
                if bold:
                    run.font.bold = True
                if italics:
                    run.font.italic = True
                if underline:
                    run.font.underline = True

        # Get the current date and time
        current_date = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")

        # Construct the file name with the current date and .pptx extension
        file_name = f"{current_date}.pptx"

        # Save the PowerPoint file with the current date as the file name
        presentation.save(file_name)

        st.success(f"Presentation generated successfully as {file_name}")

        # Ask the user to download the generated presentation with the specified MIME type
        with open(file_name, "rb") as file:
            st.download_button(
                "Download Presentation",
                file.read(),
                key=file_name,
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
            )

        # Delete the temporary file after it's downloaded
        os.remove(file_name)
