import streamlit as st
from feature import generate_presentation

# Streamlit app title
st.title("Slide Saver | PowerPoint Presentation Generator")
st.set_page_config(page_title="Slide Saver | PowerPoint Presentation Generator", layout="wide")

# User input section
st.header("version 2.0")

# Text input for content with an expanded text area
lyrics = st.text_area("Enter your lyrics here:", height=250, help="leave gaps after each stanza for the presentaion to appear on each slide")

# Call the function to generate the presentation
generate_presentation(lyrics)

# st.markdown("You can find the source code on [GitHub](https://github.com/sho6000/Slide-Saver).")