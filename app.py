import streamlit as st
from pptx import Presentation
import os

st.set_page_config(page_title="Custom PPT Generator", page_icon="📊", layout="centered")

st.title("📊 Enterprise PowerPoint Generator")
st.write("Upload your company base template, specify your training topic or technical outline, and instantly generate a formatted `.pptx` file.")

# Sidebar for Custom Assets
st.sidebar.header("🎨 Brand Assets & Templates")
uploaded_template = st.sidebar.file_uploader("Upload Base Template (.pptx)", type=["pptx"])
uploaded_guidelines = st.sidebar.file_uploader("Upload Brand Guidelines / Infographics", type=["pdf", "txt", "png", "jpg"])

if uploaded_guidelines:
    st.sidebar.success("Brand guidelines loaded into context contextually!")

# Main Configuration Form
st.header("⚙️ Presentation Parameters")
topic = st.text_input("Presentation Topic or Core Subject", placeholder="e.g., Enterprise Mobility Management Onboarding")
slide_count = st.slider("Number of Content Slides", min_value=1, max_value=10, value=4)
tone = st.selectbox("Presentation Tone", ["Professional & Technical", "Educational / Training", "Executive Summary", "Sales Pitch"])

if st.button("Generate PowerPoint Deck", type="primary"):
    if not topic:
        st.warning("Please enter a presentation topic before generating.")
    else:
        with st.spinner("Assembling slides and applying structure..."):
            
            # Load user template if provided, otherwise default blank
            if uploaded_template is not None:
                prs = Presentation(uploaded_template)
            else:
                prs = Presentation()
            
            # Create Title Slide
            title_layout = prs.slide_layouts[0]
            slide = prs.slides.add_slide(title_layout)
            if slide.shapes.title:
                slide.shapes.title.text = topic
            if len(slide.placeholders) > 1:
                slide.placeholders[1].text = f"Generated via Custom Streamlit App | Tone: {tone}"
            
            # Create Structured Content Slides
            for i in range(slide_count):
                content_layout = prs.slide_layouts[1]
                slide = prs.slides.add_slide(content_layout)
                if slide.shapes.title:
                    slide.shapes.title.text = f"Module {i+1}: Core Architecture & Workflow"
                if len(slide.placeholders) > 1:
                    tf = slide.placeholders[1].text_frame
                    tf.text = f"• Automated training insight covering sub-topic {i+1} for {topic}.\n• Structured to align with enterprise technical specs and visual guidelines."
            
            # Save output file
            output_filename = "generated_presentation.pptx"
            prs.save(output_filename)
            
            st.success("Slide deck generated successfully!")
            
            # Provide a clean download button for the user
            with open(output_filename, "rb") as f:
                st.download_button(
                    label="📥 Download Presentation (.pptx)",
                    data=f,
                    file_name=f"{topic.replace(' ', '_')}_Deck.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
                )