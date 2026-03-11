import os
import json
import requests
import streamlit as st
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.dml.color import RGBColor
import google.generativeai as genai
from dotenv import load_dotenv
from io import BytesIO

# Load environment variables
load_dotenv()

# Configure API keys
GEMINI_API_KEY = os.getenv("GEMINI_API_KEY")
UNSPLASH_ACCESS_KEY = os.getenv("UNSPLASH_ACCESS_KEY")

if GEMINI_API_KEY:
    genai.configure(api_key=GEMINI_API_KEY)

# Set Streamlit page configuration
st.set_page_config(
    page_title="AI Presentation Builder",
    page_icon="📊",
    layout="wide"
)

st.title("📊 AI Powered Presentation Builder")
st.markdown("Generate professional PowerPoint presentations using AI")

# Sidebar for configuration
with st.sidebar:
    st.header("Configuration")
    num_slides = st.slider("Number of slides", 3, 15, 5)
    presentation_style = st.selectbox(
        "Presentation style",
        ["Professional", "Educational", "Creative", "Minimal"]
    )

# Main content area
col1, col2 = st.columns([2, 1])

with col1:
    topic = st.text_input(
        "Enter presentation topic",
        placeholder="e.g., Machine Learning, Climate Change, Web Development"
    )

with col2:
    generate_btn = st.button("🚀 Generate Presentation", use_container_width=True)

# Progress tracking
progress_placeholder = st.empty()
status_placeholder = st.empty()

def fetch_image_from_unsplash(query: str) -> str:
    """Fetch a relevant image from Unsplash API."""
    try:
        url = f"https://api.unsplash.com/search/photos"
        params = {
            "query": query,
            "per_page": 1,
            "orientation": "landscape"
        }
        headers = {"Authorization": f"Client-ID {UNSPLASH_ACCESS_KEY}"}
        
        response = requests.get(url, params=params, headers=headers, timeout=10)
        if response.status_code == 200:
            data = response.json()
            if data["results"]:
                return data["results"][0]["urls"]["regular"]
    except Exception as e:
        st.warning(f"Could not fetch image: {e}")
    
    return None

def generate_presentation_content(topic: str, num_slides: int, style: str) -> dict:
    """Generate presentation content using Gemini API."""
    prompt = f"""Generate a comprehensive presentation outline for the topic: '{topic}'
    
    Style: {style}
    Number of slides: {num_slides}
    
    Return the response as a JSON object with this structure:
    {{
        "title": "Presentation Title",
        "slides": [
            {{
                "slide_number": 1,
                "title": "Slide Title",
                "content": "Main content/bullet points",
                "image_query": "relevant search term for images"
            }}
        ]
    }}
    
    Make sure to include:
    - An engaging title slide
    - {num_slides - 2} content slides with detailed information
    - A conclusion slide
    - Each slide should have an image_query for visual enhancement
    """
    
    try:
        model = genai.GenerativeModel("gemini-pro")
        response = model.generate_content(prompt)
        
        # Extract JSON from response
        response_text = response.text
        
        # Try to find JSON in the response
        start_idx = response_text.find('{')
        end_idx = response_text.rfind('}') + 1
        
        if start_idx != -1 and end_idx > start_idx:
            json_str = response_text[start_idx:end_idx]
            return json.loads(json_str)
        else:
            raise ValueError("No JSON found in response")
    
    except Exception as e:
        st.error(f"Error generating content: {e}")
        return None

def create_presentation(content: dict) -> bytes:
    """Create a PowerPoint presentation from the generated content."""
    prs = Presentation()
    prs.slide_width = Inches(10)
    prs.slide_height = Inches(7.5)
    
    # Title slide
    title_slide_layout = prs.slide_layouts[6]  # Blank layout
    slide = prs.slides.add_slide(title_slide_layout)
    
    # Add background color to title slide
    background = slide.background
    fill = background.fill
    fill.solid()
    fill.fore_color.rgb = RGBColor(41, 128, 185)
    
    # Add title
    left = Inches(0.5)
    top = Inches(2.5)
    width = Inches(9)
    height = Inches(2)
    
    title_box = slide.shapes.add_textbox(left, top, width, height)
    title_frame = title_box.text_frame
    title_frame.word_wrap = True
    p = title_frame.paragraphs[0]
    p.text = content.get("title", "Presentation")
    p.font.size = Pt(60)
    p.font.bold = True
    p.font.color.rgb = RGBColor(255, 255, 255)
    p.alignment = PP_ALIGN.CENTER
    
    # Content slides
    for slide_data in content.get("slides", []):
        slide_layout = prs.slide_layouts[1]  # Title and content layout
        slide = prs.slides.add_slide(slide_layout)
        
        # Set title
        title = slide.shapes.title
        title.text = slide_data.get("title", "Slide")
        
        # Set content
        body_shape = slide.placeholders[1]
        tf = body_shape.text_frame
        tf.clear()
        
        p = tf.paragraphs[0]
        p.text = slide_data.get("content", "")
        p.font.size = Pt(18)
        
        # Add image if available
        if slide_data.get("image_query") and UNSPLASH_ACCESS_KEY:
            image_url = fetch_image_from_unsplash(slide_data["image_query"])
            if image_url:
                try:
                    img_response = requests.get(image_url, timeout=10)
                    if img_response.status_code == 200:
                        img_stream = BytesIO(img_response.content)
                        left = Inches(5.5)
                        top = Inches(1.5)
                        height = Inches(4)
                        slide.shapes.add_picture(img_stream, left, top, height=height)
                except Exception as e:
                    st.warning(f"Could not add image: {e}")
    
    # Save to bytes
    output = BytesIO()
    prs.save(output)
    output.seek(0)
    return output.getvalue()

# Generate presentation
if generate_btn:
    if not topic:
        st.error("Please enter a presentation topic")
    elif not GEMINI_API_KEY:
        st.error("Please configure GEMINI_API_KEY in your .env file")
    else:
        with progress_placeholder.container():
            st.info("🔄 Generating presentation content...")
        
        # Generate content
        content = generate_presentation_content(topic, num_slides, presentation_style)
        
        if content:
            with progress_placeholder.container():
                st.info("🎨 Creating PowerPoint presentation...")
            
            # Create presentation
            pptx_bytes = create_presentation(content)
            
            with progress_placeholder.container():
                st.success("✅ Presentation generated successfully!")
            
            # Download button
            st.download_button(
                label="📥 Download Presentation",
                data=pptx_bytes,
                file_name=f"{topic.replace(' ', '_')}_presentation.pptx",
                mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                use_container_width=True
            )
            
            # Display generated content preview
            with st.expander("📋 View Generated Content"):
                st.json(content)
