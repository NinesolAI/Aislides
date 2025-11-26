import os
import re
import json
from typing import Dict, Any, List, Optional
from io import BytesIO

# Third-party libraries
from fastapi import FastAPI, HTTPException, status
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel, Field
from openai import OpenAI
from bs4 import BeautifulSoup
from pptx import Presentation
from pptx.util import Inches
import aiofiles
from dotenv import load_dotenv

# Load environment variables from .env file
load_dotenv()

# --- Configuration and Setup ---

# Initialize OpenAI client (API key is loaded from environment variable OPENAI_API_KEY)
try:
    # client = OpenAI()
    client = OpenAI(api_key=os.getenv("OPENAI_API_KEY"))
except Exception as e:
    # This is a placeholder for the actual error handling, as the client is initialized without args
    # because the API key is pre-configured in the environment.
    print(f"OpenAI client initialization failed: {e}")
    client = None

# Define the themes and their corresponding file paths
THEMES = {
    "Marketing": "./upload/marketing.html",
    "Education": "./upload/education.html",
    "Portfolio": "./upload/portfolio.html",
    "Technology": "./upload/technology.html",
    "Business": "./upload/business.html",
}

# --- Pydantic Models ---

class PresentationRequest(BaseModel):
    """Input model for presentation generation endpoints."""
    topic: str = Field(..., description="The topic/title of the presentation.")
    theme: str = Field(..., description="The theme of the presentation (e.g., Marketing, Education).")
    num_slides: int = Field(10, ge=1, le=10, description="The number of slides to generate content for (max 10).")

class PresentationResponse(BaseModel):
    """Output model for the HTML generation endpoint."""
    html_content: str

# --- Utility Functions ---

# def get_template_path(theme: str) -> str:
#     """Validates the theme and returns the corresponding file path."""
#     path = THEMES.get(theme)
#     if not path or not os.path.exists(path):
#         raise HTTPException(
#             status_code=status.HTTP_400_BAD_REQUEST,
#             detail=f"Invalid theme: {theme}. Must be one of {list(THEMES.keys())}"
#         )
#     return path

def get_template_path(theme: str) -> str:
    theme_key = next((k for k in THEMES.keys() if k.lower() == theme.lower()), None)
    if not theme_key:
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Invalid theme: {theme}. Must be one of {list(THEMES.keys())}"
        )
    path = THEMES[theme_key]
    if not os.path.exists(path):
        raise HTTPException(
            status_code=status.HTTP_400_BAD_REQUEST,
            detail=f"Template file not found for theme: {theme}"
        )
    return path


def extract_placeholder_text(html_content: str, num_slides: int) -> Dict[str, str]:
    """
    Parses the HTML to find all text content within the first `num_slides`
    and returns a dictionary mapping a unique ID to the original text.
    """
    soup = BeautifulSoup(html_content, 'html.parser')
    placeholders = {}
    
    # Find all slide containers
    slide_containers = soup.find_all('div', class_='slide-container')
    
    # Limit to the requested number of slides
    slides_to_process = slide_containers[:num_slides]
    
    for slide_index, slide in enumerate(slides_to_process):
        # Find all text-containing elements (e.g., h1, h2, p, li)
        text_elements = slide.find_all(re.compile(r'^(h[1-6]|p|li|span|div)$'))
        
        for element_index, element in enumerate(text_elements):
            # Clean up the text content
            text = element.get_text(strip=True)
            
            # Only consider non-empty text that is not purely navigation/style related
            if text and len(text) > 5 and not re.match(r'^\d+ / \d+$', text):
                # Create a unique ID for the placeholder
                placeholder_id = f"SLIDE_{slide_index+1}_ELEM_{element_index+1}"
                placeholders[placeholder_id] = text
                
                # Replace the original text with the unique ID marker
                # This is a complex operation with BeautifulSoup, so we'll simplify the prompt
                # and rely on the model to return a JSON structure for replacement.
                # For now, we'll just collect the text and send the whole HTML as context.
                pass
                
    return placeholders

async def generate_content_with_openai(html_template: str, topic: str, theme: str, num_slides: int) -> str:
    """
    Calls the OpenAI API to generate new content based on the HTML structure.
    Returns the full HTML with the new content.
    """
    if not client:
        raise HTTPException(
            status_code=status.HTTP_503_SERVICE_UNAVAILABLE,
            detail="OpenAI client is not initialized. Check API key configuration."
        )

    # 1. Extract the HTML structure for the first `num_slides`
    soup = BeautifulSoup(html_template, 'html.parser')
    slide_containers = soup.find_all('div', class_='slide-container')
    
    # Get the HTML for the slides to be processed
    slides_html = "".join(str(s) for s in slide_containers[:num_slides])
    
    # Get the rest of the HTML (before and after the slides)
    # This is complex, so we will use a simpler approach: send the relevant slides and ask for the *new* content only.
    # The model will be asked to return a JSON object with the new content for each element.
    
    # 2. Extract all text content from the relevant slides to be replaced
    # We will use a more robust method: find all text nodes and replace them with a unique marker.
    
    # Create a copy of the soup for marking
    marked_soup = BeautifulSoup(html_template, 'html.parser')
    marked_slides = marked_soup.find_all('div', class_='slide-container')[:num_slides]
    
    placeholder_map = {}
    
    for slide_index, slide in enumerate(marked_slides):
        # Find all text nodes that are not inside script or style tags
        for element in slide.find_all(text=True):
            if element.parent.name not in ['script', 'style', 'title'] and element.strip():
                # Check if the text is not a navigation element (like '← Previous')
                if not re.match(r'^\s*(\d+ / \d+|← Previous|Next →|Explore Our Vision|Learn More|Schedule Meeting|Download Brochure|f|in|𝕏|Company Image|Years in Business|Team Members|CAGR|Total Addressable Market|By 2025|Revenue Growth|Valuation|Series C Funding|Customer Retention|Magic Number|Average Contract Value|Gross Margin|Average Experience|Team Size|Global Offices|Product Development|Beta Testing|Market Launch|Geographic Expansion|Partnership Development|Sales Growth|M&A Activity|Integration|Synergy Realization|Regulatory Compliance|Financial Audit|IPO Launch)\s*$', element.strip()):
                    # Create a unique ID for the placeholder
                    placeholder_id = f"SLIDE_{slide_index+1}_TEXT_{len(placeholder_map) + 1}"
                    placeholder_map[placeholder_id] = element.strip()
                    
                    # Replace the text node with the unique ID marker
                    element.replace_with(f"[[{placeholder_id}]]")

    # The prompt will ask the model to fill in the content for the marked IDs.
    prompt_template = f"""
    You are an expert presentation content generator. Your task is to rewrite the content for a presentation on the topic: "{topic}" with the theme "{theme}".
    
    The presentation structure is based on the following HTML. I have replaced all text content that needs to be rewritten with unique placeholder IDs in the format [[SLIDE_X_TEXT_Y]].
    
    Your goal is to provide new, relevant content for each placeholder ID.
    
    Rules:
    1. The new content MUST be based on the topic "{topic}" and the theme "{theme}".
    2. You MUST return a single JSON object with the structure: {{"new_content": {{"SLIDE_1_TEXT_1": "New content for slide 1", ...}}}}.
    3. The keys of the 'new_content' dictionary MUST be the exact placeholder IDs (e.g., "SLIDE_1_TEXT_1").
    4. The values of the 'new_content' dictionary MUST be the new text content for that placeholder.
    5. The new content should be concise and professional, suitable for a presentation slide.
    6. DO NOT include any introductory text, markdown, code blocks, or extra text outside the single JSON object.
    
    Here is the map of placeholder IDs and their original text (for context on the element's purpose):
    {json.dumps(placeholder_map, indent=2)}
    
    
    Return ONLY the JSON object.
    """

    try:
        response = client.chat.completions.create(
            model="gpt-4o-mini", # Using a fast model for content generation
            messages=[
                {"role": "system", "content": "You are an expert presentation content generator. You return a single JSON object with the new content for the provided placeholder IDs."},
                {"role": "user", "content": prompt_template}
            ],
            # response_format={"type": "json_object"} # This is not supported by the current client version
        )
        
        # Parse the JSON response
        response_text = response.choices[0].message.content
        
        # Write the raw response to a file for debugging
        async with aiofiles.open("./openai_response.log", mode="w") as f:
            await f.write(response_text)
            
        # Attempt to extract JSON from the response text
        json_match = re.search(r'```json\s*(\{.*\})\s*```', response_text, re.DOTALL)
        if json_match:
            json_content = json_match.group(1)
        else:
            # Fallback for non-markdown JSON
            json_match = re.search(r'(\{.*\})', response_text, re.DOTALL)
            if not json_match:
                raise HTTPException(
                    status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                    detail="OpenAI did not return a valid JSON object."
                )
            json_content = json_match.group(0)
            
        try:
            new_content_map = json.loads(json_content)["new_content"]
        except (json.JSONDecodeError, KeyError) as e:
            raise HTTPException(
                status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
                detail=f"Failed to parse OpenAI response: {e}. Raw content saved to openai_response.log"
            )
            
        # 3. Replace the markers in the HTML with the new content
        final_html = str(marked_soup)
        for placeholder_id, new_text in new_content_map.items():
            # Escape special characters in the new text for safe regex replacement
            # We use re.escape on the placeholder ID to ensure it's treated literally
            # We use re.sub to replace the marker with the new content
            final_html = re.sub(re.escape(f"[[{placeholder_id}]]"), new_text, final_html)

        return final_html
        
    except Exception as e:
        raise HTTPException(
            status_code=status.HTTP_500_INTERNAL_SERVER_ERROR,
            detail=f"OpenAI API call failed: {e}"
        )

# def html_to_ppt(html_content: str) -> BytesIO:
#     """
#     Converts the AI-generated HTML slides into a PowerPoint presentation.
#     This is a simplified conversion, focusing on extracting main text elements.
#     """
#     prs = Presentation()
#     soup = BeautifulSoup(html_content, 'html.parser')
#     slide_containers = soup.find_all('div', class_='slide-container')
    
#     # Use a blank slide layout (index 6) for maximum flexibility
#     blank_slide_layout = prs.slide_layouts[6]

#     for slide_index, slide_div in enumerate(slide_containers):
#         # Add a new slide
#         slide = prs.slides.add_slide(blank_slide_layout)
        
#         # Define a text box for the content (simplified for this example)
#         left = top = width = height = Inches(0.5)
#         width = Inches(9)
#         height = Inches(6.5)
#         txBox = slide.shapes.add_textbox(left, top, width, height)
#         tf = txBox.text_frame
        
#         # Extract all relevant text from the slide
#         all_text = []
        
#         # Find main headings and paragraphs
#         main_elements = slide_div.find_all(['h1', 'h2', 'h3', 'p', 'li'])
        
#         for element in main_elements:
#             text = element.get_text(strip=True)
#             if text and not re.match(r'^\s*(\d+ / \d+|← Previous|Next →|Explore Our Vision|Learn More|Schedule Meeting|Download Brochure|f|in|𝕏|Company Image|Years in Business|Team Members|CAGR|Total Addressable Market|By 2025|Revenue Growth|Valuation|Series C Funding|Customer Retention|Magic Number|Average Contract Value|Gross Margin|Average Experience|Team Size|Global Offices|Product Development|Beta Testing|Market Launch|Geographic Expansion|Partnership Development|Sales Growth|M&A Activity|Integration|Synergy Realization|Regulatory Compliance|Financial Audit|IPO Launch)\s*$', text):
#                 all_text.append(text)

#         # Add text to the PowerPoint slide
#         if all_text:
#             # First element is usually the title
#             title_text = all_text[0]
#             p = tf.paragraphs[0]
#             p.text = title_text
#             p.font.size = Inches(0.3) # Approx 32pt
            
#             # Remaining elements as bullet points
#             for text in all_text[1:]:
#                 p = tf.add_paragraph()
#                 p.text = text
#                 p.level = 0
#                 p.font.size = Inches(0.2) # Approx 20pt

#     # Save the presentation to a BytesIO object
#     ppt_io = BytesIO()
#     prs.save(ppt_io)
#     ppt_io.seek(0)
#     return ppt_io

from pptx.util import Pt
from pptx.dml.color import RGBColor

def html_to_ppt(html_content: str) -> BytesIO:
    """
    Converts the AI-generated HTML slides into a PowerPoint presentation with
    basic styling: headings bolded, bullet points, and font sizes.
    """
    prs = Presentation()
    soup = BeautifulSoup(html_content, 'html.parser')
    slide_containers = soup.find_all('div', class_='slide-container')

    # Use a blank slide layout for flexibility
    blank_slide_layout = prs.slide_layouts[6]

    for slide_index, slide_div in enumerate(slide_containers):
        slide = prs.slides.add_slide(blank_slide_layout)

        # Define a text box for the content
        left = top = Inches(0.5)
        width = Inches(9)
        height = Inches(6.5)
        txBox = slide.shapes.add_textbox(left, top, width, height)
        tf = txBox.text_frame
        tf.word_wrap = True

        # Extract headings and paragraphs
        elements = slide_div.find_all(['h1', 'h2', 'h3', 'h4', 'h5', 'h6', 'p', 'li'])
        for i, elem in enumerate(elements):
            text = elem.get_text(strip=True)
            if not text:
                continue

            # Determine font size based on heading level
            if elem.name == 'h1':
                font_size = Pt(36)
                bold = True
            elif elem.name == 'h2':
                font_size = Pt(32)
                bold = True
            elif elem.name == 'h3':
                font_size = Pt(28)
                bold = True
            elif elem.name == 'h4':
                font_size = Pt(24)
                bold = True
            else:  # h5, h6, p, li
                font_size = Pt(20)
                bold = False

            # First element in the text frame uses the first paragraph
            if i == 0:
                p = tf.paragraphs[0]
            else:
                p = tf.add_paragraph()

            p.text = text
            p.font.size = font_size
            p.font.bold = bold
            p.level = 0 if elem.name.startswith('h') else 1  # indent bullet points for non-headings
            p.font.color.rgb = RGBColor(0, 0, 0)  # black text by default

    # Save presentation to BytesIO
    ppt_io = BytesIO()
    prs.save(ppt_io)
    ppt_io.seek(0)
    return ppt_io

# --- FastAPI Application ---

app = FastAPI(
    title="AI Presentation Generator",
    description="A service to generate themed presentations in HTML and PowerPoint formats using OpenAI.",
    version="1.0.0"
)

# Add CORS middleware to allow all origins for development/testing
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/")
async def root():
    return {"message": "AI Presentation Generator API is running. See /docs for endpoints."}

from fastapi.responses import HTMLResponse

@app.post("/generate-html", response_class=HTMLResponse)
async def generate_html_presentation(request: PresentationRequest):
    """
    Generates an HTML presentation with AI-filled content based on the topic and theme.
    """
    theme_path = get_template_path(request.theme)
    
    # Read the HTML template asynchronously
    async with aiofiles.open(theme_path, mode="r") as f:
        html_template = await f.read()
        
    # Generate content and replace placeholders
    final_html = await generate_content_with_openai(
        html_template, 
        request.topic, 
        request.theme, 
        request.num_slides
    )
    
    return HTMLResponse(content=final_html, status_code=200)

@app.post("/generate-ppt")
async def generate_ppt_presentation(request: PresentationRequest):
    """
    Generates a PowerPoint (.pptx) presentation with AI-filled content and returns it as a file download.
    """
    # 1. Generate the HTML content first
    theme_path = get_template_path(request.theme)
    async with aiofiles.open(theme_path, mode="r") as f:
        html_template = await f.read()
        
    final_html = await generate_content_with_openai(
        html_template, 
        request.topic, 
        request.theme, 
        request.num_slides
    )
    
    # 2. Convert HTML to PPTX
    ppt_io = html_to_ppt(final_html)
    
    # 3. Return the PPTX file as a download
    from fastapi.responses import StreamingResponse
    
    filename = f"{request.theme}_{request.topic.replace(' ', '_')}.pptx"
    
    return StreamingResponse(
        ppt_io,
        media_type="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

# To run the application: uvicorn app:app --host 0.0.0.0 --port 8000
