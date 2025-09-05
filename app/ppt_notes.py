
import streamlit as st
import os
import time
import base64
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image
from io import BytesIO
import google.generativeai as genai
from google.generativeai.types import GenerationConfig
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Configure API key
genai.configure(api_key=os.getenv("GOOGLE_API_KEY"))
# Choose a Gemini model.
model = genai.GenerativeModel(model_name="gemini-2.0-flash-lite")

def ExtractText_LLM(image_base64: str, image_format: str = 'png', max_retries=3, delay=1, quota_refill_delay=60):
    generation_config = GenerationConfig(max_output_tokens=150)

    for attempt in range(max_retries):
        try:
            image_part = {
                'inline_data': {
                    'mime_type': f'image/{image_format}',
                    'data': image_base64
                }
            }

            result = model.generate_content(
                contents=[
                    image_part,
                    "\n",
                    "Analyze this image and provide a comprehensive description suitable for accessibility. "
                    "Include: main subject, key elements, context, and purpose. "
                    "Be descriptive but concise (under 125 characters for alt text). "
                    "Focus on what someone who can't see the image would need to know."
                ],
                generation_config=generation_config,
                request_options={"timeout": 10}
            )

            return result.text.strip()

        except Exception as e:
            if "Resource has been exhausted" in str(e):
                print(f"Quota exhausted, waiting {quota_refill_delay} seconds for refill...")
                time.sleep(quota_refill_delay)
            else:
                print(f"Attempt {attempt + 1} failed: {str(e)}")
                if attempt < max_retries - 1:
                    time.sleep(delay)
                else:
                    raise

def create_accessible_alt_text(ai_description, slide_number, image_number, context=""):
    """Create optimized alt text following accessibility best practices"""
    
    # Clean and truncate the AI description
    clean_desc = ai_description.strip()
    
    # If description is too long, create a shorter version
    if len(clean_desc) > 125:
        # Try to keep the most important parts
        words = clean_desc.split()
        short_desc = ""
        for word in words:
            if len(short_desc + " " + word) <= 120:
                short_desc += " " + word if short_desc else word
            else:
                break
        clean_desc = short_desc + "..."
    
    # Add context if provided
    if context:
        alt_text = f"{clean_desc} - {context}"
    else:
        alt_text = f"{clean_desc} - Image {image_number} on slide {slide_number}"
    
    # Ensure it's not too long
    if len(alt_text) > 125:
        alt_text = alt_text[:122] + "..."
    
    return alt_text

def process_powerpoint_with_ai(pptx_path, output_path):
    """Process PowerPoint to add comprehensive accessibility features"""
    prs = Presentation(pptx_path)
    
    total_images = 0
    processed_images = 0
    total_shapes = 0
    processed_shapes = 0

    for slide_index, slide in enumerate(prs.slides):
        notes_texts = []
        image_counter = 1
        slide_title = f"Slide {slide_index + 1}"
        
        # Try to get actual slide title
        for shape in slide.shapes:
            if hasattr(shape, 'text') and shape.text.strip():
                if not slide_title or slide_title == f"Slide {slide_index + 1}":
                    slide_title = shape.text.strip()
                    break
        
        notes_texts.append(f"=== {slide_title} ===\n")
        
        # Process all shapes for accessibility
        for shape_index, shape in enumerate(slide.shapes):
            total_shapes += 1
            
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                total_images += 1
                try:
                    image = shape.image
                    image_bytes = image.blob

                    # Load image using PIL
                    img = Image.open(BytesIO(image_bytes))

                    # Convert to base64
                    buffered = BytesIO()
                    image_format = img.format or "PNG"
                    img.save(buffered, format=image_format)
                    img_base64 = base64.b64encode(buffered.getvalue()).decode('utf-8')

                    # Normalize format
                    fmt = image_format.lower()
                    if fmt == "jpg":
                        fmt = "jpeg"

                    # Generate AI description
                    ai_description = ExtractText_LLM(img_base64, image_format=fmt)
                    
                    # Create optimized alt text
                    context = f"Slide: {slide_title}"
                    alt_text = create_accessible_alt_text(ai_description, slide_index + 1, image_counter, context)
                    
                    # Add alt text to the image
                    shape.element.set("descr", alt_text)
                    
                    # Add to notes with detailed information
                    note_text = (
                        f"🖼️ Image {image_counter}:\n"
                        f"📝 AI Analysis: {ai_description}\n"
                        f"♿ Alt Text: {alt_text}\n"
                        f"📊 Format: {fmt.upper()}\n"
                        f"📍 Position: Shape {shape_index + 1}\n"
                    )
                    notes_texts.append(note_text)
                    image_counter += 1
                    processed_images += 1
                    processed_shapes += 1
                    
                except Exception as e:
                    st.warning(f"Could not process image on slide {slide_index + 1}: {str(e)}")
                    # Add fallback alt text
                    fallback_alt = f"Image on slide {slide_index + 1} - Unable to process"
                    shape.alternative_text = fallback_alt
                    notes_texts.append(f"⚠️ Image {image_counter}: Processing failed - {str(e)}\n")
                    image_counter += 1
            
            elif hasattr(shape, 'text') and shape.text.strip():
                # Add text content to notes for context
                text_content = shape.text.strip()
                if len(text_content) > 100:
                    text_content = text_content[:97] + "..."
                
                notes_texts.append(f"📝 Text Content: {text_content}\n")
                processed_shapes += 1
            
            elif shape.shape_type == MSO_SHAPE_TYPE.CHART:
                # Handle charts
                chart_alt = f"Chart on slide {slide_index + 1} - {slide_title}"
                shape.alternative_text = chart_alt
                notes_texts.append(f"📊 Chart: {chart_alt}\n")
                processed_shapes += 1
            
            elif shape.shape_type == MSO_SHAPE_TYPE.TABLE:
                # Handle tables
                table_alt = f"Table on slide {slide_index + 1} - {slide_title}"
                shape.alternative_text = table_alt
                notes_texts.append(f"📋 Table: {table_alt}\n")
                processed_shapes += 1

        # Add or update notes - Fixed method
        try:
            if not slide.has_notes_slide:
                # Create notes slide if it doesn't exist
                notes_slide = slide.notes_slide
            else:
                notes_slide = slide.notes_slide
            
            # Get or create the notes text frame
            if not notes_slide.notes_text_frame:
                # If no text frame exists, we'll just skip adding notes for this slide
                continue
            
            notes_frame = notes_slide.notes_text_frame
            notes_frame.text = "\n".join(notes_texts) if notes_texts else "No content found on this slide."
            
        except Exception as e:
            # If we can't add notes, just continue - alt text is more important
            print(f"Could not add notes to slide {slide_index + 1}: {str(e)}")
            continue

    # Save the processed presentation
    prs.save(output_path)
    
    return total_images, processed_images, total_shapes, processed_shapes

def main():
    st.set_page_config(
        page_title="PowerPoint Accessibility Enhancer",
        page_icon="♿",
        layout="centered"
    )
    
    st.title("♿ PowerPoint Accessibility Enhancer")
    st.markdown("Transform your PowerPoint into an accessible masterpiece with AI-generated descriptions")
    
    # Check API key
    api_key = os.getenv("GOOGLE_API_KEY")
    if not api_key:
        st.error("❌ GOOGLE_API_KEY not found. Please set it in your .env file")
        st.stop()
    
    # File upload
    uploaded_file = st.file_uploader(
        "Choose a PowerPoint file",
        type=['pptx'],
        help="Upload a .pptx file to enhance accessibility"
    )
    
    if uploaded_file is not None:
        if st.button("♿ Enhance Accessibility", type="primary", use_container_width=True):
            with st.spinner("Enhancing PowerPoint accessibility..."):
                try:
                    # Save uploaded file temporarily
                    temp_path = f"temp_{uploaded_file.name}"
                    with open(temp_path, "wb") as f:
                        f.write(uploaded_file.getvalue())
                    
                    # Process the file
                    output_path = f"accessible_{uploaded_file.name}"
                    total_images, processed_images, total_shapes, processed_shapes = process_powerpoint_with_ai(temp_path, output_path)
                    
                    # Clean up temp file
                    os.remove(temp_path)
                    
                    st.success(f"✅ Accessibility enhancement complete!")
                    
                    # Show detailed results
                    col1, col2 = st.columns(2)
                    with col1:
                        st.info(f"🖼️ Images: {processed_images}/{total_images}")
                    with col2:
                        st.info(f"📊 Shapes: {processed_shapes}/{total_shapes}")
                    
                    st.info("♿ Accessibility features added:")
                    st.write("• **Alt Text**: AI-generated descriptions for all images")
                    st.write("• **Slide Notes**: Comprehensive content descriptions")
                    st.write("• **Chart/Table Labels**: Accessibility labels for data elements")
                    st.write("• **Context Information**: Slide titles and positioning details")
                    
                    # Download button
                    with open(output_path, "rb") as f:
                        st.download_button(
                            label="📥 Download Accessible PowerPoint",
                            data=f.read(),
                            file_name=output_path,
                            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                            type="primary",
                            use_container_width=True
                        )
                    
                    # Clean up output file
                    os.remove(output_path)
                    
                except Exception as e:
                    st.error(f"Error processing file: {str(e)}")
                    # Clean up temp file if it exists
                    if os.path.exists(temp_path):
                        os.remove(temp_path)

if __name__ == "__main__":
    main()