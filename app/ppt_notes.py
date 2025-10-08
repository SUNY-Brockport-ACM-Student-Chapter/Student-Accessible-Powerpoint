
import streamlit as st
import os
import time
import base64
import io
import uuid
import csv
from datetime import datetime
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE
from PIL import Image, UnidentifiedImageError
from io import BytesIO
import google.generativeai as genai
from google.generativeai.types import GenerationConfig
from dotenv import load_dotenv

# Import Wand for WMF support
try:
    from wand.image import Image as WandImage
    WAND_AVAILABLE = True
except ImportError:
    WAND_AVAILABLE = False

def safe_open_image(img_bytes: bytes):
    """Safely load images including WMF formats that Pillow cannot handle.
    
    Args:
        img_bytes: Image data as bytes
        
    Returns:
        PIL Image object (converted to PNG if WMF detected)
        
    Raises:
        RuntimeError: If WMF images cannot be processed and Wand is not available
    """
    # Always try to open with Pillow first
    try:
        img = Image.open(io.BytesIO(img_bytes))
        if img.format and img.format.lower() == "wmf":
            raise OSError("Force WMF fallback")
        return img
    except (OSError, UnidentifiedImageError):
        if not WAND_AVAILABLE:
            raise RuntimeError("Wand (ImageMagick) is required to handle WMF/EMF images. "
                               "Install with `pip install Wand` and ensure ImageMagick is installed.")
        # Convert WMF → PNG in memory
        with WandImage(blob=img_bytes, format="wmf") as wmf:
            png_bytes = wmf.make_blob("png")
        return Image.open(io.BytesIO(png_bytes))

# Save consent response (email only when opted in)
def save_consent_email(email: str, choice: str):
    try:
        file_path = "consent_responses.csv"
        file_exists = os.path.exists(file_path)
        with open(file_path, "a", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            if not file_exists:
                writer.writerow(["timestamp_utc", "email", "choice"])  # header
            writer.writerow([datetime.utcnow().isoformat(), email, choice])
    except Exception as e:
        print(f"Error saving consent response: {e}")

# Import our RAG modules
from pptx_rag_quizzer.utils import parse_powerpoint
from pptx_rag_quizzer.rag_core import RAGCore
from pptx_rag_quizzer.image import Image as ImageProcessor
from models.models import Presentation as PresentationModel

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

def parse_powerpoint_file(file_bytes, file_name):
    """Parse PowerPoint file into our model format"""
    file_object = io.BytesIO(file_bytes)
    return parse_powerpoint(file_object, file_name)

def process_powerpoint_with_rag_enhanced(pptx_path, output_path, presentation_model, collection_id, image_descriptions):
    """Process PowerPoint with RAG-enhanced accessibility features"""
    prs = Presentation(pptx_path)
    
    # Initialize RAG core and image processor
    rag_core = RAGCore()
    image_processor = ImageProcessor(rag_core)
    
    total_images = 0
    processed_images = 0
    total_shapes = 0
    processed_shapes = 0

    for slide_index, slide in enumerate(prs.slides):
        slide_model = presentation_model.slides[slide_index] if slide_index < len(presentation_model.slides) else None
        
        notes_texts = []
        image_counter = 1
        slide_title = f"Slide {slide_index + 1}"
        
        # Try to get actual slide title from the model
        if slide_model:
            for item in slide_model.items:
                if item.type.value == 'text' and item.content.strip():
                    if not slide_title or slide_title == f"Slide {slide_index + 1}":
                        slide_title = item.content.strip()[:50]  # Limit title length
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

                    # Get image description from our enhanced processing
                    image_description = None
                    if slide_model:
                        for item in slide_model.items:
                            if item.type.value == 'image' and hasattr(item, 'image_bytes'):
                                # Compare image bytes to find matching description
                                if item.image_bytes == image_bytes:
                                    image_description = item.content
                                    break
                    
                    if not image_description or image_description.lower() in ['none', 'null', '']:
                        # Fallback to basic description
                        image_description = "Image content - detailed description not available"
                    
                    # Use the image caption/description directly for alt text
                    alt_text = (image_description or "").strip()
                    if not alt_text:
                        alt_text = f"Image {image_counter} on slide {slide_index + 1}"
                    

                    # Set native PPTX alt text via underlying cNvPr descr attribute
                    try:
                        shape._element._nvXxPr.cNvPr.attrib["descr"] = alt_text
                    except Exception:
                        # Fallback to python-pptx property if needed
                        try:
                            shape.alternative_text = alt_text
                        except Exception:
                            pass
                    
                    # Add to notes with detailed information
                    note_text = (
                        f"🖼️ Image {image_counter}:\n"
                        f"📝 AI Analysis: {image_description}\n"
                        f"♿ Alt Text: {alt_text}\n"
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
                    try:
                        shape._element._nvXxPr.cNvPr.attrib["descr"] = fallback_alt
                    except Exception:
                        try:
                            shape.alternative_text = fallback_alt
                        except Exception:
                            pass
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

        # Enhanced notes generation with RAG context
        try:
            if collection_id:
                # Get context from RAG for this slide
                try:
                    slide_context = rag_core.get_context_from_slide_number(slide_index + 1, collection_id)
                    context_docs = slide_context.get("documents", "")
                    
                    if context_docs:
                        # Generate enhanced notes with context
                        enhanced_notes = generate_enhanced_notes_with_context(
                            slide_title, notes_texts, context_docs, rag_core
                        )
                        notes_texts = enhanced_notes
                        
                except Exception as e:
                    print(f"Could not retrieve context for slide {slide_index + 1}: {e}")
                    # Continue with basic notes
            
            # Add or update notes
            if not slide.has_notes_slide:
                notes_slide = slide.notes_slide
            else:
                notes_slide = slide.notes_slide
            
            if not notes_slide.notes_text_frame:
                continue
            
            notes_frame = notes_slide.notes_text_frame
            notes_frame.text = "\n".join(notes_texts) if notes_texts else "No content found on this slide."
            
        except Exception as e:
            print(f"Could not add notes to slide {slide_index + 1}: {str(e)}")
            continue

    # Save the processed presentation
    prs.save(output_path)
    
    return total_images, processed_images, total_shapes, processed_shapes

def generate_enhanced_notes_with_context(slide_title, basic_notes, context_docs, rag_core):
    """Generate enhanced notes using RAG context"""
    try:
        context_str = context_docs if isinstance(context_docs, str) else " ".join(context_docs)
        basic_notes_str = "\n".join(basic_notes)
        
        prompt = f"""
        You are an accessibility expert creating comprehensive slide notes for students.
        
        Slide Title: {slide_title}
        
        Basic Content:
        {basic_notes_str}
        
        Additional Context from Presentation:
        {context_str}
        
        Create enhanced, comprehensive accessibility notes that:
        1. Summarize the slide's main points clearly
        2. Explain any complex concepts in simple terms
        3. Provide context for images and visual elements
        4. Include key takeaways for students
        5. Maintain a helpful, educational tone
        
        Format the response as structured notes suitable for student accessibility.
        """
        
        enhanced_notes = rag_core.prompt_gemini(prompt, max_output_tokens=500)
        
        # Return enhanced notes with original structure
        enhanced_notes_list = [f"=== {slide_title} ===", ""]
        enhanced_notes_list.extend(enhanced_notes.split('\n'))
        enhanced_notes_list.extend(["", "---", ""])
        enhanced_notes_list.extend(basic_notes)
        
        return enhanced_notes_list
        
    except Exception as e:
        print(f"Error generating enhanced notes: {e}")
        return basic_notes



def main():
    st.set_page_config(
        page_title="PowerPoint Accessibility Enhancer",
        page_icon="♿",
        layout="centered"
    )
    
    st.title("♿ PowerPoint Accessibility Enhancer")
    st.markdown("Transform your PowerPoint into an accessible masterpiece with RAG-enhanced AI descriptions")
    st.warning("AI-Generated Content: The following text is created by an AI model. It may contain errors, omissions, or outdated information. Please exercise caution and confirm details before use.")
    
    # Initialize session state
    if 'processing_stage' not in st.session_state:
        st.session_state.processing_stage = 'consent'
    if 'presentation_model' not in st.session_state:
        st.session_state.presentation_model = None
    if 'rag_core' not in st.session_state:
        st.session_state.rag_core = None
    if 'image_processor' not in st.session_state:
        st.session_state.image_processor = None
    if 'collection_id' not in st.session_state:
        st.session_state.collection_id = None
    if 'current_batch' not in st.session_state:
        st.session_state.current_batch = 0
    if 'batch_size' not in st.session_state:
        st.session_state.batch_size = 5
    if 'consent_completed' not in st.session_state:
        st.session_state.consent_completed = False
    if 'consent_choice' not in st.session_state:
        st.session_state.consent_choice = None
    if 'consent_email' not in st.session_state:
        st.session_state.consent_email = ""
    
    # Check API key
    api_key = os.getenv("GOOGLE_API_KEY")
    if not api_key:
        st.error("❌ GOOGLE_API_KEY not found. Please set it in your .env file")
        st.stop()
    
    # Consent Gate
    if st.session_state.processing_stage == 'consent':
        st.header("Consent to Participate")
        st.markdown(
            """
            Dear Colleagues,

            You are invited to participate in this research on “Enhancing Accessibility in Higher Education through Experiential Learning Opportunities” by participating a one-minute survey later (via email), or participating in one focus group discussion session, or one-on-one interview if you cannot make the focus group session. Confidentiality will be reminded at the beginning of the session.

            If you choose to volunteer to participate in the study, what you share in the focus group or interview will be included in the research analysis. The data will be aggregated and no participants will be identified.

            If you choose to opt out of the study, no further action is needed. There is no penalty if you choose not to participate.

            You can simply check one of the following boxes to indicate your participation.
            """
        )


        consent_options = [
            "Yes - I agree to participate in this research project. I am 18 years of age or older.",
            "Yes, I agree. I have provided my email address, and I confirm that I am 18 years of age or older.",
            "No - I do not agree to participate in this research project. I am 18 years of age or older.",
            "No - I am not eligible to participate as I am under the age of 18.",
        ]
        consent_choice = st.radio("Please select one option to continue:", consent_options, index=None)

        email = ""
        if consent_choice == consent_options[0]:
            email = st.text_input("Email (for 1-minute survey)", value=st.session_state.consent_email, placeholder="name@domain.edu")

        if st.button("Continue", type="primary", use_container_width=True):
            if consent_choice is None:
                st.error("Please select an option to continue.")
            elif consent_choice == consent_options[0]:
                if not email.strip():
                    st.error("Please enter your email to continue.")
                else:
                    # Save and proceed
                    st.session_state.consent_choice = 'yes'
                    st.session_state.consent_email = email.strip()
                    save_consent_email(st.session_state.consent_email, st.session_state.consent_choice)
                    st.session_state.consent_completed = True
                    st.session_state.processing_stage = 'upload'
                    st.rerun()
            elif consent_choice == consent_options[2] or consent_choice == consent_options[1]:
                st.session_state.consent_choice = 'no'
                st.session_state.consent_completed = True
                st.session_state.processing_stage = 'upload'
                st.rerun()
            else:
                st.session_state.consent_choice = 'under_18'
                st.session_state.consent_completed = False
                st.session_state.processing_stage = 'blocked'
                st.rerun()

    # Stage 0: Blocked (under 18)
    elif st.session_state.processing_stage == 'blocked':
        st.header("Access Restricted")
        st.error("You are not eligible to participate as you are under the age of 18. Please close this page.")
        st.stop()

    # Stage 1: File Upload and Initial Processing
    elif st.session_state.processing_stage == 'upload':
        if not st.session_state.consent_completed:
            st.session_state.processing_stage = 'consent'
            st.rerun()
        st.header("📁 Step 1: Upload PowerPoint")
        
        uploaded_file = st.file_uploader(
            "Choose a PowerPoint file",
            type=['pptx'],
            help="Upload a .pptx file to enhance accessibility"
        )
        
        if uploaded_file is not None:
            if st.button("🚀 Process PowerPoint", type="primary", use_container_width=True):
                with st.spinner("Processing PowerPoint and building RAG collection..."):
                    try:
                        # Parse the PowerPoint file
                        file_bytes = uploaded_file.read()
                        presentation_model = parse_powerpoint_file(file_bytes, uploaded_file.name)
                        
                        # Initialize RAG core and create collection from text content
                        rag_core = RAGCore()
                        collection_id = rag_core.create_collection(presentation_model)
                        
                        # Initialize image processor
                        image_processor = ImageProcessor(rag_core)
                        
                        # Store in session state
                        st.session_state.presentation_model = presentation_model
                        st.session_state.rag_core = rag_core
                        st.session_state.image_processor = image_processor
                        st.session_state.collection_id = collection_id
                        st.session_state.uploaded_file_name = uploaded_file.name
                        st.session_state.file_bytes = file_bytes
                        
                        st.session_state.processing_stage = 'describe_images'
                        st.rerun()
                        
                    except Exception as e:
                        st.error(f"❌ Error processing PowerPoint: {str(e)}")
                        st.exception(e)

            if st.button("⚡ Quick Generate & Download (Skip Review)", use_container_width=True):
                with st.spinner("Generating accessible PowerPoint (skipping review)..."):
                    try:
                        # Parse the PowerPoint file
                        file_bytes = uploaded_file.read()
                        presentation_model = parse_powerpoint_file(file_bytes, uploaded_file.name)

                        # Initialize RAG core and create initial collection from text content
                        rag_core = RAGCore()
                        collection_id = rag_core.create_collection(presentation_model)

                        # Initialize image processor
                        image_processor = ImageProcessor(rag_core)

                        # Auto-generate descriptions for all images without going through review UI
                        for slide in presentation_model.slides:
                            for item in slide.items:
                                if item.type.value == 'image':
                                    if not item.content or item.content.lower() in ['none', 'null', '']:
                                        try:
                                            image_description = image_processor.describe_image(
                                                item.image_bytes,
                                                item.extension,
                                                item.slide_number,
                                                collection_id
                                            )
                                            if image_description and image_description.startswith("Description: "):
                                                image_description = image_description[len("Description: "):]
                                            if image_description and image_description != "None":
                                                item.content = image_description
                                            else:
                                                item.content = "No description available"
                                        except Exception as e:
                                            item.content = f"Error describing image: {e}"

                        # Rebuild collection including image descriptions
                        rag_core.remove_collection(collection_id)
                        enhanced_collection_id = rag_core.create_collection(presentation_model)

                        # Save temporary file for processing
                        temp_path = f"temp_{uploaded_file.name}"
                        with open(temp_path, "wb") as f:
                            f.write(file_bytes)

                        # Process with enhanced RAG and write accessibility features
                        output_path = f"accessible_{uploaded_file.name}"
                        process_powerpoint_with_rag_enhanced(
                            temp_path, output_path, presentation_model, enhanced_collection_id, {}
                        )

                        # Clean up temp file
                        os.remove(temp_path)

                        # Jump straight to download stage
                        st.session_state.output_path = output_path
                        st.session_state.processing_stage = 'download'
                        st.rerun()

                    except Exception as e:
                        st.error(f"❌ Error generating accessible PowerPoint: {str(e)}")
                        st.exception(e)
    
    # Stage 2: Image Description and Batch Processing
    elif st.session_state.processing_stage == 'describe_images':
        st.header("🖼️ Step 2: Describe Images")
        
        # Back button
        if st.button("← Back to Upload", use_container_width=True):
            st.session_state.processing_stage = 'upload'
            st.rerun()
        
        presentation_model = st.session_state.presentation_model
        image_processor = st.session_state.image_processor
        collection_id = st.session_state.collection_id
        
        # Get all images from the presentation
        all_images = []
        for slide in presentation_model.slides:
            for item in slide.items:
                if item.type.value == 'image':
                    all_images.append(item)
        
        total_images = len(all_images)
        
        if total_images == 0:
            st.info("📝 No images found in the presentation.")
            st.session_state.processing_stage = 'final_processing'
            st.rerun()
            return
        
        # Batch processing
        batch_size = st.session_state.batch_size
        current_batch = st.session_state.current_batch
        batch_start = current_batch * batch_size
        batch_end = min(batch_start + batch_size, total_images)
        current_batch_images = all_images[batch_start:batch_end]
        
        # Show progress
        st.write(f"**Processing images {batch_start + 1} to {batch_end} of {total_images}**")
        st.progress((batch_end) / total_images)
        
        # Process current batch if descriptions are missing
        batch_ready = all(
            img.content and img.content.lower() not in ['none', 'null', ''] 
            for img in current_batch_images
        )
        
        if not batch_ready:
            st.write("🤖 Generating AI descriptions for images...")
            
            with st.spinner("AI is analyzing images and generating descriptions..."):
                for i, img_item in enumerate(current_batch_images):
                    if not img_item.content or img_item.content.lower() in ['none', 'null', '']:
                        try:
                            st.write(f"Describing image {batch_start + i + 1} of {total_images}...")
                            image_description = image_processor.describe_image(
                                img_item.image_bytes,
                                img_item.extension,
                                img_item.slide_number,
                                collection_id
                            )
                            
                            # Clean up description
                            if image_description and image_description.startswith("Description: "):
                                image_description = image_description[len("Description: "):]
                            
                            if image_description and image_description != "None":
                                img_item.content = image_description
                                st.write(f"✓ Image {batch_start + i + 1} described successfully")
                            else:
                                img_item.content = "No description available"
                                st.write(f"⚠️ No description generated for image {batch_start + i + 1}")
                                
                        except Exception as e:
                            img_item.content = f"Error describing image: {e}"
                            st.write(f"✗ Error describing image {batch_start + i + 1}: {e}")
            
            st.success("All descriptions generated! Please review and approve.")
            st.rerun()
        
        else:
            # Display current batch for review
            st.write("📋 **Review and approve image descriptions:**")
            
            for i, img_item in enumerate(current_batch_images):
                st.write(f"**Image {batch_start + i + 1}** (Slide {img_item.slide_number})")
                
                # Display image
                img = safe_open_image(img_item.image_bytes)
                st.image(img, width=400, caption=f"Slide {img_item.slide_number} - Image {batch_start + i + 1}")
                
                # Editable description
                new_description = st.text_area(
                    f"Description for image {batch_start + i + 1}:",
                    value=img_item.content,
                    height=100,
                    key=f"desc_{batch_start + i}"
                )
                img_item.content = new_description
                st.write("---")
            
            # Navigation buttons
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if batch_start > 0:
                    if st.button("← Previous Batch"):
                        st.session_state.current_batch = current_batch - 1
                        st.rerun()
            
            with col2:
                if st.button("💾 Save Batch"):
                    st.success("Batch saved!")
            
            with col3:
                if batch_end < total_images:
                    if st.button("Next Batch →"):
                        st.session_state.current_batch = current_batch + 1
                        st.rerun()
                else:
                    # All images processed
                    if st.button("✅ Finish Image Processing", type="primary"):
                        st.session_state.processing_stage = 'final_processing'
                        st.rerun()
    
    # Stage 3: Final Processing with Enhanced RAG
    elif st.session_state.processing_stage == 'final_processing':
        st.header("🎯 Step 3: Final Processing")
        
        # Back button
        if st.button("← Back to Image Processing", use_container_width=True):
            st.session_state.processing_stage = 'describe_images'
            st.rerun()
        
        if st.button("🚀 Generate Enhanced Accessibility Features", type="primary", use_container_width=True):
            with st.spinner("Building enhanced RAG collection and generating accessibility features..."):
                try:
                    presentation_model = st.session_state.presentation_model
                    rag_core = st.session_state.rag_core
                    collection_id = st.session_state.collection_id
                    file_bytes = st.session_state.file_bytes
                    file_name = st.session_state.uploaded_file_name
                    
                    # Remove old collection and create new one with image descriptions
                    rag_core.remove_collection(collection_id)
                    enhanced_collection_id = rag_core.create_collection(presentation_model)
                    
                    # Save temporary file for processing
                    temp_path = f"temp_{file_name}"
                    with open(temp_path, "wb") as f:
                        f.write(file_bytes)
                    
                    # Process with enhanced RAG
                    output_path = f"accessible_{file_name}"
                    total_images, processed_images, total_shapes, processed_shapes = process_powerpoint_with_rag_enhanced(
                        temp_path, output_path, presentation_model, enhanced_collection_id, {}
                    )
                    
                    # Clean up temp file
                    os.remove(temp_path)
                    
                    # Store output path for download
                    st.session_state.output_path = output_path
                    st.session_state.processing_stage = 'download'
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ Error in final processing: {str(e)}")
                    st.exception(e)
    
    # Stage 4: Download
    elif st.session_state.processing_stage == 'download':
        st.header("✅ Accessibility Enhancement Complete!")
        
        output_path = st.session_state.output_path
        
        # Show results
        col1, col2 = st.columns(2)
        with col1:
            st.success("♿ Accessibility features added:")
            st.write("• **Enhanced Alt Text**: AI-generated descriptions with context")
            st.write("• **Comprehensive Slide Notes**: RAG-enhanced content summaries")
            st.write("• **Chart/Table Labels**: Accessibility labels for data elements")
            st.write("• **Context-Aware Descriptions**: Using presentation-wide knowledge")
        
        with col2:
            st.info("📊 Processing Summary:")
            st.write("• **RAG Collection**: Built with text and image content")
            st.write("• **Image Processing**: Enhanced with contextual descriptions")
            st.write("• **Notes Generation**: Context-retrieved and AI-enhanced")
        
        # Download button
        if os.path.exists(output_path):
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Accessible PowerPoint",
                    data=f.read(),
                    file_name=output_path,
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                    type="primary",
                    use_container_width=True
                )
        
        # Reset button
        if st.button("🔄 Process Another Presentation", use_container_width=True):
            # Clean up
            if os.path.exists(output_path):
                os.remove(output_path)
            
            # Reset session state
            for key in ['processing_stage', 'presentation_model', 'rag_core', 'image_processor', 
                       'collection_id', 'current_batch', 'output_path']:
                if key in st.session_state:
                    del st.session_state[key]
            
            st.session_state.processing_stage = 'upload'
            st.rerun()

if __name__ == "__main__":
    main()
