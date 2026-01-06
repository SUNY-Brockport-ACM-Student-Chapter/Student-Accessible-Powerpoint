import streamlit as st
import google.generativeai as genai
from google.generativeai.types import GenerationConfig
import pytesseract
import subprocess
import shutil


def ExtractText_OCR(img_bytes):
    """
    Extracts text from an image using OCR (Tesseract).

    Args:
        img_bytes (bytes): The image data in bytes.

    Returns:
        str: The extracted text from the image.
    """
    # try:
    #     # Extract text using OCR (Tesseract)
    #     img = Image.open(io.BytesIO(img_bytes))
    #     text = pytesseract.image_to_string(img)
    #     return text.strip()

    # except Exception as e:
    #     print(f"Error during OCR extraction: {e}")
    #     return ""
    return "<THIS OCR TEXT IS IN DEVELOPMENT AND SHOULD BE DISREGARDED>"


def clean_text(text):
    """
    Cleans the text by removing any non-essential information.
    """
    return "\n".join(line for line in text.splitlines() if line.strip())


def clean_text_with_llm(text, model):
    """
    Cleans the text by removing any non-essential information using LLM (Gemini-2.0-flash-lite).
    """
    generation_config = GenerationConfig(max_output_tokens=100)

    result = model.generate_content(
        contents=[
            text,
            "\n",
            "given the following text, remove any non-essential information and return the text in a clean format. "
            "Only return the text in a clean format. Nothing else!",
        ],
        generation_config=generation_config,
    )
    return result.text.strip()


def convert_image_to_png_or_jpg(image_bytes, extension):
    """
    Convert arbitrary image bytes to PNG (preferred) or JPG using ImageMagick if available.

    Args:
        image_bytes (bytes): Source image bytes
        extension (str): Original file extension (e.g., 'png', 'jpg', 'svg', ...)

    Returns:
        (bytes, str) or (None, None): Tuple of (converted_bytes, new_extension), or (None, None) if conversion fails for WMF/EMF
    """
    # Normalize extension
    ext = (extension or "").lower().lstrip(".")

    # If already web-safe, return as-is
    if ext in ("jpg", "jpeg"):
        return image_bytes, "jpg"
    if ext == "png":
        return image_bytes, "png"

    # Prefer PNG as the unified output
    magick = shutil.which("magick") or shutil.which("convert")
    if not magick:
        # ImageMagick not available; skip problematic formats
        if ext in ("wmf", "emf"):
            print(f"❌ WARNING: ImageMagick not found, cannot convert {ext.upper()} image. Skipping...")
            return None, None
        return image_bytes, (ext or "png")

    # Higher density for vector-like formats to improve rasterization quality
    vector_like_exts = {"svg", "pdf", "eps", "ai", "emf", "wmf"}
    density_args = ["-density", "300"] if ext in vector_like_exts else []

    try:
        # Use stdin/stdout to avoid relying on correct file extensions
        # Command: magick [-density 300] - -colorspace sRGB png:-
        cmd = [magick]
        cmd += density_args
        cmd += ["-", "-colorspace", "sRGB", "png:-"]

        proc = subprocess.run(
            cmd,
            input=image_bytes,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
        )
        
        # Verify output is valid
        if len(proc.stdout) == 0:
            raise Exception(f"ImageMagick returned empty output for {ext}")
            
        return proc.stdout, "png"
        
    except subprocess.CalledProcessError as e:
        # Log the actual error for debugging
        stderr_output = e.stderr.decode('utf-8', errors='ignore') if e.stderr else 'N/A'
        print(f"❌ ERROR: ImageMagick conversion failed for {ext.upper()}")
        print(f"   Command: {' '.join(cmd)}")
        print(f"   Return code: {e.returncode}")
        print(f"   stderr: {stderr_output[:300]}")  # First 300 chars
        
        # For formats that REQUIRE conversion (WMF, EMF), skip the image
        if ext in ("wmf", "emf"):
            print(f"   → Skipping {ext.upper()} image (cannot be used without conversion)")
            return None, None
        
        # For other formats, try to return original (risky but might work)
        print(f"   → Falling back to original {ext} bytes")
        return image_bytes, (ext or "png")
    except Exception as e:
        print(f"❌ ERROR: Unexpected error converting {ext.upper()}: {str(e)}")
        if ext in ("wmf", "emf"):
            print(f"   → Skipping {ext.upper()} image")
            return None, None
        return image_bytes, (ext or "png")