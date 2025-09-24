import uuid
import random
import os
from dotenv import load_dotenv
import google.generativeai as genai
from google.generativeai.types import GenerationConfig
from models.models import Presentation, Type
import io
import time
from PIL import Image as PILImage
import requests
import json

load_dotenv()

_llm_model_cache = None

class ChromaHTTPClient:
    """Simple HTTP client for ChromaDB API"""
    
    def __init__(self, api_url="http://localhost:8001"):
        self.api_url = api_url.rstrip('/')
    
    def create_collection(self, collection_id):
        """Create a new collection"""
        try:
            # First check if collection exists
            response = requests.get(f"{self.api_url}/collections/{collection_id}/exists")
            if response.status_code == 200:
                exists_data = response.json()
                if exists_data.get("exists", False):
                    self.delete_collection(collection_id)
            
            request_data = {"name": collection_id, "metadata": {"description": "PowerPoint presentation collection", "type": "presentation"}}
            
            response = requests.post(
                f"{self.api_url}/collections",
                json=request_data
            )
            
            response.raise_for_status()
            return response.json()
        except Exception as e:
            raise
    
    def add_documents(self, collection_id, documents, metadatas, ids):
        """Add documents to collection"""
        try:
            response = requests.post(
                f"{self.api_url}/collections/{collection_id}/add",
                json={
                    "documents": documents,
                    "metadatas": metadatas,
                    "ids": ids
                }
            )
            response.raise_for_status()
            return response.json()
        except Exception as e:
            raise
    
    def query_collection(self, collection_id, query_texts, n_results, include):
        """Query collection"""
        response = requests.post(
            f"{self.api_url}/collections/{collection_id}/query",
            json={
                "query_texts": query_texts,
                "n_results": n_results,
                "include": include
            }
        )
        response.raise_for_status()
        return response.json()["results"]
    
    def get_collection_data(self, collection_id, include):
        """Get all data from collection"""
        response = requests.post(
            f"{self.api_url}/collections/{collection_id}/get",
            json={"include": include}
        )
        response.raise_for_status()
        return response.json()["data"]
    
    def delete_collection(self, collection_id):
        """Delete collection"""
        response = requests.delete(f"{self.api_url}/collections/{collection_id}")
        response.raise_for_status()
        return response.json()

def get_chroma_http_client_instance(api_url=None):
    """Get ChromaDB HTTP client instance"""
    if api_url is None:
        api_url = "http://localhost:8001"
    return ChromaHTTPClient(api_url)

def get_llm_model():
    """
    Configures the Google Generative AI model using the GOOGLE_API_KEY environment variable.

    Returns:
        genai.GenerativeModel: The loaded LLM model.
    """
    global _llm_model_cache

    if _llm_model_cache is None:
        try:
            load_dotenv()
            api_key = os.getenv("GOOGLE_API_KEY")
            if not api_key:
                print(
                    "GOOGLE_API_KEY environment variable not found. Please set it in your .env file."
                )
                return False

            genai.configure(api_key=api_key)
            _llm_model_cache = genai.GenerativeModel("gemini-2.0-flash-lite")
        except Exception as e:
            print(f"Error loading LLM model: {e}")
            return False
    else:
        pass

    return _llm_model_cache


class RAGCore:
    """Handles the core Retrieval-Augmented Generation pipeline."""

    def __init__(self, chroma_api_url=None):
        """
        Initializes the RAGCore.
        
        Args:
            chroma_api_url: ChromaDB API service URL (for production environments)
        """
        self.llm_model = get_llm_model()
        self.chroma_api_url = chroma_api_url
        
        # Always use HTTP API client (no ChromaDB dependencies)
        self.chroma_api = get_chroma_http_client_instance(api_url=chroma_api_url)

    def create_collection(self, data: Presentation):
        """
        Builds the vector database from the Presentation object.

        Returns:
            str: The collection id.
        """

        all_texts = []
        all_ids = []
        all_metadatas = []

        for slide in data.slides:

            all_slide_texts = []
            all_slide_metadatas = []
            all_slide_items = []  # Keep track of the actual items
            
            for item in slide.items:
                if item.type == Type.text:
                    all_slide_texts.append(item.content)
                    all_slide_metadatas.append(item.metadata())
                    all_slide_items.append(item)
                elif item.type == Type.image:
                    # Skip deleted images (marked with "__DELETED__" content)
                    if item.content != "__DELETED__":
                        all_slide_texts.append(item.content)
                        all_slide_metadatas.append(item.metadata())
                        all_slide_items.append(item)

            chunk_id = str(uuid.uuid4())
            all_texts.append(" ".join(all_slide_texts))
            all_ids.append(chunk_id)
            
            # Combine all metadata into a single dictionary for this slide
            combined_metadata = {}
            for i, (metadata, item) in enumerate(zip(all_slide_metadatas, all_slide_items)):
                item_num = i + 1
                # Ensure all metadata values are JSON-serializable for ChromaDB
                combined_metadata[f"item_{item_num}_type"] = str(metadata["type"])
                combined_metadata[f"item_{item_num}_slide_number"] = int(metadata["slide_number"])
                combined_metadata[f"item_{item_num}_order_number"] = int(metadata["order_number"])
                
                # Add additional fields for images
                if metadata["type"] == "image":
                    combined_metadata[f"item_{item_num}_image_extension"] = str(metadata["extension"])
                    # Store image reference info without database
                    if hasattr(item, 'image_bytes') and hasattr(item, 'extension'):
                        combined_metadata[f"item_{item_num}_has_image"] = True
                        combined_metadata[f"item_{item_num}_image_size"] = len(item.image_bytes)
                    else:
                        print(f"❌ Image item missing required attributes: {type(item)}")
                        combined_metadata[f"item_{item_num}_has_image"] = False

            combined_metadata["slide_number"] = int(slide.slide_number)
            combined_metadata["slide_id"] = str(slide.id)

            all_metadatas.append(combined_metadata)


        if not all_texts:
            raise ValueError("No text content available to build the knowledge base.")
        
        collection_id = f"ppt_collection_{uuid.uuid4().hex[:8].lower()}"

        # Use HTTP API client (no ChromaDB dependencies)
        self.chroma_api.create_collection(collection_id)
        self.chroma_api.add_documents(collection_id, all_texts, all_metadatas, all_ids)   

        return collection_id

    def remove_collection(self, collection_id: str):
        """
        This function is used to remove a collection.
        """
        # Use HTTP API client (no ChromaDB dependencies)
        self.chroma_api.delete_collection(collection_id)


    def query_collection(self, query_text: str, collection_id: str, n_results: int = 1):
        """
        This function is used to get the context of collection.
        """
        # Use HTTP API client (no ChromaDB dependencies)
        retrieved_results = self.chroma_api.query_collection(
            collection_id, 
            [query_text], 
            n_results, 
            ["documents", "metadatas", "embeddings"]
        )

        return retrieved_results
    

    
    def get_random_slide_context(self, collection_id: str):
        """
        This function is used to get the context of a random slide.

        returns:
            dict: The context of a random slide.

        """
        # Use HTTP API client (no ChromaDB dependencies)
        collection_data = self.chroma_api.get_collection_data(collection_id, ["documents", "metadatas"])
        
        if collection_data is None or not collection_data:
            raise ValueError(f"Collection data is None or empty for collection_id: {collection_id}")
        
        random_index = random.randint(0, len(collection_data["ids"]) - 1)
        
        # Ensure we get a single document string
        document = collection_data["documents"][random_index]
        if isinstance(document, list):
            # If it's a list of characters, join them
            document = "".join(document)
        elif not isinstance(document, str):
            # Convert to string if it's not already
            document = str(document)
        
        # Create the result structure
        result = {
            "ids": [collection_data["ids"][random_index]],
            "documents": [document],
            "metadatas": [collection_data["metadatas"][random_index]],
        }
        
        return result
    
    def get_random_slide_with_image(self, collection_id: str):
        """
        This function gets the context of a random image document
        from a Chroma collection, retrying if necessary.
        """
        max_attempts = 10
        attempts = 0

        try:
            # Use HTTP API client (no ChromaDB dependencies)
            data = self.chroma_api.get_collection_data(collection_id, ["documents", "metadatas"])

            if not data["documents"]:
                print("No documents found in the collection.")
                return None

            while attempts < max_attempts:
                idx = random.randint(0, len(data["documents"]) - 1)
                metadata = data["metadatas"][idx]

                # Check if any key ends with '_type' and its value is 'image'
                if any(k.endswith("_type") and metadata[k] == "image" for k in metadata):
                    # Ensure we get a single document string
                    document = data["documents"][idx]
                    if isinstance(document, list):
                        # If it's a list of characters, join them
                        document = "".join(document)
                    elif not isinstance(document, str):
                        # Convert to string if it's not already
                        document = str(document)
                    
                    return {
                        "metadatas": metadata,
                        "documents": document,
                        "ids": data["ids"][idx]
                    }

                attempts += 1
                

        except Exception as e:
            print(f"Error getting random slide with image: {e}")

        print("Failed to find a random image after max attempts.")
        return None
    
    def get_context_from_slide_number(self, slide_number: int, collection_id: str):
        """
        This function is used to get the context of a slide by slide number.
        """
        # Use HTTP API client (no ChromaDB dependencies)
        collection_data = self.chroma_api.get_collection_data(collection_id, ["documents", "metadatas"])

        for idx, metadata in enumerate(collection_data["metadatas"]):
            if metadata["slide_number"] == slide_number:
                # Ensure we get a single document string
                document = collection_data["documents"][idx]
                if isinstance(document, list):
                    # If it's a list of characters, join them
                    document = "".join(document)
                elif not isinstance(document, str):
                    # Convert to string if it's not already
                    document = str(document)
                
                return {
                    "metadatas": collection_data["metadatas"][idx],
                    "documents": document,
                    "ids": collection_data["ids"][idx]
                }

        raise ValueError(f"No slide with number {slide_number} found")


    def prompt_gemini(self, prompt: str, max_output_tokens: int = 200):
        """
        This function is used to prompt the Gemini model.
        It handles quota exhaustion and retries.

        Args:
            prompt (str): The prompt to use for the Gemini model.
            max_output_tokens (int): The maximum number of tokens to output.

        Returns:
            str: The response from the Gemini model.
        """
        max_retries = 3
        delay = 1
        quota_refill_delay = 60
        generation_config = GenerationConfig(max_output_tokens=max_output_tokens)
        for attempt in range(max_retries):
            try:
                response = self.llm_model.generate_content(
                    contents=[prompt], generation_config=generation_config
                )
                return response.text
            except Exception as e:
                if "Resource has been exhausted" in str(e):
                    print(
                        f"Quota exhausted, waiting {quota_refill_delay} seconds for refill..."
                    )
                    time.sleep(quota_refill_delay)
                else:
                    print(f"Attempt {attempt + 1} failed: {str(e)}")
                if attempt < max_retries - 1:
                    time.sleep(delay)
                else:
                    raise

    def prompt_gemini_with_image(
        self,
        prompt: str,
        image_bytes: bytes,
        image_format: str = "png",
        max_output_tokens: int = 200,
    ):
        """
        This function is used to prompt the Gemini model with an image.
        It handles quota exhaustion and retries.

        Args:
            prompt (str): The prompt to use for the Gemini model.
            image_bytes (bytes): The image to use for the Gemini model.
            image_format (str): The format of the image.
            max_output_tokens (int): The maximum number of tokens to output.

        Returns:
            str: The response from the Gemini model.
        """
        max_retries = 3
        delay = 1
        quota_refill_delay = 60
        generation_config = GenerationConfig(max_output_tokens=max_output_tokens)

        # Validate image format and convert if necessary
        try:

            # Open and validate the image
            img = PILImage.open(io.BytesIO(image_bytes))

            # Convert to RGB if necessary (some formats like PNG with transparency cause issues)
            if img.mode in ("RGBA", "LA", "P"):
                img = img.convert("RGB")

            # Save as PNG to ensure compatibility
            img_buffer = io.BytesIO()
            img.save(img_buffer, format="PNG")
            img_buffer.seek(0)
            validated_image_bytes = img_buffer.getvalue()
            validated_format = "png"

        except Exception as e:
            print(f"Error validating image: {e}")
            # Use original image if validation fails
            validated_image_bytes = image_bytes
            validated_format = image_format

        for attempt in range(max_retries):
            try:
                image_part = {
                    "inline_data": {
                        "mime_type": f"image/{validated_format}",
                        "data": validated_image_bytes,
                    }
                }
                response = self.llm_model.generate_content(
                    contents=[image_part, "\n", prompt],
                    generation_config=generation_config,
                )
                return response.text
            except Exception as e:
                if "Resource has been exhausted" in str(e):
                    print(
                        f"Quota exhausted, waiting {quota_refill_delay} seconds for refill..."
                    )
                    time.sleep(quota_refill_delay)
                else:
                    print(f"Attempt {attempt + 1} failed: {str(e)}")
                if attempt < max_retries - 1:
                    time.sleep(delay)
                else:
                    raise