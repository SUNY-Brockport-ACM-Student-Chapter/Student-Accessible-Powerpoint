import json
from pptx_rag_quizzer.rag_core import RAGCore
import streamlit as st
from typing import List, Dict, Any, Optional
import hashlib
import time


class Image:

    def __init__(self, rag_core: RAGCore):
        self.rag_core = rag_core
        self.chat_history = []
        self.context_cache = {}
        self.lambda_index = {}
        self.max_chat_history = 10
        self.cache_ttl = 3600  # 1 hour cache TTL

    def describe_image(self, image_bytes: bytes, image_format: str = "png", slide_number: int = 0, collection_id: str = None, use_chat: bool = True):
        """
        Describes an image using a multi-stage RAG pipeline with Lambda Index and chat optimization:
        1. OCR extraction
        2. Enhanced description using OCR + image + slide context
        3. Lambda Index context retrieval using enhanced description
        4. Final description with context and chat history

        Args:
            image_bytes (bytes): The image data
            image_format (str): The image format (png, jpg, etc.)
            slide_number (int): The slide number where the image is located
            collection_id (str): The collection ID to query for context
            use_chat (bool): Whether to use chat history for context

        Returns:
            str: Enhanced description of the image with context
        """
        # Generate image hash for caching
        image_hash = self._generate_image_hash(image_bytes)
        
        # Check cache first
        cache_key = f"{image_hash}_{slide_number}_{collection_id}"
        if cache_key in self.context_cache:
            cached_result = self.context_cache[cache_key]
            if time.time() - cached_result['timestamp'] < self.cache_ttl:
                return cached_result['description']

        # Stage 1: Get OCR description
        ocr_description = self.ocr_image(image_bytes)

        # Stage 2: Get enhanced description using OCR + image + slide context
        slide_context = None
        if collection_id:
            try:
                slide_data = self.rag_core.get_context_from_slide_number(slide_number, collection_id)
                slide_context = slide_data["documents"]
                
                # Ensure slide_context is a string
                if isinstance(slide_context, list):
                    slide_context = " ".join(slide_context)
                elif not isinstance(slide_context, str):
                    slide_context = str(slide_context) if slide_context else None
                    
            except Exception as e:
                print(f"Could not get slide context: {e}")

        enhanced_description = self.get_enhanced_description(
            ocr_description, image_bytes, image_format, slide_context, use_chat
        )

        # Stage 3: Get context using Lambda Index and enhanced description
        context = None
        if collection_id:
            context = self.get_context_with_lambda_index(enhanced_description, collection_id, image_hash)

        # Stage 4: Final description with image + enhanced description + context + chat history
        final_description = self.get_final_description_with_chat(
            enhanced_description, context, image_bytes, image_format, use_chat
        )

        # Handle JSON response if the model returns JSON instead of plain text
        if final_description and final_description.strip().startswith('{'):
            try:
                parsed = json.loads(final_description)
                if 'output' in parsed and 'Description' in parsed['output']:
                    final_description = parsed['output']['Description']
                elif 'Description' in parsed:
                    final_description = parsed['Description']
            except json.JSONDecodeError:
                # If JSON parsing fails, try to extract description from the text
                pass

        # Cache the result
        self.context_cache[cache_key] = {
            'description': final_description,
            'timestamp': time.time()
        }

        return final_description

    def _generate_image_hash(self, image_bytes: bytes) -> str:
        """Generate a hash for the image for caching purposes."""
        return hashlib.md5(image_bytes).hexdigest()

    def ocr_image(self, image_bytes: bytes):
        """
        Extract text from image using OCR.
        
        Args:
            image_bytes (bytes): The image data
            
        Returns:
            str: Extracted text from the image
        """
        try:
            from .utils import ExtractText_OCR
            return ExtractText_OCR(image_bytes)
        except ImportError:
            print("OCR functionality not available - pytesseract not installed")
            return "OCR text extraction not available"
        except Exception as e:
            print(f"Error during OCR extraction: {e}")
            return "OCR extraction failed"

    def get_enhanced_description(
        self,
        ocr_description: str,
        image_bytes: bytes,
        image_format: str,
        slide_context: str = None,
        use_chat: bool = True,
    ):
        """
        Stage 2: Get enhanced description using OCR + image + slide context with chat optimization.

        Args:
            ocr_description (str): The OCR text from the image
            image_bytes (bytes): The image data
            image_format (str): The image format
            slide_context (str): The slide context
            use_chat (bool): Whether to use chat history

        Returns:
            str: Enhanced description of the image
        """
        # Build the prompt with OCR and slide context
        context_info = ""
        if slide_context:
            context_info = f"\n\nSlide Context:\n{slide_context}"

        # Add chat history if enabled
        chat_context = ""
        if use_chat and self.chat_history:
            chat_context = "\n\nPrevious Context:\n" + "\n".join([f"- {msg}" for msg in self.chat_history[-3:]])

        prompt = f"""<prompt>
                        <instructions>
                            You are an expert visual analyst providing concise descriptions for a RAG system.
                            Your task is to describe an image from a PowerPoint slide.
                            Describe the image primarily based on its visual content.
                            The description must be 1 to 3 sentences long and focus on the most important visual elements and their core meaning.
                            Incorporate the provided OCR text or slide context only if it directly clarifies or adds significant understanding to the image's visual elements. Do not include context that merely repeats what's obvious in the image or is irrelevant.
                            IMPORTANT: Return ONLY the description text, not JSON or any other format.
                        </instructions>

                        <input_data>
                            <ocr_text>{ocr_description}</ocr_text>
                            <slide_context>{context_info}</slide_context>
                            <chat_history>{chat_context}</chat_history>
                        </input_data>

                        <output_format>
                            Return the description as plain text only, starting with "Description: "
                        </output_format>

                        <examples>
                            <example>
                            <input>
                                <ocr_text>Quarterly Revenue Trends</ocr_text>
                                <slide_context>This slide details the financial performance over the past year, highlighting growth in emerging markets.</slide_context>
                            </input>
                            <output>
                                Description: A line graph displays fluctuating quarterly revenue trends over a year, with a noticeable upward curve towards the end. The chart, titled "Quarterly Revenue Trends," visually represents the company's financial trajectory, which is further supported by the context of growth in emerging markets.
                            </output>
                            </example>
                        </examples>
                        </prompt>
                """

        enhanced_description = self.rag_core.prompt_gemini_with_image(
            prompt=prompt,
            image_bytes=image_bytes,
            image_format=image_format,
            max_output_tokens=200,
        )

        # Ensure enhanced_description is a string
        if isinstance(enhanced_description, list):
            enhanced_description = " ".join(enhanced_description)
        elif not isinstance(enhanced_description, str):
            enhanced_description = str(enhanced_description)

        # Add to chat history if enabled
        if use_chat:
            self._add_to_chat_history(f"Enhanced description: {enhanced_description}")

        return enhanced_description

    def get_context_with_lambda_index(self, enhanced_description: str, collection_id: str, image_hash: str, n_results: int = 3):
        """
        Stage 3: Get context using Lambda Index for better retrieval.

        Args:
            enhanced_description (str): The enhanced description of the image
            collection_id (str): The collection ID to query
            image_hash (str): Hash of the image for indexing
            n_results (int): Number of most relevant documents to retrieve

        Returns:
            str: Combined context from most relevant slides using Lambda Index
        """
        # Build Lambda Index query with image characteristics
        lambda_query = self._build_lambda_query(enhanced_description, image_hash)
        
        # Query the collection using the Lambda Index
        retrieved_results = self.rag_core.query_collection(
            query_text=lambda_query,
            collection_id=collection_id,
            n_results=n_results,
        )

        # Process and rank results using Lambda Index
        ranked_context = self._rank_context_with_lambda(retrieved_results, enhanced_description)

        return ranked_context

    def _build_lambda_query(self, enhanced_description: str, image_hash: str) -> str:
        """
        Build a Lambda Index query that considers image characteristics and description.
        
        Args:
            enhanced_description (str): The enhanced description
            image_hash (str): Hash of the image
            
        Returns:
            str: Lambda Index query
        """
        # Extract key terms from description
        key_terms = self._extract_key_terms(enhanced_description)
        
        # Build query with Lambda Index features
        lambda_query = f"""
        Image Analysis Query:
        - Description: {enhanced_description}
        - Key Terms: {', '.join(key_terms)}
        - Image Hash: {image_hash}
        - Context Type: visual_analysis
        """
        
        return lambda_query

    def _extract_key_terms(self, description: str) -> List[str]:
        """
        Extract key terms from description for Lambda Index.
        
        Args:
            description (str): The description text
            
        Returns:
            List[str]: List of key terms
        """
        # Simple key term extraction 
        stop_words = {'the', 'a', 'an', 'and', 'or', 'but', 'in', 'on', 'at', 'to', 'for', 'of', 'with', 'by'}
        words = description.lower().split()
        key_terms = [word for word in words if word not in stop_words and len(word) > 3]
        return key_terms[:10]  # Limit to top 10 terms

    def _rank_context_with_lambda(self, retrieved_results: Dict[str, Any], enhanced_description: str) -> str:
        """
        Rank context using Lambda Index features.
        
        Args:
            retrieved_results (Dict): Results from collection query
            enhanced_description (str): The enhanced description
            
        Returns:
            str: Ranked and combined context
        """
        if not retrieved_results or "documents" not in retrieved_results:
            return ""

        # Score each document based on Lambda Index features
        scored_docs = []
        for i, doc in enumerate(retrieved_results["documents"]):
            if doc:
                # Ensure doc is a string
                if isinstance(doc, list):
                    doc = " ".join(doc)
                elif not isinstance(doc, str):
                    doc = str(doc)
                
                # Ensure metadata is a dictionary
                metadata = retrieved_results["metadatas"][i] if i < len(retrieved_results["metadatas"]) else {}
                if not isinstance(metadata, dict):
                    metadata = {}
                
                score = self._calculate_lambda_score(doc, enhanced_description, metadata)
                scored_docs.append((score, doc))

        # Sort by score and combine
        scored_docs.sort(key=lambda x: x[0], reverse=True)
        context_parts = [doc for score, doc in scored_docs if score > 0.3]  # Threshold for relevance
        
        return "\n\n".join(context_parts)

    def _calculate_lambda_score(self, document: str, enhanced_description: str, metadata: Dict[str, Any]) -> float:
        """
        Calculate Lambda Index score for a document.
        
        Args:
            document (str): The document text
            enhanced_description (str): The enhanced description
            metadata (Dict): Document metadata
            
        Returns:
            float: Lambda Index score
        """
        score = 0.0
        
        # Ensure enhanced_description is a string
        if isinstance(enhanced_description, list):
            enhanced_description = " ".join(enhanced_description)
        elif not isinstance(enhanced_description, str):
            enhanced_description = str(enhanced_description)
        
        # Ensure document is a string
        if isinstance(document, list):
            document = " ".join(document)
        elif not isinstance(document, str):
            document = str(document)
        
        # Term overlap score
        desc_terms = set(enhanced_description.lower().split())
        doc_terms = set(document.lower().split())
        overlap = len(desc_terms.intersection(doc_terms))
        score += overlap * 0.1
        
        # Metadata relevance score
        if metadata and isinstance(metadata, dict):
            # Check if document has image-related metadata
            for key, value in metadata.items():
                if key.endswith("_type") and value == "image":
                    score += 0.3
                if key.endswith("_slide_number"):
                    score += 0.1
        
        # Length normalization
        score = score / (len(document.split()) + 1)
        
        return min(score, 1.0)  # Cap at 1.0

    def get_final_description_with_chat(
        self,
        enhanced_description: str,
        context: str,
        image_bytes: bytes,
        image_format: str,
        use_chat: bool = True,
    ):
        """
        Stage 4: Get final description with image + enhanced description + context + chat history.

        Args:
            enhanced_description (str): The enhanced description from stage 2
            context (str): The retrieved context from stage 3
            image_bytes (bytes): The image data
            image_format (str): The image format
            use_chat (bool): Whether to use chat history

        Returns:
            str: Final enhanced description with context and chat history
        """
        # Build chat context
        chat_context = ""
        if use_chat and self.chat_history:
            chat_context = "\n\nChat History:\n" + "\n".join([f"- {msg}" for msg in self.chat_history[-5:]])

        prompt = f"""<prompt>
                        <instructions>
                            You are a meticulous content refiner for a RAG system, focused on image descriptions.
                            Your task is to refine a given image description based on newly retrieved contextual information and chat history.
                            The refined description must be 1 to 3 sentences long.
                            The primary focus remains the image's visual content. Only incorporate the retrieved context if it provides new, crucial clarity or meaning that is not evident from the current description alone. Do not force context if it doesn't genuinely enhance the visual explanation.
                            Consider the chat history to maintain consistency with previous descriptions.
                            IMPORTANT: Return ONLY the description text, not JSON or any other format.
                        </instructions>

                        <input_data>
                            <current_description>{enhanced_description}</current_description>
                            <retrieved_context>{context}</retrieved_context>
                            <chat_history>{chat_context}</chat_history>
                        </input_data>

                        <output_format>
                            Return the description as plain text only, starting with "Description: "
                        </output_format>

                        <examples>
                            <example>
                            <input>
                                <current_description>A diagram shows interconnected boxes with arrows.</current_description>
                                <retrieved_context>This diagram illustrates the "Customer Journey Map," detailing touchpoints from awareness to loyalty. The slide's title is "Understanding Our Customer Funnel."</retrieved_context>
                            </input>
                            <output>
                                Description: A flowchart diagram depicts a multi-stage process with interconnected boxes and arrows, visually representing a customer journey map. The diagram, aligned with the concept of a customer funnel, illustrates sequential touchpoints from initial awareness through to loyalty.
                            </output>
                            </example>
                        </examples>
                        </prompt>
            """

        final_description = self.rag_core.prompt_gemini_with_image(
            prompt=prompt,
            image_bytes=image_bytes,
            image_format=image_format,
            max_output_tokens=200,
        )

        # Ensure final_description is a string
        if isinstance(final_description, list):
            final_description = " ".join(final_description)
        elif not isinstance(final_description, str):
            final_description = str(final_description)

        # Add to chat history if enabled
        if use_chat:
            self._add_to_chat_history(f"Final description: {final_description}")

        return final_description

    def _add_to_chat_history(self, message: str):
        """
        Add message to chat history with size management.
        
        Args:
            message (str): Message to add to chat history
        """
        self.chat_history.append(message)
        
        # Keep only the last N messages
        if len(self.chat_history) > self.max_chat_history:
            self.chat_history = self.chat_history[-self.max_chat_history:]

    def clear_chat_history(self):
        """Clear the chat history."""
        self.chat_history = []

    def get_chat_history(self) -> List[str]:
        """
        Get the current chat history.
        
        Returns:
            List[str]: List of chat messages
        """
        return self.chat_history.copy()

    def get_context_from_enhanced_description(
        self, enhanced_description: str, collection_id: str, n_results: int = 3
    ):
        """
        Legacy method for backward compatibility.
        Use get_context_with_lambda_index instead.
        """
        return self.get_context_with_lambda_index(enhanced_description, collection_id, "", n_results)

    def get_final_description(
        self,
        enhanced_description: str,
        context: str,
        image_bytes: bytes,
        image_format: str,
    ):
        """
        Legacy method for backward compatibility.
        Use get_final_description_with_chat instead.
        """
        return self.get_final_description_with_chat(enhanced_description, context, image_bytes, image_format, use_chat=False)

    def get_image_info(self, image_bytes: bytes, image_extension: str = None):
        """
        Get basic information about an image.
        
        Args:
            image_bytes (bytes): The image data
            image_extension (str): File extension like 'png', 'jpg'
            
        Returns:
            dict: Image information
        """
        try:
            return {
                "size": len(image_bytes),
                "extension": image_extension,
                "hash": self._generate_image_hash(image_bytes)
            }
        except Exception as e:
            print(f"Error getting image info: {e}")
            return None

    def get_lambda_index_stats(self) -> Dict[str, Any]:
        """
        Get statistics about the Lambda Index usage.
        
        Returns:
            Dict[str, Any]: Lambda Index statistics
        """
        return {
            "cache_size": len(self.context_cache),
            "chat_history_size": len(self.chat_history),
            "lambda_index_size": len(self.lambda_index),
            "cache_hits": sum(1 for cache in self.context_cache.values() if time.time() - cache['timestamp'] < self.cache_ttl)
        }

    def clear_cache(self):
        """Clear the context cache."""
        self.context_cache = {}

    def set_cache_ttl(self, ttl_seconds: int):
        """
        Set the cache TTL (Time To Live).
        
        Args:
            ttl_seconds (int): Cache TTL in seconds
        """
        self.cache_ttl = ttl_seconds

    def set_max_chat_history(self, max_history: int):
        """
        Set the maximum chat history size.
        
        Args:
            max_history (int): Maximum number of chat messages to keep
        """
        self.max_chat_history = max_history 
        
        