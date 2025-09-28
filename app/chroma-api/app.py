"""
ChromaDB API Service
Standalone ChromaDB API service that runs on its own process.
This service handles all ChromaDB operations and exposes a REST API.
"""

import os
import sys
import json
import uuid
import numpy as np
from typing import Dict, List, Optional
from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from pydantic import BaseModel
import chromadb
from dotenv import load_dotenv

load_dotenv()

def make_json_serializable(obj):
    """
    Convert numpy arrays and other non-serializable objects to JSON-serializable formats.
    """
    if isinstance(obj, np.ndarray):
        return obj.tolist()
    elif isinstance(obj, (np.integer, np.floating)):
        return obj.item()
    elif isinstance(obj, dict):
        return {key: make_json_serializable(value) for key, value in obj.items()}
    elif isinstance(obj, (list, tuple)):
        return [make_json_serializable(item) for item in obj]
    else:
        return obj

# Initialize FastAPI app
app = FastAPI(
    title="ChromaDB API Service",
    description="REST API for ChromaDB operations",
    version="1.0.0"
)

# Add CORS middleware
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # Configure appropriately for production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global ChromaDB client
chroma_client = None

def get_chroma_client():
    """Get or create ChromaDB client."""
    global chroma_client
    if chroma_client is None:
        try:
            chroma_client = chromadb.HttpClient(
                host=os.getenv("CHROMA_SERVER_HOST", "localhost"),
                port=int(os.getenv("CHROMA_SERVER_HTTP_PORT", "8000"))
            )
            print("✅ ChromaDB client initialized")
        except Exception as e:
            print(f"❌ Failed to initialize ChromaDB client: {e}")
            raise HTTPException(status_code=500, detail="ChromaDB client initialization failed")
    return chroma_client

# Pydantic models for request/response
class CreateCollectionRequest(BaseModel):
    name: str
    metadata: Optional[Dict] = {}

class AddDocumentsRequest(BaseModel):
    documents: List[str]
    metadatas: List[Dict]
    ids: List[str]

class QueryRequest(BaseModel):
    query_texts: List[str]
    n_results: int = 1
    include: Optional[List[str]] = ["documents", "metadatas", "embeddings"]

class GetCollectionRequest(BaseModel):
    include: Optional[List[str]] = ["documents", "metadatas", "embeddings"]

# Health check endpoint
@app.get("/health")
async def health_check():
    """Health check endpoint."""
    try:
        client = get_chroma_client()
        client.list_collections()
        return {"status": "healthy", "message": "ChromaDB API service is running"}
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Service unhealthy: {str(e)}")

# Collection management endpoints
@app.post("/collections")
async def create_collection(request: CreateCollectionRequest):
    """Create a new collection."""
    try:
        client = get_chroma_client()
        collection = client.create_collection(
            name=request.name,
            metadata=request.metadata
        )
        return {
            "success": True,
            "message": f"Collection '{request.name}' created successfully",
            "collection": {
                "name": collection.name,
                "id": collection.id,
                "metadata": collection.metadata
            }
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to create collection: {str(e)}")

@app.delete("/collections/{collection_name}")
async def delete_collection(collection_name: str):
    """Delete a collection."""
    try:
        client = get_chroma_client()
        client.delete_collection(name=collection_name)
        return {
            "success": True,
            "message": f"Collection '{collection_name}' deleted successfully"
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to delete collection: {str(e)}")

@app.get("/collections")
async def list_collections():
    """List all collections."""
    try:
        client = get_chroma_client()
        collections = client.list_collections()
        return {
            "success": True,
            "collections": [
                {
                    "name": col.name,
                    "id": col.id,
                    "metadata": col.metadata
                }
                for col in collections
            ]
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to list collections: {str(e)}")

@app.get("/collections/{collection_name}/exists")
async def collection_exists(collection_name: str):
    """Check if a collection exists."""
    try:
        client = get_chroma_client()
        collections = client.list_collections()
        collection_names = [col.name for col in collections]
        exists = collection_name in collection_names
        return {
            "success": True,
            "exists": exists,
            "collection_name": collection_name
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to check collection existence: {str(e)}")

# Document operations
@app.post("/collections/{collection_name}/add")
async def add_documents(collection_name: str, request: AddDocumentsRequest):
    """Add documents to a collection."""
    try:
        client = get_chroma_client()
        collection = client.get_collection(name=collection_name)
        collection.add(
            documents=request.documents,
            metadatas=request.metadatas,
            ids=request.ids
        )
        return {
            "success": True,
            "message": f"Added {len(request.documents)} documents to collection '{collection_name}'"
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to add documents: {str(e)}")

@app.post("/collections/{collection_name}/query")
async def query_collection(collection_name: str, request: QueryRequest):
    """Query a collection for similar documents."""
    try:
        client = get_chroma_client()
        collection = client.get_collection(name=collection_name)
        results = collection.query(
            query_texts=request.query_texts,
            n_results=request.n_results,
            include=request.include
        )
        # Convert numpy arrays and other non-serializable objects to JSON-serializable formats
        serializable_results = make_json_serializable(results)
        return {
            "success": True,
            "results": serializable_results
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to query collection: {str(e)}")

@app.post("/collections/{collection_name}/get")
async def get_collection_data(collection_name: str, request: GetCollectionRequest):
    """Get all documents from a collection."""
    try:
        client = get_chroma_client()
        collection = client.get_collection(name=collection_name)
        data = collection.get(include=request.include)
        # Convert numpy arrays and other non-serializable objects to JSON-serializable formats
        serializable_data = make_json_serializable(data)
        return {
            "success": True,
            "data": serializable_data
        }
    except Exception as e:
        raise HTTPException(status_code=400, detail=f"Failed to get collection data: {str(e)}")

# Root endpoint
@app.get("/")
async def root():
    """Root endpoint with API information."""
    return {
        "service": "ChromaDB API Service",
        "version": "1.0.0",
        "status": "running",
        "endpoints": {
            "health": "/health",
            "collections": "/collections",
            "create_collection": "POST /collections",
            "delete_collection": "DELETE /collections/{name}",
            "add_documents": "POST /collections/{name}/add",
            "query": "POST /collections/{name}/query",
            "get_data": "POST /collections/{name}/get"
        }
    }

if __name__ == "__main__":
    import uvicorn
    
    # Get configuration from environment
    host = os.getenv("API_HOST", "0.0.0.0")
    port = int(os.getenv("API_PORT", "8001"))
    
    print(f"🚀 Starting ChromaDB API Service on {host}:{port}")
    print(f"📊 ChromaDB Server: {os.getenv('CHROMA_SERVER_HOST', 'localhost')}:{os.getenv('CHROMA_SERVER_HTTP_PORT', '8000')}")
    
    uvicorn.run(app, host=host, port=port)
