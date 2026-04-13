# RAG-Enhanced PowerPoint Accessibility App

This document describes the enhanced PowerPoint accessibility application that now includes RAG (Retrieval-Augmented Generation) functionality for improved image descriptions and context-aware accessibility features.

## 🚀 New Features

### Enhanced Workflow
1. **PowerPoint Upload & Initial Processing**
   - Parse PowerPoint file into structured model
   - Build initial RAG collection from text content
   - Initialize image processing pipeline

2. **Intelligent Image Processing**
   - Batch processing of images (5 images per batch)
   - AI-generated descriptions using contextual information
   - User review and approval workflow
   - Edit descriptions before proceeding

3. **Enhanced RAG Collection**
   - Rebuild collection with both text and image descriptions
   - Context-aware retrieval for better accessibility notes
   - Lambda Index for improved relevance scoring

4. **Context-Aware Accessibility Features**
   - Enhanced slide notes using presentation-wide context
   - Improved alt text generation
   - Comprehensive accessibility descriptions

## 📋 Prerequisites

### Required Dependencies
```bash
pip install -r requirements.txt
```

### Additional Requirements
- **Google API Key**: Set `GOOGLE_API_KEY` in your `.env` file
- **ChromaDB API**: Running on `localhost:8001` (see deployment instructions)
- **Tesseract OCR**: For text extraction from images

### Environment Variables
Create a `.env` file with:
```env
GOOGLE_API_KEY=your_google_api_key_here
CHROMA_SERVER_HOST=localhost
CHROMA_SERVER_HTTP_PORT=8001
```

## 🏗️ Architecture

### Core Components

1. **RAG Core** (`app/pptx_rag_quizzer/rag_core.py`)
   - Manages ChromaDB collections
   - Handles LLM interactions
   - Provides context retrieval

2. **Image Processor** (`app/pptx_rag_quizzer/image.py`)
   - Multi-stage image analysis
   - OCR text extraction
   - Context-aware descriptions
   - Lambda Index integration

3. **PowerPoint Parser** (`app/pptx_rag_quizzer/utils.py`)
   - Extracts text and images from PowerPoint
   - Maintains slide order and structure
   - Creates structured data models

4. **Enhanced Processing** (`app/ppt_notes.py`)
   - Multi-stage workflow management
   - Batch processing interface
   - RAG-enhanced accessibility features

### Data Flow

```
PowerPoint File
       ↓
   Parse & Extract
       ↓
  Build Text RAG
       ↓
  Process Images
       ↓
   User Approval
       ↓
  Enhanced RAG
       ↓
Generate Notes
       ↓
  Download File
```

## 🎯 Usage Instructions

### Starting the Application
```bash
streamlit run app/ppt_notes.py
```

### Step-by-Step Process

#### Step 1: Upload PowerPoint
1. Upload a `.pptx` file
2. System automatically parses content and builds initial RAG collection
3. Progress to image processing stage

#### Step 2: Image Description & Approval
1. **Automatic Processing**: AI generates descriptions for images in batches
2. **Review Interface**: Review and edit descriptions for each image
3. **Batch Navigation**: Move through batches using Previous/Next buttons
4. **Save Progress**: Save descriptions before moving to next batch
5. **Finish**: Complete when all images are processed

#### Step 3: Final Processing
1. System rebuilds RAG collection with image descriptions
2. Generates enhanced accessibility features
3. Creates context-aware slide notes
4. Adds comprehensive alt text

#### Step 4: Download
1. Download the enhanced PowerPoint file
2. All accessibility features are embedded
3. Option to process another presentation

## 🔧 Configuration Options

### Batch Size
Default batch size is 5 images. Can be modified in session state:
```python
st.session_state.batch_size = 5  # Adjust as needed
```

### RAG Parameters
- **Collection ID**: Auto-generated UUID
- **Context Retrieval**: Top 3 most relevant documents
- **Lambda Index**: Enhanced relevance scoring

### Image Processing
- **OCR Integration**: Automatic text extraction
- **Context Awareness**: Uses slide and presentation context
- **Chat History**: Maintains conversation context for consistency

## 🚨 Troubleshooting

### Common Issues

1. **ChromaDB API Connection Failed**
   - Ensure ChromaDB API is running on `localhost:8001`
   - Check environment variables
   - Verify API health endpoint

2. **Google API Key Issues**
   - Verify API key is set in `.env` file
   - Check API quota and limits
   - Ensure proper authentication

3. **Image Processing Errors**
   - Check Tesseract installation
   - Verify image formats are supported
   - Review error messages in console

4. **Memory Issues with Large Presentations**
   - Reduce batch size
   - Process images in smaller groups
   - Monitor system resources

### Debug Mode
Enable debug logging by setting:
```python
import logging
logging.basicConfig(level=logging.DEBUG)
```

## 📊 Performance Considerations

### Optimization Tips
1. **Batch Processing**: Process images in manageable batches
2. **Caching**: RAG results are cached for efficiency
3. **Memory Management**: Large presentations may require more RAM
4. **API Limits**: Monitor Google API usage and quotas

### Scalability
- **Small Presentations** (< 10 slides): ~2-3 minutes
- **Medium Presentations** (10-25 slides): ~5-10 minutes
- **Large Presentations** (25+ slides): ~15-30 minutes

## 🔄 Integration with Existing Systems

### Teacher Portal Integration
The enhanced app can be integrated with the existing teacher portal (`1_Teacher_Portal.py`) by:

1. **Shared Components**: Use the same RAG core and image processor
2. **Database Integration**: Store processed presentations in database
3. **User Management**: Integrate with existing user authentication
4. **Assignment Creation**: Use enhanced presentations for homework generation

### API Integration
The ChromaDB API (`app/chroma-api/app.py`) provides:
- RESTful endpoints for collection management
- Document storage and retrieval
- Query processing with metadata

## 📈 Future Enhancements

### Planned Features
1. **Multi-language Support**: OCR and descriptions in multiple languages
2. **Advanced Analytics**: Usage statistics and performance metrics
3. **Custom Models**: Fine-tuned models for specific domains
4. **Batch Upload**: Process multiple presentations simultaneously
5. **Integration APIs**: Connect with LMS platforms

### Extensibility
The modular architecture allows for:
- Custom image processors
- Additional LLM providers
- Enhanced context retrieval methods
- Custom accessibility standards

## 🤝 Contributing

### Development Setup
1. Clone the repository
2. Install dependencies: `pip install -r requirements.txt`
3. Set up environment variables (Google API key)
4. **Easy Start**: Run `python start_app.py` and choose option 3
5. **Manual Start**: 
   - Start ChromaDB API: `cd app/chroma-api && python app.py`
   - Start PowerPoint app: `streamlit run app/ppt_notes.py`

### Code Structure
- **Models**: Data structures in `app/models.py`
- **Core Logic**: RAG functionality in `app/pptx_rag_quizzer/`
- **UI**: Streamlit interface in `app/ppt_notes.py`
- **API**: ChromaDB service in `app/chroma-api/`

## 📄 License

This project is part of the Student Accessible PowerPoint initiative and follows the same licensing terms.

---

**Note**: This enhanced version significantly improves the accessibility features by leveraging RAG technology for context-aware processing and user-controlled image description workflows.
