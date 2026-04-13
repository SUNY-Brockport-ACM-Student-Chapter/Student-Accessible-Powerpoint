import axios from 'axios';

export interface ChromaCollection {
  name: string;
  id: string;
  metadata?: Record<string, any>;
}

/**
 * Client for the ChromaDB API service.
 */
export class ChromaService {
  private baseUrl: string;

  constructor(baseUrl = 'http://localhost:8001') {
    this.baseUrl = baseUrl;
  }

  async health() {
    return axios.get(`${this.baseUrl}/health`);
  }

  async createCollection(name: string): Promise<ChromaCollection> {
    const res = await axios.post(`${this.baseUrl}/collections`, { name });
    return res.data.collection;
  }

  async addDocuments(collectionName: string, documents: string[], ids: string[], metadatas: any[]) {
    return axios.post(`${this.baseUrl}/collections/${collectionName}/add`, {
      documents,
      ids,
      metadatas
    });
  }

  async createFromPresentation(data: any): Promise<string> {
    const collectionId = `ppt_${Math.random().toString(36).substring(2, 10)}`;
    await this.createCollection(collectionId);

    const documents: string[] = [];
    const ids: string[] = [];
    const metadatas: any[] = [];

    data.slides.forEach((slide: any) => {
      const slideTexts = slide.items
        .filter((i: any) => i.type === 'text')
        .map((i: any) => i.content);
      
      const combinedText = slideTexts.join(' ');
      if (combinedText.trim()) {
        documents.push(combinedText);
        ids.push(`slide-${slide.number}`);
        metadatas.push({
          slide_number: slide.number,
          title: slide.title
        });
      }
    });

    if (documents.length > 0) {
      await this.addDocuments(collectionId, documents, ids, metadatas);
    }

    return collectionId;
  }
}
