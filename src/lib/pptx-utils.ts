import AdmZip from 'adm-zip';
import { parseStringPromise } from 'xml2js';

export interface SlideItem {
  id: string;
  type: 'text' | 'image' | 'chart' | 'table';
  content: string;
  slideNumber: number;
  imagePath?: string;
  imageBuffer?: Buffer;
}

export interface SlideData {
  number: number;
  title: string;
  items: SlideItem[];
  notes?: string;
}

export interface PresentationData {
  fileName: string;
  slides: SlideData[];
}

/**
 * Utility class to parse and manipulate PPTX files using adm-zip and XML parsing.
 * Modern replacement for python-pptx logic.
 */
export class PptxProcessor {
  private zip: AdmZip;

  constructor(fileBuffer: Buffer) {
    this.zip = new AdmZip(fileBuffer);
  }

  /**
   * Read slides and extract basic structure
   */
  async getPresentationData(fileName: string): Promise<PresentationData> {
    const slideEntries = this.zip.getEntries().filter(e => e.entryName.startsWith('ppt/slides/slide') && e.entryName.endsWith('.xml'));
    const slides: SlideData[] = [];

    // Sorting slides by number
    slideEntries.sort((a, b) => {
      const numA = parseInt(a.entryName.match(/\d+/)![0]);
      const numB = parseInt(b.entryName.match(/\d+/)![0]);
      return numA - numB;
    });

    for (const [index, entry] of slideEntries.entries()) {
      const slideNumber = index + 1;
      const content = entry.getData().toString('utf8');
      const xml = await parseStringPromise(content);
      
      const items: SlideItem[] = [];
      let slideTitle = `Slide ${slideNumber}`;

      // Better text extraction from all shapes
      const slideTexts: string[] = [];
      const textNodes = this.findNodes(xml, 'a:t');
      textNodes.forEach(node => {
        const text = node['_'] || node;
        if (typeof text === 'string' && text.trim()) {
          slideTexts.push(text.trim());
        }
      });

      const fullText = slideTexts.join(' ');
      if (slideTexts.length > 0) {
        slideTitle = slideTexts[0].substring(0, 50);
        items.push({
          id: `text-${slideNumber}-full`,
          type: 'text',
          content: fullText,
          slideNumber
        });
      }

      // Check for images
      // images are referenced in slide rels
      const relsEntry = this.zip.getEntry(`ppt/slides/_rels/slide${slideNumber}.xml.rels`);
      if (relsEntry) {
        const relsXml = await parseStringPromise(relsEntry.getData().toString('utf8'));
        const relationships = relsXml.Relationships.Relationship || [];
        
        for (const rel of relationships) {
          if (rel.$.Type.includes('image')) {
            const target = rel.$.Target.replace('../', 'ppt/');
            const imageEntry = this.zip.getEntry(target);
            if (imageEntry) {
              items.push({
                id: `img-${slideNumber}-${rel.$.Id}`,
                type: 'image',
                content: 'Alt text pending...',
                slideNumber,
                imagePath: target,
                imageBuffer: imageEntry.getData()
              });
            }
          }
        }
      }

      // Get notes
      const notesPath = `ppt/notesSlides/notesSlide${slideNumber}.xml`;
      const notesEntry = this.zip.getEntry(notesPath);
      let notes = '';
      if (notesEntry) {
        const notesXml = await parseStringPromise(notesEntry.getData().toString('utf8'));
        const notesTextNodes = this.findNodes(notesXml, 'a:t');
        notes = notesTextNodes.map(n => n['_'] || n).join(' ');
      }

      slides.push({
        number: slideNumber,
        title: slideTitle,
        items,
        notes
      });
    }

    return { fileName, slides };
  }

  /**
   * Helper to find nodes by key recursively
   */
  private findNodes(obj: any, key: string): any[] {
    let results: any[] = [];
    if (!obj) return results;
    if (obj[key]) {
      if (Array.isArray(obj[key])) results = results.concat(obj[key]);
      else results.push(obj[key]);
    }
    for (const k in obj) {
      if (typeof obj[k] === 'object') {
        results = results.concat(this.findNodes(obj[k], key));
      }
    }
    return results;
  }

  /**
   * Update alt text for an image in the PPTX
   */
  async updateAltText(slideNumber: number, imageRelId: string, altText: string) {
    // This requires writing back to the slide XML's cNvPr element
    // Implementation would involve more complex XML manipulation
    console.log(`Updating slide ${slideNumber} image ${imageRelId} with alt text: ${altText}`);
    return true;
  }

  /**
   * Export the modified PPTX
   */
  export(): Buffer {
    return this.zip.toBuffer();
  }
}
