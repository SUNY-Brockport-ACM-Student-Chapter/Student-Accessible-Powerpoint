import { NextRequest, NextResponse } from 'next/server';
import { PptxProcessor } from '@/lib/pptx-utils';
import { GeminiService } from '@/lib/gemini';

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get('file') as File;
    
    if (!file) {
      return NextResponse.json({ error: 'No file uploaded' }, { status: 400 });
    }

    const buffer = Buffer.from(await file.arrayBuffer());
    const processor = new PptxProcessor(buffer);
    const pptData = await processor.getPresentationData(file.name);

    // Initial analysis result
    // In a real app, we might start a background job or process images here
    // For this demo/conversion, we return the parsed structure
    
    return NextResponse.json({ 
      success: true, 
      presentation: {
        fileName: pptData.fileName,
        slidesCount: pptData.slides.length,
        slides: pptData.slides.map(s => ({
            number: s.number,
            title: s.title,
            imageCount: s.items.filter(i => i.type === 'image').length
        }))
      }
    });

  } catch (error: any) {
    console.error('Processing error:', error);
    return NextResponse.json({ error: error.message }, { status: 500 });
  }
}
