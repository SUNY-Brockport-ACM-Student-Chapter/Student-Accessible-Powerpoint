import axios from 'axios';

/**
 * Service to interact with Google Gemini for image descriptions.
 */
export class GeminiService {
  private apiKey: string;
  private apiEndpoint = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent";

  constructor(apiKey: string) {
    this.apiKey = apiKey;
  }

  /**
   * Generates an accessibility description for an image.
   */
  async describeImage(imageBuffer: Buffer, mimeType: string): Promise<string> {
    const base64Image = imageBuffer.toString('base64');

    const payload = {
      contents: [
        {
          parts: [
            {
              text: "Analyze this image and provide a comprehensive description suitable for accessibility. " +
                    "Include: main subject, key elements, context, and purpose. " +
                    "Be descriptive but concise (under 125 characters for alt text)."
            },
            {
              inline_data: {
                mime_type: mimeType,
                data: base64Image
              }
            }
          ]
        }
      ],
      generationConfig: {
        maxOutputTokens: 150
      }
    };

    try {
      const response = await axios.post(`${this.apiEndpoint}?key=${this.apiKey}`, payload);
      const description = response.data?.candidates?.[0]?.content?.parts?.[0]?.text;
      return description?.trim() || "No description available";
    } catch (error: any) {
      console.error("Gemini API Error:", error.response?.data || error.message);
      throw new Error("Failed to generate image description");
    }
  }
}
