import { DocumentData } from '../../types';

export class WordService {
  static async getSelectedText(): Promise<DocumentData> {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();
      return {
        selectedText: selection.text
      };
    });
  }

  static async getFullDocument(): Promise<DocumentData> {
    return Word.run(async (context) => {
      const body = context.document.body;
      body.load("text");
      await context.sync();
      return {
        selectedText: "",
        fullText: body.text
      };
    });
  }

  static async getDocumentProperties(): Promise<{title: string; author: string; wordCount: number}> {
    return Word.run(async (context) => {
      const properties = context.document.properties;
      properties.load(["title", "author"]);
      const body = context.document.body;
      body.load("text");
      await context.sync();
      
      const text = body.text || "";
      const wordCount = text.split(/\s+/).filter(word => word.length > 0).length;
      
      return {
        title: properties.title || 'Untitled',
        author: properties.author || 'Unknown',
        wordCount
      };
    });
  }

  static async replaceSelectedText(newText: string): Promise<void> {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.insertText(newText, Word.InsertLocation.replace);
      await context.sync();
    });
  }

  static async insertText(text: string, location: 'before'|'after'|'replace'): Promise<void> {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      let wordLoc = Word.InsertLocation.replace;
      if (location === 'before') wordLoc = Word.InsertLocation.before;
      else if (location === 'after') wordLoc = Word.InsertLocation.after;
      
      selection.insertText(text, wordLoc);
      await context.sync();
    });
  }

  static async insertParagraph(text: string, location: 'before'|'after'): Promise<void> {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      let wordLoc = Word.InsertLocation.after;
      if (location === 'before') wordLoc = Word.InsertLocation.before;
      
      selection.insertParagraph(text, wordLoc);
      await context.sync();
    });
  }

  static async getContextForAI(maxChars: number = 5000): Promise<string> {
    try {
      const selection = await this.getSelectedText();
      if (selection.selectedText && selection.selectedText.length > 0) {
        return `Selected Text:\n${selection.selectedText.substring(0, maxChars)}`;
      }
      const fullDoc = await this.getFullDocument();
      return `Document Text (excerpt):\n${(fullDoc.fullText || "").substring(0, maxChars)}`;
    } catch (e) {
      return `Unable to read Word context: ${(e as Error).message}`;
    }
  }
}
