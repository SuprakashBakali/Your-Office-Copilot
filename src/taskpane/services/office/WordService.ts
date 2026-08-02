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

  static async insertTable(values: string[][], location: 'before'|'after'|'end'): Promise<void> {
    if (!values?.length || !values[0]?.length) {
      throw new Error('Table data must be a non-empty 2D array');
    }
    return Word.run(async (context) => {
      const doc = context.document;
      let target: any = doc.body;
      let wordLoc = Word.InsertLocation.end;

      if (location !== 'end') {
        target = doc.getSelection();
        wordLoc = location === 'before' ? Word.InsertLocation.before : Word.InsertLocation.after;
      }

      target.insertTable(values.length, values[0].length, wordLoc, values);
      await context.sync();
    });
  }

  static async formatText(opts: { bold?: boolean; italic?: boolean; color?: string; size?: number }): Promise<void> {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      if (opts.bold !== undefined) selection.font.bold = opts.bold;
      if (opts.italic !== undefined) selection.font.italic = opts.italic;
      if (opts.color) selection.font.color = opts.color;
      if (opts.size) selection.font.size = opts.size;
      await context.sync();
    });
  }

  static async applyStyle(styleName: string): Promise<void> {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.style = styleName;
      await context.sync();
    });
  }

  static async clearFormatting(): Promise<void> {
    return Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.font.reset();
      await context.sync();
    });
  }

  static async searchReplace(findText: string, replaceText: string): Promise<void> {
    return Word.run(async (context) => {
      const searchResults = context.document.body.search(findText, { matchCase: false, matchWholeWord: false });
      searchResults.load('items');
      await context.sync();

      for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].insertText(replaceText, Word.InsertLocation.replace);
      }
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

  // ── From office-agents: Word eval_js escape hatch ─────────────────────────
  /**
   * Execute arbitrary Word.js code inside Word.run.
   * Returns the value returned by the code, or null.
   */
  static async evalOfficeJs(code: string): Promise<any> {
    return Word.run(async (context) => {
      // eslint-disable-next-line no-new-func
      const fn = new Function('context', 'Word', 'Office', `return (async () => { ${code} })()`);
      const result = await fn(context, Word, typeof Office !== 'undefined' ? Office : {});
      return result ?? null;
    });
  }

  // ── From office-agents: get_document_structure ────────────────────────────
  /** List document headings, paragraphs count, and tables summary. */
  static async getDocumentStructure(): Promise<{ paragraphsCount: number; headings: string[]; tablesCount: number }> {
    return Word.run(async (context) => {
      const paragraphs = context.document.body.paragraphs;
      const tables = context.document.body.tables;
      paragraphs.load(['style', 'text']);
      tables.load('items');
      await context.sync();

      const headings: string[] = [];
      let paragraphsCount = 0;
      paragraphs.items.forEach(p => {
        paragraphsCount++;
        const s = (p.style || '').toString().toLowerCase();
        if (s.includes('heading') || s.includes('title')) {
          const t = (p.text || '').trim();
          if (t) headings.push(`${p.style}: ${t}`);
        }
      });

      return {
        paragraphsCount,
        headings: headings.slice(0, 30),
        tablesCount: tables.items.length
      };
    });
  }

  // ── Highlight Search ──────────────────────────────────────────────────────
  /** Highlight matching text across the Word document with a highlight color. */
  static async searchAndHighlight(findText: string, highlightColor: string = 'Yellow'): Promise<number> {
    return Word.run(async (context) => {
      const searchResults = context.document.body.search(findText, { matchCase: false, matchWholeWord: false });
      searchResults.load('items');
      await context.sync();

      for (let i = 0; i < searchResults.items.length; i++) {
        searchResults.items[i].font.highlightColor = highlightColor;
      }
      await context.sync();
      return searchResults.items.length;
    });
  }
}

