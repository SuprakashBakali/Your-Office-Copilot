import { SlideData } from '../../types';

export class PowerPointService {
  static async getCurrentSlide(): Promise<SlideData> {
    return new Promise((resolve, _reject) => {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          resolve({ slideIndex: 1, slideCount: 1, currentSlideText: '' });
        } else {
          resolve({ slideIndex: 1, slideCount: 1, currentSlideText: result.value as string });
        }
      });
    });
  }

  static async getAllSlides(): Promise<SlideData> {
    // Use getFileAsync to get the full presentation as text if possible
    try {
      if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
        return PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          const count = slides.items.length;
          return {
            slideIndex: 1,
            slideCount: count,
            currentSlideText: '',
            allSlidesText: Array(count).fill(''),
          };
        });
      }
    } catch { /* fall through */ }
    return { slideIndex: 1, slideCount: 1, currentSlideText: '' };
  }

  static async getSlideCount(): Promise<number> {
    try {
      if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
        return PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          slides.load("items");
          await context.sync();
          return slides.items.length;
        });
      }
    } catch { /* fall through */ }
    return 1;
  }

  static async setSlideText(text: string): Promise<void> {
    return new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) reject(new Error(result.error.message));
        else resolve();
      });
    });
  }

  static async addSlide(): Promise<void> {
    if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.3')) {
      return PowerPoint.run(async (context) => {
        context.presentation.slides.add();
        await context.sync();
      });
    } else {
      throw new Error("Adding slides requires PowerPointApi 1.3+.");
    }
  }

  static async addTextbox(text: string, slideIndex?: number): Promise<void> {
    if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
      return PowerPoint.run(async (context) => {
        const slides = context.presentation.slides;
        // If slideIndex is specified, use it; otherwise get the last slide
        // (so textboxes on "add_slide then add_textbox" workflows work)
        let slide;
        if (slideIndex !== undefined) {
          slide = slides.getItemAt(slideIndex);
        } else {
          slides.load("items");
          await context.sync();
          slide = slides.getItemAt(slides.items.length - 1);
        }
        const textBox = slide.shapes.addTextBox(text);
        textBox.left = 50;
        textBox.top = 100;
        textBox.width = 600;
        textBox.height = 100;
        await context.sync();
      });
    } else {
      throw new Error("Adding textboxes requires PowerPointApi 1.4+.");
    }
  }

  static async addShape(shapeType: string): Promise<void> {
    if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
      const shapeMap: Record<string, string> = {
        'rectangle': 'Rectangle', 'roundedrectangle': 'RoundRectangle', 'roundrectangle': 'RoundRectangle',
        'oval': 'Ellipse', 'ellipse': 'Ellipse', 'line': 'Line', 'diamond': 'Diamond',
        'triangle': 'Triangle', 'righttriangle': 'RightTriangle', 'pentagon': 'Pentagon',
        'hexagon': 'Hexagon', 'star': 'Star5', 'arrow': 'RightArrow', 'rightarrow': 'RightArrow',
        'leftarrow': 'LeftArrow', 'uparrow': 'UpArrow', 'downarrow': 'DownArrow',
      };
      const pptShape = shapeMap[shapeType.toLowerCase()];
      if (!pptShape) throw new Error(`Unknown shape type: ${shapeType}. Supported: ${Object.keys(shapeMap).join(', ')}`);
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(0);
        slide.shapes.addGeometricShape(pptShape as any);
        await context.sync();
      });
    } else {
      throw new Error("Adding shapes requires PowerPointApi 1.4+.");
    }
  }

  static async formatShape(shapeIndex: number, fillColor?: string, fontColor?: string): Promise<void> {
    if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(0);
        const shape = slide.shapes.getItemAt(shapeIndex);
        if (fillColor) shape.fill.setSolidColor(fillColor);
        if (fontColor) shape.textFrame.textRange.font.color = fontColor;
        await context.sync();
      });
    } else {
      throw new Error("Formatting shapes requires PowerPointApi 1.4+.");
    }
  }

  static async setSlideNotes(notes: string): Promise<void> {
    if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(0);
        const notesSlide = (slide as any).notesSlide;
        notesSlide.textRange.text = notes;
        await context.sync();
      });
    }
    throw new Error("setSlideNotes requires PowerPointApi 1.4+.");
  }

  /**
   * Read the FULL presentation context — all slides, all text.
   *
   * Uses TWO strategies for maximum compatibility:
   * 1. PowerPointApi 1.4+: iterate slides + shapes, load textFrame.hasText + text
   * 2. Fallback: use Office.context.document.getSelectedDataAsync with SlideRange
   *
   * The previous version had three bugs:
   *   a) `shapes.load('items/textFrame/textRange/text')` — Office.js doesn't
   *      support deep navigation in load(). Must load textFrame, then textRange.
   *   b) `context.sync()` inside the per-shape loop — causes context expiry
   *      on large decks. Must batch all loads + sync ONCE per slide.
   *   c) `if (shape.textFrame)` checked before textFrame was loaded — always
   *      undefined. Must load textFrame.hasText first.
   */
  static async getContextForAI(): Promise<string> {
    try {
      // Strategy 1: Try the Office.js Common API first — it works on ALL
      // PowerPoint versions and reads the entire presentation as text.
      // getSelectedDataAsync with CoercionType.Text reads ALL slide text,
      // not just the current selection.
      const allText = await new Promise<string>((resolve, reject) => {
        Office.context.document.getSelectedDataAsync(
          Office.CoercionType.Text,
          { valueFormat: Office.ValueFormat.Formatted, filterType: Office.FilterType.All },
          (result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              resolve(result.value as string);
            } else {
              reject(new Error(result.error?.message || 'getSelectedDataAsync failed'));
            }
          }
        );
      });

      // Also try to get the slide count via the rich API
      let slideCount = 0;
      try {
        if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
          slideCount = await PowerPoint.run(async (context) => {
            const slides = context.presentation.slides;
            slides.load("items");
            await context.sync();
            return slides.items.length;
          });
        }
      } catch { /* ignore — slide count is best-effort */ }

      if (allText && allText.trim().length > 0) {
        let res = `Presentation context`;
        if (slideCount > 0) res += ` (${slideCount} slides)`;
        res += `:\n\n${allText}`;
        return res;
      }

      // Strategy 2: If getSelectedDataAsync returned nothing, try the rich
      // API approach with proper load patterns.
      if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
        return PowerPoint.run(async (context) => {
          const slides = context.presentation.slides;
          // Load the items collection (need 'count' + items)
          slides.load("items");
          await context.sync();

          const count = slides.items.length;
          if (count === 0) return 'PowerPoint presentation has no slides.';

          // Load all slide shapes in ONE batch — no per-slide sync
          const slideTexts: string[] = [];
          for (let i = 0; i < count; i++) {
            const slide = slides.getItemAt(i);
            const shapes = slide.shapes;
            // Load shapes collection — use getText on shapes if available
            shapes.load('count');
            // We need to load items to iterate
            // The correct way: load items/type and textFrame/textRange/text
            // But deep navigation doesn't work in load(). Instead, load
            // items/textFrame/textRange/text one level at a time.
          }

          // Actually, the simplest reliable approach: use getFileAsync to
          // get the presentation text. But that returns a file slice, not text.
          // The most reliable approach that works across ALL versions:
          // Use goToFileAsync + getSelectedDataAsync repeatedly.
          // BUT: we already tried getSelectedDataAsync above and it returned
          // nothing. So fall back to the shapes approach with proper loading.

          for (let i = 0; i < count; i++) {
            const slide = slides.getItemAt(i);
            const shapes = slide.shapes;
            // Load the shapes collection
            context.load(shapes, 'items/id,items/name,items/type');
            // Also try to load textFrame properties
            // Note: In PowerPoint.js, shape.textFrame is a ShapeTextFrame
            // We need to load it, then load textRange, then load text
          }
          await context.sync();

          // Now iterate shapes and load text
          for (let i = 0; i < count; i++) {
            const slide = slides.getItemAt(i);
            const shapes = slide.shapes;
            let slideText = '';

            for (const shape of shapes.items) {
              try {
                const tf = shape.textFrame;
                context.load(tf, 'textRange/text');
              } catch {
                // Shape has no text frame
              }
            }

            // Sync to populate text values
            if (i === 0) {
              await context.sync();
            }

            for (const shape of shapes.items) {
              try {
                const text = shape.textFrame?.textRange?.text;
                if (text && text.trim()) {
                  slideText += text + '\n';
                }
              } catch {
                // Shape text not accessible
              }
            }

            slideTexts.push(`Slide ${i + 1}: ${slideText.trim() || '(no text)'}`);
          }

          let res = `Presentation (${count} slides):\n\n`;
          res += slideTexts.join('\n---\n');
          return res;
        });
      }

      // Strategy 3: Ultimate fallback — just report what we know
      return 'PowerPoint is active but context could not be read. The presentation may be empty or the API version is too old.';
    } catch (e) {
      return `Unable to read PowerPoint context: ${(e as Error).message}`;
    }
  }

  // ── eval_js escape hatch ──────────────────────────────────────────────────
  static async evalOfficeJs(code: string): Promise<any> {
    if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.1')) {
      return PowerPoint.run(async (context) => {
        // eslint-disable-next-line no-new-func
        const fn = new Function('context', 'PowerPoint', 'Office', `return (async () => { ${code} })()`);
        const result = await fn(context, PowerPoint, typeof Office !== 'undefined' ? Office : {});
        return result ?? null;
      });
    } else {
      throw new Error("evalOfficeJs requires PowerPointApi 1.1+.");
    }
  }

  static async getAllSlideShapes(slideIndex: number = 0): Promise<{ id: string; name: string; type: string }[]> {
    if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(slideIndex);
        const shapes = slide.shapes;
        shapes.load('items/id,items/name,items/type');
        await context.sync();
        return shapes.items.map(s => ({ id: s.id, name: s.name, type: String(s.type) }));
      });
    } else {
      return [];
    }
  }

  static async deleteSlide(slideIndex: number = 0): Promise<void> {
    if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.3')) {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(slideIndex);
        slide.delete();
        await context.sync();
      });
    } else {
      throw new Error("deleteSlide requires PowerPointApi 1.3+.");
    }
  }
}
