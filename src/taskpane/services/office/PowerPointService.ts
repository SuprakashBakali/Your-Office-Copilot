import { SlideData } from '../../types';

export class PowerPointService {
  static async getCurrentSlide(): Promise<SlideData> {
    return new Promise((resolve, _reject) => {
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          resolve({
            slideIndex: 1,
            slideCount: 1,
            currentSlideText: ''
          });
        } else {
          resolve({
            slideIndex: 1,
            slideCount: 1,
            currentSlideText: result.value as string
          });
        }
      });
    });
  }

  static async getAllSlides(): Promise<SlideData> {
    return {
      slideIndex: 1,
      slideCount: 1,
      currentSlideText: '',
      allSlidesText: ['(Full document reading requires newer PowerPoint API)']
    };
  }

  static async getSlideCount(): Promise<number> {
    return 1;
  }

  static async setSlideText(text: string): Promise<void> {
    return new Promise((resolve, reject) => {
      Office.context.document.setSelectedDataAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
        if (result.status === Office.AsyncResultStatus.Failed) {
          reject(new Error(result.error.message));
        } else {
          resolve();
        }
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
      throw new Error("Adding slides is not supported in this version of PowerPoint.");
    }
  }

  static async addTextbox(text: string): Promise<void> {
    if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(0);
        slide.shapes.addTextBox(text);
        await context.sync();
      });
    } else {
      throw new Error("Adding textboxes requires PowerPointApi 1.4+.");
    }
  }

  static async addShape(shapeType: string): Promise<void> {
    if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
      // Map common string names to PowerPoint.GeometricShapeType enum values.
      // Passing a raw string like "Rectangle" to addGeometricShape throws.
      // PowerPoint.addGeometricShape accepts string literal values
      // (e.g. "Rectangle", "Ellipse") — not enum members.
      const shapeMap: Record<string, string> = {
        'rectangle': 'Rectangle',
        'roundedrectangle': 'RoundRectangle',
        'roundrectangle': 'RoundRectangle',
        'oval': 'Ellipse',
        'ellipse': 'Ellipse',
        'line': 'Line',
        'diamond': 'Diamond',
        'triangle': 'Triangle',
        'righttriangle': 'RightTriangle',
        'pentagon': 'Pentagon',
        'hexagon': 'Hexagon',
        'star': 'Star5',
        'arrow': 'RightArrow',
        'rightarrow': 'RightArrow',
        'leftarrow': 'LeftArrow',
        'uparrow': 'UpArrow',
        'downarrow': 'DownArrow',
      };
      const pptShape = shapeMap[shapeType.toLowerCase()];
      if (!pptShape) {
        throw new Error(`Unknown shape type: ${shapeType}. Supported: ${Object.keys(shapeMap).join(', ')}`);
      }
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
    // PowerPointApi 1.4+ supports reading/writing slide notes via
    // slide.notesSlide.textRange.text property.
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

  static async getContextForAI(): Promise<string> {
    try {
      const slide = await this.getCurrentSlide();
      if (slide.currentSlideText) {
        return `Selected Text: ${slide.currentSlideText}`;
      }
      return `PowerPoint active, but no text selected or readable.`;
    } catch (e) {
      return `Unable to read PowerPoint context: ${(e as Error).message}`;
    }
  }

  // ── From office-agents: PPT eval_js escape hatch ──────────────────────────
  /**
   * Execute arbitrary PowerPoint.js code inside PowerPoint.run.
   * Returns the value returned by the code, or null.
   */
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

  // ── From office-agents: list_slide_shapes ─────────────────────────────────
  /** List all shapes on the active or specified slide. */
  static async getAllSlideShapes(slideIndex: number = 0): Promise<{ id: string; name: string; type: string }[]> {
    if (typeof PowerPoint !== 'undefined' && Office.context.requirements.isSetSupported('PowerPointApi', '1.4')) {
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(slideIndex);
        const shapes = slide.shapes;
        shapes.load('items/id, items/name, items/type');
        await context.sync();
        return shapes.items.map(s => ({
          id: s.id,
          name: s.name,
          type: s.type
        }));
      });
    } else {
      return [];
    }
  }

  // ── Delete Slide ──────────────────────────────────────────────────────────
  /** Delete a slide by its 0-based index. */
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

