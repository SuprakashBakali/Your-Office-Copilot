import { SlideData } from '../../types';

export class PowerPointService {
  static async getCurrentSlide(): Promise<SlideData> {
    return new Promise((resolve, reject) => {
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
      return PowerPoint.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(0);
        slide.shapes.addGeometricShape(shapeType as any);
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
    throw new Error("Setting slide notes requires newer PowerPoint API, not fully implemented in Office JS yet.");
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
}
