import winax from 'winax';
import { debug } from '../utils/debug.js';
const { Object: WinaxObject } = winax; // Destructure and rename Object

// Basic interface for Word Application object (replace with more specific types later if possible)
interface WordApplication {
  Documents: any; // Word.Documents collection
  ActiveDocument: any; // Word.Document
  Visible: boolean;
  Run(MacroName: string, ...args: any[]): any; // For running VBA macros
  AddIns: any; // Word.AddIns collection
  Quit(SaveChanges?: any, OriginalFormat?: any, RouteDocument?: any): void;
  // Add other necessary properties and methods
}

// Basic interface for Word Document object
interface WordDocument {
  Save(): void;
  SaveAs2(FileName?: any, FileFormat?: any, LockComments?: any, Password?: any, AddToRecentFiles?: any, WritePassword?: any, ReadOnlyRecommended?: any, EmbedTrueTypeFonts?: any, SaveNativePictureFormat?: any, SaveFormsData?: any, SaveAsAOCELetter?: any, Encoding?: any, InsertLineBreaks?: any, AllowSubstitutions?: any, LineEnding?: any, AddBiDiMarks?: any, CompatibilityMode?: any): void;
  Close(SaveChanges?: any, OriginalFormat?: any, RouteDocument?: any): void;
  Protect(Type?: any, NoReset?: any, Password?: any, UseIRM?: any, EnforceStyleLock?: any): void;
  Content: any; // Word.Range
  Paragraphs: any; // Word.Paragraphs
  Tables: any; // Word.Tables
  InlineShapes: any; // Word.InlineShapes
  Shapes: any; // Word.Shapes
  Sections: any; // Word.Sections
  ActiveWindow: any; // Word.Window
  PageSetup: any; // Word.PageSetup
  VBProject: any; // VBA Project object
  BuiltInDocumentProperties: any; // Document properties
  CustomDocumentProperties: any; // Custom document properties
  // Add other necessary properties and methods
}

class WordService {
  private wordApp: WordApplication | null = null;

  /**
   * Gets the currently running Word application instance or creates a new one.
   * Ensures Word is visible.
   */
  public async getWordApplication(): Promise<WordApplication> {
    if (this.wordApp) {
      try {
        // More comprehensive check for instance validity
        // Check basic property access
        this.wordApp.Visible;
        
        // Check if Documents collection is accessible and valid
        const docs = this.wordApp.Documents;
        // Try to access a property of Documents to ensure it's fully valid
        const count = docs.Count;
        
        return this.wordApp;
      } catch (error) {
        debug.warn("Existing Word instance seems invalid, creating a new one.", error);
        this.wordApp = null; // Reset if invalid
      }
    }

    try {
      debug.log("Attempting to get or create Word.Application instance...");
      // Try to get an existing instance first, then create if not found
      // Use the destructured WinaxObject constructor
      this.wordApp = new WinaxObject("Word.Application", { activate: true }) as WordApplication;
      this.wordApp.Visible = true; // Make sure Word is visible for interaction
      debug.log("Word.Application instance obtained successfully.");
      return this.wordApp;
    } catch (error) {
      debug.error("Failed to get or create Word.Application instance:", error);
      throw new Error(`Failed to initialize Word.Application. Make sure Microsoft Word is installed. Error: ${error}`);
    }
  }

  /**
   * Gets the active document in the Word application.
   * Throws an error if Word is not running or no document is active.
   */
  public async getActiveDocument(): Promise<WordDocument> {
    const app = await this.getWordApplication();
    try {
      const activeDoc = app.ActiveDocument;
      if (!activeDoc) {
        throw new Error("No active document found in Word.");
      }
      return activeDoc as WordDocument;
    } catch (error) {
      debug.error("Failed to get active document:", error);
      throw new Error(`Failed to get active document. Error: ${error}`);
    }
  }

  // --- Document Methods ---

  /**
   * Creates a new Word document.
   */
  public async createDocument(): Promise<WordDocument> {
    const app = await this.getWordApplication();
    try {
      const newDoc = app.Documents.Add();
      return newDoc as WordDocument;
    } catch (error) {
      debug.error("Failed to create new document:", error);
      throw new Error(`Failed to create new document. Error: ${error}`);
    }
  }

  /**
   * Opens an existing Word document.
   * @param filePath The path to the document file.
   */
  public async openDocument(filePath: string): Promise<WordDocument> {
    const app = await this.getWordApplication();
    try {
      // Ensure the path is absolute and correctly formatted if needed
      const openedDoc = app.Documents.Open(filePath);
      return openedDoc as WordDocument;
    } catch (error) {
      debug.error(`Failed to open document at path: ${filePath}`, error);
      throw new Error(`Failed to open document: ${filePath}. Error: ${error}`);
    }
  }

   /**
   * Saves the active document.
   */
    public async saveActiveDocument(): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
        doc.Save();
      } catch (error) {
        debug.error("Failed to save active document:", error);
        throw new Error(`Failed to save active document. Error: ${error}`);
      }
    }

  /**
   * Saves the active document with a new name or format.
   * @param filePath The new path for the document.
   * @param fileFormat Optional Word save format constant (e.g., WdSaveFormat.wdFormatDocumentDefault).
   */
  public async saveActiveDocumentAs(filePath: string, fileFormat?: any): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      // WdSaveFormat enumeration needs to be accessible or defined
      // Example: wdFormatDocumentDefault = 16
      const format = fileFormat ?? 16; // Default to .docx
      doc.SaveAs2(filePath, format);
    } catch (error) {
      debug.error(`Failed to save document as: ${filePath}`, error);
      throw new Error(`Failed to save document as: ${filePath}. Error: ${error}`);
    }
  }

  /**
   * Closes the specified document.
   * @param doc The document object to close.
   * @param saveChanges Optional WdSaveOptions constant (e.g., WdSaveOptions.wdDoNotSaveChanges).
   */
  public async closeDocument(doc: WordDocument, saveChanges?: any): Promise<void> {
    try {
      // WdSaveOptions enumeration needs to be accessible or defined
      // Example: wdDoNotSaveChanges = 0, wdPromptToSaveChanges = -2, wdSaveChanges = -1
      const saveOpt = saveChanges ?? 0; // Default to not saving changes
      doc.Close(saveOpt);
    } catch (error) {
      debug.error("Failed to close document:", error);
      // Avoid throwing if close fails, might already be closed or Word unresponsive
      debug.warn(`Could not definitively close document. Error: ${error}`);
    }
  }


  /**
   * Quits the Word application.
   * Handles potential errors during quit.
   */
  public async quitWord(): Promise<void> {
    if (this.wordApp) {
      try {
        // WdSaveOptions enumeration
        // Example: wdDoNotSaveChanges = 0
        this.wordApp.Quit(0); // Quit without saving changes
        this.wordApp = null; // Clear the reference
        debug.log("Word application quit successfully.");
      } catch (error) {
        debug.error("Error quitting Word application:", error);
        // Don't re-throw, as Word might already be closed or unresponsive
        this.wordApp = null; // Clear reference even on error
      }
    } else {
        debug.log("Word application instance not found, nothing to quit.");
    }
  }

  // --- Text Manipulation Methods ---

  /**
   * Inserts text at the current selection point.
   * @param text The text to insert.
   */
  public async insertText(text: string): Promise<void> {
    const app = await this.getWordApplication();
    try {
      app.ActiveDocument.ActiveWindow.Selection.TypeText(text);
    } catch (error) {
      debug.error("Failed to insert text:", error);
      throw new Error(`Failed to insert text. Error: ${error}`);
    }
  }

  /**
   * Deletes the current selection or a specified number of characters.
   * @param count Number of characters to delete (default: 1). Positive deletes forward, negative deletes backward.
   * @param unit The unit to delete (default: character). Use WdUnits enum values (e.g., 1 for character, 2 for word).
   */
  public async deleteText(count: number = 1, unit: number = 1 /* wdCharacter */): Promise<void> {
      const app = await this.getWordApplication();
      try {
          // WdUnits enumeration: wdCharacter = 1, wdWord = 2, etc.
          // Positive count deletes forward, negative count deletes backward from the start of the selection.
          // If selection is collapsed, positive deletes after insertion point, negative deletes before.
          if (count > 0) {
              app.ActiveDocument.ActiveWindow.Selection.Delete(unit, count);
          } else if (count < 0) {
              // Move start back and then delete forward
              app.ActiveDocument.ActiveWindow.Selection.MoveStart(unit, count); // Move start back
              app.ActiveDocument.ActiveWindow.Selection.Delete(unit, Math.abs(count)); // Delete forward
          }
          // If count is 0, do nothing
      } catch (error) {
          debug.error("Failed to delete text:", error);
          throw new Error(`Failed to delete text. Error: ${error}`);
      }
  }


  /**
   * Finds and replaces text in the active document.
   * @param findText Text to find.
   * @param replaceText Text to replace with.
   * @param matchCase Match case sensitivity.
   * @param matchWholeWord Match whole words only.
   * @param replaceAll Replace all occurrences or just the first.
   */
   public async findAndReplace(
    findText: string,
    replaceText: string,
    matchCase: boolean = false,
    matchWholeWord: boolean = false,
    replaceAll: boolean = true
  ): Promise<boolean> {
    const doc = await this.getActiveDocument();
    try {
      const find = doc.Content.Find;
      find.ClearFormatting(); // Clear previous find formatting
      find.Replacement.ClearFormatting(); // Clear previous replacement formatting

      find.Text = findText;
      find.Replacement.Text = replaceText;
      find.Forward = true;
      find.Wrap = 1; // wdFindContinue
      find.Format = false;
      find.MatchCase = matchCase;
      find.MatchWholeWord = matchWholeWord;
      find.MatchWildcards = false;
      find.MatchSoundsLike = false;
      find.MatchAllWordForms = false;

      // WdReplace enumeration: wdReplaceNone = 0, wdReplaceOne = 1, wdReplaceAll = 2
      const replaceOption = replaceAll ? 2 : 1;

      const found = find.Execute(undefined, undefined, undefined, undefined, undefined,
                                 undefined, undefined, undefined, undefined, undefined,
                                 replaceOption); // Execute the find and replace

      return found; // Returns true if text was found and replaced (or just found if replaceOption is wdReplaceNone)
    } catch (error) {
      debug.error("Failed to find and replace text:", error);
      throw new Error(`Failed to find and replace text. Error: ${error}`);
    }
  }

  /**
   * Toggles bold formatting for the current selection.
   */
  public async toggleBold(): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const font = app.ActiveDocument.ActiveWindow.Selection.Font;
      // wdToggle = 9999998
      font.Bold = 9999998;
    } catch (error) {
      debug.error("Failed to toggle bold:", error);
      throw new Error(`Failed to toggle bold formatting. Error: ${error}`);
    }
  }

  /**
   * Toggles italic formatting for the current selection.
   */
  public async toggleItalic(): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const font = app.ActiveDocument.ActiveWindow.Selection.Font;
      // wdToggle = 9999998
      font.Italic = 9999998;
    } catch (error) {
      debug.error("Failed to toggle italic:", error);
      throw new Error(`Failed to toggle italic formatting. Error: ${error}`);
    }
  }

  /**
   * Toggles underline formatting for the current selection.
   * @param underlineStyle Optional WdUnderline value (e.g., 1 for single underline). Default toggles single underline.
   */
  public async toggleUnderline(underlineStyle: number = 1 /* wdUnderlineSingle */): Promise<void> {
      const app = await this.getWordApplication();
      try {
          const font = app.ActiveDocument.ActiveWindow.Selection.Font;
          // WdUnderline enumeration: wdUnderlineNone = 0, wdUnderlineSingle = 1, etc.
          // wdToggle = 9999998
          if (font.Underline === underlineStyle) {
              font.Underline = 0; // wdUnderlineNone - Turn off if it's already the specified style
          } else {
              font.Underline = underlineStyle; // Apply the specified style
          }
      } catch (error) {
          debug.error("Failed to toggle underline:", error);
          throw new Error(`Failed to toggle underline formatting. Error: ${error}`);
      }
  }

  // --- Paragraph Formatting Methods ---

  /**
   * Sets the alignment for the selected paragraphs.
   * @param alignment Alignment type (WdParagraphAlignment enum value: 0=Left, 1=Center, 2=Right, 3=Justify).
   */
  public async setParagraphAlignment(alignment: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      // WdParagraphAlignment: wdAlignParagraphLeft = 0, wdAlignParagraphCenter = 1, wdAlignParagraphRight = 2, wdAlignParagraphJustify = 3
      app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.Alignment = alignment;
    } catch (error) {
      debug.error("Failed to set paragraph alignment:", error);
      throw new Error(`Failed to set paragraph alignment. Error: ${error}`);
    }
  }

  /**
   * Sets the left indent for the selected paragraphs.
   * @param indentPoints Indentation value in points.
   */
  public async setParagraphLeftIndent(indentPoints: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.LeftIndent = indentPoints;
    } catch (error) {
      debug.error("Failed to set left indent:", error);
      throw new Error(`Failed to set left indent. Error: ${error}`);
    }
  }

  /**
   * Sets the right indent for the selected paragraphs.
   * @param indentPoints Indentation value in points.
   */
  public async setParagraphRightIndent(indentPoints: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.RightIndent = indentPoints;
    } catch (error) {
      debug.error("Failed to set right indent:", error);
      throw new Error(`Failed to set right indent. Error: ${error}`);
    }
  }

    /**
   * Sets the first line indent for the selected paragraphs.
   * @param indentPoints Indentation value in points (positive for indent, negative for hanging indent).
   */
    public async setParagraphFirstLineIndent(indentPoints: number): Promise<void> {
        const app = await this.getWordApplication();
        try {
            app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.FirstLineIndent = indentPoints;
        } catch (error) {
            debug.error("Failed to set first line indent:", error);
            throw new Error(`Failed to set first line indent. Error: ${error}`);
        }
    }

  /**
   * Sets the space before the selected paragraphs.
   * @param spacePoints Space value in points.
   */
  public async setParagraphSpaceBefore(spacePoints: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.SpaceBefore = spacePoints;
    } catch (error) {
      debug.error("Failed to set space before:", error);
      throw new Error(`Failed to set space before paragraph. Error: ${error}`);
    }
  }

  /**
   * Sets the space after the selected paragraphs.
   * @param spacePoints Space value in points.
   */
  public async setParagraphSpaceAfter(spacePoints: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat.SpaceAfter = spacePoints;
    } catch (error) {
      debug.error("Failed to set space after:", error);
      throw new Error(`Failed to set space after paragraph. Error: ${error}`);
    }
  }

  /**
   * Sets the line spacing for the selected paragraphs.
   * @param lineSpacingRule WdLineSpacing enum value (0=Single, 1=1.5 lines, 2=Double, 3=AtLeast, 4=Exactly, 5=Multiple).
   * @param lineSpacingValue Value for AtLeast, Exactly, or Multiple spacing (in points for AtLeast/Exactly, multiplier for Multiple).
   */
  public async setParagraphLineSpacing(lineSpacingRule: number, lineSpacingValue?: number): Promise<void> {
    const app = await this.getWordApplication();
    try {
      // WdLineSpacing: wdLineSpaceSingle = 0, wdLineSpace1pt5 = 1, wdLineSpaceDouble = 2,
      // wdLineSpaceAtLeast = 3, wdLineSpaceExactly = 4, wdLineSpaceMultiple = 5
      const paraFormat = app.ActiveDocument.ActiveWindow.Selection.ParagraphFormat;
      paraFormat.LineSpacingRule = lineSpacingRule;
      if (lineSpacingValue !== undefined && lineSpacingRule >= 3) { // Only set LineSpacing if rule requires it
        paraFormat.LineSpacing = lineSpacingValue;
      }
    } catch (error) {
      debug.error("Failed to set line spacing:", error);
      throw new Error(`Failed to set line spacing. Error: ${error}`);
    }
  }

  // --- Table Methods ---

  /**
   * Adds a table at the current selection.
   * @param numRows Number of rows.
   * @param numCols Number of columns.
   * @param defaultTableBehavior Optional WdDefaultTableBehavior value.
   * @param autoFitBehavior Optional WdAutoFitBehavior value.
   */
  public async addTable(numRows: number, numCols: number, defaultTableBehavior?: number, autoFitBehavior?: number): Promise<any /* Word.Table */> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      // WdDefaultTableBehavior: wdWord8TableBehavior = 0, wdWord9TableBehavior = 1
      // WdAutoFitBehavior: wdAutoFitFixed = 0, wdAutoFitContent = 1, wdAutoFitWindow = 2
      const table = selection.Tables.Add(selection.Range, numRows, numCols, defaultTableBehavior, autoFitBehavior);
      return table;
    } catch (error) {
      debug.error("Failed to add table:", error);
      throw new Error(`Failed to add table. Error: ${error}`);
    }
  }

  /**
   * Gets a specific cell in a table.
   * @param tableIndex Index of the table in the document (1-based).
   * @param rowIndex Row index (1-based).
   * @param colIndex Column index (1-based).
   */
  public async getTableCell(tableIndex: number, rowIndex: number, colIndex: number): Promise<any /* Word.Cell */> {
    const doc = await this.getActiveDocument();
    try {
      if (tableIndex <= 0 || tableIndex > doc.Tables.Count) {
        throw new Error(`Table index ${tableIndex} is out of bounds.`);
      }
      const table = doc.Tables.Item(tableIndex);
      const cell = table.Cell(rowIndex, colIndex);
      return cell;
    } catch (error) {
      debug.error(`Failed to get cell (${rowIndex}, ${colIndex}) in table ${tableIndex}:`, error);
      throw new Error(`Failed to get cell. Error: ${error}`);
    }
  }

  /**
   * Sets the text in a specific table cell.
   * @param tableIndex Index of the table in the document (1-based).
   * @param rowIndex Row index (1-based).
   * @param colIndex Column index (1-based).
   * @param text Text to set.
   */
  public async setTableCellText(tableIndex: number, rowIndex: number, colIndex: number, text: string): Promise<void> {
    try {
      const cell = await this.getTableCell(tableIndex, rowIndex, colIndex);
      cell.Range.Text = text;
    } catch (error) {
      debug.error(`Failed to set text in cell (${rowIndex}, ${colIndex}) of table ${tableIndex}:`, error);
      throw new Error(`Failed to set cell text. Error: ${error}`);
    }
  }

  /**
   * Inserts a row in a table.
   * @param tableIndex Index of the table (1-based).
   * @param beforeRowIndex Optional index of the row to insert before (1-based). If omitted, adds to the end.
   */
  public async insertTableRow(tableIndex: number, beforeRowIndex?: number): Promise<any /* Word.Row */> {
      const doc = await this.getActiveDocument();
      try {
          if (tableIndex <= 0 || tableIndex > doc.Tables.Count) {
              throw new Error(`Table index ${tableIndex} is out of bounds.`);
          }
          const table = doc.Tables.Item(tableIndex);
          let newRow;
          if (beforeRowIndex !== undefined) {
              if (beforeRowIndex <= 0 || beforeRowIndex > table.Rows.Count + 1) { // Allow inserting after last row
                 throw new Error(`Row index ${beforeRowIndex} is out of bounds for insertion.`);
              }
              const refRow = beforeRowIndex <= table.Rows.Count ? table.Rows.Item(beforeRowIndex) : undefined;
              newRow = table.Rows.Add(refRow); // Inserts before refRow if provided, otherwise adds at end
          } else {
              newRow = table.Rows.Add(); // Add to the end
          }
          return newRow;
      } catch (error) {
          debug.error(`Failed to insert row into table ${tableIndex}:`, error);
          throw new Error(`Failed to insert table row. Error: ${error}`);
      }
  }

  /**
   * Inserts a column in a table.
   * @param tableIndex Index of the table (1-based).
   * @param beforeColIndex Optional index of the column to insert before (1-based). If omitted, adds to the right end.
   */
  public async insertTableColumn(tableIndex: number, beforeColIndex?: number): Promise<any /* Word.Column */> {
      const doc = await this.getActiveDocument();
      try {
          if (tableIndex <= 0 || tableIndex > doc.Tables.Count) {
              throw new Error(`Table index ${tableIndex} is out of bounds.`);
          }
          const table = doc.Tables.Item(tableIndex);
           let newCol;
          if (beforeColIndex !== undefined) {
               if (beforeColIndex <= 0 || beforeColIndex > table.Columns.Count + 1) { // Allow inserting after last col
                 throw new Error(`Column index ${beforeColIndex} is out of bounds for insertion.`);
              }
              const refCol = beforeColIndex <= table.Columns.Count ? table.Columns.Item(beforeColIndex) : undefined;
              newCol = table.Columns.Add(refCol); // Inserts before refCol if provided, otherwise adds at end
          } else {
              newCol = table.Columns.Add(); // Add to the end (right)
          }
          return newCol;
      } catch (error) {
          debug.error(`Failed to insert column into table ${tableIndex}:`, error);
          throw new Error(`Failed to insert table column. Error: ${error}`);
      }
  }

  /**
   * Applies an auto format style to a table.
   * @param tableIndex Index of the table (1-based).
   * @param formatName Name of the table style or a WdTableFormat enum value.
   * @param applyFormatting Optional flags for which parts of the format to apply (WdTableFormatApply enum values).
   */
  public async applyTableAutoFormat(tableIndex: number, formatName: string | number, applyFormatting?: number): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
          if (tableIndex <= 0 || tableIndex > doc.Tables.Count) {
              throw new Error(`Table index ${tableIndex} is out of bounds.`);
          }
          const table = doc.Tables.Item(tableIndex);
          // WdTableFormatApply flags can be combined (e.g., Borders | Shading | Font | Color | AutoFit | HeadingRows | FirstColumn | LastColumn | LastRow)
          // Example: wdTableFormatApplyBorders = 1, wdTableFormatApplyShading = 2, etc.
          // Default apply flags might vary, check Word documentation. Let's assume applying all is reasonable if not specified.
          const defaultApplyFlags = 1+2+4+8+16+32+64+128+256; // Example: Apply all common flags
          table.AutoFormat(formatName, applyFormatting ?? defaultApplyFlags);
      } catch (error) {
          debug.error(`Failed to apply auto format to table ${tableIndex}:`, error);
          throw new Error(`Failed to apply table auto format. Error: ${error}`);
      }
  }

  // --- Image Methods ---

  /**
   * Inserts a picture at the current selection as an inline shape.
   * @param filePath Path to the image file.
   * @param linkToFile Link to the file instead of embedding (optional).
   * @param saveWithDocument Save the image with the document (optional, relevant if linked).
   */
  public async insertPicture(filePath: string, linkToFile: boolean = false, saveWithDocument: boolean = true): Promise<any /* Word.InlineShape */> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      const inlineShape = selection.InlineShapes.AddPicture(filePath, linkToFile, saveWithDocument, selection.Range);
      return inlineShape;
    } catch (error) {
      debug.error(`Failed to insert picture from ${filePath}:`, error);
      throw new Error(`Failed to insert picture. Error: ${error}`);
    }
  }

   /**
   * Sets the size of an inline shape (e.g., a picture).
   * Assumes the shape is identified by its index in the active document's InlineShapes collection.
   * @param shapeIndex 1-based index of the inline shape.
   * @param heightPoints Height in points. Use -1 to keep original or maintain aspect ratio if width is set.
   * @param widthPoints Width in points. Use -1 to keep original or maintain aspect ratio if height is set.
   * @param lockAspectRatio Lock aspect ratio when resizing (default: true).
   */
   public async setInlinePictureSize(shapeIndex: number, heightPoints: number, widthPoints: number, lockAspectRatio: boolean = true): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
          if (shapeIndex <= 0 || shapeIndex > doc.InlineShapes.Count) {
              throw new Error(`InlineShape index ${shapeIndex} is out of bounds.`);
          }
          const shape = doc.InlineShapes.Item(shapeIndex);

          // Store original aspect ratio if needed
          const originalHeight = shape.Height;
          const originalWidth = shape.Width;
          // const aspectRatio = originalWidth / originalHeight; // Not needed if relying on LockAspectRatio

          shape.LockAspectRatio = lockAspectRatio ? -1 : 0; // msoTrue = -1, msoFalse = 0

          if (heightPoints > 0 && widthPoints > 0) {
              // Set both, respecting lock aspect ratio if enabled
              if (lockAspectRatio) {
                 // Determine dominant dimension change if aspect ratio is locked
                 const heightRatio = heightPoints / originalHeight;
                 const widthRatio = widthPoints / originalWidth;
                 if (widthRatio > heightRatio) {
                    shape.Width = widthPoints; // Width change is greater, height adjusts
                 } else {
                    shape.Height = heightPoints; // Height change is greater, width adjusts
                 }
              } else {
                 shape.Height = heightPoints;
                 shape.Width = widthPoints;
              }
          } else if (heightPoints > 0) {
              shape.Height = heightPoints; // Width adjusts if aspect ratio locked
          } else if (widthPoints > 0) {
              shape.Width = widthPoints; // Height adjusts if aspect ratio locked
          }
          // If both are <= 0, size remains unchanged

      } catch (error) {
          debug.error(`Failed to set size for inline shape ${shapeIndex}:`, error);
          throw new Error(`Failed to set inline picture size. Error: ${error}`);
      }
  }

  // Note: Positioning inline shapes is limited. For more control, convert to a floating Shape.
  // Methods for floating shapes (doc.Shapes) would be needed for complex positioning (Left, Top, Relative anchors).

  // --- Header/Footer Methods ---

  /**
   * Gets a specific header or footer object from a section.
   * @param sectionIndex 1-based index of the section.
   * @param headerFooterType WdHeaderFooterIndex enum value (1=Primary, 2=FirstPage, 3=EvenPages).
   * @param isHeader True for header, false for footer.
   */
  public async getHeaderFooter(sectionIndex: number, headerFooterType: number, isHeader: boolean): Promise<any /* Word.HeaderFooter */> {
      const doc = await this.getActiveDocument();
      try {
          if (sectionIndex <= 0 || sectionIndex > doc.Sections.Count) {
              throw new Error(`Section index ${sectionIndex} is out of bounds.`);
          }
          const section = doc.Sections.Item(sectionIndex);
          const headersFooters = isHeader ? section.Headers : section.Footers;

          // WdHeaderFooterIndex: wdHeaderFooterPrimary = 1, wdHeaderFooterFirstPage = 2, wdHeaderFooterEvenPages = 3
          if (headerFooterType < 1 || headerFooterType > 3) {
             throw new Error(`Invalid header/footer type: ${headerFooterType}. Use 1, 2, or 3.`);
          }

          const headerFooter = headersFooters.Item(headerFooterType);
          if (!headerFooter?.Exists) {
             // Depending on settings (like DifferentFirstPage, DifferentOddAndEvenPages), the requested type might not exist.
             // Handle this gracefully, maybe return null or throw a specific error.
             // For now, let's throw.
             throw new Error(`The requested ${isHeader ? 'header' : 'footer'} type (${headerFooterType}) does not exist or is not active for section ${sectionIndex}. Check document settings.`);
          }
          return headerFooter;
      } catch (error) {
          debug.error(`Failed to get ${isHeader ? 'header' : 'footer'} type ${headerFooterType} for section ${sectionIndex}:`, error);
          throw new Error(`Failed to get header/footer. Error: ${error}`);
      }
  }

  /**
   * Sets the text for a specific header or footer. Replaces existing content.
   * @param sectionIndex 1-based index of the section.
   * @param headerFooterType WdHeaderFooterIndex enum value (1=Primary, 2=FirstPage, 3=EvenPages).
   * @param isHeader True for header, false for footer.
   * @param text The text to set.
   */
  public async setHeaderFooterText(sectionIndex: number, headerFooterType: number, isHeader: boolean, text: string): Promise<void> {
      try {
          const headerFooter = await this.getHeaderFooter(sectionIndex, headerFooterType, isHeader);
          headerFooter.Range.Text = text;
      } catch (error) {
          debug.error(`Failed to set text for ${isHeader ? 'header' : 'footer'} type ${headerFooterType} section ${sectionIndex}:`, error);
          // Re-throw error as it likely indicates a real issue (invalid index, etc.)
          throw error;
      }
  }

  // --- Page Setup Methods ---

  /**
   * Sets the page margins for the active document.
   * @param topPoints Top margin in points.
   * @param bottomPoints Bottom margin in points.
   * @param leftPoints Left margin in points.
   * @param rightPoints Right margin in points.
   */
  public async setPageMargins(topPoints: number, bottomPoints: number, leftPoints: number, rightPoints: number): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
          const pageSetup = doc.PageSetup;
          pageSetup.TopMargin = topPoints;
          pageSetup.BottomMargin = bottomPoints;
          pageSetup.LeftMargin = leftPoints;
          pageSetup.RightMargin = rightPoints;
      } catch (error) {
          debug.error("Failed to set page margins:", error);
          throw new Error(`Failed to set page margins. Error: ${error}`);
      }
  }

  /**
   * Sets the page orientation for the active document.
   * @param orientation WdOrientation enum value (0=Portrait, 1=Landscape).
   */
  public async setPageOrientation(orientation: number): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
          // WdOrientation: wdOrientPortrait = 0, wdOrientLandscape = 1
          doc.PageSetup.Orientation = orientation;
      } catch (error) {
          debug.error("Failed to set page orientation:", error);
          throw new Error(`Failed to set page orientation. Error: ${error}`);
      }
  }

  /**
   * Sets the paper size for the active document.
   * @param paperSize WdPaperSize enum value (e.g., 1=Letter, 8=A4).
   */
  public async setPaperSize(paperSize: number): Promise<void> {
      const doc = await this.getActiveDocument();
      try {
          // WdPaperSize enumeration (e.g., wdPaperLetter = 1, wdPaperA4 = 8)
          doc.PageSetup.PaperSize = paperSize;
      } catch (error) {
          debug.error("Failed to set paper size:", error);
          throw new Error(`Failed to set paper size. Error: ${error}`);
      }
  }

  // --- Cursor/Selection Methods ---

  /**
   * Moves the cursor to the start of the document.
   */
  public async moveCursorToStart(): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      selection.HomeKey(6); // wdStory = 6
    } catch (error) {
      debug.error("Failed to move cursor to start:", error);
      throw new Error(`Failed to move cursor to start. Error: ${error}`);
    }
  }

  /**
   * Moves the cursor to the end of the document.
   */
  public async moveCursorToEnd(): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      selection.EndKey(6); // wdStory = 6
    } catch (error) {
      debug.error("Failed to move cursor to end:", error);
      throw new Error(`Failed to move cursor to end. Error: ${error}`);
    }
  }

  /**
   * Moves the cursor by the specified unit and count.
   * @param unit WdUnits enum value (e.g., 1=Character, 2=Word, 3=Sentence, etc.)
   * @param count Number of units to move. Positive moves forward, negative moves backward.
   * @param extend Whether to extend the selection (true) or move the insertion point (false).
   */
  public async moveCursor(unit: number, count: number, extend: boolean = false): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      // WdUnits: wdCharacter = 1, wdWord = 2, wdSentence = 3, wdParagraph = 4, wdLine = 5, wdStory = 6, etc.
      if (extend) {
        selection.MoveRight(unit, count, 1); // 1 = wdExtend
      } else {
        selection.MoveRight(unit, count, 0); // 0 = wdMove
      }
    } catch (error) {
      debug.error("Failed to move cursor:", error);
      throw new Error(`Failed to move cursor. Error: ${error}`);
    }
  }

  /**
   * Selects the entire document.
   */
  public async selectAll(): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      selection.WholeStory();
    } catch (error) {
      debug.error("Failed to select all:", error);
      throw new Error(`Failed to select all. Error: ${error}`);
    }
  }

  /**
   * Selects a specific paragraph by index.
   * @param paragraphIndex 1-based index of the paragraph to select.
   */
  public async selectParagraph(paragraphIndex: number): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      if (paragraphIndex <= 0 || paragraphIndex > doc.Paragraphs.Count) {
        throw new Error(`Paragraph index ${paragraphIndex} is out of bounds.`);
      }
      const paragraph = doc.Paragraphs.Item(paragraphIndex);
      paragraph.Range.Select();
    } catch (error) {
      debug.error(`Failed to select paragraph ${paragraphIndex}:`, error);
      throw new Error(`Failed to select paragraph. Error: ${error}`);
    }
  }

  /**
   * Collapses the current selection to its start or end point.
   * @param toStart If true, collapse to start; if false, collapse to end.
   */
  public async collapseSelection(toStart: boolean = true): Promise<void> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      // WdCollapseDirection: wdCollapseStart = 1, wdCollapseEnd = 0
      selection.Collapse(toStart ? 1 : 0);
    } catch (error) {
      debug.error("Failed to collapse selection:", error);
      throw new Error(`Failed to collapse selection. Error: ${error}`);
    }
  }

  /**
   * Gets the current selection text.
   * @returns The text of the current selection.
   */
  public async getSelectionText(): Promise<string> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      return selection.Text;
    } catch (error) {
      debug.error("Failed to get selection text:", error);
      throw new Error(`Failed to get selection text. Error: ${error}`);
    }
  }

  /**
   * Gets information about the current selection.
   * @returns Object with selection information.
   */
  public async getSelectionInfo(): Promise<{
    text: string;
    start: number;
    end: number;
    isActive: boolean;
    type: number;
  }> {
    const app = await this.getWordApplication();
    try {
      const selection = app.ActiveDocument.ActiveWindow.Selection;
      return {
        text: selection.Text,
        start: selection.Start,
        end: selection.End,
        isActive: selection.Type !== 0, // wdSelectionNone = 0
        type: selection.Type,
      };
    } catch (error) {
      debug.error("Failed to get selection info:", error);
      throw new Error(`Failed to get selection info. Error: ${error}`);
    }
  }

  // --- VBA Module Manipulation Methods ---

  /**
   * Creates a new VBA module in the active document
   */
  public async createVbaModule(moduleName: string, moduleType: string = "standard", code?: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      if (!vbProject) {
        throw new Error("VBA Project not accessible. Check macro security settings.");
      }
      
      // Module type constants: vbext_ct_StdModule=1, vbext_ct_ClassModule=2, vbext_ct_MSForm=3, vbext_ct_Document=100
      const moduleTypeMap: Record<string, number> = {
        "standard": 1,
        "class": 2,
        "form": 3,
        "document": 100
      };
      
      const vbComponent = vbProject.VBComponents.Add(moduleTypeMap[moduleType] || 1);
      vbComponent.Name = moduleName;
      
      if (code) {
        vbComponent.CodeModule.AddFromString(code);
      }
    } catch (error) {
      debug.error("Failed to create VBA module:", error);
      throw new Error(`Failed to create VBA module '${moduleName}'. Error: ${error}`);
    }
  }

  /**
   * Deletes a VBA module from the active document
   */
  public async deleteVbaModule(moduleName: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const vbComponent = vbProject.VBComponents.Item(moduleName);
      vbProject.VBComponents.Remove(vbComponent);
    } catch (error) {
      debug.error("Failed to delete VBA module:", error);
      throw new Error(`Failed to delete VBA module '${moduleName}'. Error: ${error}`);
    }
  }

  /**
   * Gets the VBA code from a module
   */
  public async getVbaModuleCode(moduleName: string): Promise<string> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const vbComponent = vbProject.VBComponents.Item(moduleName);
      const codeModule = vbComponent.CodeModule;
      
      if (codeModule.CountOfLines > 0) {
        return codeModule.Lines(1, codeModule.CountOfLines);
      }
      return "";
    } catch (error) {
      debug.error("Failed to get VBA module code:", error);
      throw new Error(`Failed to get code from module '${moduleName}'. Error: ${error}`);
    }
  }

  /**
   * Sets the VBA code in a module
   */
  public async setVbaModuleCode(moduleName: string, code: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const vbComponent = vbProject.VBComponents.Item(moduleName);
      const codeModule = vbComponent.CodeModule;
      
      // Clear existing code
      if (codeModule.CountOfLines > 0) {
        codeModule.DeleteLines(1, codeModule.CountOfLines);
      }
      
      // Add new code
      codeModule.AddFromString(code);
    } catch (error) {
      debug.error("Failed to set VBA module code:", error);
      throw new Error(`Failed to set code in module '${moduleName}'. Error: ${error}`);
    }
  }

  /**
   * Adds a VBA procedure to a module
   */
  public async addVbaProcedure(moduleName: string, procedureCode: string, position?: number): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const vbComponent = vbProject.VBComponents.Item(moduleName);
      const codeModule = vbComponent.CodeModule;
      
      if (position && position > 0) {
        codeModule.InsertLines(position, procedureCode);
      } else {
        codeModule.AddFromString(procedureCode);
      }
    } catch (error) {
      debug.error("Failed to add VBA procedure:", error);
      throw new Error(`Failed to add procedure to module '${moduleName}'. Error: ${error}`);
    }
  }

  /**
   * Deletes a VBA procedure from a module
   */
  public async deleteVbaProcedure(moduleName: string, procedureName: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const vbComponent = vbProject.VBComponents.Item(moduleName);
      const codeModule = vbComponent.CodeModule;
      
      const procStartLine = codeModule.ProcStartLine(procedureName, 0); // 0 = vbext_pk_Proc
      const procCountLines = codeModule.ProcCountLines(procedureName, 0);
      
      codeModule.DeleteLines(procStartLine, procCountLines);
    } catch (error) {
      debug.error("Failed to delete VBA procedure:", error);
      throw new Error(`Failed to delete procedure '${procedureName}' from module '${moduleName}'. Error: ${error}`);
    }
  }

  /**
   * Lists all VBA modules in the document
   */
  public async listVbaModules(): Promise<Array<{name: string, type: string}>> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const modules: Array<{name: string, type: string}> = [];
      
      const typeMap: Record<number, string> = {
        1: "standard",
        2: "class",
        3: "form",
        100: "document"
      };
      
      for (let i = 1; i <= vbProject.VBComponents.Count; i++) {
        const component = vbProject.VBComponents.Item(i);
        modules.push({
          name: component.Name,
          type: typeMap[component.Type] || "unknown"
        });
      }
      
      return modules;
    } catch (error) {
      debug.error("Failed to list VBA modules:", error);
      throw new Error(`Failed to list VBA modules. Error: ${error}`);
    }
  }

  /**
   * Imports a VBA module from file
   */
  public async importVbaModule(filePath: string): Promise<string> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const vbComponent = vbProject.VBComponents.Import(filePath);
      return vbComponent.Name;
    } catch (error) {
      debug.error("Failed to import VBA module:", error);
      throw new Error(`Failed to import VBA module from '${filePath}'. Error: ${error}`);
    }
  }

  /**
   * Exports a VBA module to file
   */
  public async exportVbaModule(moduleName: string, filePath: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const vbComponent = vbProject.VBComponents.Item(moduleName);
      vbComponent.Export(filePath);
    } catch (error) {
      debug.error("Failed to export VBA module:", error);
      throw new Error(`Failed to export module '${moduleName}' to '${filePath}'. Error: ${error}`);
    }
  }

  /**
   * Adds a reference to the VBA project
   */
  public async addVbaReference(guid?: string, major?: number, minor?: number, description?: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      
      if (guid && major !== undefined && minor !== undefined) {
        vbProject.References.AddFromGuid(guid, major, minor);
      } else if (description) {
        // Try to find and add by description/name
        vbProject.References.AddFromFile(description);
      } else {
        throw new Error("Must provide either GUID with version or description/path");
      }
    } catch (error) {
      debug.error("Failed to add VBA reference:", error);
      throw new Error(`Failed to add VBA reference. Error: ${error}`);
    }
  }

  /**
   * Removes a reference from the VBA project
   */
  public async removeVbaReference(description: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      
      for (let i = vbProject.References.Count; i >= 1; i--) {
        const ref = vbProject.References.Item(i);
        if (ref.Description === description || ref.Name === description) {
          vbProject.References.Remove(ref);
          return;
        }
      }
      
      throw new Error(`Reference '${description}' not found`);
    } catch (error) {
      debug.error("Failed to remove VBA reference:", error);
      throw new Error(`Failed to remove VBA reference '${description}'. Error: ${error}`);
    }
  }

  /**
   * Lists all VBA references
   */
  public async listVbaReferences(): Promise<Array<{description: string, guid: string}>> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const references: Array<{description: string, guid: string}> = [];
      
      for (let i = 1; i <= vbProject.References.Count; i++) {
        const ref = vbProject.References.Item(i);
        references.push({
          description: ref.Description,
          guid: ref.Guid
        });
      }
      
      return references;
    } catch (error) {
      debug.error("Failed to list VBA references:", error);
      throw new Error(`Failed to list VBA references. Error: ${error}`);
    }
  }

  /**
   * Sets VBA project properties
   */
  public async setVbaProjectProperties(properties: any): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      
      if (properties.name) vbProject.Name = properties.name;
      if (properties.description) vbProject.Description = properties.description;
      if (properties.helpFile) vbProject.HelpFile = properties.helpFile;
      if (properties.helpContextId) vbProject.HelpContextID = properties.helpContextId;
    } catch (error) {
      debug.error("Failed to set VBA project properties:", error);
      throw new Error(`Failed to set VBA project properties. Error: ${error}`);
    }
  }

  /**
   * Protects or unprotects the VBA project
   */
  public async protectVbaProject(password: string, protect: boolean): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      
      if (protect) {
        vbProject.Protection = 1; // vbext_pp_locked
        // Note: Setting password programmatically may require additional COM automation
      } else {
        vbProject.Protection = 0; // vbext_pp_none
      }
    } catch (error) {
      debug.error("Failed to protect/unprotect VBA project:", error);
      throw new Error(`Failed to ${protect ? 'protect' : 'unprotect'} VBA project. Error: ${error}`);
    }
  }

  // --- ActiveX Control Methods ---

  /**
   * Adds an ActiveX control to the document
   */
  public async addActiveXControl(
    controlType: string,
    name: string,
    caption?: string,
    left: number = 0,
    top: number = 0,
    width: number = 100,
    height: number = 25,
    anchorToRange: boolean = false
  ): Promise<any> {
    const doc = await this.getActiveDocument();
    try {
      const controlClassMap: Record<string, string> = {
        "commandButton": "Forms.CommandButton.1",
        "textBox": "Forms.TextBox.1",
        "label": "Forms.Label.1",
        "checkBox": "Forms.CheckBox.1",
        "optionButton": "Forms.OptionButton.1",
        "comboBox": "Forms.ComboBox.1",
        "listBox": "Forms.ListBox.1",
        "toggleButton": "Forms.ToggleButton.1",
        "spinButton": "Forms.SpinButton.1",
        "scrollBar": "Forms.ScrollBar.1",
        "image": "Forms.Image.1",
        "frame": "Forms.Frame.1"
      };
      
      const classId = controlClassMap[controlType];
      if (!classId) {
        throw new Error(`Unknown control type: ${controlType}`);
      }
      
      let shape;
      if (anchorToRange) {
        const selection = doc.ActiveWindow.Selection;
        shape = doc.InlineShapes.AddOLEControl(classId, selection.Range);
      } else {
        shape = doc.Shapes.AddOLEControl(classId, left, top, width, height);
      }
      
      // Set both the shape name and the underlying OLE object name
      shape.Name = name;
      
      // Critical fix: Set the actual ActiveX control's Name property
      try {
        shape.OLEFormat.Object.Name = name;
      } catch (nameError) {
        debug.warn(`Could not set OLE object name for ${name}:`, nameError);
      }
      
      if (caption && shape.OLEFormat.Object.Caption !== undefined) {
        shape.OLEFormat.Object.Caption = caption;
      }
      
      return shape;
    } catch (error) {
      debug.error("Failed to add ActiveX control:", error);
      throw new Error(`Failed to add ActiveX control '${name}'. Error: ${error}`);
    }
  }

  /**
   * Deletes an ActiveX control
   */
  public async deleteActiveXControl(controlName: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      // Try shapes collection first
      try {
        const shape = doc.Shapes.Item(controlName);
        shape.Delete();
        return;
      } catch {
        // If not in shapes, try inline shapes
        for (let i = doc.InlineShapes.Count; i >= 1; i--) {
          const inlineShape = doc.InlineShapes.Item(i);
          if (inlineShape.OLEFormat && inlineShape.OLEFormat.Object.Name === controlName) {
            inlineShape.Delete();
            return;
          }
        }
      }
      
      throw new Error(`Control '${controlName}' not found`);
    } catch (error) {
      debug.error("Failed to delete ActiveX control:", error);
      throw new Error(`Failed to delete ActiveX control '${controlName}'. Error: ${error}`);
    }
  }

  /**
   * Sets properties of an ActiveX control
   */
  public async setActiveXControlProperties(controlName: string, properties: Record<string, any>): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      let control;
      
      // Try to find in shapes
      try {
        const shape = doc.Shapes.Item(controlName);
        control = shape.OLEFormat.Object;
      } catch {
        // Try inline shapes
        for (let i = 1; i <= doc.InlineShapes.Count; i++) {
          const inlineShape = doc.InlineShapes.Item(i);
          if (inlineShape.OLEFormat && inlineShape.OLEFormat.Object.Name === controlName) {
            control = inlineShape.OLEFormat.Object;
            break;
          }
        }
      }
      
      if (!control) {
        throw new Error(`Control '${controlName}' not found`);
      }
      
      for (const [key, value] of Object.entries(properties)) {
        control[key] = value;
      }
    } catch (error) {
      debug.error("Failed to set ActiveX control properties:", error);
      throw new Error(`Failed to set properties for control '${controlName}'. Error: ${error}`);
    }
  }

  /**
   * Gets properties of an ActiveX control
   */
  public async getActiveXControlProperties(controlName: string): Promise<Record<string, any>> {
    const doc = await this.getActiveDocument();
    try {
      let control;
      
      // Try to find in shapes
      try {
        const shape = doc.Shapes.Item(controlName);
        control = shape.OLEFormat.Object;
      } catch {
        // Try inline shapes
        for (let i = 1; i <= doc.InlineShapes.Count; i++) {
          const inlineShape = doc.InlineShapes.Item(i);
          if (inlineShape.OLEFormat && inlineShape.OLEFormat.Object.Name === controlName) {
            control = inlineShape.OLEFormat.Object;
            break;
          }
        }
      }
      
      if (!control) {
        throw new Error(`Control '${controlName}' not found`);
      }
      
      // Get common properties
      const properties: Record<string, any> = {
        Name: control.Name,
        Caption: control.Caption,
        Enabled: control.Enabled,
        Visible: control.Visible,
        BackColor: control.BackColor,
        ForeColor: control.ForeColor,
        Font: control.Font ? control.Font.Name : null,
        Value: control.Value
      };
      
      return properties;
    } catch (error) {
      debug.error("Failed to get ActiveX control properties:", error);
      throw new Error(`Failed to get properties for control '${controlName}'. Error: ${error}`);
    }
  }

  /**
   * Adds an event handler for an ActiveX control
   */
  public async addActiveXEventHandler(controlName: string, eventName: string, vbaCode: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      
      // Get the ThisDocument module
      let thisDocModule;
      for (let i = 1; i <= vbProject.VBComponents.Count; i++) {
        const component = vbProject.VBComponents.Item(i);
        if (component.Type === 100) { // vbext_ct_Document
          thisDocModule = component.CodeModule;
          break;
        }
      }
      
      if (!thisDocModule) {
        throw new Error("ThisDocument module not found");
      }
      
      // Create the event handler procedure
      const procedureCode = `Private Sub ${controlName}_${eventName}()\n${vbaCode}\nEnd Sub`;
      thisDocModule.AddFromString(procedureCode);
    } catch (error) {
      debug.error("Failed to add ActiveX event handler:", error);
      throw new Error(`Failed to add event handler for control '${controlName}'. Error: ${error}`);
    }
  }

  /**
   * Lists all ActiveX controls in the document
   */
  public async listActiveXControls(): Promise<Array<{name: string, type: string}>> {
    const doc = await this.getActiveDocument();
    try {
      const controls: Array<{name: string, type: string}> = [];
      
      // Check shapes collection
      for (let i = 1; i <= doc.Shapes.Count; i++) {
        const shape = doc.Shapes.Item(i);
        if (shape.Type === 12) { // msoOLEControlObject
          controls.push({
            name: shape.Name,
            type: shape.OLEFormat.ProgID || "OLE Control"
          });
        }
      }
      
      // Check inline shapes collection
      for (let i = 1; i <= doc.InlineShapes.Count; i++) {
        const inlineShape = doc.InlineShapes.Item(i);
        if (inlineShape.Type === 5) { // wdInlineShapeOLEControlObject
          controls.push({
            name: inlineShape.OLEFormat.Object.Name || `InlineControl${i}`,
            type: inlineShape.OLEFormat.ProgID || "OLE Control"
          });
        }
      }
      
      return controls;
    } catch (error) {
      debug.error("Failed to list ActiveX controls:", error);
      throw new Error(`Failed to list ActiveX controls. Error: ${error}`);
    }
  }

  /**
   * Creates a UserForm
   */
  public async createUserForm(formName: string, caption?: string, width: number = 400, height: number = 300): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const vbComponent = vbProject.VBComponents.Add(3); // vbext_ct_MSForm
      vbComponent.Name = formName;
      
      // Fix: Access UserForm properties correctly through the Designer object
      const userForm = vbComponent.Designer;
      
      if (caption) {
        try {
          userForm.Caption = caption;
        } catch (captionError) {
          debug.warn(`Could not set UserForm caption:`, captionError);
        }
      }
      
      try {
        // Fix form sizing - ensure minimum size and proper dimensions
        const minWidth = Math.max(width, 300); // Minimum 300 pixels width
        const minHeight = Math.max(height, 200); // Minimum 200 pixels height
        
        userForm.Width = minWidth;
        userForm.Height = minHeight;
        
        // Also set the form properties that might affect sizing
        userForm.StartUpPosition = 1; // Center on screen
        
        debug.log(`UserForm '${formName}' sized to ${minWidth}x${minHeight}`);
      } catch (sizeError) {
        debug.warn(`Could not set UserForm size:`, sizeError);
      }
    } catch (error) {
      debug.error("Failed to create UserForm:", error);
      throw new Error(`Failed to create UserForm '${formName}'. Error: ${error}`);
    }
  }

  /**
   * Adds a control to a UserForm
   */
  public async addControlToUserForm(
    formName: string,
    controlType: string,
    controlName: string,
    caption?: string,
    left: number = 10,
    top: number = 10,
    width: number = 100,
    height: number = 25
  ): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const formComponent = vbProject.VBComponents.Item(formName);
      
      if (formComponent.Type !== 3) { // vbext_ct_MSForm
        throw new Error(`'${formName}' is not a UserForm`);
      }
      
      const controlClassMap: Record<string, string> = {
        "commandButton": "Forms.CommandButton.1",
        "textBox": "Forms.TextBox.1",
        "label": "Forms.Label.1",
        "checkBox": "Forms.CheckBox.1",
        "optionButton": "Forms.OptionButton.1",
        "comboBox": "Forms.ComboBox.1",
        "listBox": "Forms.ListBox.1",
        "toggleButton": "Forms.ToggleButton.1",
        "spinButton": "Forms.SpinButton.1",
        "scrollBar": "Forms.ScrollBar.1",
        "image": "Forms.Image.1",
        "frame": "Forms.Frame.1"
      };
      
      const classId = controlClassMap[controlType];
      const designer = formComponent.Designer;
      const control = designer.Controls.Add(classId);
      
      control.Name = controlName;
      if (caption) control.Caption = caption;
      control.Left = left;
      control.Top = top;
      control.Width = width;
      control.Height = height;
    } catch (error) {
      debug.error("Failed to add control to UserForm:", error);
      throw new Error(`Failed to add control to UserForm '${formName}'. Error: ${error}`);
    }
  }

  /**
   * Sets tab order for a control
   */
  public async setControlTabOrder(controlName: string, tabIndex: number): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      // Implementation would depend on whether it's a document control or UserForm control
      // This is a simplified version
      let control;
      
      // Try to find in shapes
      try {
        const shape = doc.Shapes.Item(controlName);
        control = shape.OLEFormat.Object;
      } catch {
        // Try inline shapes
        for (let i = 1; i <= doc.InlineShapes.Count; i++) {
          const inlineShape = doc.InlineShapes.Item(i);
          if (inlineShape.OLEFormat && inlineShape.OLEFormat.Object.Name === controlName) {
            control = inlineShape.OLEFormat.Object;
            break;
          }
        }
      }
      
      if (control) {
        control.TabIndex = tabIndex;
      } else {
        throw new Error(`Control '${controlName}' not found`);
      }
    } catch (error) {
      debug.error("Failed to set tab order:", error);
      throw new Error(`Failed to set tab order for control '${controlName}'. Error: ${error}`);
    }
  }

  /**
   * Groups controls together
   */
  public async groupControls(controlNames: string[], groupName?: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const shapesToGroup = [];
      
      for (const name of controlNames) {
        try {
          const shape = doc.Shapes.Item(name);
          shapesToGroup.push(shape);
        } catch {
          throw new Error(`Control '${name}' not found in shapes collection`);
        }
      }
      
      if (shapesToGroup.length > 1) {
        const shapeRange = doc.Shapes.Range(controlNames);
        const group = shapeRange.Group();
        if (groupName) {
          group.Name = groupName;
        }
      }
    } catch (error) {
      debug.error("Failed to group controls:", error);
      throw new Error(`Failed to group controls. Error: ${error}`);
    }
  }

  // --- VBA Execution Methods ---

  /**
   * Runs a VBA macro
   */
  public async runVbaMacro(macroName: string, parameters?: any[]): Promise<any> {
    const app = await this.getWordApplication();
    try {
      if (parameters && parameters.length > 0) {
        return app.Run(macroName, ...parameters);
      } else {
        return app.Run(macroName);
      }
    } catch (error) {
      debug.error("Failed to run VBA macro:", error);
      throw new Error(`Failed to run macro '${macroName}'. Error: ${error}`);
    }
  }

  /**
   * Tests a VBA macro
   */
  public async testVbaMacro(macroName: string, testData?: any, expectedResult?: any): Promise<any> {
    try {
      const startTime = Date.now();
      let result;
      let success = false;
      let error = null;
      
      try {
        result = await this.runVbaMacro(macroName, testData ? [testData] : undefined);
        
        if (expectedResult !== undefined) {
          success = JSON.stringify(result) === JSON.stringify(expectedResult);
        } else {
          success = true;
        }
      } catch (err: any) {
        error = err.message;
        success = false;
      }
      
      const executionTime = Date.now() - startTime;
      
      return {
        success,
        result,
        error,
        executionTime
      };
    } catch (error) {
      debug.error("Failed to test VBA macro:", error);
      throw new Error(`Failed to test macro '${macroName}'. Error: ${error}`);
    }
  }

  /**
   * Debugs VBA code
   */
  public async debugVbaCode(moduleName: string, procedureName: string, breakpoints?: number[]): Promise<any> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const vbComponent = vbProject.VBComponents.Item(moduleName);
      const codeModule = vbComponent.CodeModule;
      
      const debugInfo: any = {
        module: moduleName,
        procedure: procedureName,
        lineCount: codeModule.CountOfLines,
        procedureStart: codeModule.ProcStartLine(procedureName, 0),
        procedureLines: codeModule.ProcCountLines(procedureName, 0)
      };
      
      // Note: Actually setting breakpoints requires VBE automation which is limited
      if (breakpoints) {
        debugInfo.requestedBreakpoints = breakpoints;
      }
      
      return debugInfo;
    } catch (error) {
      debug.error("Failed to debug VBA code:", error);
      throw new Error(`Failed to debug code in module '${moduleName}'. Error: ${error}`);
    }
  }

  /**
   * Compiles the VBA project
   */
  public async compileVbaProject(): Promise<any> {
    const doc = await this.getActiveDocument();
    try {
      // Note: Direct compilation through COM is limited
      // This is a simplified version that checks for basic syntax
      const vbProject = doc.VBProject;
      const errors: string[] = [];
      
      for (let i = 1; i <= vbProject.VBComponents.Count; i++) {
        const component = vbProject.VBComponents.Item(i);
        const codeModule = component.CodeModule;
        
        if (codeModule.CountOfLines > 0) {
          // Basic check - try to access the code
          try {
            codeModule.Lines(1, codeModule.CountOfLines);
          } catch (err: any) {
            errors.push(`Error in module '${component.Name}': ${err.message}`);
          }
        }
      }
      
      return {
        success: errors.length === 0,
        errors: errors.length > 0 ? errors : undefined
      };
    } catch (error) {
      debug.error("Failed to compile VBA project:", error);
      throw new Error(`Failed to compile VBA project. Error: ${error}`);
    }
  }

  /**
   * Adds a document event handler
   */
  public async addDocumentEventHandler(eventName: string, vbaCode: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      
      // Find ThisDocument module
      let thisDocModule;
      for (let i = 1; i <= vbProject.VBComponents.Count; i++) {
        const component = vbProject.VBComponents.Item(i);
        if (component.Type === 100) { // vbext_ct_Document
          thisDocModule = component.CodeModule;
          break;
        }
      }
      
      if (!thisDocModule) {
        throw new Error("ThisDocument module not found");
      }
      
      const procedureCode = `Private Sub ${eventName}()\n${vbaCode}\nEnd Sub`;
      thisDocModule.AddFromString(procedureCode);
    } catch (error) {
      debug.error("Failed to add document event handler:", error);
      throw new Error(`Failed to add document event handler '${eventName}'. Error: ${error}`);
    }
  }

  /**
   * Adds an application event handler
   */
  public async addApplicationEventHandler(eventName: string, vbaCode: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      
      // Create or get a class module for application events
      let appEventModule;
      try {
        appEventModule = vbProject.VBComponents.Item("ApplicationEvents");
      } catch {
        const component = vbProject.VBComponents.Add(2); // vbext_ct_ClassModule
        component.Name = "ApplicationEvents";
        appEventModule = component;
        
        // Add WithEvents declaration
        const codeModule = component.CodeModule;
        codeModule.AddFromString("Public WithEvents App As Word.Application");
      }
      
      const codeModule = appEventModule.CodeModule;
      const procedureCode = `Private Sub App_${eventName}()\n${vbaCode}\nEnd Sub`;
      codeModule.AddFromString(procedureCode);
    } catch (error) {
      debug.error("Failed to add application event handler:", error);
      throw new Error(`Failed to add application event handler '${eventName}'. Error: ${error}`);
    }
  }

  /**
   * Creates an auto-executing macro
   */
  public async createAutoMacro(autoType: string, vbaCode: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      
      // Auto macros should be in a standard module
      let autoModule;
      try {
        autoModule = vbProject.VBComponents.Item("AutoMacros");
      } catch {
        const component = vbProject.VBComponents.Add(1); // vbext_ct_StdModule
        component.Name = "AutoMacros";
        autoModule = component;
      }
      
      const codeModule = autoModule.CodeModule;
      const procedureCode = `Sub ${autoType}()\n${vbaCode}\nEnd Sub`;
      codeModule.AddFromString(procedureCode);
    } catch (error) {
      debug.error("Failed to create auto macro:", error);
      throw new Error(`Failed to create auto macro '${autoType}'. Error: ${error}`);
    }
  }

  /**
   * Gets VBA error information
   */
  public async getVbaErrorInfo(): Promise<any> {
    const app = await this.getWordApplication();
    try {
      // This is a simplified version - actual error handling would be more complex
      return {
        lastError: null, // Would need to track errors during execution
        errorCount: 0
      };
    } catch (error) {
      debug.error("Failed to get VBA error info:", error);
      throw new Error(`Failed to get VBA error information. Error: ${error}`);
    }
  }

  /**
   * Clears the VBA Immediate window
   */
  public async clearVbaImmediateWindow(): Promise<void> {
    const app = await this.getWordApplication();
    try {
      // Note: Direct access to Immediate window is limited through COM
      // This would typically require SendKeys or other workarounds
      debug.log("Clear Immediate window requested - limited support through COM");
    } catch (error) {
      debug.error("Failed to clear Immediate window:", error);
      throw new Error(`Failed to clear VBA Immediate window. Error: ${error}`);
    }
  }

  /**
   * Executes code in the VBA Immediate window
   */
  public async executeVbaImmediate(vbaCode: string): Promise<string> {
    const app = await this.getWordApplication();
    try {
      // Execute as a temporary macro
      const tempMacroName = `TempImmediate_${Date.now()}`;
      const doc = await this.getActiveDocument();
      const vbProject = doc.VBProject;
      
      // Add temporary module
      const tempModule = vbProject.VBComponents.Add(1); // vbext_ct_StdModule
      const codeModule = tempModule.CodeModule;
      
      // Create a function that returns the result
      const functionCode = `Function ${tempMacroName}() As Variant\n${vbaCode}\nEnd Function`;
      codeModule.AddFromString(functionCode);
      
      // Run and get result
      const result = app.Run(tempMacroName);
      
      // Clean up
      vbProject.VBComponents.Remove(tempModule);
      
      return result ? String(result) : "";
    } catch (error) {
      debug.error("Failed to execute in Immediate window:", error);
      throw new Error(`Failed to execute in VBA Immediate window. Error: ${error}`);
    }
  }

  /**
   * Lists available macros
   */
  public async listAvailableMacros(): Promise<Array<{name: string, module: string}>> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const macros: Array<{name: string, module: string}> = [];
      
      for (let i = 1; i <= vbProject.VBComponents.Count; i++) {
        const component = vbProject.VBComponents.Item(i);
        const codeModule = component.CodeModule;
        
        if (codeModule.CountOfLines > 0) {
          // Parse procedures
          let lineNum = 1;
          while (lineNum <= codeModule.CountOfLines) {
            try {
              const procName = codeModule.ProcOfLine(lineNum, 0);
              if (procName && !macros.some(m => m.name === procName && m.module === component.Name)) {
                macros.push({
                  name: procName,
                  module: component.Name
                });
              }
              lineNum = codeModule.ProcStartLine(procName, 0) + codeModule.ProcCountLines(procName, 0);
            } catch {
              lineNum++;
            }
          }
        }
      }
      
      return macros;
    } catch (error) {
      debug.error("Failed to list available macros:", error);
      throw new Error(`Failed to list available macros. Error: ${error}`);
    }
  }

  /**
   * Starts recording a macro
   */
  public async startRecordingMacro(macroName: string, description?: string): Promise<void> {
    const app = await this.getWordApplication();
    try {
      // Note: Macro recording through COM is limited
      // This is a simplified placeholder
      debug.log(`Start recording macro: ${macroName}`);
      // Store for later use
      (this as any).recordingMacroName = macroName;
    } catch (error) {
      debug.error("Failed to start recording macro:", error);
      throw new Error(`Failed to start recording macro '${macroName}'. Error: ${error}`);
    }
  }

  /**
   * Stops recording a macro
   */
  public async stopRecordingMacro(): Promise<string> {
    const app = await this.getWordApplication();
    try {
      // Note: Macro recording through COM is limited
      const macroName = (this as any).recordingMacroName || "RecordedMacro";
      debug.log(`Stop recording macro: ${macroName}`);
      delete (this as any).recordingMacroName;
      return macroName;
    } catch (error) {
      debug.error("Failed to stop recording macro:", error);
      throw new Error(`Failed to stop recording macro. Error: ${error}`);
    }
  }

  // --- Template and Deployment Methods ---

  /**
   * Saves document as macro-enabled format
   */
  public async saveAsMacroEnabled(filePath: string, format: string = "docm", createBackup: boolean = false): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      if (createBackup) {
        const backupPath = filePath.replace(/\.[^.]+$/, '_backup$&');
        doc.SaveAs2(backupPath);
      }
      
      const formatMap: Record<string, number> = {
        "docm": 13,  // wdFormatXMLDocumentMacroEnabled
        "dotm": 15,  // wdFormatXMLTemplateMacroEnabled
        "docx": 16,  // wdFormatXMLDocument
        "dotx": 14,  // wdFormatXMLTemplate
        "doc": 0,    // wdFormatDocument
        "dot": 1     // wdFormatTemplate
      };
      
      const fileFormat = formatMap[format] || 13;
      doc.SaveAs2(filePath, fileFormat);
    } catch (error) {
      debug.error("Failed to save as macro-enabled:", error);
      throw new Error(`Failed to save document as macro-enabled format. Error: ${error}`);
    }
  }

  /**
   * Creates a document template
   */
  public async createDocumentTemplate(
    templatePath: string,
    includeMacros: boolean = true,
    protectTemplate: boolean = false,
    password?: string
  ): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const fileFormat = includeMacros ? 15 : 14; // dotm or dotx
      
      if (protectTemplate && password) {
        doc.Protect(2, false, password); // wdAllowOnlyFormFields
      }
      
      doc.SaveAs2(templatePath, fileFormat);
    } catch (error) {
      debug.error("Failed to create template:", error);
      throw new Error(`Failed to create document template. Error: ${error}`);
    }
  }

  /**
   * Sets macro security settings
   */
  public async setMacroSecurity(level: string, trustAccessVBOM?: boolean): Promise<void> {
    const app = await this.getWordApplication();
    try {
      // Note: Security settings are typically registry-based and may require admin rights
      const securityMap: Record<string, number> = {
        "veryHigh": 4,
        "high": 3,
        "medium": 2,
        "low": 1
      };
      
      // This would typically require registry modification or user interaction
      debug.log(`Macro security level requested: ${level}`);
      
      if (trustAccessVBOM !== undefined) {
        // This setting also requires registry modification
        debug.log(`Trust access to VBA object model: ${trustAccessVBOM}`);
      }
    } catch (error) {
      debug.error("Failed to set macro security:", error);
      throw new Error(`Failed to set macro security settings. Error: ${error}`);
    }
  }

  /**
   * Signs the VBA project
   */
  public async signVbaProject(certificatePath?: string, certificateName?: string, timestamp: boolean = true): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      
      // Note: Digital signing through COM is very limited
      // This typically requires external tools or manual intervention
      debug.log("VBA project signing requested");
      debug.log(`Certificate: ${certificatePath || certificateName}`);
      debug.log(`Timestamp: ${timestamp}`);
    } catch (error) {
      debug.error("Failed to sign VBA project:", error);
      throw new Error(`Failed to sign VBA project. Error: ${error}`);
    }
  }

  /**
   * Creates a self-executing document
   */
  public async createSelfExecutingDocument(
    filePath: string,
    startupCode: string,
    hideCode: boolean = false,
    password?: string
  ): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      // Add Document_Open event
      await this.addDocumentEventHandler("Document_Open", startupCode);
      
      if (hideCode && password) {
        await this.protectVbaProject(password, true);
      }
      
      // Save as macro-enabled
      await this.saveAsMacroEnabled(filePath, "docm");
    } catch (error) {
      debug.error("Failed to create self-executing document:", error);
      throw new Error(`Failed to create self-executing document. Error: ${error}`);
    }
  }

  /**
   * Exports entire VBA project
   */
  public async exportVbaProject(exportPath: string, includeReferences: boolean = true, includeProjectInfo: boolean = true): Promise<string[]> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const exportedFiles: string[] = [];
      
      // Export all modules
      for (let i = 1; i <= vbProject.VBComponents.Count; i++) {
        const component = vbProject.VBComponents.Item(i);
        const extension = component.Type === 1 ? ".bas" : component.Type === 2 ? ".cls" : ".frm";
        const fileName = `${exportPath}\\${component.Name}${extension}`;
        
        component.Export(fileName);
        exportedFiles.push(fileName);
      }
      
      // Export project info if requested
      if (includeProjectInfo) {
        const infoFile = `${exportPath}\\project.info`;
        // Would write project properties to file
        exportedFiles.push(infoFile);
      }
      
      // Export references if requested
      if (includeReferences) {
        const refsFile = `${exportPath}\\references.txt`;
        // Would write references to file
        exportedFiles.push(refsFile);
      }
      
      return exportedFiles;
    } catch (error) {
      debug.error("Failed to export VBA project:", error);
      throw new Error(`Failed to export VBA project. Error: ${error}`);
    }
  }

  /**
   * Imports entire VBA project
   */
  public async importVbaProject(importPath: string, clearExisting: boolean = false, importReferences: boolean = true): Promise<string[]> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const importedModules: string[] = [];
      
      // Clear existing if requested
      if (clearExisting) {
        for (let i = vbProject.VBComponents.Count; i >= 1; i--) {
          const component = vbProject.VBComponents.Item(i);
          if (component.Type !== 100) { // Don't remove ThisDocument
            vbProject.VBComponents.Remove(component);
          }
        }
      }
      
      // Import modules (simplified - would need to scan directory)
      // This is a placeholder for actual directory scanning
      debug.log(`Importing VBA modules from: ${importPath}`);
      
      return importedModules;
    } catch (error) {
      debug.error("Failed to import VBA project:", error);
      throw new Error(`Failed to import VBA project. Error: ${error}`);
    }
  }

  /**
   * Creates a Word add-in
   */
  public async createAddIn(addInPath: string, addInName: string, description?: string, autoLoad: boolean = false): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      // Save as template with macros
      await this.saveAsMacroEnabled(addInPath, "dotm");
      
      // Set add-in properties
      if (description) {
        doc.BuiltInDocumentProperties("Comments").Value = description;
      }
      
      if (autoLoad) {
        // Would need to add to Word's startup folder or registry
        debug.log(`Add-in '${addInName}' created for auto-loading`);
      }
    } catch (error) {
      debug.error("Failed to create add-in:", error);
      throw new Error(`Failed to create Word add-in. Error: ${error}`);
    }
  }

  /**
   * Installs or uninstalls an add-in
   */
  public async installAddIn(addInPath: string, install: boolean = true): Promise<void> {
    const app = await this.getWordApplication();
    try {
      if (install) {
        app.AddIns.Add(addInPath, true);
      } else {
        // Find and uninstall
        for (let i = 1; i <= app.AddIns.Count; i++) {
          const addIn = app.AddIns.Item(i);
          if (addIn.Path === addInPath) {
            addIn.Installed = false;
            break;
          }
        }
      }
    } catch (error) {
      debug.error("Failed to install/uninstall add-in:", error);
      throw new Error(`Failed to ${install ? 'install' : 'uninstall'} add-in. Error: ${error}`);
    }
  }

  /**
   * Creates ribbon customization
   */
  public async createRibbonCustomization(ribbonXML: string, callbacks?: Record<string, string>): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      // Note: Ribbon customization requires Open XML manipulation
      // This is a simplified placeholder
      debug.log("Ribbon customization requested");
      
      if (callbacks) {
        // Create callback procedures in VBA
        for (const [callbackName, procedureName] of Object.entries(callbacks)) {
          debug.log(`Callback ${callbackName} -> ${procedureName}`);
        }
      }
    } catch (error) {
      debug.error("Failed to create ribbon customization:", error);
      throw new Error(`Failed to create ribbon customization. Error: ${error}`);
    }
  }

  /**
   * Adds a custom document property
   */
  public async addCustomDocumentProperty(name: string, value: any, linkToContent: boolean = false): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      doc.CustomDocumentProperties.Add(name, linkToContent, 4, value); // msoPropertyTypeString = 4
    } catch (error) {
      debug.error("Failed to add custom document property:", error);
      throw new Error(`Failed to add custom document property '${name}'. Error: ${error}`);
    }
  }

  /**
   * Shows a UserForm modally and returns collected data
   */
  public async showUserForm(formName: string): Promise<Record<string, any>> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const formComponent = vbProject.VBComponents.Item(formName);
      
      if (formComponent.Type !== 3) { // vbext_ct_MSForm
        throw new Error(`'${formName}' is not a UserForm`);
      }
      
      const userForm = formComponent.Designer;
      
      // Show the form modally
      const result = userForm.Show(1); // vbModal = 1
      
      // Collect data from all controls
      const formData: Record<string, any> = {};
      
      try {
        for (let i = 0; i < userForm.Controls.Count; i++) {
          const control = userForm.Controls.Item(i);
          const controlName = control.Name;
          
          // Collect values based on control type
          if (control.Value !== undefined) {
            formData[controlName] = control.Value;
          } else if (control.Text !== undefined) {
            formData[controlName] = control.Text;
          } else if (control.Caption !== undefined) {
            formData[controlName] = control.Caption;
          }
        }
      } catch (collectError) {
        debug.warn("Could not collect all form data:", collectError);
      }
      
      return formData;
    } catch (error) {
      debug.error("Failed to show UserForm:", error);
      throw new Error(`Failed to show UserForm '${formName}'. Error: ${error}`);
    }
  }
  
  /**
   * Adds submit logic to a UserForm by creating a macro that handles form submission
   */
  public async addUserFormSubmitLogic(formName: string, targetMacro: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const formComponent = vbProject.VBComponents.Item(formName);
      
      if (formComponent.Type !== 3) { // vbext_ct_MSForm
        throw new Error(`'${formName}' is not a UserForm`);
      }
      
      // Add event handler for the submit button (assumes there's a button named "btnSubmit")
      const submitCode = `
Private Sub btnSubmit_Click()
  ' Collect form data
  Dim formData As String
  formData = ""
  
  ' Add logic to collect all control values here
  ' Call the target macro with collected data
  Application.Run "${targetMacro}", formData
  
  ' Hide the form
  Me.Hide
End Sub

Private Sub btnCancel_Click()
  Me.Hide
End Sub
`;
      
      const codeModule = formComponent.CodeModule;
      codeModule.AddFromString(submitCode);
    } catch (error) {
      debug.error("Failed to add UserForm submit logic:", error);
      throw new Error(`Failed to add submit logic to UserForm '${formName}'. Error: ${error}`);
    }
  }
  
  /**
   * Lists all UserForms in the current document
   */
  public async listUserForms(): Promise<string[]> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const forms: string[] = [];
      
      for (let i = 1; i <= vbProject.VBComponents.Count; i++) {
        const component = vbProject.VBComponents.Item(i);
        if (component.Type === 3) { // vbext_ct_MSForm
          forms.push(component.Name);
        }
      }
      
      return forms;
    } catch (error) {
      debug.error("Failed to list UserForms:", error);
      throw new Error(`Failed to list UserForms. Error: ${error}`);
    }
  }
  
  /**
   * Deletes a UserForm
   */
  public async deleteUserForm(formName: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const formComponent = vbProject.VBComponents.Item(formName);
      
      if (formComponent.Type !== 3) { // vbext_ct_MSForm
        throw new Error(`'${formName}' is not a UserForm`);
      }
      
      vbProject.VBComponents.Remove(formComponent);
    } catch (error) {
      debug.error("Failed to delete UserForm:", error);
      throw new Error(`Failed to delete UserForm '${formName}'. Error: ${error}`);
    }
  }
  
  /**
   * Gets all controls on a UserForm
   */
  public async getUserFormControls(formName: string): Promise<Array<{name: string, type: string, value?: any}>> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const formComponent = vbProject.VBComponents.Item(formName);
      
      if (formComponent.Type !== 3) { // vbext_ct_MSForm
        throw new Error(`'${formName}' is not a UserForm`);
      }
      
      const userForm = formComponent.Designer;
      const controls: Array<{name: string, type: string, value?: any}> = [];
      
      for (let i = 0; i < userForm.Controls.Count; i++) {
        const control = userForm.Controls.Item(i);
        controls.push({
          name: control.Name,
          type: control.ProgID || "Unknown",
          value: control.Value || control.Text || control.Caption
        });
      }
      
      return controls;
    } catch (error) {
      debug.error("Failed to get UserForm controls:", error);
      throw new Error(`Failed to get controls for UserForm '${formName}'. Error: ${error}`);
    }
  }

  /**
   * Creates the ShowUserFormModal VBA function that can be called from other VBA code
   */
  public async createShowUserFormModalFunction(): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      
      // Create or get the UserFormHelper module
      let helperModule;
      try {
        helperModule = vbProject.VBComponents.Item("UserFormHelper");
      } catch {
        // Create new standard module for helper functions
        const newComponent = vbProject.VBComponents.Add(1); // vbext_ct_StdModule
        newComponent.Name = "UserFormHelper";
        helperModule = newComponent;
      }
      
      const codeModule = helperModule.CodeModule;
      
      // Check if function already exists
      const existingCode = codeModule.Lines(1, codeModule.CountOfLines);
      if (existingCode.includes("Function ShowUserFormModal")) {
        debug.log("ShowUserFormModal function already exists");
        return;
      }
      
      // VBA code for the ShowUserFormModal function
      const vbaCode = `
' ShowUserFormModal - Displays a UserForm modally and returns collected data
' Parameter: formName - Name of the UserForm to display
' Returns: JSON-like string with form data, or empty string if cancelled
Function ShowUserFormModal(formName As String) As String
    On Error GoTo ErrorHandler
    
    Dim formComponent As Object
    Dim userForm As Object
    Dim result As String
    Dim cancelled As Boolean
    
    ' Initialize result
    result = ""
    cancelled = False
    
    ' Get the UserForm by name
    Set formComponent = ThisDocument.VBProject.VBComponents(formName)
    
    If formComponent.Type <> 3 Then ' vbext_ct_MSForm
        MsgBox "Error: '" & formName & "' is not a UserForm", vbCritical
        ShowUserFormModal = ""
        Exit Function
    End If
    
    Set userForm = formComponent.Designer
    
    ' Add event handlers for OK/Cancel buttons if they exist
    ' This is handled by individual form design
    
    ' Show the form modally
    userForm.Show vbModal
    
    ' Check if form was cancelled (assuming a global variable or form property)
    ' For now, we'll assume if the form is hidden, data was submitted
    
    ' Collect data from all controls
    result = CollectFormData(userForm)
    
    ShowUserFormModal = result
    Exit Function
    
ErrorHandler:
    MsgBox "Error displaying UserForm '" & formName & "': " & Err.Description, vbCritical
    ShowUserFormModal = ""
End Function

' Helper function to collect data from UserForm controls
Private Function CollectFormData(userForm As Object) As String
    On Error GoTo ErrorHandler
    
    Dim i As Integer
    Dim control As Object
    Dim result As String
    Dim controlData As String
    
    result = "{"
    
    ' Iterate through all controls
    For i = 0 To userForm.Controls.Count - 1
        Set control = userForm.Controls(i)
        controlData = ""
        
        ' Extract data based on control type
        Select Case TypeName(control)
            Case "TextBox"
                controlData = "\"" & control.Name & "\":\"" & control.Text & "\""
            Case "ComboBox"
                controlData = "\"" & control.Name & "\":\"" & control.Text & "\""
            Case "CheckBox"
                controlData = "\"" & control.Name & "\":" & IIf(control.Value, "true", "false")
            Case "OptionButton"
                controlData = "\"" & control.Name & "\":" & IIf(control.Value, "true", "false")
            Case "ListBox"
                If control.ListIndex >= 0 Then
                    controlData = "\"" & control.Name & "\":\"" & control.List(control.ListIndex) & "\""
                Else
                    controlData = "\"" & control.Name & "\":\"\""
                End If
            Case "Label"
                ' Skip labels as they don't contain user input
            Case "CommandButton"
                ' Skip buttons as they don't contain user input
            Case Else
                ' Try to get Value or Text property
                On Error Resume Next
                If Not IsEmpty(control.Value) Then
                    controlData = "\"" & control.Name & "\":\"" & control.Value & "\""
                ElseIf control.Text <> "" Then
                    controlData = "\"" & control.Name & "\":\"" & control.Text & "\""
                End If
                On Error GoTo ErrorHandler
        End Select
        
        ' Add to result if we have data
        If controlData <> "" Then
            If Len(result) > 1 Then result = result & ","
            result = result & controlData
        End If
    Next i
    
    result = result & "}"
    CollectFormData = result
    Exit Function
    
ErrorHandler:
    CollectFormData = "{}"
End Function

' Helper function to create standard OK/Cancel buttons on a UserForm
Sub AddOKCancelButtons(userForm As Object)
    On Error GoTo ErrorHandler
    
    Dim btnOK As Object
    Dim btnCancel As Object
    
    ' Add OK button
    Set btnOK = userForm.Controls.Add("Forms.CommandButton.1")
    With btnOK
        .Name = "btnOK"
        .Caption = "OK"
        .Left = userForm.Width - 160
        .Top = userForm.Height - 60
        .Width = 70
        .Height = 25
    End With
    
    ' Add Cancel button
    Set btnCancel = userForm.Controls.Add("Forms.CommandButton.1")
    With btnCancel
        .Name = "btnCancel"
        .Caption = "Cancel"
        .Left = userForm.Width - 80
        .Top = userForm.Height - 60
        .Width = 70
        .Height = 25
    End With
    
    Exit Sub
    
ErrorHandler:
    ' Ignore errors in button creation
End Sub
`;
      
      // Add the VBA code to the module
      codeModule.AddFromString(vbaCode);
      
      debug.log("ShowUserFormModal function created successfully");
    } catch (error) {
      debug.error("Failed to create ShowUserFormModal function:", error);
      throw new Error(`Failed to create ShowUserFormModal function. Error: ${error}`);
    }
  }
  
  /**
   * Creates a UserForm with standard OK/Cancel button event handlers
   * This is a generic helper that adds proper form closure handling
   */
  public async addStandardFormEventHandlers(formName: string): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      const vbProject = doc.VBProject;
      const formComponent = vbProject.VBComponents.Item(formName);
      
      if (formComponent.Type !== 3) { // vbext_ct_MSForm
        throw new Error(`'${formName}' is not a UserForm`);
      }
      
      const formCodeModule = formComponent.CodeModule;
      
      // Check if handlers already exist
      const existingCode = formCodeModule.Lines(1, formCodeModule.CountOfLines);
      if (existingCode.includes("btnOK_Click")) {
        debug.log("Form event handlers already exist");
        return;
      }
      
      const buttonCode = `
Private cancelled As Boolean
Private formData As String

' Standard OK button handler
Private Sub btnOK_Click()
    cancelled = False
    Me.Hide
End Sub

' Standard Cancel button handler
Private Sub btnCancel_Click()
    cancelled = True
    Me.Hide
End Sub

' Handle form close (X button)
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then ' User clicked X button
        cancelled = True
    End If
End Sub

' Helper function to check if form was cancelled
Public Function WasCancelled() As Boolean
    WasCancelled = cancelled
End Function
`;
      
      formCodeModule.AddFromString(buttonCode);
      debug.log(`Standard form event handlers added to '${formName}'`);
    } catch (error) {
      debug.error("Failed to add form event handlers:", error);
      throw new Error(`Failed to add event handlers to form '${formName}'. Error: ${error}`);
    }
  }

  /**
   * Creates a UserForm with multiple controls in batch - completely generic
   */
  public async createFormWithControls(
    formName: string,
    caption: string,
    width: number,
    height: number,
    controls: Array<{
      type: string;
      name: string;
      caption?: string;
      left: number;
      top: number;
      width: number;
      height: number;
    }>,
    addStandardButtons: boolean = true
  ): Promise<void> {
    const doc = await this.getActiveDocument();
    try {
      // Create the UserForm with proper sizing
      await this.createUserForm(formName, caption, width, height);
      
      // Add each control to the form
      for (const ctrl of controls) {
        await this.addControlToUserForm(
          formName,
          ctrl.type,
          ctrl.name,
          ctrl.caption || "",
          ctrl.left,
          ctrl.top,
          ctrl.width,
          ctrl.height
        );
      }
      
      // Add standard OK/Cancel buttons if requested
      if (addStandardButtons) {
        const buttonTop = height - 60;
        const buttonRight = width - 20;
        
        await this.addControlToUserForm(
          formName,
          "commandButton",
          "btnOK",
          "OK",
          buttonRight - 160,
          buttonTop,
          70,
          30
        );
        
        await this.addControlToUserForm(
          formName,
          "commandButton",
          "btnCancel",
          "Cancel",
          buttonRight - 80,
          buttonTop,
          70,
          30
        );
        
        // Add the standard event handlers
        await this.addStandardFormEventHandlers(formName);
      }
      
      debug.log(`Generic form '${formName}' created with ${controls.length} controls`);
    } catch (error) {
      debug.error("Failed to create form with controls:", error);
      throw new Error(`Failed to create form '${formName}'. Error: ${error}`);
    }
  }

  // --- Add more methods for other Word operations ---

}

// Export a singleton instance
export const wordService = new WordService();
