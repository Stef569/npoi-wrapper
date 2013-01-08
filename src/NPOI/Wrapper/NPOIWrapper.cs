using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.SS.Util;
using System.Drawing;

namespace NPOI.Wrapper {
  /// <summary>
  /// Provides easy to use API for basic Excel 97+ BIFF file reading and writing.
  /// All NPOI objects are hidden.
  /// Cell styles are cached, fonts and Dataformats are reused
  /// Based on NPOIExcel Written by Hunter Beanland
  /// </summary>
  class NPOIWrapper {
    private const int MAX_ROWS = 65536;
    private IWorkbook xlWorkbook;         // The npoi workbook
    private ISheet xlSheet;               // Current sheet
    private IRow xlRow;                   // Current row

    // User defined styles. Excel has a limited number of styles.
    private CellStyleCache styleCache;
    private FontCache fontCache;

    #region IO
    /// <summary>
    /// Create a new Excel workbook with 1 sheet named 'Sheet1' in memory.
    /// Use save to persist your changes.
    /// </summary>
    public void Create() {
      xlWorkbook = new HSSFWorkbook();
      CreateSheet("Sheet1");
      styleCache = new CellStyleCache(xlWorkbook);
      fontCache = new FontCache((HSSFWorkbook)xlWorkbook);
    }

    /// <summary>
    /// Read an excel workbook from a stream. The first sheet will be selected. The stream will be closed.
    /// </summary>
    /// <param name="xlsStream">The stream to read an excel workbook from</param>
    public void Open(Stream xlsStream) {
      try {
        xlWorkbook = new HSSFWorkbook(xlsStream);
        xlSheet = xlWorkbook.GetSheetAt(0);
        styleCache = new CellStyleCache(xlWorkbook);
        fontCache = new FontCache((HSSFWorkbook)xlWorkbook);
	  } finally {
        if (xlsStream != null) xlsStream.Close();
      }
    }

    /// <summary>
    /// Save the workbook to a stream. The stream will be closed.
    /// </summary>
    /// <param name="xlsStream">Stream to save the excel workbook to</param>
    public void Save(Stream xlsStream) {
      try {
        xlWorkbook.Write(xlsStream);
      } finally {
        xlsStream.Close();
      }
    }
    #endregion

    #region Sheets
    public void CreateSheet(string SheetName) {
      xlSheet = xlWorkbook.CreateSheet(SheetName);
    }

    /// <summary>
    /// Checks to see if a sheet exists
    /// </summary>
    /// <param name="sheetName">Name of the worksheet to check for existence</param>
    /// <returns>True if the sheet exists, false otherwise</returns>
    public bool SheetExists(string sheetName) {
      try {
        ISheet s = xlWorkbook.GetSheet(sheetName);
        return s != null;
      } catch {
        return false;
      }
    }

    public void AutoSizeColumn(int col) {
      xlSheet.AutoSizeColumn(col);
    }

    /// <summary>
    /// Group a range of rows and hide them, 0 based
    /// </summary>
    public void GroupRows(int firstRow, int lastRow) {
      xlSheet.GroupRow(firstRow, lastRow);
      xlSheet.SetRowGroupCollapsed(firstRow, true);
    }

    /// <summary>
    /// Group a range of columns and hide them, 0 based
    /// </summary>
    public void GroupCols(int firstCol, int lastCol) {
      xlSheet.GroupColumn(firstCol, lastCol);
      xlSheet.SetColumnGroupCollapsed(firstCol, true);
    }
    #endregion

    #region Navigation
    /// <summary>
    /// Select the next row.
    /// </summary>
    public void NextRow() {
      SelectRow(xlRow.RowNum + 1);
    }

    /// <summary>
    /// Select the given row. This allowes to write data into the cells on this row.
    /// </summary>
    public void SelectRow(int row) {
      xlRow = xlSheet.GetRow(row);

      if (xlRow == null) {
        xlRow = xlSheet.CreateRow(row);
      }
    }

    /// <summary>
    /// Check if the next row is within the sheets bounds.
    /// </summary>
    public bool HasMoreRows() {
      int nextRow = xlRow.RowNum + 1;
      return nextRow < MAX_ROWS;
    }
    #endregion

    #region Cells
    /// <summary>
    /// Write the text value in the given column in the current selected row.
    /// The previous value is overwritten.
    /// </summary>
    /// <param name="col">The column of the cell to write the value in. 0 based index.</param>
    /// <param name="value">The text value to write into this cell</param>
    public void WriteCell(int col, string value) {
      ICell xlCell = xlRow.CreateCell(col);
      xlCell.SetCellValue(value);
    }

    /// <summary>
    /// Write the number value in the given column in the current selected row.
    /// The previous value is overwritten.
    /// </summary>
    /// <param name="col">The column of the cell to write the value in. 0 based index.</param>
    /// <param name="value">The text value to write into this cell</param>
    public void WriteNumber(int col, double value) {
      ICell xlCell = xlRow.CreateCell(col);
      xlCell.SetCellValue(value);
    }

    /// <summary>
    /// Read the cell value from the given column in the current selected row.
    /// Possible return types: string, double, bool, null. 
    /// </summary>
    /// <param name="col">The column of the cell to read the value from. 0 based index.</param>
    /// <returns>The cell value or null if the value could not be read</returns>
    public object ReadCell(int col) {
      ICell cell = xlRow.GetCell(col);

      if (cell == null) {
        return null;
      } else {
        switch (xlRow.GetCell(col).CellType) {
          case CellType.FORMULA:
          case CellType.STRING:
            return xlRow.GetCell(col).StringCellValue;
          case CellType.NUMERIC:
            return xlRow.GetCell(col).NumericCellValue;
          case CellType.BOOLEAN:
            return xlRow.GetCell(col).BooleanCellValue;
          default:
            return null;
        }
      }
    }

    /// <summary>
    /// Read the date cell value from the given column in the current selected row.
    /// </summary>
    /// <param name="col">The column of the cell to read the value from. 0 based index.</param>
    /// <returns>The cell value or DateTime.MinValue if the value could not be read</returns>
    public DateTime ReadCellAsDate(int col) {
      object cellValue = ReadCell(col);

      if (cellValue is double) {
        return xlRow.GetCell(col).DateCellValue;
      } else if (cellValue is string) {
        DateTime xlDate = default(DateTime);
        DateTime.TryParse(cellValue + "", out xlDate);
        return xlDate;
      } else {
        return DateTime.MinValue;
      }
    }
    #endregion

    #region Properties
    public int CurrentRow {
      get { return xlRow.RowNum; }
    }
    #endregion

    #region Cell Style
    public void ApplyStyle(int col, CellStyle cellStyle) {
      ICell xlCell = xlRow.GetCell(col, MissingCellPolicy.CREATE_NULL_AS_BLANK);
      ICellStyle xlStyle = styleCache.GetGeneralStyle();

      string customFormatting = cellStyle.CustomFormatting;
      if (customFormatting != null) {
        // check if this is a built-in format
        short builtinFormatId = HSSFDataFormat.GetBuiltinFormat(customFormatting);

        if (builtinFormatId != -1) {
          xlStyle.DataFormat = builtinFormatId;
        } else {
          // not a built-in format, so create a new one
          IDataFormat newDataFormat = xlWorkbook.CreateDataFormat();
          xlStyle.DataFormat = newDataFormat.GetFormat(cellStyle.CustomFormatting);
        }
      }

      IFont font = fontCache.GetFont(cellStyle);

      if (cellStyle.BackgroundColor != Color.Empty) {
        xlStyle.FillForegroundColor = Util.GetXLColor((HSSFWorkbook)xlWorkbook, cellStyle.BackgroundColor);
        xlStyle.FillPattern = FillPatternType.SOLID_FOREGROUND;
      }

      switch (cellStyle.Alignment) {
        case CellStyle.Alignments.LEFT: xlStyle.Alignment = HorizontalAlignment.LEFT; break;
        case CellStyle.Alignments.CENTER: xlStyle.Alignment = HorizontalAlignment.CENTER; break;
        case CellStyle.Alignments.RIGHT: xlStyle.Alignment = HorizontalAlignment.RIGHT; break;
        case CellStyle.Alignments.JUSTIFY: xlStyle.Alignment = HorizontalAlignment.JUSTIFY; break;
      }

      xlStyle.SetFont(font);
      styleCache.SetCellStyle(xlCell, xlStyle);
    }
    #endregion
  }
}
