using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace NPOI.Wrapper {
  /// <summary>
  /// EXCEL COORDINATES:
  /// All the col and row coordinates use a 1 based index like in excel. A1 = (1,1)
  /// When a method takes 2 coordinates, the column will always be the first parameter. 
  /// 
  /// This class reads and writes only from/to the local file system.
  /// 
  /// There are 2 ways to access cell data:
  /// 
  /// The first uses the nextRow() and SelectRow() methods to move to the desired row.
  /// You can then use the WriteXXX and ReadXXX methods that take a column and a value:
  /// WriteText(1, "50") Writes 50 to the first cell on the selected row.
  /// ReadText(1, "empty") Reads the cell value from the first cell on the selected row, 
  /// if the cell does not contain any text empty is returned.
  /// 
  /// The second way is to specify both column and row coordinates:
  /// WriteText(1, 2, "50") Writes 50 to A2
  /// ReadText(1, 1, "50") Reads the cell value from A1, if the cell does not contain any text 50 is returned.
  /// </summary>
  class ExcelWorkbook {
    private NPOIWrapper wrapper = new NPOIWrapper();
    public string XlsFilePath { get; private set; }

    /// <summary>
    /// Create a new Excel workbook at the xlsFilePath location. If the location already contains an excel workbook with the same name
    /// then that file is overwritten without warning.
    /// 
    /// The workbook will contain 1 sheet named 'Sheet1'. The first row is selected.
    /// </summary>
    /// <param name="xlsFilePath">The path where the new excel file should be created.</param>
    /// <remarks></remarks>
    public void Create(String xlsFilePath) {
      this.XlsFilePath = xlsFilePath;
      wrapper.Create();
      wrapper.Save(new FileStream(xlsFilePath, FileMode.Create));
      SelectRow(1);
    }

    /// <summary>
    /// Open the Excel XLS file for reading. The first row is selected.
    /// </summary>
    /// <param name="xlsFilePath">Path to the excel file to open.</param>
    /// <remarks></remarks>
    public void Open(String xlsFilePath) {
      this.XlsFilePath = xlsFilePath;
      wrapper.Open(new FileStream(xlsFilePath, FileMode.Open, FileAccess.ReadWrite));
      SelectRow(1);
    }

    /// <summary>
    /// Save this workbook to the XlsFilePath location given in the create or open method.
    /// </summary>
    public void Save() {
      SaveAs(XlsFilePath);
    }

    /// <summary>
    /// Save this workbook to the given local path.
    /// </summary>
    public void SaveAs(String xlsFilePath) {
      wrapper.Save(new FileStream(xlsFilePath, FileMode.Create, FileAccess.Write));
    }

    /// <summary>
    /// Select the next row.
    /// </summary>
    public void NextRow() {
      wrapper.NextRow();
    }

    /// <summary>
    /// Select the given row. This allowes to write data into the cells on this row.
    /// </summary>
    public void SelectRow(int row) {
      wrapper.SelectRow(row-1);
    }

    /// <summary>
    /// Check if the next row is within the sheets bounds.
    /// </summary>
    public bool HasMoreRows() {
      return wrapper.HasMoreRows();
    }

    /// <summary>
    /// Check if the next row is smaller or equal then maxRows and within the sheet bounds.
    /// </summary>
    public bool HasMoreRows(int maxRows) {
      return wrapper.HasMoreRows() && wrapper.CurrentRow <= maxRows;
    }

    public void GroupRows(int firstRow, int lastRow) {
      wrapper.GroupRows(firstRow-1, lastRow-1);
    }

    public void GroupCols(int firstCol, int lastCol) {
      wrapper.GroupCols(firstCol-1, lastCol-1);
    }

    /// <summary>
    /// Write the text value into the cell at the given column and current row.
    /// The previous value is overwritten.
    /// </summary>
    /// <param name="col">The column of the cell to write the value in. 1 based index.</param>
    /// <param name="value">The text value to write into this cell</param>
    public void WriteText(int col, string value) {
      wrapper.WriteCell(col - 1, value);
    }

    /// <summary>
    /// Write the number value into the cell at the given column and current row.
    /// The previous value is overwritten.
    /// </summary>
    /// <param name="col">The column of the cell to write the value in. 1 based index.</param>
    /// <param name="value">The number value to write into this cell</param>
    public void WriteNumber(int col, double value) {
      wrapper.WriteNumber(col - 1, value);
    }

    /// <summary>
    /// Write the percentage value into the cell at the given column and current row.
    /// The previous value is overwritten. The cell if formatted as Percentage
    /// </summary>
    /// <param name="col">The column of the cell to write the value in. 1 based index.</param>
    /// <param name="value">The percentage value to write into this cell</param>
    public void WritePercentage(int col, double value) {
      wrapper.WriteNumber(col - 1, (double)value / 100);
      CellStyle cellStyle = new CellStyle();
      cellStyle.CustomFormatting = "0%";
      wrapper.ApplyStyle(col - 1, cellStyle);
    }

    /// <summary>
    /// see ReadNumber(col)
    /// </summary>
    public double ReadNumber(int col, int row) {
      int currentRow = CurrentRow;
      SelectRow(row);
      double number = ReadNumber(col);
      SelectRow(currentRow);
      return number;
    }

    /// <summary>
    /// Read the value from the given column and current row as double type. Strings are attempted to be converted to double.
    /// This function can also be used for integer numbers.
    /// </summary>
    /// <param name="col">The column of the cell to read the value from. 1 based index.</param>
    /// <returns>The number value or -1 if the number could not be read from the cell.</returns>
    public double ReadNumber(int col) {
      return ReadNumber(col, (double)-1);
    }

    /// <summary>
    /// Read the value from the given column and current row as double type. Strings are attempted to be converted to double.
    /// This function can also be used for integer numbers.
    /// </summary>
    /// <param name="col">The column of the cell to read the value from. 1 based index.</param>
    /// <param name="defaultValue">The value to return if the cell does not contain a number.</param>
    /// <returns>The number value or defaultValue if the number could not be read from the cell.</returns>
    public double ReadNumber(int col, double defaultValue) {
      object cellValue = wrapper.ReadCell(col - 1);

      if (cellValue == null) {
        return defaultValue;
      }

      bool cellHasNumber;
      double parsedNumber;
      cellHasNumber = Double.TryParse(cellValue+"", out parsedNumber);

      if (cellHasNumber) {
        return parsedNumber;
      } else {
        return defaultValue;
      }
    }

    /// <summary>
    /// see ReadDate(col)
    /// </summary>
    public DateTime ReadDate(int col, int row) {
      int currentRow = CurrentRow;
      SelectRow(row);
      DateTime date = ReadDate(col);
      SelectRow(currentRow);
      return date;
    }

    /// <summary>
    /// Read the date value from the given column and current row as DateTime type. Strings are attempted to be converted to DateTime.
    /// </summary>
    /// <param name="col">The column of the cell to read the value from. 1 based index.</param>
    /// <returns>The date value or DateTime.MinValue if the date could not be read from the cell</returns>
    public DateTime ReadDate(int col) {
      return ReadDate(col, DateTime.MinValue);
    }

    /// <summary>
    /// Read the date value from the given column and current row as DateTime type. Strings are attempted to be converted to DateTime.
    /// </summary>
    /// <param name="col">The column of the cell to read the value from. 1 based index.</param>
    /// <param name="defaultValue">The value to return if the cell does not contain a date.</param>
    /// <returns>The date or the defaultValue if the date could not be read from the cell.</returns>
    public DateTime ReadDate(int col, DateTime defaultValue) {
      object cellValue = wrapper.ReadCellAsDate(col - 1);

      if (cellValue == null) {
        return defaultValue;
      } else if (cellValue is DateTime) {
        return (DateTime)cellValue;
      } else {
        return defaultValue;
      }
    }

    public string ReadText(int col, int row) {
      int currentRow = CurrentRow;
      SelectRow(row);
      string text = ReadText(col);
      SelectRow(currentRow);
      return text;
    }

    public string ReadText(int col) {
      return wrapper.ReadCell(col - 1) + "";
    }

    public string ReadText(int col, string defaultValue) {
      object cellValue = wrapper.ReadCell(col - 1);

      return cellValue != null ? cellValue + "" : defaultValue;
    }

    public int CurrentRow {
      get { return wrapper.CurrentRow+1; }
    }

    public void ApplyFormatting(int col, CellStyle cellStyle) {
      wrapper.ApplyStyle(col - 1, cellStyle);
    }
  }
}
