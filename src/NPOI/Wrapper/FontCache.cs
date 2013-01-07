using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.SS.UserModel;
using System.Drawing;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using System.Diagnostics;

namespace NPOI.Wrapper {
  class FontCache {
    private HSSFWorkbook xlWorkbook;

    public FontCache(HSSFWorkbook xlWorkbook) {
      this.xlWorkbook = xlWorkbook;
    }

    public IFont GetFont(CellStyle cellStyle) {
      return GetFont(
              cellStyle.Bold,
              cellStyle.Italic,
              cellStyle.Underlined,
              cellStyle.FontColor,
              cellStyle.FontSize,
              cellStyle.FontName.ToString()
      );
    }

    public IFont GetFont(bool bold, bool italic, bool underline, Color fontColor, int fontSize, string fontName) {
      short npoiFontColor = Util.GetXLColor(xlWorkbook, fontColor);
      IFont xlFont = xlWorkbook.FindFont(CBold(bold), npoiFontColor, (short)(fontSize * 20), fontName, italic, false, 0, CUnderline(underline));

      if (xlFont == null) {
        xlFont = xlWorkbook.CreateFont();
        xlFont.Boldweight = CBold(bold);
        xlFont.Color = npoiFontColor;
        xlFont.Underline = CUnderline(underline);
        xlFont.IsItalic = italic;
        xlFont.FontName = fontName;
        xlFont.FontHeightInPoints = (short)fontSize;
        Debug.Print("Creating Font nr " + xlWorkbook.NumberOfFonts + " " + xlFont.Boldweight + ", " + xlFont.Color + ", " + xlFont.FontHeight + ", " + xlFont.FontName + ", " + xlFont.IsItalic + ", " + xlFont.IsStrikeout + ", " + xlFont.TypeOffset + ", " + xlFont.Underline);
      }
      return xlFont;
    }

    private short CBold(bool bold) {
      return bold ? (short)NPOI.SS.UserModel.FontBoldWeight.BOLD : (short)NPOI.SS.UserModel.FontBoldWeight.NORMAL;
    }

    private byte CUnderline(bool underline) {
      return underline ? FontUnderline.SINGLE.ByteValue : FontUnderline.NONE.ByteValue;
    }

  }
}
