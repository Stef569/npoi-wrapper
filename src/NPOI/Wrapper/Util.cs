using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using NPOI.HSSF.UserModel;
using NPOI.HSSF.Util;

namespace NPOI.Wrapper {
  class Util {
    /// <summary>
    /// Lookup RGB from .NET system colour in Excel pallete - or nearest match.
    /// </summary>
    public static short GetXLColor(HSSFWorkbook xlWorkbook, Color color) {
      if (color == Color.Empty) {
        // COLOR_NORMAL=unspecified.
        return HSSFColor.COLOR_NORMAL;
      }

      HSSFPalette XlPalette = xlWorkbook.GetCustomPalette();
      HSSFColor XlColour = XlPalette.FindColor(color.R, color.G, color.B);

      if (XlColour == null) {
        XlColour = XlPalette.FindSimilarColor(color.R, color.G, color.B);
        return XlColour.GetIndex();
      } else {
        return XlColour.GetIndex();
      }
    }
  }
}
