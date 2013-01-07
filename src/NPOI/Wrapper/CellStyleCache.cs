using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using NPOI.Util.Collections;
using NPOI.SS.UserModel;
using System.Diagnostics;

namespace NPOI.Wrapper {
  public class CellStyleCache {
    private Dictionary<string, CellStyleWrapper> cellStyles;
    private ICellStyle cellStyle, defaultCellStyle;
    private IWorkbook xlWorkbook;

    public CellStyleCache(IWorkbook xlWorkbook) {
      this.xlWorkbook = xlWorkbook;
      this.cellStyles = new Dictionary<string, CellStyleWrapper>();
      this.cellStyle = xlWorkbook.CreateCellStyle();
      this.defaultCellStyle = xlWorkbook.CreateCellStyle();
    }

    public void setCellStyle(ICell cell, ICellStyle npoiCellStyle) {
      CellStyleWrapper cellStyle = new CellStyleWrapper(npoiCellStyle);
      string cellStyleKey = cellStyle.getKey();

      if (cellStyles.ContainsKey(cellStyleKey)) {
        // reuse cached styles
        CellStyleWrapper cachedCellStyle = cellStyles[cellStyleKey];
        cell.CellStyle = cachedCellStyle.CellStyle;
      } else {
        // If the style does not exist create a new one
        ICellStyle newCellStyle = xlWorkbook.CreateCellStyle();
        copyCellStyle(xlWorkbook, npoiCellStyle, newCellStyle);
        add(new CellStyleWrapper(newCellStyle));
        cell.CellStyle = newCellStyle;
      }
    }

    private void add(CellStyleWrapper wrapper) {
      string key = wrapper.getKey();
      Debug.Print("Creating style " + xlWorkbook.NumCellStyles + " " + key);
      cellStyles.Add(key, wrapper);
    }

    public ICellStyle getGeneralStyle() {
      copyCellStyle(xlWorkbook, defaultCellStyle, cellStyle);
      return cellStyle;
    }

    /// <summary>
    /// Copy c1 into c2
    /// </summary>
    public static void copyCellStyle(IWorkbook wb, ICellStyle c1, ICellStyle c2) {
      c2.Alignment = c1.Alignment;
      c2.BorderBottom = c1.BorderBottom;
      c2.BorderLeft = c1.BorderLeft;
      c2.BorderRight = c1.BorderRight;
      c2.BorderTop = c1.BorderTop;
      c2.BottomBorderColor = c1.BottomBorderColor;
      c2.DataFormat = c1.DataFormat;
      c2.FillBackgroundColor = c1.FillBackgroundColor;
      c2.FillForegroundColor = c1.FillForegroundColor;
      c2.FillPattern = c1.FillPattern;
      c2.SetFont(wb.GetFontAt(c1.FontIndex));
      // HIDDEN ?
      c2.Indention = c1.Indention;
      c2.LeftBorderColor = c1.LeftBorderColor;
      // LOCKED ?
      c2.RightBorderColor = c1.RightBorderColor;
      c2.Rotation = c1.Rotation;
      c2.TopBorderColor = c1.TopBorderColor;
      c2.VerticalAlignment = c1.VerticalAlignment;
      c2.WrapText = c1.WrapText;
    }
  }

  class CellStyleWrapper {
    public ICellStyle CellStyle { get; private set; }
    private string hash = null;

    public CellStyleWrapper(ICellStyle cellStyle) {
      this.CellStyle = cellStyle;
    }

    public bool equals(Object obj) {
      if (obj == null) {
        return false;
      }

      CellStyleWrapper otherCellStyle = obj as CellStyleWrapper;
      if ((System.Object)otherCellStyle == null) {
        return false;
      }
      ICellStyle a = this.CellStyle;
      ICellStyle b = otherCellStyle.CellStyle;

      return a.Alignment == b.Alignment &&
            a.BorderBottom == b.BorderBottom &&
            a.BorderLeft == b.BorderLeft &&
            a.BorderRight == b.BorderRight &&
            a.BorderTop == b.BorderTop &&
            a.BottomBorderColor == b.BottomBorderColor &&
            a.DataFormat == b.DataFormat &&
            a.FillBackgroundColor == b.FillBackgroundColor &&
            a.FillForegroundColor == b.FillForegroundColor &&
            a.FillPattern == b.FillPattern &&
            a.FontIndex == b.FontIndex &&
            a.Indention == b.Indention &&
            a.LeftBorderColor == b.LeftBorderColor &&
            a.RightBorderColor == b.RightBorderColor &&
            a.Rotation == b.Rotation &&
            a.TopBorderColor == b.TopBorderColor &&
            a.VerticalAlignment == b.VerticalAlignment &&
            a.WrapText == b.WrapText;
    }

    public override int GetHashCode() {
      int hash = 17;
      hash += 37 * CellStyle.Alignment.GetHashCode();
      hash += 37 * CellStyle.BorderBottom.GetHashCode();
      hash += 37 * CellStyle.BorderLeft.GetHashCode();
      hash += 37 * CellStyle.BorderRight.GetHashCode();
      hash += 37 * CellStyle.BorderTop.GetHashCode();
      hash += 37 * CellStyle.BottomBorderColor.GetHashCode();
      hash += 37 * CellStyle.DataFormat.GetHashCode();
      hash += 37 * CellStyle.FillBackgroundColor.GetHashCode();
      hash += 37 * CellStyle.FillForegroundColor.GetHashCode();
      hash += 37 * CellStyle.FillPattern.GetHashCode();
      hash += 37 * CellStyle.FontIndex.GetHashCode();
      hash += 37 * CellStyle.Indention.GetHashCode();
      hash += 37 * CellStyle.LeftBorderColor.GetHashCode();
      hash += 37 * CellStyle.RightBorderColor.GetHashCode();
      hash += 37 * CellStyle.Rotation.GetHashCode();
      hash += 37 * CellStyle.TopBorderColor.GetHashCode();
      hash += 37 * CellStyle.VerticalAlignment.GetHashCode();
      hash += 37 * 23 + CellStyle.WrapText.GetHashCode();
      return hash;
    }

    public string getKey() {
      if (hash == null) {
        hash += CellStyle.Alignment.ToString();
        hash += CellStyle.BorderBottom.ToString();
        hash += CellStyle.BorderLeft.ToString();
        hash += CellStyle.BorderRight.ToString();
        hash += CellStyle.BorderTop.ToString();
        hash += CellStyle.BottomBorderColor.ToString();
        hash += CellStyle.DataFormat.ToString();
        hash += CellStyle.FillBackgroundColor.ToString();
        hash += CellStyle.FillForegroundColor.ToString();
        hash += CellStyle.FillPattern.ToString();
        hash += CellStyle.FontIndex.ToString();
        hash += CellStyle.Indention.ToString();
        hash += CellStyle.LeftBorderColor.ToString();
        hash += CellStyle.RightBorderColor.ToString();
        hash += CellStyle.Rotation.ToString();
        hash += CellStyle.TopBorderColor.ToString();
        hash += CellStyle.VerticalAlignment.ToString();
        hash += CellStyle.WrapText.ToString();
      }
      return hash;
    }
  }
}