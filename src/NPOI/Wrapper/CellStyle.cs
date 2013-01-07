using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;

namespace NPOI.Wrapper {
  public class CellStyle {
    private const int MAXIMUM_FONT_SIZE = 409;
    private const Fonts DEFAULT_FONT = Fonts.Arial;
    private int DEFAULT_FONTSIZE = 10;
    private int fontSize;

    /// <summary>
    /// Creates a CellStyle object configured as the default excel style
    /// </summary>
    public CellStyle() {
      Bold = false;
      Italic = false;
      WrapText = false;
      Underlined = false;
      FontSize = DEFAULT_FONTSIZE;
      FontName = DEFAULT_FONT;
      BorderType = BorderTypes.NONE;
      BorderTop = BorderLeft = BorderRight = BorderBottom = false;
      Alignment = Alignments.LEFT;
      BackgroundColor = Color.Empty;
      FontColor = Color.Black;
      CustomFormatting = null;
    }

    /// <summary>
    /// String representation of Font face
    /// </summary>
    public Fonts FontName { get; set; }

    /// <summary>
    /// Numeric representation of font size
    /// </summary>
    public int FontSize {
      get {
        return fontSize;
      }
      set {
        fontSize = value;
        if (fontSize > MAXIMUM_FONT_SIZE)
          fontSize = MAXIMUM_FONT_SIZE;
      }
    }

    /// <summary>
    /// Flag indicating whether the cell should be bold
    /// </summary>
    public bool Bold { get; set; }

    /// <summary>
    /// Flag indicating whether the cell should be italicized
    /// </summary>
    public bool Italic { get; set; }

    /// <summary>
    /// Cell Font color
    /// </summary>
    public Color FontColor { get; set; }

    /// <summary>
    /// Flag indicating whether text should wrap
    /// </summary>
    public bool WrapText { get; set; }

    /// <summary>
    /// Flag indicating whether the cell should be underlined
    /// </summary>
    public bool Underlined { get; set; }

    /// <summary>
    /// Border Style
    /// </summary>
    public BorderTypes BorderType { get; set; }

    /// <summary>
    /// Flag indicating whether there should be a top border
    /// </summary>
    public bool BorderTop { get; set; }

    /// <summary>
    /// Flag indicating whether there should be a left border
    /// </summary>
    public bool BorderLeft { get; set; }

    /// <summary>
    /// Flag indicating whether there should be a right border
    /// </summary>
    public bool BorderRight { get; set; }

    /// <summary>
    /// Flag indicating whether there should be a bottom border
    /// </summary>
    public bool BorderBottom { get; set; }

    /// <summary>
    /// Cell alignment
    /// </summary>
    public Alignments Alignment { get; set; }

    /// <summary>
    /// Cell background color
    /// </summary>
    public Color BackgroundColor { get; set; }

    public string CustomFormatting { get; set; }

    public override bool Equals(object obj) {
      if (!(obj is CellStyle))
        return false;

      CellStyle otherStyle = (CellStyle)obj;
      return Bold == otherStyle.Bold &&
          Alignment == otherStyle.Alignment &&
          BackgroundColor == otherStyle.BackgroundColor &&
          BorderBottom == otherStyle.BorderBottom &&
          BorderLeft == otherStyle.BorderLeft &&
          BorderRight == otherStyle.BorderRight &&
          BorderTop == otherStyle.BorderTop &&
          BorderType == otherStyle.BorderType &&
          FontName == otherStyle.FontName &&
          FontSize == otherStyle.FontSize &&
          FontColor == otherStyle.FontColor &&
          Italic == otherStyle.Italic &&
          Underlined == otherStyle.Underlined &&
          WrapText == otherStyle.WrapText &&
          CustomFormatting.Equals(otherStyle.CustomFormatting);
    }

    public override int GetHashCode() {
      return Bold.GetHashCode() * 31 + Alignment.GetHashCode() * 31 +
          BackgroundColor.GetHashCode() + BorderBottom.GetHashCode() * 31 +
          BorderLeft.GetHashCode() * 31 + BorderRight.GetHashCode() * 31 +
          BorderTop.GetHashCode() * 31 + BorderType.GetHashCode() * 31 +
          FontName.GetHashCode() * 31 + FontSize.GetHashCode() +
          FontColor.GetHashCode() + Italic.GetHashCode() * 31 +
          Underlined.GetHashCode() * 31 + WrapText.GetHashCode() * 31 +
          (CustomFormatting != null ? CustomFormatting.GetHashCode() : 1);
    }

    /// <summary>
    /// Allowed cell alignments
    /// </summary>
    public enum Alignments {
      LEFT,
      CENTER,
      RIGHT,
      JUSTIFY
    }

    /// <summary>
    /// Allowed border types
    /// </summary>
    public enum BorderTypes {
      HAIR,
      DASHED,
      DOTTED,
      DOUBLE,
      THIN,
      MEDIUM,
      THICK,
      NONE
    }

    /// <summary>
    /// Common fonts
    /// </summary>
    public enum Fonts {
      Arial,
      Times,
      Courier,
      Tahoma,
      Calibri,
      Batang,
      Broadway,
      Cambria,
      Castellar,
      Century,
      Fixedsys,
      Garamond,
      Georgia,
      Harrington,
      Terminal,
      Wingdings
    }

    /// <summary>
    /// Defines if this Excel style defines a new Font.
    /// Setting Bold to true changes the Font, Setting a BorderType will not change the font.
    /// </summary>
    /// <returns></returns>
    public bool FontStyleDefined() {
      return Bold || Underlined || Italic || FontName != DEFAULT_FONT || FontSize != DEFAULT_FONTSIZE;
    }
  }
}
