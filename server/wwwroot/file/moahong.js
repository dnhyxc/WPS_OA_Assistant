/**
 * moa Macro
 * 宏由 dnhyxc 录制，时间: 2022/05/09
 */
function moa() {
  const wpsApp = wps.WpsApplication();
  const Selection = wpsApp.ActiveWindow.Selection;

  Selection.Font.Name = "仿宋";
  (obj => {
    obj.Size = 16;
    obj.SizeBi = 16;
  })(Selection.Font);
  wpsApp.ActiveDocument.Range(0, 0).PageSetup.LeftMargin = 56.692459;
  wpsApp.ActiveDocument.Range(0, 0).PageSetup.LeftMargin = 71.999428;
  wpsApp.ActiveDocument.Range(0, 0).PageSetup.RightMargin = 56.692459;
  wpsApp.ActiveDocument.Range(0, 0).PageSetup.RightMargin = 71.999428;
  Selection.SetRange(0, 0);
  wpsApp.ActiveDocument.Range(0, 0).Start = 0;
  (obj => {
    obj.CharacterUnitFirstLineIndent = 2;
    obj.FirstLineIndent = 0;
    obj.CharacterUnitFirstLineIndent = 2;
    obj.FirstLineIndent = 0;
    obj.DisableLineHeightGrid = 0;
    obj.ReadingOrder = wdReadingOrderLtr;
    obj.AutoAdjustRightIndent = -1;
    obj.WidowControl = 0;
    obj.KeepWithNext = 0;
    obj.KeepTogether = 0;
    obj.PageBreakBefore = 0;
    obj.FarEastLineBreakControl = -1;
    obj.WordWrap = -1;
    obj.HangingPunctuation = -1;
    obj.HalfWidthPunctuationOnTopOfLine = 0;
    obj.AddSpaceBetweenFarEastAndAlpha = -1;
    obj.AddSpaceBetweenFarEastAndDigit = -1;
    obj.BaseLineAlignment = wdBaselineAlignAuto;
  })(Selection.ParagraphFormat);
  wpsApp.ActiveDocument.AcceptAllRevisions();
}
