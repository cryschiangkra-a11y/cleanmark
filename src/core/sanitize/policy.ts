// 白名單：只有呢度列出嘅元素同屬性先會重新寫入新檔案，其他一律唔搬。
// 元素名用固定前綴（w:、m:、a: …），同原檔用咩前綴無關。

export const NAMESPACES: Record<string, string> = {
  w: "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
  r: "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
  m: "http://schemas.openxmlformats.org/officeDocument/2006/math",
  wp: "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing",
  wp14: "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing",
  a: "http://schemas.openxmlformats.org/drawingml/2006/main",
  pic: "http://schemas.openxmlformats.org/drawingml/2006/picture",
  wps: "http://schemas.microsoft.com/office/word/2010/wordprocessingShape",
  wpg: "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup",
  v: "urn:schemas-microsoft-com:vml",
  o: "urn:schemas-microsoft-com:office:office",
  w10: "urn:schemas-microsoft-com:office:word",
  mc: "http://schemas.openxmlformats.org/markup-compatibility/2006",
  xml: "http://www.w3.org/XML/1998/namespace",
};

export const PREFIX_OF: Record<string, string> = Object.fromEntries(
  Object.entries(NAMESPACES).map(([p, uri]) => [uri, p]),
);

const words = (s: string) => new Set(s.trim().split(/\s+/));

/** 保留嘅元素（連同內容） */
export const KEEP = new Set<string>([
  // 文件骨架同各種 part 嘅根元素
  ...[
    "document body hdr ftr footnotes footnote endnotes endnote separator continuationSeparator continuationNotice",
    // 段落
    "p pPr pStyle keepNext keepLines pageBreakBefore framePr widowControl numPr ilvl numId suppressLineNumbers",
    "pBdr top left bottom right start end between bar shd tabs tab suppressAutoHyphens kinsoku wordWrap",
    "overflowPunct topLinePunct autoSpaceDE autoSpaceDN bidi adjustRightInd snapToGrid spacing ind",
    "contextualSpacing mirrorIndents suppressOverlap jc textDirection textAlignment textboxTightWrap outlineLvl cnfStyle",
    // 文字
    "r rPr rStyle rFonts b bCs i iCs caps smallCaps strike dstrike outline shadow emboss imprint noProof",
    "vanish webHidden specVanish color w kern position sz szCs highlight u effect bdr fitText vertAlign rtl cs em",
    "lang eastAsianLayout oMath t noBreakHyphen softHyphen sym pgNum cr br ptab fldChar instrText",
    "footnoteReference endnoteReference footnoteRef endnoteRef dayShort monthShort yearShort dayLong monthLong yearLong",
    "ruby rubyPr rubyAlign hps hpsRaise hpsBaseText lid rt rubyBase",
    // 段落入面嘅容器
    "hyperlink bookmarkStart bookmarkEnd dir bdo drawing pict",
    // 表格
    "tbl tblPr tblStyle tblpPr tblOverlap bidiVisual tblStyleRowBandSize tblStyleColBandSize tblW tblCellSpacing tblInd",
    "tblBorders insideH insideV tblLayout tblCellMar tblLook tblGrid gridCol tr trPr gridBefore gridAfter wBefore wAfter",
    "cantSplit trHeight tblHeader hidden tblPrEx tc tcPr tcW gridSpan hMerge vMerge tcBorders tl2br tr2bl noWrap tcMar",
    "tcFitText vAlign hideMark",
    // 章節
    "sectPr headerReference footerReference footnotePr endnotePr pos numFmt numStart numRestart type pgSz pgMar",
    "pgBorders lnNumType pgNumType cols col vAlign noEndnote titlePg rtlGutter docGrid",
    // 樣式
    "styles docDefaults rPrDefault pPrDefault latentStyles lsdException style name basedOn next link autoRedefine",
    "uiPriority semiHidden unhideWhenUsed qFormat locked tblStylePr",
    // 編號
    "numbering abstractNum nsid multiLevelType styleLink numStyleLink lvl start lvlRestart isLgl suff lvlText legacy",
    "lvlJc num abstractNumId lvlOverride startOverride",
    // 設定（白名單，見 SETTINGS_KEEP）
    "settings",
    // 字型表
    "fonts font altName panose1 charset family notTrueType pitch sig",
    // 文字方塊
    "txbxContent",
  ].flatMap((line) => [...words(line)].map((n) => `w:${n}`)),
  // 方程式
  ...[
    "oMath oMathPara oMathParaPr r rPr t sty scr nor lit brk aln acc accPr chr bar barPr pos box boxPr opEmu noBreak diff",
    "argSz borderBox borderBoxPr hideTop hideBot hideLeft hideRight strikeH strikeV strikeBLTR strikeTLBR d dPr begChr",
    "sepChr endChr grow shp e eqArr eqArrPr baseJc maxDist objDist rSpRule rSp f fPr type num den func funcPr fName",
    "groupChr groupChrPr vertJc limLow limLowPr lim limUpp limUppPr m mPr mcs mc mcPr count mcJc plcHide cGpRule cSp cGp",
    "mr nary naryPr limLoc subHide supHide sub sup phant phantPr show zeroWid zeroAsc zeroDesc transp rad radPr degHide",
    "deg sPre sPrePr sSub sSubPr sSubSup sSubSupPr alnScr sSup sSupPr ctrlPr jc mathPr mathFont brkBin brkBinSub",
    "smallFrac dispDef lMargin rMargin defJc preSp postSp interSp intraSp wrapIndent wrapRight intLim naryLim",
  ].flatMap((line) => [...words(line)].map((n) => `m:${n}`)),
  // 圖片框架（DrawingML）
  ...[
    "inline anchor extent effectExtent docPr cNvGraphicFramePr simplePos positionH positionV posOffset align",
    "wrapNone wrapSquare wrapTight wrapThrough wrapTopAndBottom wrapPolygon start lineTo",
  ].flatMap((line) => [...words(line)].map((n) => `wp:${n}`)),
  ...["sizeRelH sizeRelV pctWidth pctHeight"].flatMap((line) => [...words(line)].map((n) => `wp14:${n}`)),
  ...[
    "graphic graphicData graphicFrameLocks blip stretch fillRect srcRect tile xfrm off ext chOff chExt prstGeom avLst gd",
    "custGeom gdLst ahLst cxnLst rect pathLst path moveTo lnTo arcTo quadBezTo cubicBezTo close pt",
    "noFill solidFill gradFill gsLst gs lin pattFill fgClr bgClr blipFill srgbClr schemeClr sysClr prstClr scrgbClr hslClr",
    "lumMod lumOff alpha tint shade satMod hueMod",
    "ln prstDash custDash ds round miter bevel headEnd tailEnd effectLst outerShdw innerShdw glow softEdge reflection",
    "picLocks spLocks grpSpLocks bodyPr normAutofit spAutoFit noAutofit prstTxWarp fontRef lnRef fillRef effectRef",
    // 主題
    "theme themeElements clrScheme dk1 lt1 dk2 lt2 accent1 accent2 accent3 accent4 accent5 accent6 hlink folHlink",
    "fontScheme majorFont minorFont latin ea cs font fmtScheme fillStyleLst lnStyleLst effectStyleLst effectStyle",
    "bgFillStyleLst objectDefaults extraClrSchemeLst",
  ].flatMap((line) => [...words(line)].map((n) => `a:${n}`)),
  ...["pic nvPicPr cNvPr cNvPicPr blipFill spPr"].flatMap((line) => [...words(line)].map((n) => `pic:${n}`)),
  ...["wsp cNvPr cNvSpPr cNvCnPr spPr style txbx linkedTxbx bodyPr"].flatMap((line) => [...words(line)].map((n) => `wps:${n}`)),
  ...["wgp cNvGrpSpPr grpSpPr grpSp"].flatMap((line) => [...words(line)].map((n) => `wpg:${n}`)),
  // 舊式 VML 圖片同文字方塊（WordArt 水印另外處理）
  ...["shape shapetype imagedata textbox fill stroke formulas f path"].flatMap((line) => [...words(line)].map((n) => `v:${n}`)),
  "o:lock",
  "w10:wrap",
]);

/** 只保留內容、拆走外殼嘅元素 */
export const UNWRAP = new Set<string>(
  ["ins", "moveTo", "smartTag", "customXml", "sdtContent"].map((n) => `w:${n}`),
);

/**
 * 設定檔只准呢啲元素（影響排版嘅），其餘例如 rsids、docVars、docId、attachedTemplate、
 * proofState、zoom、documentProtection 一律唔要
 */
export const SETTINGS_KEEP = new Set<string>(
  [
    "defaultTabStop autoHyphenation consecutiveHyphenLimit hyphenationZone doNotHyphenateCaps evenAndOddHeaders",
    "mirrorMargins gutterAtTop bordersDoNotSurroundHeader bordersDoNotSurroundFooter characterSpacingControl",
    "noPunctuationKerning printTwoOnOne strictFirstAndLastChars noLineBreaksAfter noLineBreaksBefore",
    "footnotePr endnotePr footnote endnote compat compatSetting themeFontLang clrSchemeMapping decimalSymbol",
    "listSeparator updateFields defaultTableStyle drawingGridHorizontalSpacing drawingGridVerticalSpacing",
    "displayHorizontalDrawingGridEvery displayVerticalDrawingGridEvery doNotUseMarginsForDrawingGridOrigin",
    "useFELayout doNotExpandShiftReturn doNotUseHTMLParagraphAutoSpacing doNotSnapToGridInCell",
    "doNotWrapTextWithPunct doNotUseEastAsianBreakRules growAutofit balanceSingleByteDoubleByteWidth",
    "adjustLineHeightInTable ulTrailSpace spaceForUL doNotLeaveBackslashAlone",
  ].flatMap((line) => [...words(line)].map((n) => `w:${n}`)),
);

/** 准搬嘅屬性（任何保留咗嘅元素都用同一份清單） */
export const ATTRS = new Set<string>([
  ...[
    "val w h type orient code top right bottom left header footer gutter start end fmt chapStyle chapSep space num",
    "equalWidth sep linePitch charSpace id name styleId default customStyle ascii hAnsi eastAsia cs hint asciiTheme",
    "hAnsiTheme eastAsiaTheme cstheme before after beforeLines afterLines beforeAutospacing afterAutospacing line",
    "lineRule firstLine firstLineChars hanging hangingChars leftChars rightChars startChars endChars pos leader color",
    "themeColor themeTint themeShade fill themeFill themeFillTint themeFillShade sz frame shadow fldCharType fldLock",
    "dirty anchor history tgtFrame docLocation abstractNumId numId ilvl tentative firstRow lastRow firstColumn",
    "lastColumn noHBand noVBand oddVBand evenVBand oddHBand evenHBand firstRowFirstColumn firstRowLastColumn",
    "lastRowFirstColumn lastRowLastColumn hRule wrap hAnchor vAnchor x xAlign y yAlign hSpace vSpace anchorLock",
    "dropCap lines leftFromText rightFromText topFromText bottomFromText tblpX tblpY tblpXSpec tblpYSpec horzAnchor",
    "vertAnchor char font uri count defLockedState defUIPriority defSemiHidden defUnhideWhenUsed defQFormat locked",
    "uiPriority semiHidden unhideWhenUsed qFormat accent1 accent2 accent3 accent4 accent5 accent6 bg1 bg2 t1 t2",
    "hyperlink followedHyperlink combine combineBrackets vert vertCompress countBy distance restart bidi",
    "usb0 usb1 usb2 usb3 csb0 csb1 displacedByCustomXml colFirst colLast instr",
  ].flatMap((line) => [...words(line)].map((n) => `w:${n}`)),
  "m:val",
  ...[
    "distT distB distL distR simplePos relativeHeight behindDoc locked layoutInCell allowOverlap hidden cx cy l t r b",
    "x y relativeFrom id name noChangeAspect edited",
  ].flatMap((line) => [...words(line)].map((n) => n)),
  "wp14:relativeFrom",
  "uri", // a:graphicData
  "cstate",
  "rot flipH flipV prst fmla w cap cmpd algn len lim dist dir sx sy kx ky blurRad rad stA endA stPos endPos fadeDir",
  "lastClr val bwMode ang scaled rotWithShape",
  "noGrp noSelect noRot noChangeArrowheads noMove noResize noEditPoints noAdjustHandles noChangeShapeType noCrop",
  "vertOverflow horzOverflow vert wrap lIns tIns rIns bIns numCol spcCol rtlCol fromWordArt anchor anchorCtr forceAA",
  "upright compatLnSpc idx typeface panose pitchFamily charset script txBox normalEastAsianFlow",
  // VML
  "style coordsize coordorigin o:spt path filled fillcolor stroked strokecolor strokeweight inset o:connecttype",
  "o:extrusionok gradientshapeok o:connectlocs eqn v:ext aspectratio type opacity color2 dashstyle",
  "anchorx anchory",
].flatMap((s) => s.split(/\s+/)));

/** r: 命名空間嘅屬性要重新編號 */
export const REL_ATTRS = new Set(["r:id", "r:embed"]);

/** 准保留嘅欄位類型（第一個字）。其他欄位只保留顯示結果，指令清空 */
export const FIELDS = new Set([
  "TOC", "TC", "PAGE", "NUMPAGES", "SECTIONPAGES", "SECTION", "SEQ", "REF", "PAGEREF", "NOTEREF", "STYLEREF",
  "HYPERLINK", "XE", "INDEX", "EQ", "SYMBOL", "LISTNUM", "DATE", "TIME", "CREATEDATE", "SAVEDATE", "PRINTDATE",
  "IF", "=", "ADVANCE", "QUOTE",
]);

/** Word 書籤名稱：字母或者底線開頭，最多 40 個字元 */
export const BOOKMARK_NAME = /^[A-Za-z_][A-Za-z0-9_]{0,39}$/;

/** mc:AlternateContent 入面，Word 2010 之後原生支援嘅 Requires */
export const SUPPORTED_REQUIRES = new Set(["wps", "wpg", "wp14", "w14", "w15", "m", "a14", "v"]);
