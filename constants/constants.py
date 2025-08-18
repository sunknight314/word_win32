# constants.py
# 一、NumberStyle
class NumberStyle:
    # 1、基础数字与字母样式
    wdListNumberStyleArabic = 0
    wdListNumberStyleUppercaseRoman = 1
    wdListNumberStyleLowercaseRoman = 2
    wdListNumberStyleUppercaseLetter = 3
    wdListNumberStyleLowercaseLetter = 4
    wdListNumberStyleOrdinal = 5
    wdListNumberStyleCardinalText = 6
    wdListNumberStyleOrdinalText = 7

    # 2、中文相关样式
    wdListNumberStyleSimpChinNum1 = 37
    wdListNumberStyleSimpChinNum2 = 38
    wdListNumberStyleSimpChinNum3 = 39
    wdListNumberStyleSimpChinNum4 = 40
    wdListNumberStyleTradChinNum1 = 33
    wdListNumberStyleTradChinNum2 = 34
    wdListNumberStyleTradChinNum3 = 35
    wdListNumberStyleTradChinNum4 = 36

    # 3、日韩相关样式
    wdListNumberStyleKanji = 10
    wdListNumberStyleKanjiDigit = 11
    wdListNumberStyleAiueo = 20
    wdListNumberStyleIroha = 21
    wdListNumberStyleHanjaRead = 41
    wdListNumberStyleHangul = 43

    # 4、特殊符号样式
    wdListNumberStyleBullet = 23
    wdListNumberStyleNumberInCircle = 18
    wdListNumberStylePictureBullet = 249
    wdListNumberStyleLegal = 253
    wdListNumberStyleLegalLZ = 254

    # 5、多语言支持样式
    wdListNumberStyleArabicFullWidth = 14
    wdListNumberStyleArabicLZ = 22
    wdListNumberStyleHebrew1 = 45
    wdListNumberStyleHebrew2 = 47
    wdListNumberStyleThaiLetter = 53
    wdListNumberStyleHindiLetter1 = 49

    # 6、其他专业样式
    wdListNumberStyleGBNum1 = 26
    wdListNumberStyleGBNum2 = 27
    wdListNumberStyleZodiac1 = 30
    wdListNumberStyleLowercaseGreek = 60
    wdListNumberStyleNone = 255


# 二、TrailingCharacter
class TrailingCharacter:
    wdTrailingTab = 0
    wdTrailingSpace = 1
    wdTrailingNone = 2


# 三、WdListLevelAlignment
class WdListLevelAlignment:
    wdListLevelAlignLeft = 0
    wdListLevelAlignCenter = 1
    wdListLevelAlignRight = 2


# 四、WdListApplyTo
class WdListApplyTo:
    wdListApplyToWholeList = 0
    wdListApplyToThisPointForward = 1
    wdListApplyToSelection = 2


# 五、WdDefaultListBehavior
class WdDefaultListBehavior:
    wdWord8ListBehavior = 0
    wdWord9ListBehavior = 1
    wdWord10ListBehavior = 2


# 六、WdStyleType
class WdStyleType:
    wdStyleTypeParagraph = 1
    wdStyleTypeCharacter = 2
    wdStyleTypeTable = 3
    wdStyleTypeList = 4
    wdStyleTypeParagraphOnly = 5
    wdStyleTypeLinked = 6


# 七、WdParagraphAlignment
class WdParagraphAlignment:
    wdAlignParagraphLeft = 0
    wdAlignParagraphCenter = 1
    wdAlignParagraphRight = 2
    wdAlignParagraphJustify = 3
    wdAlignParagraphDistribute = 4
    wdAlignParagraphJustifyMed = 5
    wdAlignParagraphJustifyHi = 7
    wdAlignParagraphJustifyLow = 8
    wdAlignParagraphThaiJustify = 9


# 八、WdBreakType
class WdBreakType:
    wdSectionBreakNextPage = 2
    wdSectionBreakContinuous = 3
    wdSectionBreakEvenPage = 4
    wdSectionBreakOddPage = 5
    wdLineBreak = 6
    wdPageBreak = 7
    wdColumnBreak = 8
    wdLineBreakClearLeft = 9
    wdLineBreakClearRight = 10
    wdTextWrappingBreak = 11


# 九、WdHeaderFooterIndex
class WdHeaderFooterIndex:
    wdHeaderFooterPrimary = 1
    wdHeaderFooterFirstPage = 2
    wdHeaderFooterEvenPages = 3

# 十、WdCollapseDirection
class WdCollapseDirection:
    wdCollapseEnd = 0
    wdCollapseStart = 1

# 十一、WdFieldType
class WdFieldType:
    wdFieldEmpty = -1
    wdFieldRef = 3
    wdFieldIndexEntry = 4
    wdFieldFootnoteRef = 5
    wdFieldSet = 6
    wdFieldIf = 7
    wdFieldIndex = 8
    wdFieldTOCEntry = 9
    wdFieldStyleRef = 10
    wdFieldRefDoc = 11
    wdFieldSequence = 12
    wdFieldTOC = 13
    wdFieldInfo = 14
    wdFieldTitle = 15
    wdFieldSubject = 16
    wdFieldAuthor = 17
    wdFieldKeyWord = 18
    wdFieldComments = 19
    wdFieldLastSavedBy = 20
    wdFieldCreateDate = 21
    wdFieldSaveDate = 22
    wdFieldPrintDate = 23
    wdFieldRevisionNum = 24
    wdFieldEditTime = 25
    wdFieldNumPages = 26
    wdFieldNumWords = 27
    wdFieldNumChars = 28
    wdFieldFileName = 29
    wdFieldTemplate = 30
    wdFieldDate = 31
    wdFieldTime = 32
    wdFieldPage = 33
    wdFieldExpression = 34
    wdFieldQuote = 35
    wdFieldInclude = 36
    wdFieldPageRef = 37
    wdFieldAsk = 38
    wdFieldFillIn = 39
    wdFieldData = 40
    wdFieldNext = 41
    wdFieldNextIf = 42
    wdFieldSkipIf = 43
    wdFieldMergeRec = 44
    wdFieldDDE = 45
    wdFieldDDEAuto = 46
    wdFieldGlossary = 47
    wdFieldPrint = 48
    wdFieldFormula = 49
    wdFieldGoToButton = 50
    wdFieldMacroButton = 51
    wdFieldAutoNumOutline = 52
    wdFieldAutoNumLegal = 53
    wdFieldAutoNum = 54
    wdFieldImport = 55
    wdFieldLink = 56
    wdFieldSymbol = 57
    wdFieldEmbed = 58
    wdFieldMergeField = 59
    wdFieldUserName = 60
    wdFieldUserInitials = 61
    wdFieldUserAddress = 62
    wdFieldBarCode = 63
    wdFieldDocVariable = 64
    wdFieldSection = 65
    wdFieldSectionPages = 66
    wdFieldIncludePicture = 67
    wdFieldIncludeText = 68
    wdFieldFileSize = 69
    wdFieldFormTextInput = 70
    wdFieldFormCheckBox = 71
    wdFieldNoteRef = 72
    wdFieldTOA = 73
    wdFieldTOAEntry = 74
    wdFieldMergeSeq = 75
    wdFieldPrivate = 77
    wdFieldDatabase = 78
    wdFieldAutoText = 79
    wdFieldCompare = 80
    wdFieldAddin = 81
    wdFieldSubscriber = 82
    wdFieldFormDropDown = 83
    wdFieldAdvance = 84
    wdFieldDocProperty = 85
    wdFieldOCX = 87
    wdFieldHyperlink = 88
    wdFieldAutoTextList = 89
    wdFieldListNum = 90
    wdFieldHTMLActiveX = 91
    wdFieldBidiOutline = 92
    wdFieldAddressBlock = 93
    wdFieldGreetingLine = 94
    wdFieldShape = 95
    wdFieldCitation = 96
    wdFieldBibliography = 97
    wdFieldMergeBarcode = 98
    wdFieldDisplayBarcode = 99

# 十二、WdBorderType
class WdBorderType:
    wdBorderDiagonalUp = -8
    wdBorderDiagonalDown = -7
    wdBorderVertical = -6
    wdBorderHorizontal = -5
    wdBorderRight = -4
    wdBorderBottom = -3
    wdBorderLeft = -2
    wdBorderTop = -1

# 十三、WdLineStyle
class WdLineStyle:
    wdLineStyleNone = 0                  # 无边框
    wdLineStyleSingle = 1                # 单实线
    wdLineStyleDot = 2                   # 点
    wdLineStyleDashSmallGap = 3          # 划线后跟小间隙
    wdLineStyleDashLargeGap = 4          # 划线后跟大间隙
    wdLineStyleDashDot = 5               # 划线后跟点
    wdLineStyleDashDotDot = 6            # 划线后跟两个点
    wdLineStyleDouble = 7                # 双实线
    wdLineStyleTriple = 8                # 三条细实线
    wdLineStyleThinThickSmallGap = 9     # 里面是一条细实线，外面是一条粗实线，两条线的间隙较小
    wdLineStyleThickThinSmallGap = 10    # 里面是一条粗实线，外面是一条细实线，两条线的间隙较小
    wdLineStyleThinThickThinSmallGap = 11 # 最里面一条细实线，其次一条粗实线，最外面是一条细实线，所有线之间的间隙较小
    wdLineStyleThinThickMedGap = 12       # 里面是一条细实线，外面是一条粗实线，两条线的间隙中等
    wdLineStyleThickThinMedGap = 13       # 里面是一条粗实线，外面是一条细实线，两条线的间隙中等
    wdLineStyleThinThickThinMedGap = 14   # 最里面是一条细实线，其次是一条粗实线，最外面是一条细实线，所有线之间的间隙中等
    wdLineStyleThinThickLargeGap = 15     # 里面是一条细实线，外面是一条粗实线，两条线的间隙较大
    wdLineStyleThickThinLargeGap = 16     # 里面是一条粗实线，外面是一条细实线，两条线的间隙较大
    wdLineStyleThinThickThinLargeGap = 17 # 最里面是一条细实线，其次是一条粗实线，最外面是一条细实线，所有线之间的间隙较大
    wdLineStyleSingleWavy = 18            # 波浪型单实线
    wdLineStyleDoubleWavy = 19            # 波浪型双实线
    wdLineStyleDashDotStroked = 20        # 划线后跟粗点，使边框的外观类似于理发店招牌
    wdLineStyleEmboss3D = 21              # 边框呈现三维阳文效果
    wdLineStyleEngrave3D = 22             # 边框呈现三维阴文效果
    wdLineStyleOutset = 23                # 边框呈现凸起效果
    wdLineStyleInset = 24                 # 边框呈现凹进效果

# 十四、WdLineWidth
class WdLineWidth:
    wdLineWidth025pt = 2    # 0.25 磅
    wdLineWidth050pt = 4    # 0.50 磅
    wdLineWidth075pt = 6    # 0.75 磅
    wdLineWidth100pt = 8    # 1.00 磅，默认值
    wdLineWidth150pt = 12   # 1.50 磅
    wdLineWidth225pt = 18   # 2.25 磅
    wdLineWidth300pt = 24   # 3.00 磅
    wdLineWidth450pt = 36   # 4.50 磅
    wdLineWidth600pt = 48   # 6.00 磅

# 十五、WdColor
class WdColor:
    wdColorAutomatic = -16777216  # 自动配色。默认值；通常为黑色
    wdColorBlack = 0              # 黑色
    wdColorDarkRed = 128          # 深红色
    wdColorRed = 255              # 红色
    wdColorDarkGreen = 13056      # 深绿色
    wdColorOliveGreen = 13107     # 橄榄色
    wdColorBrown = 13209          # 褐色
    wdColorOrange = 26367         # 橙色
    wdColorGreen = 32768          # 绿色
    wdColorDarkYellow = 32896     # 深黄色
    wdColorLightOrange = 39423    # 浅橙色
    wdColorLime = 52377           # 酸橙色
    wdColorGold = 52479           # 金色
    wdColorBrightGreen = 65280    # 鲜绿色
    wdColorYellow = 65535         # 黄色
    wdColorGray95 = 789516        # 95% 灰色底纹
    wdColorGray90 = 1644825       # 90% 灰色底纹
    wdColorGray875 = 2105376      # 87.5% 灰色底纹
    wdColorGray85 = 2500134       # 85% 灰色底纹
    wdColorGray80 = 3355443       # 80% 灰色底纹
    wdColorGray75 = 4210752       # 75% 灰色底纹
    wdColorGray70 = 5000268       # 70% 灰色底纹
    wdColorGray65 = 5855577       # 65% 灰色底纹
    wdColorGray625 = 6316128      # 62.5% 灰色底纹
    wdColorDarkTeal = 6697728     # 深青色
    wdColorPlum = 6697881         # 梅红色
    wdColorGray60 = 6710886       # 60% 灰色底纹
    wdColorSeaGreen = 6723891     # 海绿色
    wdColorGray55 = 7566195       # 55% 灰色底纹
    wdColorDarkBlue = 8388608     # 深蓝色
    wdColorViolet = 8388736       # 紫色
    wdColorTeal = 8421376         # 青色
    wdColorGray50 = 8421504       # 50% 灰色底纹
    wdColorGray45 = 9211020       # 45% 灰色底纹
    wdColorIndigo = 10040115      # 靛蓝色
    wdColorBlueGray = 10053222    # 蓝灰色
    wdColorGray40 = 10066329      # 40% 灰色底纹
    wdColorTan = 10079487         # 棕黄色
    wdColorLightYellow = 10092543 # 浅黄色
    wdColorGray375 = 10526880     # 37.5% 灰色底纹
    wdColorGray35 = 10921638      # 35% 灰色底纹
    wdColorGray30 = 11776947      # 30% 灰色底纹
    wdColorGray25 = 12632256      # 25% 灰色底纹
    wdColorRose = 13408767        # 玫瑰色
    wdColorAqua = 13421619        # 水绿色
    wdColorGray20 = 13421772      # 20% 灰色底纹
    wdColorLightGreen = 13434828  # 浅绿色
    wdColorGray15 = 14277081      # 15% 灰色底纹
    wdColorGray125 = 14737632     # 12.5% 灰色底纹
    wdColorGray10 = 15132390      # 10% 灰色底纹
    wdColorGray05 = 15987699      # 5% 灰色底纹
    wdColorBlue = 16711680        # 蓝色
    wdColorPink = 16711935        # 粉红色
    wdColorLightBlue = 16737843   # 浅蓝色
    wdColorLavender = 16751052    # 淡紫色
    wdColorSkyBlue = 16763904     # 天蓝色
    wdColorPaleBlue = 16764057    # 淡蓝色
    wdColorTurquoise = 16776960   # 青绿色
    wdColorLightTurquoise = 16777164 # 浅青绿色
    wdColorWhite = 16777215       # 白色

class WdPageNumberStyle:
    wdPageNumberStyleArabic = 0            # 阿拉伯语样式
    wdPageNumberStyleUppercaseRoman = 1    # 大写罗马样式
    wdPageNumberStyleLowercaseRoman = 2    # 小写罗马样式
    wdPageNumberStyleUppercaseLetter = 3   # 大写字母样式
    wdPageNumberStyleLowercaseLetter = 4   # 小写字母样式
    wdPageNumberStyleKanji = 10            # 日语汉字样式
    wdPageNumberStyleKanjiDigit = 11       # 日语汉字数字样式
    wdPageNumberStyleArabicFullWidth = 14  # 阿拉伯语全角样式
    wdPageNumberStyleKanjiTraditional = 16 # 日语汉字传统样式
    wdPageNumberStyleNumberInCircle = 18   # 带圈数字样式
    wdPageNumberStyleTradChinNum1 = 33     # 繁体中文数字 1 样式
    wdPageNumberStyleTradChinNum2 = 34     # 繁体中文数字 2 样式
    wdPageNumberStyleSimpChinNum1 = 37     # 简体中文数字 1 样式
    wdPageNumberStyleSimpChinNum2 = 38     # 简体中文数字 2 样式
    wdPageNumberStyleHanjaRead = 41        # 朝鲜文汉字读取样式
    wdPageNumberStyleHanjaReadDigit = 42   # 朝鲜文汉字读取数字样式
    wdPageNumberStyleHebrewLetter1 = 45    # 希伯来语字母 1 样式
    wdPageNumberStyleArabicLetter1 = 46    # 阿拉伯语字母 1 样式
    wdPageNumberStyleHebrewLetter2 = 47    # 希伯来语字母 2 样式
    wdPageNumberStyleArabicLetter2 = 48    # 阿拉伯语字母 2 样式
    wdPageNumberStyleHindiLetter1 = 49     # 印地语字母 1 样式
    wdPageNumberStyleHindiLetter2 = 50     # 印地语字母 2 样式
    wdPageNumberStyleHindiArabic = 51      # 印地语/阿拉伯语样式
    wdPageNumberStyleHindiCardinalText = 52 # 印地语基数文本样式
    wdPageNumberStyleThaiLetter = 53       # 泰语字母样式
    wdPageNumberStyleThaiArabic = 54       # 泰语/阿拉伯语样式
    wdPageNumberStyleThaiCardinalText = 55 # 泰语基数文本样式
    wdPageNumberStyleVietCardinalText = 56 # 越南语基数文本样式
    wdPageNumberStyleNumberInDash = 57     # 带划线数字样式

class WdPageNumberAlignment:
    wdAlignPageNumberLeft = 0    # 左对齐
    wdAlignPageNumberCenter = 1  # 居中
    wdAlignPageNumberRight = 2   # 右对齐
    wdAlignPageNumberInside = 3  # 只在页脚内部左对齐
    wdAlignPageNumberOutside = 4 # 只在页脚外部右对齐
