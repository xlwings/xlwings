# Excel constants: We can't use 'from win32com.client import constants' as we're dynamically dispatching
import re

RE_WORD = re.compile('^[A-Z0-9]([A-Z0-9](?![a-z]))*[a-z]*')


def _dump_enum(cls):
    constants = [c for c in dir(cls) if c.startswith("xl")]
    s2i = {}
    for c in constants:
        i = getattr(cls, c)
        c = c[2:]
        words = []
        while c:
            m = RE_WORD.match(c)
            words.append(m.group(0).lower())
            c = c[m.end():]
        s = "_".join(words)
        s2i[s] = i

    from collections import OrderedDict
    s2i = OrderedDict((
        (s, s2i[s])
        for s in sorted(s2i.keys())
    ))
    i2s = OrderedDict((
        (v, k)
        for k, v in s2i.items()
    ))

    import json
    with open("out.txt", "w") as f:
        json.dump({
            'i2s': i2s,
            's2i': s2i,
            'l': list(s2i.keys())
        }, f, indent=4)


class AboveBelow:
    xlAboveAverage = 0  # from enum XlAboveBelow
    xlAboveStdDev = 4  # from enum XlAboveBelow
    xlBelowAverage = 1  # from enum XlAboveBelow
    xlBelowStdDev = 5  # from enum XlAboveBelow
    xlEqualAboveAverage = 2  # from enum XlAboveBelow
    xlEqualBelowAverage = 3  # from enum XlAboveBelow


class Constants:
    xlNextToAxis = 4  # from enum Constants
    xlNoDocuments = 3  # from enum Constants
    xlNone = -4142  # from enum Constants
    xlNotes = -4144  # from enum Constants
    xlOff = -4146  # from enum Constants
    xl3DEffects1 = 13  # from enum Constants
    xl3DBar = -4099  # from enum Constants
    xl3DEffects2 = 14  # from enum Constants
    xl3DSurface = -4103  # from enum Constants
    xlAbove = 0  # from enum Constants
    xlAccounting1 = 4  # from enum Constants
    xlAccounting2 = 5  # from enum Constants
    xlAccounting3 = 6  # from enum Constants
    xlAccounting4 = 17  # from enum Constants
    xlAdd = 2  # from enum Constants
    xlAll = -4104  # from enum Constants
    xlAllExceptBorders = 7  # from enum Constants
    xlAutomatic = -4105  # from enum Constants
    xlBar = 2  # from enum Constants
    xlBelow = 1  # from enum Constants
    xlBidi = -5000  # from enum Constants
    xlBidiCalendar = 3  # from enum Constants
    xlBoth = 1  # from enum Constants
    xlBottom = -4107  # from enum Constants
    xlCascade = 7  # from enum Constants
    xlCenter = -4108  # from enum Constants
    xlCenterAcrossSelection = 7  # from enum Constants
    xlChart4 = 2  # from enum Constants
    xlChartSeries = 17  # from enum Constants
    xlChartShort = 6  # from enum Constants
    xlChartTitles = 18  # from enum Constants
    xlChecker = 9  # from enum Constants
    xlCircle = 8  # from enum Constants
    xlClassic1 = 1  # from enum Constants
    xlClassic2 = 2  # from enum Constants
    xlClassic3 = 3  # from enum Constants
    xlClosed = 3  # from enum Constants
    xlColor1 = 7  # from enum Constants
    xlColor2 = 8  # from enum Constants
    xlColor3 = 9  # from enum Constants
    xlColumn = 3  # from enum Constants
    xlCombination = -4111  # from enum Constants
    xlComplete = 4  # from enum Constants
    xlConstants = 2  # from enum Constants
    xlContents = 2  # from enum Constants
    xlContext = -5002  # from enum Constants
    xlCorner = 2  # from enum Constants
    xlCrissCross = 16  # from enum Constants
    xlCross = 4  # from enum Constants
    xlCustom = -4114  # from enum Constants
    xlDebugCodePane = 13  # from enum Constants
    xlDefaultAutoFormat = -1  # from enum Constants
    xlDesktop = 9  # from enum Constants
    xlDiamond = 2  # from enum Constants
    xlDirect = 1  # from enum Constants
    xlDistributed = -4117  # from enum Constants
    xlDivide = 5  # from enum Constants
    xlDoubleAccounting = 5  # from enum Constants
    xlDoubleClosed = 5  # from enum Constants
    xlDoubleOpen = 4  # from enum Constants
    xlDoubleQuote = 1  # from enum Constants
    xlDrawingObject = 14  # from enum Constants
    xlEntireChart = 20  # from enum Constants
    xlExcelMenus = 1  # from enum Constants
    xlExtended = 3  # from enum Constants
    xlFill = 5  # from enum Constants
    xlFirst = 0  # from enum Constants
    xlFixedValue = 1  # from enum Constants
    xlFloating = 5  # from enum Constants
    xlFormats = -4122  # from enum Constants
    xlFormula = 5  # from enum Constants
    xlFullScript = 1  # from enum Constants
    xlGeneral = 1  # from enum Constants
    xlGray16 = 17  # from enum Constants
    xlGray25 = -4124  # from enum Constants
    xlGray50 = -4125  # from enum Constants
    xlGray75 = -4126  # from enum Constants
    xlGray8 = 18  # from enum Constants
    xlGregorian = 2  # from enum Constants
    xlGrid = 15  # from enum Constants
    xlGridline = 22  # from enum Constants
    xlHigh = -4127  # from enum Constants
    xlHindiNumerals = 3  # from enum Constants
    xlIcons = 1  # from enum Constants
    xlImmediatePane = 12  # from enum Constants
    xlInside = 2  # from enum Constants
    xlInteger = 2  # from enum Constants
    xlJustify = -4130  # from enum Constants
    xlLTR = -5003  # from enum Constants
    xlLast = 1  # from enum Constants
    xlLastCell = 11  # from enum Constants
    xlLatin = -5001  # from enum Constants
    xlLeft = -4131  # from enum Constants
    xlLeftToRight = 2  # from enum Constants
    xlLightDown = 13  # from enum Constants
    xlLightHorizontal = 11  # from enum Constants
    xlLightUp = 14  # from enum Constants
    xlLightVertical = 12  # from enum Constants
    xlList1 = 10  # from enum Constants
    xlList2 = 11  # from enum Constants
    xlList3 = 12  # from enum Constants
    xlLocalFormat1 = 15  # from enum Constants
    xlLocalFormat2 = 16  # from enum Constants
    xlLogicalCursor = 1  # from enum Constants
    xlLong = 3  # from enum Constants
    xlLotusHelp = 2  # from enum Constants
    xlLow = -4134  # from enum Constants
    xlMacrosheetCell = 7  # from enum Constants
    xlManual = -4135  # from enum Constants
    xlMaximum = 2  # from enum Constants
    xlMinimum = 4  # from enum Constants
    xlMinusValues = 3  # from enum Constants
    xlMixed = 2  # from enum Constants
    xlMixedAuthorizedScript = 4  # from enum Constants
    xlMixedScript = 3  # from enum Constants
    xlModule = -4141  # from enum Constants
    xlMultiply = 4  # from enum Constants
    xlNarrow = 1  # from enum Constants
    xlOn = 1  # from enum Constants
    xlOpaque = 3  # from enum Constants
    xlOpen = 2  # from enum Constants
    xlOutside = 3  # from enum Constants
    xlPartial = 3  # from enum Constants
    xlPartialScript = 2  # from enum Constants
    xlPercent = 2  # from enum Constants
    xlPlus = 9  # from enum Constants
    xlPlusValues = 2  # from enum Constants
    xlRTL = -5004  # from enum Constants
    xlReference = 4  # from enum Constants
    xlRight = -4152  # from enum Constants
    xlScale = 3  # from enum Constants
    xlSemiGray75 = 10  # from enum Constants
    xlSemiautomatic = 2  # from enum Constants
    xlShort = 1  # from enum Constants
    xlShowLabel = 4  # from enum Constants
    xlShowLabelAndPercent = 5  # from enum Constants
    xlShowPercent = 3  # from enum Constants
    xlShowValue = 2  # from enum Constants
    xlSimple = -4154  # from enum Constants
    xlSingle = 2  # from enum Constants
    xlSingleAccounting = 4  # from enum Constants
    xlSingleQuote = 2  # from enum Constants
    xlSolid = 1  # from enum Constants
    xlSquare = 1  # from enum Constants
    xlStError = 4  # from enum Constants
    xlStar = 5  # from enum Constants
    xlStrict = 2  # from enum Constants
    xlSubtract = 3  # from enum Constants
    xlSystem = 1  # from enum Constants
    xlTextBox = 16  # from enum Constants
    xlTiled = 1  # from enum Constants
    xlTitleBar = 8  # from enum Constants
    xlToolbar = 1  # from enum Constants
    xlToolbarButton = 2  # from enum Constants
    xlTop = -4160  # from enum Constants
    xlTopToBottom = 1  # from enum Constants
    xlTransparent = 2  # from enum Constants
    xlTriangle = 3  # from enum Constants
    xlVeryHidden = 2  # from enum Constants
    xlVisible = 12  # from enum Constants
    xlVisualCursor = 2  # from enum Constants
    xlWatchPane = 11  # from enum Constants
    xlWide = 3  # from enum Constants
    xlWorkbookTab = 6  # from enum Constants
    xlWorksheet4 = 1  # from enum Constants
    xlWorksheetCell = 3  # from enum Constants
    xlWorksheetShort = 5  # from enum Constants


class ActionType:
    xlActionTypeDrillthrough = 256  # from enum XlActionType
    xlActionTypeReport = 128  # from enum XlActionType
    xlActionTypeRowset = 16  # from enum XlActionType
    xlActionTypeUrl = 1  # from enum XlActionType


class Allocation:
    xlAutomaticAllocation = 2  # from enum XlAllocation
    xlManualAllocation = 1  # from enum XlAllocation


class AllocationMethod:
    xlEqualAllocation = 1  # from enum XlAllocationMethod
    xlWeightedAllocation = 2  # from enum XlAllocationMethod


class AllocationValue:
    xlAllocateIncrement = 2  # from enum XlAllocationValue
    xlAllocateValue = 1  # from enum XlAllocationValue


class ApplicationInternational:
    xl24HourClock = 33  # from enum XlApplicationInternational
    xl4DigitYears = 43  # from enum XlApplicationInternational
    xlAlternateArraySeparator = 16  # from enum XlApplicationInternational
    xlColumnSeparator = 14  # from enum XlApplicationInternational
    xlCountryCode = 1  # from enum XlApplicationInternational
    xlCountrySetting = 2  # from enum XlApplicationInternational
    xlCurrencyBefore = 37  # from enum XlApplicationInternational
    xlCurrencyCode = 25  # from enum XlApplicationInternational
    xlCurrencyDigits = 27  # from enum XlApplicationInternational
    xlCurrencyLeadingZeros = 40  # from enum XlApplicationInternational
    xlCurrencyMinusSign = 38  # from enum XlApplicationInternational
    xlCurrencyNegative = 28  # from enum XlApplicationInternational
    xlCurrencySpaceBefore = 36  # from enum XlApplicationInternational
    xlCurrencyTrailingZeros = 39  # from enum XlApplicationInternational
    xlDateOrder = 32  # from enum XlApplicationInternational
    xlDateSeparator = 17  # from enum XlApplicationInternational
    xlDayCode = 21  # from enum XlApplicationInternational
    xlDayLeadingZero = 42  # from enum XlApplicationInternational
    xlDecimalSeparator = 3  # from enum XlApplicationInternational
    xlGeneralFormatName = 26  # from enum XlApplicationInternational
    xlHourCode = 22  # from enum XlApplicationInternational
    xlLeftBrace = 12  # from enum XlApplicationInternational
    xlLeftBracket = 10  # from enum XlApplicationInternational
    xlListSeparator = 5  # from enum XlApplicationInternational
    xlLowerCaseColumnLetter = 9  # from enum XlApplicationInternational
    xlLowerCaseRowLetter = 8  # from enum XlApplicationInternational
    xlMDY = 44  # from enum XlApplicationInternational
    xlMetric = 35  # from enum XlApplicationInternational
    xlMinuteCode = 23  # from enum XlApplicationInternational
    xlMonthCode = 20  # from enum XlApplicationInternational
    xlMonthLeadingZero = 41  # from enum XlApplicationInternational
    xlMonthNameChars = 30  # from enum XlApplicationInternational
    xlNonEnglishFunctions = 34  # from enum XlApplicationInternational
    xlNoncurrencyDigits = 29  # from enum XlApplicationInternational
    xlRightBrace = 13  # from enum XlApplicationInternational
    xlRightBracket = 11  # from enum XlApplicationInternational
    xlRowSeparator = 15  # from enum XlApplicationInternational
    xlSecondCode = 24  # from enum XlApplicationInternational
    xlThousandsSeparator = 4  # from enum XlApplicationInternational
    xlTimeLeadingZero = 45  # from enum XlApplicationInternational
    xlTimeSeparator = 18  # from enum XlApplicationInternational
    xlUpperCaseColumnLetter = 7  # from enum XlApplicationInternational
    xlUpperCaseRowLetter = 6  # from enum XlApplicationInternational
    xlWeekdayNameChars = 31  # from enum XlApplicationInternational
    xlYearCode = 19  # from enum XlApplicationInternational


class ApplyNamesOrder:
    xlColumnThenRow = 2  # from enum XlApplyNamesOrder
    xlRowThenColumn = 1  # from enum XlApplyNamesOrder


class ArabicModes:
    xlArabicBothStrict = 3  # from enum XlArabicModes
    xlArabicNone = 0  # from enum XlArabicModes
    xlArabicStrictAlefHamza = 1  # from enum XlArabicModes
    xlArabicStrictFinalYaa = 2  # from enum XlArabicModes


class ArrangeStyle:
    xlArrangeStyleCascade = 7  # from enum XlArrangeStyle
    xlArrangeStyleHorizontal = -4128  # from enum XlArrangeStyle
    xlArrangeStyleTiled = 1  # from enum XlArrangeStyle
    xlArrangeStyleVertical = -4166  # from enum XlArrangeStyle


class ArrowHeadLength:
    xlArrowHeadLengthLong = 3  # from enum XlArrowHeadLength
    xlArrowHeadLengthMedium = -4138  # from enum XlArrowHeadLength
    xlArrowHeadLengthShort = 1  # from enum XlArrowHeadLength


class ArrowHeadStyle:
    xlArrowHeadStyleClosed = 3  # from enum XlArrowHeadStyle
    xlArrowHeadStyleDoubleClosed = 5  # from enum XlArrowHeadStyle
    xlArrowHeadStyleDoubleOpen = 4  # from enum XlArrowHeadStyle
    xlArrowHeadStyleNone = -4142  # from enum XlArrowHeadStyle
    xlArrowHeadStyleOpen = 2  # from enum XlArrowHeadStyle


class ArrowHeadWidth:
    xlArrowHeadWidthMedium = -4138  # from enum XlArrowHeadWidth
    xlArrowHeadWidthNarrow = 1  # from enum XlArrowHeadWidth
    xlArrowHeadWidthWide = 3  # from enum XlArrowHeadWidth


class AutoFillType:
    xlFillCopy = 1  # from enum XlAutoFillType
    xlFillDays = 5  # from enum XlAutoFillType
    xlFillDefault = 0  # from enum XlAutoFillType
    xlFillFormats = 3  # from enum XlAutoFillType
    xlFillMonths = 7  # from enum XlAutoFillType
    xlFillSeries = 2  # from enum XlAutoFillType
    xlFillValues = 4  # from enum XlAutoFillType
    xlFillWeekdays = 6  # from enum XlAutoFillType
    xlFillYears = 8  # from enum XlAutoFillType
    xlGrowthTrend = 10  # from enum XlAutoFillType
    xlLinearTrend = 9  # from enum XlAutoFillType


class AutoFilterOperator:
    xlAnd = 1  # from enum XlAutoFilterOperator
    xlBottom10Items = 4  # from enum XlAutoFilterOperator
    xlBottom10Percent = 6  # from enum XlAutoFilterOperator
    xlFilterAutomaticFontColor = 13  # from enum XlAutoFilterOperator
    xlFilterCellColor = 8  # from enum XlAutoFilterOperator
    xlFilterDynamic = 11  # from enum XlAutoFilterOperator
    xlFilterFontColor = 9  # from enum XlAutoFilterOperator
    xlFilterIcon = 10  # from enum XlAutoFilterOperator
    xlFilterNoFill = 12  # from enum XlAutoFilterOperator
    xlFilterNoIcon = 14  # from enum XlAutoFilterOperator
    xlFilterValues = 7  # from enum XlAutoFilterOperator
    xlOr = 2  # from enum XlAutoFilterOperator
    xlTop10Items = 3  # from enum XlAutoFilterOperator
    xlTop10Percent = 5  # from enum XlAutoFilterOperator


class AxisCrosses:
    xlAxisCrossesAutomatic = -4105  # from enum XlAxisCrosses
    xlAxisCrossesCustom = -4114  # from enum XlAxisCrosses
    xlAxisCrossesMaximum = 2  # from enum XlAxisCrosses
    xlAxisCrossesMinimum = 4  # from enum XlAxisCrosses


class AxisGroup:
    xlPrimary = 1  # from enum XlAxisGroup
    xlSecondary = 2  # from enum XlAxisGroup


class AxisType:
    xlCategory = 1  # from enum XlAxisType
    xlSeriesAxis = 3  # from enum XlAxisType
    xlValue = 2  # from enum XlAxisType


class Background:
    xlBackgroundAutomatic = -4105  # from enum XlBackground
    xlBackgroundOpaque = 3  # from enum XlBackground
    xlBackgroundTransparent = 2  # from enum XlBackground


class BarShape:
    xlBox = 0  # from enum XlBarShape
    xlConeToMax = 5  # from enum XlBarShape
    xlConeToPoint = 4  # from enum XlBarShape
    xlCylinder = 3  # from enum XlBarShape
    xlPyramidToMax = 2  # from enum XlBarShape
    xlPyramidToPoint = 1  # from enum XlBarShape


class BorderWeight:
    xlHairline = 1  # from enum XlBorderWeight
    xlMedium = -4138  # from enum XlBorderWeight
    xlThick = 4  # from enum XlBorderWeight
    xlThin = 2  # from enum XlBorderWeight


class BordersIndex:
    xlDiagonalDown = 5  # from enum XlBordersIndex
    xlDiagonalUp = 6  # from enum XlBordersIndex
    xlEdgeBottom = 9  # from enum XlBordersIndex
    xlEdgeLeft = 7  # from enum XlBordersIndex
    xlEdgeRight = 10  # from enum XlBordersIndex
    xlEdgeTop = 8  # from enum XlBordersIndex
    xlInsideHorizontal = 12  # from enum XlBordersIndex
    xlInsideVertical = 11  # from enum XlBordersIndex


class BuiltInDialog:
    _xlDialogChartSourceData = 541  # from enum XlBuiltInDialog
    _xlDialogPhonetic = 538  # from enum XlBuiltInDialog
    xlDialogActivate = 103  # from enum XlBuiltInDialog
    xlDialogActiveCellFont = 476  # from enum XlBuiltInDialog
    xlDialogAddChartAutoformat = 390  # from enum XlBuiltInDialog
    xlDialogAddinManager = 321  # from enum XlBuiltInDialog
    xlDialogAlignment = 43  # from enum XlBuiltInDialog
    xlDialogAppMove = 170  # from enum XlBuiltInDialog
    xlDialogAppSize = 171  # from enum XlBuiltInDialog
    xlDialogApplyNames = 133  # from enum XlBuiltInDialog
    xlDialogApplyStyle = 212  # from enum XlBuiltInDialog
    xlDialogArrangeAll = 12  # from enum XlBuiltInDialog
    xlDialogAssignToObject = 213  # from enum XlBuiltInDialog
    xlDialogAssignToTool = 293  # from enum XlBuiltInDialog
    xlDialogAttachText = 80  # from enum XlBuiltInDialog
    xlDialogAttachToolbars = 323  # from enum XlBuiltInDialog
    xlDialogAutoCorrect = 485  # from enum XlBuiltInDialog
    xlDialogAxes = 78  # from enum XlBuiltInDialog
    xlDialogBorder = 45  # from enum XlBuiltInDialog
    xlDialogCalculation = 32  # from enum XlBuiltInDialog
    xlDialogCellProtection = 46  # from enum XlBuiltInDialog
    xlDialogChangeLink = 166  # from enum XlBuiltInDialog
    xlDialogChartAddData = 392  # from enum XlBuiltInDialog
    xlDialogChartLocation = 527  # from enum XlBuiltInDialog
    xlDialogChartOptionsDataLabelMultiple = 724  # from enum XlBuiltInDialog
    xlDialogChartOptionsDataLabels = 505  # from enum XlBuiltInDialog
    xlDialogChartOptionsDataTable = 506  # from enum XlBuiltInDialog
    xlDialogChartSourceData = 540  # from enum XlBuiltInDialog
    xlDialogChartTrend = 350  # from enum XlBuiltInDialog
    xlDialogChartType = 526  # from enum XlBuiltInDialog
    xlDialogChartWizard = 288  # from enum XlBuiltInDialog
    xlDialogCheckboxProperties = 435  # from enum XlBuiltInDialog
    xlDialogClear = 52  # from enum XlBuiltInDialog
    xlDialogColorPalette = 161  # from enum XlBuiltInDialog
    xlDialogColumnWidth = 47  # from enum XlBuiltInDialog
    xlDialogCombination = 73  # from enum XlBuiltInDialog
    xlDialogConditionalFormatting = 583  # from enum XlBuiltInDialog
    xlDialogConsolidate = 191  # from enum XlBuiltInDialog
    xlDialogCopyChart = 147  # from enum XlBuiltInDialog
    xlDialogCopyPicture = 108  # from enum XlBuiltInDialog
    xlDialogCreateList = 796  # from enum XlBuiltInDialog
    xlDialogCreateNames = 62  # from enum XlBuiltInDialog
    xlDialogCreatePublisher = 217  # from enum XlBuiltInDialog
    xlDialogCustomViews = 493  # from enum XlBuiltInDialog
    xlDialogCustomizeToolbar = 276  # from enum XlBuiltInDialog
    xlDialogDataDelete = 36  # from enum XlBuiltInDialog
    xlDialogDataLabel = 379  # from enum XlBuiltInDialog
    xlDialogDataLabelMultiple = 723  # from enum XlBuiltInDialog
    xlDialogDataSeries = 40  # from enum XlBuiltInDialog
    xlDialogDataValidation = 525  # from enum XlBuiltInDialog
    xlDialogDefineName = 61  # from enum XlBuiltInDialog
    xlDialogDefineStyle = 229  # from enum XlBuiltInDialog
    xlDialogDeleteFormat = 111  # from enum XlBuiltInDialog
    xlDialogDeleteName = 110  # from enum XlBuiltInDialog
    xlDialogDemote = 203  # from enum XlBuiltInDialog
    xlDialogDisplay = 27  # from enum XlBuiltInDialog
    xlDialogDocumentInspector = 862  # from enum XlBuiltInDialog
    xlDialogEditColor = 223  # from enum XlBuiltInDialog
    xlDialogEditDelete = 54  # from enum XlBuiltInDialog
    xlDialogEditSeries = 228  # from enum XlBuiltInDialog
    xlDialogEditboxProperties = 438  # from enum XlBuiltInDialog
    xlDialogEditionOptions = 251  # from enum XlBuiltInDialog
    xlDialogErrorChecking = 732  # from enum XlBuiltInDialog
    xlDialogErrorbarX = 463  # from enum XlBuiltInDialog
    xlDialogErrorbarY = 464  # from enum XlBuiltInDialog
    xlDialogEvaluateFormula = 709  # from enum XlBuiltInDialog
    xlDialogExternalDataProperties = 530  # from enum XlBuiltInDialog
    xlDialogExtract = 35  # from enum XlBuiltInDialog
    xlDialogFileDelete = 6  # from enum XlBuiltInDialog
    xlDialogFileSharing = 481  # from enum XlBuiltInDialog
    xlDialogFillGroup = 200  # from enum XlBuiltInDialog
    xlDialogFillWorkgroup = 301  # from enum XlBuiltInDialog
    xlDialogFilter = 447  # from enum XlBuiltInDialog
    xlDialogFilterAdvanced = 370  # from enum XlBuiltInDialog
    xlDialogFindFile = 475  # from enum XlBuiltInDialog
    xlDialogFont = 26  # from enum XlBuiltInDialog
    xlDialogFontProperties = 381  # from enum XlBuiltInDialog
    xlDialogFormatAuto = 269  # from enum XlBuiltInDialog
    xlDialogFormatChart = 465  # from enum XlBuiltInDialog
    xlDialogFormatCharttype = 423  # from enum XlBuiltInDialog
    xlDialogFormatFont = 150  # from enum XlBuiltInDialog
    xlDialogFormatLegend = 88  # from enum XlBuiltInDialog
    xlDialogFormatMain = 225  # from enum XlBuiltInDialog
    xlDialogFormatMove = 128  # from enum XlBuiltInDialog
    xlDialogFormatNumber = 42  # from enum XlBuiltInDialog
    xlDialogFormatOverlay = 226  # from enum XlBuiltInDialog
    xlDialogFormatSize = 129  # from enum XlBuiltInDialog
    xlDialogFormatText = 89  # from enum XlBuiltInDialog
    xlDialogFormulaFind = 64  # from enum XlBuiltInDialog
    xlDialogFormulaGoto = 63  # from enum XlBuiltInDialog
    xlDialogFormulaReplace = 130  # from enum XlBuiltInDialog
    xlDialogFunctionWizard = 450  # from enum XlBuiltInDialog
    xlDialogGallery3dArea = 193  # from enum XlBuiltInDialog
    xlDialogGallery3dBar = 272  # from enum XlBuiltInDialog
    xlDialogGallery3dColumn = 194  # from enum XlBuiltInDialog
    xlDialogGallery3dLine = 195  # from enum XlBuiltInDialog
    xlDialogGallery3dPie = 196  # from enum XlBuiltInDialog
    xlDialogGallery3dSurface = 273  # from enum XlBuiltInDialog
    xlDialogGalleryArea = 67  # from enum XlBuiltInDialog
    xlDialogGalleryBar = 68  # from enum XlBuiltInDialog
    xlDialogGalleryColumn = 69  # from enum XlBuiltInDialog
    xlDialogGalleryCustom = 388  # from enum XlBuiltInDialog
    xlDialogGalleryDoughnut = 344  # from enum XlBuiltInDialog
    xlDialogGalleryLine = 70  # from enum XlBuiltInDialog
    xlDialogGalleryPie = 71  # from enum XlBuiltInDialog
    xlDialogGalleryRadar = 249  # from enum XlBuiltInDialog
    xlDialogGalleryScatter = 72  # from enum XlBuiltInDialog
    xlDialogGoalSeek = 198  # from enum XlBuiltInDialog
    xlDialogGridlines = 76  # from enum XlBuiltInDialog
    xlDialogImportTextFile = 666  # from enum XlBuiltInDialog
    xlDialogInsert = 55  # from enum XlBuiltInDialog
    xlDialogInsertHyperlink = 596  # from enum XlBuiltInDialog
    xlDialogInsertNameLabel = 496  # from enum XlBuiltInDialog
    xlDialogInsertObject = 259  # from enum XlBuiltInDialog
    xlDialogInsertPicture = 342  # from enum XlBuiltInDialog
    xlDialogInsertTitle = 380  # from enum XlBuiltInDialog
    xlDialogLabelProperties = 436  # from enum XlBuiltInDialog
    xlDialogListboxProperties = 437  # from enum XlBuiltInDialog
    xlDialogMacroOptions = 382  # from enum XlBuiltInDialog
    xlDialogMailEditMailer = 470  # from enum XlBuiltInDialog
    xlDialogMailLogon = 339  # from enum XlBuiltInDialog
    xlDialogMailNextLetter = 378  # from enum XlBuiltInDialog
    xlDialogMainChart = 85  # from enum XlBuiltInDialog
    xlDialogMainChartType = 185  # from enum XlBuiltInDialog
    xlDialogMenuEditor = 322  # from enum XlBuiltInDialog
    xlDialogMove = 262  # from enum XlBuiltInDialog
    xlDialogMyPermission = 834  # from enum XlBuiltInDialog
    xlDialogNameManager = 977  # from enum XlBuiltInDialog
    xlDialogNew = 119  # from enum XlBuiltInDialog
    xlDialogNewName = 978  # from enum XlBuiltInDialog
    xlDialogNewWebQuery = 667  # from enum XlBuiltInDialog
    xlDialogNote = 154  # from enum XlBuiltInDialog
    xlDialogObjectProperties = 207  # from enum XlBuiltInDialog
    xlDialogObjectProtection = 214  # from enum XlBuiltInDialog
    xlDialogOpen = 1  # from enum XlBuiltInDialog
    xlDialogOpenLinks = 2  # from enum XlBuiltInDialog
    xlDialogOpenMail = 188  # from enum XlBuiltInDialog
    xlDialogOpenText = 441  # from enum XlBuiltInDialog
    xlDialogOptionsCalculation = 318  # from enum XlBuiltInDialog
    xlDialogOptionsChart = 325  # from enum XlBuiltInDialog
    xlDialogOptionsEdit = 319  # from enum XlBuiltInDialog
    xlDialogOptionsGeneral = 356  # from enum XlBuiltInDialog
    xlDialogOptionsListsAdd = 458  # from enum XlBuiltInDialog
    xlDialogOptionsME = 647  # from enum XlBuiltInDialog
    xlDialogOptionsTransition = 355  # from enum XlBuiltInDialog
    xlDialogOptionsView = 320  # from enum XlBuiltInDialog
    xlDialogOutline = 142  # from enum XlBuiltInDialog
    xlDialogOverlay = 86  # from enum XlBuiltInDialog
    xlDialogOverlayChartType = 186  # from enum XlBuiltInDialog
    xlDialogPageSetup = 7  # from enum XlBuiltInDialog
    xlDialogParse = 91  # from enum XlBuiltInDialog
    xlDialogPasteNames = 58  # from enum XlBuiltInDialog
    xlDialogPasteSpecial = 53  # from enum XlBuiltInDialog
    xlDialogPatterns = 84  # from enum XlBuiltInDialog
    xlDialogPermission = 832  # from enum XlBuiltInDialog
    xlDialogPhonetic = 656  # from enum XlBuiltInDialog
    xlDialogPivotCalculatedField = 570  # from enum XlBuiltInDialog
    xlDialogPivotCalculatedItem = 572  # from enum XlBuiltInDialog
    xlDialogPivotClientServerSet = 689  # from enum XlBuiltInDialog
    xlDialogPivotFieldGroup = 433  # from enum XlBuiltInDialog
    xlDialogPivotFieldProperties = 313  # from enum XlBuiltInDialog
    xlDialogPivotFieldUngroup = 434  # from enum XlBuiltInDialog
    xlDialogPivotShowPages = 421  # from enum XlBuiltInDialog
    xlDialogPivotSolveOrder = 568  # from enum XlBuiltInDialog
    xlDialogPivotTableOptions = 567  # from enum XlBuiltInDialog
    xlDialogPivotTableSlicerConnections = 1183  # from enum XlBuiltInDialog
    xlDialogPivotTableWhatIfAnalysisSettings = 1153  # from enum XlBuiltInDialog
    xlDialogPivotTableWizard = 312  # from enum XlBuiltInDialog
    xlDialogPlacement = 300  # from enum XlBuiltInDialog
    xlDialogPrint = 8  # from enum XlBuiltInDialog
    xlDialogPrintPreview = 222  # from enum XlBuiltInDialog
    xlDialogPrinterSetup = 9  # from enum XlBuiltInDialog
    xlDialogPromote = 202  # from enum XlBuiltInDialog
    xlDialogProperties = 474  # from enum XlBuiltInDialog
    xlDialogPropertyFields = 754  # from enum XlBuiltInDialog
    xlDialogProtectDocument = 28  # from enum XlBuiltInDialog
    xlDialogProtectSharing = 620  # from enum XlBuiltInDialog
    xlDialogPublishAsWebPage = 653  # from enum XlBuiltInDialog
    xlDialogPushbuttonProperties = 445  # from enum XlBuiltInDialog
    xlDialogReplaceFont = 134  # from enum XlBuiltInDialog
    xlDialogRoutingSlip = 336  # from enum XlBuiltInDialog
    xlDialogRowHeight = 127  # from enum XlBuiltInDialog
    xlDialogRun = 17  # from enum XlBuiltInDialog
    xlDialogSaveAs = 5  # from enum XlBuiltInDialog
    xlDialogSaveCopyAs = 456  # from enum XlBuiltInDialog
    xlDialogSaveNewObject = 208  # from enum XlBuiltInDialog
    xlDialogSaveWorkbook = 145  # from enum XlBuiltInDialog
    xlDialogSaveWorkspace = 285  # from enum XlBuiltInDialog
    xlDialogScale = 87  # from enum XlBuiltInDialog
    xlDialogScenarioAdd = 307  # from enum XlBuiltInDialog
    xlDialogScenarioCells = 305  # from enum XlBuiltInDialog
    xlDialogScenarioEdit = 308  # from enum XlBuiltInDialog
    xlDialogScenarioMerge = 473  # from enum XlBuiltInDialog
    xlDialogScenarioSummary = 311  # from enum XlBuiltInDialog
    xlDialogScrollbarProperties = 420  # from enum XlBuiltInDialog
    xlDialogSearch = 731  # from enum XlBuiltInDialog
    xlDialogSelectSpecial = 132  # from enum XlBuiltInDialog
    xlDialogSendMail = 189  # from enum XlBuiltInDialog
    xlDialogSeriesAxes = 460  # from enum XlBuiltInDialog
    xlDialogSeriesOptions = 557  # from enum XlBuiltInDialog
    xlDialogSeriesOrder = 466  # from enum XlBuiltInDialog
    xlDialogSeriesShape = 504  # from enum XlBuiltInDialog
    xlDialogSeriesX = 461  # from enum XlBuiltInDialog
    xlDialogSeriesY = 462  # from enum XlBuiltInDialog
    xlDialogSetBackgroundPicture = 509  # from enum XlBuiltInDialog
    xlDialogSetMDXEditor = 1208  # from enum XlBuiltInDialog
    xlDialogSetManager = 1109  # from enum XlBuiltInDialog
    xlDialogSetPrintTitles = 23  # from enum XlBuiltInDialog
    xlDialogSetTupleEditorOnColumns = 1108  # from enum XlBuiltInDialog
    xlDialogSetTupleEditorOnRows = 1107  # from enum XlBuiltInDialog
    xlDialogSetUpdateStatus = 159  # from enum XlBuiltInDialog
    xlDialogShowDetail = 204  # from enum XlBuiltInDialog
    xlDialogShowToolbar = 220  # from enum XlBuiltInDialog
    xlDialogSize = 261  # from enum XlBuiltInDialog
    xlDialogSlicerCreation = 1182  # from enum XlBuiltInDialog
    xlDialogSlicerPivotTableConnections = 1184  # from enum XlBuiltInDialog
    xlDialogSlicerSettings = 1179  # from enum XlBuiltInDialog
    xlDialogSort = 39  # from enum XlBuiltInDialog
    xlDialogSortSpecial = 192  # from enum XlBuiltInDialog
    xlDialogSparklineInsertColumn = 1134  # from enum XlBuiltInDialog
    xlDialogSparklineInsertLine = 1133  # from enum XlBuiltInDialog
    xlDialogSparklineInsertWinLoss = 1135  # from enum XlBuiltInDialog
    xlDialogSplit = 137  # from enum XlBuiltInDialog
    xlDialogStandardFont = 190  # from enum XlBuiltInDialog
    xlDialogStandardWidth = 472  # from enum XlBuiltInDialog
    xlDialogStyle = 44  # from enum XlBuiltInDialog
    xlDialogSubscribeTo = 218  # from enum XlBuiltInDialog
    xlDialogSubtotalCreate = 398  # from enum XlBuiltInDialog
    xlDialogSummaryInfo = 474  # from enum XlBuiltInDialog
    xlDialogTabOrder = 394  # from enum XlBuiltInDialog
    xlDialogTable = 41  # from enum XlBuiltInDialog
    xlDialogTextToColumns = 422  # from enum XlBuiltInDialog
    xlDialogUnhide = 94  # from enum XlBuiltInDialog
    xlDialogUpdateLink = 201  # from enum XlBuiltInDialog
    xlDialogVbaInsertFile = 328  # from enum XlBuiltInDialog
    xlDialogVbaMakeAddin = 478  # from enum XlBuiltInDialog
    xlDialogVbaProcedureDefinition = 330  # from enum XlBuiltInDialog
    xlDialogView3d = 197  # from enum XlBuiltInDialog
    xlDialogWebOptionsBrowsers = 773  # from enum XlBuiltInDialog
    xlDialogWebOptionsEncoding = 686  # from enum XlBuiltInDialog
    xlDialogWebOptionsFiles = 684  # from enum XlBuiltInDialog
    xlDialogWebOptionsFonts = 687  # from enum XlBuiltInDialog
    xlDialogWebOptionsGeneral = 683  # from enum XlBuiltInDialog
    xlDialogWebOptionsPictures = 685  # from enum XlBuiltInDialog
    xlDialogWindowMove = 14  # from enum XlBuiltInDialog
    xlDialogWindowSize = 13  # from enum XlBuiltInDialog
    xlDialogWorkbookAdd = 281  # from enum XlBuiltInDialog
    xlDialogWorkbookCopy = 283  # from enum XlBuiltInDialog
    xlDialogWorkbookInsert = 354  # from enum XlBuiltInDialog
    xlDialogWorkbookMove = 282  # from enum XlBuiltInDialog
    xlDialogWorkbookName = 386  # from enum XlBuiltInDialog
    xlDialogWorkbookNew = 302  # from enum XlBuiltInDialog
    xlDialogWorkbookOptions = 284  # from enum XlBuiltInDialog
    xlDialogWorkbookProtect = 417  # from enum XlBuiltInDialog
    xlDialogWorkbookTabSplit = 415  # from enum XlBuiltInDialog
    xlDialogWorkbookUnhide = 384  # from enum XlBuiltInDialog
    xlDialogWorkgroup = 199  # from enum XlBuiltInDialog
    xlDialogWorkspace = 95  # from enum XlBuiltInDialog
    xlDialogZoom = 256  # from enum XlBuiltInDialog


class CVError:
    xlErrDiv0 = 2007  # from enum XlCVError
    xlErrNA = 2042  # from enum XlCVError
    xlErrName = 2029  # from enum XlCVError
    xlErrNull = 2000  # from enum XlCVError
    xlErrNum = 2036  # from enum XlCVError
    xlErrRef = 2023  # from enum XlCVError
    xlErrValue = 2015  # from enum XlCVError


class CalcFor:
    xlAllValues = 0  # from enum XlCalcFor
    xlColGroups = 2  # from enum XlCalcFor
    xlRowGroups = 1  # from enum XlCalcFor


class CalculatedMemberType:
    xlCalculatedMember = 0  # from enum XlCalculatedMemberType
    xlCalculatedSet = 1  # from enum XlCalculatedMemberType


class Calculation:
    xlCalculationAutomatic = -4105  # from enum XlCalculation
    xlCalculationManual = -4135  # from enum XlCalculation
    xlCalculationSemiautomatic = 2  # from enum XlCalculation


calculations = (
    'automatic',
    'manual',
    'semiautomatic'
)

class CalculationInterruptKey:
    xlAnyKey = 2  # from enum XlCalculationInterruptKey
    xlEscKey = 1  # from enum XlCalculationInterruptKey
    xlNoKey = 0  # from enum XlCalculationInterruptKey


class CalculationState:
    xlCalculating = 1  # from enum XlCalculationState
    xlDone = 0  # from enum XlCalculationState
    xlPending = 2  # from enum XlCalculationState


class CategoryType:
    xlAutomaticScale = -4105  # from enum XlCategoryType
    xlCategoryScale = 2  # from enum XlCategoryType
    xlTimeScale = 3  # from enum XlCategoryType


class CellChangedState:
    xlCellChangeApplied = 3  # from enum XlCellChangedState
    xlCellChanged = 2  # from enum XlCellChangedState
    xlCellNotChanged = 1  # from enum XlCellChangedState


class CellInsertionMode:
    xlInsertDeleteCells = 1  # from enum XlCellInsertionMode
    xlInsertEntireRows = 2  # from enum XlCellInsertionMode
    xlOverwriteCells = 0  # from enum XlCellInsertionMode


class CellType:
    xlCellTypeAllFormatConditions = -4172  # from enum XlCellType
    xlCellTypeAllValidation = -4174  # from enum XlCellType
    xlCellTypeBlanks = 4  # from enum XlCellType
    xlCellTypeComments = -4144  # from enum XlCellType
    xlCellTypeConstants = 2  # from enum XlCellType
    xlCellTypeFormulas = -4123  # from enum XlCellType
    xlCellTypeLastCell = 11  # from enum XlCellType
    xlCellTypeSameFormatConditions = -4173  # from enum XlCellType
    xlCellTypeSameValidation = -4175  # from enum XlCellType
    xlCellTypeVisible = 12  # from enum XlCellType


class ChartElementPosition:
    xlChartElementPositionAutomatic = -4105  # from enum XlChartElementPosition
    xlChartElementPositionCustom = -4114  # from enum XlChartElementPosition


class ChartGallery:
    xlAnyGallery = 23  # from enum XlChartGallery
    xlBuiltIn = 21  # from enum XlChartGallery
    xlUserDefined = 22  # from enum XlChartGallery


class ChartItem:
    xlAxis = 21  # from enum XlChartItem
    xlAxisTitle = 17  # from enum XlChartItem
    xlChartArea = 2  # from enum XlChartItem
    xlChartTitle = 4  # from enum XlChartItem
    xlCorners = 6  # from enum XlChartItem
    xlDataLabel = 0  # from enum XlChartItem
    xlDataTable = 7  # from enum XlChartItem
    xlDisplayUnitLabel = 30  # from enum XlChartItem
    xlDownBars = 20  # from enum XlChartItem
    xlDropLines = 26  # from enum XlChartItem
    xlErrorBars = 9  # from enum XlChartItem
    xlFloor = 23  # from enum XlChartItem
    xlHiLoLines = 25  # from enum XlChartItem
    xlLeaderLines = 29  # from enum XlChartItem
    xlLegend = 24  # from enum XlChartItem
    xlLegendEntry = 12  # from enum XlChartItem
    xlLegendKey = 13  # from enum XlChartItem
    xlMajorGridlines = 15  # from enum XlChartItem
    xlMinorGridlines = 16  # from enum XlChartItem
    xlNothing = 28  # from enum XlChartItem
    xlPivotChartDropZone = 32  # from enum XlChartItem
    xlPivotChartFieldButton = 31  # from enum XlChartItem
    xlPlotArea = 19  # from enum XlChartItem
    xlRadarAxisLabels = 27  # from enum XlChartItem
    xlSeries = 3  # from enum XlChartItem
    xlSeriesLines = 22  # from enum XlChartItem
    xlShape = 14  # from enum XlChartItem
    xlTrendline = 8  # from enum XlChartItem
    xlUpBars = 18  # from enum XlChartItem
    xlWalls = 5  # from enum XlChartItem
    xlXErrorBars = 10  # from enum XlChartItem
    xlYErrorBars = 11  # from enum XlChartItem


class ChartLocation:
    xlLocationAsNewSheet = 1  # from enum XlChartLocation
    xlLocationAsObject = 2  # from enum XlChartLocation
    xlLocationAutomatic = 3  # from enum XlChartLocation


class ChartPicturePlacement:
    xlAllFaces = 7  # from enum XlChartPicturePlacement
    xlEnd = 2  # from enum XlChartPicturePlacement
    xlEndSides = 3  # from enum XlChartPicturePlacement
    xlFront = 4  # from enum XlChartPicturePlacement
    xlFrontEnd = 6  # from enum XlChartPicturePlacement
    xlFrontSides = 5  # from enum XlChartPicturePlacement
    xlSides = 1  # from enum XlChartPicturePlacement


class ChartPictureType:
    xlStack = 2  # from enum XlChartPictureType
    xlStackScale = 3  # from enum XlChartPictureType
    xlStretch = 1  # from enum XlChartPictureType


class ChartSplitType:
    xlSplitByCustomSplit = 4  # from enum XlChartSplitType
    xlSplitByPercentValue = 3  # from enum XlChartSplitType
    xlSplitByPosition = 1  # from enum XlChartSplitType
    xlSplitByValue = 2  # from enum XlChartSplitType


class ChartType:
    xl3DArea = -4098  # from enum XlChartType
    xl3DAreaStacked = 78  # from enum XlChartType
    xl3DAreaStacked100 = 79  # from enum XlChartType
    xl3DBarClustered = 60  # from enum XlChartType
    xl3DBarStacked = 61  # from enum XlChartType
    xl3DBarStacked100 = 62  # from enum XlChartType
    xl3DColumn = -4100  # from enum XlChartType
    xl3DColumnClustered = 54  # from enum XlChartType
    xl3DColumnStacked = 55  # from enum XlChartType
    xl3DColumnStacked100 = 56  # from enum XlChartType
    xl3DLine = -4101  # from enum XlChartType
    xl3DPie = -4102  # from enum XlChartType
    xl3DPieExploded = 70  # from enum XlChartType
    xlArea = 1  # from enum XlChartType
    xlAreaStacked = 76  # from enum XlChartType
    xlAreaStacked100 = 77  # from enum XlChartType
    xlBarClustered = 57  # from enum XlChartType
    xlBarOfPie = 71  # from enum XlChartType
    xlBarStacked = 58  # from enum XlChartType
    xlBarStacked100 = 59  # from enum XlChartType
    xlBubble = 15  # from enum XlChartType
    xlBubble3DEffect = 87  # from enum XlChartType
    xlColumnClustered = 51  # from enum XlChartType
    xlColumnStacked = 52  # from enum XlChartType
    xlColumnStacked100 = 53  # from enum XlChartType
    xlConeBarClustered = 102  # from enum XlChartType
    xlConeBarStacked = 103  # from enum XlChartType
    xlConeBarStacked100 = 104  # from enum XlChartType
    xlConeCol = 105  # from enum XlChartType
    xlConeColClustered = 99  # from enum XlChartType
    xlConeColStacked = 100  # from enum XlChartType
    xlConeColStacked100 = 101  # from enum XlChartType
    xlCylinderBarClustered = 95  # from enum XlChartType
    xlCylinderBarStacked = 96  # from enum XlChartType
    xlCylinderBarStacked100 = 97  # from enum XlChartType
    xlCylinderCol = 98  # from enum XlChartType
    xlCylinderColClustered = 92  # from enum XlChartType
    xlCylinderColStacked = 93  # from enum XlChartType
    xlCylinderColStacked100 = 94  # from enum XlChartType
    xlDoughnut = -4120  # from enum XlChartType
    xlDoughnutExploded = 80  # from enum XlChartType
    xlLine = 4  # from enum XlChartType
    xlLineMarkers = 65  # from enum XlChartType
    xlLineMarkersStacked = 66  # from enum XlChartType
    xlLineMarkersStacked100 = 67  # from enum XlChartType
    xlLineStacked = 63  # from enum XlChartType
    xlLineStacked100 = 64  # from enum XlChartType
    xlPie = 5  # from enum XlChartType
    xlPieExploded = 69  # from enum XlChartType
    xlPieOfPie = 68  # from enum XlChartType
    xlPyramidBarClustered = 109  # from enum XlChartType
    xlPyramidBarStacked = 110  # from enum XlChartType
    xlPyramidBarStacked100 = 111  # from enum XlChartType
    xlPyramidCol = 112  # from enum XlChartType
    xlPyramidColClustered = 106  # from enum XlChartType
    xlPyramidColStacked = 107  # from enum XlChartType
    xlPyramidColStacked100 = 108  # from enum XlChartType
    xlRadar = -4151  # from enum XlChartType
    xlRadarFilled = 82  # from enum XlChartType
    xlRadarMarkers = 81  # from enum XlChartType
    xlStockHLC = 88  # from enum XlChartType
    xlStockOHLC = 89  # from enum XlChartType
    xlStockVHLC = 90  # from enum XlChartType
    xlStockVOHLC = 91  # from enum XlChartType
    xlSurface = 83  # from enum XlChartType
    xlSurfaceTopView = 85  # from enum XlChartType
    xlSurfaceTopViewWireframe = 86  # from enum XlChartType
    xlSurfaceWireframe = 84  # from enum XlChartType
    xlXYScatter = -4169  # from enum XlChartType
    xlXYScatterLines = 74  # from enum XlChartType
    xlXYScatterLinesNoMarkers = 75  # from enum XlChartType
    xlXYScatterSmooth = 72  # from enum XlChartType
    xlXYScatterSmoothNoMarkers = 73  # from enum XlChartType


chart_types = (
    '3d_area',
    '3d_area_stacked',
    '3d_area_stacked_100',
    '3d_bar_clustered',
    '3d_bar_stacked',
    '3d_bar_stacked_100',
    '3d_column',
    '3d_column_clustered',
    '3d_column_stacked',
    '3d_column_stacked_100',
    '3d_line',
    '3d_pie',
    '3d_pie_exploded',
    'area',
    'area_stacked',
    'area_stacked_100',
    'bar_clustered',
    'bar_of_pie',
    'bar_stacked',
    'bar_stacked_100',
    'bubble',
    'bubble_3d_effect',
    'column_clustered',
    'column_stacked',
    'column_stacked_100',
    'combination',
    'cone_bar_clustered',
    'cone_bar_stacked',
    'cone_bar_stacked_100',
    'cone_col',
    'cone_col_clustered',
    'cone_col_stacked',
    'cone_col_stacked_100',
    'cylinder_bar_clustered',
    'cylinder_bar_stacked',
    'cylinder_bar_stacked_100',
    'cylinder_col',
    'cylinder_col_clustered',
    'cylinder_col_stacked',
    'cylinder_col_stacked_100',
    'doughnut',
    'doughnut_exploded',
    'line',
    'line_markers',
    'line_markers_stacked',
    'line_markers_stacked_100',
    'line_stacked',
    'line_stacked_100',
    'pie',
    'pie_exploded',
    'pie_of_pie',
    'pyramid_bar_clustered',
    'pyramid_bar_stacked',
    'pyramid_bar_stacked_100',
    'pyramid_col',
    'pyramid_col_clustered',
    'pyramid_col_stacked',
    'pyramid_col_stacked_100',
    'radar',
    'radar_filled',
    'radar_markers',
    'stock_hlc',
    'stock_ohlc',
    'stock_vhlc',
    'stock_vohlc',
    'surface',
    'surface_top_view',
    'surface_top_view_wireframe',
    'surface_wireframe',
    'xy_scatter',
    'xy_scatter_lines',
    'xy_scatter_lines_no_markers',
    'xy_scatter_smooth',
    'xy_scatter_smooth_no_markers'
)


class CheckInVersionType:
    xlCheckInMajorVersion = 1  # from enum XlCheckInVersionType
    xlCheckInMinorVersion = 0  # from enum XlCheckInVersionType
    xlCheckInOverwriteVersion = 2  # from enum XlCheckInVersionType


class ClipboardFormat:
    xlClipboardFormatBIFF = 8  # from enum XlClipboardFormat
    xlClipboardFormatBIFF12 = 63  # from enum XlClipboardFormat
    xlClipboardFormatBIFF2 = 18  # from enum XlClipboardFormat
    xlClipboardFormatBIFF3 = 20  # from enum XlClipboardFormat
    xlClipboardFormatBIFF4 = 30  # from enum XlClipboardFormat
    xlClipboardFormatBinary = 15  # from enum XlClipboardFormat
    xlClipboardFormatBitmap = 9  # from enum XlClipboardFormat
    xlClipboardFormatCGM = 13  # from enum XlClipboardFormat
    xlClipboardFormatCSV = 5  # from enum XlClipboardFormat
    xlClipboardFormatDIF = 4  # from enum XlClipboardFormat
    xlClipboardFormatDspText = 12  # from enum XlClipboardFormat
    xlClipboardFormatEmbedSource = 22  # from enum XlClipboardFormat
    xlClipboardFormatEmbeddedObject = 21  # from enum XlClipboardFormat
    xlClipboardFormatLink = 11  # from enum XlClipboardFormat
    xlClipboardFormatLinkSource = 23  # from enum XlClipboardFormat
    xlClipboardFormatLinkSourceDesc = 32  # from enum XlClipboardFormat
    xlClipboardFormatMovie = 24  # from enum XlClipboardFormat
    xlClipboardFormatNative = 14  # from enum XlClipboardFormat
    xlClipboardFormatObjectDesc = 31  # from enum XlClipboardFormat
    xlClipboardFormatObjectLink = 19  # from enum XlClipboardFormat
    xlClipboardFormatOwnerLink = 17  # from enum XlClipboardFormat
    xlClipboardFormatPICT = 2  # from enum XlClipboardFormat
    xlClipboardFormatPrintPICT = 3  # from enum XlClipboardFormat
    xlClipboardFormatRTF = 7  # from enum XlClipboardFormat
    xlClipboardFormatSYLK = 6  # from enum XlClipboardFormat
    xlClipboardFormatScreenPICT = 29  # from enum XlClipboardFormat
    xlClipboardFormatStandardFont = 28  # from enum XlClipboardFormat
    xlClipboardFormatStandardScale = 27  # from enum XlClipboardFormat
    xlClipboardFormatTable = 16  # from enum XlClipboardFormat
    xlClipboardFormatText = 0  # from enum XlClipboardFormat
    xlClipboardFormatToolFace = 25  # from enum XlClipboardFormat
    xlClipboardFormatToolFacePICT = 26  # from enum XlClipboardFormat
    xlClipboardFormatVALU = 1  # from enum XlClipboardFormat
    xlClipboardFormatWK1 = 10  # from enum XlClipboardFormat


class CmdType:
    xlCmdCube = 1  # from enum XlCmdType
    xlCmdDefault = 4  # from enum XlCmdType
    xlCmdList = 5  # from enum XlCmdType
    xlCmdSql = 2  # from enum XlCmdType
    xlCmdTable = 3  # from enum XlCmdType


class ColorIndex:
    xlColorIndexAutomatic = -4105  # from enum XlColorIndex
    xlColorIndexNone = -4142  # from enum XlColorIndex


class ColumnDataType:
    xlDMYFormat = 4  # from enum XlColumnDataType
    xlDYMFormat = 7  # from enum XlColumnDataType
    xlEMDFormat = 10  # from enum XlColumnDataType
    xlGeneralFormat = 1  # from enum XlColumnDataType
    xlMDYFormat = 3  # from enum XlColumnDataType
    xlMYDFormat = 6  # from enum XlColumnDataType
    xlSkipColumn = 9  # from enum XlColumnDataType
    xlTextFormat = 2  # from enum XlColumnDataType
    xlYDMFormat = 8  # from enum XlColumnDataType
    xlYMDFormat = 5  # from enum XlColumnDataType


class CommandUnderlines:
    xlCommandUnderlinesAutomatic = -4105  # from enum XlCommandUnderlines
    xlCommandUnderlinesOff = -4146  # from enum XlCommandUnderlines
    xlCommandUnderlinesOn = 1  # from enum XlCommandUnderlines


class CommentDisplayMode:
    xlCommentAndIndicator = 1  # from enum XlCommentDisplayMode
    xlCommentIndicatorOnly = -1  # from enum XlCommentDisplayMode
    xlNoIndicator = 0  # from enum XlCommentDisplayMode


class ConditionValueTypes:
    xlConditionValueAutomaticMax = 7  # from enum XlConditionValueTypes
    xlConditionValueAutomaticMin = 6  # from enum XlConditionValueTypes
    xlConditionValueFormula = 4  # from enum XlConditionValueTypes
    xlConditionValueHighestValue = 2  # from enum XlConditionValueTypes
    xlConditionValueLowestValue = 1  # from enum XlConditionValueTypes
    xlConditionValueNone = -1  # from enum XlConditionValueTypes
    xlConditionValueNumber = 0  # from enum XlConditionValueTypes
    xlConditionValuePercent = 3  # from enum XlConditionValueTypes
    xlConditionValuePercentile = 5  # from enum XlConditionValueTypes


class ConnectionType:
    xlConnectionTypeODBC = 2  # from enum XlConnectionType
    xlConnectionTypeOLEDB = 1  # from enum XlConnectionType
    xlConnectionTypeTEXT = 4  # from enum XlConnectionType
    xlConnectionTypeWEB = 5  # from enum XlConnectionType
    xlConnectionTypeXMLMAP = 3  # from enum XlConnectionType


class ConsolidationFunction:
    xlAverage = -4106  # from enum XlConsolidationFunction
    xlCount = -4112  # from enum XlConsolidationFunction
    xlCountNums = -4113  # from enum XlConsolidationFunction
    xlMax = -4136  # from enum XlConsolidationFunction
    xlMin = -4139  # from enum XlConsolidationFunction
    xlProduct = -4149  # from enum XlConsolidationFunction
    xlStDev = -4155  # from enum XlConsolidationFunction
    xlStDevP = -4156  # from enum XlConsolidationFunction
    xlSum = -4157  # from enum XlConsolidationFunction
    xlUnknown = 1000  # from enum XlConsolidationFunction
    xlVar = -4164  # from enum XlConsolidationFunction
    xlVarP = -4165  # from enum XlConsolidationFunction


class ContainsOperator:
    xlBeginsWith = 2  # from enum XlContainsOperator
    xlContains = 0  # from enum XlContainsOperator
    xlDoesNotContain = 1  # from enum XlContainsOperator
    xlEndsWith = 3  # from enum XlContainsOperator


class CopyPictureFormat:
    xlBitmap = 2  # from enum XlCopyPictureFormat
    xlPicture = -4147  # from enum XlCopyPictureFormat


class CorruptLoad:
    xlExtractData = 2  # from enum XlCorruptLoad
    xlNormalLoad = 0  # from enum XlCorruptLoad
    xlRepairFile = 1  # from enum XlCorruptLoad


class Creator:
    xlCreatorCode = 1480803660  # from enum XlCreator


class CredentialsMethod:
    xlCredentialsMethodIntegrated = 0  # from enum XlCredentialsMethod
    xlCredentialsMethodNone = 1  # from enum XlCredentialsMethod
    xlCredentialsMethodStored = 2  # from enum XlCredentialsMethod


class CubeFieldSubType:
    xlCubeAttribute = 4  # from enum XlCubeFieldSubType
    xlCubeCalculatedMeasure = 5  # from enum XlCubeFieldSubType
    xlCubeHierarchy = 1  # from enum XlCubeFieldSubType
    xlCubeKPIGoal = 7  # from enum XlCubeFieldSubType
    xlCubeKPIStatus = 8  # from enum XlCubeFieldSubType
    xlCubeKPITrend = 9  # from enum XlCubeFieldSubType
    xlCubeKPIValue = 6  # from enum XlCubeFieldSubType
    xlCubeKPIWeight = 10  # from enum XlCubeFieldSubType
    xlCubeMeasure = 2  # from enum XlCubeFieldSubType
    xlCubeSet = 3  # from enum XlCubeFieldSubType


class CubeFieldType:
    xlHierarchy = 1  # from enum XlCubeFieldType
    xlMeasure = 2  # from enum XlCubeFieldType
    xlSet = 3  # from enum XlCubeFieldType


class CutCopyMode:
    xlCopy = 1  # from enum XlCutCopyMode
    xlCut = 2  # from enum XlCutCopyMode


class DVAlertStyle:
    xlValidAlertInformation = 3  # from enum XlDVAlertStyle
    xlValidAlertStop = 1  # from enum XlDVAlertStyle
    xlValidAlertWarning = 2  # from enum XlDVAlertStyle


class DVType:
    xlValidateCustom = 7  # from enum XlDVType
    xlValidateDate = 4  # from enum XlDVType
    xlValidateDecimal = 2  # from enum XlDVType
    xlValidateInputOnly = 0  # from enum XlDVType
    xlValidateList = 3  # from enum XlDVType
    xlValidateTextLength = 6  # from enum XlDVType
    xlValidateTime = 5  # from enum XlDVType
    xlValidateWholeNumber = 1  # from enum XlDVType


class DataBarAxisPosition:
    xlDataBarAxisAutomatic = 0  # from enum XlDataBarAxisPosition
    xlDataBarAxisMidpoint = 1  # from enum XlDataBarAxisPosition
    xlDataBarAxisNone = 2  # from enum XlDataBarAxisPosition


class DataBarBorderType:
    xlDataBarBorderNone = 0  # from enum XlDataBarBorderType
    xlDataBarBorderSolid = 1  # from enum XlDataBarBorderType


class DataBarFillType:
    xlDataBarFillGradient = 1  # from enum XlDataBarFillType
    xlDataBarFillSolid = 0  # from enum XlDataBarFillType


class DataBarNegativeColorType:
    xlDataBarColor = 0  # from enum XlDataBarNegativeColorType
    xlDataBarSameAsPositive = 1  # from enum XlDataBarNegativeColorType


class DataLabelPosition:
    xlLabelPositionAbove = 0  # from enum XlDataLabelPosition
    xlLabelPositionBelow = 1  # from enum XlDataLabelPosition
    xlLabelPositionBestFit = 5  # from enum XlDataLabelPosition
    xlLabelPositionCenter = -4108  # from enum XlDataLabelPosition
    xlLabelPositionCustom = 7  # from enum XlDataLabelPosition
    xlLabelPositionInsideBase = 4  # from enum XlDataLabelPosition
    xlLabelPositionInsideEnd = 3  # from enum XlDataLabelPosition
    xlLabelPositionLeft = -4131  # from enum XlDataLabelPosition
    xlLabelPositionMixed = 6  # from enum XlDataLabelPosition
    xlLabelPositionOutsideEnd = 2  # from enum XlDataLabelPosition
    xlLabelPositionRight = -4152  # from enum XlDataLabelPosition


class DataLabelSeparator:
    xlDataLabelSeparatorDefault = 1  # from enum XlDataLabelSeparator


class DataLabelsType:
    xlDataLabelsShowBubbleSizes = 6  # from enum XlDataLabelsType
    xlDataLabelsShowLabel = 4  # from enum XlDataLabelsType
    xlDataLabelsShowLabelAndPercent = 5  # from enum XlDataLabelsType
    xlDataLabelsShowNone = -4142  # from enum XlDataLabelsType
    xlDataLabelsShowPercent = 3  # from enum XlDataLabelsType
    xlDataLabelsShowValue = 2  # from enum XlDataLabelsType


class DataSeriesDate:
    xlDay = 1  # from enum XlDataSeriesDate
    xlMonth = 3  # from enum XlDataSeriesDate
    xlWeekday = 2  # from enum XlDataSeriesDate
    xlYear = 4  # from enum XlDataSeriesDate
    xlAutoFill = 4  # from enum XlDataSeriesType
    xlChronological = 3  # from enum XlDataSeriesType
    xlDataSeriesLinear = -4132  # from enum XlDataSeriesType
    xlGrowth = 2  # from enum XlDataSeriesType


class DeleteShiftDirection:
    xlShiftToLeft = -4159  # from enum XlDeleteShiftDirection
    xlShiftUp = -4162  # from enum XlDeleteShiftDirection


class Direction:
    xlDown = -4121  # from enum XlDirection
    xlToLeft = -4159  # from enum XlDirection
    xlToRight = -4161  # from enum XlDirection
    xlUp = -4162  # from enum XlDirection


directions = (
    'down',
    'left',
    'right'
    'up'
)


class DisplayBlanksAs:
    xlInterpolated = 3  # from enum XlDisplayBlanksAs
    xlNotPlotted = 1  # from enum XlDisplayBlanksAs
    xlZero = 2  # from enum XlDisplayBlanksAs


class DisplayDrawingObjects:
    xlDisplayShapes = -4104  # from enum XlDisplayDrawingObjects
    xlHide = 3  # from enum XlDisplayDrawingObjects
    xlPlaceholders = 2  # from enum XlDisplayDrawingObjects


class DisplayUnit:
    xlHundredMillions = -8  # from enum XlDisplayUnit
    xlHundredThousands = -5  # from enum XlDisplayUnit
    xlHundreds = -2  # from enum XlDisplayUnit
    xlMillionMillions = -10  # from enum XlDisplayUnit
    xlMillions = -6  # from enum XlDisplayUnit
    xlTenMillions = -7  # from enum XlDisplayUnit
    xlTenThousands = -4  # from enum XlDisplayUnit
    xlThousandMillions = -9  # from enum XlDisplayUnit
    xlThousands = -3  # from enum XlDisplayUnit


class DupeUnique:
    xlDuplicate = 1  # from enum XlDupeUnique
    xlUnique = 0  # from enum XlDupeUnique


class DynamicFilterCriteria:
    xlFilterAboveAverage = 33  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodApril = 24  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodAugust = 28  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodDecember = 32  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodFebruray = 22  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodJanuary = 21  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodJuly = 27  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodJune = 26  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodMarch = 23  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodMay = 25  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodNovember = 31  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodOctober = 30  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodQuarter1 = 17  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodQuarter2 = 18  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodQuarter3 = 19  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodQuarter4 = 20  # from enum XlDynamicFilterCriteria
    xlFilterAllDatesInPeriodSeptember = 29  # from enum XlDynamicFilterCriteria
    xlFilterBelowAverage = 34  # from enum XlDynamicFilterCriteria
    xlFilterLastMonth = 8  # from enum XlDynamicFilterCriteria
    xlFilterLastQuarter = 11  # from enum XlDynamicFilterCriteria
    xlFilterLastWeek = 5  # from enum XlDynamicFilterCriteria
    xlFilterLastYear = 14  # from enum XlDynamicFilterCriteria
    xlFilterNextMonth = 9  # from enum XlDynamicFilterCriteria
    xlFilterNextQuarter = 12  # from enum XlDynamicFilterCriteria
    xlFilterNextWeek = 6  # from enum XlDynamicFilterCriteria
    xlFilterNextYear = 15  # from enum XlDynamicFilterCriteria
    xlFilterThisMonth = 7  # from enum XlDynamicFilterCriteria
    xlFilterThisQuarter = 10  # from enum XlDynamicFilterCriteria
    xlFilterThisWeek = 4  # from enum XlDynamicFilterCriteria
    xlFilterThisYear = 13  # from enum XlDynamicFilterCriteria
    xlFilterToday = 1  # from enum XlDynamicFilterCriteria
    xlFilterTomorrow = 3  # from enum XlDynamicFilterCriteria
    xlFilterYearToDate = 16  # from enum XlDynamicFilterCriteria
    xlFilterYesterday = 2  # from enum XlDynamicFilterCriteria


class EditionFormat:
    xlBIFF = 2  # from enum XlEditionFormat
    xlPICT = 1  # from enum XlEditionFormat
    xlRTF = 4  # from enum XlEditionFormat
    xlVALU = 8  # from enum XlEditionFormat


class EditionOptionsOption:
    xlAutomaticUpdate = 4  # from enum XlEditionOptionsOption
    xlCancel = 1  # from enum XlEditionOptionsOption
    xlChangeAttributes = 6  # from enum XlEditionOptionsOption
    xlManualUpdate = 5  # from enum XlEditionOptionsOption
    xlOpenSource = 3  # from enum XlEditionOptionsOption
    xlSelect = 3  # from enum XlEditionOptionsOption
    xlSendPublisher = 2  # from enum XlEditionOptionsOption
    xlUpdateSubscriber = 2  # from enum XlEditionOptionsOption


class EditionType:
    xlPublisher = 1  # from enum XlEditionType
    xlSubscriber = 2  # from enum XlEditionType


class EnableCancelKey:
    xlDisabled = 0  # from enum XlEnableCancelKey
    xlErrorHandler = 2  # from enum XlEnableCancelKey
    xlInterrupt = 1  # from enum XlEnableCancelKey


class EnableSelection:
    xlNoRestrictions = 0  # from enum XlEnableSelection
    xlNoSelection = -4142  # from enum XlEnableSelection
    xlUnlockedCells = 1  # from enum XlEnableSelection


class EndStyleCap:
    xlCap = 1  # from enum XlEndStyleCap
    xlNoCap = 2  # from enum XlEndStyleCap


class ErrorBarDirection:
    xlX = -4168  # from enum XlErrorBarDirection
    xlY = 1  # from enum XlErrorBarDirection


class ErrorBarInclude:
    xlErrorBarIncludeBoth = 1  # from enum XlErrorBarInclude
    xlErrorBarIncludeMinusValues = 3  # from enum XlErrorBarInclude
    xlErrorBarIncludeNone = -4142  # from enum XlErrorBarInclude
    xlErrorBarIncludePlusValues = 2  # from enum XlErrorBarInclude


class ErrorBarType:
    xlErrorBarTypeCustom = -4114  # from enum XlErrorBarType
    xlErrorBarTypeFixedValue = 1  # from enum XlErrorBarType
    xlErrorBarTypePercent = 2  # from enum XlErrorBarType
    xlErrorBarTypeStDev = -4155  # from enum XlErrorBarType
    xlErrorBarTypeStError = 4  # from enum XlErrorBarType


class ErrorChecks:
    xlEmptyCellReferences = 7  # from enum XlErrorChecks
    xlEvaluateToError = 1  # from enum XlErrorChecks
    xlInconsistentFormula = 4  # from enum XlErrorChecks
    xlInconsistentListFormula = 9  # from enum XlErrorChecks
    xlListDataValidation = 8  # from enum XlErrorChecks
    xlNumberAsText = 3  # from enum XlErrorChecks
    xlOmittedCells = 5  # from enum XlErrorChecks
    xlTextDate = 2  # from enum XlErrorChecks
    xlUnlockedFormulaCells = 6  # from enum XlErrorChecks


class FileAccess:
    xlReadOnly = 3  # from enum XlFileAccess
    xlReadWrite = 2  # from enum XlFileAccess


class FileFormat:
    xlAddIn = 18  # from enum XlFileFormat
    xlAddIn8 = 18  # from enum XlFileFormat
    xlCSV = 6  # from enum XlFileFormat
    xlCSVMSDOS = 24  # from enum XlFileFormat
    xlCSVMac = 22  # from enum XlFileFormat
    xlCSVWindows = 23  # from enum XlFileFormat
    xlCurrentPlatformText = -4158  # from enum XlFileFormat
    xlDBF2 = 7  # from enum XlFileFormat
    xlDBF3 = 8  # from enum XlFileFormat
    xlDBF4 = 11  # from enum XlFileFormat
    xlDIF = 9  # from enum XlFileFormat
    xlExcel12 = 50  # from enum XlFileFormat
    xlExcel2 = 16  # from enum XlFileFormat
    xlExcel2FarEast = 27  # from enum XlFileFormat
    xlExcel3 = 29  # from enum XlFileFormat
    xlExcel4 = 33  # from enum XlFileFormat
    xlExcel4Workbook = 35  # from enum XlFileFormat
    xlExcel5 = 39  # from enum XlFileFormat
    xlExcel7 = 39  # from enum XlFileFormat
    xlExcel8 = 56  # from enum XlFileFormat
    xlExcel9795 = 43  # from enum XlFileFormat
    xlHtml = 44  # from enum XlFileFormat
    xlIntlAddIn = 26  # from enum XlFileFormat
    xlIntlMacro = 25  # from enum XlFileFormat
    xlOpenDocumentSpreadsheet = 60  # from enum XlFileFormat
    xlOpenXMLAddIn = 55  # from enum XlFileFormat
    xlOpenXMLTemplate = 54  # from enum XlFileFormat
    xlOpenXMLTemplateMacroEnabled = 53  # from enum XlFileFormat
    xlOpenXMLWorkbook = 51  # from enum XlFileFormat
    xlOpenXMLWorkbookMacroEnabled = 52  # from enum XlFileFormat
    xlSYLK = 2  # from enum XlFileFormat
    xlTemplate = 17  # from enum XlFileFormat
    xlTemplate8 = 17  # from enum XlFileFormat
    xlTextMSDOS = 21  # from enum XlFileFormat
    xlTextMac = 19  # from enum XlFileFormat
    xlTextPrinter = 36  # from enum XlFileFormat
    xlTextWindows = 20  # from enum XlFileFormat
    xlUnicodeText = 42  # from enum XlFileFormat
    xlWJ2WD1 = 14  # from enum XlFileFormat
    xlWJ3 = 40  # from enum XlFileFormat
    xlWJ3FJ3 = 41  # from enum XlFileFormat
    xlWK1 = 5  # from enum XlFileFormat
    xlWK1ALL = 31  # from enum XlFileFormat
    xlWK1FMT = 30  # from enum XlFileFormat
    xlWK3 = 15  # from enum XlFileFormat
    xlWK3FM3 = 32  # from enum XlFileFormat
    xlWK4 = 38  # from enum XlFileFormat
    xlWKS = 4  # from enum XlFileFormat
    xlWQ1 = 34  # from enum XlFileFormat
    xlWebArchive = 45  # from enum XlFileFormat
    xlWorkbookDefault = 51  # from enum XlFileFormat
    xlWorkbookNormal = -4143  # from enum XlFileFormat
    xlWorks2FarEast = 28  # from enum XlFileFormat
    xlXMLSpreadsheet = 46  # from enum XlFileFormat


class FileValidationPivotMode:
    xlFileValidationPivotDefault = 0  # from enum XlFileValidationPivotMode
    xlFileValidationPivotRun = 1  # from enum XlFileValidationPivotMode
    xlFileValidationPivotSkip = 2  # from enum XlFileValidationPivotMode


class FillWith:
    xlFillWithAll = -4104  # from enum XlFillWith
    xlFillWithContents = 2  # from enum XlFillWith
    xlFillWithFormats = -4122  # from enum XlFillWith


class FilterAction:
    xlFilterCopy = 2  # from enum XlFilterAction
    xlFilterInPlace = 1  # from enum XlFilterAction


class FilterAllDatesInPeriod:
    xlFilterAllDatesInPeriodDay = 2  # from enum XlFilterAllDatesInPeriod
    xlFilterAllDatesInPeriodHour = 3  # from enum XlFilterAllDatesInPeriod
    xlFilterAllDatesInPeriodMinute = 4  # from enum XlFilterAllDatesInPeriod
    xlFilterAllDatesInPeriodMonth = 1  # from enum XlFilterAllDatesInPeriod
    xlFilterAllDatesInPeriodSecond = 5  # from enum XlFilterAllDatesInPeriod
    xlFilterAllDatesInPeriodYear = 0  # from enum XlFilterAllDatesInPeriod


class FindLookIn:
    xlComments = -4144  # from enum XlFindLookIn
    xlFormulas = -4123  # from enum XlFindLookIn
    xlValues = -4163  # from enum XlFindLookIn


class FixedFormatQuality:
    xlQualityMinimum = 1  # from enum XlFixedFormatQuality
    xlQualityStandard = 0  # from enum XlFixedFormatQuality


class FixedFormatType:
    xlTypePDF = 0  # from enum XlFixedFormatType
    xlTypeXPS = 1  # from enum XlFixedFormatType


class FormControl:
    xlButtonControl = 0  # from enum XlFormControl
    xlCheckBox = 1  # from enum XlFormControl
    xlDropDown = 2  # from enum XlFormControl
    xlEditBox = 3  # from enum XlFormControl
    xlGroupBox = 4  # from enum XlFormControl
    xlLabel = 5  # from enum XlFormControl
    xlListBox = 6  # from enum XlFormControl
    xlOptionButton = 7  # from enum XlFormControl
    xlScrollBar = 8  # from enum XlFormControl
    xlSpinner = 9  # from enum XlFormControl


class FormatConditionOperator:
    xlBetween = 1  # from enum XlFormatConditionOperator
    xlEqual = 3  # from enum XlFormatConditionOperator
    xlGreater = 5  # from enum XlFormatConditionOperator
    xlGreaterEqual = 7  # from enum XlFormatConditionOperator
    xlLess = 6  # from enum XlFormatConditionOperator
    xlLessEqual = 8  # from enum XlFormatConditionOperator
    xlNotBetween = 2  # from enum XlFormatConditionOperator
    xlNotEqual = 4  # from enum XlFormatConditionOperator


class FormatConditionType:
    xlAboveAverageCondition = 12  # from enum XlFormatConditionType
    xlBlanksCondition = 10  # from enum XlFormatConditionType
    xlCellValue = 1  # from enum XlFormatConditionType
    xlColorScale = 3  # from enum XlFormatConditionType
    xlDatabar = 4  # from enum XlFormatConditionType
    xlErrorsCondition = 16  # from enum XlFormatConditionType
    xlExpression = 2  # from enum XlFormatConditionType
    xlIconSets = 6  # from enum XlFormatConditionType
    xlNoBlanksCondition = 13  # from enum XlFormatConditionType
    xlNoErrorsCondition = 17  # from enum XlFormatConditionType
    xlTextString = 9  # from enum XlFormatConditionType
    xlTimePeriod = 11  # from enum XlFormatConditionType
    xlTop10 = 5  # from enum XlFormatConditionType
    xlUniqueValues = 8  # from enum XlFormatConditionType

class FormatFilterTypes:
    xlFilterBottom = 0  # from enum XlFormatFilterTypes
    xlFilterBottomPercent = 2  # from enum XlFormatFilterTypes
    xlFilterTop = 1  # from enum XlFormatFilterTypes
    xlFilterTopPercent = 3  # from enum XlFormatFilterTypes


class FormulaLabel:
    xlColumnLabels = 2  # from enum XlFormulaLabel
    xlMixedLabels = 3  # from enum XlFormulaLabel
    xlNoLabels = -4142  # from enum XlFormulaLabel
    xlRowLabels = 1  # from enum XlFormulaLabel


class GenerateTableRefs:
    xlGenerateTableRefA1 = 0  # from enum XlGenerateTableRefs
    xlGenerateTableRefStruct = 1  # from enum XlGenerateTableRefs


class GradientFillType:
    xlGradientFillLinear = 0  # from enum XlGradientFillType
    xlGradientFillPath = 1  # from enum XlGradientFillType


class HAlign:
    xlHAlignCenter = -4108  # from enum XlHAlign
    xlHAlignCenterAcrossSelection = 7  # from enum XlHAlign
    xlHAlignDistributed = -4117  # from enum XlHAlign
    xlHAlignFill = 5  # from enum XlHAlign
    xlHAlignGeneral = 1  # from enum XlHAlign
    xlHAlignJustify = -4130  # from enum XlHAlign
    xlHAlignLeft = -4131  # from enum XlHAlign
    xlHAlignRight = -4152  # from enum XlHAlign


class HebrewModes:
    xlHebrewFullScript = 0  # from enum XlHebrewModes
    xlHebrewMixedAuthorizedScript = 3  # from enum XlHebrewModes
    xlHebrewMixedScript = 2  # from enum XlHebrewModes
    xlHebrewPartialScript = 1  # from enum XlHebrewModes


class HighlightChangesTime:
    xlAllChanges = 2  # from enum XlHighlightChangesTime
    xlNotYetReviewed = 3  # from enum XlHighlightChangesTime
    xlSinceMyLastSave = 1  # from enum XlHighlightChangesTime


class HtmlType:
    xlHtmlCalc = 1  # from enum XlHtmlType
    xlHtmlChart = 3  # from enum XlHtmlType
    xlHtmlList = 2  # from enum XlHtmlType
    xlHtmlStatic = 0  # from enum XlHtmlType


class IMEMode:
    xlIMEModeAlpha = 8  # from enum XlIMEMode
    xlIMEModeAlphaFull = 7  # from enum XlIMEMode
    xlIMEModeDisable = 3  # from enum XlIMEMode
    xlIMEModeHangul = 10  # from enum XlIMEMode
    xlIMEModeHangulFull = 9  # from enum XlIMEMode
    xlIMEModeHiragana = 4  # from enum XlIMEMode
    xlIMEModeKatakana = 5  # from enum XlIMEMode
    xlIMEModeKatakanaHalf = 6  # from enum XlIMEMode
    xlIMEModeNoControl = 0  # from enum XlIMEMode
    xlIMEModeOff = 2  # from enum XlIMEMode
    xlIMEModeOn = 1  # from enum XlIMEMode


class Icon:
    xlIcon0Bars = 37  # from enum XlIcon
    xlIcon0FilledBoxes = 52  # from enum XlIcon
    xlIcon1Bar = 38  # from enum XlIcon
    xlIcon1FilledBox = 51  # from enum XlIcon
    xlIcon2Bars = 39  # from enum XlIcon
    xlIcon2FilledBoxes = 50  # from enum XlIcon
    xlIcon3Bars = 40  # from enum XlIcon
    xlIcon3FilledBoxes = 49  # from enum XlIcon
    xlIcon4Bars = 41  # from enum XlIcon
    xlIcon4FilledBoxes = 48  # from enum XlIcon
    xlIconBlackCircle = 32  # from enum XlIcon
    xlIconBlackCircleWithBorder = 13  # from enum XlIcon
    xlIconCircleWithOneWhiteQuarter = 33  # from enum XlIcon
    xlIconCircleWithThreeWhiteQuarters = 35  # from enum XlIcon
    xlIconCircleWithTwoWhiteQuarters = 34  # from enum XlIcon
    xlIconGoldStar = 42  # from enum XlIcon
    xlIconGrayCircle = 31  # from enum XlIcon
    xlIconGrayDownArrow = 6  # from enum XlIcon
    xlIconGrayDownInclineArrow = 28  # from enum XlIcon
    xlIconGraySideArrow = 5  # from enum XlIcon
    xlIconGrayUpArrow = 4  # from enum XlIcon
    xlIconGrayUpInclineArrow = 27  # from enum XlIcon
    xlIconGreenCheck = 22  # from enum XlIcon
    xlIconGreenCheckSymbol = 19  # from enum XlIcon
    xlIconGreenCircle = 10  # from enum XlIcon
    xlIconGreenFlag = 7  # from enum XlIcon
    xlIconGreenTrafficLight = 14  # from enum XlIcon
    xlIconGreenUpArrow = 1  # from enum XlIcon
    xlIconGreenUpTriangle = 45  # from enum XlIcon
    xlIconHalfGoldStar = 43  # from enum XlIcon
    xlIconNoCellIcon = -1  # from enum XlIcon
    xlIconPinkCircle = 30  # from enum XlIcon
    xlIconRedCircle = 29  # from enum XlIcon
    xlIconRedCircleWithBorder = 12  # from enum XlIcon
    xlIconRedCross = 24  # from enum XlIcon
    xlIconRedCrossSymbol = 21  # from enum XlIcon
    xlIconRedDiamond = 18  # from enum XlIcon
    xlIconRedDownArrow = 3  # from enum XlIcon
    xlIconRedDownTriangle = 47  # from enum XlIcon
    xlIconRedFlag = 9  # from enum XlIcon
    xlIconRedTrafficLight = 16  # from enum XlIcon
    xlIconSilverStar = 44  # from enum XlIcon
    xlIconWhiteCircleAllWhiteQuarters = 36  # from enum XlIcon
    xlIconYellowCircle = 11  # from enum XlIcon
    xlIconYellowDash = 46  # from enum XlIcon
    xlIconYellowDownInclineArrow = 26  # from enum XlIcon
    xlIconYellowExclamation = 23  # from enum XlIcon
    xlIconYellowExclamationSymbol = 20  # from enum XlIcon
    xlIconYellowFlag = 8  # from enum XlIcon
    xlIconYellowSideArrow = 2  # from enum XlIcon
    xlIconYellowTrafficLight = 15  # from enum XlIcon
    xlIconYellowTriangle = 17  # from enum XlIcon
    xlIconYellowUpInclineArrow = 25  # from enum XlIcon


class IconSet:
    xl3Arrows = 1  # from enum XlIconSet
    xl3ArrowsGray = 2  # from enum XlIconSet
    xl3Flags = 3  # from enum XlIconSet
    xl3Signs = 6  # from enum XlIconSet
    xl3Stars = 18  # from enum XlIconSet
    xl3Symbols = 7  # from enum XlIconSet
    xl3Symbols2 = 8  # from enum XlIconSet
    xl3TrafficLights1 = 4  # from enum XlIconSet
    xl3TrafficLights2 = 5  # from enum XlIconSet
    xl3Triangles = 19  # from enum XlIconSet
    xl4Arrows = 9  # from enum XlIconSet
    xl4ArrowsGray = 10  # from enum XlIconSet
    xl4CRV = 12  # from enum XlIconSet
    xl4RedToBlack = 11  # from enum XlIconSet
    xl4TrafficLights = 13  # from enum XlIconSet
    xl5Arrows = 14  # from enum XlIconSet
    xl5ArrowsGray = 15  # from enum XlIconSet
    xl5Boxes = 20  # from enum XlIconSet
    xl5CRV = 16  # from enum XlIconSet
    xl5Quarters = 17  # from enum XlIconSet
    xlCustomSet = -1  # from enum XlIconSet


class ImportDataAs:
    xlPivotTableReport = 1  # from enum XlImportDataAs
    xlQueryTable = 0  # from enum XlImportDataAs
    xlTable = 2  # from enum XlImportDataAs


class InsertFormatOrigin:
    xlFormatFromLeftOrAbove = 0  # from enum XlInsertFormatOrigin
    xlFormatFromRightOrBelow = 1  # from enum XlInsertFormatOrigin


class InsertShiftDirection:
    xlShiftDown = -4121  # from enum XlInsertShiftDirection
    xlShiftToRight = -4161  # from enum XlInsertShiftDirection


class LayoutFormType:
    xlOutline = 1  # from enum XlLayoutFormType
    xlTabular = 0  # from enum XlLayoutFormType


class LayoutRowType:
    xlCompactRow = 0  # from enum XlLayoutRowType
    xlOutlineRow = 2  # from enum XlLayoutRowType
    xlTabularRow = 1  # from enum XlLayoutRowType


class LegendPosition:
    xlLegendPositionBottom = -4107  # from enum XlLegendPosition
    xlLegendPositionCorner = 2  # from enum XlLegendPosition
    xlLegendPositionCustom = -4161  # from enum XlLegendPosition
    xlLegendPositionLeft = -4131  # from enum XlLegendPosition
    xlLegendPositionRight = -4152  # from enum XlLegendPosition
    xlLegendPositionTop = -4160  # from enum XlLegendPosition


class LineStyle:
    xlContinuous = 1  # from enum XlLineStyle
    xlDash = -4115  # from enum XlLineStyle
    xlDashDot = 4  # from enum XlLineStyle
    xlDashDotDot = 5  # from enum XlLineStyle
    xlDot = -4118  # from enum XlLineStyle
    xlDouble = -4119  # from enum XlLineStyle
    xlLineStyleNone = -4142  # from enum XlLineStyle
    xlSlantDashDot = 13  # from enum XlLineStyle


class Link:
    xlExcelLinks = 1  # from enum XlLink
    xlOLELinks = 2  # from enum XlLink
    xlPublishers = 5  # from enum XlLink
    xlSubscribers = 6  # from enum XlLink


class LinkInfo:
    xlEditionDate = 2  # from enum XlLinkInfo
    xlLinkInfoStatus = 3  # from enum XlLinkInfo
    xlUpdateState = 1  # from enum XlLinkInfo


class LinkInfoType:
    xlLinkInfoOLELinks = 2  # from enum XlLinkInfoType
    xlLinkInfoPublishers = 5  # from enum XlLinkInfoType
    xlLinkInfoSubscribers = 6  # from enum XlLinkInfoType


class LinkStatus:
    xlLinkStatusCopiedValues = 10  # from enum XlLinkStatus
    xlLinkStatusIndeterminate = 5  # from enum XlLinkStatus
    xlLinkStatusInvalidName = 7  # from enum XlLinkStatus
    xlLinkStatusMissingFile = 1  # from enum XlLinkStatus
    xlLinkStatusMissingSheet = 2  # from enum XlLinkStatus
    xlLinkStatusNotStarted = 6  # from enum XlLinkStatus
    xlLinkStatusOK = 0  # from enum XlLinkStatus
    xlLinkStatusOld = 3  # from enum XlLinkStatus
    xlLinkStatusSourceNotCalculated = 4  # from enum XlLinkStatus
    xlLinkStatusSourceNotOpen = 8  # from enum XlLinkStatus
    xlLinkStatusSourceOpen = 9  # from enum XlLinkStatus


class LinkType:
    xlLinkTypeExcelLinks = 1  # from enum XlLinkType
    xlLinkTypeOLELinks = 2  # from enum XlLinkType


class ListConflict:
    xlListConflictDialog = 0  # from enum XlListConflict
    xlListConflictDiscardAllConflicts = 2  # from enum XlListConflict
    xlListConflictError = 3  # from enum XlListConflict
    xlListConflictRetryAllConflicts = 1  # from enum XlListConflict


class ListDataType:
    xlListDataTypeCheckbox = 9  # from enum XlListDataType
    xlListDataTypeChoice = 6  # from enum XlListDataType
    xlListDataTypeChoiceMulti = 7  # from enum XlListDataType
    xlListDataTypeCounter = 11  # from enum XlListDataType
    xlListDataTypeCurrency = 4  # from enum XlListDataType
    xlListDataTypeDateTime = 5  # from enum XlListDataType
    xlListDataTypeHyperLink = 10  # from enum XlListDataType
    xlListDataTypeListLookup = 8  # from enum XlListDataType
    xlListDataTypeMultiLineRichText = 12  # from enum XlListDataType
    xlListDataTypeMultiLineText = 2  # from enum XlListDataType
    xlListDataTypeNone = 0  # from enum XlListDataType
    xlListDataTypeNumber = 3  # from enum XlListDataType
    xlListDataTypeText = 1  # from enum XlListDataType


class ListObjectSourceType:
    xlSrcExternal = 0  # from enum XlListObjectSourceType
    xlSrcQuery = 3  # from enum XlListObjectSourceType
    xlSrcRange = 1  # from enum XlListObjectSourceType
    xlSrcXml = 2  # from enum XlListObjectSourceType


class LocationInTable:
    xlColumnHeader = -4110  # from enum XlLocationInTable
    xlColumnItem = 5  # from enum XlLocationInTable
    xlDataHeader = 3  # from enum XlLocationInTable
    xlDataItem = 7  # from enum XlLocationInTable
    xlPageHeader = 2  # from enum XlLocationInTable
    xlPageItem = 6  # from enum XlLocationInTable
    xlRowHeader = -4153  # from enum XlLocationInTable
    xlRowItem = 4  # from enum XlLocationInTable
    xlTableBody = 8  # from enum XlLocationInTable


class LookAt:
    xlPart = 2  # from enum XlLookAt
    xlWhole = 1  # from enum XlLookAt


class LookFor:
    xlLookForBlanks = 0  # from enum XlLookFor
    xlLookForErrors = 1  # from enum XlLookFor
    xlLookForFormulas = 2  # from enum XlLookFor


class MSApplication:
    xlMicrosoftAccess = 4  # from enum XlMSApplication
    xlMicrosoftFoxPro = 5  # from enum XlMSApplication
    xlMicrosoftMail = 3  # from enum XlMSApplication
    xlMicrosoftPowerPoint = 2  # from enum XlMSApplication
    xlMicrosoftProject = 6  # from enum XlMSApplication
    xlMicrosoftSchedulePlus = 7  # from enum XlMSApplication
    xlMicrosoftWord = 1  # from enum XlMSApplication


class MailSystem:
    xlMAPI = 1  # from enum XlMailSystem
    xlNoMailSystem = 0  # from enum XlMailSystem
    xlPowerTalk = 2  # from enum XlMailSystem


class MarkerStyle:
    xlMarkerStyleAutomatic = -4105  # from enum XlMarkerStyle
    xlMarkerStyleCircle = 8  # from enum XlMarkerStyle
    xlMarkerStyleDash = -4115  # from enum XlMarkerStyle
    xlMarkerStyleDiamond = 2  # from enum XlMarkerStyle
    xlMarkerStyleDot = -4118  # from enum XlMarkerStyle
    xlMarkerStyleNone = -4142  # from enum XlMarkerStyle
    xlMarkerStylePicture = -4147  # from enum XlMarkerStyle
    xlMarkerStylePlus = 9  # from enum XlMarkerStyle
    xlMarkerStyleSquare = 1  # from enum XlMarkerStyle
    xlMarkerStyleStar = 5  # from enum XlMarkerStyle
    xlMarkerStyleTriangle = 3  # from enum XlMarkerStyle
    xlMarkerStyleX = -4168  # from enum XlMarkerStyle

class MeasurementUnits:
    xlCentimeters = 1  # from enum XlMeasurementUnits
    xlInches = 0  # from enum XlMeasurementUnits
    xlMillimeters = 2  # from enum XlMeasurementUnits


class MouseButton:
    xlNoButton = 0  # from enum XlMouseButton
    xlPrimaryButton = 1  # from enum XlMouseButton
    xlSecondaryButton = 2  # from enum XlMouseButton


class MousePointer:
    xlDefault = -4143  # from enum XlMousePointer
    xlIBeam = 3  # from enum XlMousePointer
    xlNorthwestArrow = 1  # from enum XlMousePointer
    xlWait = 2  # from enum XlMousePointer


class OLEType:
    xlOLEControl = 2  # from enum XlOLEType
    xlOLEEmbed = 1  # from enum XlOLEType
    xlOLELink = 0  # from enum XlOLEType


class OLEVerb:
    xlVerbOpen = 2  # from enum XlOLEVerb
    xlVerbPrimary = 1  # from enum XlOLEVerb


class OartHorizontalOverflow:
    xlOartHorizontalOverflowClip = 1  # from enum XlOartHorizontalOverflow
    xlOartHorizontalOverflowOverflow = 0  # from enum XlOartHorizontalOverflow


class OartVerticalOverflow:
    xlOartVerticalOverflowClip = 1  # from enum XlOartVerticalOverflow
    xlOartVerticalOverflowEllipsis = 2  # from enum XlOartVerticalOverflow
    xlOartVerticalOverflowOverflow = 0  # from enum XlOartVerticalOverflow


class ObjectSize:
    xlFitToPage = 2  # from enum XlObjectSize
    xlFullPage = 3  # from enum XlObjectSize
    xlScreenSize = 1  # from enum XlObjectSize


class Order:
    xlDownThenOver = 1  # from enum XlOrder
    xlOverThenDown = 2  # from enum XlOrder


class Orientation:
    xlDownward = -4170  # from enum XlOrientation
    xlHorizontal = -4128  # from enum XlOrientation
    xlUpward = -4171  # from enum XlOrientation
    xlVertical = -4166  # from enum XlOrientation


class PTSelectionMode:
    xlBlanks = 4  # from enum XlPTSelectionMode
    xlButton = 15  # from enum XlPTSelectionMode
    xlDataAndLabel = 0  # from enum XlPTSelectionMode
    xlDataOnly = 2  # from enum XlPTSelectionMode
    xlFirstRow = 256  # from enum XlPTSelectionMode
    xlLabelOnly = 1  # from enum XlPTSelectionMode
    xlOrigin = 3  # from enum XlPTSelectionMode


class PageBreak:
    xlPageBreakAutomatic = -4105  # from enum XlPageBreak
    xlPageBreakManual = -4135  # from enum XlPageBreak
    xlPageBreakNone = -4142  # from enum XlPageBreak


class PageBreakExtent:
    xlPageBreakFull = 1  # from enum XlPageBreakExtent
    xlPageBreakPartial = 2  # from enum XlPageBreakExtent


class PageOrientation:
    xlLandscape = 2  # from enum XlPageOrientation
    xlPortrait = 1  # from enum XlPageOrientation


class PaperSize:
    xlPaper10x14 = 16  # from enum XlPaperSize
    xlPaper11x17 = 17  # from enum XlPaperSize
    xlPaperA3 = 8  # from enum XlPaperSize
    xlPaperA4 = 9  # from enum XlPaperSize
    xlPaperA4Small = 10  # from enum XlPaperSize
    xlPaperA5 = 11  # from enum XlPaperSize
    xlPaperB4 = 12  # from enum XlPaperSize
    xlPaperB5 = 13  # from enum XlPaperSize
    xlPaperCsheet = 24  # from enum XlPaperSize
    xlPaperDsheet = 25  # from enum XlPaperSize
    xlPaperEnvelope10 = 20  # from enum XlPaperSize
    xlPaperEnvelope11 = 21  # from enum XlPaperSize
    xlPaperEnvelope12 = 22  # from enum XlPaperSize
    xlPaperEnvelope14 = 23  # from enum XlPaperSize
    xlPaperEnvelope9 = 19  # from enum XlPaperSize
    xlPaperEnvelopeB4 = 33  # from enum XlPaperSize
    xlPaperEnvelopeB5 = 34  # from enum XlPaperSize
    xlPaperEnvelopeB6 = 35  # from enum XlPaperSize
    xlPaperEnvelopeC3 = 29  # from enum XlPaperSize
    xlPaperEnvelopeC4 = 30  # from enum XlPaperSize
    xlPaperEnvelopeC5 = 28  # from enum XlPaperSize
    xlPaperEnvelopeC6 = 31  # from enum XlPaperSize
    xlPaperEnvelopeC65 = 32  # from enum XlPaperSize
    xlPaperEnvelopeDL = 27  # from enum XlPaperSize
    xlPaperEnvelopeItaly = 36  # from enum XlPaperSize
    xlPaperEnvelopeMonarch = 37  # from enum XlPaperSize
    xlPaperEnvelopePersonal = 38  # from enum XlPaperSize
    xlPaperEsheet = 26  # from enum XlPaperSize
    xlPaperExecutive = 7  # from enum XlPaperSize
    xlPaperFanfoldLegalGerman = 41  # from enum XlPaperSize
    xlPaperFanfoldStdGerman = 40  # from enum XlPaperSize
    xlPaperFanfoldUS = 39  # from enum XlPaperSize
    xlPaperFolio = 14  # from enum XlPaperSize
    xlPaperLedger = 4  # from enum XlPaperSize
    xlPaperLegal = 5  # from enum XlPaperSize
    xlPaperLetter = 1  # from enum XlPaperSize
    xlPaperLetterSmall = 2  # from enum XlPaperSize
    xlPaperNote = 18  # from enum XlPaperSize
    xlPaperQuarto = 15  # from enum XlPaperSize
    xlPaperStatement = 6  # from enum XlPaperSize
    xlPaperTabloid = 3  # from enum XlPaperSize
    xlPaperUser = 256  # from enum XlPaperSize


class ParameterDataType:
    xlParamTypeBigInt = -5  # from enum XlParameterDataType
    xlParamTypeBinary = -2  # from enum XlParameterDataType
    xlParamTypeBit = -7  # from enum XlParameterDataType
    xlParamTypeChar = 1  # from enum XlParameterDataType
    xlParamTypeDate = 9  # from enum XlParameterDataType
    xlParamTypeDecimal = 3  # from enum XlParameterDataType
    xlParamTypeDouble = 8  # from enum XlParameterDataType
    xlParamTypeFloat = 6  # from enum XlParameterDataType
    xlParamTypeInteger = 4  # from enum XlParameterDataType
    xlParamTypeLongVarBinary = -4  # from enum XlParameterDataType
    xlParamTypeLongVarChar = -1  # from enum XlParameterDataType
    xlParamTypeNumeric = 2  # from enum XlParameterDataType
    xlParamTypeReal = 7  # from enum XlParameterDataType
    xlParamTypeSmallInt = 5  # from enum XlParameterDataType
    xlParamTypeTime = 10  # from enum XlParameterDataType
    xlParamTypeTimestamp = 11  # from enum XlParameterDataType
    xlParamTypeTinyInt = -6  # from enum XlParameterDataType
    xlParamTypeUnknown = 0  # from enum XlParameterDataType
    xlParamTypeVarBinary = -3  # from enum XlParameterDataType
    xlParamTypeVarChar = 12  # from enum XlParameterDataType
    xlParamTypeWChar = -8  # from enum XlParameterDataType


class ParameterType:
    xlConstant = 1  # from enum XlParameterType
    xlPrompt = 0  # from enum XlParameterType
    xlRange = 2  # from enum XlParameterType


class PasteSpecialOperation:
    xlPasteSpecialOperationAdd = 2  # from enum XlPasteSpecialOperation
    xlPasteSpecialOperationDivide = 5  # from enum XlPasteSpecialOperation
    xlPasteSpecialOperationMultiply = 4  # from enum XlPasteSpecialOperation
    xlPasteSpecialOperationNone = -4142  # from enum XlPasteSpecialOperation
    xlPasteSpecialOperationSubtract = 3  # from enum XlPasteSpecialOperation


class PasteType:
    xlPasteAll = -4104  # from enum XlPasteType
    xlPasteAllExceptBorders = 7  # from enum XlPasteType
    xlPasteAllMergingConditionalFormats = 14  # from enum XlPasteType
    xlPasteAllUsingSourceTheme = 13  # from enum XlPasteType
    xlPasteColumnWidths = 8  # from enum XlPasteType
    xlPasteComments = -4144  # from enum XlPasteType
    xlPasteFormats = -4122  # from enum XlPasteType
    xlPasteFormulas = -4123  # from enum XlPasteType
    xlPasteFormulasAndNumberFormats = 11  # from enum XlPasteType
    xlPasteValidation = 6  # from enum XlPasteType
    xlPasteValues = -4163  # from enum XlPasteType
    xlPasteValuesAndNumberFormats = 12  # from enum XlPasteType


class Pattern:
    xlPatternAutomatic = -4105  # from enum XlPattern
    xlPatternChecker = 9  # from enum XlPattern
    xlPatternCrissCross = 16  # from enum XlPattern
    xlPatternDown = -4121  # from enum XlPattern
    xlPatternGray16 = 17  # from enum XlPattern
    xlPatternGray25 = -4124  # from enum XlPattern
    xlPatternGray50 = -4125  # from enum XlPattern
    xlPatternGray75 = -4126  # from enum XlPattern
    xlPatternGray8 = 18  # from enum XlPattern
    xlPatternGrid = 15  # from enum XlPattern
    xlPatternHorizontal = -4128  # from enum XlPattern
    xlPatternLightDown = 13  # from enum XlPattern
    xlPatternLightHorizontal = 11  # from enum XlPattern
    xlPatternLightUp = 14  # from enum XlPattern
    xlPatternLightVertical = 12  # from enum XlPattern
    xlPatternLinearGradient = 4000  # from enum XlPattern
    xlPatternNone = -4142  # from enum XlPattern
    xlPatternRectangularGradient = 4001  # from enum XlPattern
    xlPatternSemiGray75 = 10  # from enum XlPattern
    xlPatternSolid = 1  # from enum XlPattern
    xlPatternUp = -4162  # from enum XlPattern
    xlPatternVertical = -4166  # from enum XlPattern


class PhoneticAlignment:
    xlPhoneticAlignCenter = 2  # from enum XlPhoneticAlignment
    xlPhoneticAlignDistributed = 3  # from enum XlPhoneticAlignment
    xlPhoneticAlignLeft = 1  # from enum XlPhoneticAlignment
    xlPhoneticAlignNoControl = 0  # from enum XlPhoneticAlignment


class PhoneticCharacterType:
    xlHiragana = 2  # from enum XlPhoneticCharacterType
    xlKatakana = 1  # from enum XlPhoneticCharacterType
    xlKatakanaHalf = 0  # from enum XlPhoneticCharacterType
    xlNoConversion = 3  # from enum XlPhoneticCharacterType


class PictureAppearance:
    xlPrinter = 2  # from enum XlPictureAppearance
    xlScreen = 1  # from enum XlPictureAppearance


class PictureConvertorType:
    xlBMP = 1  # from enum XlPictureConvertorType
    xlCGM = 7  # from enum XlPictureConvertorType
    xlDRW = 4  # from enum XlPictureConvertorType
    xlDXF = 5  # from enum XlPictureConvertorType
    xlEPS = 8  # from enum XlPictureConvertorType
    xlHGL = 6  # from enum XlPictureConvertorType
    xlPCT = 13  # from enum XlPictureConvertorType
    xlPCX = 10  # from enum XlPictureConvertorType
    xlPIC = 11  # from enum XlPictureConvertorType
    xlPLT = 12  # from enum XlPictureConvertorType
    xlTIF = 9  # from enum XlPictureConvertorType
    xlWMF = 2  # from enum XlPictureConvertorType
    xlWPG = 3  # from enum XlPictureConvertorType


class PieSliceIndex:
    xlCenterPoint = 5  # from enum XlPieSliceIndex
    xlInnerCenterPoint = 8  # from enum XlPieSliceIndex
    xlInnerClockwisePoint = 7  # from enum XlPieSliceIndex
    xlInnerCounterClockwisePoint = 9  # from enum XlPieSliceIndex
    xlMidClockwiseRadiusPoint = 4  # from enum XlPieSliceIndex
    xlMidCounterClockwiseRadiusPoint = 6  # from enum XlPieSliceIndex
    xlOuterCenterPoint = 2  # from enum XlPieSliceIndex
    xlOuterClockwisePoint = 3  # from enum XlPieSliceIndex
    xlOuterCounterClockwisePoint = 1  # from enum XlPieSliceIndex


class PieSliceLocation:
    xlHorizontalCoordinate = 1  # from enum XlPieSliceLocation
    xlVerticalCoordinate = 2  # from enum XlPieSliceLocation


class PivotCellType:
    xlPivotCellBlankCell = 9  # from enum XlPivotCellType
    xlPivotCellCustomSubtotal = 7  # from enum XlPivotCellType
    xlPivotCellDataField = 4  # from enum XlPivotCellType
    xlPivotCellDataPivotField = 8  # from enum XlPivotCellType
    xlPivotCellGrandTotal = 3  # from enum XlPivotCellType
    xlPivotCellPageFieldItem = 6  # from enum XlPivotCellType
    xlPivotCellPivotField = 5  # from enum XlPivotCellType
    xlPivotCellPivotItem = 1  # from enum XlPivotCellType
    xlPivotCellSubtotal = 2  # from enum XlPivotCellType
    xlPivotCellValue = 0  # from enum XlPivotCellType


class PivotConditionScope:
    xlDataFieldScope = 2  # from enum XlPivotConditionScope
    xlFieldsScope = 1  # from enum XlPivotConditionScope
    xlSelectionScope = 0  # from enum XlPivotConditionScope


class PivotFieldCalculation:
    xlDifferenceFrom = 2  # from enum XlPivotFieldCalculation
    xlIndex = 9  # from enum XlPivotFieldCalculation
    xlNoAdditionalCalculation = -4143  # from enum XlPivotFieldCalculation
    xlPercentDifferenceFrom = 4  # from enum XlPivotFieldCalculation
    xlPercentOf = 3  # from enum XlPivotFieldCalculation
    xlPercentOfColumn = 7  # from enum XlPivotFieldCalculation
    xlPercentOfParent = 12  # from enum XlPivotFieldCalculation
    xlPercentOfParentColumn = 11  # from enum XlPivotFieldCalculation
    xlPercentOfParentRow = 10  # from enum XlPivotFieldCalculation
    xlPercentOfRow = 6  # from enum XlPivotFieldCalculation
    xlPercentOfTotal = 8  # from enum XlPivotFieldCalculation
    xlPercentRunningTotal = 13  # from enum XlPivotFieldCalculation
    xlRankAscending = 14  # from enum XlPivotFieldCalculation
    xlRankDecending = 15  # from enum XlPivotFieldCalculation
    xlRunningTotal = 5  # from enum XlPivotFieldCalculation


class PivotFieldDataType:
    xlDate = 2  # from enum XlPivotFieldDataType
    xlNumber = -4145  # from enum XlPivotFieldDataType
    xlText = -4158  # from enum XlPivotFieldDataType


class PivotFieldOrientation:
    xlColumnField = 2  # from enum XlPivotFieldOrientation
    xlDataField = 4  # from enum XlPivotFieldOrientation
    xlHidden = 0  # from enum XlPivotFieldOrientation
    xlPageField = 3  # from enum XlPivotFieldOrientation
    xlRowField = 1  # from enum XlPivotFieldOrientation


class PivotFieldRepeatLabels:
    xlDoNotRepeatLabels = 1  # from enum XlPivotFieldRepeatLabels
    xlRepeatLabels = 2  # from enum XlPivotFieldRepeatLabels


class PivotFilterType:
    xlAfter = 33  # from enum XlPivotFilterType
    xlAfterOrEqualTo = 34  # from enum XlPivotFilterType
    xlAllDatesInPeriodApril = 60  # from enum XlPivotFilterType
    xlAllDatesInPeriodAugust = 64  # from enum XlPivotFilterType
    xlAllDatesInPeriodDecember = 68  # from enum XlPivotFilterType
    xlAllDatesInPeriodFebruary = 58  # from enum XlPivotFilterType
    xlAllDatesInPeriodJanuary = 57  # from enum XlPivotFilterType
    xlAllDatesInPeriodJuly = 63  # from enum XlPivotFilterType
    xlAllDatesInPeriodJune = 62  # from enum XlPivotFilterType
    xlAllDatesInPeriodMarch = 59  # from enum XlPivotFilterType
    xlAllDatesInPeriodMay = 61  # from enum XlPivotFilterType
    xlAllDatesInPeriodNovember = 67  # from enum XlPivotFilterType
    xlAllDatesInPeriodOctober = 66  # from enum XlPivotFilterType
    xlAllDatesInPeriodQuarter1 = 53  # from enum XlPivotFilterType
    xlAllDatesInPeriodQuarter2 = 54  # from enum XlPivotFilterType
    xlAllDatesInPeriodQuarter3 = 55  # from enum XlPivotFilterType
    xlAllDatesInPeriodQuarter4 = 56  # from enum XlPivotFilterType
    xlAllDatesInPeriodSeptember = 65  # from enum XlPivotFilterType
    xlBefore = 31  # from enum XlPivotFilterType
    xlBeforeOrEqualTo = 32  # from enum XlPivotFilterType
    xlBottomCount = 2  # from enum XlPivotFilterType
    xlBottomPercent = 4  # from enum XlPivotFilterType
    xlBottomSum = 6  # from enum XlPivotFilterType
    xlCaptionBeginsWith = 17  # from enum XlPivotFilterType
    xlCaptionContains = 21  # from enum XlPivotFilterType
    xlCaptionDoesNotBeginWith = 18  # from enum XlPivotFilterType
    xlCaptionDoesNotContain = 22  # from enum XlPivotFilterType
    xlCaptionDoesNotEndWith = 20  # from enum XlPivotFilterType
    xlCaptionDoesNotEqual = 16  # from enum XlPivotFilterType
    xlCaptionEndsWith = 19  # from enum XlPivotFilterType
    xlCaptionEquals = 15  # from enum XlPivotFilterType
    xlCaptionIsBetween = 27  # from enum XlPivotFilterType
    xlCaptionIsGreaterThan = 23  # from enum XlPivotFilterType
    xlCaptionIsGreaterThanOrEqualTo = 24  # from enum XlPivotFilterType
    xlCaptionIsLessThan = 25  # from enum XlPivotFilterType
    xlCaptionIsLessThanOrEqualTo = 26  # from enum XlPivotFilterType
    xlCaptionIsNotBetween = 28  # from enum XlPivotFilterType
    xlDateBetween = 35  # from enum XlPivotFilterType
    xlDateLastMonth = 45  # from enum XlPivotFilterType
    xlDateLastQuarter = 48  # from enum XlPivotFilterType
    xlDateLastWeek = 42  # from enum XlPivotFilterType
    xlDateLastYear = 51  # from enum XlPivotFilterType
    xlDateNextMonth = 43  # from enum XlPivotFilterType
    xlDateNextQuarter = 46  # from enum XlPivotFilterType
    xlDateNextWeek = 40  # from enum XlPivotFilterType
    xlDateNextYear = 49  # from enum XlPivotFilterType
    xlDateNotBetween = 36  # from enum XlPivotFilterType
    xlDateThisMonth = 44  # from enum XlPivotFilterType
    xlDateThisQuarter = 47  # from enum XlPivotFilterType
    xlDateThisWeek = 41  # from enum XlPivotFilterType
    xlDateThisYear = 50  # from enum XlPivotFilterType
    xlDateToday = 38  # from enum XlPivotFilterType
    xlDateTomorrow = 37  # from enum XlPivotFilterType
    xlDateYesterday = 39  # from enum XlPivotFilterType
    xlNotSpecificDate = 30  # from enum XlPivotFilterType
    xlSpecificDate = 29  # from enum XlPivotFilterType
    xlTopCount = 1  # from enum XlPivotFilterType
    xlTopPercent = 3  # from enum XlPivotFilterType
    xlTopSum = 5  # from enum XlPivotFilterType
    xlValueDoesNotEqual = 8  # from enum XlPivotFilterType
    xlValueEquals = 7  # from enum XlPivotFilterType
    xlValueIsBetween = 13  # from enum XlPivotFilterType
    xlValueIsGreaterThan = 9  # from enum XlPivotFilterType
    xlValueIsGreaterThanOrEqualTo = 10  # from enum XlPivotFilterType
    xlValueIsLessThan = 11  # from enum XlPivotFilterType
    xlValueIsLessThanOrEqualTo = 12  # from enum XlPivotFilterType
    xlValueIsNotBetween = 14  # from enum XlPivotFilterType
    xlYearToDate = 52  # from enum XlPivotFilterType


class PivotFormatType:
    xlPTClassic = 20  # from enum XlPivotFormatType
    xlPTNone = 21  # from enum XlPivotFormatType
    xlReport1 = 0  # from enum XlPivotFormatType
    xlReport10 = 9  # from enum XlPivotFormatType
    xlReport2 = 1  # from enum XlPivotFormatType
    xlReport3 = 2  # from enum XlPivotFormatType
    xlReport4 = 3  # from enum XlPivotFormatType
    xlReport5 = 4  # from enum XlPivotFormatType
    xlReport6 = 5  # from enum XlPivotFormatType
    xlReport7 = 6  # from enum XlPivotFormatType
    xlReport8 = 7  # from enum XlPivotFormatType
    xlReport9 = 8  # from enum XlPivotFormatType
    xlTable1 = 10  # from enum XlPivotFormatType
    xlTable10 = 19  # from enum XlPivotFormatType
    xlTable2 = 11  # from enum XlPivotFormatType
    xlTable3 = 12  # from enum XlPivotFormatType
    xlTable4 = 13  # from enum XlPivotFormatType
    xlTable5 = 14  # from enum XlPivotFormatType
    xlTable6 = 15  # from enum XlPivotFormatType
    xlTable7 = 16  # from enum XlPivotFormatType
    xlTable8 = 17  # from enum XlPivotFormatType
    xlTable9 = 18  # from enum XlPivotFormatType


class PivotLineType:
    xlPivotLineBlank = 3  # from enum XlPivotLineType
    xlPivotLineGrandTotal = 2  # from enum XlPivotLineType
    xlPivotLineRegular = 0  # from enum XlPivotLineType
    xlPivotLineSubtotal = 1  # from enum XlPivotLineType


class PivotTableMissingItems:
    xlMissingItemsDefault = -1  # from enum XlPivotTableMissingItems
    xlMissingItemsMax = 32500  # from enum XlPivotTableMissingItems
    xlMissingItemsMax2 = 1048576  # from enum XlPivotTableMissingItems
    xlMissingItemsNone = 0  # from enum XlPivotTableMissingItems


class PivotTableSourceType:
    xlConsolidation = 3  # from enum XlPivotTableSourceType
    xlDatabase = 1  # from enum XlPivotTableSourceType
    xlExternal = 2  # from enum XlPivotTableSourceType
    xlPivotTable = -4148  # from enum XlPivotTableSourceType
    xlScenario = 4  # from enum XlPivotTableSourceType


class PivotTableVersionList:
    xlPivotTableVersion10 = 1  # from enum XlPivotTableVersionList
    xlPivotTableVersion11 = 2  # from enum XlPivotTableVersionList
    xlPivotTableVersion12 = 3  # from enum XlPivotTableVersionList
    xlPivotTableVersion14 = 4  # from enum XlPivotTableVersionList
    xlPivotTableVersion2000 = 0  # from enum XlPivotTableVersionList
    xlPivotTableVersionCurrent = -1  # from enum XlPivotTableVersionList


class Placement:
    xlFreeFloating = 3  # from enum XlPlacement
    xlMove = 2  # from enum XlPlacement
    xlMoveAndSize = 1  # from enum XlPlacement


class Platform:
    xlMSDOS = 3  # from enum XlPlatform
    xlMacintosh = 1  # from enum XlPlatform
    xlWindows = 2  # from enum XlPlatform


class PortugueseReform:
    xlPortugueseBoth = 3  # from enum XlPortugueseReform
    xlPortuguesePostReform = 2  # from enum XlPortugueseReform
    xlPortuguesePreReform = 1  # from enum XlPortugueseReform


class PrintErrors:
    xlPrintErrorsBlank = 1  # from enum XlPrintErrors
    xlPrintErrorsDash = 2  # from enum XlPrintErrors
    xlPrintErrorsDisplayed = 0  # from enum XlPrintErrors
    xlPrintErrorsNA = 3  # from enum XlPrintErrors


class PrintLocation:
    xlPrintInPlace = 16  # from enum XlPrintLocation
    xlPrintNoComments = -4142  # from enum XlPrintLocation
    xlPrintSheetEnd = 1  # from enum XlPrintLocation


class Priority:
    xlPriorityHigh = -4127  # from enum XlPriority
    xlPriorityLow = -4134  # from enum XlPriority
    xlPriorityNormal = -4143  # from enum XlPriority


class PropertyDisplayedIn:
    xlDisplayPropertyInPivotTable = 1  # from enum XlPropertyDisplayedIn
    xlDisplayPropertyInPivotTableAndTooltip = 3  # from enum XlPropertyDisplayedIn
    xlDisplayPropertyInTooltip = 2  # from enum XlPropertyDisplayedIn


class ProtectedViewCloseReason:
    xlProtectedViewCloseEdit = 1  # from enum XlProtectedViewCloseReason
    xlProtectedViewCloseForced = 2  # from enum XlProtectedViewCloseReason
    xlProtectedViewCloseNormal = 0  # from enum XlProtectedViewCloseReason


class ProtectedViewWindowState:
    xlProtectedViewWindowMaximized = 2  # from enum XlProtectedViewWindowState
    xlProtectedViewWindowMinimized = 1  # from enum XlProtectedViewWindowState
    xlProtectedViewWindowNormal = 0  # from enum XlProtectedViewWindowState


class QueryType:
    xlADORecordset = 7  # from enum XlQueryType
    xlDAORecordset = 2  # from enum XlQueryType
    xlODBCQuery = 1  # from enum XlQueryType
    xlOLEDBQuery = 5  # from enum XlQueryType
    xlTextImport = 6  # from enum XlQueryType
    xlWebQuery = 4  # from enum XlQueryType


class RangeAutoFormat:
    xlRangeAutoFormat3DEffects1 = 13  # from enum XlRangeAutoFormat
    xlRangeAutoFormat3DEffects2 = 14  # from enum XlRangeAutoFormat
    xlRangeAutoFormatAccounting1 = 4  # from enum XlRangeAutoFormat
    xlRangeAutoFormatAccounting2 = 5  # from enum XlRangeAutoFormat
    xlRangeAutoFormatAccounting3 = 6  # from enum XlRangeAutoFormat
    xlRangeAutoFormatAccounting4 = 17  # from enum XlRangeAutoFormat
    xlRangeAutoFormatClassic1 = 1  # from enum XlRangeAutoFormat
    xlRangeAutoFormatClassic2 = 2  # from enum XlRangeAutoFormat
    xlRangeAutoFormatClassic3 = 3  # from enum XlRangeAutoFormat
    xlRangeAutoFormatClassicPivotTable = 31  # from enum XlRangeAutoFormat
    xlRangeAutoFormatColor1 = 7  # from enum XlRangeAutoFormat
    xlRangeAutoFormatColor2 = 8  # from enum XlRangeAutoFormat
    xlRangeAutoFormatColor3 = 9  # from enum XlRangeAutoFormat
    xlRangeAutoFormatList1 = 10  # from enum XlRangeAutoFormat
    xlRangeAutoFormatList2 = 11  # from enum XlRangeAutoFormat
    xlRangeAutoFormatList3 = 12  # from enum XlRangeAutoFormat
    xlRangeAutoFormatLocalFormat1 = 15  # from enum XlRangeAutoFormat
    xlRangeAutoFormatLocalFormat2 = 16  # from enum XlRangeAutoFormat
    xlRangeAutoFormatLocalFormat3 = 19  # from enum XlRangeAutoFormat
    xlRangeAutoFormatLocalFormat4 = 20  # from enum XlRangeAutoFormat
    xlRangeAutoFormatNone = -4142  # from enum XlRangeAutoFormat
    xlRangeAutoFormatPTNone = 42  # from enum XlRangeAutoFormat
    xlRangeAutoFormatReport1 = 21  # from enum XlRangeAutoFormat
    xlRangeAutoFormatReport10 = 30  # from enum XlRangeAutoFormat
    xlRangeAutoFormatReport2 = 22  # from enum XlRangeAutoFormat
    xlRangeAutoFormatReport3 = 23  # from enum XlRangeAutoFormat
    xlRangeAutoFormatReport4 = 24  # from enum XlRangeAutoFormat
    xlRangeAutoFormatReport5 = 25  # from enum XlRangeAutoFormat
    xlRangeAutoFormatReport6 = 26  # from enum XlRangeAutoFormat
    xlRangeAutoFormatReport7 = 27  # from enum XlRangeAutoFormat
    xlRangeAutoFormatReport8 = 28  # from enum XlRangeAutoFormat
    xlRangeAutoFormatReport9 = 29  # from enum XlRangeAutoFormat
    xlRangeAutoFormatSimple = -4154  # from enum XlRangeAutoFormat
    xlRangeAutoFormatTable1 = 32  # from enum XlRangeAutoFormat
    xlRangeAutoFormatTable10 = 41  # from enum XlRangeAutoFormat
    xlRangeAutoFormatTable2 = 33  # from enum XlRangeAutoFormat
    xlRangeAutoFormatTable3 = 34  # from enum XlRangeAutoFormat
    xlRangeAutoFormatTable4 = 35  # from enum XlRangeAutoFormat
    xlRangeAutoFormatTable5 = 36  # from enum XlRangeAutoFormat
    xlRangeAutoFormatTable6 = 37  # from enum XlRangeAutoFormat
    xlRangeAutoFormatTable7 = 38  # from enum XlRangeAutoFormat
    xlRangeAutoFormatTable8 = 39  # from enum XlRangeAutoFormat
    xlRangeAutoFormatTable9 = 40  # from enum XlRangeAutoFormat


class RangeValueDataType:
    xlRangeValueDefault = 10  # from enum XlRangeValueDataType
    xlRangeValueMSPersistXML = 12  # from enum XlRangeValueDataType
    xlRangeValueXMLSpreadsheet = 11  # from enum XlRangeValueDataType


class ReferenceStyle:
    xlA1 = 1  # from enum XlReferenceStyle
    xlR1C1 = -4150  # from enum XlReferenceStyle


class ReferenceType:
    xlAbsRowRelColumn = 2  # from enum XlReferenceType
    xlAbsolute = 1  # from enum XlReferenceType
    xlRelRowAbsColumn = 3  # from enum XlReferenceType
    xlRelative = 4  # from enum XlReferenceType


class RemoveDocInfoType:
    xlRDIAll = 99  # from enum XlRemoveDocInfoType
    xlRDIComments = 1  # from enum XlRemoveDocInfoType
    xlRDIContentType = 16  # from enum XlRemoveDocInfoType
    xlRDIDefinedNameComments = 18  # from enum XlRemoveDocInfoType
    xlRDIDocumentManagementPolicy = 15  # from enum XlRemoveDocInfoType
    xlRDIDocumentProperties = 8  # from enum XlRemoveDocInfoType
    xlRDIDocumentServerProperties = 14  # from enum XlRemoveDocInfoType
    xlRDIDocumentWorkspace = 10  # from enum XlRemoveDocInfoType
    xlRDIEmailHeader = 5  # from enum XlRemoveDocInfoType
    xlRDIInactiveDataConnections = 19  # from enum XlRemoveDocInfoType
    xlRDIInkAnnotations = 11  # from enum XlRemoveDocInfoType
    xlRDIPrinterPath = 20  # from enum XlRemoveDocInfoType
    xlRDIPublishInfo = 13  # from enum XlRemoveDocInfoType
    xlRDIRemovePersonalInformation = 4  # from enum XlRemoveDocInfoType
    xlRDIRoutingSlip = 6  # from enum XlRemoveDocInfoType
    xlRDIScenarioComments = 12  # from enum XlRemoveDocInfoType
    xlRDISendForReview = 7  # from enum XlRemoveDocInfoType


class RgbColor:
    rgbAliceBlue = 16775408  # from enum XlRgbColor
    rgbAntiqueWhite = 14150650  # from enum XlRgbColor
    rgbAqua = 16776960  # from enum XlRgbColor
    rgbAquamarine = 13959039  # from enum XlRgbColor
    rgbAzure = 16777200  # from enum XlRgbColor
    rgbBeige = 14480885  # from enum XlRgbColor
    rgbBisque = 12903679  # from enum XlRgbColor
    rgbBlack = 0  # from enum XlRgbColor
    rgbBlanchedAlmond = 13495295  # from enum XlRgbColor
    rgbBlue = 16711680  # from enum XlRgbColor
    rgbBlueViolet = 14822282  # from enum XlRgbColor
    rgbBrown = 2763429  # from enum XlRgbColor
    rgbBurlyWood = 8894686  # from enum XlRgbColor
    rgbCadetBlue = 10526303  # from enum XlRgbColor
    rgbChartreuse = 65407  # from enum XlRgbColor
    rgbCoral = 5275647  # from enum XlRgbColor
    rgbCornflowerBlue = 15570276  # from enum XlRgbColor
    rgbCornsilk = 14481663  # from enum XlRgbColor
    rgbCrimson = 3937500  # from enum XlRgbColor
    rgbDarkBlue = 9109504  # from enum XlRgbColor
    rgbDarkCyan = 9145088  # from enum XlRgbColor
    rgbDarkGoldenrod = 755384  # from enum XlRgbColor
    rgbDarkGray = 11119017  # from enum XlRgbColor
    rgbDarkGreen = 25600  # from enum XlRgbColor
    rgbDarkGrey = 11119017  # from enum XlRgbColor
    rgbDarkKhaki = 7059389  # from enum XlRgbColor
    rgbDarkMagenta = 9109643  # from enum XlRgbColor
    rgbDarkOliveGreen = 3107669  # from enum XlRgbColor
    rgbDarkOrange = 36095  # from enum XlRgbColor
    rgbDarkOrchid = 13382297  # from enum XlRgbColor
    rgbDarkRed = 139  # from enum XlRgbColor
    rgbDarkSalmon = 8034025  # from enum XlRgbColor
    rgbDarkSeaGreen = 9419919  # from enum XlRgbColor
    rgbDarkSlateBlue = 9125192  # from enum XlRgbColor
    rgbDarkSlateGray = 5197615  # from enum XlRgbColor
    rgbDarkSlateGrey = 5197615  # from enum XlRgbColor
    rgbDarkTurquoise = 13749760  # from enum XlRgbColor
    rgbDarkViolet = 13828244  # from enum XlRgbColor
    rgbDeepPink = 9639167  # from enum XlRgbColor
    rgbDeepSkyBlue = 16760576  # from enum XlRgbColor
    rgbDimGray = 6908265  # from enum XlRgbColor
    rgbDimGrey = 6908265  # from enum XlRgbColor
    rgbDodgerBlue = 16748574  # from enum XlRgbColor
    rgbFireBrick = 2237106  # from enum XlRgbColor
    rgbFloralWhite = 15792895  # from enum XlRgbColor
    rgbForestGreen = 2263842  # from enum XlRgbColor
    rgbFuchsia = 16711935  # from enum XlRgbColor
    rgbGainsboro = 14474460  # from enum XlRgbColor
    rgbGhostWhite = 16775416  # from enum XlRgbColor
    rgbGold = 55295  # from enum XlRgbColor
    rgbGoldenrod = 2139610  # from enum XlRgbColor
    rgbGray = 8421504  # from enum XlRgbColor
    rgbGreen = 32768  # from enum XlRgbColor
    rgbGreenYellow = 3145645  # from enum XlRgbColor
    rgbGrey = 8421504  # from enum XlRgbColor
    rgbHoneydew = 15794160  # from enum XlRgbColor
    rgbHotPink = 11823615  # from enum XlRgbColor
    rgbIndianRed = 6053069  # from enum XlRgbColor
    rgbIndigo = 8519755  # from enum XlRgbColor
    rgbIvory = 15794175  # from enum XlRgbColor
    rgbKhaki = 9234160  # from enum XlRgbColor
    rgbLavender = 16443110  # from enum XlRgbColor
    rgbLavenderBlush = 16118015  # from enum XlRgbColor
    rgbLawnGreen = 64636  # from enum XlRgbColor
    rgbLemonChiffon = 13499135  # from enum XlRgbColor
    rgbLightBlue = 15128749  # from enum XlRgbColor
    rgbLightCoral = 8421616  # from enum XlRgbColor
    rgbLightCyan = 9145088  # from enum XlRgbColor
    rgbLightGoldenrodYellow = 13826810  # from enum XlRgbColor
    rgbLightGray = 13882323  # from enum XlRgbColor
    rgbLightGreen = 9498256  # from enum XlRgbColor
    rgbLightGrey = 13882323  # from enum XlRgbColor
    rgbLightPink = 12695295  # from enum XlRgbColor
    rgbLightSalmon = 8036607  # from enum XlRgbColor
    rgbLightSeaGreen = 11186720  # from enum XlRgbColor
    rgbLightSkyBlue = 16436871  # from enum XlRgbColor
    rgbLightSlateGray = 10061943  # from enum XlRgbColor
    rgbLightSlateGrey = 10061943  # from enum XlRgbColor
    rgbLightSteelBlue = 14599344  # from enum XlRgbColor
    rgbLightYellow = 14745599  # from enum XlRgbColor
    rgbLime = 65280  # from enum XlRgbColor
    rgbLimeGreen = 3329330  # from enum XlRgbColor
    rgbLinen = 15134970  # from enum XlRgbColor
    rgbMaroon = 128  # from enum XlRgbColor
    rgbMediumAquamarine = 11206502  # from enum XlRgbColor
    rgbMediumBlue = 13434880  # from enum XlRgbColor
    rgbMediumOrchid = 13850042  # from enum XlRgbColor
    rgbMediumPurple = 14381203  # from enum XlRgbColor
    rgbMediumSeaGreen = 7451452  # from enum XlRgbColor
    rgbMediumSlateBlue = 15624315  # from enum XlRgbColor
    rgbMediumSpringGreen = 10156544  # from enum XlRgbColor
    rgbMediumTurquoise = 13422920  # from enum XlRgbColor
    rgbMediumVioletRed = 8721863  # from enum XlRgbColor
    rgbMidnightBlue = 7346457  # from enum XlRgbColor
    rgbMintCream = 16449525  # from enum XlRgbColor
    rgbMistyRose = 14804223  # from enum XlRgbColor
    rgbMoccasin = 11920639  # from enum XlRgbColor
    rgbNavajoWhite = 11394815  # from enum XlRgbColor
    rgbNavy = 8388608  # from enum XlRgbColor
    rgbNavyBlue = 8388608  # from enum XlRgbColor
    rgbOldLace = 15136253  # from enum XlRgbColor
    rgbOlive = 32896  # from enum XlRgbColor
    rgbOliveDrab = 2330219  # from enum XlRgbColor
    rgbOrange = 42495  # from enum XlRgbColor
    rgbOrangeRed = 17919  # from enum XlRgbColor
    rgbOrchid = 14053594  # from enum XlRgbColor
    rgbPaleGoldenrod = 7071982  # from enum XlRgbColor
    rgbPaleGreen = 10025880  # from enum XlRgbColor
    rgbPaleTurquoise = 15658671  # from enum XlRgbColor
    rgbPaleVioletRed = 9662683  # from enum XlRgbColor
    rgbPapayaWhip = 14020607  # from enum XlRgbColor
    rgbPeachPuff = 12180223  # from enum XlRgbColor
    rgbPeru = 4163021  # from enum XlRgbColor
    rgbPink = 13353215  # from enum XlRgbColor
    rgbPlum = 14524637  # from enum XlRgbColor
    rgbPowderBlue = 15130800  # from enum XlRgbColor
    rgbPurple = 8388736  # from enum XlRgbColor
    rgbRed = 255  # from enum XlRgbColor
    rgbRosyBrown = 9408444  # from enum XlRgbColor
    rgbRoyalBlue = 14772545  # from enum XlRgbColor
    rgbSalmon = 7504122  # from enum XlRgbColor
    rgbSandyBrown = 6333684  # from enum XlRgbColor
    rgbSeaGreen = 5737262  # from enum XlRgbColor
    rgbSeashell = 15660543  # from enum XlRgbColor
    rgbSienna = 2970272  # from enum XlRgbColor
    rgbSilver = 12632256  # from enum XlRgbColor
    rgbSkyBlue = 15453831  # from enum XlRgbColor
    rgbSlateBlue = 13458026  # from enum XlRgbColor
    rgbSlateGray = 9470064  # from enum XlRgbColor
    rgbSlateGrey = 9470064  # from enum XlRgbColor
    rgbSnow = 16448255  # from enum XlRgbColor
    rgbSpringGreen = 8388352  # from enum XlRgbColor
    rgbSteelBlue = 11829830  # from enum XlRgbColor
    rgbTan = 9221330  # from enum XlRgbColor
    rgbTeal = 8421376  # from enum XlRgbColor
    rgbThistle = 14204888  # from enum XlRgbColor
    rgbTomato = 4678655  # from enum XlRgbColor
    rgbTurquoise = 13688896  # from enum XlRgbColor
    rgbViolet = 15631086  # from enum XlRgbColor
    rgbWheat = 11788021  # from enum XlRgbColor
    rgbWhite = 16777215  # from enum XlRgbColor
    rgbWhiteSmoke = 16119285  # from enum XlRgbColor
    rgbYellow = 65535  # from enum XlRgbColor
    rgbYellowGreen = 3329434  # from enum XlRgbColor

class RobustConnect:
    xlAlways = 1  # from enum XlRobustConnect
    xlAsRequired = 0  # from enum XlRobustConnect
    xlNever = 2  # from enum XlRobustConnect


class RoutingSlipDelivery:
    xlAllAtOnce = 2  # from enum XlRoutingSlipDelivery
    xlOneAfterAnother = 1  # from enum XlRoutingSlipDelivery


class RoutingSlipStatus:
    xlNotYetRouted = 0  # from enum XlRoutingSlipStatus
    xlRoutingComplete = 2  # from enum XlRoutingSlipStatus
    xlRoutingInProgress = 1  # from enum XlRoutingSlipStatus


class RowCol:
    xlColumns = 2  # from enum XlRowCol
    xlRows = 1  # from enum XlRowCol


class RunAutoMacro:
    xlAutoActivate = 3  # from enum XlRunAutoMacro
    xlAutoClose = 2  # from enum XlRunAutoMacro
    xlAutoDeactivate = 4  # from enum XlRunAutoMacro
    xlAutoOpen = 1  # from enum XlRunAutoMacro


class SaveAction:
    xlDoNotSaveChanges = 2  # from enum XlSaveAction
    xlSaveChanges = 1  # from enum XlSaveAction


class SaveAsAccessMode:
    xlExclusive = 3  # from enum XlSaveAsAccessMode
    xlNoChange = 1  # from enum XlSaveAsAccessMode
    xlShared = 2  # from enum XlSaveAsAccessMode


class SaveConflictResolution:
    xlLocalSessionChanges = 2  # from enum XlSaveConflictResolution
    xlOtherSessionChanges = 3  # from enum XlSaveConflictResolution
    xlUserResolution = 1  # from enum XlSaveConflictResolution


class ScaleType:
    xlScaleLinear = -4132  # from enum XlScaleType
    xlScaleLogarithmic = -4133  # from enum XlScaleType


class SearchDirection:
    xlNext = 1  # from enum XlSearchDirection
    xlPrevious = 2  # from enum XlSearchDirection


class SearchOrder:
    xlByColumns = 2  # from enum XlSearchOrder
    xlByRows = 1  # from enum XlSearchOrder


class SearchWithin:
    xlWithinSheet = 1  # from enum XlSearchWithin
    xlWithinWorkbook = 2  # from enum XlSearchWithin


class SheetType:
    xlChart = -4109  # from enum XlSheetType
    xlDialogSheet = -4116  # from enum XlSheetType
    xlExcel4IntlMacroSheet = 4  # from enum XlSheetType
    xlExcel4MacroSheet = 3  # from enum XlSheetType
    xlWorksheet = -4167  # from enum XlSheetType


class SheetVisibility:
    xlSheetHidden = 0  # from enum XlSheetVisibility
    xlSheetVeryHidden = 2  # from enum XlSheetVisibility
    xlSheetVisible = -1  # from enum XlSheetVisibility


class SizeRepresents:
    xlSizeIsArea = 1  # from enum XlSizeRepresents
    xlSizeIsWidth = 2  # from enum XlSizeRepresents


class SlicerCrossFilterType:
    xlSlicerCrossFilterShowItemsWithDataAtTop = 2  # from enum XlSlicerCrossFilterType
    xlSlicerCrossFilterShowItemsWithNoData = 3  # from enum XlSlicerCrossFilterType
    xlSlicerNoCrossFilter = 1  # from enum XlSlicerCrossFilterType


class SlicerSort:
    xlSlicerSortAscending = 2  # from enum XlSlicerSort
    xlSlicerSortDataSourceOrder = 1  # from enum XlSlicerSort
    xlSlicerSortDescending = 3  # from enum XlSlicerSort


class SmartTagControlType:
    xlSmartTagControlActiveX = 13  # from enum XlSmartTagControlType
    xlSmartTagControlButton = 6  # from enum XlSmartTagControlType
    xlSmartTagControlCheckbox = 9  # from enum XlSmartTagControlType
    xlSmartTagControlCombo = 12  # from enum XlSmartTagControlType
    xlSmartTagControlHelp = 3  # from enum XlSmartTagControlType
    xlSmartTagControlHelpURL = 4  # from enum XlSmartTagControlType
    xlSmartTagControlImage = 8  # from enum XlSmartTagControlType
    xlSmartTagControlLabel = 7  # from enum XlSmartTagControlType
    xlSmartTagControlLink = 2  # from enum XlSmartTagControlType
    xlSmartTagControlListbox = 11  # from enum XlSmartTagControlType
    xlSmartTagControlRadioGroup = 14  # from enum XlSmartTagControlType
    xlSmartTagControlSeparator = 5  # from enum XlSmartTagControlType
    xlSmartTagControlSmartTag = 1  # from enum XlSmartTagControlType
    xlSmartTagControlTextbox = 10  # from enum XlSmartTagControlType


class SmartTagDisplayMode:
    xlButtonOnly = 2  # from enum XlSmartTagDisplayMode
    xlDisplayNone = 1  # from enum XlSmartTagDisplayMode
    xlIndicatorAndButton = 0  # from enum XlSmartTagDisplayMode


class SortDataOption:
    xlSortNormal = 0  # from enum XlSortDataOption
    xlSortTextAsNumbers = 1  # from enum XlSortDataOption


class SortMethod:
    xlPinYin = 1  # from enum XlSortMethod
    xlStroke = 2  # from enum XlSortMethod


class SortMethodOld:
    xlCodePage = 2  # from enum XlSortMethodOld
    xlSyllabary = 1  # from enum XlSortMethodOld


class SortOn:
    xlSortOnCellColor = 1  # from enum XlSortOn
    xlSortOnFontColor = 2  # from enum XlSortOn
    xlSortOnIcon = 3  # from enum XlSortOn
    xlSortOnValues = 0  # from enum XlSortOn


class SortOrder:
    xlAscending = 1  # from enum XlSortOrder
    xlDescending = 2  # from enum XlSortOrder


class SortOrientation:
    xlSortColumns = 1  # from enum XlSortOrientation
    xlSortRows = 2  # from enum XlSortOrientation


class SortType:
    xlSortLabels = 2  # from enum XlSortType
    xlSortValues = 1  # from enum XlSortType


class SourceType:
    xlSourceAutoFilter = 3  # from enum XlSourceType
    xlSourceChart = 5  # from enum XlSourceType
    xlSourcePivotTable = 6  # from enum XlSourceType
    xlSourcePrintArea = 2  # from enum XlSourceType
    xlSourceQuery = 7  # from enum XlSourceType
    xlSourceRange = 4  # from enum XlSourceType
    xlSourceSheet = 1  # from enum XlSourceType
    xlSourceWorkbook = 0  # from enum XlSourceType


class SpanishModes:
    xlSpanishTuteoAndVoseo = 1  # from enum XlSpanishModes
    xlSpanishTuteoOnly = 0  # from enum XlSpanishModes
    xlSpanishVoseoOnly = 2  # from enum XlSpanishModes


class SparkScale:
    xlSparkScaleCustom = 3  # from enum XlSparkScale
    xlSparkScaleGroup = 1  # from enum XlSparkScale
    xlSparkScaleSingle = 2  # from enum XlSparkScale


class SparkType:
    xlSparkColumn = 2  # from enum XlSparkType
    xlSparkColumnStacked100 = 3  # from enum XlSparkType
    xlSparkLine = 1  # from enum XlSparkType


class SparklineRowCol:
    xlSparklineColumnsSquare = 2  # from enum XlSparklineRowCol
    xlSparklineNonSquare = 0  # from enum XlSparklineRowCol
    xlSparklineRowsSquare = 1  # from enum XlSparklineRowCol


class SpeakDirection:
    xlSpeakByColumns = 1  # from enum XlSpeakDirection
    xlSpeakByRows = 0  # from enum XlSpeakDirection


class SpecialCellsValue:
    xlErrors = 16  # from enum XlSpecialCellsValue
    xlLogical = 4  # from enum XlSpecialCellsValue
    xlNumbers = 1  # from enum XlSpecialCellsValue
    xlTextValues = 2  # from enum XlSpecialCellsValue


class StdColorScale:
    xlColorScaleBlackWhite = 3  # from enum XlStdColorScale
    xlColorScaleGYR = 2  # from enum XlStdColorScale
    xlColorScaleRYG = 1  # from enum XlStdColorScale
    xlColorScaleWhiteBlack = 4  # from enum XlStdColorScale


class SubscribeToFormat:
    xlSubscribeToPicture = -4147  # from enum XlSubscribeToFormat
    xlSubscribeToText = -4158  # from enum XlSubscribeToFormat


class SubtotalLocationType:
    xlAtBottom = 2  # from enum XlSubtototalLocationType
    xlAtTop = 1  # from enum XlSubtototalLocationType


class SummaryColumn:
    xlSummaryOnLeft = -4131  # from enum XlSummaryColumn
    xlSummaryOnRight = -4152  # from enum XlSummaryColumn


class SummaryReportType:
    xlStandardSummary = 1  # from enum XlSummaryReportType
    xlSummaryPivotTable = -4148  # from enum XlSummaryReportType


class SummaryRow:
    xlSummaryAbove = 0  # from enum XlSummaryRow
    xlSummaryBelow = 1  # from enum XlSummaryRow


class TabPosition:
    xlTabPositionFirst = 0  # from enum XlTabPosition
    xlTabPositionLast = 1  # from enum XlTabPosition


class TableStyleElementType:
    xlBlankRow = 19  # from enum XlTableStyleElementType
    xlColumnStripe1 = 7  # from enum XlTableStyleElementType
    xlColumnStripe2 = 8  # from enum XlTableStyleElementType
    xlColumnSubheading1 = 20  # from enum XlTableStyleElementType
    xlColumnSubheading2 = 21  # from enum XlTableStyleElementType
    xlColumnSubheading3 = 22  # from enum XlTableStyleElementType
    xlFirstColumn = 3  # from enum XlTableStyleElementType
    xlFirstHeaderCell = 9  # from enum XlTableStyleElementType
    xlFirstTotalCell = 11  # from enum XlTableStyleElementType
    xlGrandTotalColumn = 4  # from enum XlTableStyleElementType
    xlGrandTotalRow = 2  # from enum XlTableStyleElementType
    xlHeaderRow = 1  # from enum XlTableStyleElementType
    xlLastColumn = 4  # from enum XlTableStyleElementType
    xlLastHeaderCell = 10  # from enum XlTableStyleElementType
    xlLastTotalCell = 12  # from enum XlTableStyleElementType
    xlPageFieldLabels = 26  # from enum XlTableStyleElementType
    xlPageFieldValues = 27  # from enum XlTableStyleElementType
    xlRowStripe1 = 5  # from enum XlTableStyleElementType
    xlRowStripe2 = 6  # from enum XlTableStyleElementType
    xlRowSubheading1 = 23  # from enum XlTableStyleElementType
    xlRowSubheading2 = 24  # from enum XlTableStyleElementType
    xlRowSubheading3 = 25  # from enum XlTableStyleElementType
    xlSlicerHoveredSelectedItemWithData = 33  # from enum XlTableStyleElementType
    xlSlicerHoveredSelectedItemWithNoData = 35  # from enum XlTableStyleElementType
    xlSlicerHoveredUnselectedItemWithData = 32  # from enum XlTableStyleElementType
    xlSlicerHoveredUnselectedItemWithNoData = 34  # from enum XlTableStyleElementType
    xlSlicerSelectedItemWithData = 30  # from enum XlTableStyleElementType
    xlSlicerSelectedItemWithNoData = 31  # from enum XlTableStyleElementType
    xlSlicerUnselectedItemWithData = 28  # from enum XlTableStyleElementType
    xlSlicerUnselectedItemWithNoData = 29  # from enum XlTableStyleElementType
    xlSubtotalColumn1 = 13  # from enum XlTableStyleElementType
    xlSubtotalColumn2 = 14  # from enum XlTableStyleElementType
    xlSubtotalColumn3 = 15  # from enum XlTableStyleElementType
    xlSubtotalRow1 = 16  # from enum XlTableStyleElementType
    xlSubtotalRow2 = 17  # from enum XlTableStyleElementType
    xlSubtotalRow3 = 18  # from enum XlTableStyleElementType
    xlTotalRow = 2  # from enum XlTableStyleElementType
    xlWholeTable = 0  # from enum XlTableStyleElementType


class TextParsingType:
    xlDelimited = 1  # from enum XlTextParsingType
    xlFixedWidth = 2  # from enum XlTextParsingType


class TextQualifier:
    xlTextQualifierDoubleQuote = 1  # from enum XlTextQualifier
    xlTextQualifierNone = -4142  # from enum XlTextQualifier
    xlTextQualifierSingleQuote = 2  # from enum XlTextQualifier


class TextVisualLayoutType:
    xlTextVisualLTR = 1  # from enum XlTextVisualLayoutType
    xlTextVisualRTL = 2  # from enum XlTextVisualLayoutType


class ThemeColor:
    xlThemeColorAccent1 = 5  # from enum XlThemeColor
    xlThemeColorAccent2 = 6  # from enum XlThemeColor
    xlThemeColorAccent3 = 7  # from enum XlThemeColor
    xlThemeColorAccent4 = 8  # from enum XlThemeColor
    xlThemeColorAccent5 = 9  # from enum XlThemeColor
    xlThemeColorAccent6 = 10  # from enum XlThemeColor
    xlThemeColorDark1 = 1  # from enum XlThemeColor
    xlThemeColorDark2 = 3  # from enum XlThemeColor
    xlThemeColorFollowedHyperlink = 12  # from enum XlThemeColor
    xlThemeColorHyperlink = 11  # from enum XlThemeColor
    xlThemeColorLight1 = 2  # from enum XlThemeColor
    xlThemeColorLight2 = 4  # from enum XlThemeColor


class ThemeFont:
    xlThemeFontMajor = 1  # from enum XlThemeFont
    xlThemeFontMinor = 2  # from enum XlThemeFont
    xlThemeFontNone = 0  # from enum XlThemeFont


class ThreadMode:
    xlThreadModeAutomatic = 0  # from enum XlThreadMode
    xlThreadModeManual = 1  # from enum XlThreadMode


class TickLabelOrientation:
    xlTickLabelOrientationAutomatic = -4105  # from enum XlTickLabelOrientation
    xlTickLabelOrientationDownward = -4170  # from enum XlTickLabelOrientation
    xlTickLabelOrientationHorizontal = -4128  # from enum XlTickLabelOrientation
    xlTickLabelOrientationUpward = -4171  # from enum XlTickLabelOrientation
    xlTickLabelOrientationVertical = -4166  # from enum XlTickLabelOrientation
    xlTickLabelPositionHigh = -4127  # from enum XlTickLabelPosition
    xlTickLabelPositionLow = -4134  # from enum XlTickLabelPosition
    xlTickLabelPositionNextToAxis = 4  # from enum XlTickLabelPosition
    xlTickLabelPositionNone = -4142  # from enum XlTickLabelPosition


class TickMark:
    xlTickMarkCross = 4  # from enum XlTickMark
    xlTickMarkInside = 2  # from enum XlTickMark
    xlTickMarkNone = -4142  # from enum XlTickMark
    xlTickMarkOutside = 3  # from enum XlTickMark


class TimePeriods:
    xlLast7Days = 2  # from enum XlTimePeriods
    xlLastMonth = 5  # from enum XlTimePeriods
    xlLastWeek = 4  # from enum XlTimePeriods
    xlNextMonth = 8  # from enum XlTimePeriods
    xlNextWeek = 7  # from enum XlTimePeriods
    xlThisMonth = 9  # from enum XlTimePeriods
    xlThisWeek = 3  # from enum XlTimePeriods
    xlToday = 0  # from enum XlTimePeriods
    xlTomorrow = 6  # from enum XlTimePeriods
    xlYesterday = 1  # from enum XlTimePeriods


class TimeUnit:
    xlDays = 0  # from enum XlTimeUnit
    xlMonths = 1  # from enum XlTimeUnit
    xlYears = 2  # from enum XlTimeUnit


class ToolbarProtection:
    xlNoButtonChanges = 1  # from enum XlToolbarProtection
    xlNoChanges = 4  # from enum XlToolbarProtection
    xlNoDockingChanges = 3  # from enum XlToolbarProtection
    xlNoShapeChanges = 2  # from enum XlToolbarProtection
    xlToolbarProtectionNone = -4143  # from enum XlToolbarProtection


class TopBottom:
    xlTop10Bottom = 0  # from enum XlTopBottom
    xlTop10Top = 1  # from enum XlTopBottom


class TotalsCalculation:
    xlTotalsCalculationAverage = 2  # from enum XlTotalsCalculation
    xlTotalsCalculationCount = 3  # from enum XlTotalsCalculation
    xlTotalsCalculationCountNums = 4  # from enum XlTotalsCalculation
    xlTotalsCalculationCustom = 9  # from enum XlTotalsCalculation
    xlTotalsCalculationMax = 6  # from enum XlTotalsCalculation
    xlTotalsCalculationMin = 5  # from enum XlTotalsCalculation
    xlTotalsCalculationNone = 0  # from enum XlTotalsCalculation
    xlTotalsCalculationStdDev = 7  # from enum XlTotalsCalculation
    xlTotalsCalculationSum = 1  # from enum XlTotalsCalculation
    xlTotalsCalculationVar = 8  # from enum XlTotalsCalculation


class TrendlineType:
    xlExponential = 5  # from enum XlTrendlineType
    xlLinear = -4132  # from enum XlTrendlineType
    xlLogarithmic = -4133  # from enum XlTrendlineType
    xlMovingAvg = 6  # from enum XlTrendlineType
    xlPolynomial = 3  # from enum XlTrendlineType
    xlPower = 4  # from enum XlTrendlineType


class UnderlineStyle:
    xlUnderlineStyleDouble = -4119  # from enum XlUnderlineStyle
    xlUnderlineStyleDoubleAccounting = 5  # from enum XlUnderlineStyle
    xlUnderlineStyleNone = -4142  # from enum XlUnderlineStyle
    xlUnderlineStyleSingle = 2  # from enum XlUnderlineStyle
    xlUnderlineStyleSingleAccounting = 4  # from enum XlUnderlineStyle


class UpdateLinks:
    xlUpdateLinksAlways = 3  # from enum XlUpdateLinks
    xlUpdateLinksNever = 2  # from enum XlUpdateLinks
    xlUpdateLinksUserSetting = 1  # from enum XlUpdateLinks


class VAlign:
    xlVAlignBottom = -4107  # from enum XlVAlign
    xlVAlignCenter = -4108  # from enum XlVAlign
    xlVAlignDistributed = -4117  # from enum XlVAlign
    xlVAlignJustify = -4130  # from enum XlVAlign
    xlVAlignTop = -4160  # from enum XlVAlign


class WBATemplate:
    xlWBATChart = -4109  # from enum XlWBATemplate
    xlWBATExcel4IntlMacroSheet = 4  # from enum XlWBATemplate
    xlWBATExcel4MacroSheet = 3  # from enum XlWBATemplate
    xlWBATWorksheet = -4167  # from enum XlWBATemplate


class WebFormatting:
    xlWebFormattingAll = 1  # from enum XlWebFormatting
    xlWebFormattingNone = 3  # from enum XlWebFormatting
    xlWebFormattingRTF = 2  # from enum XlWebFormatting


class WebSelectionType:
    xlAllTables = 2  # from enum XlWebSelectionType
    xlEntirePage = 1  # from enum XlWebSelectionType
    xlSpecifiedTables = 3  # from enum XlWebSelectionType


class WindowState:
    xlMaximized = -4137  # from enum XlWindowState
    xlMinimized = -4140  # from enum XlWindowState
    xlNormal = -4143  # from enum XlWindowState


class WindowType:
    xlChartAsWindow = 5  # from enum XlWindowType
    xlChartInPlace = 4  # from enum XlWindowType
    xlClipboard = 3  # from enum XlWindowType
    xlInfo = -4129  # from enum XlWindowType
    xlWorkbook = 1  # from enum XlWindowType


class WindowView:
    xlNormalView = 1  # from enum XlWindowView
    xlPageBreakPreview = 2  # from enum XlWindowView
    xlPageLayoutView = 3  # from enum XlWindowView


class XlmMacroType:
    xlCommand = 2  # from enum XlXLMMacroType
    xlFunction = 1  # from enum XlXLMMacroType
    xlNotXLM = 3  # from enum XlXLMMacroType


class XmlExportResult:
    xlXmlExportSuccess = 0  # from enum XlXmlExportResult
    xlXmlExportValidationFailed = 1  # from enum XlXmlExportResult
    xlXmlImportElementsTruncated = 1  # from enum XlXmlImportResult


class XmlImportResult:
    xlXmlImportSuccess = 0  # from enum XlXmlImportResult
    xlXmlImportValidationFailed = 2  # from enum XlXmlImportResult


class XmlLoadOption:
    xlXmlLoadImportToList = 2  # from enum XlXmlLoadOption
    xlXmlLoadMapXml = 3  # from enum XlXmlLoadOption
    xlXmlLoadOpenXml = 1  # from enum XlXmlLoadOption
    xlXmlLoadPromptUser = 0  # from enum XlXmlLoadOption


class YesNoGuess:
    xlGuess = 0  # from enum XlYesNoGuess
    xlNo = 2  # from enum XlYesNoGuess
    xlYes = 1  # from enum XlYesNoGuess


shape_types = [
    "auto_shape",
    "callout",
    "canvas",
    "chart",
    "comment",
    "content_app",
    "diagram",
    "embedded_ole_object",
    "form_control",
    "free_form",
    "group",
    "igx_graphic",
    "ink",
    "ink_comment",
    "line",
    "linked_ole_object",
    "linked_picture",
    "media",
    "ole_control_object",
    "picture",
    "placeholder",
    "script_anchor",
    "shape_type_mixed",
    "table",
    "text_box",
    "text_effect",
    "web_video"
]
