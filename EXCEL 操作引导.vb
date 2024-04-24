Sub create()

End Sub
Sub main()
'
' main 函数
'
' 快捷键: Ctrl+m
'
    ' 选择区域
    Range("A1").Select
    
    ' 选择列
    Columns("A:A").Select
    
    ' 选择行
    Rows("1:1").Select
    
    ' 插入内容
    ActiveCell.FormulaR1C1 = "1"
    '' 1. 字体
    
    ' 设置字体
    With Selection.Font
        .Name = "等线"
        .Size = 11
        .Strikethrough = False
        .Superscript = False
        .Subscript = False
        .OutlineFont = False
        .Shadow = False
        .Underline = xlUnderlineStyleNone
        .ThemeColor = xlThemeColorLight1
        .Color = -16776961
        .TintAndShade = 0
        .ThemeFont = xlThemeFontMinor
    End With
    
    ' 加粗
    Selection.Font.Bold = True
    
    ' 倾斜
    Selection.Font.Italic = True
    
    '下划线
    Selection.Font.Underline = xlUnderlineStyleSingle
    
    ' 设置底部边框
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    ' 设置底纹
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    ' 设置字体颜色
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    
    ' 设置拼音
    Selection.Phonetics.Visible = True
    
    '' 2. 对齐方式
    
    ' 顶端对齐
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlTop
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    
    ' 垂直居中
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' 底端对齐
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' 左对齐 + 底端对齐
    With Selection
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' 旋转文字
    With Selection
        .HorizontalAlignment = xlGeneral
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 45
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    ' 合并后居中
    Range("A1:D1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    
    
    '' 3. 格式
    
    ' 通用格式
    Selection.NumberFormatLocal = "G/通用格式"
    
    ' 数字格式
    Selection.NumberFormatLocal = "0.00_);[红色](0.00)"
    
    ' 货币格式
    Selection.NumberFormatLocal = "￥#,##0.00_);[红色](￥#,##0.00)"
    
    ' 文本格式
    Selection.NumberFormatLocal = "@"
    
    ' 百分比格式
    Selection.NumberFormatLocal = "0.00%"
    
    '新建条件格式
    Selection.FormatConditions.AddColorScale ColorScaleType:=2
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).ColorScaleCriteria(1).Type = _
        xlConditionValueLowestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(1).FormatColor
        .Color = 2650623
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).ColorScaleCriteria(2).Type = _
        xlConditionValueHighestValue
    With Selection.FormatConditions(1).ColorScaleCriteria(2).FormatColor
        .Color = 10285055
        .TintAndShade = 0
    End With
    
    ' 插入公式
    ActiveCell.FormulaR1C1 = "=SUM(10)"
    Range("A2").Select

    ' 升序
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Sheet1").Sort.SortFields.Add2 Key:=Range("A1"), _
        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Sheet1").Sort
        .SetRange Range("A1")
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    ' 替换
    ActiveCell.Replace What:="1", Replacement:="1", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False, FormulaVersion:=xlReplaceFormula2
    
    ' 选择空值
    Selection.SpecialCells(xlCellTypeBlanks).Select
        
    ' 清除
    Selection.Clear
    
    ' 选择公式错误单元格
    Selection.SpecialCells(xlCellTypeFormulas, 16).Select
    
    ' 选择公式数字单元格
    Selection.SpecialCells(xlCellTypeFormulas, 1).Select
    
    ' 粘贴
    ActiveSheet.Paste
    
    ' 剪切
    Selection.Cut
    
    ' 插入柱状图
    ActiveSheet.Shapes.AddChart2(201, xlColumnClustered).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$I$18:$J$19")
    

    ' 插入饼图
    ActiveSheet.Shapes.AddChart2(251, xlPie).Select
    ActiveChart.SetSourceData Source:=Range("Sheet1!$I$18:$J$19")
    
    ' 删除图
    ActiveChart.Parent.Delete
    
    ' 自动筛选
    Selection.AutoFilter
    
    '打印
    ActiveWindow.SelectedSheets.PrintOut Copies:=1, Collate:=True, _
        IgnorePrintAreas:=False
        
    ' 冻结窗格
    ActiveWindow.SmallScroll Down:=-21
    Range("B2").Select
    ActiveWindow.FreezePanes = True
    
    ' 解除冻结窗格
    ActiveWindow.FreezePanes = False
    
    ' 页面设置
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .PrintTitleRows = ""
        .PrintTitleColumns = ""
    End With
    Application.PrintCommunication = True
    ActiveSheet.PageSetup.PrintArea = ""
    Application.PrintCommunication = False
    With ActiveSheet.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .LeftMargin = Application.InchesToPoints(0.708661417322835)
        .RightMargin = Application.InchesToPoints(0.708661417322835)
        .TopMargin = Application.InchesToPoints(0.748031496062992)
        .BottomMargin = Application.InchesToPoints(0.748031496062992)
        .HeaderMargin = Application.InchesToPoints(0.31496062992126)
        .FooterMargin = Application.InchesToPoints(0.31496062992126)
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .PrintQuality = 1200
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlPortrait
        .Draft = False
        .PaperSize = xlPaperA4
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 100
        .PrintErrors = xlPrintErrorsDisplayed
        .OddAndEvenPagesHeaderFooter = False
        .DifferentFirstPageHeaderFooter = False
        .ScaleWithDocHeaderFooter = True
        .AlignMarginsHeaderFooter = True
        .EvenPage.LeftHeader.Text = ""
        .EvenPage.CenterHeader.Text = ""
        .EvenPage.RightHeader.Text = ""
        .EvenPage.LeftFooter.Text = ""
        .EvenPage.CenterFooter.Text = ""
        .EvenPage.RightFooter.Text = ""
        .FirstPage.LeftHeader.Text = ""
        .FirstPage.CenterHeader.Text = ""
        .FirstPage.RightHeader.Text = ""
        .FirstPage.LeftFooter.Text = ""
        .FirstPage.CenterFooter.Text = ""
        .FirstPage.RightFooter.Text = ""
    End With
    Application.PrintCommunication = True
    
    ' 保存
    ChDir "D:\Users\Desktop"
    ActiveWorkbook.SaveAs Filename:="D:\Users\Desktop\工作簿1.xlsx", FileFormat:= _
        xlOpenXMLWorkbook, CreateBackup:=False
End Sub
