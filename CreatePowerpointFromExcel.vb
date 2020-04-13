Option Explicit

Public blankSlideLayoutId As Integer
Public startingRow As Integer
Public currentSlideIndex As Integer
Public slideTitleDistanceFromTop As Integer
Public chartMargin As Integer
Public chartWidth As Single
Public presentation As Object

Sub SetGlobalVariables()
  blankSlideLayoutId = 12
  startingRow = 5
  slideTitleDistanceFromTop = 40
  currentSlideIndex = 1
End Sub

' Main function
Sub CreatePowerpointFromExcel()
  Call SetGlobalVariables()

  ' Start PowerPoint
  Dim app as Object
  Set app = CreateObject("PowerPoint.Application")
  app.Activate
  app.Visible = True
  
  ' Create presentation
  Set presentation = app.Presentations.Add
  chartMargin = presentation.PageSetup.SlideWidth / 40
  chartWidth = presentation.PageSetup.SlideWidth / 40 * 12
  
  'Variables to track which rows tie to which slide
  Dim firstComponentSet, lastRowFound As Boolean
  Dim currentComponent, currentRowComponent As String
  Dim currentComponentStartRow, currentComponentEndRow As Integer
  Static currentRow As Integer
  currentRow = startingRow
  
  Do While lastRowFound = False
    currentRowComponent = GetCellVal("D", currentRow)

    If currentComponent = "" And firstComponentSet = False Then
      firstComponentSet = True
      currentComponent = currentRowComponent
      currentComponentStartRow = currentRow

    ElseIf currentRowComponent <> currentComponent Then
      currentComponentEndRow = currentRow - 1
      Call AddIngredientSlide(currentComponent, currentComponentStartRow, currentComponentEndRow)
      Call Increment(currentSlideIndex)

      ' temporary
      If currentRowComponent = "" Or LCase(currentRowComponent) = "korean fried chicken" Then
        lastRowFound = True
      Else
        currentComponentStartRow = currentRow
        currentComponent = currentRowComponent
      End If
    End If

    Call Increment(currentRow)
  Loop 
End Sub

Sub AddIngredientSlide( _
  ByVal component as String, _
  ByVal startRow As Integer, _
  ByVal endRow As Integer _
) 
  Dim slide As Object
  Set slide = presentation.Slides.Add(currentSlideIndex, blankSlideLayoutId)

  Call AddTitleText(slide, component)
  Call AddAmtPerRecipeChart(slide, startRow, endRow)
  Call AddRecipePerGroupChart(slide, startRow, endRow)
  Call AddDailyParChart(slide, startRow, endRow)
End Sub

Sub AddTitleText(ByRef slide As Object, ByVal title As String)
  Dim titleText As Object
  Set titleText = slide.Shapes.AddTextbox( _
    msoTextOrientationHorizontal, _
    0, _
    slideTitleDistanceFromTop, _
    presentation.PageSetup.SlideWidth, _
    60 _
  )
  With titleText.TextFrame
    .TextRange.Text = title
    .TextRange.ParagraphFormat.Alignment = 2
    .TextRange.Font.Size = 40
    .TextRange.Font.Name = "Tungsten Book"
    .VerticalAnchor = msoAnchorMiddle
  End With
End Sub

Sub AddAmtPerRecipeChart( _
  ByRef slide As Object, _
  ByVal startRow As Integer, _
  ByVal endRow As Integer _
)
  Dim chart As Object
  Set chart = slide.Shapes.AddTable( _
    endRow - startRow + 1, _
    3, _
    chartMargin, _
    slideTitleDistanceFromTop + 120, _
    chartWidth _
  )

  Dim chartTitle As Object
  Set chartTitle = slide.Shapes.AddTextbox( _
    msoTextOrientationHorizontal, _
    chartMargin, _
    slideTitleDistanceFromTop + 90, _
    chartWidth, _
    60 _
  )
  Call FormatChartTitle(chartTitle, "Amount per Recipe")

  Call FormatChartEntry(chart.Table.cell(1, 1), "INGREDIENT")
  Call FormatChartEntry(chart.Table.cell(1, 2), "AMOUNT")
  Call FormatChartEntry(chart.Table.cell(1, 3), "METHOD")

  Dim currentChartRow As Integer
  currentChartRow = 2
  Dim row As Integer
  For row = startRow To endRow
    If GetCellVal("F", row) <> "Recipe" Then
      Call FormatChartEntry(chart.Table.cell(currentChartRow, 1), GetCellVal("F", row))
      Call FormatChartEntry(chart.Table.cell(currentChartRow, 2), GetCellVal("H", row) & Lcase(GetCellVal("I", row)))
      Call FormatChartEntry(chart.Table.cell(currentChartRow, 3), GetCellVal("G", row))
      Call Increment(currentChartRow)
    End If
  Next row
End Sub

Sub AddRecipePerGroupChart( _
  ByRef slide As Object, _
  ByVal startRow As Integer, _
  ByVal endRow As Integer _
)
  Dim chart As Object
  Set chart = slide.Shapes.AddTable( _
    endRow - startRow + 1, _
    2, _
    chartWidth + (chartMargin * 2), _
    slideTitleDistanceFromTop + 120, _
    chartWidth _
  )
  Call FormatChartEntry(chart.Table.cell(1, 1), "INGREDIENT")
  Call FormatChartEntry(chart.Table.cell(1, 2), "AMOUNT")

  Dim currentChartRow As Integer
  currentChartRow = 2
  Dim row As Integer
  For row = startRow To endRow
    If GetCellVal("F", row) = "Recipe" Then
      Dim chartTitle As Object
      Set chartTitle = slide.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        chartWidth + (chartMargin * 2), _
        slideTitleDistanceFromTop + 90, _
        chartWidth, _
        60 _
      )
      Call FormatChartTitle(chartTitle, GetCellVal("K", row) & " " & GetCellVal("L",row))

    Else
      Call FormatChartEntry(chart.Table.cell(currentChartRow, 1), GetCellVal("F", row))
      Call FormatChartEntry(chart.Table.cell(currentChartRow, 2), GetCellVal("K", row) & Lcase(GetCellVal("L", row)))
      Call Increment(currentChartRow)
    End If
  Next row
End Sub

Sub AddDailyParChart( _
  ByRef slide As Object, _
  ByVal startRow As Integer, _
  ByVal endRow As Integer _
)
  Dim chart As Object
  Set chart = slide.Shapes.AddTable( _
    7, _
    2, _
    (chartWidth * 2) + (chartMargin * 3), _
    slideTitleDistanceFromTop + 120, _
    chartWidth _
  )

  Dim chartTitle As Object
  Set chartTitle = slide.Shapes.AddTextbox( _
    msoTextOrientationHorizontal, _
    (chartWidth * 2) + (chartMargin * 3), _
    slideTitleDistanceFromTop + 90, _
    chartWidth, _
    60 _
  )
  Call FormatChartTitle(chartTitle, "Quantity")

  Dim row As Integer
  For row = startRow To endRow
    If GetCellVal("F", row) = "Recipe" Then
      Call FormatChartEntry(chart.Table.cell(1, 1), "DAY")
      Call FormatChartEntry(chart.Table.cell(2, 1), "MON")
      Call FormatChartEntry(chart.Table.cell(3, 1), "TUES")
      Call FormatChartEntry(chart.Table.cell(4, 1), "WED")
      Call FormatChartEntry(chart.Table.cell(5, 1), "THURS")
      Call FormatChartEntry(chart.Table.cell(6, 1), "FRI")
      Call FormatChartEntry(chart.Table.cell(7, 1), "TOTAL", True)

      Call FormatChartEntry(chart.Table.cell(1, 2), GetCellVal("AG", row))
      Call FormatChartEntry(chart.Table.cell(2, 2), GetCellVal("Y", row))
      Call FormatChartEntry(chart.Table.cell(3, 2), GetCellVal("Z", row))
      Call FormatChartEntry(chart.Table.cell(4, 2), GetCellVal("AA", row))
      Call FormatChartEntry(chart.Table.cell(5, 2), GetCellVal("AB", row))
      Call FormatChartEntry(chart.Table.cell(6, 2), GetCellVal("AC", row))
      Call FormatChartEntry(chart.Table.cell(7, 2), GetCellVal("AF", row), True)
    End If
  Next row
End Sub

Sub FormatChartTitle(ByRef chartTitle As Object, ByVal text As String)
  With chartTitle.TextFrame
    .TextRange.Text = text
    .TextRange.ParagraphFormat.Alignment = 2
    .TextRange.Font.Size = 20
    .TextRange.Font.Bold = True
    .TextRange.Font.Underline = True
    .TextRange.Font.Name = "Roboto Condensed"
    .VerticalAnchor = msoAnchorMiddle
  End With
End Sub

Sub FormatChartEntry( _
  ByRef chartCell As Object,  _
  ByVal entry As String, _
  Optional ByVal isBold As Boolean = false _
)
  With chartCell.Shape.TextFrame
    .TextRange.Text = entry
    .TextRange.Font.Size = 14
    .TextRange.Font.Bold = isBold
  End With
End Sub

Function GetCellVal(ByVal row As String, ByVal col As Integer) As String
  GetCellVal = Range(row & col).Value
End Function

Sub Increment(ByRef i As Integer)
  i = i + 1
End Sub
