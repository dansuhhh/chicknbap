Option Explicit



Public blankSlideLayoutId As Integer

Public startingRow As Integer

Public currentSlideIndex As Integer

Public slideTitleDistanceFromTop As Integer

Public chartMargin As Integer

Public chartWidth As Single

Public presentation As Object



Public riceColor As Long

Public kfcColor As Long

Public containerColor As Long

Public friesColor As Long

Public carbsColor As Long

Public proteinColor As Long

Public sauceColor As Long

Public addonColor As Long



Sub SetGlobalVariables()

  blankSlideLayoutId = 12

  startingRow = 5

  slideTitleDistanceFromTop = 40

  currentSlideIndex = 1

  chartMargin = 20



  riceColor = RGB(156, 136, 86)

  kfcColor = RGB(0,0,205)

  containerColor = RGB(172,172,172)

  friesColor = RGB(252,247,94)

  carbsColor = RGB(255,127,80)

  proteinColor = RGB(255, 65, 65)

  sauceColor = RGB(91, 107, 137)

  addonColor = RGB(134, 160, 111)

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

  chartWidth = (presentation.PageSetup.SlideWidth - (chartMargin * 4)) / 5



  'Variables to track which rows tie to which slide

  Dim lastRowFound As Boolean

  Dim currentComponent, currentRowComponent, currentRowSubComponent As String

  Dim prevRowComponent, prevSubComponent, prevRowSubComponent As String

  Dim currentComponentStartRow, currentComponentEndRow As Integer

  Static currentRow As Integer

  currentRow = startingRow



  Do While lastRowFound = False

    currentRowComponent = GetCellVal("D", currentRow)

    currentRowSubComponent = GetCellVal("E", currentRow)



    If currentComponent <> "" Then

      prevRowComponent = GetCellVal("D", currentRow - 1)

      prevRowSubComponent = GetCellVal("E", currentComponentStartRow - 1)

      prevSubComponent = GetCellVal("E", currentComponentStartRow)



      If currentRowSubComponent = "" Then

        If currentComponent <> currentRowComponent Then

          currentComponentEndRow = currentRow - 1



          If GetCellVal("E", currentComponentStartRow) = "" And _

            currentComponent <> "Tortilla" And _

            currentComponent <> "Containers" And _

            currentComponent <> "Lids" Then

            Call AddIngredientTitleSlide(currentComponent, currentComponentStartRow)

            Call Increment(currentSlideIndex)

          End If



          Call AddIngredientSlide(currentComponent, currentComponentStartRow, currentComponentEndRow)

          Call Increment(currentSlideIndex)



          If currentRowComponent = "" Then

            lastRowFound = True

          Else

            currentComponentStartRow = currentRow

            currentComponent = currentRowComponent

          End If

        End If

      Else

        If currentComponent <> currentRowSubComponent Then

          currentComponentEndRow = currentRow - 1



          If prevSubComponent = "" And Lcase(prevRowComponent) <> Lcase(currentRowComponent) Then

            Call AddIngredientTitleSlide(GetCellVal("D", currentComponentStartRow), currentComponentStartRow)

            Call Increment(currentSlideIndex)

          ElseIf prevRowSubComponent = "" And Lcase(prevRowComponent) = Lcase(currentRowComponent) Then

            Call AddIngredientTitleSlide(GetCellVal("D", currentComponentStartRow), currentComponentStartRow)

            Call Increment(currentSlideIndex)

          End If



          Call AddIngredientSlide(currentComponent, currentComponentStartRow, currentComponentEndRow)

          Call Increment(currentSlideIndex)



          If currentRowComponent = "" Then

            lastRowFound = True

          Else

            currentComponentStartRow = currentRow

            currentComponent = currentRowSubComponent

          End If

        End If

      End If

    Else

      currentComponent = currentRowComponent

      currentComponentStartRow = currentRow

    End If



    Call Increment(currentRow)



  Loop

End Sub



Sub AddIngredientTitleSlide(ByVal title as String, ByVal row As Integer)

  Dim slide As Object

  Set slide = presentation.Slides.Add(currentSlideIndex, blankSlideLayoutId)

  slide.FollowMasterBackground = False



  Dim category As String

  category = GetCellVal("C", row)



  If category = "1. Rice" Then

    slide.Background.Fill.ForeColor.RGB = riceColor

  ElseIf title = "Pita Bread" Then

    slide.Background.Fill.ForeColor.RGB = carbsColor

  ElseIf title = "Waffle Fries" Then

    slide.Background.Fill.ForeColor.RGB = friesColor

  ElseIf title = "Aluminum Foil" Then

    slide.Background.Fill.ForeColor.RGB = containerColor

  ElseIf Lcase(title) = "korean fried chicken" Then

    slide.Background.Fill.ForeColor.RGB = kfcColor

  ElseIf category = "2. Protein" Then

    slide.Background.Fill.ForeColor.RGB = proteinColor

  ElseIf category = "3. Sauce" Then

    slide.Background.Fill.ForeColor.RGB = sauceColor

  Else

    slide.Background.Fill.ForeColor.RGB = addonColor

  End If



  Dim titleText As Object

  Set titleText = slide.Shapes.AddTextbox( _

    msoTextOrientationHorizontal, _

    25, _

    presentation.PageSetup.SlideHeight - 200, _

    presentation.PageSetup.SlideWidth, _

    60 _

  )



  With titleText.TextFrame

    .TextRange.ParagraphFormat.Alignment = 1

    .TextRange.Font.Size = 66

    .TextRange.Font.Color.RGB = RGB(255,255,255)

    .TextRange.Font.Name = "Roboto Condensed"

    .VerticalAnchor = msoAnchorMiddle

  End With



  If title = "Pita Bread" Then

    titleText.TextFrame.TextRange.Text = "SIDE - CARBS"

  ElseIf title = "Waffle Fries" Then

    titleText.TextFrame.TextRange.Text = "SIDE - FRIES"

    titleText.TextFrame.TextRange.Font.Color.RGB = RGB(0,0,0)

  ElseIf title = "Aluminum Foil" Then

    titleText.TextFrame.TextRange.Text = "CONTAINERS"

  ElseIf title = "Daikon Radish" Then

    titleText.TextFrame.TextRange.Text = "KOREAN FRIED CHICKEN"

  Else

    titleText.TextFrame.TextRange.Text = Ucase(title)

  End If

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

  Call AddServingsChart(slide, startRow, endRow)

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

    .TextRange.Font.Name = "Roboto Condensed"

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

    chartWidth * 2 _

  )



  Dim chartTitle As Object

  Set chartTitle = slide.Shapes.AddTextbox( _

    msoTextOrientationHorizontal, _

    chartMargin, _

    slideTitleDistanceFromTop + 90, _

    chartWidth * 2, _

    60 _

  )

  Call FormatChartTitle(chartTitle, "Amount per Recipe")



  Call FormatChartEntry(chart.Table.cell(1, 1), "INGREDIENT", False, True)

  Call FormatChartEntry(chart.Table.cell(1, 2), "AMOUNT", False, True)

  Call FormatChartEntry(chart.Table.cell(1, 3), "METHOD", False, True)



  Dim currentChartRow As Integer

  currentChartRow = 2

  Dim row As Integer

  For row = startRow To endRow

    If GetCellVal("F", row) <> "Recipe" Then

      Call FormatChartEntry(chart.Table.cell(currentChartRow, 1), GetCellVal("F", row))

      Call FormatChartEntry(chart.Table.cell(currentChartRow, 2), Round(Val(GetCellVal("H", row)), 3) & " " & Lcase(GetCellVal("I", row)))

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

    3, _

    (chartWidth * 2) + (chartMargin * 2), _

    slideTitleDistanceFromTop + 120, _

    chartWidth * 2 _

  )

  Call FormatChartEntry(chart.Table.cell(1, 1), "INGREDIENT", False, True)

  Call FormatChartEntry(chart.Table.cell(1, 2), "AMOUNT", False, True)

  Call FormatChartEntry(chart.Table.cell(1, 3), "METHOD", False, True)



  Dim currentChartRow As Integer

  currentChartRow = 2

  Dim row As Integer

  For row = startRow To endRow

    If GetCellVal("F", row) = "Recipe" Then

      Dim chartTitle As Object

      Set chartTitle = slide.Shapes.AddTextbox( _

        msoTextOrientationHorizontal, _

        (chartWidth * 2) + (chartMargin * 2), _

        slideTitleDistanceFromTop + 90, _

        chartWidth * 2, _

        60 _

      )

      Call FormatChartTitle(chartTitle, GetCellVal("K", row) & " " & GetCellVal("L",row))



    Else

      Call FormatChartEntry(chart.Table.cell(currentChartRow, 1), GetCellVal("F", row))

      Call FormatChartEntry(chart.Table.cell(currentChartRow, 2), GetCellVal("K", row) & " " & Lcase(GetCellVal("L", row)))

      Call FormatChartEntry(chart.Table.cell(currentChartRow, 3), GetCellVal("G", row))

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

    (chartWidth * 4) + (chartMargin * 3), _

    slideTitleDistanceFromTop + 120, _

    chartWidth _

  )



  Dim chartTitle As Object

  Set chartTitle = slide.Shapes.AddTextbox( _

    msoTextOrientationHorizontal, _

    (chartWidth * 4) + (chartMargin * 3), _

    slideTitleDistanceFromTop + 90, _

    chartWidth, _

    60 _

  )

  Call FormatChartTitle(chartTitle, "Quantity")



  Dim row As Integer

  For row = startRow To endRow

    If GetCellVal("F", row) = "Recipe" Then

      Call FormatChartEntry(chart.Table.cell(1, 1), "DAY", False, True)

      Call FormatChartEntry(chart.Table.cell(2, 1), "MON")

      Call FormatChartEntry(chart.Table.cell(3, 1), "TUES")

      Call FormatChartEntry(chart.Table.cell(4, 1), "WED")

      Call FormatChartEntry(chart.Table.cell(5, 1), "THURS")

      Call FormatChartEntry(chart.Table.cell(6, 1), "FRI")

      Call FormatChartEntry(chart.Table.cell(7, 1), "TOTAL", True)



      Call FormatChartEntry(chart.Table.cell(1, 2), UCase(GetCellVal("AG", row)), False, True)

      Call FormatChartEntry(chart.Table.cell(2, 2), GetCellVal("Y", row))

      Call FormatChartEntry(chart.Table.cell(3, 2), GetCellVal("Z", row))

      Call FormatChartEntry(chart.Table.cell(4, 2), GetCellVal("AA", row))

      Call FormatChartEntry(chart.Table.cell(5, 2), GetCellVal("AB", row))

      Call FormatChartEntry(chart.Table.cell(6, 2), GetCellVal("AC", row))

      Call FormatChartEntry(chart.Table.cell(7, 2), GetCellVal("AF", row), True)

    End If

  Next row

End Sub



Sub AddServingsChart( _

  ByRef slide As Object, _

  ByVal startRow As Integer, _

  ByVal endRow As Integer _

)

  Dim chart As Object

  Set chart = slide.Shapes.AddTable( _

    3, _

    2, _

    (chartWidth * 4) + (chartMargin * 3), _

    slideTitleDistanceFromTop + 360, _

    chartWidth _

  )



  Dim chartTitle As Object

  Set chartTitle = slide.Shapes.AddTextbox( _

    msoTextOrientationHorizontal, _

    (chartWidth * 4) + (chartMargin * 3), _

    slideTitleDistanceFromTop + 330, _

    chartWidth, _

    60 _

  )

  Call FormatChartTitle(chartTitle, "Servings")



  Dim af As Variant

  Dim ak As Variant

  ak = -1



  Dim row As Integer

  For row = startRow To endRow

    If GetCellVal("F", row) = "Recipe" Then

      af = Val(GetCellVal("AF", row))

      Call FormatChartEntry(chart.Table.cell(1, 1), "PER", False, True)

      Call FormatChartEntry(chart.Table.cell(1, 2), "SERVINGS", False, True)

      Call FormatChartEntry(chart.Table.cell(2, 1), GetCellVal("AG", row))

    Else

      If ak = -1 Then

        ak = Val(GetCellVal("AK", row))

      End If

    End If

  Next row



  If Round(ak, 3) =  0 Or Round(ak, 3) = 0 Then

    Call FormatChartEntry(chart.Table.cell(2, 2), 0)

  Else

    Call FormatChartEntry(chart.Table.cell(2, 2), Round(ak / af, 2))

  End If



  Call FormatChartEntry(chart.Table.cell(3, 1), "Week")

  Call FormatChartEntry(chart.Table.cell(3, 2), Round(ak, 2))

End Sub





Sub FormatChartTitle(ByRef chartTitle As Object, ByVal text As String)

  With chartTitle.TextFrame

    .TextRange.Text = text

    .TextRange.ParagraphFormat.Alignment = 2

    .TextRange.Font.Size = 20

    .TextRange.Font.Color.RGB = RGB(193, 0, 0)

    .TextRange.Font.Bold = True

    .TextRange.Font.Name = "Roboto Condensed"

    .VerticalAnchor = msoAnchorMiddle

  End With

End Sub



Sub FormatChartEntry( _

  ByRef chartCell As Object,  _

  ByVal entry As String, _

  Optional ByVal isBold As Boolean = False, _

  Optional ByVal isHeader As Boolean = False _

)

  chartCell.Shape.Fill.ForeColor.RGB = RGB(255,255,255)

  chartCell.Borders(1).ForeColor.RGB = RGB(0,0,0)

  chartCell.Borders(2).ForeColor.RGB = RGB(0,0,0)

  chartCell.Borders(3).ForeColor.RGB = RGB(0,0,0)

  chartCell.Borders(4).ForeColor.RGB = RGB(0,0,0)

  With chartCell.Shape.TextFrame

    .TextRange.Text = entry

    .TextRange.Font.Size = 12

    .TextRange.Font.Name = "Roboto Condensed"

    .TextRange.Font.Bold = isBold

  End With



  If isHeader = False Then

    chartCell.Shape.TextFrame.TextRange.Font.Color.RGB = RGB(0,0,0)

  Else

    chartCell.Shape.TextFrame.TextRange.Font.Color.RGB = RGB(193, 0, 0)

  End If

End Sub



Function GetCellVal(ByVal row As String, ByVal col As Integer) As String

  GetCellVal = Range(row & col).Value

End Function



Sub Increment(ByRef i As Integer)

  i = i + 1

End Sub