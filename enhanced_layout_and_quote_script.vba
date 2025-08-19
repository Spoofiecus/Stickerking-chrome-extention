' VBA Script for CorelDRAW to create sticker layouts and generate quotes.
'
' This macro arranges selected shapes in a serpentine layout, automatically
' calculates the layout, fills the last row, prices the job, and places
' a quote on the page.

' Forward declaration for the Point type if needed by other modules
Private Type Point
    X As Double
    Y As Double
End Type

' Main subroutine to be called by the user
Sub CreateLayoutAndQuote()
    ' --- Pricing and Layout Constants ---
    Const VINYL_COST_PER_M2 As Double = 460.0
    Const VAT_RATE As Double = 0.15 ' 15%
    Const ROLL_WIDTH As Double = 650 ' mm
    Const BLEED As Double = 1 ' mm
    Const MIN_PRICE_PER_STICKER As Double = 0.2
    Const MIN_ORDER_AMOUNT As Double = 100.0
    ' --- End Constants ---

    ' Set the document units to millimeters for consistency
    ActiveDocument.Unit = cdrMillimeter

    ' Check for an active selection
    If ActiveDocument Is Nothing Or ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Please select at least one shape to serve as the sticker template.", vbExclamation, "No Selection"
        Exit Sub
    End If

    If ActiveSelection.Shapes.Count > 1 Then
        MsgBox "More than one shape is selected. Only the first shape will be used as the template.", vbInformation, "Multiple Shapes Selected"
    End If

    ' Prompt for the initial total number of stickers required
    Dim originalStickers As Long
    On Error Resume Next
    originalStickers = CLng(InputBox("Enter the approximate number of stickers required:", "Sticker Quantity", 10))
    On Error GoTo 0
    If originalStickers <= 0 Then
        MsgBox "Invalid input. Please enter a positive number.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Get sticker and page dimensions
    Dim baseShape As Shape
    Set baseShape = ActiveSelection.Shapes(1)
    Dim stickerWidth As Double, stickerHeight As Double
    stickerWidth = baseShape.SizeWidth
    stickerHeight = baseShape.SizeHeight

    Dim pageWidth As Double, pageHeight As Double
    pageWidth = ActivePage.SizeWidth
    pageHeight = ActivePage.SizeHeight

    ' Automatically calculate stickers per row
    Dim stickersPerRow As Long
    If stickerWidth > 0 Then
        stickersPerRow = Int(pageWidth / stickerWidth)
    Else
        stickersPerRow = 0
    End If

    If stickersPerRow <= 0 Then
        MsgBox "The selected sticker is wider than the page.", vbExclamation, "Sticker Too Wide"
        Exit Sub
    End If

    ' --- Adjust quantity to fill the last row ---
    Dim numRows As Long
    numRows = (originalStickers + stickersPerRow - 1) \ stickersPerRow

    Dim totalStickers As Long
    totalStickers = numRows * stickersPerRow

    If originalStickers <> totalStickers Then
        MsgBox "Original quantity was " & originalStickers & ". Adjusted to " & totalStickers & " to fill the final row.", vbInformation, "Quantity Adjusted"
    End If
    ' --- End adjustment ---

    ' Prompt for vertical spacing
    Dim spacingY As Double
    On Error Resume Next
    spacingY = CDbl(InputBox("Enter the spacing between rows (in mm):", "Vertical Spacing", 0.5))
    On Error GoTo 0
    If spacingY < 0 Then
        MsgBox "Invalid input for spacing.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Check for vertical page overflow
    Dim totalLayoutHeight As Double
    totalLayoutHeight = (numRows * stickerHeight) + ((numRows - 1) * spacingY)
    If totalLayoutHeight > pageHeight Then
        If MsgBox("Warning: The layout may exceed the page height. Continue anyway?", vbYesNo + vbExclamation, "Layout May Not Fit") = vbNo Then
            Exit Sub
        End If
    End If

    ' --- Generate Quote ---
    Dim pricePerSticker As Double
    pricePerSticker = CalculatePrice(stickerWidth, stickerHeight, VINYL_COST_PER_M2, ROLL_WIDTH, BLEED, MIN_PRICE_PER_STICKER)

    Dim totalCostExclVat As Double
    totalCostExclVat = pricePerSticker * totalStickers

    Dim totalCostInclVat As Double
    totalCostInclVat = totalCostExclVat * (1 + VAT_RATE)

    Dim quoteText As String
    quoteText = "Quote Summary" & vbCrLf
    quoteText = quoteText & "----------------------------------" & vbCrLf
    quoteText = quoteText & "Sticker Dimensions: " & Format(stickerWidth, "0.00") & "mm x " & Format(stickerHeight, "0.00") & "mm" & vbCrLf
    quoteText = quoteText & "Adjusted Quantity: " & totalStickers & " stickers" & vbCrLf
    quoteText = quoteText & "Layout: " & numRows & " rows of " & stickersPerRow & " stickers" & vbCrLf
    quoteText = quoteText & "----------------------------------" & vbCrLf
    quoteText = quoteText & "Price per Sticker (excl. VAT): R " & Format(pricePerSticker, "0.00") & vbCrLf
    quoteText = quoteText & "Total (excl. VAT): R " & Format(totalCostExclVat, "0.00") & vbCrLf
    quoteText = quoteText & "Total (incl. VAT): R " & Format(totalCostInclVat, "0.00") & vbCrLf
    quoteText = quoteText & "----------------------------------" & vbCrLf

    If totalCostExclVat < MIN_ORDER_AMOUNT Then
        quoteText = quoteText & "NOTE: Order is below the minimum of R " & Format(MIN_ORDER_AMOUNT, "0.00") & "." & vbCrLf
    End If
    ' --- End Quote Generation ---

    ' --- Perform Layout ---
    Dim spacingX As Double
    If stickersPerRow > 1 Then
        spacingX = (pageWidth - (stickersPerRow * stickerWidth)) / (stickersPerRow - 1)
    Else
        spacingX = 0
    End If

    Dim positions() As Point
    ReDim positions(totalStickers - 1)

    Dim i As Long
    For i = 0 To totalStickers - 1
        Dim rowCounter As Long, colCounter As Long
        rowCounter = i \ stickersPerRow
        colCounter = i Mod stickersPerRow

        Dim currentX As Double, currentY As Double
        currentY = ActivePage.TopY - rowCounter * (stickerHeight + spacingY)

        If (rowCounter Mod 2) = 0 Then
            currentX = ActivePage.LeftX + colCounter * (stickerWidth + spacingX)
        Else
            currentX = ActivePage.LeftX + (stickersPerRow - 1 - colCounter) * (stickerWidth + spacingX)
        End If

        positions(i).X = currentX
        positions(i).Y = currentY
    Next i

    Dim duplicateShape As Shape
    For i = totalStickers - 1 To 1 Step -1
        Set duplicateShape = baseShape.Duplicate
        duplicateShape.SetPosition positions(i).X, positions(i).Y
    Next i
    baseShape.SetPosition positions(0).X, positions(0).Y
    ' --- End Layout ---

    ' --- Add Quote Text to Page ---
    Dim quoteBox As Shape
    Dim quoteX As Double, quoteY As Double
    quoteX = baseShape.PositionX + pageWidth + 10 ' Position to the right of the page
    quoteY = ActivePage.TopY

    Set quoteBox = ActiveLayer.CreateParagraphText(quoteX, quoteY, quoteX + 100, quoteY - 100, quoteText)
    If Not quoteBox Is Nothing Then
        quoteBox.Paragraph.Font = "Arial"
        quoteBox.Paragraph.Size = 10
    End If
    ' --- End Add Quote ---

    MsgBox "Layout and quote created successfully!", vbInformation, "Success"
End Sub

Private Function CalculatePrice(ByVal W As Double, ByVal H As Double, ByVal costPerM2 As Double, ByVal rollWidthMM As Double, ByVal bleedMM As Double, ByVal minPrice As Double) As Double
    ' Calculates the most cost-effective price per sticker based on orientation.

    Dim P_horizontal As Double, P_vertical As Double

    ' Horizontal Orientation
    Dim W_bleed_h As Double, S_rounded_h As Long
    W_bleed_h = W + bleedMM
    If W_bleed_h > 0 Then
        S_rounded_h = Int(rollWidthMM / W_bleed_h)
    Else
        S_rounded_h = 0
    End If

    If S_rounded_h > 0 Then
        Dim H_meters_h As Double, Area_h As Double, Row_Cost_h As Double
        H_meters_h = H / 1000
        Area_h = (rollWidthMM / 1000) * H_meters_h ' Area of one row in m^2
        Row_Cost_h = Area_h * costPerM2
        P_horizontal = Row_Cost_h / S_rounded_h
    Else
        P_horizontal = 999999 ' Represents infinity
    End If

    ' Vertical Orientation
    Dim H_bleed_v As Double, S_rounded_v As Long
    H_bleed_v = H + bleedMM
    If H_bleed_v > 0 Then
        S_rounded_v = Int(rollWidthMM / H_bleed_v)
    Else
        S_rounded_v = 0
    End If

    If S_rounded_v > 0 Then
        Dim W_meters_v As Double, Area_v As Double, Row_Cost_v As Double
        W_meters_v = W / 1000
        Area_v = (rollWidthMM / 1000) * W_meters_v ' Area of one row in m^2
        Row_Cost_v = Area_v * costPerM2
        P_vertical = Row_Cost_v / S_rounded_v
    Else
        P_vertical = 999999 ' Represents infinity
    End If

    ' Determine the minimum price between the two orientations
    Dim price As Double
    If P_horizontal < P_vertical Then
        price = P_horizontal
    Else
        price = P_vertical
    End If

    ' Return the greater of the calculated price or the minimum allowed price
    If price > minPrice Then
        CalculatePrice = price
    Else
        CalculatePrice = minPrice
    End If
End Function
