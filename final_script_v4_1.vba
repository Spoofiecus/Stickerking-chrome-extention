' VBA Script for CorelDRAW to create sticker layouts and generate quotes.
' Version 4.1: Fixes layout bug by using Printable Area.
'
' This module contains the main logic for the layout and quoting tool.
' It should be placed in a standard Module in the VBA Editor.

' AppName is used as the root key for saving settings in the Windows Registry.
Private Const AppName As String = "StickerKingVBAScript"

Private Type Point
    X As Double
    Y As Double
End Type

' --- Public Subroutines (to be run by the user) ---

Public Sub ShowSettingsPanel()
    ' This sub opens the settings panel. The user can run this macro
    ' to view or change the saved pricing settings.
    frmSettings.Show
End Sub

Public Sub CreateLayoutAndQuote()
    ' This is the main macro that performs the layout and quote generation.

    ' Declare variables for settings that will be loaded.
    Dim vinylCost As Double, vatRate As Double, rollWidth As Double
    Dim bleed As Double, minStickerPrice As Double, minOrder As Double

    ' Load settings from registry or use defaults.
    LoadSettings vinylCost, vatRate, rollWidth, bleed, minStickerPrice, minOrder

    ActiveDocument.Unit = cdrMillimeter

    ' Check for a selected shape.
    If ActiveDocument Is Nothing Or ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Please select a sticker template shape.", vbExclamation, "No Selection"
        Exit Sub
    End If

    Dim baseShape As Shape
    Set baseShape = ActiveSelection.Shapes(1)

    ' Get user input for quantity.
    Dim originalStickers As Long
    On Error Resume Next
    originalStickers = CLng(InputBox("Enter approximate sticker quantity:", "Sticker Quantity", 10))
    On Error GoTo 0
    If originalStickers <= 0 Then
        MsgBox "Invalid quantity. Please enter a positive number.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Get dimensions from the selected shape and the active page's printable area.
    Dim stickerWidth As Double, stickerHeight As Double
    stickerWidth = baseShape.SizeWidth
    stickerHeight = baseShape.SizeHeight

    Dim pageLeft As Double, pageTop As Double, pageWidth As Double, pageHeight As Double
    pageLeft = ActivePage.PrintableArea.Left
    pageTop = ActivePage.PrintableArea.Top
    pageWidth = ActivePage.PrintableArea.Width
    pageHeight = ActivePage.PrintableArea.Height

    ' --- Best Fit Logic: Determine the most efficient orientation ---
    Dim rotated As Boolean
    Dim effectiveWidth As Double, effectiveHeight As Double
    Dim stickers_as_is As Long, stickers_rotated As Long

    If stickerWidth > 0 Then stickers_as_is = Int(pageWidth / stickerWidth) Else stickers_as_is = 0
    If stickerHeight > 0 Then stickers_rotated = Int(pageWidth / stickerHeight) Else stickers_rotated = 0

    ' Check if rotating is better AND if the rotated sticker will fit on the page.
    If stickers_rotated > stickers_as_is And stickerWidth <= pageHeight Then
        rotated = True
        effectiveWidth = stickerHeight
        effectiveHeight = stickerWidth
    Else
        rotated = False
        effectiveWidth = stickerWidth
        effectiveHeight = stickerHeight
    End If
    ' --- End Best Fit ---

    Dim stickersPerRow As Long
    stickersPerRow = Int(pageWidth / effectiveWidth)

    If stickersPerRow <= 0 Then
        MsgBox "The selected sticker is wider than the printable area, even when rotated.", vbExclamation, "Sticker Too Wide"
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
    ' --- End Adjustment ---

    ' Get vertical spacing from user.
    Dim spacingY As Double
    On Error Resume Next
    spacingY = CDbl(InputBox("Enter vertical spacing between rows (mm):", "Vertical Spacing", 0.5))
    On Error GoTo 0
    If spacingY < 0 Then MsgBox "Invalid spacing.", vbExclamation, "Invalid Input": Exit Sub

    ' Check for vertical page overflow.
    Dim totalLayoutHeight As Double
    totalLayoutHeight = (numRows * effectiveHeight) + ((numRows - 1) * spacingY)
    If totalLayoutHeight > pageHeight Then
        If MsgBox("Warning: The layout is projected to exceed the printable height. Continue anyway?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    ' --- Generate Quote ---
    Dim pricePerSticker As Double
    pricePerSticker = CalculatePrice(stickerWidth, stickerHeight, vinylCost, rollWidth, bleed, minStickerPrice)
    Dim totalCostExclVat As Double
    totalCostExclVat = pricePerSticker * totalStickers
    Dim totalCostInclVat As Double
    totalCostInclVat = totalCostExclVat * (1 + vatRate)

    Dim quoteText As String
    quoteText = "Quote Summary" & vbCrLf & "----------------------------------" & vbCrLf
    quoteText = quoteText & "Sticker Dimensions: " & Format(stickerWidth, "0.00") & "mm x " & Format(stickerHeight, "0.00") & "mm" & vbCrLf
    If rotated Then quoteText = quoteText & "Orientation: Rotated for Best Fit" & vbCrLf
    quoteText = quoteText & "Adjusted Quantity: " & totalStickers & " stickers" & vbCrLf
    quoteText = quoteText & "Layout: " & numRows & " rows of " & stickersPerRow & " stickers" & vbCrLf
    quoteText = quoteText & "----------------------------------" & vbCrLf
    quoteText = quoteText & "Price per Sticker (excl. VAT): R " & Format(pricePerSticker, "0.00") & vbCrLf
    quoteText = quoteText & "Total (excl. VAT): R " & Format(totalCostExclVat, "0.00") & vbCrLf
    quoteText = quoteText & "Total (incl. VAT): R " & Format(totalCostInclVat, "0.00") & vbCrLf
    quoteText = quoteText & "----------------------------------" & vbCrLf
    If totalCostExclVat < minOrder Then quoteText = quoteText & "NOTE: Order is below the minimum of R " & Format(minOrder, "0.00") & "." & vbCrLf
    ' --- End Quote ---

    ' --- Perform Layout ---
    Dim spacingX As Double
    If stickersPerRow > 1 Then spacingX = (pageWidth - (stickersPerRow * effectiveWidth)) / (stickersPerRow - 1) Else spacingX = 0

    Dim positions() As Point
    ReDim positions(totalStickers - 1)
    Dim i As Long, rowCounter As Long, colCounter As Long
    Dim currentX As Double, currentY As Double

    For i = 0 To totalStickers - 1
        rowCounter = i \ stickersPerRow
        colCounter = i Mod stickersPerRow
        currentY = pageTop - rowCounter * (effectiveHeight + spacingY)
        If (rowCounter Mod 2) = 0 Then
            currentX = pageLeft + colCounter * (effectiveWidth + spacingX)
        Else
            currentX = pageLeft + (stickersPerRow - 1 - colCounter) * (effectiveWidth + spacingX)
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
    quoteX = pageLeft + pageWidth + 10 ' Position to the right of the printable area
    quoteY = pageTop
    Set quoteBox = ActiveLayer.CreateParagraphText(quoteX, quoteY, quoteX + 100, quoteY - 100, quoteText)
    If Not quoteBox Is Nothing Then
        quoteBox.Paragraph.Font = "Arial"
        quoteBox.Paragraph.Size = 10
    End If
    ' --- End Add Quote ---

    MsgBox "Layout and quote created successfully!", vbInformation, "Success"
End Sub

' --- Settings Management (Private and Public) ---

Private Sub LoadSettings(ByRef vinylCost As Double, ByRef vatRate As Double, ByRef rollWidth As Double, ByRef bleed As Double, ByRef minStickerPrice As Double, ByRef minOrder As Double)
    On Error Resume Next
    vinylCost = CDbl(GetSetting(AppName, "Pricing", "VinylCost", "460.0"))
    vatRate = CDbl(GetSetting(AppName, "Pricing", "VatRate", "0.15"))
    rollWidth = CDbl(GetSetting(AppName, "Pricing", "RollWidth", "650.0"))
    bleed = CDbl(GetSetting(AppName, "Pricing", "Bleed", "1.0"))
    minStickerPrice = CDbl(GetSetting(AppName, "Pricing", "MinStickerPrice", "0.2"))
    minOrder = CDbl(GetSetting(AppName, "Pricing", "MinOrderAmount", "100.0"))
    On Error GoTo 0
End Sub

Public Sub SaveSettings(ByVal vinylCost As String, ByVal vatRate As String, ByVal rollWidth As String, ByVal bleed As String, ByVal minStickerPrice As String, ByVal minOrder As String)
    On Error Resume Next
    SaveSetting AppName, "Pricing", "VinylCost", vinylCost
    SaveSetting AppName, "Pricing", "VatRate", vatRate
    SaveSetting AppName, "Pricing", "RollWidth", rollWidth
    SaveSetting AppName, "Pricing", "Bleed", bleed
    SaveSetting AppName, "Pricing", "MinStickerPrice", minStickerPrice
    SaveSetting AppName, "Pricing", "MinOrderAmount", minOrder
    On Error GoTo 0
End Sub

' --- Helper Functions (Private) ---

Private Function CalculatePrice(ByVal W As Double, ByVal H As Double, ByVal costPerM2 As Double, ByVal rollWidthMM As Double, ByVal bleedMM As Double, ByVal minPrice As Double) As Double
    Dim P_horizontal As Double, P_vertical As Double
    Dim W_bleed_h As Double, S_rounded_h As Long, H_bleed_v As Double, S_rounded_v As Long

    ' Horizontal Orientation
    W_bleed_h = W + bleedMM
    If W_bleed_h > 0 Then S_rounded_h = Int(rollWidthMM / W_bleed_h) Else S_rounded_h = 0
    If S_rounded_h > 0 Then
        P_horizontal = (costPerM2 * (H / 1000) * (rollWidthMM / 1000)) / S_rounded_h
    Else
        P_horizontal = 999999
    End If

    ' Vertical Orientation
    H_bleed_v = H + bleedMM
    If H_bleed_v > 0 Then S_rounded_v = Int(rollWidthMM / H_bleed_v) Else S_rounded_v = 0
    If S_rounded_v > 0 Then
        P_vertical = (costPerM2 * (W / 1000) * (rollWidthMM / 1000)) / S_rounded_v
    Else
        P_vertical = 999999
    End If

    Dim price As Double
    If P_horizontal < P_vertical Then price = P_horizontal Else price = P_vertical
    If price > minPrice Then CalculatePrice = price Else CalculatePrice = minPrice
End Function
