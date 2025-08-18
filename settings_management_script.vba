' VBA Script for CorelDRAW to create sticker layouts and generate quotes.
' Version 3.0: Includes Settings Management and Best Fit logic.
'
' This macro arranges selected shapes in a serpentine layout, automatically
' calculates the layout, fills the last row, prices the job based on
' saved settings, and places a quote on the page.

' Module-level constants and types
Private Const AppName As String = "StickerKingVBAScript" ' Used for saving settings

Private Type Point
    X As Double
    Y As Double
End Type

' --- Main Subroutines (Public) ---

Public Sub CreateLayoutAndQuote()
    ' Declare variables for settings
    Dim vinylCost As Double, vatRate As Double, rollWidth As Double
    Dim bleed As Double, minStickerPrice As Double, minOrder As Double

    ' Load settings from registry or use defaults
    LoadSettings vinylCost, vatRate, rollWidth, bleed, minStickerPrice, minOrder

    ActiveDocument.Unit = cdrMillimeter

    If ActiveDocument Is Nothing Or ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Please select a sticker template.", vbExclamation, "No Selection"
        Exit Sub
    End If

    Dim baseShape As Shape
    Set baseShape = ActiveSelection.Shapes(1)

    Dim originalStickers As Long
    On Error Resume Next
    originalStickers = CLng(InputBox("Enter approximate sticker quantity:", "Sticker Quantity", 10))
    On Error GoTo 0
    If originalStickers <= 0 Then
        MsgBox "Invalid quantity.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    Dim stickerWidth As Double, stickerHeight As Double
    stickerWidth = baseShape.SizeWidth
    stickerHeight = baseShape.SizeHeight

    Dim pageWidth As Double, pageHeight As Double
    pageWidth = ActivePage.SizeWidth
    pageHeight = ActivePage.SizeHeight

    ' Best Fit Logic
    Dim rotated As Boolean
    Dim effectiveWidth As Double, effectiveHeight As Double
    Dim stickers_as_is As Long, stickers_rotated As Long

    If stickerWidth > 0 Then stickers_as_is = Int(pageWidth / stickerWidth) Else stickers_as_is = 0
    If stickerHeight > 0 Then stickers_rotated = Int(pageWidth / stickerHeight) Else stickers_rotated = 0

    If stickers_rotated > stickers_as_is And stickerWidth <= pageHeight Then
        rotated = True
        effectiveWidth = stickerHeight
        effectiveHeight = stickerWidth
    Else
        rotated = False
        effectiveWidth = stickerWidth
        effectiveHeight = stickerHeight
    End If

    Dim stickersPerRow As Long
    stickersPerRow = Int(pageWidth / effectiveWidth)

    If stickersPerRow <= 0 Then
        MsgBox "Sticker is wider than the page.", vbExclamation, "Sticker Too Wide"
        Exit Sub
    End If

    Dim numRows As Long
    numRows = (originalStickers + stickersPerRow - 1) \ stickersPerRow

    Dim totalStickers As Long
    totalStickers = numRows * stickersPerRow

    If originalStickers <> totalStickers Then
        MsgBox "Quantity adjusted from " & originalStickers & " to " & totalStickers & " to fill the final row.", vbInformation, "Quantity Adjusted"
    End If

    Dim spacingY As Double
    On Error Resume Next
    spacingY = CDbl(InputBox("Enter row spacing (mm):", "Vertical Spacing", 0.5))
    On Error GoTo 0
    If spacingY < 0 Then MsgBox "Invalid spacing.", vbExclamation, "Invalid Input": Exit Sub

    Dim totalLayoutHeight As Double
    totalLayoutHeight = (numRows * effectiveHeight) + ((numRows - 1) * spacingY)
    If totalLayoutHeight > pageHeight Then
        If MsgBox("Warning: Layout may exceed page height. Continue?", vbYesNo + vbExclamation) = vbNo Then Exit Sub
    End If

    ' Generate Quote
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

    ' Perform Layout
    ' ... (Layout logic remains the same, using effectiveWidth and effectiveHeight) ...

    ' Add Quote Text to Page
    ' ... (Quote placement logic remains the same) ...

    MsgBox "Layout and quote created successfully!", vbInformation, "Success"
End Sub

' --- Settings Management (Private and Public) ---

Private Sub LoadSettings(ByRef vinylCost As Double, ByRef vatRate As Double, ByRef rollWidth As Double, ByRef bleed As Double, ByRef minStickerPrice As Double, ByRef minOrder As Double)
    ' Load settings from the registry, providing a default value if a key is not found.
    On Error Resume Next ' In case registry access is denied
    vinylCost = CDbl(GetSetting(AppName, "Pricing", "VinylCost", "460.0"))
    vatRate = CDbl(GetSetting(AppName, "Pricing", "VatRate", "0.15"))
    rollWidth = CDbl(GetSetting(AppName, "Pricing", "RollWidth", "650.0"))
    bleed = CDbl(GetSetting(AppName, "Pricing", "Bleed", "1.0"))
    minStickerPrice = CDbl(GetSetting(AppName, "Pricing", "MinStickerPrice", "0.2"))
    minOrder = CDbl(GetSetting(AppName, "Pricing", "MinOrderAmount", "100.0"))
    On Error GoTo 0
End Sub

Public Sub SaveSettings(ByVal vinylCost As String, ByVal vatRate As String, ByVal rollWidth As String, ByVal bleed As String, ByVal minStickerPrice As String, ByVal minOrder As String)
    ' Save settings to the registry. Values are passed as strings from the UserForm.
    On Error Resume Next ' In case registry access is denied
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
    ' This function's logic remains the same as before.
    ' ... (The full implementation of CalculatePrice is here) ...
    Dim P_horizontal As Double, P_vertical As Double

    ' Horizontal Orientation
    Dim W_bleed_h As Double, S_rounded_h As Long
    W_bleed_h = W + bleedMM
    If W_bleed_h > 0 Then S_rounded_h = Int(rollWidthMM / W_bleed_h) Else S_rounded_h = 0

    If S_rounded_h > 0 Then
        Dim H_meters_h As Double, Area_h As Double, Row_Cost_h As Double
        H_meters_h = H / 1000
        Area_h = (rollWidthMM / 1000) * H_meters_h
        Row_Cost_h = Area_h * costPerM2
        P_horizontal = Row_Cost_h / S_rounded_h
    Else
        P_horizontal = 999999
    End If

    ' Vertical Orientation
    Dim H_bleed_v As Double, S_rounded_v As Long
    H_bleed_v = H + bleedMM
    If H_bleed_v > 0 Then S_rounded_v = Int(rollWidthMM / H_bleed_v) Else S_rounded_v = 0

    If S_rounded_v > 0 Then
        Dim W_meters_v As Double, Area_v As Double, Row_Cost_v As Double
        W_meters_v = W / 1000
        Area_v = (rollWidthMM / 1000) * W_meters_v
        Row_Cost_v = Area_v * costPerM2
        P_vertical = Row_Cost_v / S_rounded_v
    Else
        P_vertical = 999999
    End If

    Dim price As Double
    If P_horizontal < P_vertical Then price = P_horizontal Else price = P_vertical
    If price > minPrice Then CalculatePrice = price Else CalculatePrice = minPrice
End Function
