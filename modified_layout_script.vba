' VBA Script for CorelDRAW to create sticker layouts
'
' This macro arranges selected shapes in a boustrophedon (serpentine)
' layout, ensuring the creation order allows for a top-left start
' for cutting machines. This version automatically calculates the
' number of stickers per row.

' Defines a structure to hold X, Y coordinates
Private Type Point
    X As Double
    Y As Double
End Type

Sub CreateStickerLayout()
    ' Set the document units to millimeters for consistency
    ActiveDocument.Unit = cdrMillimeter

    ' Check for an active selection
    If ActiveDocument Is Nothing Or ActiveSelection.Shapes.Count = 0 Then
        MsgBox "Please select at least one shape to serve as the sticker template.", vbExclamation, "No Selection"
        Exit Sub
    End If

    ' If more than one shape is selected, inform the user only the first will be used
    If ActiveSelection.Shapes.Count > 1 Then
        MsgBox "More than one shape is selected. Only the first shape in the selection will be used as the template.", vbInformation, "Multiple Shapes Selected"
    End If

    ' Prompt for the total number of stickers required
    Dim totalStickers As Long
    On Error Resume Next ' Handle non-numeric input
    totalStickers = CLng(InputBox("Enter the total number of stickers (including the selected one):", "Total Stickers", 10))
    On Error GoTo 0
    If totalStickers <= 0 Then
        MsgBox "Invalid input. Please enter a positive number for the total amount of stickers.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Get the first selected shape and its dimensions
    Dim baseShape As Shape
    Set baseShape = ActiveSelection.Shapes(1)
    Dim stickerWidth As Double, stickerHeight As Double
    stickerWidth = baseShape.SizeWidth
    stickerHeight = baseShape.SizeHeight

    ' Get page dimensions
    Dim pageWidth As Double, pageHeight As Double
    pageWidth = ActivePage.SizeWidth
    pageHeight = ActivePage.SizeHeight

    ' Automatically calculate the number of stickers per row
    Dim stickersPerRow As Long
    If stickerWidth > 0 Then
        stickersPerRow = Int(pageWidth / stickerWidth)
    Else
        stickersPerRow = 0
    End If

    If stickersPerRow <= 0 Then
        MsgBox "The selected sticker is wider than the page. Cannot generate layout.", vbExclamation, "Sticker Too Wide"
        Exit Sub
    End If

    Dim spacingX As Double
    If stickersPerRow > 1 Then
        spacingX = (pageWidth - (stickersPerRow * stickerWidth)) / (stickersPerRow - 1)
    Else
        spacingX = 0 ' No horizontal spacing if only one sticker per row
    End If

    ' Prompt for vertical spacing between rows
    Dim spacingY As Double
    On Error Resume Next ' Handle non-numeric input
    spacingY = CDbl(InputBox("Enter the spacing between rows (in mm):", "Vertical Spacing", 0.5))
    On Error GoTo 0
    If spacingY < 0 Then
        MsgBox "Invalid input. Please enter a non-negative number for spacing.", vbExclamation, "Invalid Input"
        Exit Sub
    End If

    ' Check for vertical page overflow before creating stickers
    Dim numRows As Long
    numRows = (totalStickers + stickersPerRow - 1) \ stickersPerRow

    Dim totalLayoutHeight As Double
    totalLayoutHeight = (numRows * stickerHeight) + ((numRows - 1) * spacingY)

    If totalLayoutHeight > pageHeight Then
        If MsgBox("Warning: The layout is projected to exceed the page height. This may result in clipped stickers. Do you want to continue anyway?", vbYesNo + vbExclamation, "Layout May Not Fit") = vbNo Then
            Exit Sub
        End If
    End If

    ' Array to hold all sticker positions
    Dim positions() As Point
    ReDim positions(totalStickers - 1)

    ' Define page starting coordinates
    Dim startX As Double, startY As Double
    startX = ActivePage.LeftX
    startY = ActivePage.TopY

    Dim rowCounter As Long, colCounter As Long
    Dim i As Long

    ' === Step 1: Calculate all positions first ===
    For i = 0 To totalStickers - 1
        rowCounter = i \ stickersPerRow
        colCounter = i Mod stickersPerRow

        Dim currentX As Double, currentY As Double

        ' Calculate Y position for the current row
        currentY = startY - rowCounter * (stickerHeight + spacingY)

        ' Calculate X position, accounting for boustrophedon layout
        If (rowCounter Mod 2) = 0 Then
            ' Even row (0, 2, ...): layout is left-to-right
            currentX = startX + colCounter * (stickerWidth + spacingX)
        Else
            ' Odd row (1, 3, ...): layout is right-to-left
            currentX = startX + (stickersPerRow - 1 - colCounter) * (stickerWidth + spacingX)
        End If

        positions(i).X = currentX
        positions(i).Y = currentY
    Next i

    ' === Step 2: Create shapes in reverse order for correct cutting sequence ===
    Dim duplicateShape As Shape
    ' Create duplicates for positions N-1 down to 1
    For i = totalStickers - 1 To 1 Step -1
        Set duplicateShape = baseShape.Duplicate
        duplicateShape.SetPosition positions(i).X, positions(i).Y
    Next i

    ' Move the original shape to the first position (top-left) LAST
    baseShape.SetPosition positions(0).X, positions(0).Y

    MsgBox "Sticker layout created successfully. The printer should now start from the top-left.", vbInformation, "Success"
End Sub
