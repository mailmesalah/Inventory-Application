Attribute VB_Name = "MPrintPriceTag"
Public charArray(43) As String
Public Sub printFromPrintMGrid(PrintMGrid As MSFlexGrid, gPrintPurchaseRate As Single, gPrintSize As Single)
Dim r As Long, x As Long, y As Long, height As Long
    
    x = 800 'CHANGE THIS TO ADJUST STARTING X
    y = 650 'CHANGE THIS TO ADJUST STARTING Y
    r = 0

    Do While r < PrintMGrid.Rows
        If y >= Printer.height - 2000 Then
            Printer.EndDoc
            x = 800 'CHANGE THIS TO ADJUST STARTING X
            y = 650 'CHANGE THIS TO ADJUST STARTING Y
        End If
        
      
        'FIRST COL
        If r >= PrintMGrid.Rows Then
            Exit Do
        End If
        
        Printer.FontName = "Arial"
        Printer.FontUnderline = True
        Printer.FontSize = 12
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Silky Days"

        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.CurrentX = x
        Printer.CurrentY = y + 300
        Printer.Print PrintMGrid.TextMatrix(r, 0)

        Printer.FontName = "Rupee"
        Printer.FontSize = 11
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y + 500
        Printer.Print "`." & Format(PrintMGrid.TextMatrix(r, 2), "0")

        Printer.FontBold = False
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 750
        Printer.CurrentY = y + 520
        Printer.Print getUniversaloFor(Format(PrintMGrid.TextMatrix(r, 1), "0"))

        Printer.FontName = "Arial"
        Printer.FontSize = 10
        Printer.CurrentX = x + 100
        Printer.CurrentY = y + 740
        Printer.Print "Size : " & Trim(PrintMGrid.TextMatrix(r, 3))
        Printer.CurrentX = x
        Printer.CurrentY = y
        
        x = x + 2300
        r = r + 1
        
        'SECOND COL
        If r >= PrintMGrid.Rows Then
            Exit Do
        End If
        
        Printer.FontName = "Arial"
        Printer.FontUnderline = True
        Printer.FontSize = 12
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Silky Days"
        
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.CurrentX = x
        Printer.CurrentY = y + 300
        Printer.Print PrintMGrid.TextMatrix(r, 0)
        
        Printer.FontName = "Rupee"
        Printer.FontSize = 11
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y + 500
        Printer.Print "`." & Format(PrintMGrid.TextMatrix(r, 2), "0")
        
        Printer.FontBold = False
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 750
        Printer.CurrentY = y + 520
        Printer.Print getUniversaloFor(Format(PrintMGrid.TextMatrix(r, 1), "0"))
        
        Printer.FontName = "Arial"
        Printer.FontSize = 10
        Printer.CurrentX = x + 100
        Printer.CurrentY = y + 740
        Printer.Print "Size : " & Trim(PrintMGrid.TextMatrix(r, 3))
        Printer.CurrentX = x
        Printer.CurrentY = y
        
        x = x + 2300
        r = r + 1
        
        
        'THIRD COL
        If r >= PrintMGrid.Rows Then
            Exit Do
        End If
        
        Printer.FontName = "Arial"
        Printer.FontUnderline = True
        Printer.FontSize = 12
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Silky Days"
               
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.CurrentX = x
        Printer.CurrentY = y + 300
        Printer.Print PrintMGrid.TextMatrix(r, 0)
        
        Printer.FontName = "Rupee"
        Printer.FontSize = 11
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y + 500
        Printer.Print "`." & Format(PrintMGrid.TextMatrix(r, 2), "0")
        
        Printer.FontBold = False
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 750
        Printer.CurrentY = y + 520
        Printer.Print getUniversaloFor(Format(PrintMGrid.TextMatrix(r, 1), "0"))
        
        Printer.FontName = "Arial"
        Printer.FontSize = 10
        Printer.CurrentX = x + 100
        Printer.CurrentY = y + 740
        Printer.Print "Size : " & Trim(PrintMGrid.TextMatrix(r, 3))
        Printer.CurrentX = x
        Printer.CurrentY = y
        
        x = x + 2300
        r = r + 1
        
        
        'FOURTH COL
        If r >= PrintMGrid.Rows Then
            Exit Do
        End If
        
        Printer.FontName = "Arial"
        Printer.FontUnderline = True
        Printer.FontSize = 12
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Silky Days"
               
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.CurrentX = x
        Printer.CurrentY = y + 300
        Printer.Print PrintMGrid.TextMatrix(r, 0)
        
        Printer.FontName = "Rupee"
        Printer.FontSize = 11
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y + 500
        Printer.Print "`." & Format(PrintMGrid.TextMatrix(r, 2), "0")
        
        Printer.FontBold = False
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 750
        Printer.CurrentY = y + 520
        Printer.Print getUniversaloFor(Format(PrintMGrid.TextMatrix(r, 1), "0"))
        
        Printer.FontName = "Arial"
        Printer.FontSize = 10
        Printer.CurrentX = x + 100
        Printer.CurrentY = y + 740
        Printer.Print "Size : " & Trim(PrintMGrid.TextMatrix(r, 3))
        Printer.CurrentX = x
        Printer.CurrentY = y
        
        x = x + 2300
        r = r + 1
        
        
        'FIFTH COL
        If r >= PrintMGrid.Rows Then
            Exit Do
        End If
        
        Printer.FontName = "Arial"
        Printer.FontUnderline = True
        Printer.FontSize = 12
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Silky Days"
               
        Printer.FontSize = 10
        Printer.FontBold = False
        Printer.FontUnderline = False
        Printer.CurrentX = x
        Printer.CurrentY = y + 300
        Printer.Print PrintMGrid.TextMatrix(r, 0)
        
        Printer.FontName = "Rupee"
        Printer.FontSize = 11
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y + 500
        Printer.Print "`." & Format(PrintMGrid.TextMatrix(r, 2), "0")
        
        Printer.FontBold = False
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 750
        Printer.CurrentY = y + 520
        Printer.Print getUniversaloFor(Format(PrintMGrid.TextMatrix(r, 1), "0"))
        
        Printer.FontName = "Arial"
        Printer.FontSize = 10
        Printer.CurrentX = x + 100
        Printer.CurrentY = y + 740
        Printer.Print "Size : " & Trim(PrintMGrid.TextMatrix(r, 3))
        Printer.CurrentX = x
        Printer.CurrentY = y
        
        x = 800 'CHANGE THIS TO ADJUST STARTING X
        y = y + 1200 ' CHANGE THIS ONE TO ADJUST ROW PRINTING
        r = r + 1
        
    Loop
    Printer.EndDoc

    MsgBox "Successfully Send to Printer !", vbInformation
End Sub

Public Sub printBarcodesFromGrid(PrintMGrid As MSFlexGrid, gPrintItem As Single, gPrintPurchaseRate As Single, gPrintSize As Single, gBarCodeCol As Single)

Dim r As Long, x As Long, y As Long, Cy As Long, Cx As Long, height As Long

    initialiseBinaryCodes
    height = 350
    
    x = 800 'CHANGE THIS TO ADJUST STARTING X
    y = 650 'CHANGE THIS TO ADJUST STARTING Y
    r = 0

    Do While r < PrintMGrid.Rows
        If y >= Printer.height - 2000 Then
            Printer.EndDoc
            x = 800 'CHANGE THIS TO ADJUST STARTING X
            y = 650 'CHANGE THIS TO ADJUST STARTING Y
        End If
        
      
        'FIRST COL
        If r >= PrintMGrid.Rows Then
            Exit Do
        End If
        
        Printer.FontName = "Arial"
        Printer.FontSize = 10
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Silky Days"
        Printer.FontBold = False

        Printer.FontSize = 8
        Printer.CurrentX = x - 150
        Printer.CurrentY = y + 200
        Printer.Print PrintMGrid.TextMatrix(r, 0)
        
        Cy = Printer.CurrentY
        printBarcode PrintMGrid.TextMatrix(r, gBarCodeCol), x - 150, Cy, height
        Printer.CurrentX = x
        Printer.CurrentY = y + 300

        Printer.FontName = "Rupee"
        Printer.FontSize = 9
        Printer.FontBold = True
        Printer.CurrentX = x - 150
        Printer.CurrentY = y + 750
        Printer.Print "`." & Format(PrintMGrid.TextMatrix(r, 2), "0")
        Printer.FontBold = False
        Printer.CurrentX = x + 750
        
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 400
        Printer.CurrentY = y + 750
        Printer.Print getUniversaloFor(Format(PrintMGrid.TextMatrix(r, 1), "0"))

        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 850
        Printer.CurrentY = y + 750
        Printer.Print Trim(PrintMGrid.TextMatrix(r, 3))
        Printer.CurrentX = x
        Printer.CurrentY = y
        
        x = x + 2300
        r = r + 1
        
        'SECOND COL
        If r >= PrintMGrid.Rows Then
            Exit Do
        End If
        
        Printer.FontName = "Arial"
        Printer.FontSize = 10
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Silky Days"
        Printer.FontBold = False

        Printer.FontSize = 8
        Printer.CurrentX = x - 150
        Printer.CurrentY = y + 200
        Printer.Print PrintMGrid.TextMatrix(r, 0)
        
        Cy = Printer.CurrentY
        printBarcode PrintMGrid.TextMatrix(r, gBarCodeCol), x - 150, Cy, height
        Printer.CurrentX = x
        Printer.CurrentY = y + 300

        Printer.FontName = "Rupee"
        Printer.FontSize = 9
        Printer.FontBold = True
        Printer.CurrentX = x - 150
        Printer.CurrentY = y + 750
        Printer.Print "`." & Format(PrintMGrid.TextMatrix(r, 2), "0")
        Printer.FontBold = False
        Printer.CurrentX = x + 750
        
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 400
        Printer.CurrentY = y + 750
        Printer.Print getUniversaloFor(Format(PrintMGrid.TextMatrix(r, 1), "0"))

        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 850
        Printer.CurrentY = y + 750
        Printer.Print Trim(PrintMGrid.TextMatrix(r, 3))
        Printer.CurrentX = x
        Printer.CurrentY = y

        x = x + 2300
        r = r + 1
        
        'THIRD COL
        If r >= PrintMGrid.Rows Then
            Exit Do
        End If
        
        Printer.FontName = "Arial"
        Printer.FontSize = 10
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Silky Days"
        Printer.FontBold = False

        Printer.FontSize = 8
        Printer.CurrentX = x - 150
        Printer.CurrentY = y + 200
        Printer.Print PrintMGrid.TextMatrix(r, 0)
        
        Cy = Printer.CurrentY
        printBarcode PrintMGrid.TextMatrix(r, gBarCodeCol), x - 150, Cy, height
        Printer.CurrentX = x
        Printer.CurrentY = y + 300

        Printer.FontName = "Rupee"
        Printer.FontSize = 9
        Printer.FontBold = True
        Printer.CurrentX = x - 150
        Printer.CurrentY = y + 750
        Printer.Print "`." & Format(PrintMGrid.TextMatrix(r, 2), "0")
        Printer.FontBold = False
        Printer.CurrentX = x + 750
        
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 400
        Printer.CurrentY = y + 750
        Printer.Print getUniversaloFor(Format(PrintMGrid.TextMatrix(r, 1), "0"))

        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 850
        Printer.CurrentY = y + 750
        Printer.Print Trim(PrintMGrid.TextMatrix(r, 3))
        Printer.CurrentX = x
        Printer.CurrentY = y
        
        x = x + 2300
        r = r + 1
        
        'FOURTH COL
        If r >= PrintMGrid.Rows Then
            Exit Do
        End If
        
        Printer.FontName = "Arial"
        Printer.FontSize = 10
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Silky Days"
        Printer.FontBold = False

        Printer.FontSize = 8
        Printer.CurrentX = x - 150
        Printer.CurrentY = y + 200
        Printer.Print PrintMGrid.TextMatrix(r, 0)
        
        Cy = Printer.CurrentY
        printBarcode PrintMGrid.TextMatrix(r, gBarCodeCol), x - 150, Cy, height
        Printer.CurrentX = x
        Printer.CurrentY = y + 300

        Printer.FontName = "Rupee"
        Printer.FontSize = 9
        Printer.FontBold = True
        Printer.CurrentX = x - 150
        Printer.CurrentY = y + 750
        Printer.Print "`." & Format(PrintMGrid.TextMatrix(r, 2), "0")
        Printer.FontBold = False
        Printer.CurrentX = x + 750
        
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 400
        Printer.CurrentY = y + 750
        Printer.Print getUniversaloFor(Format(PrintMGrid.TextMatrix(r, 1), "0"))

        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 850
        Printer.CurrentY = y + 750
        Printer.Print Trim(PrintMGrid.TextMatrix(r, 3))
        Printer.CurrentX = x
        Printer.CurrentY = y
        
        x = x + 2300
        r = r + 1
        
        
        'FIFTH COL
        If r >= PrintMGrid.Rows Then
            Exit Do
        End If
        
        Printer.FontName = "Arial"
        Printer.FontSize = 10
        Printer.FontBold = True
        Printer.CurrentX = x
        Printer.CurrentY = y
        Printer.Print "Silky Days"
        Printer.FontBold = False

        Printer.FontSize = 8
        Printer.CurrentX = x - 150
        Printer.CurrentY = y + 200
        Printer.Print PrintMGrid.TextMatrix(r, 0)
        
        Cy = Printer.CurrentY
        printBarcode PrintMGrid.TextMatrix(r, gBarCodeCol), x - 150, Cy, height
        Printer.CurrentX = x
        Printer.CurrentY = y + 300

        Printer.FontName = "Rupee"
        Printer.FontSize = 9
        Printer.FontBold = True
        Printer.CurrentX = x - 150
        Printer.CurrentY = y + 750
        Printer.Print "`." & Format(PrintMGrid.TextMatrix(r, 2), "0")
        Printer.FontBold = False
        Printer.CurrentX = x + 750
        
        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 400
        Printer.CurrentY = y + 750
        Printer.Print getUniversaloFor(Format(PrintMGrid.TextMatrix(r, 1), "0"))

        Printer.FontName = "Arial"
        Printer.FontSize = 8
        Printer.CurrentX = x + 850
        Printer.CurrentY = y + 750
        Printer.Print Trim(PrintMGrid.TextMatrix(r, 3))
        Printer.CurrentX = x
        Printer.CurrentY = y
        
        x = 800 'CHANGE THIS TO ADJUST STARTING X
        y = y + 1200 ' CHANGE THIS ONE TO ADJUST ROW PRINTING
        r = r + 1
        
    Loop
    Printer.EndDoc

    MsgBox "Successfully Send to Printer !", vbInformation
End Sub
Private Sub initialiseBinaryCodes()
    charArray(0) = "000110100"
    charArray(1) = "100100001"
    charArray(2) = "001100001"
    charArray(3) = "101100000"
    charArray(4) = "000110001"
    charArray(5) = "100110000"
    charArray(6) = "001110000"
    charArray(7) = "000100101"
    charArray(8) = "100100100"
    charArray(9) = "001100100"
    charArray(10) = "100001001"
    charArray(11) = "001001001"
    charArray(12) = "101001000"
    charArray(13) = "000011001"
    charArray(14) = "100011000"
    charArray(15) = "001011000"
    charArray(16) = "000001101"
    charArray(17) = "100001100"
    charArray(18) = "001001100"
    charArray(19) = "000011100"
    charArray(20) = "100000011"
    charArray(21) = "001000011"
    charArray(22) = "101000010"
    charArray(23) = "000010011"
    charArray(24) = "100010010"
    charArray(25) = "001010010"
    charArray(26) = "000000111"
    charArray(27) = "100000110"
    charArray(28) = "001000110"
    charArray(29) = "000010110"
    charArray(30) = "110000001"
    charArray(31) = "011000001"
    charArray(32) = "111000000"
    charArray(33) = "010010001"
    charArray(34) = "110010000"
    charArray(35) = "011010000"
    charArray(36) = "010000101"
    charArray(37) = "110000100"
    charArray(38) = "011000100"
    charArray(39) = "010101000"
    charArray(40) = "010100010"
    charArray(41) = "010001010"
    charArray(42) = "000101010"
    charArray(43) = "010010100"
End Sub
Private Sub printBarcode(inputString As String, x As Long, y As Long, height As Long)
    Dim lineTop As Integer      'top of the line
    Dim lineBottom As Integer   'bottom of the line
    Dim leftSide As Integer     'left side of barcode
    Dim zerosWidth As Integer  'length of the narrow bars
    Dim binArray As Integer     'retrievs the binary code for the current letter
    Dim i, j, k As Integer
    Dim newInputString As String      'hold the modified string
    
    lineTop = y
    lineBottom = y + height
    leftSide = x
    zerosWidth = 20
    
    newInputString = "*" + StrConv(inputString, vbUpperCase) + "*"  'add start and stop char changes string to uppercase
    
    For i = 1 To Len(newInputString)  'loops through each char of the string
        
        binArray = integerValueOfCharactes(Mid(newInputString, i, 1))
        
        For j = 1 To Len(charArray(binArray))   'loops through the binary array
            
            For k = 1 To IIf(CInt(Mid(charArray(binArray), j, 1)) = 0, 1 * zerosWidth, 2 * zerosWidth) 'For 0 (1*zerosWidth) and for 1 (2*zerosWidth)
                
                If j Mod 2 Then  'odd char of the string are black, even are white
                    Printer.Line (leftSide, lineTop)-(leftSide, lineBottom), vbBlack
                Else
                    Printer.Line (leftSide, lineTop)-(leftSide, lineBottom), vbWhite
                End If
                leftSide = leftSide + 1
            Next k
        Next j
        
        For j = 1 To zerosWidth  'adds a white space between char
            Printer.Line (leftSide, lineTop)-(leftSide, lineBottom), vbWhite
            leftSide = leftSide + 1
        Next j
    Next i
End Sub
Function integerValueOfCharactes(inputCharacter As String) As Integer
    Select Case inputCharacter
        Case "0"
            integerValueOfCharactes = 0
            Exit Function
        Case "1"
            integerValueOfCharactes = 1
            Exit Function
        Case "2"
            integerValueOfCharactes = 2
            Exit Function
        Case "3"
            integerValueOfCharactes = 3
            Exit Function
        Case "4"
            integerValueOfCharactes = 4
            Exit Function
        Case "5"
            integerValueOfCharactes = 5
            Exit Function
        Case "6"
            integerValueOfCharactes = 6
            Exit Function
        Case "7"
            integerValueOfCharactes = 7
            Exit Function
        Case "8"
            integerValueOfCharactes = 8
            Exit Function
        Case "9"
            integerValueOfCharactes = 9
            Exit Function
        Case "A"
            integerValueOfCharactes = 10
            Exit Function
        Case "B"
            integerValueOfCharactes = 11
            Exit Function
        Case "C"
            integerValueOfCharactes = 12
            Exit Function
        Case "D"
            integerValueOfCharactes = 13
            Exit Function
        Case "E"
            integerValueOfCharactes = 14
            Exit Function
        Case "F"
            integerValueOfCharactes = 15
            Exit Function
        Case "G"
            integerValueOfCharactes = 16
            Exit Function
        Case "H"
            integerValueOfCharactes = 17
            Exit Function
        Case "I"
            integerValueOfCharactes = 18
            Exit Function
        Case "J"
            integerValueOfCharactes = 19
            Exit Function
        Case "K"
            integerValueOfCharactes = 20
            Exit Function
        Case "L"
            integerValueOfCharactes = 21
            Exit Function
        Case "M"
            integerValueOfCharactes = 22
            Exit Function
        Case "N"
            integerValueOfCharactes = 23
            Exit Function
        Case "O"
            integerValueOfCharactes = 24
            Exit Function
        Case "P"
            integerValueOfCharactes = 25
            Exit Function
        Case "Q"
            integerValueOfCharactes = 26
            Exit Function
        Case "R"
            integerValueOfCharactes = 27
            Exit Function
        Case "S"
            integerValueOfCharactes = 28
            Exit Function
        Case "T"
            integerValueOfCharactes = 29
            Exit Function
        Case "U"
            integerValueOfCharactes = 30
            Exit Function
        Case "V"
            integerValueOfCharactes = 31
            Exit Function
        Case "W"
            integerValueOfCharactes = 32
            Exit Function
        Case "X"
            integerValueOfCharactes = 33
            Exit Function
        Case "Y"
            integerValueOfCharactes = 34
            Exit Function
        Case "Z"
            integerValueOfCharactes = 35
            Exit Function
        Case "-"
            integerValueOfCharactes = 36
            Exit Function
        Case "."
            integerValueOfCharactes = 37
            Exit Function
        Case " "
            integerValueOfCharactes = 38
            Exit Function
        Case "$"
            integerValueOfCharactes = 39
            Exit Function
        Case "/"
            integerValueOfCharactes = 40
            Exit Function
        Case "+"
            integerValueOfCharactes = 41
            Exit Function
        Case "%"
            integerValueOfCharactes = 42
            Exit Function
        Case "*"
            integerValueOfCharactes = 43
            Exit Function
    End Select
End Function
Public Function getUniversaloFor(sPurchaseRate As String) As String
Dim tempPurchase As String, sCode As String
        tempPurchase = sPurchaseRate
        sCode = ""
        While Len(tempPurchase) > 0
            sCode = sCode & getCode(Left(tempPurchase, 1))
            tempPurchase = Right(tempPurchase, Len(tempPurchase) - 1)
        Wend
        getUniversaloFor = sCode
End Function
Private Function getCode(sChar As String) As String
    Select Case sChar
        Case "0"
            getCode = "O"
            Exit Function
        Case "1"
            getCode = "U"
            Exit Function
        Case "2"
            getCode = "N"
            Exit Function
        Case "3"
            getCode = "I"
            Exit Function
        Case "4"
            getCode = "V"
            Exit Function
        Case "5"
            getCode = "E"
            Exit Function
        Case "6"
            getCode = "R"
            Exit Function
        Case "7"
            getCode = "S"
            Exit Function
        Case "8"
            getCode = "A"
            Exit Function
        Case "9"
            getCode = "L"
            Exit Function
    End Select
End Function
