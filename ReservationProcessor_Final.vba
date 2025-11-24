Sub ProcessReservations()
    '=====================================================================
    ' Macro: ProcessReservations
    ' Purpose: Process data from DELIMITED DATA to ENTERED ON sheet
    ' Features:
    '   - TDF calculation (20 AED for 1BA, 40 AED for 2BA per night)
    '   - Duplicate RESV_ID check
    '   - Spillover row handling
    '   - Date format preservation (dd/mm/yyyy)
    '   - Season formula (Column T)
    '   - Booking Lead Time formula (Column U)
    '   - Events Dates formula (Column V)
    '=====================================================================

    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowSource As Long
    Dim lastRowTarget As Long
    Dim i As Long
    Dim targetRow As Long
    Dim resvID As String
    Dim roomCategory As String
    Dim nights As Long
    Dim tdfCharge As Double
    Dim netAmount As Double
    Dim totalAmount As Double
    Dim adrAmount As Double
    Dim shareAmount As Double
    Dim fullName As String
    Dim firstName As String
    Dim arrivalDate As Date
    Dim departureDate As Date
    Dim existingIDs As Object
    Dim processCount As Long
    Dim skipCount As Long

    ' Initialize
    Set wb = ThisWorkbook
    Set existingIDs = CreateObject("Scripting.Dictionary")
    processCount = 0
    skipCount = 0

    ' Set source and target worksheets
    On Error Resume Next
    Set wsSource = wb.Worksheets("DELIMITED DATA")
    Set wsTarget = wb.Worksheets("ENTERED ON")
    On Error GoTo 0

    If wsSource Is Nothing Then
        MsgBox "Error: 'DELIMITED DATA' sheet not found!", vbCritical
        Exit Sub
    End If

    If wsTarget Is Nothing Then
        MsgBox "Error: 'ENTERED ON' sheet not found!", vbCritical
        Exit Sub
    End If

    ' Turn off screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    ' Fix spillover rows in source data first
    Call FixSpilloverRows(wsSource)

    ' Get existing RESV IDs from ENTERED ON sheet (Column S)
    lastRowTarget = wsTarget.Cells(wsTarget.Rows.Count, "S").End(xlUp).Row
    If lastRowTarget > 1 Then
        For i = 2 To lastRowTarget
            resvID = Trim(CStr(wsTarget.Cells(i, "S").Value))
            If Len(resvID) > 0 Then
                existingIDs(resvID) = True
            End If
        Next i
    End If

    ' Get last row in source
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "M").End(xlUp).Row

    ' Find next available row in target
    targetRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row + 1
    If targetRow = 2 And Len(Trim(wsTarget.Cells(2, "A").Value)) = 0 Then
        targetRow = 2
    End If

    ' Process each row from DELIMITED DATA
    For i = 2 To lastRowSource
        ' Get RESV_NAME_ID (Column M) and INSERT_DATE (Column Y)
        resvID = Trim(CStr(wsSource.Cells(i, "M").Value))
        Dim insertDate As String
        insertDate = Trim(CStr(wsSource.Cells(i, "Y").Value))

        ' Create combined RESV ID (RESV_NAME_ID + INSERT_DATE)
        If Len(resvID) > 0 And Len(insertDate) > 0 Then
            resvID = resvID & insertDate
        End If

        ' Check if RESV_ID already exists (skip duplicates)
        If Len(resvID) > 0 And existingIDs.exists(resvID) Then
            skipCount = skipCount + 1
            GoTo NextRow
        End If

        ' Get data from source columns
        fullName = Trim(CStr(wsSource.Cells(i, "Q").Value)) ' FULL_NAME (Column Q)

        ' Split name into Last Name and First Name
        Call SplitName(fullName, fullName, firstName)

        ' Get dates
        On Error Resume Next
        arrivalDate = wsSource.Cells(i, "AC").Value ' ARRIVAL (Column AC)
        departureDate = wsSource.Cells(i, "R").Value ' DEPARTURE (Column R)
        On Error GoTo 0

        ' Get nights (Column AD)
        On Error Resume Next
        nights = CLng(wsSource.Cells(i, "AD").Value)
        If Err.Number <> 0 Then nights = 0
        On Error GoTo 0

        ' Get room category (Column V)
        roomCategory = Trim(CStr(wsSource.Cells(i, "V").Value))

        ' Calculate TDF
        tdfCharge = CalculateTDF(roomCategory, nights)

        ' Get SHARE_AMOUNT (Column AF) - this becomes AMOUNT (Column P)
        On Error Resume Next
        shareAmount = CDbl(wsSource.Cells(i, "AF").Value)
        If Err.Number <> 0 Then shareAmount = 0
        On Error GoTo 0

        ' Get SHARE_AMOUNT_PER_STAY (Column AI) - this becomes NET (Column I)
        On Error Resume Next
        netAmount = CDbl(wsSource.Cells(i, "AI").Value)
        If Err.Number <> 0 Then netAmount = 0
        On Error GoTo 0

        ' Calculate TOTAL (NET + TDF)
        totalAmount = netAmount + tdfCharge

        ' Calculate ADR (AMOUNT / NIGHTS)
        If nights > 0 Then
            adrAmount = shareAmount / nights
        Else
            adrAmount = 0
        End If

        ' Write data to ENTERED ON sheet
        ' Column A: FULL_NAME (Last Name)
        wsTarget.Cells(targetRow, "A").Value = fullName

        ' Column B: FIRST NAME
        wsTarget.Cells(targetRow, "B").Value = firstName

        ' Column C: ARRIVAL (maintain dd/mm/yyyy format)
        wsTarget.Cells(targetRow, "C").Value = arrivalDate
        wsTarget.Cells(targetRow, "C").NumberFormat = "dd/mm/yyyy"

        ' Column D: DEPARTURE (maintain dd/mm/yyyy format)
        wsTarget.Cells(targetRow, "D").Value = departureDate
        wsTarget.Cells(targetRow, "D").NumberFormat = "dd/mm/yyyy"

        ' Column E: NIGHTS
        wsTarget.Cells(targetRow, "E").Value = nights

        ' Column F: PERSONS (Column S from source)
        wsTarget.Cells(targetRow, "F").Value = wsSource.Cells(i, "S").Value

        ' Column G: ROOM (Column V from source - room category)
        wsTarget.Cells(targetRow, "G").Value = roomCategory

        ' Column H: TDF
        wsTarget.Cells(targetRow, "H").Value = tdfCharge
        wsTarget.Cells(targetRow, "H").NumberFormat = "0"

        ' Column I: NET
        wsTarget.Cells(targetRow, "I").Value = netAmount
        wsTarget.Cells(targetRow, "I").NumberFormat = "0.000"

        ' Column J: TOTAL
        wsTarget.Cells(targetRow, "J").Value = totalAmount
        wsTarget.Cells(targetRow, "J").NumberFormat = "0.000"

        ' Column K: RATE_CODE (Column W from source)
        wsTarget.Cells(targetRow, "K").Value = wsSource.Cells(i, "W").Value

        ' Column L: INSERT_USER (Column X from source)
        wsTarget.Cells(targetRow, "L").Value = wsSource.Cells(i, "X").Value

        ' Column M: C_T_S_NAME (Column AG from source)
        wsTarget.Cells(targetRow, "M").Value = wsSource.Cells(i, "AG").Value

        ' Column N: SHORT_RESV_STATUS (Column AH from source)
        wsTarget.Cells(targetRow, "N").Value = wsSource.Cells(i, "AH").Value

        ' Column O: ADR
        wsTarget.Cells(targetRow, "O").Value = adrAmount
        wsTarget.Cells(targetRow, "O").NumberFormat = "0"

        ' Column P: AMOUNT
        wsTarget.Cells(targetRow, "P").Value = shareAmount
        wsTarget.Cells(targetRow, "P").NumberFormat = "0"

        ' Column Q: COMMENT (leave blank)
        wsTarget.Cells(targetRow, "Q").Value = ""

        ' Column R: C=CHECK (leave blank)
        wsTarget.Cells(targetRow, "R").Value = ""

        ' Column S: RESV ID
        wsTarget.Cells(targetRow, "S").Value = resvID

        ' Column T: Season (Formula)
        Call AddSeasonFormula(wsTarget, targetRow)

        ' Column U: Booking Lead Time (Formula)
        Call AddLeadTimeFormula(wsTarget, targetRow)

        ' Column V: Events Dates (Formula)
        Call AddEventsFormula(wsTarget, targetRow)

        ' Mark this ID as processed
        If Len(resvID) > 0 Then
            existingIDs(resvID) = True
        End If

        processCount = processCount + 1
        targetRow = targetRow + 1

NextRow:
    Next i

    ' Restore settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Show completion message
    MsgBox "Processing Complete!" & vbCrLf & vbCrLf & _
           "Records processed: " & processCount & vbCrLf & _
           "Duplicates skipped: " & skipCount, vbInformation, "Success"

End Sub

Function CalculateTDF(roomCategory As String, nights As Long) As Double
    '=====================================================================
    ' Calculate Tourism Dirham Fee (TDF)
    ' Rules:
    '   - 1BA: 20 AED per night (capped at 600 AED for 30+ nights)
    '   - 2BA: 40 AED per night (capped at 1200 AED for 30+ nights)
    '   - Other room types: 0 AED
    '=====================================================================

    Dim tdf As Double
    tdf = 0

    If nights <= 0 Then
        CalculateTDF = 0
        Exit Function
    End If

    Select Case UCase(Trim(roomCategory))
        Case "1BA"
            If nights <= 30 Then
                tdf = nights * 20
            Else
                tdf = 600 ' Cap at 600 for > 30 nights
            End If

        Case "2BA"
            If nights <= 30 Then
                tdf = nights * 40
            Else
                tdf = 1200 ' Cap at 1200 for > 30 nights
            End If

        Case Else
            tdf = 0
    End Select

    CalculateTDF = tdf
End Function

Sub FixSpilloverRows(ws As Worksheet)
    '=====================================================================
    ' Fix spillover rows where data continues on the next line
    ' Pattern: Rows starting with "[Letter]- " (e.g., "T- ", "S- ", "C- ")
    '=====================================================================

    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String

    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row

    ' Process from bottom to top to avoid row shifting issues
    For i = lastRow To 2 Step -1
        cellValue = Trim(CStr(ws.Cells(i, "A").Value))

        ' Check if this is a spillover row
        If Len(cellValue) > 2 Then
            If Mid(cellValue, 2, 1) = "-" And cellValue Like "[A-Z]- *" Then
                ' This is a spillover row - merge with previous row
                If i > 1 Then
                    ' Append to Column AG of previous row (C_T_S_NAME)
                    ws.Cells(i - 1, "AG").Value = ws.Cells(i - 1, "AG").Value & " " & cellValue

                    ' Move spillover data to previous row
                    If Len(Trim(CStr(ws.Cells(i, "B").Value))) > 0 Then
                        ws.Cells(i - 1, "AH").Value = ws.Cells(i, "B").Value ' SHORT_RESV_STATUS
                    End If

                    If Len(Trim(CStr(ws.Cells(i, "C").Value))) > 0 Then
                        ws.Cells(i - 1, "AI").Value = ws.Cells(i, "C").Value ' SHARE_AMOUNT_PER_STAY
                    End If

                    ' Delete the spillover row
                    ws.Rows(i).Delete
                End If
            End If
        End If
    Next i
End Sub

Sub SplitName(fullName As String, ByRef lastName As String, ByRef firstName As String)
    '=====================================================================
    ' Split full name into Last Name and First Name
    ' Format: "LastName,FirstName,Title" or "LastName,FirstName"
    '=====================================================================

    Dim nameParts() As String
    Dim commaPos As Long

    lastName = fullName
    firstName = ""

    ' Remove quotes if present
    fullName = Replace(fullName, """", "")

    ' Check if comma exists
    commaPos = InStr(fullName, ",")
    If commaPos > 0 Then
        ' Split by comma
        nameParts = Split(fullName, ",")

        If UBound(nameParts) >= 0 Then
            lastName = Trim(nameParts(0))
        End If

        If UBound(nameParts) >= 1 Then
            firstName = Trim(nameParts(1))
        End If
    End If
End Sub

Sub AddSeasonFormula(ws As Worksheet, rowNum As Long)
    '=====================================================================
    ' Add Season formula to Column T
    ' Logic: Winter (Oct-Apr), Summer (May-Sep)
    '=====================================================================

    Dim formula As String
    formula = "=IF(OR(AND(ISNUMBER(C" & rowNum & "), OR(MONTH(C" & rowNum & ")<=4, MONTH(C" & rowNum & ")>=10)), " & _
              "AND(ISNUMBER(D" & rowNum & "), OR(MONTH(D" & rowNum & ")<=4, MONTH(D" & rowNum & ")>=10))), " & _
              """Winter"", IF(OR(AND(ISNUMBER(C" & rowNum & "), MONTH(C" & rowNum & ")>=5, MONTH(C" & rowNum & ")<=9), " & _
              "AND(ISNUMBER(D" & rowNum & "), MONTH(D" & rowNum & ")>=5, MONTH(D" & rowNum & ")<=9)), ""Summer"", """"))"

    ws.Cells(rowNum, "T").Formula = formula
End Sub

Sub AddLeadTimeFormula(ws As Worksheet, rowNum As Long)
    '=====================================================================
    ' Add Booking Lead Time formula to Column U
    ' Logic: Arrival Date - Today's Date
    '=====================================================================

    Dim formula As String
    formula = "=IF(ISNUMBER(C" & rowNum & "), C" & rowNum & " - TODAY(), """")"

    ws.Cells(rowNum, "U").Formula = formula
End Sub

Sub AddEventsFormula(ws As Worksheet, rowNum As Long)
    '=====================================================================
    ' Add Events Dates formula to Column V
    ' Checks if arrival date falls within any event period
    '=====================================================================

    Dim formula As String
    formula = "=IF(AND(C" & rowNum & ">=DATE(2025,1,26),C" & rowNum & "<=DATE(2025,1,31)),""Arab Health""," & _
              "IF(AND(C" & rowNum & ">=DATE(2025,2,16),C" & rowNum & "<=DATE(2025,2,21)),""Gulf Food""," & _
              "IF(AND(C" & rowNum & ">=DATE(2025,3,1),C" & rowNum & "<=DATE(2025,3,29)),""Ramadan""," & _
              "IF(AND(C" & rowNum & ">=DATE(2025,3,30),C" & rowNum & "<=DATE(2025,4,2)),""Eid Al Fitr""," & _
              "IF(AND(C" & rowNum & ">=DATE(2025,6,6),C" & rowNum & "<=DATE(2025,6,9)),""Eid Al Adha""," & _
              "IF(AND(C" & rowNum & ">=DATE(2025,10,12),C" & rowNum & "<=DATE(2025,10,17)),""GITEX""," & _
              "IF(AND(C" & rowNum & ">=DATE(2025,11,3),C" & rowNum & "<=DATE(2025,11,7)),""Gulf Food Manufacturing""," & _
              "IF(AND(C" & rowNum & ">=DATE(2025,11,16),C" & rowNum & "<=DATE(2025,11,21)),""Air Show""," & _
              "IF(AND(C" & rowNum & ">=DATE(2025,11,23),C" & rowNum & "<=DATE(2025,11,28)),""Big 5""," & _
              "IF(AND(C" & rowNum & ">=DATE(2025,11,29),C" & rowNum & "<=DATE(2025,12,2)),""National Day Holidays""," & _
              "IF(AND(C" & rowNum & ">=DATE(2025,12,4),C" & rowNum & "<=DATE(2025,12,7)),""F1 Yas Island""," & _
              "IF(AND(C" & rowNum & ">=DATE(2025,12,26),C" & rowNum & "<=DATE(2025,12,31)),""New Year's Eve"",""""))))))))))))"

    ws.Cells(rowNum, "V").Formula = formula
End Sub
