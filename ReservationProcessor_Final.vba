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
    '   - Auto-sort by Column M (C_T_S_NAME) descending
    '=====================================================================

    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRowSource As Long
    Dim lastRowTarget As Long
    Dim i As Long
    Dim targetRow As Long
    Dim resvID As String
    Dim insertDate As String
    Dim roomCategory As String
    Dim rateCode As String
    Dim resvStatus As String
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

    ' Ask user if they want to clear existing data
    Dim clearResponse As VbMsgBoxResult
    clearResponse = MsgBox("Do you want to CLEAR all existing data in ENTERED ON sheet before processing?" & vbCrLf & vbCrLf & _
                          "Click YES to clear and start fresh" & vbCrLf & _
                          "Click NO to append to existing data", vbYesNoCancel + vbQuestion, "Clear Existing Data?")

    If clearResponse = vbCancel Then
        Application.ScreenUpdating = True
        Application.Calculation = xlCalculationAutomatic
        Exit Sub
    End If

    If clearResponse = vbYes Then
        Dim clearLastRow As Long
        clearLastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
        If clearLastRow < 2 Then
            clearLastRow = wsTarget.Cells(wsTarget.Rows.Count, "S").End(xlUp).Row
        End If

        If clearLastRow > 1 Then
            On Error Resume Next
            wsTarget.Range("A2:V" & clearLastRow).ClearContents
            wsTarget.Range("A2:V" & clearLastRow).Borders.LineStyle = xlNone
            ' Clear specific formatting but PRESERVE conditional formatting
            With wsTarget.Range("A2:V" & clearLastRow)
                .Interior.ColorIndex = xlNone ' Clear fill colors
                .Font.Bold = False
                .Font.ColorIndex = xlAutomatic
                .NumberFormat = "General"
            End With
            If Err.Number <> 0 Then
                MsgBox "Error clearing data: " & Err.Description, vbCritical
                Application.ScreenUpdating = True
                Application.Calculation = xlCalculationAutomatic
                Exit Sub
            End If
            On Error GoTo 0
        End If
    End If

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

    ' Get last row in source - check multiple columns to find actual data
    Dim tempRow As Long
    lastRowSource = 1

    ' Check Column A
    tempRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    If tempRow > lastRowSource Then lastRowSource = tempRow

    ' Check Column M (RESV_NAME_ID)
    tempRow = wsSource.Cells(wsSource.Rows.Count, "M").End(xlUp).Row
    If tempRow > lastRowSource Then lastRowSource = tempRow

    ' Check Column Q (FULL_NAME)
    tempRow = wsSource.Cells(wsSource.Rows.Count, "Q").End(xlUp).Row
    If tempRow > lastRowSource Then lastRowSource = tempRow

    ' Find next available row in target - always start at row 2 if empty
    targetRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row
    If targetRow = 1 Or Len(Trim(wsTarget.Cells(2, "A").Value)) = 0 Then
        targetRow = 2
    Else
        targetRow = targetRow + 1
    End If

    ' Process each row from DELIMITED DATA (exclude last 2 summary rows)
    For i = 2 To lastRowSource - 2
        ' Get RESV_NAME_ID (Column M) and INSERT_DATE (Column Y)
        resvID = Trim(CStr(wsSource.Cells(i, "M").Value))

        ' Convert INSERT_DATE to string - treat as text, not date
        insertDate = Trim(CStr(wsSource.Cells(i, "Y").Text))

        ' Create combined RESV ID (RESV_NAME_ID + INSERT_DATE)
        If Len(resvID) > 0 And Len(insertDate) > 0 Then
            resvID = resvID & insertDate
        End If

        ' Get room category (Column V), rate code (Column W), and status (Column AH)
        roomCategory = Trim(CStr(wsSource.Cells(i, "V").Value))
        rateCode = Trim(CStr(wsSource.Cells(i, "W").Value))
        resvStatus = Trim(CStr(wsSource.Cells(i, "AH").Value))

        ' Skip PM rooms
        If UCase(roomCategory) = "PM" Then
            skipCount = skipCount + 1
            GoTo NextRow
        End If

        ' Skip HOUSEUSE rate codes
        If UCase(rateCode) = "HOUSEUSE" Then
            skipCount = skipCount + 1
            GoTo NextRow
        End If

        ' Skip CXL (cancelled) reservations
        If UCase(resvStatus) = "CXL" Then
            skipCount = skipCount + 1
            GoTo NextRow
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

        ' Get dates - parse from text format dd.mm.yy
        Dim arrivalText As String
        Dim departureText As String
        Dim dateParts() As String

        On Error Resume Next
        arrivalText = Trim(wsSource.Cells(i, "AC").Text)
        departureText = Trim(wsSource.Cells(i, "R").Text)

        ' Parse dd.mm.yy format to actual date
        If InStr(arrivalText, ".") > 0 Then
            dateParts = Split(arrivalText, ".")
            If UBound(dateParts) = 2 Then
                arrivalDate = DateSerial(2000 + CInt(dateParts(2)), CInt(dateParts(1)), CInt(dateParts(0)))
            End If
        End If

        If InStr(departureText, ".") > 0 Then
            dateParts = Split(departureText, ".")
            If UBound(dateParts) = 2 Then
                departureDate = DateSerial(2000 + CInt(dateParts(2)), CInt(dateParts(1)), CInt(dateParts(0)))
            End If
        End If
        On Error GoTo 0

        ' Get nights (Column AD)
        On Error Resume Next
        nights = CLng(wsSource.Cells(i, "AD").Value)
        If Err.Number <> 0 Then nights = 0
        On Error GoTo 0

        ' Calculate TDF (roomCategory already retrieved earlier)
        tdfCharge = CalculateTDF(roomCategory, nights)

        ' Get SHARE_AMOUNT_PER_STAY (Column AI) - this is the main amount
        On Error Resume Next
        shareAmount = CDbl(wsSource.Cells(i, "AI").Value)
        If Err.Number <> 0 Then shareAmount = 0
        On Error GoTo 0

        ' NET is Column AI * 1.225 (Column I in ENTERED ON)
        netAmount = shareAmount * 1.225

        ' Calculate TOTAL (NET + TDF)
        totalAmount = netAmount + tdfCharge

        ' Calculate ADR (Column AI / NIGHTS)
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

        ' Column I: NET (with bold formatting and color)
        wsTarget.Cells(targetRow, "I").Value = netAmount
        wsTarget.Cells(targetRow, "I").NumberFormat = "0"
        wsTarget.Cells(targetRow, "I").Font.Bold = True
        ' Add #00FFCC background color if value is not zero
        If netAmount <> 0 Then
            wsTarget.Cells(targetRow, "I").Interior.Color = RGB(0, 255, 204)
        End If

        ' Column J: TOTAL
        wsTarget.Cells(targetRow, "J").Value = totalAmount
        wsTarget.Cells(targetRow, "J").NumberFormat = "0"

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

        ' Add borders to all cells in this row (columns A to V)
        wsTarget.Range(wsTarget.Cells(targetRow, "A"), wsTarget.Cells(targetRow, "V")).Borders.LineStyle = xlContinuous
        wsTarget.Range(wsTarget.Cells(targetRow, "A"), wsTarget.Cells(targetRow, "V")).Borders.Weight = xlThin

        ' Mark this ID as processed
        If Len(resvID) > 0 Then
            existingIDs(resvID) = True
        End If

        processCount = processCount + 1
        targetRow = targetRow + 1

NextRow:
    Next i

    ' Sort the ENTERED ON sheet by Column M (C_T_S_NAME) in descending order
    If processCount > 0 Then
        Dim sortLastRow As Long
        sortLastRow = wsTarget.Cells(wsTarget.Rows.Count, "A").End(xlUp).Row

        If sortLastRow > 1 Then
            ' Sort Column M in descending order
            With wsTarget.Sort
                .SortFields.Clear
                .SortFields.Add Key:=wsTarget.Range("M2:M" & sortLastRow), _
                    SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
                .SetRange wsTarget.Range("A1:V" & sortLastRow)
                .Header = xlYes
                .MatchCase = False
                .Orientation = xlTopToBottom
                .SortMethod = xlPinYin
                .Apply
            End With
        End If
    End If

    ' Restore settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

    ' Automatically refresh all sheets, pivot tables, and charts
    Call RefreshAllWorkbookData

    ' Save the workbook in its current location
    On Error Resume Next
    wb.Save
    If Err.Number <> 0 Then
        MsgBox "Warning: Could not save file. Error: " & Err.Description, vbExclamation
        Err.Clear
    End If
    On Error GoTo 0

    ' Show completion message with processing stats
    MsgBox "Processing complete!" & vbCrLf & vbCrLf & _
           "Processed: " & processCount & " rows" & vbCrLf & _
           "Skipped: " & skipCount & " rows" & vbCrLf & vbCrLf & _
           "File saved successfully.", vbInformation, "Complete"

End Sub

Function CalculateTDF(roomCategory As String, nights As Long) As Double
    '=====================================================================
    ' Calculate Tourism Dirham Fee (TDF)
    ' Rules:
    '   - 1BA: 20 AED per night (capped at 600 AED for 30+ nights)
    '   - 2BA: 40 AED per night (capped at 1200 AED for 30+ nights)
    '=====================================================================

    Dim tdf As Double
    Dim ratePerNight As Double
    Dim maxNights As Long

    tdf = 0
    maxNights = 30

    If nights <= 0 Then
        CalculateTDF = 0
        Exit Function
    End If

    ' Determine rate based on room category
    If InStr(1, roomCategory, "2BA", vbTextCompare) > 0 Then
        ' 2 Bedroom Apartment: 40 AED per night
        ratePerNight = 40
    Else
        ' 1 Bedroom Apartment (or other): 20 AED per night
        ratePerNight = 20
    End If

    ' Calculate TDF with 30-night cap
    If nights <= maxNights Then
        tdf = nights * ratePerNight
    Else
        tdf = maxNights * ratePerNight ' Cap at 30 nights
    End If

    CalculateTDF = tdf
End Function

Sub FixSpilloverRows(ws As Worksheet)
    '=====================================================================
    ' Fix spillover rows where data continues on the next line
    ' Pattern: Rows starting with "[Letter]- " (e.g., "T- ", "S- ", "C- ")
    ' Also fixes title spillover in Column R (e.g., ",Mr.", ",Ms.")
    '=====================================================================

    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String
    Dim departureValue As String

    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row

    ' First pass: Fix title spillover where entire row spills to next line
    ' Pattern: Row ends with incomplete name like "Pereyda,Enrique Antonio"
    ' Next row starts with ",Mr." in column A followed by actual data in columns B onwards
    For i = lastRow To 2 Step -1
        cellValue = Trim(CStr(ws.Cells(i, "A").Value))

        ' Check if row starts with a title (,Mr., ,Ms., etc.)
        If Len(cellValue) > 0 And Left(cellValue, 1) = "," Then
            If i > 1 Then
                ' This is a spillover row - append title to previous row's name (Column Q)
                ws.Cells(i - 1, "Q").Value = ws.Cells(i - 1, "Q").Value & cellValue

                ' Move all data from this spillover row to the previous row
                ' The data in columns B onwards should go to columns R onwards
                ws.Cells(i - 1, "R").Value = ws.Cells(i, "B").Value   ' DEPARTURE
                ws.Cells(i - 1, "S").Value = ws.Cells(i, "C").Value   ' PERSONS
                ws.Cells(i - 1, "T").Value = ws.Cells(i, "D").Value   ' GROUP_NAME
                ws.Cells(i - 1, "U").Value = ws.Cells(i, "E").Value   ' NO_OF_ROOMS
                ws.Cells(i - 1, "V").Value = ws.Cells(i, "F").Value   ' ROOM_CATEGORY
                ws.Cells(i - 1, "W").Value = ws.Cells(i, "G").Value   ' RATE_CODE
                ws.Cells(i - 1, "X").Value = ws.Cells(i, "H").Value   ' INSERT_USER
                ws.Cells(i - 1, "Y").Value = ws.Cells(i, "I").Value   ' INSERT_DATE
                ws.Cells(i - 1, "Z").Value = ws.Cells(i, "J").Value   ' GUARANTEE_CODE_DESC
                ws.Cells(i - 1, "AA").Value = ws.Cells(i, "K").Value  ' COMPANY_NAME
                ws.Cells(i - 1, "AB").Value = ws.Cells(i, "L").Value  ' TRAVEL_AGENT_NAME
                ws.Cells(i - 1, "AC").Value = ws.Cells(i, "M").Value  ' ARRIVAL
                ws.Cells(i - 1, "AD").Value = ws.Cells(i, "N").Value  ' NIGHTS
                ws.Cells(i - 1, "AE").Value = ws.Cells(i, "O").Value  ' COMP_HOUSE_YN
                ws.Cells(i - 1, "AF").Value = ws.Cells(i, "P").Value  ' SHARE_AMOUNT
                ws.Cells(i - 1, "AG").Value = ws.Cells(i, "Q").Value  ' C_T_S_NAME
                ws.Cells(i - 1, "AH").Value = ws.Cells(i, "R").Value  ' SHORT_RESV_STATUS
                ws.Cells(i - 1, "AI").Value = ws.Cells(i, "S").Value  ' SHARE_AMOUNT_PER_STAY

                ' Delete the spillover row
                ws.Rows(i).Delete
            End If
        End If
    Next i

    ' Recalculate lastRow after deletions
    lastRow = ws.Cells(ws.Rows.Count, "M").End(xlUp).Row

    ' Second pass: Fix title spillover in Column R (DEPARTURE)
    For i = 2 To lastRow
        departureValue = Trim(CStr(ws.Cells(i, "R").Value))

        ' Check if Column R contains a title instead of a date
        ' Titles look like: ",Mr.", ",Ms.", ",Mrs.", etc.
        If Len(departureValue) > 0 And Left(departureValue, 1) = "," Then
            ' This is a title spillover - append to name in Column Q and shift data left
            ws.Cells(i, "Q").Value = ws.Cells(i, "Q").Value & departureValue

            ' Shift all data from Column S onward back to Column R
            ' R = DEPARTURE, S = PERSONS, T = GROUP_NAME, U = NO_OF_ROOMS, V = ROOM_CATEGORY_LABEL
            ' W = RATE_CODE, X = INSERT_USER, Y = INSERT_DATE, Z = GUARANTEE_CODE_DESC
            ' AA = COMPANY_NAME, AB = TRAVEL_AGENT_NAME, AC = ARRIVAL, AD = NIGHTS
            ' AE = COMP_HOUSE_YN, AF = SHARE_AMOUNT, AG = C_T_S_NAME, AH = SHORT_RESV_STATUS
            ' AI = SHARE_AMOUNT_PER_STAY

            ws.Cells(i, "R").Value = ws.Cells(i, "S").Value   ' DEPARTURE from PERSONS
            ws.Cells(i, "S").Value = ws.Cells(i, "T").Value   ' PERSONS from GROUP_NAME
            ws.Cells(i, "T").Value = ws.Cells(i, "U").Value   ' GROUP_NAME from NO_OF_ROOMS
            ws.Cells(i, "U").Value = ws.Cells(i, "V").Value   ' NO_OF_ROOMS from ROOM_CATEGORY
            ws.Cells(i, "V").Value = ws.Cells(i, "W").Value   ' ROOM_CATEGORY from RATE_CODE
            ws.Cells(i, "W").Value = ws.Cells(i, "X").Value   ' RATE_CODE from INSERT_USER
            ws.Cells(i, "X").Value = ws.Cells(i, "Y").Value   ' INSERT_USER from INSERT_DATE
            ws.Cells(i, "Y").Value = ws.Cells(i, "Z").Value   ' INSERT_DATE from GUARANTEE_CODE_DESC
            ws.Cells(i, "Z").Value = ws.Cells(i, "AA").Value  ' GUARANTEE_CODE_DESC from COMPANY_NAME
            ws.Cells(i, "AA").Value = ws.Cells(i, "AB").Value ' COMPANY_NAME from TRAVEL_AGENT_NAME
            ws.Cells(i, "AB").Value = ws.Cells(i, "AC").Value ' TRAVEL_AGENT_NAME from ARRIVAL
            ws.Cells(i, "AC").Value = ws.Cells(i, "AD").Value ' ARRIVAL from NIGHTS
            ws.Cells(i, "AD").Value = ws.Cells(i, "AE").Value ' NIGHTS from COMP_HOUSE_YN
            ws.Cells(i, "AE").Value = ws.Cells(i, "AF").Value ' COMP_HOUSE_YN from SHARE_AMOUNT
            ws.Cells(i, "AF").Value = ws.Cells(i, "AG").Value ' SHARE_AMOUNT from C_T_S_NAME
            ws.Cells(i, "AG").Value = ws.Cells(i, "AH").Value ' C_T_S_NAME from SHORT_RESV_STATUS
            ws.Cells(i, "AH").Value = ws.Cells(i, "AI").Value ' SHORT_RESV_STATUS from SHARE_AMOUNT_PER_STAY
            ws.Cells(i, "AI").Value = ws.Cells(i, "AJ").Value ' SHARE_AMOUNT_PER_STAY from column AJ
            ws.Cells(i, "AJ").Value = ""                      ' Clear column AJ
        End If
    Next i

    ' Third pass: Process company name spillover rows (rows starting with "[Letter]- ")
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

Sub ClearEnteredOnSheet()
    '=====================================================================
    ' Helper macro to clear all data from ENTERED ON sheet (keeps headers)
    '=====================================================================
    Dim wb As Workbook
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim response As VbMsgBoxResult

    Set wb = ThisWorkbook

    On Error Resume Next
    Set ws = wb.Worksheets("ENTERED ON")
    On Error GoTo 0

    If ws Is Nothing Then
        MsgBox "Error: 'ENTERED ON' sheet not found!", vbCritical
        Exit Sub
    End If

    ' Find last row with data
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then
        lastRow = ws.Cells(ws.Rows.Count, "S").End(xlUp).Row
    End If

    ' Confirm before clearing
    response = MsgBox("This will clear ALL data from the ENTERED ON sheet (rows 2 to " & lastRow & ")." & vbCrLf & vbCrLf & _
                      "Are you sure you want to continue?", vbYesNo + vbExclamation, "Confirm Clear")

    If response = vbNo Then
        MsgBox "Operation cancelled.", vbInformation
        Exit Sub
    End If

    ' Clear all data (keep row 1 header)
    If lastRow > 1 Then
        On Error Resume Next
        ws.Range("A2:V" & lastRow).ClearContents
        ws.Range("A2:V" & lastRow).Borders.LineStyle = xlNone
        ' Clear specific formatting but PRESERVE conditional formatting
        With ws.Range("A2:V" & lastRow)
            .Interior.ColorIndex = xlNone ' Clear fill colors
            .Font.Bold = False
            .Font.ColorIndex = xlAutomatic
            .NumberFormat = "General"
        End With
        If Err.Number <> 0 Then
            MsgBox "Error clearing data: " & Err.Description, vbCritical
            Exit Sub
        End If
        On Error GoTo 0
    End If

End Sub

Sub RefreshAllWorkbookData()
    '=====================================================================
    ' Macro: RefreshAllWorkbookData
    ' Purpose: Refresh all sheets, pivot tables, charts, and calculations
    ' Features:
    '   - Recalculate all formulas
    '   - Refresh all pivot tables
    '   - Refresh all data connections
    '   - Refresh all charts
    '   - Update all links
    '=====================================================================

    Dim wb As Workbook
    Dim ws As Worksheet
    Dim pt As PivotTable
    Dim cht As ChartObject
    Dim pivotCount As Long
    Dim chartCount As Long
    Dim startTime As Double

    ' Initialize
    Set wb = ThisWorkbook
    pivotCount = 0
    chartCount = 0
    startTime = Timer

    ' Turn off screen updating for performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error Resume Next

    ' 1. Recalculate all formulas
    Application.CalculateFull

    ' 2. Refresh all pivot tables in all sheets
    For Each ws In wb.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshTable
            pivotCount = pivotCount + 1
        Next pt
    Next ws

    ' 3. Refresh all charts in all sheets
    For Each ws In wb.Worksheets
        For Each cht In ws.ChartObjects
            cht.Chart.Refresh
            chartCount = chartCount + 1
        Next cht
    Next ws

    ' 4. Refresh all data connections (queries, external data)
    Dim conn As WorkbookConnection
    For Each conn In wb.Connections
        conn.Refresh
    Next conn

    ' 5. Update all links (if any)
    wb.UpdateLink Name:=wb.LinkSources(xlExcelLinks)

    On Error GoTo 0

    ' Restore settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    Application.EnableEvents = True

End Sub
