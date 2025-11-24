Sub ProcessEnteredOnReport()
    Dim wb As Workbook
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim lastRow As Long
    Dim i As Long, targetRow As Long
    Dim arrData() As Variant
    Dim roomCategory As String
    Dim nights As Long
    Dim tdfCharge As Double
    Dim amountWithoutTDF As Double
    Dim amountWithTDF As Double
    Dim shareAmount As Double
    Dim adr As Double

    ' Set this workbook
    Set wb = ThisWorkbook

    ' Check if "Entered On" sheet exists, if not create it
    On Error Resume Next
    Set wsTarget = wb.Worksheets("Entered On")
    On Error GoTo 0

    If wsTarget Is Nothing Then
        Set wsTarget = wb.Worksheets.Add
        wsTarget.Name = "Entered On"
    Else
        ' Clear existing data except headers
        If wsTarget.Cells(1, 1).Value <> "" Then
            lastRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row
            If lastRow > 1 Then
                wsTarget.Rows("2:" & lastRow).Delete
            End If
        End If
    End If

    ' Import CSV file
    Dim csvPath As String
    csvPath = wb.Path & "\resenteredon102243710-lpo.csv"

    ' Open CSV and load to temporary sheet
    Dim wbCSV As Workbook
    Set wbCSV = Workbooks.Open(Filename:=csvPath, Local:=True)
    Set wsSource = wbCSV.Worksheets(1)

    ' Fix spillover rows first
    Call FixSpilloverRows(wsSource)

    ' Set up target sheet headers
    Call SetupHeaders(wsTarget)

    ' Get last row in source
    lastRow = wsSource.Cells(wsSource.Rows.Count, 1).End(xlUp).Row

    ' Initialize target row
    targetRow = 2

    ' Process each row
    For i = 2 To lastRow
        ' Skip if first column is empty or starts with specific prefixes that should have been merged
        If Len(Trim(wsSource.Cells(i, 1).Value)) = 0 Then
            GoTo NextRow
        End If

        ' Get room category from column 22 (ROOM_CATEGORY_LABEL)
        roomCategory = Trim(wsSource.Cells(i, 22).Value)

        ' Get nights from column 30 (NIGHTS)
        On Error Resume Next
        nights = CLng(wsSource.Cells(i, 30).Value)
        If Err.Number <> 0 Then nights = 0
        On Error GoTo 0

        ' Calculate TDF based on room category and nights
        tdfCharge = CalculateTDF(roomCategory, nights)

        ' Get share amount from column 32 (SHARE_AMOUNT) - this becomes column P
        On Error Resume Next
        shareAmount = CDbl(wsSource.Cells(i, 32).Value)
        If Err.Number <> 0 Then shareAmount = 0
        On Error GoTo 0

        ' Get amount from column 35 (SHARE_AMOUNT_PER_STAY) - this becomes Share column (AI)
        On Error Resume Next
        amountWithoutTDF = CDbl(wsSource.Cells(i, 35).Value)
        If Err.Number <> 0 Then amountWithoutTDF = 0
        On Error GoTo 0

        ' Calculate amount with TDF (column J)
        amountWithTDF = amountWithoutTDF + tdfCharge

        ' Calculate ADR (average daily rate)
        If nights > 0 Then
            adr = shareAmount / nights
        Else
            adr = 0
        End If

        ' Map data to target sheet
        ' Column A: RESORT (arrival date from source column 1)
        wsTarget.Cells(targetRow, 1).Value = wsSource.Cells(i, 1).Value

        ' Column B: GRPBY_DISP1 (source column 2)
        wsTarget.Cells(targetRow, 2).Value = wsSource.Cells(i, 2).Value

        ' Column C: RESV_NAME_ID (source column 13)
        wsTarget.Cells(targetRow, 3).Value = wsSource.Cells(i, 13).Value

        ' Column D: GUARANTEE_CODE (source column 14)
        wsTarget.Cells(targetRow, 4).Value = wsSource.Cells(i, 14).Value

        ' Column E: RESV_STATUS (source column 15)
        wsTarget.Cells(targetRow, 5).Value = wsSource.Cells(i, 15).Value

        ' Column F: ROOM (source column 16)
        wsTarget.Cells(targetRow, 6).Value = wsSource.Cells(i, 16).Value

        ' Column G: FULL_NAME (source column 17)
        wsTarget.Cells(targetRow, 7).Value = wsSource.Cells(i, 17).Value

        ' Column H: DEPARTURE (source column 18)
        wsTarget.Cells(targetRow, 8).Value = wsSource.Cells(i, 18).Value

        ' Column I: NET (amount without TDF)
        wsTarget.Cells(targetRow, 9).Value = amountWithoutTDF

        ' Column J: TOTAL (amount with TDF)
        wsTarget.Cells(targetRow, 10).Value = amountWithTDF

        ' Column K: PERSONS (source column 19)
        wsTarget.Cells(targetRow, 11).Value = wsSource.Cells(i, 19).Value

        ' Column L: GROUP_NAME (source column 20)
        wsTarget.Cells(targetRow, 12).Value = wsSource.Cells(i, 20).Value

        ' Column M: NO_OF_ROOMS (source column 21)
        wsTarget.Cells(targetRow, 13).Value = wsSource.Cells(i, 21).Value

        ' Column N: ROOM_CATEGORY_LABEL (source column 22)
        wsTarget.Cells(targetRow, 14).Value = roomCategory

        ' Column O: RATE_CODE (source column 23)
        wsTarget.Cells(targetRow, 15).Value = wsSource.Cells(i, 23).Value

        ' Column P: SHARE (from source column 32 - SHARE_AMOUNT)
        wsTarget.Cells(targetRow, 16).Value = shareAmount

        ' Column Q: INSERT_USER (source column 24)
        wsTarget.Cells(targetRow, 17).Value = wsSource.Cells(i, 24).Value

        ' Column R: INSERT_DATE (source column 25)
        wsTarget.Cells(targetRow, 18).Value = wsSource.Cells(i, 25).Value

        ' Column S: GUARANTEE_CODE_DESC (source column 26)
        wsTarget.Cells(targetRow, 19).Value = wsSource.Cells(i, 26).Value

        ' Column T: COMPANY_NAME (source column 27)
        wsTarget.Cells(targetRow, 20).Value = wsSource.Cells(i, 27).Value

        ' Column U: TRAVEL_AGENT_NAME (source column 28 + 33 if merged)
        wsTarget.Cells(targetRow, 21).Value = wsSource.Cells(i, 28).Value

        ' Column V: ARRIVAL (source column 29)
        wsTarget.Cells(targetRow, 22).Value = wsSource.Cells(i, 29).Value

        ' Column W: NIGHTS (source column 30)
        wsTarget.Cells(targetRow, 23).Value = nights

        ' Column X: COMP_HOUSE_YN (source column 31)
        wsTarget.Cells(targetRow, 24).Value = wsSource.Cells(i, 31).Value

        ' Column Y: TDF Charge (calculated)
        wsTarget.Cells(targetRow, 25).Value = tdfCharge

        ' Column Z: ADR (calculated)
        wsTarget.Cells(targetRow, 26).Value = adr

        ' Column AA: C_T_S_NAME (source column 33)
        wsTarget.Cells(targetRow, 27).Value = wsSource.Cells(i, 33).Value

        ' Column AB: SHORT_RESV_STATUS (source column 34)
        wsTarget.Cells(targetRow, 28).Value = wsSource.Cells(i, 34).Value

        ' Column AC onwards - add remaining fields as needed

        targetRow = targetRow + 1

NextRow:
    Next i

    ' Close CSV workbook without saving
    wbCSV.Close SaveChanges:=False

    ' Format the target sheet
    Call FormatSheet(wsTarget, targetRow - 1)

    MsgBox "Processing complete! " & (targetRow - 2) & " records processed.", vbInformation

End Sub

Function CalculateTDF(roomCategory As String, nights As Long) As Double
    ' TDF Logic:
    ' - 1BA: 20 AED per night for first 30 nights, then 600 AED total
    ' - 2BA: 40 AED per night for first 30 nights, then 1200 AED total
    ' - Reservations > 30 nights: Cap at 600 (1BA) or 1200 (2BA)

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
    ' Fix rows where data spills over to next row
    ' Pattern: Row starts with "T- ", "S- ", "C- " etc. in column A
    ' These should be merged with previous row column AG (27) and AH (28)

    Dim lastRow As Long
    Dim i As Long
    Dim cellValue As String

    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row

    ' Process from bottom to top to avoid row shifting issues
    For i = lastRow To 2 Step -1
        cellValue = Trim(ws.Cells(i, 1).Value)

        ' Check if this is a spillover row (starts with letter- pattern)
        If Len(cellValue) > 2 Then
            If Mid(cellValue, 2, 1) = "-" And cellValue Like "[A-Z]- *" Then
                ' This is a spillover row
                ' Copy to previous row's column 28 (Travel Agent Name continuation)
                If i > 1 Then
                    ' Append to column 28 of previous row
                    ws.Cells(i - 1, 28).Value = ws.Cells(i - 1, 28).Value & " " & cellValue

                    ' Move column B and C of spillover row to columns 34 and 35 of previous row
                    If Len(Trim(ws.Cells(i, 2).Value)) > 0 Then
                        ws.Cells(i - 1, 34).Value = ws.Cells(i, 2).Value
                    End If

                    If Len(Trim(ws.Cells(i, 3).Value)) > 0 Then
                        ws.Cells(i - 1, 35).Value = ws.Cells(i, 3).Value
                    End If

                    ' Delete the spillover row
                    ws.Rows(i).Delete
                End If
            End If
        End If
    Next i
End Sub

Sub SetupHeaders(ws As Worksheet)
    ' Set up column headers
    ws.Cells(1, 1).Value = "ARRIVAL"
    ws.Cells(1, 2).Value = "DATE"
    ws.Cells(1, 3).Value = "RESV_NAME_ID"
    ws.Cells(1, 4).Value = "GUARANTEE_CODE"
    ws.Cells(1, 5).Value = "RESV_STATUS"
    ws.Cells(1, 6).Value = "ROOM"
    ws.Cells(1, 7).Value = "FULL_NAME"
    ws.Cells(1, 8).Value = "DEPARTURE"
    ws.Cells(1, 9).Value = "NET"
    ws.Cells(1, 10).Value = "TOTAL"
    ws.Cells(1, 11).Value = "PERSONS"
    ws.Cells(1, 12).Value = "GROUP_NAME"
    ws.Cells(1, 13).Value = "NO_OF_ROOMS"
    ws.Cells(1, 14).Value = "ROOM_CATEGORY"
    ws.Cells(1, 15).Value = "RATE_CODE"
    ws.Cells(1, 16).Value = "SHARE"
    ws.Cells(1, 17).Value = "INSERT_USER"
    ws.Cells(1, 18).Value = "INSERT_DATE"
    ws.Cells(1, 19).Value = "GUARANTEE_DESC"
    ws.Cells(1, 20).Value = "COMPANY_NAME"
    ws.Cells(1, 21).Value = "TRAVEL_AGENT"
    ws.Cells(1, 22).Value = "ARRIVAL_DATE"
    ws.Cells(1, 23).Value = "NIGHTS"
    ws.Cells(1, 24).Value = "COMP_HOUSE"
    ws.Cells(1, 25).Value = "TDF"
    ws.Cells(1, 26).Value = "ADR"
    ws.Cells(1, 27).Value = "SOURCE"
    ws.Cells(1, 28).Value = "STATUS"

    ' Format header row
    With ws.Rows(1)
        .Font.Bold = True
        .Interior.Color = RGB(217, 217, 217)
        .HorizontalAlignment = xlCenter
    End With
End Sub

Sub FormatSheet(ws As Worksheet, lastRow As Long)
    ' Auto-fit columns
    ws.Columns.AutoFit

    ' Format number columns
    ws.Range("I2:I" & lastRow).NumberFormat = "#,##0.00"
    ws.Range("J2:J" & lastRow).NumberFormat = "#,##0.00"
    ws.Range("P2:P" & lastRow).NumberFormat = "#,##0.00"
    ws.Range("Y2:Y" & lastRow).NumberFormat = "#,##0.00"
    ws.Range("Z2:Z" & lastRow).NumberFormat = "#,##0.00"

    ' Add borders
    With ws.Range("A1:AB" & lastRow)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With

    ' Freeze header row
    ws.Rows(2).Select
    ActiveWindow.FreezePanes = True
    ws.Cells(1, 1).Select
End Sub
