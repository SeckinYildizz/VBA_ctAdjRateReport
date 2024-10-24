Sub ctAdjRep()
'Creates adjustment report from visit maintenance, service note and Sandata adjustment data.

    'Determine the variables
    Dim i1, i10, i2, i3, i4, lRow1, lRow2, lRow3, lRow4 As Integer
    Dim vms, sns, sas, ts As Worksheet
    
    'Set some of the variables
    Set sas = ThisWorkbook.Sheets(2)
    Set vms = ThisWorkbook.Sheets(3)
    Set sns = ThisWorkbook.Sheets(4)
    Sheets.Add after:=Sheets(Sheets.Count)
    Sheets(Sheets.Count).Name = "AdjRep"
    Set ts = Sheets(Sheets.Count)
    lRow1 = sas.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row - 1
    lRow2 = vms.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    lRow3 = sns.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
    
    'Adjust dates in the visit maintenance sheet
    For i2 = lRow2 To 2 Step -1
        vms.Cells(i2, "C").NumberFormat = "@"
        vms.Cells(i2, "C").Value = Left(vms.Cells(i2, "C").Value, 10)
    Next i2
    
    'Create a column and format PTP names as C2S in the adjustment sheet
    With sas
        .Columns("F:G").Insert shift:=xlRight
        For i1 = lRow1 To 5 Step -1
            .Cells(i1, "F").Value = .Cells(i1, "E").Value & ", " & .Cells(i1, "D").Value
            'Format the date column
            .Cells(i1, "A").NumberFormat = "@"
            i10 = InStr(.Cells(i1, "A").Value, "/")
            If i10 = 2 Then .Cells(i1, "A").Value = "0" & .Cells(i1, "A").Value
            i10 = InStr(4, .Cells(i1, "A").Value, "/")
            If i10 = 5 Then .Cells(i1, "A").Value = Left(.Cells(i1, "A").Value, 3) & "0" & Right(.Cells(i1, "A"), 6)
            
            For i2 = lRow2 To 2 Step -1
                If .Cells(i1, "F").Value = vms.Cells(i2, "AI").Value And .Cells(i1, "A").Value = vms.Cells(i2, "C").Value Then
                    .Cells(i1, "G").Value = vms.Cells(i2, "L").Value
                End If
            Next i2
        Next i1
        'Copy unique DSP names into the target sheet
        .Range("G5:G" & lRow1).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ts.Range("A2"), Unique:=True
    End With
    
    'Find the last row of the target sheet
    lRow4 = ts.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
                    
    With sns
        'Delete notes that require EVV
        For i3 = lRow3 To 2 Step -1
            If InStr(.Cells(i3, "L"), "NER") = 0 Or .Cells(i3, "G").Value <> " Approved" Then
                .Rows(i3).EntireRow.Delete
            End If
        Next i3
        lRow3 = .Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        
        'Paste DSP names into the target sheet
        .Range("C2:C" & lRow3).AdvancedFilter Action:=xlFilterCopy, CopyToRange:=ts.Range("A" & lRow4 + 1), Unique:=True
    End With
    
    With ts
        'Remove duplicates and sort
        .Range("A:A").RemoveDuplicates Columns:=1, Header:=xlYes
        lRow4 = ts.Cells.Find("*", searchorder:=xlByRows, searchdirection:=xlPrevious).Row
        .Range("A:A").Sort key1:=.Range("A2"), Order1:=xlAscending, Header:=xlYes
        
        'Format the sheet
        .Cells.WrapText = False
        .Cells.Font.Size = 11
        .Cells.Font.FontStyle = "Calibri"
        
        'Put column names and format them
        .Range("A1:D1").Value = Array("DSP", "Visit #", "Adj #", "Adj Rate")
        .Range("A1:D1").Font.Bold = True
        .Range("A1:D1").Borders(xlEdgeBottom).LineStyle = xlContinuous
        .Range("A1:D1").Interior.Color = RGB(255, 255, 102)
    
        'Fill the report
        For i4 = 2 To lRow4
            .Cells(i4, "C").Value = 0
            For i1 = 5 To lRow1
                'Count the DSP names in the Sandata report
                If .Cells(i4, "A").Value = sas.Cells(i1, "G").Value Then
                    .Cells(i4, "B").Value = .Cells(i4, "B").Value + 1
                End If
                'Count the adjustments
                If .Cells(i4, "A").Value = sas.Cells(i1, "G").Value And sas.Cells(i1, "I").Value = "M" Then
                    .Cells(i4, "C").Value = .Cells(i4, "C").Value + 1
                End If
            Next i1
            'Count the DSP names in the not EVV required services
            For i3 = 2 To lRow3
                If .Cells(i4, "A").Value = sns.Cells(i3, "C").Value Then
                    .Cells(i4, "B").Value = .Cells(i4, "B").Value + 1
                End If
            Next i3
            
            'Fill the adjustment rates
            .Cells(i4, "D").Value = .Cells(i4, "C").Value / .Cells(i4, "B").Value
            .Cells(i4, "D").NumberFormat = "0.00%"
        Next i4
    End With
End Sub
