Function copyHeader(sht As Worksheet, startRow As Integer, searchColumn As String, pasteTargetColumn As String, searchTextPattern As String)
'Finds all text that contains as pattern in a given data column and pastes the found text in another column for all rows between the found text row and then next found text row
'Data Input Sample - Find "Header" keyword
'A   B
'    Header1
'    Data1
'    Data2
'    Data3
'    Header2
'    Data1

'Data Output Sample
'A          B
'           Header1
'Header1    Data1
'Header1    Data2
'Header1    Data3
'           Header2
'Header2    Data1

shtLastRow = sht.Cells(sht.Rows.Count, searchColumn).End(xlUp).Row
Dim pairsArr()
Dim pair(1)
counter = 0

For r = startRow To shtLastRow
    cellText = sht.Cells(r, searchColumn)
    If InStr(1, cellText, searchTextPattern) Then
    
        pair(0) = cellText
        pair(1) = r
        
        ReDim Preserve pairsArr(counter)
        pairsArr(counter) = pair
        counter = counter + 1
    End If
Next

For i = 0 To UBound(pairsArr)
    startRow = pairsArr(i)(1) + 1
    If i = UBound(pairsArr) Then
        lastRow = shtLastRow
    Else
        lastRow = pairsArr(i + 1)(1) - 1
    End If
    For r = startRow To lastRow
        sht.Cells(r, pasteTargetColumn) = pairsArr(i)(0)
    Next
Next
End Function
