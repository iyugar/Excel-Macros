'Deletes table rows for a given value. Requires the value to search, the table as a ListObject and the column name where to perform the search.

Function deleteTableRows(columnName As String, searchValue As Variant, dataTable As ListObject)
   
    colNumber = dataTable.ListColumns(columnName).DataBodyRange.Column
    
    With dataTable.DataBodyRange
        lastRow = .Rows.Count
        
        For i = lastRow To 1 Step -1
            tableValue = .Rows(i).Cells(1, colNumber)
            If tableValue = searchValue Then
                .Rows(i).Delete
            End If
        Next
    End With
End Function
