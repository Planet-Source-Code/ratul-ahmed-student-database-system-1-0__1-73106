Attribute VB_Name = "List_View_mod"
'-----------------------------------------------------------------------------
Public Function lstConfig(lstdb As ListView)
'-----------------------------------------------------------------------------

lstdb.View = lvwReport
lstdb.Gridlines = True
lstdb.FullRowSelect = True

    With lstdb
    
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "ID", .Width * 0.12, 0
        .ColumnHeaders.Add 2, , "Name", .Width * 0.3, 2
        .ColumnHeaders.Add 3, , "Roll", .Width * 0.165, 2
        .ColumnHeaders.Add 4, , "Department", .Width * 0.2, 2
        .ColumnHeaders.Add 5, , "Grade", .Width * 0.2, 2
              
    End With

End Function

'-----------------------------------------------------------------------------
Public Function lstConfig_result(lstdb As ListView)
'-----------------------------------------------------------------------------

lstdb.View = lvwReport
lstdb.Gridlines = True
lstdb.FullRowSelect = True

    With lstdb
    
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, , "ID", .Width * 0.12, 0
        .ColumnHeaders.Add 2, , "Name", .Width * 0.3, 2
        .ColumnHeaders.Add 3, , "Roll", .Width * 0.165, 2
        .ColumnHeaders.Add 4, , "Department", .Width * 0.2, 2
        .ColumnHeaders.Add 5, , "Computer Architecture", .Width * 0.4, 2
        .ColumnHeaders.Add 6, , "Microprocessor", .Width * 0.4, 2
        .ColumnHeaders.Add 7, , "DataBase Management", .Width * 0.4, 2
        .ColumnHeaders.Add 8, , "Visual Programming", .Width * 0.4, 2
        .ColumnHeaders.Add 9, , "Data Communication Fund.", .Width * 0.4, 2
        .ColumnHeaders.Add 10, , "Environmental Manag.", .Width * 0.4, 2
        .ColumnHeaders.Add 11, , "Book Keeping", .Width * 0.4, 2
        .ColumnHeaders.Add 12, , "Business Organization", .Width * 0.4, 2
        .ColumnHeaders.Add 13, , "Score", .Width * 0.2, 2
        .ColumnHeaders.Add 14, , "Grade", .Width * 0.2, 2
        .ColumnHeaders.Add 15, , "cGPA", .Width * 0.2, 2
              
    End With

End Function



Public Function search(strText As String, listv As ListView)
Dim itm As ListItem
Dim itmName As ListItem

    With listv
        Set itm = .FindItem(strText, lvwText, , lvwPartial)
            If Not itm Is Nothing Then
                .ListItems(itm.Index).selected = True
                .SetFocus
            End If
    End With

        Set itm = Nothing
End Function
