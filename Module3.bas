Attribute VB_Name = "Module3"
Sub test2()

    Dim c_row As Long, c_col As Long, first_item_row As Long, first_item_col As Long, last_row As Long, last_col As Long
    Dim store_unique_num As String
    Dim is_done As Boolean
    
    first_item_row = Cells.Find(What:="*").Row
    first_item_col = Cells.Find(What:="*").Column
    
    last_row = Cells.Find(What:="*", SearchOrder:=xlRows, SearchDirection:=xlPrevious, LookIn:=xlValues).Row
    last_col = Cells.Find(What:="*", SearchOrder:=xlColumns, SearchDirection:=xlPrevious, LookIn:=xlValues).Column
    
    
    store_num_row = -1
    store_num_col = -1
    
    is_done = False
       

    For c_row = first_item_row To last_row
        For c_col = first_item_col To last_col
            If store_num_row = -1 And StrComp(Cells(c_row, c_col).Value, "Store Number", vbTextCompare) = 0 Then
                store_num_row = c_row
                store_num_col = c_col
                
                is_done = True
                Exit For
            End If
        
        Next c_col
        
        If is_done Then
            Exit For
        End If
                   
    Next c_row
    
    
    '
    ' Populating the correct store number (first five digits)
    '

    Cells(store_num_row, last_col + 2).Value = "Unique_store_num"
    Cells(store_num_row, last_col + 4).Value = "Identifier"
    
    Dim i As Long
    Dim j As Long
    
    For i = store_num_row + 1 To last_row
        j = Left(Cells(i, store_num_col), 5)
        Cells(i, last_col + 2).Value = j
    Next i
    
    

    ' Get all the columns with header that starts with "Answer"
    
    Dim col As Long
    Dim answer_cols() As Long
    Dim count As Long
    count = 0
    
    For col = first_item_col To last_col
        If InStr(1, Cells(first_item_row, col), "Answer") = 1 Then
            count = count + 1
            ReDim Preserve answer_cols(1 To count)
            answer_cols(count) = col
            
        End If
    Next col
    
    For i = LBound(answer_cols) To UBound(answer_cols)
        MsgBox answer_cols(i)

    Next


    '
    ' Check if the store num are same if so check else ski p
    '

    For c_row = (first_item_row + 2) To last_row
        If Cells(c_row - 1, last_col + 2).Value = Cells(c_row, last_col + 2).Value Then
            comparison c_row, last_col, answer_cols()
        Else
            Cells(c_row, last_col + 4) = "Unique"
        End If
        
    Next c_row
    
    
End Sub


Sub comparison(c_row As Long, last_col As Long, answer_cols() As Long)
    '
    ' The first_item_row would serve as the rows for the data headers
    '

    Dim groupd_cols() As Long
    Dim c_col As Long
    Dim i As Long
    Dim j As Long
    Dim is_done As Boolean
    Dim count As Long
    is_done = False
    count = 1
        
    For i = LBound(answer_cols) To UBound(answer_cols)
        c_col = answer_cols(i).Value
        
        If InStr(1, Cells(c_row, c_col), ("Answer " & count)) = 1 Or _
            InStr(1, Cells(c_row, c_col), ("Answer" & count)) = 1 Then
            
            ReDim Preserve groupd_cols(1 To i)
            groupd_cols(i) = c_col
        
        Else
            count = count + 1

    
            last_index = UBound(grouped_cols) - 1

            'ActiveSheet.Sort.SortFields.Add2 Key:=Range(Cells(c_row - 1, group_cols(1)), Cells(c_row - 1, group_cols(last_index))), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                xlSortNormal

            'ActiveSheet.Sort.SortFields.Add2 Key:=Range(Cells(c_row, group_cols(1)), Cells(c_row, group_cols(last_index))), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
                xlSortNormal


            For c_col = LBound(groupd_cols) To UBound(groupd_cols)
                If StrComp(Cells(c_row - 1, c_col).Value, Cells(c_row, c_col).Value, vbTextCompare) = 0 Then
                
                Else
                    Cells(c_row, last_col + 4) = "Unique"
                    is_done = True
                    Exit For
                End If

            Next c_col

        End If

        If is_done Then
            Exit For
        End If
        
    Next i

    If Not is_done Then
        Cells(c_row, last_col + 4) = "Duplicated"
    End If

End Sub

