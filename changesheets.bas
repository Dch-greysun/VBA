Attribute VB_Name = "模块1"
Sub changeSheets()
Dim n As Integer
Dim sht As Worksheet


For Each sht In Worksheets
        sht.Select
    
        For n = 100 To 2 Step -1
         
            '填写专业代号
            If sht.Range("b" & n) = "理工" Then
                 sht.Range("c" & n) = "LG"
            ElseIf sht.Range("b" & n) = "文科" Then
                 sht.Range("c" & n) = "WK"
            Else
                sht.Range("c" & n) = "CJ"
            End If
            
            '填写称呼
            If sht.Range("e" & n) = "女" Then
                 sht.Range("f" & n) = "女士"
            Else
                sht.Range("f" & n) = "先生"
            End If
            
               '删除空白行最后删除空白行可以把前两轮因为循环多执行的工作去掉
             If sht.Range("d" & n) = "" Then
                sht.Range("d" & n).EntireRow.Delete
            End If
    
        Next
    
         '拆分并生成新表格
        sht.Copy
        ActiveWorkbook.SaveAs FileName:= _
        "/Users/greysun/Desktop/" & sht.Name & ".xlsx"
        ActiveWorkbook.Close
Next



End Sub




