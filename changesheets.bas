Attribute VB_Name = "ģ��1"
Sub changeSheets()
Dim n As Integer
Dim sht As Worksheet


For Each sht In Worksheets
        sht.Select
    
        For n = 100 To 2 Step -1
         
            '��дרҵ����
            If sht.Range("b" & n) = "��" Then
                 sht.Range("c" & n) = "LG"
            ElseIf sht.Range("b" & n) = "�Ŀ�" Then
                 sht.Range("c" & n) = "WK"
            Else
                sht.Range("c" & n) = "CJ"
            End If
            
            '��д�ƺ�
            If sht.Range("e" & n) = "Ů" Then
                 sht.Range("f" & n) = "Ůʿ"
            Else
                sht.Range("f" & n) = "����"
            End If
            
               'ɾ���հ������ɾ���հ��п��԰�ǰ������Ϊѭ����ִ�еĹ���ȥ��
             If sht.Range("d" & n) = "" Then
                sht.Range("d" & n).EntireRow.Delete
            End If
    
        Next
    
         '��ֲ������±��
        sht.Copy
        ActiveWorkbook.SaveAs FileName:= _
        "/Users/greysun/Desktop/" & sht.Name & ".xlsx"
        ActiveWorkbook.Close
Next



End Sub




