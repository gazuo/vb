Attribute VB_Name = "ģ��1"
Sub RemoveDuplicatesAndMerge()
    Dim lastRow As Long
    Dim i As Long
    
    lastRow = Cells(Rows.Count, "A").End(xlUp).Row
    
    Range("A1:B" & lastRow).Sort key1:=Range("A1"), order1:=xlAscending, Header:=xlYes
        
    For i = lastRow To 2 Step -1
        If Cells(i, 1).Value = Cells(i - 1, 1).Value Then
         If Cells(i, 2).Value = Cells(i - 1, 2).Value Then
            Cells(i - 1, 3).Value = Cells(i - 1, 3).Value & " " & Cells(i, 3).Value
            Rows(i).EntireRow.Delete
        End If
        End If
    Next i
    
    cc = Cells(Rows.Count, "C").End(xlUp).Row
    
    For j = 2 To cc
    
    '�жϵ�Ԫ��ո��ж��ٸ�

'    Range("d" & j) = "=LEN(C" & j & ")-LEN(SUBSTITUTE(C" & j & "," & """" & " " & """" & "," & """" & "" & """" & "))+1"
    sr = Range("c" & j)
    Range("e" & j) = Len(sr) - Len(Replace(sr, " ", "")) + 1
    
     Next j
    
    
End Sub


'�˴���ᰴ��A�е���������Ȼ����������У�����ڵ�ǰ�к�ǰһ���о�����ͬ����ֵ��
'����Ӧ��B��ֵ�ϲ���һ�𣬲�ɾ����ǰ�С���ע�⣬�ںϲ�B������ʱ��
'�˴���ʹ�ö�����ո�ָ������������Ҫ��ͬ�ķָ��������޸ĸ��������������Ҫ��
