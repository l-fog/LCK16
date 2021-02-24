Attribute VB_Name = "Mdl_LCK16"
'Copyright (C) 2021 Locked_Fog Studio. All Rights Reserved.

Public code
Public decode
Public decode16(15)

Sub main()              '������ȴ�main���̿�ʼ��ʹ��[Form.show]�ķ������ô��塣��Ϊ���������ʼ��
    code = Array(")", "!", "@", "#", "$", "%", "^", "&", _
                        "*", "(", ":", ";", ",", ".", "<", ">")
    decode = Array("0", "1", "2", "3", "4", "5", "6", _
                            "7", "8", "9", "A", "B", "C", "D", _
                            "E", "F")
                            
    For i = 0 To 15
        decode16(i) = i
    Next i
    '****************���²���ȫ����ɾ��********************
    Dim str As String
    str = Coding("U:\Codes\vb6\code&decode\code&decode.zip")
    Open "e:\cod.txt" For Output As #1
    Print #1, str & Chr(0)
    Close #1
    str = DeCoding("e:\cod.txt", "e:\decod.zip")
    '********************************************************
    
End Sub

Public Function DeCoding(ByVal Path As String, ByVal OutPut As String) As String    '���ܣ���һ�������������ļ�·�����ڶ�������������ļ�·��������ֵ������ļ�·��
    Open Path For Input As #1
    Open OutPut For Binary As #2
    Dim point As Long
    ind = FileLen(Path) / 2 - 1
    Dim str() As String
    ReDim str(0 To ind, 1)
    Dim back As Byte
    Dim az As Byte
    Do While Not EOF(1)
        str(point, 0) = Input(1, 1)
        str(point, 1) = Input(1, 1)
        point = point + 1
        For i = 0 To 15
            If str(point - 1, 0) = code(i) Then
                back = decode16(i) * 16
                For j = 0 To 15
                    If str(point - 1, 1) = code(j) Then
                        back = back + decode16(j)
                        Put #2, , back
                        Exit For
                    End If
                Next j
                Exit For
            End If
        Next i
    Loop
    
    DeCoding = OutPut
    Close #1
    Close #2
End Function

Public Function Coding(ByVal Path As String) As String      '���ܣ������Ǳ������ļ�·��������ֵ�������ı�
    Open Path For Binary As #1
    Dim point As Long
    point = 1
    Dim Bin() As Byte
    Dim str(1) As String
    Dim back As String
    
    ReDim Bin(0 To FileLen(Path) - 1)
    Get #1, point, Bin
    For j = 0 To FileLen(Path) - 1
        If Len(Hex(Bin(j))) = 1 Then str(0) = "0" Else str(0) = Left(Hex(Bin(j)), 1)
        str(1) = Right(Hex(Bin(j)), 1)
        For i = 0 To 15
            If str(0) = decode(i) Then
                back = back & code(i)
                Exit For
            End If
        Next i
        For i = 0 To 15
            If str(1) = decode(i) Then
                back = back & code(i)
                Exit For
            End If
        Next i
        point = point + 1
    Next j
    Coding = back
    Close #1
End Function
