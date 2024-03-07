Attribute VB_Name = "Q_InToFile"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : Q_InToFile - ���������� �����
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *

Option Explicit
Option Private Module

    Public Sub InToFile()
5:    Dim strPath As String
6:    On Error GoTo errMsg
7:
8:    strPath = O_XML.OpenAndCloseExcelFileInFolder(bOpenFile:=True, bBackUp:=False)
9:    If strPath = vbNullString Then Exit Sub
10:   Q_InToFile.FilenamesCollectionToPath (strPath)
11:
12:    If MsgBox("Delete the folder of the unpacked Excel file" & vbNewLine & "The Excel file is not deleted!", vbYesNo + vbCritical, "Deleting a folder:") = vbYes Then
13:        Call Q_InToFile.RemoveFolderWithContent(strPath)
14:    End If
15:    Exit Sub
errMsg:
17:    Debug.Print "Error in InToFile!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
18:    Call WriteErrorLog("InToFile")
19: End Sub
    Private Sub FilenamesCollectionToPath(ByVal StrPathToFile As String)
21:    ' ���� �� ������� ����� ��� ����� TXT, � ������� �� ���� ������ �� ���.
22:    ' ��������������� ����� � �������� �������� �� ����� ���.
23:    Dim i      As Long
24:    Dim coll   As Collection
25:    On Error GoTo errMsg
26:    ' ��������� � �������� coll ������ ����� ������
27:    Set coll = FilenamesCollection(StrPathToFile, "*.*", 3)
28:
29:    Application.ScreenUpdating = False    ' ��������� ���������� ������
30:    ' ������ ����� �����
31:    Dim SH     As Worksheet: Set SH = ActiveWorkbook.Sheets.Add(After:=ActiveWorkbook.Sheets(ActiveWorkbook.Sheets.Count))
32:    ' ��������� ��������� �������
33:    With SH.Range("a1").Resize(, 5)
34:        .Value = Array("�", "File name", "Full path", "File Size", "File Extension")
35:        .Font.Bold = True: .Interior.ColorIndex = 17
36:    End With
37:
38:    ' ������� ���������� �� ����
39:    For i = 1 To coll.Count    ' ���������� ��� �������� ���������, ���������� ���� � ������
40:        SH.Range("a" & SH.Rows.Count).End(xlUp).Offset(1).Resize(, 5).Value = _
                      Array(i, C_PublicFunctions.sGetFileName(coll(i)), coll(i), C_PublicFunctions.FileSize(coll(i)), C_PublicFunctions.sGetExtensionName(coll(i)))    ' ������� �� ���� ��������� ������
42:        DoEvents    ' �������� ������� ���������� ��
43:    Next
44:    SH.Range("a:e").EntireColumn.AutoFit    ' ���������� ������ ��������
45:    [a2].Activate: ActiveWindow.FreezePanes = True    ' ���������� ������ ������ �����
46:    Application.ScreenUpdating = True    ' ��������� ���������� ������
47:    Exit Sub
errMsg:
49:    Debug.Print "Error in FilenamesCollectionToPath!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
50:    Call WriteErrorLog("FilenamesCollectionToPath")
51: End Sub

    Private Function FilenamesCollection(ByVal FolderPath As String, Optional ByVal Mask As String = "", _
                Optional ByVal SearchDeep As Long = 999) As Collection
55:    ' � EducatedFool  excelvba.ru/code/FilenamesCollection
56:    ' �������� � �������� ��������� ���� � ����� FolderPath,
57:    ' ����� ����� ������� ������ Mask (����� �������� ������ ����� � ����� ������/�����������)
58:    ' � ������� ������ SearchDeep � ��������� (���� SearchDeep=1, �� �������� �� ���������������).
59:    ' ���������� ���������, ���������� ������ ���� ��������� ������
60:    ' (����������� ����������� ����� ��������� GetAllFileNamesUsingFSO)
61:    Dim FSO    As Object
62:    On Error GoTo errMsg
63:    Set FilenamesCollection = New Collection    ' ������ ������ ���������
64:    Set FSO = CreateObject("Scripting.FileSystemObject")    ' ������ ��������� FileSystemObject
65:    Call GetAllFileNamesUsingFSO(FolderPath, Mask, FSO, FilenamesCollection, SearchDeep)  ' �����
66:    Set FSO = Nothing: Application.StatusBar = False    ' ������� ������ ��������� Excel
67:    Exit Function
errMsg:
69:    Debug.Print "Error in FilenamesCollection!" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
70:    Call WriteErrorLog("FilenamesCollection")
71: End Function

    Private Function GetAllFileNamesUsingFSO(ByVal FolderPath As String, ByVal Mask As String, ByRef FSO, _
                ByRef FileNamesColl As Collection, ByVal SearchDeep As Long)
75:    ' ���������� ��� ����� � �������� � ����� FolderPath, ��������� ������ FSO
76:    ' ������� ����� �������������� � ��� ������, ���� SearchDeep > 1
77:    ' ��������� ���� ��������� ������ � ��������� FileNamesColl
78:    Dim curfold As Object, fil As Object, sfol As Object
79:    On Error Resume Next: Set curfold = FSO.GetFolder(FolderPath)
80:    If Not curfold Is Nothing Then    ' ���� ������� �������� ������ � �����
81:
82:        ' ���������������� ��� ������ ��� ������ ���� � ���������������
83:        ' � ������� ������ ����� � ������ ��������� Excel
84:        ' Application.StatusBar = "����� � �����: " & FolderPath
85:
86:        For Each fil In curfold.Files    ' ���������� ��� ����� � ����� FolderPath
87:            If fil.Name Like "*" & Mask Then FileNamesColl.Add fil.Path
88:        Next
89:        SearchDeep = SearchDeep - 1    ' ��������� ������� ������ � ���������
90:        If SearchDeep Then    ' ���� ���� ������ ������
91:            For Each sfol In curfold.SubFolders    ' ���������� ��� �������� � ����� FolderPath
92:                GetAllFileNamesUsingFSO sfol.Path, Mask, FSO, FileNamesColl, SearchDeep
93:            Next
94:        End If
95:        Set fil = Nothing: Set curfold = Nothing: Set sfol = Nothing   ' ������� ����������
96:    End If
97: End Function

     Private Sub RemoveFolderWithContent(ByVal sFolder As String)
102: '    '���� � ����� ����� ������ ��������, ���� �� ������� �������� � �� ����������
104:   Shell "cmd /c rd /S/Q """ & sFolder & """"
105: End Sub

