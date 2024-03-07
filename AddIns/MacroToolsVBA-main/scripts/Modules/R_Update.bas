Attribute VB_Name = "R_Update"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : R_Update - ������ ���������� ����������
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Explicit
Option Private Module
'�������� ���������� ��� 10 ����
Public Const DayAfterCheck As Byte = 10

    Public Sub StartUpdate()
'14:    If C_Const.NAME_ADDIN & ".xlam" = ThisWorkbook.Name Then
'15:        C_Const.FlagVisible = R_Update.GetUpdate()
'16:        If C_Const.FlagVisible Then Call ShowUpdateMsg
'17:    End If
18: End Sub

    Private Sub ShowUpdateMsg()
21:    Dim TextUpdate  As String
22:    Dim TbRange     As Range
23:
24:    On Error GoTo ErrorHandler
25:
26:
27:    Set TbRange = SHSNIPPETS.ListObjects(C_Const.TB_UPDATE).DataBodyRange
28:    TextUpdate = TbRange.Cells(1, 3).Value2
29:
30:    If TextUpdate <> vbNullString And TbRange.Cells(1, 2).Value2 + R_Update.DayAfterCheck < Now() Then
31:        If MsgBox("Good afternoon!" & vbNewLine & _
                   "An update of the MACROTools VBA add-in is available" & vbNewLine & vbNewLine & _
                   "To update, go to the website VBATools.ru" & vbNewLine & vbNewLine & _
                   TextUpdate & vbNewLine & vbNewLine & _
                   "Remind me later ?" & vbNewLine, vbYesNo, "Updating MACROTools") = vbYes Then
36:            TbRange.Cells(1, 2).Value2 = Now()
37:            Workbooks(C_Const.NAME_ADDIN & ".xlam").Save
38:        End If
39:    End If
40:    If TextUpdate <> TbRange.Cells(1, 1).Value2 And TextUpdate <> vbNullString Then C_Const.FlagVisible = True
41:    Exit Sub
ErrorHandler:
43:    Select Case Err.Number
        Case 1004:
45:            '���� �� �������� ������ ������
46:            '�� ���� �� ������
47:        Case Else:
48:            Debug.Print "Mistake! in ShowUpdateMsg" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
49:            Call WriteErrorLog("ShowUpdateMsg")
50:    End Select
51:    Err.Clear
52: End Sub
     Private Function GetUpdate() As Boolean
54:    Dim NewVersion As String, CurentVersion As String
55:    Dim TbRange     As Range
56:
57:    On Error GoTo ErrorHandler
58:
59:    Set TbRange = SHSNIPPETS.ListObjects(C_Const.TB_UPDATE).DataBodyRange
60:    Application.DisplayAlerts = False
61:
62:    '�������� ���������� ���������� ����
63:    If ChekDateUpdate Then
64:        NewVersion = Split(ResponseTextHttp(C_Const.URL_UPDATE), vbNewLine)(0)
65:    End If
66:    '������ ����������
67:    If NewVersion <> vbNullString Then
68:        CurentVersion = C_Const.NAME_VERSION
69:        '���� ���� ����� ������ �� �������������, ����� ������ ���� ���������� �������
70:        If CurentVersion <> NewVersion Then
71:            GetUpdate = True
72:            TbRange.Cells(1, 3).Value2 = NewVersion
73:            Workbooks(C_Const.NAME_ADDIN & ".xlam").Save
74:        Else
75:            '��������� ����� ���� ��������
76:            GoTo SaveLabel
77:        End If
78:    Else
79:        '��������� ����� ���� ��������
80:        '���� ���� �� ������� �� �������� ���� �������
81:        GoTo SaveLabel
82:    End If
83:    Application.DisplayAlerts = True
84:    Exit Function
SaveLabel:
86:    GetUpdate = False
87:    '��������� ����� ���� ��������
88:    TbRange.Cells(1, 2).Value2 = Now() + DayAfterCheck
89:    Workbooks(C_Const.NAME_ADDIN & ".xlam").Save
90:    Application.DisplayAlerts = True
91:    Exit Function
ErrorHandler:
93:    Select Case Err.Number
        Case 1004, -2146697211:
95:            '���� �� �������� ������ ������
96:            '�� ���� �� ������
97:        Case Else:
98:            Debug.Print "Mistake! in GetUpdate" & vbLf & Err.Number & vbLf & Err.Description & vbCrLf & "in the line" & Erl
99:            Call WriteErrorLog("GetUpdate")
100:    End Select
101:    Err.Clear
102:    GetUpdate = False
103: End Function
'�������� ���� ������ ��� � 10 ����
     Private Function ChekDateUpdate() As Boolean
106:    Dim TbRange     As Range
107:    Dim DateCurentUpdate As Date
108:
109:    ChekDateUpdate = False
110:    Set TbRange = SHSNIPPETS.ListObjects(C_Const.TB_UPDATE).DataBodyRange
111:    DateCurentUpdate = CDate(TbRange.Cells(1, 2).Value2)
112:    '������ �������� ���������� ��� � ������ ����
113:    If Now < DateCurentUpdate + DayAfterCheck Then
114:        Exit Function
115:    End If
116:    ChekDateUpdate = True
117: End Function

Private Function ResponseTextHttp(ByVal URL As String) As String
120:    Dim oHttp       As Object
121:
122:    '������ ����� ������
123:    On Error Resume Next
124:    Set oHttp = CreateObject("MSXML2.XMLHTTP")
125:    If Err.Number <> 0 Then
126:        Set oHttp = CreateObject("MSXML.XMLHTTPRequest")
127:    End If
128:    On Error GoTo 0
129:    If oHttp Is Nothing Then
130:        ResponseTextHttp = vbNullString
131:        Exit Function
132:    End If
133:
134:    With oHttp
135:        .Open "GET", URL, False
136:        .send
137:        If .Status = 200 Then
138:            ResponseTextHttp = .responseText
139:        Else
140:            ResponseTextHttp = vbNullString
141:        End If
142:    End With
143:    Set oHttp = Nothing
End Function

