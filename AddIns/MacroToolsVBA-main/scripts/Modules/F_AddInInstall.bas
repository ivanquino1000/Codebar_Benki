Attribute VB_Name = "F_AddInInstall"
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
'* Module     : F_AddInInstall - ������ ��������� ����������
'* Created    : 15-09-2019 15:48
'* Author     : VBATools
'* Contacts   : http://vbatools.ru/ https://vk.com/vbatools
'* Copyright  : VBATools.ru
'* * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * * *
Option Private Module
Option Explicit
' ��������� ����������
    Public Sub InstallationAddMacro()
10:    Dim AddFolder As String
11:    On Error GoTo InstallationAdd_Err
12:    ' ��������� ������� �� ������ ����������
13:    AddFolder = Replace(Application.UserLibraryPath & "\", "\\", "\")
14:    '�������� �� ������� ����������
15:    If Dir(AddFolder, vbDirectory) = vbNullString Then
16:        Call MsgBox("Unfortunately, the program cannot install the add-on on this computer." _
                      & vbCrLf & "There is no directory with add-ons." & vbCrLf & _
                      "Contact the developer of the program.", vbCritical, _
                      "Add-on installation failed")
20:        Exit Sub
21:    End If
22:    '��������� ����� ������������ ����������
23:    If FileHave(AddFolder & C_Const.NAME_ADDIN & ".xlam") Then AddIns(C_Const.NAME_ADDIN).Installed = False
24:    ' ��������� ������� �� ����������
25:    If WorkbookIsOpen(C_Const.NAME_ADDIN & ".xlam") Then
26:        Call MsgBox("The file with the add-in is already open." & vbCrLf & _
                      "It may have already been installed earlier.", vbCritical, _
                      "Program installation failure")
29:        Exit Sub
30:    End If
31:    ' ��������� ���
32:    Application.EnableEvents = 0
33:    Application.DisplayAlerts = False
34:    If Workbooks.Count = 0 Then Workbooks.Add
35:    ThisWorkbook.SaveAs AddFolder & C_Const.NAME_ADDIN & ".xlam", FileFormat:=xlOpenXMLAddIn
36:    AddIns.Add Filename:=AddFolder & C_Const.NAME_ADDIN & ".xlam"
37:    AddIns(C_Const.NAME_ADDIN).Installed = True
38:    Application.EnableEvents = 1
39:    Application.DisplayAlerts = True
40:    Call MsgBox("The program has been successfully installed!" & vbCrLf & _
                  "Just open or create a new document.", vbInformation, _
                  "Installing the add-in:" & C_Const.NAME_ADDIN)
43:    ThisWorkbook.Close False
44:    Exit Sub
InstallationAdd_Err:
46:    If Err.Number = 1004 Then
47:        MsgBox "To install the add-on, please close this file and run it again.", _
                      64, "Installation"
49:    Else
50:        MsgBox Err.Description & vbCrLf & "� F_AddInInstall.InstallationAdd " & vbCrLf & "in the line" & Erl, vbExclamation + vbOKOnly, "Mistake:"
51:        Call WriteErrorLog("F_AddInInstall.InstallationAdd")
52:    End If
53: End Sub

