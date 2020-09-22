Attribute VB_Name = "modGeneral"
'-------------------------------------------------------------------------
'
' Copyright Notice :
' All rights reserved by Bora SAVAS 2005, 2007
' Osaka University, Japan
'
' Contact : borasavas@gmail.com
'           http://www.japanalyzer.com
'
'-------------------------------------------------------------------------
'
' Module Name  : AutoRegExTool.modGeneral
'
' Description  : General functions
'
'-------------------------------------------------------------------------
'
'   Redistribution and use in source and binary forms, with or without
'   modification, are permitted provided that the following conditions are
'   met:
'
'   -Redistributions of source code must retain the above copyright notice,
'    this list of conditions and the following disclaimer.
'
'   -Redistributions in binary form must reproduce the above copyright
'    notice, this list of conditions and the following disclaimer in the
'    documentation and/or other materials provided with the distribution.
'
'   THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS
'   "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT
'   LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR
'   A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL THE
'   CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL,
'   EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO,
'   PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR
'   PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF
'   LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING
'   NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
'   SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'-------------------------------------------------------------------------

Option Explicit

'Get tick count
Public Declare Function GetTickCount Lib "kernel32" () As Long

Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, _
                        ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, _
                        ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'For Dragging Borderless Forms...
Public Declare Function ReleaseCapture Lib "user32" () As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, _
                        ByVal wParam As Long, ByVal lParam As Long) As Long

'AlwaysOnTop()
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
                        ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, _
                        ByVal wFlags As Long) As Long

'For AlwaysOnTop()
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2

Public Const WM_NCLBUTTONDOWN = &HA1
Public Const HTCAPTION = 2

Public cDialog As New CommonDialog

'---- Tooltip related: START ----
Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type LVHITTESTINFO
   pt       As POINTAPI
   flags    As Long
   iItem    As Long
   iSubItem As Long
End Type

Public TT As CTooltip
Public Const LVM_FIRST = &H1000&
Public Const LVM_HITTEST = LVM_FIRST + 18
'---- Tooltip related: END ----

Public justLoaded       As Boolean
Public editMode         As Boolean
Public closePressed     As Boolean
Public splashShowed     As Boolean
Public isCaseSensitive  As Boolean

Public processTime              As Long      'For GetTickCount
Public totalRegEx               As Integer
Public currentSelectedItem      As Integer   'Current selected item in the ListView
Public currentFile              As String    'Current file name opened
Public currentRegExFile         As String
Public DEFAULT_SETTINGS_FILE    As String

Public Sub changeFontSettings(fName As String, fSize As Integer)

    'Main form
    frmMain.txtEdit.Font.Name = fName
    frmMain.txtEdit.Font.Size = fSize
    
    'frmEntry
    With frmEntry
        .txtRegEx.Font.Name = fName
        .txtRegEx.Font.Size = fSize
        .txtRegEx.Font.Charset = 1
        .txtDefinition.Font.Name = fName
        .txtDefinition.Font.Size = fSize
        .txtDefinition.Font.Charset = 1
    End With
    
    'frmRegEx
    With frmRegEx
        .lstRegEx.Font.Name = fName
        .lstRegEx.Font.Size = fSize
        .lstRegEx.Font.Charset = 1
    End With
    
    WriteINI "RegExSettings", "CurrentFont", fName, DEFAULT_SETTINGS_FILE
    WriteINI "RegExSettings", "CurrentFontSize", CStr(fSize), DEFAULT_SETTINGS_FILE
    
End Sub

Public Function saveListView(lW As ListView, fName As String)

    Dim FileId As Integer
    Dim x, i As Integer
    Dim sIdx As Integer
    
    sIdx = lW.ColumnHeaders.Count - 1
    FileId = FreeFile
    
    On Error Resume Next
    
    Open fName For Output As #FileId
    For i = 1 To lW.ListItems.Count
        Write #FileId, lW.ListItems.Item(i).Text
        For x = 1 To sIdx
            Write #FileId, lW.ListItems.Item(i).SubItems(x)
        Next
    Next
    
    Close #FileId

End Function
Public Function loadListView(lW As ListView, fName As String)
    
    Dim FileId As Integer
    Dim LVI As ListItem
    Dim fData As New Collection
    Dim Buffer As String
    Dim x, y As Integer
    Dim i As Integer
    Dim sIdx As Integer
    
    sIdx = lW.ColumnHeaders.Count - 1
    i = 0
    FileId = FreeFile
    
    On Error Resume Next
    
    Open fName For Input As #FileId
    While Not EOF(FileId)
        i = i + 1
        For x = 0 To sIdx
            Input #FileId, Buffer
            fData.Add Buffer
        Next
        Set LVI = lW.ListItems.Add
        LVI.Text = fData.Item(fData.Count - sIdx)
        For y = 1 To sIdx
            LVI.SubItems(y) = fData.Item(fData.Count + y - sIdx)
        Next
    Wend
    
    Close #FileId
    
End Function

Public Function initRegExPatterns(patternFile As String) As Boolean

    
End Function

Public Function saveFile(sPath As String, sData As String) As Boolean
    
    On Error GoTo HANDLER
    
    Dim iFF As Integer
    
    'If file exists, delete it first
    If Dir(sPath) <> vbNullString Then Kill sPath
    
    iFF = FreeFile
    Open sPath For Output As iFF
        Print #iFF, "-- Output of AutoRegEx. (" & Now & ")"
        Print #iFF, "-- All Rights Reserved 2005-2007 { Bora Savas - borasavas@gmail.com }" & vbCrLf & vbCrLf
        Print #iFF, sData
    Close #iFF
    
    saveFile = True
    Exit Function
    
HANDLER:
    saveFile = False
        
End Function

Public Function getWordFreq(strSource, strSearchString) As Long
    
    'Calculates frequency of the character(s)
    On Error Resume Next
    
    If Len(strSource) Then
        getWordFreq = UBound(Split(strSource, strSearchString))
    End If

End Function

Public Sub DragForm(Frm As Form)
    
    On Local Error Resume Next
    'Move the borderless form...
    Call ReleaseCapture
    Call SendMessage(Frm.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0)

End Sub

Public Sub AlwaysOnTop(TheForm As Form, SetOnTop As Boolean)
    
    Dim lflag As Long
    Dim SWP_NOACTIVATE As Long
    Dim SWP_SHOWWINDOW As Long
    
    If SetOnTop Then
        lflag = HWND_TOPMOST
    Else
        lflag = HWND_NOTOPMOST
    End If
    SetWindowPos TheForm.hWnd, lflag, TheForm.Left / Screen.TwipsPerPixelX, TheForm.Top / Screen.TwipsPerPixelY, TheForm.Width / Screen.TwipsPerPixelX, TheForm.Height / Screen.TwipsPerPixelY, SWP_NOACTIVATE Or SWP_SHOWWINDOW

End Sub

