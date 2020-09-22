Attribute VB_Name = "modINI"
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
' Module Name  : AutoRegExTool.modINI
'
' Description  : INI Related API & implemention
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

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpdefault As String, ByVal lpreturnedstring As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Public Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationname As String, ByVal lpKeyname As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Function ReadINI(Section As String, KeyName As String, filename As String) As String
    Dim sRet As String
    sRet = String(255, Chr(0))
    ReadINI = Left(sRet, GetPrivateProfileString(Section, ByVal KeyName$, "", sRet, Len(sRet), filename))
End Function

Function WriteINI(sSection As String, sKeyName As String, sNewString As String, sFileName) As Integer
    Dim r
    r = WritePrivateProfileString(sSection, sKeyName, sNewString, sFileName)
End Function

Function DeleteSection(ByVal INIFileLoc As String, ByVal Section As String)
    'This Function Deletes a specified Secti
    '     on from an INI file
    'INIFileLoc = The location of the INI Fi
    '     le (ex. "C:\Windows\INIFile.ini")
    'Section = The name of the Section you w
    '     ish to remove (ex. "Section Number 1")
    'Checking to see if the INI File specifi
    '     ed exists
    If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc, vbExclamation, "INI File Not Found": Exit Function
    'If INI File exists then proceed to dele
    '     te Section
    WritePrivateProfileString Section, vbNullString, vbNullString, INIFileLoc
    'NOTE: vbNullString is the coding in whi
    '     ch to delete a Section, or Key
End Function


Function DeleteKey(ByVal INIFileLoc As String, ByVal Section As String, ByVal Key As String)
    
    If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc, vbExclamation, "INI File Not Found": Exit Function
    WritePrivateProfileString Section, Key, vbNullString, INIFileLoc
    
End Function


Function DeleteKeyValue(ByVal INIFileLoc As String, ByVal Section As String, ByVal Key As String)
    
    If Dir(INIFileLoc) = "" Then MsgBox "File Not Found: " & INIFileLoc & vbCrLf & "Please refer To code in Function 'DeleteKeyValue'", vbExclamation, "INI File Not Found": Exit Function
    WritePrivateProfileString Section, Key, "", INIFileLoc

End Function



