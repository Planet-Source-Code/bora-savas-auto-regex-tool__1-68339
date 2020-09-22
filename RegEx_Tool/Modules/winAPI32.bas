Attribute VB_Name = "winAPI32"
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
' Module Name  : AutoRegExTool.winAPI32
'
' Description  : Windows related API
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

' To get the word under the mouse pointer
Private Const EM_CHARFROMPOS& = &HD7
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

' (!) This function is currently removed from the project
Public Function getWordUnderMousePointer(rch As RichTextBox, x As Single, y As Single) As String
    
    Dim pt As POINTAPI
    Dim pos As Long
    Dim start_pos As Long
    Dim end_pos As Long
    Dim ch As String
    Dim txt As String
    Dim txtlen As Long

    ' Convert the position to pixels.
    pt.x = x \ Screen.TwipsPerPixelX
    pt.y = y \ Screen.TwipsPerPixelY

    ' Get the character number
    pos = SendMessage(rch.hWnd, EM_CHARFROMPOS, 0&, pt)
    If pos <= 0 Then Exit Function

    ' Find the start of the word.
    txt = rch.Text
    For start_pos = pos To 1 Step -1
        ch = Mid$(rch.Text, start_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ( _
            (ch >= "0" And ch <= "9") Or _
            (ch >= "a" And ch <= "z") Or _
            (ch >= "A" And ch <= "Z") Or _
            ch = "_" _
        ) Then Exit For
    Next start_pos
    start_pos = start_pos + 1

    ' Find the end of the word.
    txtlen = Len(txt)
    For end_pos = pos To txtlen
        ch = Mid$(txt, end_pos, 1)
        ' Allow digits, letters, and underscores.
        If Not ( _
            (ch >= "0" And ch <= "9") Or _
            (ch >= "a" And ch <= "z") Or _
            (ch >= "A" And ch <= "Z") Or _
            ch = "_" _
        ) Then Exit For
    Next end_pos
    end_pos = end_pos - 1

    If start_pos <= end_pos Then getWordUnderMousePointer = Mid$(txt, start_pos, end_pos - start_pos + 1)
    
End Function

