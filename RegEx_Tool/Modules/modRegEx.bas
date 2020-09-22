Attribute VB_Name = "modRegEx"
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
' Module Name  : AutoRegExTool.modRegEx
'
' Description  : RegEx related functions
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

Public Type SpecialPattern
    pattern             As String       'Regular expression pattern
    conditionOperator   As String * 1   'Operator (<,>)
    conditionalValue    As Long         'Value for the condition
    isSpecialRegEx      As Boolean
End Type

Public Type SpecialMatch
    matchIndex  As Long
    matchLength As Long
    matchValue  As String
End Type
    
Public myRegExp As RegExp
Public digits As RegExp
Public myMatches As MatchCollection
Public digitMatch As MatchCollection
Public myMatch As Match

Public Function applyRegEx()

    Dim cnt         As Integer
    Dim expression  As String
    Dim definition  As String
    Dim action      As String
    Dim parameter   As String
    
    DoEvents
    Load frmProgress
    frmProgress.Show
    
    For cnt = 1 To totalRegEx
        
        With frmRegEx.lstRegEx
            expression = .ListItems(cnt).Text
            definition = .ListItems(cnt).SubItems(1)
            action = .ListItems(cnt).SubItems(2)
            parameter = .ListItems(cnt).SubItems(3)
            
            DoEvents
            frmProgress.Status.Caption = "Status: Processing " & cnt & "/" & totalRegEx
            
            'Get Action
            Select Case action
                Case "Replace"
                    'Replace function
                    frmProgress.List1.ListIndex = cnt - 1
                    frmMain.txtEdit.Text = regExReplace(frmMain.txtEdit.Text, expression, parameter)
                    saveFile App.Path & "\output\" & cnt & "." & currentFile & "[" & action & "].txt", frmMain.txtEdit.Text
                    
                Case "Remove"
                    'Remove function
                    frmProgress.List1.ListIndex = cnt - 1
                    frmMain.txtEdit.Text = regExReplace(frmMain.txtEdit.Text, expression, vbNullString)
                    saveFile App.Path & "\output\" & cnt & "." & currentFile & "[" & action & "].txt", frmMain.txtEdit.Text

                Case "Highlight"
                    'Higlight function
                    highlightMatches frmMain.txtEdit, expression
                
                Case "WriteToFile"
                    'WriteToFile function
                    writeAllMatchesToFile frmMain.txtEdit, expression, parameter
                    
                Case Else
                    'Nothing
            End Select
        End With
        
    Next cnt
    
    processTime = GetTickCount - processTime
    frmProgress.Status.Caption = frmProgress.Status.Caption & vbCrLf & _
                                    "Processed in " & processTime & " miliseconds."
    
End Function

Public Function regExReplace(iString As String, pattern As String, repString As String) As String
    
    Set myRegExp = New RegExp
    myRegExp.Global = True
    myRegExp.IgnoreCase = Not isCaseSensitive
    
    myRegExp.pattern = pattern '"([^(\r\n)])\r\n"
    repString = Replace(repString, "$n", vbCrLf)
    regExReplace = myRegExp.Replace(iString, repString)
    
    Set myRegExp = Nothing
    
    frmProgress.Status.Caption = frmProgress.Status.Caption & vbCrLf & vbCrLf & "Log files saved in the 'output/' folder"

End Function

Public Function highlightMatches(rTextBox As RichTextBox, pattern As String) As Long
    
    Dim cnt                 As Long
    Dim parsedPattern       As SpecialPattern
    Dim isSpecialPattern    As Boolean
    Dim localProcessTime    As Long
    Dim spMatch             As SpecialMatch
    
    On Error GoTo HANDLER
    
    cnt = 0
    localProcessTime = GetTickCount
    
    parsedPattern = parseRegularExpressions(pattern)
    
    Set myRegExp = New RegExp
    Set digits = New RegExp
    
    myRegExp.pattern = parsedPattern.pattern
    myRegExp.Global = True
    myRegExp.IgnoreCase = Not isCaseSensitive
    digits.pattern = "\d+"
    digits.Global = True
    
    Set myMatches = myRegExp.Execute(rTextBox.Text)
    
    For Each myMatch In myMatches
        If (parsedPattern.isSpecialRegEx) Then
            Set digitMatch = digits.Execute(myMatch.value)
            If (parsedPattern.conditionOperator = ">") Then
                If (digitMatch.Item(0) > parsedPattern.conditionalValue) Then
                    spMatch.matchIndex = myMatch.FirstIndex
                    spMatch.matchLength = myMatch.Length
                    spMatch.matchValue = myMatch.value
                End If
            Else
                If (digitMatch.Item(0) < parsedPattern.conditionalValue) Then
                    spMatch.matchIndex = myMatch.FirstIndex
                    spMatch.matchLength = myMatch.Length
                    spMatch.matchValue = myMatch.value
                End If
            End If
        Else
            spMatch.matchIndex = myMatch.FirstIndex
            spMatch.matchLength = myMatch.Length
            spMatch.matchValue = myMatch.value
        End If
        
        rTextBox.SelStart = spMatch.matchIndex
        rTextBox.SelLength = spMatch.matchLength
        rTextBox.SelColor = vbRed
        DoEvents
        cnt = cnt + 1
        frmProgress.Status = "Status: " & cnt & "/" & myMatches.Count & " item(s) in progress"
        frmMain.lblStatus.Caption = cnt & "/" & myMatches.Count & " item(s) in progress"
        If closePressed Then Exit For
        'backreference n text: myMatch.SubMatches(n-1)
    Next
        
    frmProgress.Status.Caption = frmProgress.Status.Caption & vbCrLf & myMatches.Count & " items has beed highlighted"
    
    Set myRegExp = Nothing
    highlightMatches = GetTickCount - localProcessTime
    
    Exit Function
    
HANDLER:
    MsgBox "Following error occured while applying the regular expression." & vbCrLf & _
            "[" & Err.Number & "] " & Err.Description, vbCritical

End Function

Public Function parseRegularExpressions(pattern As String) As SpecialPattern
    
    Dim conditionalVal  As Long
    Dim numOperatorPos  As Integer
    Dim closingBracket  As Integer
    Dim operatorPos     As Integer
    Dim operatorField   As String
    Dim operator        As String * 1
    Dim parsedPattern   As SpecialPattern
    
    On Error GoTo HANDLER
    
    If (Trim$(Len(pattern)) = 0) Then
        parsedPattern.pattern = pattern
        parseRegularExpressions = parsedPattern
    End If
    
    'Look for "\d{} pattern
    numOperatorPos = InStr(1, pattern, "\d{")
    If (numOperatorPos > 0) Then closingBracket = InStr(numOperatorPos, pattern, "}")
    'If there is no closing bracket
    If (closingBracket = 0) Then
        parsedPattern.pattern = pattern
        parseRegularExpressions = parsedPattern
        Exit Function
    End If
    'Get string range for a possible operator
    operatorField = Mid$(pattern, numOperatorPos + 3, closingBracket - numOperatorPos - 3)
    'Look for operator
    operatorPos = InStr(1, operatorField, "<")
    If (operatorPos = 0) Then operatorPos = InStr(1, operatorField, ">")
    'No operator found
    If (operatorPos = 0) Then
        parsedPattern.pattern = pattern
        parseRegularExpressions = parsedPattern
        Exit Function
    End If
    
    'Get operator
    operator = Mid$(operatorField, operatorPos, 1)
    'Get conditional value
    conditionalVal = Mid$(operatorField, operatorPos + 1, Len(operatorField) - 1)
    
    pattern = Replace(pattern, operator & conditionalVal, vbNullString)
    
    parsedPattern.pattern = pattern
    parsedPattern.conditionalValue = conditionalVal
    parsedPattern.conditionOperator = operator
    parsedPattern.isSpecialRegEx = True
    
    parseRegularExpressions = parsedPattern
    
    Exit Function
    
HANDLER:
    MsgBox "[" & Err.Description & "]" & vbCrLf & _
            "Conditional value must be a number, not a character", vbCritical, "RegEx Format Error (" & Err.Number & ")"
    
End Function

Public Sub writeAllMatchesToFile(iString As String, pattern As String, filename As String)
    
    Dim cnt As Long
    
    cnt = 0
    Set myRegExp = New RegExp
    
    myRegExp.pattern = pattern
    myRegExp.Global = True
    myRegExp.IgnoreCase = Not isCaseSensitive
    
    Set myMatches = myRegExp.Execute(iString)
    clearBuffer (1)
    
    For Each myMatch In myMatches
        appendToBuffer myMatch.value & vbCrLf, 1
        DoEvents
        cnt = cnt + 1
        frmProgress.Status = "Status: " & cnt & "/" & myMatches.Count & " item(s) in progress"
        If closePressed Then Exit For
    Next
    
    saveFile App.Path & "\output\" & filename & "[WriteToFile].txt", getBuffer(1)
    
    clearBuffer (1)
    
    frmProgress.Status.Caption = frmProgress.Status.Caption & vbCrLf & "File saved to: " & App.Path & "\output\" & filename & _
                                    "[WriteToFile].txt"
        
    Set myRegExp = Nothing

End Sub

Public Sub fillList(listObject As ListBox)

    Dim cnt As Integer
    
    For cnt = 1 To totalRegEx
        With frmRegEx.lstRegEx
            listObject.AddItem cnt & ". " & .ListItems(cnt).Text
        End With
    Next cnt
    
End Sub
