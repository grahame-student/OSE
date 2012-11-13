Attribute VB_Name = "HelperFunctions"
' OSE - Oblivion Save Editor
' Copyright (C) 2012  Grahame White
'
' This program is free software; you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation; either version 2 of the License, or
' (at your option) any later version.
'
' This program is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License along
' with this program; if not, write to the Free Software Foundation, Inc.,
' 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA.

Option Explicit
DefObj A-Z

Public Function Ordinal(ByVal DayValue As Integer) As String

    Select Case DayValue
         Case 1, 21, 31
            Ordinal = DayValue & "st"
         Case 2, 22
            Ordinal = DayValue & "nd"
         Case 3, 23
            Ordinal = DayValue & "rd"
         Case Else
            Ordinal = DayValue & "th"
      End Select

End Function

Public Sub ValidateInput(ByRef CheckTextBox As TextBox, ByVal MinVal As Long, ByVal MaxVal As Long)
    
    If Not (IsNumeric(CheckTextBox.Text)) Then
        SendMessage CheckTextBox.hwnd, EM_UNDO, 0, ByVal 0&
        SendMessage CheckTextBox.hwnd, EM_EMPTYUNDOBUFFER, 0, ByVal 0&
    ElseIf CheckTextBox.Text < MinVal Or CheckTextBox.Text > MaxVal Then
        SendMessage CheckTextBox.hwnd, EM_UNDO, 0, ByVal 0&
        SendMessage CheckTextBox.hwnd, EM_EMPTYUNDOBUFFER, 0, ByVal 0&
    Else
        SendMessage CheckTextBox.hwnd, EM_EMPTYUNDOBUFFER, 0, ByVal 0&
        CheckTextBox.Text = Int(CheckTextBox.Text)
    End If

End Sub

Public Function GetFormID(ByVal Iref As Long) As Long

    If Iref < SaveFileData.FormIDs.NumberOfFormIDs Then
        GetFormID = SaveFileData.FormIDs.FormIDsList(Iref)
    End If

End Function

Public Function GetModIndex(ByVal ModName As String) As LongType

    Dim i As Integer
    Dim Index As ByteArray
    
    For i = 0 To SaveFileData.PlugIns.NumberOfPlugins - 1
        If ModName = SaveFileData.PlugIns.PlugInNames(i) Then
            Index.Bytes(3) = i
            Exit For
        End If
    Next i

    LSet GetModIndex = Index

End Function

