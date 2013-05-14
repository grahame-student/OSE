VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OSE - Oblivion Save Editor
' Copyright (C) 2012, 2013 Grahame White
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

Private Const STRING_SUPERSTRING_NAME   As String = "SuperString"

Private Const STRING_ERROR_BASE         As Long = vbObjectError + &H400&

Private Const STRING_BUFFER_SIZE        As Integer = 15000

Private iValue As String

Private Sub Class_Initialize()

    iValue = 0

End Sub


''' Properties
Public Property Get Value() As String
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Value.VB_UserMemId = 0

    Value = iValue

End Property

Public Property Let Value(NewValue As String)

    iValue = NewValue

End Property


''' Methods
Public Sub AppendString(ExtraString As String)

    iValue = iValue & ExtraString

End Sub

Public Sub AppendStrings(ExtraStrings() As String, Optional Separator As String = "")

    ' In a test of 32000 concatenations this proved to be ~400 times faster than simply using the & operator

    Dim StringNumber As Integer
    Dim ExtraStringLength As Long
    Dim SeparatorLength As Long
    Dim Offset As Long
    Dim Buffer As String

    Buffer = iValue
    Offset = Len(Buffer)

    SeparatorLength = Len(Separator)
    
    For StringNumber = 0 To UBound(ExtraStrings)
        ExtraStringLength = Len(ExtraStrings(StringNumber))
        If Offset + ExtraStringLength >= Len(Buffer) Then
            Buffer = Buffer & String$(STRING_BUFFER_SIZE, 0)
        End If
        Mid(Buffer, Offset + 1, ExtraStringLength) = ExtraStrings(StringNumber)
        Offset = Offset + ExtraStringLength
        
        If SeparatorLength > 0 Then
            If StringNumber < UBound(ExtraStrings) Then
                If Offset + SeparatorLength >= Len(Buffer) Then
                    Buffer = Buffer & String$(STRING_BUFFER_SIZE, 0)
                End If
                Mid(Buffer, Offset + 1, SeparatorLength) = Separator
                Offset = Offset + SeparatorLength
            End If
        End If
    Next StringNumber

    iValue = Left(Buffer, Offset)

End Sub
