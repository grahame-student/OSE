VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperInt"
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

' This class is inteneded to be a drop in alternative for the Integer datatype
' The extra features offered are:
'   A cleaner method for incrementing the variables value ('Add' method)
'   A cleaner method for decreasing the variables value ('Minus' method)
'   Overflow errors will list the direction of the overflow (i.e. which limit has been exceeded, upper or lower)
'   Converts to/from a raw byte array

Private Const INT_SUPERINT_NAME As String = "SuperInt"

Private Const INT_MIN As Integer = -32768
Private Const INT_MAX As Integer = 32767

Private Const INT_MINUS As Integer = -1
Private Const INT_ADD As Integer = 1

Private Const INT_ERROR_BASE                        As Long = vbObjectError + &H100
Private Const INT_ERROR_OVERFLOW                    As Long = INT_ERROR_BASE + 1
Private Const INT_ERROR_UNDERFLOW                   As Long = INT_ERROR_BASE + 2
Private Const INT_ERROR_INVALID_ARRAY_POSITION      As Long = INT_ERROR_BASE + 3

Private Const INT_ERROR_DESC_OVERFLOW               As String = "This operation would cause an 'Integer' overflow"
Private Const INT_ERROR_DESC_UNDERFLOW              As String = "This operation would cause an 'Integer' underflow"
Private Const INT_ERROR_DESC_INVALID_ARRAY_POSITION As String = "Trying to access invalid array position"

Private Const INT_BYTE_SIZE As Integer = 2

Private Type TwoBytes
    RawValue As Integer
End Type

Private Type ByteArray
    RawByteValue(INT_BYTE_SIZE - 1) As Byte
End Type

Private iValue As TwoBytes
Private iByte As ByteArray


''' Initialise
Private Sub Class_Initialize()

    iValue.RawValue = 0
    UpdateByteArray

End Sub


''' Properties
' Value property
Public Property Get Value() As Integer
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Value.VB_UserMemId = 0

    Value = iValue.RawValue

End Property

Public Property Let Value(NewValue As Integer)

    iValue.RawValue = NewValue
    UpdateByteArray

End Property

' ByteValue property
Public Property Get ByteValue(Index As Integer) As Byte

    If Index < 0 Then
        Err.Raise INT_ERROR_INVALID_ARRAY_POSITION, INT_SUPERINT_NAME, INT_ERROR_DESC_INVALID_ARRAY_POSITION
    ElseIf Index >= INT_BYTE_SIZE Then
        Err.Raise INT_ERROR_INVALID_ARRAY_POSITION, INT_SUPERINT_NAME, INT_ERROR_DESC_INVALID_ARRAY_POSITION
    End If

    ByteValue = iByte.RawByteValue(Index)

End Property

Public Property Let ByteValue(Index As Integer, NewByteValue As Byte)

    If Index < 0 Then
        Err.Raise INT_ERROR_INVALID_ARRAY_POSITION, INT_SUPERINT_NAME, INT_ERROR_DESC_INVALID_ARRAY_POSITION
    ElseIf Index >= INT_BYTE_SIZE Then
        Err.Raise INT_ERROR_INVALID_ARRAY_POSITION, INT_SUPERINT_NAME, INT_ERROR_DESC_INVALID_ARRAY_POSITION
    End If

    iByte.RawByteValue(Index) = NewByteValue
    UpdateValue

End Property


''' Methods
Public Sub Add(Optional Increment As Integer = 1)

    ' Add Increment to Value
    ' If no Increment is listed then add 1 instead (Similar to ++ and += in C-like languages)

    If Overflow(Increment, INT_ADD) Then
        Err.Raise INT_ERROR_OVERFLOW, INT_SUPERINT_NAME, INT_ERROR_DESC_OVERFLOW
        Exit Sub
    ElseIf Underflow(Increment, INT_ADD) Then
        Err.Raise INT_ERROR_UNDERFLOW, INT_SUPERINT_NAME, INT_ERROR_DESC_UNDERFLOW
        Exit Sub
    End If

    iValue.RawValue = iValue.RawValue + Increment
    UpdateByteArray

End Sub

Public Sub Minus(Optional Decrement As Integer = 1)

    ' Subtract Decrement from Value
    ' If no Decrement is listed then subtract 1 instead (Similar to -- and -= in C-like languages)

    If Overflow(Decrement, INT_MINUS) Then
        Err.Raise INT_ERROR_OVERFLOW, INT_SUPERINT_NAME, INT_ERROR_DESC_OVERFLOW
        Exit Sub
    ElseIf Underflow(Decrement, INT_MINUS) Then
        Err.Raise INT_ERROR_UNDERFLOW, INT_SUPERINT_NAME, INT_ERROR_DESC_UNDERFLOW
        Exit Sub
    End If

    iValue.RawValue = iValue.RawValue - Decrement
    UpdateByteArray

End Sub

Private Function Overflow(TestValue As Integer, Direction As Integer) As Boolean

    ' Catch attempted adjustments that exceed the base datatype upper limits (Integer)

    Dim tmpValue As Long
    
    tmpValue = iValue.RawValue
    tmpValue = tmpValue + (TestValue * Direction)

    If tmpValue > INT_MAX Then
        Overflow = True
    End If

End Function

Private Function Underflow(TestValue As Integer, Direction As Integer) As Boolean

    ' Catch attempted adjustments that exceed the base datatype lower limits (Integer)

    Dim tmpValue As Long

    tmpValue = iValue.RawValue
    tmpValue = tmpValue + (TestValue * Direction)

    If tmpValue < INT_MIN Then
        Underflow = True
    End If

End Function

Private Sub UpdateByteArray()

    LSet iByte = iValue

End Sub

Private Sub UpdateValue()

    LSet iValue = iByte

End Sub
