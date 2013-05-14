VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperDouble"
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

Private Const DOUBLE_SUPERDOUBLE_NAME As String = "SuperDouble"

Private Const DOUBLE_MIN_NEG As Double = -1.79769313486231E+308
Private Const DOUBLE_MIN_POS As Double = 1.79769313486231E+308
Private Const DOUBLE_MAX_NEG As Double = -4.94065645841247E-324
Private Const DOUBLE_MAX_POS As Double = 4.94065645841247E-324

Private Const DOUBLE_MINUS As Integer = -1
Private Const DOUBLE_ADD As Integer = 1

Private Const DOUBLE_ERROR_BASE                         As Long = vbObjectError + &H500&
Private Const DOUBLE_ERROR_OVERFLOW                     As Long = DOUBLE_ERROR_BASE + 1
Private Const DOUBLE_ERROR_UNDERFLOW                    As Long = DOUBLE_ERROR_BASE + 2
Private Const DOUBLE_ERROR_INVALID_ARRAY_POSITION       As Long = DOUBLE_ERROR_BASE + 3

Private Const DOUBLE_ERROR_DESC_OVERFLOW                As String = "This operation would cause a 'Double' overflow"
Private Const DOUBLE_ERROR_DESC_UNDERFLOW               As String = "This operation would cause a 'Double' underflow"
Private Const DOUBLE_ERROR_DESC_INVALID_ARRAY_POSITION  As String = "Trying to access invalid array position"

Private Const DOUBLE_BYTE_SIZE As Integer = 8

Private Type EightBytes
    RawValue As Double
End Type

Private Type ByteArray
    RawByteValue(DOUBLE_BYTE_SIZE - 1) As Byte
End Type

Private iValue As EightBytes
Private iByte As ByteArray

Private Sub Class_Initialize()

    iValue.RawValue = 0
    UpdateByteArray

End Sub


''' Properties
Public Property Get Value() As Double

    Value = iValue.RawValue

End Property

Public Property Let Value(NewValue As Double)

    iValue.RawValue = NewValue
    UpdateByteArray

End Property

' ByteValue property
Public Property Get ByteValue(Index As Integer) As Byte

    If Index < 0 Then
        Err.Raise DOUBLE_ERROR_INVALID_ARRAY_POSITION, DOUBLE_SUPERDOUBLE_NAME, DOUBLE_ERROR_DESC_INVALID_ARRAY_POSITION
    ElseIf Index >= DOUBLE_BYTE_SIZE Then
        Err.Raise DOUBLE_ERROR_INVALID_ARRAY_POSITION, DOUBLE_SUPERDOUBLE_NAME, DOUBLE_ERROR_DESC_INVALID_ARRAY_POSITION
    End If

    ByteValue = iByte.RawByteValue(Index)

End Property

Public Property Let ByteValue(Index As Integer, NewByteValue As Byte)

    If Index < 0 Then
        Err.Raise DOUBLE_ERROR_INVALID_ARRAY_POSITION, DOUBLE_SUPERDOUBLE_NAME, DOUBLE_ERROR_DESC_INVALID_ARRAY_POSITION
    ElseIf Index >= DOUBLE_BYTE_SIZE Then
        Err.Raise DOUBLE_ERROR_INVALID_ARRAY_POSITION, DOUBLE_SUPERDOUBLE_NAME, DOUBLE_ERROR_DESC_INVALID_ARRAY_POSITION
    End If

    iByte.RawByteValue(Index) = NewByteValue
    UpdateValue

End Property

''' Methods
Public Sub Add(Optional Increment As Double = 1)

    If Overflow(Increment, DOUBLE_ADD) Then
        Err.Raise DOUBLE_ERROR_OVERFLOW, DOUBLE_SUPERDOUBLE_NAME, DOUBLE_ERROR_DESC_OVERFLOW
        Exit Sub
    ElseIf Underflow(Increment, DOUBLE_ADD) Then
        Err.Raise DOUBLE_ERROR_UNDERFLOW, DOUBLE_SUPERDOUBLE_NAME, DOUBLE_ERROR_DESC_UNDERFLOW
        Exit Sub
    End If

    iValue.RawValue = iValue.RawValue + Increment
    UpdateByteArray

End Sub

Public Sub Minus(Optional Decrement As Double = 1)

    If Overflow(Decrement, DOUBLE_MINUS) Then
        Err.Raise DOUBLE_ERROR_OVERFLOW, DOUBLE_SUPERDOUBLE_NAME, DOUBLE_ERROR_DESC_OVERFLOW
        Exit Sub
    ElseIf Underflow(Decrement, DOUBLE_MINUS) Then
        Err.Raise DOUBLE_ERROR_UNDERFLOW, DOUBLE_SUPERDOUBLE_NAME, DOUBLE_ERROR_DESC_UNDERFLOW
        Exit Sub
    End If

    iValue.RawValue = iValue.RawValue - Decrement
    UpdateByteArray

End Sub

Private Function Overflow(TestValue As Double, Direction As Integer) As Boolean

    Dim tmpValue As Variant
    
    tmpValue = CDec(iValue.RawValue)
    tmpValue = CDec(tmpValue + (TestValue * Direction))

    If tmpValue > DOUBLE_MAX_POS Then
        Overflow = True
    End If

End Function

Private Function Underflow(TestValue As Double, Direction As Integer) As Boolean

    Dim tmpValue As Double

    tmpValue = CDec(iValue.RawValue)
    tmpValue = CDec(tmpValue + (TestValue * Direction))

    If tmpValue < DOUBLE_MIN_NEG Then
        Underflow = True
    End If

End Function

Private Sub UpdateByteArray()

    LSet iByte = iValue

End Sub

Private Sub UpdateValue()

    LSet iValue = iByte

End Sub

