VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperLong"
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

Private Const LONG_SUPERLONG_NAME As String = "SuperLong"

Private Const LONG_MIN As Long = -2147483648#
Private Const LONG_MAX As Long = 2147483647

Private Const LONG_MINUS As Integer = -1
Private Const LONG_ADD As Integer = 1

Private Const LONG_ERROR_BASE                           As Long = vbObjectError + &H200&
Private Const LONG_ERROR_OVERFLOW                       As Long = LONG_ERROR_BASE + 1
Private Const LONG_ERROR_UNDERFLOW                      As Long = LONG_ERROR_BASE + 2
Private Const LONG_ERROR_INVALID_ARRAY_POSITION         As Long = LONG_ERROR_BASE + 3

Private Const LONG_ERROR_DESC_OVERFLOW                  As String = "This operation would cause a 'Long' overflow"
Private Const LONG_ERROR_DESC_UNDERFLOW                 As String = "This operation would cause a 'Long' underflow"
Private Const LONG_ERROR_DESC_INVALID_ARRAY_POSITION    As String = "Trying to access invalid array position"

Private Const LONG_BYTE_SIZE As Integer = 4

Private Type FourBytes
    RawValue As Long
End Type

Private Type ByteArray
    RawByteValue(LONG_BYTE_SIZE - 1) As Byte
End Type

Private iValue As FourBytes
Private iByte As ByteArray

Private Sub Class_Initialize()

    iValue.RawValue = 0
    UpdateByteArray

End Sub


''' Properties
Public Property Get Value() As Long
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Value.VB_UserMemId = 0

    Value = iValue.RawValue

End Property

Public Property Let Value(NewValue As Long)

    iValue.RawValue = NewValue
    UpdateByteArray

End Property

' ByteValue property
Public Property Get ByteValue(Index As Integer) As Byte

    If Index < 0 Then
        Err.Raise LONG_ERROR_INVALID_ARRAY_POSITION, LONG_SUPERLONG_NAME, LONG_ERROR_DESC_INVALID_ARRAY_POSITION
    ElseIf Index >= LONG_BYTE_SIZE Then
        Err.Raise LONG_ERROR_INVALID_ARRAY_POSITION, LONG_SUPERLONG_NAME, LONG_ERROR_DESC_INVALID_ARRAY_POSITION
    End If

    ByteValue = iByte.RawByteValue(Index)

End Property

Public Property Let ByteValue(Index As Integer, NewByteValue As Byte)

    If Index < 0 Then
        Err.Raise LONG_ERROR_INVALID_ARRAY_POSITION, LONG_SUPERLONG_NAME, LONG_ERROR_DESC_INVALID_ARRAY_POSITION
    ElseIf Index >= LONG_BYTE_SIZE Then
        Err.Raise LONG_ERROR_INVALID_ARRAY_POSITION, LONG_SUPERLONG_NAME, LONG_ERROR_DESC_INVALID_ARRAY_POSITION
    End If

    iByte.RawByteValue(Index) = NewByteValue
    UpdateValue

End Property


''' Methods
Public Sub Add(Optional Increment As Long = 1)

    If Overflow(Increment, LONG_ADD) Then
        Err.Raise LONG_ERROR_OVERFLOW, LONG_SUPERLONG_NAME, LONG_ERROR_DESC_OVERFLOW
        Exit Sub
    ElseIf Underflow(Increment, LONG_ADD) Then
        Err.Raise LONG_ERROR_UNDERFLOW, LONG_SUPERLONG_NAME, LONG_ERROR_DESC_UNDERFLOW
        Exit Sub
    End If

    iValue.RawValue = iValue.RawValue + Increment
    UpdateByteArray

End Sub

Public Sub Minus(Optional Decrement As Long = 1)

    If Overflow(Decrement, LONG_MINUS) Then
        Err.Raise LONG_ERROR_OVERFLOW, LONG_SUPERLONG_NAME, LONG_ERROR_DESC_OVERFLOW
        Exit Sub
    ElseIf Underflow(Decrement, LONG_MINUS) Then
        Err.Raise LONG_ERROR_UNDERFLOW, LONG_SUPERLONG_NAME, LONG_ERROR_DESC_UNDERFLOW
        Exit Sub
    End If

    iValue.RawValue = iValue.RawValue - Decrement
    UpdateByteArray

End Sub

Private Function Overflow(TestValue As Long, Direction As Integer) As Boolean

    Dim tmpValue As Currency
    
    tmpValue = iValue.RawValue
    tmpValue = tmpValue + (TestValue * Direction)

    If tmpValue > LONG_MAX Then
        Overflow = True
    End If

End Function

Private Function Underflow(TestValue As Long, Direction As Integer) As Boolean

    Dim tmpValue As Currency

    tmpValue = iValue.RawValue
    tmpValue = tmpValue + (TestValue * Direction)

    If tmpValue < LONG_MIN Then
        Underflow = True
    End If

End Function

Private Sub UpdateByteArray()

    LSet iByte = iValue

End Sub

Private Sub UpdateValue()

    LSet iValue = iByte

End Sub

