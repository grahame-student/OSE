VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SuperSingle"
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

Private Const SINGLE_SUPERSINGLE_NAME As String = "SuperSingle"

Private Const SINGLE_MIN_NEG As Single = -3.402823E+38
Private Const SINGLE_MIN_POS As Single = 1.401298E-45
Private Const SINGLE_MAX_NEG As Single = -1.401298E-45
Private Const SINGLE_MAX_POS As Single = 3.402823E+38

Private Const SINGLE_MINUS As Integer = -1
Private Const SINGLE_ADD As Integer = 1

Private Const SINGLE_ERROR_BASE        As Long = vbObjectError + &H300&
Private Const SINGLE_ERROR_OVERFLOW    As Long = SINGLE_ERROR_BASE + 1
Private Const SINGLE_ERROR_UNDERFLOW   As Long = SINGLE_ERROR_BASE + 2
Private Const SINGLE_ERROR_INVALID_ARRAY_POSITION       As Long = SINGLE_ERROR_BASE + 3

Private Const SINGLE_ERROR_DESC_OVERFLOW   As String = "This operation would cause a 'Single' overflow"
Private Const SINGLE_ERROR_DESC_UNDERFLOW  As String = "This operation would cause a 'Single' underflow"
Private Const SINGLE_ERROR_DESC_INVALID_ARRAY_POSITION  As String = "Trying to access invalid array position"

Private Const SINGLE_BYTE_SIZE As Integer = 4

Private Type FourBytes
    RawValue As Single
End Type

Private Type ByteArray
    RawByteValue(SINGLE_BYTE_SIZE - 1) As Byte
End Type

Private iValue As FourBytes
Private iByte As ByteArray

Private Sub Class_Initialize()

    iValue.RawValue = 0
    UpdateByteArray

End Sub


''' Properties
Public Property Get Value() As Single
Attribute Value.VB_ProcData.VB_Invoke_Property = ";Data"
Attribute Value.VB_UserMemId = 0

    Value = iValue.RawValue

End Property

Public Property Let Value(NewValue As Single)

    iValue.RawValue = NewValue
    UpdateByteArray

End Property

' ByteValue property
Public Property Get ByteValue(Index As Integer) As Byte

    If Index < 0 Then
        Err.Raise SINGLE_ERROR_INVALID_ARRAY_POSITION, SINGLE_SUPERSINGLE_NAME, SINGLE_ERROR_DESC_INVALID_ARRAY_POSITION
    ElseIf Index >= SINGLE_BYTE_SIZE Then
        Err.Raise SINGLE_ERROR_INVALID_ARRAY_POSITION, SINGLE_SUPERSINGLE_NAME, SINGLE_ERROR_DESC_INVALID_ARRAY_POSITION
    End If

    ByteValue = iByte.RawByteValue(Index)

End Property

Public Property Let ByteValue(Index As Integer, NewByteValue As Byte)

    If Index < 0 Then
        Err.Raise SINGLE_ERROR_INVALID_ARRAY_POSITION, SINGLE_SUPERSINGLE_NAME, SINGLE_ERROR_DESC_INVALID_ARRAY_POSITION
    ElseIf Index >= SINGLE_BYTE_SIZE Then
        Err.Raise SINGLE_ERROR_INVALID_ARRAY_POSITION, SINGLE_SUPERSINGLE_NAME, SINGLE_ERROR_DESC_INVALID_ARRAY_POSITION
    End If

    iByte.RawByteValue(Index) = NewByteValue
    UpdateValue

End Property

''' Methods
Public Sub Add(Optional Increment As Single = 1)

    If Overflow(Increment, SINGLE_ADD) Then
        Err.Raise SINGLE_ERROR_OVERFLOW, SINGLE_SUPERSINGLE_NAME, SINGLE_ERROR_DESC_OVERFLOW
        Exit Sub
    ElseIf Underflow(Increment, SINGLE_ADD) Then
        Err.Raise SINGLE_ERROR_UNDERFLOW, SINGLE_SUPERSINGLE_NAME, SINGLE_ERROR_DESC_UNDERFLOW
        Exit Sub
    End If

    iValue.RawValue = iValue.RawValue + Increment
    UpdateByteArray

End Sub

Public Sub Minus(Optional Decrement As Single = 1)

    If Overflow(Decrement, SINGLE_MINUS) Then
        Err.Raise SINGLE_ERROR_OVERFLOW, SINGLE_SUPERSINGLE_NAME, SINGLE_ERROR_DESC_OVERFLOW
        Exit Sub
    ElseIf Underflow(Decrement, SINGLE_MINUS) Then
        Err.Raise SINGLE_ERROR_UNDERFLOW, SINGLE_SUPERSINGLE_NAME, SINGLE_ERROR_DESC_UNDERFLOW
        Exit Sub
    End If

    iValue.RawValue = iValue.RawValue - Decrement
    UpdateByteArray

End Sub

Private Function Overflow(TestValue As Single, Direction As Integer) As Boolean

    Dim tmpValue As Double
    
    tmpValue = iValue.RawValue
    tmpValue = tmpValue + (TestValue * Direction)

    If tmpValue > SINGLE_MAX_POS Then
        Overflow = True
    End If

End Function

Private Function Underflow(TestValue As Single, Direction As Integer) As Boolean

    Dim tmpValue As Double

    tmpValue = iValue.RawValue
    tmpValue = tmpValue + (TestValue * Direction)

    If tmpValue < SINGLE_MIN_NEG Then
        Underflow = True
    End If

End Function

Private Sub UpdateByteArray()

    LSet iByte = iValue

End Sub

Private Sub UpdateValue()

    LSet iValue = iByte

End Sub

