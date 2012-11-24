Attribute VB_Name = "DataTypes"
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

Public Type SystemTime
    Year As Integer
    Month As Integer
    DayOfWeek As Integer
    Day As Integer
    Hour As Integer
    Minute As Integer
    Second As Integer
    MilliSecond As Integer
End Type

Public Type Pixel
    Red As Byte
    Green As Byte
    Blue As Byte
End Type

Public Type ScreenShot
    Size As Long
    Width As Long
    Height As Long
    Pixel() As Pixel
End Type

Public Type LongType
    Result As Long
End Type

Public Type ByteArray
    Bytes(3) As Byte
End Type
          
Public Type FactionEntry
    FormID As Long
    Name As String
    PlugIn As String
    MaxRank As Integer
    Ranks() As String
End Type

Public Type SpellEntry
    FormID As Long
    Name As String
    PlugIn As String
End Type
