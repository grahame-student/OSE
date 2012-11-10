Attribute VB_Name = "DataTypes"
Option Explicit
DefObj A-Z

Public Type bByteType
    Bytes(3) As Byte
End Type
 
Public Type SingleType
    Result As Single
End Type

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


