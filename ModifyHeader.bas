Attribute VB_Name = "ModifyHeader"
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

' The routines to actually modify the header portion of the data structure
' all live in this file (this includes the FileHeader and SaveHeader sections)

Public Sub ModifySaveFileVersionMajor(ByVal NewValue As Byte)

    SaveFileData.FileHeader.MajorVersion = NewValue

End Sub

Public Sub ModifySaveFileVersionMinor(ByVal NewValue As Byte)

    SaveFileData.FileHeader.MinorVersion = NewValue

End Sub

Public Sub ModifySaveFileHeaderVersion(ByVal NewValue As Byte)

    SaveFileData.SaveHeader.HeaderVersion = NewValue

End Sub

Public Sub ModifySaveFileNumber(ByVal NewValue As Long)

    SaveFileData.SaveHeader.SaveNumber = NewValue

End Sub

Public Sub ModifySaveFilePlayerName(ByVal NewName As String)

    Dim LengthDifference As Integer

    ' Make sure the new name is null terminated
    If Right(NewName, 1) <> Chr$(0) Then
        NewName = NewName + Chr$(0)
    End If

    ' Calculate the difference in length between the new name and the old name
    LengthDifference = Len(NewName) - Len(SaveFileData.SaveHeader.PlayerName)

    SaveFileData.SaveHeader.PlayerName = NewName
    
    ' Add the length difference to the headersize variable
    SaveFileData.SaveHeader.SaveHeaderSize = SaveFileData.SaveHeader.SaveHeaderSize + LengthDifference

End Sub

Public Sub ModifySaveFilePlayerLocation(ByVal NewLocation As String)

    Dim LengthDifference As Integer

    ' Make sure the new location is null terminated
    If Right(NewLocation, 1) <> Chr$(0) Then
        NewLocation = NewLocation + Chr$(0)
    End If

    ' Calculate the difference in length between the new location and the old location
    LengthDifference = Len(NewLocation) - Len(SaveFileData.SaveHeader.PlayerLocation)

    SaveFileData.SaveHeader.PlayerLocation = NewLocation
    
    ' Add the length difference to the headersize variable
    SaveFileData.SaveHeader.SaveHeaderSize = SaveFileData.SaveHeader.SaveHeaderSize + LengthDifference

End Sub

Public Sub ModifySaveFilePlayerLevel(ByVal NewValue As Integer)

    SaveFileData.SaveHeader.PlayerLevel = NewValue

End Sub

Public Sub ModifySaveFileGameTime(ByRef CallingForm As Form)

    SaveFileData.SaveHeader.GameTime.Year = CallingForm.txtYear.Text
    SaveFileData.SaveHeader.GameTime.Month = CallingForm.txtMonth.Text
    SaveFileData.SaveHeader.GameTime.DayOfWeek = CallingForm.txtDayOfWeek.Text
    SaveFileData.SaveHeader.GameTime.Day = CallingForm.txtDay.Text
    SaveFileData.SaveHeader.GameTime.Hour = CallingForm.txtHour.Text
    SaveFileData.SaveHeader.GameTime.Minute = CallingForm.txtMinute.Text
    SaveFileData.SaveHeader.GameTime.Second = CallingForm.txtSecond.Text
    SaveFileData.SaveHeader.GameTime.MilliSecond = CallingForm.txtMillisecond.Text

End Sub

Public Sub ModifySaveFileEXETime(ByRef CallingForm As Form)

    SaveFileData.FileHeader.EXETime.Year = CallingForm.txtYear.Text
    SaveFileData.FileHeader.EXETime.Month = CallingForm.txtMonth.Text
    SaveFileData.FileHeader.EXETime.DayOfWeek = CallingForm.txtDayOfWeek.Text
    SaveFileData.FileHeader.EXETime.Day = CallingForm.txtDay.Text
    SaveFileData.FileHeader.EXETime.Hour = CallingForm.txtHour.Text
    SaveFileData.FileHeader.EXETime.Minute = CallingForm.txtMinute.Text
    SaveFileData.FileHeader.EXETime.Second = CallingForm.txtSecond.Text
    SaveFileData.FileHeader.EXETime.MilliSecond = CallingForm.txtMillisecond.Text

End Sub

Public Sub ModifySavefileScreenShot(ByRef ScreenShot As PictureBox)

    Dim PixelColour As Long
    Dim PixelAddress As Long
    Dim X As Long
    Dim Y As Long

    SaveFileData.SaveHeader.ScreenShot.Height = ScreenShot.ScaleHeight
    SaveFileData.SaveHeader.ScreenShot.Width = ScreenShot.Width
    SaveFileData.SaveHeader.ScreenShot.Size = SaveFileData.SaveHeader.ScreenShot.Height * _
                                              SaveFileData.SaveHeader.ScreenShot.Width + 8

    ReDim SaveFileData.SaveHeader.ScreenShot.Pixel(SaveFileData.SaveHeader.ScreenShot.Size - 8)

    For Y = 0 To SaveFileData.SaveHeader.ScreenShot.Height - 1
        For X = 0 To SaveFileData.SaveHeader.ScreenShot.Width - 1
            PixelAddress = Y * SaveFileData.SaveHeader.ScreenShot.Width + X
            PixelColour = ScreenShot.Point(X, Y)
            SaveFileData.SaveHeader.ScreenShot.Pixel(PixelAddress).Red = CByte(PixelColour And &HFF&)
            SaveFileData.SaveHeader.ScreenShot.Pixel(PixelAddress).Green = CByte((PixelColour And &HFF00&) / BYTE_2)
            SaveFileData.SaveHeader.ScreenShot.Pixel(PixelAddress).Blue = CByte((PixelColour And &HFF0000) / BYTE_3)
        Next X
    Next Y

End Sub

