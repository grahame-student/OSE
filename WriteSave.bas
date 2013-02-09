Attribute VB_Name = "WriteSave"
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

Public Sub WriteSaveFile(ByVal SaveFilePath As String, ByRef Status As StatusBar, ByRef Progress As ProgressBar)

    If SaveFilePath = "" Then Exit Sub

    Progress.Value = Progress.Min
    
    FF = FreeFile

    Open SaveFilePath For Binary Access Read Write Lock Write As FF

    Status.Panels(STB_STATUS).Text = "Saving FileHeader..."
    WriteSaveFileHeader
    
    Status.Panels(STB_STATUS).Text = "Saving SaveHeader..."
    WriteSaveSaveHeader
    
    Status.Panels(STB_STATUS).Text = "Saving PlugIns..."
    WriteSavePlugIns

    Status.Panels(STB_STATUS).Text = "Saving Globals..."
    WriteSaveGlobals

    Status.Panels(STB_STATUS).Text = "Saving Change Records..."
    WriteSaveChangeRecords Progress

    Status.Panels(STB_STATUS).Text = "Loading Temporary Effects..."
    WriteSaveTempEffects
    
    Status.Panels(STB_STATUS).Text = "Loading FormIDs..."
    WriteSaveFormIDs
    
    Status.Panels(STB_STATUS).Text = "Loading World Spaces..."
    WriteSaveWorldSpaces
    
    Progress.Value = Progress.Max
    
    Status.Panels(STB_STATUS).Text = "Save Completed..."
    Close #FF
    
End Sub

Private Sub WriteSaveFileHeader()

    PutNextFixedLengthString SaveFileData.FileHeader.FileID
    PutNextUByte SaveFileData.FileHeader.MajorVersion
    PutNextUByte SaveFileData.FileHeader.MinorVersion
    PutNextSystemTime SaveFileData.FileHeader.EXETime

End Sub

Private Sub WriteSaveSaveHeader()

    PutNext32BitULong SaveFileData.SaveHeader.HeaderVersion
    PutNext32BitULong SaveFileData.SaveHeader.SaveHeaderSize
    PutNext32BitULong SaveFileData.SaveHeader.SaveNumber
    
    PutNextUByte Len(SaveFileData.SaveHeader.PlayerName)
    PutNextFixedLengthString SaveFileData.SaveHeader.PlayerName
    
    PutNext16BitUInteger SaveFileData.SaveHeader.PlayerLevel
    
    PutNextUByte Len(SaveFileData.SaveHeader.PlayerLocation)
    PutNextFixedLengthString SaveFileData.SaveHeader.PlayerLocation
    PutNext32BitSingle SaveFileData.SaveHeader.GameDays
    PutNext32BitULong SaveFileData.SaveHeader.GameTicks
    PutNextSystemTime SaveFileData.SaveHeader.GameTime
    PutNextScreenShot SaveFileData.SaveHeader.ScreenShot

End Sub

Private Sub WriteSavePlugIns()

    Dim i As Integer

    PutNextUByte SaveFileData.PlugIns.NumberOfPlugins

    For i = 0 To SaveFileData.PlugIns.NumberOfPlugins - 1
        PutNextUByte Len(SaveFileData.PlugIns.PlugInNames(i))
        PutNextFixedLengthString SaveFileData.PlugIns.PlugInNames(i)
    Next i

End Sub

Private Sub WriteSaveGlobals()

    PutNext32BitULong SaveFileData.Globals.FormIDOffset
    PutNext32BitULong SaveFileData.Globals.NumberOfChangeRecords
    PutNext32BitULong SaveFileData.Globals.NextObjectID
    PutNext32BitULong SaveFileData.Globals.WorldID
    PutNext32BitULong SaveFileData.Globals.WorldX
    PutNext32BitULong SaveFileData.Globals.WorldY

    PutNext32BitULong SaveFileData.Globals.PlayerLocation.Cell
    PutNext32BitSingle SaveFileData.Globals.PlayerLocation.X
    PutNext32BitSingle SaveFileData.Globals.PlayerLocation.Y
    PutNext32BitSingle SaveFileData.Globals.PlayerLocation.Z

    PutNext16BitUInteger SaveFileData.Globals.GlobalsNumber
    WriteSaveGlobalsGlobals
    
    PutNext16BitUInteger SaveFileData.Globals.ClassSize

    PutNext32BitULong SaveFileData.Globals.NumberOfDeathCounts
    WriteSaveGlobalsDeathCounts
    
    PutNext32BitSingle SaveFileData.Globals.GameModeSeconds

    PutNext16BitUInteger SaveFileData.Globals.ProcessesSize
    WriteSaveGlobalsProcessesData

    PutNext16BitUInteger SaveFileData.Globals.SpectatorEventSize
    WriteSaveGlobalsSpectatorEventData

    PutNext16BitUInteger SaveFileData.Globals.WeatherSize
    WriteSaveGlobalsWeatherData

    PutNext32BitULong SaveFileData.Globals.PlayerCombatCount
    PutNext32BitULong SaveFileData.Globals.CreatedNumber
    WriteSaveGlobalsCreatedData

    PutNext16BitUInteger SaveFileData.Globals.QuickKeySize
    WriteSaveGlobalsQuickKeyData

    PutNext16BitUInteger SaveFileData.Globals.ReticuleSize
    WriteSaveGlobalsReticuleData

    PutNext16BitUInteger SaveFileData.Globals.InterfaceSize
    WriteSaveGlobalsInterfaceData

    PutNext16BitUInteger SaveFileData.Globals.RegionSize
    PutNext16BitUInteger SaveFileData.Globals.RegionNumber
    WriteSaveGlobalsRegionData

End Sub

Private Sub WriteSaveGlobalsGlobals()

    Dim i As Integer

    For i = 0 To SaveFileData.Globals.GlobalsNumber - 1
        PutNext32BitULong SaveFileData.Globals.Globals(i).iRef
        PutNext32BitSingle SaveFileData.Globals.Globals(i).Value
    Next i

End Sub

Private Sub WriteSaveGlobalsDeathCounts()

    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.NumberOfDeathCounts - 1
        PutNext32BitULong SaveFileData.Globals.DeathCounts(i).Actor
        PutNext16BitUInteger SaveFileData.Globals.DeathCounts(i).Count
    Next i

End Sub

Private Sub WriteSaveGlobalsProcessesData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.ProcessesSize - 1
        PutNextUByte SaveFileData.Globals.ProcessesData(i)
    Next i
    
End Sub

Private Sub WriteSaveGlobalsSpectatorEventData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.SpectatorEventSize - 1
        PutNextUByte SaveFileData.Globals.SpectatorEventData(i)
    Next i
    
End Sub

Private Sub WriteSaveGlobalsWeatherData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.WeatherSize - 1
        PutNextUByte SaveFileData.Globals.WeatherData(i)
    Next i
    
End Sub

Private Sub WriteSaveGlobalsCreatedData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.CreatedNumber - 1
        PutNextFixedLengthString SaveFileData.Globals.CreatedData(i).Type
        PutNext32BitULong SaveFileData.Globals.CreatedData(i).Size
        PutNext32BitULong SaveFileData.Globals.CreatedData(i).Flags
        PutNext32BitULong SaveFileData.Globals.CreatedData(i).FormID
        PutNext32BitULong SaveFileData.Globals.CreatedData(i).VersionControlInfo
        WriteSaveGlobalsCreatedDataData i
    Next i

End Sub

Private Sub WriteSaveGlobalsCreatedDataData(ByVal ItemNumber As Integer)

    Dim i As Integer

    For i = 0 To SaveFileData.Globals.CreatedData(ItemNumber).Size - 1
        PutNextUByte SaveFileData.Globals.CreatedData(ItemNumber).Data(i)
    Next i

End Sub

Private Sub WriteSaveGlobalsQuickKeyData()

    Dim i As Integer
    Dim Size As Integer
    
    i = 0
    
    Do Until Size = SaveFileData.Globals.QuickKeySize
        PutNextUByte SaveFileData.Globals.QuickKeyData(i).Flag
        Size = Size + 1
        If SaveFileData.Globals.QuickKeyData(i).Flag = 0 Then
'            PutNextUByte SaveFileData.Globals.QuickKeyData(i).Reference
'            Size = Size + 1
        Else
            PutNext32BitULong SaveFileData.Globals.QuickKeyData(i).Reference
            Size = Size + 4
        End If
        i = i + 1
    Loop

End Sub

Private Sub WriteSaveGlobalsReticuleData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.ReticuleSize - 1
        PutNextUByte SaveFileData.Globals.ReticuleData(i)
    Next i
    
End Sub

Private Sub WriteSaveGlobalsInterfaceData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.InterfaceSize - 1
        PutNextUByte SaveFileData.Globals.InterfaceData(i)
    Next i
    
End Sub

Private Sub WriteSaveGlobalsRegionData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.RegionNumber - 1
        PutNext32BitULong SaveFileData.Globals.Regions(i).Reference
        PutNext32BitULong SaveFileData.Globals.Regions(i).Unknown
    Next i
    
End Sub

Private Sub WriteSaveChangeRecords(ByRef Progress As ProgressBar)

    Dim i As Long
    Dim ProgressValue As Single

    Progress.Value = Progress.Min

    For i = 0 To SaveFileData.Globals.NumberOfChangeRecords - 1
        PutNext32BitULong SaveFileData.ChangeRecords(i).FormID
        PutNextUByte SaveFileData.ChangeRecords(i).Type
        PutNext32BitULong SaveFileData.ChangeRecords(i).Flags
        PutNextUByte SaveFileData.ChangeRecords(i).Version
        PutNext16BitUInteger SaveFileData.ChangeRecords(i).DataSize
        WriteSaveChangeRecordsData i
        ProgressValue = Int((i / SaveFileData.Globals.NumberOfChangeRecords) * Progress.Max)
        Progress.Value = ProgressValue
        DoEvents
    Next i

    Progress.Value = ProgressValue

End Sub

Private Sub WriteSaveChangeRecordsData(ByVal RecordNumber As Long)

    Dim i As Integer

    For i = 0 To SaveFileData.ChangeRecords(RecordNumber).DataSize - 1
        PutNextUByte SaveFileData.ChangeRecords(RecordNumber).Data(i)
    Next i

End Sub

Private Sub WriteSaveTempEffects()

    Dim i As Long

    PutNext32BitULong SaveFileData.TempEffects.Size
    For i = 0 To SaveFileData.TempEffects.Size - 1
        PutNextUByte SaveFileData.TempEffects.Data(i)
    Next i

End Sub

Private Sub WriteSaveFormIDs()

    Dim i As Long

    PutNext32BitULong SaveFileData.FormIDs.NumberOfFormIDs
    For i = 0 To SaveFileData.FormIDs.NumberOfFormIDs - 1
        PutNext32BitULong SaveFileData.FormIDs.FormIDsList(i)
    Next i

End Sub

Private Sub WriteSaveWorldSpaces()

    Dim i As Long

    PutNext32BitULong SaveFileData.WorldSpaces.NumberOfWorldSpaces
    For i = 0 To SaveFileData.WorldSpaces.NumberOfWorldSpaces - 1
        PutNext32BitULong SaveFileData.WorldSpaces.WorldSpaces(i)
    Next i

End Sub

Private Sub PutNextFixedLengthString(ByVal StringValue As String)

    Put #FF, , StringValue

End Sub

Private Sub PutNextUByte(ByVal ByteValue As Byte)

    Put #FF, , ByteValue

End Sub

Private Sub PutNextSystemTime(ByRef SystemTimeValue As SystemTime)

    PutNext16BitUInteger SystemTimeValue.Year
    PutNext16BitUInteger SystemTimeValue.Month
    PutNext16BitUInteger SystemTimeValue.DayOfWeek
    PutNext16BitUInteger SystemTimeValue.Day
    PutNext16BitUInteger SystemTimeValue.Hour
    PutNext16BitUInteger SystemTimeValue.Minute
    PutNext16BitUInteger SystemTimeValue.Second
    PutNext16BitUInteger SystemTimeValue.MilliSecond

End Sub

Private Sub PutNext16BitUInteger(ByVal IntergerValue As Integer)
    
    Put #FF, , IntergerValue

End Sub

Private Sub PutNext32BitULong(ByVal LongValue As Long)
    
    Put #FF, , LongValue

End Sub

Private Sub PutNext32BitSingle(ByVal SingleValue As Single)

    Put #FF, , SingleValue
        
End Sub

Private Sub PutNextScreenShot(ByRef ScreenShotValue As ScreenShot)

    Dim X As Long
    Dim Y As Long
    
    PutNext32BitULong ScreenShotValue.Size
    PutNext32BitULong ScreenShotValue.Width
    PutNext32BitULong ScreenShotValue.Height

    For Y = 0 To ScreenShotValue.Height - 1
        For X = 0 To ScreenShotValue.Width - 1
            PutNextUByte ScreenShotValue.Pixel((Y * ScreenShotValue.Width) + X).Red
            PutNextUByte ScreenShotValue.Pixel((Y * ScreenShotValue.Width) + X).Green
            PutNextUByte ScreenShotValue.Pixel((Y * ScreenShotValue.Width) + X).Blue
        Next X
        DoEvents
    Next Y

End Sub

