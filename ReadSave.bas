Attribute VB_Name = "ReadSave"
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

' All the code to read an oblivion save file
Public Sub ReadSaveFile(ByVal SaveFilePath As String, ByRef Status As StatusBar, ByRef Progress As ProgressBar)

    Progress.Value = Progress.Min
    
    FF = FreeFile

    Open SaveFilePath For Binary Access Read Write Lock Write As FF

    Status.Panels(STB_STATUS).Text = "Loading FileHeader..."
    ReadSaveFileHeader
    Select Case FileIDInvalid
        Case FILE_ID_OK
            ' All okay, continue with the rest of the load routine
        Case FILE_ID_XBOX360
            MsgBox "XBox360 container detected, unable to proceed", vbOKOnly, "XBox360 Container Detected"
            Status.Panels(1).Text = "XBox360 container detected"
            SaveFileData.OSE.LoadSuccessful = False
            Exit Sub
        Case FILE_ID_UNKNOWN
            MsgBox "Unknown FileID detected, unable to proceed", vbOKOnly, "Unknown FileID Detected"
            Status.Panels(STB_STATUS).Text = "Unknown FileID detected"
            SaveFileData.OSE.LoadSuccessful = False
            Exit Sub
    End Select
    
    Status.Panels(STB_STATUS).Text = "Loading SaveHeader..."
    ReadSaveSaveHeader
    
    Status.Panels(STB_STATUS).Text = "Loading PlugIns..."
    ReadSavePlugIns
    
    Status.Panels(STB_STATUS).Text = "Loading Globals..."
    ReadSaveGlobals
    ReDim SaveFileData.ChangeRecords(SaveFileData.Globals.NumberOfChangeRecords - 1)
    
    Status.Panels(STB_STATUS).Text = "Loading Change Records..."
    ReadSaveChangeRecords Progress
    
    Status.Panels(STB_STATUS).Text = "Loading Temporary Effects..."
    ReadSaveTempEffects
    
    Status.Panels(STB_STATUS).Text = "Loading FormIDs..."
    ReadSaveFormIDs
    
    Status.Panels(STB_STATUS).Text = "Loading World Spaces..."
    ReadSaveWorldSpaces
        
    Progress.Value = Progress.Max
    SaveFileData.OSE.LoadSuccessful = True
        
    Close #FF
        
End Sub

Private Sub ReadSaveFileHeader()

    SaveFileData.FileHeader.FileID = GetNextFixedLengthString(12)
    SaveFileData.FileHeader.MajorVersion = GetNextUByte
    SaveFileData.FileHeader.MinorVersion = GetNextUByte
    SaveFileData.FileHeader.EXETime = GetNextSystemTime

End Sub

Private Function FileIDInvalid() As Integer

    If SaveFileData.FileHeader.FileID = "TES4SAVEGAME" Then
        FileIDInvalid = FILE_ID_OK
    ElseIf Mid$(SaveFileData.FileHeader.FileID, 1, 3) = "CON" Then
        FileIDInvalid = FILE_ID_XBOX360
    Else
        FileIDInvalid = FILE_ID_UNKNOWN
    End If

End Function

Private Sub ReadSaveSaveHeader()

    SaveFileData.SaveHeader.HeaderVersion = GetNext32BitULong
    SaveFileData.SaveHeader.SaveHeaderSize = GetNext32BitULong
    SaveFileData.SaveHeader.SaveNumber = GetNext32BitULong
    SaveFileData.SaveHeader.PlayerName = GetNextFixedLengthString(GetNextUByte)
    SaveFileData.SaveHeader.PlayerLevel = GetNext16BitUInteger
    SaveFileData.SaveHeader.PlayerLocation = GetNextFixedLengthString(GetNextUByte)
    SaveFileData.SaveHeader.GameDays = GetNext32BitSingle
    SaveFileData.SaveHeader.GameTicks = GetNext32BitULong
    SaveFileData.SaveHeader.GameTime = GetNextSystemTime
    SaveFileData.SaveHeader.ScreenShot = GetNextScreenShot

End Sub

Private Sub ReadSavePlugIns()

    Dim i As Integer

    SaveFileData.PlugIns.NumberOfPlugins = GetNextUByte

    ' Now we know how many plugins are used make room for their names
    ReDim SaveFileData.PlugIns.PlugInNames(SaveFileData.PlugIns.NumberOfPlugins - 1)

    For i = 0 To SaveFileData.PlugIns.NumberOfPlugins - 1
        SaveFileData.PlugIns.PlugInNames(i) = GetNextFixedLengthString(GetNextUByte)
    Next i

End Sub

Private Sub ReadSaveGlobals()

    SaveFileData.Globals.FormIDOffset = GetNext32BitULong
    SaveFileData.Globals.NumberOfChangeRecords = GetNext32BitULong
    SaveFileData.Globals.NextObjectID = GetNext32BitULong
    SaveFileData.Globals.WorldID = GetNext32BitULong
    SaveFileData.Globals.WorldX = GetNext32BitULong
    SaveFileData.Globals.WorldY = GetNext32BitULong

    SaveFileData.Globals.PlayerLocation.Cell = GetNext32BitULong
    SaveFileData.Globals.PlayerLocation.X = GetNext32BitSingle
    SaveFileData.Globals.PlayerLocation.Y = GetNext32BitSingle
    SaveFileData.Globals.PlayerLocation.Z = GetNext32BitSingle

    SaveFileData.Globals.GlobalsNumber = GetNext16BitUInteger
    
    ' Allocate room for the globals now that we know how many there are
    ReDim SaveFileData.Globals.Globals(SaveFileData.Globals.GlobalsNumber - 1)

    ReadSaveGlobalsGlobals
    SaveFileData.Globals.ClassSize = GetNext16BitUInteger
    
    SaveFileData.Globals.NumberOfDeathCounts = GetNext32BitULong
    ReDim SaveFileData.Globals.DeathCounts(SaveFileData.Globals.NumberOfDeathCounts - 1)
    ReadSaveGlobalsDeathCounts
    SaveFileData.Globals.GameModeSeconds = GetNext32BitSingle
    
    SaveFileData.Globals.ProcessesSize = GetNext16BitUInteger
    ReDim SaveFileData.Globals.ProcessesData(SaveFileData.Globals.ProcessesSize - 1)
    ReadSaveGlobalsProcessesData
    
    SaveFileData.Globals.SpectatorEventSize = GetNext16BitUInteger
    ReDim SaveFileData.Globals.SpectatorEventData(SaveFileData.Globals.SpectatorEventSize - 1)
    ReadSaveGlobalsSpectatorEventData

    SaveFileData.Globals.WeatherSize = GetNext16BitUInteger
    ReDim SaveFileData.Globals.WeatherData(SaveFileData.Globals.WeatherSize - 1)
    ReadSaveGlobalsWeatherData

    SaveFileData.Globals.PlayerCombatCount = GetNext32BitULong
    SaveFileData.Globals.CreatedNumber = GetNext32BitULong
    ReDim SaveFileData.Globals.CreatedData(SaveFileData.Globals.CreatedNumber - 1)
    ReadSaveGlobalsCreatedData

    SaveFileData.Globals.QuickKeySize = GetNext16BitUInteger
    ReadSaveGlobalsQuickKeyData

    SaveFileData.Globals.ReticuleSize = GetNext16BitUInteger
    ReDim SaveFileData.Globals.ReticuleData(SaveFileData.Globals.ReticuleSize - 1)
    ReadSaveGlobalsReticuleData

    SaveFileData.Globals.InterfaceSize = GetNext16BitUInteger
    ReDim SaveFileData.Globals.InterfaceData(SaveFileData.Globals.InterfaceSize - 1)
    ReadSaveGlobalsInterfaceData
    
    SaveFileData.Globals.RegionSize = GetNext16BitUInteger
    SaveFileData.Globals.RegionNumber = GetNext16BitUInteger
    ReDim SaveFileData.Globals.Regions(SaveFileData.Globals.RegionNumber - 1)
    ReadSaveGlobalsRegionData

End Sub

Private Sub ReadSaveGlobalsGlobals()

    Dim i As Integer

    For i = 0 To SaveFileData.Globals.GlobalsNumber - 1
        SaveFileData.Globals.Globals(i).Iref = GetNext32BitULong
        SaveFileData.Globals.Globals(i).Value = GetNext32BitSingle
    Next i

End Sub

Private Sub ReadSaveGlobalsDeathCounts()

    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.NumberOfDeathCounts - 1
        SaveFileData.Globals.DeathCounts(i).Actor = GetNext32BitULong
        SaveFileData.Globals.DeathCounts(i).Count = GetNext16BitUInteger
    Next i

End Sub

Private Sub ReadSaveGlobalsProcessesData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.ProcessesSize - 1
        SaveFileData.Globals.ProcessesData(i) = GetNextUByte
    Next i
    
End Sub

Private Sub ReadSaveGlobalsSpectatorEventData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.SpectatorEventSize - 1
        SaveFileData.Globals.SpectatorEventData(i) = GetNextUByte
    Next i
    
End Sub

Private Sub ReadSaveGlobalsWeatherData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.WeatherSize - 1
        SaveFileData.Globals.WeatherData(i) = GetNextUByte
    Next i
    
End Sub

Private Sub ReadSaveGlobalsCreatedData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.CreatedNumber - 1
        SaveFileData.Globals.CreatedData(i).Type = GetNextFixedLengthString(4)
        SaveFileData.Globals.CreatedData(i).Size = GetNext32BitULong
        SaveFileData.Globals.CreatedData(i).Flags = GetNext32BitULong
        SaveFileData.Globals.CreatedData(i).FormID = GetNext32BitULong
        SaveFileData.Globals.CreatedData(i).VersionControlInfo = GetNext32BitULong
        ReDim SaveFileData.Globals.CreatedData(i).Data(SaveFileData.Globals.CreatedData(i).Size - 1)
        ReadSaveGlobalsCreatedDataData i
    Next i

End Sub

Private Sub ReadSaveGlobalsCreatedDataData(ByVal ItemNumber As Integer)

    Dim i As Integer

    For i = 0 To SaveFileData.Globals.CreatedData(ItemNumber).Size - 1
        SaveFileData.Globals.CreatedData(ItemNumber).Data(i) = GetNextUByte
    Next i

End Sub

Private Sub ReadSaveGlobalsQuickKeyData()

    Dim i As Integer
    Dim Size As Integer
    
    i = 0
    
    Do Until Size = SaveFileData.Globals.QuickKeySize
        ReDim Preserve SaveFileData.Globals.QuickKeyData(i)
        SaveFileData.Globals.QuickKeyData(i).Flag = GetNextUByte
        Size = Size + 1
        If SaveFileData.Globals.QuickKeyData(i).Flag = 0 Then
            SaveFileData.Globals.QuickKeyData(i).Reference = GetNextUByte
            Size = Size + 1
        Else
            SaveFileData.Globals.QuickKeyData(i).Reference = GetNext32BitULong
            Size = Size + 4
        End If
        i = i + 1
    Loop

End Sub

Private Sub ReadSaveGlobalsReticuleData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.ReticuleSize - 1
        SaveFileData.Globals.ReticuleData(i) = GetNextUByte
    Next i
    
End Sub

Private Sub ReadSaveGlobalsInterfaceData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.InterfaceSize - 1
        SaveFileData.Globals.InterfaceData(i) = GetNextUByte
    Next i
    
End Sub

Private Sub ReadSaveGlobalsRegionData()
    
    Dim i As Integer
    
    For i = 0 To SaveFileData.Globals.RegionNumber - 1
        SaveFileData.Globals.Regions(i).Reference = GetNext32BitULong
        SaveFileData.Globals.Regions(i).Unknown = GetNext32BitULong
    Next i
    
End Sub

Private Sub ReadSaveChangeRecords(ByRef Progress As ProgressBar)

    Dim i As Long
    Dim ProgressValue As Single

    For i = 0 To SaveFileData.Globals.NumberOfChangeRecords - 1
        SaveFileData.ChangeRecords(i).FormID = GetNext32BitULong
        SaveFileData.ChangeRecords(i).Type = GetNextUByte
        SaveFileData.ChangeRecords(i).Flags = GetNext32BitULong
        SaveFileData.ChangeRecords(i).Version = GetNextUByte
        SaveFileData.ChangeRecords(i).DataSize = GetNext16BitUInteger
        ReDim SaveFileData.ChangeRecords(i).Data(SaveFileData.Globals.NumberOfChangeRecords - 1)
        ReadSaveChangeRecordsData i
        ProgressValue = Int((i / SaveFileData.Globals.NumberOfChangeRecords) * Progress.Max)
        Progress.Value = ProgressValue
        DoEvents
    Next i

End Sub

Private Sub ReadSaveChangeRecordsData(ByVal RecordNumber As Long)

    Dim i As Integer

    For i = 0 To SaveFileData.ChangeRecords(RecordNumber).DataSize - 1
        SaveFileData.ChangeRecords(RecordNumber).Data(i) = GetNextUByte
    Next i

End Sub

Private Sub ReadSaveTempEffects()

    Dim i As Long

    SaveFileData.TempEffects.Size = GetNext32BitULong
    ReDim SaveFileData.TempEffects.Data(SaveFileData.TempEffects.Size - 1)
    For i = 0 To SaveFileData.TempEffects.Size - 1
        SaveFileData.TempEffects.Data(i) = GetNextUByte
    Next i

End Sub

Private Sub ReadSaveFormIDs()

    Dim i As Long

    SaveFileData.FormIDs.NumberOfFormIDs = GetNext32BitULong
    ReDim SaveFileData.FormIDs.FormIDsList(SaveFileData.FormIDs.NumberOfFormIDs - 1)
    For i = 0 To SaveFileData.FormIDs.NumberOfFormIDs - 1
        SaveFileData.FormIDs.FormIDsList(i) = GetNext32BitULong
    Next i

End Sub

Private Sub ReadSaveWorldSpaces()

    Dim i As Long

    SaveFileData.WorldSpaces.NumberOfWorldSpaces = GetNext32BitULong
    ReDim SaveFileData.WorldSpaces.WorldSpaces(SaveFileData.WorldSpaces.NumberOfWorldSpaces - 1)
    For i = 0 To SaveFileData.WorldSpaces.NumberOfWorldSpaces - 1
        SaveFileData.WorldSpaces.WorldSpaces(i) = GetNext32BitULong
    Next i

End Sub

Private Function GetNextFixedLengthString(ByVal Length As Integer) As String

    Dim i As Integer
    Dim NextByte As Byte

    GetNextFixedLengthString = ""
        
    For i = 0 To Length - 1
        Get #FF, , NextByte
        GetNextFixedLengthString = GetNextFixedLengthString & Chr$(NextByte)
    Next i

End Function

Private Function GetNextUByte() As Byte

    Get #FF, , GetNextUByte

End Function

Private Function GetNext16BitUInteger() As Integer
    
    Get #FF, , GetNext16BitUInteger

End Function

Private Function GetNext32BitULong() As Long
    
    Get #FF, , GetNext32BitULong

End Function

Private Function GetNextSystemTime() As SystemTime

    GetNextSystemTime.Year = GetNext16BitUInteger
    GetNextSystemTime.Month = GetNext16BitUInteger
    GetNextSystemTime.DayOfWeek = GetNext16BitUInteger
    GetNextSystemTime.Day = GetNext16BitUInteger
    GetNextSystemTime.Hour = GetNext16BitUInteger
    GetNextSystemTime.Minute = GetNext16BitUInteger
    GetNextSystemTime.Second = GetNext16BitUInteger
    GetNextSystemTime.MilliSecond = GetNext16BitUInteger

End Function

Private Function GetNext32BitSingle() As Single

    Get #FF, , GetNext32BitSingle
        
End Function

Private Function GetNextScreenShot() As ScreenShot

    Dim X As Long
    Dim Y As Long
    
    GetNextScreenShot.Size = GetNext32BitULong       ' Size includes the 8 bits for the next 2 variables
    GetNextScreenShot.Width = GetNext32BitULong
    GetNextScreenShot.Height = GetNext32BitULong

    ' Now that we know how big the screenshot is, assign some room for it
    ReDim GetNextScreenShot.Pixel((GetNextScreenShot.Width * GetNextScreenShot.Height) - 1)

    For Y = 0 To GetNextScreenShot.Height - 1
        For X = 0 To GetNextScreenShot.Width - 1
            GetNextScreenShot.Pixel((Y * GetNextScreenShot.Width) + X).Red = GetNextUByte
            GetNextScreenShot.Pixel((Y * GetNextScreenShot.Width) + X).Green = GetNextUByte
            GetNextScreenShot.Pixel((Y * GetNextScreenShot.Width) + X).Blue = GetNextUByte
        Next X
        DoEvents
    Next Y

End Function

