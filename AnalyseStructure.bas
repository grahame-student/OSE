Attribute VB_Name = "AnalyseStructure"
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

Public Sub ScanForMarkers()

    SaveFileData.OSE.Player.PlayerRecord = LocatePlayerRecord
    If SaveFileData.OSE.Player.PlayerRecord <> -1 Then
        ScanForPlayerMarkers
        If SaveFileData.OSE.Player.Factions <> -1 Then
            InitPlayerFactions
        End If
    End If

End Sub

Private Function LocatePlayerRecord() As Long

    Dim i As Long

    For i = 0 To SaveFileData.Globals.NumberOfChangeRecords - 1
        ' Look for the player's change record
        If SaveFileData.ChangeRecords(i).Type = 35 And SaveFileData.ChangeRecords(i).FormID = 7 Then
            LocatePlayerRecord = i
            Exit Function
        End If
    Next i

    LocatePlayerRecord = -1

End Function

Private Sub ScanForPlayerMarkers()

    ' Scan the player record for specific blocks, we need to rescan when the data
    ' structure changes size but it speeds up finding things.

    Dim Offset As Integer
    Dim i As Integer

    ' Check for Form Flags
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_0) <> 0) Then
        SaveFileData.OSE.Player.FormFlags = Offset
        Offset = Offset + 4
    Else
        SaveFileData.OSE.Player.FormFlags = -1
    End If
    
    ' Check for Base Attributes
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_3) <> 0) Then
        SaveFileData.OSE.Player.BaseAttributes = Offset
        Offset = Offset + 8
    Else
        SaveFileData.OSE.Player.BaseAttributes = -1
    End If
    
    ' Check for Base Data
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_4) <> 0) Then
        SaveFileData.OSE.Player.BaseData = Offset
        Offset = Offset + 16
    Else
        SaveFileData.OSE.Player.BaseData = -1
    End If
    
    ' Check for Factions
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_6) <> 0) Then
        SaveFileData.OSE.Player.Factions = Offset
        SaveFileData.OSE.Player.FactionCount = ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset) + _
                                                 SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1) * BYTE_2))
        Offset = Offset + (SaveFileData.OSE.Player.FactionCount * 5) + 2
    Else
        SaveFileData.OSE.Player.Factions = -1
    End If
    
    ' Check for spell list
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_5) <> 0) Then
        SaveFileData.OSE.Player.SpellList = Offset
        Offset = Offset + _
                 ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset) + _
                  SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1) * BYTE_2)) * 4
        Offset = Offset + 2
    Else
        SaveFileData.OSE.Player.SpellList = -1
    End If
    
    ' Check for AI Data
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_8) <> 0) Then
        SaveFileData.OSE.Player.AI = Offset
        Offset = Offset + 4
    Else
        SaveFileData.OSE.Player.AI = -1
    End If
    
    ' Check for base health
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_2) <> 0) Then
        SaveFileData.OSE.Player.BaseHealth = Offset
        Offset = Offset + 4
    Else
        SaveFileData.OSE.Player.BaseHealth = -1
    End If
    
    ' Check for base modifiers
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_28) <> 0) Then
        SaveFileData.OSE.Player.BaseModifiers = Offset
        Offset = Offset + _
                 ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset) + _
                  SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1) * BYTE_2)) * 5
        Offset = Offset + 2
    Else
        SaveFileData.OSE.Player.BaseModifiers = -1
    End If
    
    ' Check for full name
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_7) <> 0) Then
        ' Not used for player will need to be fixed for records that do require this sub-record
        SaveFileData.OSE.Player.FullName = Offset
        Offset = Offset ' + length of name
    Else
        SaveFileData.OSE.Player.FullName = -1
    End If
    
    ' Check for skills
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_9) <> 0) Then
        SaveFileData.OSE.Player.Skills = Offset
        Offset = Offset + 21
    Else
        SaveFileData.OSE.Player.Skills = -1
    End If
    
    ' Check for combat style
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_10) <> 0) Then
        SaveFileData.OSE.Player.CombatStyle = Offset
    Else
        SaveFileData.OSE.Player.CombatStyle = -1
    End If

End Sub

Private Sub InitPlayerFactions()

    Dim i As Long
    Dim Offset As Long
    Dim RawRef As ByteArray
    Dim Ref As LongType
    Dim Level As Byte

    Offset = SaveFileData.OSE.Player.Factions + 2

    ReDim SaveFileData.OSE.Player.FactionList(SaveFileData.OSE.Player.FactionCount - 1)

    For i = 0 To SaveFileData.OSE.Player.FactionCount - 1
        RawRef.Bytes(0) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset)
        RawRef.Bytes(1) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1)
        RawRef.Bytes(2) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 2)
        RawRef.Bytes(3) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 3)
        
        LSet Ref = RawRef
        
        Level = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 4)
        SaveFileData.OSE.Player.FactionList(i).Ref = GetFormID(Ref.Result)
        SaveFileData.OSE.Player.FactionList(i).Level = Level
        GetFaction SaveFileData.OSE.Player.FactionList(i).Ref, i
        Offset = Offset + 5
    Next i

End Sub

Private Sub GetFaction(ByVal Reference As Long, ByVal IndexNumber As Integer)

    Dim i As Integer
    
    If Int(Reference / BYTE_4) > 0 Then
        GetPluginFaction Reference, IndexNumber
        Exit Sub
    End If
    
    For i = 0 To UBound(FactionData())
        If Reference = FactionData(i).Reference Then
            SaveFileData.OSE.Player.FactionList(IndexNumber).Name = FactionData(i).Name
            Exit Sub
        End If
    Next i
    
    MsgBox "Reference not recognised (" & Reference & ")", vbOKOnly, "Unknown Reference"
    

End Sub

Private Sub GetPluginFaction(ByVal Reference As Long, ByVal IndexNumber As Integer)

    Dim i As Integer
    Dim ModReference As Long

    For i = 0 To UBound(FactionData())
        If FactionData(i).PlugIn <> "None" Then
            ModReference = (FactionData(i).Reference Or GetModIndex(FactionData(i).PlugIn).Result)
            If Reference = ModReference Then
                SaveFileData.OSE.Player.FactionList(IndexNumber).Name = FactionData(i).Name
                Exit Sub
            End If
        End If
    Next i

    MsgBox "Reference not recognised (" & Reference & ")", vbOKOnly, "Unknown Reference"

End Sub
