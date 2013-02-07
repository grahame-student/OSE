Attribute VB_Name = "AnalyseStructure"
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

Public Sub ScanForMarkers(ByRef MainForm As Form)

    SaveFileData.OSE.Player.PlayerRecord = LocatePlayerRecord
    SaveFileData.OSE.PlayerChange.PlayerChangeRecord = LocatePlayerChangeRecord
    SaveFileData.OSE.CustomItems.SpellRecords = LocateCustomSpells
    
    If SaveFileData.OSE.CustomItems.SpellRecords(0) <> -1 Then
        ProcessCustomSpells
    End If
    
    If SaveFileData.OSE.Player.PlayerRecord <> -1 Then
        ProcessPlayerRecord MainForm
    End If

    If SaveFileData.OSE.PlayerChange.PlayerChangeRecord <> -1 Then
        ProcessPlayerChangeRecord MainForm
    End If

End Sub

Private Function LocatePlayerRecord() As Long

    Dim i As Long

    For i = 0 To SaveFileData.Globals.NumberOfChangeRecords - 1
        ' Look for the player's change record
        If SaveFileData.ChangeRecords(i).Type = CHANGE_RECORD_NPC_ And _
           SaveFileData.ChangeRecords(i).FormID = PLAYER_FORMID Then
            LocatePlayerRecord = i
            Exit Function
        End If
    Next i

    LocatePlayerRecord = -1

End Function

Private Function LocatePlayerChangeRecord() As Long

    Dim i As Long

    For i = 0 To SaveFileData.Globals.NumberOfChangeRecords - 1
        ' Look for the player's change record
        If SaveFileData.ChangeRecords(i).Type = CHANGE_RECORD_ACHR And _
           SaveFileData.ChangeRecords(i).FormID = PLAYER_CHANGE_FORMID Then
            LocatePlayerChangeRecord = i
            Exit Function
        End If
    Next i

    LocatePlayerChangeRecord = -1

End Function

Private Function LocateCustomSpells() As Long()

    Dim i As Long
    Dim tmpSpellRecords() As Long
    Dim SpellRecordCount As Long
    
    ReDim tmpSpellRecords(0)

    For i = 0 To SaveFileData.Globals.CreatedNumber - 1
        ' Look for spell change records
        If SaveFileData.Globals.CreatedData(i).Type = CREATED_DATA_SPELL Then
            ReDim Preserve tmpSpellRecords(SpellRecordCount)
            tmpSpellRecords(SpellRecordCount) = i
            SpellRecordCount = SpellRecordCount + 1
        End If
    Next i

    If SpellRecordCount = 0 Then
        tmpSpellRecords(0) = -1
    End If

    LocateCustomSpells = tmpSpellRecords

End Function

Private Sub ProcessCustomSpells()

    Dim MaxSpellIndex As Integer
    Dim CurrentRecord As Integer
    Dim i As Integer
    
    MaxSpellIndex = UBound(SaveFileData.OSE.CustomItems.SpellRecords)
    
    ReDim SaveFileData.OSE.CustomItems.Spells(MaxSpellIndex)
    
    For i = 0 To MaxSpellIndex
        CurrentRecord = SaveFileData.OSE.CustomItems.SpellRecords(i)
        
        If ((SaveFileData.Globals.CreatedData(CurrentRecord).FormID And BIT_18) = 0) Then
            ExtractCustomSpellData CurrentRecord, i
        Else
            MsgBox "Created data compressed, cannot process in this version", vbOKOnly, "Not Supported in this version"
        End If
    Next i

End Sub

Private Sub ProcessPlayerRecord(ByRef MainForm As Form)

    ScanForPlayerMarkers
    If SaveFileData.OSE.Player.Factions <> -1 Then
        FixFactionReferences
        PopulateFactionListBox MainForm
        InitPlayerFactions
    End If
    If SaveFileData.OSE.Player.Spells <> -1 Then
        FixSpellReferences
        PopulateSpellListBox MainForm
        InitPlayerSpells
    End If
    If SaveFileData.OSE.Player.BaseModifiers <> -1 Then
        InitPlayerBaseMods
    End If

End Sub

Public Sub ScanForPlayerMarkers()

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
        SaveFileData.OSE.Player.Spells = Offset
        SaveFileData.OSE.Player.SpellCount = ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset) + _
                                               SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1) * BYTE_2))
        Offset = Offset + (SaveFileData.OSE.Player.SpellCount * 4) + 2
    Else
        SaveFileData.OSE.Player.Spells = -1
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
        SaveFileData.OSE.Player.BaseModCount = ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset) + _
                                                 SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1) * BYTE_2))
        Offset = Offset + (SaveFileData.OSE.Player.BaseModCount * 5) + 2
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

Private Sub ProcessPlayerChangeRecord(ByRef MainForm As Form)

    ScanForPlayerChangeMarkers
    
End Sub

Public Sub ScanForPlayerChangeMarkers()
    
    ' Scan the player change record for specific blocks, we need to rescan when the data
    ' structure changes size but it speeds up finding things.

    Dim Offset As Integer
    Dim i As Integer

    ' Check for Cell Changed
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Flags And BIT_31) <> 0) Then
        SaveFileData.OSE.PlayerChange.CellChanged = Offset
        Offset = Offset + 16
    Else
        SaveFileData.OSE.PlayerChange.CellChanged = -1
    End If

    ' Check for Created
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Flags And BIT_1) <> 0) Then
        SaveFileData.OSE.PlayerChange.Created = Offset
        Offset = Offset + 36
    Else
        SaveFileData.OSE.PlayerChange.Created = -1
    End If

    ' Check for Moved
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Flags And BIT_2) <> 0) Then
        SaveFileData.OSE.PlayerChange.Moved = Offset
        Offset = Offset + 28
    Else
        SaveFileData.OSE.PlayerChange.Moved = -1
    End If

    ' Check for HavokMoved
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Flags And BIT_3) <> 0) Then
        SaveFileData.OSE.PlayerChange.HavokMoved = Offset
        Offset = Offset + 28
    Else
        SaveFileData.OSE.PlayerChange.HavokMoved = -1
    End If

    If Not CreatedOrMoved Then
        ' Check for OblivionFlag
        If ((SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Flags And BIT_23) <> 0) Then
            SaveFileData.OSE.PlayerChange.OblivionFlag = Offset
            Offset = Offset + 4
        Else
            SaveFileData.OSE.PlayerChange.OblivionFlag = -1
        End If
    Else
        SaveFileData.OSE.PlayerChange.OblivionFlag = -1
    End If
    
    ' Temporary Attribute Changes
    SaveFileData.OSE.PlayerChange.TempAttributeChanges = Offset
    Offset = Offset + 876
    
    ' Actor Flag
    SaveFileData.OSE.PlayerChange.ActorFlag = Offset
    Offset = Offset + 1

    ' Check for FormFlags
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Flags And BIT_0) <> 0) Then
        SaveFileData.OSE.PlayerChange.FormFlags = Offset
        Offset = Offset + 4
    Else
        SaveFileData.OSE.PlayerChange.FormFlags = -1
    End If

    ' Check for Inventory
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Flags And BIT_27) <> 0) Then
        SaveFileData.OSE.PlayerChange.Inventory = Offset
        SaveFileData.OSE.PlayerChange.InventoryCount = ((SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset) + _
                                                         SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + 1) * BYTE_2))
        Offset = Offset + InventorySize
    Else
        SaveFileData.OSE.PlayerChange.Inventory = -1
    End If



End Sub

Private Function CreatedOrMoved() As Boolean

    CreatedOrMoved = False
    
    If SaveFileData.OSE.PlayerChange.Created <> -1 Then CreatedOrMoved = True
    If SaveFileData.OSE.PlayerChange.Moved <> -1 Then CreatedOrMoved = True
    If SaveFileData.OSE.PlayerChange.HavokMoved <> -1 Then CreatedOrMoved = True

End Function

Private Function InventorySize()

    ' TODO calculate the size of the inventory changes

    InventorySize = SaveFileData.OSE.PlayerChange.InventoryCount * 12

End Function

Private Sub FixFactionReferences()

    Dim FactionNumber As Integer

    For FactionNumber = 0 To UBound(FactionData())
        If FactionData(FactionNumber).PlugIn <> "None" Then
            FactionData(FactionNumber).FormID = (FactionData(FactionNumber).FormID Or GetModIndex(FactionData(FactionNumber).PlugIn).Result)
        End If
    Next FactionNumber

End Sub

Private Sub FixSpellReferences()

    Dim SpellNumber As Integer

    For SpellNumber = 0 To UBound(SpellData())
        If SpellData(SpellNumber).PlugIn <> "None" Then
            SpellData(SpellNumber).FormID = (SpellData(SpellNumber).FormID Or GetModIndex(SpellData(SpellNumber).PlugIn).Result)
        End If
    Next SpellNumber

End Sub

Public Sub PopulateFactionListBox(ByRef MainForm As Form)

    Dim i As Integer

    For i = 0 To UBound(FactionData())
        MainForm.lstAllFactions.AddItem FactionData(i).Name, i
        MainForm.lstAllFactions.ItemData(i) = FactionData(i).FormID
    Next i

End Sub

Public Sub PopulateSpellListBox(ByRef MainForm As Form)

    Dim i As Integer

    For i = 0 To UBound(SpellData())
        MainForm.lstAllSpells.AddItem SpellData(i).Name, i
        MainForm.lstAllSpells.ItemData(i) = SpellData(i).FormID
    Next i

End Sub

Private Sub InitPlayerFactions()

    Dim i As Long
    Dim Offset As Long
    Dim RawiRef As ByteArray
    Dim iRef As LongType
    Dim Level As Byte

    Offset = SaveFileData.OSE.Player.Factions + 2

    ReDim SaveFileData.OSE.Player.FactionList(SaveFileData.OSE.Player.FactionCount - 1)

    For i = 0 To SaveFileData.OSE.Player.FactionCount - 1
        RawiRef.Bytes(0) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset)
        RawiRef.Bytes(1) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1)
        RawiRef.Bytes(2) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 2)
        RawiRef.Bytes(3) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 3)
        
        LSet iRef = RawiRef
        
        Level = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 4)
        SaveFileData.OSE.Player.FactionList(i).FormID = GetFormID(iRef.Result)
        SaveFileData.OSE.Player.FactionList(i).Level = Level
        
        GetFaction SaveFileData.OSE.Player.FactionList(i).FormID, i
        Offset = Offset + 5
    Next i

End Sub

Public Sub GetFaction(ByVal FormID As Long, ByVal IndexNumber As Integer)

    Dim i As Integer
    
    For i = 0 To UBound(FactionData())
        If FormID = FactionData(i).FormID Then
            SaveFileData.OSE.Player.FactionList(IndexNumber).Name = FactionData(i).Name
            SaveFileData.OSE.Player.FactionList(IndexNumber).MaxRank = FactionData(i).MaxRank
            If SaveFileData.OSE.Player.FactionList(IndexNumber).Level = &HFF& Then
                SaveFileData.OSE.Player.FactionList(IndexNumber).Suspended = True
            End If
            SaveFileData.OSE.Player.FactionList(IndexNumber).Ranks() = FactionData(i).Ranks()
            Exit Sub
        End If
    Next i
    
    MsgBox "FormID not recognised (" & FormID & ")", vbOKOnly, "Unknown FormID"
    
End Sub

Private Sub InitPlayerSpells()

    Dim i As Long
    Dim Offset As Long
    Dim RawiRef As ByteArray
    Dim iRef As LongType
    
    Offset = SaveFileData.OSE.Player.Spells + 2
    
    ReDim SaveFileData.OSE.Player.SpellList(SaveFileData.OSE.Player.SpellCount - 1)

    For i = 0 To SaveFileData.OSE.Player.SpellCount - 1
        RawiRef.Bytes(0) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset)
        RawiRef.Bytes(1) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1)
        RawiRef.Bytes(2) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 2)
        RawiRef.Bytes(3) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 3)
        
        LSet iRef = RawiRef
                
        ' Add extra code to handle custom spells here
        SaveFileData.OSE.Player.SpellList(i).iRef = iRef.Result
        SaveFileData.OSE.Player.SpellList(i).FormID = GetFormID(iRef.Result)
        SaveFileData.OSE.Player.SpellList(i).Name = GetSpell(SaveFileData.OSE.Player.SpellList(i).FormID)
            
        Offset = Offset + 4
    Next i

End Sub

Public Function GetSpell(ByVal FormID As Long) As String

    Dim i As Integer
    
    If FormID < 0 Then
        GetSpell = GetCustomSpell(FormID)
        Exit Function
    End If
    
    For i = 0 To UBound(SpellData())
        If FormID = SpellData(i).FormID Then
            GetSpell = SpellData(i).Name
            Exit Function
        End If
    Next i
    
    MsgBox "FormID not recognised (" & FormID & ")" & vbNewLine & _
           "Please report this FormID to mrloquax@googlemail.com" & vbNewLine & _
           "so that it can be added in future versions", vbOKOnly, "Unknown FormID"
    
End Function

Private Function GetCustomSpell(ByVal FormID As Long) As String

    Dim i As Integer

    For i = 0 To UBound(SaveFileData.OSE.CustomItems.SpellRecords())
        If FormID = SaveFileData.Globals.CreatedData(SaveFileData.OSE.CustomItems.SpellRecords(i)).FormID Then
            ' Return custom spell name
            GetCustomSpell = SaveFileData.OSE.CustomItems.Spells(i).Name
            Exit Function
        End If
    Next i

End Function

Private Sub InitPlayerBaseMods()

    ReDim SaveFileData.OSE.Player.BaseModList(SaveFileData.OSE.Player.BaseModCount - 1)

End Sub

Private Sub ExtractCustomSpellData(ByVal CreatedItemRecordIndex As Long, ByVal CustomSpellIndex As Integer)

    Dim SubRecord As Integer
    Dim Offset As Integer
            
    Do Until Offset >= SaveFileData.Globals.CreatedData(CreatedItemRecordIndex).Size
        Select Case GetType(SaveFileData.Globals.CreatedData(CreatedItemRecordIndex).Data, Offset)
            Case SUB_RECORD_FULL_NAME
                SaveFileData.OSE.CustomItems.Spells(CustomSpellIndex).NameLength = GetInteger(SaveFileData.Globals.CreatedData(CreatedItemRecordIndex).Data, Offset)
                SaveFileData.OSE.CustomItems.Spells(CustomSpellIndex).Name = GetZString(SaveFileData.Globals.CreatedData(CreatedItemRecordIndex).Data, Offset)
            Case SUB_RECORD_EDITOR_ID
            Case SUB_RECORD_SPELL_DATA
            Case Else
                Exit Do
        End Select
    Loop

End Sub

