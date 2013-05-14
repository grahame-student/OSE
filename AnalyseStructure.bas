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
        InitPlayerFactions
    End If
    
    PopulateFactionListBox MainForm
    
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
    
    If SaveFileData.OSE.PlayerChange.Inventory <> -1 Then
        FixItemReferences
        PopulateItemListBox MainForm
        InitPlayerItems
    End If
    
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
        ' TODO: Needs to account for default items as only changes are recorded!
        SaveFileData.OSE.Player.ItemCount = SaveFileData.OSE.PlayerChange.InventoryCount
        Offset = Offset + InventorySize
    Else
        SaveFileData.OSE.PlayerChange.Inventory = -1
    End If

'    DebugPlayerData

End Sub

Private Function CreatedOrMoved() As Boolean

    CreatedOrMoved = False
    
    If SaveFileData.OSE.PlayerChange.Created <> -1 Then CreatedOrMoved = True
    If SaveFileData.OSE.PlayerChange.Moved <> -1 Then CreatedOrMoved = True
    If SaveFileData.OSE.PlayerChange.HavokMoved <> -1 Then CreatedOrMoved = True

End Function

Private Function InventorySize() As Long

    ' TODO calculate the size of the inventory changes

    ' Calculate the base inventory size
    InventorySize = SaveFileData.OSE.PlayerChange.InventoryCount * 12

    ' Add the changes to the base

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

Private Sub FixItemReferences()

    Dim ItemNumber As Integer

    For ItemNumber = 0 To UBound(ItemData())
        If ItemData(ItemNumber).PlugIn <> "None" Then
            ItemData(ItemNumber).FormID = (ItemData(ItemNumber).FormID Or GetModIndex(ItemData(ItemNumber).PlugIn).Result)
        End If
    Next ItemNumber

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

Public Sub PopulateItemListBox(ByRef MainForm As Form)

    Dim i As Integer

    For i = 0 To UBound(ItemData())
        MainForm.lstAllItems.AddItem ItemData(i).Name, i
        MainForm.lstAllItems.ItemData(i) = ItemData(i).FormID
    Next i

End Sub

Private Sub InitPlayerFactions()

    Dim Offset As New SuperLong
    Dim i As Long
    Dim ByteNumber As Integer

    Offset = SaveFileData.OSE.Player.Factions + 2

    ReDim SaveFileData.OSE.Player.FactionList(SaveFileData.OSE.Player.FactionCount - 1)

    For i = 0 To SaveFileData.OSE.Player.FactionCount - 1
        For ByteNumber = 0 To 3
            SaveFileData.OSE.Player.FactionList(i).iRef.ByteValue(ByteNumber) = _
            SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + ByteNumber)
        Next ByteNumber
        Offset.Add 4
                
        SaveFileData.OSE.Player.FactionList(i).FormID = GetFormID(SaveFileData.OSE.Player.FactionList(i).iRef)
        
        SaveFileData.OSE.Player.FactionList(i).Level = _
        SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset)
        Offset.Add
        
        GetFaction SaveFileData.OSE.Player.FactionList(i).FormID, i
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
    
    MsgBox "FormID not recognised (" & FormID & ")" & vbNewLine & _
           "Please report this FormID to mrloquax@googlemail.com" & vbNewLine & _
           "so that it can be added in future versions", vbOKOnly, "Unknown Faction FormID"
    
End Sub

Private Sub InitPlayerSpells()

    Dim i As Long
    Dim Offset As New SuperLong
    Dim ByteNumber As Integer
    
    Offset = SaveFileData.OSE.Player.Spells + 2
    
    ReDim SaveFileData.OSE.Player.SpellList(SaveFileData.OSE.Player.SpellCount - 1)

    For i = 0 To SaveFileData.OSE.Player.SpellCount - 1
        For ByteNumber = 0 To 3
            SaveFileData.OSE.Player.SpellList(i).iRef.ByteValue(ByteNumber) = _
            SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + ByteNumber)
        Next ByteNumber
        Offset.Add 4
                
        SaveFileData.OSE.Player.SpellList(i).FormID = GetFormID(SaveFileData.OSE.Player.SpellList(i).iRef)
        SaveFileData.OSE.Player.SpellList(i).Name = GetSpell(SaveFileData.OSE.Player.SpellList(i).FormID)
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
           "so that it can be added in future versions", vbOKOnly, "Unknown Spell FormID"
    
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

Private Sub InitPlayerItems()

    Dim i As Long
    Dim j As Integer
    Dim k As Integer
    Dim Offset As New SuperLong
    Dim ByteNumber As Integer
    
    Offset = SaveFileData.OSE.PlayerChange.Inventory + 2
    
    ReDim SaveFileData.OSE.Player.ItemList(SaveFileData.OSE.PlayerChange.InventoryCount - 1)

    For i = 0 To SaveFileData.OSE.PlayerChange.InventoryCount - 1
        For ByteNumber = 0 To 3
            SaveFileData.OSE.Player.ItemList(i).iRef.ByteValue(ByteNumber) = _
            SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
        Next ByteNumber
                        
        SaveFileData.OSE.Player.ItemList(i).FormID = GetFormID(SaveFileData.OSE.Player.ItemList(i).iRef)
        SaveFileData.OSE.Player.ItemList(i).Name = GetItem(SaveFileData.OSE.Player.ItemList(i).FormID)
        Offset.Add 4
        
        For ByteNumber = 0 To 3
            SaveFileData.OSE.Player.ItemList(i).StackedItemsCount.ByteValue(ByteNumber) = _
            SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
        Next ByteNumber
        Offset.Add 4
    
        For ByteNumber = 0 To 3
            SaveFileData.OSE.Player.ItemList(i).ChangedEntriesCount.ByteValue(ByteNumber) = _
            SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
        Next ByteNumber
        Offset.Add 4
    
        If SaveFileData.OSE.Player.ItemList(i).ChangedEntriesCount > 0 Then
            ReDim Preserve SaveFileData.OSE.Player.ItemList(i).InventoryChangedEntries(SaveFileData.OSE.Player.ItemList(i).ChangedEntriesCount - 1)
            For j = 0 To SaveFileData.OSE.Player.ItemList(i).ChangedEntriesCount - 1
                ' Read in the number of properties in the change entry
                For ByteNumber = 0 To 1
                    SaveFileData.OSE.Player.ItemList(i).InventoryChangedEntries(j).PropertiesNumber.ByteValue(ByteNumber) = _
                    SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
                Next ByteNumber
                Offset.Add 2
                
                If SaveFileData.OSE.Player.ItemList(i).InventoryChangedEntries(j).PropertiesNumber > 0 Then
                    ReDim Preserve SaveFileData.OSE.Player.ItemList(i).InventoryChangedEntries(j).Properties(SaveFileData.OSE.Player.ItemList(i).InventoryChangedEntries(j).PropertiesNumber - 1)
                    For k = 0 To SaveFileData.OSE.Player.ItemList(i).InventoryChangedEntries(j).PropertiesNumber - 1
                        AddProperty i, j, k, Offset
                    Next k
                End If
            Next j
        End If
    Next i

    Set Offset = Nothing

End Sub

Public Function GetItem(ByVal FormID As Long) As String

    Dim i As Integer
    
    If FormID < 0 Then
'        GetItem = GetCustomItem(FormID)
        Exit Function
    End If
    
    For i = 0 To UBound(ItemData())
        If FormID = ItemData(i).FormID Then
            GetItem = ItemData(i).Name
            Exit Function
        End If
    Next i
    
    MsgBox "FormID not recognised (" & FormID & ")" & vbNewLine & _
           "Please report this FormID to mrloquax@googlemail.com" & vbNewLine & _
           "so that it can be added in future versions", vbOKOnly, "Unknown Item FormID"
    
End Function

Private Sub AddProperty(ItemNumber As Long, ChangeSetEntry As Integer, PropertyNumber As Integer, Offset As SuperLong)
    
    SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetEntry).Properties(PropertyNumber).Flag = _
    SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset)
    Offset.Add 1

    Select Case SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetEntry).Properties(PropertyNumber).Flag
        Case PROPERTY_FLAG_SCRIPT                       ' d18 h12
            ReadScript ItemNumber, ChangeSetEntry, PropertyNumber, Offset
        Case PROPERTY_FLAG_EQUIPPED_1                   ' d27 h1B
        Case PROPERTY_FLAG_EQUIPPED_2                   ' d28 h1C
        Case PROPERTY_FLAG_ITEM_HEALTH                  ' d43 h2B
            ReadItemHealth ItemNumber, ChangeSetEntry, PropertyNumber, Offset
        Case PROPERTY_FLAG_CURRENT_ENCHANTMENT_POINTS   ' d46 h2E
            ReadCurEP ItemNumber, ChangeSetEntry, PropertyNumber, Offset
        Case PROPERTY_FLAG_SCALE                        ' d55 h37
            ReadScale ItemNumber, ChangeSetEntry, PropertyNumber, Offset
        Case Else
            Debug.Print "Property Flag: " & SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetEntry).Properties(PropertyNumber).Flag
            End
    End Select

End Sub

Private Sub ReadScript(ItemNumber As Long, ChangeSetNumber As Integer, PropertyNumber As Integer, Offset As SuperLong)

    Dim ByteNumber As Integer
    Dim i As Integer
    
    For ByteNumber = 0 To 3
        SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.ScriptRef.ByteValue(ByteNumber) = _
        SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
    Next ByteNumber
    Offset.Add 4

    For ByteNumber = 0 To 1
        SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.VariableCount.ByteValue(ByteNumber) = _
        SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
    Next ByteNumber
    Offset.Add 2

    If SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.VariableCount > 0 Then
        ReDim Preserve SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.VariableList(SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.VariableCount - 1)
        For i = 0 To SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.VariableCount - 1
            For ByteNumber = 0 To 1
                SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.VariableList(i).Index.ByteValue(ByteNumber) = _
                SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
            Next ByteNumber
            Offset.Add 2
                        
            For ByteNumber = 0 To 1
                SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.VariableList(i).Type.ByteValue(ByteNumber) = _
                SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
            Next ByteNumber
            Offset.Add 2
                                    
            Select Case SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.VariableList(i).Type
                Case SCRIPT_VARIABLE_TYPE_IREF
                    For ByteNumber = 0 To 3
                        SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.VariableList(i).RefVariable.ByteValue(ByteNumber) = _
                        SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
                    Next ByteNumber
                    Offset.Add 4
                Case SCRIPT_VARIABLE_TYPE_LOCAL
                    For ByteNumber = 0 To 7
                        SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.VariableList(i).LocalVariable.ByteValue(ByteNumber) = _
                        SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
                    Next ByteNumber
                    Offset.Add 8
            End Select
        Next i
    End If

    SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).Script.Unknown = _
    SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset)
    Offset.Add

End Sub

Private Sub ReadItemHealth(ItemNumber As Long, ChangeSetNumber As Integer, PropertyNumber As Integer, Offset As SuperLong)

    Dim ByteNumber As Integer

    For ByteNumber = 0 To 3
        SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).ItemHealth.ByteValue(ByteNumber) = _
        SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
    Next ByteNumber
    Offset.Add 4

End Sub

Private Sub ReadCurEP(ItemNumber As Long, ChangeSetNumber As Integer, PropertyNumber As Integer, Offset As SuperLong)

    Dim ByteNumber As Integer

    For ByteNumber = 0 To 3
        SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).EnchantmentPoints.ByteValue(ByteNumber) = _
        SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
    Next ByteNumber
    Offset.Add 4

End Sub

Private Sub ReadScale(ItemNumber As Long, ChangeSetNumber As Integer, PropertyNumber As Integer, Offset As SuperLong)

    Dim ByteNumber As Integer

    For ByteNumber = 0 To 3
        SaveFileData.OSE.Player.ItemList(ItemNumber).InventoryChangedEntries(ChangeSetNumber).Properties(PropertyNumber).ScaleValue.ByteValue(ByteNumber) = _
        SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(Offset + ByteNumber)
    Next ByteNumber
    Offset.Add 4

End Sub

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

Private Sub DebugPlayerData()

    Dim FF As Integer
    Dim i As Integer

    FF = FreeFile

    Open App.Path & "\Debug\PlayerData.csv" For Output As #FF
    For i = 0 To UBound(SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data)
        If i Mod 16 = 0 Then Print #FF, ""
        Print #FF, SaveFileData.ChangeRecords(SaveFileData.OSE.PlayerChange.PlayerChangeRecord).Data(i) & ",";
    Next i
    Close #FF

End Sub

