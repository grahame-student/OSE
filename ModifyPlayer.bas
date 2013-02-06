Attribute VB_Name = "ModifyPlayer"
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

' The routines to actually modify the player portion of the data structure
' all live in this file

Public Sub ModifyPlayerAttribute(ByVal PlayerAttribute As Integer, ByVal AttributeValue As Byte)

    SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.BaseAttributes + PlayerAttribute) = AttributeValue

End Sub

Public Sub ModifyPlayerSkill(ByVal PlayerSkill As Integer, ByVal SkillValue As Byte)

    SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.Skills + PlayerSkill) = SkillValue

End Sub

Public Sub ModifyPlayerBaseHealth(ByVal BaseHealthValue As Long)

    Dim tmpHealth(3) As Byte
    Dim i As Integer

    tmpHealth(3) = CByte((BaseHealthValue And &HFF000000) / BYTE_4)
    tmpHealth(2) = CByte((BaseHealthValue And &HFF0000) / BYTE_3)
    tmpHealth(1) = CByte((BaseHealthValue And &HFF00&) / BYTE_2)
    tmpHealth(0) = CByte(BaseHealthValue And &HFF&)

    For i = 0 To 3
        SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.BaseHealth + i) = tmpHealth(i)
    Next i

End Sub

Public Sub ModifyPlayerBaseMagicka(ByVal BaseMagickaValue As Long)

    Dim tmpMagicka(1) As Byte
    Dim i As Integer

    tmpMagicka(1) = CByte((BaseMagickaValue And &HFF00&) / BYTE_2)
    tmpMagicka(0) = CByte(BaseMagickaValue And &HFF&)

    For i = 4 To 5
        SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.BaseData + i) = tmpMagicka(i - 4)
    Next i

End Sub

Public Sub ModifyPlayerBaseFatigue(ByVal BaseFatigueValue As Long)

    Dim tmpFatigue(1) As Byte
    Dim i As Integer

    tmpFatigue(1) = CByte((BaseFatigueValue And &HFF00&) / BYTE_2)
    tmpFatigue(0) = CByte(BaseFatigueValue And &HFF&)

    For i = 6 To 7
        SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.BaseData + i) = tmpFatigue(i - 6)
    Next i

End Sub

Public Sub ModifyPlayerFaction(ByVal FactionNumber As Byte, ByVal FactionRank As Byte)

    ' Can only modify existing factions, other routines will allow adding and removing factions

    Dim Offset As Integer

    Offset = (FactionNumber * 5) + 6

    ' If a strange rank was detected when the file was loaded then restore it when saving
    If SaveFileData.OSE.Player.FactionList(FactionNumber).Suspended Then
        SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.Factions + Offset) = &HFF&
    Else
        SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.Factions + Offset) = FactionRank
    End If

End Sub

Public Sub RebuildPlayerChangeRecord()

    Dim tmpPlayerChangeRecord As ChangeRecord
    Dim Position As Integer
    Dim Start As Integer
        
    tmpPlayerChangeRecord.FormID = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).FormID
    tmpPlayerChangeRecord.Type = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Type
    tmpPlayerChangeRecord.Flags = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags
    tmpPlayerChangeRecord.Version = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Version
    
    tmpPlayerChangeRecord.DataSize = GetNewPlayerRecordSize
    ReDim tmpPlayerChangeRecord.Data(tmpPlayerChangeRecord.DataSize - 1)
    
    Position = 0
    
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_0) <> 0) Then
        SetPlayerChangeRecordFormFlags tmpPlayerChangeRecord, Position, 4
        Position = Position + 4
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_3) <> 0) Then
        SetPlayerChangeRecordBaseAttributes tmpPlayerChangeRecord, Position, 8
        Position = Position + 8
    End If
    
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_4) <> 0) Then
        SetPlayerChangeRecordBaseData tmpPlayerChangeRecord, Position, 16
        Position = Position + 16
    End If
    
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_6) <> 0) Then
        SetPlayerChangeRecordFactions tmpPlayerChangeRecord, Position
        Position = Position + (SaveFileData.OSE.Player.FactionCount * 5) + 2
    End If
    
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_5) <> 0) Then
        SetPlayerChangeRecordSpellList tmpPlayerChangeRecord, Position
        Position = Position + (SaveFileData.OSE.Player.SpellCount * 4) + 2
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_8) <> 0) Then
        SetPlayerChangeRecordAI tmpPlayerChangeRecord, Position, 4
        Position = Position + 4
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_2) <> 0) Then
        SetPlayerChangeRecordBaseHealth tmpPlayerChangeRecord, Position, 4
        Position = Position + 4
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_28) <> 0) Then
        SetPlayerChangeRecordBaseMods tmpPlayerChangeRecord, Position
        Position = Position + (SaveFileData.OSE.Player.BaseModCount * 5) + 2
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_7) <> 0) Then
        SetPlayerChangeRecordFullName tmpPlayerChangeRecord
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_9) <> 0) Then
        SetPlayerChangeRecordSkills tmpPlayerChangeRecord, Position, 21
        Position = Position + 21
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_10) <> 0) Then
        SetPlayerChangeRecordCombatStyle tmpPlayerChangeRecord, Position, 21
        Position = Position + 4
    End If

    SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord) = tmpPlayerChangeRecord

    ScanForPlayerMarkers

End Sub

Private Function GetNewPlayerRecordSize() As Integer

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_0) <> 0) Then
        GetNewPlayerRecordSize = GetNewPlayerRecordSize + 4
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_3) <> 0) Then
        GetNewPlayerRecordSize = GetNewPlayerRecordSize + 8
    End If
    
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_4) <> 0) Then
        GetNewPlayerRecordSize = GetNewPlayerRecordSize + 16
    End If
    
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_6) <> 0) Then
        GetNewPlayerRecordSize = GetNewPlayerRecordSize + (SaveFileData.OSE.Player.FactionCount * 5) + 2
    End If
    
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_5) <> 0) Then
        GetNewPlayerRecordSize = GetNewPlayerRecordSize + (SaveFileData.OSE.Player.SpellCount * 4) + 2
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_8) <> 0) Then
        GetNewPlayerRecordSize = GetNewPlayerRecordSize + 4
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_2) <> 0) Then
        GetNewPlayerRecordSize = GetNewPlayerRecordSize + 4
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_28) <> 0) Then
        GetNewPlayerRecordSize = GetNewPlayerRecordSize + (SaveFileData.OSE.Player.BaseModCount * 5) + 2
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_7) <> 0) Then
        GetNewPlayerRecordSize = GetNewPlayerRecordSize + 1 + Len(SaveFileData.OSE.Player.FullNameString)
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_9) <> 0) Then
        GetNewPlayerRecordSize = GetNewPlayerRecordSize + 21
    End If

    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_10) <> 0) Then
        GetNewPlayerRecordSize = GetNewPlayerRecordSize + 4
    End If

End Function

Private Sub SetPlayerChangeRecordFormFlags(ByRef tmpPlayerChangeRecord As ChangeRecord, ByVal Start As Integer, ByVal Length As Integer)

    Dim NewOffset As Integer
    Dim OriginalOffset As Integer
    
    OriginalOffset = 0
    
    For NewOffset = Start To Start + (Length - 1)
        tmpPlayerChangeRecord.Data(NewOffset) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.FormFlags + OriginalOffset)
        OriginalOffset = OriginalOffset + 1
    Next NewOffset

End Sub

Private Sub SetPlayerChangeRecordBaseAttributes(ByRef tmpPlayerChangeRecord As ChangeRecord, ByVal Start As Integer, ByVal Length As Integer)

    Dim NewOffset As Integer
    Dim OriginalOffset As Integer
    
    OriginalOffset = 0
    
    For NewOffset = Start To Start + (Length - 1)
        tmpPlayerChangeRecord.Data(NewOffset) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.BaseAttributes + OriginalOffset)
        OriginalOffset = OriginalOffset + 1
    Next NewOffset

End Sub

Private Sub SetPlayerChangeRecordBaseData(ByRef tmpPlayerChangeRecord As ChangeRecord, ByVal Start As Integer, ByVal Length As Integer)

    Dim NewOffset As Integer
    Dim OriginalOffset As Integer
    
    OriginalOffset = 0
    
    For NewOffset = Start To Start + (Length - 1)
        tmpPlayerChangeRecord.Data(NewOffset) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.BaseData + OriginalOffset)
        OriginalOffset = OriginalOffset + 1
    Next NewOffset

End Sub

Private Sub SetPlayerChangeRecordFactions(ByRef tmpPlayerChangeRecord As ChangeRecord, ByVal Start As Integer)

    Dim NewOffset As Integer
    Dim RecordNumber As Integer
    Dim iRef As Long

    NewOffset = Start
    tmpPlayerChangeRecord.Data(NewOffset) = CByte(SaveFileData.OSE.Player.FactionCount And &HFF&)
    tmpPlayerChangeRecord.Data(NewOffset + 1) = CByte((SaveFileData.OSE.Player.FactionCount And &HFF00&) / BYTE_2)
    NewOffset = NewOffset + 2
    
    For RecordNumber = 0 To SaveFileData.OSE.Player.FactionCount - 1
        iRef = GetIref(SaveFileData.OSE.Player.FactionList(RecordNumber).FormID)
        tmpPlayerChangeRecord.Data(NewOffset) = CByte(iRef And &HFF&)
        tmpPlayerChangeRecord.Data(NewOffset + 1) = CByte((iRef And &HFF00&) / BYTE_2)
        tmpPlayerChangeRecord.Data(NewOffset + 2) = CByte((iRef And &HFF0000) / BYTE_3)
        tmpPlayerChangeRecord.Data(NewOffset + 3) = CByte((iRef And &HFF000000) / BYTE_4)
        tmpPlayerChangeRecord.Data(NewOffset + 4) = CByte(SaveFileData.OSE.Player.FactionList(RecordNumber).Level)
        NewOffset = NewOffset + 5
    Next RecordNumber

End Sub

Private Sub SetPlayerChangeRecordSpellList(ByRef tmpPlayerChangeRecord As ChangeRecord, ByVal Start As Integer)

    Dim NewOffset As Integer
    Dim RecordNumber As Integer
    Dim RawiRef As ByteArray
    Dim iRef As LongType

    NewOffset = Start
    Debug.Print UBound(tmpPlayerChangeRecord.Data())
    tmpPlayerChangeRecord.Data(NewOffset) = CByte(SaveFileData.OSE.Player.SpellCount And &HFF&)
    tmpPlayerChangeRecord.Data(NewOffset + 1) = CByte((SaveFileData.OSE.Player.SpellCount And &HFF00&) / BYTE_2)
    NewOffset = NewOffset + 2
        
    For RecordNumber = 0 To SaveFileData.OSE.Player.SpellCount - 1
        iRef.Result = SaveFileData.OSE.Player.SpellList(RecordNumber).iRef
        
        LSet RawiRef = iRef
        
        tmpPlayerChangeRecord.Data(NewOffset) = RawiRef.Bytes(0)
        tmpPlayerChangeRecord.Data(NewOffset + 1) = RawiRef.Bytes(1)
        tmpPlayerChangeRecord.Data(NewOffset + 2) = RawiRef.Bytes(2)
        tmpPlayerChangeRecord.Data(NewOffset + 3) = RawiRef.Bytes(3)
        NewOffset = NewOffset + 4
    Next RecordNumber

End Sub

Private Sub SetPlayerChangeRecordAI(ByRef tmpPlayerChangeRecord As ChangeRecord, ByVal Start As Integer, ByVal Length As Integer)

    Dim NewOffset As Integer
    Dim OriginalOffset As Integer
    
    OriginalOffset = 0
    
    For NewOffset = Start To Start + (Length - 1)
        tmpPlayerChangeRecord.Data(NewOffset) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.AI + OriginalOffset)
        OriginalOffset = OriginalOffset + 1
    Next NewOffset

End Sub

Private Sub SetPlayerChangeRecordBaseHealth(ByRef tmpPlayerChangeRecord As ChangeRecord, ByVal Start As Integer, ByVal Length As Integer)

    Dim NewOffset As Integer
    Dim OriginalOffset As Integer
    
    OriginalOffset = 0
    
    For NewOffset = Start To Start + (Length - 1)
        tmpPlayerChangeRecord.Data(NewOffset) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.BaseHealth + OriginalOffset)
        OriginalOffset = OriginalOffset + 1
    Next NewOffset

End Sub

Private Sub SetPlayerChangeRecordBaseMods(ByRef tmpPlayerChangeRecord As ChangeRecord, ByVal Start As Integer)

    Dim NewOffset As Integer
    Dim RecordNumber As Integer

    NewOffset = Start
    tmpPlayerChangeRecord.Data(NewOffset) = CByte(SaveFileData.OSE.Player.BaseModCount And &HFF&)
    tmpPlayerChangeRecord.Data(NewOffset + 1) = CByte((SaveFileData.OSE.Player.BaseModCount And &HFF00&) / BYTE_2)
    NewOffset = NewOffset + 2
    
    For RecordNumber = 0 To SaveFileData.OSE.Player.BaseModCount - 1
        tmpPlayerChangeRecord.Data(NewOffset) = CByte(SaveFileData.OSE.Player.BaseModList(RecordNumber).Index)
        tmpPlayerChangeRecord.Data(NewOffset + 1) = CByte(SaveFileData.OSE.Player.BaseModList(RecordNumber).ModValue And &HFF&)
        tmpPlayerChangeRecord.Data(NewOffset + 2) = CByte(SaveFileData.OSE.Player.BaseModList(RecordNumber).ModValue And &HFF00&) / BYTE_2
        tmpPlayerChangeRecord.Data(NewOffset + 3) = CByte(SaveFileData.OSE.Player.BaseModList(RecordNumber).ModValue And &HFF0000) / BYTE_3
        tmpPlayerChangeRecord.Data(NewOffset + 4) = CByte(SaveFileData.OSE.Player.BaseModList(RecordNumber).ModValue And &HFF000000) / BYTE_4
        NewOffset = NewOffset + 5
    Next RecordNumber

End Sub

Private Sub SetPlayerChangeRecordFullName(ByRef tmpPlayerChangeRecord As ChangeRecord)

    Dim NewOffset As Integer
    Dim NamePosition As Integer
        
    tmpPlayerChangeRecord.Data(NewOffset) = CByte(Len(SaveFileData.OSE.Player.FullNameString))
        
    For NamePosition = 1 To Len(SaveFileData.OSE.Player.FullNameString)
        tmpPlayerChangeRecord.Data(NewOffset + NamePosition) = Asc(Mid(SaveFileData.OSE.Player.FullNameString, NamePosition, 1))
    Next NamePosition

End Sub

Private Sub SetPlayerChangeRecordSkills(ByRef tmpPlayerChangeRecord As ChangeRecord, ByVal Start As Integer, ByVal Length As Integer)

    Dim NewOffset As Integer
    Dim OriginalOffset As Integer
    
    OriginalOffset = 0
    
    For NewOffset = Start To Start + (Length - 1)
        tmpPlayerChangeRecord.Data(NewOffset) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.Skills + OriginalOffset)
        OriginalOffset = OriginalOffset + 1
    Next NewOffset

End Sub

Private Sub SetPlayerChangeRecordCombatStyle(ByRef tmpPlayerChangeRecord As ChangeRecord, ByVal Start As Integer, ByVal Length As Integer)

    Dim NewOffset As Integer
    Dim OriginalOffset As Integer
    
    OriginalOffset = 0
    
    For NewOffset = Start To Start + (Length - 1)
        tmpPlayerChangeRecord.Data(NewOffset) = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.CombatStyle + OriginalOffset)
        OriginalOffset = OriginalOffset + 1
    Next NewOffset

End Sub

