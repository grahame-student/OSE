Attribute VB_Name = "ModifyPlayer"
Option Explicit
DefObj A-Z

' The routines to actually modify the player portion of the data structure
' all live in this file

Public Sub ModifyPlayerAttribute(ByVal PlayerAttribute As Integer, ByVal AttributeValue As Byte)

    SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.BaseAttributes + PlayerAttribute) = AttributeValue

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

