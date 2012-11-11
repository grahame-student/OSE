Attribute VB_Name = "ModifyChangeRecord"
Option Explicit
DefObj A-Z

Public Sub ModifyChangeRecordVersion(ByVal NewValue As Byte)

    Dim i As Long

    For i = 0 To SaveFileData.Globals.NumberOfChangeRecords - 1
        SaveFileData.ChangeRecords(i).Version = NewValue
    Next i

End Sub

