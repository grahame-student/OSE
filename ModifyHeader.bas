Attribute VB_Name = "ModifyHeader"
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
