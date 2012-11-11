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

Public Sub ModifySaveFileNumber(ByVal NewValue As Long)

    SaveFileData.SaveHeader.SaveNumber = NewValue

End Sub

