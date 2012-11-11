VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OSE"
   ClientHeight    =   6015
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   7815
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pnlSaveFile 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4545
      ScaleWidth      =   7305
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin VB.PictureBox pnlAll 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   1320
         ScaleHeight     =   4305
         ScaleWidth      =   5865
         TabIndex        =   39
         Top             =   120
         Width           =   5895
         Begin VB.PictureBox picScreenShot 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2250
            Left            =   1935
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   150
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   256
            TabIndex        =   50
            ToolTipText     =   $"frmMain.frx":1CCA
            Top             =   1920
            Width           =   3840
         End
         Begin VB.ListBox lstPlugIns 
            Height          =   1035
            Left            =   1935
            TabIndex        =   48
            ToolTipText     =   "The plugins in use when the save was created."
            Top             =   840
            Width           =   3840
         End
         Begin VB.TextBox txtSaveFileNumber 
            Height          =   285
            Left            =   1935
            TabIndex        =   40
            Text            =   "0"
            ToolTipText     =   "The number of saves made since starting the character. (0 to 2147483647)"
            Top             =   315
            Width           =   3840
         End
         Begin VB.TextBox txtSaveFileVersionMajor 
            Height          =   285
            Left            =   1935
            TabIndex        =   42
            Text            =   "0"
            ToolTipText     =   "The major version of the savefile format. (0 to 255)"
            Top             =   75
            Width           =   735
         End
         Begin VB.TextBox txtSaveFileVersionMinor 
            Height          =   285
            Left            =   5040
            TabIndex        =   41
            Text            =   "0"
            ToolTipText     =   "The major version of the savefile format. (0 to 255)"
            Top             =   75
            Width           =   735
         End
         Begin VB.Label lblScreenshot 
            AutoSize        =   -1  'True
            Caption         =   "Screenshot"
            Height          =   195
            Left            =   960
            TabIndex        =   51
            Top             =   1920
            Width           =   810
         End
         Begin VB.Label lblPlugins 
            AutoSize        =   -1  'True
            Caption         =   "PlugIns"
            Height          =   195
            Left            =   1245
            TabIndex        =   49
            Top             =   960
            Width           =   525
         End
         Begin VB.Label lblSaveTime 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   1935
            TabIndex        =   46
            ToolTipText     =   "The time that the save file was created. Click to enter a new time manually."
            Top             =   600
            Width           =   3840
         End
         Begin VB.Label lblSavetimeTag 
            AutoSize        =   -1  'True
            Caption         =   "Save Time"
            Height          =   195
            Left            =   1005
            TabIndex        =   47
            Top             =   630
            Width           =   765
         End
         Begin VB.Label lblSaveFileNumberTag 
            AutoSize        =   -1  'True
            Caption         =   "Save File Number"
            Height          =   195
            Left            =   510
            TabIndex        =   45
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label lblSaveFileVersionMajor 
            AutoSize        =   -1  'True
            Caption         =   "Save File Major Version"
            Height          =   195
            Left            =   105
            TabIndex        =   44
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label lblSaveFileVersionMinor 
            AutoSize        =   -1  'True
            Caption         =   "Save File Minor Version"
            Height          =   195
            Left            =   3255
            TabIndex        =   43
            Top             =   120
            Width           =   1665
         End
      End
      Begin MSComctlLib.TabStrip tabSaveSubCategory 
         Height          =   4335
         Left            =   120
         TabIndex        =   5
         Top             =   120
         Width           =   1030
         _ExtentX        =   1826
         _ExtentY        =   7646
         TabWidthStyle   =   2
         MultiRow        =   -1  'True
         Style           =   1
         TabFixedWidth   =   1764
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "All"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   6120
      TabIndex        =   6
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox pnlPlayer 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4545
      ScaleWidth      =   7305
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   7335
      Begin VB.PictureBox pnlBaseStats 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   1320
         ScaleHeight     =   945
         ScaleWidth      =   2865
         TabIndex        =   31
         Top             =   1320
         Width           =   2895
         Begin VB.TextBox txtBaseHealth 
            Height          =   285
            Left            =   1080
            TabIndex        =   32
            Text            =   "0"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtBaseFatigue 
            Height          =   285
            Left            =   1080
            TabIndex        =   33
            Text            =   "0"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtBaseMagicka 
            Height          =   285
            Left            =   1080
            TabIndex        =   34
            Text            =   "0"
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblBaseMagicka 
            Caption         =   "B. Magicka"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   135
            Width           =   855
         End
         Begin VB.Label lblBaseFatigue 
            Caption         =   "B. Fatigue"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   375
            Width           =   855
         End
         Begin VB.Label lblBaseHealth 
            Caption         =   "B. Health"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   615
            Width           =   855
         End
      End
      Begin VB.PictureBox pnlBasicInformation 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   1320
         ScaleHeight     =   1185
         ScaleWidth      =   2865
         TabIndex        =   24
         Top             =   120
         Width           =   2895
         Begin VB.TextBox txtLocation 
            Height          =   525
            Left            =   1080
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   25
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtLevel 
            Height          =   285
            Left            =   1080
            TabIndex        =   26
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1080
            MaxLength       =   254
            TabIndex        =   27
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblName 
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   135
            Width           =   855
         End
         Begin VB.Label lblLevel 
            Caption         =   "Level"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   375
            Width           =   855
         End
         Begin VB.Label lblLocation 
            Caption         =   "Location"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   615
            Width           =   855
         End
      End
      Begin VB.PictureBox pnlBaseAttributes 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   1320
         ScaleHeight     =   2145
         ScaleWidth      =   2865
         TabIndex        =   7
         Top             =   2280
         Width           =   2895
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   8
            Text            =   "0"
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   9
            Text            =   "0"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   10
            Text            =   "0"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   11
            Text            =   "0"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   12
            Text            =   "0"
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   13
            Text            =   "0"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   14
            Text            =   "0"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   15
            Text            =   "0"
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblStrength 
            Caption         =   "Strength"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   135
            Width           =   855
         End
         Begin VB.Label lblIntelligence 
            Caption         =   "Intelligence"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   375
            Width           =   855
         End
         Begin VB.Label lblAgility 
            Caption         =   "Agility"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   855
            Width           =   855
         End
         Begin VB.Label lblWillPower 
            Caption         =   "Willpower"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   615
            Width           =   855
         End
         Begin VB.Label lblEndurance 
            Caption         =   "Endurance"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1335
            Width           =   855
         End
         Begin VB.Label lblSpeed 
            Caption         =   "Speed"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label lblLuck 
            Caption         =   "Luck"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   1815
            Width           =   855
         End
         Begin VB.Label lblPersonality 
            Caption         =   "Personality"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1575
            Width           =   855
         End
      End
      Begin MSComctlLib.TabStrip tabPlayerSubCategory 
         Height          =   4335
         Left            =   120
         TabIndex        =   38
         Top             =   120
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   7646
         TabWidthStyle   =   2
         MultiRow        =   -1  'True
         Style           =   1
         TabFixedWidth   =   1764
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   1
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Attributes"
               ImageVarType    =   2
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.TabStrip tabCategory 
      Height          =   5055
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8916
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   2
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Save File"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Player"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ProgressBar prgProgress 
      Height          =   255
      Left            =   9120
      TabIndex        =   2
      Top             =   6720
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
      Max             =   10000
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6853
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6853
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuQuit 
         Caption         =   "Quit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

Private Sub Form_Load()

    frmMain.Caption = PROGRAM_TITLE

    SetUpProgressBarInStatusBar

End Sub

Private Sub SetUpProgressBarInStatusBar()
    
    ' Make prgProgress a child of StatusBar
    SetParent prgProgress.hwnd, StatusBar.hwnd
    
    PositionProgressBarWithInStatusBar

End Sub

Private Sub PositionProgressBarWithInStatusBar()

    Dim BorderX As Long
    Dim BorderY As Long

    Dim ProgressBarLeft As Long
    Dim ProgressBarTop As Long
    Dim ProgressBarWidth As Long
    Dim ProgressBarHeight As Long

    ' Get border sizes based on system metrics, converted to twips
    BorderX = GetSystemMetrics(SM_CXBORDER) * Screen.TwipsPerPixelX
    BorderY = GetSystemMetrics(SM_CYBORDER) * Screen.TwipsPerPixelY

    ' Make the progressbar look like it belongs in the statusbar
    With StatusBar
        ProgressBarLeft = .Panels(STB_PROGRESS).Left + BorderX
        ProgressBarTop = 3 * BorderY
        ProgressBarWidth = .Panels(STB_PROGRESS).Width - (2 * BorderX)
        ProgressBarHeight = .Height - (4 * BorderY)
        
        prgProgress.Move ProgressBarLeft, ProgressBarTop, ProgressBarWidth, ProgressBarHeight
    End With

End Sub

Private Sub Form_Unload(Cancel As Integer)

    mnuQuit_Click

End Sub

Private Sub lblSaveTime_Click()

    frmNewDateStamp.Mode = DATE_STAMP_MODE_GAME
    frmNewDateStamp.Show vbModal
    Unload frmNewDateStamp

    UpdateDisplay
    
End Sub

Private Sub mnuQuit_Click()

    End

End Sub

Private Sub mnuAbout_Click()

    frmAbout.Show

End Sub

Private Sub mnuOpen_Click()

    ClearSaveFileData

    OpenSaveFile

End Sub

Private Sub ClearSaveFileData()

    Dim EmptyESS As ESS

    SaveFileData = EmptyESS

End Sub

Private Sub OpenSaveFile()

    cmdSave.Enabled = False
    
    ReadSave.ReadSaveFile StatusBar, prgProgress
    
    If Not SaveFileData.OSE.LoadSuccessful Then
        Exit Sub
    End If
    
    StatusBar.Panels(1).Text = "Scanning for markers"
    AnalyseStructure.ScanForMarkers

    StatusBar.Panels(1).Text = "Read complete"

    tabCategory.SelectedItem = TAB_CAT_SAVE_FILE
    tabCategory_Click

    cmdSave.Enabled = True

End Sub

Private Sub picScreenShot_Click()

    SavePicture picScreenShot.Image, App.Path & "\ScreenShot\screenshot.bmp"

End Sub

Private Sub picScreenShot_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' TODO: Catch invalid file formats
    ' TODO: Save new bitmap to the data structure

    picScreenShot.Picture = LoadPicture(Data.Files(1))

End Sub

Private Sub tabCategory_Click()

    If Not SaveFileData.OSE.LoadSuccessful Then Exit Sub

    HideAllPanels

    Select Case tabCategory.SelectedItem
        Case TAB_CAT_SAVE_FILE
            pnlSaveFile.Visible = True
        Case TAB_CAT_PLAYER
            pnlPlayer.Visible = True
    End Select

    UpdateDisplay

End Sub

Private Sub HideAllPanels()

    pnlSaveFile.Visible = False
    pnlPlayer.Visible = False

End Sub

Private Sub UpdateDisplay()
        
    Select Case tabCategory.SelectedItem
        Case TAB_CAT_SAVE_FILE
            UpdateDisplaySaveFile
        Case TAB_CAT_PLAYER
            UpdateDisplayPlayer
    End Select
    
    txtName.Text = SaveFileData.SaveHeader.PlayerName
    txtLevel.Text = SaveFileData.SaveHeader.PlayerLevel
    txtLocation.Text = SaveFileData.SaveHeader.PlayerLocation
            
End Sub

Private Sub UpdateDisplaySaveFile()

    Select Case tabSaveSubCategory.SelectedItem
        Case TAB_SUB_CAT_SAVE_FILE_ALL
            UpdateDisplaySaveFileAll
    End Select

End Sub

Private Sub UpdateDisplaySaveFileAll()

    Dim i As Integer

    txtSaveFileVersionMajor.Text = SaveFileData.FileHeader.MajorVersion
    txtSaveFileVersionMinor.Text = SaveFileData.FileHeader.MinorVersion
    txtSaveFileNumber.Text = SaveFileData.SaveHeader.SaveNumber
    
    lblSaveTime.Caption = SaveFileData.SaveHeader.GameTime.Hour & ":" & _
                          SaveFileData.SaveHeader.GameTime.Minute & ":" & _
                          SaveFileData.SaveHeader.GameTime.Second & " " & _
                          WeekdayName(SaveFileData.SaveHeader.GameTime.DayOfWeek) & " " & _
                          Ordinal(SaveFileData.SaveHeader.GameTime.Day) & ", " & _
                          MonthName(SaveFileData.SaveHeader.GameTime.Month) & " " & _
                          SaveFileData.SaveHeader.GameTime.Year
    
    lstPlugIns.Clear
    For i = 0 To SaveFileData.PlugIns.NumberOfPlugins - 1
        lstPlugIns.AddItem SaveFileData.PlugIns.PlugInNames(i)
    Next i

    If Not SaveFileData.OSE.ScreenShotLoaded Then
        DisplayScreenShot
    End If

End Sub

Private Sub DisplayScreenShot()

    Dim X As Long
    Dim Y As Long

    ' Chuck the screenshot into a picturebox
    For Y = 0 To SaveFileData.SaveHeader.ScreenShot.Height - 1
        For X = 0 To SaveFileData.SaveHeader.ScreenShot.Width - 1
            picScreenShot.PSet (X, Y), RGB(SaveFileData.SaveHeader.ScreenShot.Pixel((Y * SaveFileData.SaveHeader.ScreenShot.Width) + X).Red, _
                                           SaveFileData.SaveHeader.ScreenShot.Pixel((Y * SaveFileData.SaveHeader.ScreenShot.Width) + X).Green, _
                                           SaveFileData.SaveHeader.ScreenShot.Pixel((Y * SaveFileData.SaveHeader.ScreenShot.Width) + X).Blue)
        Next X
    Next Y

    SaveFileData.OSE.ScreenShotLoaded = True

End Sub

Private Sub UpdateDisplayPlayer()

    Select Case tabPlayerSubCategory.SelectedItem
        Case TAB_SUB_CAT_PLAYER_ATTRIBUTES
            UpdateDisplayPlayerAttributes
    End Select

End Sub

Private Sub UpdateDisplayPlayerAttributes()

    If SaveFileData.OSE.Player.BaseAttributes <> -1 Then
        UpdateDisplayPlayerDataBaseAttributes
    End If

    If SaveFileData.OSE.Player.BaseData <> -1 Then
        UpdateDisplayPlayerDataBaseData
    End If

    If SaveFileData.OSE.Player.BaseHealth <> -1 Then
        UpdateDisplayPlayerDataBaseHealth
    End If

End Sub

Private Sub UpdateDisplayPlayerDataBaseAttributes()

    Dim i As Integer
    Dim Offset As Long
    
    Offset = SaveFileData.OSE.Player.BaseAttributes

    For i = 0 To 7
        txtAttribute(i).Text = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + i)
    Next i

End Sub

Private Sub UpdateDisplayPlayerDataBaseData()

    Dim Offset As Long
    
    Offset = SaveFileData.OSE.Player.BaseData

    txtBaseMagicka.Text = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 4) + _
                          SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 5) * BYTE_2
    txtBaseFatigue.Text = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 6) + _
                          SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 7) * BYTE_2

End Sub

Private Sub UpdateDisplayPlayerDataBaseHealth()

    Dim Offset As Long
    
    Offset = SaveFileData.OSE.Player.BaseHealth

    txtBaseHealth.Text = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset) + _
                         SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1) * BYTE_2 + _
                         SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 2) * BYTE_3 + _
                         SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 3) * BYTE_4

End Sub

Private Sub mnuSave_Click()

    cmdSave_Click

End Sub

Private Sub cmdSave_Click()

    cmdSave.Enabled = False
    WriteSaveFile StatusBar, prgProgress
    cmdSave.Enabled = True

End Sub

Private Sub txtAttribute_Change(ByRef Index As Integer)

    ValidateInput txtAttribute(Index), BYTE_MIN, BYTE_MAX

    ' Input is valid, update the data structure
    ModifyPlayerAttribute Index, txtAttribute(Index).Text

End Sub

Private Sub txtBaseHealth_Change()

    ValidateInput txtBaseHealth, HEALTH_MIN, HEALTH_MAX

    ' Input is valid, update the data structure
    ModifyPlayerBaseHealth txtBaseHealth.Text

End Sub

Private Sub txtBaseMagicka_Change()

    ValidateInput txtBaseMagicka, MAGICKA_MIN, MAGICKA_MAX

    ' Input is valid, update the data structure
    ModifyPlayerBaseMagicka txtBaseMagicka.Text

End Sub

Private Sub txtBaseFatigue_Change()

    ValidateInput txtBaseFatigue, FATIGUE_MIN, FATIGUE_MAX

    ' Input is valid, update the data structure
    ModifyPlayerBaseFatigue txtBaseFatigue.Text

End Sub

Private Sub txtLevel_Change()

    ValidateInput txtLevel, INTEGER_MIN, INTEGER_MAX

    ' Input is valid, update the data structure
    ModifySaveFilePlayerLevel txtLevel.Text

End Sub

Private Sub txtLocation_Change()

    ModifySaveFilePlayerLocation txtLocation.Text

End Sub

Private Sub txtName_Change()

    ModifySaveFilePlayerName txtName.Text

End Sub

Private Sub txtSaveFileNumber_Change()

    ValidateInput txtSaveFileNumber, LONG_MIN, LONG_MAX

    ' Input is valid, update the data structure
    ModifySaveFileNumber txtSaveFileNumber.Text

End Sub

Private Sub txtSaveFileVersionMajor_Change()

    ValidateInput txtSaveFileVersionMajor, BYTE_MIN, BYTE_MAX

    ' Input is valid, update the data structure
    ModifySaveFileVersionMajor txtSaveFileVersionMajor.Text

End Sub

Private Sub txtSaveFileVersionMinor_Change()

    ValidateInput txtSaveFileVersionMinor, BYTE_MIN, BYTE_MAX

    ' Input is valid, update the data structure
    ModifySaveFileVersionMinor txtSaveFileVersionMinor.Text
    ModifySaveFileHeaderVersion txtSaveFileVersionMinor.Text
    ModifyChangeRecordVersion txtSaveFileVersionMinor.Text

End Sub

