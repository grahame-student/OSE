VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OSE"
   ClientHeight    =   6015
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   13260
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   13260
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox pnlSaveFile 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4545
      ScaleWidth      =   12705
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   12735
      Begin VB.TextBox txtSaveFileNumber 
         Height          =   285
         Left            =   8760
         TabIndex        =   13
         Text            =   "0"
         Top             =   480
         Width           =   3825
      End
      Begin VB.TextBox txtSaveFileVersionMinor 
         Height          =   285
         Left            =   11850
         TabIndex        =   16
         Text            =   "0"
         ToolTipText     =   "The major version of the savefile format (0 to 255)"
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox txtSaveFileVersionMajor 
         Height          =   285
         Left            =   8760
         TabIndex        =   14
         Text            =   "0"
         ToolTipText     =   "The major version of the savefile format (0 to 255)"
         Top             =   240
         Width           =   735
      End
      Begin MSComctlLib.TabStrip tabSaveSubCategory 
         Height          =   4335
         Left            =   120
         TabIndex        =   12
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
      Begin VB.PictureBox picScreenShot 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2250
         Left            =   8760
         ScaleHeight     =   150
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   255
         TabIndex        =   10
         Top             =   2160
         Width           =   3825
      End
      Begin VB.ListBox lstPlugIns 
         Height          =   1035
         Left            =   8760
         TabIndex        =   5
         Top             =   1080
         Width           =   3825
      End
      Begin VB.Label lblSaveTime 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   8760
         TabIndex        =   17
         Top             =   840
         Width           =   3825
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Save File Minor Version"
         Height          =   195
         Left            =   10080
         TabIndex        =   15
         Top             =   285
         Width           =   1665
      End
      Begin VB.Label lblScreenshot 
         AutoSize        =   -1  'True
         Caption         =   "Screenshot"
         Height          =   195
         Left            =   7800
         TabIndex        =   11
         Top             =   2160
         Width           =   810
      End
      Begin VB.Label lblVersionTag 
         AutoSize        =   -1  'True
         Caption         =   "Save File Major Version"
         Height          =   195
         Left            =   6945
         TabIndex        =   9
         Top             =   285
         Width           =   1665
      End
      Begin VB.Label lblSaveFileNumberTag 
         AutoSize        =   -1  'True
         Caption         =   "Save File Number"
         Height          =   195
         Left            =   7350
         TabIndex        =   8
         Top             =   525
         Width           =   1260
      End
      Begin VB.Label lblSavetimeTag 
         AutoSize        =   -1  'True
         Caption         =   "Save Time"
         Height          =   195
         Left            =   7845
         TabIndex        =   7
         Top             =   870
         Width           =   765
      End
      Begin VB.Label lblPlugins 
         AutoSize        =   -1  'True
         Caption         =   "PlugIns"
         Height          =   195
         Left            =   8085
         TabIndex        =   6
         Top             =   1080
         Width           =   525
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   11520
      TabIndex        =   18
      Top             =   5280
      Width           =   1575
   End
   Begin VB.PictureBox pnlPlayer 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4545
      ScaleWidth      =   12705
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   12735
      Begin VB.PictureBox pnlBaseStats 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   9720
         ScaleHeight     =   945
         ScaleWidth      =   2865
         TabIndex        =   43
         Top             =   1320
         Width           =   2895
         Begin VB.TextBox txtBaseHealth 
            Height          =   285
            Left            =   1080
            TabIndex        =   44
            Text            =   "0"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtBaseFatigue 
            Height          =   285
            Left            =   1080
            TabIndex        =   45
            Text            =   "0"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtBaseMagicka 
            Height          =   285
            Left            =   1080
            TabIndex        =   46
            Text            =   "0"
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblBaseMagicka 
            Caption         =   "B. Magicka"
            Height          =   255
            Left            =   120
            TabIndex        =   49
            Top             =   135
            Width           =   855
         End
         Begin VB.Label lblBaseFatigue 
            Caption         =   "B. Fatigue"
            Height          =   255
            Left            =   120
            TabIndex        =   48
            Top             =   375
            Width           =   855
         End
         Begin VB.Label lblBaseHealth 
            Caption         =   "B. Health"
            Height          =   255
            Left            =   120
            TabIndex        =   47
            Top             =   615
            Width           =   855
         End
      End
      Begin VB.PictureBox pnlBasicInformation 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   9720
         ScaleHeight     =   1185
         ScaleWidth      =   2865
         TabIndex        =   36
         Top             =   120
         Width           =   2895
         Begin VB.TextBox txtLocation 
            Height          =   525
            Left            =   1080
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   37
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtLevel 
            Height          =   285
            Left            =   1080
            TabIndex        =   38
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1080
            MaxLength       =   254
            TabIndex        =   39
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblName 
            Caption         =   "Name"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   135
            Width           =   855
         End
         Begin VB.Label lblLevel 
            Caption         =   "Level"
            Height          =   255
            Left            =   120
            TabIndex        =   41
            Top             =   375
            Width           =   855
         End
         Begin VB.Label lblLocation 
            Caption         =   "Location"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   615
            Width           =   855
         End
      End
      Begin VB.PictureBox pnlBaseAttributes 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   9720
         ScaleHeight     =   2145
         ScaleWidth      =   2865
         TabIndex        =   19
         Top             =   2280
         Width           =   2895
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   20
            Text            =   "0"
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   21
            Text            =   "0"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   22
            Text            =   "0"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   23
            Text            =   "0"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   24
            Text            =   "0"
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   25
            Text            =   "0"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   26
            Text            =   "0"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   27
            Text            =   "0"
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblStrength 
            Caption         =   "Strength"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   135
            Width           =   855
         End
         Begin VB.Label lblIntelligence 
            Caption         =   "Intelligence"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   375
            Width           =   855
         End
         Begin VB.Label lblAgility 
            Caption         =   "Agility"
            Height          =   255
            Left            =   120
            TabIndex        =   33
            Top             =   855
            Width           =   855
         End
         Begin VB.Label lblWillPower 
            Caption         =   "Willpower"
            Height          =   255
            Left            =   120
            TabIndex        =   32
            Top             =   615
            Width           =   855
         End
         Begin VB.Label lblEndurance 
            Caption         =   "Endurance"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   1335
            Width           =   855
         End
         Begin VB.Label lblSpeed 
            Caption         =   "Speed"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   1095
            Width           =   855
         End
         Begin VB.Label lblLuck 
            Caption         =   "Luck"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   1815
            Width           =   855
         End
         Begin VB.Label lblPersonality 
            Caption         =   "Personality"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   1575
            Width           =   855
         End
      End
      Begin MSComctlLib.TabStrip tabPlayerSubCategory 
         Height          =   4335
         Left            =   120
         TabIndex        =   50
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
      Width           =   12975
      _ExtentX        =   22886
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
      Width           =   13260
      _ExtentX        =   23389
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11642
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   11642
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

Private Sub mnuQuit_Click()

    End

End Sub

Private Sub mnuSave_Click()

    cmdSave_Click

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

Private Sub OpenSaveFile()

    cmdSave.Enabled = False
    
    ReadSaveFile StatusBar, prgProgress
    
    If Not SaveFileData.OSE.LoadSuccessful Then
        Exit Sub
    End If
    
    StatusBar.Panels(1).Text = "Scanning for markers"
    ScanForMarkers

    StatusBar.Panels(1).Text = "Read complete"

    tabCategory.SelectedItem = TAB_CAT_SAVE_FILE
    tabCategory_Click

    cmdSave.Enabled = True

End Sub

Private Sub ScanForMarkers()

    SaveFileData.OSE.Player.PlayerRecord = LocatePlayerRecord
    If SaveFileData.OSE.Player.PlayerRecord <> -1 Then
        ScanForPlayerMarkers
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
        Offset = Offset + _
                 (SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset) + _
                  SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1) * BYTE_2) * 5
        Offset = Offset + 2
    Else
        SaveFileData.OSE.Player.Factions = -1
    End If
    
    ' Check for spell list
    If ((SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Flags And BIT_5) <> 0) Then
        SaveFileData.OSE.Player.SpellList = Offset
        Offset = Offset + _
                 (SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset) + _
                  SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1) * BYTE_2) * 4
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
                 (SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset) + _
                  SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + 1) * BYTE_2) * 5
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

    DisplayScreenShot

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
'    ModifySaveFileVersionMinor txtSaveFileVersionMinor.Text

End Sub

