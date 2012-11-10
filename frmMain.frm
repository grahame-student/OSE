VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OSE"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7710
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   7710
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar prgProgress 
      Height          =   255
      Left            =   4080
      TabIndex        =   23
      Top             =   4800
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
      TabIndex        =   22
      Top             =   4815
      Width           =   7710
      _ExtentX        =   13600
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6747
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   6747
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pnlSaveHeader 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   120
      ScaleHeight     =   4545
      ScaleWidth      =   7425
      TabIndex        =   0
      Top             =   120
      Width           =   7455
      Begin VB.TextBox txtBaseHealth 
         Height          =   285
         Left            =   1080
         TabIndex        =   14
         Text            =   "0"
         Top             =   3720
         Width           =   1695
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
         Height          =   375
         Left            =   120
         TabIndex        =   39
         Top             =   4080
         Width           =   2655
      End
      Begin VB.TextBox txtBaseFatigue 
         Height          =   285
         Left            =   1080
         TabIndex        =   13
         Text            =   "0"
         Top             =   3480
         Width           =   1695
      End
      Begin VB.TextBox txtBaseMagicka 
         Height          =   285
         Left            =   1080
         TabIndex        =   12
         Text            =   "0"
         Top             =   3240
         Width           =   1695
      End
      Begin VB.TextBox txtAttribute 
         Height          =   285
         Index           =   7
         Left            =   1080
         TabIndex        =   11
         Text            =   "0"
         Top             =   2880
         Width           =   1695
      End
      Begin VB.TextBox txtAttribute 
         Height          =   285
         Index           =   6
         Left            =   1080
         TabIndex        =   10
         Text            =   "0"
         Top             =   2640
         Width           =   1695
      End
      Begin VB.TextBox txtAttribute 
         Height          =   285
         Index           =   5
         Left            =   1080
         TabIndex        =   9
         Text            =   "0"
         Top             =   2400
         Width           =   1695
      End
      Begin VB.TextBox txtAttribute 
         Height          =   285
         Index           =   4
         Left            =   1080
         TabIndex        =   8
         Text            =   "0"
         Top             =   2160
         Width           =   1695
      End
      Begin VB.TextBox txtAttribute 
         Height          =   285
         Index           =   3
         Left            =   1080
         TabIndex        =   7
         Text            =   "0"
         Top             =   1920
         Width           =   1695
      End
      Begin VB.TextBox txtAttribute 
         Height          =   285
         Index           =   2
         Left            =   1080
         TabIndex        =   6
         Text            =   "0"
         Top             =   1680
         Width           =   1695
      End
      Begin VB.TextBox txtAttribute 
         Height          =   285
         Index           =   1
         Left            =   1080
         TabIndex        =   5
         Text            =   "0"
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtAttribute 
         Height          =   285
         Index           =   0
         Left            =   1080
         TabIndex        =   4
         Text            =   "0"
         Top             =   1200
         Width           =   1695
      End
      Begin VB.TextBox txtLocation 
         Height          =   525
         Left            =   1080
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtLevel 
         Height          =   285
         Left            =   1080
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.TextBox txtName 
         Height          =   285
         Left            =   1080
         TabIndex        =   1
         Top             =   120
         Width           =   1695
      End
      Begin VB.ListBox lstPlugIns 
         Height          =   1035
         Left            =   4320
         TabIndex        =   24
         Top             =   960
         Width           =   3015
      End
      Begin VB.Timer tmrLoad 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   6960
         Top             =   4080
      End
      Begin VB.PictureBox picScreenShot 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2250
         Left            =   3240
         ScaleHeight     =   150
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   2160
         Width           =   3825
      End
      Begin VB.Label lblBaseHealth 
         Caption         =   "B. Health"
         Height          =   255
         Left            =   120
         TabIndex        =   40
         Top             =   3735
         Width           =   855
      End
      Begin VB.Label lblBaseFatigue 
         Caption         =   "B. Fatigue"
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   3495
         Width           =   855
      End
      Begin VB.Label lblBaseMagicka 
         Caption         =   "B. Magicka"
         Height          =   255
         Left            =   120
         TabIndex        =   37
         Top             =   3255
         Width           =   855
      End
      Begin VB.Label lblPersonality 
         Caption         =   "Personality"
         Height          =   255
         Left            =   120
         TabIndex        =   36
         Top             =   2655
         Width           =   855
      End
      Begin VB.Label lblLuck 
         Caption         =   "Luck"
         Height          =   255
         Left            =   120
         TabIndex        =   35
         Top             =   2895
         Width           =   855
      End
      Begin VB.Label lblSpeed 
         Caption         =   "Speed"
         Height          =   255
         Left            =   120
         TabIndex        =   34
         Top             =   2175
         Width           =   855
      End
      Begin VB.Label lblEndurance 
         Caption         =   "Endurance"
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   2415
         Width           =   855
      End
      Begin VB.Label lblWillPower 
         Caption         =   "Willpower"
         Height          =   255
         Left            =   120
         TabIndex        =   32
         Top             =   1695
         Width           =   855
      End
      Begin VB.Label lblAgility 
         Caption         =   "Agility"
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   1935
         Width           =   855
      End
      Begin VB.Label lblIntelligence 
         Caption         =   "Intelligence"
         Height          =   255
         Left            =   120
         TabIndex        =   30
         Top             =   1455
         Width           =   855
      End
      Begin VB.Label lblStrength 
         Caption         =   "Strength"
         Height          =   255
         Left            =   120
         TabIndex        =   29
         Top             =   1215
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
      Begin VB.Label lblLevel 
         Caption         =   "Level"
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   375
         Width           =   855
      End
      Begin VB.Label lblName 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   135
         Width           =   855
      End
      Begin VB.Label lblPlugins 
         AutoSize        =   -1  'True
         Caption         =   "PlugIns"
         Height          =   195
         Left            =   3000
         TabIndex        =   25
         Top             =   960
         Width           =   525
      End
      Begin VB.Label lblSavetimeTag 
         AutoSize        =   -1  'True
         Caption         =   "Save Time"
         Height          =   195
         Left            =   3000
         TabIndex        =   21
         Top             =   630
         Width           =   765
      End
      Begin VB.Label lblSavetime 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   20
         Top             =   600
         Width           =   3060
      End
      Begin VB.Label lblSaveFileNumber 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   360
         Width           =   3060
      End
      Begin VB.Label lblSaveFileNumberTag 
         AutoSize        =   -1  'True
         Caption         =   "Save File Number"
         Height          =   195
         Left            =   3000
         TabIndex        =   19
         Top             =   390
         Width           =   1260
      End
      Begin VB.Line linDivider 
         BorderColor     =   &H80000000&
         X1              =   2880
         X2              =   2880
         Y1              =   120
         Y2              =   4440
      End
      Begin VB.Label lblVersion 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4320
         TabIndex        =   16
         Top             =   120
         Width           =   3060
      End
      Begin VB.Label lblVersionTag 
         AutoSize        =   -1  'True
         Caption         =   "Save File Version"
         Height          =   195
         Left            =   3000
         TabIndex        =   15
         Top             =   150
         Width           =   1230
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
DefObj A-Z

Private Sub Form_Load()

    frmMain.Caption = PROGRAM_TITLE
    tmrLoad.Enabled = True

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

    tmrLoad.Enabled = False
    End

End Sub

Private Sub tmrLoad_Timer()

    cmdSave.Enabled = False
    
    ReadSaveFile StatusBar, prgProgress

    
    StatusBar.Panels(1).Text = "Scanning for markers"
    ScanForMarkers

    StatusBar.Panels(1).Text = "Read complete"
    UpdateDisplay

    cmdSave.Enabled = True
    tmrLoad.Enabled = False

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
    
    Dim i As Integer
    
    lblVersion.Caption = SaveFileData.FileHeader.MajorVersion & "." & SaveFileData.FileHeader.MinorVersion
    lblSaveFileNumber.Caption = SaveFileData.SaveHeader.SaveNumber
    lblSavetime.Caption = SaveFileData.SaveHeader.GameTime.Hour & ":" & _
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
        
    txtName.Text = SaveFileData.SaveHeader.PlayerName
    txtLevel.Text = SaveFileData.SaveHeader.PlayerLevel
    txtLocation.Text = SaveFileData.SaveHeader.PlayerLocation
        
    UpdateDisplayPlayerData
        
    DisplayScreenShot
    
End Sub

Private Sub UpdateDisplayPlayerData()

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

    ' Remove once happy things work
    UpdateDisplay

End Sub

Private Sub ModifyPlayerAttribute(ByVal PlayerAttribute As Integer, ByVal AttributeValue As Byte)

    SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(SaveFileData.OSE.Player.BaseAttributes + PlayerAttribute) = AttributeValue

End Sub

