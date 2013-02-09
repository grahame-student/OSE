VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "OSE"
   ClientHeight    =   6015
   ClientLeft      =   150
   ClientTop       =   780
   ClientWidth     =   8055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6015
   ScaleWidth      =   8055
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog CommonDialog 
      Left            =   120
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pnlSaveFile 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   4575
      Left            =   240
      ScaleHeight     =   4545
      ScaleWidth      =   7545
      TabIndex        =   4
      Top             =   480
      Visible         =   0   'False
      Width           =   7575
      Begin VB.PictureBox pnlAll 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   1320
         ScaleHeight     =   4305
         ScaleWidth      =   6105
         TabIndex        =   39
         Top             =   120
         Width           =   6135
         Begin VB.PictureBox picScreenShot 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   2250
            Left            =   2175
            OLEDropMode     =   1  'Manual
            ScaleHeight     =   150
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   256
            TabIndex        =   50
            ToolTipText     =   "Screenshot created at time of save. Click to save to a file. Dropping a bitmap here will replace the image"
            Top             =   1920
            Width           =   3840
         End
         Begin VB.ListBox lstPlugIns 
            Height          =   1035
            Left            =   2175
            TabIndex        =   48
            ToolTipText     =   "The plugins in use when the save was created."
            Top             =   840
            Width           =   3840
         End
         Begin VB.TextBox txtSaveFileNumber 
            Height          =   285
            Left            =   2175
            TabIndex        =   40
            Text            =   "0"
            ToolTipText     =   "The number of saves made since starting the character. (0 to 2147483647)"
            Top             =   315
            Width           =   3840
         End
         Begin VB.TextBox txtSaveFileVersionMajor 
            Height          =   285
            Left            =   2175
            TabIndex        =   42
            Text            =   "0"
            ToolTipText     =   "The major version of the savefile format. (0 to 255)"
            Top             =   75
            Width           =   735
         End
         Begin VB.TextBox txtSaveFileVersionMinor 
            Height          =   285
            Left            =   5280
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
            Left            =   1200
            TabIndex        =   51
            Top             =   1920
            Width           =   810
         End
         Begin VB.Label lblPlugins 
            AutoSize        =   -1  'True
            Caption         =   "PlugIns"
            Height          =   195
            Left            =   1485
            TabIndex        =   49
            Top             =   960
            Width           =   525
         End
         Begin VB.Label lblSaveTime 
            BorderStyle     =   1  'Fixed Single
            Height          =   285
            Left            =   2175
            TabIndex        =   46
            ToolTipText     =   "The time that the save file was created. Click to enter a new time manually."
            Top             =   600
            Width           =   3840
         End
         Begin VB.Label lblSavetimeTag 
            AutoSize        =   -1  'True
            Caption         =   "Save Time"
            Height          =   195
            Left            =   1245
            TabIndex        =   47
            Top             =   630
            Width           =   765
         End
         Begin VB.Label lblSaveFileNumberTag 
            AutoSize        =   -1  'True
            Caption         =   "Save File Number"
            Height          =   195
            Left            =   750
            TabIndex        =   45
            Top             =   360
            Width           =   1260
         End
         Begin VB.Label lblSaveFileVersionMajor 
            AutoSize        =   -1  'True
            Caption         =   "Save File Major Version"
            Height          =   195
            Left            =   345
            TabIndex        =   44
            Top             =   120
            Width           =   1665
         End
         Begin VB.Label lblSaveFileVersionMinor 
            AutoSize        =   -1  'True
            Caption         =   "Save File Minor Version"
            Height          =   195
            Left            =   3495
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
      Left            =   6360
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
      ScaleWidth      =   7545
      TabIndex        =   0
      Top             =   480
      Visible         =   0   'False
      Width           =   7575
      Begin VB.PictureBox pnlSpells 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   1320
         ScaleHeight     =   4305
         ScaleWidth      =   6105
         TabIndex        =   104
         Top             =   120
         Width           =   6135
         Begin VB.CommandButton cmdAddSpell 
            Caption         =   "Add"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4800
            TabIndex        =   110
            Top             =   3840
            Width           =   1215
         End
         Begin VB.CommandButton cmdRemoveSpell 
            Caption         =   "Remove"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4800
            TabIndex        =   109
            Top             =   840
            Width           =   1215
         End
         Begin VB.ListBox lstAllSpells 
            Height          =   2205
            Left            =   1320
            TabIndex        =   107
            ToolTipText     =   "The list of all available factions"
            Top             =   1440
            Width           =   4695
         End
         Begin VB.ComboBox cboPlayerSpells 
            Height          =   315
            Left            =   1320
            TabIndex        =   105
            ToolTipText     =   "The factions the character is a member of"
            Top             =   120
            Width           =   4695
         End
         Begin VB.Label lblAllSpells 
            Caption         =   "All Spells"
            Height          =   255
            Left            =   120
            TabIndex        =   108
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Line linDividor2 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   6000
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label lblSpells 
            Caption         =   "Spells"
            Height          =   255
            Left            =   120
            TabIndex        =   106
            Top             =   120
            Width           =   1095
         End
      End
      Begin VB.PictureBox pnlPosition 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   1335
         Left            =   4320
         ScaleHeight     =   1305
         ScaleWidth      =   3105
         TabIndex        =   111
         Top             =   120
         Width           =   3135
         Begin VB.TextBox txtLocationZ 
            Height          =   285
            Left            =   840
            TabIndex        =   118
            Text            =   "0"
            ToolTipText     =   "The character's level (0-32767)"
            Top             =   960
            Width           =   2295
         End
         Begin VB.TextBox txtLocationY 
            Height          =   285
            Left            =   840
            TabIndex        =   116
            Text            =   "0"
            ToolTipText     =   "The character's level (0-32767)"
            Top             =   720
            Width           =   2295
         End
         Begin VB.TextBox txtLocationX 
            Height          =   285
            Left            =   840
            TabIndex        =   114
            Text            =   "0"
            ToolTipText     =   "The character's level (0-32767)"
            Top             =   480
            Width           =   2295
         End
         Begin VB.ComboBox cboLocationCell 
            Height          =   315
            Left            =   840
            TabIndex        =   113
            Text            =   "cboLocationCell"
            Top             =   120
            Width           =   2295
         End
         Begin VB.Label lblLocationZ 
            Caption         =   "Z"
            Height          =   255
            Left            =   120
            TabIndex        =   119
            Top             =   975
            Width           =   735
         End
         Begin VB.Label lblLocationY 
            Caption         =   "Y"
            Height          =   255
            Left            =   120
            TabIndex        =   117
            Top             =   735
            Width           =   735
         End
         Begin VB.Label lblLocationX 
            Caption         =   "X"
            Height          =   255
            Left            =   120
            TabIndex        =   115
            Top             =   495
            Width           =   735
         End
         Begin VB.Label lblCell 
            Caption         =   "Cell"
            Height          =   255
            Left            =   120
            TabIndex        =   112
            Top             =   150
            Width           =   735
         End
      End
      Begin VB.PictureBox pnlFactions 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   1320
         ScaleHeight     =   4305
         ScaleWidth      =   6105
         TabIndex        =   95
         Top             =   120
         Visible         =   0   'False
         Width           =   6135
         Begin VB.CommandButton cmdRemoveFaction 
            Caption         =   "Remove"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4800
            TabIndex        =   102
            Top             =   840
            Width           =   1215
         End
         Begin VB.CommandButton cmdAddFaction 
            Caption         =   "Add"
            Enabled         =   0   'False
            Height          =   375
            Left            =   4800
            TabIndex        =   101
            Top             =   3840
            Width           =   1215
         End
         Begin VB.ListBox lstAllFactions 
            Height          =   2205
            Left            =   1320
            TabIndex        =   100
            ToolTipText     =   "The list of all available factions"
            Top             =   1440
            Width           =   4695
         End
         Begin VB.ComboBox cboFactionRank 
            Height          =   315
            Left            =   1320
            TabIndex        =   99
            ToolTipText     =   "The character's rank within the faction"
            Top             =   480
            Width           =   4695
         End
         Begin VB.ComboBox cboFactions 
            Height          =   315
            Left            =   1320
            TabIndex        =   96
            ToolTipText     =   "The factions the character is a member of"
            Top             =   120
            Width           =   4695
         End
         Begin VB.Label lblAvailableFactions 
            Caption         =   "All Factions"
            Height          =   255
            Left            =   120
            TabIndex        =   103
            Top             =   1440
            Width           =   1095
         End
         Begin VB.Line linDivider1 
            BorderColor     =   &H00C0C0C0&
            X1              =   120
            X2              =   6000
            Y1              =   1320
            Y2              =   1320
         End
         Begin VB.Label lblFactionRank 
            Caption         =   "Rank"
            Height          =   255
            Left            =   120
            TabIndex        =   98
            Top             =   510
            Width           =   1095
         End
         Begin VB.Label lblFaction 
            Caption         =   "Faction"
            Height          =   255
            Left            =   120
            TabIndex        =   97
            Top             =   150
            Width           =   1095
         End
      End
      Begin VB.PictureBox pnlSkills 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4335
         Left            =   1320
         ScaleHeight     =   4305
         ScaleWidth      =   6105
         TabIndex        =   52
         Top             =   120
         Visible         =   0   'False
         Width           =   6135
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   20
            Left            =   4320
            TabIndex        =   93
            Text            =   "0"
            ToolTipText     =   "The character's skill at talking to others"
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   19
            Left            =   4320
            TabIndex        =   87
            Text            =   "0"
            ToolTipText     =   "The character's skill at sneaking (0-255)"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   18
            Left            =   4320
            TabIndex        =   88
            Text            =   "0"
            ToolTipText     =   "The character's skill at picking locks (0-255)"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   17
            Left            =   4320
            TabIndex        =   89
            Text            =   "0"
            ToolTipText     =   "The character's skill as a merchant (0-255)"
            Top             =   120
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   16
            Left            =   1320
            TabIndex        =   85
            Text            =   "0"
            ToolTipText     =   "The character's skill as a marksman (0-255)"
            Top             =   3960
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   15
            Left            =   1320
            TabIndex        =   69
            Text            =   "0"
            ToolTipText     =   "The character's skill at using light armour (0-255)"
            Top             =   3720
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   14
            Left            =   1320
            TabIndex        =   70
            Text            =   "0"
            ToolTipText     =   "The character's skill at acrobatics (0-255)"
            Top             =   3480
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   13
            Left            =   1320
            TabIndex        =   71
            Text            =   "0"
            ToolTipText     =   "The character's skill at restoration magic (0-255)"
            Top             =   3240
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   12
            Left            =   1320
            TabIndex        =   72
            Text            =   "0"
            ToolTipText     =   "The character's skill at mysticism magic (0-255)"
            Top             =   3000
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   11
            Left            =   1320
            TabIndex        =   73
            Text            =   "0"
            ToolTipText     =   "The character's skill at illusion magic (0-255)"
            Top             =   2760
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   10
            Left            =   1320
            TabIndex        =   74
            Text            =   "0"
            ToolTipText     =   "The character's skill at destruction magic (0-255)"
            Top             =   2520
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   9
            Left            =   1320
            TabIndex        =   75
            Text            =   "0"
            ToolTipText     =   "The character's skill at conjuration magic (0-255)"
            Top             =   2280
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   8
            Left            =   1320
            TabIndex        =   76
            Text            =   "0"
            ToolTipText     =   "The character's skill at alteration magic (0-255)"
            Top             =   2040
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   7
            Left            =   1320
            TabIndex        =   65
            Text            =   "0"
            ToolTipText     =   "The character's skill at alchemy (0-255)"
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   6
            Left            =   1320
            TabIndex        =   66
            Text            =   "0"
            ToolTipText     =   "The character's skill at using heavy armour (0-255)"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   5
            Left            =   1320
            TabIndex        =   61
            Text            =   "0"
            ToolTipText     =   "The character's skill at hand to hand combat (0-255)"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   4
            Left            =   1320
            TabIndex        =   62
            Text            =   "0"
            ToolTipText     =   "The character's skill with a blunt weapon (0-255)"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   3
            Left            =   1320
            TabIndex        =   57
            Text            =   "0"
            ToolTipText     =   "The character's skill at blocking (0-255)"
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   2
            Left            =   1320
            TabIndex        =   58
            Text            =   "0"
            ToolTipText     =   "The character's skill with a bladed weapon (0-255)"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   1
            Left            =   1320
            TabIndex        =   55
            Text            =   "0"
            ToolTipText     =   "The character's skill at athletics (0-255)"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtSkill 
            Height          =   285
            Index           =   0
            Left            =   1320
            TabIndex        =   54
            Text            =   "0"
            ToolTipText     =   "The character's skill as an armourer (0-255)"
            Top             =   120
            Width           =   1695
         End
         Begin VB.Label lblSpeechcraft 
            Caption         =   "Speechcraft"
            Height          =   255
            Left            =   3120
            TabIndex        =   94
            Top             =   855
            Width           =   1095
         End
         Begin VB.Label lblMercantile 
            Caption         =   "Mercantile"
            Height          =   255
            Left            =   3120
            TabIndex        =   92
            Top             =   135
            Width           =   1095
         End
         Begin VB.Label lblSneak 
            Caption         =   "Sneak"
            Height          =   255
            Left            =   3120
            TabIndex        =   91
            Top             =   615
            Width           =   1095
         End
         Begin VB.Label lblSecurity 
            Caption         =   "Security"
            Height          =   255
            Left            =   3120
            TabIndex        =   90
            Top             =   375
            Width           =   1095
         End
         Begin VB.Label lblMarksman 
            Caption         =   "Marksman"
            Height          =   255
            Left            =   120
            TabIndex        =   86
            Top             =   3975
            Width           =   1095
         End
         Begin VB.Label lblAlteration 
            Caption         =   "Alteration"
            Height          =   255
            Left            =   120
            TabIndex        =   84
            Top             =   2055
            Width           =   1095
         End
         Begin VB.Label lblConjuration 
            Caption         =   "Conjuration"
            Height          =   255
            Left            =   120
            TabIndex        =   83
            Top             =   2295
            Width           =   1095
         End
         Begin VB.Label lblIllusion 
            Caption         =   "Illusion"
            Height          =   255
            Left            =   120
            TabIndex        =   82
            Top             =   2775
            Width           =   1095
         End
         Begin VB.Label lblDestruction 
            Caption         =   "Destruction"
            Height          =   255
            Left            =   120
            TabIndex        =   81
            Top             =   2535
            Width           =   1095
         End
         Begin VB.Label lblRestoration 
            Caption         =   "Restoration"
            Height          =   255
            Left            =   120
            TabIndex        =   80
            Top             =   3255
            Width           =   1095
         End
         Begin VB.Label lblMysticism 
            Caption         =   "Mysticism"
            Height          =   255
            Left            =   120
            TabIndex        =   79
            Top             =   3015
            Width           =   1095
         End
         Begin VB.Label lblLightArmour 
            Caption         =   "Light Armour"
            Height          =   255
            Left            =   120
            TabIndex        =   78
            Top             =   3735
            Width           =   1095
         End
         Begin VB.Label lblAcrobatics 
            Caption         =   "Acrobatics"
            Height          =   255
            Left            =   120
            TabIndex        =   77
            Top             =   3495
            Width           =   1095
         End
         Begin VB.Label lblHeavyArmour 
            Caption         =   "Heavy Armour"
            Height          =   255
            Left            =   120
            TabIndex        =   68
            Top             =   1575
            Width           =   1095
         End
         Begin VB.Label lblAlchemy 
            Caption         =   "Alchemy"
            Height          =   255
            Left            =   120
            TabIndex        =   67
            Top             =   1815
            Width           =   1095
         End
         Begin VB.Label lblBlunt 
            Caption         =   "Blunt"
            Height          =   255
            Left            =   120
            TabIndex        =   64
            Top             =   1095
            Width           =   1095
         End
         Begin VB.Label lblHandToHand 
            Caption         =   "Hand to Hand"
            Height          =   255
            Left            =   120
            TabIndex        =   63
            Top             =   1335
            Width           =   1095
         End
         Begin VB.Label lblBlade 
            Caption         =   "Blade"
            Height          =   255
            Left            =   120
            TabIndex        =   60
            Top             =   615
            Width           =   1095
         End
         Begin VB.Label lblBlock 
            Caption         =   "Block"
            Height          =   255
            Left            =   120
            TabIndex        =   59
            Top             =   855
            Width           =   1095
         End
         Begin VB.Label lblAthletics 
            Caption         =   "Athletics"
            Height          =   255
            Left            =   120
            TabIndex        =   56
            Top             =   375
            Width           =   1095
         End
         Begin VB.Label lblArmourer 
            Caption         =   "Armourer"
            Height          =   255
            Left            =   120
            TabIndex        =   53
            Top             =   135
            Width           =   1095
         End
      End
      Begin VB.PictureBox pnlBaseStats 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   1320
         ScaleHeight     =   945
         ScaleWidth      =   2865
         TabIndex        =   31
         Top             =   1320
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox txtBaseHealth 
            Height          =   285
            Left            =   1080
            TabIndex        =   32
            Text            =   "0"
            ToolTipText     =   "The character's base health (0-30000)"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtBaseFatigue 
            Height          =   285
            Left            =   1080
            TabIndex        =   33
            Text            =   "0"
            ToolTipText     =   "The character's base fatigue (0-10000)"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtBaseMagicka 
            Height          =   285
            Left            =   1080
            TabIndex        =   34
            Text            =   "0"
            ToolTipText     =   "The player's base magicka (0-10000)"
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
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox txtLocation 
            Height          =   525
            Left            =   1080
            MaxLength       =   254
            MultiLine       =   -1  'True
            TabIndex        =   25
            ToolTipText     =   "The character's location at the time the save was made (0-254 characters)"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtLevel 
            Height          =   285
            Left            =   1080
            TabIndex        =   26
            ToolTipText     =   "The character's level (0-32767)"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtName 
            Height          =   285
            Left            =   1080
            MaxLength       =   254
            TabIndex        =   27
            ToolTipText     =   "The character's name (0-254 characters)"
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
         Visible         =   0   'False
         Width           =   2895
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   7
            Left            =   1080
            TabIndex        =   8
            Text            =   "0"
            ToolTipText     =   "The character's base luck (0-255)"
            Top             =   1800
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   6
            Left            =   1080
            TabIndex        =   9
            Text            =   "0"
            ToolTipText     =   "The character's base personality (0-255)"
            Top             =   1560
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   5
            Left            =   1080
            TabIndex        =   10
            Text            =   "0"
            ToolTipText     =   "The character's base endurance (0-255)"
            Top             =   1320
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   4
            Left            =   1080
            TabIndex        =   11
            Text            =   "0"
            ToolTipText     =   "The character's base speed (0-255)"
            Top             =   1080
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   3
            Left            =   1080
            TabIndex        =   12
            Text            =   "0"
            ToolTipText     =   "The character's base agility (0-255)"
            Top             =   840
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   2
            Left            =   1080
            TabIndex        =   13
            Text            =   "0"
            ToolTipText     =   "The character's base willpower (0-255)"
            Top             =   600
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   1
            Left            =   1080
            TabIndex        =   14
            Text            =   "0"
            ToolTipText     =   "The character's base intelligence (0-255)"
            Top             =   360
            Width           =   1695
         End
         Begin VB.TextBox txtAttribute 
            Height          =   285
            Index           =   0
            Left            =   1080
            TabIndex        =   15
            Text            =   "0"
            ToolTipText     =   "The character's base strength (0-255)"
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
            NumTabs         =   4
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Attributes"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Skills"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Factions"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Spells"
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
      Width           =   7815
      _ExtentX        =   13785
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
      Width           =   8055
      _ExtentX        =   14208
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7064
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   7064
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
         Enabled         =   0   'False
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

Private Sub PositionPanels()

    pnlPlayer.Top = TAB_CAT_TOP
    pnlPlayer.Left = TAB_CAT_LEFT
    
    pnlSaveFile.Top = TAB_CAT_TOP
    pnlSaveFile.Left = TAB_CAT_LEFT

    pnlAll.Top = TAB_SUB_CAT_TOP
    pnlAll.Left = TAB_SUB_CAT_LEFT

    pnlSpells.Top = TAB_SUB_CAT_TOP
    pnlSpells.Left = TAB_SUB_CAT_LEFT

    pnlFactions.Top = TAB_SUB_CAT_TOP
    pnlFactions.Left = TAB_SUB_CAT_LEFT

    pnlSkills.Top = TAB_SUB_CAT_TOP
    pnlSkills.Left = TAB_SUB_CAT_LEFT

End Sub

Private Sub cboFactionRank_Click()

    If SaveFileData.OSE.Player.FactionList(cboFactions.ListIndex).Level <> cboFactionRank.ListIndex Then
        If cboFactionRank.ListIndex = cboFactionRank.ListCount - 1 Then
            SaveFileData.OSE.Player.FactionList(cboFactions.ListIndex).Suspended = True
            SaveFileData.OSE.Player.FactionList(cboFactions.ListIndex).Level = &HFF&
        Else
            SaveFileData.OSE.Player.FactionList(cboFactions.ListIndex).Suspended = False
            SaveFileData.OSE.Player.FactionList(cboFactions.ListIndex).Level = cboFactionRank.ListIndex
        End If
        ModifyPlayerFaction cboFactions.ListIndex, cboFactionRank.ListIndex
    End If

End Sub

Private Sub cboFactions_Click()

    Dim i As Integer

    cboFactionRank.Clear
    
    For i = 0 To SaveFileData.OSE.Player.FactionList(cboFactions.ListIndex).MaxRank
        cboFactionRank.AddItem SaveFileData.OSE.Player.FactionList(cboFactions.ListIndex).Ranks(i)
    Next i
    
    cboFactionRank.AddItem "Suspended"
    
    UpdateDisplay

    If cboFactions.ListCount > 0 Then
        cmdRemoveFaction.Enabled = True
    Else
        cmdRemoveFaction.Enabled = False
    End If

End Sub

Private Sub cmdAddFaction_Click()

    Dim NewFaction As String
    Dim i As Integer
    Dim iRef As Long

    ' Check player's faction list to see if the selected faction is in there
    ' If it's not there then:
    NewFaction = lstAllFactions.Text

    For i = 0 To cboFactions.ListCount - 1
        If cboFactions.List(i) = NewFaction Then
            MsgBox "Player is already a member of this faction", vbOKOnly, "Unable To Add Faction"
            Exit Sub
        End If
    Next i

    '   Add the selected faction to the player's list
    '   Add the faction at the lowest available level
    cboFactions.AddItem lstAllFactions.Text
    
    iRef = GetIref(lstAllFactions.ItemData(lstAllFactions.ListIndex))
    ' The FormID isn't in the array therefore a new Iref needs to ba added to the FormID array
    If iRef = -1 Then
        ' Make room for the new Iref
        ReDim Preserve SaveFileData.FormIDs.FormIDsList(SaveFileData.FormIDs.NumberOfFormIDs)
        ' Put the Faction's FormID into the new position
        SaveFileData.FormIDs.FormIDsList(SaveFileData.FormIDs.NumberOfFormIDs) = lstAllFactions.ItemData(lstAllFactions.ListIndex)
        ' Increase the number of FormIDs by 1
        iRef = SaveFileData.FormIDs.NumberOfFormIDs
        SaveFileData.FormIDs.NumberOfFormIDs = SaveFileData.FormIDs.NumberOfFormIDs + 1
    End If
    
    ReDim Preserve SaveFileData.OSE.Player.FactionList(SaveFileData.OSE.Player.FactionCount)
    SaveFileData.OSE.Player.FactionList(SaveFileData.OSE.Player.FactionCount).FormID = lstAllFactions.ItemData(lstAllFactions.ListIndex)
    SaveFileData.OSE.Player.FactionList(SaveFileData.OSE.Player.FactionCount).Level = 0
    GetFaction GetFormID(iRef), SaveFileData.OSE.Player.FactionCount
    SaveFileData.OSE.Player.FactionCount = SaveFileData.OSE.Player.FactionCount + 1
    
    cboFactions.ItemData(cboFactions.NewIndex) = iRef

    RebuildPlayerChangeRecord

End Sub

Private Sub cmdAddSpell_Click()

    Dim NewSpell As String
    Dim i As Integer
    Dim iRef As Long

    ' Check player's faction list to see if the selected faction is in there
    ' If it's not there then:
    NewSpell = lstAllSpells.Text

    For i = 0 To cboPlayerSpells.ListCount - 1
        If cboPlayerSpells.List(i) = NewSpell Then
            MsgBox "Player already knows this spell", vbOKOnly, "Unable To Add Spell"
            Exit Sub
        End If
    Next i
    
    '   Add the selected faction to the player's list
    cboPlayerSpells.AddItem lstAllSpells.Text

    iRef = GetIref(lstAllSpells.ItemData(lstAllSpells.ListIndex))
    
    ' The FormID isn't in the array therefore a new Iref needs to ba added to the FormID array
    If iRef = -1 Then
        ' Make room for the new Iref
        ReDim Preserve SaveFileData.FormIDs.FormIDsList(SaveFileData.FormIDs.NumberOfFormIDs)
        ' Put the Spells's FormID into the new position
        SaveFileData.FormIDs.FormIDsList(SaveFileData.FormIDs.NumberOfFormIDs) = lstAllSpells.ItemData(lstAllSpells.ListIndex)
        ' Increase the number of FormIDs by 1
        iRef = SaveFileData.FormIDs.NumberOfFormIDs
        SaveFileData.FormIDs.NumberOfFormIDs = SaveFileData.FormIDs.NumberOfFormIDs + 1
    End If

    ReDim Preserve SaveFileData.OSE.Player.SpellList(SaveFileData.OSE.Player.SpellCount)
    SaveFileData.OSE.Player.SpellList(SaveFileData.OSE.Player.SpellCount).iRef = iRef
    SaveFileData.OSE.Player.SpellList(SaveFileData.OSE.Player.SpellCount).FormID = lstAllSpells.ItemData(lstAllSpells.ListIndex)
    SaveFileData.OSE.Player.SpellCount = SaveFileData.OSE.Player.SpellCount + 1
    
    cboPlayerSpells.ItemData(cboPlayerSpells.NewIndex) = iRef

    RebuildPlayerChangeRecord

End Sub

Private Sub cmdRemoveFaction_Click()

    Dim NewFactionList() As Faction
    Dim ListIndex As Integer
    Dim i As Integer
    Dim FactionNumber As Integer

    ' Remove Faction from the SaveFileData.OSE.Player.FactionList array
    ' Remove Faction from combobox
    ' Rebuild Player's ChangeRecord

    ReDim NewFactionList(SaveFileData.OSE.Player.FactionCount - 2)

    FactionNumber = 0

    ListIndex = cboFactions.ListIndex

    For i = 0 To SaveFileData.OSE.Player.FactionCount - 1
        If i <> cboFactions.ListIndex Then
            NewFactionList(FactionNumber) = SaveFileData.OSE.Player.FactionList(i)
            FactionNumber = FactionNumber + 1
        End If
    Next i

    SaveFileData.OSE.Player.FactionCount = FactionNumber
    SaveFileData.OSE.Player.FactionList = NewFactionList

    cboFactions.RemoveItem cboFactions.ListIndex
        
    If cboFactions.ListCount > ListIndex Then
        cboFactions.ListIndex = ListIndex
    ElseIf cboFactions.ListCount > 0 Then
        ListIndex = ListIndex - 1
        If ListIndex < 0 Then
            ListIndex = 0
        End If
        cboFactions.ListIndex = ListIndex
    End If
    
    ' cboFactions_Click
    
    RebuildPlayerChangeRecord

End Sub

Private Sub Form_Load()

    frmMain.Caption = PROGRAM_TITLE

    GetHomePath

    SetUpProgressBarInStatusBar

    PositionPanels
    
End Sub

Private Sub GetHomePath()

    HomePath = GetSpecialFolder(CSIDL_DOCUMENTS, frmMain)

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

Private Sub LoadFactionData()

    Dim FactionFF As Integer
    Dim NextLine As String
        
    FactionFF = FreeFile

    Open App.Path & "\factions.data" For Input As #FactionFF
    Do Until EOF(FactionFF)
        Line Input #FactionFF, NextLine
        If Mid$(NextLine, 1, 1) <> "#" Then
            ParseFactionData NextLine
        End If
    Loop
    Close #FactionFF

End Sub

Private Sub ParseFactionData(ByVal DataLine As String)

    Static FactionCount As Long
    Dim FactionVariables() As String
    Dim i As Integer
    
    ReDim Preserve FactionData(FactionCount)

    FactionVariables() = Split(DataLine, ",")

    FactionData(FactionCount).FormID = Val(FactionVariables(0))
    FactionData(FactionCount).Name = FactionVariables(1)
    FactionData(FactionCount).PlugIn = FactionVariables(2)
    FactionData(FactionCount).MaxRank = Val(FactionVariables(3))

    ReDim FactionData(FactionCount).Ranks(FactionData(FactionCount).MaxRank)

    For i = 4 To UBound(FactionVariables())
        FactionData(FactionCount).Ranks(i - 4) = FactionVariables(i)
    Next i

    FactionCount = FactionCount + 1

End Sub

Private Sub LoadSpellData()

    Dim SpellFF As Integer
    Dim NextLine As String
        
    SpellFF = FreeFile

    Open App.Path & "\Spells.data" For Input As #SpellFF
    Do Until EOF(SpellFF)
        Line Input #SpellFF, NextLine
        If Mid$(NextLine, 1, 1) <> "#" Then
            ParseSpellData NextLine
        End If
    Loop
    Close #SpellFF

End Sub

Private Sub ParseSpellData(ByVal DataLine As String)

    Static SpellCount As Long
    Dim SpellVariables() As String
    Dim i As Integer
    
    ReDim Preserve SpellData(SpellCount)

    SpellVariables() = Split(DataLine, ",")

    SpellData(SpellCount).FormID = Val(SpellVariables(0))
    SpellData(SpellCount).Name = SpellVariables(1)
    SpellData(SpellCount).PlugIn = SpellVariables(2)

    SpellCount = SpellCount + 1

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

Private Sub lstAllFactions_Click()

    cmdAddFaction.Enabled = True

End Sub

Private Sub lstAllSpells_Click()

    cmdAddSpell.Enabled = True

End Sub

Private Sub mnuQuit_Click()

    End

End Sub

Private Sub mnuAbout_Click()

    frmAbout.Show

End Sub

Private Sub mnuOpen_Click()

    lstAllFactions.Clear
    LoadFactionData

    lstAllSpells.Clear
    LoadSpellData

    ClearSaveFileData
    OpenSaveFile GetFilename

End Sub

Private Sub ClearSaveFileData()

    Dim EmptyESS As ESS

    SaveFileData = EmptyESS

End Sub

Private Function GetFilename() As String

    On Error GoTo CancelError
    
    CommonDialog.Filter = "Oblivion save file (*.dat)|*.dat"
        
    CommonDialog.InitDir = HomePath
    CommonDialog.ShowOpen
    GetFilename = CommonDialog.FileName

    Exit Function

CancelError:
    
    Select Case Err.Number
        Case 91
            MsgBox Err.Number & "  - in GetFilename", vbOKOnly, "Error"
    End Select

End Function

Private Sub OpenSaveFile(ByVal SaveFilePath As String)

    If SaveFilePath = "" Then
        Exit Sub
    End If

    cmdSave.Enabled = False
    
    ReadSave.ReadSaveFile SaveFilePath, StatusBar, prgProgress
    
    If Not SaveFileData.OSE.LoadSuccessful Then
        mnuSave.Enabled = False
        Exit Sub
    End If
    
    StatusBar.Panels(1).Text = "Scanning for markers"
    AnalyseStructure.ScanForMarkers frmMain

    StatusBar.Panels(1).Text = "Read complete"

    tabCategory_Click
        
    mnuSave.Enabled = True

    cmdSave.Enabled = True

End Sub

Private Sub picScreenShot_Click()

    SavePicture picScreenShot.Image, App.Path & "\ScreenShot\screenshot.bmp"

End Sub

Private Sub picScreenShot_OLEDragDrop(Data As DataObject, Effect As Long, Button As Integer, Shift As Integer, X As Single, Y As Single)

    ' TODO: Catch invalid file formats

    picScreenShot.Picture = LoadPicture(Data.Files(1))
    
    ModifySavefileScreenShot picScreenShot

End Sub


Private Sub tabCategory_Click()

    If Not SaveFileData.OSE.LoadSuccessful Then Exit Sub

    HideAllPanels

    Select Case tabCategory.SelectedItem
        Case TAB_CAT_SAVE_FILE
            pnlSaveFile.Visible = True
        Case TAB_CAT_PLAYER
            pnlPlayer.Visible = True
            tabPlayerSubCategory_Click
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
        Case TAB_SUB_CAT_PLAYER_SKILLS
            UpdateDisplayPlayerSkills
        Case TAB_SUB_CAT_PLAYER_FACTIONS
            UpdateDisplayPlayerFactions
        Case TAB_SUB_CAT_PLAYER_SPELLS
            UpdateDisplayPlayerSpells
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

Private Sub UpdateDisplayPlayerSkills()

    Dim i As Integer
    Dim Offset As Long
    
    Offset = SaveFileData.OSE.Player.Skills

    For i = 0 To 20
        txtSkill(i).Text = SaveFileData.ChangeRecords(SaveFileData.OSE.Player.PlayerRecord).Data(Offset + i)
    Next i

End Sub

Private Sub UpdateDisplayPlayerFactions()
    
    Dim i As Integer
        
    If SaveFileData.OSE.Player.FactionCount = 0 Then
        Exit Sub
    End If
        
    If cboFactions.ListCount <> SaveFileData.OSE.Player.FactionCount Then
        For i = 0 To SaveFileData.OSE.Player.FactionCount - 1
            cboFactions.AddItem SaveFileData.OSE.Player.FactionList(i).Name
        Next i
        cboFactions.ListIndex = 0
    End If
        
    If SaveFileData.OSE.Player.FactionList(cboFactions.ListIndex).Suspended Then
        cboFactionRank.ListIndex = cboFactionRank.ListCount - 1
        Exit Sub
    End If
    
    cboFactionRank.ListIndex = SaveFileData.OSE.Player.FactionList(cboFactions.ListIndex).Level
                
End Sub

Private Sub UpdateDisplayPlayerSpells()

    Dim i As Integer

    If cboPlayerSpells.ListCount <> SaveFileData.OSE.Player.SpellCount Then
        For i = 0 To SaveFileData.OSE.Player.SpellCount - 1
            cboPlayerSpells.AddItem SaveFileData.OSE.Player.SpellList(i).Name
        Next i
        cboPlayerSpells.ListIndex = 0
    End If

End Sub

Private Sub mnuSave_Click()

    cmdSave_Click

End Sub

Private Sub cmdSave_Click()

    cmdSave.Enabled = False
    
    WriteSaveFile GetSaveFilename, StatusBar, prgProgress
    
    cmdSave.Enabled = True

End Sub

Private Function GetSaveFilename() As String

    On Error GoTo CancelError
    
    CommonDialog.Filter = "Oblivion save file (*.dat)|*.dat"
        
    CommonDialog.InitDir = HomePath
    CommonDialog.ShowSave
    GetSaveFilename = CommonDialog.FileName

    Exit Function

CancelError:
    
    Select Case Err.Number
        Case 91
            MsgBox Err.Number & "  - in GetSaveFilename", vbOKOnly, "Error"
    End Select

End Function

Private Sub tabPlayerSubCategory_Click()

    HideAllPlayerSubPanels

    Select Case tabPlayerSubCategory.SelectedItem
        Case TAB_SUB_CAT_PLAYER_ATTRIBUTES
            pnlBasicInformation.Visible = True
            pnlBaseStats.Visible = True
            pnlBaseAttributes.Visible = True
            pnlPosition.Visible = False             ' Make true once coded in
        Case TAB_SUB_CAT_PLAYER_SKILLS
            pnlSkills.Visible = True
        Case TAB_SUB_CAT_PLAYER_FACTIONS
            pnlFactions.Visible = True
        Case TAB_SUB_CAT_PLAYER_SPELLS
            pnlSpells.Visible = True
    End Select

    UpdateDisplay

End Sub

Private Sub HideAllPlayerSubPanels()

    pnlBasicInformation.Visible = False
    pnlBaseStats.Visible = False
    pnlBaseAttributes.Visible = False
    pnlPosition.Visible = False
    pnlSkills.Visible = False
    pnlFactions.Visible = False
    pnlSpells.Visible = False

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

    ' Also modify level in player change record?

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

Private Sub txtSkill_Change(Index As Integer)

    ValidateInput txtSkill(Index), 0, 100

    ModifyPlayerSkill Index, txtSkill(Index).Text

End Sub
