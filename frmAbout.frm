VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About MyApp"
   ClientHeight    =   3060
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2112.066
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   750
      Left            =   120
      Picture         =   "frmAbout.frx":1CCA
      ScaleHeight     =   505.68
      ScaleMode       =   0  'User
      ScaleWidth      =   505.68
      TabIndex        =   1
      Top             =   240
      Width           =   750
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   345
      Left            =   4320
      TabIndex        =   0
      Top             =   2640
      Width           =   1260
   End
   Begin VB.Label lblDescription 
      Caption         =   "App Description"
      ForeColor       =   &H00000000&
      Height          =   1650
      Left            =   1080
      TabIndex        =   2
      Top             =   960
      Width           =   4485
   End
   Begin VB.Label lblTitle 
      Caption         =   "Application Title"
      ForeColor       =   &H00000000&
      Height          =   240
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Label lblVersion 
      Caption         =   "Version"
      Height          =   225
      Left            =   1050
      TabIndex        =   4
      Top             =   600
      Width           =   3885
   End
End
Attribute VB_Name = "frmAbout"
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
                     
Private Sub cmdOK_Click()
  
  Unload Me

End Sub

Private Sub Form_Load()
    
    frmAbout.Caption = "About - " & PROGRAM_TITLE
    lblVersion.Caption = "Version - " & VERSION_MAJOR & "." & VERSION_MINOR & "." & VERSION_REVISION
    lblTitle.Caption = PROGRAM_NAME

    lblDescription.Caption = PROGRAM_TITLE & ", Copyright (C) 2012 Grahame White" & vbNewLine & vbNewLine
    lblDescription.Caption = lblDescription.Caption & "OSE comes with ABSOLUTELY NO WARRANTY;" & vbNewLine
    lblDescription.Caption = lblDescription.Caption & "for details see the LICENSE file" & vbNewLine & vbNewLine
    lblDescription.Caption = lblDescription.Caption & "This is free software, and you are welcome to redistribute it" & vbNewLine
    lblDescription.Caption = lblDescription.Caption & "under certain conditions; " & vbNewLine
    lblDescription.Caption = lblDescription.Caption & "for details see the LICENSE file"

End Sub
