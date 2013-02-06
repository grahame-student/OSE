VERSION 5.00
Begin VB.Form frmNewDateStamp 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Date Stamp"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   2445
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   2445
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   120
      TabIndex        =   17
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   16
      Top             =   3000
      Width           =   975
   End
   Begin VB.TextBox txtMillisecond 
      Height          =   285
      Left            =   1320
      TabIndex        =   15
      Text            =   "0"
      Top             =   2640
      Width           =   975
   End
   Begin VB.TextBox txtSecond 
      Height          =   285
      Left            =   1320
      TabIndex        =   14
      Text            =   "0"
      Top             =   2280
      Width           =   975
   End
   Begin VB.TextBox txtMinute 
      Height          =   285
      Left            =   1320
      TabIndex        =   13
      Text            =   "0"
      Top             =   1920
      Width           =   975
   End
   Begin VB.TextBox txtHour 
      Height          =   285
      Left            =   1320
      TabIndex        =   12
      Text            =   "0"
      Top             =   1560
      Width           =   975
   End
   Begin VB.TextBox txtDay 
      Height          =   285
      Left            =   1320
      TabIndex        =   11
      Text            =   "0"
      Top             =   1200
      Width           =   975
   End
   Begin VB.TextBox txtDayOfWeek 
      Height          =   285
      Left            =   1320
      TabIndex        =   10
      Text            =   "0"
      Top             =   840
      Width           =   975
   End
   Begin VB.TextBox txtMonth 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Text            =   "1"
      Top             =   480
      Width           =   975
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Text            =   "0"
      Top             =   120
      Width           =   975
   End
   Begin VB.Label lblMillisecond 
      AutoSize        =   -1  'True
      Caption         =   "Millisecond"
      Height          =   195
      Left            =   360
      TabIndex        =   7
      Top             =   2640
      Width           =   780
   End
   Begin VB.Label lblSecond 
      AutoSize        =   -1  'True
      Caption         =   "Second"
      Height          =   195
      Left            =   585
      TabIndex        =   6
      Top             =   2280
      Width           =   555
   End
   Begin VB.Label lblMinute 
      AutoSize        =   -1  'True
      Caption         =   "Minute"
      Height          =   195
      Left            =   660
      TabIndex        =   5
      Top             =   1920
      Width           =   480
   End
   Begin VB.Label lblHour 
      AutoSize        =   -1  'True
      Caption         =   "Hour"
      Height          =   195
      Left            =   795
      TabIndex        =   4
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label lblDay 
      AutoSize        =   -1  'True
      Caption         =   "Day"
      Height          =   195
      Left            =   855
      TabIndex        =   3
      Top             =   1200
      Width           =   285
   End
   Begin VB.Label lblDayOfWeek 
      AutoSize        =   -1  'True
      Caption         =   "Day of Week"
      Height          =   195
      Left            =   195
      TabIndex        =   2
      Top             =   840
      Width           =   945
   End
   Begin VB.Label lblMonth 
      AutoSize        =   -1  'True
      Caption         =   "Month"
      Height          =   195
      Left            =   690
      TabIndex        =   1
      Top             =   480
      Width           =   450
   End
   Begin VB.Label lblYear 
      AutoSize        =   -1  'True
      Caption         =   "Year"
      Height          =   195
      Left            =   810
      TabIndex        =   0
      Top             =   120
      Width           =   330
   End
End
Attribute VB_Name = "frmNewDateStamp"
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

Public Mode As Integer

Private Sub Form_Load()

    Select Case Mode
        Case DATE_STAMP_MODE_EXE
            UpdateDisplay SaveFileData.FileHeader.EXETime
        Case DATE_STAMP_MODE_GAME
            UpdateDisplay SaveFileData.SaveHeader.GameTime
    End Select

End Sub

Private Sub UpdateDisplay(ByRef DateStamp As SystemTime)

    txtYear.Text = DateStamp.Year
    txtMonth.Text = DateStamp.Month
    txtDayOfWeek.Text = DateStamp.DayOfWeek
    txtDay.Text = DateStamp.Day
    txtHour.Text = DateStamp.Hour
    txtMinute.Text = DateStamp.Minute
    txtSecond.Text = DateStamp.Second
    txtMillisecond.Text = DateStamp.MilliSecond

End Sub

Private Sub txtYear_Change()

    ValidateInput txtYear, YEAR_MIN, YEAR_MAX

End Sub

Private Sub txtMonth_Change()

    ValidateInput txtMonth, MONTH_MIN, MONTH_MAX

End Sub

Private Sub txtDayOfWeek_Change()

    ValidateInput txtDayOfWeek, DAY_OF_WEEK_MIN, DAY_OF_WEEK_MAX

End Sub

Private Sub txtDay_Change()

    ValidateInput txtDay, DAY_MIN, DAY_MAX
    
End Sub

Private Sub txtHour_Change()

    ValidateInput txtHour, HOUR_MIN, HOUR_MAX

End Sub

Private Sub txtMinute_Change()

    ValidateInput txtMinute, MINUTE_MIN, MINUTE_MAX

End Sub

Private Sub txtSecond_Change()

    ValidateInput txtSecond, SECOND_MIN, SECOND_MAX

End Sub

Private Sub txtMillisecond_Change()

    ValidateInput txtMillisecond, MILLISECOND_MIN, MILLISECOND_MAX

End Sub

Private Sub cmdCancel_Click()

    frmNewDateStamp.Hide

End Sub

Private Sub cmdOK_Click()

    ' Fix the minimum year if it's too low
    If txtYear.Text < 1601 Then
        txtYear.Text = 1601
    End If
    
    ' Fix any problems in the day number
    CheckDaysInMonth
    
    ' Set new date
    Select Case Mode
        Case DATE_STAMP_MODE_EXE
            ModifySaveFileEXETime frmNewDateStamp
        Case DATE_STAMP_MODE_GAME
            ModifySaveFileGameTime frmNewDateStamp
    End Select
    
    frmNewDateStamp.Hide

End Sub

Private Sub CheckDaysInMonth()

    Select Case txtMonth.Text
        Case 4, 6, 9, 11
            If txtDay.Text > 30 Then
                txtDay.Text = 30
            End If
        Case 2
            If IsLeapYear Then
                If txtDay.Text > 29 Then
                    txtDay.Text = 29
                End If
            Else
                If txtDay.Text > 28 Then
                    txtDay.Text = 28
                End If
            End If
    End Select
        

End Sub

Private Function IsLeapYear() As Boolean

    If txtYear.Text Mod 4 = 0 And txtYear.Text Mod 100 <> 0 Then
        IsLeapYear = True
    ElseIf (txtYear.Text / 100) Mod 4 = 0 Then
        IsLeapYear = True
    Else
        IsLeapYear = False
    End If

End Function
