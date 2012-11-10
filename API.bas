Attribute VB_Name = "API"
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

' Constants used in the SendMessage API
Public Const EM_CANUNDO = &HC6
Public Const EM_UNDO = &HC7
Public Const EM_EMPTYUNDOBUFFER = &HCD

' Constants for use with reparenting progress bars
Public Const SM_CXBORDER = 5     ' Width of non-sizable borders
Public Const SM_CYBORDER = 6     ' Height of non-sizable borders

' API to re-parent a control
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long

' API to aquire system metrics
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

' SendMessage API (ANSI entrypoint) used for the Undo functions
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
    ByVal hwnd As Long, _
    ByVal wMsg As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

