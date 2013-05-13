Attribute VB_Name = "Constants"
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

' Application constants
Public Const PROGRAM_NAME As String = "OSE"

Public Const VERSION_MAJOR As Integer = 0
Public Const VERSION_MINOR As Integer = 1
Public Const VERSION_REVISION As Integer = 8
Public Const VERSION_EXTRA As String = ""

Public Const PROGRAM_TITLE As String = PROGRAM_NAME & " - " & _
                                       VERSION_MAJOR & "." & _
                                       VERSION_MINOR & "." & _
                                       VERSION_REVISION & _
                                       VERSION_EXTRA

' Bit values to help work out which flags are set
Public Const BIT_0 As Long = &H1&
Public Const BIT_1 As Long = &H2&
Public Const BIT_2 As Long = &H4&
Public Const BIT_3 As Long = &H8&
Public Const BIT_4 As Long = &H10&
Public Const BIT_5 As Long = &H20&
Public Const BIT_6 As Long = &H40&
Public Const BIT_7 As Long = &H80&
Public Const BIT_8 As Long = &H100&
Public Const BIT_9 As Long = &H200&
Public Const BIT_10 As Long = &H400&
Public Const BIT_11 As Long = &H800&
Public Const BIT_12 As Long = &H1000&
Public Const BIT_13 As Long = &H2000&
Public Const BIT_14 As Long = &H4000&
Public Const BIT_15 As Long = &H8000&
Public Const BIT_16 As Long = &H10000
Public Const BIT_17 As Long = &H20000
Public Const BIT_18 As Long = &H40000
Public Const BIT_19 As Long = &H80000
Public Const BIT_20 As Long = &H100000
Public Const BIT_21 As Long = &H200000
Public Const BIT_22 As Long = &H400000
Public Const BIT_23 As Long = &H800000
Public Const BIT_24 As Long = &H1000000
Public Const BIT_25 As Long = &H2000000
Public Const BIT_26 As Long = &H4000000
Public Const BIT_27 As Long = &H8000000
Public Const BIT_28 As Long = &H10000000
Public Const BIT_29 As Long = &H20000000
Public Const BIT_30 As Long = &H40000000
Public Const BIT_31 As Long = &H80000000

Public Const BYTE_2 As Long = &H100&
Public Const BYTE_3 As Long = &H10000
Public Const BYTE_4 As Long = &H1000000

' FileIDInvalid Return Values
Public Const FILE_ID_OK As Integer = 0 ' Valid FileID
Public Const FILE_ID_XBOX360 As Integer = 1 ' Xbox360 Container detected
Public Const FILE_ID_UNKNOWN As Integer = 2 ' Unknown error

' StatusBar Panels
Public Const STB_STATUS As Integer = 1
Public Const STB_PROGRESS As Integer = 2

' Min / Max values
Public Const BYTE_MIN As Byte = 0               ' Smallest value to fit in an unsigned byte
Public Const BYTE_MAX As Byte = 255             ' Largest value to fit in an unsigned byte

Public Const INTEGER_MIN As Integer = 0         ' Smallest value to fit into an unsigned integer
Public Const INTEGER_MAX As Integer = 32767     ' Largest vaule to fit into a signed integer

Public Const LONG_MIN As Long = 0               ' Smallest value to fit into an unsigned long
Public Const LONG_MAX As Long = 2147483647      ' Largest value to fit into a signed long

Public Const HEALTH_MIN As Integer = 0          ' Smallest base health value
Public Const HEALTH_MAX As Integer = 30000      ' Largest base health value (needs testing)

Public Const MAGICKA_MIN As Integer = 0         ' Smallest base health value
Public Const MAGICKA_MAX As Integer = 10000     ' Largest base health value (needs testing)

Public Const FATIGUE_MIN As Integer = 0         ' Smallest base health value
Public Const FATIGUE_MAX As Integer = 10000     ' Largest base health value (needs testing)

Public Const YEAR_MIN As Integer = 0            ' Smallest year value
Public Const YEAR_MAX As Integer = 30827        ' Largest year value

Public Const MONTH_MIN As Integer = 1           ' Smallest month value
Public Const MONTH_MAX As Integer = 12          ' Largest month value

Public Const DAY_OF_WEEK_MIN As Integer = 0     ' Smallest day of week value
Public Const DAY_OF_WEEK_MAX As Integer = 6     ' Largest day of week value

Public Const DAY_MIN As Integer = 1             ' Smallest day value
Public Const DAY_MAX As Integer = 31            ' Largest day value

Public Const HOUR_MIN As Integer = 0            ' Smallest hour value
Public Const HOUR_MAX As Integer = 23           ' Largest hour value

Public Const MINUTE_MIN As Integer = 0          ' Smallest minute value
Public Const MINUTE_MAX As Integer = 59         ' Largest minute value

Public Const SECOND_MIN As Integer = 0          ' Smallest second value
Public Const SECOND_MAX As Integer = 59         ' Largest second value

Public Const MILLISECOND_MIN As Integer = 0     ' Smallest millisecond value
Public Const MILLISECOND_MAX As Integer = 999   ' Largest millisecond value

' Tabstrip category constants
Public Const TAB_CAT_SAVE_FILE As String = "Save File"
Public Const TAB_CAT_PLAYER As String = "Player"

' Tabstrip sub-category constants
Public Const TAB_SUB_CAT_SAVE_FILE_ALL As String = "All"
Public Const TAB_SUB_CAT_PLAYER_ATTRIBUTES As String = "Attributes"
Public Const TAB_SUB_CAT_PLAYER_SKILLS As String = "Skills"
Public Const TAB_SUB_CAT_PLAYER_FACTIONS As String = "Factions"
Public Const TAB_SUB_CAT_PLAYER_SPELLS As String = "Spells"
Public Const TAB_SUB_CAT_PLAYER_ITEMS As String = "Items"

' Panel Positions
Public Const TAB_CAT_TOP As Integer = 480
Public Const TAB_CAT_LEFT As Integer = 240
Public Const TAB_SUB_CAT_TOP As Integer = 120
Public Const TAB_SUB_CAT_LEFT As Integer = 1320

' Date Stamp Modes
Public Const DATE_STAMP_MODE_EXE As Integer = 1
Public Const DATE_STAMP_MODE_GAME As Integer = 2

' Special Change Records
Public Const PLAYER_FORMID As Long = 7
Public Const PLAYER_CHANGE_FORMID As Long = 20

' Change Record Types
Public Const CHANGE_RECORD_NPC_ As Byte = 35
Public Const CHANGE_RECORD_ACHR As Byte = 50

' Created Data Types
Public Const CREATED_DATA_SPELL As String = "SPEL"

' Created Data Sub Records
Public Const SUB_RECORD_EDITOR_ID As String = "EDID"        ' EditorID
Public Const SUB_RECORD_FULL_NAME As String = "FULL"        ' Name
Public Const SUB_RECORD_SPELL_DATA As String = "SPIT"       ' Basic Spell Data

' Property Flags
Public Const PROPERTY_FLAG_WORLD_SPACE                  As Byte = &H11&
Public Const PROPERTY_FLAG_SCRIPT                       As Byte = &H12&
Public Const PROPERTY_FLAG_EQUIPPED_1                   As Byte = &H1B&
Public Const PROPERTY_FLAG_EQUIPPED_2                   As Byte = &H1C&
Public Const PROPERTY_FLAG_ITEM_HEALTH                  As Byte = &H2B&
Public Const PROPERTY_FLAG_CURRENT_ENCHANTMENT_POINTS   As Byte = &H2E&
Public Const PROPERTY_FLAG_SCALE                        As Byte = &H37&

' Property Variable Types
Public Const SCRIPT_VARIABLE_TYPE_IREF As Integer = &HF000
Public Const SCRIPT_VARIABLE_TYPE_LOCAL As Integer = &H0&

