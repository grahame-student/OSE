Attribute VB_Name = "OSE_ESS"
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

' The data structure for the ESS file format used in Elder Scrolls IV: Oblivion
' is defined here.
' Based very heavily on the documentation on http://www.uesp.net/wiki/Tes4Mod:Save_File_Format

''' Sub-structures located within the Savefile '''
Private Type FileHeaderStructure
    FileID As String
    MajorVersion As Byte
    MinorVersion As Byte
    EXETime As SystemTime
End Type

Private Type SaveHeaderStructure
    HeaderVersion As Long
    SaveHeaderSize As Long
    SaveNumber As Long
    PlayerName As String
    PlayerLevel As Integer
    PlayerLocation As String
    GameDays As Single
    GameTicks As Long
    GameTime As SystemTime
    ScreenShot As ScreenShot
End Type

Private Type PlugIns
    NumberOfPlugins As Byte
    PlugInNames() As String
End Type

Private Type PlayerLocation
    Cell As Long
    X As Single
    Y As Single
    Z As Single
End Type

Private Type GlobalsStructure
    iRef As Long
    Value As Single
End Type

Private Type DeathCount
    Actor As Long
    Count As Integer
End Type

Private Type CreatedData
    Type As String
    Size As Long
    Flags As Long
    FormID As Long
    VersionControlInfo As Long
    Data() As Byte
End Type

Private Type QuickKeyData
    Flag As Byte
    Reference As Long
End Type

Private Type Region
    Reference As Long
    Unknown As Long
End Type

Private Type GlobalStructure
    FormIDOffset As Long
    NumberOfChangeRecords As Long
    NextObjectID As Long
    WorldID As Long
    WorldX As Long
    WorldY As Long
    PlayerLocation As PlayerLocation
    GlobalsNumber As Integer
    Globals() As GlobalsStructure
    ClassSize As Integer
    NumberOfDeathCounts As Long
    DeathCounts() As DeathCount
    GameModeSeconds As Single
    ProcessesSize As Integer
    ProcessesData() As Byte
    SpectatorEventSize As Integer
    SpectatorEventData() As Byte
    WeatherSize As Integer
    WeatherData() As Byte
    PlayerCombatCount As Long
    CreatedNumber As Long
    CreatedData() As CreatedData
    QuickKeySize As Integer
    QuickKeyData() As QuickKeyData
    ReticuleSize As Integer
    ReticuleData() As Byte
    InterfaceSize As Integer
    InterfaceData() As Byte
    RegionSize As Integer
    RegionNumber As Integer
    Regions() As Region
End Type

Public Type ChangeRecord
    FormID As Long
    Type As Byte
    Flags As Long
    Version As Byte
    DataSize As Integer
    Data() As Byte
End Type

Private Type TemporaryEffects
    Size As Long
    Data() As Byte
End Type

Private Type FormIDs
    NumberOfFormIDs As Long
    FormIDsList() As Long
End Type

Private Type WorldSpaces
    NumberOfWorldSpaces As Long
    WorldSpaces() As Long
End Type

Private Type PropertyScriptVariable
    Index           As New SuperInt
    Type            As New SuperInt
    RefVariable     As New SuperLong                ' Only if Type = 0xF000
    LocalVariable   As New SuperDouble              ' Only if Type = 0x0000
End Type

Private Type PropertyScript
    ScriptRef       As New SuperLong                ' iRef
    VariableCount   As New SuperInt
    VariableList()  As PropertyScriptVariable
    Unknown         As Byte
End Type

Private Type PropertyMarkerHeadingRef
    Cell As Long
    X As Single
    Y As Single
    Z As Single
    Flags As Long
End Type

Private Type PropertyAllPack
    Package As Long
    Flags As Long
    Package2 As Long
    Unknown As Integer
End Type

Private Type PropertyUnknown1Data
    iRef As Long
    Unknown As Byte
End Type

Private Type PropertyUnknown1
    BlockCount As Integer
    BlockData() As PropertyUnknown1Data
End Type

Private Type PropertyUnknown2
    BlockCount As Integer
    Blocks() As Long                                ' iRefs
End Type

Private Type PropertyLock
    LockLevel As Byte
    Key As Long
    Flag As Byte
End Type

Private Type PropertyTeleport
    X As Single
    Y As Single
    Z As Single
    RX As Single
    RY As Single
    RZ As Single
    DestinationDoor As Long
End Type

Private Type PropertyUnknown5
    iRef As Long
    BlockCount As Integer
    Block(60) As Byte
End Type

Private Type PropertyOblivionEntry
    Door As Long
    X As Single
    Y As Single
    Z As Single
End Type

Private Type PropertyMovementExtra
    Unknown As Long
    BlockCount As Single
    Block() As Byte
    Blank() As Byte
End Type

Private Type PropertyUnknown7
    BlockCount As Integer
    Block(9) As Byte
End Type

Private Type ConversationBlock
    Number As Byte
    Quest As Long                                   ' iRef
    Dialog As Long                                  ' iRef
    Info As Long                                    ' iRef
End Type

Private Type PropertyConversation
    Topic As String
    ConversationCount As Integer
End Type

Private Type PropertyChange
    Flag As Byte                                    ' The property flag
    WorldspaceiRef As Long                          ' Used for flag 0x11
    Script As PropertyScript                        ' Used for flag 0x12
    MarkerHeadingRef As PropertyMarkerHeadingRef    ' Used for flag 0x1E
    AllPack As PropertyAllPack                      ' Used for flag 0x1F
    Trespass(62) As Byte                            ' Used for flag 0x20
    Unknown1 As PropertyUnknown1                    ' Used for flag 0x21
    UnknowniRef As Long                             ' Used for flag 0x22
    Unknown2 As PropertyUnknown2                    ' Used for flag 0x23
    Owner As Long                                   ' Used for flag 0x27
    GlobalVariable As Long                          ' Used for flag 0x28
    FactionRank As Long                             ' Used for flag 0x29
    AffectedItemsNumber As Integer                  ' Used for flag 0x2A
    ItemHealth          As New SuperSingle          ' Used for flag 0x2B
    Time As Single                                  ' Used for flag 0x2D
    EnchantmentPoints   As New SuperSingle          ' Used for flag 0x2E
    Soul As Byte                                    ' Used for flag 0x2F
    Lock As PropertyLock                            ' Used for flag 0x31
    Teleport As PropertyTeleport                    ' Used for flag 0x32
    MapMarkerFlag As Byte                           ' Used for flag 0x33
    Unknown3(4) As Byte                             ' Used for flag 0x36
    ScaleValue          As New SuperSingle          ' Used for flag 0x37
    Unknown4(11) As Byte                            ' Used for flag 0x39
    Unknown5 As PropertyUnknown5                    ' Used for flag 0x3A
    CrimeGold As Single                             ' Used for flag 0x3D
    OblivionEntry As PropertyOblivionEntry          ' Used for flag 0x3E
    Unknown6 As Single                              ' Used for flag 0x41
    Poison As Long                                  ' Used for flag 0x48
    Animation As String                             ' Used for flag 0x4A
    MovementExtra As PropertyMovementExtra          ' Used for flag 0x4B
    Unknown7 As PropertyUnknown7                    ' Used for flag 0x4E
    Unknown8(3) As Byte                             ' Used for flag 0x4F
    InvestmentGold As Long                          ' Used for flag 0x520
    Unknown9 As Long                                ' Used for flag 0x53
    ShortcutKey As Byte                             ' Used for flag 0x55
    Conversation As PropertyConversation            ' Used for flag 0x59
    Essential As Byte                               ' Used for flag 0x5A
    Unknown10 As Single                             ' Used for flag 0x5C
End Type

Private Type InventoryChangeEntry
    PropertiesNumber As New SuperInt
    Properties() As PropertyChange
End Type

Private Type InventoryEntry
    iRef As Long
    StackedItemsCount As Long
    ChangedEntriesCount As Long
    InventoryChangedEntries() As InventoryChangeEntry
End Type


' These blocks hold additional information to make it easier to navigate
' various records. DO NOT WRITE THEM DIRECTLY TO THE NEW SAVE FILE
Public Type Faction
    iRef        As New SuperLong
    FormID      As Long
    Level       As Byte
    Name        As String
    MaxRank     As Integer
    Ranks()     As String
    Suspended   As Boolean
End Type

Public Type Spell
    iRef        As New SuperLong
    FormID      As Long
    Name        As String
End Type

Public Type Item
    iRef As New SuperLong                               ' the iRef of the item
    FormID As Long                                      ' The item's FormID
    Name As String                                      ' The item name
    Size As Integer                                     ' Size of the item entry in bytes
    StackedItemsCount As New SuperLong                  ' Number of items in the stack
    ChangedEntriesCount As New SuperLong                ' Number of change entries
    InventoryChangedEntries() As InventoryChangeEntry   ' The actual change entries
End Type

Private Type BaseMod
    Index As Byte
    ModValue As Single
End Type

' SPIT
Private Type BasicSpellData
    Type As Long
    Cost As Long
    Level As Long
    Flags As Long
End Type

' EFIT
Private Type EffectData
    EffectID As String
    Magnitude As Long
    Area As Long
    Duration As Long
    Type As Long
    ActorValue As Long
End Type

' SCIT
Private Type ScriptEffectData
    FormID As Long
    School As Long
    VisualEffect As String
    Flags As Long
End Type

Private Type RegularEffect
    EffectID As String          ' EFID
    EffectData As EffectData    ' EFIT
End Type

Private Type ScriptEffect
    EffectID As String
    EffectData As EffectData
    ScriptEffectData As ScriptEffectData
    Name As String
End Type

Private Type CreatedItem
    NameLength As Integer
    Name As String
    BasicSpellData As BasicSpellData
    RegularEffects() As RegularEffect
    ScriptEffects() As ScriptEffect
End Type

Private Type CreatedItems
    SpellRecords() As Long          ' The change records containing spells (-1 in element 0 for not present)
    Spells() As CreatedItem
End Type

Private Type Player
    PlayerRecord As Long            ' The change record containing the player data (-1 for not present)
    FormFlags As Long               ' Offset of the start of the FormFlags within the data block (-1 for not present)
    BaseAttributes As Long          ' Offset of the start of the BaseAttributes within the data block (-1 for not present)
    BaseData As Long                ' Offset of the start of the BaseData within the data block (-1 for not present)
    Factions As Long                ' Offset of the start of the Factions within the data block (-1 for not present)
    Spells As Long                  ' Offset of the start of the SpellList within the data block (-1 for not present)
    AI As Long                      ' Offset of the start of the AI within the data block (-1 for not present)
    BaseHealth As Long              ' Offset of the start of the BaseHealth within the data block (-1 for not present)
    BaseModifiers As Long           ' Offset of the start of the BaseModifiers within the data block (-1 for not present)
    FullName As Long                ' Offset of the start of the FullName within the data block (-1 for not present)
    Skills As Long                  ' Offset of the start of the Skills within the data block (-1 for not present)
    CombatStyle As Long             ' Offset of the start of the CombatStyle within the data block (-1 for not present)
    FactionCount As Integer         ' Number of factions the player is in
    FactionList() As Faction        ' The list of factions the player is in
    SpellCount As Integer           ' Number of spells the player has
    SpellList() As Spell            ' The list of spells the player has
    BaseModCount As Integer         ' Number of BaseMods
    BaseModList() As BaseMod        ' List of the BaseMods
    FullNameString As String        ' The FullName
    ItemCount As Integer            ' Number of items the player has
    ItemList() As Item              ' The list of the items the player has
End Type

Private Type PlayerChange
    PlayerChangeRecord As Long      ' The change record containing the player change record (-1 for not present)
    CellChanged As Long             ' Offset of the start of the CellChanged within the data block (-1 for not present)
    Created As Long                 ' Offset of the start of the Created within the data block (-1 for not present)
    Moved As Long                   ' Offset of the start of the Moved within the data block (-1 for not present)
    HavokMoved As Long              ' Offset of the start of the HavokMoved within the data block (-1 for not present)
    OblivionFlag As Long            ' Offset of the start of the OblivionFlag within the data block (-1 for not present)
    TempAttributeChanges As Long    ' Offset of the start of the TempAttributeChanges within the data block (-1 for not present)
    ActorFlag As Long               ' Offset of the start of the ActorFlag within the data block (-1 for not present)
    FormFlags As Long               ' Offset of the start of the FormFlags within the data block (-1 for not present)
    Inventory As Long               ' Offset of the start of the Inventory within the data block (-1 for not present)
    InventoryCount As Integer       ' Number of things in the inventory sub record
End Type

Private Type OSEExtra
    LoadSuccessful As Boolean
    ScreenShotLoaded As Boolean
    Player As Player
    PlayerChange As PlayerChange
    CustomItems As CreatedItems
End Type


''' The main Savefile data structure '''
Public Type ESS
    FileHeader As FileHeaderStructure
    SaveHeader As SaveHeaderStructure
    PlugIns As PlugIns
    Globals As GlobalStructure
    ChangeRecords() As ChangeRecord
    TempEffects As TemporaryEffects
    FormIDs As FormIDs
    WorldSpaces As WorldSpaces
    OSE As OSEExtra
End Type


