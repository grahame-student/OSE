Attribute VB_Name = "OSE_ESS"
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
    IRef As Long
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

Private Type ChangeRecord
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


' These blocks hold additional information to make it easier to navigate
' various records. DO NOT WRITE THEM TO THE NEW SAVE FILE
Private Type Player
    PlayerRecord As Long        ' The change record containing the player data (-1 for not present)
    FormFlags As Long           ' Offset of the start of the FormFlags within the data block (-1 for not present)
    BaseAttributes As Long      ' Offset of the start of the BaseAttributes within the data block (-1 for not present)
    BaseData As Long            ' Offset of the start of the BaseData within the data block (-1 for not present)
    Factions As Long            ' Offset of the start of the Factions within the data block (-1 for not present)
    SpellList As Long           ' Offset of the start of the SpellList within the data block (-1 for not present)
    AI As Long                  ' Offset of the start of the AI within the data block (-1 for not present)
    BaseHealth As Long          ' Offset of the start of the BaseHealth within the data block (-1 for not present)
    BaseModifiers As Long       ' Offset of the start of the BaseModifiers within the data block (-1 for not present)
    FullName As Long            ' Offset of the start of the FullName within the data block (-1 for not present)
    Skills As Long              ' Offset of the start of the Skills within the data block (-1 for not present)
    CombatStyle As Long         ' Offset of the start of the CombatStyle within the data block (-1 for not present)
End Type

Private Type OSEExtra
    Player As Player
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

