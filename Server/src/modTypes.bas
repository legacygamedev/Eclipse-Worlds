Attribute VB_Name = "modTypes"
Option Explicit

' Public data structures
Public Map(1 To MAX_MAPS) As MapRec
Public TempEventMap(1 To MAX_MAPS) As GlobalEventsRec
Public MapCache(1 To MAX_MAPS) As Cache
Public PlayersOnMap(1 To MAX_MAPS) As Long
Public ResourceCache(1 To MAX_MAPS) As ResourceCacheRec
Public Account(1 To MAX_PLAYERS) As AccountRec
Public TempPlayer(1 To MAX_PLAYERS) As TempPlayerRec
Public TempChar(1 To MAX_GUILD_MEMBERS) As PlayerRec

Public Item(1 To MAX_ITEMS) As ItemRec
Public NPC(1 To MAX_NPCS) As NpcRec
Public MapItem(1 To MAX_MAPS, 1 To MAX_MAP_ITEMS) As MapItemRec
Public MapNpc(1 To MAX_MAPS) As MapDataRec
Public Shop(1 To MAX_SHOPS) As ShopRec
Public Spell(1 To MAX_SPELLS) As SpellRec
Public Resource(1 To MAX_RESOURCES) As ResourceRec
Public Animation(1 To MAX_ANIMATIONS) As AnimationRec
Public Guild(1 To MAX_GUILDS) As GuildRec
Public Party(1 To MAX_PARTYS) As PartyRec
Public Ban(1 To MAX_BANS) As BanRec
Public Title(1 To MAX_TITLES) As TitleRec
Public Moral(1 To MAX_MORALS) As MoralRec
Public Class(1 To MAX_CLASSES) As ClassRec
Public Emoticon(1 To MAX_EMOTICONS) As EmoticonRec

Public Switches(1 To MAX_SWITCHES) As String
Public Variables(1 To MAX_VARIABLES) As String
Public MapBlocks(1 To MAX_MAPS) As MapBlockRec

' Logs
Public Log As LogRec

' Options
Public Options As OptionsRec

Private Type MoveRouteRec
    Index As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    data5 As Long
    data6 As Long
End Type

Private Type GuildMemberRec
    Index As Long
    Access As Byte
End Type

Private Type GlobalEventRec
    X As Long
    Y As Long
    Dir As Long
    Active As Long
    
    WalkingAnim As Long
    FixedDir As Long
    WalkThrough As Long
    ShowName As Long
    
    Position As Long
    
    GraphicType As Long
    GraphicNum As Long
    GraphicX As Long
    GraphicX2 As Long
    GraphicY As Long
    GraphicY2 As Long
    
    ' Server only options
    MoveType As Long
    MoveSpeed As Long
    MoveFreq As Long
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    MoveRouteStep As Long
    
    RepeatMoveRoute As Long
    IgnoreIfCannotMove As Long
    
    MoveTimer As Long
End Type

Public Type GlobalEventsRec
    EventCount As Long
    Events() As GlobalEventRec
End Type

Private Type OptionsRec
    Name As String
    MOTD As String
    SMOTD As String
    Port As Long
    Website As String
    PKLevel As Byte
    MultipleSerial As Byte
    MultipleIP As Byte
    GuildCost As Long
    News As String
    MissSound As String
    DodgeSound As String
    DeflectSound As String
    BlockSound As String
    CriticalSound As String
    ResistSound As String
    BuySound As String
    SellSound As String
End Type

Public Type PartyRec
    Leader As Long
    Member(1 To MAX_PARTY_MEMBERS) As Long
    MemberCount As Long
End Type

Public Type PlayerItemRec
    Num As Byte
    Value As Long
    Durability As Integer
    Bind As Byte
End Type

Private Type Cache
    Data() As Byte
End Type

Private Type BankRec
    Item(1 To MAX_BANK) As PlayerItemRec
End Type

Public Type BuffRec
    id As Long
    Behavior As Long
    Vital As Long
    Timer As Long
End Type

Private Type BanRec
    Date As String * NAME_LENGTH
    Time As String * NAME_LENGTH
    PlayerName As String * NAME_LENGTH
    PlayerLogin As String * NAME_LENGTH
    IP As String * NAME_LENGTH
    HDSerial As String * NAME_LENGTH
    Reason As String * 100
    By As String * NAME_LENGTH
End Type

Public Type TitleRec
    Name As String * NAME_LENGTH
    Color As Byte
    LevelReq As Byte
    PKReq As Integer
    Desc As String * 100
End Type

Public Type MoralRec
    Name As String * NAME_LENGTH
    Color As Byte
    CanPK As Byte
    CanCast As Byte
    CanUseItem As Byte
    LoseExp As Byte
    DropItems As Byte
    CanPickupItem As Byte
    CanDropItem As Byte
    PlayerBlocked As Byte
End Type

Public Type HotbarRec
    Slot As Byte
    SType As Byte
End Type

Public Type SkillRec
    Level As Integer
    Exp As Long
    
    ' For future use...
    ObjectiveLevel As Integer
    ObjectiveExp As Long
End Type

Public Type FriendsRec
    AmountOfFriends As Byte
    Members(1 To MAX_PEOPLE) As String
End Type

Public Type FoesRec
    Amount As Byte
    Members(1 To MAX_PEOPLE) As String
End Type

Private Type PlayerRec
    ' Face
    Face As Integer
    
    Level As Integer
    Exp As Long
    
    ' Stats
    Stat(1 To Stats.Stat_count - 1) As Integer
    Points As Integer
    
    ' Spells
    Spell(1 To MAX_PLAYER_SPELLS) As Byte
    SpellCD(1 To MAX_PLAYER_SPELLS) As Long
    AmountOfCasts(1 To MAX_PLAYER_SPELLS) As Integer
    
    ' Tree
    CurrentCombatTree As Byte
    
    ' General
    Name As String * NAME_LENGTH
    Gender As Byte
    Class As Byte
    Sprite As Integer
    Access As Byte
    PK As Byte
    Status As String * NAME_LENGTH
    
    ' Position
    Map As Integer
    X As Byte
    Y As Byte
    Dir As Byte
    
    ' Vitals
    Vital(1 To Vitals.Vital_Count - 1) As Long
    
    ' Amount of titles
    AmountOfTitles As Byte
    
    ' Current Title
    CurrentTitle As Byte
    
    ' Titles
    Title(1 To MAX_TITLES) As Byte
    
    ' Worn equipment
    Equipment(1 To Equipment.Equipment_Count - 1) As PlayerItemRec
    
    ' Inventory
    Inv(1 To MAX_INV) As PlayerItemRec
    
    ' Buffs
    Buff(1 To MAX_BUFFS) As BuffRec
    
    ' Hotbar
    Hotbar(1 To MAX_HOTBAR) As HotbarRec
    
    ' Guild
    Guild As GuildMemberRec
    
    ' Skill
    Skills(1 To Skill_Count - 1) As SkillRec
    
    ' Events
    Switches(0 To MAX_SWITCHES) As Byte
    Variables(0 To MAX_VARIABLES) As Long

    ' Server use only
    CheckPointMap As Integer
    CheckPointX As Byte
    CheckPointY As Byte
    CanTrade As Boolean
    TempSprite As Integer
    PlayerKills As Integer
End Type

Type AccountRec
    ' Account
    Login As String * NAME_LENGTH
    Password As String * NAME_LENGTH
    
    ' Bank
    Bank As BankRec
    
    ' Friends
    Friends As FriendsRec
    
    ' Foes
    Foes As FoesRec

    CurrentChar As Byte
    
    ' Character
    Chars(1 To MAX_CHARS) As PlayerRec
End Type

Public Type DoTRec
    Used As Boolean
    Spell As Long
    Timer As Long
    Caster As Long
    StartTime As Long
End Type

Public Type SpellBufferRec
    Spell As Long
    Timer As Long
    Target As Long
    TType As Byte
End Type

Public Type ConditionalBranchRec
    Condition As Long
    Data1 As Long
    Data2 As Long
    Data3 As Long
    CommandList As Long
    ElseCommandList As Long
End Type

Private Type EventCommandRec
    Index As Byte
    Text1 As String
    Text2 As String
    Text3 As String
    Text4 As String
    Text5 As String
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As Long
    data5 As Long
    data6 As Long
    ConditionalBranch As ConditionalBranchRec
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
End Type

Private Type CommandListRec
    CommandCount As Long
    ParentList As Long
    Commands() As EventCommandRec
End Type

Private Type EventPageRec
    ' These are condition variables that decide if the event even appears to the player.
    chkVariable As Long
    VariableIndex As Long
    VariableCondition As Long
    VariableCompare As Long
    
    chkSwitch As Long
    SwitchIndex As Long
    SwitchCompare As Long
    
    chkHasItem As Long
    HasItemIndex As Long
    
    chkSelfSwitch As Long
    SelfSwitchIndex As Long
    SelfSwitchCompare As Long
    ' End Conditions
    
    ' Handles the Event Sprite
    GraphicType As Byte
    Graphic As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    
    ' Handles Movement - Move Routes to come soon.
    MoveType As Byte
    MoveSpeed As Byte
    MoveFreq As Byte
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    IgnoreMoveRoute As Long
    RepeatMoveRoute As Long
    
    ' Guidelines for the event
    WalkAnim As Long
    DirFix As Long
    WalkThrough As Long
    ShowName As Long
    
    ' Trigger for the event
    Trigger As Byte
    
    'Commands for the event
    CommandListCount As Long
    CommandList() As CommandListRec
    
    Position As Byte
    
    ' For EventMap
    X As Long
    Y As Long
End Type

Private Type EventRec
    Name As String
    Global As Byte
    PageCount As Long
    Pages() As EventPageRec
    X As Long
    Y As Long
    ' Self Switches re-set on restart.
    SelfSwitches(0 To 4) As Long
End Type

Public Type GlobalMapEvents
    eventID As Long
    pageID As Long
    X As Long
    Y As Long
End Type

Private Type MapEventRec
    Dir As Long
    X As Long
    Y As Long
    
    WalkingAnim As Long
    FixedDir As Long
    WalkThrough As Long
    ShowName As Long
    
    GraphicType As Long
    GraphicX As Long
    GraphicY As Long
    GraphicX2 As Long
    GraphicY2 As Long
    GraphicNum As Long
    
    movementspeed As Long
    Position As Long
    Visible As Long
    eventID As Long
    pageID As Long
    
    ' Server Only Options
    MoveType As Long
    MoveSpeed As Long
    MoveFreq As Long
    MoveRouteCount As Long
    MoveRoute() As MoveRouteRec
    MoveRouteStep As Long
    
    RepeatMoveRoute As Long
    IgnoreIfCannotMove As Long
    
    MoveTimer As Long
    SelfSwitches(0 To 4) As Long
End Type

Private Type EventMapRec
    CurrentEvents As Long
    EventPages() As MapEventRec
End Type

Private Type EventProcessingRec
    CurList As Long
    CurSlot As Long
    eventID As Long
    pageID As Long
    WaitingForResponse As Long
    ActionTimer As Long
    ListLeftOff() As Long
End Type

Public Type TempPlayerRec
    ' Non saved local vars
    buffer As clsBuffer
    HDSerial As String * NAME_LENGTH
    InGame As Boolean
    AttackTimer As Long
    DataTimer As Long
    DataBytes As Long
    DataPackets As Long
    TargetType As Byte
    Target As Byte
    PartyStarter As Byte
    GettingMap As Byte
    InShop As Long
    StunTimer As Long
    StunDuration As Long
    InBank As Boolean
    VitalCycle(1 To Vital_Count - 1) As Byte
    VitalPotion(1 To Vital_Count - 1) As Long
    VitalPotionTimer(1 To Vital_Count - 1) As Long
    
    ' Trade
    TradeRequest As Long
    InTrade As Long
    TradeOffer(1 To MAX_INV) As PlayerItemRec
    AcceptTrade As Boolean
    
    ' Regen
    StopRegen As Boolean
    StopRegenTimer As Long
    
    ' Dot/Hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    
    ' Spell Buffer
    SpellBuffer As SpellBufferRec
    
    ' Party
    InParty As Long
    PartyInvite As Long
    
    ' Guild
    GuildInvite As Long
    
    ' Events
    EventMap As EventMapRec
    EventProcessingCount As Long
    EventProcessing() As EventProcessingRec
End Type

Private Type TileDataRec
    X As Integer
    Y As Integer
    Tileset As Byte
End Type

Private Type TileRec
    Layer(1 To MapLayer.Layer_Count - 1) As TileDataRec
    Autotile(1 To MapLayer.Layer_Count - 1) As Byte
    Type As Byte
    Data1 As Long
    Data2 As Long
    Data3 As Long
    Data4 As String
    DirBlock As Byte
End Type

Private Type MapRec
    Name As String * NAME_LENGTH
    Music As String * FILE_LENGTH
    BGS As String * FILE_LENGTH
    
    Revision As Long
    Moral As Byte
    
    Up As Long
    Down As Long
    Left As Long
    Right As Long
    
    BootMap As Long
    BootX As Byte
    BootY As Byte
    
    Weather As Long
    WeatherIntensity As Long
    
    Fog As Long
    FogSpeed As Long
    FogOpacity As Long
    
    Panorama As Long
    
    Red As Long
    Green As Long
    Blue As Long
    Alpha As Long
    
    MaxX As Byte
    MaxY As Byte
    
    Npc_HighIndex As Byte
    
    Tile() As TileRec
    NPC(1 To MAX_MAP_NPCS) As Long
    NpcSpawnType(1 To MAX_MAP_NPCS) As Long
    EventCount As Long
    Events() As EventRec
End Type

Private Type ClassRec
    Name As String * NAME_LENGTH
    Stat(1 To Stats.Stat_count - 1) As Integer
    MaleSprite As Integer
    FemaleSprite As Integer
    
    StartItem(1 To MAX_INV) As Long
    StartItemValue(1 To MAX_INV) As Long
    StartSpell(1 To MAX_PLAYER_SPELLS) As Long

    Locked As Byte
    
    ' Faces
    MaleFace As Integer
    FemaleFace As Integer
    
    ' Color
    Color As Long
End Type

Private Type ItemRec
    Name As String * NAME_LENGTH
    Desc As String * 256
    Sound As String * FILE_LENGTH
    
    Pic As Integer
    Type As Byte
    
    EquipSlot As Byte
    
    Data1 As Integer
    Data2 As Integer
    Data3 As Integer
    
    ClassReq As Byte
    AccessReq As Byte
    LevelReq As Byte
    GenderReq As Byte
    ProficiencyReq As Byte
    
    Price As Long
    Add_Stat(1 To Stats.Stat_count - 1) As Integer
    Rarity As Byte
    WeaponSpeed As Long
    Handed As Long
    BindType As Byte
    Stat_Req(1 To Stats.Stat_count - 1) As Integer
    Animation As Long
    Paperdoll As Long
    AddHP As Long
    AddMP As Long
    AddEXP As Long
    CastSpell As Long
    InstaCast As Byte
    ChanceModifier As Byte
    IsReusable As Boolean
    Tool As Integer
    HoT As Byte
End Type

Private Type MapItemRec
    PlayerName As String * NAME_LENGTH
    Num As Byte
    Value As Long
    Durability As Integer
    X As Byte
    Y As Byte
    YOffset As Integer
    
    ' Ownership & Despawn
    PlayerTimer As Long
    CanDespawn As Boolean
    DespawnTimer As Long
End Type

Private Type NpcRec
    Name As String * NAME_LENGTH
    Title As String * NAME_LENGTH
    Music As String * FILE_LENGTH
    Sound As String * FILE_LENGTH

    Sprite As Integer
    SpawnSecs As Long
    Behavior As Byte
    Range As Byte
    DropChance(1 To MAX_NPC_DROPS) As Double
    DropItem(1 To MAX_NPC_DROPS) As Byte
    DropValue(1 To MAX_NPC_DROPS) As Integer
    Damage As Long
    Stat(1 To Stats.Stat_count - 1) As Integer
    HP As Long
    MP As Long
    Exp As Long
    Animation As Long
    Level As Byte
    Face As Long
    Spell(1 To MAX_NPC_SPELLS) As Integer
    Faction As Byte
    AttackSay As String * 100
    FactionThreat As Boolean
End Type

Private Type MapNpcRec
    Num As Byte
    Target As Byte
    TargetType As Byte
    Vital(1 To Vitals.Vital_Count - 1) As Long
    X As Byte
    Y As Byte
    Dir As Byte
    
    ' For server use only
    SpawnWait As Long
    AttackTimer As Long
    StunDuration As Long
    StunTimer As Long
    
    ' Regen
    StopRegen As Boolean
    StopRegenTimer As Long
    
    ' Dot/Hot
    DoT(1 To MAX_DOTS) As DoTRec
    HoT(1 To MAX_DOTS) As DoTRec
    
    ' Spells
    SpellTimer(1 To MAX_NPC_SPELLS) As Long
    SpellBuffer As SpellBufferRec

    ' Cache
    ActiveSpell As Integer
End Type

Private Type TradeItemRec
    Item As Long
    ItemValue As Long
    CostItem As Long
    CostValue As Long
    CostItem2 As Long
    CostValue2 As Long
End Type

Private Type ShopRec
    Name As String * NAME_LENGTH
    BuyRate As Integer
    SellRate As Integer
    TradeItem(1 To MAX_TRADES) As TradeItemRec
    CanFix As Byte
End Type

Private Type SpellRec
    Name As String * NAME_LENGTH
    Desc As String * 256
    Sound As String * FILE_LENGTH
    
    Type As Byte
    MPCost As Long
    LevelReq As Byte
    AccessReq As Byte
    ClassReq As Byte
    CastTime As Long
    CDTime As Long
    Icon As Long
    Map As Long
    X As Long
    Y As Long
    Dir As Byte
    Vital As Long
    Duration As Long
    Interval As Long
    Range As Byte
    AoE As Long
    IsAoe As Boolean
    CastAnim As Long
    SpellAnim As Long
    StunDuration As Long
    Sprite As Integer
    WeaponDamage As Boolean
    CastRequired As Integer
    NewSpell As Integer
End Type

Private Type GuildRec
    Name As String * NAME_LENGTH
    MOTD As String * 512
    Members(1 To MAX_GUILD_MEMBERS) As String
End Type

Private Type MapDataRec
    NPC() As MapNpcRec
End Type

Private Type MapResourceRec
    ResourceState As Byte
    ResourceTimer As Long
    X As Long
    Y As Long
    Cur_Reward As Byte
End Type

Private Type ResourceCacheRec
    Resource_Count As Long
    ResourceData() As MapResourceRec
End Type

Private Type ResourceRec
    Name As String * NAME_LENGTH
    SuccessMessage As String * NAME_LENGTH
    EmptyMessage As String * NAME_LENGTH
    FailMessage As String * NAME_LENGTH
    Sound As String * FILE_LENGTH
    
    Skill As Byte
    Exp As Integer
    ResourceImage As Byte
    ExhaustedImage As Byte
    ItemReward As Long
    ToolRequired As Long
    Reward_Min As Byte
    Reward_Max As Byte
    RespawnTime As Long
    Animation As Long
    LowChance As Byte
    HighChance As Byte
End Type

Private Type AnimationRec
    Name As String * NAME_LENGTH
    Sound As String * FILE_LENGTH
    Sprite(0 To 1) As Long
    Frames(0 To 1) As Long
    LoopCount(0 To 1) As Long
    LoopTime(0 To 1) As Long
End Type

Public Type Vector
    X As Long
    Y As Long
End Type

Public Type MapBlockRec
    Blocks() As Long
End Type

Private Type EmoticonRec
    Command As String * NAME_LENGTH
    Pic As Long
End Type

Private Type LogRec
    Msg As String * 512
    File As String * NAME_LENGTH
End Type
