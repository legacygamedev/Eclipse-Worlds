Attribute VB_Name = "modCommands"
Option Explicit

Function GetPlayerLogin(ByVal Index As Long) As String

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerLogin = Trim$(Account(Index).Login)
End Function

Sub SetPlayerLogin(ByVal Index As Long, ByVal Login As String)
    Account(Index).Login = Login
End Sub

Function GetPlayerPassword(ByVal Index As Long) As String

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerPassword = Trim$(Account(Index).Password)
End Function

Sub SetPlayerPassword(ByVal Index As Long, ByVal Password As String)
    Account(Index).Password = Password
End Sub

Function GetPlayerName(ByVal Index As Long) As String

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerName = Trim$(Account(Index).Chars(GetPlayerChar(Index)).Name)
End Function

Sub SetPlayerName(ByVal Index As Long, ByVal Name As String)
    Account(Index).Chars(GetPlayerChar(Index)).Name = Name
End Sub

Function GetPlayerChar(ByVal Index As Byte) As Byte

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerChar = Account(Index).CurrentChar
End Function

Sub SetPlayerChar(ByVal Index As Long, ByVal Char As Byte)
    Account(Index).CurrentChar = Index
End Sub

Function GetPlayerGuildName(ByVal Index As Long) As String
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerGuildName = Trim$(Guild(Account(Index).Chars(GetPlayerChar(Index)).Guild.Index).Name)
End Function

Function GetPlayerGuild(ByVal Index As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerGuild = Account(Index).Chars(GetPlayerChar(Index)).Guild.Index
End Function

Sub SetPlayerGuild(ByVal Index As Long, ByVal GuildNum As Long)
    Account(Index).Chars(GetPlayerChar(Index)).Guild.Index = GuildNum
End Sub

Function GetPlayerGuildAccess(ByVal Index As Long) As Byte

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerGuildAccess = Account(Index).Chars(GetPlayerChar(Index)).Guild.Access
End Function

Sub SetPlayerGuildAccess(ByVal Index As Long, ByVal Access As Byte)
    Account(Index).Chars(GetPlayerChar(Index)).Guild.Access = Access
End Sub

Function GetPlayerClass(ByVal Index As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerClass = Account(Index).Chars(GetPlayerChar(Index)).Class
End Function

Sub SetPlayerClass(ByVal Index As Long, ByVal ClassNum As Long)
    Account(Index).Chars(GetPlayerChar(Index)).Class = ClassNum
End Sub

Function GetPlayerSprite(ByVal Index As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerSprite = Account(Index).Chars(GetPlayerChar(Index)).Sprite
End Function

Sub SetPlayerSprite(ByVal Index As Long, ByVal Sprite As Long)
    Account(Index).Chars(GetPlayerChar(Index)).Sprite = Sprite
End Sub

Function GetPlayerTitle(ByVal Index As Long, ByVal TitleNum As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerTitle = Account(Index).Chars(GetPlayerChar(Index)).Title(TitleNum)
End Function

Sub SetPlayerTitle(ByVal Index As Long, ByVal Title As Long, ByVal TitleNum As Long)
    Account(Index).Chars(GetPlayerChar(Index)).Title(Title) = TitleNum
End Sub

Function GetPlayerFace(ByVal Index As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerFace = Account(Index).Chars(GetPlayerChar(Index)).Face
End Function

Sub SetPlayerFace(ByVal Index As Long, ByVal Face As Long)
    Account(Index).Chars(GetPlayerChar(Index)).Face = Face
End Sub

Function GetPlayerLevel(ByVal Index As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerLevel = Account(Index).Chars(GetPlayerChar(Index)).Level
End Function

Function GetPlayerSkill(ByVal Index As Long, ByVal SkillNum As Byte) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerSkill = Account(Index).Chars(GetPlayerChar(Index)).Skills(SkillNum).Level
End Function

Sub SetPlayerLevel(ByVal Index As Long, ByVal Level As Byte)

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Account(Index).Chars(GetPlayerChar(Index)).Level = Level
End Sub

Sub SetPlayerSkill(ByVal Index As Long, ByVal Level As Byte, ByVal SkillNum As Byte)

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Account(Index).Chars(GetPlayerChar(Index)).Skills(SkillNum).Level = Level
End Sub

Function GetPlayerNextLevel(ByVal Index As Long) As Long
    GetPlayerNextLevel = (50 / 3) * ((GetPlayerLevel(Index) + 1) ^ 3 - (6 * (GetPlayerLevel(Index) + 1) ^ 2) + 17 * (GetPlayerLevel(Index) + 1) - 12)
End Function

Function GetPlayerNextSkillLevel(ByVal Index As Long, ByVal SkillNum As Byte) As Long
    GetPlayerNextSkillLevel = (50 / 3) * ((GetPlayerSkill(Index, SkillNum) + 1) ^ 3 - (6 * (GetPlayerSkill(Index, SkillNum) + 1) ^ 2) + 17 * (GetPlayerSkill(Index, SkillNum) + 1) - 12)
End Function

Function GetPlayerExp(ByVal Index As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerExp = Account(Index).Chars(GetPlayerChar(Index)).Exp
End Function

Function GetPlayerSkillExp(ByVal Index As Long, ByVal SkillNum As Byte) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerSkillExp = Account(Index).Chars(GetPlayerChar(Index)).Skills(SkillNum).Exp
End Function

Sub SetPlayerExp(ByVal Index As Long, ByVal Exp As Long)
    Account(Index).Chars(GetPlayerChar(Index)).Exp = Exp
End Sub

Sub SetPlayerSkillExp(ByVal Index As Long, ByVal Exp As Long, ByVal SkillNum As Byte)
    Account(Index).Chars(GetPlayerChar(Index)).Skills(SkillNum).Exp = Exp
End Sub

Function GetPlayerAccess(ByVal Index As Long) As Byte

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerAccess = Account(Index).Chars(GetPlayerChar(Index)).Access
End Function

Sub SetPlayerAccess(ByVal Index As Long, ByVal Access As Byte)
    Account(Index).Chars(GetPlayerChar(Index)).Access = Access
End Sub

Function GetPlayerPK(ByVal Index As Long) As Byte

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerPK = Account(Index).Chars(GetPlayerChar(Index)).PK
End Function

Sub SetPlayerPK(ByVal Index As Long, ByVal PK As Byte)
    Account(Index).Chars(GetPlayerChar(Index)).PK = PK
End Sub

Function GetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerVital = Account(Index).Chars(GetPlayerChar(Index)).Vital(Vital)
End Function

Sub SetPlayerVital(ByVal Index As Long, ByVal Vital As Vitals, ByVal Value As Long)
    Account(Index).Chars(GetPlayerChar(Index)).Vital(Vital) = Value

    If GetPlayerVital(Index, Vital) > GetPlayerMaxVital(Index, Vital) Then
        Account(Index).Chars(GetPlayerChar(Index)).Vital(Vital) = GetPlayerMaxVital(Index, Vital)
    End If

    If GetPlayerVital(Index, Vital) < 0 Then
        Account(Index).Chars(GetPlayerChar(Index)).Vital(Vital) = 0
    End If
End Sub

Function GetPlayerStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    Dim X As Long, i As Long
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    
    X = Account(Index).Chars(GetPlayerChar(Index)).Stat(Stat)
    
    For i = 1 To Equipment.Equipment_Count - 1
        If Account(Index).Chars(GetPlayerChar(Index)).Equipment(i).Num > 0 Then
            If Item(Account(Index).Chars(GetPlayerChar(Index)).Equipment(i).Num).Add_Stat(Stat) > 0 Then
                X = X + Item(Account(Index).Chars(GetPlayerChar(Index)).Equipment(i).Num).Add_Stat(Stat)
            End If
        End If
    Next
    
    GetPlayerStat = X
End Function

Function GetPlayerRawStat(ByVal Index As Long, ByVal Stat As Stats) As Long
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerRawStat = Account(Index).Chars(GetPlayerChar(Index)).Stat(Stat)
End Function

Sub SetPlayerStat(ByVal Index As Long, ByVal Stat As Stats, ByVal Value As Long)
    Account(Index).Chars(GetPlayerChar(Index)).Stat(Stat) = Value
End Sub

Function GetPlayerPoints(ByVal Index As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerPoints = Account(Index).Chars(GetPlayerChar(Index)).Points
End Function

Sub SetPlayerPoints(ByVal Index As Long, ByVal Points As Integer)
    Account(Index).Chars(GetPlayerChar(Index)).Points = Points
End Sub

Function GetPlayerMap(ByVal Index As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerMap = Account(Index).Chars(GetPlayerChar(Index)).Map
End Function

Sub SetPlayerMap(ByVal Index As Long, ByVal MapNum As Integer)

    If MapNum > 0 And MapNum <= MAX_MAPS Then
        Account(Index).Chars(GetPlayerChar(Index)).Map = MapNum
    End If
End Sub

Function GetPlayerX(ByVal Index As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerX = Account(Index).Chars(GetPlayerChar(Index)).X
End Function

Sub SetPlayerX(ByVal Index As Long, ByVal X As Long)
    
    Account(Index).Chars(GetPlayerChar(Index)).X = X
End Sub

Function GetPlayerY(ByVal Index As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerY = Account(Index).Chars(GetPlayerChar(Index)).Y
End Function

Sub SetPlayerY(ByVal Index As Long, ByVal Y As Long)
    Account(Index).Chars(GetPlayerChar(Index)).Y = Y
End Sub

Function GetPlayerDir(ByVal Index As Long) As Long
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerDir = Account(Index).Chars(GetPlayerChar(Index)).Dir
End Function

Sub SetPlayerDir(ByVal Index As Long, ByVal Dir As Byte)
    Account(Index).Chars(GetPlayerChar(Index)).Dir = Dir
End Sub

Function GetPlayerIP(ByVal Index As Long) As String

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerIP = frmServer.Socket(Index).RemoteHostIP
End Function

Function GetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Byte) As Long
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    If InvSlot = 0 Then Exit Function
    
    GetPlayerInvItemNum = Account(Index).Chars(GetPlayerChar(Index)).Inv(InvSlot).Num
End Function

Sub SetPlayerInvItemNum(ByVal Index As Long, ByVal InvSlot As Byte, ByVal ItemNum As Integer)
    Account(Index).Chars(GetPlayerChar(Index)).Inv(InvSlot).Num = ItemNum
End Sub

Function GetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Byte) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemValue = Account(Index).Chars(GetPlayerChar(Index)).Inv(InvSlot).Value
End Function

Sub SetPlayerInvItemValue(ByVal Index As Long, ByVal InvSlot As Byte, ByVal ItemValue As Long)
    Account(Index).Chars(GetPlayerChar(Index)).Inv(InvSlot).Value = ItemValue
End Sub

Function GetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Byte) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerSpell = Account(Index).Chars(GetPlayerChar(Index)).Spell(SpellSlot)
End Function

Sub SetPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Byte, ByVal SpellNum As Long)
    Account(Index).Chars(GetPlayerChar(Index)).Spell(SpellSlot) = SpellNum
End Sub

Function GetPlayerSpellCD(ByVal Index As Long, ByVal SpellSlot As Byte) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerSpellCD = Account(Index).Chars(GetPlayerChar(Index)).SpellCD(SpellSlot)
End Function

Sub SetPlayerSpellCD(ByVal Index As Long, ByVal SpellSlot As Byte, ByVal NewCD As Long)
    Account(Index).Chars(GetPlayerChar(Index)).SpellCD(SpellSlot) = NewCD
End Sub

Function GetPlayerEquipment(ByVal Index As Long, ByVal EquipmentSlot As Byte) As Byte

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipment = Account(Index).Chars(GetPlayerChar(Index)).Equipment(EquipmentSlot).Num
End Function

Sub SetPlayerEquipment(ByVal Index As Long, ByVal InvNum As Byte, ByVal EquipmentSlot As Byte)
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Account(Index).Chars(GetPlayerChar(Index)).Equipment(EquipmentSlot).Num = InvNum
End Sub

Function GetPlayerEquipmentDur(ByVal Index As Long, ByVal EquipmentSlot As Byte) As Byte

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentDur = Account(Index).Chars(GetPlayerChar(Index)).Equipment(EquipmentSlot).Durability
End Function

Sub SetPlayerEquipmentDur(ByVal Index As Long, ByVal DurValue As Integer, ByVal EquipmentSlot As Byte)
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Account(Index).Chars(GetPlayerChar(Index)).Equipment(EquipmentSlot).Durability = DurValue
End Sub

Function GetPlayerEquipmentBind(ByVal Index As Long, ByVal EquipmentSlot As Byte) As Byte

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    If EquipmentSlot = 0 Then Exit Function
    GetPlayerEquipmentBind = Account(Index).Chars(GetPlayerChar(Index)).Equipment(EquipmentSlot).Bind
End Function

Sub SetPlayerEquipmentBind(ByVal Index As Long, ByVal BindType As Byte, ByVal EquipmentSlot As Byte)
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    Account(Index).Chars(GetPlayerChar(Index)).Equipment(EquipmentSlot).Bind = BindType
End Sub

Function GetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Byte) As Integer
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerBankItemNum = Account(Index).Bank.Item(BankSlot).Num
End Function

Sub SetPlayerBankItemNum(ByVal Index As Long, ByVal BankSlot As Byte, ByVal ItemNum As Integer)
    Account(Index).Bank.Item(BankSlot).Num = ItemNum
End Sub

Function GetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Byte) As Long
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerBankItemValue = Account(Index).Bank.Item(BankSlot).Value
End Function

Sub SetPlayerBankItemValue(ByVal Index As Long, ByVal BankSlot As Byte, ByVal ItemValue As Long)
    Account(Index).Bank.Item(BankSlot).Value = ItemValue
End Sub

Function GetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Byte) As Long
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerBankItemDur = Account(Index).Bank.Item(BankSlot).Durability
End Function

Sub SetPlayerBankItemDur(ByVal Index As Long, ByVal BankSlot As Byte, ByVal DurValue As Long)
    Account(Index).Bank.Item(BankSlot).Durability = DurValue
End Sub

Function GetPlayerBankItemBind(ByVal Index As Long, ByVal BankSlot As Byte) As Long
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerBankItemBind = Account(Index).Bank.Item(BankSlot).Bind
End Function

Sub SetPlayerBankItemBind(ByVal Index As Long, ByVal BankSlot As Byte, ByVal BindValue As Long)
    Account(Index).Bank.Item(BankSlot).Bind = BindValue
End Sub

Function GetPlayerGender(ByVal Index As Long) As Long

    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerGender = Account(Index).Chars(GetPlayerChar(Index)).Gender
    Exit Function
End Function

Sub SetPlayerGender(ByVal Index As Long, GenderNum As Byte)
    Account(Index).Chars(GetPlayerChar(Index)).Gender = GenderNum
End Sub

Function GetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Byte) As Integer
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemDur = Account(Index).Chars(GetPlayerChar(Index)).Inv(InvSlot).Durability
End Function

Sub SetPlayerInvItemDur(ByVal Index As Long, ByVal InvSlot As Byte, ByVal ItemDur As Integer)
    Account(Index).Chars(GetPlayerChar(Index)).Inv(InvSlot).Durability = ItemDur
End Sub

Function GetPlayerInvItemBind(ByVal Index As Long, ByVal InvSlot As Byte) As Integer
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerInvItemBind = Account(Index).Chars(GetPlayerChar(Index)).Inv(InvSlot).Bind
End Function

Sub SetPlayerInvItemBind(ByVal Index As Long, ByVal InvSlot As Byte, ByVal BindType As Byte)
    Account(Index).Chars(GetPlayerChar(Index)).Inv(InvSlot).Bind = BindType
End Sub

Function GetMapItemX(ByVal MapNum As Integer, ByVal MapItemNum As Integer)
    GetMapItemX = MapItem(MapNum, MapItemNum).X
End Function

Sub SetMapItemX(ByVal MapNum As Integer, ByVal MapItemNum As Integer, ByVal Value As Long)
    MapItem(MapNum, MapItemNum).X = Value
End Sub

Function GetMapItemY(ByVal MapNum As Integer, ByVal MapItemNum As Integer)
    GetMapItemY = MapItem(MapNum, MapItemNum).Y
End Function

Sub SetMapItemY(ByVal MapNum As Integer, ByVal MapItemNum As Integer, ByVal Value As Long)
    MapItem(MapNum, MapItemNum).Y = Value
End Sub

Function GetPlayerHDSerial(ByVal Index As Long) As String
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerHDSerial = Trim$(TempPlayer(Index).HDSerial)
End Function

Function GetClassName(ByVal ClassesNum As Long) As String
    GetClassName = Trim$(Class(ClassesNum).Name)
End Function

Function GetClasseStat(ByVal ClassesNum As Long, ByVal Stat As Stats) As Long
    GetClasseStat = Class(ClassesNum).Stat(Stat)
End Function

Function GetPlayerProficiency(ByVal Index As Long, ByVal ProficiencyNum As Byte) As Long
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    
    Select Case Class(GetPlayerClass(Index)).CombatTree
        Case 1: ' Melee
            If ProficiencyNum = Proficiency.Axe Or ProficiencyNum = Proficiency.Dagger Or ProficiencyNum = Proficiency.Mace Or ProficiencyNum = Proficiency.Spear Or ProficiencyNum = Proficiency.Sword Or ProficiencyNum = Proficiency.Heavy Or ProficiencyNum = Proficiency.Light Or ProficiencyNum = Proficiency.Medium Then
                GetPlayerProficiency = 1
            Else
                GetPlayerProficiency = 0
            End If
        Case 2: ' Range
            If ProficiencyNum = Proficiency.Dagger Or ProficiencyNum = Proficiency.Bow Or ProficiencyNum = Proficiency.Crossbow Or ProficiencyNum = Proficiency.Light Or ProficiencyNum = Proficiency.Medium Then
                GetPlayerProficiency = 1
            Else
                GetPlayerProficiency = 0
            End If
        Case 3: ' Magic
            If ProficiencyNum = Proficiency.Staff Or ProficiencyNum = Proficiency.Mace Or ProficiencyNum = Proficiency.Light Then
                GetPlayerProficiency = 1
            Else
                GetPlayerProficiency = 0
            End If
    End Select
End Function
