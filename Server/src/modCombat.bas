Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################
Function GetPlayerMaxVital(ByVal index As Long, ByVal Vital As Vitals) As Long
    If index > Player_HighIndex Or index < 1 Then Exit Function
    
    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (Account(index).Chars(GetPlayerChar(index)).Stat(Stats.Endurance) / 3)) * 15 + 135
    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (Account(index).Chars(GetPlayerChar(index)).Stat(Stats.Intelligence) / 3)) * 5 + 75
    Exit Function
    
    Select Case Vital
        Case HP
            Select Case Class(GetPlayerClass(index)).CombatTree
                Case 1 ' Melee
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (Account(index).Chars(GetPlayerChar(index)).Stat(Stats.Endurance) / 3)) * 15 + 135
                Case 2 ' Range
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (Account(index).Chars(GetPlayerChar(index)).Stat(Stats.Endurance) / 3)) * 10 + 100
                Case 3 ' Magic
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (Account(index).Chars(GetPlayerChar(index)).Stat(Stats.Endurance) / 3)) * 5 + 75
            End Select

        Case MP
            Select Case Class(GetPlayerClass(index)).CombatTree
                Case 1 ' Melee
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (Account(index).Chars(GetPlayerChar(index)).Stat(Stats.Intelligence) / 3)) * 5 + 75
                Case 2 ' Range
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (Account(index).Chars(GetPlayerChar(index)).Stat(Stats.Intelligence) / 3)) * 10 + 100
                Case 3 ' Magic
                    GetPlayerMaxVital = ((GetPlayerLevel(index) / 2) + (Account(index).Chars(GetPlayerChar(index)).Stat(Stats.Intelligence) / 3)) * 15 + 135
            End Select
    End Select
End Function

Function GetPlayerVitalRegen(ByVal index As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If IsPlaying(index) = False Or index < 1 Or index > Player_HighIndex Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            i = (GetPlayerStat(index, Stats.Spirit) * 0.8) + 7
            If i > GetPlayerMaxVital(index, HP) / 25 Then
                i = GetPlayerMaxVital(index, HP) / 25
            End If
        Case MP
            i = (GetPlayerStat(index, Stats.Spirit) / 4) + 12
            If i > GetPlayerMaxVital(index, MP) / 25 Then
                i = GetPlayerMaxVital(index, MP) / 25
            End If
    End Select

    Round i
    GetPlayerVitalRegen = i
End Function

Public Sub selectValue(ByRef textBox As textBox)
    textBox.SelStart = 0
    textBox.SelLength = Len(textBox.Text)
End Sub

Function GetPlayerDamage(ByVal index As Long) As Long
    Dim WeaponNum As Long
    
    ' Check for subscript out of range
    If IsPlaying(index) = False Or index < 1 Or index > Player_HighIndex Then Exit Function
    
    If GetPlayerEquipment(index, Weapon) > 0 Then
        If Not GetPlayerEquipmentDur(index, GetPlayerEquipment(index, Weapon)) = 0 Or Item(GetPlayerEquipment(index, Weapon)).Data1 = 0 Then
            WeaponNum = GetPlayerEquipment(index, Weapon)
            GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, Strength) * Item(WeaponNum).Data2 + (GetPlayerLevel(index) / 5)
            Exit Function
        End If
    End If
    
    GetPlayerDamage = 0.085 * 5 * GetPlayerStat(index, Strength) + (GetPlayerLevel(index) / 5)
End Function

Public Function GetNpcSpellVital(ByVal MapNum As Integer, ByVal MapNpcNum As Byte, ByVal Victim As Byte, ByVal SpellNum As Long, Optional ByVal HealingSpell As Boolean = False) As Long
    If Victim < 1 Or MapNpcNum < 1 Or MapNum < 1 Then Exit Function
    If MapNpc(MapNum).NPC(MapNpcNum).Num < 1 Then Exit Function
    
    GetNpcSpellVital = Spell(SpellNum).Vital + (NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Stat(Stats.Intelligence) / 3)
    
    ' Randomize damage
    GetNpcSpellVital = Random(GetNpcSpellVital - (GetNpcSpellVital / 2), GetNpcSpellVital)
    
    ' 1.5 times the damage if it's a critical
    If CanNpcSpellCritical(MapNpcNum) Then
        GetNpcSpellVital = GetNpcSpellVital * 1.5
        Call SendSoundToMap(MapNum, Options.CriticalSound)
        SendAnimation MapNum, Options.CriticalAnimation, 0, 0, TARGET_TYPE_NPC, MapNpcNum
    End If
    
    If HealingSpell = False Then
        If MapNpc(MapNum).NPC(MapNpcNum).TargetType = TARGET_TYPE_PLAYER Then
            GetNpcSpellVital = GetNpcSpellVital - GetPlayerStat(Victim, Spirit)
        Else
            GetNpcSpellVital = GetNpcSpellVital - NPC(MapNpc(MapNum).NPC(Victim).Num).Stat(Stats.Spirit)
        End If
    End If
End Function

Function GetNpcMaxVital(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim X As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then Exit Function

    Select Case Vital
        Case HP
            GetNpcMaxVital = NPC(NpcNum).HP
        Case MP
            GetNpcMaxVital = NPC(NpcNum).MP
    End Select
End Function

Function GetNpcVitalRegen(ByVal NpcNum As Long, ByVal Vital As Vitals) As Long
    Dim i As Long

    ' Prevent subscript out of range
    If NpcNum <= 0 Or NpcNum > MAX_NPCS Then Exit Function

    Select Case Vital
        Case HP
            i = (NPC(NpcNum).Stat(Stats.Spirit) * 0.8) + 7
            If i > GetNpcMaxVital(NpcNum, HP) / 25 Then
                i = GetNpcMaxVital(NpcNum, HP) / 25
            End If
        Case MP
            i = (NPC(NpcNum).Stat(Stats.Spirit) / 4) + 12
            If i > GetNpcMaxVital(NpcNum, MP) / 25 Then
                i = GetNpcMaxVital(NpcNum, MP) / 25
            End If
    End Select
    
    Round i
    GetNpcVitalRegen = i
End Function

Function GetNpcDamage(ByVal NpcNum As Long) As Long
    GetNpcDamage = 0.085 * 5 * NPC(NpcNum).Stat(Stats.Strength) * NPC(NpcNum).Damage + (NPC(NpcNum).Level / 5)
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################
Public Function CanPlayerCritical(ByVal index As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanPlayerCritical = False

    Rate = GetPlayerStat(index, Agility) / 52.08
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanPlayerCritical = True
    End If
End Function

Public Function CanPlayerSpellCritical(ByVal index As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanPlayerSpellCritical = False

    Rate = Account(index).Chars(GetPlayerChar(index)).Stat(Stats.Intelligence) / 78.16
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanPlayerSpellCritical = True
    End If
End Function

Public Function CanPlayerDodge(ByVal index As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanPlayerDodge = False

    Rate = GetPlayerStat(index, Agility) / 83.3
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerDeflect(ByVal index As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanPlayerDeflect = False

    Rate = GetPlayerStat(index, Strength) * 0.25
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanPlayerDeflect = True
    End If
End Function

Public Function CanNpcCritical(ByVal NpcNum As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanNpcCritical = False

    Rate = NPC(NpcNum).Stat(Stats.Agility) / 52.08
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanNpcCritical = True
    End If
End Function

Public Function CanNpcSpellCritical(ByVal NpcNum As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanNpcSpellCritical = False

    Rate = NPC(NpcNum).Stat(Stats.Intelligence) / 78.16
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanNpcSpellCritical = True
    End If
End Function

Function GetPlayerProtection(ByVal index As Long) As Long
    Dim Body As Long
    Dim Helm As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or index <= 0 Or index > Player_HighIndex Then Exit Function

    Body = GetPlayerEquipment(index, Body)
    Helm = GetPlayerEquipment(index, Head)
    GetPlayerProtection = (GetPlayerStat(index, Stats.Endurance) \ 4)

    If Body > 0 Then
        If Not GetPlayerEquipmentDur(index, Body) = 0 Or Item(GetPlayerEquipment(index, Body)).Data1 = 0 Then
            GetPlayerProtection = GetPlayerProtection + Item(Body).Data2
        End If
    End If

    If Helm > 0 Then
        If Not GetPlayerEquipmentDur(index, Helm) = 0 Or Item(GetPlayerEquipment(index, Helm)).Data1 = 0 Then
            GetPlayerProtection = GetPlayerProtection + Item(Helm).Data2
        End If
    End If
End Function

Public Function CanPlayerBlock(ByVal index As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long
    Dim ShieldNum As Long

    CanPlayerBlock = False

    If GetPlayerEquipment(index, Shield) > 0 Then
        ShieldNum = GetPlayerEquipment(index, Shield)
        Rate = Item(ShieldNum).Data2 / 9
        RandomNum = Random(1, 100)
        
        If Rate > 25 Then Rate = 25
        
        If RandomNum <= Rate Then
            CanPlayerBlock = True
        End If
    End If
End Function

Function CanPlayerMitigatePlayer(ByVal Attacker As Long, Victim As Long) As Boolean
    If Account(Attacker).Chars(GetPlayerChar(Attacker)).Dir = DIR_UP Then
        If Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_DOWN Then
            CanPlayerMitigatePlayer = True
        End If
    ElseIf Account(Attacker).Chars(GetPlayerChar(Attacker)).Dir = DIR_DOWN Then
        If Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_UP Then
            CanPlayerMitigatePlayer = True
        End If
    ElseIf Account(Attacker).Chars(GetPlayerChar(Attacker)).Dir = DIR_LEFT Then
        If Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_RIGHT Then
            CanPlayerMitigatePlayer = True
        End If
    ElseIf Account(Attacker).Chars(GetPlayerChar(Attacker)).Dir = DIR_RIGHT Then
        If Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_LEFT Then
            CanPlayerMitigatePlayer = True
        End If
    Else
        CanPlayerMitigatePlayer = False
    End If
End Function

Function CanPlayerMitigateNpc(ByVal index As Long, MapNpcNum As Long) As Boolean
    If Account(index).Chars(GetPlayerChar(index)).Dir = DIR_UP Then
        If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_DOWN Then
            CanPlayerMitigateNpc = True
        End If
    ElseIf Account(index).Chars(GetPlayerChar(index)).Dir = DIR_DOWN Then
        If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_UP Then
            CanPlayerMitigateNpc = True
        End If
    ElseIf Account(index).Chars(GetPlayerChar(index)).Dir = DIR_LEFT Then
        If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_RIGHT Then
            CanPlayerMitigateNpc = True
        End If
    ElseIf Account(index).Chars(GetPlayerChar(index)).Dir = DIR_RIGHT Then
        If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_LEFT Then
            CanPlayerMitigateNpc = True
        End If
    Else
        CanPlayerMitigateNpc = False
    End If
End Function

Function CanNpcMitigatePlayer(ByVal MapNpcNum As Long, index As Long) As Boolean
    If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_UP Then
        If Account(index).Chars(GetPlayerChar(index)).Dir = DIR_DOWN Then
            CanNpcMitigatePlayer = True
        End If
    ElseIf MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_DOWN Then
        If Account(index).Chars(GetPlayerChar(index)).Dir = DIR_UP Then
            CanNpcMitigatePlayer = True
        End If
    ElseIf MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_LEFT Then
        If Account(index).Chars(GetPlayerChar(index)).Dir = DIR_RIGHT Then
            CanNpcMitigatePlayer = True
        End If
    ElseIf MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_RIGHT Then
        If Account(index).Chars(GetPlayerChar(index)).Dir = DIR_LEFT Then
            CanNpcMitigatePlayer = True
        End If
    Else
        CanNpcMitigatePlayer = False
    End If
End Function

Function CanNpcMitigateNpc(ByVal Attacker As Long, Victim As Long, MapNum As Integer) As Boolean
    If MapNpc(MapNum).NPC(Attacker).Dir = DIR_UP Then
        If MapNpc(MapNum).NPC(Victim).Dir = DIR_DOWN Then
            CanNpcMitigateNpc = True
        End If
    ElseIf MapNpc(MapNum).NPC(Attacker).Dir = DIR_DOWN Then
        If MapNpc(MapNum).NPC(Victim).Dir = DIR_UP Then
            CanNpcMitigateNpc = True
        End If
    ElseIf MapNpc(MapNum).NPC(Attacker).Dir = DIR_LEFT Then
        If MapNpc(MapNum).NPC(Victim).Dir = DIR_RIGHT Then
            CanNpcMitigateNpc = True
        End If
    ElseIf MapNpc(MapNum).NPC(Attacker).Dir = DIR_RIGHT Then
        If MapNpc(MapNum).NPC(Victim).Dir = DIR_LEFT Then
            CanNpcMitigateNpc = True
        End If
    Else
        CanNpcMitigateNpc = False
    End If
End Function

Public Function CanNpcDodge(ByVal NpcNum As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanNpcDodge = False

    Rate = NPC(NpcNum).Stat(Stats.Agility) / 83.3
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanNpcDodge = True
    End If
End Function

Public Function CanNpcDeflect(ByVal NpcNum As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanNpcDeflect = False

    Rate = NPC(NpcNum).Stat(Stats.Strength) * 0.25
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanNpcDeflect = True
    End If
End Function

' ###################################
' ##      Player Attacking Npc     ##
' ###################################
Public Sub TryPlayerAttackNpc(ByVal index As Long, ByVal MapNpcNum As Long)
    Dim NpcNum As Long
    Dim MapNum As Integer
    Dim Damage As Long
    
    ' Can we attack the npc?
    If CanPlayerAttackNpc(index, MapNpcNum, False) Then
        MapNum = GetPlayerMap(index)
        NpcNum = MapNpc(MapNum).NPC(MapNpcNum).Num
    
        ' Get the damage we can do
        Damage = GetPlayerDamage(index)
        
        ' Add damage based on direction
        If Account(index).Chars(GetPlayerChar(index)).Dir = DIR_UP Then
            If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_LEFT Or MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_UP Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf Account(index).Chars(GetPlayerChar(index)).Dir = DIR_DOWN Then
            If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_LEFT Or MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf Account(index).Chars(GetPlayerChar(index)).Dir = DIR_LEFT Then
            If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_UP Or MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_LEFT Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf Account(index).Chars(GetPlayerChar(index)).Dir = DIR_RIGHT Then
            If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_UP Or MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 4)
            End If
        End If
        
        ' 1.5 times the damage if it's a critical
        If CanPlayerCritical(index) Then
            Damage = Damage * 1.5
            Call SendSoundToMap(MapNum, Options.CriticalSound)
            SendAnimation MapNum, Options.CriticalAnimation, 0, 0, TARGET_TYPE_PLAYER, index
        End If
        
        ' Take away protection from the damage
        Damage = Damage - (NPC(MapNpcNum).Stat(Stats.Endurance) / 4)
        
        ' Randomize damage
        Damage = Random(Damage - (Damage / 2), Damage)
        
        Round Damage
        
        If Damage < 1 Then
            Call SendSoundToMap(MapNum, Options.MissSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_NPC, MapNpcNum
            Exit Sub
        End If
    
        Call PlayerAttackNpc(index, MapNpcNum, Damage)
    End If
End Sub

Public Function CanPlayerAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal UsingBow As Boolean) As Boolean
    Dim MapNum As Integer
    Dim NpcNum As Long
    Dim NpcX As Long
    Dim NpcY As Long
    Dim Attackspeed As Long
    Dim WeaponSlot As Long
    Dim Range As Byte
    Dim DistanceToNpc As Integer

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Then Exit Function

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Attacker)).NPC(MapNpcNum).Num < 1 Then Exit Function
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).NPC(MapNpcNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.HP) < 1 Then
        If Not NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_FRIENDLY And Not NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_SHOPKEEPER And Not NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_QUEST And Not NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_GUIDE Then Exit Function
    End If

    ' Make sure they are a player killer or else they can't attack a guard
    If NPC(NpcNum).Behavior = NPC_BEHAVIOR_GUARD And GetPlayerPK(Attacker) = NO Then Exit Function

    ' Attack speed from weapon
    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        Attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).WeaponSpeed
    Else
        Attackspeed = 1000
    End If
    
    If NpcNum > 0 And timeGetTime > TempPlayer(Attacker).AttackTimer + Attackspeed Then
        If Not IsSpell Then ' Melee attack
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    If Not ((MapNpc(MapNum).NPC(MapNpcNum).Y + 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum).NPC(MapNpcNum).X = GetPlayerX(Attacker))) Then Exit Function
                Case DIR_DOWN
                    If Not ((MapNpc(MapNum).NPC(MapNpcNum).Y - 1 = GetPlayerY(Attacker)) And (MapNpc(MapNum).NPC(MapNpcNum).X = GetPlayerX(Attacker))) Then Exit Function
                Case DIR_LEFT
                    If Not ((MapNpc(MapNum).NPC(MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNpc(MapNum).NPC(MapNpcNum).X + 1 = GetPlayerX(Attacker))) Then Exit Function
                Case DIR_RIGHT
                    If Not ((MapNpc(MapNum).NPC(MapNpcNum).Y = GetPlayerY(Attacker)) And (MapNpc(MapNum).NPC(MapNpcNum).X - 1 = GetPlayerX(Attacker))) Then Exit Function
                Case Else
                    Exit Function
            End Select
        End If

        If Not IsSpell And Not UsingBow Then
            If Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_FRIENDLY And Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_SHOPKEEPER And Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_QUEST And Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_GUIDE Then
                If DidNpcMitigatePlayer(Attacker, MapNpcNum) = False Then
                    CanPlayerAttackNpc = True
                End If
            ElseIf Len(Trim$(NPC(NpcNum).AttackSay)) > 0 Then
                Call SendChatBubble(MapNum, MapNpcNum, TARGET_TYPE_NPC, Trim$(NPC(NpcNum).AttackSay), White)
            End If
        ElseIf UsingBow Then
            If Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_FRIENDLY And Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_SHOPKEEPER And Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_QUEST And Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_GUIDE Then
                If DidNpcMitigatePlayer(Attacker, MapNpcNum) = False Then
                    CanPlayerAttackNpc = True
                End If
            End If
        ElseIf IsSpell Then
            If Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_FRIENDLY And Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_SHOPKEEPER And Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_QUEST And Not NPC(NpcNum).Behavior = NPC_BEHAVIOR_GUIDE Then
                If DidNpcMitigatePlayer(Attacker, MapNpcNum) = False Then
                    CanPlayerAttackNpc = True
                End If
            End If
        End If
    End If
End Function

Public Function DidNpcMitigatePlayer(ByVal Attacker As Long, ByVal MapNpcNum As Long) As Boolean
    Dim MapNum As Integer
    Dim NpcNum As Long
    
    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).NPC(MapNpcNum).Num
    
    If CanNpcMitigatePlayer(MapNpcNum, Attacker) = True Or TempPlayer(Attacker).SpellBuffer.Spell > 0 Then
        ' Check if NPC can avoid the attack
        If CanNpcDodge(NpcNum) Then
            Call SendSoundToMap(MapNum, Options.DodgeSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_NPC, MapNpcNum
            DidNpcMitigatePlayer = True
            Exit Function
        End If
        
        If CanNpcDeflect(NpcNum) Then
            Call SendSoundToMap(MapNum, Options.DeflectSound)
            SendAnimation MapNum, Options.DeflectAnimation, 0, 0, TARGET_TYPE_NPC, MapNpcNum
            DidNpcMitigatePlayer = True
            Exit Function
        End If
    End If
    
    DidNpcMitigatePlayer = False
End Function

Public Sub PlayerAttackNpc(ByVal Attacker As Long, ByVal MapNpcNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal OverTime As Boolean = False)
    Dim Name As String
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim STR As Long
    Dim DEF As Long
    Dim MapNum As Integer
    Dim NpcNum As Long
    Dim Value As Long
    Dim LevelDiff As Byte

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNpcNum <= 0 Or MapNpcNum > MAX_MAP_NPCS Then Exit Sub

    MapNum = GetPlayerMap(Attacker)
    NpcNum = MapNpc(MapNum).NPC(MapNpcNum).Num
    Name = Trim$(NPC(NpcNum).Name)
    
    ' Set the attacker's target
    If SpellNum = 0 Then
        TempPlayer(Attacker).TargetType = TARGET_TYPE_NPC
        TempPlayer(Attacker).Target = MapNpcNum
        Call SendPlayerTarget(Attacker)
    End If
    
    ' Set their target if they are being hit
    MapNpc(MapNum).NPC(MapNpcNum).TargetType = TARGET_TYPE_PLAYER
    MapNpc(MapNum).NPC(MapNpcNum).Target = Attacker
    Call SendMapNpcTarget(MapNum, MapNpcNum, MapNpc(MapNum).NPC(MapNpcNum).Target, MapNpc(GetPlayerMap(Attacker)).NPC(MapNpcNum).TargetType)
    
    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' Set the regen timer
    TempPlayer(Attacker).StopRegen = True
    TempPlayer(Attacker).StopRegenTimer = timeGetTime
    
    ' Send the sound
    If SpellNum > 0 Then
        Call SendMapSound(MapNum, Attacker, MapNpc(MapNum).NPC(MapNpcNum).X, MapNpc(MapNum).NPC(MapNpcNum).Y, SoundEntity.seSpell, SpellNum)
     Else
        Call SendMapSound(MapNum, Attacker, MapNpc(MapNum).NPC(MapNpcNum).X, MapNpc(MapNum).NPC(MapNpcNum).Y, SoundEntity.seAnimation, 1)
     End If

    If Damage >= MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.HP) Then
        ' Set the damage to the npc's health so that it doesn't appear that it's overkilling it
        Damage = MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.HP)
    
        SendActionMsg GetPlayerMap(Attacker), "-" & Damage, BrightRed, 1, (MapNpc(MapNum).NPC(MapNpcNum).X * 32), (MapNpc(MapNum).NPC(MapNpcNum).Y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(MapNum).NPC(MapNpcNum).X, MapNpc(MapNum).NPC(MapNpcNum).Y
         
        ' Send animation
        If SpellNum < 1 Then
            If GetPlayerEquipment(Attacker, Weapon) > 0 Then
                If Item(GetPlayerEquipment(Attacker, Weapon)).Animation > 0 Then
                    If n > 0 Then
                        If Not OverTime Then
                            Call SendAnimation(MapNum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
                        End If
                    End If
                Else
                    Call SendAnimation(MapNum, 1, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
                End If
            Else
                Call SendAnimation(MapNum, 1, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
            End If
        End If

        ' Calculate experience to give attacker
        Exp = NPC(NpcNum).Exp
        
        ' Find the level difference between the npc and player
        LevelDiff = GetPlayerLevel(Attacker) - NPC(NpcNum).Level
        
        If Exp > 0 Then
            If LevelDiff > 0 Then
                Exp = Exp / (Exp / (LevelDiff * 10))
            ElseIf LevelDiff < 0 Then
                Exp = Exp + (Exp * (LevelDiff * -2.5))
            End If
        End If
        
        ' Adjust the exp based on the server's rate for experience
        Exp = Exp * EXP_RATE
        
        ' Randomize
        Exp = Random(Exp * 0.95, Exp * 1.05)
        
        Round Exp
        
        ' Make sure we dont get less then 0
        If Exp < 0 Then Exp = 0
        
        ' In party
        If TempPlayer(Attacker).InParty > 0 Then
            ' Pass through party sharing function
            Party_ShareExp TempPlayer(Attacker).InParty, Exp, Attacker
        ElseIf GetPlayerLevel(Attacker) < MAX_LEVEL Then
            ' No party - keep exp for self
            Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
            SendPlayerExp Attacker
            SendActionMsg GetPlayerMap(Attacker), "+" & Exp & " Exp", White, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
            
            ' Check if we've leveled
            CheckPlayerLevelUp Attacker
        Else
            SendActionMsg GetPlayerMap(Attacker), "+0 Exp", White, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
        End If

        ' Drop the goods if they have anything to drop
        For n = 1 To MAX_NPC_DROPS
            If NPC(NpcNum).DropItem(n) = 0 Then Exit For
            
            Value = NPC(NpcNum).DropValue(n)
            Value = Random(Value * 0.25, Value * 1.5)
            Round Value
            
            If Value < 1 Then Value = 1
            
            If Rnd <= NPC(NpcNum).DropChance(n) Then
                If TempPlayer(Attacker).InParty > 0 Then
                    Call Party_GetLoot(TempPlayer(Attacker).InParty, NPC(NpcNum).DropItem(n), NPC(NpcNum).DropValue(n), MapNpc(MapNum).NPC(MapNpcNum).X, MapNpc(MapNum).NPC(MapNpcNum).Y)
                Else
                    Call SpawnItem(NPC(NpcNum).DropItem(n), Value, Item(NPC(NpcNum).DropItem(n)).Data1, MapNum, MapNpc(MapNum).NPC(MapNpcNum).X, MapNpc(MapNum).NPC(MapNpcNum).Y, GetPlayerName(Attacker))
                End If
            End If
        Next

        ' Now set HP to 0, so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).NPC(MapNpcNum).Num = 0
        MapNpc(MapNum).NPC(MapNpcNum).SpawnWait = timeGetTime
        MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.HP) = 0
        UpdateMapBlock MapNum, MapNpc(MapNum).NPC(MapNpcNum).X, MapNpc(MapNum).NPC(MapNpcNum).Y, False
        
        ' Clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).NPC(MapNpcNum).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(MapNum).NPC(MapNpcNum).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' Send death to the map
        Call SendNpcDeath(MapNpcNum, MapNum)
        
        ' Loop through entire map and purge npcs from players
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If Account(i).Chars(GetPlayerChar(i)).Map = MapNum Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_NPC Then
                        If TempPlayer(i).Target = MapNpcNum Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendPlayerTarget i
                        End If
                    End If
                End If
            End If
        Next
    Else
        ' Npc not dead, just do the damage
        MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.HP) = MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.HP) - Damage
        Call SendMapNpcVitals(MapNum, MapNpcNum)

        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).NPC(MapNpcNum).X * 32), (MapNpc(MapNum).NPC(MapNpcNum).Y * 32)
        SendBlood GetPlayerMap(Attacker), MapNpc(MapNum).NPC(MapNpcNum).X, MapNpc(MapNum).NPC(MapNpcNum).Y
        
        ' Send animation
        If SpellNum < 1 Then
            If GetPlayerEquipment(Attacker, Weapon) > 0 Then
                If Item(GetPlayerEquipment(Attacker, Weapon)).Animation > 0 Then
                    If n > 0 Then
                        If Not OverTime Then
                            Call SendAnimation(MapNum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
                        End If
                    End If
                Else
                    Call SendAnimation(MapNum, 1, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
                End If
            Else
                Call SendAnimation(MapNum, 1, 0, 0, TARGET_TYPE_NPC, MapNpcNum)
            End If
        End If
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_GUARD Then
            For i = 1 To Map(MapNum).Npc_HighIndex
                If MapNpc(MapNum).NPC(i).Num = MapNpc(MapNum).NPC(MapNpcNum).Num Then
                    MapNpc(MapNum).NPC(i).Target = Attacker
                    MapNpc(MapNum).NPC(i).TargetType = TARGET_TYPE_PLAYER
                    Call SendMapNpcTarget(MapNum, i, MapNpc(MapNum).NPC(i).Target, MapNpc(MapNum).NPC(i).TargetType)
                End If
            Next
        End If
        
        ' Set the regen timer
        MapNpc(MapNum).NPC(MapNpcNum).StopRegen = True
        MapNpc(MapNum).NPC(MapNpcNum).StopRegenTimer = timeGetTime
        
        ' If stunning spell then stun the npc
        If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
            If Spell(SpellNum).StunDuration > 0 Then StunNPC MapNpcNum, MapNum, SpellNum
            
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Npc MapNum, MapNpcNum, SpellNum, Attacker
            End If
        End If
        SendMapNpcVitals MapNum, MapNpcNum
    End If
    
    If SpellNum = 0 Then
        ' Reset the attack timer
        TempPlayer(Attacker).AttackTimer = timeGetTime
    End If
    
    ' Reduce durability of weapon
    Call DamagePlayerEquipment(Attacker, Weapon)
End Sub

' ###################################
' ##      Npc Attacking Npc        ##
' ###################################
Public Sub TryNpcAttackNpc(ByVal MapNum As Integer, ByVal Attacker As Long, ByVal Victim As Long)
    Dim NpcNum As Long, Damage As Long, i As Long
    
    NpcNum = MapNpc(MapNum).NPC(Attacker).Num
    Damage = GetNpcDamage(NpcNum)

    ' Set the npc target to the npc
    If MapNpc(MapNum).NPC(Victim).Target = 0 Then
        MapNpc(MapNum).NPC(Victim).Target = Attacker
        MapNpc(MapNum).NPC(Victim).TargetType = TARGET_TYPE_NPC
        Call SendMapNpcTarget(MapNum, Victim, MapNpc(MapNum).NPC(Victim).Target, MapNpc(MapNum).NPC(Victim).TargetType)
    End If
        
    ' Can the npc attack the player
    If CanNpcAttackNpc(MapNum, Attacker, Victim) Then
        ' Set attack timer
        MapNpc(MapNum).NPC(Attacker).AttackTimer = timeGetTime
        
        If NPC(MapNpc(MapNum).NPC(Victim).Num).FactionThreat = True Then
            ' Send threat to all of the same faction if they have the option enabled
            For i = 1 To Map(MapNum).Npc_HighIndex
                If MapNpc(MapNum).NPC(i).Num > 0 Then
                    If NPC(MapNpc(MapNum).NPC(Victim).Num).Faction > 0 And NPC(MapNpc(MapNum).NPC(i).Num).Faction > 0 Then
                        If NPC(MapNpc(MapNum).NPC(Victim).Num).Faction = NPC(MapNpc(MapNum).NPC(i).Num).Faction Then
                            If NPC(MapNpc(MapNum).NPC(Victim).Num).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Then
                                If MapNpc(MapNum).NPC(i).Target = 0 Then
                                    MapNpc(MapNum).NPC(i).TargetType = TARGET_TYPE_NPC
                                    MapNpc(MapNum).NPC(i).Target = Attacker
                                    Call SendMapNpcTarget(MapNum, i, MapNpc(MapNum).NPC(i).Target, MapNpc(MapNum).NPC(i).TargetType)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        
        ' Add damage based on direction
        If MapNpc(MapNum).NPC(Attacker).Dir = DIR_UP Then
            If MapNpc(MapNum).NPC(Victim).Dir = DIR_LEFT Or MapNpc(MapNum).NPC(Victim).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNpc(MapNum).NPC(Victim).Dir = DIR_UP Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNpc(MapNum).NPC(Attacker).Dir = DIR_DOWN Then
            If MapNpc(MapNum).NPC(Victim).Dir = DIR_LEFT Or MapNpc(MapNum).NPC(Victim).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNpc(MapNum).NPC(Victim).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNpc(MapNum).NPC(Attacker).Dir = DIR_LEFT Then
            If MapNpc(MapNum).NPC(Victim).Dir = DIR_UP Or MapNpc(MapNum).NPC(Victim).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNpc(MapNum).NPC(Victim).Dir = DIR_LEFT Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNpc(MapNum).NPC(Attacker).Dir = DIR_RIGHT Then
            If MapNpc(MapNum).NPC(Victim).Dir = DIR_UP Or MapNpc(MapNum).NPC(Victim).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNpc(MapNum).NPC(Victim).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 4)
            End If
        End If
        
        ' 1.5 times the damage if it's a critical
        If CanNpcCritical(Attacker) Then
            Damage = Damage * 1.5
            Call SendSoundToMap(MapNum, Options.CriticalSound)
            SendAnimation MapNum, Options.CriticalAnimation, 0, 0, TARGET_TYPE_NPC, Attacker
        End If
        
        ' Take away protection from the damage
        Damage = Damage - (NPC(Victim).Stat(Stats.Endurance) / 4)
        
        ' Randomize damage
        Damage = Random(Damage - (Damage / 2), Damage)
        
        Round Damage
        
        ' Make sure the damage isn't 0
        If Damage < 1 Then
            Call SendSoundToMap(MapNum, Options.MissSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_NPC, Attacker
            Exit Sub
        End If
        
        ' Send the sound
        If Trim$(NPC(MapNpc(MapNum).NPC(Victim).Num).Sound) = vbNullString Then
            Call SendMapSound(MapNum, Victim, MapNpc(MapNum).NPC(Victim).X, MapNpc(MapNum).NPC(Victim).Y, SoundEntity.seAnimation, 1)
        Else
            Call SendMapSound(MapNum, Victim, MapNpc(MapNum).NPC(Victim).X, MapNpc(MapNum).NPC(Victim).Y, SoundEntity.seNpc, MapNpc(MapNum).NPC(Victim).Num)
        End If
        
        Call NpcAttackNpc(MapNum, Attacker, Victim, Damage)
    End If
End Sub

Function CanNpcAttackNpc(ByVal MapNum As Integer, ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal Spell As Boolean = False) As Boolean
    Dim aNpcNum As Long
    Dim vNpcNum As Long
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long
    
    CanNpcAttackNpc = False

    ' Check for subscript out of range
    If Attacker < 1 Or Attacker > MAX_MAP_NPCS Then Exit Function
    If Victim < 1 Or Victim > MAX_MAP_NPCS Then Exit Function
    
    aNpcNum = MapNpc(MapNum).NPC(Attacker).Num
    vNpcNum = MapNpc(MapNum).NPC(Victim).Num
    
     ' Check for subscript out of range
    If aNpcNum < 1 Or vNpcNum < 1 Then Exit Function
    
    ' Can't attack itself
    If aNpcNum = vNpcNum Then Exit Function

    ' Make sure the Npcs aren't already dead
    If MapNpc(MapNum).NPC(Attacker).Vital(Vitals.HP) < 1 Or MapNpc(MapNum).NPC(Victim).Vital(Vitals.HP) < 1 Then Exit Function
    
    ' Make sure they aren't trying to attack a friendly or shopkeeper NPC
    If NPC(MapNpc(MapNum).NPC(Victim).Num).Behavior = NPC_BEHAVIOR_FRIENDLY Or NPC(MapNpc(MapNum).NPC(Victim).Num).Behavior = NPC_BEHAVIOR_SHOPKEEPER Or NPC(MapNpc(MapNum).NPC(Victim).Num).Behavior = NPC_BEHAVIOR_QUEST Or NPC(MapNpc(MapNum).NPC(Victim).Num).Behavior = NPC_BEHAVIOR_GUIDE Then Exit Function
    
    ' Make sure they aren't casting a spell
    If MapNpc(MapNum).NPC(Attacker).SpellBuffer.Timer > 0 And Spell = False Then Exit Function

    ' Check if they have the same faction if they do exit
    If NPC(MapNpc(MapNum).NPC(Attacker).Num).Faction > 0 Then
        If NPC(MapNpc(MapNum).NPC(Attacker).Num).Faction = NPC(MapNpc(MapNum).NPC(Victim).Num).Faction Then Exit Function
    End If
    
    If Spell Then
        CanNpcAttackNpc = True
        Exit Function
    End If

    ' Make sure npcs don't attack more than once a second
    If timeGetTime < MapNpc(MapNum).NPC(Attacker).AttackTimer + 1000 Then Exit Function
    
    AttackerX = MapNpc(MapNum).NPC(Attacker).X
    AttackerY = MapNpc(MapNum).NPC(Attacker).Y
    VictimX = MapNpc(MapNum).NPC(Victim).X
    VictimY = MapNpc(MapNum).NPC(Victim).Y
    
    ' Check if they are going to cast
    If Random(1, 2) = 1 And CanNpcCastSpell(MapNum, Attacker) Then
        Call BufferNpcSpell(MapNum, Attacker, Victim)
        Exit Function
    End If
        
    ' Check if at same coordinates
    If AttackerY + 1 = VictimY And AttackerX = VictimX Then
    ElseIf AttackerY - 1 = VictimY And AttackerX = VictimX Then
    ElseIf AttackerY = VictimY And AttackerX + 1 = VictimX Then
    ElseIf AttackerY = VictimY And AttackerX - 1 = VictimX Then
    Else
        Exit Function
    End If
    
    CanNpcAttackNpc = True
End Function

Private Function DidNpcMitigateNpc(ByVal MapNum As Integer, Attacker As Long, Victim As Long) As Boolean
    If CanNpcMitigateNpc(Attacker, Victim, MapNum) = True Or MapNpc(MapNum).NPC(Attacker).SpellBuffer.Spell > 0 Then
        ' Check if npc can avoid the attack
        If CanNpcDodge(Victim) Then
            Call SendSoundToMap(MapNum, Options.DodgeSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_NPC, Victim
            DidNpcMitigateNpc = True
            Exit Function
        End If
        
        ' Check if npc can deflect the attack
        If CanNpcDeflect(Victim) Then
            Call SendSoundToMap(MapNum, Options.DeflectSound)
            SendAnimation MapNum, Options.DeflectAnimation, 0, 0, TARGET_TYPE_NPC, Victim
            DidNpcMitigateNpc = True
            Exit Function
        End If
    End If
    
    DidNpcMitigateNpc = False
End Function

Sub NpcAttackNpc(ByVal MapNum As Integer, ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim i As Long, n As Byte
    Dim aVictim As Long
    Dim vVictim As Long
    Dim DistanceX As Byte
    Dim DistanceY As Byte
    Dim Value As Long
    
    ' Check for subscript out of range
    If Attacker < 1 Or Attacker > MAX_MAP_NPCS Or Victim < 1 Or Victim > MAX_MAP_NPCS Then Exit Sub
    
    If DidNpcMitigateNpc(MapNum, Attacker, Victim) = True Then Exit Sub
    
    aVictim = MapNpc(MapNum).NPC(Attacker).Num
    vVictim = MapNpc(MapNum).NPC(Victim).Num
    
    ' Check for subscript out of range
    If aVictim < 1 Then Exit Sub
    If vVictim < 1 Then Exit Sub
    
    ' Send this packet so they can see the person attacking
    Call SendNpcAttack(Attacker, MapNum)
    
     ' Set the regen timer
    MapNpc(MapNum).NPC(Attacker).StopRegen = True
    MapNpc(MapNum).NPC(Attacker).StopRegenTimer = timeGetTime

    If Damage >= MapNpc(MapNum).NPC(Victim).Vital(Vitals.HP) Then
        ' Set the damage to the target's health exactly so it's not overkilling them
        Damage = MapNpc(MapNum).NPC(Victim).Vital(Vitals.HP)
        
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).NPC(Victim).X * 32), (MapNpc(MapNum).NPC(Victim).Y * 32)
        SendBlood MapNum, MapNpc(MapNum).NPC(Victim).X, MapNpc(MapNum).NPC(Victim).Y
        
        Call SendMapNpcTarget(MapNum, Attacker, 0, 0)
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNpc(MapNum).NPC(Victim).Num = 0
        MapNpc(MapNum).NPC(Victim).SpawnWait = timeGetTime
        MapNpc(MapNum).NPC(Victim).Vital(Vitals.HP) = 0
        UpdateMapBlock MapNum, MapNpc(MapNum).NPC(Victim).X, MapNpc(MapNum).NPC(Victim).Y, False
        
        ' Clear DoTs and HoTs
        For i = 1 To MAX_DOTS
            With MapNpc(MapNum).NPC(Victim).DoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNpc(MapNum).NPC(Victim).HoT(i)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' Send npc death packet to map
        Call SendNpcDeath(Victim, MapNum)
    Else
        ' Npc not dead, just do the damage
        If MapNpc(MapNum).NPC(Attacker).SpellBuffer.Spell = 0 Then
            If NPC(MapNpc(MapNum).NPC(Attacker).Num).Animation < 1 Then
                Call SendAnimation(MapNum, 1, 0, 0, TARGET_TYPE_NPC, Victim)
            Else
                Call SendAnimation(MapNum, NPC(MapNpc(MapNum).NPC(Attacker).Num).Animation, 0, 0, TARGET_TYPE_NPC, Victim)
            End If
        End If
        
        MapNpc(MapNum).NPC(Victim).Vital(Vitals.HP) = MapNpc(MapNum).NPC(Victim).Vital(Vitals.HP) - Damage
        Call SendMapNpcVitals(MapNum, Victim)
        
        ' Say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNpc(MapNum).NPC(Victim).X * 32), (MapNpc(MapNum).NPC(Victim).Y * 32)
        SendBlood MapNum, MapNpc(MapNum).NPC(Victim).X, MapNpc(MapNum).NPC(Victim).Y
        
        ' Set the regen timer
        TempPlayer(Victim).StopRegen = True
        TempPlayer(Victim).StopRegenTimer = timeGetTime
    End If
End Sub

' ###################################
' ##      Npc Attacking Player     ##
' ###################################
Public Sub TryNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long)
    Dim MapNum As Integer, NpcNum As Long, Damage As Long, n As Byte, DistanceX As Byte, DistanceY As Byte
    
    MapNum = GetPlayerMap(index)
    NpcNum = MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Num
    
    ' Can the npc attack the player
    If CanNpcAttackPlayer(MapNpcNum, index) Then
        ' Set attack timer
        MapNpc(MapNum).NPC(MapNpcNum).AttackTimer = timeGetTime
        
        If NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).FactionThreat = True Then
            ' Send threat to all of the same faction if they have the option enabled
            For n = 1 To Map(MapNum).Npc_HighIndex
                If MapNpc(MapNum).NPC(n).Num > 0 Then
                    If NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Faction > 0 And NPC(MapNpc(MapNum).NPC(n).Num).Faction > 0 Then
                        If NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Faction = NPC(MapNpc(MapNum).NPC(n).Num).Faction Then
                            If NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Then
                                If MapNpc(MapNum).NPC(n).Target = 0 Then
                                    MapNpc(MapNum).NPC(n).TargetType = TARGET_TYPE_NPC
                                    MapNpc(MapNum).NPC(n).Target = MapNpcNum
                                    Call SendMapNpcTarget(MapNum, n, MapNpc(MapNum).NPC(n).Target, MapNpc(MapNum).NPC(n).TargetType)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        
        ' Don't help player killers
        If GetPlayerPK(index) = NO Then
            ' Send threat to all guards which are in range
            For n = 1 To Map(MapNum).Npc_HighIndex
                If MapNpc(MapNum).NPC(n).Num > 0 Then
                    If NPC(MapNpc(MapNum).NPC(n).Num).Behavior = NPC_BEHAVIOR_GUARD Then
                        ' X range
                        If MapNpc(MapNum).NPC(n).X > GetPlayerX(index) Then
                            DistanceX = MapNpc(MapNum).NPC(n).X - GetPlayerX(index)
                        Else
                            DistanceX = GetPlayerX(index) - MapNpc(MapNum).NPC(n).X
                        End If
                        
                        ' Y range
                        If MapNpc(MapNum).NPC(n).Y > GetPlayerY(index) Then
                            DistanceY = MapNpc(MapNum).NPC(n).Y - GetPlayerY(index)
                        Else
                            DistanceY = GetPlayerY(index) - MapNpc(MapNum).NPC(n).Y
                        End If
                        
                        n = NPC(MapNpc(MapNum).NPC(n).Num).Range
                                
                        ' Are they in range
                        If DistanceX <= n And DistanceY <= n Then
                            If MapNpc(MapNum).NPC(n).Target = 0 Then
                                MapNpc(MapNum).NPC(n).TargetType = TARGET_TYPE_NPC
                                MapNpc(MapNum).NPC(n).Target = MapNpcNum
                                Call SendMapNpcTarget(MapNum, n, MapNpc(MapNum).NPC(n).Target, MapNpc(MapNum).NPC(n).TargetType)
                            End If
                        End If
                    End If
                End If
            Next
        End If
        
        ' Get the damage we can do
        Damage = GetNpcDamage(NpcNum)
        
        ' Add damage based on direction
        If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_UP Then
            If Account(index).Chars(GetPlayerChar(index)).Dir = DIR_LEFT Or Account(index).Chars(GetPlayerChar(index)).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(index).Chars(GetPlayerChar(index)).Dir = DIR_UP Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_DOWN Then
            If Account(index).Chars(GetPlayerChar(index)).Dir = DIR_LEFT Or Account(index).Chars(GetPlayerChar(index)).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(index).Chars(GetPlayerChar(index)).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_LEFT Then
            If Account(index).Chars(GetPlayerChar(index)).Dir = DIR_UP Or Account(index).Chars(GetPlayerChar(index)).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(index).Chars(GetPlayerChar(index)).Dir = DIR_LEFT Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Dir = DIR_RIGHT Then
            If Account(index).Chars(GetPlayerChar(index)).Dir = DIR_UP Or Account(index).Chars(GetPlayerChar(index)).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(index).Chars(GetPlayerChar(index)).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 4)
            End If
        End If
        
        ' 1.5 times the damage if it's a critical
        If CanNpcCritical(MapNpcNum) Then
            Damage = Damage * 1.5
            Call SendSoundToMap(MapNum, Options.CriticalSound)
            SendAnimation MapNum, Options.CriticalAnimation, 0, 0, TARGET_TYPE_NPC, MapNpcNum
            Exit Sub
        End If
        
        ' Take away protection from the damage
        Damage = Damage - GetPlayerProtection(index)
        
        ' Randomize damage
        Damage = Random(Damage - (Damage / 2), Damage)
        
        Round Damage
        
        If Damage < 1 Then
            Call SendSoundToMap(MapNum, Options.MissSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_PLAYER, index
            Exit Sub
        End If
        
        ' Send the sound
        Call SendMapSound(MapNum, index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seAnimation, 1)

        Call NpcAttackPlayer(MapNpcNum, index, Damage)
    End If
End Sub

Function CanNpcAttackPlayer(ByVal MapNpcNum As Long, ByVal index As Long, Optional ByVal Spell As Boolean = False) As Boolean
    Dim MapNum As Integer
    Dim NpcNum As Long

    ' Check for subscript out of range
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Or Not IsPlaying(index) Then Exit Function

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(index)).NPC(MapNpcNum).Num < 1 Then Exit Function

    MapNum = GetPlayerMap(index)
    NpcNum = MapNpc(MapNum).NPC(MapNpcNum).Num

    ' Make sure the npc isn't already dead
    If MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.HP) < 1 Then Exit Function
    
    ' Make sure they aren't casting a spell
    If MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Timer > 0 And Spell = False Then Exit Function
    
    ' Can't attack if shopkeeper, friendly, quest, or guide
    If NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_FRIENDLY Or NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_SHOPKEEPER Or NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_QUEST Or NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_GUIDE Then Exit Function
    
    ' Don't attack players who are not Player Killers if the attack is a guard
    If NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Behavior = NPC_BEHAVIOR_GUARD Then
        If GetPlayerPK(index) = NO Then Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(index).GettingMap = YES Then Exit Function
    
    ' Make sure npcs don't attack more than once a second
    If timeGetTime < MapNpc(MapNum).NPC(MapNpcNum).AttackTimer + 1000 And Spell = False Then Exit Function

    If Spell Then
        CanNpcAttackPlayer = True
        Exit Function
    End If
    
    ' Adjust target if they have none
    If TempPlayer(index).Target = 0 Then
        TempPlayer(index).Target = MapNpcNum
        TempPlayer(index).TargetType = TARGET_TYPE_NPC
        Call SendPlayerTarget(index)
    End If
    
    ' Make sure they are on the same map
    If IsPlaying(index) Then
        If NpcNum > 0 Then
            ' Check if they are going to cast
            If Random(1, 2) = 1 And CanNpcCastSpell(MapNum, MapNpcNum) Then
                Call BufferNpcSpell(MapNum, MapNpcNum, MapNpc(MapNum).NPC(MapNpcNum).Target)
                Exit Function
            End If
            
            ' Check if at same coordinates
            If (GetPlayerY(index) + 1 = MapNpc(MapNum).NPC(MapNpcNum).Y) And (GetPlayerX(index) = MapNpc(MapNum).NPC(MapNpcNum).X) Then
            ElseIf (GetPlayerY(index) - 1 = MapNpc(MapNum).NPC(MapNpcNum).Y) And (GetPlayerX(index) = MapNpc(MapNum).NPC(MapNpcNum).X) Then
            ElseIf (GetPlayerY(index) = MapNpc(MapNum).NPC(MapNpcNum).Y) And (GetPlayerX(index) + 1 = MapNpc(MapNum).NPC(MapNpcNum).X) Then
            ElseIf (GetPlayerY(index) = MapNpc(MapNum).NPC(MapNpcNum).Y) And (GetPlayerX(index) - 1 = MapNpc(MapNum).NPC(MapNpcNum).X) Then
            Else
                Exit Function
            End If
            
            CanNpcAttackPlayer = True
        End If
    End If
End Function

Private Function DidPlayerMitigateNpc(ByVal MapNum As Integer, ByVal index As Long, ByVal MapNpcNum As Long) As Boolean
    If CanPlayerMitigateNpc(index, MapNpcNum) = True Or MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Spell > 0 Then
        ' Check if player can avoid the attack
        If CanPlayerDodge(index) Then
            Call SendSoundToMap(MapNum, Options.DodgeSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_PLAYER, index
            DidPlayerMitigateNpc = True
            Exit Function
        End If
        
        ' Check if player can deflect the attack
        If CanPlayerDeflect(index) Then
            Call SendSoundToMap(MapNum, Options.DeflectSound)
            SendAnimation MapNum, Options.DeflectAnimation, 0, 0, TARGET_TYPE_PLAYER, index
            DidPlayerMitigateNpc = True
            Exit Function
        End If
        
        ' Check if player can block the attack
        If CanPlayerBlock(index) Then
            Call SendSoundToMap(MapNum, Options.BlockSound)
            SendAnimation MapNum, Options.DeflectAnimation, 0, 0, TARGET_TYPE_PLAYER, index
            DidPlayerMitigateNpc = True
            Exit Function
        End If
    End If
End Function

Sub NpcAttackPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim MapNum As Integer
    Dim i As Long
    Dim n As Long
    Dim DistanceX As Byte, DistanceY As Byte

    ' Check for subscript out of range
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then Exit Sub

    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).NPC(MapNpcNum).Num < 1 Or MapNpc(GetPlayerMap(Victim)).NPC(MapNpcNum).Num > MAX_MAP_NPCS Then Exit Sub

    If DidPlayerMitigateNpc(GetPlayerMap(Victim), Victim, MapNpcNum) = True Then Exit Sub
    
    MapNum = GetPlayerMap(Victim)
    Name = Trim$(NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Name)
    
    ' Send this packet so they can see the npc attacking
    Call SendNpcAttack(MapNpcNum, MapNum)
    
    ' Reduce durability on the victim's equipment
    If Random(1, 2) = 1 Then ' Which one it will affect
        Call DamagePlayerEquipment(Victim, Body)
    Else
        Call DamagePlayerEquipment(Victim, Head)
    End If
    
    ' Set the regen timer
    MapNpc(MapNum).NPC(MapNpcNum).StopRegen = True
    MapNpc(MapNum).NPC(MapNpcNum).StopRegenTimer = timeGetTime

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Set the damage to the player's health so that it doesn't appear that it's overkilling it
        Damage = GetPlayerVital(Victim, Vitals.HP)
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' Kill player
        KillPlayer Victim
        
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & CheckGrammar(Name) & "!", BrightRed)

        ' Set npc target to 0
        MapNpc(MapNum).NPC(MapNpcNum).Target = 0
        MapNpc(MapNum).NPC(MapNpcNum).TargetType = TARGET_TYPE_NONE
        Call SendMapNpcTarget(MapNum, MapNpcNum, 0, 0)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' Send animation
        If MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Spell = 0 Then
            If NPC(MapNpc(GetPlayerMap(Victim)).NPC(MapNpcNum).Num).Animation < 1 Then
                Call SendAnimation(GetPlayerMap(Victim), 1, 0, 0, TARGET_TYPE_PLAYER, Victim)
            Else
               Call SendAnimation(GetPlayerMap(Victim), NPC(MapNpc(GetPlayerMap(Victim)).NPC(MapNpcNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
            End If
        End If
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' Set the regen timer
        TempPlayer(Victim).StopRegen = True
        TempPlayer(Victim).StopRegenTimer = timeGetTime
    End If
End Sub

Public Sub BufferNpcSpell(ByVal MapNum As Integer, ByVal MapNpcNum As Long, ByVal Target As Long)
    Dim SpellNum As Long
    Dim SpellCastType As Byte
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    SpellNum = MapNpc(MapNum).NPC(MapNpcNum).ActiveSpell
    
    If SpellNum < 1 Or SpellNum > MAX_SPELLS Then Exit Sub
    
    Range = Spell(SpellNum).Range
    
    HasBuffered = False
    
    ' Find out what kind of spell it is self cast, target or AOE
    If Range > 0 Then
        ' Ranged attack, single target or aoe?
        If Spell(SpellNum).IsAoe = False Then
            SpellCastType = 2 ' Targetted
        Else
            SpellCastType = 3 ' Targetted AoE
        End If
    Else
        If Spell(SpellNum).IsAoe = False Then
            SpellCastType = 0 ' Self-cast
        Else
            SpellCastType = 1 ' Self-cast AoE
        End If
    End If
    
    Select Case SpellCastType
        Case 0, 1 ' Self-cast & self-cast AoE
            HasBuffered = True
        Case 2, 3 ' Targeted & targeted AoE
            ' Go through spell types
            If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEHP Or Spell(SpellNum).Type <> SPELL_TYPE_HEALHP Then
                HasBuffered = True
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_NPC, MapNpcNum
        
        If Spell(SpellNum).CastTime > 0 Then
            SendActionMsg MapNum, "Casting " & Trim$(Spell(SpellNum).Name), BrightBlue, ACTIONMSG_SCROLL, MapNpc(MapNum).NPC(MapNpcNum).X * 32, MapNpc(MapNum).NPC(MapNpcNum).Y * 32
        End If
        
        MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Spell = SpellNum
        MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Timer = timeGetTime
        MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Target = MapNpc(MapNum).NPC(MapNpcNum).Target
        MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.TType = MapNpc(MapNum).NPC(MapNpcNum).TargetType
        Call SendNpcSpellBuffer(MapNum, MapNpcNum)
    End If
End Sub

Sub NpcSpellPlayer(ByVal MapNpcNum As Long, ByVal Victim As Long)
    Dim MapNum As Integer
    Dim i As Long
    Dim Damage As Long
    Dim SpellNum As Long
    Dim DidCast As Boolean, X As Byte, Y As Byte, AoE As Long

    ' Check for subscript out of range
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then Exit Sub
        
    ' Check for subscript out of range
    If MapNpc(GetPlayerMap(Victim)).NPC(MapNpcNum).Num < 1 Then Exit Sub

    ' Set the map number
    MapNum = GetPlayerMap(Victim)
     
    ' Set the spell that they are going to cast
    SpellNum = MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Spell
   
    ' Send this packet so they can see the person attacking
    Call SendNpcAttack(MapNpcNum, MapNum)
    
    ' Play the sound
    Call SendMapSound(MapNum, Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, SpellNum)
    
    DidCast = False
    AoE = Spell(SpellNum).AoE
    X = MapNpc(MapNum).NPC(MapNpc(MapNum).NPC(MapNpcNum).Target).X
    Y = MapNpc(MapNum).NPC(MapNpc(MapNum).NPC(MapNpcNum).Target).Y
    
    ' Check if the spell they are going to cast is valid
    If SpellNum > 0 Then
        If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Or Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
            Damage = GetNpcSpellVital(MapNum, MapNpcNum, Victim, SpellNum, True)
            
            If Spell(SpellNum).IsAoe = True Then ' AoE
                If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = MapNum Then
                                If IsInRange(AoE, X, Y, Account(i).Chars(GetPlayerChar(i)).X, Account(i).Chars(GetPlayerChar(i)).Y) Then
                                    Account(i).Chars(GetPlayerChar(i)).Vital(Vitals.HP) = Account(i).Chars(GetPlayerChar(i)).Vital(Vitals.HP) + Damage
                                    SendActionMsg MapNum, "+" & Damage, BrightGreen, 1, (Account(i).Chars(GetPlayerChar(i)).X * 32), (Account(i).Chars(GetPlayerChar(i)).Y * 32)
                                    
                                    Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i)
                                    DidCast = True
                                    
                                    ' Prevent overhealing
                                    If Account(i).Chars(GetPlayerChar(i)).Vital(Vitals.HP) > GetPlayerMaxVital(i, HP) Then
                                        Account(i).Chars(GetPlayerChar(i)).Vital(Vitals.HP) = GetPlayerMaxVital(i, HP)
                                    End If
                                End If
                            End If
                        End If
                    Next
                Else
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = MapNum Then
                                If IsInRange(AoE, X, Y, Account(i).Chars(GetPlayerChar(i)).X, Account(i).Chars(GetPlayerChar(i)).Y) Then
                                    Account(i).Chars(GetPlayerChar(i)).Vital(Vitals.MP) = Account(i).Chars(GetPlayerChar(i)).Vital(Vitals.MP) + Damage
                                    SendActionMsg MapNum, "+" & Damage, BrightBlue, 1, (Account(i).Chars(GetPlayerChar(i)).X * 32), (Account(i).Chars(GetPlayerChar(i)).Y * 32)
                                    
                                    Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i)
                                    DidCast = True
                                    
                                    ' Prevent overhealing
                                    If Account(i).Chars(GetPlayerChar(i)).Vital(Vitals.MP) > GetPlayerMaxVital(i, MP) Then
                                        Account(i).Chars(GetPlayerChar(i)).Vital(Vitals.MP) = GetPlayerMaxVital(i, MP)
                                    End If
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                    Account(Victim).Chars(GetPlayerChar(Victim)).Vital(Vitals.HP) = Account(Victim).Chars(GetPlayerChar(Victim)).Vital(Vitals.HP) + Damage
                    SendActionMsg MapNum, "+" & Damage, BrightGreen, 1, (Account(Victim).Chars(GetPlayerChar(Victim)).X * 32), (Account(Victim).Chars(GetPlayerChar(Victim)).Y * 32)
                    
                    ' Prevent overhealing
                    If Account(Victim).Chars(GetPlayerChar(Victim)).Vital(Vitals.HP) > GetPlayerMaxVital(Victim, HP) Then
                        Account(Victim).Chars(GetPlayerChar(Victim)).Vital(Vitals.HP) = GetPlayerMaxVital(Victim, HP)
                    End If
                Else
                    Account(Victim).Chars(GetPlayerChar(Victim)).Vital(Vitals.MP) = Account(Victim).Chars(GetPlayerChar(Victim)).Vital(Vitals.MP) + Damage
                    SendActionMsg MapNum, "+" & Damage, BrightBlue, 1, (Account(Victim).Chars(GetPlayerChar(Victim)).X * 32), (Account(Victim).Chars(GetPlayerChar(Victim)).Y * 32)
    
                    ' Prevent overhealing
                    If Account(Victim).Chars(GetPlayerChar(Victim)).Vital(Vitals.MP) > GetPlayerMaxVital(Victim, MP) Then
                        Account(Victim).Chars(GetPlayerChar(Victim)).Vital(Vitals.MP) = GetPlayerMaxVital(Victim, MP)
                    End If
                End If
            End If
            
            Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Victim)
            DidCast = True
        Else
            If Spell(SpellNum).IsAoe = True Then ' AoE
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If GetPlayerMap(i) = MapNum Then
                            If IsInRange(AoE, X, Y, Account(i).Chars(GetPlayerChar(i)).X, Account(i).Chars(GetPlayerChar(i)).Y) Then
                                If CanNpcAttackPlayer(MapNpcNum, i, True) Then
                                    Damage = GetNpcSpellVital(MapNum, MapNpcNum, i, SpellNum)
                            
                                    If Damage < 1 Then
                                        Call SendSoundToMap(GetPlayerMap(i), Options.ResistSound)
                                        SendActionMsg GetPlayerMap(i), "Resist", Pink, 1, (GetPlayerX(i) * 32), (GetPlayerY(i) * 32)
                                    Else
                                        Call NpcAttackPlayer(MapNpcNum, i, Damage)
                                        Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i)
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            Else ' Non AoE
                Damage = GetNpcSpellVital(MapNum, MapNpcNum, Victim, SpellNum)
                
                If Damage < 1 Then
                    Call SendSoundToMap(GetPlayerMap(Victim), Options.ResistSound)
                    SendActionMsg GetPlayerMap(Victim), "Resist", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
                Else
                    Call NpcAttackPlayer(MapNpcNum, Victim, Damage)
                    Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Victim)
                    DidCast = True
                End If
            End If
        End If
    End If
    
    If DidCast Then
        MapNpc(MapNum).NPC(MapNpcNum).Vital(MP) = MapNpc(MapNum).NPC(MapNpcNum).Vital(MP) - Spell(SpellNum).MPCost
        Call SendMapNpcVitals(MapNum, MapNpcNum)
        MapNpc(MapNum).NPC(MapNpcNum).AttackTimer = timeGetTime
        
        MapNpc(MapNum).NPC(MapNpcNum).SpellTimer(MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Spell) = timeGetTime + Spell(SpellNum).CDTime * 1000
        SendActionMsg MapNum, Trim$(Spell(SpellNum).Name), BrightBlue, ACTIONMSG_SCROLL, MapNpc(MapNum).NPC(MapNpcNum).X * 32, MapNpc(MapNum).NPC(MapNpcNum).Y * 32
    End If
End Sub

Sub NpcSpellNpc(ByVal MapNpcNum As Long, ByVal Victim As Long, MapNum As Integer)
    Dim i As Long
    Dim Damage As Long
    Dim SpellNum As Long
    Dim DidCast As Boolean, AoE As Long, X As Byte, Y As Byte

    ' Check for subscript out of range
    If MapNpcNum < 1 Or MapNpcNum > MAX_MAP_NPCS Or Victim < 1 Or Victim > MAX_MAP_NPCS Then Exit Sub
        
    ' Check for subscript out of range
    If MapNpc(MapNum).NPC(MapNpcNum).Num < 1 Or MapNpc(MapNum).NPC(Victim).Num < 1 Then Exit Sub
    
    ' Set the spell that they are going to cast
    SpellNum = MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Spell
   
    ' Send this packet so they can see the person attacking
    Call SendNpcAttack(MapNpcNum, MapNum)
    
    ' Play the sound
    Call SendMapSound(MapNum, Victim, MapNpc(MapNum).NPC(Victim).X, MapNpc(MapNum).NPC(Victim).Y, SoundEntity.seSpell, SpellNum)
    
    DidCast = False
    AoE = Spell(SpellNum).AoE
    X = MapNpc(MapNum).NPC(MapNpc(MapNum).NPC(MapNpcNum).Target).X
    Y = MapNpc(MapNum).NPC(MapNpc(MapNum).NPC(MapNpcNum).Target).Y
    
    ' Check if the spell they are going to cast is valid
    If SpellNum > 0 Then
        If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Or Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
            Damage = GetNpcSpellVital(MapNum, MapNpcNum, Victim, SpellNum, True)
            
            If Spell(SpellNum).IsAoe = True Then ' AoE
                If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                    For i = 1 To Map(MapNum).Npc_HighIndex
                        If MapNpc(MapNum).NPC(i).Num > 0 Then
                            If IsInRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                                MapNpc(MapNum).NPC(i).Vital(Vitals.HP) = MapNpc(MapNum).NPC(i).Vital(Vitals.HP) + Damage
                                SendActionMsg MapNum, "+" & Damage, BrightGreen, 1, (MapNpc(MapNum).NPC(i).X * 32), (MapNpc(MapNum).NPC(i).Y * 32)
                                
                                Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i)
                                DidCast = True
                                
                                ' Prevent overhealing
                                If MapNpc(MapNum).NPC(i).Vital(Vitals.HP) > GetNpcMaxVital(MapNpc(MapNum).NPC(i).Num, HP) Then
                                    MapNpc(MapNum).NPC(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(MapNum).NPC(i).Num, HP)
                                End If
                            End If
                        End If
                    Next
                Else
                    For i = 1 To Map(MapNum).Npc_HighIndex
                        If MapNpc(MapNum).NPC(i).Num > 0 Then
                            If IsInRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                                MapNpc(MapNum).NPC(i).Vital(Vitals.MP) = MapNpc(MapNum).NPC(i).Vital(Vitals.MP) + Damage
                                SendActionMsg MapNum, "+" & Damage, BrightBlue, 1, (MapNpc(MapNum).NPC(i).X * 32), (MapNpc(MapNum).NPC(i).Y * 32)
                                
                                Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i)
                                DidCast = True
                                
                                ' Prevent overhealing
                                If MapNpc(MapNum).NPC(i).Vital(Vitals.MP) > GetNpcMaxVital(MapNpc(MapNum).NPC(i).Num, MP) Then
                                    MapNpc(MapNum).NPC(i).Vital(Vitals.MP) = GetNpcMaxVital(MapNpc(MapNum).NPC(i).Num, MP)
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                    MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.HP) = MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.HP) + Damage
                    SendActionMsg MapNum, "+" & Damage, BrightGreen, 1, (MapNpc(MapNum).NPC(MapNpcNum).X * 32), (MapNpc(MapNum).NPC(MapNpcNum).Y * 32)
                    
                    ' Prevent overhealing
                    If MapNpc(MapNum).NPC(Victim).Vital(Vitals.HP) > GetNpcMaxVital(MapNpc(MapNum).NPC(Victim).Num, HP) Then
                        MapNpc(MapNum).NPC(Victim).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(MapNum).NPC(Victim).Num, HP)
                    End If
                ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                    MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.MP) = MapNpc(MapNum).NPC(MapNpcNum).Vital(Vitals.MP) + Damage
                    SendActionMsg MapNum, "+" & Damage, BrightBlue, 1, (MapNpc(MapNum).NPC(MapNpcNum).X * 32), (MapNpc(MapNum).NPC(MapNpcNum).Y * 32)
                    
                    ' Prevent overhealing
                    If MapNpc(MapNum).NPC(Victim).Vital(Vitals.MP) > GetNpcMaxVital(MapNpc(MapNum).NPC(Victim).Num, MP) Then
                        MapNpc(MapNum).NPC(Victim).Vital(Vitals.MP) = GetNpcMaxVital(MapNpc(MapNum).NPC(Victim).Num, MP)
                    End If
                End If
            End If
            
            Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Victim)
            DidCast = True
        Else
            If Spell(SpellNum).IsAoe = True Then ' AoE
                For i = 1 To Map(MapNum).Npc_HighIndex
                    If MapNpc(MapNum).NPC(i).Num > 0 Then
                        If IsInRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                            If CanNpcAttackNpc(MapNum, MapNpcNum, i, True) Then
                                Damage = GetNpcSpellVital(MapNum, MapNpcNum, i, SpellNum)
                                
                                If Damage < 1 Then
                                    Call SendSoundToMap(MapNum, Options.ResistSound)
                                    SendActionMsg MapNum, "Resist", Pink, 1, (MapNpc(MapNum).NPC(i).X * 32), (MapNpc(MapNum).NPC(i).Y * 32)
                                Else
                                    Call NpcAttackNpc(MapNum, MapNpcNum, i, Damage)
                                    Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i)
                                    DidCast = True
                                End If
                            End If
                        End If
                    End If
                Next
            Else ' Non AoE
                Damage = GetNpcSpellVital(MapNum, MapNpcNum, Victim, SpellNum)
                
                If Damage < 1 Then
                    Call SendSoundToMap(MapNum, Options.ResistSound)
                    SendActionMsg MapNum, "Resist", Pink, 1, (MapNpc(MapNum).NPC(MapNpc(MapNum).NPC(MapNpcNum).Target).X * 32), (MapNpc(MapNum).NPC(MapNpc(MapNum).NPC(MapNpcNum).Target).Y * 32)
                Else
                    Call NpcAttackNpc(MapNum, MapNpcNum, Victim, Damage)
                    Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Victim)
                    DidCast = True
                End If
            End If
        End If
    End If
    
    If DidCast Then
        MapNpc(MapNum).NPC(MapNpcNum).Vital(MP) = MapNpc(MapNum).NPC(MapNpcNum).Vital(MP) - Spell(SpellNum).MPCost
        Call SendMapNpcVitals(MapNum, MapNpcNum)
        MapNpc(MapNum).NPC(MapNpcNum).AttackTimer = timeGetTime
        
        MapNpc(MapNum).NPC(MapNpcNum).SpellTimer(MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Spell) = timeGetTime + Spell(SpellNum).CDTime * 1000
        SendActionMsg MapNum, Trim$(Spell(SpellNum).Name), BrightBlue, ACTIONMSG_SCROLL, MapNpc(MapNum).NPC(MapNpcNum).X * 32, MapNpc(MapNum).NPC(MapNpcNum).Y * 32
    End If
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################
Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long)
    Dim NpcNum As Long
    Dim MapNum As Integer
    Dim Damage As Long

    ' Can we attack the player
    If CanPlayerAttackPlayer(Attacker, Victim, False) Then
        MapNum = GetPlayerMap(Attacker)

        ' Get the damage we can do
        Damage = GetPlayerDamage(Attacker)
        
        ' Add damage based on direction
        If Account(Attacker).Chars(GetPlayerChar(Attacker)).Dir = DIR_UP Then
            If Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_LEFT Or Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_UP Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf Account(Attacker).Chars(GetPlayerChar(Attacker)).Dir = DIR_DOWN Then
            If Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_LEFT Or Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf Account(Attacker).Chars(GetPlayerChar(Attacker)).Dir = DIR_LEFT Then
            If Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_UP Or Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_LEFT Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf Account(Attacker).Chars(GetPlayerChar(Attacker)).Dir = DIR_RIGHT Then
            If Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_UP Or Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(Victim).Chars(GetPlayerChar(Victim)).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 4)
            End If
        End If
        
        ' 1.5 times the damage if it's a critical
        If CanPlayerCritical(Attacker) Then
            Damage = Damage * 1.5
            Call SendSoundToMap(MapNum, Options.CriticalSound)
            SendAnimation MapNum, Options.CriticalAnimation, 0, 0, TARGET_TYPE_PLAYER, Attacker
        End If
        
        ' Take away protection from the damage
        Damage = Damage - GetPlayerProtection(Victim)
        
        ' Randomize damage
        Damage = Random(Damage - (Damage / 2), Damage)
        
        Round Damage

        If Damage < 1 Then
            Call SendSoundToMap(MapNum, Options.MissSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_PLAYER, Victim
        End If
        
        Call PlayerAttackPlayer(Attacker, Victim, Damage)
    End If
End Sub

Function CanPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal IsSpell As Boolean = False, Optional ByVal UsingBow As Boolean) As Boolean
Dim WeaponSlot As Long
Dim Range As Byte
Dim DistanceToPlayer As Integer

    ' Check for subscript out of range
    If Not IsPlaying(Victim) Then Exit Function

    ' Can't attack yourself
    If Attacker = Victim Then Exit Function
    
    ' Make sure they are on the same map
    If Not GetPlayerMap(Attacker) = GetPlayerMap(Victim) Then Exit Function

    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Victim).GettingMap = YES Then Exit Function

    ' Make sure they have at least 1 HP
    If GetPlayerVital(Victim, Vitals.HP) < 1 Then Exit Function

    If Not IsSpell And Not UsingBow Then ' Melee attack
        ' Check if at same coordinates
        Select Case GetPlayerDir(Attacker)
            Case DIR_UP
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWN
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_LEFT
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_RIGHT
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If
        
    ' Check if map is attackable
    If Moral(Map(GetPlayerMap(Attacker)).Moral).CanPK = 0 Then
        If GetPlayerPK(Victim) = NO Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If
    
    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > STAFF_MODERATOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > STAFF_MODERATOR Then
        Call PlayerMsg(Attacker, "You cannot attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < Options.PKLevel Then
        Call PlayerMsg(Attacker, "You are below level " & Options.PKLevel & ", you cannot attack another player!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < Options.PKLevel Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & Options.PKLevel & ", you cannot attack this player!", BrightRed)
        Exit Function
    End If
    
    ' Don't attack a party member
    If TempPlayer(Attacker).InParty > 0 And TempPlayer(Victim).InParty > 0 Then
        If TempPlayer(Attacker).InParty = TempPlayer(Victim).InParty Then
            Call PlayerMsg(Attacker, "You can't attack another party member!", BrightRed)
            Exit Function
        End If
    End If
    
    ' Don't attack a guild member
    If GetPlayerGuild(Attacker) > 0 Then
        If GetPlayerGuild(Attacker) = GetPlayerGuild(Victim) Then
            Call PlayerMsg(Attacker, "You can't attack another guild member!", BrightRed)
            Exit Function
        End If
    End If
    
    ' Adjust target if they have none
    If TempPlayer(Victim).Target = 0 Then
        TempPlayer(Victim).Target = Attacker
        TempPlayer(Victim).TargetType = TARGET_TYPE_PLAYER
        Call SendPlayerTarget(Victim)
    End If
    
    If Not IsSpell Then
        ' Set the attack's target
        TempPlayer(Attacker).TargetType = TARGET_TYPE_PLAYER
        TempPlayer(Attacker).Target = Victim
        Call SendPlayerTarget(Attacker)
    
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If timeGetTime < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).WeaponSpeed Then Exit Function
        Else
            If timeGetTime < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If
    
    If CanPlayerMitigatePlayer(Attacker, Victim) = True Or TempPlayer(Attacker).SpellBuffer.Spell > 0 Then
        ' Check if player can avoid the attack
        If CanPlayerDodge(Victim) Then
            Call SendSoundToMap(GetPlayerMap(Victim), Options.DodgeSound)
            SendAnimation GetPlayerMap(Victim), Options.DodgeAnimation, 0, 0, TARGET_TYPE_PLAYER, Victim
            Exit Function
        End If
        
        ' Check if player can deflect the attack
        If CanPlayerDeflect(Victim) Then
            Call SendSoundToMap(GetPlayerMap(Victim), Options.DeflectSound)
            SendAnimation GetPlayerMap(Victim), Options.DeflectAnimation, 0, 0, TARGET_TYPE_PLAYER, Victim
            Exit Function
        End If
        
        ' Check if player can block the attack
        If CanPlayerBlock(Victim) Then
            Call SendSoundToMap(GetPlayerMap(Victim), Options.BlockSound)
            SendAnimation GetPlayerMap(Victim), Options.DeflectAnimation, 0, 0, TARGET_TYPE_PLAYER, Victim
            Exit Function
        End If
    End If

    CanPlayerAttackPlayer = True
End Function

Sub PlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal OverTime As Boolean = False)
    Dim Exp As Long
    Dim n As Long
    Dim i As Long
    Dim LevelDiff As Byte

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or IsPlaying(Victim) = False Or Damage < 0 Then Exit Sub

    ' Check for weapon
    n = 0

    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        n = GetPlayerEquipment(Attacker, Weapon)
    End If
    
    ' Set the regen timer
    TempPlayer(Attacker).StopRegen = True
    TempPlayer(Attacker).StopRegenTimer = timeGetTime
    
    ' Send the sound
    If SpellNum > 0 Then
        Call SendMapSound(GetPlayerMap(Victim), Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, SpellNum)
    Else
        Call SendMapSound(GetPlayerMap(Victim), Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seAnimation, 1)
    End If
    
    ' Reduce durability on the victim's equipment
    If Random(1, 2) = 1 Then ' Which one it will affect
        Call DamagePlayerEquipment(Victim, Body)
    Else
        Call DamagePlayerEquipment(Victim, Head)
    End If
    
    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Set the damage to the player's health so that it doesn't appear that it's overkilling it
        Damage = Damage = GetPlayerVital(Victim, Vitals.HP)
          
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' Player is dead
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & GetPlayerName(Attacker) & "!", BrightRed)
        
        ' Calculate exp to give attacker
        Exp = (GetPlayerExp(Victim) \ 10)
        
        ' Find the level difference between both the players
        LevelDiff = GetPlayerLevel(Attacker) - GetPlayerLevel(Victim)
        
        If Exp > 0 Then
            If LevelDiff > 0 Then
                Exp = Exp / (Exp / (LevelDiff * 10))
            ElseIf LevelDiff < 0 Then
                Exp = Exp + (Exp * (LevelDiff * -2.5))
            End If
        End If
        
        ' Adjust the exp based on the server's rate for experience
        Exp = Exp * EXP_RATE
        
        ' Randomize
        Exp = Random(Exp * 0.95, Exp * 1.05)

        Round Exp

        ' Make sure we dont get less then 0
        If Exp < 0 Then Exp = 0

        If Exp = 0 Or Moral(Map(GetPlayerMap(Attacker)).Moral).LoseExp = 0 Or GetPlayerLevel(Victim) < MAX_LEVEL Then
            SendActionMsg GetPlayerMap(Attacker), "+0 Exp", White, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
            Call PlayerMsg(Victim, "You did not lose any experience.", Grey)
        Else
            Call SetPlayerExp(Victim, GetPlayerExp(Victim) - Exp)
            SendPlayerExp Victim
            Call PlayerMsg(Victim, "You lost " & Exp & " experience.", BrightRed)
            
            ' Check if we're in a party
            If TempPlayer(Attacker).InParty > 0 Then
                ' Pass through party exp share function
                Party_ShareExp TempPlayer(Attacker).InParty, Exp, Attacker
            ElseIf GetPlayerLevel(Attacker) < MAX_LEVEL Then
                ' Not in party, get exp for self
                Call SetPlayerExp(Attacker, GetPlayerExp(Attacker) + Exp)
                
                SendPlayerExp Attacker
                SendActionMsg GetPlayerMap(Attacker), "+" & Exp & " Exp", White, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
                
                ' Check if we've leveled
                CheckPlayerLevelUp Attacker
            Else
                SendActionMsg GetPlayerMap(Attacker), "+0 Exp", White, 1, (GetPlayerX(Attacker) * 32), (GetPlayerY(Attacker) * 32)
            End If
        End If
        
        ' Purge target info of anyone who targetted dead player
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If Account(i).Chars(GetPlayerChar(i)).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(i).TargetType = TARGET_TYPE_PLAYER Then
                        If TempPlayer(i).Target = Victim Then
                            TempPlayer(i).Target = 0
                            TempPlayer(i).TargetType = TARGET_TYPE_NONE
                            SendPlayerTarget i
                        End If
                    End If
                End If
            End If
        Next

        If GetPlayerPK(Victim) = NO Then
            If GetPlayerPK(Attacker) = NO Then
                Call SetPlayerPK(Attacker, YES)
                Call SendPlayerPK(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Killer!", BrightRed)
            End If
        Else
            Call GlobalMsg(GetPlayerName(Victim) & " has paid the price for being a Player Killer!", BrightRed)
        End If
        
        Call OnDeath(Victim, Attacker)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' Send animation
        If SpellNum < 1 Then
            If Item(GetPlayerEquipment(Attacker, Weapon)).Animation > 0 Then
                If n > 0 Then
                    If Not OverTime Then
                        Call SendAnimation(GetPlayerMap(Victim), Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
                    End If
                End If
            Else
                Call SendAnimation(GetPlayerMap(Victim), 1, 0, 0, TARGET_TYPE_PLAYER, Victim)
            End If
        End If
        
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        SendBlood GetPlayerMap(Victim), GetPlayerX(Victim), GetPlayerY(Victim)
        
        ' Set the regen timer
        TempPlayer(Victim).StopRegen = True
        TempPlayer(Victim).StopRegenTimer = timeGetTime
        
        ' If a stunning spell, stun the player
        If SpellNum > 0 Then
            If Spell(SpellNum).StunDuration > 0 Then StunPlayer Victim, SpellNum
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Player Victim, SpellNum, Attacker
            End If
        End If
    End If

    ' Reset attack timer
    TempPlayer(Attacker).AttackTimer = timeGetTime
    
    ' Reduce durability of weapon
    Call DamagePlayerEquipment(Attacker, Weapon)
End Sub

' ############
' ## Spells ##
' ############
Public Sub BufferPlayerSpell(ByVal index As Long, ByVal SpellSlot As Byte)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Integer
    Dim SpellCastType As Byte
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    Dim TargetType As Byte
    Dim Target As Long
    
    ' Prevent subscript out of range
    If SpellSlot < 1 Or SpellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    SpellNum = GetPlayerSpell(index, SpellSlot)
    
    If SpellNum < 1 Or SpellNum > MAX_SPELLS Then Exit Sub
    
    MapNum = GetPlayerMap(index)
    
    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then Exit Sub
    
    ' See if cooldown has finished
    If Account(index).Chars(GetPlayerChar(index)).SpellCD(SpellSlot) > timeGetTime Then
        PlayerMsg index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(index, Vitals.MP) < MPCost Then
        Call PlayerMsg(index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        Call PlayerMsg(index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(SpellNum).AccessReq
    
    ' Make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        Call PlayerMsg(index, "You must be a staff member to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(SpellNum).ClassReq
    
    ' Make sure the ClassReq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(index) Then
            Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' Can't use items while in a map that doesn't allow it
    If Moral(Map(GetPlayerMap(index)).Moral).CanCast = 0 Then Exit Sub

    ' Find out what kind of spell it is (Self cast, Target or AOE)
    If Spell(SpellNum).Range > 0 Then
        ' Ranged attack, single target or aoe?
        If Not Spell(SpellNum).IsAoe Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Not Spell(SpellNum).IsAoe Then
            SpellCastType = 0 ' Self-cast
        Else
            SpellCastType = 1 ' Self-cast AoE
        End If
    End If
    
    TargetType = TempPlayer(index).TargetType
    Target = TempPlayer(index).Target
    Range = Spell(SpellNum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' Self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' Check if have target
            If Not Target > 0 Then Exit Sub

            If TargetType = TARGET_TYPE_PLAYER Then
                ' If have target, check in range
                If Not IsInRange(Range, GetPlayerX(index), GetPlayerY(index), GetPlayerX(Target), GetPlayerY(Target)) Then
                    PlayerMsg index, "Target is not in range!", BrightRed
                Else
                    ' Go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(index, Target, False, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf TargetType = TARGET_TYPE_NPC Then
                ' If have target, check in range
                If Not IsInRange(Range, GetPlayerX(index), GetPlayerY(index), MapNpc(MapNum).NPC(Target).X, MapNpc(MapNum).NPC(Target).Y) Then
                    PlayerMsg index, "Target is not in range!", BrightRed
                Else
                    ' Go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNpc(index, Target, False, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, index
        
        If Spell(SpellNum).CastTime > 0 Then
            SendActionMsg MapNum, "Casting " & Trim$(Spell(SpellNum).Name), BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        End If
        
        TempPlayer(index).SpellBuffer.Spell = SpellSlot
        TempPlayer(index).SpellBuffer.Timer = timeGetTime
        TempPlayer(index).SpellBuffer.Target = TempPlayer(index).Target
        TempPlayer(index).SpellBuffer.TType = TempPlayer(index).TargetType
    End If
End Sub

Public Sub CastSpell(ByVal index As Long, ByVal SpellSlot As Byte, ByVal Target As Long, ByVal TargetType As Byte)
    Dim SpellNum As Long
    Dim MapNum As Integer
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim i As Long
    Dim AoE As Long
    Dim Range As Byte
    Dim VitalType As Byte
    Dim Increment As Boolean
    Dim X As Long, Y As Long
    Dim SpellCastType As Long
    Dim MPCost As Long

    ' Prevent subscript out of range
    If SpellSlot < 1 Or SpellSlot > MAX_PLAYER_SPELLS Then Exit Sub

    SpellNum = GetPlayerSpell(index, SpellSlot)
    MapNum = GetPlayerMap(index)
    MPCost = Spell(SpellNum).MPCost
    
    ' Make sure player has the spell
    If Not HasSpell(index, SpellNum) Then Exit Sub
    
    ' Make sure they meet the requirements
    If CanPlayerCastSpell(index, SpellNum) = False Then Exit Sub
    
    ' Find out what kind of spell it is, self cast, target, or AoE
    If Spell(SpellNum).Range > 0 Then
        ' Ranged attack, single target or aoe?
        If Spell(SpellNum).IsAoe = False Then
            SpellCastType = 2 ' targetted
        Else
            SpellCastType = 3 ' targetted aoe
        End If
    Else
        If Spell(SpellNum).IsAoe = False Then
            SpellCastType = 0 ' Self-cast
        Else
            SpellCastType = 1 ' Self-cast AoE
        End If
    End If
    
    ' Set the vital
    If Spell(SpellNum).WeaponDamage = True Then
        Vital = Spell(SpellNum).Vital + GetPlayerDamage(index)
    Else
        Vital = Spell(SpellNum).Vital
    End If
    
    AoE = Spell(SpellNum).AoE
    Range = Spell(SpellNum).Range
    
    ' Add damage based on intelligence
    Vital = Vital + GetPlayerStat(index, Intelligence) / 3
    
    ' Randomize the vital
    Vital = Random(Vital - (Vital / 2), Vital)
    
    ' 1.5 times the damage if it's a critical
    If CanPlayerSpellCritical(index) Then
        Vital = Vital * 1.5
        Call SendSoundToMap(MapNum, Options.CriticalSound)
        SendAnimation MapNum, Options.CriticalAnimation, 0, 0, TARGET_TYPE_PLAYER, index
    End If
    
    Select Case SpellCastType
        Case 0 ' Self-cast target
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_HEALHP
                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellPlayer_Effect Vitals.HP, True, index, Vital, SpellNum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    SpellPlayer_Effect Vitals.MP, True, index, Vital, SpellNum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    PlayerWarp index, Spell(SpellNum).Map, Spell(SpellNum).X, Spell(SpellNum).Y
                    DidCast = True
                Case SPELL_TYPE_RECALL
                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
                    WarpToCheckPoint (index)
                    DidCast = True
                Case SPELL_TYPE_WARPTOTARGET
                    Call PlayerMsg(index, "This spell has been made incorrectly, report this to a staff member!", BrightRed)
                    Exit Sub
            End Select
            
        Case 1, 3 ' Self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                X = GetPlayerX(index)
                Y = GetPlayerY(index)
            ElseIf SpellCastType = 3 Then
                If TargetType = 0 Then Exit Sub
                If Target = 0 Then Exit Sub
                
                If TargetType = TARGET_TYPE_PLAYER Then
                    X = GetPlayerX(Target)
                    Y = GetPlayerY(Target)
                Else
                    X = MapNpc(MapNum).NPC(Target).X
                    Y = MapNpc(MapNum).NPC(Target).Y
                End If
                
                If Not IsInRange(Range, GetPlayerX(index), GetPlayerY(index), X, Y) Then
                    PlayerMsg index, "Target is not in range!", BrightRed
                    SendClearAccountSpellBuffer index
                End If
            End If
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If Not i = index Then
                                If GetPlayerMap(i) = GetPlayerMap(index) Then
                                    If IsInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                        If CanPlayerAttackPlayer(index, i, False, True) Then
                                            SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, i
                                            PlayerAttackPlayer index, i, Vital, SpellNum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To Map(MapNum).Npc_HighIndex
                        If MapNpc(MapNum).NPC(i).Num > 0 Then
                            If MapNpc(MapNum).NPC(i).Vital(HP) > 0 Then
                                If IsInRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                                    ' Friendly and Shopkeeper
                                    If Not NPC(MapNpc(MapNum).NPC(i).Num).Behavior = NPC_BEHAVIOR_FRIENDLY And Not NPC(MapNpc(MapNum).NPC(i).Num).Behavior = NPC_BEHAVIOR_SHOPKEEPER And Not NPC(MapNpc(MapNum).NPC(i).Num).Behavior = NPC_BEHAVIOR_QUEST And Not NPC(MapNpc(MapNum).NPC(i).Num).Behavior = NPC_BEHAVIOR_GUIDE Then
                                        ' Guard
                                        If Not NPC(MapNpc(MapNum).NPC(i).Num).Behavior = NPC_BEHAVIOR_GUARD Or (NPC(MapNpc(MapNum).NPC(i).Num).Behavior = NPC_BEHAVIOR_GUARD And GetPlayerPK(index) = YES) Then
                                            If CanPlayerAttackNpc(index, i, False, True) Then
                                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, i
                                                PlayerAttackNpc index, i, Vital, SpellNum
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                Case SPELL_TYPE_HEALHP, SPELL_TYPE_HEALMP, SPELL_TYPE_DAMAGEMP
                    If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        Increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        Increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        Increment = False
                    End If
                    
                    DidCast = True
                    
                    For i = 1 To Player_HighIndex
                        If IsPlaying(i) Then
                            If GetPlayerMap(i) = GetPlayerMap(index) Then
                                If IsInRange(AoE, X, Y, GetPlayerX(i), GetPlayerY(i)) Then
                                    SpellPlayer_Effect VitalType, Increment, i, Vital, SpellNum
                                End If
                            End If
                        End If
                    Next
                    
                    For i = 1 To Map(MapNum).Npc_HighIndex
                        If MapNpc(MapNum).NPC(i).Num > 0 Then
                            If (Increment = True And NPC(MapNpc(MapNum).NPC(i).Num).Behavior = NPC_BEHAVIOR_GUARD And Account(index).Chars(GetPlayerChar(index)).PK = NO) Or Increment = False Then
                                If MapNpc(MapNum).NPC(i).Vital(HP) > 0 Then
                                    If IsInRange(AoE, X, Y, MapNpc(MapNum).NPC(i).X, MapNpc(MapNum).NPC(i).Y) Then
                                        SpellNpc_Effect VitalType, Increment, i, Vital, SpellNum, MapNum
                                    End If
                                End If
                            End If
                        End If
                    Next
            End Select
            
        Case 2 ' Targetted
            If TargetType = 0 Then Exit Sub
            If Target = 0 Then Exit Sub
            
            If TargetType = TARGET_TYPE_PLAYER Then
                X = GetPlayerX(Target)
                Y = GetPlayerY(Target)
            Else
                X = MapNpc(MapNum).NPC(Target).X
                Y = MapNpc(MapNum).NPC(Target).Y
            End If
            
            If Not IsInRange(Range, GetPlayerX(index), GetPlayerY(index), X, Y) Then
                SendClearAccountSpellBuffer index
                Exit Sub
            End If
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If CanPlayerAttackPlayer(index, Target, False, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Target
                                PlayerAttackPlayer index, Target, Vital, SpellNum
                                DidCast = True
                            Else
                                Call SendSoundToMap(GetPlayerMap(i), Options.ResistSound)
                                SendActionMsg GetPlayerMap(i), "Resist", Pink, 1, (Account(Target).Chars(GetPlayerChar(Target)).X * 32), (Account(Target).Chars(GetPlayerChar(Target)).Y * 32)
                            End If
                        End If
                    Else
                        If CanPlayerAttackNpc(index, Target, False, True) Then
                            If Vital > 0 Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Target
                                PlayerAttackNpc index, Target, Vital, SpellNum
                                DidCast = True
                            Else
                                Call SendSoundToMap(MapNum, Options.ResistSound)
                                SendActionMsg MapNum, "Resist", Pink, 1, (MapNpc(MapNum).NPC(Target).X * 32), (MapNpc(MapNum).NPC(Target).Y * 32)
                            End If
                        End If
                    End If
    
                Case SPELL_TYPE_DAMAGEMP, SPELL_TYPE_HEALMP, SPELL_TYPE_HEALHP
                    If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                        VitalType = Vitals.MP
                        Increment = False
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                        VitalType = Vitals.HP
                        Increment = True
                    ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                        VitalType = Vitals.MP
                        Increment = True
                    End If
                    
                    DidCast = True
                    
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(index, Target, False, True) Then
                                SpellPlayer_Effect VitalType, Increment, Target, Vital, SpellNum
                            End If
                        Else
                            SpellPlayer_Effect VitalType, Increment, Target, Vital, SpellNum
                        End If
                    ElseIf TargetType = TARGET_TYPE_NPC And Increment = False Or NPC(MapNpc(MapNum).NPC(Target).Num).Behavior = NPC_BEHAVIOR_GUARD And Account(index).Chars(GetPlayerChar(index)).PK = NO Then
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNpc(index, Target, False, True) Then
                                SpellNpc_Effect VitalType, Increment, Target, Vital, SpellNum, MapNum
                            End If
                        Else
                            SpellNpc_Effect VitalType, Increment, Target, Vital, SpellNum, MapNum
                        End If
                    Else
                        Call PlayerMsg(index, "You are unable to cast your spell on this target!", 12)
                        Exit Sub
                    End If
                    
                Case SPELL_TYPE_WARPTOTARGET
                    Call PlayerWarp(index, MapNum, X, Y)
            End Select
    End Select
    
    If DidCast Then
        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) - MPCost)
        Call SendVital(index, Vitals.MP)
        TempPlayer(index).SpellBuffer.Timer = timeGetTime + (Spell(SpellNum).CDTime * 1000)
        Account(index).Chars(GetPlayerChar(index)).SpellCD(SpellSlot) = timeGetTime + (Spell(SpellNum).CDTime * 1000)
        Call SendSpellCooldown(index, SpellSlot)
        SendActionMsg MapNum, Trim$(Spell(SpellNum).Name), BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        ' Set the sprite
        If Spell(SpellNum).Sprite > 0 Then
            Call SetPlayerSprite(index, Spell(SpellNum).Sprite)
            Call SendPlayerSprite(index)
        End If
        
        If Spell(SpellNum).NewSpell > 0 And Spell(SpellNum).NewSpell <= MAX_SPELLS Then
            If Spell(Spell(SpellNum).NewSpell).CastRequired > 0 Then
                ' Add 1 to the amount of casts
                Account(index).Chars(GetPlayerChar(index)).AmountOfCasts(SpellSlot) = Account(index).Chars(GetPlayerChar(index)).AmountOfCasts(SpellSlot) + 1
                
                ' Check if a spell can rank up
                Call CheckSpellRankUp(index, SpellNum, SpellSlot)
            End If
        End If
    End If
    
    Call ClearAccountSpellBuffer(index)
    Call SendClearAccountSpellBuffer(index)
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal Increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal SpellNum As Long)
    Dim sSymbol As String * 1
    Dim Color As Long

    ' Randomize damage
    Damage = Random(Damage - (Damage / 2), Damage)

    If Damage > 0 Then
        If Increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Color = BrightGreen
            If Vital = Vitals.MP Then Color = BrightBlue
        Else
            sSymbol = "-"
            If Vital = Vitals.HP Then Color = Red
            If Vital = Vitals.MP Then Color = Blue
        End If
    
        SendAnimation GetPlayerMap(index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, index
        SendActionMsg GetPlayerMap(index), sSymbol & Damage, Color, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
        
        ' Send the sound
        SendMapSound GetPlayerMap(index), index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seSpell, SpellNum
        
        If Increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) + Damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Player index, SpellNum
            End If
        ElseIf Not Increment Then
            SetPlayerVital index, Vital, GetPlayerVital(index, Vital) - Damage
        End If
    End If
End Sub

Public Sub SpellNpc_Effect(ByVal Vital As Byte, ByVal Increment As Boolean, ByVal index As Long, ByVal Damage As Long, ByVal SpellNum As Long, ByVal MapNum As Integer)
    Dim sSymbol As String * 1
    Dim Color As Long

    ' Randomize damage
    Damage = Random(Damage - (Damage / 2), Damage)

    If Damage > 0 Then
        If Increment Then
            sSymbol = "+"
            If Vital = Vitals.HP Then Color = BrightGreen
            If Vital = Vitals.MP Then Color = BrightBlue
        Else
            sSymbol = "-"
            If Vital = Vitals.HP Then Color = Red
            If Vital = Vitals.MP Then Color = Blue
        End If
    
        SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, index
        SendActionMsg MapNum, sSymbol & Damage, Color, ACTIONMSG_SCROLL, MapNpc(MapNum).NPC(index).X * 32, MapNpc(MapNum).NPC(index).Y * 32
        
        ' Send the sound
        SendMapSound MapNum, index, MapNpc(MapNum).NPC(index).X, MapNpc(MapNum).NPC(index).Y, SoundEntity.seSpell, SpellNum
        
        If Increment Then
            MapNpc(MapNum).NPC(index).Vital(Vital) = MapNpc(MapNum).NPC(index).Vital(Vital) + Damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Npc MapNum, index, SpellNum
            End If
        ElseIf Not Increment Then
            MapNpc(MapNum).NPC(index).Vital(Vital) = MapNpc(MapNum).NPC(index).Vital(Vital) - Damage
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Player(ByVal index As Long, ByVal SpellNum As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With TempPlayer(index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = timeGetTime
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddDoT_Npc(ByVal MapNum As Integer, ByVal index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).NPC(index).DoT(i)
            If .Spell = SpellNum Then
                .Timer = timeGetTime
                .Caster = Caster
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = timeGetTime
                .Caster = Caster
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub AddHoT_Npc(ByVal MapNum As Integer, ByVal index As Long, ByVal SpellNum As Long)
    Dim i As Long

    For i = 1 To MAX_DOTS
        With MapNpc(MapNum).NPC(index).HoT(i)
            If .Spell = SpellNum Then
                .Timer = timeGetTime
                .StartTime = timeGetTime
                Exit Sub
            End If
            
            If .Used = False Then
                .Spell = SpellNum
                .Timer = timeGetTime
                .Used = True
                .StartTime = timeGetTime
                Exit Sub
            End If
        End With
    Next
End Sub

Public Sub HandleDoT_Player(ByVal index As Long, ByVal dotNum As Long)
    With TempPlayer(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' Time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackPlayer(.Caster, index, True) Then
                    PlayerAttackPlayer .Caster, index, Spell(.Spell).Vital
                End If
                .Timer = timeGetTime
                ' Check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy DoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Player(ByVal index As Long, ByVal hotNum As Long)
    With TempPlayer(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' Time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg Account(index).Chars(GetPlayerChar(index)).Map, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, Account(index).Chars(GetPlayerChar(index)).X * 32, Account(index).Chars(GetPlayerChar(index)).Y * 32
                Account(index).Chars(GetPlayerChar(index)).Vital(Vitals.HP) = Account(index).Chars(GetPlayerChar(index)).Vital(Vitals.HP) + Spell(.Spell).Vital
                .Timer = timeGetTime
                ' Check if DoT is still active - if player died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' destroy hoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleDoT_Npc(ByVal MapNum As Integer, ByVal index As Long, ByVal dotNum As Long)
    With MapNpc(MapNum).NPC(index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' Time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                If CanPlayerAttackNpc(.Caster, index, True) Then
                    PlayerAttackNpc .Caster, index, Spell(.Spell).Vital, , True
                End If
                .Timer = timeGetTime
                ' Check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' Destroy DoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub HandleHoT_Npc(ByVal MapNum As Integer, ByVal index As Long, ByVal hotNum As Long)
    With MapNpc(MapNum).NPC(index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' Time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendActionMsg MapNum, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, MapNpc(MapNum).NPC(index).X * 32, MapNpc(MapNum).NPC(index).Y * 32
                MapNpc(MapNum).NPC(index).Vital(Vitals.HP) = MapNpc(MapNum).NPC(index).Vital(Vitals.HP) + Spell(.Spell).Vital
                .Timer = timeGetTime
                ' Check if DoT is still active - if NPC died it'll have been purged
                If .Used And .Spell > 0 Then
                    ' Destroy hoT if finished
                    If timeGetTime - .StartTime >= (Spell(.Spell).Duration * 1000) Then
                        .Used = False
                        .Spell = 0
                        .Timer = 0
                        .Caster = 0
                        .StartTime = 0
                    End If
                End If
            End If
        End If
    End With
End Sub

Public Sub StunPlayer(ByVal index As Long, ByVal SpellNum As Long)
    ' Check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' Set the values on Index
        TempPlayer(index).StunDuration = Spell(SpellNum).StunDuration
        TempPlayer(index).StunTimer = timeGetTime
        ' Send it to the Index
        SendStunned index
        ' Tell him he's stunned
        SendActionMsg GetPlayerMap(index), "Stunned", RGB(255, 128, 0), 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
    End If
End Sub

Public Sub StunNPC(ByVal index As Long, ByVal MapNum As Integer, ByVal SpellNum As Long)
    Dim NpcNum As Long
    
    NpcNum = MapNpc(MapNum).NPC(index).Num
    
    ' Check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' Set the values on Index
        MapNpc(MapNum).NPC(index).StunDuration = Spell(SpellNum).StunDuration
        MapNpc(MapNum).NPC(index).StunTimer = timeGetTime
    End If
    
     ' Tell other players its stunned
    SendActionMsg MapNum, "Stunned", RGB(255, 128, 0), 1, (MapNpc(MapNum).NPC(index).X * 32), (MapNpc(MapNum).NPC(index).Y * 32)
End Sub

Public Sub ClearNpcSpellBuffer(ByVal MapNum As Integer, ByVal MapNpcNum As Byte)
    MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Spell = 0
    MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Timer = 0
    MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Target = 0
    MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.TType = 0
End Sub

Public Sub ClearAccountSpellBuffer(ByVal index As Long)
    TempPlayer(index).SpellBuffer.Spell = 0
    TempPlayer(index).SpellBuffer.Timer = 0
    TempPlayer(index).SpellBuffer.Target = 0
    TempPlayer(index).SpellBuffer.TType = 0
End Sub

Private Function CanNpcHealSelf(ByVal MapNum As Integer, ByVal MapNpcNum As Byte, ByVal SpellNum As Long) As Boolean
    Dim NpcNum As Long, Target As Long, TargetType As Byte
    
    NpcNum = MapNpc(MapNum).NPC(MapNpcNum).Num
    Target = MapNpcNum
    TargetType = TARGET_TYPE_NPC
    
    ' Valid spell
    If NPC(NpcNum).Spell(SpellNum) > 0 And NPC(NpcNum).Spell(SpellNum) <= MAX_SPELLS Then
        ' Check for cooldown
        If MapNpc(MapNum).NPC(MapNpcNum).SpellTimer(MapNpcNum) <= timeGetTime Then
            ' Have enough mana
            If MapNpc(MapNum).NPC(MapNpcNum).Vital(MP) - Spell(SpellNum).MPCost >= 0 Or Spell(SpellNum).MPCost = 0 Then
                If Spell(NPC(NpcNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALHP Or Spell(NPC(NpcNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALMP Then
                    ' Don't want to overheal
                    ' If Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD And TargetType = TARGET_TYPE_PLAYER Then
                    '    If Spell(Npc(NpcNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALHP Then
                    '        If Target > 0 And Target <= MAX_PLAYERS Then
                    '            If Account(Target).Chars(GetPlayerChar(Target)).Vital(HP) + GetNpcSpellVital(MapNum, MapNpcNum, Target, Npc(MapNpc(MapNum).Npc(Target).Num).Spell(SpellNum), True) > GetPlayerMaxVital(MapNpc(MapNum).Npc(Target).Num, HP) Then Exit Function
                    '        End If
                    '    ElseIf Spell(Npc(NpcNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALMP Then
                    '        If Target > 0 And Target <= MAX_PLAYERS Then
                    '            If Account(Target).Chars(GetPlayerChar(Target)).Vital(MP) + GetNpcSpellVital(MapNum, MapNpcNum, Target, Npc(MapNpc(MapNum).Npc(Target).Num).Spell(SpellNum), True) > GetPlayerMaxVital(MapNpc(MapNum).Npc(Target).Num, MP) Then Exit Function
                    '        End If
                    '    End If
                    ' If TargetType = TARGET_TYPE_NPC And Not Npc(NpcNum).Behavior = NPC_BEHAVIOR_GUARD Then
                        If Spell(NPC(NpcNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALHP Then
                            If Target = MapNpcNum Then
                                If MapNpc(MapNum).NPC(MapNpcNum).Vital(HP) + GetNpcSpellVital(MapNum, MapNpcNum, MapNpcNum, NPC(NpcNum).Spell(SpellNum), True) > GetNpcMaxVital(NpcNum, HP) Then Exit Function
                            ElseIf Target > 0 And Target <= MAX_MAP_NPCS Then
                                If MapNpc(MapNum).NPC(Target).Vital(HP) + GetNpcSpellVital(MapNum, MapNpcNum, Target, NPC(MapNpc(MapNum).NPC(Target).Num).Spell(SpellNum), True) > GetNpcMaxVital(MapNpc(MapNum).NPC(Target).Num, HP) Then Exit Function
                            End If
                        ElseIf Spell(NPC(NpcNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALMP Then
                            If Target = MapNpcNum Then
                                If MapNpc(MapNum).NPC(MapNpcNum).Vital(MP) + GetNpcSpellVital(MapNum, MapNpcNum, MapNpcNum, NPC(NpcNum).Spell(SpellNum), True) > GetNpcMaxVital(NpcNum, MP) Then Exit Function
                            ElseIf Target > 0 And Target <= MAX_MAP_NPCS Then
                                If MapNpc(MapNum).NPC(Target).Vital(MP) + GetNpcSpellVital(MapNum, MapNpcNum, Target, NPC(MapNpc(MapNum).NPC(Target).Num).Spell(SpellNum), True) > GetNpcMaxVital(MapNpc(MapNum).NPC(Target).Num, MP) Then Exit Function
                            End If
                        End If
                    ' End If
                    
                    CanNpcHealSelf = True
                    Exit Function
                End If
            End If
        End If
    End If
End Function

Private Function CanNpcCastSpell(ByVal MapNum As Integer, ByVal MapNpcNum As Byte) As Boolean
    Dim i As Long, Target As Long, TargetType As Byte, Range As Byte, NpcNum As Long
    Dim RndNum As Byte
    
    TargetType = MapNpc(MapNum).NPC(MapNpcNum).TargetType
    Target = MapNpc(MapNum).NPC(MapNpcNum).Target
    NpcNum = MapNpc(MapNum).NPC(MapNpcNum).Num
    
    CanNpcCastSpell = False

    ' Self-healing mechanics
    ' If MapNpc(MapNum).Npc(MapNpcNum).Vital(HP) < GetNpcMaxVital(MapNpc(MapNum).Npc(MapNpcNum).Num, HP) / 2 Then
    '    For i = 1 To MAX_NPC_SPELLS
    '        If CanNpcHealSelf(MapNum, MapNpcNum, i) Then
    '            MapNpc(MapNum).Npc(MapNpcNum).Target = MapNpcNum
    '            MapNpc(MapNum).Npc(MapNpcNum).TargetType = TARGET_TYPE_NPC
    '            CanNpcCastSpell = True
    '            MapNpc(MapNum).Npc(MapNpcNum).ActiveSpell = Npc(MapNpc(MapNum).Npc(MapNpcNum).Num).Spell(i)
    '            Exit Function
    '        End If
    '    Next
    ' End If
    
    For i = 1 To MAX_NPC_SPELLS
        ' Valid spell
        If NPC(NpcNum).Spell(i) > 0 And NPC(NpcNum).Spell(i) <= MAX_SPELLS Then
            ' Check for cooldown
            If MapNpc(MapNum).NPC(MapNpcNum).SpellTimer(i) <= timeGetTime Then
                ' Have enough mana
                If MapNpc(MapNum).NPC(MapNpcNum).Vital(MP) - Spell(NPC(NpcNum).Spell(i)).MPCost >= 0 Or Spell(NPC(NpcNum).Spell(i)).MPCost = 0 Then
                    Range = Spell(NPC(NpcNum).Spell(i)).Range
                    
                    ' Are they in range
                    If TargetType = TARGET_TYPE_PLAYER Then
                        If Target > 0 And Target <= MAX_PLAYERS Then
                            If Not IsInRange(Range, GetPlayerX(Target), GetPlayerY(Target), MapNpc(MapNum).NPC(MapNpcNum).X, MapNpc(MapNum).NPC(MapNpcNum).Y) Then Exit Function
                        End If
                    ElseIf TargetType = TARGET_TYPE_NPC Then
                        If Target > 0 And Target <= MAX_MAP_NPCS Then
                            If Not Target = MapNpcNum Then
                                If Not IsInRange(Range, MapNpc(MapNum).NPC(Target).X, MapNpc(MapNum).NPC(Target).Y, MapNpc(MapNum).NPC(MapNpcNum).X, MapNpc(MapNum).NPC(MapNpcNum).Y) Then Exit Function
                            End If
                        End If
                    End If
                    
                    CanNpcCastSpell = True
                    MapNpc(MapNum).NPC(MapNpcNum).ActiveSpell = NPC(MapNpc(MapNum).NPC(MapNpcNum).Num).Spell(i)
                    
                    ' Add random exits
                    If Random(1, MAX_NPC_SPELLS) = 1 Then Exit Function
                End If
            End If
        End If
    Next
End Function
