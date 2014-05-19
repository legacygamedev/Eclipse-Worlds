Attribute VB_Name = "modCombat"
Option Explicit

' ################################
' ##      Basic Calculations    ##
' ################################
Function GetPlayerMaxVital(ByVal Index As Long, ByVal Vital As Vitals) As Long
    If Index > Player_HighIndex Or Index < 1 Then Exit Function

    Select Case Vital
        Case HP
            Select Case Class(GetPlayerClass(Index)).CombatTree
                Case 1 ' Melee
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (Account(Index).Chars(GetPlayerChar(Index)).Stat(Stats.Endurance) / 3)) * 15 + 135
                Case 2 ' Range
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (Account(Index).Chars(GetPlayerChar(Index)).Stat(Stats.Endurance) / 3)) * 10 + 100
                Case 3 ' Magic
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (Account(Index).Chars(GetPlayerChar(Index)).Stat(Stats.Endurance) / 3)) * 5 + 75
            End Select

        Case MP
            Select Case Class(GetPlayerClass(Index)).CombatTree
                Case 1 ' Melee
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (Account(Index).Chars(GetPlayerChar(Index)).Stat(Stats.Intelligence) / 3)) * 5 + 75
                Case 2 ' Range
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (Account(Index).Chars(GetPlayerChar(Index)).Stat(Stats.Intelligence) / 3)) * 10 + 100
                Case 3 ' Magic
                    GetPlayerMaxVital = ((GetPlayerLevel(Index) / 2) + (Account(Index).Chars(GetPlayerChar(Index)).Stat(Stats.Intelligence) / 3)) * 15 + 135
            End Select
    End Select
End Function

Function GetPlayerVitalRegen(ByVal Index As Long, ByVal Vital As Vitals) As Long
    Dim I As Long

    ' Prevent subscript out of range
    If IsPlaying(Index) = False Or Index < 1 Or Index > Player_HighIndex Then
        GetPlayerVitalRegen = 0
        Exit Function
    End If

    Select Case Vital
        Case HP
            I = (GetPlayerStat(Index, Stats.Spirit) * 0.8) + 7
            If I > GetPlayerMaxVital(Index, HP) / 25 Then
                I = GetPlayerMaxVital(Index, HP) / 25
            End If
        Case MP
            I = (GetPlayerStat(Index, Stats.Spirit) / 4) + 12
            If I > GetPlayerMaxVital(Index, MP) / 25 Then
                I = GetPlayerMaxVital(Index, MP) / 25
            End If
    End Select

    Round I
    GetPlayerVitalRegen = I
End Function

Public Sub selectValue(ByRef textBox As textBox)
    textBox.SelStart = 0
    textBox.SelLength = Len(textBox.Text)
End Sub

Function GetPlayerDamage(ByVal Index As Long) As Long
    Dim WeaponNum As Long
    
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index < 1 Or Index > Player_HighIndex Then Exit Function
    
    If GetPlayerEquipment(Index, Equipment.Weapon) > 0 Then
        If Not GetPlayerEquipmentDur(Index, GetPlayerEquipment(Index, Equipment.Weapon)) = 0 Or Item(GetPlayerEquipment(Index, Equipment.Weapon)).Indestructable = 1 Then
            WeaponNum = GetPlayerEquipment(Index, Equipment.Weapon)
            GetPlayerDamage = 0.085 * 5 * GetPlayerStat(Index, Strength) * Item(WeaponNum).Data2 + (GetPlayerLevel(Index) * 0.2)
        Else
            GetPlayerDamage = 0.085 * 5 * GetPlayerStat(Index, Strength) + (GetPlayerLevel(Index) * 0.2)
        End If
    Else
        GetPlayerDamage = 0.085 * 5 * GetPlayerStat(Index, Strength) + (GetPlayerLevel(Index) * 0.2)
    End If
End Function

Public Function GetNPCSpellVital(ByVal MapNum As Integer, ByVal MapNPCNum As Byte, ByVal Victim As Byte, ByVal SpellNum As Long, Optional ByVal HealingSpell As Boolean = False) As Long
    If Victim < 1 Or MapNPCNum < 1 Or MapNum < 1 Then Exit Function
    If MapNPC(MapNum).NPC(MapNPCNum).Num < 1 Then Exit Function
    
    GetNPCSpellVital = Spell(SpellNum).Vital + (NPC(MapNPC(MapNum).NPC(MapNPCNum).Num).Stat(Stats.Intelligence) / 3)
    
    ' Randomize damage
    GetNPCSpellVital = Random(GetNPCSpellVital - (GetNPCSpellVital / 2), GetNPCSpellVital)
    
    ' 1.5 times the damage if it's a critical
    If CanNPCSpellCritical(MapNPCNum) Then
        GetNPCSpellVital = GetNPCSpellVital * 1.5
        Call SendSoundToMap(MapNum, Options.CriticalSound)
        SendAnimation MapNum, Options.CriticalAnimation, 0, 0, TARGET_TYPE_NPC, MapNPCNum
    End If
    
    If HealingSpell = False Then
        If MapNPC(MapNum).NPC(MapNPCNum).targetType = TARGET_TYPE_PLAYER Then
            GetNPCSpellVital = GetNPCSpellVital - GetPlayerStat(Victim, Spirit)
        Else
            GetNPCSpellVital = GetNPCSpellVital - NPC(MapNPC(MapNum).NPC(Victim).Num).Stat(Stats.Spirit)
        End If
    End If
End Function

Function GetNPCMaxVital(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Dim X As Long

    ' Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then Exit Function

    Select Case Vital
        Case HP
            GetNPCMaxVital = NPC(NPCNum).HP
        Case MP
            GetNPCMaxVital = NPC(NPCNum).MP
    End Select
End Function

Function GetNPCVitalRegen(ByVal NPCNum As Long, ByVal Vital As Vitals) As Long
    Dim I As Long

    ' Prevent subscript out of range
    If NPCNum <= 0 Or NPCNum > MAX_NPCS Then Exit Function

    Select Case Vital
        Case HP
            I = (NPC(NPCNum).Stat(Stats.Spirit) * 0.8) + 7
            If I > GetNPCMaxVital(NPCNum, HP) / 25 Then
                I = GetNPCMaxVital(NPCNum, HP) / 25
            End If
        Case MP
            I = (NPC(NPCNum).Stat(Stats.Spirit) / 4) + 12
            If I > GetNPCMaxVital(NPCNum, MP) / 25 Then
                I = GetNPCMaxVital(NPCNum, MP) / 25
            End If
    End Select
    
    Round I
    GetNPCVitalRegen = I
End Function

Function GetNPCDamage(ByVal NPCNum As Long) As Long
    GetNPCDamage = 0.085 * 5 * NPC(NPCNum).Stat(Stats.Strength) * NPC(NPCNum).Damage + (NPC(NPCNum).Level / 5)
End Function

' ###############################
' ##      Luck-based rates     ##
' ###############################
Public Function CanPlayerCritical(ByVal Index As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanPlayerCritical = False

    Rate = GetPlayerStat(Index, Agility) / 52.08
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanPlayerCritical = True
    End If
End Function

Public Function CanPlayerSpellCritical(ByVal Index As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanPlayerSpellCritical = False

    Rate = Account(Index).Chars(GetPlayerChar(Index)).Stat(Stats.Intelligence) / 78.16
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanPlayerSpellCritical = True
    End If
End Function

Public Function CanPlayerDodge(ByVal Index As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanPlayerDodge = False

    Rate = GetPlayerStat(Index, Agility) / 83.3
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanPlayerDodge = True
    End If
End Function

Public Function CanPlayerDeflect(ByVal Index As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanPlayerDeflect = False

    Rate = GetPlayerStat(Index, Strength) * 0.25
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanPlayerDeflect = True
    End If
End Function

Public Function CanNPCCritical(ByVal NPCNum As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanNPCCritical = False

    Rate = NPC(NPCNum).Stat(Stats.Agility) / 52.08
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanNPCCritical = True
    End If
End Function

Public Function CanNPCSpellCritical(ByVal NPCNum As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanNPCSpellCritical = False

    Rate = NPC(NPCNum).Stat(Stats.Intelligence) / 78.16
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanNPCSpellCritical = True
    End If
End Function

Function GetPlayerProtection(ByVal Index As Long) As Long
    ' Check for subscript out of range
    If IsPlaying(Index) = False Or Index <= 0 Or Index > Player_HighIndex Then Exit Function

    GetPlayerProtection = (GetPlayerStat(Index, Stats.Endurance) \ 4)

    If GetPlayerEquipment(Index, Equipment.Body) > 0 Then
        If Not GetPlayerEquipmentDur(Index, Equipment.Body) = 0 Or Item(GetPlayerEquipment(Index, Equipment.Body)).Indestructable = 1 Then
            GetPlayerProtection = GetPlayerProtection + Item(Body).Data2
        End If
    End If

    If GetPlayerEquipment(Index, Equipment.Head) > 0 Then
        If Not GetPlayerEquipmentDur(Index, Equipment.Head) = 0 Or Item(GetPlayerEquipment(Index, Equipment.Head)).Indestructable = 1 Then
            GetPlayerProtection = GetPlayerProtection + Item(Equipment.Head).Data2
        End If
    End If
End Function

Public Function CanPlayerBlock(ByVal Index As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long
    Dim ShieldNum As Long

    CanPlayerBlock = False

    If GetPlayerEquipment(Index, Shield) > 0 Then
        ShieldNum = GetPlayerEquipment(Index, Shield)
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

Function CanPlayerMitigateNPC(ByVal Index As Long, MapNPCNum As Long) As Boolean
    If Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_UP Then
        If MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_DOWN Then
            CanPlayerMitigateNPC = True
        End If
    ElseIf Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_DOWN Then
        If MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_UP Then
            CanPlayerMitigateNPC = True
        End If
    ElseIf Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_LEFT Then
        If MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_RIGHT Then
            CanPlayerMitigateNPC = True
        End If
    ElseIf Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_RIGHT Then
        If MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_LEFT Then
            CanPlayerMitigateNPC = True
        End If
    Else
        CanPlayerMitigateNPC = False
    End If
End Function

Function CanNPCMitigatePlayer(ByVal MapNPCNum As Long, Index As Long) As Boolean
    If MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_UP Then
        If Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_DOWN Then
            CanNPCMitigatePlayer = True
        End If
    ElseIf MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_DOWN Then
        If Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_UP Then
            CanNPCMitigatePlayer = True
        End If
    ElseIf MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_LEFT Then
        If Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_RIGHT Then
            CanNPCMitigatePlayer = True
        End If
    ElseIf MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_RIGHT Then
        If Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_LEFT Then
            CanNPCMitigatePlayer = True
        End If
    Else
        CanNPCMitigatePlayer = False
    End If
End Function

Function CanNPCMitigateNPC(ByVal Attacker As Long, Victim As Long, MapNum As Integer) As Boolean
    If MapNPC(MapNum).NPC(Attacker).Dir = DIR_UP Then
        If MapNPC(MapNum).NPC(Victim).Dir = DIR_DOWN Then
            CanNPCMitigateNPC = True
        End If
    ElseIf MapNPC(MapNum).NPC(Attacker).Dir = DIR_DOWN Then
        If MapNPC(MapNum).NPC(Victim).Dir = DIR_UP Then
            CanNPCMitigateNPC = True
        End If
    ElseIf MapNPC(MapNum).NPC(Attacker).Dir = DIR_LEFT Then
        If MapNPC(MapNum).NPC(Victim).Dir = DIR_RIGHT Then
            CanNPCMitigateNPC = True
        End If
    ElseIf MapNPC(MapNum).NPC(Attacker).Dir = DIR_RIGHT Then
        If MapNPC(MapNum).NPC(Victim).Dir = DIR_LEFT Then
            CanNPCMitigateNPC = True
        End If
    Else
        CanNPCMitigateNPC = False
    End If
End Function

Public Function CanNPCDodge(ByVal NPCNum As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanNPCDodge = False

    Rate = NPC(NPCNum).Stat(Stats.Agility) / 83.3
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanNPCDodge = True
    End If
End Function

Public Function CanNPCDeflect(ByVal NPCNum As Long) As Boolean
    Dim Rate As Long
    Dim RandomNum As Long

    CanNPCDeflect = False

    Rate = NPC(NPCNum).Stat(Stats.Strength) * 0.25
    RandomNum = Random(1, 100)
    
    If RandomNum <= Rate Then
        CanNPCDeflect = True
    End If
End Function

' ###################################
' ##      Player Attacking NPC     ##
' ###################################
Public Sub TryPlayerAttackNPC(ByVal Index As Long, ByVal MapNPCNum As Long)
    Dim NPCNum As Long
    Dim MapNum As Integer
    Dim Damage As Long
    
    ' Can we attack the npc?
    If CanPlayerAttackNPC(Index, MapNPCNum, False) Then
        MapNum = GetPlayerMap(Index)
        NPCNum = MapNPC(MapNum).NPC(MapNPCNum).Num
    
        ' Get the damage we can do
        Damage = GetPlayerDamage(Index)
        
        ' Add damage based on direction
        If Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_UP Then
            If MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_LEFT Or MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_UP Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_DOWN Then
            If MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_LEFT Or MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_LEFT Then
            If MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_UP Or MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_LEFT Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_RIGHT Then
            If MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_UP Or MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 4)
            End If
        End If
        
        ' 1.5 times the damage if it's a critical
        If CanPlayerCritical(Index) Then
            Damage = Damage * 1.5
            Call SendSoundToMap(MapNum, Options.CriticalSound)
            SendAnimation MapNum, Options.CriticalAnimation, 0, 0, TARGET_TYPE_PLAYER, Index
        End If
        
        ' Take away protection from the damage
        Damage = Damage - (NPC(MapNPCNum).Stat(Stats.Endurance) / 4)
        
        ' Randomize damage
        Damage = Random(Damage - (Damage / 2), Damage)
        
        Round Damage
        
        If Damage < 1 Then
            Call SendSoundToMap(MapNum, Options.MissSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_NPC, MapNPCNum
            Exit Sub
        End If
    
        Call PlayerAttackNPC(Index, MapNPCNum, Damage)
    End If
End Sub

Public Function CanPlayerAttackNPC(ByVal Attacker As Long, ByVal MapNPCNum As Long, Optional ByVal IsSpell As Boolean = False) As Boolean
    Dim MapNum As Integer
    Dim NPCNum As Long
    Dim NPCX As Long
    Dim NPCY As Long
    Dim Attackspeed As Long
    Dim WeaponSlot As Long
    Dim Range As Byte
    Dim DistanceToNPC As Integer
    Dim FindQuest As FindQuestRec

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Then Exit Function

    ' Check for subscript out of range
    If MapNPC(GetPlayerMap(Attacker)).NPC(MapNPCNum).Num < 1 Then Exit Function
    
    MapNum = GetPlayerMap(Attacker)
    NPCNum = MapNPC(MapNum).NPC(MapNPCNum).Num
    
    ' Make sure the npc isn't already dead
    If MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.HP) < 1 Then
        If NPC(NPCNum).Behavior = NPC_BEHAVIOR_QUEST Then
            FindQuest = IsQuestCLI(Attacker, NPCNum)
            If Not FindQuest.QuestIndex > 0 Then Exit Function
        Else
            Exit Function
        End If
    End If

    ' Make sure they are a player killer or else they can't attack a guard
    If NPC(NPCNum).Behavior = NPC_BEHAVIOR_GUARD And GetPlayerPK(Attacker) = NO Then Exit Function

    ' Attack speed from weapon
    If GetPlayerEquipment(Attacker, Weapon) > 0 Then
        Attackspeed = Item(GetPlayerEquipment(Attacker, Weapon)).WeaponSpeed
    Else
        Attackspeed = 1000
    End If
    
    If NPCNum > 0 And timeGetTime > TempPlayer(Attacker).AttackTimer + Attackspeed Then
        If Not IsSpell Then ' Melee attack
            ' Check if at same coordinates
            Select Case GetPlayerDir(Attacker)
                Case DIR_UP
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y + 1 = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X = GetPlayerX(Attacker))) Then Exit Function
                Case DIR_DOWN
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y - 1 = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X = GetPlayerX(Attacker))) Then Exit Function
                Case DIR_LEFT
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X + 1 = GetPlayerX(Attacker))) Then Exit Function
                Case DIR_RIGHT
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X - 1 = GetPlayerX(Attacker))) Then Exit Function
                Case DIR_UPLEFT
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y + 1 = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X = GetPlayerX(Attacker))) Then Exit Function
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X + 1 = GetPlayerX(Attacker))) Then Exit Function
                Case DIR_UPRIGHT
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y + 1 = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X = GetPlayerX(Attacker))) Then Exit Function
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X - 1 = GetPlayerX(Attacker))) Then Exit Function
                Case DIR_DOWNLEFT
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X + 1 = GetPlayerX(Attacker))) Then Exit Function
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y - 1 = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X = GetPlayerX(Attacker))) Then Exit Function
                Case DIR_DOWNRIGHT
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X - 1 = GetPlayerX(Attacker))) Then Exit Function
                    If Not ((MapNPC(MapNum).NPC(MapNPCNum).Y - 1 = GetPlayerY(Attacker)) And (MapNPC(MapNum).NPC(MapNPCNum).X = GetPlayerX(Attacker))) Then Exit Function
                Case Else
                    Exit Function
            End Select
        End If
        
        'Handle quest part.
        If FindQuest.QuestIndex > 0 Then
            Call CheckQuest(Attacker, FindQuest.QuestIndex, FindQuest.CLIIndex, FindQuest.ActionIndex)
        End If

        If Not IsSpell Then
            If Not NPC(NPCNum).Behavior = NPC_BEHAVIOR_QUEST Then
                If DidNPCMitigatePlayer(Attacker, MapNPCNum) = False Then
                    CanPlayerAttackNPC = True
                End If
            ElseIf Len(Trim$(NPC(NPCNum).AttackSay)) > 0 Then
                Call SendChatBubble(MapNum, MapNPCNum, TARGET_TYPE_NPC, Trim$(NPC(NPCNum).AttackSay), White)
            End If
        ElseIf IsSpell Then
            If Not NPC(NPCNum).Behavior = NPC_BEHAVIOR_QUEST Then
                If DidNPCMitigatePlayer(Attacker, MapNPCNum) = False Then
                    CanPlayerAttackNPC = True
                End If
            End If
        End If
    End If
End Function

Public Function DidNPCMitigatePlayer(ByVal Attacker As Long, ByVal MapNPCNum As Long) As Boolean
    Dim MapNum As Integer
    Dim NPCNum As Long
    
    MapNum = GetPlayerMap(Attacker)
    NPCNum = MapNPC(MapNum).NPC(MapNPCNum).Num
    
    If CanNPCMitigatePlayer(MapNPCNum, Attacker) = True Or TempPlayer(Attacker).SpellBuffer.Spell > 0 Then
        ' Check if NPC can avoid the attack
        If CanNPCDodge(NPCNum) Then
            Call SendSoundToMap(MapNum, Options.DodgeSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_NPC, MapNPCNum
            DidNPCMitigatePlayer = True
            Exit Function
        End If
        
        If CanNPCDeflect(NPCNum) Then
            Call SendSoundToMap(MapNum, Options.DeflectSound)
            SendAnimation MapNum, Options.DeflectAnimation, 0, 0, TARGET_TYPE_NPC, MapNPCNum
            DidNPCMitigatePlayer = True
            Exit Function
        End If
    End If
    
    DidNPCMitigatePlayer = False
End Function

Public Sub PlayerAttackNPC(ByVal Attacker As Long, ByVal MapNPCNum As Long, ByVal Damage As Long, Optional ByVal SpellNum As Long, Optional ByVal OverTime As Boolean = False)
    Dim Name As String
    Dim Exp As Long
    Dim n As Long
    Dim I As Long
    Dim STR As Long
    Dim DEF As Long
    Dim MapNum As Integer
    Dim NPCNum As Long
    Dim Value As Long
    Dim LevelDiff As Long

    ' Check for subscript out of range
    If IsPlaying(Attacker) = False Or MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Then Exit Sub

    MapNum = GetPlayerMap(Attacker)
    NPCNum = MapNPC(MapNum).NPC(MapNPCNum).Num
    Name = Trim$(NPC(NPCNum).Name)
    
    ' Set the attacker's target
    If SpellNum = 0 Then
        TempPlayer(Attacker).targetType = TARGET_TYPE_NPC
        TempPlayer(Attacker).target = MapNPCNum
        Call SendPlayerTarget(Attacker)
    End If
    
    ' Set their target if they are being hit
    MapNPC(MapNum).NPC(MapNPCNum).targetType = TARGET_TYPE_PLAYER
    MapNPC(MapNum).NPC(MapNPCNum).target = Attacker
    Call SendMapNPCTarget(MapNum, MapNPCNum, MapNPC(MapNum).NPC(MapNPCNum).target, MapNPC(GetPlayerMap(Attacker)).NPC(MapNPCNum).targetType)
    
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
        Call SendMapSound(MapNum, Attacker, MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y, SoundEntity.seSpell, SpellNum)
     Else
        Call SendMapSound(MapNum, Attacker, MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y, SoundEntity.seAnimation, 1)
     End If

    If Damage >= MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.HP) Then
        ' Set the damage to the npc's health so that it doesn't appear that it's overkilling it
        Damage = MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.HP)
    
        SendActionMsg GetPlayerMap(Attacker), "-" & Damage, BrightRed, 1, (MapNPC(MapNum).NPC(MapNPCNum).X * 32), (MapNPC(MapNum).NPC(MapNPCNum).Y * 32)
        SendBlood GetPlayerMap(Attacker), MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y
         
        ' Send animation
        If SpellNum < 1 Then
            If GetPlayerEquipment(Attacker, Weapon) > 0 Then
                If Item(GetPlayerEquipment(Attacker, Weapon)).Animation > 0 Then
                    If n > 0 Then
                        If Not OverTime Then
                            Call SendAnimation(MapNum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, MapNPCNum)
                        End If
                    End If
                Else
                    Call SendAnimation(MapNum, 1, 0, 0, TARGET_TYPE_NPC, MapNPCNum)
                End If
            Else
                Call SendAnimation(MapNum, 1, 0, 0, TARGET_TYPE_NPC, MapNPCNum)
            End If
        End If

        ' Calculate experience to give attacker
        Exp = NPC(NPCNum).Exp
        
        ' Find the level difference between the npc and player
        LevelDiff = GetPlayerLevel(Attacker) - NPC(NPCNum).Level
        
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
            If NPC(NPCNum).DropItem(n) = 0 Then Exit For
            
            Value = NPC(NPCNum).DropValue(n)
            Value = Random(Value * 0.25, Value * 1.5)
            Round Value
            
            If Value < 1 Then Value = 1
            
            If Rnd <= NPC(NPCNum).DropChance(n) Then
                If TempPlayer(Attacker).InParty > 0 Then
                    Call Party_GetLoot(TempPlayer(Attacker).InParty, NPC(NPCNum).DropItem(n), NPC(NPCNum).DropValue(n), MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y)
                Else
                    Call SpawnItem(NPC(NPCNum).DropItem(n), Value, Item(NPC(NPCNum).DropItem(n)).Data1, MapNum, MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y, GetPlayerName(Attacker))
                End If
            End If
        Next

        ' Now set HP to 0, so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNPC(MapNum).NPC(MapNPCNum).Num = 0
        MapNPC(MapNum).NPC(MapNPCNum).SpawnWait = timeGetTime
        MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.HP) = 0
        UpdateMapBlock MapNum, MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y, False
        
        ' Clear DoTs and HoTs
        For I = 1 To MAX_DOTS
            With MapNPC(MapNum).NPC(MapNPCNum).DoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNPC(MapNum).NPC(MapNPCNum).HoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' Send death to the map
        Call SendNPCDeath(MapNPCNum, MapNum)
        
        ' Set the player variables/swithces
        If NPC(NPCNum).SwitchNum > 0 Then
            Call SetPlayerSwitch(Attacker, NPC(NPCNum).SwitchNum, NPC(NPCNum).SwitchVal)
        End If

        If NPC(NPCNum).VariableNum > 0 Then
            If NPC(NPCNum).AddToVariable = 1 Then
                Call SetPlayerVariable(Attacker, NPC(NPCNum).VariableNum, GetPlayerVariable(Attacker, NPC(NPCNum).VariableNum) + NPC(NPCNum).VariableVal)
            Else
                Call SetPlayerVariable(Attacker, NPC(NPCNum).VariableNum, NPC(NPCNum).VariableVal)
            End If
        End If
        
        ' Loop through entire map and purge npcs from players
        For I = 1 To Player_HighIndex
            If IsPlaying(I) Then
                If Account(I).Chars(GetPlayerChar(I)).Map = MapNum Then
                    If TempPlayer(I).targetType = TARGET_TYPE_NPC Then
                        If TempPlayer(I).target = MapNPCNum Then
                            TempPlayer(I).target = 0
                            TempPlayer(I).targetType = TARGET_TYPE_NONE
                            SendPlayerTarget I
                        End If
                    End If
                End If
            End If
        Next
        
        'See if we need to add to the player's kill count for a quest
        Call CheckIfQuestKill(Attacker, NPCNum)
    Else
        ' NPC not dead, just do the damage
        MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.HP) = MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.HP) - Damage
        Call SendMapNPCVitals(MapNum, MapNPCNum)

        ' Check for a weapon and say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNPC(MapNum).NPC(MapNPCNum).X * 32), (MapNPC(MapNum).NPC(MapNPCNum).Y * 32)
        SendBlood GetPlayerMap(Attacker), MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y
        
        ' Send animation
        If SpellNum < 1 Then
            If GetPlayerEquipment(Attacker, Weapon) > 0 Then
                If Item(GetPlayerEquipment(Attacker, Weapon)).Animation > 0 Then
                    If n > 0 Then
                        If Not OverTime Then
                            Call SendAnimation(MapNum, Item(GetPlayerEquipment(Attacker, Weapon)).Animation, 0, 0, TARGET_TYPE_NPC, MapNPCNum)
                        End If
                    End If
                Else
                    Call SendAnimation(MapNum, 1, 0, 0, TARGET_TYPE_NPC, MapNPCNum)
                End If
            Else
                Call SendAnimation(MapNum, 1, 0, 0, TARGET_TYPE_NPC, MapNPCNum)
            End If
        End If
        
        ' Now check for guard ai and if so have all onmap guards come after'm
        If NPC(MapNPC(MapNum).NPC(MapNPCNum).Num).Behavior = NPC_BEHAVIOR_GUARD Then
            For I = 1 To Map(MapNum).NPC_HighIndex
                If MapNPC(MapNum).NPC(I).Num = MapNPC(MapNum).NPC(MapNPCNum).Num Then
                    MapNPC(MapNum).NPC(I).target = Attacker
                    MapNPC(MapNum).NPC(I).targetType = TARGET_TYPE_PLAYER
                    Call SendMapNPCTarget(MapNum, I, MapNPC(MapNum).NPC(I).target, MapNPC(MapNum).NPC(I).targetType)
                End If
            Next
        End If
        
        ' Set the regen timer
        MapNPC(MapNum).NPC(MapNPCNum).StopRegen = True
        MapNPC(MapNum).NPC(MapNPCNum).StopRegenTimer = timeGetTime
        
        ' If stunning spell then stun the npc
        If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
            If Spell(SpellNum).StunDuration > 0 Then StunNPC MapNPCNum, MapNum, SpellNum
            
            ' DoT
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_NPC MapNum, MapNPCNum, SpellNum, Attacker
            End If
        End If
        SendMapNPCVitals MapNum, MapNPCNum
    End If
    
    If SpellNum = 0 Then
        ' Reset the attack timer
        TempPlayer(Attacker).AttackTimer = timeGetTime
    End If
    
    ' Reduce durability of weapon
    Call DamagePlayerEquipment(Attacker, Equipment.Weapon)
End Sub

' ###################################
' ##      NPC Attacking NPC        ##
' ###################################
Public Sub TryNPCAttackNPC(ByVal MapNum As Integer, ByVal Attacker As Long, ByVal Victim As Long)
    Dim NPCNum As Long, Damage As Long, I As Long
    
    NPCNum = MapNPC(MapNum).NPC(Attacker).Num
    Damage = GetNPCDamage(NPCNum)

    ' Set the npc target to the npc
    If MapNPC(MapNum).NPC(Victim).target = 0 Then
        MapNPC(MapNum).NPC(Victim).target = Attacker
        MapNPC(MapNum).NPC(Victim).targetType = TARGET_TYPE_NPC
        Call SendMapNPCTarget(MapNum, Victim, MapNPC(MapNum).NPC(Victim).target, MapNPC(MapNum).NPC(Victim).targetType)
    End If
        
    ' Can the npc attack the player
    If CanNPCAttackNPC(MapNum, Attacker, Victim) Then
        ' Set attack timer
        MapNPC(MapNum).NPC(Attacker).AttackTimer = timeGetTime
        
        If NPC(MapNPC(MapNum).NPC(Victim).Num).FactionThreat = True Then
            ' Send threat to all of the same faction if they have the option enabled
            For I = 1 To Map(MapNum).NPC_HighIndex
                If MapNPC(MapNum).NPC(I).Num > 0 Then
                    If NPC(MapNPC(MapNum).NPC(Victim).Num).Faction > 0 And NPC(MapNPC(MapNum).NPC(I).Num).Faction > 0 Then
                        If NPC(MapNPC(MapNum).NPC(Victim).Num).Faction = NPC(MapNPC(MapNum).NPC(I).Num).Faction Then
                            If NPC(MapNPC(MapNum).NPC(Victim).Num).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Then
                                If MapNPC(MapNum).NPC(I).target = 0 Then
                                    MapNPC(MapNum).NPC(I).targetType = TARGET_TYPE_NPC
                                    MapNPC(MapNum).NPC(I).target = Attacker
                                    Call SendMapNPCTarget(MapNum, I, MapNPC(MapNum).NPC(I).target, MapNPC(MapNum).NPC(I).targetType)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        
        ' Add damage based on direction
        If MapNPC(MapNum).NPC(Attacker).Dir = DIR_UP Then
            If MapNPC(MapNum).NPC(Victim).Dir = DIR_LEFT Or MapNPC(MapNum).NPC(Victim).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNPC(MapNum).NPC(Victim).Dir = DIR_UP Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNPC(MapNum).NPC(Attacker).Dir = DIR_DOWN Then
            If MapNPC(MapNum).NPC(Victim).Dir = DIR_LEFT Or MapNPC(MapNum).NPC(Victim).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNPC(MapNum).NPC(Victim).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNPC(MapNum).NPC(Attacker).Dir = DIR_LEFT Then
            If MapNPC(MapNum).NPC(Victim).Dir = DIR_UP Or MapNPC(MapNum).NPC(Victim).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNPC(MapNum).NPC(Victim).Dir = DIR_LEFT Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNPC(MapNum).NPC(Attacker).Dir = DIR_RIGHT Then
            If MapNPC(MapNum).NPC(Victim).Dir = DIR_UP Or MapNPC(MapNum).NPC(Victim).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf MapNPC(MapNum).NPC(Victim).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 4)
            End If
        End If
        
        ' 1.5 times the damage if it's a critical
        If CanNPCCritical(Attacker) Then
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
        If Trim$(NPC(MapNPC(MapNum).NPC(Victim).Num).Sound) = vbNullString Then
            Call SendMapSound(MapNum, Victim, MapNPC(MapNum).NPC(Victim).X, MapNPC(MapNum).NPC(Victim).Y, SoundEntity.seAnimation, 1)
        Else
            Call SendMapSound(MapNum, Victim, MapNPC(MapNum).NPC(Victim).X, MapNPC(MapNum).NPC(Victim).Y, SoundEntity.seNPC, MapNPC(MapNum).NPC(Victim).Num)
        End If
        
        Call NPCAttackNPC(MapNum, Attacker, Victim, Damage)
    End If
End Sub

Function CanNPCAttackNPC(ByVal MapNum As Integer, ByVal Attacker As Long, ByVal Victim As Long, Optional ByVal Spell As Boolean = False) As Boolean
    Dim aNPCNum As Long
    Dim vNPCNum As Long
    Dim VictimX As Long
    Dim VictimY As Long
    Dim AttackerX As Long
    Dim AttackerY As Long
    
    CanNPCAttackNPC = False

    ' Check for subscript out of range
    If Attacker < 1 Or Attacker > MAX_MAP_NPCS Then Exit Function
    If Victim < 1 Or Victim > MAX_MAP_NPCS Then Exit Function
    
    aNPCNum = MapNPC(MapNum).NPC(Attacker).Num
    vNPCNum = MapNPC(MapNum).NPC(Victim).Num
    
     ' Check for subscript out of range
    If aNPCNum < 1 Or vNPCNum < 1 Then Exit Function
    
    ' Can't attack itself
    If aNPCNum = vNPCNum Then Exit Function

    ' Make sure the NPCs aren't already dead
    If MapNPC(MapNum).NPC(Attacker).Vital(Vitals.HP) < 1 Or MapNPC(MapNum).NPC(Victim).Vital(Vitals.HP) < 1 Then Exit Function
    
    ' Make sure they aren't trying to attack a friendly or shopkeeper NPC
    If NPC(MapNPC(MapNum).NPC(Victim).Num).Behavior = NPC_BEHAVIOR_QUEST Then Exit Function
    
    ' Make sure they aren't casting a spell
    If MapNPC(MapNum).NPC(Attacker).SpellBuffer.Timer > 0 And Spell = False Then Exit Function

    ' Check if they have the same faction if they do exit
    If NPC(MapNPC(MapNum).NPC(Attacker).Num).Faction > 0 Then
        If NPC(MapNPC(MapNum).NPC(Attacker).Num).Faction = NPC(MapNPC(MapNum).NPC(Victim).Num).Faction Then Exit Function
    End If
    
    If Spell Then
        CanNPCAttackNPC = True
        Exit Function
    End If

    ' Make sure npcs don't attack more than once a second
    If timeGetTime < MapNPC(MapNum).NPC(Attacker).AttackTimer + 1000 Then Exit Function
    
    AttackerX = MapNPC(MapNum).NPC(Attacker).X
    AttackerY = MapNPC(MapNum).NPC(Attacker).Y
    VictimX = MapNPC(MapNum).NPC(Victim).X
    VictimY = MapNPC(MapNum).NPC(Victim).Y
    
    ' Check if they are going to cast
    If Random(1, 2) = 1 And CanNPCCastSpell(MapNum, Attacker) Then
        Call BufferNPCSpell(MapNum, Attacker, Victim)
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
    
    ' Send this packet so they can see the person attacking
    Call SendNPCAttack(Attacker, MapNum)
    
    CanNPCAttackNPC = True
End Function

Private Function DidNPCMitigateNPC(ByVal MapNum As Integer, Attacker As Long, Victim As Long) As Boolean
    If CanNPCMitigateNPC(Attacker, Victim, MapNum) = True Or MapNPC(MapNum).NPC(Attacker).SpellBuffer.Spell > 0 Then
        ' Check if npc can avoid the attack
        If CanNPCDodge(Victim) Then
            Call SendSoundToMap(MapNum, Options.DodgeSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_NPC, Victim
            DidNPCMitigateNPC = True
            Exit Function
        End If
        
        ' Check if npc can deflect the attack
        If CanNPCDeflect(Victim) Then
            Call SendSoundToMap(MapNum, Options.DeflectSound)
            SendAnimation MapNum, Options.DeflectAnimation, 0, 0, TARGET_TYPE_NPC, Victim
            DidNPCMitigateNPC = True
            Exit Function
        End If
    End If
    
    DidNPCMitigateNPC = False
End Function

Sub NPCAttackNPC(ByVal MapNum As Integer, ByVal Attacker As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim I As Long, n As Byte
    Dim aVictim As Long
    Dim vVictim As Long
    Dim DistanceX As Byte
    Dim DistanceY As Byte
    Dim Value As Long
    
    ' Check for subscript out of range
    If Attacker < 1 Or Attacker > MAX_MAP_NPCS Or Victim < 1 Or Victim > MAX_MAP_NPCS Then Exit Sub
    
    If DidNPCMitigateNPC(MapNum, Attacker, Victim) = True Then Exit Sub
    
    aVictim = MapNPC(MapNum).NPC(Attacker).Num
    vVictim = MapNPC(MapNum).NPC(Victim).Num
    
    ' Check for subscript out of range
    If aVictim < 1 Then Exit Sub
    If vVictim < 1 Then Exit Sub
    
     ' Set the regen timer
    MapNPC(MapNum).NPC(Attacker).StopRegen = True
    MapNPC(MapNum).NPC(Attacker).StopRegenTimer = timeGetTime

    If Damage >= MapNPC(MapNum).NPC(Victim).Vital(Vitals.HP) Then
        ' Set the damage to the target's health exactly so it's not overkilling them
        Damage = MapNPC(MapNum).NPC(Victim).Vital(Vitals.HP)
        
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNPC(MapNum).NPC(Victim).X * 32), (MapNPC(MapNum).NPC(Victim).Y * 32)
        SendBlood MapNum, MapNPC(MapNum).NPC(Victim).X, MapNPC(MapNum).NPC(Victim).Y
        
        Call SendMapNPCTarget(MapNum, Attacker, 0, 0)
        
        ' Now set HP to 0 so we know to actually kill them in the server loop (this prevents subscript out of range)
        MapNPC(MapNum).NPC(Victim).Num = 0
        MapNPC(MapNum).NPC(Victim).SpawnWait = timeGetTime
        MapNPC(MapNum).NPC(Victim).Vital(Vitals.HP) = 0
        UpdateMapBlock MapNum, MapNPC(MapNum).NPC(Victim).X, MapNPC(MapNum).NPC(Victim).Y, False
        
        ' Clear DoTs and HoTs
        For I = 1 To MAX_DOTS
            With MapNPC(MapNum).NPC(Victim).DoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
            
            With MapNPC(MapNum).NPC(Victim).HoT(I)
                .Spell = 0
                .Timer = 0
                .Caster = 0
                .StartTime = 0
                .Used = False
            End With
        Next
        
        ' Send npc death packet to map
        Call SendNPCDeath(Victim, MapNum)
    Else
        ' NPC not dead, just do the damage
        If MapNPC(MapNum).NPC(Attacker).SpellBuffer.Spell = 0 Then
            If NPC(MapNPC(MapNum).NPC(Attacker).Num).Animation < 1 Then
                Call SendAnimation(MapNum, 1, 0, 0, TARGET_TYPE_NPC, Victim)
            Else
                Call SendAnimation(MapNum, NPC(MapNPC(MapNum).NPC(Attacker).Num).Animation, 0, 0, TARGET_TYPE_NPC, Victim)
            End If
        End If
        
        MapNPC(MapNum).NPC(Victim).Vital(Vitals.HP) = MapNPC(MapNum).NPC(Victim).Vital(Vitals.HP) - Damage
        Call SendMapNPCVitals(MapNum, Victim)
        
        ' Say damage
        SendActionMsg MapNum, "-" & Damage, BrightRed, 1, (MapNPC(MapNum).NPC(Victim).X * 32), (MapNPC(MapNum).NPC(Victim).Y * 32)
        SendBlood MapNum, MapNPC(MapNum).NPC(Victim).X, MapNPC(MapNum).NPC(Victim).Y
        
        ' Set the regen timer
        TempPlayer(Victim).StopRegen = True
        TempPlayer(Victim).StopRegenTimer = timeGetTime
    End If
End Sub

' ###################################
' ##      NPC Attacking Player     ##
' ###################################
Public Sub TryNPCAttackPlayer(ByVal MapNPCNum As Long, ByVal Index As Long)
    Dim MapNum As Integer, NPCNum As Long, Damage As Long, n As Byte, DistanceX As Byte, DistanceY As Byte
    
    MapNum = GetPlayerMap(Index)
    NPCNum = MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Num
    
    ' Can the npc attack the player
    If CanNPCAttackPlayer(MapNPCNum, Index) Then
        ' Set attack timer
        MapNPC(MapNum).NPC(MapNPCNum).AttackTimer = timeGetTime
        
        If NPC(MapNPC(MapNum).NPC(MapNPCNum).Num).FactionThreat = True Then
            ' Send threat to all of the same faction if they have the option enabled
            For n = 1 To Map(MapNum).NPC_HighIndex
                If MapNPC(MapNum).NPC(n).Num > 0 Then
                    If NPC(MapNPC(MapNum).NPC(MapNPCNum).Num).Faction > 0 And NPC(MapNPC(MapNum).NPC(n).Num).Faction > 0 Then
                        If NPC(MapNPC(MapNum).NPC(MapNPCNum).Num).Faction = NPC(MapNPC(MapNum).NPC(n).Num).Faction Then
                            If NPC(MapNPC(MapNum).NPC(MapNPCNum).Num).Behavior = NPC_BEHAVIOR_ATTACKONSIGHT Then
                                If MapNPC(MapNum).NPC(n).target = 0 Then
                                    MapNPC(MapNum).NPC(n).targetType = TARGET_TYPE_NPC
                                    MapNPC(MapNum).NPC(n).target = MapNPCNum
                                    Call SendMapNPCTarget(MapNum, n, MapNPC(MapNum).NPC(n).target, MapNPC(MapNum).NPC(n).targetType)
                                End If
                            End If
                        End If
                    End If
                End If
            Next
        End If
        
        ' Don't help player killers
        If GetPlayerPK(Index) = NO Then
            ' Send threat to all guards which are in range
            For n = 1 To Map(MapNum).NPC_HighIndex
                If MapNPC(MapNum).NPC(n).Num > 0 Then
                    If NPC(MapNPC(MapNum).NPC(n).Num).Behavior = NPC_BEHAVIOR_GUARD Then
                        ' X range
                        If MapNPC(MapNum).NPC(n).X > GetPlayerX(Index) Then
                            DistanceX = MapNPC(MapNum).NPC(n).X - GetPlayerX(Index)
                        Else
                            DistanceX = GetPlayerX(Index) - MapNPC(MapNum).NPC(n).X
                        End If
                        
                        ' Y range
                        If MapNPC(MapNum).NPC(n).Y > GetPlayerY(Index) Then
                            DistanceY = MapNPC(MapNum).NPC(n).Y - GetPlayerY(Index)
                        Else
                            DistanceY = GetPlayerY(Index) - MapNPC(MapNum).NPC(n).Y
                        End If
                        
                        n = NPC(MapNPC(MapNum).NPC(n).Num).Range
                                
                        ' Are they in range
                        If DistanceX <= n And DistanceY <= n Then
                            If MapNPC(MapNum).NPC(n).target = 0 Then
                                MapNPC(MapNum).NPC(n).targetType = TARGET_TYPE_NPC
                                MapNPC(MapNum).NPC(n).target = MapNPCNum
                                Call SendMapNPCTarget(MapNum, n, MapNPC(MapNum).NPC(n).target, MapNPC(MapNum).NPC(n).targetType)
                            End If
                        End If
                    End If
                End If
            Next
        End If
        
        ' Get the damage we can do
        Damage = GetNPCDamage(NPCNum)
        
        ' Add damage based on direction
        If MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_UP Then
            If Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_LEFT Or Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_UP Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_DOWN Then
            If Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_LEFT Or Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_LEFT Then
            If Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_UP Or Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_LEFT Then
                Damage = Damage + (Damage / 4)
            End If
        ElseIf MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Dir = DIR_RIGHT Then
            If Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_UP Or Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_DOWN Then
                Damage = Damage + (Damage / 10)
            ElseIf Account(Index).Chars(GetPlayerChar(Index)).Dir = DIR_RIGHT Then
                Damage = Damage + (Damage / 4)
            End If
        End If
        
        ' 1.5 times the damage if it's a critical
        If CanNPCCritical(MapNPCNum) Then
            Damage = Damage * 1.5
            Call SendSoundToMap(MapNum, Options.CriticalSound)
            SendAnimation MapNum, Options.CriticalAnimation, 0, 0, TARGET_TYPE_NPC, MapNPCNum
            Exit Sub
        End If
        
        ' Take away protection from the damage
        Damage = Damage - GetPlayerProtection(Index)
        
        ' Randomize damage
        Damage = Random(Damage - (Damage / 2), Damage)
        
        Round Damage
        
        If Damage < 1 Then
            Call SendSoundToMap(MapNum, Options.MissSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_PLAYER, Index
            Exit Sub
        End If
        
        ' Send the sound
        Call SendMapSound(MapNum, Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seAnimation, 1)

        Call NPCAttackPlayer(MapNPCNum, Index, Damage)
    End If
End Sub

Function CanNPCAttackPlayer(ByVal MapNPCNum As Long, ByVal Index As Long, Optional ByVal Spell As Boolean = False) As Boolean
    Dim MapNum As Integer
    Dim NPCNum As Long

    ' Check for subscript out of range
    If MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Or Not IsPlaying(Index) Then Exit Function

    ' Check for subscript out of range
    If MapNPC(GetPlayerMap(Index)).NPC(MapNPCNum).Num < 1 Then Exit Function

    MapNum = GetPlayerMap(Index)
    NPCNum = MapNPC(MapNum).NPC(MapNPCNum).Num

    ' Make sure the npc isn't already dead
    If MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.HP) < 1 Then Exit Function
    
    ' Make sure they aren't casting a spell
    If MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Timer > 0 And Spell = False Then Exit Function
    
    ' Can't attack if shopkeeper, friendly, or quest
    If NPC(MapNPC(MapNum).NPC(MapNPCNum).Num).Behavior = NPC_BEHAVIOR_QUEST Then Exit Function
    
    ' Don't attack players who are not Player Killers if the attack is a guard
    If NPC(MapNPC(MapNum).NPC(MapNPCNum).Num).Behavior = NPC_BEHAVIOR_GUARD Then
        If GetPlayerPK(Index) = NO Then Exit Function
    End If
    
    ' Make sure we dont attack the player if they are switching maps
    If TempPlayer(Index).GettingMap = YES Then Exit Function
    
    ' Make sure npcs don't attack more than once a second
    If timeGetTime < MapNPC(MapNum).NPC(MapNPCNum).AttackTimer + 1000 And Spell = False Then Exit Function

    If Spell Then
        CanNPCAttackPlayer = True
        Exit Function
    End If
    
    ' Adjust target if they have none
    If TempPlayer(Index).target = 0 Then
        TempPlayer(Index).target = MapNPCNum
        TempPlayer(Index).targetType = TARGET_TYPE_NPC
        Call SendPlayerTarget(Index)
    End If
    
    ' Make sure they are on the same map
    If IsPlaying(Index) Then
        If NPCNum > 0 Then
            ' Check if they are going to cast
            If Random(1, 2) = 1 And CanNPCCastSpell(MapNum, MapNPCNum) Then
                Call BufferNPCSpell(MapNum, MapNPCNum, MapNPC(MapNum).NPC(MapNPCNum).target)
                Exit Function
            End If
            
            ' Check if at same coordinates
            If (GetPlayerY(Index) + 1 = MapNPC(MapNum).NPC(MapNPCNum).Y) And (GetPlayerX(Index) = MapNPC(MapNum).NPC(MapNPCNum).X) Then
            ElseIf (GetPlayerY(Index) - 1 = MapNPC(MapNum).NPC(MapNPCNum).Y) And (GetPlayerX(Index) = MapNPC(MapNum).NPC(MapNPCNum).X) Then
            ElseIf (GetPlayerY(Index) = MapNPC(MapNum).NPC(MapNPCNum).Y) And (GetPlayerX(Index) + 1 = MapNPC(MapNum).NPC(MapNPCNum).X) Then
            ElseIf (GetPlayerY(Index) = MapNPC(MapNum).NPC(MapNPCNum).Y) And (GetPlayerX(Index) - 1 = MapNPC(MapNum).NPC(MapNPCNum).X) Then
            Else
                Exit Function
            End If
            
            ' Send this packet so they can see the npc attacking
            Call SendNPCAttack(MapNPCNum, MapNum)
            
            CanNPCAttackPlayer = True
        End If
    End If
End Function

Private Function DidPlayerMitigateNPC(ByVal MapNum As Integer, ByVal Index As Long, ByVal MapNPCNum As Long) As Boolean
    If CanPlayerMitigateNPC(Index, MapNPCNum) = True Or MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Spell > 0 Then
        ' Check if player can avoid the attack
        If CanPlayerDodge(Index) Then
            Call SendSoundToMap(MapNum, Options.DodgeSound)
            SendAnimation MapNum, Options.DodgeAnimation, 0, 0, TARGET_TYPE_PLAYER, Index
            DidPlayerMitigateNPC = True
            Exit Function
        End If
        
        ' Check if player can deflect the attack
        If CanPlayerDeflect(Index) Then
            Call SendSoundToMap(MapNum, Options.DeflectSound)
            SendAnimation MapNum, Options.DeflectAnimation, 0, 0, TARGET_TYPE_PLAYER, Index
            DidPlayerMitigateNPC = True
            Exit Function
        End If
        
        ' Check if player can block the attack
        If CanPlayerBlock(Index) Then
            Call SendSoundToMap(MapNum, Options.BlockSound)
            SendAnimation MapNum, Options.DeflectAnimation, 0, 0, TARGET_TYPE_PLAYER, Index
            DidPlayerMitigateNPC = True
            Exit Function
        End If
    End If
End Function

Sub NPCAttackPlayer(ByVal MapNPCNum As Long, ByVal Victim As Long, ByVal Damage As Long)
    Dim Name As String
    Dim Exp As Long
    Dim MapNum As Integer
    Dim I As Long
    Dim n As Long
    Dim DistanceX As Byte, DistanceY As Byte

    ' Check for subscript out of range
    If MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then Exit Sub

    ' Check for subscript out of range
    If MapNPC(GetPlayerMap(Victim)).NPC(MapNPCNum).Num < 1 Or MapNPC(GetPlayerMap(Victim)).NPC(MapNPCNum).Num > MAX_MAP_NPCS Then Exit Sub

    If DidPlayerMitigateNPC(GetPlayerMap(Victim), Victim, MapNPCNum) = True Then Exit Sub
    
    MapNum = GetPlayerMap(Victim)
    Name = Trim$(NPC(MapNPC(MapNum).NPC(MapNPCNum).Num).Name)
    
    ' Reduce durability on the victim's equipment
    If Random(1, 2) = 1 Then ' Which one it will affect
        Call DamagePlayerEquipment(Victim, Equipment.Body)
    Else
        Call DamagePlayerEquipment(Victim, Equipment.Head)
    End If
    
    ' Set the regen timer
    MapNPC(MapNum).NPC(MapNPCNum).StopRegen = True
    MapNPC(MapNum).NPC(MapNPCNum).StopRegenTimer = timeGetTime

    If Damage >= GetPlayerVital(Victim, Vitals.HP) Then
        ' Set the damage to the player's health so that it doesn't appear that it's overkilling it
        Damage = GetPlayerVital(Victim, Vitals.HP)
        
        ' Say damage
        SendActionMsg GetPlayerMap(Victim), "-" & Damage, BrightRed, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
        
        ' Kill player
        KillPlayer Victim
        
        Call GlobalMsg(GetPlayerName(Victim) & " has been killed by " & CheckGrammar(Trim$(Name)) & "!", BrightRed)

        ' Set npc target to 0
        MapNPC(MapNum).NPC(MapNPCNum).target = 0
        MapNPC(MapNum).NPC(MapNPCNum).targetType = TARGET_TYPE_NONE
        Call SendMapNPCTarget(MapNum, MapNPCNum, 0, 0)
    Else
        ' Player not dead, just do the damage
        Call SetPlayerVital(Victim, Vitals.HP, GetPlayerVital(Victim, Vitals.HP) - Damage)
        Call SendVital(Victim, Vitals.HP)
        
        ' Send animation
        If MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Spell = 0 Then
            If NPC(MapNPC(GetPlayerMap(Victim)).NPC(MapNPCNum).Num).Animation < 1 Then
                Call SendAnimation(GetPlayerMap(Victim), 1, 0, 0, TARGET_TYPE_PLAYER, Victim)
            Else
               Call SendAnimation(GetPlayerMap(Victim), NPC(MapNPC(GetPlayerMap(Victim)).NPC(MapNPCNum).Num).Animation, 0, 0, TARGET_TYPE_PLAYER, Victim)
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

Public Sub BufferNPCSpell(ByVal MapNum As Integer, ByVal MapNPCNum As Long, ByVal target As Long)
    Dim SpellNum As Long
    Dim SpellCastType As Byte
    Dim Range As Long
    Dim HasBuffered As Boolean
    
    SpellNum = MapNPC(MapNum).NPC(MapNPCNum).ActiveSpell
    
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
        SendAnimation MapNum, Spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_NPC, MapNPCNum
        
        If Spell(SpellNum).CastTime > 0 Then
            SendActionMsg MapNum, "Casting " & Trim$(Spell(SpellNum).Name), BrightBlue, ACTIONMSG_SCROLL, MapNPC(MapNum).NPC(MapNPCNum).X * 32, MapNPC(MapNum).NPC(MapNPCNum).Y * 32
        End If
        
        MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Spell = SpellNum
        MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Timer = timeGetTime
        MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.target = MapNPC(MapNum).NPC(MapNPCNum).target
        MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.TType = MapNPC(MapNum).NPC(MapNPCNum).targetType
        Call SendNPCSpellBuffer(MapNum, MapNPCNum)
    End If
End Sub

Sub NPCSpellPlayer(ByVal MapNPCNum As Long, ByVal Victim As Long)
    Dim MapNum As Integer
    Dim I As Long
    Dim Damage As Long
    Dim SpellNum As Long
    Dim DidCast As Boolean, X As Byte, Y As Byte, AoE As Long

    ' Check for subscript out of range
    If MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Or IsPlaying(Victim) = False Then Exit Sub
        
    ' Check for subscript out of range
    If MapNPC(GetPlayerMap(Victim)).NPC(MapNPCNum).Num < 1 Then Exit Sub

    ' Set the map number
    MapNum = GetPlayerMap(Victim)
     
    ' Set the spell that they are going to cast
    SpellNum = MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Spell
   
    ' Send this packet so they can see the person attacking
    Call SendNPCAttack(MapNPCNum, MapNum)
    
    ' Play the sound
    Call SendMapSound(MapNum, Victim, GetPlayerX(Victim), GetPlayerY(Victim), SoundEntity.seSpell, SpellNum)
    
    DidCast = False
    AoE = Spell(SpellNum).AoE
    X = MapNPC(MapNum).NPC(MapNPC(MapNum).NPC(MapNPCNum).target).X
    Y = MapNPC(MapNum).NPC(MapNPC(MapNum).NPC(MapNPCNum).target).Y
    
    ' Check if the spell they are going to cast is valid
    If SpellNum > 0 Then
        If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Or Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
            Damage = GetNPCSpellVital(MapNum, MapNPCNum, Victim, SpellNum, True)
            
            If Spell(SpellNum).IsAoe = True Then ' AoE
                If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = MapNum Then
                                If IsInRange(AoE, X, Y, Account(I).Chars(GetPlayerChar(I)).X, Account(I).Chars(GetPlayerChar(I)).Y) Then
                                    Account(I).Chars(GetPlayerChar(I)).Vital(Vitals.HP) = Account(I).Chars(GetPlayerChar(I)).Vital(Vitals.HP) + Damage
                                    SendActionMsg MapNum, "+" & Damage, BrightGreen, 1, (Account(I).Chars(GetPlayerChar(I)).X * 32), (Account(I).Chars(GetPlayerChar(I)).Y * 32)
                                    
                                    Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, I)
                                    DidCast = True
                                    
                                    ' Prevent overhealing
                                    If Account(I).Chars(GetPlayerChar(I)).Vital(Vitals.HP) > GetPlayerMaxVital(I, HP) Then
                                        Account(I).Chars(GetPlayerChar(I)).Vital(Vitals.HP) = GetPlayerMaxVital(I, HP)
                                    End If
                                End If
                            End If
                        End If
                    Next
                Else
                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = MapNum Then
                                If IsInRange(AoE, X, Y, Account(I).Chars(GetPlayerChar(I)).X, Account(I).Chars(GetPlayerChar(I)).Y) Then
                                    Account(I).Chars(GetPlayerChar(I)).Vital(Vitals.MP) = Account(I).Chars(GetPlayerChar(I)).Vital(Vitals.MP) + Damage
                                    SendActionMsg MapNum, "+" & Damage, BrightBlue, 1, (Account(I).Chars(GetPlayerChar(I)).X * 32), (Account(I).Chars(GetPlayerChar(I)).Y * 32)
                                    
                                    Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, I)
                                    DidCast = True
                                    
                                    ' Prevent overhealing
                                    If Account(I).Chars(GetPlayerChar(I)).Vital(Vitals.MP) > GetPlayerMaxVital(I, MP) Then
                                        Account(I).Chars(GetPlayerChar(I)).Vital(Vitals.MP) = GetPlayerMaxVital(I, MP)
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
                For I = 1 To Player_HighIndex
                    If IsPlaying(I) Then
                        If GetPlayerMap(I) = MapNum Then
                            If IsInRange(AoE, X, Y, Account(I).Chars(GetPlayerChar(I)).X, Account(I).Chars(GetPlayerChar(I)).Y) Then
                                If CanNPCAttackPlayer(MapNPCNum, I, True) Then
                                    Damage = GetNPCSpellVital(MapNum, MapNPCNum, I, SpellNum)
                                    Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, I)
                                    
                                    If Damage < 1 Then
                                        Call SendSoundToMap(GetPlayerMap(I), Options.ResistSound)
                                        SendActionMsg GetPlayerMap(I), "Resist", Pink, 1, (GetPlayerX(I) * 32), (GetPlayerY(I) * 32)
                                    Else
                                        Call NPCAttackPlayer(MapNPCNum, I, Damage)
                                        DidCast = True
                                    End If
                                End If
                            End If
                        End If
                    End If
                Next
            Else ' Non AoE
                Damage = GetNPCSpellVital(MapNum, MapNPCNum, Victim, SpellNum)
                Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Victim)
                
                If Damage < 1 Then
                    Call SendSoundToMap(GetPlayerMap(Victim), Options.ResistSound)
                    SendActionMsg GetPlayerMap(Victim), "Resist", Pink, 1, (GetPlayerX(Victim) * 32), (GetPlayerY(Victim) * 32)
                Else
                    Call NPCAttackPlayer(MapNPCNum, Victim, Damage)
                    DidCast = True
                End If
            End If
        End If
    End If
    
    If DidCast Then
        MapNPC(MapNum).NPC(MapNPCNum).Vital(MP) = MapNPC(MapNum).NPC(MapNPCNum).Vital(MP) - Spell(SpellNum).MPCost
        Call SendMapNPCVitals(MapNum, MapNPCNum)
        MapNPC(MapNum).NPC(MapNPCNum).AttackTimer = timeGetTime
        
        MapNPC(MapNum).NPC(MapNPCNum).SpellTimer(MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Spell) = timeGetTime + Spell(SpellNum).CDTime * 1000
        SendActionMsg MapNum, Trim$(Spell(SpellNum).Name), BrightBlue, ACTIONMSG_SCROLL, MapNPC(MapNum).NPC(MapNPCNum).X * 32, MapNPC(MapNum).NPC(MapNPCNum).Y * 32
    End If
End Sub

Sub NPCSpellNPC(ByVal MapNPCNum As Long, ByVal Victim As Long, MapNum As Integer)
    Dim I As Long
    Dim Damage As Long
    Dim SpellNum As Long
    Dim DidCast As Boolean, AoE As Long, X As Byte, Y As Byte

    ' Check for subscript out of range
    If MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Or Victim < 1 Or Victim > MAX_MAP_NPCS Then Exit Sub
        
    ' Check for subscript out of range
    If MapNPC(MapNum).NPC(MapNPCNum).Num < 1 Or MapNPC(MapNum).NPC(Victim).Num < 1 Then Exit Sub
    
    ' Set the spell that they are going to cast
    SpellNum = MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Spell
   
    ' Send this packet so they can see the person attacking
    Call SendNPCAttack(MapNPCNum, MapNum)
    
    ' Play the sound
    Call SendMapSound(MapNum, Victim, MapNPC(MapNum).NPC(Victim).X, MapNPC(MapNum).NPC(Victim).Y, SoundEntity.seSpell, SpellNum)
    
    DidCast = False
    AoE = Spell(SpellNum).AoE
    X = MapNPC(MapNum).NPC(MapNPC(MapNum).NPC(MapNPCNum).target).X
    Y = MapNPC(MapNum).NPC(MapNPC(MapNum).NPC(MapNPCNum).target).Y
    
    ' Check if the spell they are going to cast is valid
    If SpellNum > 0 Then
        If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Or Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
            Damage = GetNPCSpellVital(MapNum, MapNPCNum, Victim, SpellNum, True)
            
            If Spell(SpellNum).IsAoe = True Then ' AoE
                If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                    For I = 1 To Map(MapNum).NPC_HighIndex
                        If MapNPC(MapNum).NPC(I).Num > 0 Then
                            If IsInRange(AoE, X, Y, MapNPC(MapNum).NPC(I).X, MapNPC(MapNum).NPC(I).Y) Then
                                MapNPC(MapNum).NPC(I).Vital(Vitals.HP) = MapNPC(MapNum).NPC(I).Vital(Vitals.HP) + Damage
                                SendActionMsg MapNum, "+" & Damage, BrightGreen, 1, (MapNPC(MapNum).NPC(I).X * 32), (MapNPC(MapNum).NPC(I).Y * 32)
                                
                                Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, I)
                                DidCast = True
                                
                                ' Prevent overhealing
                                If MapNPC(MapNum).NPC(I).Vital(Vitals.HP) > GetNPCMaxVital(MapNPC(MapNum).NPC(I).Num, HP) Then
                                    MapNPC(MapNum).NPC(I).Vital(Vitals.HP) = GetNPCMaxVital(MapNPC(MapNum).NPC(I).Num, HP)
                                End If
                            End If
                        End If
                    Next
                Else
                    For I = 1 To Map(MapNum).NPC_HighIndex
                        If MapNPC(MapNum).NPC(I).Num > 0 Then
                            If IsInRange(AoE, X, Y, MapNPC(MapNum).NPC(I).X, MapNPC(MapNum).NPC(I).Y) Then
                                MapNPC(MapNum).NPC(I).Vital(Vitals.MP) = MapNPC(MapNum).NPC(I).Vital(Vitals.MP) + Damage
                                SendActionMsg MapNum, "+" & Damage, BrightBlue, 1, (MapNPC(MapNum).NPC(I).X * 32), (MapNPC(MapNum).NPC(I).Y * 32)
                                
                                Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, I)
                                DidCast = True
                                
                                ' Prevent overhealing
                                If MapNPC(MapNum).NPC(I).Vital(Vitals.MP) > GetNPCMaxVital(MapNPC(MapNum).NPC(I).Num, MP) Then
                                    MapNPC(MapNum).NPC(I).Vital(Vitals.MP) = GetNPCMaxVital(MapNPC(MapNum).NPC(I).Num, MP)
                                End If
                            End If
                        End If
                    Next
                End If
            Else
                If Spell(SpellNum).Type = SPELL_TYPE_HEALHP Then
                    MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.HP) = MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.HP) + Damage
                    SendActionMsg MapNum, "+" & Damage, BrightGreen, 1, (MapNPC(MapNum).NPC(MapNPCNum).X * 32), (MapNPC(MapNum).NPC(MapNPCNum).Y * 32)
                    
                    ' Prevent overhealing
                    If MapNPC(MapNum).NPC(Victim).Vital(Vitals.HP) > GetNPCMaxVital(MapNPC(MapNum).NPC(Victim).Num, HP) Then
                        MapNPC(MapNum).NPC(Victim).Vital(Vitals.HP) = GetNPCMaxVital(MapNPC(MapNum).NPC(Victim).Num, HP)
                    End If
                ElseIf Spell(SpellNum).Type = SPELL_TYPE_HEALMP Then
                    MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.MP) = MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.MP) + Damage
                    SendActionMsg MapNum, "+" & Damage, BrightBlue, 1, (MapNPC(MapNum).NPC(MapNPCNum).X * 32), (MapNPC(MapNum).NPC(MapNPCNum).Y * 32)
                    
                    ' Prevent overhealing
                    If MapNPC(MapNum).NPC(Victim).Vital(Vitals.MP) > GetNPCMaxVital(MapNPC(MapNum).NPC(Victim).Num, MP) Then
                        MapNPC(MapNum).NPC(Victim).Vital(Vitals.MP) = GetNPCMaxVital(MapNPC(MapNum).NPC(Victim).Num, MP)
                    End If
                End If
            End If
            
            Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Victim)
            DidCast = True
        Else
            If Spell(SpellNum).IsAoe = True Then ' AoE
                For I = 1 To Map(MapNum).NPC_HighIndex
                    If MapNPC(MapNum).NPC(I).Num > 0 Then
                        If IsInRange(AoE, X, Y, MapNPC(MapNum).NPC(I).X, MapNPC(MapNum).NPC(I).Y) Then
                            If CanNPCAttackNPC(MapNum, MapNPCNum, I, True) Then
                                Damage = GetNPCSpellVital(MapNum, MapNPCNum, I, SpellNum)
                                Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, I)
                                
                                If Damage < 1 Then
                                    Call SendSoundToMap(MapNum, Options.ResistSound)
                                    SendActionMsg MapNum, "Resist", Pink, 1, (MapNPC(MapNum).NPC(I).X * 32), (MapNPC(MapNum).NPC(I).Y * 32)
                                Else
                                    Call NPCAttackNPC(MapNum, MapNPCNum, I, Damage)
                                    DidCast = True
                                End If
                            End If
                        End If
                    End If
                Next
            Else ' Non AoE
                Damage = GetNPCSpellVital(MapNum, MapNPCNum, Victim, SpellNum)
                Call SendAnimation(MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Victim)
                
                If Damage < 1 Then
                    Call SendSoundToMap(MapNum, Options.ResistSound)
                    SendActionMsg MapNum, "Resist", Pink, 1, (MapNPC(MapNum).NPC(MapNPC(MapNum).NPC(MapNPCNum).target).X * 32), (MapNPC(MapNum).NPC(MapNPC(MapNum).NPC(MapNPCNum).target).Y * 32)
                Else
                    Call NPCAttackNPC(MapNum, MapNPCNum, Victim, Damage)
                    DidCast = True
                End If
            End If
        End If
    End If
    
    If DidCast Then
        MapNPC(MapNum).NPC(MapNPCNum).Vital(MP) = MapNPC(MapNum).NPC(MapNPCNum).Vital(MP) - Spell(SpellNum).MPCost
        Call SendMapNPCVitals(MapNum, MapNPCNum)
        MapNPC(MapNum).NPC(MapNPCNum).AttackTimer = timeGetTime
        
        MapNPC(MapNum).NPC(MapNPCNum).SpellTimer(MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Spell) = timeGetTime + Spell(SpellNum).CDTime * 1000
        SendActionMsg MapNum, Trim$(Spell(SpellNum).Name), BrightBlue, ACTIONMSG_SCROLL, MapNPC(MapNum).NPC(MapNPCNum).X * 32, MapNPC(MapNum).NPC(MapNPCNum).Y * 32
    End If
End Sub

' ###################################
' ##    Player Attacking Player    ##
' ###################################
Public Sub TryPlayerAttackPlayer(ByVal Attacker As Long, ByVal Victim As Long)
    Dim NPCNum As Long
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
            Case DIR_UPLEFT
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_UPRIGHT
                If Not ((GetPlayerY(Victim) + 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWNLEFT
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) + 1 = GetPlayerX(Attacker))) Then Exit Function
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case DIR_DOWNRIGHT
                If Not ((GetPlayerY(Victim) = GetPlayerY(Attacker)) And (GetPlayerX(Victim) - 1 = GetPlayerX(Attacker))) Then Exit Function
                If Not ((GetPlayerY(Victim) - 1 = GetPlayerY(Attacker)) And (GetPlayerX(Victim) = GetPlayerX(Attacker))) Then Exit Function
            Case Else
                Exit Function
        End Select
    End If
        
    ' Check if map is attackable
    If Moral(Map(GetPlayerMap(Attacker)).Moral).CanPK = 0 Then
        If (GetPlayerPK(Victim) <> PLAYER_KILLER Or GetPlayerPK(Attacker) <> PLAYER_DEFENDER) And (GetPlayerPK(Victim) <> PLAYER_DEFENDER Or GetPlayerPK(Attacker) <> PLAYER_KILLER) Then
            Call PlayerMsg(Attacker, "This is a safe zone!", BrightRed)
            Exit Function
        End If
    End If
    
    ' Check to make sure that they dont have access
    If GetPlayerAccess(Attacker) > STAFF_MODERATOR Then
        Call PlayerMsg(Attacker, "You can't attack " & GetPlayerName(Victim) & "!", BrightBlue)
        Exit Function
    End If

    ' Check to make sure the victim isn't an admin
    If GetPlayerAccess(Victim) > STAFF_MODERATOR Then
        Call PlayerMsg(Attacker, "You can't attack " & GetPlayerName(Victim) & "!", BrightRed)
        Exit Function
    End If

    ' Make sure attacker is high enough level
    If GetPlayerLevel(Attacker) < Options.PKLevel Then
        Call PlayerMsg(Attacker, "You are below level " & Options.PKLevel & ", you Can't attack another player!", BrightRed)
        Exit Function
    End If

    ' Make sure victim is high enough level
    If GetPlayerLevel(Victim) < Options.PKLevel Then
        Call PlayerMsg(Attacker, GetPlayerName(Victim) & " is below level " & Options.PKLevel & ", you Can't attack this player!", BrightRed)
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
    If TempPlayer(Victim).target = 0 Then
        TempPlayer(Victim).target = Attacker
        TempPlayer(Victim).targetType = TARGET_TYPE_PLAYER
        Call SendPlayerTarget(Victim)
    End If
    
    If Not IsSpell Then
        ' Set the attack's target
        TempPlayer(Attacker).targetType = TARGET_TYPE_PLAYER
        TempPlayer(Attacker).target = Victim
        Call SendPlayerTarget(Attacker)
    
        ' Check attack timer
        If GetPlayerEquipment(Attacker, Weapon) > 0 Then
            If timeGetTime < TempPlayer(Attacker).AttackTimer + Item(GetPlayerEquipment(Attacker, Weapon)).WeaponSpeed Then Exit Function
        Else
            If timeGetTime < TempPlayer(Attacker).AttackTimer + 1000 Then Exit Function
        End If
    End If
    
    If CanPlayerMitigatePlayer(Attacker, Victim) Or TempPlayer(Attacker).SpellBuffer.Spell > 0 Then
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
    Dim I As Long
    Dim LevelDiff As Long

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
        Call DamagePlayerEquipment(Victim, Equipment.Body)
    Else
        Call DamagePlayerEquipment(Victim, Equipment.Head)
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
        For I = 1 To Player_HighIndex
            If IsPlaying(I) Then
                If Account(I).Chars(GetPlayerChar(I)).Map = GetPlayerMap(Attacker) Then
                    If TempPlayer(I).targetType = TARGET_TYPE_PLAYER Then
                        If TempPlayer(I).target = Victim Then
                            TempPlayer(I).target = 0
                            TempPlayer(I).targetType = TARGET_TYPE_NONE
                            SendPlayerTarget I
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
            If GetPlayerPK(Victim) = PLAYER_DEFENDER Then
                Call SetPlayerPK(Attacker, PLAYER_DEFENDER)
                Call SendPlayerPK(Attacker)
                Call GlobalMsg(GetPlayerName(Attacker) & " has been deemed a Player Defender!", BrightBlue)
            End If
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
    Call DamagePlayerEquipment(Attacker, Equipment.Weapon)
End Sub

' ############
' ## Spells ##
' ############
Public Sub BufferPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Byte)
    Dim SpellNum As Long
    Dim MPCost As Long
    Dim LevelReq As Long
    Dim MapNum As Integer
    Dim SpellCastType As Byte
    Dim ClassReq As Long
    Dim AccessReq As Long
    Dim Range As Long
    Dim HasBuffered As Boolean
    Dim targetType As Byte
    Dim target As Long
    
    ' Prevent subscript out of range
    If SpellSlot < 1 Or SpellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    SpellNum = GetPlayerSpell(Index, SpellSlot)
    
    If SpellNum < 1 Or SpellNum > MAX_SPELLS Then Exit Sub
    
    MapNum = GetPlayerMap(Index)
    
    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub
    
    ' See if cooldown has finished
    If Account(Index).Chars(GetPlayerChar(Index)).SpellCD(SpellSlot) > timeGetTime Then
        PlayerMsg Index, "Spell hasn't cooled down yet!", BrightRed
        Exit Sub
    End If

    MPCost = Spell(SpellNum).MPCost

    ' Check if they have enough MP
    If GetPlayerVital(Index, Vitals.MP) < MPCost Then
        Call PlayerMsg(Index, "Not enough mana!", BrightRed)
        Exit Sub
    End If
    
    LevelReq = Spell(SpellNum).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(Index) Then
        Call PlayerMsg(Index, "You must be level " & LevelReq & " to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    AccessReq = Spell(SpellNum).AccessReq
    
    ' Make sure they have the right access
    If AccessReq > GetPlayerAccess(Index) Then
        Call PlayerMsg(Index, "You must be a staff member to cast this spell.", BrightRed)
        Exit Sub
    End If
    
    ClassReq = Spell(SpellNum).ClassReq
    
    ' Make sure the ClassReq > 0
    If ClassReq > 0 Then ' 0 = no req
        If ClassReq <> GetPlayerClass(Index) Then
            Call PlayerMsg(Index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this spell.", BrightRed)
            Exit Sub
        End If
    End If
    
    ' Can't use items while in a map that doesn't allow it
    If Moral(Map(GetPlayerMap(Index)).Moral).CanCast = 0 Then Exit Sub

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
    
    targetType = TempPlayer(Index).targetType
    target = TempPlayer(Index).target
    Range = Spell(SpellNum).Range
    HasBuffered = False
    
    Select Case SpellCastType
        Case 0, 1 ' Self-cast & self-cast AOE
            HasBuffered = True
        Case 2, 3 ' targeted & targeted AOE
            ' Check if have target
            If Not target > 0 Then Exit Sub

            If targetType = TARGET_TYPE_PLAYER Then
                ' If have target, check in range
                If Not IsInRange(Range, GetPlayerX(Index), GetPlayerY(Index), GetPlayerX(target), GetPlayerY(target)) And Spell(SpellNum).CastTime = 0 Then
                    PlayerMsg Index, "Target is not in range!", BrightRed
                Else
                    ' Go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackPlayer(Index, target, False, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            ElseIf targetType = TARGET_TYPE_NPC Then
                ' If have target, check in range
                If Not IsInRange(Range, GetPlayerX(Index), GetPlayerY(Index), MapNPC(MapNum).NPC(target).X, MapNPC(MapNum).NPC(target).Y) And Spell(SpellNum).CastTime = 0 Then
                    PlayerMsg Index, "Target is not in range!", BrightRed
                Else
                    ' Go through spell types
                    If Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEHP And Spell(SpellNum).Type <> SPELL_TYPE_DAMAGEMP Then
                        HasBuffered = True
                    Else
                        If CanPlayerAttackNPC(Index, target, True) Then
                            HasBuffered = True
                        End If
                    End If
                End If
            End If
    End Select
    
    If HasBuffered Then
        SendAnimation MapNum, Spell(SpellNum).CastAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        TempPlayer(Index).SpellBuffer.Spell = SpellSlot
        TempPlayer(Index).SpellBuffer.Timer = timeGetTime
        TempPlayer(Index).SpellBuffer.target = TempPlayer(Index).target
        TempPlayer(Index).SpellBuffer.TType = TempPlayer(Index).targetType
    End If
End Sub

Public Sub CastSpell(ByVal Index As Long, ByVal SpellSlot As Byte, ByVal target As Long, ByVal targetType As Byte)
    Dim SpellNum As Long
    Dim MapNum As Integer
    Dim Vital As Long
    Dim DidCast As Boolean
    Dim I As Long
    Dim AoE As Long
    Dim Range As Byte
    Dim VitalType As Byte
    Dim Increment As Boolean
    Dim X As Long, Y As Long
    Dim SpellCastType As Long
    Dim MPCost As Long

    ' Prevent subscript out of range
    If SpellSlot < 1 Or SpellSlot > MAX_PLAYER_SPELLS Then Exit Sub

    SpellNum = GetPlayerSpell(Index, SpellSlot)
    MapNum = GetPlayerMap(Index)
    MPCost = Spell(SpellNum).MPCost
    
    ' Make sure player has the spell
    If Not HasSpell(Index, SpellNum) Then Exit Sub
    
    ' They are stunned, don't allow them to cast
    If TempPlayer(Index).StunDuration > 0 Then Exit Sub
    
    ' Make sure they meet the requirements
    If CanPlayerCastSpell(Index, SpellNum) = False Then Exit Sub
    
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
        Vital = Spell(SpellNum).Vital + GetPlayerDamage(Index)
    Else
        Vital = Spell(SpellNum).Vital
    End If
    
    AoE = Spell(SpellNum).AoE
    Range = Spell(SpellNum).Range
    
    ' Add damage based on intelligence
    Vital = Vital + GetPlayerStat(Index, Intelligence) / 3
    
    ' Randomize the vital
    Vital = Random(Vital - (Vital / 2), Vital)
    
    ' 1.5 times the damage if it's a critical
    If CanPlayerSpellCritical(Index) Then
        Vital = Vital * 1.5
        Call SendSoundToMap(MapNum, Options.CriticalSound)
        SendAnimation MapNum, Options.CriticalAnimation, 0, 0, TARGET_TYPE_PLAYER, Index
    End If
    
    Select Case SpellCastType
        Case 0 ' Self-cast target
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_HEALHP
                    SpellPlayer_Effect Vitals.HP, True, Index, Vital, SpellNum, Index
                    ' Send the sound
                    SendMapSound MapNum, Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
                    DidCast = True
                Case SPELL_TYPE_HEALMP
                    SpellPlayer_Effect Vitals.MP, True, Index, Vital, SpellNum, Index
                    ' Send the sound
                    SendMapSound MapNum, Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
                    DidCast = True
                Case SPELL_TYPE_WARP
                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    PlayerWarp Index, Spell(SpellNum).Map, Spell(SpellNum).X, Spell(SpellNum).Y
                    ' Send the sound
                    SendMapSound MapNum, Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
                    DidCast = True
                Case SPELL_TYPE_RECALL
                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    WarpToCheckPoint (Index)
                    ' Send the sound
                    SendMapSound MapNum, Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
                    DidCast = True
                Case SPELL_TYPE_WARPTOTARGET
                    Call PlayerMsg(Index, "This spell has been made incorrectly, report this to a staff member!", BrightRed)
                    Exit Sub
            End Select
            
        Case 1, 3 ' Self-cast AOE & targetted AOE
            If SpellCastType = 1 Then
                X = GetPlayerX(Index)
                Y = GetPlayerY(Index)
            ElseIf SpellCastType = 3 Then
                If targetType = 0 Then Exit Sub
                If target = 0 Then Exit Sub
                
                If targetType = TARGET_TYPE_PLAYER Then
                    X = GetPlayerX(target)
                    Y = GetPlayerY(target)
                Else
                    X = MapNPC(MapNum).NPC(target).X
                    Y = MapNPC(MapNum).NPC(target).Y
                End If
                
                If Not IsInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                    PlayerMsg Index, "Target is not in range!", BrightRed
                    ClearAccountSpellBuffer Index
                End If
            End If
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    DidCast = True
                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If Not I = Index Then
                                If GetPlayerMap(I) = GetPlayerMap(Index) Then
                                    If IsInRange(AoE, X, Y, GetPlayerX(I), GetPlayerY(I)) Then
                                        If CanPlayerAttackPlayer(Index, I, False, True) Then
                                            SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, I
                                            PlayerAttackPlayer Index, I, Vital, SpellNum
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    For I = 1 To Map(MapNum).NPC_HighIndex
                        If MapNPC(MapNum).NPC(I).Num > 0 Then
                            If MapNPC(MapNum).NPC(I).Vital(HP) > 0 Then
                                If IsInRange(AoE, X, Y, MapNPC(MapNum).NPC(I).X, MapNPC(MapNum).NPC(I).Y) Then
                                    ' Friendly and Shopkeeper
                                    If Not NPC(MapNPC(MapNum).NPC(I).Num).Behavior = NPC_BEHAVIOR_QUEST Then
                                        ' Guard
                                        If Not NPC(MapNPC(MapNum).NPC(I).Num).Behavior = NPC_BEHAVIOR_GUARD Or (NPC(MapNPC(MapNum).NPC(I).Num).Behavior = NPC_BEHAVIOR_GUARD And GetPlayerPK(Index) = PLAYER_KILLER) Then
                                            If CanPlayerAttackNPC(Index, I, True) Then
                                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, I
                                                PlayerAttackNPC Index, I, Vital, SpellNum
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
                    
                    For I = 1 To Player_HighIndex
                        If IsPlaying(I) Then
                            If GetPlayerMap(I) = GetPlayerMap(Index) Then
                                If IsInRange(AoE, X, Y, GetPlayerX(I), GetPlayerY(I)) Then
                                    SpellPlayer_Effect VitalType, Increment, I, Vital, SpellNum, Index
                                End If
                            End If
                        End If
                    Next
                    
                    For I = 1 To Map(MapNum).NPC_HighIndex
                        If MapNPC(MapNum).NPC(I).Num > 0 Then
                            If (Increment = True And NPC(MapNPC(MapNum).NPC(I).Num).Behavior = NPC_BEHAVIOR_GUARD And Account(Index).Chars(GetPlayerChar(Index)).PK = NO) Or Increment = False Then
                                If MapNPC(MapNum).NPC(I).Vital(HP) > 0 Then
                                    If IsInRange(AoE, X, Y, MapNPC(MapNum).NPC(I).X, MapNPC(MapNum).NPC(I).Y) Then
                                        SpellNPC_Effect VitalType, Increment, I, Vital, SpellNum, MapNum, Index
                                    End If
                                End If
                            End If
                        End If
                    Next
            End Select
            
        Case 2 ' Targetted
            If targetType = 0 Then Exit Sub
            If target = 0 Then Exit Sub
            
            If targetType = TARGET_TYPE_PLAYER Then
                X = GetPlayerX(target)
                Y = GetPlayerY(target)
            Else
                X = MapNPC(MapNum).NPC(target).X
                Y = MapNPC(MapNum).NPC(target).Y
            End If
            
            If Not IsInRange(Range, GetPlayerX(Index), GetPlayerY(Index), X, Y) Then
                ClearAccountSpellBuffer Index
                Exit Sub
            End If
            
            Select Case Spell(SpellNum).Type
                Case SPELL_TYPE_DAMAGEHP
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Index <> target Then
                            If CanPlayerAttackPlayer(Index, target, False, True) Then
                                SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, target
                                
                                If Vital > 0 Then
                                    PlayerAttackPlayer Index, target, Vital, SpellNum
                                    DidCast = True
                                Else
                                    Call SendSoundToMap(GetPlayerMap(I), Options.ResistSound)
                                    SendActionMsg GetPlayerMap(I), "Resist", Pink, 1, (Account(target).Chars(GetPlayerChar(target)).X * 32), (Account(target).Chars(GetPlayerChar(target)).Y * 32)
                                End If
                            End If
                        Else
                            Call PlayerMsg(Index, "You can't cast that spell on yourself!", 12)
                        End If
                    Else
                        If CanPlayerAttackNPC(Index, target, True) Then
                            SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, target
                            
                            If Vital > 0 Then
                                PlayerAttackNPC Index, target, Vital, SpellNum
                                DidCast = True
                            Else
                                Call SendSoundToMap(MapNum, Options.ResistSound)
                                SendActionMsg MapNum, "Resist", Pink, 1, (MapNPC(MapNum).NPC(target).X * 32), (MapNPC(MapNum).NPC(target).Y * 32)
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
                    
                    If targetType = TARGET_TYPE_PLAYER Then
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackPlayer(Index, target, False, True) Then
                                SpellPlayer_Effect VitalType, Increment, target, Vital, SpellNum, Index
                            End If
                        Else
                            SpellPlayer_Effect VitalType, Increment, target, Vital, SpellNum, Index
                        End If
                    ElseIf targetType = TARGET_TYPE_NPC And Increment = False Or NPC(MapNPC(MapNum).NPC(target).Num).Behavior = NPC_BEHAVIOR_GUARD And Account(Index).Chars(GetPlayerChar(Index)).PK = NO Then
                        If Spell(SpellNum).Type = SPELL_TYPE_DAMAGEMP Then
                            If CanPlayerAttackNPC(Index, target, True) Then
                                SpellNPC_Effect VitalType, Increment, target, Vital, SpellNum, MapNum, Index
                            End If
                        Else
                            SpellNPC_Effect VitalType, Increment, target, Vital, SpellNum, MapNum, Index
                        End If
                    Else
                        Call PlayerMsg(Index, "You are unable to cast your spell on this target!", 12)
                        Exit Sub
                    End If
                    
                Case SPELL_TYPE_WARPTOTARGET
                    Call PlayerWarp(Index, MapNum, X, Y)
                    SendAnimation MapNum, Spell(SpellNum).SpellAnim, X, Y
                    ' Send the sound
                    SendMapSound MapNum, Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum
                    DidCast = True
            End Select
    End Select
    
    If DidCast Then
        Call SetPlayerVital(Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - MPCost)
        Call SendVital(Index, Vitals.MP)
        TempPlayer(Index).SpellBuffer.Timer = timeGetTime + (Spell(SpellNum).CDTime * 1000)
        Call SetPlayerSpellCD(Index, SpellSlot, timeGetTime + (Spell(SpellNum).CDTime * 1000))
        Call SendSpellCooldown(Index, SpellSlot)
        SendActionMsg MapNum, Trim$(Spell(SpellNum).Name), BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
        
        ' Set the sprite
        If Spell(SpellNum).Sprite > 0 Then
            Call SetPlayerSprite(Index, Spell(SpellNum).Sprite)
            Call SendPlayerSprite(Index)
        End If
        
        If Spell(SpellNum).NewSpell > 0 And Spell(SpellNum).NewSpell <= MAX_SPELLS Then
            If Spell(Spell(SpellNum).NewSpell).CastRequired > 0 Then
                ' Add 1 to the amount of casts
                Account(Index).Chars(GetPlayerChar(Index)).AmountOfCasts(SpellSlot) = Account(Index).Chars(GetPlayerChar(Index)).AmountOfCasts(SpellSlot) + 1
                
                ' Check if a spell can rank up
                Call CheckSpellRankUp(Index, SpellNum, SpellSlot)
            End If
        End If
    End If
    
    Call ClearAccountSpellBuffer(Index)
End Sub

Public Sub SpellPlayer_Effect(ByVal Vital As Byte, ByVal Increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long, ByVal Caster As Long)
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

        SendAnimation GetPlayerMap(Index), Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
        SendActionMsg GetPlayerMap(Index), sSymbol & Damage, Color, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32

        ' send the sound
        SendMapSound GetPlayerMap(Index), Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seSpell, SpellNum

        If Increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) + Damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_Player Index, SpellNum
            End If
        ElseIf Not Increment Then
            SetPlayerVital Index, Vital, GetPlayerVital(Index, Vital) - Damage
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_Player Index, SpellNum, Caster
            End If
        End If
    End If
End Sub

Public Sub SpellNPC_Effect(ByVal Vital As Byte, ByVal Increment As Boolean, ByVal Index As Long, ByVal Damage As Long, ByVal SpellNum As Long, ByVal MapNum As Integer, ByVal Caster As Long)
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

        SendAnimation MapNum, Spell(SpellNum).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
        SendActionMsg MapNum, sSymbol & Damage, Color, ACTIONMSG_SCROLL, MapNPC(MapNum).NPC(Index).X * 32, MapNPC(MapNum).NPC(Index).Y * 32

        ' send the sound
        SendMapSound MapNum, Index, MapNPC(MapNum).NPC(Index).X, MapNPC(MapNum).NPC(Index).Y, SoundEntity.seSpell, SpellNum

        If Increment Then
            MapNPC(MapNum).NPC(Index).Vital(Vital) = MapNPC(MapNum).NPC(Index).Vital(Vital) + Damage
            If Spell(SpellNum).Duration > 0 Then
                AddHoT_NPC MapNum, Index, SpellNum
            End If
        ElseIf Not Increment Then
            MapNPC(MapNum).NPC(Index).Vital(Vital) = MapNPC(MapNum).NPC(Index).Vital(Vital) - Damage
            If Spell(SpellNum).Duration > 0 Then
                AddDoT_NPC MapNum, Index, SpellNum, Caster
            End If
        End If
    End If
End Sub

Public Sub AddDoT_Player(ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
    Dim I As Long

    For I = 1 To MAX_DOTS
        With TempPlayer(Index).DoT(I)
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

Public Sub AddHoT_Player(ByVal Index As Long, ByVal SpellNum As Long)
    Dim I As Long

    For I = 1 To MAX_DOTS
        With TempPlayer(Index).HoT(I)
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

Public Sub AddDoT_NPC(ByVal MapNum As Integer, ByVal Index As Long, ByVal SpellNum As Long, ByVal Caster As Long)
    Dim I As Long

    For I = 1 To MAX_DOTS
        With MapNPC(MapNum).NPC(Index).DoT(I)
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

Public Sub AddHoT_NPC(ByVal MapNum As Integer, ByVal Index As Long, ByVal SpellNum As Long)
    Dim I As Long

    For I = 1 To MAX_DOTS
        With MapNPC(MapNum).NPC(Index).HoT(I)
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

Public Sub HandleDoT_Player(ByVal Index As Long, ByVal dotNum As Long)
    With TempPlayer(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' Time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendAnimation GetPlayerMap(Index), Spell(.Spell).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                If CanPlayerAttackPlayer(.Caster, Index, True) Then
                    SendAnimation GetPlayerMap(Index), Spell(.Spell).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                    If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                        SendActionMsg GetPlayerMap(Index), "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                        SetPlayerVital Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) + Spell(.Spell).Vital
                        Call SendVital(Index, Vitals.HP)
                    Else
                        SendActionMsg GetPlayerMap(Index), "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                        SetPlayerVital Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) + Spell(.Spell).Vital
                        Call SendVital(Index, Vitals.MP)
                    End If
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

Public Sub HandleHoT_Player(ByVal Index As Long, ByVal hotNum As Long)
    With TempPlayer(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' Time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendAnimation GetPlayerMap(Index), Spell(.Spell).SpellAnim, 0, 0, TARGET_TYPE_PLAYER, Index
                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                    SendActionMsg GetPlayerMap(Index), "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SetPlayerVital Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) + Spell(.Spell).Vital
                    Call SendVital(Index, Vitals.HP)
                Else
                    SendActionMsg GetPlayerMap(Index), "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                    SetPlayerVital Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) + Spell(.Spell).Vital
                    Call SendVital(Index, Vitals.MP)
                End If
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

Public Sub HandleDoT_NPC(ByVal MapNum As Integer, ByVal Index As Long, ByVal dotNum As Long)
    With MapNPC(MapNum).NPC(Index).DoT(dotNum)
        If .Used And .Spell > 0 Then
            ' Time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendAnimation MapNum, Spell(.Spell).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
                If CanPlayerAttackNPC(.Caster, Index, True) Then
                    If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                        SendActionMsg GetPlayerMap(Index), "-" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                        SetPlayerVital Index, Vitals.HP, GetPlayerVital(Index, Vitals.HP) - Spell(.Spell).Vital
                        Call SendVital(Index, Vitals.HP)
                    Else
                        SendActionMsg GetPlayerMap(Index), "-" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(Index) * 32, GetPlayerY(Index) * 32
                        SetPlayerVital Index, Vitals.MP, GetPlayerVital(Index, Vitals.MP) - Spell(.Spell).Vital
                        Call SendVital(Index, Vitals.MP)
                    End If
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

Public Sub HandleHoT_NPC(ByVal MapNum As Integer, ByVal Index As Long, ByVal hotNum As Long)
    With MapNPC(MapNum).NPC(Index).HoT(hotNum)
        If .Used And .Spell > 0 Then
            ' Time to tick?
            If timeGetTime > .Timer + (Spell(.Spell).Interval * 1000) Then
                SendAnimation MapNum, Spell(.Spell).SpellAnim, 0, 0, TARGET_TYPE_NPC, Index
                If Spell(.Spell).Type = SPELL_TYPE_HEALHP Then
                    SendActionMsg MapNum, "+" & Spell(.Spell).Vital, BrightGreen, ACTIONMSG_SCROLL, MapNPC(MapNum).NPC(Index).X * 32, MapNPC(MapNum).NPC(Index).Y * 32
                    MapNPC(MapNum).NPC(Index).Vital(Vitals.HP) = MapNPC(MapNum).NPC(Index).Vital(Vitals.HP) + Spell(.Spell).Vital
                Else
                    SendActionMsg MapNum, "+" & Spell(.Spell).Vital, BrightBlue, ACTIONMSG_SCROLL, MapNPC(MapNum).NPC(Index).X * 32, MapNPC(MapNum).NPC(Index).Y * 32
                    MapNPC(MapNum).NPC(Index).Vital(Vitals.MP) = MapNPC(MapNum).NPC(Index).Vital(Vitals.MP) + Spell(.Spell).Vital
                End If
                
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

Public Sub StunPlayer(ByVal Index As Long, ByVal SpellNum As Long, Optional ByVal Interrupt As Boolean = True)
    ' Check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' Set the values on Index
        TempPlayer(Index).StunDuration = Spell(SpellNum).StunDuration
        TempPlayer(Index).StunTimer = timeGetTime
        
        ' Send it to the Index
        SendStunned Index

        ' tell him he's stunned
        If Interrupt Then
            SendActionMsg GetPlayerMap(Index), "Stunned", RGB(255, 128, 0), 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
            ClearAccountSpellBuffer Index
        End If
    End If
End Sub

Public Sub StunNPC(ByVal Index As Long, ByVal MapNum As Integer, ByVal SpellNum As Long)
    Dim NPCNum As Long
    
    NPCNum = MapNPC(MapNum).NPC(Index).Num
    
    ' Check if it's a stunning spell
    If Spell(SpellNum).StunDuration > 0 Then
        ' Set the values on index
        MapNPC(MapNum).NPC(Index).StunDuration = Spell(SpellNum).StunDuration
        MapNPC(MapNum).NPC(Index).StunTimer = timeGetTime
    End If
    
     ' Tell other players its stunned
    SendActionMsg MapNum, "Stunned", RGB(255, 128, 0), 1, (MapNPC(MapNum).NPC(Index).X * 32), (MapNPC(MapNum).NPC(Index).Y * 32)
End Sub

Public Sub ClearNPCSpellBuffer(ByVal MapNum As Integer, ByVal MapNPCNum As Byte)
    MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Spell = 0
    MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Timer = 0
    MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.target = 0
    MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.TType = 0
End Sub

Public Sub ClearAccountSpellBuffer(ByVal Index As Long)
    If TempPlayer(Index).SpellBuffer.Spell = 0 Then Exit Sub
    
    TempPlayer(Index).SpellBuffer.Spell = 0
    TempPlayer(Index).SpellBuffer.Timer = 0
    TempPlayer(Index).SpellBuffer.target = 0
    TempPlayer(Index).SpellBuffer.TType = 0
    Call SendClearAccountSpellBuffer(Index)
End Sub

Private Function CanNPCHealSelf(ByVal MapNum As Integer, ByVal MapNPCNum As Byte, ByVal SpellNum As Long) As Boolean
    Dim NPCNum As Long, target As Long, targetType As Byte
    
    NPCNum = MapNPC(MapNum).NPC(MapNPCNum).Num
    target = MapNPCNum
    targetType = TARGET_TYPE_NPC
    
    ' Valid spell
    If NPC(NPCNum).Spell(SpellNum) > 0 And NPC(NPCNum).Spell(SpellNum) <= MAX_SPELLS Then
        ' Check for cooldown
        If MapNPC(MapNum).NPC(MapNPCNum).SpellTimer(MapNPCNum) <= timeGetTime Then
            ' Have enough mana
            If MapNPC(MapNum).NPC(MapNPCNum).Vital(MP) - Spell(SpellNum).MPCost >= 0 Or Spell(SpellNum).MPCost = 0 Then
                If Spell(NPC(NPCNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALHP Or Spell(NPC(NPCNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALMP Then
                    ' Don't want to overheal
                    If NPC(NPCNum).Behavior = NPC_BEHAVIOR_GUARD And targetType = TARGET_TYPE_PLAYER Then
                        If Spell(NPC(NPCNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALHP Then
                            If target > 0 And target <= MAX_PLAYERS Then
                                If GetPlayerVital(target, Vitals.HP) + GetNPCSpellVital(MapNum, MapNPCNum, target, NPC(MapNPC(MapNum).NPC(target).Num).Spell(SpellNum), True) > GetPlayerMaxVital(MapNPC(MapNum).NPC(target).Num, HP) Then Exit Function
                            End If
                        ElseIf Spell(NPC(NPCNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALMP Then
                            If target > 0 And target <= MAX_PLAYERS Then
                                If GetPlayerVital(target, Vitals.MP) + GetNPCSpellVital(MapNum, MapNPCNum, target, NPC(MapNPC(MapNum).NPC(target).Num).Spell(SpellNum), True) > GetPlayerMaxVital(MapNPC(MapNum).NPC(target).Num, MP) Then Exit Function
                            End If
                        End If
                    ElseIf targetType = TARGET_TYPE_NPC And Not NPC(NPCNum).Behavior = NPC_BEHAVIOR_GUARD Then
                        If Spell(NPC(NPCNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALHP Then
                            If target = MapNPCNum Then
                                If MapNPC(MapNum).NPC(MapNPCNum).Vital(HP) + GetNPCSpellVital(MapNum, MapNPCNum, MapNPCNum, NPC(NPCNum).Spell(SpellNum), True) > GetNPCMaxVital(NPCNum, HP) Then Exit Function
                            ElseIf target > 0 And target <= MAX_MAP_NPCS Then
                                If MapNPC(MapNum).NPC(target).Vital(HP) + GetNPCSpellVital(MapNum, MapNPCNum, target, NPC(MapNPC(MapNum).NPC(target).Num).Spell(SpellNum), True) > GetNPCMaxVital(MapNPC(MapNum).NPC(target).Num, HP) Then Exit Function
                            End If
                        ElseIf Spell(NPC(NPCNum).Spell(SpellNum)).Type = SPELL_TYPE_HEALMP Then
                            If target = MapNPCNum Then
                                If MapNPC(MapNum).NPC(MapNPCNum).Vital(MP) + GetNPCSpellVital(MapNum, MapNPCNum, MapNPCNum, NPC(NPCNum).Spell(SpellNum), True) > GetNPCMaxVital(NPCNum, MP) Then Exit Function
                            ElseIf target > 0 And target <= MAX_MAP_NPCS Then
                                If MapNPC(MapNum).NPC(target).Vital(MP) + GetNPCSpellVital(MapNum, MapNPCNum, target, NPC(MapNPC(MapNum).NPC(target).Num).Spell(SpellNum), True) > GetNPCMaxVital(MapNPC(MapNum).NPC(target).Num, MP) Then Exit Function
                            End If
                        End If
                    End If
                    
                    CanNPCHealSelf = True
                    Exit Function
                End If
            End If
        End If
    End If
End Function

Private Function CanNPCCastSpell(ByVal MapNum As Integer, ByVal MapNPCNum As Byte) As Boolean
    Dim I As Long, target As Long, targetType As Byte, Range As Byte, NPCNum As Long
    Dim RndNum As Byte
    
    targetType = MapNPC(MapNum).NPC(MapNPCNum).targetType
    target = MapNPC(MapNum).NPC(MapNPCNum).target
    NPCNum = MapNPC(MapNum).NPC(MapNPCNum).Num

    ' Self-healing mechanics
    If MapNPC(MapNum).NPC(MapNPCNum).Vital(HP) < GetNPCMaxVital(MapNPC(MapNum).NPC(MapNPCNum).Num, HP) / 2 Then
        For I = 1 To MAX_NPC_SPELLS
            If CanNPCHealSelf(MapNum, MapNPCNum, I) Then
                MapNPC(MapNum).NPC(MapNPCNum).target = MapNPCNum
                MapNPC(MapNum).NPC(MapNPCNum).targetType = TARGET_TYPE_NPC
                CanNPCCastSpell = True
                MapNPC(MapNum).NPC(MapNPCNum).ActiveSpell = NPC(MapNPC(MapNum).NPC(MapNPCNum).Num).Spell(I)
                Exit Function
            End If
        Next
    End If
    
    For I = 1 To MAX_NPC_SPELLS
        ' Valid spell
        If NPC(NPCNum).Spell(I) > 0 And NPC(NPCNum).Spell(I) <= MAX_SPELLS Then
            ' Check for cooldown
            If MapNPC(MapNum).NPC(MapNPCNum).SpellTimer(I) <= timeGetTime Then
                ' Have enough mana?
                If MapNPC(MapNum).NPC(MapNPCNum).Vital(MP) - Spell(NPC(NPCNum).Spell(I)).MPCost >= 0 Or Spell(NPC(NPCNum).Spell(I)).MPCost = 0 Then
                    Range = Spell(NPC(NPCNum).Spell(I)).Range
                    
                    ' Are they in range
                    If targetType = TARGET_TYPE_PLAYER Then
                        If target > 0 And target <= MAX_PLAYERS Then
                            If Not IsInRange(Range, GetPlayerX(target), GetPlayerY(target), MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y) Then Exit Function
                        End If
                    ElseIf targetType = TARGET_TYPE_NPC Then
                        If target > 0 And target <= MAX_MAP_NPCS Then
                            If Not target = MapNPCNum Then
                                If Not IsInRange(Range, MapNPC(MapNum).NPC(target).X, MapNPC(MapNum).NPC(target).Y, MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y) Then Exit Function
                            End If
                        End If
                    End If
                    
                    CanNPCCastSpell = True
                    MapNPC(MapNum).NPC(MapNPCNum).ActiveSpell = NPC(MapNPC(MapNum).NPC(MapNPCNum).Num).Spell(I)
                    
                    ' Add random exits
                    If Random(1, MAX_NPC_SPELLS) = 1 Then Exit Function
                End If
            End If
        End If
    Next
End Function
