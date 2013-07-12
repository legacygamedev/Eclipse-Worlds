Attribute VB_Name = "modPlayer"
Option Explicit

Sub HandleUseChar(ByVal index As Long)
    If Not IsPlaying(index) Then
        Call JoinGame(index)
        Call AddLog(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Name & ".", "Player")
        Call TextAdd(GetPlayerLogin(index) & "/" & GetPlayerName(index) & " has began playing " & Options.Name & ".")
        Call UpdateCaption
    End If
End Sub

Sub JoinGame(ByVal index As Long)
    Dim i As Long
    Dim n As Long
    Dim Color As Long

    ' Set the flag so we know the person is in the game
    tempPlayer(index).InGame = True

    ' Update the log
    frmServer.lvwInfo.ListItems(index).SubItems(1) = GetPlayerIP(index)
    frmServer.lvwInfo.ListItems(index).SubItems(2) = GetPlayerLogin(index)
    frmServer.lvwInfo.ListItems(index).SubItems(3) = GetPlayerName(index)
    
    ' Send an ok to client to start receiving in game data
    Call SendLogin(index)

    TotalPlayersOnline = TotalPlayersOnline + 1
    
    ' Send data
    Call SendItems(index)
    Call SendAnimations(index)
    Call SendNpcs(index)
    Call SendShops(index)
    Call SendSpells(index)
    Call SendResources(index)
    Call SendInventory(index)
    Call CheckEquippedItems(index)
    Call SendPlayerEquipmentTo(index)
    Call SendHotbar(index)
    Call SendTitles(index)
    Call SendMorals(index)
    Call SendEmoticons(index)
    
    ' Spell Cooldowns
    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) > 0 And GetPlayerSpellCD(index, i) > 0 Then
            ' Check if the CD has expired
            If GetPlayerSpellCD(index, i) - timeGetTime < 1 Then Call SetPlayerSpellCD(index, i, 0)
            If GetPlayerSpellCD(index, i) - timeGetTime >= Spell(GetPlayerSpell(index, i)).CDTime * 1000 Then Call SetPlayerSpellCD(index, i, 0)
            If GetPlayerSpellCD(index, i) <= timeGetTime Then Call SetPlayerSpellCD(index, i, 0)
            
            ' Send it
            Call SendSpellCooldown(index, i)
        End If
    Next
    
    ' Check for glitches in the inventory
    Call UpdatePlayerInvItems(index)
    
    ' Send the player's data
    Call SendPlayerData(index)
    
    ' Send vitals to player of all other players online
    For n = 1 To Player_HighIndex
        For i = 1 To Vitals.Vital_Count - 1
            If IsPlaying(n) Then
                Call SendVitalTo(index, n, i) ' Sends all players to new player
                If Not index = n Then
                    Call SendVitalTo(n, index, i) ' Sends new player to logged in players
                End If
            End If
        Next
    Next
    
    ' Send other data
    Call SendPlayerStatus(index)
    Call SendPlayerExp(index)
    
    ' Warp the player to their saved location
    Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), True)
    
    ' Send welcome messages
    Call SendWelcome(index)

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next

    ' Send a global message that they joined
    If GetPlayerAccess(index) <= STAFF_MODERATOR Then
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Name & "!", Class(GetPlayerClass(index)).Color)
    Else
         ' Color for access
        Select Case GetPlayerAccess(index)
            Case 0
                Color = 15
            Case 1
                Color = 3
            Case 2
                Color = 2
            Case 3
                Color = BrightBlue
            Case 4
                Color = Yellow
            Case 5
                Color = RGB(255, 165, 0)
        End Select
            
        Call GlobalMsg(GetPlayerName(index) & " has joined " & Options.Name & "!", Color)
    End If

    ' Send the flag so they know they can start doing stuff
    Call SendInGame(index)
    
    ' Refresh the friends list to all players online
    For i = 1 To Player_HighIndex
        Call UpdateFriendsList(i)
    Next
    
    ' Refresh the foes list to all players online
    For i = 1 To Player_HighIndex
        Call UpdateFoesList(i)
    Next
    
    ' Update guild list
    If GetPlayerGuild(index) > 0 Then
        Call SendPlayerGuildMembers(index)
    End If
End Sub

Sub LeftGame(ByVal index As Long)
    Dim n As Long, i As Long
    Dim TradeTarget As Long

    If tempPlayer(index).InGame Then
        tempPlayer(index).InGame = False
        ' Check if player was the only player on the map and stop npc processing if so
        If GetTotalMapPlayers(GetPlayerMap(index)) < 1 Then
            PlayersOnMap(GetPlayerMap(index)) = NO
        End If
        
        ' Clear any invites out
        If tempPlayer(index).TradeRequest > 0 Or tempPlayer(index).PartyInvite > 0 Or tempPlayer(index).GuildInvite > 0 Then
            If tempPlayer(index).TradeRequest > 0 Then
                Call DeclineTradeRequest(index)
            End If
            
            If tempPlayer(index).PartyInvite > 0 Then
                Call Party_InviteDecline(tempPlayer(index).PartyInvite, index)
            End If
            
            If tempPlayer(index).GuildInvite > 0 Then
                Call DeclineGuildInvite(index)
            End If
        End If
        
        ' Cancel any trade they're in
        If tempPlayer(index).InTrade > 0 Then
            TradeTarget = tempPlayer(index).InTrade
            PlayerMsg TradeTarget, Trim$(GetPlayerName(index)) & " has declined the trade!", BrightRed
            
            ' Clear out trade
            For i = 1 To MAX_INV
                tempPlayer(TradeTarget).TradeOffer(i).Num = 0
                tempPlayer(TradeTarget).TradeOffer(i).Value = 0
            Next
            
            tempPlayer(TradeTarget).InTrade = 0
            SendCloseTrade TradeTarget
        End If
        
        ' Leave party
        Party_PlayerLeave index

        ' Loop through entire map and purge npc targets from player
        For i = 1 To Map(GetPlayerMap(index)).Npc_HighIndex
            If MapNpc(GetPlayerMap(index)).NPC(i).Num > 0 Then
                If MapNpc(GetPlayerMap(index)).NPC(i).TargetType = TARGET_TYPE_PLAYER Then
                    If MapNpc(GetPlayerMap(index)).NPC(i).Target = index Then
                        MapNpc(GetPlayerMap(index)).NPC(i).Target = 0
                        MapNpc(GetPlayerMap(index)).NPC(i).TargetType = TARGET_TYPE_NONE
                        Call SendMapNpcTarget(GetPlayerMap(index), i, 0, 0)
                    End If
                End If
            End If
        Next
        
        ' Refresh guild members
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                If Not i = index Then
                    If GetPlayerGuild(i) = GetPlayerGuild(index) Then
                        SendPlayerGuildMembers i, index
                    End If
                End If
            End If
        Next
        
        ' Send a global message that they left
        If GetPlayerAccess(index) <= STAFF_MODERATOR Then
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Name & "!", Grey)
        Else
            Call GlobalMsg(GetPlayerName(index) & " has left " & Options.Name & "!", DarkGrey)
        End If
        
        Call TextAdd(GetPlayerName(index) & " has disconnected from " & Options.Name & ".")
        Call SendLeftGame(index)
        TotalPlayersOnline = TotalPlayersOnline - 1
        
        ' Save and clear data
        Call SaveAccount(index)
        Call ClearAccount(index)
        
        ' Refresh the friends list of all players online
        For i = 1 To Player_HighIndex
            Call UpdateFriendsList(i)
        Next
        
        ' Refresh the foes list of all players online
        For i = 1 To Player_HighIndex
            Call UpdateFoesList(i)
        Next
    End If
End Sub

Sub PlayerWarp(ByVal index As Long, ByVal MapNum As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal NeedMap = False)
    Dim ShopNum As Long
    Dim OldMap As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub

    ' Check if you are out of bounds
    If X > Map(MapNum).MaxX Then X = Map(MapNum).MaxX
    If Y > Map(MapNum).MaxY Then Y = Map(MapNum).MaxY
    If X < 0 Then X = 0
    If Y < 0 Then Y = 0
    
    ' Save old map to send erase player data to
    OldMap = GetPlayerMap(index)
    
    UpdateMapBlock OldMap, GetPlayerX(index), GetPlayerY(index), False
    Call SetPlayerX(index, X)
    Call SetPlayerY(index, Y)
    UpdateMapBlock MapNum, X, Y, True
    
    ' if same map then just send their co-ordinates
    If MapNum = GetPlayerMap(index) And Not NeedMap Then
        Call SendPlayerWarp(index)
        Exit Sub
    End If
    
    ' Clear target
    tempPlayer(index).Target = 0
    tempPlayer(index).TargetType = TARGET_TYPE_NONE
    SendPlayerTarget index

    ' Loop through entire map and purge npc targets from player
    For i = 1 To Map(GetPlayerMap(index)).Npc_HighIndex
        If MapNpc(GetPlayerMap(index)).NPC(i).Num > 0 Then
            If MapNpc(GetPlayerMap(index)).NPC(i).TargetType = TARGET_TYPE_PLAYER Then
                If MapNpc(GetPlayerMap(index)).NPC(i).Target = index Then
                    MapNpc(GetPlayerMap(index)).NPC(i).Target = 0
                    MapNpc(GetPlayerMap(index)).NPC(i).TargetType = TARGET_TYPE_NONE
                    Call SendMapNpcTarget(OldMap, i, 0, 0)
                End If
            End If
        End If
    Next
    
    ' Leave the old map
    If Not OldMap = MapNum Then
        Call SendLeaveMap(index, OldMap)
    End If

    If Not OldMap = MapNum Then
        ' Set the new map
        Call SetPlayerMap(index, MapNum)
    End If
    
    ' Now we check if there were any players left on the map the player just left, and if not stop processing npcs
    If GetTotalMapPlayers(OldMap) = 0 Then
        PlayersOnMap(OldMap) = NO
        
        ' Get all npcs' vitals
        For i = 1 To Map(OldMap).Npc_HighIndex
            If MapNpc(OldMap).NPC(i).Num > 0 Then
                MapNpc(OldMap).NPC(i).Vital(Vitals.HP) = GetNpcMaxVital(MapNpc(OldMap).NPC(i).Num, Vitals.HP)
            End If
        Next
    End If

    ' Sets it so we know to process npcs on the map
    PlayersOnMap(MapNum) = YES
    tempPlayer(index).GettingMap = YES
    Set buffer = New clsBuffer
    Call SendCheckForMap(index, MapNum)
End Sub

Sub PlayerMove(ByVal index As Long, ByVal Dir As Long, ByVal Movement As Long, Optional ByVal SendToSelf As Boolean = False)
    Dim buffer As clsBuffer, MapNum As Integer
    Dim X As Long, Y As Long, i As Long
    Dim Moved As Byte, MovedSoFar As Boolean
    Dim TileType As Long, VitalType As Long, Color As Long, Amount As Long, BeginEventProcessing As Boolean

    ' Check for subscript out of range
    If IsPlaying(index) = False Or Dir < DIR_UP Or Dir > DIR_RIGHT Or Movement < 1 Or Movement > 2 Then Exit Sub

    Call SetPlayerDir(index, Dir)
    
    Moved = NO
    MapNum = GetPlayerMap(index)
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_UP + 1) Then
                    If Not isPlayerBlocked(index, 0, -1) Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_BLOCKED Then
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) - 1).Type <> TILE_TYPE_RESOURCE Then
                                Call SetPlayerY(index, GetPlayerY(index) - 1)
                                SendPlayerMove index, Movement, SendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Up > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Up, GetPlayerX(index), Map(MapNum).MaxY)
                    Moved = YES
                    
                    ' Clear their target
                    tempPlayer(index).Target = 0
                    tempPlayer(index).TargetType = TARGET_TYPE_NONE
                    SendPlayerTarget index
                End If
            End If

        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If GetPlayerY(index) < Map(MapNum).MaxY Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_DOWN + 1) Then
                    If Not isPlayerBlocked(index, 0, 1) Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_BLOCKED Then
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index) + 1).Type <> TILE_TYPE_RESOURCE Then
                                Call SetPlayerY(index, GetPlayerY(index) + 1)
                                SendPlayerMove index, Movement, SendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Down > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Down, GetPlayerX(index), 0)
                    Moved = YES
                    
                    ' Clear their target
                    tempPlayer(index).Target = 0
                    tempPlayer(index).TargetType = TARGET_TYPE_NONE
                    SendPlayerTarget index
                End If
            End If

        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) > 0 Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_LEFT + 1) Then
                    If Not isPlayerBlocked(index, -1, 0) Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) - 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                                Call SetPlayerX(index, GetPlayerX(index) - 1)
                                SendPlayerMove index, Movement, SendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Left > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Left, Map(MapNum).MaxX, GetPlayerY(index))
                    Moved = YES
                    
                    ' Clear their target
                    tempPlayer(index).Target = 0
                    tempPlayer(index).TargetType = TARGET_TYPE_NONE
                    SendPlayerTarget index
                End If
            End If

        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If GetPlayerX(index) < Map(MapNum).MaxX Then
                ' Check to make sure that the tile is walkable
                If Not IsDirBlocked(Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index)).DirBlock, DIR_RIGHT + 1) Then
                    If Not isPlayerBlocked(index, 1, 0) Then
                        If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_BLOCKED Then
                            If Map(GetPlayerMap(index)).Tile(GetPlayerX(index) + 1, GetPlayerY(index)).Type <> TILE_TYPE_RESOURCE Then
                                Call SetPlayerX(index, GetPlayerX(index) + 1)
                                SendPlayerMove index, Movement, SendToSelf
                                Moved = YES
                            End If
                        End If
                    End If
                End If
            Else
                ' Check to see if we can move them to the another map
                If Map(GetPlayerMap(index)).Right > 0 Then
                    Call PlayerWarp(index, Map(GetPlayerMap(index)).Right, 0, GetPlayerY(index))
                    Moved = YES
                    
                    ' Clear their target
                    tempPlayer(index).Target = 0
                    tempPlayer(index).TargetType = TARGET_TYPE_NONE
                    SendPlayerTarget index
                End If
            End If
    End Select
    
    With Map(GetPlayerMap(index)).Tile(GetPlayerX(index), GetPlayerY(index))
        ' Check to see if the tile is a warp tile, and if so warp them
        If .Type = TILE_TYPE_WARP Then
            MapNum = .Data1
            X = .Data2
            Y = .Data3
            Call PlayerWarp(index, MapNum, X, Y)
            Moved = YES
        End If
        
        ' Check for a shop, and if so open it
        If .Type = TILE_TYPE_SHOP Then
            X = .Data1
            If X > 0 Then ' Shop exists?
                If Len(Trim$(Shop(X).Name)) > 0 Then ' Name exists?
                    SendOpenShop index, X
                    tempPlayer(index).InShop = X ' Stops movement and the like
                End If
            End If
        End If
        
        ' Check to see if the tile is a bank, and if so send bank
        If .Type = TILE_TYPE_BANK Then
            SendBank index
            tempPlayer(index).InBank = True
            Moved = YES
        End If
        
        ' Check if it's a heal tile
        If .Type = TILE_TYPE_HEAL Then
            VitalType = .Data1
            Amount = .Data2
            
            If VitalType = Int(Vitals.HP) Then
                Color = BrightGreen
            ElseIf VitalType = Int(Vitals.MP) Then
                Color = BrightBlue
            End If
            
            If Not GetPlayerVital(index, VitalType) = GetPlayerMaxVital(index, VitalType) Then
                If GetPlayerVital(index, VitalType) + Amount > GetPlayerMaxVital(index, VitalType) Then
                    Amount = GetPlayerMaxVital(index, VitalType) - GetPlayerVital(index, VitalType)
                End If
                SendActionMsg GetPlayerMap(index), "+" & Amount, Color, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) + Amount
                Call SendVital(index, VitalType)
            Else
                SendActionMsg GetPlayerMap(index), "+0", Color, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                If tempPlayer(index).InParty > 0 Then SendPartyVitals tempPlayer(index).InParty, index
            End If
            Moved = YES
        End If
        
        ' Check if it's a trap tile
        If .Type = TILE_TYPE_TRAP Then
            VitalType = .Data1
            Amount = .Data2
            
            If VitalType = Int(Vitals.HP) Then
                Color = BrightRed
            ElseIf VitalType = Int(Vitals.MP) Then
                Color = Magenta
            End If
            
            If Not GetPlayerVital(index, VitalType) < 1 Then
                If GetPlayerVital(index, VitalType) - Amount < 1 Then
                    Amount = GetPlayerVital(index, VitalType)
                End If
                SendActionMsg GetPlayerMap(index), "-" & Amount, Color, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32, 1
                If GetPlayerVital(index, HP) - Amount < 1 And VitalType = 1 Then
                    KillPlayer index
                    Call GlobalMsg(GetPlayerName(index) & " has been killed by a trap!", BrightRed)
                Else
                    SetPlayerVital index, VitalType, GetPlayerVital(index, VitalType) - Amount
                    Call SendVital(index, VitalType)
                End If
            Else
                SetPlayerVital index, HP, GetPlayerVital(index, HP) - Amount
                PlayerMsg index, "You're injured by a trap.", BrightRed
                Call SendVital(index, HP)
                ' Send vitals to party if in one
                If tempPlayer(index).InParty > 0 Then SendPartyVitals tempPlayer(index).InParty, index
            End If
            Moved = YES
        End If
            
        ' Checkpoint
        If .Type = TILE_TYPE_CHECKPOINT Then
            SetCheckpoint index, .Data1, .Data2, .Data3
            Moved = YES
        End If
        
        ' Slide
        If .Type = TILE_TYPE_SLIDE Then
            ForcePlayerMove index, MOVING_WALKING, GetPlayerDir(index)
            Moved = YES
        End If
    End With

    ' They tried to hack
    If Moved = NO Then
        Call PlayerWarp(index, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
    Else
        If Trim$(Account(index).Chars(GetPlayerChar(index)).Status) = "AFK" Then
            Account(index).Chars(GetPlayerChar(index)).Status = vbNullString
            Call SendPlayerStatus(index)
        End If
        
        If tempPlayer(index).EventMap.CurrentEvents > 0 Then
            For i = 1 To tempPlayer(index).EventMap.CurrentEvents
                If Map(GetPlayerMap(index)).Events(tempPlayer(index).EventMap.EventPages(i).eventID).Global = 1 Then
                    If Map(GetPlayerMap(index)).Events(tempPlayer(index).EventMap.EventPages(i).eventID).X = X And Map(GetPlayerMap(index)).Events(tempPlayer(index).EventMap.EventPages(i).eventID).Y = Y And Map(GetPlayerMap(index)).Events(tempPlayer(index).EventMap.EventPages(i).eventID).Pages(tempPlayer(index).EventMap.EventPages(i).pageID).Trigger = 1 And tempPlayer(index).EventMap.EventPages(i).Visible = 1 Then BeginEventProcessing = True
                Else
                    If tempPlayer(index).EventMap.EventPages(i).X = X And tempPlayer(index).EventMap.EventPages(i).Y = Y And Map(GetPlayerMap(index)).Events(tempPlayer(index).EventMap.EventPages(i).eventID).Pages(tempPlayer(index).EventMap.EventPages(i).pageID).Trigger = 1 And tempPlayer(index).EventMap.EventPages(i).Visible = 1 Then BeginEventProcessing = True
                End If
                
                If BeginEventProcessing = True Then
                    'Process this event, it is on-touch and everything checks out.
                    If Map(GetPlayerMap(index)).Events(tempPlayer(index).EventMap.EventPages(i).eventID).Pages(tempPlayer(index).EventMap.EventPages(i).pageID).CommandListCount > 0 Then
                        tempPlayer(index).EventProcessingCount = tempPlayer(index).EventProcessingCount + 1
                        ReDim Preserve tempPlayer(index).EventProcessing(tempPlayer(index).EventProcessingCount)
                        tempPlayer(index).EventProcessing(tempPlayer(index).EventProcessingCount).ActionTimer = timeGetTime
                        tempPlayer(index).EventProcessing(tempPlayer(index).EventProcessingCount).CurList = 1
                        tempPlayer(index).EventProcessing(tempPlayer(index).EventProcessingCount).CurSlot = 1
                        tempPlayer(index).EventProcessing(tempPlayer(index).EventProcessingCount).eventID = tempPlayer(index).EventMap.EventPages(i).eventID
                        tempPlayer(index).EventProcessing(tempPlayer(index).EventProcessingCount).pageID = tempPlayer(index).EventMap.EventPages(i).pageID
                        tempPlayer(index).EventProcessing(tempPlayer(index).EventProcessingCount).WaitingForResponse = 0
                        ReDim tempPlayer(index).EventProcessing(tempPlayer(index).EventProcessingCount).ListLeftOff(0 To Map(GetPlayerMap(index)).Events(tempPlayer(index).EventMap.EventPages(i).eventID).Pages(tempPlayer(index).EventMap.EventPages(i).pageID).CommandListCount)
                    End If
                    BeginEventProcessing = False
                End If
            Next
        End If
    End If
End Sub

Sub ForcePlayerMove(ByVal index As Long, ByVal Movement As Long, ByVal Direction As Long)
    If Direction < DIR_UP Or Direction > DIR_RIGHT Then Exit Sub
    If Movement < 1 Or Movement > 2 Then Exit Sub
    
    Select Case Direction
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
        Case DIR_DOWN
            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
    End Select
    
    PlayerMove index, Direction, Movement, True
End Sub

Sub CheckEquippedItems(ByVal index As Long)
    Dim Slot As Long
    Dim ItemNum As Integer
    Dim i As Long

    ' We want to check incase an admin takes away an object but they had it equipped
    For i = 1 To Equipment.Equipment_Count - 1
        ItemNum = GetPlayerEquipment(index, i)

        If ItemNum > 0 Then
            If Not Item(ItemNum).Type = ITEM_TYPE_EQUIPMENT Or Not Item(ItemNum).EquipSlot = i Then SetPlayerEquipment index, 0, i
        Else
            SetPlayerEquipment index, 0, i
        End If
    Next

End Sub

Function FindOpenInvSlot(ByVal index As Long, ByVal ItemNum As Long) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then

        ' If currency then check to see if they already have an instance of the item and add it to that
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = ItemNum Then
                FindOpenInvSlot = i
                Exit Function
            End If
        Next

    End If

    For i = 1 To MAX_INV
        ' Try to find an open free slot
        If GetPlayerInvItemNum(index, i) = 0 Then
            FindOpenInvSlot = i
            Exit Function
        End If
    Next
End Function

Function FindOpenBankSlot(ByVal index As Long, ByVal ItemNum As Integer) As Byte
    Dim i As Long

    If Not IsPlaying(index) Then Exit Function
    
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Then Exit Function

    If Not Item(ItemNum).Type = ITEM_TYPE_EQUIPMENT Then
        For i = 1 To MAX_BANK
            If GetPlayerBankItemNum(index, i) = ItemNum Then
                FindOpenBankSlot = i
                Exit Function
            End If
        Next
    End If

    For i = 1 To MAX_BANK
        If GetPlayerBankItemNum(index, i) = 0 Then
            FindOpenBankSlot = i
            Exit Function
        End If
    Next
End Function

Function HasItem(ByVal index As Long, ByVal ItemNum As Integer) As Long
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                HasItem = GetPlayerInvItemValue(index, i)
            Else
                HasItem = 1
            End If
            
            Exit Function
        End If
    Next
End Function

Function TakeInvItem(ByVal index As Long, ByVal ItemNum As Integer, ByVal ItemVal As Long, Optional Update As Boolean = True) As Boolean
    Dim i As Long
    Dim n As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

    For i = 1 To MAX_INV
        ' Check to see if the player has the item
        If GetPlayerInvItemNum(index, i) = ItemNum Then
            If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
                ' Is what we are trying to take away more then what they have?  If so just set it to zero
                If ItemVal >= GetPlayerInvItemValue(index, i) Then
                    TakeInvItem = True
                Else
                    Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) - ItemVal)
                    
                    If Update Then
                        Call SendInventoryUpdate(index, i)
                    End If
                    
                    Exit Function
                End If
            Else
                TakeInvItem = True
            End If

            If TakeInvItem Then
                Call SetPlayerInvItemNum(index, i, 0)
                Call SetPlayerInvItemValue(index, i, 0)
                Call SetPlayerInvItemDur(index, i, 0)
                Call SetPlayerInvItemBind(index, i, 0)
                
                ' Send the inventory update
                If Update Then Call SendInventoryUpdate(index, i)
                Exit Function
            End If
        End If
    Next
End Function

Function TakeInvSlot(ByVal index As Long, ByVal InvSlot As Byte, ByVal ItemVal As Long, Optional ByVal Update As Boolean = True) As Boolean
    Dim i As Long
    Dim n As Long
    Dim ItemNum As Integer

    ' Check for subscript out of range
    If IsPlaying(index) = False Or InvSlot <= 0 Or InvSlot > MAX_ITEMS Then Exit Function
    
    ItemNum = GetPlayerInvItemNum(index, InvSlot)

    ' Prevent subscript out of range
    If ItemNum < 1 Then Exit Function
    
    If Item(ItemNum).Type = ITEM_TYPE_CURRENCY Then
        ' Is what we are trying to take away more then what they have?  If so just set it to zero
        If ItemVal >= GetPlayerInvItemValue(index, InvSlot) Then
            TakeInvSlot = True
        Else
            Call SetPlayerInvItemValue(index, InvSlot, GetPlayerInvItemValue(index, InvSlot) - ItemVal)
            
            ' Send the inventory update
            If Update Then
                Call SendInventoryUpdate(index, InvSlot)
            End If
            Exit Function
        End If
    Else
        TakeInvSlot = True
    End If

    If TakeInvSlot Then
        Call SetPlayerInvItemNum(index, InvSlot, 0)
        Call SetPlayerInvItemValue(index, InvSlot, 0)
        Call SetPlayerInvItemDur(index, InvSlot, 0)
        Call SetPlayerInvItemBind(index, InvSlot, 0)
        
        ' Send the inventory update
        If Update Then
            Call SendInventoryUpdate(index, InvSlot)
        End If
    End If
End Function

Function GiveInvItem(ByVal index As Long, ByVal ItemNum As Integer, ByVal ItemVal As Long, Optional ByVal ItemDur As Integer = -1, Optional ByVal ItemBind As Integer = 0, Optional ByVal SendUpdate As Boolean = True) As Boolean
    Dim i As Long

    ' Check for subscript out of range
    If IsPlaying(index) = False Or ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Function

    i = FindOpenInvSlot(index, ItemNum)

    ' Check to see if inventory is full
    If Not i = 0 Then
        Call SetPlayerInvItemNum(index, i, ItemNum)
        Call SetPlayerInvItemValue(index, i, GetPlayerInvItemValue(index, i) + ItemVal)
        
        If Item(GetPlayerInvItemNum(index, i)).Type = ITEM_TYPE_EQUIPMENT Then
            If ItemDur = -1 Then
                Call SetPlayerInvItemDur(index, i, Item(ItemNum).Data1)
            Else
                Call SetPlayerInvItemDur(index, i, ItemDur)
            End If
        End If
        
        If ItemBind = BIND_ON_PICKUP Or Item(GetPlayerInvItemNum(index, i)).BindType = BIND_ON_PICKUP Then
            Call SetPlayerInvItemBind(index, i, BIND_ON_PICKUP)
        Else
            Call SetPlayerInvItemBind(index, i, 0)
        End If
        
        If SendUpdate Then Call SendInventoryUpdate(index, i)
        GiveInvItem = True
    Else
        Call PlayerMsg(index, "Your inventory is full!", BrightRed)
    End If
End Function

Function HasSpell(ByVal index As Long, ByVal SpellNum As Long) As Boolean
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = SpellNum Then
            HasSpell = True
            Exit Function
        End If
    Next
End Function

Function FindOpenSpellSlot(ByVal index As Long) As Long
    Dim i As Long

    For i = 1 To MAX_PLAYER_SPELLS
        If GetPlayerSpell(index, i) = 0 Then
            FindOpenSpellSlot = i
            Exit Function
        End If
    Next
End Function

Sub PlayerMapGetItem(ByVal index As Long, ByVal i As Byte)
    Dim n As Long
    Dim MapNum As Integer
    Dim Msg As String

    If Not IsPlaying(index) Then Exit Sub
    MapNum = GetPlayerMap(index)

    ' See if there's even an item here
    If (MapItem(MapNum, i).Num > 0) And (MapItem(MapNum, i).Num <= MAX_ITEMS) Then
        ' Can we pick the item up?
        If CanPlayerPickupItem(index, i) Then
            ' Check if item is at the same location as the player
            If (MapItem(MapNum, i).X = GetPlayerX(index)) Then
                If (MapItem(MapNum, i).Y = GetPlayerY(index)) Then
                    ' Find open slot
                    n = FindOpenInvSlot(index, MapItem(MapNum, i).Num)

                    ' Open slot available?
                    If Not n = 0 Then
                        ' Set item in the player's inventory
                        Call SetPlayerInvItemNum(index, n, MapItem(MapNum, i).Num)

                        If Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_CURRENCY Then
                            Call SetPlayerInvItemValue(index, n, GetPlayerInvItemValue(index, n) + MapItem(MapNum, i).Value)
                            Msg = MapItem(MapNum, i).Value & " " & Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                        Else
                            Call SetPlayerInvItemValue(index, n, 0)
                            Msg = Trim$(Item(GetPlayerInvItemNum(index, n)).Name)
                        End If
                        
                        Call SetPlayerInvItemDur(index, n, MapItem(MapNum, i).Durability)
                        
                        If Item(GetPlayerInvItemNum(index, n)).BindType = BIND_ON_PICKUP Then
                            Call SetPlayerInvItemBind(index, i, BIND_ON_PICKUP)
                        Else
                            Call SetPlayerInvItemBind(index, i, 0)
                        End If
                        
                        ' Erase the item from the map
                        MapItem(MapNum, i).Num = 0
                        MapItem(MapNum, i).Value = 0
                        MapItem(MapNum, i).Durability = 0
                        MapItem(MapNum, i).X = 0
                        MapItem(MapNum, i).Y = 0
                        
                        Call SendInventoryUpdate(index, n)
                        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), 0, 0)
                        SendActionMsg GetPlayerMap(index), Msg, Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                    Else
                        Call PlayerMsg(index, "Your inventory is full.", BrightRed)
                    End If
                End If
            End If
        End If
    End If
End Sub

Function CanPlayerPickupItem(ByVal index As Long, ByVal MapItemNum As Integer)
    Dim MapNum As Integer

    MapNum = GetPlayerMap(index)
    
    If Moral(Map(MapNum).Moral).CanPickupItem = 1 Then
        ' No lock or locked to player?
        If Trim$(MapItem(MapNum, MapItemNum).playerName) = vbNullString Or Trim$(MapItem(MapNum, MapItemNum).playerName) = GetPlayerName(index) Then
            CanPlayerPickupItem = True
            Exit Function
        End If
    End If
End Function

Sub PlayerMapDropItem(ByVal index As Long, ByVal InvNum As Byte, ByVal Amount As Long)
    Dim i As Long
    Dim Msg As String
    
    If (GetPlayerInvItemNum(index, InvNum) <= MAX_ITEMS) Then
        i = FindOpenMapItemSlot(GetPlayerMap(index))

        If Not i = 0 Then
            MapItem(GetPlayerMap(index), i).Num = GetPlayerInvItemNum(index, InvNum)
            MapItem(GetPlayerMap(index), i).X = GetPlayerX(index)
            MapItem(GetPlayerMap(index), i).Y = GetPlayerY(index)
            MapItem(GetPlayerMap(index), i).playerName = Trim$(GetPlayerName(index))
            MapItem(GetPlayerMap(index), i).PlayerTimer = timeGetTime + ITEM_SPAWN_TIME
            MapItem(GetPlayerMap(index), i).CanDespawn = True
            MapItem(GetPlayerMap(index), i).DespawnTimer = timeGetTime + ITEM_DESPAWN_TIME

            If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_EQUIPMENT Then
                MapItem(GetPlayerMap(index), i).Durability = GetPlayerInvItemDur(index, InvNum)
            Else
                MapItem(GetPlayerMap(index), i).Durability = 0
            End If
            
            If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_CURRENCY Then
                ' Check if its more then they have and if so drop it all
                If Amount >= GetPlayerInvItemValue(index, InvNum) Then
                    MapItem(GetPlayerMap(index), i).Value = GetPlayerInvItemValue(index, InvNum)
                    Msg = GetPlayerInvItemValue(index, InvNum) & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name)
                    Call SetPlayerInvItemNum(index, InvNum, 0)
                    Call SetPlayerInvItemValue(index, InvNum, 0)
                    Call SetPlayerInvItemDur(index, InvNum, 0)
                    Call SetPlayerInvItemBind(index, InvNum, 0)
                Else
                    MapItem(GetPlayerMap(index), i).Value = Amount
                    Msg = Amount & " " & Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name)
                    Call SetPlayerInvItemValue(index, InvNum, GetPlayerInvItemValue(index, InvNum) - Amount)
                End If
            Else
                ' It's not a currency object so this is easy
                Msg = Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name)
                MapItem(GetPlayerMap(index), i).Value = 0
                Call SetPlayerInvItemNum(index, InvNum, 0)
                Call SetPlayerInvItemValue(index, InvNum, 0)
                Call SetPlayerInvItemDur(index, InvNum, 0)
                Call SetPlayerInvItemBind(index, InvNum, 0)
            End If
            
            ' Send message
            SendActionMsg GetPlayerMap(index), Msg, BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
            
            ' Send inventory update
            Call SendInventoryUpdate(index, InvNum)
            
            ' Spawn the item before we set the num or we'll get a different free map item slot
            Call SpawnItemSlot(i, MapItem(GetPlayerMap(index), i).Num, Amount, MapItem(GetPlayerMap(index), i).Durability, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
        Else
            Call PlayerMsg(index, "There are too many items on the ground to drop anything else.", BrightRed)
        End If
    End If
End Sub

Sub CheckPlayerLevelUp(ByVal index As Long)
    Dim i As Long
    Dim ExpRollOver As Long
    Dim Level_Count As Long

    If GetPlayerLevel(index) > 0 And GetPlayerLevel(index) < MAX_LEVEL Then
        Do While GetPlayerExp(index) >= GetPlayerNextLevel(index)
            ExpRollOver = GetPlayerExp(index) - GetPlayerNextLevel(index)
            Call SetPlayerLevel(index, GetPlayerLevel(index) + 1)
            Call SetPlayerPoints(index, GetPlayerPoints(index) + 5)
            Call SetPlayerExp(index, ExpRollOver)
            Level_Count = Level_Count + 1
        Loop
        
        If Level_Count > 0 Then
            SendActionMsg GetPlayerMap(index), "Level Up", Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
            SendPlayerExp index
            Call SendAnimation(GetPlayerMap(index), 2, 0, 0, TARGET_TYPE_PLAYER, index)
            
            If Level_Count > 1 Then
                Call GlobalMsg(GetPlayerName(index) & " has gained " & Level_Count & " levels!", Yellow)
            Else
                Call GlobalMsg(GetPlayerName(index) & " has gained a level!", Yellow)
            End If
            
            ' Restore and send vitals
            For i = 1 To Vitals.Vital_Count - 1
                Call SetPlayerVital(index, i, GetPlayerMaxVital(index, i))
                Call SendVital(index, i)
            Next
            
            ' Check for new title
            Call CheckPlayerNewTitle(index)
            
            ' Check if any of the player's spells can rank up
            For i = 1 To MAX_PLAYER_SPELLS
                If GetPlayerSpell(index, i) > 0 Then
                    If Spell(GetPlayerSpell(index, i)).NewSpell > 0 Then
                        If Spell(Spell(GetPlayerSpell(index, i)).NewSpell).CastRequired > 0 Then
                            Call CheckSpellRankUp(index, GetPlayerSpell(index, i), i)
                        End If
                    End If
                End If
            Next
            
            ' Send other data
            Call SendPlayerStats(index)
            Call SendPlayerPoints(index)
            Call SendPlayerLevel(index)
        End If
    End If
End Sub

Sub CheckPlayerSkillLevelUp(ByVal index As Long, ByVal SkillNum As Byte)
    Dim ExpRollOver As Long
    Dim Level_Count As Long
    
    Level_Count = 0

    If GetPlayerSkillLevel(index, SkillNum) > 0 And GetPlayerSkillLevel(index, SkillNum) < MAX_LEVEL Then
        Do While GetPlayerSkillExp(index, SkillNum) >= GetPlayerNextSkillLevel(index, SkillNum)
            ExpRollOver = GetPlayerSkillExp(index, SkillNum) - GetPlayerNextSkillLevel(index, SkillNum)
            Call SetPlayerSkillLevel(index, GetPlayerSkillLevel(index, SkillNum) + 1, SkillNum)
            Call SetPlayerSkillExp(index, ExpRollOver, SkillNum)
            Level_Count = Level_Count + 1
        Loop
        
        If Level_Count > 0 Then
            SendActionMsg GetPlayerMap(index), "Level Up", Yellow, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
            Call PlayerMsg(index, "Your " & GetSkillName(SkillNum) & " level is now " & GetPlayerSkillLevel(index, SkillNum) & ".", BrightGreen)
            Call SendAnimation(GetPlayerMap(index), 2, 0, 0, TARGET_TYPE_PLAYER, index)
        End If
    End If
End Sub

Private Function AutoLife(ByVal index As Long) As Boolean
    Dim i As Byte
    
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) > 0 Then
            If Item(GetPlayerInvItemNum(index, i)).Type = ITEM_TYPE_AUTOLIFE Then
                If CanPlayerUseItem(index, GetPlayerInvItemNum(index, i), False) Then
                    ' HP
                    If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddHP > 0 Then
                        If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddHP > GetPlayerMaxVital(index, HP) Then
                            SendActionMsg GetPlayerMap(index), "+" & GetPlayerMaxVital(index, HP), BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                        Else
                            SendActionMsg GetPlayerMap(index), "+" & Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                        End If
                        Call SetPlayerVital(index, Vitals.HP, GetPlayerVital(index, Vitals.HP) + Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddHP)
                        Call SendVital(index, Vitals.HP)
                    End If
                    
                    ' MP
                    If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddMP > 0 Then
                        If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddMP > GetPlayerMaxVital(index, MP) Then
                            SendActionMsg GetPlayerMap(index), "+" & GetPlayerMaxVital(index, MP), BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                        Else
                            SendActionMsg GetPlayerMap(index), "+" & Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                        End If
                        Call SendVital(index, Vitals.MP)
                        Call SetPlayerVital(index, Vitals.MP, GetPlayerVital(index, Vitals.MP) + Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).AddMP)
                    End If
                    
                    ' If it is not reusable then take the item away
                    If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).IsReusable = False Then
                        Call TakeInvItem(index, Account(index).Chars(GetPlayerChar(index)).Inv(i).Num, 0)
                    End If
                    
                    Call SendAnimation(GetPlayerMap(index), Item(GetPlayerInvItemNum(index, i)).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                    
                    ' Warp player away
                    If Item(Account(index).Chars(GetPlayerChar(index)).Inv(i).Num).Data1 = 1 Then
                        Call WarpPlayer(index)
                    End If
                    
                    Call PlayerMsg(index, "You have been given another life!", Yellow)
                    
                    AutoLife = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Sub OnDeath(ByVal index As Long, Optional ByVal Attacker As Long)
    Dim i As Long
   
    ' Set HP to 0
    Call SetPlayerVital(index, Vitals.HP, 0)
    
    ' Exit out if they were saved
    If AutoLife(index) Then Exit Sub
    
    ' If map moral can drop items or not
    If Moral(Map(GetPlayerMap(index)).Moral).DropItems = 1 Then
        If GetPlayerPK(index) = YES Then
            Call SetPlayerPK(index, NO)
            Call SendPlayerPK(index)
        End If

        ' Drop all worn items
        For i = 1 To Equipment.Equipment_Count - 1
            If GetPlayerEquipment(index, i) > 0 Then
                If tempPlayer(Attacker).InParty > 0 Then
                    Call Party_GetLoot(tempPlayer(Attacker).InParty, GetPlayerEquipment(index, i), 1, GetPlayerX(index), GetPlayerY(index))
                Else
                    If Moral(GetPlayerMap(index)).CanDropItem = 0 Then
                        Call SpawnItem(GetPlayerEquipment(index, i), 1, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(Attacker))
                    Else
                        Call GiveInvItem(Attacker, GetPlayerEquipment(index, i), 1)
                    End If
                End If
                
                ' Send a message to the world indicating that they dropped an item
                Call GlobalMsg(GetPlayerName(index) & " drops " & CheckGrammar(Item(GetPlayerEquipment(index, i)).Name) & "!", Yellow)
                
                ' Remove equipment item
                SetPlayerEquipment index, 0, i
                SetPlayerEquipmentDur index, 0, i
                SetPlayerEquipmentBind index, 0, i
            End If
        Next
        
        ' Drop 10% of their gold
        For i = 1 To MAX_INV
            If GetPlayerInvItemNum(index, i) = 1 Then
                If Round(GetPlayerInvItemValue(index, i) / 10) > 0 Then
                    Call TakeInvItem(index, GetPlayerInvItemNum(index, i), Round(GetPlayerInvItemValue(index, i) / 10))
                    Call SpawnItem(1, Round(GetPlayerInvItemValue(index, i) / 10), GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(Attacker))
                    Exit For
                End If
            End If
        Next
        
        ' Add the player kill
        If Attacker > 0 Then
            Account(FindPlayer(GetPlayerName(Attacker))).Chars(GetPlayerChar(i)).PlayerKills = Account(FindPlayer(GetPlayerName(Attacker))).Chars(GetPlayerChar(i)).PlayerKills + 1
        End If
        
        ' Check for new title
        Call CheckPlayerNewTitle(index)
    End If
    
    ' Loop through entire map and purge npc targets from player
    For i = 1 To Map(GetPlayerMap(index)).Npc_HighIndex
        If MapNpc(GetPlayerMap(index)).NPC(i).Num > 0 Then
            If MapNpc(GetPlayerMap(index)).NPC(i).TargetType = TARGET_TYPE_PLAYER Then
                If MapNpc(GetPlayerMap(index)).NPC(i).Target = index Then
                    MapNpc(GetPlayerMap(index)).NPC(i).Target = 0
                    MapNpc(GetPlayerMap(index)).NPC(i).TargetType = TARGET_TYPE_NONE
                    Call SendMapNpcTarget(GetPlayerMap(index), i, 0, 0)
                End If
            End If
        End If
    Next

    ' Set player direction
    Call SetPlayerDir(index, DIR_DOWN)
    
    ' Warp away player
    Call WarpPlayer(index)
    
    ' Clear all DoTs and HoTs
    For i = 1 To MAX_DOTS
        With tempPlayer(index).DoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
        
        With tempPlayer(index).HoT(i)
            .Used = False
            .Spell = 0
            .Timer = 0
            .Caster = 0
            .StartTime = 0
        End With
    Next
    
    ' Clear spell casting
    Call ClearAccountSpellBuffer(index)
    Call SendClearAccountSpellBuffer(index)
    
    ' Restore vitals
    Call SetPlayerVital(index, Vitals.HP, GetPlayerMaxVital(index, Vitals.HP))
    Call SetPlayerVital(index, Vitals.MP, GetPlayerMaxVital(index, Vitals.MP))

    ' Send vitals to party if in one
    If tempPlayer(index).InParty > 0 Then SendPartyVitals tempPlayer(index).InParty, index
    
    ' Send vitals
    For i = 1 To Vitals.Vital_Count - 1
        Call SendVital(index, i)
    Next
End Sub

Private Sub WarpPlayer(ByVal index As Long)
     With Map(GetPlayerMap(index))
        If .BootMap = 0 Then
            ' Warp to the checkpoint
            Call WarpToCheckPoint(index)
        Else
            ' Warp to the boot map
            If .BootMap > 0 And .BootMap <= MAX_MAPS Then
                PlayerWarp index, .BootMap, .BootX, .BootY
            Else
                ' Warp to the start map
                Call PlayerWarp(index, START_MAP, START_X, START_Y)
            End If
        End If
     End With
End Sub

Sub CheckResource(ByVal index As Long, ByVal X As Long, ByVal Y As Long)
    Dim Resource_Num As Long
    Dim Resource_Index As Long
    Dim rX As Long, rY As Long
    Dim i As Long
    Dim Damage As Long
    Dim RndNum As Long
    
    If Map(GetPlayerMap(index)).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        Resource_Num = 0
        Resource_Index = Map(GetPlayerMap(index)).Tile(X, Y).Data1

        ' Get the cache number
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            If ResourceCache(GetPlayerMap(index)).ResourceData(i).X = X Then
                If ResourceCache(GetPlayerMap(index)).ResourceData(i).Y = Y Then
                    Resource_Num = i
                End If
            End If
        Next

        If Resource_Num > 0 Then
            If GetPlayerEquipment(index, Weapon) > 0 Then
                If Item(GetPlayerEquipment(index, Weapon)).Data3 = Resource(Resource_Index).ToolRequired Then
                    If Not GetPlayerEquipmentDur(index, Weapon) = 0 Or Item(GetPlayerEquipment(index, Weapon)).Data1 = 0 Then
                        ' Enough space in inventory?
                        If Resource(Resource_Index).ItemReward > 0 Then
                            If FindOpenInvSlot(index, Resource(Resource_Index).ItemReward) = 0 Then
                                PlayerMsg index, "You do not have enough inventory space!", BrightRed
                                Exit Sub
                            End If
                        End If
    
                        ' Check if the resource has already been deplenished
                        If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).ResourceState = 0 Then
                            rX = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).X
                            rY = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).Y
                        
                            ' Reduce weapon's durability
                            Call DamagePlayerEquipment(index, Weapon)
                            
                            ' Give the reward random when they deal damage
                            RndNum = Random(Resource(Resource_Index).LowChance, Resource(Resource_Index).HighChance)
                              
                            If Not RndNum = Resource(Resource_Index).LowChance Then
                                ' Subtract the RndNum by the random value of the weapon's chance modifier
                                RndNum = RndNum - Round(Random((Item(GetPlayerEquipment(index, Weapon)).ChanceModifier / 2), Item(GetPlayerEquipment(index, Weapon)).ChanceModifier))
                                
                                ' If value is less than the resource low chance then set it to it
                                If RndNum < Resource(Resource_Index).LowChance Then
                                    RndNum = Resource(Resource_Index).LowChance
                                End If
                            End If
                            
                            If RndNum = Resource(Resource_Index).LowChance Then
                                ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).Cur_Reward = ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).Cur_Reward - 1
                                GiveInvItem index, Resource(Resource_Index).ItemReward, 1
                                
                                If GetPlayerSkillLevel(index, Resource(Resource_Index).Skill) < MAX_LEVEL Then
                                    ' Add the experience to the skill
                                    Call SetPlayerSkillExp(index, GetPlayerSkillExp(index, Resource(Resource_Index).Skill) + Resource(Resource_Index).Exp * EXP_RATE, Resource(Resource_Index).Skill)
                                    
                                    ' Check for skill level up
                                    Call CheckPlayerSkillLevelUp(index, Resource(Resource_Index).Skill)
                                End If
                                
                                ' Send message if it exists
                                If Len(Trim$(Resource(Resource_Index).SuccessMessage)) > 0 Then
                                    SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_Index).SuccessMessage), BrightGreen, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                End If
                                
                                ' If the resource is empty then clear it
                                If ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).Cur_Reward = 0 Then
                                    ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).ResourceState = 1
                                    ResourceCache(GetPlayerMap(index)).ResourceData(Resource_Num).ResourceTimer = timeGetTime
                                    SendResourceCacheToMap GetPlayerMap(index), Resource_Num
                                End If
                            Else
                                ' Send message if it exists
                                If Len(Trim$(Resource(Resource_Index).FailMessage)) > 0 Then
                                    SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_Index).FailMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                                End If
                            End If
                            
                            SendAnimation GetPlayerMap(index), Resource(Resource_Index).Animation, rX, rY
                            
                            ' Send the sound
                            SendMapSound GetPlayerMap(index), index, rX, rY, SoundEntity.seResource, Resource_Index
                        Else
                            ' Send message if it exists
                            If Len(Trim$(Resource(Resource_Index).EmptyMessage)) > 0 Then
                                SendActionMsg GetPlayerMap(index), Trim$(Resource(Resource_Index).EmptyMessage), BrightRed, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)
                            End If
                        End If
                    Else
                        PlayerMsg index, "The tool your using is broken!", BrightRed
                    End If
                Else
                    PlayerMsg index, "You have the wrong type of tool equipped.", BrightRed
                End If
            Else
                PlayerMsg index, "You need a tool to interact with this resource.", BrightRed
            End If
        End If
    End If
End Sub

Sub GiveBankItem(ByVal index As Long, ByVal InvSlot As Byte, ByVal Amount As Long, ByVal Durability As Integer)
    Dim BankSlot
    
    BankSlot = FindOpenBankSlot(index, GetPlayerInvItemNum(index, InvSlot))
        
    If BankSlot > 0 Then
        If Item(GetPlayerInvItemNum(index, InvSlot)).Type = ITEM_TYPE_CURRENCY Then
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, InvSlot) Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + Amount)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), Amount)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, InvSlot))
                Call SetPlayerBankItemValue(index, BankSlot, Amount)
                Call SetPlayerBankItemBind(index, BankSlot, GetPlayerInvItemBind(index, InvSlot))
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), Amount)
            End If
        Else
            If GetPlayerBankItemNum(index, BankSlot) = GetPlayerInvItemNum(index, InvSlot) And Not Item(GetPlayerInvItemNum(index, InvSlot)).Type = ITEM_TYPE_EQUIPMENT Then
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) + 1)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), 0)
            Else
                Call SetPlayerBankItemNum(index, BankSlot, GetPlayerInvItemNum(index, InvSlot))
                Call SetPlayerBankItemValue(index, BankSlot, 1)
                Call SetPlayerBankItemBind(index, BankSlot, GetPlayerInvItemBind(index, InvSlot))
                Call SetPlayerBankItemDur(index, BankSlot, Durability)
                Call TakeInvItem(index, GetPlayerInvItemNum(index, InvSlot), 0)
            End If
        End If
    End If
    
    ' Send update
    SaveAccount index
    SendBank index
End Sub

Sub TakeBankItem(ByVal index As Long, ByVal BankSlot As Byte, ByVal Amount As Long, Durability As Integer)
    Dim InvSlot

    If BankSlot < 1 Or BankSlot > MAX_BANK Then Exit Sub
    
    ' Hack prevention
    If Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
        If GetPlayerBankItemValue(index, BankSlot) < Amount Then Amount = GetPlayerBankItemValue(index, BankSlot)
        If Amount < 1 Then Exit Sub
    Else
        If Not Amount = 1 Then Exit Sub
    End If
    
    InvSlot = FindOpenInvSlot(index, GetPlayerBankItemNum(index, BankSlot))
        
    If InvSlot > 0 Then
        If Item(GetPlayerBankItemNum(index, BankSlot)).Type = ITEM_TYPE_CURRENCY Then
            Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), Amount)
            Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - Amount)
            
            If GetPlayerBankItemValue(index, BankSlot) <= 0 Then
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
                Call SetPlayerBankItemBind(index, BankSlot, 0)
            End If
        Else
            If GetPlayerBankItemValue(index, BankSlot) > 1 Then
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0)
                Call SetPlayerBankItemValue(index, BankSlot, GetPlayerBankItemValue(index, BankSlot) - 1)
            Else
                Call GiveInvItem(index, GetPlayerBankItemNum(index, BankSlot), 0, GetPlayerBankItemDur(index, BankSlot), GetPlayerBankItemBind(index, BankSlot))
                Call SetPlayerBankItemNum(index, BankSlot, 0)
                Call SetPlayerBankItemValue(index, BankSlot, 0)
                Call SetPlayerBankItemDur(index, BankSlot, 0)
                Call SetPlayerBankItemBind(index, BankSlot, 0)
            End If
        End If
    End If
    
    SaveAccount index
    SendBank index
End Sub

Public Sub KillPlayer(ByVal index As Long)
    Dim Exp As Long

    ' Calculate exp to give to attacker
    Exp = GetPlayerExp(index) \ 4
    
    ' Randomize
    Exp = Random(Exp * 0.95, Exp * 1.05)

    ' Make sure the exp we get isn't less than 0
    If Exp < 0 Then Exp = 0
    
    If Exp = 0 Or Moral(Map(GetPlayerMap(index)).Moral).LoseExp = 0 Then
        Call PlayerMsg(index, "You did not lose any experience.", Grey)
    ElseIf GetPlayerLevel(index) < MAX_LEVEL Then
        Call SetPlayerExp(index, GetPlayerExp(index) - Exp)
        SendPlayerExp index
        Call PlayerMsg(index, "You lost " & Exp & " experience.", Grey)
    End If
    
    Call OnDeath(index)
End Sub

Public Sub UseItem(ByVal index As Long, ByVal InvNum As Byte)
    Dim n As Long, i As Long, X As Long, Y As Long, TotalPoints As Integer, filename As String, EquipSlot As Byte
    
    ' Check subscript out of range
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    
    ' Check if they can use the item
    If Not CanPlayerUseItem(index, InvNum) Then Exit Sub
    
    n = Item(GetPlayerInvItemNum(index, InvNum)).Data2

    ' Set the bind
    If Item(GetPlayerInvItemNum(index, InvNum)).Type = ITEM_TYPE_EQUIPMENT Then
        If Item(GetPlayerInvItemNum(index, InvNum)).BindType = BIND_ON_EQUIP Then
            Call SetPlayerInvItemBind(index, InvNum, BIND_ON_EQUIP)
        End If
    End If
            
    ' Find out what kind of item it is
    Select Case Item(GetPlayerInvItemNum(index, InvNum)).Type
        Case ITEM_TYPE_EQUIPMENT
            EquipSlot = Item(GetPlayerInvItemNum(index, InvNum)).EquipSlot
            
            If EquipSlot >= 1 And EquipSlot <= Equipment.Equipment_Count - 1 Then
                Call PlayerUnequipItem(index, EquipSlot, False, False)
                
                PlayerMsg index, "You equip " & CheckGrammar(Trim$(Item(GetPlayerInvItemNum(index, InvNum)).Name)) & ".", BrightGreen
                SetPlayerEquipment index, GetPlayerInvItemNum(index, InvNum), EquipSlot
                SetPlayerEquipmentDur index, GetPlayerInvItemDur(index, InvNum), EquipSlot
                SetPlayerEquipmentBind index, GetPlayerInvItemBind(index, InvNum), EquipSlot
                TakeInvSlot index, InvNum, 0, True
                
                ' Send update
                SendInventoryUpdate index, InvNum
                SendPlayerEquipmentTo index
                SendPlayerEquipmentToMapBut index
                SendPlayerStats index
                
                ' Send vitals
                For i = 1 To Vitals.Vital_Count - 1
                    Call SendVital(index, i)
                Next
                
            ' Send vitals to party if in one
            If tempPlayer(index).InParty > 0 Then SendPartyVitals tempPlayer(index).InParty, index
                
                 ' Send the sound
                SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerInvItemNum(index, InvNum)
            End If
        
        Case ITEM_TYPE_CONSUME
            If GetPlayerLevel(index) = MAX_LEVEL And Item(GetPlayerInvItemNum(index, InvNum)).AddEXP > 0 Then
                Call PlayerMsg(index, "You can't use items which modify your experience when your at the max level!", BrightRed)
                Exit Sub
            End If
            
            ' Add HP
            If Item(GetPlayerInvItemNum(index, InvNum)).AddHP > 0 Then
                If Not GetPlayerVital(index, HP) = GetPlayerMaxVital(index, HP) Then
                    If tempPlayer(index).VitalPotionTimer(HP) > timeGetTime Then
                        Call PlayerMsg(index, "You must wait before you can use another potion that modifies your health!", BrightRed)
                        Exit Sub
                    Else
                        If Item(GetPlayerInvItemNum(index, InvNum)).HoT = 1 Then
                            tempPlayer(index).VitalCycle(HP) = Item(GetPlayerInvItemNum(index, InvNum)).Data1
                            tempPlayer(index).VitalPotion(HP) = GetPlayerInvItemNum(index, InvNum)
                            tempPlayer(index).VitalPotionTimer(HP) = timeGetTime + (Item(GetPlayerInvItemNum(index, InvNum)).Data1 * 1000)
                        Else
                            Account(index).Chars(GetPlayerChar(index)).Vital(Vitals.HP) = Account(index).Chars(GetPlayerChar(index)).Vital(Vitals.HP) + Item(GetPlayerInvItemNum(index, InvNum)).AddHP
                            SendActionMsg GetPlayerMap(index), "+" & Item(GetPlayerInvItemNum(index, InvNum)).AddHP, BrightGreen, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                            SendVital index, HP
                            tempPlayer(index).VitalPotionTimer(HP) = timeGetTime + PotionWaitTimer
                        End If
                    End If
                ElseIf Item(GetPlayerInvItemNum(index, InvNum)).AddMP < 1 Then
                    Call PlayerMsg(index, "Using this item will have no effect!", BrightRed)
                    Exit Sub
                End If
            End If
            
            ' Add MP
            If Item(GetPlayerInvItemNum(index, InvNum)).AddMP > 0 Then
                If Not GetPlayerVital(index, MP) = GetPlayerMaxVital(index, MP) Then
                    If tempPlayer(index).VitalPotionTimer(MP) > timeGetTime And Item(GetPlayerInvItemNum(index, InvNum)).AddHP < 1 Then
                        Call PlayerMsg(index, "You must wait before you can use another potion that modifies your mana!", BrightRed)
                        Exit Sub
                    Else
                        If Item(GetPlayerInvItemNum(index, InvNum)).HoT = 1 Then
                            tempPlayer(index).VitalCycle(MP) = Item(GetPlayerInvItemNum(index, InvNum)).Data1
                            tempPlayer(index).VitalPotion(MP) = GetPlayerInvItemNum(index, InvNum)
                            tempPlayer(index).VitalPotionTimer(MP) = timeGetTime + (Item(GetPlayerInvItemNum(index, InvNum)).Data1 * 1000)
                        Else
                            Account(index).Chars(GetPlayerChar(index)).Vital(Vitals.MP) = Account(index).Chars(GetPlayerChar(index)).Vital(Vitals.MP) + Item(GetPlayerInvItemNum(index, InvNum)).AddMP
                            SendActionMsg GetPlayerMap(index), "+" & Item(GetPlayerInvItemNum(index, InvNum)).AddMP, BrightBlue, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
                            SendVital index, MP
                            tempPlayer(index).VitalPotionTimer(MP) = timeGetTime + PotionWaitTimer
                        End If
                    End If
                ElseIf Item(GetPlayerInvItemNum(index, InvNum)).AddHP < 1 Then
                    Call PlayerMsg(index, "Using this item will have no effect!", BrightRed)
                    Exit Sub
                End If
            End If
            
            ' Add Exp
            If Item(GetPlayerInvItemNum(index, InvNum)).AddEXP > 0 Then
                SetPlayerExp index, GetPlayerExp(index) + Item(GetPlayerInvItemNum(index, InvNum)).AddEXP
                SendPlayerExp index
                CheckPlayerLevelUp index
                SendActionMsg GetPlayerMap(index), "+" & Item(GetPlayerInvItemNum(index, InvNum)).AddEXP & " Exp", White, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32
            End If
            
            Call SendAnimation(GetPlayerMap(index), Item(GetPlayerInvItemNum(index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
            
            ' Send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerInvItemNum(index, InvNum)
            
            ' Is it reusable, if not take the item away
            If Item(GetPlayerInvItemNum(index, InvNum)).IsReusable = False Then
                Call TakeInvItem(index, Account(index).Chars(GetPlayerChar(index)).Inv(InvNum).Num, 0)
            End If
        
        Case ITEM_TYPE_SPELL
            ' Get the spell number
            n = Item(GetPlayerInvItemNum(index, InvNum)).Data1

            If n > 0 Then
                i = FindOpenSpellSlot(index)

                ' Make sure they have an open spell slot
                If i > 0 Then
                    ' Make sure they don't already have the spell
                    If Not HasSpell(index, n) Then
                        ' Make sure it's a valid name and their is an icon
                        If Not Trim$(Spell(n).Name) = vbNullString And Not Spell(n).Icon = 0 Then
                            ' Send the sound
                            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerInvItemNum(index, InvNum)
                            Call SetPlayerSpell(index, i, n)
                            Call SendAnimation(GetPlayerMap(index), Item(GetPlayerInvItemNum(index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
                            Call TakeInvItem(index, GetPlayerInvItemNum(index, InvNum), 0)
                            Call PlayerMsg(index, "You have learned a new spell!", BrightGreen)
                            Call SendPlayerSpell(index, i)
                        Else
                            Call PlayerMsg(index, "This spell either does not have a name or icon, report this to a staff member.", BrightRed)
                            Exit Sub
                        End If
                    Else
                        Call PlayerMsg(index, "You have already learned this spell!", BrightRed)
                        Exit Sub
                    End If
                Else
                    Call PlayerMsg(index, "You have learned all that you can learn!", BrightRed)
                    Exit Sub
                End If
            Else
                Call PlayerMsg(index, "This item does not have a spell, please inform a staff member!", BrightRed)
                Exit Sub
            End If
        
        Case ITEM_TYPE_TELEPORT
            If Moral(Map(GetPlayerMap(index)).Moral).CanPK = 1 Then
                Call PlayerMsg(index, "You can't teleport while in a PvP area!", BrightRed)
                Exit Sub
            End If
            
            Call SendAnimation(GetPlayerMap(index), Item(GetPlayerInvItemNum(index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
            Call PlayerWarp(index, Item(GetPlayerInvItemNum(index, InvNum)).Data1, Item(GetPlayerInvItemNum(index, InvNum)).Data2, Item(GetPlayerInvItemNum(index, InvNum)).Data3)
            
            ' Send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerInvItemNum(index, InvNum)
            
            ' Is it reusable, if not take item away
            If Item(GetPlayerInvItemNum(index, InvNum)).IsReusable = False Then
                Call TakeInvItem(index, Account(index).Chars(GetPlayerChar(index)).Inv(InvNum).Num, 1)
            End If
            
        Case ITEM_TYPE_RESETSTATS
            TotalPoints = GetPlayerPoints(index)
            filename = App.path & "\data\classes.ini"
            
            For i = 1 To Stats.Stat_count - 1
                TotalPoints = TotalPoints + (GetPlayerStat(index, i) - (Val(GetVar(filename, "Class" & index, "" & i & ""))))
                Call SetPlayerStat(index, i, Val(GetVar(filename, "Class" & index, "" & i & "")))
            Next
            
            ' Send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerInvItemNum(index, InvNum)
            
            Call SendAnimation(GetPlayerMap(index), Item(GetPlayerInvItemNum(index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
            Call SetPlayerPoints(index, TotalPoints)
            Call SendPlayerStats(index)
            Call SendPlayerPoints(index)
            Call PlayerMsg(index, "Your stats have been reset!", Yellow)
            Call TakeInvItem(index, Account(index).Chars(GetPlayerChar(index)).Inv(InvNum).Num, 1)

        Case ITEM_TYPE_SPRITE
            Call SendAnimation(GetPlayerMap(index), Item(GetPlayerInvItemNum(index, InvNum)).Animation, 0, 0, TARGET_TYPE_PLAYER, index)
            Call SetPlayerSprite(index, Item(GetPlayerInvItemNum(index, InvNum)).Data1)
            Call SendPlayerSprite(index)
            
            ' Send the sound
            SendPlayerSound index, GetPlayerX(index), GetPlayerY(index), SoundEntity.seItem, GetPlayerInvItemNum(index, InvNum)
        
            ' Is it reusable, if not take item away
            If Item(GetPlayerInvItemNum(index, InvNum)).IsReusable = False Then
                Call TakeInvItem(index, Account(index).Chars(GetPlayerChar(index)).Inv(InvNum).Num, 1)
            End If
    End Select
End Sub

Public Sub SetCheckpoint(ByVal index As Long, ByVal MapNum As Integer, ByVal X As Long, ByVal Y As Long)
    ' Check if their checkpoint is already set here
    If Account(index).Chars(GetPlayerChar(index)).CheckPointMap = MapNum And Account(index).Chars(GetPlayerChar(index)).CheckPointX = X And Account(index).Chars(GetPlayerChar(index)).CheckPointY = Y Then
        Call PlayerMsg(index, "Your checkpoint is already saved here!", BrightRed)
        Exit Sub
    End If
   
    PlayerMsg index, "Your checkpoint has been saved.", BrightGreen
    
    ' Save the Checkpoint
    Account(index).Chars(GetPlayerChar(index)).CheckPointMap = MapNum
    Account(index).Chars(GetPlayerChar(index)).CheckPointX = X
    Account(index).Chars(GetPlayerChar(index)).CheckPointY = Y
End Sub

Private Sub UpdatePlayerInvItems(ByVal index As Long)
    Dim TmpItem As Long
    Dim i As Byte, X As Byte
    
    X = 0

    ' Make sure the inv items are not cached as a currency
    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(index, i) > 0 And GetPlayerInvItemNum(index, i) <= MAX_INV Then
            If Not Item(GetPlayerInvItemNum(index, i)).Type = ITEM_TYPE_CURRENCY Then
                If GetPlayerInvItemValue(index, i) > 1 Then
                    TmpItem = Account(index).Chars(GetPlayerChar(index)).Inv(i).Num
                    Call TakeInvItem(index, TmpItem, 1)
                    Call GiveInvItem(index, TmpItem, 1)
                End If
            End If
            
            If GetPlayerInvItemNum(index, i) > 0 And GetPlayerInvItemNum(index, i) <= MAX_INV Then
                If Item(GetPlayerInvItemNum(index, i)).Type = ITEM_TYPE_CURRENCY Then
                    If GetPlayerInvItemValue(index, i) = 0 Then
                        TmpItem = Account(index).Chars(GetPlayerChar(index)).Inv(i).Num
                        Call TakeInvItem(index, TmpItem, 1)
                        X = X + 1
                    End If
                End If
            End If
        End If
    Next
    
    If X > 0 Then
        Call GiveInvItem(index, TmpItem, X)
    End If
End Sub

Public Sub UpdateAllPlayerInvItems(ByVal ItemNum As Integer)
    Dim TmpItem As Long
    Dim n As Long, i As Byte, X As Byte
    
    X = 0

    For n = 1 To Player_HighIndex
        If IsPlaying(n) Then
            ' Make sure the inv items are not cached as a currency
            For i = 1 To MAX_INV
                If GetPlayerInvItemNum(n, i) > 0 And GetPlayerInvItemNum(n, i) <= MAX_INV Then
                    If GetPlayerInvItemNum(n, i) = ItemNum Then
                        TmpItem = Account(n).Chars(GetPlayerChar(i)).Inv(i).Num
                        If Not Item(GetPlayerInvItemNum(n, i)).Type = ITEM_TYPE_CURRENCY Then
                            If GetPlayerInvItemValue(n, i) > 1 Then
                                Call TakeInvItem(n, TmpItem, 1)
                                Call GiveInvItem(n, TmpItem, 1)
                            End If
                        End If
                        
                        If GetPlayerInvItemNum(n, i) > 0 And GetPlayerInvItemNum(n, i) <= MAX_INV Then
                            If Item(GetPlayerInvItemNum(n, i)).Type = ITEM_TYPE_CURRENCY Then
                                If GetPlayerInvItemValue(n, i) = 0 Then
                                    Call TakeInvItem(n, TmpItem, 1)
                                    X = X + 1
                                End If
                            End If
                        End If
                    End If
                End If
            Next
            
            If X > 0 Then
                Call GiveInvItem(n, TmpItem, X)
            End If
        End If
    Next
End Sub

Function CanPlayerTrade(ByVal index As Long, ByVal TradeTarget As Long) As Boolean
    Dim sX As Long, sY As Long, tX As Long, tY As Long
    
    CanPlayerTrade = False
    
    ' Can't trade with yourself
    If TradeTarget = index Then
        PlayerMsg index, "You can't trade with yourself.", BrightRed
        Exit Function
    End If
    
    ' Make sure they're on the same map
    If Not Account(TradeTarget).Chars(GetPlayerChar(TradeTarget)).Map = Account(index).Chars(GetPlayerChar(index)).Map Then Exit Function
    
    ' Make sure they are allowed to trade
    If Account(TradeTarget).Chars(GetPlayerChar(index)).CanTrade = False Then
        PlayerMsg index, Trim$(GetPlayerName(TradeTarget)) & " has their trading turned off.", BrightRed
        Exit Function
    End If

    ' Make sure they're stood next to each other
    tX = Account(TradeTarget).Chars(GetPlayerChar(TradeTarget)).X
    tY = Account(TradeTarget).Chars(GetPlayerChar(TradeTarget)).Y
    sX = Account(index).Chars(GetPlayerChar(index)).X
    sY = Account(index).Chars(GetPlayerChar(index)).Y
    
    ' Within range?
    If tX < sX - 1 Or tX > sX + 1 And tY < sY - 1 Or tY > sY + 1 Then
        PlayerMsg index, "You need to be standing next to someone to request or accept a trade.", BrightRed
        Exit Function
    End If
    
    CanPlayerTrade = True
End Function

Function CanPlayerUseItem(ByVal index As Long, ByVal InvItem As Integer, Optional Message As Boolean = True) As Boolean
    Dim LevelReq As Byte
    Dim AccessReq As Byte
    Dim ClassReq As Byte
    Dim GenderReq As Byte
    Dim i As Long

    ' Can't use items while in a map that doesn't allow it
    If Moral(Map(GetPlayerMap(index)).Moral).CanUseItem = 0 Then
        Call PlayerMsg(index, "You can't use items here!", BrightRed)
        Exit Function
    End If
    
    LevelReq = Item(InvItem).LevelReq

    ' Make sure they are the right level
    If LevelReq > GetPlayerLevel(index) Then
        If Message Then
            Call PlayerMsg(index, "You must be level " & LevelReq & " to use this item.", BrightRed)
        End If
        Exit Function
    End If
    
    AccessReq = Item(InvItem).AccessReq
    
    ' Make sure they have the right access
    If AccessReq > GetPlayerAccess(index) Then
        If Message Then
            Call PlayerMsg(index, "You must be a staff member to use this item.", BrightRed)
        End If
        Exit Function
    End If
    
    ClassReq = Item(InvItem).ClassReq
    
    ' Make sure the Classes req > 0
    If ClassReq > 0 Then ' 0 = no req
        If Not ClassReq = GetPlayerClass(index) Then
            If Message Then
                Call PlayerMsg(index, "Only " & CheckGrammar(Trim$(Class(ClassReq).Name)) & " can use this item.", BrightRed)
            End If
            Exit Function
        End If
    End If
    
    GenderReq = Item(InvItem).GenderReq
    
    ' Make sure the Gender req > 0
    If GenderReq > 0 Then ' 0 = no req
        If Not GenderReq - 1 = GetPlayerGender(index) Then
            If Message Then
                If GetPlayerGender(index) = 0 Then
                    Call PlayerMsg(index, "You need to be a female to use this item!", BrightRed)
                Else
                    Call PlayerMsg(index, "You need to be a male to use this item!", BrightRed)
                End If
            End If
            Exit Function
        End If
    End If
    
    ' Check if they have the stats required to use this item
    For i = 1 To Stats.Stat_count - 1
        If GetPlayerRawStat(index, i) < Item(InvItem).Stat_Req(i) Then
            If Message Then
                PlayerMsg index, "You do not meet the stat requirements to use this item.", BrightRed
            End If
            Exit Function
        End If
    Next
    
    ' Check if they have the proficiency required to use this item
    If Item(InvItem).ProficiencyReq > 0 Then
        If GetPlayerProficiency(index, Item(InvItem).ProficiencyReq) = 0 Then
            If Message Then
                PlayerMsg index, "You lack the proficiency to use this item!", BrightRed
            End If
            Exit Function
        End If
    End If
    
    CanPlayerUseItem = True
End Function

Public Sub DamagePlayerEquipment(ByVal index As Long, ByVal EquipmentSlot As Byte)
    Dim Slot As Long, RandomNum As Byte
    
    Slot = GetPlayerEquipment(index, EquipmentSlot)
    
    If Slot = 0 Then Exit Sub
    
    ' Make sure the item isn't indestructable
    If Item(Slot).Data1 = 0 Then Exit Sub
    
    ' Don't subtract past 0
    If GetPlayerEquipmentDur(index, Slot) = 0 Then Exit Sub
    
    RandomNum = Random(1, 7)
    
    ' 1 in 7 chance it will actually damage the equipment if it's not a shield type item
    If RandomNum = 1 Or EquipmentSlot = Shield Then
        If Item(Slot).Type = ITEM_TYPE_EQUIPMENT Then
        
            ' Take away 1 durability
            Call SetPlayerEquipmentDur(index, GetPlayerEquipmentDur(index, Slot) - 1, Slot)
            Call SendPlayerEquipmentTo(index)
                
            If GetPlayerEquipmentDur(index, Slot) < 1 Then
                Call PlayerMsg(index, "Your " & Trim$(Item(Slot).Name) & " has broken.", BrightRed)
            ElseIf GetPlayerEquipmentDur(index, Slot) = 10 Then
                Call PlayerMsg(index, "Your " & Trim$(Item(Slot).Name) & " is about to break!", BrightRed)
            End If
        End If
    End If
End Sub

Public Sub WarpToCheckPoint(index As Long)
    Dim MapNum As Integer
    Dim X As Long, Y As Long
    
    MapNum = Account(index).Chars(GetPlayerChar(index)).CheckPointMap
    X = Account(index).Chars(GetPlayerChar(index)).CheckPointX
    Y = Account(index).Chars(GetPlayerChar(index)).CheckPointY
    
    PlayerWarp index, MapNum, X, Y
End Sub

Function IsAFriend(ByVal index As Long, ByVal OtherPlayer As Long) As Boolean
    Dim i As Long
    
    ' Are they on the user's friend list
    For i = 1 To Account(OtherPlayer).Friends.AmountOfFriends
        If Trim$(Account(OtherPlayer).Friends.Members(i)) = GetPlayerName(index) Then
            IsAFriend = True
            Exit Function
        End If
    Next
End Function

Function IsAFoe(ByVal index As Long, ByVal OtherPlayer As Long) As Boolean
    Dim i As Long
    
    ' Are they on the user's foe list
    For i = 1 To Account(OtherPlayer).Foes.Amount
        If Trim$(Account(OtherPlayer).Foes.Members(i)) = GetPlayerName(index) Then
            Call PlayerMsg(index, "You are being ignored by " & GetPlayerName(OtherPlayer) & "!", BrightRed)
            IsAFoe = True
            Exit Function
        End If
    Next
End Function

Function IsPlayerBusy(ByVal index As Long, ByVal OtherPlayer As Long) As Boolean
    ' Make sure they're not busy doing something else
    If IsPlaying(OtherPlayer) Then
        If tempPlayer(OtherPlayer).InBank Or tempPlayer(OtherPlayer).InShop > 0 Or tempPlayer(OtherPlayer).InTrade > 0 Or tempPlayer(OtherPlayer).PartyInvite > 0 Or tempPlayer(OtherPlayer).TradeRequest > 0 Or tempPlayer(OtherPlayer).GuildInvite > 0 Then
            IsPlayerBusy = True
            PlayerMsg index, GetPlayerName(OtherPlayer) & " is busy!", BrightRed
            Exit Function
        End If
    End If
End Function

