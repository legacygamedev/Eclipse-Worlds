Attribute VB_Name = "modGameLogic"
Option Explicit

Function FindOpenPlayerSlot() As Long
    Dim i As Long
    FindOpenPlayerSlot = 0

    For i = 1 To MAX_PLAYERS
        If Not IsConnected(i) Then
            FindOpenPlayerSlot = i
            Exit Function
        End If
    Next
End Function

Function FindOpenMapItemSlot(ByVal MapNum As Integer) As Long
    Dim i As Long
    FindOpenMapItemSlot = 0

    ' Check for subscript out of range
    If MapNum < 1 Or MapNum > MAX_MAPS Then Exit Function

    For i = 1 To MAX_MAP_ITEMS
        If MapItem(MapNum, i).Num = 0 Then
            FindOpenMapItemSlot = i
            Exit Function
        End If
    Next
End Function

Function TotalOnlinePlayers() As Long
    Dim i As Long
    
    TotalOnlinePlayers = 0

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            TotalOnlinePlayers = TotalOnlinePlayers + 1
        End If
    Next
End Function

Function FindPlayer(ByVal Name As String) As Long
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            ' Make sure we dont try to check a name that's too small
            If Len(GetPlayerName(i)) >= Len(Trim$(Name)) Then
                If UCase$(Mid$(GetPlayerName(i), 1, Len(Trim$(Name)))) = UCase$(Trim$(Name)) Then
                    FindPlayer = i
                    Exit Function
                End If
            End If
        End If
    Next

    FindPlayer = 0
End Function

Sub SpawnItem(ByVal ItemNum As Integer, ByVal ItemVal As Long, ByVal ItemDur As Integer, ByVal MapNum As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString)
    Dim i As Long, II As Long

    ' Check for subscript out of range
    If ItemNum < 1 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    
    If Item(ItemNum).stackable = 1 Then
        ' Find open map item slot
        i = FindOpenMapItemSlot(MapNum)
        
        Call SpawnItemSlot(i, ItemNum, ItemVal, ItemDur, MapNum, X, Y, playerName)
    Else
        If Item(ItemNum).Type = ITEM_TYPE_EQUIPMENT Then
            If ItemVal = 0 Then ItemVal = 1
        End If
        
        For II = 1 To ItemVal
            ' Find open map item slot
            i = FindOpenMapItemSlot(MapNum)

            If i <= 0 Or i > MAX_MAP_ITEMS Then Exit For
            
            If Item(ItemNum).Type <> ITEM_TYPE_EQUIPMENT Then
                Call SpawnItemSlot(i, ItemNum, 1, ItemDur, MapNum, X, Y, playerName)
            Else
                Call SpawnItemSlot(i, ItemNum, 0, ItemDur, MapNum, X, Y, playerName)
            End If
        Next
    End If
End Sub

Sub SpawnItemSlot(ByVal MapItemSlot As Long, ByVal ItemNum As Integer, ByVal ItemVal As Long, ByVal ItemDur As Integer, ByVal MapNum As Integer, ByVal X As Long, ByVal Y As Long, Optional ByVal playerName As String = vbNullString, Optional ByVal CanDespawn As Boolean = True)
    Dim i As Long
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If MapItemSlot <= 0 Or MapItemSlot > MAX_MAP_ITEMS Or ItemNum < 0 Or ItemNum > MAX_ITEMS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub

    i = MapItemSlot

    If Not i = 0 Then
        If ItemNum >= 0 And ItemNum <= MAX_ITEMS Then
            MapItem(MapNum, i).playerName = playerName
            MapItem(MapNum, i).PlayerTimer = timeGetTime + ITEM_SPAWN_TIME
            MapItem(MapNum, i).CanDespawn = CanDespawn
            MapItem(MapNum, i).DespawnTimer = timeGetTime + ITEM_DESPAWN_TIME
            MapItem(MapNum, i).Num = ItemNum
            MapItem(MapNum, i).Value = ItemVal
            MapItem(MapNum, i).Durability = ItemDur
            
            Call SetMapItemX(MapNum, i, X)
            Call SetMapItemY(MapNum, i, Y)
            
            ' Send to map
            SendMapItemToMap MapNum, i
        End If
    End If

End Sub

Sub SpawnAllMapsItems()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapItems(i)
    Next
End Sub

Sub SpawnMapItems(ByVal MapNum As Integer)
    Dim X As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum < 1 Or MapNum > MAX_MAPS Then Exit Sub
    
    If Map(MapNum).MaxX = 0 Or Map(MapNum).MaxY = 0 Then Exit Sub
    ReDim Map(MAX_MAPS).Tile(Map(MapNum).MaxX, Map(MapNum).MaxY)
    
    ' Spawn what we have
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            ' Check if the tile type is an item or a saved tile incase someone drops something
            If (Map(MapNum).Tile(X, Y).Type = TILE_TYPE_ITEM) Then
                ' Check to see if its a currency and if they set the value to 0 set it to 1 automatically
                If Item(Map(MapNum).Tile(X, Y).Data1).stackable = 1 And Map(MapNum).Tile(X, Y).Data2 <= 0 Then
                    Call SpawnItem(Map(MapNum).Tile(X, Y).Data1, 1, Item(Map(MapNum).Tile(X, Y).Data1).Data1, MapNum, X, Y)
                Else
                    Call SpawnItem(Map(MapNum).Tile(X, Y).Data1, Map(MapNum).Tile(X, Y).Data2, Item(Map(MapNum).Tile(X, Y).Data1).Data1, MapNum, X, Y)
                End If
            End If
        Next
    Next
End Sub

Public Sub SpawnNPC(ByVal MapNPCNum As Long, ByVal MapNum As Integer, Optional ForcedSpawn As Boolean = False, Optional ByVal SetX As Integer, Optional ByVal SetY As Integer)
    Dim buffer As clsBuffer
    Dim NPCNum As Long
    Dim i As Long
    Dim X As Long
    Dim Y As Long
    Dim Spawned As Boolean

    ' Check for subscript out of range
    If MapNPCNum <= 0 Or MapNPCNum > MAX_MAP_NPCS Or MapNum <= 0 Or MapNum > MAX_MAPS Then Exit Sub
    
    NPCNum = Map(MapNum).NPC(MapNPCNum)
    If ForcedSpawn = False And Map(MapNum).NPCSpawnType(MapNPCNum) = 1 Then NPCNum = 0
    
    If NPCNum > 0 Then
        MapNPC(MapNum).NPC(MapNPCNum).Num = NPCNum
        MapNPC(MapNum).NPC(MapNPCNum).target = 0
        MapNPC(MapNum).NPC(MapNPCNum).targetType = TARGET_TYPE_NONE ' Clear
        Call SendMapNPCTarget(MapNum, MapNPCNum, 0, 0)
       
        MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.HP) = GetNPCMaxVital(NPCNum, Vitals.HP)
        MapNPC(MapNum).NPC(MapNPCNum).Vital(Vitals.MP) = GetNPCMaxVital(NPCNum, Vitals.MP)

        MapNPC(MapNum).NPC(MapNPCNum).Dir = Int(Rnd * 4)
        
        ' Check if theres a spawn tile for the specific npc
        For X = 0 To Map(MapNum).MaxX
            For Y = 0 To Map(MapNum).MaxY
                If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_NPCSPAWN Then
                    If Map(MapNum).Tile(X, Y).Data1 = MapNPCNum Then
                        MapNPC(MapNum).NPC(MapNPCNum).X = X
                        MapNPC(MapNum).NPC(MapNPCNum).Y = Y
                        MapNPC(MapNum).NPC(MapNPCNum).Dir = Map(MapNum).Tile(X, Y).Data2
                        Spawned = True
                        Exit For
                    End If
                End If
            Next Y
        Next X
       
        If Not Spawned Then
            ' Well try 100 times to randomly place the sprite
            For i = 1 To 100
                If SetX = 0 And SetY = 0 Then
                    X = Random(0, Map(MapNum).MaxX)
                    Y = Random(0, Map(MapNum).MaxY)
                Else
                    X = SetX
                    Y = SetY
                End If
   
                ' Check if the tile is walkable
                If NPCTileIsOpen(MapNum, X, Y) Then
                    MapNPC(MapNum).NPC(MapNPCNum).X = X
                    MapNPC(MapNum).NPC(MapNPCNum).Y = Y
                    Spawned = True
                    Exit For
                End If
            Next
        End If

        ' Didn't spawn, so now we'll just try to find a free tile
        If Not Spawned Then
            For X = 0 To Map(MapNum).MaxX
                For Y = 0 To Map(MapNum).MaxY
                    If NPCTileIsOpen(MapNum, X, Y) Then
                        MapNPC(MapNum).NPC(MapNPCNum).X = X
                        MapNPC(MapNum).NPC(MapNPCNum).Y = Y
                        Spawned = True
                    End If
                Next
            Next
        End If

        ' If we suceeded in spawning then send it to everyone
        If Spawned Then
            Set buffer = New clsBuffer
            buffer.WriteLong SSpawnNPC
            buffer.WriteLong MapNPCNum
            buffer.WriteLong MapNPC(MapNum).NPC(MapNPCNum).Num
            buffer.WriteLong MapNPC(MapNum).NPC(MapNPCNum).X
            buffer.WriteLong MapNPC(MapNum).NPC(MapNPCNum).Y
            buffer.WriteLong MapNPC(MapNum).NPC(MapNPCNum).Dir
            SendDataToMap MapNum, buffer.ToArray()
            Set buffer = Nothing
            UpdateMapBlock MapNum, MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y, True
        End If
        SendMapNPCVitals MapNum, MapNPCNum
    End If
End Sub

Public Function NPCTileIsOpen(ByVal MapNum As Integer, ByVal X As Long, ByVal Y As Long) As Boolean
    Dim LoopI As Long
    NPCTileIsOpen = True

    If PlayersOnMap(MapNum) Then

        For LoopI = 1 To Player_HighIndex
            If GetPlayerMap(LoopI) = MapNum Then
                If GetPlayerX(LoopI) = X Then
                    If GetPlayerY(LoopI) = Y Then
                        NPCTileIsOpen = False
                        Exit Function
                    End If
                End If
            End If
        Next
    End If

    For LoopI = 1 To Map(MapNum).NPC_HighIndex
        If MapNPC(MapNum).NPC(LoopI).Num > 0 Then
            If MapNPC(MapNum).NPC(LoopI).X = X Then
                If MapNPC(MapNum).NPC(LoopI).Y = Y Then
                    NPCTileIsOpen = False
                    Exit Function
                End If
            End If
        End If
    Next

    If Map(MapNum).Tile(X, Y).Type <> TILE_TYPE_WALKABLE Then
        If Map(MapNum).Tile(X, Y).Type <> TILE_TYPE_NPCSPAWN Then
            If Map(MapNum).Tile(X, Y).Type <> TILE_TYPE_ITEM Then
                NPCTileIsOpen = False
            End If
        End If
    End If
End Function

Sub SpawnMapNPCs(ByVal MapNum As Integer)
    Dim i As Long

    For i = 1 To MAX_MAP_NPCS
        Call SpawnNPC(i, MapNum)
    Next
    
    CacheMapBlocks MapNum
End Sub

Sub SpawnAllMapNPCS()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnMapNPCs(i)
    Next
End Sub

Function CanNPCMove(ByVal MapNum As Integer, ByVal MapNPCNum As Long, ByVal Dir As Byte) As Boolean
    Dim i As Long
    Dim n As Long
    Dim X As Long
    Dim Y As Long

    ' Check for subscript out of range
    If MapNum < 1 Or MapNum > MAX_MAPS Or MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_DOWNRIGHT Or MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Spell > 0 Then Exit Function

    X = MapNPC(MapNum).NPC(MapNPCNum).X
    Y = MapNPC(MapNum).NPC(MapNPCNum).Y
    
    CanNPCMove = True

    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If Y > 0 Then
                n = Map(MapNum).Tile(X, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNPC(MapNum).NPC(MapNPCNum).X) And (GetPlayerY(i) = MapNPC(MapNum).NPC(MapNPCNum).Y - 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To Map(MapNum).NPC_HighIndex
                    If (i <> MapNPCNum) And (MapNPC(MapNum).NPC(i).Num > 0) And (MapNPC(MapNum).NPC(i).X = MapNPC(MapNum).NPC(MapNPCNum).X) And (MapNPC(MapNum).NPC(i).Y = MapNPC(MapNum).NPC(MapNPCNum).Y - 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y).DirBlock, DIR_UP + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Event blocking
                If IsEventBlocked(MapNPCNum, 0, -1, MapNum) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If

        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(X, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNPC(MapNum).NPC(MapNPCNum).X) And (GetPlayerY(i) = MapNPC(MapNum).NPC(MapNPCNum).Y + 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To Map(MapNum).NPC_HighIndex
                    If (i <> MapNPCNum) And (MapNPC(MapNum).NPC(i).Num > 0) And (MapNPC(MapNum).NPC(i).X = MapNPC(MapNum).NPC(MapNPCNum).X) And (MapNPC(MapNum).NPC(i).Y = MapNPC(MapNum).NPC(MapNPCNum).Y + 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y).DirBlock, DIR_DOWN + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Event blocking
                If IsEventBlocked(MapNPCNum, 0, 1, MapNum) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If

        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If X > 0 Then
                n = Map(MapNum).Tile(X - 1, Y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNPC(MapNum).NPC(MapNPCNum).X - 1) And (GetPlayerY(i) = MapNPC(MapNum).NPC(MapNPCNum).Y) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To Map(MapNum).NPC_HighIndex
                    If (i <> MapNPCNum) And (MapNPC(MapNum).NPC(i).Num > 0) And (MapNPC(MapNum).NPC(i).X = MapNPC(MapNum).NPC(MapNPCNum).X - 1) And (MapNPC(MapNum).NPC(i).Y = MapNPC(MapNum).NPC(MapNPCNum).Y) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y).DirBlock, DIR_LEFT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Event blocking
                If IsEventBlocked(MapNPCNum, -1, 0, MapNum) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If

        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If X < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(X + 1, Y).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNPC(MapNum).NPC(MapNPCNum).X + 1) And (GetPlayerY(i) = MapNPC(MapNum).NPC(MapNPCNum).Y) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To Map(MapNum).NPC_HighIndex
                    If (i <> MapNPCNum) And (MapNPC(MapNum).NPC(i).Num > 0) And (MapNPC(MapNum).NPC(i).X = MapNPC(MapNum).NPC(MapNPCNum).X + 1) And (MapNPC(MapNum).NPC(i).Y = MapNPC(MapNum).NPC(MapNPCNum).Y) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y).DirBlock, DIR_RIGHT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Event blocking
                If IsEventBlocked(MapNPCNum, 1, 0, MapNum) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
            
        Case DIR_UPLEFT
            ' Check to make sure not outside of boundries
            If X > 0 And Y > 0 Then
                n = Map(MapNum).Tile(X - 1, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNPC(MapNum).NPC(MapNPCNum).X - 1) And (GetPlayerY(i) = MapNPC(MapNum).NPC(MapNPCNum).Y - 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To Map(MapNum).NPC_HighIndex
                    If (i <> MapNPCNum) And (MapNPC(MapNum).NPC(i).Num > 0) And (MapNPC(MapNum).NPC(i).X = MapNPC(MapNum).NPC(MapNPCNum).X - 1) And (MapNPC(MapNum).NPC(i).Y = MapNPC(MapNum).NPC(MapNPCNum).Y - 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(MapNPC(MapNum).NPC(MapNPCNum).X - 1, MapNPC(MapNum).NPC(MapNPCNum).Y - 1).DirBlock, DIR_UPLEFT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Event blocking
                If IsEventBlocked(MapNPCNum, -1, -1, MapNum) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
            
        Case DIR_UPRIGHT
            ' Check to make sure not outside of boundries
            If X < Map(MapNum).MaxX And Y > 0 Then
                n = Map(MapNum).Tile(X + 1, Y - 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNPC(MapNum).NPC(MapNPCNum).X + 1) And (GetPlayerY(i) = MapNPC(MapNum).NPC(MapNPCNum).Y - 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To Map(MapNum).NPC_HighIndex
                    If (i <> MapNPCNum) And (MapNPC(MapNum).NPC(i).Num > 0) And (MapNPC(MapNum).NPC(i).X = MapNPC(MapNum).NPC(MapNPCNum).X + 1) And (MapNPC(MapNum).NPC(i).Y = MapNPC(MapNum).NPC(MapNPCNum).Y - 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(MapNPC(MapNum).NPC(MapNPCNum).X + 1, MapNPC(MapNum).NPC(MapNPCNum).Y - 1).DirBlock, DIR_UPRIGHT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Event blocking
                If IsEventBlocked(MapNPCNum, 1, -1, MapNum) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
            
        Case DIR_DOWNLEFT
            ' Check to make sure not outside of boundries
            If X > 0 And Y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(X - 1, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNPC(MapNum).NPC(MapNPCNum).X - 1) And (GetPlayerY(i) = MapNPC(MapNum).NPC(MapNPCNum).Y + 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To Map(MapNum).NPC_HighIndex
                    If (i <> MapNPCNum) And (MapNPC(MapNum).NPC(i).Num > 0) And (MapNPC(MapNum).NPC(i).X = MapNPC(MapNum).NPC(MapNPCNum).X - 1) And (MapNPC(MapNum).NPC(i).Y = MapNPC(MapNum).NPC(MapNPCNum).Y + 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(MapNPC(MapNum).NPC(MapNPCNum).X - 1, MapNPC(MapNum).NPC(MapNPCNum).Y + 1).DirBlock, DIR_DOWNLEFT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Event blocking
                If IsEventBlocked(MapNPCNum, -1, 1, MapNum) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
        
        Case DIR_DOWNRIGHT
            ' Check to make sure not outside of boundries
            If X < Map(MapNum).MaxX And Y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(X + 1, Y + 1).Type

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_ITEM And n <> TILE_TYPE_NPCSPAWN Then
                    CanNPCMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = MapNPC(MapNum).NPC(MapNPCNum).X + 1) And (GetPlayerY(i) = MapNPC(MapNum).NPC(MapNPCNum).Y + 1) Then
                            CanNPCMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To Map(MapNum).NPC_HighIndex
                    If (i <> MapNPCNum) And (MapNPC(MapNum).NPC(i).Num > 0) And (MapNPC(MapNum).NPC(i).X = MapNPC(MapNum).NPC(MapNPCNum).X + 1) And (MapNPC(MapNum).NPC(i).Y = MapNPC(MapNum).NPC(MapNPCNum).Y + 1) Then
                        CanNPCMove = False
                        Exit Function
                    End If
                Next
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(MapNPC(MapNum).NPC(MapNPCNum).X + 1, MapNPC(MapNum).NPC(MapNPCNum).Y + 1).DirBlock, DIR_DOWNRIGHT + 1) Then
                    CanNPCMove = False
                    Exit Function
                End If
                
                ' Event blocking
                If IsEventBlocked(MapNPCNum, 1, 1, MapNum) Then
                    CanNPCMove = False
                    Exit Function
                End If
            Else
                CanNPCMove = False
            End If
    End Select
End Function

Sub NPCMove(ByVal MapNum As Integer, ByVal MapNPCNum As Long, ByVal Dir As Byte, ByVal movement As Byte)
    ' Check for subscript out of range
    If MapNum < 1 Or MapNum > MAX_MAPS Or MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_DOWNRIGHT Or movement < 1 Or movement > 3 Then Exit Sub

    MapNPC(MapNum).NPC(MapNPCNum).Dir = Dir
    UpdateMapBlock MapNum, MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y, False
    
    Select Case Dir
        Case DIR_UP
            MapNPC(MapNum).NPC(MapNPCNum).Y = MapNPC(MapNum).NPC(MapNPCNum).Y - 1
        Case DIR_DOWN
            MapNPC(MapNum).NPC(MapNPCNum).Y = MapNPC(MapNum).NPC(MapNPCNum).Y + 1
        Case DIR_LEFT
            MapNPC(MapNum).NPC(MapNPCNum).X = MapNPC(MapNum).NPC(MapNPCNum).X - 1
        Case DIR_RIGHT
            MapNPC(MapNum).NPC(MapNPCNum).X = MapNPC(MapNum).NPC(MapNPCNum).X + 1
        Case DIR_UPLEFT
            MapNPC(MapNum).NPC(MapNPCNum).Y = MapNPC(MapNum).NPC(MapNPCNum).Y - 1
            MapNPC(MapNum).NPC(MapNPCNum).X = MapNPC(MapNum).NPC(MapNPCNum).X - 1
        Case DIR_UPRIGHT
            MapNPC(MapNum).NPC(MapNPCNum).Y = MapNPC(MapNum).NPC(MapNPCNum).Y - 1
            MapNPC(MapNum).NPC(MapNPCNum).X = MapNPC(MapNum).NPC(MapNPCNum).X + 1
        Case DIR_DOWNLEFT
            MapNPC(MapNum).NPC(MapNPCNum).Y = MapNPC(MapNum).NPC(MapNPCNum).Y + 1
            MapNPC(MapNum).NPC(MapNPCNum).X = MapNPC(MapNum).NPC(MapNPCNum).X - 1
        Case DIR_DOWNRIGHT
            MapNPC(MapNum).NPC(MapNPCNum).Y = MapNPC(MapNum).NPC(MapNPCNum).Y + 1
            MapNPC(MapNum).NPC(MapNPCNum).X = MapNPC(MapNum).NPC(MapNPCNum).X + 1
    End Select
    
    Call SendNPCMove(MapNPCNum, movement, MapNum)
    UpdateMapBlock MapNum, MapNPC(MapNum).NPC(MapNPCNum).X, MapNPC(MapNum).NPC(MapNPCNum).Y, True
End Sub

Sub NPCDir(ByVal MapNum As Integer, ByVal MapNPCNum As Long, ByVal Dir As Byte)
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum < 1 Or MapNum > MAX_MAPS Or MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Or Dir < DIR_UP Or Dir > DIR_DOWNRIGHT Then Exit Sub

    MapNPC(MapNum).NPC(MapNPCNum).Dir = Dir
    Set buffer = New clsBuffer
    buffer.WriteLong SNPCDir
    buffer.WriteLong MapNPCNum
    buffer.WriteByte Dir
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Function GetTotalMapPlayers(ByVal MapNum As Integer) As Long
    Dim i As Long
    Dim n As Long
    
    n = 0

    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            n = n + 1
        End If
    Next

    GetTotalMapPlayers = n
End Function

Public Sub CacheResources(ByVal MapNum As Integer)
    Dim X As Long, Y As Long, Resource_Count As Long

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            If Map(MapNum).Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve ResourceCache(MapNum).ResourceData(0 To Resource_Count)
                ResourceCache(MapNum).ResourceData(Resource_Count).X = X
                ResourceCache(MapNum).ResourceData(Resource_Count).Y = Y
                ResourceCache(MapNum).ResourceData(Resource_Count).Cur_Reward = Random(Resource(Map(MapNum).Tile(X, Y).Data1).Reward_Min, Resource(Map(MapNum).Tile(X, Y).Data1).Reward_Max)
            End If
        Next
    Next

    ResourceCache(MapNum).Resource_Count = Resource_Count
End Sub

Sub PlayerSwapBankSlots(ByVal Index As Long, ByVal OldSlot As Byte, ByVal NewSlot As Byte)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim NewNum As Long
    Dim NewValue As Long
    Dim OldBind As Byte
    Dim NewBind As Byte
    Dim OldDur As Integer
    Dim NewDur As Integer
    
    OldNum = GetPlayerBankItemNum(Index, OldSlot)
    OldValue = GetPlayerBankItemValue(Index, OldSlot)
    NewNum = GetPlayerBankItemNum(Index, NewSlot)
    NewValue = GetPlayerBankItemValue(Index, NewSlot)
    
    SetPlayerBankItemNum Index, NewSlot, OldNum
    SetPlayerBankItemValue Index, NewSlot, OldValue
    
    SetPlayerBankItemNum Index, OldSlot, NewNum
    SetPlayerBankItemValue Index, OldSlot, NewValue
    
    SetPlayerBankItemBind Index, OldSlot, NewBind
    SetPlayerBankItemBind Index, NewSlot, OldBind
    
    SetPlayerBankItemDur Index, OldSlot, NewDur
    SetPlayerBankItemDur Index, NewSlot, OldDur
        
    SendBank Index
End Sub

Sub PlayerSwitchSpellSlots(ByVal Index As Long, ByVal OldSlot As Byte, ByVal NewSlot As Byte)
    Dim OldNum As Long, NewNum As Long
    Dim OldSpellCD As Long, NewSpellCD As Long
    Dim OldSpellCasts As Integer, NewSpellCasts As Integer
    
    ' Switch the actual spells
    OldNum = GetPlayerSpell(Index, OldSlot)
    NewNum = GetPlayerSpell(Index, NewSlot)
    SetPlayerSpell Index, OldSlot, NewNum
    SetPlayerSpell Index, NewSlot, OldNum
    
    ' Switch the spell cooldowns
    OldSpellCD = GetPlayerSpellCD(Index, OldSlot)
    NewSpellCD = GetPlayerSpellCD(Index, NewSlot)
    Call SetPlayerSpellCD(Index, OldSlot, NewSpellCD)
    Call SetPlayerSpellCD(Index, NewSlot, OldSpellCD)
    
    ' Switch the spell casts
    OldSpellCasts = Account(Index).Chars(GetPlayerChar(Index)).AmountOfCasts(OldSlot)
    NewSpellCasts = Account(Index).Chars(GetPlayerChar(Index)).AmountOfCasts(NewSlot)
    Account(Index).Chars(GetPlayerChar(Index)).AmountOfCasts(OldSlot) = Account(Index).Chars(GetPlayerChar(Index)).AmountOfCasts(NewSlot)
    Account(Index).Chars(GetPlayerChar(Index)).AmountOfCasts(NewSlot) = Account(Index).Chars(GetPlayerChar(Index)).AmountOfCasts(OldSlot)
    
    ' Update the spells
    Call SendPlayerSpell(Index, OldSlot)
    Call SendPlayerSpell(Index, NewSlot)
    Call SendSpellCooldown(Index, OldSlot)
    Call SendSpellCooldown(Index, NewSlot)
End Sub

Sub PlayerSwitchInvSlots(ByVal Index As Long, ByVal OldSlot As Byte, ByVal NewSlot As Byte)
    Dim OldNum As Long
    Dim OldValue As Long
    Dim OldDur As Integer
    Dim OldBind As Byte
    Dim NewNum As Long
    Dim NewValue As Long
    Dim NewDur As Integer
    Dim NewBind As Byte

    ' Set the item
    OldNum = GetPlayerInvItemNum(Index, OldSlot)
    NewNum = GetPlayerInvItemNum(Index, NewSlot)
    SetPlayerInvItemNum Index, OldSlot, NewNum
    SetPlayerInvItemNum Index, NewSlot, OldNum
    
    ' Set the item's value
    OldValue = GetPlayerInvItemValue(Index, OldSlot)
    NewValue = GetPlayerInvItemValue(Index, NewSlot)
    SetPlayerInvItemValue Index, OldSlot, NewValue
    SetPlayerInvItemValue Index, NewSlot, OldValue
    
    ' Set the item's durability
    OldDur = GetPlayerInvItemDur(Index, OldSlot)
    NewDur = GetPlayerInvItemDur(Index, NewSlot)
    SetPlayerInvItemDur Index, OldSlot, NewDur
    SetPlayerInvItemDur Index, NewSlot, OldDur
    
    ' Set the item's bind
    OldBind = GetPlayerInvItemBind(Index, OldSlot)
    NewBind = GetPlayerInvItemBind(Index, NewSlot)
    SetPlayerInvItemBind Index, OldSlot, NewBind
    SetPlayerInvItemBind Index, NewSlot, OldBind
    
    SendInventory Index
End Sub

Sub PlayerSwitchHotbarSlots(ByVal Index As Long, ByVal OldSlot As Byte, ByVal NewSlot As Byte)
    Dim OldNum As Long
    Dim NewNum As Long
    Dim OldSType As Long
    Dim NewSType As Long

    ' Set the number
    OldNum = Account(Index).Chars(GetPlayerChar(Index)).Hotbar(OldSlot).Slot
    NewNum = Account(Index).Chars(GetPlayerChar(Index)).Hotbar(NewSlot).Slot
    Account(Index).Chars(GetPlayerChar(Index)).Hotbar(OldSlot).Slot = NewNum
    Account(Index).Chars(GetPlayerChar(Index)).Hotbar(NewSlot).Slot = OldNum
    
    ' Set the type
    OldSType = Account(Index).Chars(GetPlayerChar(Index)).Hotbar(OldSlot).SType
    NewSType = Account(Index).Chars(GetPlayerChar(Index)).Hotbar(NewSlot).SType
    Account(Index).Chars(GetPlayerChar(Index)).Hotbar(OldSlot).SType = NewSType
    Account(Index).Chars(GetPlayerChar(Index)).Hotbar(NewSlot).SType = OldSType
    
    SendHotbar Index
End Sub

Sub PlayerUnequipItem(ByVal Index As Long, ByVal EqSlot As Long, Optional ByVal SendUpdate As Boolean = True, Optional ByVal SendSound As Boolean = True, Optional ByVal Ignore As Boolean = False)
    Dim i As Long
    
    ' Check for subscript out of range
    If EqSlot < 1 Or EqSlot > Equipment.Equipment_Count - 1 Then Exit Sub
    
    If GetPlayerEquipment(Index, EqSlot) < 1 Then Exit Sub
    
    ' Don't allow a player killer to unequip their equipment
    'If GetPlayerPK(Index) = PLAYER_KILLER Then
    '    Call PlayerMsg(Index, "You can't unequip your items as a Player Killer!", BrightRed)
    '    Exit Sub
    'End If
    
    If FindOpenInvSlot(Index, GetPlayerEquipment(Index, EqSlot)) > 0 Then
        i = GiveInvItem(Index, GetPlayerEquipment(Index, EqSlot), 0, GetPlayerEquipmentDur(Index, EqSlot), GetPlayerEquipmentBind(Index, EqSlot), False)

        ' Prevent subscript out of range
        If i < 1 Or i > MAX_ITEMS Then Exit Sub
        
        PlayerMsg Index, "You unequip " & CheckGrammar(Trim$(Item(GetPlayerEquipment(Index, EqSlot)).Name)) & ".", Yellow
        
        ' Send the sound
        If SendSound Then
            SendPlayerSound Index, GetPlayerX(Index), GetPlayerY(Index), SoundEntity.seItem, GetPlayerEquipment(Index, EqSlot)
        End If
        
        ' Remove equipment
        SetPlayerEquipment Index, 0, EqSlot
        SetPlayerEquipmentDur Index, 0, EqSlot
        SetPlayerEquipmentBind Index, 0, EqSlot
        
        SendInventoryUpdate Index, i
        
        If SendUpdate Then
            Call SendWornEquipment(Index)
            Call SendMapEquipment(Index)
            SendPlayerStats Index
            
            ' Send vitals
            For i = 1 To Vitals.Vital_Count - 1
                Call SendVital(Index, i)
            Next
            
            ' Send vitals to party if in one
            If tempplayer(Index).InParty > 0 Then SendPartyVitals tempplayer(Index).InParty, Index
        End If
    ElseIf Not Ignore Then
        PlayerMsg Index, "Your inventory is full.", BrightRed
    Else
        i = FindOpenMapItemSlot(GetPlayerMap(Index))
        If i <> 0 Then
            Call SpawnItemSlot(i, GetPlayerEquipment(Index, EqSlot), 1, GetPlayerEquipmentDur(Index, EqSlot), GetPlayerMap(Index), GetPlayerX(Index), GetPlayerY(Index), GetPlayerName(Index))
        End If
    End If
End Sub

Public Sub CheckSpellRankUp(ByVal Index As Long, ByVal SpellNum As Long, ByVal SpellSlot As Byte)
    Dim i As Long
    
    ' Check if they have enough to rank up their spell
    If Spell(Spell(SpellNum).NewSpell).CastRequired <= Account(Index).Chars(GetPlayerChar(Index)).AmountOfCasts(SpellSlot) Then
        ' Check if they meet the level to learn the spell
        If Spell(Spell(SpellNum).NewSpell).LevelReq <= GetPlayerLevel(Index) Then
            ' Send the message update
            Call PlayerMsg(Index, "You have ranked up the spell " & Trim$(Spell(SpellNum).Name) & "!", BrightGreen)
            
            ' Set the hotbar to the new spell
            For i = 1 To MAX_HOTBAR
                If Account(Index).Chars(GetPlayerChar(Index)).Hotbar(i).Slot = SpellNum And Account(Index).Chars(GetPlayerChar(Index)).Hotbar(i).SType = 2 Then
                    Account(Index).Chars(GetPlayerChar(Index)).Hotbar(i).Slot = Spell(SpellNum).NewSpell
                    SendHotbar Index
                End If
            Next
            
            ' Set it to the new spell
            Call SetPlayerSpell(Index, SpellSlot, Spell(SpellNum).NewSpell)
            
            ' Reset the cooldown
            Call SetPlayerSpellCD(Index, SpellSlot, 0)
            
            ' Reset the amount of casts
            Account(Index).Chars(GetPlayerChar(Index)).AmountOfCasts(SpellSlot) = 0
            
            ' Update the cooldown
            Call SendSpellCooldown(Index, SpellSlot)
            
            ' Update the spell
            Call SendPlayerSpell(Index, SpellSlot)
        End If
    End If
End Sub

Public Sub CheckPlayerNewTitle(ByVal Index As Long, Optional ByVal Message As Boolean = True)
    Dim i As Byte, X As Byte
    
    For i = 1 To MAX_TITLES
        For X = 1 To MAX_TITLES
            If CanAddTitle(Index, i) Then
                ' Find an empty slot
                If GetPlayerTitle(Index, X) = 0 Then
                    ' Set the title
                    Account(Index).Chars(GetPlayerChar(Index)).Title(X) = i
                    Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles = Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles + 1

                    If Message Then
                        Call PlayerMsg(Index, "You have unlocked the " & Trim$(Title(i).Name) & " title!", BrightGreen)
                    End If
                    
                    ' Set the current title
                    If Account(Index).Chars(GetPlayerChar(Index)).CurrentTitle = 0 Then
                        Account(Index).Chars(GetPlayerChar(Index)).CurrentTitle = i
                    End If
                    
                    ' Send updated title
                    Call SendPlayerTitles(Index)
                    Exit For
                End If
            End If
        Next
    Next
End Sub

Public Sub AddPlayerTitle(ByVal Index As Long, ByVal TitleNum As Long, Optional ByVal InvNum As Byte, Optional ByVal Message As Boolean = True)
    Dim i As Byte, X As Byte
    
    If CanAddTitle(Index, TitleNum, False, True) Then
        For i = 1 To MAX_TITLES
            ' Find an empty slot
            If GetPlayerTitle(Index, i) = 0 Then
                ' Set the title
                Call SetPlayerTitle(Index, i, TitleNum)
                Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles = Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles + 1

                If Message Then Call PlayerMsg(Index, "You have unlocked the " & Trim$(Title(TitleNum).Name) & " title!", BrightGreen)

                ' Set the current title
                If Account(Index).Chars(GetPlayerChar(Index)).CurrentTitle = 0 Then
                    Account(Index).Chars(GetPlayerChar(Index)).CurrentTitle = i
                End If
                
                ' Send updated title
                Call SendPlayerTitles(Index)
                
                ' Is it reusable, if not take item away
                If InvNum > 0 Then
                    If Item(GetPlayerInvItemNum(Index, InvNum)).IsReusable = False Then
                        Call TakeInvItem(Index, GetPlayerInvItemNum(Index, InvNum), 1)
                    End If
                End If
                Exit For
            End If
        Next
    End If
End Sub

Public Sub RemovePlayerTitle(ByVal Index As Long, ByVal TitleNum As Long, Optional ByVal InvNum As Byte, Optional ByVal Message As Boolean = True)
    Dim i As Byte, X As Byte
    
    For i = 1 To MAX_TITLES
        ' Find an empty slot
        If GetPlayerTitle(Index, i) > 0 Then
            ' Set the title
            Call SetPlayerTitle(Index, i, 0)
            Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles = Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles - 1

            If Message Then Call PlayerMsg(Index, "You have lost the " & Trim$(Title(TitleNum).Name) & " title!", BrightGreen)

            ' Set the current title
            If Account(Index).Chars(GetPlayerChar(Index)).CurrentTitle = i Then
                Account(Index).Chars(GetPlayerChar(Index)).CurrentTitle = 0
                
                For X = MAX_TITLES To 1 Step -1
                    If Account(Index).Chars(GetPlayerChar(Index)).Title(X) > 0 Then
                        Account(Index).Chars(GetPlayerChar(Index)).CurrentTitle = X
                        Exit For
                    End If
                Next
            End If
            
            ' Send updated title
            Call SendPlayerTitles(Index)
            
            If InvNum > 0 Then Call TakeInvItem(Index, GetPlayerInvItemNum(Index, InvNum), 1)
            Exit For
        End If
    Next
End Sub

Private Function CanAddTitle(ByVal Index As Long, ByVal TitleNum As Byte, Optional ByVal NoRequirement As Boolean = True, Optional ByVal Message As Byte = False) As Boolean
    Dim i As Byte
    
    ' Don't add a title if it has no requirements
    If Title(TitleNum).LevelReq = 0 And Title(TitleNum).PKReq = 0 And NoRequirement Then Exit Function

    ' Check if we don't have one of the possible titles
    If GetPlayerLevel(Index) >= Title(TitleNum).LevelReq And Account(Index).Chars(GetPlayerChar(Index)).PlayerKills >= Title(TitleNum).PKReq Then
        If Len(Trim$(Title(TitleNum).Name)) > 0 Then
            For i = 1 To Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles
                If Account(Index).Chars(GetPlayerChar(Index)).Title(i) = TitleNum Then
                    If Message Then Call PlayerMsg(Index, "You already have this title!", 12)
                    Exit Function
                End If
            Next
        Else
            Exit Function
        End If
    Else
        If Message Then Call PlayerMsg(Index, "You do not meet the requirements to earn this title!", 12)
        Exit Function
    End If
    
    CanAddTitle = True
End Function

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
    Dim FirstLetter As String * 1
   
    FirstLetter = LCase$(Left$(Word, 1))
   
    If FirstLetter = "$" Then
        CheckGrammar = (Mid$(Word, 2, Len(Word) - 1))
        Exit Function
    End If
   
    If FirstLetter Like "*[aeiou]*" Then
        If Caps Then CheckGrammar = "An " & Word Else CheckGrammar = "an " & Word
    Else
        If Caps Then CheckGrammar = "A " & Word Else CheckGrammar = "a " & Word
    End If
End Function

Public Function IsInRange(ByVal Range As Byte, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Boolean
    Dim nVal As Long
    
    IsInRange = False
    
    nVal = Sqr((x1 - x2) ^ 2 + (y1 - y2) ^ 2)
    
    If nVal <= Range Then IsInRange = True
End Function

Public Function IsDirBlocked(ByRef BlockVar As Byte, ByRef Dir As Byte) As Boolean
    If Not BlockVar And (2 ^ Dir) Then
        IsDirBlocked = False
    Else
        IsDirBlocked = True
    End If
End Function

Public Function Random(ByVal Low As Long, ByVal High As Long) As Long
    ' Randomize rnd's seed
    Randomize
    
    Random = Int((High - Low + 1) * Rnd) + Low
End Function

' #####################
' ## Party Functions ##
' #####################
Public Sub Party_PlayerLeave(ByVal Index As Long)
    Dim PartyNum As Long, i As Long

    PartyNum = tempplayer(Index).InParty
    
    If PartyNum > 0 Then
        ' Find out how many members we have
        Party_CountMembers PartyNum
        ' Make sure there's more than 2 people
        If Party(PartyNum).MemberCount > 2 Then
            ' Check if leader
            If Party(PartyNum).Leader = Index Then
                ' Set next person down as leader
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(PartyNum).Member(i) > 0 And Party(PartyNum).Member(i) <> Index Then
                        Party(PartyNum).Leader = Party(PartyNum).Member(i)
                        PartyMsg PartyNum, GetPlayerName(i) & " is now the party leader.", BrightBlue
                        Exit For
                    End If
                Next
                
                ' Leave party
                PartyMsg PartyNum, GetPlayerName(Index) & " has left the party.", BrightRed
                
                ' Clear the PartyNum
                tempplayer(Index).InParty = 0
                
                ' Remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(PartyNum).Member(i) = Index Then
                        Party(PartyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                
                ' Recount party
                Party_CountMembers PartyNum
                
                ' Set update to all
                SendPartyUpdate PartyNum
                
                ' Send clear to player
                SendPartyUpdateTo Index
            Else
                ' Not the leader, just leave
                PartyMsg PartyNum, GetPlayerName(Index) & " has left the party.", BrightRed
                
                ' Remove from array
                For i = 1 To MAX_PARTY_MEMBERS
                    If Party(PartyNum).Member(i) = Index Then
                        Party(PartyNum).Member(i) = 0
                        Exit For
                    End If
                Next
                
                ' Clear the PartyNum
                tempplayer(Index).InParty = 0
                
                ' Recount party
                Party_CountMembers PartyNum
                
                ' Set update to all
                SendPartyUpdate PartyNum
                
                ' Send clear to player
                SendPartyUpdateTo Index
            End If
        Else
            ' Find out how many members we have
            Party_CountMembers PartyNum
            
            ' Only 2 people, disband
            PartyMsg PartyNum, "Party disbanded.", BrightRed
                
            ' Clear out everyone's party
            For i = 1 To MAX_PARTY_MEMBERS
                Index = Party(PartyNum).Member(i)
                ' Player exist?
                If Index > 0 Then
                    ' Remove them
                    tempplayer(i).InParty = 0
                    Party(PartyNum).Member(i) = 0
                    
                    ' Send clear to players
                    SendPartyUpdateTo i
                End If
            Next
            
            ' Clear out the party itself
            ClearParty PartyNum
        End If
    End If
End Sub

Public Sub Party_Invite(ByVal Index As Long, ByVal OtherPlayer As Long)
    Dim PartyNum As Long, i As Long
    
    ' Make sure they're not in a party
    If tempplayer(OtherPlayer).InParty > 0 Then
        ' They're already in a party
        PlayerMsg Index, "This player is already in a party!", BrightRed
        Exit Sub
    End If
    
    ' Check if there doing another action
    If IsPlayerBusy(Index, OtherPlayer) Then Exit Sub
    
    ' Check if we're in a party
    If tempplayer(Index).InParty > 0 Then
        PartyNum = tempplayer(Index).InParty
        ' Make sure we're the leader
        If Party(PartyNum).Leader = Index Then
            ' Got a blank slot?
            For i = 1 To MAX_PARTY_MEMBERS
                If Party(PartyNum).Member(i) = 0 Then
                    ' Send the invitation
                    SendPartyInvite OtherPlayer, Index
                    
                    ' Set the invite target
                    tempplayer(OtherPlayer).PartyInvite = Index
                    
                    ' Let them know
                    PlayerMsg Index, "Party invitation sent.", Pink
                    Exit Sub
                End If
            Next
            
            ' No room
            PlayerMsg Index, "Party is full!", BrightRed
            Exit Sub
        Else
            ' Not the leader
            PlayerMsg Index, "You are not the party leader!", BrightRed
            Exit Sub
        End If
    Else
        ' Not in a party - doesn't matter
        SendPartyInvite OtherPlayer, Index
        
        ' Set the invite target
        tempplayer(OtherPlayer).PartyInvite = Index
        
        ' Let them know
        PlayerMsg Index, "Party invitation sent.", Pink
        Exit Sub
    End If
End Sub

Public Sub Party_InviteAccept(ByVal Index As Long, ByVal OtherPlayer As Long)
    Dim PartyNum As Byte, i As Long

    ' Check if already in a party
    If tempplayer(Index).InParty > 0 Then
        ' Get the PartyNumber
        PartyNum = tempplayer(Index).InParty
        ' Got a blank slot?
        For i = 1 To MAX_PARTY_MEMBERS
            If Party(PartyNum).Member(i) = 0 Then
                ' Clear party invite
                tempplayer(OtherPlayer).PartyInvite = 0
                
                ' Add to the party
                Party(PartyNum).Member(i) = OtherPlayer
                
                ' Recount party
                Party_CountMembers PartyNum
                
                ' Send update to all - including new player
                SendPartyUpdate PartyNum
                SendPartyVitals PartyNum, OtherPlayer
                
                ' Let everyone know they've joined
                PartyMsg PartyNum, GetPlayerName(OtherPlayer) & " has joined the party.", BrightGreen
                
                ' Add them in
                tempplayer(OtherPlayer).InParty = PartyNum
                Exit Sub
            End If
        Next
        
        ' No empty slots - let them know
        PlayerMsg Index, "Party is full!", BrightRed
        PlayerMsg OtherPlayer, "Party is full!", BrightRed
        Exit Sub
    Else
        ' Not in a party. Create one with the new person.
        For i = 1 To MAX_PARTYS
            ' Find blank party
            If Not Party(i).Leader > 0 Then
                PartyNum = i
                Exit For
            End If
        Next
        
        ' Create the party
        Party(PartyNum).MemberCount = 2
        Party(PartyNum).Leader = Index
        Party(PartyNum).Member(1) = Index
        Party(PartyNum).Member(2) = OtherPlayer
        SendPartyUpdate PartyNum
        SendPartyVitals PartyNum, Index
        SendPartyVitals PartyNum, OtherPlayer
        
        ' Let them know it's created
        PartyMsg PartyNum, "Party created.", BrightGreen
        
        ' Clear the invitation
        tempplayer(OtherPlayer).PartyInvite = 0
       
       ' Add them to the party
        tempplayer(OtherPlayer).InParty = PartyNum
        tempplayer(Index).InParty = PartyNum
    End If
End Sub

Public Sub Party_InviteDecline(ByVal Index As Long, ByVal OtherPlayer As Long)
    If IsPlaying(Index) Then
        PlayerMsg Index, GetPlayerName(OtherPlayer) & " has declined to join the party!", BrightRed
    End If
    
    PlayerMsg OtherPlayer, "You declined to join the party!", BrightRed
    
    ' Clear the invitation
    tempplayer(OtherPlayer).PartyInvite = 0
End Sub

Public Sub Party_CountMembers(ByVal PartyNum As Long)
    Dim i As Long, highIndex As Long, X As Long
    
    ' Find the high Index
    For i = MAX_PARTY_MEMBERS To 1 Step -1
        If Party(PartyNum).Member(i) > 0 Then
            highIndex = i
            Exit For
        End If
    Next
    
    ' Count the members
    For i = 1 To MAX_PARTY_MEMBERS
        ' We've got a blank member
        If Party(PartyNum).Member(i) = 0 Then
            ' Is it lower than the high Index?
            If i < highIndex Then
                ' Move everyone down a slot
                For X = i To MAX_PARTY_MEMBERS - 1
                    Party(PartyNum).Member(X) = Party(PartyNum).Member(X + 1)
                    Party(PartyNum).Member(X + 1) = 0
                Next
            Else
                ' Not lower - highIndex is count
                Party(PartyNum).MemberCount = highIndex
                Exit Sub
            End If
        End If
        
        ' Check if we've reached the max
        If i = MAX_PARTY_MEMBERS Then
            If highIndex = i Then
                Party(PartyNum).MemberCount = MAX_PARTY_MEMBERS
                Exit Sub
            End If
        End If
    Next
    
    ' If we're here it means that we need to re-count again
    Party_CountMembers PartyNum
End Sub

Public Sub Party_ShareExp(ByVal PartyNum As Long, ByVal Exp As Long, ByVal Index As Long)
    Dim ExpShare As Long, LeftOver As Long, i As Long, tmpIndex As Long

    ' Check if it's worth sharing
    If Not Exp >= Party(PartyNum).MemberCount Then
        ' No party - keep exp for self
        GivePlayerEXP Index, Exp
        Exit Sub
    End If
    
    ' Find out the equal share
    ExpShare = Exp \ Party(PartyNum).MemberCount
    LeftOver = Exp Mod Party(PartyNum).MemberCount
    
    ' Loop through and give everyone exp
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(PartyNum).Member(i)
        
        ' Existing member?
        If tmpIndex > 0 Then
            ' Playing?
            If IsConnected(tmpIndex) And IsPlaying(tmpIndex) Then
                ' Give them their share
                GivePlayerEXP tmpIndex, ExpShare
            End If
        End If
    Next
    
    ' Give the remainder to a random member
    tmpIndex = Party(PartyNum).Member(Random(1, Party(PartyNum).MemberCount))
    
    ' Give the exp
    GivePlayerEXP tmpIndex, LeftOver
End Sub

Public Sub GivePlayerEXP(ByVal Index As Long, ByVal Exp As Long)
    If GetPlayerLevel(Index) = Options.MaxLevel Then Exit Sub
    
    ' Give the exp
    Call SetPlayerExp(Index, GetPlayerExp(Index) + Exp)
    
    SendPlayerExp Index
    SendActionMsg GetPlayerMap(Index), "+" & Exp & " Exp", White, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 32)
    
    ' Check if we've leveled
    CheckPlayerLevelUp Index
End Sub

Public Function Clamp(ByVal Value As Long, ByVal Min As Long, ByVal Max As Long) As Long
    Clamp = Value
    
    If Value < Min Then Clamp = Min
    If Value > Max Then Clamp = Max
End Function

Public Function GetSkillName(ByVal SkillNum As Byte) As String
    Select Case SkillNum
        Case Skills.Alchemy: GetSkillName = "Alchemy"
        Case Skills.Cooking: GetSkillName = "Cooking"
        Case Skills.Crafting: GetSkillName = "Crafting"
        Case Skills.Farming: GetSkillName = "Farming"
        Case Skills.Firemaking: GetSkillName = "Firemaking"
        Case Skills.Fishing: GetSkillName = "Fishing"
        Case Skills.Fletching: GetSkillName = "Fletching"
        Case Skills.Herbalism: GetSkillName = "Herbalism"
        Case Skills.Prayer: GetSkillName = "Prayer"
        Case Skills.Smithing: GetSkillName = "Smithing"
        Case Skills.Woodcutting: GetSkillName = "Woodcutting"
        Case Skills.Mining: GetSkillName = "Mining"
    End Select
End Function

Public Function GetProficiencyName(ByVal ProficiencyNum As Byte) As String
    Select Case ProficiencyNum
        Case Proficiency.Medium: GetProficiencyName = "Medium"
        Case Proficiency.Sword: GetProficiencyName = "Sword"
        Case Proficiency.Staff: GetProficiencyName = "Staff"
        Case Proficiency.Spear: GetProficiencyName = "Spear"
        Case Proficiency.Heavy: GetProficiencyName = "Heavy"
        Case Proficiency.Mace: GetProficiencyName = "Mace"
        Case Proficiency.Dagger: GetProficiencyName = "Dagger"
        Case Proficiency.Crossbow: GetProficiencyName = "Crossbow"
        Case Proficiency.Light: GetProficiencyName = "Light"
        Case Proficiency.Bow: GetProficiencyName = "Bow"
        Case Proficiency.Axe: GetProficiencyName = "Axe"
    End Select
End Function

Public Sub DeclineTradeRequest(ByVal Index As Long)
    If IsPlaying(tempplayer(Index).TradeRequest) Then
        PlayerMsg tempplayer(Index).TradeRequest, GetPlayerName(Index) & " has declined your trade request!", BrightRed
    End If
    PlayerMsg Index, "You decline the trade request.", BrightRed
    
    ' Clear the tradeRequest server-side
    tempplayer(Index).TradeRequest = 0
End Sub

Public Sub DeclineGuildInvite(ByVal Index As Long)
    If IsPlaying(tempplayer(Index).GuildInvite) Then
        Call PlayerMsg(tempplayer(Index).GuildInvite, GetPlayerName(Index) & " has declined the guild invitation!", BrightRed)
    End If
    
    PlayerMsg Index, "You declined to join the guild.", BrightRed
    
    ' Clear the guild invite server-side
    tempplayer(Index).GuildInvite = 0
End Sub

Sub Guild_Disband(ByVal Index As Long)
    Dim i As Long, tmpIndex As Long, tmpGuild As Long
    
    ' Subscript out of range
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Sub
    
    ' Make sure they are actually in a guild
    If GetPlayerGuild(Index) = 0 Then Exit Sub
    
    ' Make sure they are the guild leader
    If GetPlayerGuildAccess(Index) < MAX_GUILDACCESS Then Exit Sub

    tmpGuild = GetPlayerGuild(Index)
    
    Call GlobalMsg(GetPlayerName(Index) & " has disbanded the guild " & GetPlayerGuildName(Index) & "!", BrightRed)
    Guild(tmpGuild).Name = vbNullString
    Guild(tmpGuild).MOTD = vbNullString
    
    ' Remove them
    For i = 1 To MAX_GUILD_MEMBERS
        If Not Guild(tmpGuild).Members(i) = vbNullString Then
            tmpIndex = FindPlayer(Guild(tmpGuild).Members(i))
            
            Call LoadTempGuildMember(i, Guild(tmpGuild).Members(i))
            TempGuildMember(i).Guild.Index = 0
            TempGuildMember(i).Guild.Access = 0
            Call SaveTempGuildMember(i, Guild(tmpGuild).Members(i))
            Guild(tmpGuild).Members(i) = vbNullString

            Call ClearTempGuildMember(i)
            
            ' Send update
            If IsPlaying(tmpIndex) Then
                Call SetPlayerGuild(tmpIndex, 0)
                Call SetPlayerGuildAccess(tmpIndex, 0)
                Call SendPlayerGuild(tmpIndex)
            End If
        End If
    Next
    
    Call SaveGuild(tmpGuild)
End Sub

Public Sub SpawnMapEventsFor(Index As Long, MapNum As Long)
    Dim i As Long, X As Long, Y As Long, z As Long, spawncurrentevent As Boolean, p As Long
    Dim buffer As clsBuffer
    
    tempplayer(Index).EventMap.CurrentEvents = 0
    ReDim tempplayer(Index).EventMap.EventPages(0)
    
    If Map(MapNum).EventCount <= 0 Then Exit Sub
    For i = 1 To Map(MapNum).EventCount
        If Map(MapNum).Events(i).PageCount > 0 Then
            For z = Map(MapNum).Events(i).PageCount To 1 Step -1
                With Map(MapNum).Events(i).Pages(z)
                    spawncurrentevent = True
                    
                    If .chkVariable = 1 Then
                        If Account(Index).Chars(GetPlayerChar(Index)).Variables(.VariableIndex) < .VariableCondition Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSwitch = 1 Then
                        If Account(Index).Chars(GetPlayerChar(Index)).Switches(.SwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkHasItem = 1 Then
                        If HasItem(Index, .HasItemIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If .chkSelfSwitch = 1 Then
                        If Map(MapNum).Events(i).SelfSwitches(.SelfSwitchIndex) = 0 Then
                            spawncurrentevent = False
                        End If
                    End If
                    
                    If spawncurrentevent = True Or (spawncurrentevent = False And z = 1) Then
                        ' Spawn the event and send data to player
                        tempplayer(Index).EventMap.CurrentEvents = tempplayer(Index).EventMap.CurrentEvents + 1
                        
                        ReDim Preserve tempplayer(Index).EventMap.EventPages(tempplayer(Index).EventMap.CurrentEvents)
                        
                        With tempplayer(Index).EventMap.EventPages(tempplayer(Index).EventMap.CurrentEvents)
                            If Map(MapNum).Events(i).Pages(z).GraphicType = 1 Then
                                Select Case Map(MapNum).Events(i).Pages(z).GraphicY
                                    Case 0
                                        .Dir = DIR_DOWN
                                    Case 1
                                        .Dir = DIR_LEFT
                                    Case 2
                                        .Dir = DIR_RIGHT
                                    Case 3
                                        .Dir = DIR_UP
                                End Select
                            Else
                                .Dir = 0
                            End If
                            
                            .GraphicNum = Map(MapNum).Events(i).Pages(z).Graphic
                            .GraphicType = Map(MapNum).Events(i).Pages(z).GraphicType
                            .GraphicX = Map(MapNum).Events(i).Pages(z).GraphicX
                            .GraphicY = Map(MapNum).Events(i).Pages(z).GraphicY
                            .GraphicX2 = Map(MapNum).Events(i).Pages(z).GraphicX2
                            .GraphicY2 = Map(MapNum).Events(i).Pages(z).GraphicY2
                            
                            Select Case Map(MapNum).Events(i).Pages(z).MoveSpeed
                                Case 0
                                    .MovementSpeed = 2
                                Case 1
                                    .MovementSpeed = 3
                                Case 2
                                    .MovementSpeed = 4
                                Case 3
                                    .MovementSpeed = 6
                                Case 4
                                    .MovementSpeed = 12
                                Case 5
                                    .MovementSpeed = 24
                            End Select
                            
                            If Map(MapNum).Events(i).Global Then
                                .X = TempEventMap(MapNum).Events(i).X
                                .Y = TempEventMap(MapNum).Events(i).Y
                                .Dir = TempEventMap(MapNum).Events(i).Dir
                                .MoveRouteStep = TempEventMap(MapNum).Events(i).MoveRouteStep
                            Else
                                .X = Map(MapNum).Events(i).X
                                .Y = Map(MapNum).Events(i).Y
                                .MoveRouteStep = 0
                            End If
                            
                            .Position = Map(MapNum).Events(i).Pages(z).Position
                            .eventID = i
                            .PageID = z
                            
                            If spawncurrentevent = True Then
                                .Visible = 1
                            Else
                                .Visible = 0
                            End If
                            
                            .MoveType = Map(MapNum).Events(i).Pages(z).MoveType
                            
                            If .MoveType = 2 Then
                                .MoveRouteCount = Map(MapNum).Events(i).Pages(z).MoveRouteCount
                                ReDim .MoveRoute(0 To Map(MapNum).Events(i).Pages(z).MoveRouteCount)
                                If Map(MapNum).Events(i).Pages(z).MoveRouteCount > 0 Then
                                    For p = 0 To Map(MapNum).Events(i).Pages(z).MoveRouteCount
                                        .MoveRoute(p) = Map(MapNum).Events(i).Pages(z).MoveRoute(p)
                                    Next
                                End If
                            End If
                            
                            .RepeatMoveRoute = Map(MapNum).Events(i).Pages(z).RepeatMoveRoute
                            .IgnoreIfCannotMove = Map(MapNum).Events(i).Pages(z).IgnoreMoveRoute
                                
                            .MoveFreq = Map(MapNum).Events(i).Pages(z).MoveFreq
                            .MoveSpeed = Map(MapNum).Events(i).Pages(z).MoveSpeed
                            
                            .WalkingAnim = Map(MapNum).Events(i).Pages(z).WalkAnim
                            .WalkThrough = Map(MapNum).Events(i).Pages(z).WalkThrough
                            .ShowName = Map(MapNum).Events(i).Pages(z).ShowName
                            .FixedDir = Map(MapNum).Events(i).Pages(z).DirFix
                            .Trigger = Map(MapNum).Events(i).Pages(z).Trigger
                        End With
                        GoTo nextevent
                    End If
                End With
            Next
        End If
nextevent:
    Next
    
    If tempplayer(Index).EventMap.CurrentEvents > 0 Then
        For i = 1 To tempplayer(Index).EventMap.CurrentEvents
            Set buffer = New clsBuffer
            
            buffer.WriteLong SSpawnEvent
            buffer.WriteLong i
            
            With tempplayer(Index).EventMap.EventPages(i)
                buffer.WriteString Map(GetPlayerMap(Index)).Events(i).Name
                buffer.WriteLong .Dir
                buffer.WriteLong .GraphicNum
                buffer.WriteLong .GraphicType
                buffer.WriteLong .GraphicX
                buffer.WriteLong .GraphicX2
                buffer.WriteLong .GraphicY
                buffer.WriteLong .GraphicY2
                buffer.WriteLong .MovementSpeed
                buffer.WriteLong .X
                buffer.WriteLong .Y
                buffer.WriteLong .Position
                buffer.WriteLong .Visible
                buffer.WriteLong Map(MapNum).Events(.eventID).Pages(.PageID).WalkAnim
                buffer.WriteLong Map(MapNum).Events(.eventID).Pages(.PageID).DirFix
                buffer.WriteLong Map(MapNum).Events(.eventID).Pages(.PageID).WalkThrough
                buffer.WriteLong Map(MapNum).Events(.eventID).Pages(.PageID).ShowName
                buffer.WriteByte Map(MapNum).Events(.eventID).Pages(.PageID).Trigger
            End With
            
            SendDataTo Index, buffer.ToArray
            Set buffer = Nothing
        Next
    End If
End Sub

Sub SpawnAllMapGlobalEvents()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call SpawnGlobalEvents(i)
    Next
End Sub

Sub SpawnGlobalEvents(ByVal MapNum As Long)
    Dim i As Long, z As Long
    TempEventMap(MapNum).EventCount = 0
    ReDim TempEventMap(MapNum).Events(0)
    If Map(MapNum).EventCount > 0 Then
        For i = 1 To Map(MapNum).EventCount
            TempEventMap(MapNum).EventCount = TempEventMap(MapNum).EventCount + 1
            ReDim Preserve TempEventMap(MapNum).Events(0 To TempEventMap(MapNum).EventCount)
            If Map(MapNum).Events(i).PageCount > 0 Then
                If Map(MapNum).Events(i).Global = 1 Then
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).X = Map(MapNum).Events(i).X
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Y = Map(MapNum).Events(i).Y
                    If Map(MapNum).Events(i).Pages(1).GraphicType = 1 Then
                        Select Case Map(MapNum).Events(i).Pages(1).GraphicY
                            Case 0
                                TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Dir = DIR_DOWN
                            Case 1
                                TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Dir = DIR_LEFT
                            Case 2
                                TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Dir = DIR_RIGHT
                            Case 3
                                TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Dir = DIR_UP
                        End Select
                    Else
                        TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Dir = DIR_DOWN
                    End If
                    
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Active = 1
                    
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveType = Map(MapNum).Events(i).Pages(1).MoveType
                    
                    If TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveType = 2 Then
                        TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveRouteCount = Map(MapNum).Events(i).Pages(1).MoveRouteCount
                        ReDim TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveRoute(0 To Map(MapNum).Events(i).Pages(1).MoveRouteCount)
                        For z = 0 To Map(MapNum).Events(i).Pages(1).MoveRouteCount
                            TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveRoute(z) = Map(MapNum).Events(i).Pages(1).MoveRoute(z)
                        Next
                    End If
                    
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).RepeatMoveRoute = Map(MapNum).Events(i).Pages(1).RepeatMoveRoute
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).IgnoreIfCannotMove = Map(MapNum).Events(i).Pages(1).IgnoreMoveRoute
                    
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveFreq = Map(MapNum).Events(i).Pages(1).MoveFreq
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).MoveSpeed = Map(MapNum).Events(i).Pages(1).MoveSpeed
                    
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).WalkThrough = Map(MapNum).Events(i).Pages(1).WalkThrough
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).FixedDir = Map(MapNum).Events(i).Pages(1).DirFix
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).WalkingAnim = Map(MapNum).Events(i).Pages(1).WalkAnim
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).ShowName = Map(MapNum).Events(i).Pages(1).ShowName
                    TempEventMap(MapNum).Events(TempEventMap(MapNum).EventCount).Trigger = Map(MapNum).Events(i).Pages(1).Trigger
                End If
            End If
        Next
    End If
End Sub

Function CanEventMove(Index As Long, ByVal MapNum As Long, X As Long, Y As Long, eventID As Long, WalkThrough As Long, ByVal Dir As Byte, Optional GlobalEvent As Boolean = False) As Boolean
    Dim i As Long
    Dim n As Long, z As Long

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_DOWNRIGHT Then Exit Function
    
    CanEventMove = True
    
    Select Case Dir
        Case DIR_UP
            ' Check to make sure not outside of boundries
            If Y > 0 Then
                n = Map(MapNum).Tile(X, Y - 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If
                
                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = X) And (GetPlayerY(i) = Y - 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNPC(MapNum).NPC(i).X = X) And (MapNPC(MapNum).NPC(i).Y = Y - 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If GlobalEvent = True Then
                    If TempEventMap(MapNum).EventCount > 0 Then
                        For z = 1 To TempEventMap(MapNum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(MapNum).Events(z).X = X) And (TempEventMap(MapNum).Events(z).Y = Y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If tempplayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To tempplayer(Index).EventMap.CurrentEvents
                            If (tempplayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (tempplayer(Index).EventMap.EventPages(z).X = tempplayer(Index).EventMap.EventPages(eventID).X) And (tempplayer(Index).EventMap.EventPages(z).Y = tempplayer(Index).EventMap.EventPages(eventID).Y - 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(X, Y).DirBlock, DIR_UP + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_DOWN
            ' Check to make sure not outside of boundries
            If Y < Map(MapNum).MaxY Then
                n = Map(MapNum).Tile(X, Y + 1).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = X) And (GetPlayerY(i) = Y + 1) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNPC(MapNum).NPC(i).X = X) And (MapNPC(MapNum).NPC(i).Y = Y + 1) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If GlobalEvent = True Then
                    If TempEventMap(MapNum).EventCount > 0 Then
                        For z = 1 To TempEventMap(MapNum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(MapNum).Events(z).X = X) And (TempEventMap(MapNum).Events(z).Y = Y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If tempplayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To tempplayer(Index).EventMap.CurrentEvents
                            If (tempplayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (tempplayer(Index).EventMap.EventPages(z).X = tempplayer(Index).EventMap.EventPages(eventID).X) And (tempplayer(Index).EventMap.EventPages(z).Y = tempplayer(Index).EventMap.EventPages(eventID).Y + 1) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(X, Y).DirBlock, DIR_DOWN + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_LEFT
            ' Check to make sure not outside of boundries
            If X > 0 Then
                n = Map(MapNum).Tile(X - 1, Y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = X - 1) And (GetPlayerY(i) = Y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNPC(MapNum).NPC(i).X = X - 1) And (MapNPC(MapNum).NPC(i).Y = Y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If GlobalEvent = True Then
                    If TempEventMap(MapNum).EventCount > 0 Then
                        For z = 1 To TempEventMap(MapNum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(MapNum).Events(z).X = X - 1) And (TempEventMap(MapNum).Events(z).Y = Y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If tempplayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To tempplayer(Index).EventMap.CurrentEvents
                            If (tempplayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (tempplayer(Index).EventMap.EventPages(z).X = tempplayer(Index).EventMap.EventPages(eventID).X - 1) And (tempplayer(Index).EventMap.EventPages(z).Y = tempplayer(Index).EventMap.EventPages(eventID).Y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(X, Y).DirBlock, DIR_LEFT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If

        Case DIR_RIGHT
            ' Check to make sure not outside of boundries
            If X < Map(MapNum).MaxX Then
                n = Map(MapNum).Tile(X + 1, Y).Type
                
                If WalkThrough = 1 Then
                    CanEventMove = True
                    Exit Function
                End If

                ' Check to make sure that the tile is walkable
                If n <> TILE_TYPE_WALKABLE And n <> TILE_TYPE_NPCSPAWN Then
                    CanEventMove = False
                    Exit Function
                End If

                ' Check to make sure that there is not a player in the way
                For i = 1 To Player_HighIndex
                    If IsPlaying(i) Then
                        If (GetPlayerMap(i) = MapNum) And (GetPlayerX(i) = X + 1) And (GetPlayerY(i) = Y) Then
                            CanEventMove = False
                            Exit Function
                        End If
                    End If
                Next

                ' Check to make sure that there is not another npc in the way
                For i = 1 To MAX_MAP_NPCS
                    If (MapNPC(MapNum).NPC(i).X = X + 1) And (MapNPC(MapNum).NPC(i).Y = Y) Then
                        CanEventMove = False
                        Exit Function
                    End If
                Next
                
                If GlobalEvent = True Then
                    If TempEventMap(MapNum).EventCount > 0 Then
                        For z = 1 To TempEventMap(MapNum).EventCount
                            If (z <> eventID) And (z > 0) And (TempEventMap(MapNum).Events(z).X = X + 1) And (TempEventMap(MapNum).Events(z).Y = Y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                Else
                    If tempplayer(Index).EventMap.CurrentEvents > 0 Then
                        For z = 1 To tempplayer(Index).EventMap.CurrentEvents
                            If (tempplayer(Index).EventMap.EventPages(z).eventID <> eventID) And (eventID > 0) And (tempplayer(Index).EventMap.EventPages(z).X = tempplayer(Index).EventMap.EventPages(eventID).X + 1) And (tempplayer(Index).EventMap.EventPages(z).Y = tempplayer(Index).EventMap.EventPages(eventID).Y) Then
                                CanEventMove = False
                                Exit Function
                            End If
                        Next
                    End If
                End If
                
                ' Directional blocking
                If IsDirBlocked(Map(MapNum).Tile(X, Y).DirBlock, DIR_RIGHT + 1) Then
                    CanEventMove = False
                    Exit Function
                End If
            Else
                CanEventMove = False
            End If
    End Select
End Function

Sub EventDir(playerindex As Long, ByVal MapNum As Long, ByVal eventID As Long, ByVal Dir As Long, Optional GlobalEvent As Boolean = False)
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_DOWNRIGHT Then Exit Sub
    
    If GlobalEvent Then
        If Map(MapNum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(MapNum).Events(eventID).Dir = Dir
    Else
        If Map(MapNum).Events(eventID).Pages(tempplayer(playerindex).EventMap.EventPages(eventID).PageID).DirFix = 0 Then tempplayer(playerindex).EventMap.EventPages(eventID).Dir = Dir
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEventDir
    buffer.WriteLong eventID
    
    If GlobalEvent Then
        buffer.WriteLong TempEventMap(MapNum).Events(eventID).Dir
    Else
        buffer.WriteLong tempplayer(playerindex).EventMap.EventPages(eventID).Dir
    End If
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub EventMove(Index As Long, MapNum As Long, ByVal eventID As Long, ByVal Dir As Long, MovementSpeed As Long, Optional GlobalEvent As Boolean = False)
    Dim packet As String
    Dim buffer As clsBuffer

    ' Check for subscript out of range
    If MapNum <= 0 Or MapNum > MAX_MAPS Or Dir < DIR_UP Or Dir > DIR_DOWNRIGHT Then Exit Sub
    
    If GlobalEvent Then
        If Map(MapNum).Events(eventID).Pages(1).DirFix = 0 Then TempEventMap(MapNum).Events(eventID).Dir = Dir
        UpdateMapBlock MapNum, TempEventMap(MapNum).Events(eventID).X, TempEventMap(MapNum).Events(eventID).Y, False
    Else
        If Map(MapNum).Events(eventID).Pages(tempplayer(Index).EventMap.EventPages(eventID).PageID).DirFix = 0 Then tempplayer(Index).EventMap.EventPages(eventID).Dir = Dir
    End If

    Select Case Dir
        Case DIR_UP
            If GlobalEvent Then
                TempEventMap(MapNum).Events(eventID).Y = TempEventMap(MapNum).Events(eventID).Y - 1
                UpdateMapBlock MapNum, TempEventMap(MapNum).Events(eventID).X, TempEventMap(MapNum).Events(eventID).Y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).X
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).Y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).Dir
                buffer.WriteLong MovementSpeed
                
                If GlobalEvent Then
                    SendDataToMap MapNum, buffer.ToArray()
                Else
                    SendDataTo Index, buffer.ToArray
                End If
                
                Set buffer = Nothing
            Else
                tempplayer(Index).EventMap.EventPages(eventID).Y = tempplayer(Index).EventMap.EventPages(eventID).Y - 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).X
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).Y
                buffer.WriteLong Dir
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong MovementSpeed
                
                If GlobalEvent Then
                    SendDataToMap MapNum, buffer.ToArray()
                Else
                    SendDataTo Index, buffer.ToArray
                End If
                
                Set buffer = Nothing
            End If
            
        Case DIR_DOWN
            If GlobalEvent Then
                TempEventMap(MapNum).Events(eventID).Y = TempEventMap(MapNum).Events(eventID).Y + 1
                UpdateMapBlock MapNum, TempEventMap(MapNum).Events(eventID).X, TempEventMap(MapNum).Events(eventID).Y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).X
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).Y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).Dir
                buffer.WriteLong MovementSpeed
                If GlobalEvent Then
                    SendDataToMap MapNum, buffer.ToArray()
                Else
                    SendDataTo Index, buffer.ToArray
                End If
                Set buffer = Nothing
            Else
                tempplayer(Index).EventMap.EventPages(eventID).Y = tempplayer(Index).EventMap.EventPages(eventID).Y + 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).X
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).Y
                buffer.WriteLong Dir
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong MovementSpeed
                
                If GlobalEvent Then
                    SendDataToMap MapNum, buffer.ToArray()
                Else
                    SendDataTo Index, buffer.ToArray
                End If
                
                Set buffer = Nothing
            End If
        Case DIR_LEFT
            If GlobalEvent Then
                TempEventMap(MapNum).Events(eventID).X = TempEventMap(MapNum).Events(eventID).X - 1
                UpdateMapBlock MapNum, TempEventMap(MapNum).Events(eventID).X, TempEventMap(MapNum).Events(eventID).Y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).X
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).Y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).Dir
                buffer.WriteLong MovementSpeed
                
                If GlobalEvent Then
                    SendDataToMap MapNum, buffer.ToArray()
                Else
                    SendDataTo Index, buffer.ToArray
                End If
                
                Set buffer = Nothing
            Else
                tempplayer(Index).EventMap.EventPages(eventID).X = tempplayer(Index).EventMap.EventPages(eventID).X - 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).X
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).Y
                buffer.WriteLong Dir
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong MovementSpeed
                
                If GlobalEvent Then
                    SendDataToMap MapNum, buffer.ToArray()
                Else
                    SendDataTo Index, buffer.ToArray
                End If
                
                Set buffer = Nothing
            End If
        Case DIR_RIGHT
            If GlobalEvent Then
                TempEventMap(MapNum).Events(eventID).X = TempEventMap(MapNum).Events(eventID).X + 1
                UpdateMapBlock MapNum, TempEventMap(MapNum).Events(eventID).X, TempEventMap(MapNum).Events(eventID).Y, True
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).X
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).Y
                buffer.WriteLong Dir
                buffer.WriteLong TempEventMap(MapNum).Events(eventID).Dir
                buffer.WriteLong MovementSpeed
                
                If GlobalEvent Then
                    SendDataToMap MapNum, buffer.ToArray()
                Else
                    SendDataTo Index, buffer.ToArray
                End If
                
                Set buffer = Nothing
            Else
                tempplayer(Index).EventMap.EventPages(eventID).X = tempplayer(Index).EventMap.EventPages(eventID).X + 1
                Set buffer = New clsBuffer
                buffer.WriteLong SEventMove
                buffer.WriteLong eventID
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).X
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).Y
                buffer.WriteLong Dir
                buffer.WriteLong tempplayer(Index).EventMap.EventPages(eventID).Dir
                buffer.WriteLong MovementSpeed
                
                If GlobalEvent Then
                    SendDataToMap MapNum, buffer.ToArray()
                Else
                    SendDataTo Index, buffer.ToArray
                End If
                
                Set buffer = Nothing
            End If
    End Select
End Sub

Public Sub Party_GetLoot(ByVal PartyNum As Long, ByVal ItemNum As Long, ByVal ItemValue As Long, X As Byte, Y As Byte)
    Dim i As Long, tmpIndex As Long, foundMember As Boolean
    
    For i = 1 To MAX_PARTY_MEMBERS
        tmpIndex = Party(PartyNum).Member(Random(1, Party(PartyNum).MemberCount))
        
        If IsPlaying(tmpIndex) Then
            foundMember = True
            Exit For
        End If
    Next
    
    ' Prevent subscript out of range
    If foundMember = False Then Exit Sub
    
    If Moral(GetPlayerMap(tmpIndex)).CanDropItem = 1 Then
        Call SpawnItem(ItemNum, ItemValue, Item(ItemNum).Data1, GetPlayerMap(tmpIndex), X, Y, GetPlayerName(tmpIndex))
    Else
        GiveInvItem tmpIndex, ItemNum, ItemValue, Item(ItemNum).Data1
    End If
End Sub

Public Function IsPlayerBlocked(Index As Long, ByVal X As Long, ByVal Y As Long)
    Dim i As Long
    
    ' Does the map block players?
    If Moral(Map(GetPlayerMap(Index)).Moral).PlayerBlocked = 1 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) And Not i = Index Then
                If GetPlayerMap(i) = GetPlayerMap(Index) Then
                    If (X > 0 And GetPlayerX(i) = GetPlayerX(Index) + X) And GetPlayerY(Index) = GetPlayerY(i) Then
                        IsPlayerBlocked = True
                        Exit For
                    ElseIf (Y > 0 And GetPlayerY(i) = GetPlayerX(Index) + Y) And GetPlayerX(Index) = GetPlayerX(i) Then
                        IsPlayerBlocked = True
                        Exit For
                    End If
                End If
            End If
        Next
    End If
End Function

Public Function IsEventBlocked(Index As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal MapNum As Integer)
    Dim i As Long
    
    If MapNum = 0 Then
        MapNum = GetPlayerMap(Index)
        X = X + GetPlayerX(Index)
        Y = Y + GetPlayerY(Index)
        
        ' Check to see if a player event is already on that tile
        For i = 1 To TempEventMap(MapNum).EventCount
            If tempplayer(Index).EventMap.EventPages(i).X = X Then
                If tempplayer(Index).EventMap.EventPages(i).Y = Y Then
                    If tempplayer(Index).EventMap.EventPages(i).WalkThrough = 0 Then
                        IsEventBlocked = True
                        Exit Function
                    End If
                End If
            End If
        Next
    Else
        X = X + MapNPC(MapNum).NPC(Index).X
        Y = Y + MapNPC(MapNum).NPC(Index).Y
    End If
    
    ' Check to see if a global event is already on that tile
    For i = 1 To TempEventMap(MapNum).EventCount
        If TempEventMap(MapNum).Events(i).X = X Then
            If TempEventMap(MapNum).Events(i).Y = Y Then
                If TempEventMap(MapNum).Events(i).WalkThrough = 0 Then
                    IsEventBlocked = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function
