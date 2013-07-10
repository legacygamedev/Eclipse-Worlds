Attribute VB_Name = "modServerTCP"
Option Explicit

Sub UpdateCaption()
    If Not TotalOnlinePlayers = 1 Then
        frmServer.Caption = Options.Name & " (" & TotalOnlinePlayers & " Players Online)"
    Else
        frmServer.Caption = Options.Name & " (" & TotalOnlinePlayers & " Player Online)"
    End If
End Sub

Sub CreateFullMapCache()
    Dim i As Long

    For i = 1 To MAX_MAPS
        Call MapCache_Create(i)
    Next
End Sub

Function IsConnected(ByVal index As Long) As Boolean
    If frmServer.Socket(index).State = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal index As Long) As Boolean
    If IsConnected(index) Then
        If TempPlayer(index).InGame Then
            IsPlaying = True
        End If
    End If
End Function

Function IsLoggedIn(ByVal index As Integer) As Boolean
    If IsConnected(index) Then
        If Len(Trim$(Account(index).Login)) > 0 Then
            IsLoggedIn = True
        End If
    End If
End Function

Function IsMultiAccounts(ByVal Login As String) As Boolean
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsConnected(i) Then
            If LCase$(Trim$(Account(i).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If
    Next
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim i As Long
    Dim n As Long

    For i = 1 To Player_HighIndex
        If IsConnected(i) Then
            If Trim$(GetPlayerIP(i)) = IP Then
                n = n + 1

                If (n > 1) Then
                    IsMultiIPOnline = True
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Public Function IsBanned(ByVal index As Long, Serial As String) As Boolean
    Dim SendIP As String
    Dim n As Long

    IsBanned = False

    ' Cut off last portion of IP
    SendIP = GetPlayerIP(index)
    For n = Len(SendIP) To 1 Step -1
        If Mid$(SendIP, n, 1) = "." Then Exit For
    Next n
    
    SendIP = Mid$(SendIP, 1, n)

    For n = 1 To MAX_BANS
        If Len(GetPlayerLogin(index)) > 0 Then
            If GetPlayerLogin(index) = Trim$(Ban(n).PlayerLogin) Then
                IsBanned = True
                Exit For
            End If
        End If

        If Len(Trim$(SendIP)) > 0 Then
            If Trim$(SendIP) = Left$(Trim$(Ban(n).IP), Len(SendIP)) Then
                IsBanned = True
                Exit For
            End If
        End If

        If Len(Serial) > 0 Then
            If Serial = Trim$(Ban(n).HDSerial) Then
                IsBanned = True
                Exit For
            End If
        Else
            IsBanned = True
            Exit For
        End If
    Next n
    
    If IsBanned = True Then
        Call AlertMsg(index, "You are banned from " & Options.Name & " and can no longer play.")
    End If
End Function

Sub SendDataTo(ByVal index As Long, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Dim TempData() As Byte
    
    If IsConnected(index) Then
        Set buffer = New clsBuffer
        TempData = Data
        
        buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteBytes TempData()
        
        If IsConnected(index) Then
            frmServer.Socket(index).SendData buffer.ToArray()
        End If
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            Call SendDataTo(i, Data)
        End If
    Next
End Sub

Sub SendDataToAllBut(ByVal index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not i = index Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next
End Sub

Sub SendDataToMap(ByVal MapNum As Integer, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                Call SendDataTo(i, Data)
            End If
        End If
    Next
End Sub

Sub SendDataToMapBut(ByVal index As Long, ByVal MapNum As Integer, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If Not i = index Then
                    Call SendDataTo(i, Data)
                End If
            End If
        End If
    Next
End Sub

Sub SendDataToParty(ByVal PartyNum As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Party(PartyNum).MemberCount
        If Party(PartyNum).Member(i) > 0 Then
            Call SendDataTo(Party(PartyNum).Member(i), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SGlobalMsg
    buffer.WriteString Msg
    buffer.WriteLong Color
    SendDataToAll buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, Color As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim LogMsg As String
    
    Set buffer = New clsBuffer
    
    ' Add server log
    Call AddLog(Msg, "Player")
    
    LogMsg = Msg
    Msg = "[Admin] " & Msg
    
    buffer.WriteLong SAdminMsg
    buffer.WriteString Msg
    buffer.WriteByte Color
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) >= STAFF_MODERATOR Then
            SendDataTo i, buffer.ToArray
            Call SendLogs(i, LogMsg, "Admin")
        End If
    Next
    
    Set buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerMsg
    buffer.WriteString Msg
    buffer.WriteByte Color
    SendDataTo index, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub MapMsg(ByVal MapNum As Integer, ByVal Msg As String, ByVal Color As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    Set buffer = New clsBuffer

    buffer.WriteLong SMapMsg
    buffer.WriteString Msg
    buffer.WriteByte Color
    SendDataToMap MapNum, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Public Sub GuildMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long, Optional HideName As Boolean = False)
    Dim i As Long
    Dim LogMsg As String
    
    ' Add server log
    Call AddLog(Msg, "Player")
    
    ' Set the LogMsg
    If HideName = True Then
        LogMsg = Msg
        Msg = "[Guild] " & Msg
    Else
        LogMsg = GetPlayerName(index) & ": " & Msg
        Msg = "[Guild] " & GetPlayerName(index) & ": " & Msg
    End If

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerGuild(i) = GetPlayerGuild(index) Then
                PlayerMsg i, Msg, Color
                Call SendLogs(i, LogMsg, "Guild")
            End If
        End If
    Next
End Sub

Public Sub AlertMsg(ByVal index As Long, ByVal Msg As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong SAlertMsg
    buffer.WriteString Msg
    SendDataTo index, buffer.ToArray
    DoEvents
    
    Set buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal PartyNum As Long, ByVal Msg As String, ByVal Color As Long)
    Dim i As Long
    Dim LogMsg As String
    
    ' Add server log
    Call AddLog(Msg, "Player")
    
    LogMsg = Msg
    
    Msg = "[Party] " & Msg
    
    ' Send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' Exist?
        If Party(PartyNum).Member(i) > 0 Then
            ' Make sure they're logged on
            If IsConnected(Party(PartyNum).Member(i)) And IsPlaying(Party(PartyNum).Member(i)) Then
                PlayerMsg Party(PartyNum).Member(i), Msg, Color
                Call SendLogs(Party(PartyNum).Member(i), LogMsg, "Party")
            End If
        End If
    Next
End Sub

Sub AcceptConnection(ByVal index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (index = 0) Then
        i = FindOpenPlayerSlot

        If Not i = 0 Then
            If Not IsConnected(i) Then
                ' We can connect them
                frmServer.Socket(i).Close
                frmServer.Socket(i).Accept SocketId
                Call SocketConnected(i)
            End If
        End If
    End If
End Sub

Sub SocketConnected(ByVal index As Long)
    Dim i As Long
    
    If index > 0 And index <= MAX_PLAYERS Then
        ' Are they trying to connect more then one connection?
        If Not IsMultiIPOnline(GetPlayerIP(index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(index) & ".")
        ElseIf Options.MultipleIP = 0 And IsMultiIPOnline(GetPlayerIP(index)) Then
            ' Tried Multiple connections
            Call AlertMsg(index, "Multiple account logins are not authorized.")
            frmServer.Socket(index).Close
        End If
    Else
        Call AlertMsg(index, "The server is full! Try back again later.")
        frmServer.Socket(index).Close
    End If

    ' Re-set the high Index
    Player_HighIndex = 0
    
    For i = MAX_PLAYERS To 1 Step -1
        If IsConnected(i) Then
            Player_HighIndex = i
            Exit For
        End If
    Next
    
    ' Send the new highIndex to all logged in players
    SendPlayer_HighIndex
End Sub

Sub IncomingData(ByVal index As Long, ByVal DataLength As Long)
    Dim buffer() As Byte
    Dim pLength As Long

    If GetPlayerAccess(index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(index).DataBytes > 1000 Then Exit Sub
    
        ' Check for Packet flooding
        If TempPlayer(index).DataPackets > 105 Then Exit Sub
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(index).DataBytes = TempPlayer(index).DataBytes + DataLength
    
    If timeGetTime >= TempPlayer(index).DataTimer Then
        TempPlayer(index).DataTimer = timeGetTime + 1000
        TempPlayer(index).DataBytes = 0
        TempPlayer(index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(index).GetData buffer(), vbUnicode, DataLength
    TempPlayer(index).buffer.WriteBytes buffer()
    
    If TempPlayer(index).buffer.Length >= 4 Then
        pLength = TempPlayer(index).buffer.ReadLong(False)
    
        If pLength < 0 Then Exit Sub
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).buffer.Length - 4
        If pLength <= TempPlayer(index).buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).buffer.ReadLong
            HandleData index, TempPlayer(index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        
        If TempPlayer(index).buffer.Length >= 4 Then
            pLength = TempPlayer(index).buffer.ReadLong(False)
        
            If pLength < 0 Then Exit Sub
        End If
    Loop
            
    TempPlayer(index).buffer.Trim
End Sub

Sub CloseSocket(ByVal index As Long, Optional ByVal NoMessage As Boolean = False)
    Dim i As Long

    If index > 0 And index <= MAX_PLAYERS Then
        Call LeftGame(index)
        
        If NoMessage = False Then
            Call TextAdd("Connection from " & GetPlayerIP(index) & " has been terminated.")
        End If

        frmServer.Socket(index).Close
        Call UpdateCaption
    End If
    
    ' Re-set the high Index
    Player_HighIndex = 0
    
    ' Set the new high index
    For i = MAX_PLAYERS To 1 Step -1
        If IsConnected(i) Then
            Player_HighIndex = i
            Exit For
        End If
    Next
    
    ' Send the new highIndex to all logged in players
    SendPlayer_HighIndex
End Sub

Public Sub MapCache_Create(ByVal MapNum As Integer)
    Dim MapData As String
    Dim X As Long
    Dim Y As Long
    Dim i As Long, z As Long, w As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong MapNum
    buffer.WriteString Trim$(Map(MapNum).Name)
    buffer.WriteString Trim$(Map(MapNum).Music)
    buffer.WriteString Trim$(Map(MapNum).BGS)
    buffer.WriteLong Map(MapNum).Revision
    buffer.WriteByte Map(MapNum).Moral
    buffer.WriteLong Map(MapNum).Up
    buffer.WriteLong Map(MapNum).Down
    buffer.WriteLong Map(MapNum).Left
    buffer.WriteLong Map(MapNum).Right
    buffer.WriteLong Map(MapNum).BootMap
    buffer.WriteByte Map(MapNum).BootX
    buffer.WriteByte Map(MapNum).BootY
    
    buffer.WriteLong Map(MapNum).Weather
    buffer.WriteLong Map(MapNum).WeatherIntensity
    
    buffer.WriteLong Map(MapNum).Fog
    buffer.WriteLong Map(MapNum).FogSpeed
    buffer.WriteLong Map(MapNum).FogOpacity
    
    buffer.WriteLong Map(MapNum).Panorama
    
    buffer.WriteLong Map(MapNum).Red
    buffer.WriteLong Map(MapNum).Green
    buffer.WriteLong Map(MapNum).Blue
    buffer.WriteLong Map(MapNum).Alpha
    
    buffer.WriteByte Map(MapNum).MaxX
    buffer.WriteByte Map(MapNum).MaxY
    
    buffer.WriteByte Map(MapNum).Npc_HighIndex

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            With Map(MapNum).Tile(X, Y)
                For i = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Layer(i).X
                    buffer.WriteLong .Layer(i).Y
                    buffer.WriteLong .Layer(i).Tileset
                Next
                
                For z = 1 To MapLayer.Layer_Count - 1
                    buffer.WriteLong .Autotile(z)
                Next
                
                buffer.WriteByte .Type
                buffer.WriteLong .Data1
                buffer.WriteLong .Data2
                buffer.WriteLong .Data3
                buffer.WriteString .Data4
                buffer.WriteByte .DirBlock
            End With
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        buffer.WriteLong Map(MapNum).NPC(X)
        buffer.WriteLong Map(MapNum).NpcSpawnType(X)
    Next

    MapCache(MapNum).Data = buffer.ToArray()
    Set buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not i = index Then
                s = s & GetPlayerName(i) & ", "
                n = n + 1
            End If
        End If
    Next

    If n = 0 Then
        s = "There are no other players online."
    ElseIf n = 1 Then
        s = Mid$(s, 1, Len(s) - 2)
        s = "There is " & n & " other player online: " & s & "."
    Else
        s = Mid$(s, 1, Len(s) - 2)
        s = "There are " & n & " other players online: " & s & "."
    End If

    Call PlayerMsg(index, s, WhoColor)
End Sub

Function PlayerData(ByVal index As Long) As Byte()
    Dim buffer As clsBuffer, i As Long

    If index > Player_HighIndex Or index < 1 Then Exit Function
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerData
    buffer.WriteLong index
    buffer.WriteInteger Account(index).Chars(GetPlayerChar(index)).Face
    buffer.WriteString GetPlayerName(index)
    buffer.WriteByte GetPlayerGender(index)
    buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Class
    buffer.WriteInteger GetPlayerLevel(index)
    buffer.WriteInteger GetPlayerPoints(index)
    buffer.WriteInteger GetPlayerSprite(index)
    buffer.WriteInteger GetPlayerMap(index)
    buffer.WriteByte GetPlayerX(index)
    buffer.WriteByte GetPlayerY(index)
    buffer.WriteByte GetPlayerDir(index)
    buffer.WriteByte GetPlayerAccess(index)
    buffer.WriteByte GetPlayerPK(index)
    
    If GetPlayerGuild(index) > 0 Then
        buffer.WriteString Guild(GetPlayerGuild(index)).Name
    Else
        buffer.WriteString vbNullString
    End If
    
    buffer.WriteByte GetPlayerGuildAccess(index)
    
    For i = 1 To Stats.Stat_count - 1
        buffer.WriteInteger GetPlayerStat(index, i)
    Next
    
    ' Amount of titles
    buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
    
    ' Send player titles
    For i = 1 To Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
        buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Title(i)
    Next
    
    ' Send the player's current title
    buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).CurrentTitle
    
    ' Send player status
    buffer.WriteString Account(index).Chars(GetPlayerChar(index)).Status
    
    PlayerData = buffer.ToArray()
    Set buffer = Nothing
End Function

Sub SendJoinMap(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    ' Send all players on current map to index
    If GetTotalMapPlayers(GetPlayerMap(index)) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                SendDataToMap GetPlayerMap(index), PlayerData(i)
                Call SendPlayerEquipmentToMapBut(i)
            End If
        Next
    End If
    
    ' Send Index's player data to everyone on the map including themself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    ' Send player's equipment to new map
    SendPlayerEquipmentTo index
    
    ' Send the npc targets to the player
    For i = 1 To Map(GetPlayerMap(index)).Npc_HighIndex
        If MapNpc(GetPlayerMap(index)).NPC(i).Num > 0 Then
            Call SendMapNpcTarget(GetPlayerMap(index), i, MapNpc(GetPlayerMap(index)).NPC(i).Target, MapNpc(GetPlayerMap(index)).NPC(i).TargetType)
        Else
            ' Send 0 so it uncaches any old data
            Call SendMapNpcTarget(GetPlayerMap(index), i, 0, 0)
        End If
    Next
    
    Set buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Integer)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SLeft
    buffer.WriteLong index
    SendDataToMapBut index, MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendPlayerData(ByVal index As Long)
    SendDataToMap GetPlayerMap(index), PlayerData(index)
End Sub

Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    buffer.WriteLong SMapData
    buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Integer)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData
    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteString MapItem(MapNum, i).PlayerName
        buffer.WriteLong MapItem(MapNum, i).Num
        buffer.WriteLong MapItem(MapNum, i).Value
        buffer.WriteInteger MapItem(MapNum, i).Durability
        buffer.WriteByte MapItem(MapNum, i).X
        buffer.WriteByte MapItem(MapNum, i).Y
        buffer.WriteInteger MapItem(MapNum, i).YOffset
    Next

    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemToMap(ByVal MapNum As Integer, ByVal MapSlotNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SSpawnItem
    buffer.WriteLong MapSlotNum
    buffer.WriteString MapItem(MapNum, MapSlotNum).PlayerName
    buffer.WriteLong MapItem(MapNum, MapSlotNum).Num
    buffer.WriteLong MapItem(MapNum, MapSlotNum).Value
    buffer.WriteInteger MapItem(MapNum, MapSlotNum).Durability
    buffer.WriteLong MapItem(MapNum, MapSlotNum).X
    buffer.WriteLong MapItem(MapNum, MapSlotNum).Y
    buffer.WriteInteger MapItem(MapNum, MapSlotNum).YOffset
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapNpcVitals(ByVal MapNum As Integer, ByVal MapNpcNum As Byte)
    Dim buffer As clsBuffer, i As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNPCVitals
    buffer.WriteByte MapNpcNum
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong MapNpc(MapNum).NPC(MapNpcNum).Vital(i)
    Next

    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapNpcTarget(ByVal MapNum As Integer, ByVal MapNpcNum As Byte, ByVal Target As Byte, ByVal TargetType As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNPCTarget
    buffer.WriteByte MapNpcNum
    buffer.WriteByte Target
    buffer.WriteByte TargetType

    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapNpcsTo(ByVal index As Long, ByVal MapNum As Integer)
    Dim i As Long, X As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNPCData

    For i = 1 To MAX_MAP_NPCS
        With MapNpc(MapNum).NPC(i)
            buffer.WriteLong .Num
            buffer.WriteLong .X
            buffer.WriteLong .Y
            buffer.WriteLong .Dir
            For X = 1 To Vitals.Vital_Count - 1
                buffer.WriteLong .Vital(X)
            Next
        End With
    Next

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapNpcsToMap(ByVal MapNum As Integer)
    Dim i As Long, X As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapNPCData

    For i = 1 To MAX_MAP_NPCS
        With MapNpc(MapNum).NPC(i)
            buffer.WriteLong .Num
            buffer.WriteLong .X
            buffer.WriteLong .Y
            buffer.WriteLong .Dir
            For X = 1 To Vitals.Vital_Count - 1
                buffer.WriteLong .Vital(X)
            Next
        End With
    Next

    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMorals(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_MORALS
        If Len(Trim$(Moral(i).Name)) > 0 Then
            Call SendUpdateMoralTo(index, i)
        End If
    Next
End Sub

Sub SendClasses(ByVal index As Long)
    Dim i As Long
    
    For i = 1 To MAX_CLASSES
        If Len(Trim$(Class(i).Name)) > 0 Then
            Call SendUpdateClassTo(index, i)
        End If
    Next
End Sub

Sub SendEmoticons(ByVal index As Long)
    Dim i As Long
    
    For i = 1 To MAX_EMOTICONS
        If Len(Trim$(Emoticon(i).Command)) > 0 Then
            Call SendUpdateEmoticonTo(index, i)
        End If
    Next
End Sub

Sub SendItems(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS
        If Len(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(index, i)
        End If
    Next
End Sub

Sub SendTitles(ByVal index As Long)
    Dim i As Long
    
    For i = 1 To MAX_TITLES
        If Len(Trim$(Title(i).Name)) > 0 Then
            Call SendUpdateTitleTo(index, i)
        End If
    Next
End Sub

Sub SendAnimations(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        If Len(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(index, i)
        End If
    Next
End Sub

Sub SendNpcs(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS
        If Len(Trim$(NPC(i).Name)) > 0 Then
            Call SendUpdateNpcTo(index, i)
        End If
    Next
End Sub

Sub SendResources(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Len(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(index, i)
        End If
    Next
End Sub

Sub SendInventory(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        buffer.WriteLong GetPlayerInvItemNum(index, i)
        buffer.WriteLong GetPlayerInvItemValue(index, i)
        buffer.WriteInteger GetPlayerInvItemDur(index, i)
        buffer.WriteByte GetPlayerInvItemBind(index, i)
    Next

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInvUpdate
    buffer.WriteByte InvSlot
    buffer.WriteLong GetPlayerInvItemNum(index, InvSlot)
    buffer.WriteLong GetPlayerInvItemValue(index, InvSlot)
    buffer.WriteInteger GetPlayerInvItemDur(index, InvSlot)
    buffer.WriteByte GetPlayerInvItemBind(index, InvSlot)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerEquipmentTo(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim i As Byte
    
    buffer.WriteLong SPlayerWornEq
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteLong GetPlayerEquipment(index, i)
    Next
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteInteger GetPlayerEquipmentDur(index, i)
    Next
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerEquipmentToMap(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim i As Byte
    
    buffer.WriteLong SMapWornEq
    buffer.WriteLong index
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteLong GetPlayerEquipment(index, i)
    Next

    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerEquipmentToMapBut(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Byte
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapWornEq
    buffer.WriteLong index
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteLong GetPlayerEquipment(index, i)
    Next
    
    SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    Select Case Vital
        Case HP
            buffer.WriteLong SPlayerHP
            buffer.WriteLong index
            buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case MP
            buffer.WriteLong SPlayerMP
            buffer.WriteLong index
            buffer.WriteLong GetPlayerMaxVital(index, Vitals.MP)
            buffer.WriteLong GetPlayerVital(index, Vitals.MP)
    End Select

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendVitalTo(ByVal index As Long, Player As Long, ByVal Vital As Vitals)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    Select Case Vital
        Case HP
            buffer.WriteLong SPlayerHP
            buffer.WriteLong Player
            buffer.WriteLong GetPlayerMaxVital(Player, Vitals.HP)
            buffer.WriteLong GetPlayerVital(Player, Vitals.HP)
        Case MP
            buffer.WriteLong SPlayerMP
            buffer.WriteLong Player
            buffer.WriteLong GetPlayerMaxVital(Player, Vitals.MP)
            buffer.WriteLong GetPlayerVital(Player, Vitals.MP)
    End Select

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerExp(ByVal index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerEXP
    buffer.WriteLong index
    buffer.WriteLong GetPlayerExp(index)
    buffer.WriteLong GetPlayerNextLevel(index)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerStats(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerStats
    buffer.WriteLong index
    
    For i = 1 To Stats.Stat_count - 1
        buffer.WriteInteger GetPlayerStat(index, i)
    Next
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerPoints(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerPoints
    buffer.WriteLong index
    buffer.WriteInteger GetPlayerPoints(index)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerLevel(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerLevel
    buffer.WriteLong index
    buffer.WriteInteger GetPlayerLevel(index)
    
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerGuild(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerGuild
    buffer.WriteLong index
    
    If GetPlayerGuild(index) > 0 Then
        buffer.WriteString Guild(GetPlayerGuild(index)).Name
    Else
        buffer.WriteString vbNullString
    End If
    
    buffer.WriteByte GetPlayerGuildAccess(index)
    
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendGuildInvite(ByVal index As Long, ByVal OtherPlayer As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SGuildInvite
    
    buffer.WriteString Trim$(Account(index).Chars(GetPlayerChar(index)).Name)
    buffer.WriteString Trim$(Guild(GetPlayerGuild(index)).Name)
    
    SendDataTo OtherPlayer, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerGuildMembers(ByVal index As Long, Optional ByVal Ignore As Byte = 0)
    Dim i As Long
    Dim PlayerArray() As String
    Dim PlayerCount As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SGuildMembers
    
    PlayerCount = 0
    
    ' Count members online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not Ignore = i And Not i = index Then
                If GetPlayerGuild(i) = GetPlayerGuild(index) Then
                    PlayerCount = PlayerCount + 1
                    ReDim Preserve PlayerArray(1 To PlayerCount)
                    PlayerArray(UBound(PlayerArray)) = GetPlayerName(i)
                End If
            End If
        End If
    Next
    
    ' Add to Packet
    buffer.WriteLong PlayerCount
    
    If PlayerCount > 0 Then
        For i = 1 To PlayerCount
            buffer.WriteString PlayerArray(i)
        Next
    End If
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub
Sub SendPlayerSprite(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerSprite
    buffer.WriteLong index
    buffer.WriteInteger GetPlayerSprite(index)
      
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerTitles(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerTitles
    buffer.WriteLong index
    
    ' Amount of titles
    buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
    
    ' Send player titles
    For i = 1 To Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
        buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Title(i)
    Next
    
    ' Send the player's current title
    buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).CurrentTitle
    
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerStatus(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerStatus
    buffer.WriteLong index
    
    buffer.WriteString Account(index).Chars(GetPlayerChar(index)).Status
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerPK(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerPK
    buffer.WriteLong index
    
    buffer.WriteByte GetPlayerPK(index)
    
    SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendWelcome(ByVal index As Long)
    ' Send the MOTD
    If Not Trim$(Options.MOTD) = vbNullString Then
        Call PlayerMsg(index, Options.MOTD, BrightCyan)
    End If
    
    ' Send the SMOTD
    If Not Trim$(Options.SMOTD) = vbNullString Then
        If GetPlayerAccess(index) >= STAFF_MODERATOR Then
            Call PlayerMsg(index, Options.SMOTD, Cyan)
        End If
    End If
    
    ' Send the GMOTD
    If GetPlayerGuild(index) > 0 Then
        If Not Trim$(Guild(GetPlayerGuild(index)).MOTD) = vbNullString Then
            Call PlayerMsg(index, Trim$(Guild(GetPlayerGuild(index)).MOTD), BrightGreen)
        End If
    End If
End Sub

Sub SendLeftGame(ByVal index As Long)
    Dim buffer As clsBuffer, i As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SLeaveGame
    buffer.WriteLong index
    
    SendDataToAllBut index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Integer)
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    buffer.WriteLong SUpdateItem
    buffer.WriteLong ItemNum
    buffer.WriteBytes ItemData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Integer)
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    
    Set buffer = New clsBuffer
    
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    buffer.WriteLong SUpdateItem
    buffer.WriteLong ItemNum
    buffer.WriteBytes ItemData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    
    Set buffer = New clsBuffer
    
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal index As Long, ByVal AnimationNum As Long)
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    
    Set buffer = New clsBuffer
    
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    buffer.WriteLong SUpdateAnimation
    buffer.WriteLong AnimationNum
    buffer.WriteBytes AnimationData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcToAll(ByVal npcnum As Long)
    Dim buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    
    Set buffer = New clsBuffer
    
    NpcSize = LenB(NPC(npcnum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(NPC(npcnum)), NpcSize
    buffer.WriteLong SUpdateNPC
    buffer.WriteLong npcnum
    buffer.WriteBytes NpcData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateNpcTo(ByVal index As Long, ByVal npcnum As Long)
    Dim buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    
    Set buffer = New clsBuffer
    
    NpcSize = LenB(NPC(npcnum))
    ReDim NpcData(NpcSize - 1)
    CopyMemory NpcData(0), ByVal VarPtr(NPC(npcnum)), NpcSize
    buffer.WriteLong SUpdateNPC
    buffer.WriteLong npcnum
    buffer.WriteBytes NpcData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal index As Long, ByVal ResourceNum As Long)
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    buffer.WriteLong SUpdateResource
    buffer.WriteLong ResourceNum
    buffer.WriteBytes ResourceData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendShops(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS
        If Len(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(index, i)
        End If
    Next
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong ShopNum
    buffer.WriteBytes ShopData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum As Long)
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    
    buffer.WriteLong SUpdateShop
    buffer.WriteLong ShopNum
    buffer.WriteBytes ShopData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpells(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS
        If Len(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(index, i)
        End If
    Next
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong SpellNum
    buffer.WriteBytes SpellData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    buffer.WriteLong SUpdateSpell
    buffer.WriteLong SpellNum
    buffer.WriteBytes SpellData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SSpells
    For i = 1 To MAX_PLAYER_SPELLS
        buffer.WriteLong GetPlayerSpell(index, i)
    Next
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSpell(ByVal index As Long, ByVal SpellSlot As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SSpell
    buffer.WriteByte SpellSlot
    buffer.WriteLong GetPlayerSpell(index, SpellSlot)
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal index As Long, ByVal Resource_Num As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(i).ResourceState
            buffer.WriteInteger ResourceCache(GetPlayerMap(index)).ResourceData(i).X
            buffer.WriteInteger ResourceCache(GetPlayerMap(index)).ResourceData(i).Y
        Next
    End If

    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal MapNum As Integer, ByVal Resource_Num As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(MapNum).Resource_Count

    If ResourceCache(MapNum).Resource_Count > 0 Then
        For i = 0 To ResourceCache(MapNum).Resource_Count
            buffer.WriteByte ResourceCache(MapNum).ResourceData(i).ResourceState
            buffer.WriteInteger ResourceCache(MapNum).ResourceData(i).X
            buffer.WriteInteger ResourceCache(MapNum).ResourceData(i).Y
        Next
    End If

    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal MapNum As Integer, ByVal X As Long, ByVal Y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SDoorAnimation
    buffer.WriteLong X
    buffer.WriteLong Y
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendActionMsg(ByVal MapNum As Integer, ByVal Message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal X As Long, ByVal Y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SActionMsg
    buffer.WriteString Message
    buffer.WriteLong Color
    buffer.WriteLong MsgType
    buffer.WriteLong X
    buffer.WriteLong Y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, buffer.ToArray()
    Else
        SendDataToMap MapNum, buffer.ToArray()
    End If
    
    Set buffer = Nothing
End Sub

Sub SendBlood(ByVal MapNum As Integer, ByVal X As Long, ByVal Y As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SBlood
    buffer.WriteLong X
    buffer.WriteLong Y
    
    SendDataToMap MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendAnimation(ByVal MapNum As Integer, ByVal Anim As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0, Optional ByVal OnlyTo As Long = 0)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAnimation
    buffer.WriteLong Anim
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteByte LockType
    buffer.WriteLong LockIndex
    
    If OnlyTo > 0 Then
        SendDataTo OnlyTo, buffer.ToArray
    Else
        SendDataToMap MapNum, buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Sub SendSpellCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSpellCooldown
    buffer.WriteByte Slot
    buffer.WriteLong Account(index).Chars(GetPlayerChar(index)).SpellCD(Slot)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendClearAccountSpellBuffer(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal MapNum As Integer, ByVal index As Long, ByVal Message As String, ByVal SayColor As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString Message
    buffer.WriteString "[Map] "
    buffer.WriteLong SayColor
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal index As Long, ByVal Message As String, ByVal SayColor As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    
    buffer.WriteString GetPlayerName(index)
    buffer.WriteLong GetPlayerAccess(index)
    buffer.WriteLong GetPlayerPK(index)
    buffer.WriteString Message
    buffer.WriteString "[Global] "
    buffer.WriteLong SayColor
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub ResetShopAction(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SResetShopAction
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendStunned(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStunned
    
    buffer.WriteLong TempPlayer(index).StunDuration
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendBank(ByVal index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong sbank
    
    For i = 1 To MAX_BANK
        buffer.WriteLong Account(index).Bank.Item(i).Num
        buffer.WriteLong Account(index).Bank.Item(i).Value
    Next
    
    SendDataTo index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapKey(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapKey
    
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteByte Value
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapKeyToMap(ByVal MapNum As Integer, ByVal X As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapKey
    
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteByte Value
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal ShopNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SOpenShop
    
    buffer.WriteLong ShopNum
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal Movement As Byte, Optional ByVal SendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMove
    
    buffer.WriteLong index
    buffer.WriteByte GetPlayerX(index)
    buffer.WriteByte GetPlayerY(index)
    buffer.WriteByte GetPlayerDir(index)
    buffer.WriteByte Movement
    
    If Not SendToSelf Then
        SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Else
        SendDataTo index, buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Sub SendPlayerWarp(ByVal index As Long, Optional ByVal SendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerWarp
    
    buffer.WriteLong index
    buffer.WriteByte GetPlayerX(index)
    buffer.WriteByte GetPlayerY(index)
    buffer.WriteByte GetPlayerDir(index)
    
    If Not SendToSelf Then
        SendDataToMap GetPlayerMap(index), buffer.ToArray()
    Else
        SendDataTo index, buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Sub SendNpcMove(ByVal MapNpcNum As Long, ByVal Movement As Byte, MapNum As Integer)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNPCMove
    
    buffer.WriteLong MapNpcNum
    buffer.WriteByte MapNpc(MapNum).NPC(MapNpcNum).X
    buffer.WriteByte MapNpc(MapNum).NPC(MapNpcNum).Y
    buffer.WriteByte MapNpc(MapNum).NPC(MapNpcNum).Dir
    buffer.WriteByte Movement
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTrade(ByVal index As Long, ByVal TradeTarget As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STrade
    
    buffer.WriteLong TradeTarget
    buffer.WriteString Trim$(GetPlayerName(TradeTarget))
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCloseTrade
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal index As Long, ByVal DataType As Byte)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim TradeTarget As Long
    Dim TotalWorth As Long
    
    TradeTarget = TempPlayer(index).InTrade
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeUpdate
    
    buffer.WriteByte DataType
    If DataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            buffer.WriteLong TempPlayer(index).TradeOffer(i).Num
            buffer.WriteLong TempPlayer(index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(index).TradeOffer(i).Num).Type = ITEM_TYPE_CURRENCY Then
                    TotalWorth = TotalWorth + (Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).Price * TempPlayer(index).TradeOffer(i).Value)
                Else
                    TotalWorth = TotalWorth + Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    ElseIf DataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            buffer.WriteLong GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num)
            buffer.WriteLong TempPlayer(TradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num)).Type = ITEM_TYPE_CURRENCY Then
                    TotalWorth = TotalWorth + (Item(GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num)).Price * TempPlayer(TradeTarget).TradeOffer(i).Value)
                Else
                    TotalWorth = TotalWorth + Item(GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    End If
    ' Send total worth of trade
    buffer.WriteLong TotalWorth
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal index As Long, ByVal Status As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeStatus
    
    buffer.WriteByte Status
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendAttack(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAttack
    buffer.WriteLong index
    
    SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendChatUpdate(ByVal index As Long, ByVal npcnum As Long, ByVal mT As String, ByVal o1 As String, ByVal o2 As String, ByVal o3 As String, ByVal o4 As String)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SChatUpdate
    
    buffer.WriteLong npcnum
    buffer.WriteString mT
    buffer.WriteString o1
    buffer.WriteString o2
    buffer.WriteString o3
    buffer.WriteString o4
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerTarget(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STarget
    
    buffer.WriteByte TempPlayer(index).Target
    buffer.WriteByte TempPlayer(index).TargetType
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendHotbar(ByVal index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SHotbar
    
    For i = 1 To MAX_HOTBAR
        buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Hotbar(i).Slot
        buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Hotbar(i).SType
    Next
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNpcSpellBuffer(MapNum, MapNpcNum)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNPCSpellBuffer
    
    buffer.WriteLong MapNpcNum
    buffer.WriteLong MapNpc(MapNum).NPC(MapNpcNum).SpellBuffer.Spell
    
    Call SendDataToMap(MapNum, buffer.ToArray)
    Set buffer = Nothing
End Sub

Sub SendLogs(ByVal index As Long, Msg As String, Name As String)
    Dim buffer As clsBuffer
    Dim LogSize As Long
    Dim LogData() As Byte
    
    Set buffer = New clsBuffer
    
    Log.Msg = Msg
    Log.File = Name
    LogSize = LenB(Log)
    ReDim LogData(LogSize - 1)
    CopyMemory LogData(0), ByVal VarPtr(Log), LogSize
    buffer.WriteLong SUpdateLogs
    
    buffer.WriteBytes LogData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub UpdateFriendsList(index)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim i As Long, n As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SFriendsList
    
    If Account(index).Friends.AmountOfFriends = 0 Then
        buffer.WriteByte Account(index).Friends.AmountOfFriends
        GoTo Finish
    End If
   
    ' Sends the amount of friends in friends list
    buffer.WriteByte Account(index).Friends.AmountOfFriends
   
    ' Check to see if they are online
    For i = 1 To Account(index).Friends.AmountOfFriends
        Name = Trim$(Account(index).Friends.Members(i))
        buffer.WriteString Name
        For n = 1 To Player_HighIndex
            If IsPlaying(FindPlayer(Name)) Then
                buffer.WriteString Name & " Online"
            Else
                buffer.WriteString Name & " Offline"
            End If
        Next
    Next
    
Finish:
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub UpdateFoesList(index)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim i As Long, n As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SFoesList
    
    If Account(index).Foes.Amount = 0 Then
        buffer.WriteByte Account(index).Foes.Amount
        GoTo Finish
    End If
   
    ' Sends the amount of Foes in Foes list
    buffer.WriteByte Account(index).Foes.Amount
   
    ' Check to see if they are online
    For i = 1 To Account(index).Foes.Amount
        Name = Trim$(Account(index).Foes.Members(i))
        buffer.WriteString Name
        For n = 1 To Player_HighIndex
            If IsPlaying(FindPlayer(Name)) Then
                buffer.WriteString Name & " Online"
            Else
                buffer.WriteString Name & " Offline"
            End If
        Next
    Next
    
Finish:
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendInGame(ByVal index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SInGame
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayer_HighIndex()
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SHighIndex
    
    buffer.WriteLong Player_HighIndex
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSoundTo(ByVal index As Integer, Sound As String)
    Dim buffer As clsBuffer
    
    ' Don't send it if there's nothing to send
    If Sound = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    
    buffer.WriteString Sound
    SendDataTo index, buffer.ToArray()
End Sub

Sub SendSoundToMap(ByVal MapNum As Integer, Sound As String)
    Dim buffer As clsBuffer
    
    ' Don't send it if there's nothing to send
    If Sound = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    
    buffer.WriteString Sound
    SendDataToMap MapNum, buffer.ToArray()
End Sub

Sub SendSoundToAll(ByVal MapNum As Integer, Sound As String)
    Dim buffer As clsBuffer
    
    ' Don't send it if there's nothing to send
    If Sound = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    
    buffer.WriteString Sound
    SendDataToAll buffer.ToArray()
End Sub

Sub SendPlayerSound(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal EntityType As Long, ByVal EntityNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SEntitySound
    
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteLong EntityType
    buffer.WriteLong EntityNum
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapSound(ByVal MapNum As Integer, ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal EntityType As Long, ByVal EntityNum As Long)
    Dim buffer As clsBuffer

    If EntityNum <= 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEntitySound
    
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteLong EntityType
    buffer.WriteLong EntityNum
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNpcDeath(ByVal MapNpcNum As Long, MapNum As Integer)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNPCDead
    
    buffer.WriteLong MapNpcNum
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNpcAttack(ByVal Attacker As Long, MapNum As Integer)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNPCAttack
    
    buffer.WriteLong Attacker
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendLogin(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SLoginOk
    
    buffer.WriteLong index
    buffer.WriteLong Player_HighIndex
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNewCharClasses
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNews(ByVal index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSendNews
    
    buffer.WriteString Options.News
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCheckForMap(ByVal index As Long, MapNum As Integer)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCheckForMap
    
    buffer.WriteInteger MapNum
    buffer.WriteInteger Map(MapNum).Revision
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STradeRequest
    
    buffer.WriteLong TradeRequest
    buffer.WriteString Trim$(Account(TradeRequest).Chars(GetPlayerChar(TradeRequest)).Name)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal index As Long, ByVal OtherPlayer As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyInvite
    
    buffer.WriteString Trim$(Account(OtherPlayer).Chars(GetPlayerChar(OtherPlayer)).Name)
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal PartyNum As Long)
    Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    
    buffer.WriteByte 1
    For i = 1 To MAX_PARTY_MEMBERS
        buffer.WriteLong Party(PartyNum).Member(i)
    Next
    buffer.WriteLong Party(PartyNum).MemberCount
    buffer.WriteLong PartyNum
    
    SendDataToParty PartyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal index As Long)
    Dim buffer As clsBuffer, i As Long, PartyNum As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    
    ' Check if we're in a party
    PartyNum = TempPlayer(index).InParty
    
    If PartyNum > 0 Then
        ' Send party data
        buffer.WriteByte 1
        buffer.WriteLong Party(PartyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            buffer.WriteLong Party(PartyNum).Member(i)
        Next
        buffer.WriteLong Party(PartyNum).MemberCount
    Else
        ' Send clear command
        buffer.WriteByte 0
    End If
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal PartyNum As Long, ByVal index As Long)
    Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyVitals
    
    buffer.WriteLong index
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong GetPlayerMaxVital(index, i)
        buffer.WriteLong Account(index).Chars(GetPlayerChar(index)).Vital(i)
    Next
    
    SendDataToParty PartyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateTitleToAll(ByVal TitleNum As Long)
    Dim buffer As clsBuffer
    Dim TitleSize As Long
    Dim TitleData() As Byte
    
    Set buffer = New clsBuffer
    
    TitleSize = LenB(Title(TitleNum))
    ReDim TitleData(TitleSize - 1)
    CopyMemory TitleData(0), ByVal VarPtr(Title(TitleNum)), TitleSize
    buffer.WriteLong SUpdateTitle
    
    buffer.WriteLong TitleNum
    buffer.WriteBytes TitleData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateTitleTo(ByVal index As Long, ByVal TitleNum As Long)
    Dim buffer As clsBuffer
    Dim TitleSize As Long
    Dim TitleData() As Byte
    
    Set buffer = New clsBuffer
    
    TitleSize = LenB(Title(TitleNum))
    ReDim TitleData(TitleSize - 1)
    CopyMemory TitleData(0), ByVal VarPtr(Title(TitleNum)), TitleSize
    buffer.WriteLong SUpdateTitle
    
    buffer.WriteLong TitleNum
    buffer.WriteBytes TitleData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCloseClient(index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SCloseClient
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateMoralToAll(ByVal MoralNum As Long)
    Dim buffer As clsBuffer
    Dim MoralSize As Long
    Dim MoralData() As Byte
    
    Set buffer = New clsBuffer
    
    MoralSize = LenB(Moral(MoralNum))
    ReDim MoralData(MoralSize - 1)
    CopyMemory MoralData(0), ByVal VarPtr(Moral(MoralNum)), MoralSize
    buffer.WriteLong SUpdateMoral
    
    buffer.WriteLong MoralNum
    buffer.WriteBytes MoralData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateMoralTo(ByVal index As Long, ByVal MoralNum As Long)
    Dim buffer As clsBuffer
    Dim MoralSize As Long
    Dim MoralData() As Byte
    
    Set buffer = New clsBuffer
    
    MoralSize = LenB(Moral(MoralNum))
    ReDim MoralData(MoralSize - 1)
    CopyMemory MoralData(0), ByVal VarPtr(Moral(MoralNum)), MoralSize
    buffer.WriteLong SUpdateMoral
    
    buffer.WriteLong MoralNum
    buffer.WriteBytes MoralData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateClassTo(ByVal index As Long, ByVal ClassNum As Long)
    Dim buffer As clsBuffer
    Dim Classesize As Long
    Dim ClassData() As Byte
    
    Set buffer = New clsBuffer
    
    Classesize = LenB(Class(ClassNum))
    ReDim ClassData(Classesize - 1)
    CopyMemory ClassData(0), ByVal VarPtr(Class(ClassNum)), Classesize
    buffer.WriteLong SUpdateClass
    
    buffer.WriteLong ClassNum
    buffer.WriteBytes ClassData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateEmoticonTo(ByVal index As Long, ByVal EmoticonNum As Long)
    Dim buffer As clsBuffer
    Dim EmoticonSize As Long
    Dim EmoticonData() As Byte
    
    Set buffer = New clsBuffer

    EmoticonSize = LenB(Emoticon(EmoticonNum))
    ReDim EmoticonData(EmoticonSize - 1)
    CopyMemory EmoticonData(0), ByVal VarPtr(Emoticon(EmoticonNum)), EmoticonSize
    buffer.WriteLong SUpdateEmoticon
    
    buffer.WriteLong EmoticonNum
    buffer.WriteBytes EmoticonData
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateEmoticonToAll(ByVal EmoticonNum As Integer)
    Dim buffer As clsBuffer
    Dim EmoticonSize As Long
    Dim EmoticonData() As Byte
    Set buffer = New clsBuffer
    EmoticonSize = LenB(Emoticon(EmoticonNum))
    
    ReDim EmoticonData(EmoticonSize - 1)
    
    CopyMemory EmoticonData(0), ByVal VarPtr(Emoticon(EmoticonNum)), EmoticonSize
    
    buffer.WriteLong SUpdateEmoticon
    buffer.WriteLong EmoticonNum
    buffer.WriteBytes EmoticonData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCheckEmoticon(ByVal index As Long, ByVal MapNum As Long, ByVal EmoticonNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SCheckEmoticon
    
    buffer.WriteLong index
    buffer.WriteLong EmoticonNum
    
    SendDataToMap MapNum, buffer.ToArray()
End Sub

Sub SendChatBubble(ByVal MapNum As Long, ByVal Target As Long, ByVal TargetType As Long, ByVal Message As String, ByVal Color As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SChatBubble
    
    buffer.WriteLong Target
    buffer.WriteLong TargetType
    buffer.WriteString Message
    buffer.WriteLong Color
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpecialEffect(ByVal index As Long, EffectType As Long, Optional Data1 As Long = 0, Optional Data2 As Long = 0, Optional Data3 As Long = 0, Optional Data4 As Long = 0)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SSpecialEffect
    
    Select Case EffectType
        Case EFFECT_TYPE_FADEIN
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FADEOUT
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FLASH
            buffer.WriteLong EffectType
        Case EFFECT_TYPE_FOG
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 ' Fog num
            buffer.WriteLong Data2 ' Fog movement speed
            buffer.WriteLong Data3 ' Opacity
        Case EFFECT_TYPE_WEATHER
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 ' Weather type
            buffer.WriteLong Data2 ' Weather intensity
        Case EFFECT_TYPE_TINT
            buffer.WriteLong EffectType
            buffer.WriteLong Data1 ' Red
            buffer.WriteLong Data2 ' Green
            buffer.WriteLong Data3 ' Blue
            buffer.WriteLong Data4 ' Alpha
    End Select
    
    SendDataTo index, buffer.ToArray
    Set buffer = Nothing
End Sub
