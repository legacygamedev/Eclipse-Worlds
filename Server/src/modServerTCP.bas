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

Function IsConnected(ByVal Index As Long) As Boolean
    If frmServer.Socket(Index).State = sckConnected Then
        IsConnected = True
    End If
End Function

Function IsPlaying(ByVal Index As Long) As Boolean
    If IsConnected(Index) Then
        If TempPlayer(Index).InGame Then
            IsPlaying = True
        End If
    End If
End Function

Function IsLoggedIn(ByVal Index As Integer) As Boolean
    If IsConnected(Index) Then
        If Len(Trim$(Account(Index).Login)) > 0 Then
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

Public Function IsBanned(ByVal Index As Long, Serial As String) As Boolean
    Dim SendIP As String
    Dim n As Long

    IsBanned = False

    ' Cut off last portion of IP
    SendIP = GetPlayerIP(Index)
    For n = Len(SendIP) To 1 Step -1
        If Mid$(SendIP, n, 1) = "." Then Exit For
    Next n
    
    SendIP = Mid$(SendIP, 1, n)

    For n = 1 To MAX_BANS
        If Len(GetPlayerLogin(Index)) > 0 Then
            If GetPlayerLogin(Index) = Trim$(Ban(n).PlayerLogin) Then
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
        Call AlertMsg(Index, "You are banned from " & Options.Name & " and can no longer play.")
    End If
End Function

Sub SendDataTo(ByVal Index As Long, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Dim TempData() As Byte
    
    If IsConnected(Index) Then
        Set buffer = New clsBuffer
        TempData = Data
        
        buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        buffer.WriteBytes TempData()
        
        If IsConnected(Index) Then
            frmServer.Socket(Index).SendData buffer.ToArray()
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

Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not i = Index Then
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

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Integer, ByRef Data() As Byte)
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = MapNum Then
                If Not i = Index Then
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

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerMsg
    buffer.WriteString Msg
    buffer.WriteByte Color
    SendDataTo Index, buffer.ToArray
    
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

Public Sub GuildMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Long, Optional HideName As Boolean = False)
    Dim i As Long
    Dim LogMsg As String
    
    ' Add server log
    Call AddLog(Msg, "Player")
    
    ' Set the LogMsg
    If HideName = True Then
        LogMsg = Msg
        Msg = "[Guild] " & Msg
    Else
        LogMsg = GetPlayerName(Index) & ": " & Msg
        Msg = "[Guild] " & GetPlayerName(Index) & ": " & Msg
    End If

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerGuild(i) = GetPlayerGuild(Index) Then
                PlayerMsg i, Msg, Color
                Call SendLogs(i, LogMsg, "Guild")
            End If
        End If
    Next
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    buffer.WriteLong SAlertMsg
    buffer.WriteString Msg
    SendDataTo Index, buffer.ToArray
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

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim i As Long

    If (Index = 0) Then
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

Sub SocketConnected(ByVal Index As Long)
    Dim i As Long
    
    If Index > 0 And Index <= MAX_PLAYERS Then
        ' Are they trying to connect more then one connection?
        If Not IsMultiIPOnline(GetPlayerIP(Index)) Then
            Call TextAdd("Received connection from " & GetPlayerIP(Index) & ".")
        ElseIf Options.MultipleIP = 0 And IsMultiIPOnline(GetPlayerIP(Index)) Then
            ' Tried Multiple connections
            Call AlertMsg(Index, "Multiple account logins are not authorized.")
            frmServer.Socket(Index).Close
        End If
    Else
        Call AlertMsg(Index, "The server is full! Try back again later.")
        frmServer.Socket(Index).Close
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

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
    Dim buffer() As Byte
    Dim pLength As Long

    If GetPlayerAccess(Index) <= 0 Then
        ' Check for data flooding
        If TempPlayer(Index).DataBytes > 1000 Then Exit Sub
    
        ' Check for Packet flooding
        If TempPlayer(Index).DataPackets > 105 Then Exit Sub
    End If
            
    ' Check if elapsed time has passed
    TempPlayer(Index).DataBytes = TempPlayer(Index).DataBytes + DataLength
    
    If timeGetTime >= TempPlayer(Index).DataTimer Then
        TempPlayer(Index).DataTimer = timeGetTime + 1000
        TempPlayer(Index).DataBytes = 0
        TempPlayer(Index).DataPackets = 0
    End If
    
    ' Get the data from the socket now
    frmServer.Socket(Index).GetData buffer(), vbUnicode, DataLength
    TempPlayer(Index).buffer.WriteBytes buffer()
    
    If TempPlayer(Index).buffer.Length >= 4 Then
        pLength = TempPlayer(Index).buffer.ReadLong(False)
    
        If pLength < 0 Then Exit Sub
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(Index).buffer.Length - 4
        If pLength <= TempPlayer(Index).buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).buffer.ReadLong
            HandleData Index, TempPlayer(Index).buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        
        If TempPlayer(Index).buffer.Length >= 4 Then
            pLength = TempPlayer(Index).buffer.ReadLong(False)
        
            If pLength < 0 Then Exit Sub
        End If
    Loop
            
    TempPlayer(Index).buffer.Trim
End Sub

Sub CloseSocket(ByVal Index As Long, Optional ByVal NoMessage As Boolean = False)
    Dim i As Long

    If Index > 0 And Index <= MAX_PLAYERS Then
        Call LeftGame(Index)
        
        If NoMessage = False Then
            Call TextAdd("Connection from " & GetPlayerIP(Index) & " has been terminated.")
        End If

        frmServer.Socket(Index).Close
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
Sub SendWhosOnline(ByVal Index As Long)
    Dim s As String
    Dim n As Long
    Dim i As Long

    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If Not i = Index Then
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

    Call PlayerMsg(Index, s, WhoColor)
End Sub
'Character Editor
Sub SendPlayersOnline(ByVal Index As Long)
    Dim buffer As clsBuffer, i As Long
    Dim list As String

    If Index > Player_HighIndex Or Index < 1 Then Exit Sub
    Set buffer = New clsBuffer
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
                If i <> Player_HighIndex Then
                    list = list & GetPlayerName(i) & ", "
                Else
                    list = list & GetPlayerName(i)
                End If
        End If
    Next
    
    buffer.WriteLong SPlayersOnline
    buffer.WriteString list
 
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub
'Character Editor
Sub SendAllCharacters(Index As Long, Optional everyone As Boolean = False)
    Dim buffer As clsBuffer, i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAllCharacters
    
    buffer.WriteString GetCharList
    
    SendDataTo Index, buffer.ToArray
    
    Set buffer = Nothing
End Sub

Function PlayerData(ByVal Index As Long) As Byte()
    Dim buffer As clsBuffer, i As Long

    If Index > Player_HighIndex Or Index < 1 Then Exit Function
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerData
    buffer.WriteLong Index
    buffer.WriteInteger Account(Index).Chars(GetPlayerChar(Index)).Face
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteByte GetPlayerGender(Index)
    buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).Class
    buffer.WriteInteger GetPlayerLevel(Index)
    buffer.WriteInteger GetPlayerPoints(Index)
    buffer.WriteInteger GetPlayerSprite(Index)
    buffer.WriteInteger GetPlayerMap(Index)
    buffer.WriteByte GetPlayerX(Index)
    buffer.WriteByte GetPlayerY(Index)
    buffer.WriteByte GetPlayerDir(Index)
    buffer.WriteByte GetPlayerAccess(Index)
    buffer.WriteByte GetPlayerPK(Index)
    
    If GetPlayerGuild(Index) > 0 Then
        buffer.WriteString Guild(GetPlayerGuild(Index)).Name
    Else
        buffer.WriteString vbNullString
    End If
    
    buffer.WriteByte GetPlayerGuildAccess(Index)
    
    For i = 1 To Stats.Stat_count - 1
        buffer.WriteInteger GetPlayerStat(Index, i)
    Next
    
    ' Amount of titles
    buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles
    
    ' Send player titles
    For i = 1 To Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles
        buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).title(i)
    Next
    
    ' Send the player's current title
    buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).CurrentTitle
    
    ' Send player status
    buffer.WriteString Account(Index).Chars(GetPlayerChar(Index)).Status
    
    PlayerData = buffer.ToArray()
    Set buffer = Nothing
End Function


Sub SendJoinMap(ByVal Index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    ' Send all players on current map to index
    If GetTotalMapPlayers(GetPlayerMap(Index)) > 0 Then
        For i = 1 To Player_HighIndex
            If IsPlaying(i) Then
                SendDataToMap GetPlayerMap(Index), PlayerData(i)
                Call SendPlayerEquipmentToMapBut(i)
            End If
        Next
    End If
    
    ' Send Index's player data to everyone on the map including themself
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
    
    ' Send player's equipment to new map
    SendPlayerEquipmentTo Index
    
    ' Send the npc targets to the player
    For i = 1 To Map(GetPlayerMap(Index)).Npc_HighIndex
        If MapNpc(GetPlayerMap(Index)).NPC(i).Num > 0 Then
            Call SendMapNpcTarget(GetPlayerMap(Index), i, MapNpc(GetPlayerMap(Index)).NPC(i).Target, MapNpc(GetPlayerMap(Index)).NPC(i).TargetType)
        Else
            ' Send 0 so it uncaches any old data
            Call SendMapNpcTarget(GetPlayerMap(Index), i, 0, 0)
        End If
    Next
    
    Set buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Integer)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SLeft
    buffer.WriteLong Index
    SendDataToMapBut Index, MapNum, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendPlayerData(ByVal Index As Long)
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
End Sub

'Character Editor
Sub SendExtendedPlayerData(Index As Long, playerName As String)
    'Check if He is Online
    Dim i As Long, j As Long, tempPlayer_ As PlayerRec
    For i = 1 To MAX_PLAYERS
        For j = 1 To MAX_CHARS
            If Account(i).Login = "" Then GoTo use_offline_player
            If Trim(Account(i).Chars(j).Name) = playerName Then
                tempPlayer_ = Account(i).Chars(j)
                GoTo use_online_player
            End If
        Next
    Next
use_offline_player:
    'Find associated Account Name
    Dim F As Long
    Dim s As String
    Dim charLogin() As String
    
    F = FreeFile
    
    Open App.path & "\data\accounts\charlist.txt" For Input As #F
        Do While Not EOF(F)
            Input #F, s
            charLogin = Split(s, ":")
            If charLogin(0) = playerName Then Exit Do
        Loop
    Close #F
    
    'Load Character into temp variable - charLogin(0) -> Character Name | charLogin(1) -> Account/Login Name
    Dim tempAccount As AccountRec
    Dim filename As String
    
    filename = App.path & "\data\accounts\" & charLogin(1) & "\data.bin"
    
    F = FreeFile
    
    Open filename For Binary As #F
        Get #F, , tempAccount
    Close #F
    
    'Get Character info, that we are requesting -> playerName
    Dim requestedClientPlayer As PlayerEditableRec
    For i = 1 To MAX_CHARS
        If Trim$(tempAccount.Chars(i).Name) = playerName Then
            tempPlayer_ = tempAccount.Chars(i)
            Exit For
        End If
    Next
use_online_player:
    'Copy over data that's available...
    requestedClientPlayer.Name = tempPlayer_.Name
    requestedClientPlayer.Level = tempPlayer_.Level
    requestedClientPlayer.Class = tempPlayer_.Class
    requestedClientPlayer.Access = tempPlayer_.Access
    requestedClientPlayer.Exp = tempPlayer_.Exp
    requestedClientPlayer.Gender = tempPlayer_.Gender
    requestedClientPlayer.Login = "XXXX" ' Do we really want to edit it in client? Is it safe to send it?
    requestedClientPlayer.Password = "XXXX" ' Do we really want to edit it in client? Is it safe to send it?
    requestedClientPlayer.Points = tempPlayer_.Points
    requestedClientPlayer.Sprite = tempPlayer_.Sprite
    Dim tempSize As Long
    tempSize = LenB(tempPlayer_.Stat(1)) * UBound(tempPlayer_.Stat)
    CopyMemory ByVal VarPtr(requestedClientPlayer.Stat(1)), ByVal VarPtr(tempPlayer_.Stat(1)), tempSize
    tempSize = LenB(tempPlayer_.Vital(1)) * UBound(tempPlayer_.Vital)
    CopyMemory ByVal VarPtr(requestedClientPlayer.Vital(1)), ByVal VarPtr(tempPlayer_.Vital(1)), tempSize
    
    'Send Data Over Network to Admin
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SExtendedPlayerData
    
    Dim PlayerSize As Long
    Dim PlayerData() As Byte
    
    PlayerSize = LenB(requestedClientPlayer)
    ReDim PlayerData(PlayerSize - 1)
    CopyMemory PlayerData(0), ByVal VarPtr(requestedClientPlayer), PlayerSize
    buffer.WriteBytes PlayerData
    
    SendDataTo Index, buffer.ToArray
    
    Set buffer = Nothing
    
    
End Sub
Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    buffer.WriteLong SMapData
    buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Integer)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapItemData
    For i = 1 To MAX_MAP_ITEMS
        buffer.WriteString MapItem(MapNum, i).playerName
        buffer.WriteLong MapItem(MapNum, i).Num
        buffer.WriteLong MapItem(MapNum, i).Value
        buffer.WriteInteger MapItem(MapNum, i).Durability
        buffer.WriteByte MapItem(MapNum, i).X
        buffer.WriteByte MapItem(MapNum, i).Y
    Next

    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapItemToMap(ByVal MapNum As Integer, ByVal MapSlotNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    
    buffer.WriteLong SSpawnItem
    buffer.WriteLong MapSlotNum
    buffer.WriteString MapItem(MapNum, MapSlotNum).playerName
    buffer.WriteLong MapItem(MapNum, MapSlotNum).Num
    buffer.WriteLong MapItem(MapNum, MapSlotNum).Value
    buffer.WriteInteger MapItem(MapNum, MapSlotNum).Durability
    buffer.WriteLong MapItem(MapNum, MapSlotNum).X
    buffer.WriteLong MapItem(MapNum, MapSlotNum).Y
    
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

Sub SendMapNpcsTo(ByVal Index As Long, ByVal MapNum As Integer)
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

    SendDataTo Index, buffer.ToArray()
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

Sub SendMorals(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_MORALS
        If Len(Trim$(Moral(i).Name)) > 0 Then
            Call SendUpdateMoralTo(Index, i)
        End If
    Next
End Sub

Sub SendClasses(ByVal Index As Long)
    Dim i As Long
    
    For i = 1 To MAX_CLASSES
        If Len(Trim$(Class(i).Name)) > 0 Then
            Call SendUpdateClassTo(Index, i)
        End If
    Next
End Sub

Sub SendEmoticons(ByVal Index As Long)
    Dim i As Long
    
    For i = 1 To MAX_EMOTICONS
        If Len(Trim$(Emoticon(i).Command)) > 0 Then
            Call SendUpdateEmoticonTo(Index, i)
        End If
    Next
End Sub

Sub SendItems(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ITEMS
        If Len(Trim$(Item(i).Name)) > 0 Then
            Call SendUpdateItemTo(Index, i)
        End If
    Next
End Sub

Sub SendTitles(ByVal Index As Long)
    Dim i As Long
    
    For i = 1 To MAX_TITLES
        If Len(Trim$(title(i).Name)) > 0 Then
            Call SendUpdateTitleTo(Index, i)
        End If
    Next
End Sub

Sub SendAnimations(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_ANIMATIONS
        If Len(Trim$(Animation(i).Name)) > 0 Then
            Call SendUpdateAnimationTo(Index, i)
        End If
    Next
End Sub

Sub SendNpcs(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS
        If Len(Trim$(NPC(i).Name)) > 0 Then
            Call SendUpdateNpcTo(Index, i)
        End If
    Next
End Sub

Sub SendResources(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_RESOURCES
        If Len(Trim$(Resource(i).Name)) > 0 Then
            Call SendUpdateResourceTo(Index, i)
        End If
    Next
End Sub

Sub SendInventory(ByVal Index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        buffer.WriteLong GetPlayerInvItemNum(Index, i)
        buffer.WriteLong GetPlayerInvItemValue(Index, i)
        buffer.WriteInteger GetPlayerInvItemDur(Index, i)
        buffer.WriteByte GetPlayerInvItemBind(Index, i)
    Next

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Byte)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerInvUpdate
    buffer.WriteByte InvSlot
    buffer.WriteLong GetPlayerInvItemNum(Index, InvSlot)
    buffer.WriteLong GetPlayerInvItemValue(Index, InvSlot)
    buffer.WriteInteger GetPlayerInvItemDur(Index, InvSlot)
    buffer.WriteByte GetPlayerInvItemBind(Index, InvSlot)
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerEquipmentTo(ByVal Index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim i As Byte
    
    buffer.WriteLong SPlayerWornEq
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteLong GetPlayerEquipment(Index, i)
    Next
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteInteger GetPlayerEquipmentDur(Index, i)
    Next
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerEquipmentToMap(ByVal Index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    Dim i As Byte
    
    buffer.WriteLong SMapWornEq
    buffer.WriteLong Index
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteLong GetPlayerEquipment(Index, i)
    Next

    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerEquipmentToMapBut(ByVal Index As Long)
    Dim buffer As clsBuffer
    Dim i As Byte
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMapWornEq
    buffer.WriteLong Index
    
    For i = 1 To Equipment.Equipment_Count - 1
        buffer.WriteLong GetPlayerEquipment(Index, i)
    Next
    
    SendDataToMapBut Index, GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer

    Select Case Vital
        Case HP
            buffer.WriteLong SPlayerHP
            buffer.WriteLong Index
            buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
            buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
        Case MP
            buffer.WriteLong SPlayerMP
            buffer.WriteLong Index
            buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
            buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
    End Select

    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendVitalTo(ByVal Index As Long, Player As Long, ByVal Vital As Vitals)
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

    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerExp(ByVal Index As Long)
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerEXP
    buffer.WriteLong Index
    buffer.WriteLong GetPlayerExp(Index)
    buffer.WriteLong GetPlayerNextLevel(Index)
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerStats(ByVal Index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerStats
    buffer.WriteLong Index
    
    For i = 1 To Stats.Stat_count - 1
        buffer.WriteInteger GetPlayerStat(Index, i)
    Next
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerPoints(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerPoints
    buffer.WriteLong Index
    buffer.WriteInteger GetPlayerPoints(Index)
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerLevel(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerLevel
    buffer.WriteLong Index
    buffer.WriteInteger GetPlayerLevel(Index)
    
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerGuild(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerGuild
    buffer.WriteLong Index
    
    If GetPlayerGuild(Index) > 0 Then
        buffer.WriteString Guild(GetPlayerGuild(Index)).Name
    Else
        buffer.WriteString vbNullString
    End If
    
    buffer.WriteByte GetPlayerGuildAccess(Index)
    
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendGuildInvite(ByVal Index As Long, ByVal OtherPlayer As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SGuildInvite
    
    buffer.WriteString Trim$(Account(Index).Chars(GetPlayerChar(Index)).Name)
    buffer.WriteString Trim$(Guild(GetPlayerGuild(Index)).Name)
    
    SendDataTo OtherPlayer, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerGuildMembers(ByVal Index As Long, Optional ByVal Ignore As Byte = 0)
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
            If Not Ignore = i And Not i = Index Then
                If GetPlayerGuild(i) = GetPlayerGuild(Index) Then
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
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub
Sub SendPlayerSprite(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerSprite
    buffer.WriteLong Index
    buffer.WriteInteger GetPlayerSprite(Index)
      
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerTitles(ByVal Index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerTitles
    buffer.WriteLong Index
    
    ' Amount of titles
    buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles
    
    ' Send player titles
    For i = 1 To Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles
        buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).title(i)
    Next
    
    ' Send the player's current title
    buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).CurrentTitle
    
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerStatus(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerStatus
    buffer.WriteLong Index
    
    buffer.WriteString Account(Index).Chars(GetPlayerChar(Index)).Status
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerPK(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SPlayerPK
    buffer.WriteLong Index
    
    buffer.WriteByte GetPlayerPK(Index)
    
    SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendWelcome(ByVal Index As Long)
    ' Send the MOTD
    If Not Trim$(Options.MOTD) = vbNullString Then
        Call PlayerMsg(Index, Options.MOTD, BrightCyan)
    End If
    
    ' Send the SMOTD
    If Not Trim$(Options.SMOTD) = vbNullString Then
        If GetPlayerAccess(Index) >= STAFF_MODERATOR Then
            Call PlayerMsg(Index, Options.SMOTD, Cyan)
        End If
    End If
    
    ' Send the GMOTD
    If GetPlayerGuild(Index) > 0 Then
        If Not Trim$(Guild(GetPlayerGuild(Index)).MOTD) = vbNullString Then
            Call PlayerMsg(Index, Trim$(Guild(GetPlayerGuild(Index)).MOTD), BrightGreen)
        End If
    End If
End Sub

Sub SendLeftGame(ByVal Index As Long)
    Dim buffer As clsBuffer, i As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SLeaveGame
    buffer.WriteLong Index
    
    SendDataToAllBut Index, buffer.ToArray()
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

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Integer)
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
    
    SendDataTo Index, buffer.ToArray()
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

Sub SendUpdateAnimationTo(ByVal Index As Long, ByVal AnimationNum As Long)
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
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub
Sub SendAssociatedCharacters()

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

Sub SendUpdateNpcTo(ByVal Index As Long, ByVal npcnum As Long)
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
    
    SendDataTo Index, buffer.ToArray()
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

Sub SendUpdateResourceTo(ByVal Index As Long, ByVal ResourceNum As Long)
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
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendShops(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SHOPS
        If Len(Trim$(Shop(i).Name)) > 0 Then
            Call SendUpdateShopTo(Index, i)
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

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum As Long)
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
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim i As Long

    For i = 1 To MAX_SPELLS
        If Len(Trim$(Spell(i).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, i)
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

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
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
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SSpells
    For i = 1 To MAX_PLAYER_SPELLS
        buffer.WriteLong GetPlayerSpell(Index, i)
    Next
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SSpell
    buffer.WriteByte SpellSlot
    buffer.WriteLong GetPlayerSpell(Index, SpellSlot)
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal Index As Long, ByVal Resource_Num As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Set buffer = New clsBuffer
    
    buffer.WriteLong SResourceCache
    buffer.WriteLong ResourceCache(GetPlayerMap(Index)).Resource_Count

    If ResourceCache(GetPlayerMap(Index)).Resource_Count > 0 Then
        For i = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            buffer.WriteByte ResourceCache(GetPlayerMap(Index)).ResourceData(i).ResourceState
            buffer.WriteInteger ResourceCache(GetPlayerMap(Index)).ResourceData(i).X
            buffer.WriteInteger ResourceCache(GetPlayerMap(Index)).ResourceData(i).Y
        Next
    End If

    SendDataTo Index, buffer.ToArray()
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

Sub SendSpellCooldown(ByVal Index As Long, ByVal Slot As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSpellCooldown
    buffer.WriteByte Slot
    buffer.WriteLong Account(Index).Chars(GetPlayerChar(Index)).SpellCD(Slot)
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendClearAccountSpellBuffer(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SClearSpellBuffer
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal MapNum As Integer, ByVal Index As Long, ByVal Message As String, ByVal SayColor As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    buffer.WriteString Message
    buffer.WriteString "[Map] "
    buffer.WriteLong SayColor
    
    SendDataToMap MapNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal Index As Long, ByVal Message As String, ByVal SayColor As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSayMsg
    
    buffer.WriteString GetPlayerName(Index)
    buffer.WriteLong GetPlayerAccess(Index)
    buffer.WriteLong GetPlayerPK(Index)
    buffer.WriteString Message
    buffer.WriteString "[Global] "
    buffer.WriteLong SayColor
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub ResetShopAction(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SResetShopAction
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendStunned(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SStunned
    
    buffer.WriteLong TempPlayer(Index).StunDuration
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendBank(ByVal Index As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong sbank
    
    For i = 1 To MAX_BANK
        buffer.WriteLong Account(Index).Bank.Item(i).Num
        buffer.WriteLong Account(Index).Bank.Item(i).Value
    Next
    
    SendDataTo Index, buffer.ToArray()
    
    Set buffer = Nothing
End Sub

Sub SendMapKey(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal Value As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SMapKey
    
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteByte Value
    
    SendDataTo Index, buffer.ToArray()
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

Sub SendOpenShop(ByVal Index As Long, ByVal ShopNum As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SOpenShop
    
    buffer.WriteLong ShopNum
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal Index As Long, ByVal Movement As Byte, Optional ByVal SendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerMove
    
    buffer.WriteLong Index
    buffer.WriteByte GetPlayerX(Index)
    buffer.WriteByte GetPlayerY(Index)
    buffer.WriteByte GetPlayerDir(Index)
    buffer.WriteByte Movement
    
    If Not SendToSelf Then
        SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Else
        SendDataTo Index, buffer.ToArray()
    End If
    Set buffer = Nothing
End Sub

Sub SendPlayerPosition(ByVal Index As Long, Optional ByVal SendToSelf As Boolean = False)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerWarp
    
    buffer.WriteLong Index
    buffer.WriteByte GetPlayerX(Index)
    buffer.WriteByte GetPlayerY(Index)
    buffer.WriteByte GetPlayerDir(Index)
    
    If Not SendToSelf Then
        SendDataToMap GetPlayerMap(Index), buffer.ToArray()
    Else
        SendDataTo Index, buffer.ToArray()
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

Sub SendTrade(ByVal Index As Long, ByVal TradeTarget As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STrade
    
    buffer.WriteLong TradeTarget
    buffer.WriteString Trim$(GetPlayerName(TradeTarget))
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCloseTrade
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal Index As Long, ByVal DataType As Byte)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim TradeTarget As Long
    Dim TotalWorth As Long
    
    TradeTarget = TempPlayer(Index).InTrade
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeUpdate
    
    buffer.WriteByte DataType
    If DataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            buffer.WriteLong TempPlayer(Index).TradeOffer(i).Num
            buffer.WriteLong TempPlayer(Index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(Index).TradeOffer(i).Num > 0 Then
                ' currency?
                If Item(TempPlayer(Index).TradeOffer(i).Num).Type = ITEM_TYPE_CURRENCY Then
                    TotalWorth = TotalWorth + (Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).Price * TempPlayer(Index).TradeOffer(i).Value)
                Else
                    TotalWorth = TotalWorth + Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(i).Num)).Price
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
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal Index As Long, ByVal Status As Byte)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STradeStatus
    
    buffer.WriteByte Status
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendAttack(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SAttack
    buffer.WriteLong Index
    
    SendDataToMapBut Index, GetPlayerMap(Index), buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPlayerTarget(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong STarget
    
    buffer.WriteByte TempPlayer(Index).Target
    buffer.WriteByte TempPlayer(Index).TargetType
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendHotbar(ByVal Index As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SHotbar
    
    For i = 1 To MAX_HOTBAR
        buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).Hotbar(i).Slot
        buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).Hotbar(i).SType
    Next
    
    SendDataTo Index, buffer.ToArray()
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

Sub SendLogs(ByVal Index As Long, Msg As String, Name As String)
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
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub UpdateFriendsList(Index)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim i As Long, n As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SFriendsList
    
    If Account(Index).Friends.AmountOfFriends = 0 Then
        buffer.WriteByte Account(Index).Friends.AmountOfFriends
        GoTo Finish
    End If
   
    ' Sends the amount of friends in friends list
    buffer.WriteByte Account(Index).Friends.AmountOfFriends
   
    ' Check to see if they are online
    For i = 1 To Account(Index).Friends.AmountOfFriends
        Name = Trim$(Account(Index).Friends.Members(i))
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
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub UpdateFoesList(Index)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim i As Long, n As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SFoesList
    
    If Account(Index).Foes.Amount = 0 Then
        buffer.WriteByte Account(Index).Foes.Amount
        GoTo Finish
    End If
   
    ' Sends the amount of Foes in Foes list
    buffer.WriteByte Account(Index).Foes.Amount
   
    ' Check to see if they are online
    For i = 1 To Account(Index).Foes.Amount
        Name = Trim$(Account(Index).Foes.Members(i))
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
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendInGame(ByVal Index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SInGame
    
    SendDataTo Index, buffer.ToArray()
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

Sub SendSoundTo(ByVal Index As Integer, Sound As String)
    Dim buffer As clsBuffer
    
    ' Don't send it if there's nothing to send
    If Sound = vbNullString Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSound
    
    buffer.WriteString Sound
    SendDataTo Index, buffer.ToArray()
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

Sub SendPlayerSound(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal EntityType As Long, ByVal EntityNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SEntitySound
    
    buffer.WriteLong X
    buffer.WriteLong Y
    buffer.WriteLong EntityType
    buffer.WriteLong EntityNum
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendMapSound(ByVal MapNum As Integer, ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal EntityType As Long, ByVal EntityNum As Long)
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

Sub SendLogin(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SLoginOk
    
    buffer.WriteLong Index
    buffer.WriteLong Player_HighIndex
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SNewCharClasses
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendNews(ByVal Index As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSendNews
    
    buffer.WriteString Options.News
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCheckForMap(ByVal Index As Long, MapNum As Integer)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SCheckForMap
    
    buffer.WriteInteger MapNum
    buffer.WriteInteger Map(MapNum).Revision
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal Index As Long, ByVal TradeRequest As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong STradeRequest
    
    buffer.WriteLong TradeRequest
    buffer.WriteString Trim$(Account(TradeRequest).Chars(GetPlayerChar(TradeRequest)).Name)
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal Index As Long, ByVal OtherPlayer As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyInvite
    
    buffer.WriteString Trim$(Account(OtherPlayer).Chars(GetPlayerChar(OtherPlayer)).Name)
    
    SendDataTo Index, buffer.ToArray()
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

Sub SendPartyUpdateTo(ByVal Index As Long)
    Dim buffer As clsBuffer, i As Long, PartyNum As Long
    
    Set buffer = New clsBuffer
    buffer.WriteLong SPartyUpdate
    
    ' Check if we're in a party
    PartyNum = TempPlayer(Index).InParty
    
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
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal PartyNum As Long, ByVal Index As Long)
    Dim buffer As clsBuffer, i As Long

    Set buffer = New clsBuffer
    buffer.WriteLong SPartyVitals
    
    buffer.WriteLong Index
    For i = 1 To Vitals.Vital_Count - 1
        buffer.WriteLong GetPlayerMaxVital(Index, i)
        buffer.WriteLong Account(Index).Chars(GetPlayerChar(Index)).Vital(i)
    Next
    
    SendDataToParty PartyNum, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateTitleToAll(ByVal TitleNum As Long)
    Dim buffer As clsBuffer
    Dim TitleSize As Long
    Dim TitleData() As Byte
    
    Set buffer = New clsBuffer
    
    TitleSize = LenB(title(TitleNum))
    ReDim TitleData(TitleSize - 1)
    CopyMemory TitleData(0), ByVal VarPtr(title(TitleNum)), TitleSize
    buffer.WriteLong SUpdateTitle
    
    buffer.WriteLong TitleNum
    buffer.WriteBytes TitleData
    
    SendDataToAll buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateTitleTo(ByVal Index As Long, ByVal TitleNum As Long)
    Dim buffer As clsBuffer
    Dim TitleSize As Long
    Dim TitleData() As Byte
    
    Set buffer = New clsBuffer
    
    TitleSize = LenB(title(TitleNum))
    ReDim TitleData(TitleSize - 1)
    CopyMemory TitleData(0), ByVal VarPtr(title(TitleNum)), TitleSize
    buffer.WriteLong SUpdateTitle
    
    buffer.WriteLong TitleNum
    buffer.WriteBytes TitleData
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendCloseClient(Index As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SCloseClient
    
    SendDataTo Index, buffer.ToArray()
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

Sub SendUpdateMoralTo(ByVal Index As Long, ByVal MoralNum As Long)
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
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateClassTo(ByVal Index As Long, ByVal ClassNum As Long)
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
    
    SendDataTo Index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub SendUpdateEmoticonTo(ByVal Index As Long, ByVal EmoticonNum As Long)
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
    
    SendDataTo Index, buffer.ToArray()
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

Sub SendCheckEmoticon(ByVal Index As Long, ByVal MapNum As Long, ByVal EmoticonNum As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteLong SCheckEmoticon
    
    buffer.WriteLong Index
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

Sub SendSpecialEffect(ByVal Index As Long, EffectType As Long, Optional Data1 As Long = 0, Optional Data2 As Long = 0, Optional Data3 As Long = 0, Optional Data4 As Long = 0)
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
    
    SendDataTo Index, buffer.ToArray
    Set buffer = Nothing
End Sub
