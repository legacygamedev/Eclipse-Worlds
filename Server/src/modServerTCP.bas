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
    Dim I As Long

    For I = 1 To MAX_MAPS
        Call MapCache_Create(I)
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
    Dim I As Long

    For I = 1 To Player_HighIndex
        If IsConnected(I) Then
            If LCase$(Trim$(Account(I).Login)) = LCase$(Login) Then
                IsMultiAccounts = True
                Exit Function
            End If
        End If
    Next
End Function

Function IsMultiIPOnline(ByVal IP As String) As Boolean
    Dim I As Long
    Dim n As Long

    For I = 1 To Player_HighIndex
        If IsConnected(I) Then
            If Trim$(GetPlayerIP(I)) = IP Then
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
    Dim Buffer As clsBuffer
    Dim TempData() As Byte
    
    If IsConnected(Index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
        
        If IsConnected(Index) Then
            frmServer.Socket(Index).SendData Buffer.ToArray()
        End If
    End If
End Sub

Sub SendDataToAll(ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            Call SendDataTo(I, Data)
        End If
    Next
End Sub

Sub SendDataToAllBut(ByVal Index As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If Not I = Index Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next
End Sub

Sub SendDataToMap(ByVal MapNum As Integer, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                Call SendDataTo(I, Data)
            End If
        End If
    Next
End Sub

Sub SendDataToMapBut(ByVal Index As Long, ByVal MapNum As Integer, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If GetPlayerMap(I) = MapNum Then
                If Not I = Index Then
                    Call SendDataTo(I, Data)
                End If
            End If
        End If
    Next
End Sub

Sub SendDataToParty(ByVal PartyNum As Long, ByRef Data() As Byte)
    Dim I As Long

    For I = 1 To Party(PartyNum).MemberCount
        If Party(PartyNum).Member(I) > 0 Then
            Call SendDataTo(Party(PartyNum).Member(I), Data)
        End If
    Next
End Sub

Public Sub GlobalMsg(ByVal Msg As String, ByVal Color As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SGlobalMsg
    Buffer.WriteString Msg
    Buffer.WriteLong Color
    SendDataToAll Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub AdminMsg(ByVal Msg As String, Color As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    Dim LogMsg As String
    
    Set Buffer = New clsBuffer
    
    ' Add server log
    Call AddLog(Msg, "Player")
    
    LogMsg = Msg
    Msg = "[Admin] " & Msg
    
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    
    For I = 1 To Player_HighIndex
        If IsPlaying(I) And GetPlayerAccess(I) >= STAFF_MODERATOR Then
            SendDataTo I, Buffer.ToArray
            Call SendLogs(I, LogMsg, "Admin")
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Long, Optional ByVal QuestMsg As Boolean = False, Optional ByVal QuestNum As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    Buffer.WriteLong QuestMsg
    Buffer.WriteLong QuestNum
    SendDataTo Index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub MapMsg(ByVal MapNum As Integer, ByVal Msg As String, ByVal Color As Long)
    Dim Buffer As clsBuffer
    Dim I As Long

    Set Buffer = New clsBuffer

    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    SendDataToMap MapNum, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub GuildMsg(ByVal Index As Long, ByVal Msg As String, ByVal Color As Long, Optional HideName As Boolean = False)
    Dim I As Long
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

    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If GetPlayerGuild(I) = GetPlayerGuild(Index) Then
                PlayerMsg I, Msg, Color
                Call SendLogs(I, LogMsg, "Guild")
            End If
        End If
    Next
End Sub

Public Sub AlertMsg(ByVal Index As Long, ByVal Msg As String)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteString Msg
    SendDataTo Index, Buffer.ToArray
    DoEvents
    
    Set Buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal PartyNum As Long, ByVal Msg As String, ByVal Color As Long)
    Dim I As Long
    Dim LogMsg As String
    
    ' Add server log
    Call AddLog(Msg, "Player")
    
    LogMsg = Msg
    
    Msg = "[Party] " & Msg
    
    ' Send message to all people
    For I = 1 To MAX_PARTY_MEMBERS
        ' Exist?
        If Party(PartyNum).Member(I) > 0 Then
            ' Make sure they're logged on
            If IsConnected(Party(PartyNum).Member(I)) And IsPlaying(Party(PartyNum).Member(I)) Then
                PlayerMsg Party(PartyNum).Member(I), Msg, Color
                Call SendLogs(Party(PartyNum).Member(I), LogMsg, "Party")
            End If
        End If
    Next
End Sub

Sub AcceptConnection(ByVal Index As Long, ByVal SocketId As Long)
    Dim I As Long

    If (Index = 0) Then
        I = FindOpenPlayerSlot

        If Not I = 0 Then
            If Not IsConnected(I) Then
                ' We can connect them
                frmServer.Socket(I).Close
                frmServer.Socket(I).Accept SocketId
                Call SocketConnected(I)
            End If
        End If
    End If
End Sub

Sub SocketConnected(ByVal Index As Long)
    Dim I As Long
    
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
    
    For I = MAX_PLAYERS To 1 Step -1
        If IsConnected(I) Then
            Player_HighIndex = I
            Exit For
        End If
    Next
    
    ' Send the new highIndex to all logged in players
    SendPlayer_HighIndex
End Sub

Sub IncomingData(ByVal Index As Long, ByVal DataLength As Long)
    Dim Buffer() As Byte
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
    frmServer.Socket(Index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(Index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(Index).Buffer.Length >= 4 Then
        pLength = TempPlayer(Index).Buffer.ReadLong(False)
    
        If pLength < 0 Then Exit Sub
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(Index).Buffer.Length - 4
        If pLength <= TempPlayer(Index).Buffer.Length - 4 Then
            TempPlayer(Index).DataPackets = TempPlayer(Index).DataPackets + 1
            TempPlayer(Index).Buffer.ReadLong
            HandleData Index, TempPlayer(Index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        
        If TempPlayer(Index).Buffer.Length >= 4 Then
            pLength = TempPlayer(Index).Buffer.ReadLong(False)
        
            If pLength < 0 Then Exit Sub
        End If
    Loop
            
    TempPlayer(Index).Buffer.Trim
End Sub

Sub CloseSocket(ByVal Index As Long, Optional ByVal NoMessage As Boolean = False)
    Dim I As Long

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
    For I = MAX_PLAYERS To 1 Step -1
        If IsConnected(I) Then
            Player_HighIndex = I
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
    Dim I As Long, z As Long, w As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong MapNum
    Buffer.WriteString Trim$(Map(MapNum).Name)
    Buffer.WriteString Trim$(Map(MapNum).Music)
    Buffer.WriteString Trim$(Map(MapNum).BGS)
    Buffer.WriteLong Map(MapNum).Revision
    Buffer.WriteByte Map(MapNum).Moral
    Buffer.WriteLong Map(MapNum).Up
    Buffer.WriteLong Map(MapNum).Down
    Buffer.WriteLong Map(MapNum).Left
    Buffer.WriteLong Map(MapNum).Right
    Buffer.WriteLong Map(MapNum).BootMap
    Buffer.WriteByte Map(MapNum).BootX
    Buffer.WriteByte Map(MapNum).BootY
    
    Buffer.WriteLong Map(MapNum).Weather
    Buffer.WriteLong Map(MapNum).WeatherIntensity
    
    Buffer.WriteLong Map(MapNum).Fog
    Buffer.WriteLong Map(MapNum).FogSpeed
    Buffer.WriteLong Map(MapNum).FogOpacity
    
    Buffer.WriteLong Map(MapNum).Panorama
    
    Buffer.WriteLong Map(MapNum).Red
    Buffer.WriteLong Map(MapNum).Green
    Buffer.WriteLong Map(MapNum).Blue
    Buffer.WriteLong Map(MapNum).Alpha
    
    Buffer.WriteByte Map(MapNum).MaxX
    Buffer.WriteByte Map(MapNum).MaxY
    
    Buffer.WriteByte Map(MapNum).NPC_HighIndex

    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            With Map(MapNum).Tile(X, Y)
                For I = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(I).X
                    Buffer.WriteLong .Layer(I).Y
                    Buffer.WriteLong .Layer(I).Tileset
                Next
                
                For z = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Autotile(z)
                Next
                
                Buffer.WriteByte .Type
                Buffer.WriteLong .Data1
                Buffer.WriteLong .Data2
                Buffer.WriteLong .Data3
                Buffer.WriteString .Data4
                Buffer.WriteByte .DirBlock
            End With
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Buffer.WriteLong Map(MapNum).NPC(X)
        Buffer.WriteLong Map(MapNum).NPCSpawnType(X)
    Next

    MapCache(MapNum).Data = Buffer.ToArray()
    Set Buffer = Nothing
End Sub

' *****************************
' ** Outgoing Server Packets **
' *****************************
Sub SendWhosOnline(ByVal Index As Long)
    Dim s As String
    Dim n As Long
    Dim I As Long

    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If Not I = Index Then
                s = s & GetPlayerName(I) & ", "
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
    Dim Buffer As clsBuffer, I As Long
    Dim list As String

    If Index > Player_HighIndex Or Index < 1 Then Exit Sub
    Set Buffer = New clsBuffer
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
                If I <> Player_HighIndex Then
                    list = list & GetPlayerName(I) & ":" & Account(I).Chars(GetPlayerChar(I)).Access & ":" & Account(I).Chars(GetPlayerChar(I)).Sprite & ", "
                Else
                    list = list & GetPlayerName(I) & ":" & Account(I).Chars(GetPlayerChar(I)).Access & ":" & Account(I).Chars(GetPlayerChar(I)).Sprite
                End If
        End If
    Next
    
    Buffer.WriteLong SPlayersOnline
    Buffer.WriteString list
 
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
'Character Editor
Sub SendAllCharacters(Index As Long, Optional everyone As Boolean = False)
    Dim Buffer As clsBuffer, I As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAllCharacters
    
    Buffer.WriteString GetCharList
    
    SendDataTo Index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Function PlayerData(ByVal Index As Long) As Byte()
    Dim Buffer As clsBuffer, I As Long

    If Index < 1 Or Index > Player_HighIndex Then Exit Function
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong Index
    Buffer.WriteInteger Account(Index).Chars(GetPlayerChar(Index)).Face
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteByte GetPlayerGender(Index)
    Buffer.WriteByte GetPlayerClass(Index)
    Buffer.WriteByte GetPlayerLevel(Index)
    Buffer.WriteInteger GetPlayerPoints(Index)
    Buffer.WriteInteger GetPlayerSprite(Index)
    Buffer.WriteInteger GetPlayerMap(Index)
    Buffer.WriteByte GetPlayerX(Index)
    Buffer.WriteByte GetPlayerY(Index)
    Buffer.WriteByte GetPlayerDir(Index)
    Buffer.WriteByte GetPlayerAccess(Index)
    Buffer.WriteByte GetPlayerPK(Index)
    
    If GetPlayerGuild(Index) > 0 Then
        Buffer.WriteString Guild(GetPlayerGuild(Index)).Name
    Else
        Buffer.WriteString vbNullString
    End If
    
    Buffer.WriteByte GetPlayerGuildAccess(Index)
    
    For I = 1 To Stats.Stat_count - 1
        Buffer.WriteInteger GetPlayerStat(Index, I)
    Next
    
    ' Amount of titles
    Buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles
    
    ' Send player titles
    For I = 1 To Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles
        Buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).Title(I)
    Next
    
    ' Send the player's current title
    Buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).CurrentTitle
    
    ' Send player status
    Buffer.WriteString Account(Index).Chars(GetPlayerChar(Index)).Status
    
    For I = 1 To Skills.Skill_Count - 1
        Buffer.WriteByte GetPlayerSkill(Index, I)
        Buffer.WriteLong GetPlayerSkillExp(Index, I)
    Next
    
    For I = 1 To MAX_QUESTS
        Buffer.WriteLong GetPlayerQuestCLIID(Index, I)
        Buffer.WriteLong GetPlayerQuestTaskID(Index, I)
        Buffer.WriteLong GetPlayerQuestAmount(Index, I)
        Buffer.WriteLong IsQuestCompleted(Index, I)
    Next I
    
    PlayerData = Buffer.ToArray()
    Set Buffer = Nothing
End Function

Sub SendJoinMap(ByVal Index As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If I <> Index Then
                If GetPlayerMap(I) = GetPlayerMap(Index) Then
                    SendDataTo Index, PlayerData(I)
                End If
            End If
        End If
    Next
    
    ' Send index's player data to everyone on the map including themself
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
    
    ' Send the NPC targets to the player
    For I = 1 To Map(GetPlayerMap(Index)).NPC_HighIndex
        If MapNPC(GetPlayerMap(Index)).NPC(I).Num > 0 Then
            Call SendMapNPCTarget(GetPlayerMap(Index), I, MapNPC(GetPlayerMap(Index)).NPC(I).target, MapNPC(GetPlayerMap(Index)).NPC(I).targetType)
        Else
            ' Send 0 so it uncaches any old data
            Call SendMapNPCTarget(GetPlayerMap(Index), I, 0, 0)
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal Index As Long, ByVal MapNum As Integer)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong Index
    SendDataToMapBut Index, MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerData(ByVal Index As Long)
    SendDataToMap GetPlayerMap(Index), PlayerData(Index)
End Sub

Sub SendAccessVerificator(ByVal Index As Long, success As Byte, Message As String, currentAccess As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAccessVerificator
    Buffer.WriteByte success
    Buffer.WriteString Message
    Buffer.WriteByte currentAccess
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

'Character Editor
Sub SendExtendedPlayerData(Index As Long, playerName As String)
    'Check if He is Online
    Dim I As Long, j As Long, tempPlayer_ As PlayerRec
    For I = 1 To MAX_PLAYERS
        For j = 1 To MAX_CHARS
            If Account(I).Login = "" Then GoTo use_offline_player
            If Trim(Account(I).Chars(j).Name) = playerName Then
                tempPlayer_ = Account(I).Chars(j)
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
    For I = 1 To MAX_CHARS
        If Trim$(tempAccount.Chars(I).Name) = playerName Then
            tempPlayer_ = tempAccount.Chars(I)
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
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SExtendedPlayerData
    
    Dim PlayerSize As Long
    Dim PlayerData() As Byte
    
    PlayerSize = LenB(requestedClientPlayer)
    ReDim PlayerData(PlayerSize - 1)
    CopyMemory PlayerData(0), ByVal VarPtr(requestedClientPlayer), PlayerSize
    Buffer.WriteBytes PlayerData
    
    SendDataTo Index, Buffer.ToArray
    
    Set Buffer = Nothing
    
    
End Sub
Sub SendMap(ByVal Index As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal Index As Long, ByVal MapNum As Integer)
    Dim I As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData
    For I = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(MapNum, I).playerName
        Buffer.WriteLong MapItem(MapNum, I).Num
        Buffer.WriteLong MapItem(MapNum, I).Value
        Buffer.WriteInteger MapItem(MapNum, I).Durability
        Buffer.WriteByte MapItem(MapNum, I).X
        Buffer.WriteByte MapItem(MapNum, I).Y
    Next

    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemToMap(ByVal MapNum As Integer, ByVal MapSlotNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SSpawnItem
    Buffer.WriteLong MapSlotNum
    Buffer.WriteString MapItem(MapNum, MapSlotNum).playerName
    Buffer.WriteLong MapItem(MapNum, MapSlotNum).Num
    Buffer.WriteLong MapItem(MapNum, MapSlotNum).Value
    Buffer.WriteInteger MapItem(MapNum, MapSlotNum).Durability
    Buffer.WriteLong MapItem(MapNum, MapSlotNum).X
    Buffer.WriteLong MapItem(MapNum, MapSlotNum).Y
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapNPCVitals(ByVal MapNum As Integer, ByVal MapNPCNum As Byte)
    Dim Buffer As clsBuffer, I As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNPCVitals
    Buffer.WriteByte MapNPCNum
    For I = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNPC(MapNum).NPC(MapNPCNum).Vital(I)
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapNPCTarget(ByVal MapNum As Integer, ByVal MapNPCNum As Byte, ByVal target As Byte, ByVal targetType As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNPCTarget
    Buffer.WriteByte MapNPCNum
    Buffer.WriteByte target
    Buffer.WriteByte targetType

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapNPCsTo(ByVal Index As Long, ByVal MapNum As Integer)
    Dim I As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNPCData

    For I = 1 To MAX_MAP_NPCS
        With MapNPC(MapNum).NPC(I)
            Buffer.WriteLong .Num
            Buffer.WriteLong .X
            Buffer.WriteLong .Y
            Buffer.WriteLong .Dir
            For X = 1 To Vitals.Vital_Count - 1
                Buffer.WriteLong .Vital(X)
            Next
        End With
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapNPCsToMap(ByVal MapNum As Integer)
    Dim I As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNPCData

    For I = 1 To MAX_MAP_NPCS
        With MapNPC(MapNum).NPC(I)
            Buffer.WriteLong .Num
            Buffer.WriteLong .X
            Buffer.WriteLong .Y
            Buffer.WriteLong .Dir
            For X = 1 To Vitals.Vital_Count - 1
                Buffer.WriteLong .Vital(X)
            Next
        End With
    Next

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMorals(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_MORALS
        If Len(Trim$(Moral(I).Name)) > 0 Then
            Call SendUpdateMoralTo(Index, I)
        End If
    Next
End Sub

Sub SendClasses(ByVal Index As Long)
    Dim I As Long
    
    For I = 1 To MAX_CLASSES
        If Len(Trim$(Class(I).Name)) > 0 Then
            Call SendUpdateClassTo(Index, I)
        End If
    Next
End Sub

Sub SendEmoticons(ByVal Index As Long)
    Dim I As Long
    
    For I = 1 To MAX_EMOTICONS
        If Len(Trim$(Emoticon(I).Command)) > 0 Then
            Call SendUpdateEmoticonTo(Index, I)
        End If
    Next
End Sub

Sub SendItems(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_ITEMS
        If Len(Trim$(Item(I).Name)) > 0 Then
            Call SendUpdateItemTo(Index, I)
        End If
    Next
End Sub

Sub SendQuests(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_QUESTS
        If Len(Trim$(Quest(I).Name)) > 0 Then
            Call SendUpdateQuestTo(Index, I)
        End If
    Next
End Sub

Sub SendTitles(ByVal Index As Long)
    Dim I As Long
    
    For I = 1 To MAX_TITLES
        If Len(Trim$(Title(I).Name)) > 0 Then
            Call SendUpdateTitleTo(Index, I)
        End If
    Next
End Sub

Sub SendAnimations(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_ANIMATIONS
        If Len(Trim$(Animation(I).Name)) > 0 Then
            Call SendUpdateAnimationTo(Index, I)
        End If
    Next
End Sub

Sub SendNPCs(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_NPCS
        If Len(Trim$(NPC(I).Name)) > 0 Then
            Call SendUpdateNPCTo(Index, I)
        End If
    Next
End Sub

Sub SendResources(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_RESOURCES
        If Len(Trim$(Resource(I).Name)) > 0 Then
            Call SendUpdateResourceTo(Index, I)
        End If
    Next
End Sub

Sub SendInventory(ByVal Index As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For I = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(Index, I)
        Buffer.WriteLong GetPlayerInvItemValue(Index, I)
        Buffer.WriteInteger GetPlayerInvItemDur(Index, I)
        Buffer.WriteByte GetPlayerInvItemBind(Index, I)
    Next

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal Index As Long, ByVal InvSlot As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteByte InvSlot
    Buffer.WriteLong GetPlayerInvItemNum(Index, InvSlot)
    Buffer.WriteLong GetPlayerInvItemValue(Index, InvSlot)
    Buffer.WriteInteger GetPlayerInvItemDur(Index, InvSlot)
    Buffer.WriteByte GetPlayerInvItemBind(Index, InvSlot)
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim I As Byte
    
    Buffer.WriteLong SPlayerWornEq
    
    For I = 1 To Equipment.Equipment_Count - 1
        Buffer.WriteLong GetPlayerEquipment(Index, I)
    Next
    
    For I = 1 To Equipment.Equipment_Count - 1
        Buffer.WriteInteger GetPlayerEquipmentDur(Index, I)
    Next
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim I As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapWornEq
    
    Buffer.WriteLong Index
    
    For I = 1 To Equipment.Equipment_Count - 1
        Buffer.WriteLong GetPlayerEquipment(Index, I)
    Next

    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim I As Byte
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong PlayerNum
    
    For I = 1 To Equipment.Equipment_Count - 1
        Buffer.WriteLong GetPlayerEquipment(PlayerNum, I)
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendVital(ByVal Index As Long, ByVal Vital As Vitals)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHP
            Buffer.WriteLong Index
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMP
            Buffer.WriteLong Index
            Buffer.WriteLong GetPlayerMaxVital(Index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(Index, Vitals.MP)
    End Select

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendVitalTo(ByVal Index As Long, player As Long, ByVal Vital As Vitals)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHP
            Buffer.WriteLong player
            Buffer.WriteLong GetPlayerMaxVital(player, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(player, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMP
            Buffer.WriteLong player
            Buffer.WriteLong GetPlayerMaxVital(player, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(player, Vitals.MP)
    End Select

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerExp(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong Index
    Buffer.WriteLong GetPlayerExp(Index)
    Buffer.WriteLong GetPlayerNextLevel(Index)
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerStats(ByVal Index As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStats
    Buffer.WriteLong Index
    
    For I = 1 To Stats.Stat_count - 1
        Buffer.WriteInteger GetPlayerStat(Index, I)
    Next
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerPoints(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerPoints
    Buffer.WriteLong Index
    Buffer.WriteInteger GetPlayerPoints(Index)
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerLevel(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerLevel
    Buffer.WriteLong Index
    Buffer.WriteByte GetPlayerLevel(Index)
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerGuild(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerGuild
    Buffer.WriteLong Index
    
    If GetPlayerGuild(Index) > 0 Then
        Buffer.WriteString Guild(GetPlayerGuild(Index)).Name
    Else
        Buffer.WriteString vbNullString
    End If
    
    Buffer.WriteByte GetPlayerGuildAccess(Index)
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendGuildInvite(ByVal Index As Long, ByVal OtherPlayer As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SGuildInvite
    
    Buffer.WriteString Trim$(Account(Index).Chars(GetPlayerChar(Index)).Name)
    Buffer.WriteString Trim$(Guild(GetPlayerGuild(Index)).Name)
    
    SendDataTo OtherPlayer, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerGuildMembers(ByVal Index As Long, Optional ByVal Ignore As Byte = 0)
    Dim I As Long
    Dim PlayerArray() As String
    Dim PlayerCount As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SGuildMembers
    
    PlayerCount = 0
    
    ' Count members online
    For I = 1 To Player_HighIndex
        If IsPlaying(I) Then
            If Not Ignore = I And Not I = Index Then
                If GetPlayerGuild(I) = GetPlayerGuild(Index) Then
                    PlayerCount = PlayerCount + 1
                    ReDim Preserve PlayerArray(1 To PlayerCount)
                    PlayerArray(UBound(PlayerArray)) = GetPlayerName(I)
                End If
            End If
        End If
    Next
    
    ' Add to Packet
    Buffer.WriteLong PlayerCount
    
    If PlayerCount > 0 Then
        For I = 1 To PlayerCount
            Buffer.WriteString PlayerArray(I)
        Next
    End If
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Sub SendPlayerSprite(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerSprite
    Buffer.WriteLong Index
    Buffer.WriteInteger GetPlayerSprite(Index)
      
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendShowTaskCompleteOnNPC(ByVal Index As Long, ByVal NPCNum As Long, ShowIt As Boolean)
    If NPCNum < 1 Or NPCNum > MAX_NPCS Then Exit Sub
    If Index < 1 Or Index > Player_HighIndex Then Exit Sub
    
    NPC(NPCNum).ShowQuestCompleteIcon = Abs(ShowIt)
    Call SendNPCs(Index)
    Call SaveNPCs
End Sub

Sub SendPlayerTitles(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerTitles
    Buffer.WriteLong Index
    
    ' Amount of titles
    Buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles
    
    ' Send player titles
    For I = 1 To Account(Index).Chars(GetPlayerChar(Index)).AmountOfTitles
        Buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).Title(I)
    Next
    
    ' Send the player's current title
    Buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).CurrentTitle
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerStatus(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerStatus
    Buffer.WriteLong Index
    
    Buffer.WriteString Account(Index).Chars(GetPlayerChar(Index)).Status
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerPK(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerPK
    Buffer.WriteLong Index
    
    Buffer.WriteByte GetPlayerPK(Index)
    
    SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
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
    Dim Buffer As clsBuffer, I As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeaveGame
    Buffer.WriteLong Index
    
    SendDataToAllBut Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateItemToAll(ByVal ItemNum As Integer)
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    Set Buffer = New clsBuffer
    ItemSize = LenB(Item(ItemNum))
    
    ReDim ItemData(ItemSize - 1)
    
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateQuestTo(ByVal Index As Long, ByVal QuestNum As Integer)
    Dim Buffer As clsBuffer
    Dim I As Long, II As Long
    
    Set Buffer = New clsBuffer
    
        Buffer.WriteLong SUpdateQuest
        
        With Quest(QuestNum)
            Buffer.WriteLong QuestNum
            Buffer.WriteString .Name
            Buffer.WriteString .Description
            Buffer.WriteLong .CanBeRetaken
            Buffer.WriteLong .Max_CLI
            
            For I = 1 To .Max_CLI
                Buffer.WriteLong .CLI(I).ItemIndex
                Buffer.WriteLong .CLI(I).isNPC
                Buffer.WriteLong .CLI(I).Max_Actions
                
                For II = 1 To .CLI(I).Max_Actions
                    Buffer.WriteString .CLI(I).Action(II).TextHolder
                    Buffer.WriteLong .CLI(I).Action(II).ActionID
                    Buffer.WriteLong .CLI(I).Action(II).Amount
                    Buffer.WriteLong .CLI(I).Action(II).MainData
                    Buffer.WriteLong .CLI(I).Action(II).QuadData
                    Buffer.WriteLong .CLI(I).Action(II).SecondaryData
                    Buffer.WriteLong .CLI(I).Action(II).TertiaryData
                Next II
            Next I
            
            Buffer.WriteLong .Requirements.AccessReq
            Buffer.WriteLong .Requirements.ClassReq
            Buffer.WriteLong .Requirements.GenderReq
            Buffer.WriteLong .Requirements.LevelReq
            Buffer.WriteLong .Requirements.SkillLevelReq
            Buffer.WriteLong .Requirements.SkillReq
            
            For I = 1 To Stats.Stat_count - 1
                Buffer.WriteLong .Requirements.Stat_Req(I)
            Next I
        End With
        
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub SendUpdateItemTo(ByVal Index As Long, ByVal ItemNum As Integer)
    Dim Buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    
    Set Buffer = New clsBuffer
    
    ItemSize = LenB(Item(ItemNum))
    ReDim ItemData(ItemSize - 1)
    CopyMemory ItemData(0), ByVal VarPtr(Item(ItemNum)), ItemSize
    Buffer.WriteLong SUpdateItem
    Buffer.WriteLong ItemNum
    Buffer.WriteBytes ItemData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationToAll(ByVal AnimationNum As Long)
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    
    Set Buffer = New clsBuffer
    
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateAnimationTo(ByVal Index As Long, ByVal AnimationNum As Long)
    Dim Buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    
    Set Buffer = New clsBuffer
    
    AnimationSize = LenB(Animation(AnimationNum))
    ReDim AnimationData(AnimationSize - 1)
    CopyMemory AnimationData(0), ByVal VarPtr(Animation(AnimationNum)), AnimationSize
    Buffer.WriteLong SUpdateAnimation
    Buffer.WriteLong AnimationNum
    Buffer.WriteBytes AnimationData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Sub SendAssociatedCharacters()

End Sub
Sub SendUpdateNPCToAll(ByVal NPCNum As Long)
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    
    Set Buffer = New clsBuffer
    
    NPCSize = LenB(NPC(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(NPCNum)), NPCSize
    Buffer.WriteLong SUpdateNPC
    Buffer.WriteLong NPCNum
    Buffer.WriteBytes NPCData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateNPCTo(ByVal Index As Long, ByVal NPCNum As Long)
    Dim Buffer As clsBuffer
    Dim NPCSize As Long
    Dim NPCData() As Byte
    
    Set Buffer = New clsBuffer
    
    NPCSize = LenB(NPC(NPCNum))
    ReDim NPCData(NPCSize - 1)
    CopyMemory NPCData(0), ByVal VarPtr(NPC(NPCNum)), NPCSize
    Buffer.WriteLong SUpdateNPC
    Buffer.WriteLong NPCNum
    Buffer.WriteBytes NPCData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceToAll(ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateResourceTo(ByVal Index As Long, ByVal ResourceNum As Long)
    Dim Buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    Set Buffer = New clsBuffer
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    CopyMemory ResourceData(0), ByVal VarPtr(Resource(ResourceNum)), ResourceSize
    
    Buffer.WriteLong SUpdateResource
    Buffer.WriteLong ResourceNum
    Buffer.WriteBytes ResourceData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendShops(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_SHOPS
        If Len(Trim$(Shop(I).Name)) > 0 Then
            Call SendUpdateShopTo(Index, I)
        End If
    Next
End Sub

Sub SendUpdateShopToAll(ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteBytes ShopData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateShopTo(ByVal Index As Long, ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    Set Buffer = New clsBuffer
    
    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    CopyMemory ShopData(0), ByVal VarPtr(Shop(ShopNum)), ShopSize
    
    Buffer.WriteLong SUpdateShop
    Buffer.WriteLong ShopNum
    Buffer.WriteBytes ShopData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpells(ByVal Index As Long)
    Dim I As Long

    For I = 1 To MAX_SPELLS
        If Len(Trim$(Spell(I).Name)) > 0 Then
            Call SendUpdateSpellTo(Index, I)
        End If
    Next
End Sub

Sub SendUpdateSpellToAll(ByVal SpellNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes SpellData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateSpellTo(ByVal Index As Long, ByVal SpellNum As Long)
    Dim Buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    
    Set Buffer = New clsBuffer
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    CopyMemory SpellData(0), ByVal VarPtr(Spell(SpellNum)), SpellSize
    
    Buffer.WriteLong SUpdateSpell
    Buffer.WriteLong SpellNum
    Buffer.WriteBytes SpellData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal Index As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SSpells
    For I = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong GetPlayerSpell(Index, I)
    Next
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpell(ByVal Index As Long, ByVal SpellSlot As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SSpell
    Buffer.WriteByte SpellSlot
    Buffer.WriteLong GetPlayerSpell(Index, SpellSlot)
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal Index As Long, ByVal Resource_Num As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(GetPlayerMap(Index)).Resource_Count

    If ResourceCache(GetPlayerMap(Index)).Resource_Count > 0 Then
        For I = 0 To ResourceCache(GetPlayerMap(Index)).Resource_Count
            Buffer.WriteByte ResourceCache(GetPlayerMap(Index)).ResourceData(I).ResourceState
            Buffer.WriteInteger ResourceCache(GetPlayerMap(Index)).ResourceData(I).X
            Buffer.WriteInteger ResourceCache(GetPlayerMap(Index)).ResourceData(I).Y
        Next
    End If

    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal MapNum As Integer, ByVal Resource_Num As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(MapNum).Resource_Count

    If ResourceCache(MapNum).Resource_Count > 0 Then
        For I = 0 To ResourceCache(MapNum).Resource_Count
            Buffer.WriteByte ResourceCache(MapNum).ResourceData(I).ResourceState
            Buffer.WriteInteger ResourceCache(MapNum).ResourceData(I).X
            Buffer.WriteInteger ResourceCache(MapNum).ResourceData(I).Y
        Next
    End If

    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendDoorAnimation(ByVal MapNum As Integer, ByVal X As Long, ByVal Y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SDoorAnimation
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendActionMsg(ByVal MapNum As Integer, ByVal Message As String, ByVal Color As Long, ByVal MsgType As Long, ByVal X As Long, ByVal Y As Long, Optional PlayerOnlyNum As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SActionMsg
    Buffer.WriteString Message
    Buffer.WriteLong Color
    Buffer.WriteLong MsgType
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    If PlayerOnlyNum > 0 Then
        SendDataTo PlayerOnlyNum, Buffer.ToArray()
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
    
    Set Buffer = Nothing
End Sub

Sub SendBlood(ByVal MapNum As Integer, ByVal X As Long, ByVal Y As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SBlood
    Buffer.WriteLong X
    Buffer.WriteLong Y
    
    SendDataToMap MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendAnimation(ByVal MapNum As Integer, ByVal Anim As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal LockType As Byte = 0, Optional ByVal LockIndex As Long = 0, Optional ByVal OnlyTo As Long = 0)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAnimation
    Buffer.WriteLong Anim
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteByte LockType
    Buffer.WriteLong LockIndex
    
    If OnlyTo > 0 Then
        SendDataTo OnlyTo, Buffer.ToArray
    Else
        SendDataToMap MapNum, Buffer.ToArray()
    End If
    Set Buffer = Nothing
End Sub

Sub SendSpellCooldown(ByVal Index As Long, ByVal Slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpellCooldown
    Buffer.WriteByte Slot
    Buffer.WriteLong GetPlayerSpellCD(Index, Slot)
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendClearAccountSpellBuffer(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal MapNum As Integer, ByVal Index As Long, ByVal Message As String, ByVal SayColor As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString Message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong SayColor
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal Index As Long, ByVal Message As String, ByVal SayColor As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    
    Buffer.WriteString GetPlayerName(Index)
    Buffer.WriteLong GetPlayerAccess(Index)
    Buffer.WriteLong GetPlayerPK(Index)
    Buffer.WriteString Message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong SayColor
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub ResetShopAction(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendStunned(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    
    Buffer.WriteLong TempPlayer(Index).StunDuration
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendBank(ByVal Index As Long)
    Dim Buffer As clsBuffer
    Dim I As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong sbank
    
    For I = 1 To MAX_BANK
        Buffer.WriteLong Account(Index).Bank.Item(I).Num
        Buffer.WriteLong Account(Index).Bank.Item(I).Value
    Next
    
    SendDataTo Index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendOpenShop(ByVal Index As Long, ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    
    Buffer.WriteLong ShopNum
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal Index As Long, ByVal movement As Byte, Optional ByVal SendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    
    Buffer.WriteLong Index
    Buffer.WriteByte GetPlayerX(Index)
    Buffer.WriteByte GetPlayerY(Index)
    Buffer.WriteByte GetPlayerDir(Index)
    Buffer.WriteByte movement
    
    If Not SendToSelf Then
        SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Else
        SendDataTo Index, Buffer.ToArray()
    End If
    Set Buffer = Nothing
End Sub

Sub SendPlayerPosition(ByVal Index As Long, Optional ByVal SendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerWarp
    
    Buffer.WriteLong Index
    Buffer.WriteByte GetPlayerX(Index)
    Buffer.WriteByte GetPlayerY(Index)
    Buffer.WriteByte GetPlayerDir(Index)
    
    If Not SendToSelf Then
        SendDataToMap GetPlayerMap(Index), Buffer.ToArray()
    Else
        SendDataTo Index, Buffer.ToArray()
    End If
    Set Buffer = Nothing
End Sub

Sub SendNPCMove(ByVal MapNPCNum As Long, ByVal movement As Byte, MapNum As Integer)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNPCMove
    
    Buffer.WriteLong MapNPCNum
    Buffer.WriteByte MapNPC(MapNum).NPC(MapNPCNum).X
    Buffer.WriteByte MapNPC(MapNum).NPC(MapNPCNum).Y
    Buffer.WriteByte MapNPC(MapNum).NPC(MapNPCNum).Dir
    Buffer.WriteByte movement
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTrade(ByVal Index As Long, ByVal TradeTarget As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    
    Buffer.WriteLong TradeTarget
    Buffer.WriteString Trim$(GetPlayerName(TradeTarget))
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal Index As Long, ByVal DataType As Byte)
    Dim Buffer As clsBuffer
    Dim I As Long
    Dim TradeTarget As Long
    Dim TotalWorth As Long
    
    TradeTarget = TempPlayer(Index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    
    Buffer.WriteByte DataType
    If DataType = 0 Then ' own inventory
        For I = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(Index).TradeOffer(I).Num
            Buffer.WriteLong TempPlayer(Index).TradeOffer(I).Value
            ' add total worth
            If TempPlayer(Index).TradeOffer(I).Num > 0 Then
                 If GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(I).Num) > 0 Then
                    ' currency?
                    If Item(TempPlayer(Index).TradeOffer(I).Num).Stackable = 1 Then
                        TotalWorth = TotalWorth + (Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(I).Num)).Price * TempPlayer(Index).TradeOffer(I).Value)
                    Else
                        TotalWorth = TotalWorth + Item(GetPlayerInvItemNum(Index, TempPlayer(Index).TradeOffer(I).Num)).Price
                    End If
                End If
            End If
        Next
    ElseIf DataType = 1 Then ' other inventory
        For I = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(I).Num)
            Buffer.WriteLong TempPlayer(TradeTarget).TradeOffer(I).Value
            ' add total worth
            If GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(I).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(I).Num)).Stackable = 1 Then
                    TotalWorth = TotalWorth + (Item(GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(I).Num)).Price * TempPlayer(TradeTarget).TradeOffer(I).Value)
                Else
                    TotalWorth = TotalWorth + Item(GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(I).Num)).Price
                End If
            End If
        Next
    End If
    ' Send total worth of trade
    Buffer.WriteLong TotalWorth
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal Index As Long, ByVal Status As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    
    Buffer.WriteByte Status
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAttack(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAttack
    Buffer.WriteLong Index
    
    SendDataToMapBut Index, GetPlayerMap(Index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerTarget(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    
    Buffer.WriteByte TempPlayer(Index).target
    Buffer.WriteByte TempPlayer(Index).targetType
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHotbar(ByVal Index As Long)
    Dim I As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    
    For I = 1 To MAX_HOTBAR
        Buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).Hotbar(I).Slot
        Buffer.WriteByte Account(Index).Chars(GetPlayerChar(Index)).Hotbar(I).SType
    Next
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNPCSpellBuffer(MapNum, MapNPCNum)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNPCSpellBuffer
    
    Buffer.WriteLong MapNPCNum
    Buffer.WriteLong MapNPC(MapNum).NPC(MapNPCNum).SpellBuffer.Spell
    
    Call SendDataToMap(MapNum, Buffer.ToArray)
    Set Buffer = Nothing
End Sub

Sub SendLogs(ByVal Index As Long, Msg As String, Name As String)
    Dim Buffer As clsBuffer
    Dim LogSize As Long
    Dim LogData() As Byte
    
    Set Buffer = New clsBuffer
    
    Log.Msg = Msg
    Log.File = Name
    LogSize = LenB(Log)
    ReDim LogData(LogSize - 1)
    CopyMemory LogData(0), ByVal VarPtr(Log), LogSize
    Buffer.WriteLong SUpdateLogs
    
    Buffer.WriteBytes LogData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub UpdateFriendsList(Index)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim I As Long, n As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SFriendsList
    
    If Account(Index).Friends.AmountOfFriends = 0 Then
        Buffer.WriteByte Account(Index).Friends.AmountOfFriends
        GoTo Finish
    End If
   
    ' Sends the amount of friends in friends list
    Buffer.WriteByte Account(Index).Friends.AmountOfFriends
   
    ' Check to see if they are online
    For I = 1 To Account(Index).Friends.AmountOfFriends
        Name = Trim$(Account(Index).Friends.Members(I))
        Buffer.WriteString Name
        For n = 1 To Player_HighIndex
            If IsPlaying(FindPlayer(Name)) Then
                Buffer.WriteString Name & " Online"
            Else
                Buffer.WriteString Name & " Offline"
            End If
        Next
    Next
    
Finish:
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub UpdateFoesList(Index)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim I As Long, n As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SFoesList
    
    If Account(Index).Foes.Amount = 0 Then
        Buffer.WriteByte Account(Index).Foes.Amount
        GoTo Finish
    End If
   
    ' Sends the amount of Foes in Foes list
    Buffer.WriteByte Account(Index).Foes.Amount
   
    ' Check to see if they are online
    For I = 1 To Account(Index).Foes.Amount
        Name = Trim$(Account(Index).Foes.Members(I))
        Buffer.WriteString Name
        For n = 1 To Player_HighIndex
            If IsPlaying(FindPlayer(Name)) Then
                Buffer.WriteString Name & " Online"
            Else
                Buffer.WriteString Name & " Offline"
            End If
        Next
    Next
    
Finish:
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInGame(ByVal Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayer_HighIndex()
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SHighIndex
    
    Buffer.WriteLong Player_HighIndex
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSoundTo(ByVal Index As Integer, Sound As String)
    Dim Buffer As clsBuffer
    
    ' Don't send it if there's nothing to send
    If Sound = vbNullString Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    
    Buffer.WriteString Sound
    SendDataTo Index, Buffer.ToArray()
End Sub

Sub SendSoundToMap(ByVal MapNum As Integer, Sound As String)
    Dim Buffer As clsBuffer
    
    ' Don't send it if there's nothing to send
    If Sound = vbNullString Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    
    Buffer.WriteString Sound
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendSoundToAll(ByVal MapNum As Integer, Sound As String)
    Dim Buffer As clsBuffer
    
    ' Don't send it if there's nothing to send
    If Sound = vbNullString Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    
    Buffer.WriteString Sound
    SendDataToAll Buffer.ToArray()
End Sub

Sub SendPlayerSound(ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal EntityType As Long, ByVal EntityNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEntitySound
    
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong EntityType
    Buffer.WriteLong EntityNum
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapSound(ByVal MapNum As Integer, ByVal Index As Long, ByVal X As Long, ByVal Y As Long, ByVal EntityType As Long, ByVal EntityNum As Long)
    Dim Buffer As clsBuffer

    If EntityNum <= 0 Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SEntitySound
    
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong EntityType
    Buffer.WriteLong EntityNum
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNPCDeath(ByVal MapNPCNum As Long, MapNum As Integer)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNPCDead
    
    Buffer.WriteLong MapNPCNum
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNPCAttack(ByVal Attacker As Long, MapNum As Integer)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNPCAttack
    
    Buffer.WriteLong Attacker
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendLogin(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    
    Buffer.WriteLong Index
    Buffer.WriteLong Player_HighIndex
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewCharClasses
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNews(ByVal Index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendNews
    
    Buffer.WriteString Options.News
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendCheckForMap(ByVal Index As Long, MapNum As Integer)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    
    Buffer.WriteInteger MapNum
    Buffer.WriteInteger Map(MapNum).Revision
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal Index As Long, ByVal TradeRequest As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    
    Buffer.WriteString Trim$(GetPlayerName(TradeRequest))
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal Index As Long, ByVal OtherPlayer As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    
    Buffer.WriteString Trim$(Account(OtherPlayer).Chars(GetPlayerChar(OtherPlayer)).Name)
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal PartyNum As Long)
    Dim Buffer As clsBuffer, I As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    Buffer.WriteByte 1
    For I = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(PartyNum).Member(I)
    Next
    Buffer.WriteLong Party(PartyNum).MemberCount
    Buffer.WriteLong PartyNum
    
    SendDataToParty PartyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal Index As Long)
    Dim Buffer As clsBuffer, I As Long, PartyNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' Check if we're in a party
    PartyNum = TempPlayer(Index).InParty
    
    If PartyNum > 0 Then
        ' Send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(PartyNum).Leader
        For I = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(PartyNum).Member(I)
        Next
        Buffer.WriteLong Party(PartyNum).MemberCount
    Else
        ' Send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal PartyNum As Long, ByVal Index As Long)
    Dim Buffer As clsBuffer, I As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    
    Buffer.WriteLong Index
    For I = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(Index, I)
        Buffer.WriteLong Account(Index).Chars(GetPlayerChar(Index)).Vital(I)
    Next
    
    SendDataToParty PartyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateTitleToAll(ByVal TitleNum As Long)
    Dim Buffer As clsBuffer
    Dim TitleSize As Long
    Dim TitleData() As Byte
    
    Set Buffer = New clsBuffer
    
    TitleSize = LenB(Title(TitleNum))
    ReDim TitleData(TitleSize - 1)
    CopyMemory TitleData(0), ByVal VarPtr(Title(TitleNum)), TitleSize
    Buffer.WriteLong SUpdateTitle
    
    Buffer.WriteLong TitleNum
    Buffer.WriteBytes TitleData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateTitleTo(ByVal Index As Long, ByVal TitleNum As Long)
    Dim Buffer As clsBuffer
    Dim TitleSize As Long
    Dim TitleData() As Byte
    
    Set Buffer = New clsBuffer
    
    TitleSize = LenB(Title(TitleNum))
    ReDim TitleData(TitleSize - 1)
    CopyMemory TitleData(0), ByVal VarPtr(Title(TitleNum)), TitleSize
    Buffer.WriteLong SUpdateTitle
    
    Buffer.WriteLong TitleNum
    Buffer.WriteBytes TitleData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendCloseClient(Index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseClient
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateMoralToAll(ByVal MoralNum As Long)
    Dim Buffer As clsBuffer
    Dim MoralSize As Long
    Dim MoralData() As Byte
    
    Set Buffer = New clsBuffer
    
    MoralSize = LenB(Moral(MoralNum))
    ReDim MoralData(MoralSize - 1)
    CopyMemory MoralData(0), ByVal VarPtr(Moral(MoralNum)), MoralSize
    Buffer.WriteLong SUpdateMoral
    
    Buffer.WriteLong MoralNum
    Buffer.WriteBytes MoralData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateMoralTo(ByVal Index As Long, ByVal MoralNum As Long)
    Dim Buffer As clsBuffer
    Dim MoralSize As Long
    Dim MoralData() As Byte
    
    Set Buffer = New clsBuffer
    
    MoralSize = LenB(Moral(MoralNum))
    ReDim MoralData(MoralSize - 1)
    CopyMemory MoralData(0), ByVal VarPtr(Moral(MoralNum)), MoralSize
    Buffer.WriteLong SUpdateMoral
    
    Buffer.WriteLong MoralNum
    Buffer.WriteBytes MoralData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateClassTo(ByVal Index As Long, ByVal ClassNum As Long)
    Dim Buffer As clsBuffer
    Dim Classesize As Long
    Dim ClassData() As Byte
    
    Set Buffer = New clsBuffer
    
    Classesize = LenB(Class(ClassNum))
    ReDim ClassData(Classesize - 1)
    CopyMemory ClassData(0), ByVal VarPtr(Class(ClassNum)), Classesize
    Buffer.WriteLong SUpdateClass
    
    Buffer.WriteLong ClassNum
    Buffer.WriteBytes ClassData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateEmoticonTo(ByVal Index As Long, ByVal EmoticonNum As Long)
    Dim Buffer As clsBuffer
    Dim EmoticonSize As Long
    Dim EmoticonData() As Byte
    
    Set Buffer = New clsBuffer

    EmoticonSize = LenB(Emoticon(EmoticonNum))
    ReDim EmoticonData(EmoticonSize - 1)
    CopyMemory EmoticonData(0), ByVal VarPtr(Emoticon(EmoticonNum)), EmoticonSize
    Buffer.WriteLong SUpdateEmoticon
    
    Buffer.WriteLong EmoticonNum
    Buffer.WriteBytes EmoticonData
    
    SendDataTo Index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateEmoticonToAll(ByVal EmoticonNum As Integer)
    Dim Buffer As clsBuffer
    Dim EmoticonSize As Long
    Dim EmoticonData() As Byte
    Set Buffer = New clsBuffer
    EmoticonSize = LenB(Emoticon(EmoticonNum))
    
    ReDim EmoticonData(EmoticonSize - 1)
    
    CopyMemory EmoticonData(0), ByVal VarPtr(Emoticon(EmoticonNum)), EmoticonSize
    
    Buffer.WriteLong SUpdateEmoticon
    Buffer.WriteLong EmoticonNum
    Buffer.WriteBytes EmoticonData
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendCheckEmoticon(ByVal Index As Long, ByVal MapNum As Long, ByVal EmoticonNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckEmoticon
    
    Buffer.WriteLong Index
    Buffer.WriteLong EmoticonNum
    
    SendDataToMap MapNum, Buffer.ToArray()
End Sub

Sub SendChatBubble(ByVal MapNum As Long, ByVal target As Long, ByVal targetType As Long, ByVal Message As String, ByVal Color As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SChatBubble
    
    Buffer.WriteLong target
    Buffer.WriteLong targetType
    Buffer.WriteString Message
    Buffer.WriteLong Color
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendSpecialEffect(ByVal Index As Long, EffectType As Long, Optional Data1 As Long = 0, Optional Data2 As Long = 0, Optional Data3 As Long = 0, Optional Data4 As Long = 0)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpecialEffect
    
    Select Case EffectType
        Case EFFECT_TYPE_FADEIN
            Buffer.WriteLong EffectType
        Case EFFECT_TYPE_FADEOUT
            Buffer.WriteLong EffectType
        Case EFFECT_TYPE_FLASH
            Buffer.WriteLong EffectType
        Case EFFECT_TYPE_FOG
            Buffer.WriteLong EffectType
            Buffer.WriteLong Data1 ' Fog num
            Buffer.WriteLong Data2 ' Fog movement speed
            Buffer.WriteLong Data3 ' Opacity
        Case EFFECT_TYPE_WEATHER
            Buffer.WriteLong EffectType
            Buffer.WriteLong Data1 ' Weather type
            Buffer.WriteLong Data2 ' Weather intensity
        Case EFFECT_TYPE_TINT
            Buffer.WriteLong EffectType
            Buffer.WriteLong Data1 ' Red
            Buffer.WriteLong Data2 ' Green
            Buffer.WriteLong Data3 ' Blue
            Buffer.WriteLong Data4 ' Alpha
    End Select
    
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing
End Sub
