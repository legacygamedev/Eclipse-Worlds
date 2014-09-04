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
    Dim Buffer As clsBuffer
    Dim TempData() As Byte
    
    If IsConnected(index) Then
        Set Buffer = New clsBuffer
        TempData = Data
        
        Buffer.PreAllocate 4 + (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteLong (UBound(TempData) - LBound(TempData)) + 1
        Buffer.WriteBytes TempData()
        
        If IsConnected(index) Then
            frmServer.Socket(index).SendData Buffer.ToArray()
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
    Dim i As Long
    Dim LogMsg As String
    
    Set Buffer = New clsBuffer
    
    ' Add server log
    Call AddLog(Msg, "Player")
    
    LogMsg = Msg
    Msg = "[Admin] " & Msg
    
    ' Prevent hacking
    For i = 1 To Len(Msg)

        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then

            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then

                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If

    Next
    
    Buffer.WriteLong SAdminMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerAccess(i) >= STAFF_MODERATOR Then
            SendDataTo i, Buffer.ToArray
            Call SendLogs(i, LogMsg, "Admin")
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Public Sub PlayerMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long, Optional ByVal QuestMsg As Boolean = False, Optional ByVal QuestNum As Long = 0)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    
    ' Prevent hacking
    For i = 1 To Len(Msg)

        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then

            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then

                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If

    Next
    
    Buffer.WriteLong SPlayerMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    Buffer.WriteLong QuestMsg
    Buffer.WriteLong QuestNum
    SendDataTo index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub MapMsg(ByVal MapNum As Integer, ByVal Msg As String, ByVal Color As Long)
    Dim Buffer As clsBuffer
    Dim i As Long

    Set Buffer = New clsBuffer

    ' Prevent hacking
    For i = 1 To Len(Msg)

        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then

            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then

                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If

    Next
    
    Buffer.WriteLong SMapMsg
    Buffer.WriteString Msg
    Buffer.WriteByte Color
    SendDataToMap MapNum, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Public Sub GuildMsg(ByVal index As Long, ByVal Msg As String, ByVal Color As Long, Optional HideName As Boolean = False)
    Dim i As Long
    Dim LogMsg As String
    
    ' Prevent hacking
    For i = 1 To Len(Msg)

        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then

            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then

                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If

    Next
    
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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Buffer.WriteLong SAlertMsg
    Buffer.WriteString Msg
    SendDataTo index, Buffer.ToArray
    DoEvents
    
    Set Buffer = Nothing
End Sub

Public Sub PartyMsg(ByVal PartyNum As Long, ByVal Msg As String, ByVal Color As Long)
    Dim i As Long
    Dim LogMsg As String
    
    ' Prevent hacking
    For i = 1 To Len(Msg)

        ' limit the ASCII
        If AscW(Mid$(Msg, i, 1)) < 32 Or AscW(Mid$(Msg, i, 1)) > 126 Then

            ' limit the extended ASCII
            If AscW(Mid$(Msg, i, 1)) < 128 Or AscW(Mid$(Msg, i, 1)) > 168 Then

                ' limit the extended ASCII
                If AscW(Mid$(Msg, i, 1)) < 224 Or AscW(Mid$(Msg, i, 1)) > 253 Then
                    Mid$(Msg, i, 1) = ""
                End If
            End If
        End If

    Next
    
    ' Add server log
    Call AddLog(Msg, "Player")
    
    LogMsg = Msg
    
    Msg = "[Party] " & Msg
    
    ' Send message to all people
    For i = 1 To MAX_PARTY_MEMBERS
        ' Exist?
        If Party(PartyNum).Member(i) > 0 Then
            ' Make sure they're logged on
            If IsPlaying(Party(PartyNum).Member(i)) Then
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
    Dim Buffer() As Byte
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
    frmServer.Socket(index).GetData Buffer(), vbUnicode, DataLength
    TempPlayer(index).Buffer.WriteBytes Buffer()
    
    If TempPlayer(index).Buffer.Length >= 4 Then
        pLength = TempPlayer(index).Buffer.ReadLong(False)
    
        If pLength < 0 Then Exit Sub
    End If
    
    Do While pLength > 0 And pLength <= TempPlayer(index).Buffer.Length - 4
        If pLength <= TempPlayer(index).Buffer.Length - 4 Then
            TempPlayer(index).DataPackets = TempPlayer(index).DataPackets + 1
            TempPlayer(index).Buffer.ReadLong
            HandleData index, TempPlayer(index).Buffer.ReadBytes(pLength)
        End If
        
        pLength = 0
        
        If TempPlayer(index).Buffer.Length >= 4 Then
            pLength = TempPlayer(index).Buffer.ReadLong(False)
        
            If pLength < 0 Then Exit Sub
        End If
    Loop
            
    TempPlayer(index).Buffer.Trim
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
                For i = 1 To MapLayer.Layer_Count - 1
                    Buffer.WriteLong .Layer(i).X
                    Buffer.WriteLong .Layer(i).Y
                    Buffer.WriteLong .Layer(i).Tileset
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
'Character Editor
Sub SendPlayersOnline(ByVal index As Long)
    Dim Buffer As clsBuffer, i As Long
    Dim list As String

    If index > Player_HighIndex Or index < 1 Then Exit Sub
    Set Buffer = New clsBuffer
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
                If i <> Player_HighIndex Then
                    list = list & GetPlayerName(i) & ":" & Account(i).Chars(GetPlayerChar(i)).Access & ":" & Account(i).Chars(GetPlayerChar(i)).Sprite & ", "
                Else
                    list = list & GetPlayerName(i) & ":" & Account(i).Chars(GetPlayerChar(i)).Access & ":" & Account(i).Chars(GetPlayerChar(i)).Sprite
                End If
        End If
    Next
    
    Buffer.WriteLong SPlayersOnline
    Buffer.WriteString list
 
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

'Character Editor
Sub SendAllCharacters(index As Long, Optional everyone As Boolean = False)
    Dim Buffer As clsBuffer, i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAllCharacters
    
    Buffer.WriteString GetCharList
    
    SendDataTo index, Buffer.ToArray
    
    Set Buffer = Nothing
End Sub

Function PlayerData(ByVal index As Long) As Byte()
    Dim Buffer As clsBuffer, i As Long

    If index < 1 Or index > Player_HighIndex Then Exit Function
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerData
    Buffer.WriteLong index
    Buffer.WriteInteger Account(index).Chars(GetPlayerChar(index)).Face
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteByte GetPlayerGender(index)
    Buffer.WriteByte GetPlayerClass(index)
    Buffer.WriteByte GetPlayerLevel(index)
    Buffer.WriteInteger GetPlayerPoints(index)
    Buffer.WriteInteger GetPlayerSprite(index)
    Buffer.WriteInteger GetPlayerMap(index)
    Buffer.WriteByte GetPlayerX(index)
    Buffer.WriteByte GetPlayerY(index)
    Buffer.WriteByte GetPlayerDir(index)
    Buffer.WriteByte GetPlayerAccess(index)
    Buffer.WriteByte GetPlayerPK(index)
    
    If GetPlayerGuild(index) > 0 Then
        Buffer.WriteString Guild(GetPlayerGuild(index)).Name
    Else
        Buffer.WriteString vbNullString
    End If
    
    Buffer.WriteByte GetPlayerGuildAccess(index)
    
    For i = 1 To Stats.Stat_count - 1
        Buffer.WriteInteger GetPlayerStat(index, i)
    Next
    
    ' Amount of titles
    Buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
    
    ' Send player titles
    For i = 1 To Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
        Buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Title(i)
    Next
    
    ' Send the player's current title
    Buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).CurrentTitle
    
    ' Send player status
    Buffer.WriteString Account(index).Chars(GetPlayerChar(index)).Status
    
    For i = 1 To Skills.Skill_Count - 1
        Buffer.WriteByte GetPlayerSkill(index, i)
        Buffer.WriteLong GetPlayerSkillExp(index, i)
    Next
    
    For i = 1 To MAX_QUESTS
        Buffer.WriteLong GetPlayerQuestCLIID(index, i)
        Buffer.WriteLong GetPlayerQuestTaskID(index, i)
        Buffer.WriteLong GetPlayerQuestAmount(index, i)
        Buffer.WriteLong IsQuestCompleted(index, i)
    Next i
    
    PlayerData = Buffer.ToArray()
    Set Buffer = Nothing
End Function

Sub SendJoinMap(ByVal index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    ' Send all players on current map to index
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If i <> index Then
                If GetPlayerMap(i) = GetPlayerMap(index) Then
                    SendDataTo index, PlayerData(i)
                End If
            End If
        End If
    Next
    
    ' Send index's player data to everyone on the map including themself
    SendDataToMap GetPlayerMap(index), PlayerData(index)
    
    ' Send the NPC targets to the player
    For i = 1 To Map(GetPlayerMap(index)).NPC_HighIndex
        If MapNPC(GetPlayerMap(index)).NPC(i).Num > 0 Then
            Call SendMapNPCTarget(GetPlayerMap(index), i, MapNPC(GetPlayerMap(index)).NPC(i).target, MapNPC(GetPlayerMap(index)).NPC(i).targetType)
        Else
            ' Send 0 so it uncaches any old data
            Call SendMapNPCTarget(GetPlayerMap(index), i, 0, 0)
        End If
    Next
    
    Set Buffer = Nothing
End Sub

Sub SendLeaveMap(ByVal index As Long, ByVal MapNum As Integer)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeft
    Buffer.WriteLong index
    SendDataToMapBut index, MapNum, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendPlayerData(ByVal index As Long)
    SendDataToMap GetPlayerMap(index), PlayerData(index)
End Sub

Sub SendAccessVerificator(ByVal index As Long, success As Byte, Message As String, currentAccess As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SAccessVerificator
    Buffer.WriteByte success
    Buffer.WriteString Message
    Buffer.WriteByte currentAccess
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

'Character Editor
Sub SendExtendedPlayerData(index As Long, playerName As String)
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
    
    filename = App.path & "\data\accounts\" & charLogin(0) & "\data.bin"
    Call ChkDir(App.path & "\data\accounts\", charLogin(0))
    
    F = FreeFile
    
    If Not FileExist(filename, True) Then
        ' Erase that char name
        Call DeleteName(playerName)
        Call PlayerMsg(index, "This character doesn't exist and has been wiped from charlist.txt.", BrightRed)
        Call SendRefreshCharEditor(index)
        Exit Sub
    End If
    
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
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SExtendedPlayerData
    
    Dim PlayerSize As Long
    Dim PlayerData() As Byte
    
    PlayerSize = LenB(requestedClientPlayer)
    ReDim PlayerData(PlayerSize - 1)
    CopyMemory PlayerData(0), ByVal VarPtr(requestedClientPlayer), PlayerSize
    Buffer.WriteBytes PlayerData
    
    SendDataTo index, Buffer.ToArray
    
    Set Buffer = Nothing
    
    
End Sub
Sub SendMap(ByVal index As Long, ByVal MapNum As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.PreAllocate (UBound(MapCache(MapNum).Data) - LBound(MapCache(MapNum).Data)) + 5
    Buffer.WriteLong SMapData
    Buffer.WriteBytes MapCache(MapNum).Data()
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendMapItemsTo(ByVal index As Long, ByVal MapNum As Integer)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapItemData
    For i = 1 To MAX_MAP_ITEMS
        Buffer.WriteString MapItem(MapNum, i).playerName
        Buffer.WriteLong MapItem(MapNum, i).Num
        Buffer.WriteLong MapItem(MapNum, i).Value
        Buffer.WriteInteger MapItem(MapNum, i).Durability
        Buffer.WriteByte MapItem(MapNum, i).X
        Buffer.WriteByte MapItem(MapNum, i).Y
    Next

    SendDataTo index, Buffer.ToArray()
    
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
    Dim Buffer As clsBuffer, i As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNPCVitals
    Buffer.WriteByte MapNPCNum
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong MapNPC(MapNum).NPC(MapNPCNum).Vital(i)
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

Sub SendMapNPCsTo(ByVal index As Long, ByVal MapNum As Integer)
    Dim i As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNPCData

    For i = 1 To MAX_MAP_NPCS
        With MapNPC(MapNum).NPC(i)
            Buffer.WriteLong .Num
            Buffer.WriteLong .X
            Buffer.WriteLong .Y
            Buffer.WriteLong .Dir
            For X = 1 To Vitals.Vital_Count - 1
                Buffer.WriteLong .Vital(X)
            Next
        End With
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapNPCsToMap(ByVal MapNum As Integer)
    Dim i As Long, X As Long
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapNPCData

    For i = 1 To MAX_MAP_NPCS
        With MapNPC(MapNum).NPC(i)
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

Sub SendQuests(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_QUESTS
        If Len(Trim$(Quest(i).Name)) > 0 Then
            Call SendUpdateQuestTo(index, i)
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

Sub SendNPCs(ByVal index As Long)
    Dim i As Long

    For i = 1 To MAX_NPCS
        If Len(Trim$(NPC(i).Name)) > 0 Then
            Call SendUpdateNPCTo(index, i)
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
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInv

    For i = 1 To MAX_INV
        Buffer.WriteLong GetPlayerInvItemNum(index, i)
        Buffer.WriteLong GetPlayerInvItemValue(index, i)
        Buffer.WriteInteger GetPlayerInvItemDur(index, i)
        Buffer.WriteByte GetPlayerInvItemBind(index, i)
    Next

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInventoryUpdate(ByVal index As Long, ByVal InvSlot As Byte)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerInvUpdate
    Buffer.WriteByte InvSlot
    Buffer.WriteLong GetPlayerInvItemNum(index, InvSlot)
    Buffer.WriteLong GetPlayerInvItemValue(index, InvSlot)
    Buffer.WriteInteger GetPlayerInvItemDur(index, InvSlot)
    Buffer.WriteByte GetPlayerInvItemBind(index, InvSlot)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendWornEquipment(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    Dim i As Byte
    
    Buffer.WriteLong SPlayerWornEq
    
    For i = 1 To Equipment.Equipment_Count - 1
        Buffer.WriteLong GetPlayerEquipment(index, i)
    Next
    
    For i = 1 To Equipment.Equipment_Count - 1
        Buffer.WriteInteger GetPlayerEquipmentDur(index, i)
    Next
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapEquipment(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Byte
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapWornEq
    
    Buffer.WriteLong index
    
    For i = 1 To Equipment.Equipment_Count - 1
        Buffer.WriteLong GetPlayerEquipment(index, i)
    Next

    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapEquipmentTo(ByVal PlayerNum As Long, ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Byte
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SMapWornEq
    Buffer.WriteLong PlayerNum
    
    For i = 1 To Equipment.Equipment_Count - 1
        Buffer.WriteLong GetPlayerEquipment(PlayerNum, i)
    Next
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendVital(ByVal index As Long, ByVal Vital As Vitals)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer

    Select Case Vital
        Case HP
            Buffer.WriteLong SPlayerHP
            Buffer.WriteLong index
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.HP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.HP)
        Case MP
            Buffer.WriteLong SPlayerMP
            Buffer.WriteLong index
            Buffer.WriteLong GetPlayerMaxVital(index, Vitals.MP)
            Buffer.WriteLong GetPlayerVital(index, Vitals.MP)
    End Select

    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendVitalTo(ByVal index As Long, player As Long, ByVal Vital As Vitals)
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

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerExp(ByVal index As Long)
    Dim Buffer As clsBuffer
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerEXP
    Buffer.WriteLong index
    Buffer.WriteLong GetPlayerExp(index)
    Buffer.WriteLong GetPlayerNextLevel(index)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerStats(ByVal index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerStats
    Buffer.WriteLong index
    
    For i = 1 To Stats.Stat_count - 1
        Buffer.WriteInteger GetPlayerStat(index, i)
    Next
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerPoints(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerPoints
    Buffer.WriteLong index
    Buffer.WriteInteger GetPlayerPoints(index)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerLevel(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerLevel
    Buffer.WriteLong index
    Buffer.WriteByte GetPlayerLevel(index)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerGuild(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerGuild
    Buffer.WriteLong index
    
    If GetPlayerGuild(index) > 0 Then
        Buffer.WriteString Guild(GetPlayerGuild(index)).Name
    Else
        Buffer.WriteString vbNullString
    End If
    
    Buffer.WriteByte GetPlayerGuildAccess(index)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendGuildInvite(ByVal index As Long, ByVal OtherPlayer As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SGuildInvite
    
    Buffer.WriteString Trim$(Account(index).Chars(GetPlayerChar(index)).Name)
    Buffer.WriteString Trim$(Guild(GetPlayerGuild(index)).Name)
    
    SendDataTo OtherPlayer, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerGuildMembers(ByVal index As Long, Optional ByVal Ignore As Byte = 0)
    Dim i As Long
    Dim PlayerArray() As String
    Dim PlayerCount As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SGuildMembers
    
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
    Buffer.WriteLong PlayerCount
    
    If PlayerCount > 0 Then
        For i = 1 To PlayerCount
            Buffer.WriteString PlayerArray(i)
        Next
    End If
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub
Sub SendPlayerSprite(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerSprite
    Buffer.WriteLong index
    Buffer.WriteInteger GetPlayerSprite(index)
      
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendShowTaskCompleteOnNPC(ByVal index As Long, ByVal NPCNum As Long, ShowIt As Boolean)
    If NPCNum < 1 Or NPCNum > MAX_NPCS Then Exit Sub
    If index < 1 Or index > Player_HighIndex Then Exit Sub
    
    NPC(NPCNum).ShowQuestCompleteIcon = Abs(ShowIt)
    Call SendNPCs(index)
    Call SaveNPCs
End Sub

Sub SendPlayerTitles(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerTitles
    Buffer.WriteLong index
    
    ' Amount of titles
    Buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
    
    ' Send player titles
    For i = 1 To Account(index).Chars(GetPlayerChar(index)).AmountOfTitles
        Buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Title(i)
    Next
    
    ' Send the player's current title
    Buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).CurrentTitle
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerStatus(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerStatus
    Buffer.WriteLong index
    
    Buffer.WriteString Account(index).Chars(GetPlayerChar(index)).Status
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerPK(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SPlayerPK
    Buffer.WriteLong index
    
    Buffer.WriteByte GetPlayerPK(index)
    
    SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
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
    Dim Buffer As clsBuffer, i As Long
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SLeaveGame
    Buffer.WriteLong index
    
    SendDataToAllBut index, Buffer.ToArray()
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

Sub SendUpdateQuestTo(ByVal index As Long, ByVal QuestNum As Integer)
    Dim Buffer As clsBuffer
    Dim i As Long, II As Long
    
    Set Buffer = New clsBuffer
    
        Buffer.WriteLong SUpdateQuest
        
        With Quest(QuestNum)
            Buffer.WriteLong QuestNum
            Buffer.WriteString .Name
            Buffer.WriteString .Description
            Buffer.WriteLong .CanBeRetaken
            Buffer.WriteLong .Max_CLI
            
            For i = 1 To .Max_CLI
                Buffer.WriteLong .CLI(i).ItemIndex
                Buffer.WriteLong .CLI(i).isNPC
                Buffer.WriteLong .CLI(i).Max_Actions
                
                For II = 1 To .CLI(i).Max_Actions
                    Buffer.WriteString .CLI(i).Action(II).TextHolder
                    Buffer.WriteLong .CLI(i).Action(II).ActionID
                    Buffer.WriteLong .CLI(i).Action(II).Amount
                    Buffer.WriteLong .CLI(i).Action(II).MainData
                    Buffer.WriteLong .CLI(i).Action(II).QuadData
                    Buffer.WriteLong .CLI(i).Action(II).SecondaryData
                    Buffer.WriteLong .CLI(i).Action(II).TertiaryData
                Next II
            Next i
            
            Buffer.WriteLong .Requirements.AccessReq
            Buffer.WriteLong .Requirements.ClassReq
            Buffer.WriteLong .Requirements.GenderReq
            Buffer.WriteLong .Requirements.LevelReq
            Buffer.WriteLong .Requirements.SkillLevelReq
            Buffer.WriteLong .Requirements.SkillReq
            
            For i = 1 To Stats.Stat_count - 1
                Buffer.WriteLong .Requirements.Stat_Req(i)
            Next i
        End With
        
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing

End Sub

Sub SendUpdateItemTo(ByVal index As Long, ByVal ItemNum As Integer)
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
    
    SendDataTo index, Buffer.ToArray()
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

Sub SendUpdateAnimationTo(ByVal index As Long, ByVal AnimationNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
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

Sub SendUpdateNPCTo(ByVal index As Long, ByVal NPCNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
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

Sub SendUpdateResourceTo(ByVal index As Long, ByVal ResourceNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
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

Sub SendUpdateShopTo(ByVal index As Long, ByVal ShopNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
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

Sub SendUpdateSpellTo(ByVal index As Long, ByVal SpellNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpells(ByVal index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SSpells
    For i = 1 To MAX_PLAYER_SPELLS
        Buffer.WriteLong GetPlayerSpell(index, i)
    Next
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerSpell(ByVal index As Long, ByVal SpellSlot As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SSpell
    Buffer.WriteByte SpellSlot
    Buffer.WriteLong GetPlayerSpell(index, SpellSlot)
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheTo(ByVal index As Long, ByVal Resource_Num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(GetPlayerMap(index)).Resource_Count

    If ResourceCache(GetPlayerMap(index)).Resource_Count > 0 Then
        For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
            Buffer.WriteByte ResourceCache(GetPlayerMap(index)).ResourceData(i).ResourceState
            Buffer.WriteInteger ResourceCache(GetPlayerMap(index)).ResourceData(i).X
            Buffer.WriteInteger ResourceCache(GetPlayerMap(index)).ResourceData(i).Y
        Next
    End If

    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendResourceCacheToMap(ByVal MapNum As Integer, ByVal Resource_Num As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    Set Buffer = New clsBuffer
    
    Buffer.WriteLong SResourceCache
    Buffer.WriteLong ResourceCache(MapNum).Resource_Count

    If ResourceCache(MapNum).Resource_Count > 0 Then
        For i = 0 To ResourceCache(MapNum).Resource_Count
            Buffer.WriteByte ResourceCache(MapNum).ResourceData(i).ResourceState
            Buffer.WriteInteger ResourceCache(MapNum).ResourceData(i).X
            Buffer.WriteInteger ResourceCache(MapNum).ResourceData(i).Y
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

Sub SendSpellCooldown(ByVal index As Long, ByVal Slot As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSpellCooldown
    Buffer.WriteByte Slot
    Buffer.WriteLong GetPlayerSpellCD(index, Slot)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendClearAccountSpellBuffer(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SClearSpellBuffer
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SayMsg_Map(ByVal MapNum As Integer, ByVal index As Long, ByVal Message As String, ByVal SayColor As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString Message
    Buffer.WriteString "[Map] "
    Buffer.WriteLong SayColor
    
    SendDataToMap MapNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SayMsg_Global(ByVal index As Long, ByVal Message As String, ByVal SayColor As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSayMsg
    
    Buffer.WriteString GetPlayerName(index)
    Buffer.WriteLong GetPlayerAccess(index)
    Buffer.WriteLong GetPlayerPK(index)
    Buffer.WriteString Message
    Buffer.WriteString "[Global] "
    Buffer.WriteLong SayColor
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub ResetShopAction(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SResetShopAction
    
    SendDataToAll Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendStunned(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SStunned
    
    Buffer.WriteLong TempPlayer(index).StunDuration
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendBank(ByVal index As Long)
    Dim Buffer As clsBuffer
    Dim i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong sbank
    
    For i = 1 To MAX_BANK
        Buffer.WriteLong Account(index).Bank.Item(i).Num
        Buffer.WriteLong Account(index).Bank.Item(i).Value
    Next
    
    SendDataTo index, Buffer.ToArray()
    
    Set Buffer = Nothing
End Sub

Sub SendOpenShop(ByVal index As Long, ByVal ShopNum As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SOpenShop
    
    Buffer.WriteLong ShopNum
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerMove(ByVal index As Long, ByVal movement As Byte, Optional ByVal SendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerMove
    
    Buffer.WriteLong index
    Buffer.WriteByte GetPlayerX(index)
    Buffer.WriteByte GetPlayerY(index)
    Buffer.WriteByte GetPlayerDir(index)
    Buffer.WriteByte movement
    
    If Not SendToSelf Then
        SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Else
        SendDataTo index, Buffer.ToArray()
    End If
    Set Buffer = Nothing
End Sub

Sub SendPlayerPosition(ByVal index As Long, Optional ByVal SendToSelf As Boolean = False)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPlayerWarp
    
    Buffer.WriteLong index
    Buffer.WriteByte GetPlayerX(index)
    Buffer.WriteByte GetPlayerY(index)
    Buffer.WriteByte GetPlayerDir(index)
    
    If Not SendToSelf Then
        SendDataToMap GetPlayerMap(index), Buffer.ToArray()
    Else
        SendDataTo index, Buffer.ToArray()
    End If
    Set Buffer = Nothing
End Sub

Sub SendNPCMove(ByVal MapNPCNum As Long, ByVal movement As Byte, ByVal MapNum As Integer)
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

Sub SendTrade(ByVal index As Long, ByVal TradeTarget As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STrade
    
    Buffer.WriteLong TradeTarget
    Buffer.WriteString Trim$(GetPlayerName(TradeTarget))
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendCloseTrade(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseTrade
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeUpdate(ByVal index As Long, ByVal DataType As Byte)
    Dim Buffer As clsBuffer
    Dim i As Long
    Dim TradeTarget As Long
    Dim TotalWorth As Long
    
    TradeTarget = TempPlayer(index).InTrade
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeUpdate
    
    Buffer.WriteByte DataType
    If DataType = 0 Then ' own inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Num
            Buffer.WriteLong TempPlayer(index).TradeOffer(i).Value
            ' add total worth
            If TempPlayer(index).TradeOffer(i).Num > 0 Then
                 If GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num) > 0 Then
                    ' currency?
                    If Item(TempPlayer(index).TradeOffer(i).Num).Stackable = 1 Then
                        TotalWorth = TotalWorth + (Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).Price * TempPlayer(index).TradeOffer(i).Value)
                    Else
                        TotalWorth = TotalWorth + Item(GetPlayerInvItemNum(index, TempPlayer(index).TradeOffer(i).Num)).Price
                    End If
                End If
            End If
        Next
    ElseIf DataType = 1 Then ' other inventory
        For i = 1 To MAX_INV
            Buffer.WriteLong GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num)
            Buffer.WriteLong TempPlayer(TradeTarget).TradeOffer(i).Value
            ' add total worth
            If GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num) > 0 Then
                ' currency?
                If Item(GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num)).Stackable = 1 Then
                    TotalWorth = TotalWorth + (Item(GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num)).Price * TempPlayer(TradeTarget).TradeOffer(i).Value)
                Else
                    TotalWorth = TotalWorth + Item(GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num)).Price
                End If
            End If
        Next
    End If
    ' Send total worth of trade
    Buffer.WriteLong TotalWorth
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeStatus(ByVal index As Long, ByVal Status As Byte)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeStatus
    
    Buffer.WriteByte Status
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendAttack(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SAttack
    Buffer.WriteLong index
    
    SendDataToMapBut index, GetPlayerMap(index), Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPlayerTarget(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong STarget
    
    Buffer.WriteByte TempPlayer(index).target
    Buffer.WriteByte TempPlayer(index).targetType
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendHotbar(ByVal index As Long)
    Dim i As Long
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SHotbar
    
    For i = 1 To MAX_HOTBAR
        Buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Hotbar(i).Slot
        Buffer.WriteByte Account(index).Chars(GetPlayerChar(index)).Hotbar(i).SType
    Next
    
    SendDataTo index, Buffer.ToArray()
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

Sub SendLogs(ByVal index As Long, Msg As String, Name As String)
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
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub UpdateFriendsList(index)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim i As Long, n As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SFriendsList
    
    If Account(index).Friends.AmountOfFriends = 0 Then
        Buffer.WriteByte Account(index).Friends.AmountOfFriends
        GoTo Finish
    End If
   
    ' Sends the amount of friends in friends list
    Buffer.WriteByte Account(index).Friends.AmountOfFriends
   
    ' Check to see if they are online
    For i = 1 To Account(index).Friends.AmountOfFriends
        Name = Trim$(Account(index).Friends.Members(i))
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
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub UpdateFoesList(index)
    Dim Buffer As clsBuffer
    Dim Name As String
    Dim i As Long, n As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SFoesList
    
    If Account(index).Foes.Amount = 0 Then
        Buffer.WriteByte Account(index).Foes.Amount
        GoTo Finish
    End If
   
    ' Sends the amount of Foes in Foes list
    Buffer.WriteByte Account(index).Foes.Amount
   
    ' Check to see if they are online
    For i = 1 To Account(index).Foes.Amount
        Name = Trim$(Account(index).Foes.Members(i))
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
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendInGame(ByVal index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SInGame
    
    SendDataTo index, Buffer.ToArray()
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

Sub SendSoundTo(ByVal index As Integer, Sound As String)
    Dim Buffer As clsBuffer
    
    ' Don't send it if there's nothing to send
    If Sound = vbNullString Then Exit Sub
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSound
    
    Buffer.WriteString Sound
    SendDataTo index, Buffer.ToArray()
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

Sub SendPlayerSound(ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal EntityType As Long, ByVal EntityNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SEntitySound
    
    Buffer.WriteLong X
    Buffer.WriteLong Y
    Buffer.WriteLong EntityType
    Buffer.WriteLong EntityNum
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendMapSound(ByVal MapNum As Integer, ByVal index As Long, ByVal X As Long, ByVal Y As Long, ByVal EntityType As Long, ByVal EntityNum As Long)
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

Sub SendLogin(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SLoginOk
    
    Buffer.WriteLong index
    Buffer.WriteLong Player_HighIndex
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNewCharClasses(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SNewCharClasses
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendNews(ByVal index As Long)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSendNews
    
    Buffer.WriteString Options.News
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendCheckForMap(ByVal index As Long, MapNum As Integer)
    Dim Buffer As clsBuffer
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckForMap
    
    Buffer.WriteInteger MapNum
    Buffer.WriteInteger Map(MapNum).Revision
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendTradeRequest(ByVal index As Long, ByVal TradeRequest As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong STradeRequest
    
    Buffer.WriteString Trim$(GetPlayerName(TradeRequest))
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyInvite(ByVal index As Long, ByVal OtherPlayer As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyInvite
    
    Buffer.WriteString Trim$(Account(OtherPlayer).Chars(GetPlayerChar(OtherPlayer)).Name)
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdate(ByVal PartyNum As Long)
    Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    Buffer.WriteByte 1
    For i = 1 To MAX_PARTY_MEMBERS
        Buffer.WriteLong Party(PartyNum).Member(i)
    Next
    Buffer.WriteLong Party(PartyNum).MemberCount
    Buffer.WriteLong PartyNum
    
    SendDataToParty PartyNum, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyUpdateTo(ByVal index As Long)
    Dim Buffer As clsBuffer, i As Long, PartyNum As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyUpdate
    
    ' Check if we're in a party
    PartyNum = TempPlayer(index).InParty
    
    If PartyNum > 0 Then
        ' Send party data
        Buffer.WriteByte 1
        Buffer.WriteLong Party(PartyNum).Leader
        For i = 1 To MAX_PARTY_MEMBERS
            Buffer.WriteLong Party(PartyNum).Member(i)
        Next
        Buffer.WriteLong Party(PartyNum).MemberCount
    Else
        ' Send clear command
        Buffer.WriteByte 0
    End If
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendPartyVitals(ByVal PartyNum As Long, ByVal index As Long)
    Dim Buffer As clsBuffer, i As Long

    Set Buffer = New clsBuffer
    Buffer.WriteLong SPartyVitals
    
    Buffer.WriteLong index
    For i = 1 To Vitals.Vital_Count - 1
        Buffer.WriteLong GetPlayerMaxVital(index, i)
        Buffer.WriteLong Account(index).Chars(GetPlayerChar(index)).Vital(i)
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

Sub SendUpdateTitleTo(ByVal index As Long, ByVal TitleNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendCloseClient(index As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SCloseClient
    
    SendDataTo index, Buffer.ToArray()
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

Sub SendUpdateMoralTo(ByVal index As Long, ByVal MoralNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateClassTo(ByVal index As Long, ByVal ClassNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
    Set Buffer = Nothing
End Sub

Sub SendUpdateEmoticonTo(ByVal index As Long, ByVal EmoticonNum As Long)
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
    
    SendDataTo index, Buffer.ToArray()
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

Sub SendCheckEmoticon(ByVal index As Long, ByVal MapNum As Long, ByVal EmoticonNum As Long)
    Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SCheckEmoticon
    
    Buffer.WriteLong index
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

Sub SendSpecialEffect(ByVal index As Long, EffectType As Long, Optional Data1 As Long = 0, Optional Data2 As Long = 0, Optional Data3 As Long = 0, Optional Data4 As Long = 0)
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
    
    SendDataTo index, Buffer.ToArray
    Set Buffer = Nothing
End Sub
