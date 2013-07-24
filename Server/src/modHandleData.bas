Attribute VB_Name = "modHandleData"
Option Explicit

Private Function GetAddress(FunAddr As Long) As Long
    GetAddress = FunAddr
End Function

Public Sub InitMessages()
    HandleDataSub(CNewAccount) = GetAddress(AddressOf HandleNewAccount)
    HandleDataSub(CDelAccount) = GetAddress(AddressOf HandleDelAccount)
    HandleDataSub(CLogin) = GetAddress(AddressOf HandleLogin)
    HandleDataSub(CAddChar) = GetAddress(AddressOf HandleAddChar)
    HandleDataSub(CUseChar) = GetAddress(AddressOf HandleUseChar)
    HandleDataSub(CSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(CEmoteMsg) = GetAddress(AddressOf HandleEmoteMsg)
    HandleDataSub(CGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(CAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(CPrivateMsg) = GetAddress(AddressOf HandlePrivateMsg)
    HandleDataSub(CPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(CPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(CUseItem) = GetAddress(AddressOf HandleUseItem)
    HandleDataSub(CAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(CWarpMeTo) = GetAddress(AddressOf HandleWarpMeTo)
    HandleDataSub(CWarpToMe) = GetAddress(AddressOf HandleWarpToMe)
    HandleDataSub(CWarpTo) = GetAddress(AddressOf HandleWarpTo)
    HandleDataSub(CSetSprite) = GetAddress(AddressOf HandleSetSprite)
    HandleDataSub(CSetPlayerSprite) = GetAddress(AddressOf HandleSetPlayerSprite)
    HandleDataSub(CRequestNewMap) = GetAddress(AddressOf HandleRequestNewMap)
    HandleDataSub(CMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(CNeedMap) = GetAddress(AddressOf HandleNeedMap)
    HandleDataSub(CMapGetItem) = GetAddress(AddressOf HandleMapGetItem)
    HandleDataSub(CMapDropItem) = GetAddress(AddressOf HandleMapDropItem)
    HandleDataSub(CMapRespawn) = GetAddress(AddressOf HandleMapRespawn)
    HandleDataSub(CMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(COpenMaps) = GetAddress(AddressOf HandleOpenMaps)
    HandleDataSub(CKickPlayer) = GetAddress(AddressOf HandleKickPlayer)
    HandleDataSub(CMutePlayer) = GetAddress(AddressOf HandleMutePlayer)
    HandleDataSub(CBanPlayer) = GetAddress(AddressOf HandleBanPlayer)
    HandleDataSub(CRequestPlayerData) = GetAddress(AddressOf HandleRequestPlayerData)
    HandleDataSub(CRequestPlayerStats) = GetAddress(AddressOf HandleRequestPlayerStats)
    HandleDataSub(CRequestBans) = GetAddress(AddressOf HandleRequestBans)
    HandleDataSub(CRequestSpellCooldown) = GetAddress(AddressOf HandleRequestSpellCooldown)
    HandleDataSub(CRequestEditMap) = GetAddress(AddressOf HandleRequestEditMap)
    HandleDataSub(CRequestEditEvent) = GetAddress(AddressOf HandleRequestEditEvent)
    HandleDataSub(CRequestEditItem) = GetAddress(AddressOf HandleRequestEditItem)
    HandleDataSub(CSaveItem) = GetAddress(AddressOf HandleSaveItem)
    HandleDataSub(CRequestEditNPC) = GetAddress(AddressOf HandleRequestEditNpc)
    HandleDataSub(CSaveNPC) = GetAddress(AddressOf HandleSaveNpc)
    HandleDataSub(CRequestEditShop) = GetAddress(AddressOf HandleRequestEditShop)
    HandleDataSub(CSaveShop) = GetAddress(AddressOf HandleSaveShop)
    HandleDataSub(CRequestEditSpell) = GetAddress(AddressOf HandleRequestEditSpell)
    HandleDataSub(CSaveSpell) = GetAddress(AddressOf HandleSaveSpell)
    HandleDataSub(CSetAccess) = GetAddress(AddressOf HandleSetAccess)
    HandleDataSub(CWhosOnline) = GetAddress(AddressOf HandleWhosOnline)
    HandleDataSub(CSetMOTD) = GetAddress(AddressOf HandleSetMOTD)
    HandleDataSub(CSetSMotd) = GetAddress(AddressOf HandleSetSMotd)
    HandleDataSub(CSetGMotd) = GetAddress(AddressOf HandleSetGMotd)
    HandleDataSub(CSearch) = GetAddress(AddressOf HandleSearch)
    HandleDataSub(CSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(CCastSpell) = GetAddress(AddressOf HandleCastSpell)
    HandleDataSub(CLeaveGame) = GetAddress(AddressOf HandleLeaveGame)
    HandleDataSub(CSwapInvSlots) = GetAddress(AddressOf HandleSwapInvSlots)
    HandleDataSub(CSwapSpellSlots) = GetAddress(AddressOf HandleSwapSpellSlots)
    HandleDataSub(CSwapHotbarSlots) = GetAddress(AddressOf HandleSwapHotbarSlots)
    HandleDataSub(CRequestEditResource) = GetAddress(AddressOf HandleRequestEditResource)
    HandleDataSub(CSaveResource) = GetAddress(AddressOf HandleSaveResource)
    HandleDataSub(CCheckPing) = GetAddress(AddressOf HandleCheckPing)
    HandleDataSub(CUnequip) = GetAddress(AddressOf HandleUnequip)
    HandleDataSub(CRequestItems) = GetAddress(AddressOf HandleRequestItems)
    HandleDataSub(CRequestNPCs) = GetAddress(AddressOf HandleRequestNpcs)
    HandleDataSub(CRequestResources) = GetAddress(AddressOf HandleRequestResources)
    HandleDataSub(CSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(CUseStatPoint) = GetAddress(AddressOf HandleUseStatPoint)
    HandleDataSub(CRequestEditAnimation) = GetAddress(AddressOf HandleRequestEditAnimation)
    HandleDataSub(CSaveAnimation) = GetAddress(AddressOf HandleSaveAnimation)
    HandleDataSub(CRequestAnimations) = GetAddress(AddressOf HandleRequestAnimations)
    HandleDataSub(CRequestSpells) = GetAddress(AddressOf HandleRequestSpells)
    HandleDataSub(CRequestShops) = GetAddress(AddressOf HandleRequestShops)
    HandleDataSub(CRequestLevelUp) = GetAddress(AddressOf HandleRequestLevelUp)
    HandleDataSub(CForgetSpell) = GetAddress(AddressOf HandleForgetSpell)
    HandleDataSub(CCloseShop) = GetAddress(AddressOf HandleCloseShop)
    HandleDataSub(CBuyItem) = GetAddress(AddressOf HandleBuyItem)
    HandleDataSub(CSellItem) = GetAddress(AddressOf HandleSellItem)
    HandleDataSub(CSwapBankSlots) = GetAddress(AddressOf HandleSwapBankSlots)
    HandleDataSub(CDepositItem) = GetAddress(AddressOf HandleDepositItem)
    HandleDataSub(CWithdrawItem) = GetAddress(AddressOf HandleWithdrawItem)
    HandleDataSub(CCloseBank) = GetAddress(AddressOf HandleCloseBank)
    HandleDataSub(CAdminWarp) = GetAddress(AddressOf HandleAdminWarp)
    HandleDataSub(CFixItem) = GetAddress(AddressOf HandleFixItem)
    HandleDataSub(CTradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(CAcceptTrade) = GetAddress(AddressOf HandleAcceptTrade)
    HandleDataSub(CDeclineTrade) = GetAddress(AddressOf HandleDeclineTrade)
    HandleDataSub(CTradeItem) = GetAddress(AddressOf HandleTradeItem)
    HandleDataSub(CUntradeItem) = GetAddress(AddressOf HandleUntradeItem)
    HandleDataSub(CCanTrade) = GetAddress(AddressOf HandleCanTrade)
    HandleDataSub(CHotbarChange) = GetAddress(AddressOf HandleHotbarChange)
    HandleDataSub(CAcceptTradeRequest) = GetAddress(AddressOf HandleAcceptTradeRequest)
    HandleDataSub(CDeclineTradeRequest) = GetAddress(AddressOf HandleDeclineTradeRequest)
    HandleDataSub(CPartyRequest) = GetAddress(AddressOf HandlePartyRequest)
    HandleDataSub(CAcceptParty) = GetAddress(AddressOf HandleAcceptParty)
    HandleDataSub(CDeclineParty) = GetAddress(AddressOf HandleDeclineParty)
    HandleDataSub(CPartyLeave) = GetAddress(AddressOf HandlePartyLeave)
    HandleDataSub(CPartyMsg) = GetAddress(AddressOf HandlePartyMsg)
    HandleDataSub(CGuildCreate) = GetAddress(AddressOf HandleGuildCreate)
    HandleDataSub(CGuildChangeAccess) = GetAddress(AddressOf HandleGuildChangeAccess)
    HandleDataSub(CGuildInvite) = GetAddress(AddressOf HandleGuildInvite)
    HandleDataSub(CAcceptGuild) = GetAddress(AddressOf HandleAcceptGuild)
    HandleDataSub(CDeclineGuild) = GetAddress(AddressOf HandleDeclineGuild)
    HandleDataSub(CGuildRemove) = GetAddress(AddressOf HandleGuildRemove)
    HandleDataSub(CGuildDisband) = GetAddress(AddressOf HandleGuildDisband)
    HandleDataSub(CGuildMsg) = GetAddress(AddressOf HandleGuildMsg)
    HandleDataSub(CBreakSpell) = GetAddress(AddressOf HandleBreakSpell)
    HandleDataSub(CAddFriend) = GetAddress(AddressOf HandleAddFriend)
    HandleDataSub(CRemoveFriend) = GetAddress(AddressOf HandleRemoveFriend)
    HandleDataSub(CFriendsList) = GetAddress(AddressOf HandleUpdateFriendsList)
    HandleDataSub(CAddFoe) = GetAddress(AddressOf HandleAddFoe)
    HandleDataSub(CRemoveFoe) = GetAddress(AddressOf HandleRemoveFoe)
    HandleDataSub(CFoesList) = GetAddress(AddressOf HandleUpdateFoesList)
    HandleDataSub(CUpdateData) = GetAddress(AddressOf HandleUpdateData)
    HandleDataSub(CSaveBan) = GetAddress(AddressOf HandleSaveBan)
    HandleDataSub(CRequestEditBans) = GetAddress(AddressOf HandleBanEditor)
    HandleDataSub(CSetTitle) = GetAddress(AddressOf HandleSetTitle)
    HandleDataSub(CRequestEditTitles) = GetAddress(AddressOf HandleTitleEditor)
    HandleDataSub(CSaveTitle) = GetAddress(AddressOf HandleSaveTitle)
    HandleDataSub(CRequestTitles) = GetAddress(AddressOf HandleRequestTitles)
    HandleDataSub(CChangeStatus) = GetAddress(AddressOf HandleChangeStatus)
    HandleDataSub(CRequestEditMorals) = GetAddress(AddressOf HandleMoralEditor)
    HandleDataSub(CSaveMoral) = GetAddress(AddressOf HandleSaveMoral)
    HandleDataSub(CRequestMorals) = GetAddress(AddressOf HandleRequestMorals)
    HandleDataSub(CRequestEditClasses) = GetAddress(AddressOf HandleClassEditor)
    HandleDataSub(CSaveClass) = GetAddress(AddressOf HandleSaveClass)
    HandleDataSub(CRequestClasses) = GetAddress(AddressOf HandleRequestClasses)
    HandleDataSub(CDestoryItem) = GetAddress(AddressOf HandleDestroyItem)
    HandleDataSub(CRequestEditEmoticons) = GetAddress(AddressOf HandleEmoticonEditor)
    HandleDataSub(CSaveEmoticon) = GetAddress(AddressOf HandleSaveEmoticon)
    HandleDataSub(CRequestEmoticons) = GetAddress(AddressOf HandleRequestEmoticons)
    HandleDataSub(CCheckEmoticon) = GetAddress(AddressOf HandleCheckEmoticon)
    
    HandleDataSub(CEventChatReply) = GetAddress(AddressOf HandleEventChatReply)
    HandleDataSub(CEvent) = GetAddress(AddressOf HandleEvent)
    HandleDataSub(CRequestSwitchesAndVariables) = GetAddress(AddressOf HandleRequestSwitchesAndVariables)
    HandleDataSub(CSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    
    ' Character Editor
    HandleDataSub(CRequestAllCharacters) = GetAddress(AddressOf HandleRequestAllCharacters)
    HandleDataSub(CRequestPlayersOnline) = GetAddress(AddressOf HandleRequestPlayersOnline)
    HandleDataSub(CRequestExtendedPlayerData) = GetAddress(AddressOf HandleRequestExtendedPlayerData)
    HandleDataSub(CCharacterUpdate) = GetAddress(AddressOf HandleCharacterUpdate)
    
    HandleDataSub(CTarget) = GetAddress(AddressOf HandleTarget)
End Sub

' Will handle the packet data
Sub HandleData(ByVal index As Long, ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Dim MsgType As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MsgType = buffer.ReadLong
    
    If MsgType < 0 Then Exit Sub
    If MsgType >= CMSG_COUNT Then Exit Sub

    CallWindowProc HandleDataSub(MsgType), index, buffer.ReadBytes(buffer.Length), 0, 0
End Sub

Private Sub HandleNewAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim HDSerial As String
    Dim i As Long
    Dim n As Long

    ' Clear out old data
    If IsLoggedIn(index) Then Call ClearAccount(index)
    
    If Not IsPlaying(index) Then
        ' Make sure the server isn't being shutdown or restarted
        If IsShuttingDown Then
            Call AlertMsg(index, "Server is either rebooting or being shutdown.")
            Exit Sub
        End If
        
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        
        HDSerial = buffer.ReadString
        
        ' Check for ban
        If IsBanned(index, HDSerial) Then Exit Sub

        ' Check version
        If Not App.Major = buffer.ReadLong Or Not App.Minor = buffer.ReadLong Or Not App.Revision = buffer.ReadLong Then
            Call AlertMsg(index, "Version outdated, please visit " & Options.Website & " for more information on new releases and run the updater.")
            Exit Sub
        End If

        ' Get the data
        Name = buffer.ReadString
        Password = buffer.ReadString

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Or Len(Trim$(Name)) > NAME_LENGTH Then Exit Sub
        If Len(Trim$(Password)) < 3 Or Len(Trim$(Password)) > NAME_LENGTH Then Exit Sub
        
        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))
            If Not IsNameLegal(n) Then
                Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ are allowed.")
                Exit Sub
            End If
        Next
        
        For i = 1 To Len(Password)
            n = AscW(Mid$(Password, i, 1))
            If Not IsNameLegal(n) Then
                Call AlertMsg(index, "Invalid password, only letters, numbers, spaces, and _ are allowed.")
                Exit Sub
            End If
        Next

        ' Check to see if account already exists
        If Not AccountExist(Name) Then
            Call AddAccount(index, Name, Password)
            Call TextAdd("Account " & Name & " has been created.")
            Call AddLog("Account " & Name & " has been created.", "Player")
            
            ' Load the player
            Call LoadAccount(index, Name)
            
            ' Check if character data has been created
            If Len(Trim$(Account(index).Chars(GetPlayerChar(index)).Name)) > 0 Then
                ' We have a character
                HandleUseChar index
            Else
                If Not IsPlaying(index) Then
                    Call SendNewCharClasses(index)
                End If
            End If
                    
            ' Show the player up on the socket status
            Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", "Player")
            Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
        Else
            Call AlertMsg(index, "That account name is already in use!")
        End If
        Set buffer = Nothing
    End If
End Sub

' :::::::::::::::::::::::::::
' :: Delete account packet ::
' :::::::::::::::::::::::::::
Private Sub HandleDelAccount(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim HDSerial As String
    Dim i As Long
    
    ' Clear out old data
    If IsLoggedIn(index) Then Call ClearAccount(index)

    If Not IsPlaying(index) Then
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        
        HDSerial = buffer.ReadString
        
        ' Check for ban
        If IsBanned(index, HDSerial) Then Exit Sub
        
        ' Check version
        If Not App.Major = buffer.ReadLong Or Not App.Minor = buffer.ReadLong Or Not App.Revision = buffer.ReadLong Then
            Call AlertMsg(index, "Version outdated, please visit " & Options.Website & " for more information on new releases and run the updater.")
            Exit Sub
        End If
        
        ' Get the data
        Name = buffer.ReadString
        Password = buffer.ReadString

        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Or Len(Trim$(Name)) > NAME_LENGTH Then Exit Sub
        If Len(Trim$(Password)) < 3 Or Len(Trim$(Password)) > NAME_LENGTH Then Exit Sub
        
        If Not AccountExist(Name) Then
            Call AlertMsg(index, "That account name does not exist.")
            Exit Sub
        End If

        If Not PasswordOK(Name, Password) Then
            Call AlertMsg(index, "Incorrect password.")
            Exit Sub
        End If

        ' Load the player
        Call LoadAccount(index, Name)
        
        ' Check for ban
        If IsBanned(index, GetPlayerHDSerial(index)) Then Exit Sub
        
        ' Delete names from master name file
        If Len(Trim$(Account(index).Chars(GetPlayerChar(index)).Name)) > 0 Then
            Call DeleteName(Account(index).Chars(GetPlayerChar(index)).Name)
        End If

        Call ClearAccount(index)
        
        ' Everything went ok
        Call Kill(App.path & "\data\Accounts\" & Trim$(Name) & ".bin")
        Call AddLog("Account " & Trim$(Name) & " has been deleted.", "Player")
        Call AlertMsg(index, "Your account has been deleted.")
        
        Set buffer = Nothing
    End If
End Sub

' ::::::::::::::::::
' :: Login packet ::
' ::::::::::::::::::
Private Sub HandleLogin(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim i As Long
    Dim n As Long
    Dim HDSerial As String

    ' Clear out old data
    If IsLoggedIn(index) Then Call ClearAccount(index)

    If Not IsPlaying(index) Then
        ' Make sure the server isn't being shutdown or restarted
        If IsShuttingDown Then
            Call AlertMsg(index, "Server is either rebooting or being shutdown.")
            Exit Sub
        End If
        
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()

        HDSerial = buffer.ReadString
        
        ' Check for ban
        If IsBanned(index, HDSerial) Then Exit Sub

        ' Check version
        If Not App.Major = buffer.ReadLong Or Not App.Minor = buffer.ReadLong Or Not App.Revision = buffer.ReadLong Then
            Call AlertMsg(index, "Version outdated, please visit " & Options.Website & " for more information on new releases and run the updater.")
            Exit Sub
        End If
        
        ' Get the data
        Name = buffer.ReadString
        Password = buffer.ReadString
        
        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Or Len(Trim$(Name)) > NAME_LENGTH Then Exit Sub
        If Len(Trim$(Password)) < 3 Or Len(Trim$(Password)) > NAME_LENGTH Then Exit Sub
        
        If Not AccountExist(Name) Then
            Call AlertMsg(index, "That account name does not exist.")
            Exit Sub
        End If

        If Not PasswordOK(Name, Password) Then
            Call AlertMsg(index, "Incorrect password.")
            Exit Sub
        End If

        If IsMultiAccounts(Name) Then
            Call AlertMsg(index, "Multiple account logins are not authorized.")
            Exit Sub
        End If
        
        ' Load the player
        Call LoadAccount(index, Name)
        
        TempPlayer(index).HDSerial = HDSerial
        
        ' Check if character data has been created
        If Len(Trim$(Account(index).Chars(GetPlayerChar(index)).Name)) > 0 Then
            ' Load character
            HandleUseChar index
        Else
            If Not IsPlaying(index) Then
                Call SendNewCharClasses(index)
            End If
        End If
        
        ' Show the player up on the socket status
        Call AddLog(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".", "Player")
        Call TextAdd(GetPlayerLogin(index) & " has logged in from " & GetPlayerIP(index) & ".")
        
        Set buffer = Nothing
    End If
End Sub

' ::::::::::::::::::::::::::
' :: Add character packet ::
' ::::::::::::::::::::::::::
Private Sub HandleAddChar(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim Password As String
    Dim Gender As Byte
    Dim ClassNum As Byte
    Dim i As Long
    Dim n As Long

    If Not IsPlaying(index) Then
        ' Make sure the server isn't being shutdown or restarted
        If IsShuttingDown Then
            Call AlertMsg(index, "Server is either rebooting or being shutdown.")
            Exit Sub
        End If
        
        Set buffer = New clsBuffer
        buffer.WriteBytes Data()
        
        Name = buffer.ReadString
        Gender = buffer.ReadByte
        ClassNum = buffer.ReadByte
        
        ' Prevent hacking
        If Len(Trim$(Name)) < 3 Or Len(Trim$(Name)) > NAME_LENGTH Then Exit Sub

        ' Prevent hacking
        For i = 1 To Len(Name)
            n = AscW(Mid$(Name, i, 1))
            
            If Not IsNameLegal(n) Then
                Call AlertMsg(index, "Invalid name, only letters, numbers, spaces, and _ are allowed.")
                Exit Sub
            End If
        Next

        ' Prevent hacking
        If (Gender < GENDER_MALE) Or (Gender > GENDER_FEMALE) Then Exit Sub
        If ClassNum < 1 Or ClassNum > MAX_CLASSES Then Exit Sub
        If Class(ClassNum).Locked = 1 Then Exit Sub
        If Trim$(Class(ClassNum).Name) = vbNullString And Not ClassNum = 1 Then Exit Sub

        ' Check if char already exists in slot
        If CharExist(index) Then
            Call AlertMsg(index, "Character already exists!")
            Exit Sub
        End If

        ' Check if name is already in use
        If FindChar(Name) Then
            Call AlertMsg(index, "That name is already in use!")
            Exit Sub
        End If

        ' Everything went ok, add the character
        Call AddChar(index, Name, Gender, ClassNum)
        Call AddLog("Character " & Name & " added to " & GetPlayerLogin(index) & "'s account.", "Player")
        
        ' Log them in
        HandleUseChar index
        
        Set buffer = Nothing
    End If
End Sub

' ::::::::::::::::::::
' :: Social packets ::
' ::::::::::::::::::::
Private Sub HandleSayMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim LogMsg As String
    Dim i As Long
    Dim buffer As clsBuffer
    Dim MapNum As Integer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MapNum = GetPlayerMap(index)
    Msg = buffer.ReadString
    Set buffer = Nothing
    
    If Msg = vbNullString Then Exit Sub
    
    If Trim$(Account(index).Chars(GetPlayerChar(index)).Status) = "Muted" Then
        Call PlayerMsg(index, "You are muted!", BrightRed)
        Exit Sub
    End If
    
    LogMsg = GetPlayerName(index) & ": " & Msg

    ' Add the logs
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(index) Then
                Call SendLogs(i, LogMsg, "Map")
            End If
        End If
    Next
    
    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " says, '" & Msg & "'", "Player")
    Call SayMsg_Map(MapNum, index, Msg, White)
    Call SendChatBubble(GetPlayerMap(index), index, TARGET_TYPE_PLAYER, Msg, White)
End Sub

Private Sub HandleEmoteMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim LogMsg As String
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Msg = buffer.ReadString
    
    If Msg = vbNullString Then Exit Sub
    
    If Trim$(Account(index).Chars(GetPlayerChar(index)).Status) = "Muted" Then
        Call PlayerMsg(index, "You are muted!", BrightRed)
        Exit Sub
    End If
    
    LogMsg = GetPlayerName(index) & " " & Right$(Msg, Len(Msg) - 1)

    ' Add the logs
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(i) = GetPlayerMap(index) Then
                Call SendLogs(i, LogMsg, "Map")
            End If
        End If
    Next

    Call AddLog("Map #" & GetPlayerMap(index) & ": " & GetPlayerName(index) & " " & Msg, "Player")
    Call MapMsg(GetPlayerMap(index), GetPlayerName(index) & " " & Msg, EmoteColor)
    
    Set buffer = Nothing
End Sub

Private Sub HandleGlobalMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim LogMsg As String
    Dim s As String
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Msg = buffer.ReadString
    
    If Msg = vbNullString Then Exit Sub
    
    If Trim$(Account(index).Chars(GetPlayerChar(index)).Status) = "Muted" Then
        Call PlayerMsg(index, "You are muted!", BrightRed)
        Exit Sub
    End If
    
    LogMsg = GetPlayerName(index) & ": " & Msg
    
    ' Add the logs
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            Call SendLogs(i, LogMsg, "Global")
        End If
    Next

    s = "[Global] " & GetPlayerName(index) & ": " & Msg
    
    Call SayMsg_Global(index, Msg, White)
    Call AddLog(s, "Player")
    Call TextAdd(s)
    
    Set buffer = Nothing
End Sub

Private Sub HandlePrivateMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim MsgTo As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MsgTo = FindPlayer(buffer.ReadString)
    Msg = buffer.ReadString
    
    If Msg = vbNullString Then Exit Sub
    
    ' Check if they are trying to talk to themselves
    If MsgTo <> index Then
        If MsgTo > 0 Then
            ' Can't invite if the player is a foe
            If IsAFoe(index, MsgTo) = True Then Exit Sub
            
            ' Add server log
            Call AddLog(GetPlayerName(index) & " whispers " & GetPlayerName(MsgTo) & ", '" & Msg & "'", "Player")
            
            ' Send the messages
            Call PlayerMsg(MsgTo, "[Private] " & GetPlayerName(index) & " whispers you, '" & Msg & "'", Pink)
            Call PlayerMsg(index, "[Private] You whisper " & GetPlayerName(MsgTo) & ", '" & Msg & "'", Pink)
        Else
            Call PlayerMsg(index, "Player is not online!", BrightRed)
        End If
    Else
        Call PlayerMsg(index, "Cannot message yourself.", BrightRed)
    End If
    
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerMove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Byte, i As Long
    Dim Movement As Byte
    Dim buffer As clsBuffer
    Dim TmpX As Integer, TmpY As Integer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    Dir = buffer.ReadByte
    Movement = buffer.ReadByte
    TmpX = buffer.ReadInteger
    TmpY = buffer.ReadInteger
    Set buffer = Nothing

    Call PlayerMove(index, Dir, Movement)
End Sub

' :::::::::::::::::::::::::::::
' :: Moving character packet ::
' :::::::::::::::::::::::::::::
Sub HandlePlayerDir(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    If TempPlayer(index).GettingMap = YES Then Exit Sub

    Dir = buffer.ReadLong
    Set buffer = Nothing

    ' Prevent hacking
    If Dir < DIR_UP Or Dir > DIR_RIGHT Then Exit Sub

    Call SetPlayerDir(index, Dir)
    Set buffer = New clsBuffer
    buffer.WriteLong SPlayerDir
    buffer.WriteLong index
    buffer.WriteByte GetPlayerDir(index)
    SendDataToMapBut index, GetPlayerMap(index), buffer.ToArray()
End Sub

' :::::::::::::::::::::
' :: Use item packet ::
' :::::::::::::::::::::
Sub HandleUseItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim InvNum As Byte
    Dim InvItem As Integer
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    InvNum = buffer.ReadByte
    Set buffer = Nothing

    ' Check for subscript out of range
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    
    InvItem = GetPlayerInvItemNum(index, InvNum)
            
    UseItem index, InvNum
End Sub

' ::::::::::::::::::::::::::
' :: Player attack packet ::
' ::::::::::::::::::::::::::
Sub HandleAttack(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim MapNum As Integer, DirReq As Long, ChatNpc As Long
    Dim X As Long, Y As Long
    Dim WeaponSlot As Long
    
    ' Can't attack while casting
    If TempPlayer(index).SpellBuffer.Spell > 0 Then Exit Sub
    
    ' Can't attack while stunned
    If TempPlayer(index).StunDuration > 0 Then Exit Sub

    ' Send this packet so they can see the person attacking
    Call SendAttack(index)
    
    ' Try to attack a player
    For i = 1 To Player_HighIndex
        TempIndex = i
    
        ' Make sure we dont try to attack ourselves
        If Not TempIndex = index Then
            TryPlayerAttackPlayer index, i
        End If
    Next
    
    ' Try to attack a npc
    For i = 1 To Map(GetPlayerMap(index)).Npc_HighIndex
        TryPlayerAttackNpc index, i
    Next
    
    ' Check if we've got a remote chat tile
    MapNum = GetPlayerMap(index)
    X = GetPlayerX(index)
    Y = GetPlayerY(index)

    Select Case GetPlayerDir(index)
        Case DIR_UP
            If GetPlayerY(index) = 0 Then Exit Sub
            X = GetPlayerX(index)
            Y = GetPlayerY(index) - 1
        Case DIR_DOWN
            If GetPlayerY(index) = Map(MapNum).MaxY Then Exit Sub
            X = GetPlayerX(index)
            Y = GetPlayerY(index) + 1
        Case DIR_LEFT
            If GetPlayerX(index) = 0 Then Exit Sub
            X = GetPlayerX(index) - 1
            Y = GetPlayerY(index)
        Case DIR_RIGHT
            If GetPlayerX(index) = Map(MapNum).MaxX Then Exit Sub
            X = GetPlayerX(index) + 1
            Y = GetPlayerY(index)
    End Select
    
    ' Check trade skills
    CheckResource index, X, Y
End Sub

' ::::::::::::::::::::::
' :: Use stats packet ::
' ::::::::::::::::::::::
Sub HandleUseStatPoint(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim PointType As Byte
    Dim buffer As clsBuffer
    Dim sMes As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    PointType = buffer.ReadByte
    Set buffer = Nothing

    ' Prevent hacking
    If (PointType < 0) Or (PointType > Stats.Stat_count) Then Exit Sub

    ' Make sure they have points
    If GetPlayerPoints(index) < 1 Then Exit Sub
    
    ' Make sure they're not maxed
    If GetPlayerRawStat(index, PointType) >= MAX_STAT Then
        PlayerMsg index, "You cannot spend any more points on that stat.", BrightRed
        Exit Sub
    End If
    
    ' Take away a stat point
    Call SetPlayerPoints(index, GetPlayerPoints(index) - 1)

    ' Add the stat
    Select Case PointType
        Case Stats.Strength
            Call SetPlayerStat(index, Stats.Strength, GetPlayerRawStat(index, Stats.Strength) + 1)
            sMes = "Strength"
        Case Stats.Endurance
            Call SetPlayerStat(index, Stats.Endurance, GetPlayerRawStat(index, Stats.Endurance) + 1)
            sMes = "Endurance"
        Case Stats.Intelligence
            Call SetPlayerStat(index, Stats.Intelligence, GetPlayerRawStat(index, Stats.Intelligence) + 1)
            sMes = "Intelligence"
        Case Stats.Agility
            Call SetPlayerStat(index, Stats.Agility, GetPlayerRawStat(index, Stats.Agility) + 1)
            sMes = "Agility"
        Case Stats.Spirit
            Call SetPlayerStat(index, Stats.Spirit, GetPlayerRawStat(index, Stats.Spirit) + 1)
            sMes = "Spirit"
    End Select
    
    ' Send the message
    SendActionMsg GetPlayerMap(index), "+1 " & sMes, White, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 32)

    ' Send the update
    Call SendPlayerStats(index)
    Call SendPlayerPoints(index)
End Sub

' ::::::::::::::::::::::::::::::::
' :: Player info request packet ::
' ::::::::::::::::::::::::::::::::
Sub HandlePlayerInfoRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Name As String
    Dim i As Long
    Dim n As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Name = buffer.ReadString
    
    Set buffer = Nothing
    
    i = FindPlayer(Name)
End Sub

' :::::::::::::::::::::::
' :: Warp me to packet ::
' :::::::::::::::::::::::
Sub HandleWarpMeTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MAPPER Then Exit Sub

    ' The player
    n = FindPlayer(buffer.ReadString)
    
    Set buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(index, GetPlayerMap(n), GetPlayerX(n), GetPlayerY(n))
            Call PlayerMsg(n, GetPlayerName(index) & " has warped to you.", BrightBlue)
            Call PlayerMsg(index, "You have been warped to " & GetPlayerName(n) & ".", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped to " & GetPlayerName(n) & ", map #" & GetPlayerMap(n) & ".", "Staff")
        Else
            Call PlayerMsg(index, "Player is not online!", BrightRed)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp to yourself!", BrightRed)
    End If

End Sub

' :::::::::::::::::::::::
' :: Warp to me packet ::
' :::::::::::::::::::::::
Sub HandleWarpToMe(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MAPPER Then Exit Sub

    ' The player
    n = FindPlayer(buffer.ReadString)
    
    Set buffer = Nothing

    If n <> index Then
        If n > 0 Then
            Call PlayerWarp(n, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index))
            Call PlayerMsg(n, "You have been summoned by " & GetPlayerName(index) & ".", BrightBlue)
            Call PlayerMsg(index, GetPlayerName(n) & " has been summoned.", BrightBlue)
            Call AddLog(GetPlayerName(index) & " has warped " & GetPlayerName(n) & " to self, map #" & GetPlayerMap(index) & ".", "Staff")
        Else
            Call PlayerMsg(index, "Player is not online!", BrightRed)
        End If

    Else
        Call PlayerMsg(index, "You cannot warp to yourself!", BrightRed)
    End If
End Sub

' ::::::::::::::::::::::::
' :: Warp to map packet ::
' ::::::::::::::::::::::::
Sub HandleWarpTo(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MAPPER Then Exit Sub

    ' The map
    n = buffer.ReadInteger
    Set buffer = Nothing

    ' Prevent hacking
    If n < 1 Or n > MAX_MAPS Then Exit Sub

    Call PlayerWarp(index, n, GetPlayerX(index), GetPlayerY(index))
    Call PlayerMsg(index, "You have been warped to map #" & n, BrightBlue)
    Call AddLog(GetPlayerName(index) & " warped to map #" & n & ".", "Staff")
End Sub

' :::::::::::::::::::::::
' :: Set sprite packet ::
' :::::::::::::::::::::::
Sub HandleSetSprite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpriteNum As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MAPPER Then Exit Sub

    ' Sprite
    SpriteNum = buffer.ReadLong
    
    Set buffer = Nothing
    
    Call SetPlayerSprite(index, SpriteNum)
    Call SendPlayerSprite(index)
End Sub

Sub HandleSetPlayerSprite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpriteNum As Long, Name As String
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_ADMIN Then Exit Sub

    ' Sprite
    SpriteNum = buffer.ReadLong
    
    ' Player
    Name = buffer.ReadString
    
    Set buffer = Nothing
    
    If Not IsPlaying(FindPlayer(Name)) Then
        Call PlayerMsg(index, "Player is not online!", BrightRed)
        Exit Sub
    End If
    
    Call SetPlayerSprite(FindPlayer(Name), SpriteNum)
    Call SendPlayerSprite(FindPlayer(Name))
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Player request for a new map ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestNewMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Dir As Byte
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Dir = buffer.ReadLong
    Set buffer = Nothing

    Call PlayerMove(index, Dir, 1)
End Sub

' :::::::::::::::::::::
' :: Map data packet ::
' :::::::::::::::::::::
Sub HandleMapData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
   Dim i As Long
    Dim MapNum As Long
    Dim X As Long
    Dim Y As Long, z As Long, w As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MAPPER Then Exit Sub

    MapNum = GetPlayerMap(index)
    i = Map(MapNum).Revision + 1
    Call ClearMap(MapNum)
    
    Map(MapNum).Name = buffer.ReadString
    Map(MapNum).Music = buffer.ReadString
    Map(MapNum).BGS = buffer.ReadString
    Map(MapNum).Revision = i
    Map(MapNum).Moral = buffer.ReadByte
    Map(MapNum).Up = buffer.ReadLong
    Map(MapNum).Down = buffer.ReadLong
    Map(MapNum).Left = buffer.ReadLong
    Map(MapNum).Right = buffer.ReadLong
    Map(MapNum).BootMap = buffer.ReadLong
    Map(MapNum).BootX = buffer.ReadByte
    Map(MapNum).BootY = buffer.ReadByte
    
    Map(MapNum).Weather = buffer.ReadLong
    Map(MapNum).WeatherIntensity = buffer.ReadLong
    
    Map(MapNum).Fog = buffer.ReadLong
    Map(MapNum).FogSpeed = buffer.ReadLong
    Map(MapNum).FogOpacity = buffer.ReadLong
    
    Map(MapNum).Panorama = buffer.ReadLong
    
    Map(MapNum).Red = buffer.ReadLong
    Map(MapNum).Green = buffer.ReadLong
    Map(MapNum).Blue = buffer.ReadLong
    Map(MapNum).Alpha = buffer.ReadLong
    
    Map(MapNum).MaxX = buffer.ReadByte
    Map(MapNum).MaxY = buffer.ReadByte
    
    ReDim Map(MapNum).Tile(0 To Map(MapNum).MaxX, 0 To Map(MapNum).MaxY)

    Map(MapNum).Npc_HighIndex = buffer.ReadByte
    
    For X = 0 To Map(MapNum).MaxX
        For Y = 0 To Map(MapNum).MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map(MapNum).Tile(X, Y).Layer(i).X = buffer.ReadLong
                Map(MapNum).Tile(X, Y).Layer(i).Y = buffer.ReadLong
                Map(MapNum).Tile(X, Y).Layer(i).Tileset = buffer.ReadLong
            Next
            
            For z = 1 To MapLayer.Layer_Count - 1
                Map(MapNum).Tile(X, Y).Autotile(z) = buffer.ReadLong
            Next
            
            Map(MapNum).Tile(X, Y).Type = buffer.ReadByte
            Map(MapNum).Tile(X, Y).Data1 = buffer.ReadLong
            Map(MapNum).Tile(X, Y).Data2 = buffer.ReadLong
            Map(MapNum).Tile(X, Y).Data3 = buffer.ReadLong
            Map(MapNum).Tile(X, Y).Data4 = buffer.ReadString
            Map(MapNum).Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map(MapNum).NPC(X) = buffer.ReadLong
        Map(MapNum).NpcSpawnType(X) = buffer.ReadLong
        Call ClearMapNpc(X, MapNum)
    Next
    
    ' Event data
    Map(MapNum).EventCount = buffer.ReadLong
        
    If Map(MapNum).EventCount > 0 Then
        ReDim Map(MapNum).Events(0 To Map(MapNum).EventCount)
        For i = 1 To Map(MapNum).EventCount
            With Map(MapNum).Events(i)
                .Name = buffer.ReadString
                .Global = buffer.ReadLong
                .X = buffer.ReadLong
                .Y = buffer.ReadLong
                .PageCount = buffer.ReadLong
            End With
            
            If Map(MapNum).Events(i).PageCount > 0 Then
                ReDim Map(MapNum).Events(i).Pages(0 To Map(MapNum).Events(i).PageCount)
                For X = 1 To Map(MapNum).Events(i).PageCount
                    With Map(MapNum).Events(i).Pages(X)
                        .chkVariable = buffer.ReadLong
                        .VariableIndex = buffer.ReadLong
                        .VariableCondition = buffer.ReadLong
                        .VariableCompare = buffer.ReadLong
                            
                        .chkSwitch = buffer.ReadLong
                        .SwitchIndex = buffer.ReadLong
                        .SwitchCompare = buffer.ReadLong
                            
                        .chkHasItem = buffer.ReadLong
                        .HasItemIndex = buffer.ReadLong
                            
                        .chkSelfSwitch = buffer.ReadLong
                        .SelfSwitchIndex = buffer.ReadLong
                        .SelfSwitchCompare = buffer.ReadLong
                            
                        .GraphicType = buffer.ReadLong
                        .Graphic = buffer.ReadLong
                        .GraphicX = buffer.ReadLong
                        .GraphicY = buffer.ReadLong
                        .GraphicX2 = buffer.ReadLong
                        .GraphicY2 = buffer.ReadLong
                            
                        .MoveType = buffer.ReadLong
                        .MoveSpeed = buffer.ReadLong
                        .MoveFreq = buffer.ReadLong
                            
                        .MoveRouteCount = buffer.ReadLong
                        
                        .IgnoreMoveRoute = buffer.ReadLong
                        .RepeatMoveRoute = buffer.ReadLong
                            
                        If .MoveRouteCount > 0 Then
                            ReDim Map(MapNum).Events(i).Pages(X).MoveRoute(0 To .MoveRouteCount)
                            For Y = 1 To .MoveRouteCount
                                .MoveRoute(Y).index = buffer.ReadLong
                                .MoveRoute(Y).Data1 = buffer.ReadLong
                                .MoveRoute(Y).Data2 = buffer.ReadLong
                                .MoveRoute(Y).Data3 = buffer.ReadLong
                                .MoveRoute(Y).Data4 = buffer.ReadLong
                                .MoveRoute(Y).Data5 = buffer.ReadLong
                                .MoveRoute(Y).Data6 = buffer.ReadLong
                            Next
                        End If
                            
                        .WalkAnim = buffer.ReadLong
                        .DirFix = buffer.ReadLong
                        .WalkThrough = buffer.ReadLong
                        .ShowName = buffer.ReadLong
                        .Trigger = buffer.ReadLong
                        .CommandListCount = buffer.ReadLong
                            
                        .Position = buffer.ReadLong
                    End With
                        
                    If Map(MapNum).Events(i).Pages(X).CommandListCount > 0 Then
                        ReDim Map(MapNum).Events(i).Pages(X).CommandList(0 To Map(MapNum).Events(i).Pages(X).CommandListCount)
                        For Y = 1 To Map(MapNum).Events(i).Pages(X).CommandListCount
                            Map(MapNum).Events(i).Pages(X).CommandList(Y).CommandCount = buffer.ReadLong
                            Map(MapNum).Events(i).Pages(X).CommandList(Y).ParentList = buffer.ReadLong
                            If Map(MapNum).Events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                ReDim Map(MapNum).Events(i).Pages(X).CommandList(Y).Commands(1 To Map(MapNum).Events(i).Pages(X).CommandList(Y).CommandCount)
                                For z = 1 To Map(MapNum).Events(i).Pages(X).CommandList(Y).CommandCount
                                    With Map(MapNum).Events(i).Pages(X).CommandList(Y).Commands(z)
                                        .index = buffer.ReadLong
                                        .Text1 = buffer.ReadString
                                        .Text2 = buffer.ReadString
                                        .Text3 = buffer.ReadString
                                        .Text4 = buffer.ReadString
                                        .Text5 = buffer.ReadString
                                        .Data1 = buffer.ReadLong
                                        .Data2 = buffer.ReadLong
                                        .Data3 = buffer.ReadLong
                                        .Data4 = buffer.ReadLong
                                        .Data5 = buffer.ReadLong
                                        .Data6 = buffer.ReadLong
                                        .ConditionalBranch.CommandList = buffer.ReadLong
                                        .ConditionalBranch.Condition = buffer.ReadLong
                                        .ConditionalBranch.Data1 = buffer.ReadLong
                                        .ConditionalBranch.Data2 = buffer.ReadLong
                                        .ConditionalBranch.Data3 = buffer.ReadLong
                                        .ConditionalBranch.ElseCommandList = buffer.ReadLong
                                        .MoveRouteCount = buffer.ReadLong
                                        If .MoveRouteCount > 0 Then
                                            ReDim Preserve .MoveRoute(.MoveRouteCount)
                                            For w = 1 To .MoveRouteCount
                                                .MoveRoute(w).index = buffer.ReadLong
                                                .MoveRoute(w).Data1 = buffer.ReadLong
                                                .MoveRoute(w).Data2 = buffer.ReadLong
                                                .MoveRoute(w).Data3 = buffer.ReadLong
                                                .MoveRoute(w).Data4 = buffer.ReadLong
                                                .MoveRoute(w).Data5 = buffer.ReadLong
                                                .MoveRoute(w).Data6 = buffer.ReadLong
                                            Next
                                        End If
                                    End With
                                Next
                            End If
                        Next
                    End If
                Next
            End If
        Next
    End If
    
    Call SendMapNpcsToMap(MapNum)
    Call SpawnMapNpcs(MapNum)
    Call SpawnGlobalEvents(MapNum)
    
    For i = 1 To Player_HighIndex
        If Account(i).Chars(GetPlayerChar(i)).Map = MapNum Then
            SpawnMapEventsFor i, MapNum
        End If
    Next

    ' Save the map
    Call SaveMap(MapNum)
    Call MapCache_Create(MapNum)
    Call CacheResources(MapNum)

    ' Refresh map for everyone online
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = MapNum Then
            Call PlayerWarp(i, MapNum, GetPlayerX(i), GetPlayerY(i), True)
        End If
    Next i
    
    Call CacheMapBlocks(MapNum)

    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::
' :: Need map yes/no packet ::
' ::::::::::::::::::::::::::::
Sub HandleNeedMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    ' Send the map
    Call SendMap(index, GetPlayerMap(index))

    Call SendMapItemsTo(index, GetPlayerMap(index))
    Call SendMapNpcsTo(index, GetPlayerMap(index))
    Call SpawnMapEventsFor(index, GetPlayerMap(index))
    Call SendJoinMap(index)

    ' Send Resource cache
    For i = 0 To ResourceCache(GetPlayerMap(index)).Resource_Count
        SendResourceCacheTo index, i
    Next

    TempPlayer(index).GettingMap = NO
    Set buffer = New clsBuffer
    buffer.WriteLong SMapDone
    SendDataTo index, buffer.ToArray()
End Sub

' :::::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to pick up something packet ::
' :::::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapGetItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Call PlayerMapGetItem(index, buffer.ReadByte)
End Sub

' ::::::::::::::::::::::::::::::::::::::::::::
' :: Player trying to drop something packet ::
' ::::::::::::::::::::::::::::::::::::::::::::
Sub HandleMapDropItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim InvNum As Byte
    Dim Amount As Long
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    InvNum = buffer.ReadByte
    Amount = buffer.ReadLong
    Set buffer = Nothing

    ' Prevent hacking
    If InvNum < 1 Or InvNum > MAX_INV Or IsPlaying(index) = False Then Exit Sub

    ' Check the player isn't doing something
    If TempPlayer(index).InBank Or TempPlayer(index).InShop Or TempPlayer(index).InTrade > 0 Then Exit Sub
    
    If GetPlayerInvItemNum(index, InvNum) < 1 Or GetPlayerInvItemNum(index, InvNum) > MAX_ITEMS Then Exit Sub
    
    If Item(GetPlayerInvItemNum(index, InvNum)).Stackable = 1 Then
        If GetPlayerInvItemValue(index, InvNum) < Amount Then Amount = GetPlayerInvItemValue(index, InvNum)
        If Amount < 1 Or Amount > GetPlayerInvItemValue(index, InvNum) Then Exit Sub
    Else
        If Not Amount = 0 Then Exit Sub
    End If
    
    ' Check if the item is binded
    If GetPlayerInvItemBind(index, InvNum) = 1 Then Exit Sub

    ' Check if on a map that forbids dropping items
    If Moral(Map(GetPlayerMap(index)).Moral).CanDropItem = 0 Then
        Call PlayerMsg(index, "You can't drop items here!", BrightRed)
        Exit Sub
    End If
    
    ' Everything worked out fine
    Call PlayerMapDropItem(index, InvNum, Amount)
End Sub

' ::::::::::::::::::::::::
' :: Respawn map packet ::
' ::::::::::::::::::::::::
Sub HandleMapRespawn(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MAPPER Then Exit Sub

    ' Clear out it all
    For i = 1 To MAX_MAP_ITEMS
        Call SpawnItemSlot(i, 0, 0, 0, GetPlayerMap(index), MapItem(GetPlayerMap(index), i).X, MapItem(GetPlayerMap(index), i).Y)
        Call ClearMapItem(i, GetPlayerMap(index))
    Next

    ' Respawn
    Call SpawnMapItems(GetPlayerMap(index))

    ' Respawn Npcs
    For i = 1 To MAX_MAP_NPCS
        Call SpawnNpc(i, GetPlayerMap(index))
    Next

    CacheResources GetPlayerMap(index)
    Call PlayerMsg(index, "Map respawned.", BrightBlue)
    Call AddLog(GetPlayerName(index) & " has respawned map #" & GetPlayerMap(index), "Staff")
End Sub

' :::::::::::::::::::::::
' :: Map Report packet ::
' :::::::::::::::::::::::
Sub HandleMapReport(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MAPPER Then Exit Sub
   
    Set buffer = New clsBuffer
    buffer.WriteLong SMapReport
   
    For i = 1 To MAX_MAPS
        buffer.WriteString Trim$(Map(i).Name)
    Next
   
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleOpenMaps(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim s As String
    Dim i As Long
    Dim tMapStart As Long
    Dim tMapEnd As Long
    
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MAPPER Then Exit Sub
    
    s = "Open Maps: "
    tMapStart = 1
    tMapEnd = 1

    For i = 1 To MAX_MAPS

        If Len(Trim$(Map(i).Name)) = 0 Then
            tMapEnd = tMapEnd + 1
        Else
            If tMapEnd - tMapStart > 0 Then
                s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
            End If

            tMapStart = i + 1
            tMapEnd = i + 1
        End If
    Next
    
    s = s & Trim$(CStr(tMapStart)) & "-" & Trim$(CStr(tMapEnd - 1)) & ", "
    s = Mid$(s, 1, Len(s) - 2)
    s = s & "."
    Call PlayerMsg(index, s, Brown)
End Sub

' ::::::::::::::::::::::::
' :: Kick player packet ::
' ::::::::::::::::::::::::
Sub HandleKickPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MODERATOR Then Exit Sub

    ' The player Index
    n = FindPlayer(buffer.ReadString)
    Set buffer = Nothing

    If Not n = index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call GlobalMsg(GetPlayerName(n) & " has been kicked from " & Options.Name & " by " & GetPlayerName(index) & "!", White)
                Call AddLog(GetPlayerName(index) & " has kicked " & GetPlayerName(n) & ".", "Staff")
                Call AlertMsg(n, "You have been kicked by " & GetPlayerName(index) & "!")
                Call LeftGame(index)
            Else
                Call PlayerMsg(index, "They are a higher or same access admin as you!", BrightRed)
            End If

        Else
            Call PlayerMsg(index, "Player is not online!", BrightRed)
        End If

    Else
        Call PlayerMsg(index, "You can't kick yourself!", BrightRed)
    End If
End Sub

' ::::::::::::::::::::::::
' :: Mute Player packet ::
' ::::::::::::::::::::::::
Sub HandleMutePlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim n As Long, Name As String
    
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MODERATOR Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Name = buffer.ReadString
    
    If Name = vbNullString Then Exit Sub
    
    n = FindPlayer(Name)

    ' Prevent subscript out of range
    If n < 1 Or n > Player_HighIndex Or Not IsPlaying(n) Then
        Call PlayerMsg(index, "Player is not online!", BrightRed)
        Exit Sub
    End If
    
    If n = index Then
        Call PlayerMsg(index, "You can't mute yourself!", BrightRed)
        Exit Sub
    End If
    
    Call PlayerMsg(n, "You have been muted by " & GetPlayerName(index) & "!", BrightRed)
    Account(index).Chars(GetPlayerChar(index)).Status = "Muted"
    Call SendPlayerStatus(index)
End Sub

Public Sub LoadBans()
    Dim i As Long

    CheckBans
    
    For i = 1 To MAX_BANS
        Call LoadBan(i)
    Next
End Sub

Public Sub LoadBan(index As Long)
    Dim F As Long
    Dim filename  As String

    
    F = FreeFile
    filename = App.path & "\data\bans\" & index & ".dat"
    
    Open filename For Binary As #F
        Get #F, , Ban(index)
    Close #F
End Sub

Private Sub CheckBans()
    Dim i As Long

    For i = 1 To MAX_BANS
        If Not FileExist("\data\bans\ban" & i & ".dat") Then
            SaveBan i
        End If
    Next
End Sub

' :::::::::::::::::::::::
' :: Ban player packet ::
' :::::::::::::::::::::::
Sub HandleBanPlayer(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim Reason As String
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_ADMIN Then Exit Sub

    ' The player Index
    n = FindPlayer(buffer.ReadString)
    Reason = buffer.ReadString
    
    Set buffer = Nothing

    If n <> index Then
        If n > 0 Then
            If GetPlayerAccess(n) < GetPlayerAccess(index) Then
                Call BanIndex(n, index, Reason)
            Else
                Call PlayerMsg(index, "That is a higher or same access admin then you!", White)
            End If
        Else
            Call PlayerMsg(index, "Player is not online!", BrightRed)
        End If
    Else
        Call PlayerMsg(index, "You cannot ban yourself!", BrightRed)
    End If
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit map oacket ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditMap(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MAPPER Then Exit Sub

    SendMapEventData (index)
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEditMap
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::::::::::::
' :: Request edit event packet ::
' :::::::::::::::::::::::::::::::
Sub HandleRequestEditEvent(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MAPPER Then Exit Sub

    ' TODO Add common event sending
    'SendMapEventData (Index)
    
    Set buffer = New clsBuffer
    buffer.WriteLong SEditEvent
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit item packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteLong SItemEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save item packet ::
' ::::::::::::::::::::::
Sub HandleSaveItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    n = buffer.ReadLong

    If n < 1 Or n > MAX_ITEMS Then Exit Sub

    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing
    
    ' Save It
    Call SendUpdateItemToAll(n)
    Call UpdateAllPlayerItems(n)
    Call UpdateAllPlayerEquipmentItems
    Call SaveItem(n)
    Call AddLog(GetPlayerName(index) & " saved Item #" & n & ".", "Staff")
End Sub

' :::::::::::::::::::::::::::::::::::
' :: Request edit animation packet ::
' :::::::::::::::::::::::::::::::::::
Sub HandleRequestEditAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SAnimationEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::::::::
' :: Save animation packet ::
' :::::::::::::::::::::::::::
Sub HandleSaveAnimation(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    n = buffer.ReadLong

    If n < 1 Or n > MAX_ANIMATIONS Then Exit Sub

    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateAnimationToAll(n)
    Call SaveAnimation(n)
    Call AddLog(GetPlayerName(index) & " saved Animation #" & n & ".", "Staff")
End Sub

' :::::::::::::::::::::::::::::
' :: Request edit NPC packet ::
' :::::::::::::::::::::::::::::
Sub HandleRequestEditNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SNpcEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Save NPC packet ::
' :::::::::::::::::::::
Private Sub HandleSaveNpc(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim NpcNum As Long
    Dim buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    NpcNum = buffer.ReadLong

    ' Prevent hacking
    If NpcNum < 1 Or NpcNum > MAX_NPCS Then Exit Sub
    
    NpcSize = LenB(NPC(NpcNum))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(NPC(NpcNum)), ByVal VarPtr(NpcData(0)), NpcSize
    
    ' Save it
    Call SendUpdateNpcToAll(NpcNum)
    Call SaveNpc(NpcNum)
    Call AddLog(GetPlayerName(index) & " saved Npc #" & NpcNum & ".", "Staff")
End Sub

' ::::::::::::::::::::::::::::::::::
' :: Request edit resource packet ::
' ::::::::::::::::::::::::::::::::::
Sub HandleRequestEditResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SResourceEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::::::
' :: Save resource packet ::
' ::::::::::::::::::::::::::
Private Sub HandleSaveResource(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    ResourceNum = buffer.ReadLong

    ' Prevent hacking
    If ResourceNum < 1 Or ResourceNum > MAX_RESOURCES Then Exit Sub

    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    ' Save it
    Call SendUpdateResourceToAll(ResourceNum)
    Call SaveResource(ResourceNum)
    Call AddLog(GetPlayerName(index) & " saved Resource #" & ResourceNum & ".", "Staff")
End Sub

' ::::::::::::::::::::::::::::::
' :: Request edit shop packet ::
' ::::::::::::::::::::::::::::::
Sub HandleRequestEditShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SShopEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save shop packet ::
' ::::::::::::::::::::::
Sub HandleSaveShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ShopNum As Long
    Dim i As Long
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub
    
    ShopNum = buffer.ReadLong

    ' Prevent hacking
    If ShopNum < 1 Or ShopNum > MAX_SHOPS Then Exit Sub

    ShopSize = LenB(Shop(ShopNum))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(ShopNum)), ByVal VarPtr(ShopData(0)), ShopSize

    Set buffer = Nothing
    
    ' Save it
    Call SendUpdateShopToAll(ShopNum)
    Call SaveShop(ShopNum)
    Call AddLog(GetPlayerName(index) & " saving shop #" & ShopNum & ".", "Staff")
End Sub

' :::::::::::::::::::::::::::::::
' :: Request edit spell packet ::
' :::::::::::::::::::::::::::::::
Sub HandleRequestEditSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteLong SSpellEditor
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::::
' :: Save spell packet ::
' :::::::::::::::::::::::
Sub HandleSaveSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpellNum As Long
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    SpellNum = buffer.ReadLong

    ' Prevent hacking
    If SpellNum < 1 Or SpellNum > MAX_SPELLS Then Exit Sub

    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    
    ' Save it
    Call SendUpdateSpellToAll(SpellNum)
    Call SaveSpell(SpellNum)
    Call AddLog(GetPlayerName(index) & " saved Spell #" & SpellNum & ".", "Staff")
End Sub

' :::::::::::::::::::::::
' :: Set access packet ::
' :::::::::::::::::::::::
Sub HandleSetAccess(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim buffer As clsBuffer, playerToChange As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    playerToChange = buffer.ReadString
    
    ' The Index
    n = FindPlayer(playerToChange)
    
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_ADMIN Then
        SendAccessVerificator index, 0, "You access level is too low!:" & playerToChange, GetPlayerAccess(n)
        Exit Sub
    End If


    
    ' The access
    i = buffer.ReadLong
    
    Set buffer = Nothing

    ' Check for invalid access level
    If i >= 0 And i <= 4 Then
        ' Check if player is on
        If n > 0 Then
            ' Check to see if same level access is trying to change another access of the very same level and boot them if they are.
            'If GetPlayerAccess(n) = GetPlayerAccess(index) And index <> n Then
                'Call PlayerMsg(index, "That player has the same access level as you!", BrightRed)
                'SendAccessVerificator index, 0, "Cannot set a player to the same access level as yourself!:" & playerToChange, GetPlayerAccess(n)
                'Exit Sub
            'End If
            If GetPlayerAccess(n) = i Then
                Call PlayerMsg(index, "That player already has that access level!", BrightRed)
                SendAccessVerificator index, 1, "Access level saved!:" & playerToChange, GetPlayerAccess(n)
                Exit Sub
            End If
            
            If GetPlayerAccess(index) = i Then
                Call PlayerMsg(index, "Cannot set a player to the same access level as yourself!", BrightRed)
                SendAccessVerificator index, 0, "Cannot set a player to the same access level as yourself!:" & playerToChange, GetPlayerAccess(n)
                Exit Sub
            End If
            If GetPlayerAccess(index) < i Then
                Call PlayerMsg(index, "Cannot set a player's access level higher than yourself!", BrightRed)
                SendAccessVerificator index, 0, "Cannot set a player's access level higher than yourself!:" & playerToChange, GetPlayerAccess(n)
                Exit Sub
            End If
            
            If GetPlayerAccess(n) <= 0 Then
                Call GlobalMsg(GetPlayerName(n) & " has been blessed with administrative access.", BrightBlue)
            End If
            
            SendAccessVerificator index, 1, "Access level saved!:" & playerToChange, GetPlayerAccess(n)
            
            Call SetPlayerAccess(n, i)
            Call SendPlayerData(n)
            Call AddLog(GetPlayerName(index) & " has modified " & GetPlayerName(n) & "'s access.", "Staff")

        Else
            Call PlayerMsg(index, "Player is not online!", BrightRed)
            SendAccessVerificator index, 0, "Player is Offline!:" & playerToChange, GetPlayerAccess(n)
        End If
    Else
        Call PlayerMsg(index, "Invalid access level.", BrightRed)
        SendAccessVerificator index, 0, "Invalid access level!:" & playerToChange, GetPlayerAccess(n)
    End If

End Sub

' :::::::::::::::::::::::::
' :: Who's online packet ::
' :::::::::::::::::::::::::
Sub HandleWhosOnline(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendWhosOnline(index)
End Sub

' Character Editor
Sub HandleRequestPlayersOnline(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayersOnline(index)
End Sub

' :::::::::::::::::::::
' :: Set MOTD packet ::
' :::::::::::::::::::::
Sub HandleSetMOTD(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_ADMIN Then Exit Sub

    ' Save options
    Options.MOTD = Trim$(buffer.ReadString)
    SaveOptions
    
    Set buffer = Nothing
    Call GlobalMsg("MOTD changed to: " & Options.MOTD, BrightCyan)
    
    Call AddLog(GetPlayerName(index) & " changed MOTD to: " & Options.MOTD, "Staff")
End Sub

' ::::::::::::::::::::::
' :: Set SMOTD packet ::
' ::::::::::::::::::::::
Sub HandleSetSMotd(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_ADMIN Then Exit Sub

    ' Save options
    Options.SMOTD = Trim$(buffer.ReadString)
    SaveOptions
    
    Set buffer = Nothing
    Call AdminMsg("Staff MOTD changed to: " & Options.SMOTD, Cyan)
    
    Call AddLog(GetPlayerName(index) & " changed Staff MOTD to: " & Options.SMOTD, "Staff")
End Sub

' ::::::::::::::::::::::
' :: Set GMOTD packet ::
' ::::::::::::::::::::::
Sub HandleSetGMotd(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim message As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerGuild(index) = 0 Then Exit Sub
    If GetPlayerGuildAccess(index) < 3 Then Exit Sub
        
    message = buffer.ReadString
    Guild(GetPlayerGuild(index)).MOTD = message

    Set buffer = Nothing
    
    Call GuildMsg(index, GetPlayerName(index) & " has changed the MOTD to: " & message, BrightGreen, True)
    Call AddLog(GetPlayerName(index) & " changed MOTD to: " & message, "Player")
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleSearch(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Set buffer = Nothing

    ' Prevent subscript out of range
    If X < 0 Or X > Map(GetPlayerMap(index)).MaxX Or Y < 0 Or Y > Map(GetPlayerMap(index)).MaxY Then Exit Sub
    If Not IsPlaying(index) Then Exit Sub

    ' Check for a player
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerMap(index) = GetPlayerMap(i) Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        ' Change target
                        If TempPlayer(index).TargetType = TARGET_TYPE_PLAYER And TempPlayer(index).Target = i Then
                            TempPlayer(index).Target = 0
                            TempPlayer(index).TargetType = TARGET_TYPE_NONE
                            
                            ' Send target to player
                            SendPlayerTarget index
                        Else
                            TempPlayer(index).Target = i
                            TempPlayer(index).TargetType = TARGET_TYPE_PLAYER
                            
                            ' Send target to player
                            SendPlayerTarget index
                        End If
                        Exit Sub
                    End If
                End If
            End If
        End If
    Next
    
    ' Check for an npc
    For i = 1 To Map(GetPlayerMap(index)).Npc_HighIndex
        If MapNpc(GetPlayerMap(index)).NPC(i).Num > 0 Then
            If MapNpc(GetPlayerMap(index)).NPC(i).X = X Then
                If MapNpc(GetPlayerMap(index)).NPC(i).Y = Y Then
                    If TempPlayer(index).Target = i And TempPlayer(index).TargetType = TARGET_TYPE_NPC Then
                        ' Change target
                        TempPlayer(index).Target = 0
                        TempPlayer(index).TargetType = TARGET_TYPE_NONE
                       
                        ' Send target to player
                        SendPlayerTarget index
                    Else
                        ' Change target
                        TempPlayer(index).Target = i
                        TempPlayer(index).TargetType = TARGET_TYPE_NPC
                        
                        ' Send target to player
                        SendPlayerTarget index
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next
    
    ' Check for an item
    For i = 1 To MAX_MAP_ITEMS
        If MapItem(GetPlayerMap(index), i).Num > 0 Then
            If MapItem(GetPlayerMap(index), i).X = X And MapItem(GetPlayerMap(index), i).Y = Y Then
                If CanPlayerPickupItem(index, i) Then
                    If Item(MapItem(GetPlayerMap(index), i).Num).Stackable = 1 Then
                        Call PlayerMsg(index, "You see " & MapItem(GetPlayerMap(index), i).Value & " " & Trim$(Item(MapItem(GetPlayerMap(index), i).Num).Name) & ".", Yellow)
                    Else
                        Call PlayerMsg(index, "You see " & CheckGrammar(Trim$(Item(MapItem(GetPlayerMap(index), i).Num).Name)) & ".", Yellow)
                    End If
                    Exit Sub
                End If
            End If
        End If
    Next
End Sub

' :::::::::::::::::::
' :: Spells packet ::
' :::::::::::::::::::
Sub HandleSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call SendPlayerSpells(index)
End Sub

' :::::::::::::::::
' :: Cast packet ::
' :::::::::::::::::
Sub HandleCastSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Spell slot
    n = buffer.ReadLong
    
    Set buffer = Nothing
    
    ' Set the spell buffer before castin
    Call BufferPlayerSpell(index, n)
End Sub

' ::::::::::::::::::::::
' :: Quit game packet ::
' ::::::::::::::::::::::
Sub HandleLeaveGame(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If IsPlaying(index) Then
        Call LeftGame(index)
    End If
End Sub

' ::::::::::::::::::::::::::
' :: Swap Inventory Slots ::
' ::::::::::::::::::::::::::
Sub HandleSwapInvSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim OldSlot As Byte, NewSlot As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Old Slot
    OldSlot = buffer.ReadByte
    
    ' New Slot
    NewSlot = buffer.ReadByte
    
    Set buffer = Nothing
    
    ' Make sure their valid
    If OldSlot < 1 Or OldSlot > MAX_INV Then Exit Sub
    If NewSlot < 1 Or NewSlot > MAX_INV Then Exit Sub
    If TempPlayer(index).InTrade > 0 Then Exit Sub
    
    PlayerSwitchInvSlots index, OldSlot, NewSlot
End Sub

' ::::::::::::::::::::::
' :: Swap Spell Slots ::
' ::::::::::::::::::::::
Sub HandleSwapSpellSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim OldSlot As Byte, NewSlot As Byte
    
    ' Prevent subscript if someone tries to cast
    If TempPlayer(index).SpellBuffer.Spell > 0 Then
        If TempPlayer(index).SpellBuffer.Spell = Account(index).Chars(GetPlayerChar(index)).Spell(OldSlot) Or Account(index).Chars(GetPlayerChar(index)).Spell(NewSlot) Then Exit Sub
    End If
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Old Slot
    OldSlot = buffer.ReadByte
    
    ' New Slot
    NewSlot = buffer.ReadByte
    
    Set buffer = Nothing
    
    ' Make sure their valid
    If OldSlot < 1 Or OldSlot > MAX_PLAYER_SPELLS Then Exit Sub
    If NewSlot < 1 Or NewSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    PlayerSwitchSpellSlots index, OldSlot, NewSlot
End Sub

' :::::::::::::::::::::::
' :: Swap Hotbar Slots ::
' :::::::::::::::::::::::
Sub HandleSwapHotbarSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim OldSlot As Byte, NewSlot As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Old Slot
    OldSlot = buffer.ReadByte
    
    ' New Slot
    NewSlot = buffer.ReadByte
    
    Set buffer = Nothing
    
    ' Make sure their valid
    If OldSlot < 1 Or OldSlot > MAX_HOTBAR Then Exit Sub
    If NewSlot < 1 Or NewSlot > MAX_HOTBAR Then Exit Sub
    
    PlayerSwitchHotbarSlots index, OldSlot, NewSlot
End Sub

' ::::::::::::::::
' :: Check Ping ::
' ::::::::::::::::
Sub HandleCheckPing(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteLong SSendPing
    
    SendDataTo index, buffer.ToArray()
    Set buffer = Nothing
End Sub

Sub HandleUnequip(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    PlayerUnequipItem index, buffer.ReadLong
    Set buffer = Nothing
End Sub

Sub HandleRequestPlayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerData index
End Sub

Sub HandleRequestPlayerStats(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendPlayerStats index
End Sub

Sub HandleRequestSpellCooldown(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Slot As Byte
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    Slot = buffer.ReadByte
    
    Call SendSpellCooldown(index, Slot)
End Sub

Sub HandleRequestBans(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim BanSize As Long
    Dim BanData() As Byte
    Dim i As Long
    
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_ADMIN Then Exit Sub
    
    Set buffer = New clsBuffer
    
    For i = 1 To MAX_BANS
        If Len(Trim$(Ban(i).playerName)) > 0 Then
            BanSize = LenB(Ban(i))
            ReDim BanData(BanSize - 1)
            CopyMemory BanData(0), ByVal VarPtr(Ban(i)), BanSize
            buffer.WriteLong SUpdateBan
            buffer.WriteLong i
            buffer.WriteBytes BanData
            SendDataTo index, buffer.ToArray()
        End If
    Next
    Set buffer = Nothing
End Sub

Sub HandleRequestTitles(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendTitles index
End Sub

Sub HandleRequestMorals(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendMorals index
End Sub

Sub HandleRequestClasses(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendClasses index
End Sub

Sub HandleRequestEmoticons(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendEmoticons index
End Sub

Sub HandleRequestItems(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendItems index
End Sub

Sub HandleRequestAnimations(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendAnimations index
End Sub

Sub HandleRequestNpcs(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendNpcs index
End Sub

Sub HandleRequestResources(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendResources index
End Sub

Sub HandleRequestSpells(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSpells index
End Sub

Sub HandleRequestShops(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendShops index
End Sub

Sub HandleSpawnItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim TmpItem As Long
    Dim TmpAmount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub
    
    ' Item
    TmpItem = buffer.ReadLong
    TmpAmount = buffer.ReadLong
    
    If Moral(GetPlayerMap(index)).CanDropItem = 0 Then
        SpawnItem TmpItem, TmpAmount, Item(TmpItem).Data1, GetPlayerMap(index), GetPlayerX(index), GetPlayerY(index), GetPlayerName(index)
    Else
        GiveInvItem index, TmpItem, TmpAmount
    End If
    
    Set buffer = Nothing
End Sub

Sub HandleRequestLevelUp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub
    
    ' Make sure their not the max level
    If GetPlayerLevel(index) = MAX_LEVEL Then Exit Sub
    
    SetPlayerExp index, GetPlayerNextLevel(index)
    CheckPlayerLevelUp index
End Sub

Sub HandleForgetSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim SpellSlot As Byte, i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    SpellSlot = buffer.ReadByte
    
    ' Check for subscript out of range
    If SpellSlot < 1 Or SpellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    ' Don't let them forget a spell which is in CD
    If GetPlayerSpellCD(index, SpellSlot) > timeGetTime Then
        PlayerMsg index, "Cannot forget a spell which is cooling down!", BrightRed
        Exit Sub
    End If
    
    ' Don't let them forget a spell which is buffered
    If TempPlayer(index).SpellBuffer.Spell = SpellSlot Then
        PlayerMsg index, "Cannot forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    ' Check if we need to remove anything from the botbar
    For i = 1 To MAX_HOTBAR
        If Account(index).Chars(GetPlayerChar(index)).Hotbar(i).Slot = SpellSlot And Account(index).Chars(GetPlayerChar(index)).Hotbar(i).SType = 2 Then
            Account(index).Chars(GetPlayerChar(index)).Hotbar(i).Slot = 0
            Account(index).Chars(GetPlayerChar(index)).Hotbar(i).SType = 0
            SendHotbar index
        End If
    Next
    
    Call SetPlayerSpell(index, SpellSlot, 0)
    SendPlayerSpells index
    
    Set buffer = Nothing
End Sub

Sub HandleCloseShop(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    TempPlayer(index).InShop = 0
End Sub

Sub HandleBuyItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim ShopSlot As Long
    Dim ShopNum As Long
    Dim ItemAmount As Long
    Dim ItemAmount2 As Long
    Dim Multiplier As Integer
    Dim ItemPrice As Long
    Dim ItemPrice2 As Long
   
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
   
    ShopSlot = buffer.ReadLong
    ShopNum = TempPlayer(index).InShop
    
    ' Exit shop if not in it
    If ShopNum < 1 Or ShopNum > MAX_SHOPS Then Exit Sub
   
    With Shop(ShopNum).TradeItem(ShopSlot)
        ' Check that trade exists
        If .Item < 1 Then Exit Sub
        
        ' Work out price
        Multiplier = Shop(TempPlayer(index).InShop).BuyRate / 100
        ItemPrice = Item(.CostItem).Price * Multiplier
        ItemPrice2 = Item(.CostItem2).Price * Multiplier
        
        ' Calculate how much of the item they have
        ItemAmount = HasItem(index, .CostItem)
        ItemAmount2 = HasItem(index, .CostItem2)
        
        If .CostItem > 0 And .CostItem2 = 0 Then
            If ItemAmount = 0 Or ItemAmount < ItemPrice Then
                PlayerMsg index, "You do not have enough " & Trim$(Item(1).Name) & " to buy this item.", BrightRed
                ResetShopAction index
                Exit Sub
            End If
        ElseIf .CostItem = 0 And .CostItem2 > 0 Then
            If ItemAmount2 = 0 Or ItemAmount2 < ItemPrice2 Then
                PlayerMsg index, "You do not have enough " & Trim$(Item(1).Name) & " to buy this item.", BrightRed
                ResetShopAction index
                Exit Sub
            End If
        ElseIf .CostItem > 0 And .CostItem2 > 0 Then
            If ItemAmount2 = 0 Or ItemAmount2 < ItemPrice2 Or ItemAmount = 0 Or ItemAmount < ItemPrice Then
                PlayerMsg index, "You do not have enough " & Trim$(Item(1).Name) & " to buy this item.", BrightRed
                ResetShopAction index
                Exit Sub
            End If
        End If
       
        ' It's fine, let's go ahead
        If .CostItem > 0 And .CostItem2 = 0 Then
            TakeInvItem index, .CostItem, ItemPrice
            GiveInvItem index, .Item, .ItemValue
        ElseIf .CostItem2 > 0 And .CostItem = 0 Then
            TakeInvItem index, .CostItem2, ItemPrice2
            GiveInvItem index, .Item, .ItemValue
        ElseIf .CostItem > 0 And .CostItem2 > 0 Then
            TakeInvItem index, .CostItem, ItemPrice
            TakeInvItem index, .CostItem2, ItemPrice2
            GiveInvItem index, .Item, .ItemValue
        End If
    End With
   
    ' Send confirmation message & reset their shop action
    Call SendSoundTo(index, Options.BuySound)
    PlayerMsg index, "Trade successful.", Yellow
    ResetShopAction index
   
    Set buffer = Nothing
End Sub

Sub HandleSellItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim InvSlot As Byte
    Dim ItemNum As Integer
    Dim Price As Long
    Dim Multiplier As Integer
    Dim Amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Prevent hacking
    If TempPlayer(index).InShop < 1 Or TempPlayer(index).InShop > MAX_SHOPS Then Exit Sub
    
    InvSlot = buffer.ReadByte
    
    ' If invalid, exit out
    If InvSlot < 1 Or InvSlot > MAX_INV Then Exit Sub
    
    ' Has item?
    If GetPlayerInvItemNum(index, InvSlot) < 1 Or GetPlayerInvItemNum(index, InvSlot) > MAX_ITEMS Then Exit Sub
    
    ' Seems to be valid
    ItemNum = GetPlayerInvItemNum(index, InvSlot)
    
    ' Work out price
    Multiplier = Shop(TempPlayer(index).InShop).SellRate / 100
    Price = Item(ItemNum).Price * Multiplier
    
    ' Item has cost?
    If Price < 1 Or ItemNum = 1 Then
        PlayerMsg index, "The shop doesn't want that item.", BrightRed
        ResetShopAction index
        Exit Sub
    End If

    ' Take item and give `
    TakeInvItem index, ItemNum, 1
    GiveInvItem index, 1, Price
    
    ' Send confirmation message and reset their shop action
    Call SendSoundTo(index, Options.SellSound)
    PlayerMsg index, "Trade successful.", Yellow
    ResetShopAction index
    
    Set buffer = Nothing
End Sub

Sub HandleSwapBankSlots(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim NewSlot As Byte
    Dim OldSlot As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    OldSlot = buffer.ReadByte
    NewSlot = buffer.ReadByte
    Set buffer = Nothing
    
    ' Make sure their valid
    If OldSlot < 1 Or OldSlot > MAX_BANK Then Exit Sub
    If NewSlot < 1 Or NewSlot > MAX_BANK Then Exit Sub
    
    PlayerSwapBankSlots index, OldSlot, NewSlot
End Sub

Sub HandleWithdrawItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim BankSlot As Byte
    Dim Amount As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    BankSlot = buffer.ReadByte
    Amount = buffer.ReadLong
    
    TakeBankItem index, BankSlot, Amount
    
    Set buffer = Nothing
End Sub

Sub HandleDepositItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim InvSlot As Byte
    Dim Amount As Long
    Dim Durability As Integer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    InvSlot = buffer.ReadByte
    Amount = buffer.ReadLong
    Durability = GetPlayerInvItemDur(index, InvSlot)
    
    ' Prevent subscript out of range
    If InvSlot < 1 Or InvSlot > MAX_INV Then Exit Sub
    
    ' Hack prevention
    If Item(GetPlayerInvItemNum(index, InvSlot)).Stackable = 1 Then
        If GetPlayerInvItemValue(index, InvSlot) < Amount Then Amount = GetPlayerInvItemValue(index, InvSlot)
        If Amount < 1 Then Exit Sub
    Else
        If Not Amount = 0 Then Exit Sub
    End If
    
    GiveBankItem index, InvSlot, Amount, Durability
    
    Set buffer = Nothing
End Sub

Sub HandleCloseBank(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    
    SaveAccount index
    
    TempPlayer(index).InBank = False
    
    Set buffer = Nothing
End Sub

Sub HandleAdminWarp(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long
    Dim Y As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    X = buffer.ReadLong
    Y = buffer.ReadLong
    
    If GetPlayerAccess(index) >= STAFF_MAPPER Then
        SetPlayerX index, X
        SetPlayerY index, Y
        Call SendPlayerPosition(index)
    End If
    
    Set buffer = Nothing
End Sub

' :::::::::::::::::::::
' :: Fix item packet ::
' :::::::::::::::::::::
Private Sub HandleFixItem(ByVal index As Integer, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Byte
    Dim i As Long
    Dim ItemNum As Long
    Dim DurNeeded As Long
    Dim GoldNeeded As Long
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Prevent hacking
    If TempPlayer(index).InShop < 1 Or TempPlayer(index).InShop > MAX_SHOPS Then Exit Sub
    If Shop(TempPlayer(index).InShop).CanFix = 0 Then Exit Sub
    
    ' Inv num
    n = buffer.ReadByte
    
    ' Prevent hacking
    If n <= 0 Or n > MAX_INV Then Exit Sub
    
    ' Check for bad data
    If GetPlayerInvItemNum(index, n) <= 0 Or GetPlayerInvItemNum(index, n) > MAX_ITEMS Then Exit Sub

    ' Make sure its a equipable item
    If Not Item(GetPlayerInvItemNum(index, n)).Type = ITEM_TYPE_EQUIPMENT Then
        Call PlayerMsg(index, "You can only fix weapons, armors, helmets, and shields.", BrightRed)
        Exit Sub
    End If
    
    ' Now check the rate of pay
    ItemNum = GetPlayerInvItemNum(index, n)
    i = (Item(GetPlayerInvItemNum(index, n)).Data2 \ 5)
    If i <= 0 Then i = 1
    
    DurNeeded = Item(ItemNum).Data1 - GetPlayerInvItemDur(index, n)
    GoldNeeded = Int(DurNeeded * i / 2)
    If GoldNeeded <= 0 Then GoldNeeded = 1
    
    ' Check if they even need it repaired
    If DurNeeded <= 0 Then
        Call PlayerMsg(index, "This item is in perfect condition!", White)
        Exit Sub
    End If
    
    ' Check if they have enough for at least one point
    If HasItem(index, 1) >= i Then
        ' Check if they have enough for a total restoration
        If HasItem(index, 1) >= GoldNeeded Then
            Call TakeInvItem(index, 1, GoldNeeded)
            Call SetPlayerInvItemDur(index, n, Item(ItemNum).Data1)
            Call PlayerMsg(index, "Item has been totally restored for " & GoldNeeded & " " & Trim$(Item(1).Name) & "!", BrightBlue)
        Else
            ' They dont so restore as much as we can
            DurNeeded = (HasItem(index, 1) / i)
            GoldNeeded = Int(DurNeeded * i \ 2)
            If GoldNeeded <= 0 Then GoldNeeded = 1
            
            Call TakeInvItem(index, 1, GoldNeeded)
            Call SetPlayerInvItemDur(index, n, GetPlayerInvItemDur(index, n) + DurNeeded)
            Call PlayerMsg(index, "Item has been partially fixed for " & GoldNeeded & Trim$(Item(1).Name) & "!", BrightBlue)
        End If
    Else
        Call PlayerMsg(index, "Insufficient " & Trim$(Item(1).Name) & " to fix this item!", BrightRed)
    End If
End Sub

Sub HandleTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim TradeTarget As Long
    
    ' Can't trade with npcs
    If Not TempPlayer(index).TargetType = TARGET_TYPE_PLAYER Then Exit Sub

    ' Find the target
    TradeTarget = TempPlayer(index).Target
    
    ' Make sure we don't error
    If TradeTarget < 1 Or TradeTarget > MAX_PLAYERS Then Exit Sub
    
    ' Can't invite if the player is a foe
    If IsAFoe(index, TradeTarget) Then Exit Sub
    
    ' Make sure they're not in a trade
    If TempPlayer(TradeTarget).InTrade > 0 Then
        ' They're already in a trade
        PlayerMsg index, "This player is already in a trade!", BrightRed
        Exit Sub
    End If
    
    ' Check if there doing another action
    If IsPlayerBusy(index, TradeTarget) Then Exit Sub
    
    ' Let them know
    PlayerMsg index, "Trade invitation sent.", Pink

    ' Send the trade request
    TempPlayer(TradeTarget).TradeRequest = index
    SendTradeRequest TradeTarget, index
End Sub

Sub HandleAcceptTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim TradeTarget As Long
    Dim i As Long

    TradeTarget = TempPlayer(index).TradeRequest
    
    ' See if the player can trade
    If CanPlayerTrade(index, TradeTarget) = False Then
        ' Clear the tradeRequest server-side
        TempPlayer(index).TradeRequest = 0
        TempPlayer(TradeTarget).TradeRequest = 0
        Exit Sub
    End If
    
    ' Let them know they're trading
    PlayerMsg index, "You have accepted " & Trim$(GetPlayerName(TradeTarget)) & "'s trade request.", BrightGreen
    PlayerMsg TradeTarget, Trim$(GetPlayerName(index)) & " has accepted your trade request.", BrightGreen
    
    ' Clear the trade request server-side
    TempPlayer(index).TradeRequest = 0
    TempPlayer(TradeTarget).TradeRequest = 0
    
    ' Set that they're trading with each other
    TempPlayer(index).InTrade = TradeTarget
    TempPlayer(TradeTarget).InTrade = index
    
    ' Clear out their trade offers
    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).Num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(TradeTarget).TradeOffer(i).Num = 0
        TempPlayer(TradeTarget).TradeOffer(i).Value = 0
        TempPlayer(TradeTarget).TradeOffer(i).Bind = 0
        TempPlayer(TradeTarget).TradeOffer(i).Durability = 0
    Next
    
    ' Used to init the trade window clientside
    SendTrade index, TradeTarget
    SendTrade TradeTarget, index
    
    ' Send the offer data - Used to clear their client
    SendTradeUpdate index, 0
    SendTradeUpdate index, 1
    SendTradeUpdate TradeTarget, 0
    SendTradeUpdate TradeTarget, 1
End Sub

Sub HandleDeclineTradeRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call DeclineTradeRequest(index)
End Sub

Sub HandleAcceptTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim TradeTarget As Long
    Dim i As Long
    Dim TmpTradeItem(1 To MAX_INV) As PlayerItemRec
    Dim TmpTradeItem2(1 To MAX_INV) As PlayerItemRec
    Dim ItemNum As Integer
    
    TempPlayer(index).AcceptTrade = True
    TradeTarget = TempPlayer(index).InTrade
    
    ' If not both of them accept, then exit
    If Not TempPlayer(TradeTarget).AcceptTrade Then
        SendTradeStatus index, 2
        SendTradeStatus TradeTarget, 1
        Exit Sub
    End If
    
    ' Take their items
    For i = 1 To MAX_INV
        ' Player
        If TempPlayer(index).TradeOffer(i).Num > 0 Then
            ItemNum = Account(index).Chars(GetPlayerChar(index)).Inv(TempPlayer(index).TradeOffer(i).Num).Num
            If ItemNum > 0 Then
                ' Store temp
                TmpTradeItem(i).Num = ItemNum
                TmpTradeItem(i).Value = TempPlayer(index).TradeOffer(i).Value
                TmpTradeItem(i).Bind = TempPlayer(index).TradeOffer(i).Bind
                TmpTradeItem(i).Durability = TempPlayer(index).TradeOffer(i).Durability
                
                ' Take item
                TakeInvSlot index, TempPlayer(index).TradeOffer(i).Num, TmpTradeItem(i).Value, False
            End If
        End If
        
        ' Target
        If TempPlayer(TradeTarget).TradeOffer(i).Num > 0 Then
            ItemNum = GetPlayerInvItemNum(TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num)
            If ItemNum > 0 Then
                ' Store temp
                TmpTradeItem2(i).Num = ItemNum
                TmpTradeItem2(i).Value = TempPlayer(TradeTarget).TradeOffer(i).Value
                TmpTradeItem2(i).Bind = TempPlayer(TradeTarget).TradeOffer(i).Bind
                TmpTradeItem2(i).Durability = TempPlayer(TradeTarget).TradeOffer(i).Durability
                
                ' Take item
                TakeInvSlot TradeTarget, TempPlayer(TradeTarget).TradeOffer(i).Num, TmpTradeItem2(i).Value, False
            End If
        End If
    Next
    
    ' Taken all items, now they can't get items because of no inventory space
    For i = 1 To MAX_INV
        ' Player
        If TmpTradeItem2(i).Num > 0 Then
            ' Give away
            GiveInvItem index, TmpTradeItem2(i).Num, TmpTradeItem2(i).Value, -1, 0, False
        End If
        
        ' Target
        If TmpTradeItem(i).Num > 0 Then
            ' Give away
            GiveInvItem TradeTarget, TmpTradeItem(i).Num, TmpTradeItem(i).Value, -1, 0, False
        End If
    Next
    
    ' Refresh inventory
    SendInventory index
    SendInventory TradeTarget
    
    ' They now have all the items. Clear out values + let them out of the trade.
    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).Num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(index).TradeOffer(i).Bind = 0
        TempPlayer(index).TradeOffer(i).Durability = 0
        TempPlayer(TradeTarget).TradeOffer(i).Num = 0
        TempPlayer(TradeTarget).TradeOffer(i).Value = 0
        TempPlayer(TradeTarget).TradeOffer(i).Bind = 0
        TempPlayer(TradeTarget).TradeOffer(i).Durability = 0
    Next

    TempPlayer(index).InTrade = 0
    TempPlayer(TradeTarget).InTrade = 0
    
    PlayerMsg index, "Trade completed.", BrightGreen
    PlayerMsg TradeTarget, "Trade completed.", BrightGreen
    
    SendCloseTrade index
    SendCloseTrade TradeTarget
End Sub

Sub HandleDeclineTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim TradeTarget As Long

    TradeTarget = TempPlayer(index).InTrade

    For i = 1 To MAX_INV
        TempPlayer(index).TradeOffer(i).Num = 0
        TempPlayer(index).TradeOffer(i).Value = 0
        TempPlayer(index).TradeOffer(i).Bind = 0
        TempPlayer(index).TradeOffer(i).Durability = 0
        TempPlayer(TradeTarget).TradeOffer(i).Num = 0
        TempPlayer(TradeTarget).TradeOffer(i).Value = 0
        TempPlayer(TradeTarget).TradeOffer(i).Bind = 0
        TempPlayer(TradeTarget).TradeOffer(i).Durability = 0
    Next

    TempPlayer(index).InTrade = 0
    TempPlayer(TradeTarget).InTrade = 0
    
    PlayerMsg index, "You declined the trade.", BrightRed
    PlayerMsg TradeTarget, GetPlayerName(index) & " has declined the trade!", BrightRed
    
    SendCloseTrade index
    SendCloseTrade TradeTarget
End Sub

Sub HandleTradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim InvSlot As Byte
    Dim Amount As Long
    Dim EmptySlot As Long
    Dim ItemNum As Integer
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    InvSlot = buffer.ReadByte
    Amount = buffer.ReadLong
    
    Set buffer = Nothing
    
    If InvSlot <= 0 Or InvSlot > MAX_INV Then Exit Sub
    
    ItemNum = GetPlayerInvItemNum(index, InvSlot)
    If ItemNum <= 0 Or ItemNum > MAX_ITEMS Then Exit Sub
    
    ' Hack prevention
    If Item(GetPlayerInvItemNum(index, InvSlot)).Stackable = 1 Then
        If GetPlayerInvItemValue(index, InvSlot) < Amount Then Amount = GetPlayerInvItemValue(index, InvSlot)
        If Amount < 1 Then Exit Sub
    Else
        If Not Amount = 0 Then Exit Sub
    End If

    If Item(ItemNum).Stackable = 1 Then
        ' Check if already offering same currency item
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).Num = InvSlot Then
                ' Add amount
                TempPlayer(index).TradeOffer(i).Value = TempPlayer(index).TradeOffer(i).Value + Amount
                
                ' Clamp to limits
                If TempPlayer(index).TradeOffer(i).Value > GetPlayerInvItemValue(index, InvSlot) Then
                    TempPlayer(index).TradeOffer(i).Value = GetPlayerInvItemValue(index, InvSlot)
                End If
                
                TempPlayer(index).TradeOffer(i).Bind = GetPlayerInvItemBind(index, InvSlot)
                TempPlayer(index).TradeOffer(i).Durability = GetPlayerInvItemDur(index, InvSlot)
                
                ' Cancel any trade agreement
                TempPlayer(index).AcceptTrade = False
                TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
                
                SendTradeStatus index, 0
                SendTradeStatus TempPlayer(index).InTrade, 0
                
                SendTradeUpdate index, 0
                SendTradeUpdate TempPlayer(index).InTrade, 1
                ' Exit early
                Exit Sub
            End If
        Next
    Else
        ' Make sure they're not already offering it
        For i = 1 To MAX_INV
            If TempPlayer(index).TradeOffer(i).Num = InvSlot Then
                PlayerMsg index, "You've already offered this item.", BrightRed
                Exit Sub
            End If
        Next
    End If
    
    ' Not already offering - find earliest empty slot
    For i = 1 To MAX_INV
        If TempPlayer(index).TradeOffer(i).Num = 0 Then
            EmptySlot = i
            Exit For
        End If
    Next
    
    TempPlayer(index).TradeOffer(EmptySlot).Num = InvSlot
    TempPlayer(index).TradeOffer(EmptySlot).Value = Amount
    TempPlayer(index).TradeOffer(EmptySlot).Bind = GetPlayerInvItemBind(index, InvSlot)
    TempPlayer(index).TradeOffer(EmptySlot).Durability = GetPlayerInvItemDur(index, InvSlot)
    
    ' Cancel any trade agreement and send new data
    TempPlayer(index).AcceptTrade = False
    TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleUntradeItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim TradeSlot As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    TradeSlot = buffer.ReadByte
    Set buffer = Nothing
    
    ' Make sure there in trade
    If TempPlayer(index).InTrade = 0 Then Exit Sub
    
    If TradeSlot < 1 Or TradeSlot > MAX_INV Then Exit Sub
    If TempPlayer(index).TradeOffer(TradeSlot).Num < 1 Then Exit Sub
    
    TempPlayer(index).TradeOffer(TradeSlot).Num = 0
    TempPlayer(index).TradeOffer(TradeSlot).Value = 0
    TempPlayer(index).TradeOffer(TradeSlot).Bind = 0
    TempPlayer(index).TradeOffer(TradeSlot).Durability = 0
    
    If TempPlayer(index).AcceptTrade Then TempPlayer(index).AcceptTrade = False
    If TempPlayer(TempPlayer(index).InTrade).AcceptTrade Then TempPlayer(TempPlayer(index).InTrade).AcceptTrade = False
    
    SendTradeStatus index, 0
    SendTradeStatus TempPlayer(index).InTrade, 0
    
    SendTradeUpdate index, 0
    SendTradeUpdate TempPlayer(index).InTrade, 1
End Sub

Sub HandleHotbarChange(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim SType As Byte
    Dim Slot As Byte
    Dim HotbarNum As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    SType = buffer.ReadByte
    Slot = buffer.ReadByte
    HotbarNum = buffer.ReadByte
    
    If HotbarNum < 1 Or HotbarNum > MAX_HOTBAR Then Exit Sub
    
    Select Case SType
        Case 0 ' Clear
            Account(index).Chars(GetPlayerChar(index)).Hotbar(HotbarNum).Slot = 0
            Account(index).Chars(GetPlayerChar(index)).Hotbar(HotbarNum).SType = 0
        Case 1 ' Inventory
            If Slot > 0 And Slot <= MAX_INV Then
                ' Don't add None/Currency/Auto Life type items
                If Item(GetPlayerInvItemNum(index, Slot)).Stackable = 1 Or Item(GetPlayerInvItemNum(index, Slot)).Type = ITEM_TYPE_NONE Or Item(GetPlayerInvItemNum(index, Slot)).Type = ITEM_TYPE_AUTOLIFE Then Exit Sub
                
                If Account(index).Chars(GetPlayerChar(index)).Inv(Slot).Num > 0 Then
                    If Len(Trim$(Item(GetPlayerInvItemNum(index, Slot)).Name)) > 0 Then
                        Account(index).Chars(GetPlayerChar(index)).Hotbar(HotbarNum).Slot = Account(index).Chars(GetPlayerChar(index)).Inv(Slot).Num
                        Account(index).Chars(GetPlayerChar(index)).Hotbar(HotbarNum).SType = SType
                    End If
                End If
            End If
        Case 2 ' Spell
            If Slot > 0 And Slot <= MAX_PLAYER_SPELLS Then
                If Account(index).Chars(GetPlayerChar(index)).Spell(Slot) > 0 Then
                    If Len(Trim$(Spell(Account(index).Chars(GetPlayerChar(index)).Spell(Slot)).Name)) > 0 Then
                        Account(index).Chars(GetPlayerChar(index)).Hotbar(HotbarNum).Slot = Account(index).Chars(GetPlayerChar(index)).Spell(Slot)
                        Account(index).Chars(GetPlayerChar(index)).Hotbar(HotbarNum).SType = SType
                    End If
                End If
            End If
    End Select
    
    SendHotbar index
    
    Set buffer = Nothing
End Sub

Sub HandlePartyRequest(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, Name As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Name = buffer.ReadString
    
    ' Check if it is invalid
    If Name = vbNullString Then Exit Sub
    If FindPlayer(Name) = index Then Exit Sub
    
    If IsPlaying(FindPlayer(Name)) = False Then
        Call PlayerMsg(index, "Player is not online!", BrightRed)
        Exit Sub
    End If
    
    ' Can't invite if the player is a foe
    If IsAFoe(index, FindPlayer(Name)) Then Exit Sub
    
    ' Init the request
    Party_Invite index, FindPlayer(Name)
End Sub

Sub HandleAcceptParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteAccept TempPlayer(index).PartyInvite, index
End Sub

Sub HandleDeclineParty(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_InviteDecline TempPlayer(index).PartyInvite, index
End Sub

Sub HandlePartyLeave(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Party_PlayerLeave index
End Sub

Sub HandlePartyMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    
    ' Make sure there in a party
    If TempPlayer(index).InParty = 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Msg = buffer.ReadString
    
    If Msg = vbNullString Then Exit Sub
    
    If Trim$(Account(index).Chars(GetPlayerChar(index)).Status) = "Muted" Then
        Call PlayerMsg(index, "You are muted!", BrightRed)
        Exit Sub
    End If
    
    PartyMsg TempPlayer(index).InParty, Msg, BrightBlue
    Set buffer = Nothing
End Sub

Sub HandleAdminMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_MODERATOR Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Msg = buffer.ReadString
    
    If Msg = vbNullString Then Exit Sub
    
    If Trim$(Account(index).Chars(GetPlayerChar(index)).Status) = "Muted" Then
        Call PlayerMsg(index, "You are muted!", BrightRed)
        Exit Sub
    End If
    
    Call AdminMsg(Msg, BrightCyan)
    Set buffer = Nothing
End Sub

Sub HandleGuildCreate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, Name As String, i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Name = Trim$(buffer.ReadString)
    Set buffer = Nothing
    
    If Len(Name) > NAME_LENGTH Then
        Call PlayerMsg(index, "You have entered a guild name that is too long!", BrightRed)
        Exit Sub
    End If

    For i = 1 To MAX_GUILDS
        If Trim$(LCase$(Guild(i).Name)) = Name Then
            Call PlayerMsg(index, "This guild name has already been used!", BrightRed)
            Exit Sub
        End If
    Next

    If HasItem(index, 1) < Options.GuildCost Then
        Call PlayerMsg(index, "You do not have enough " & Trim$(Item(1).Name) & " to purchase a guild!", BrightRed)
        Exit Sub
    Else
        For i = 1 To MAX_GUILDS
            If Len(Trim$(Guild(i).Name)) = 0 Then
                Guild(i).Name = Name
                Guild(i).Members(1) = GetPlayerLogin(index)
                Call SetPlayerGuild(index, i)
                Call SetPlayerGuildAccess(index, MAX_GUILDACCESS)
                Call TakeInvItem(index, 1, Options.GuildCost)
                Call GlobalMsg(GetPlayerName(index) & " has founded the guild " & Name & "!", Yellow)
                Call SendPlayerGuild(index)
                Call SaveGuilds
                Call SaveAccount(index)
                Exit Sub
            End If
        Next
        
        Call PlayerMsg(index, "There are too many guilds already! You must join another guild or wait until the amount of guilds permitted is increased.", BrightRed)
    End If
End Sub

Sub HandleGuildInvite(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, Name As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Name = buffer.ReadString
    
    ' Check if it is invalid
    If Name = vbNullString Then Exit Sub
    If FindPlayer(Name) = index Then Exit Sub
    
    If IsPlaying(FindPlayer(Name)) = False Then
        Call PlayerMsg(index, "Player is not online!", BrightRed)
        Exit Sub
    End If
    
    ' Make sure they are actually in a guild
    If GetPlayerGuild(index) = 0 Then Exit Sub
    
    ' Can't invite if the player is a foe
    If IsAFoe(index, FindPlayer(Name)) = True Then Exit Sub
    
    ' Init the request
    Guild_Invite index, FindPlayer(Name)
End Sub

Sub Guild_Invite(ByVal index As Long, ByVal OtherPlayer As Long)
    ' Is the other player in a guild already
    If GetPlayerGuild(OtherPlayer) > 0 Then
        Call PlayerMsg(index, GetPlayerName(OtherPlayer) & " is already in a guild!", BrightRed)
        Exit Sub
    End If
    
    ' Check if there doing another action
    If IsPlayerBusy(index, OtherPlayer) Then Exit Sub
    
    ' Make sure they have a high enough access
    If GetPlayerGuildAccess(index) < 2 Then
        Call PlayerMsg(index, "You are not allowed to invite members to the guild!", BrightRed)
        Exit Sub
    End If
    
    ' Send the invite
    Call SendGuildInvite(index, OtherPlayer)
    
    ' Set the invite target
    TempPlayer(OtherPlayer).GuildInvite = index
    
    ' Let them know
    PlayerMsg index, "Guild invitation sent.", Pink
End Sub

Sub HandleGuildRemove(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, Name As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Name = buffer.ReadString
    
    ' Check if it is invalid
    If Name = vbNullString Then Exit Sub
    If FindPlayer(Name) = index Then Exit Sub
    
    ' Make sure they are actually in a guild
    If GetPlayerGuild(index) = 0 Or GetPlayerGuild(FindPlayer(Name)) = 0 Then Exit Sub
    
    ' Init the request
    Guild_Remove index, FindPlayer(Name)
End Sub

Sub Guild_Remove(ByVal index As Long, ByVal OtherPlayer As Long)
    Dim i As Long
    
    If IsPlaying(index) = False Then
        Call PlayerMsg(index, "Player is not online!", BrightRed)
        Exit Sub
    End If
    
    ' Is the other player not in a guild
    If GetPlayerGuild(OtherPlayer) = 0 Then
        Call PlayerMsg(index, GetPlayerName(OtherPlayer) & " is not in a guild!", BrightRed)
        Exit Sub
    End If
    
    ' Is the other player not in our guild
    If Not GetPlayerGuild(OtherPlayer) = GetPlayerGuild(index) Then
        Call PlayerMsg(index, GetPlayerName(OtherPlayer) & " is not in our guild!", BrightRed)
        Exit Sub
    End If

    ' Make sure they have a high enough access
    If GetPlayerGuildAccess(index) < 2 Then
        Call PlayerMsg(index, "You are not allowed to remove other guild members!", BrightRed)
        Exit Sub
    End If

    ' Can't remove someone from guild if they have a higher access
    If GetPlayerGuildAccess(index) <= GetPlayerGuildAccess(OtherPlayer) Then
        Call PlayerMsg(index, "You can't change the guild rank of someone who has same or higher rank!", BrightRed)
        Exit Sub
    End If
    
    Call GuildMsg(index, GetPlayerName(OtherPlayer) & " has been removed from the guild by " & GetPlayerName(index) & "!", BrightRed, True)
    
    ' Remove them
    Call SetPlayerGuild(OtherPlayer, 0)
    Call SetPlayerGuildAccess(OtherPlayer, 0)
    
    ' Send the update
    Call SendPlayerGuild(OtherPlayer)
    
    ' Update other player's guild information
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerGuild(i) = GetPlayerGuild(index) Then
                SendPlayerGuildMembers i
            End If
        End If
    Next
End Sub

Sub HandleGuildChangeAccess(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, Name As String, X As Long, i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Name = buffer.ReadString
    X = buffer.ReadByte
    i = FindPlayer(Name)
    Set buffer = Nothing
    
    ' Make sure they are actually in a guild
    If GetPlayerGuild(index) = 0 Or GetPlayerGuild(i) = 0 Then Exit Sub
    
    If X < 1 Or X > MAX_GUILDACCESS Then
        Call PlayerMsg(index, "Invalid access level!", BrightRed)
        Exit Sub
    End If
    
    If Not IsPlaying(i) Then
        Call PlayerMsg(index, "Player is not online!", BrightRed)
        Exit Sub
    End If
    
    If i = index Then
        Call PlayerMsg(index, "You can't change your own access!", BrightRed)
        Exit Sub
    End If
    
    If X < GetPlayerGuildAccess(index) Then
        If X = GetPlayerGuildAccess(i) Then
            Call PlayerMsg(index, "That player is already that access level!", BrightRed)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(index) < 3 Then
            Call PlayerMsg(index, "You need to have a higher guild rank to change that player's rank!", BrightRed)
            Exit Sub
        End If
        
        If GetPlayerGuildAccess(index) <= GetPlayerGuildAccess(i) Then
            PlayerMsg index, "You can't change the guild rank of someone who has the same or higher rank!", BrightRed
            Exit Sub
        End If
        
        ' Set access
        Call SetPlayerGuildAccess(i, X)

        Call GuildMsg(i, GetPlayerName(index) & " has changed " & GetPlayerName(i) & "'s guild rank to " & X & "!", Yellow, True)
    Else
        Call PlayerMsg(index, "You can't promote players to the same or higher guild rank as yourself!", BrightRed)
        Exit Sub
    End If
    
    ' Send guild to player
    Call SendPlayerGuild(i)
End Sub

Sub HandleAcceptGuild(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    
    Call GuildMsg(index, GetPlayerName(index) & " has joined " & Trim$(Guild(Account(TempPlayer(index).GuildInvite).Chars(GetPlayerChar(index)).Guild.index).Name) & "!", Yellow, True)
    Call SetPlayerGuildAccess(index, 1)
    Call SetPlayerGuild(index, GetPlayerGuild(TempPlayer(index).GuildInvite))
    TempPlayer(index).GuildInvite = 0
    
     ' Send data
    Call SendPlayerGuild(index)
    
    ' Update other player's guild information
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If GetPlayerGuild(i) = GetPlayerGuild(index) Then
                SendPlayerGuildMembers i
            End If
        End If
    Next
End Sub

Sub HandleDeclineGuild(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call DeclineGuildInvite(index)
End Sub

Sub HandleGuildDisband(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Guild_Disband index
End Sub

Sub HandleGuildMsg(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    
    ' Can't send messgae if not in a guild
    If GetPlayerGuild(index) = 0 Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Msg = buffer.ReadString
    Set buffer = Nothing
    
    If Msg = vbNullString Then Exit Sub
    
    If Trim$(Account(index).Chars(GetPlayerChar(index)).Status) = "Muted" Then
        Call PlayerMsg(index, "You are muted!", BrightRed)
        Exit Sub
    End If
    
    Call GuildMsg(index, Msg, Green)
End Sub

Sub HandleBreakSpell(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If TempPlayer(index).SpellBuffer.Spell > 0 Then
        Call SendActionMsg(GetPlayerMap(index), "Interrupted", BrightRed, ACTIONMSG_SCROLL, GetPlayerX(index) * 32, GetPlayerY(index) * 32)
        Call ClearAccountSpellBuffer(index)
    End If
End Sub

Sub HandleCanTrade(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If Account(index).Chars(GetPlayerChar(index)).CanTrade = False Then
        Call PlayerMsg(index, "Other players are now able to trade with you.", BrightGreen)
        Account(index).Chars(GetPlayerChar(index)).CanTrade = True
    Else
        Call PlayerMsg(index, "Other players are now unable to trade with you.", BrightRed)
        Account(index).Chars(GetPlayerChar(index)).CanTrade = False
    End If
End Sub

Sub HandleAddFriend(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Name = buffer.ReadString
    Set buffer = Nothing
    
    ' Make sure the name isn't empty
    If Trim$(Name) = vbNullString Then
        Call PlayerMsg(index, "Invalid name!", BrightRed)
        Exit Sub
    End If
    
    ' Check to see if they have more friends then they can hold
    If Account(index).Friends.AmountOfFriends = MAX_PEOPLE Then
        Call PlayerMsg(index, "Tour friends list is full!", BrightRed)
        Exit Sub
    End If
    
    ' See if character exists
    If FindPlayer(Name) = 0 Then
        Call PlayerMsg(index, "Player is not online!", 12)
        Exit Sub
    End If
    
    If FindPlayer(Name) = index Then
        Call PlayerMsg(index, "You can't add yourself as a friend!", 12)
        Exit Sub
    End If
    
    If GetPlayerAccess(FindPlayer(Name)) > STAFF_MODERATOR Then
        Call PlayerMsg(index, "You can't add a friend who is a staff member!", BrightRed)
        Exit Sub
    End If
    
    ' Check if they already have that as their friend
    If Account(index).Friends.AmountOfFriends > 0 Then
        For i = 1 To Account(index).Friends.AmountOfFriends
            If Trim$(Account(index).Friends.Members(i)) = Name Then
                Call PlayerMsg(index, "You already have that player as your friend!", 12)
                Exit Sub
            End If
        Next
    End If
    
    ' Add friend to List
    If Trim$(Account(index).Friends.Members(Account(index).Friends.AmountOfFriends + 1)) = vbNullString Then
        Account(index).Friends.Members(Account(index).Friends.AmountOfFriends + 1) = Name
        Account(index).Friends.AmountOfFriends = Account(index).Friends.AmountOfFriends + 1
        Call PlayerMsg(index, "You have added " & Trim$(Account(index).Friends.Members(Account(index).Friends.AmountOfFriends)) & " to your friends list!", BrightGreen)
    End If
   
    ' Update Friend List
    Call UpdateFriendsList(index)
End Sub

Sub HandleRemoveFriend(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim i As Long, X As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Name = buffer.ReadString
    Set buffer = Nothing
   
    ' If the name is blank then exit
    If Name = vbNullString Then Exit Sub
    
    ' If they don't have any friends then exit
    If Account(index).Friends.AmountOfFriends = 0 Then
        Call PlayerMsg(index, "You don't have any friends to remove!", BrightRed)
        Exit Sub
    End If
    
    X = 0
    
    For i = 1 To Account(index).Friends.AmountOfFriends
        If Trim$(Account(index).Friends.Members(i)) = Name Then
            X = 1
            Exit For
        End If
    Next
    
    If Not X = 1 Then
        Call PlayerMsg(index, "You don't have a friend with that name!", BrightRed)
    End If
    
    For i = 1 To Account(index).Friends.AmountOfFriends
        If Trim$(Account(index).Friends.Members(i)) = Name Then
            ' They successfully removed the friend, send the message
            Call PlayerMsg(index, "You have removed " & Trim$(Account(index).Friends.Members(i)) & " from your friends list!", BrightRed)
            Account(index).Friends.Members(i) = vbNullString
            Account(index).Friends.AmountOfFriends = Account(index).Friends.AmountOfFriends - 1
            Exit For
        End If
    Next
   
    ' Update Friend List
    Call UpdateFriendsList(index)
End Sub

Sub HandleUpdateFriendsList(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call UpdateFriendsList(index)
End Sub

Sub HandleAddFoe(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Name = buffer.ReadString
    Set buffer = Nothing
    
    ' Make sure the name isn't empty
    If Trim$(Name) = vbNullString Then
        Call PlayerMsg(index, "Invalid name!", BrightRed)
        Exit Sub
    End If
    
    ' Check to see if they have more Foes then they can hold
    If Account(index).Foes.Amount = MAX_PEOPLE Then
        Call PlayerMsg(index, "Tour foes list is full!", BrightRed)
        Exit Sub
    End If
    
    ' See if character exists
    If FindPlayer(Name) = 0 Then
        Call PlayerMsg(index, "Player is not online!", 12)
        Exit Sub
    End If
    
    If FindPlayer(Name) = index Then
        Call PlayerMsg(index, "You can't add yourself as a foe!", 12)
        Exit Sub
    End If
    
    If GetPlayerAccess(FindPlayer(Name)) > STAFF_MODERATOR Then
        Call PlayerMsg(index, "You can't add a foe who is a staff member!", BrightRed)
        Exit Sub
    End If
    
    ' Check if they already have that as their Foe
    If Account(index).Foes.Amount > 0 Then
        For i = 1 To Account(index).Foes.Amount
            If Trim$(Account(index).Foes.Members(i)) = Name Then
                Call PlayerMsg(index, "You already have that player as your foe!", 12)
                Exit Sub
            End If
        Next
    End If
    
    ' Add Foe to List
    If Trim$(Account(index).Foes.Members(Account(index).Foes.Amount + 1)) = vbNullString Then
        Account(index).Foes.Members(Account(index).Foes.Amount + 1) = Name
        Account(index).Foes.Amount = Account(index).Foes.Amount + 1
        Call PlayerMsg(index, "You have added " & Trim$(Account(index).Foes.Members(Account(index).Foes.Amount)) & " to your foes list!", BrightGreen)
    End If
   
    ' Update Foe List
    Call UpdateFoesList(index)
End Sub

Sub HandleRemoveFoe(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Name As String
    Dim i As Long, X As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Name = buffer.ReadString
    Set buffer = Nothing
   
    ' If the name is blank then exit
    If Name = vbNullString Then Exit Sub
    
    ' If they don't have any Foes then exit
    If Account(index).Foes.Amount = 0 Then
        Call PlayerMsg(index, "You don't have any foes to remove!", BrightRed)
        Exit Sub
    End If
    
    X = 0
    
    For i = 1 To Account(index).Foes.Amount
        If Account(index).Foes.Members(i) = Name Then
            X = 1
            Exit For
        End If
    Next
    
    If Not X = 1 Then
        Call PlayerMsg(index, "You don't have a foe with that name!", BrightRed)
    End If
    
    For i = 1 To Account(index).Foes.Amount
        If Trim$(Account(index).Foes.Members(i)) = Name Then
            ' They successfully removed the foe, send the message
            Call PlayerMsg(index, "You have removed " & Trim$(Account(index).Foes.Members(i)) & " from your foes list!", BrightRed)
            Account(index).Foes.Members(i) = vbNullString
            Account(index).Foes.Amount = Account(index).Foes.Amount - 1
            Exit For
        End If
    Next
   
    ' Update Foe List
    Call UpdateFoesList(index)
End Sub

Sub HandleUpdateFoesList(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Call UpdateFoesList(index)
End Sub

Private Sub HandleUpdateData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    TempPlayer(index).HDSerial = buffer.ReadString
    
    ' Close any clients that have the same serial
    For i = 1 To Player_HighIndex
        If Not i = index And Options.MultipleSerial = 0 Then
            If GetPlayerHDSerial(i) = GetPlayerHDSerial(index) Then
                Call SendCloseClient(index)
                Exit Sub
            End If
        End If
    Next
    
    ' Check version
    If Not App.Major = buffer.ReadLong Or Not App.Minor = buffer.ReadLong Or Not App.Revision = buffer.ReadLong Then
        Call AlertMsg(index, "Version outdated, please visit " & Options.Website & " for more information on new releases and run the updater.")
    End If
    
    ' Send the news
    Call SendNews(index)
    
    ' Send classes
    Call SendClasses(index)
    
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Ban packet ::
' ::::::::::::::::::::::
Sub HandleSaveBan(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim BanSize As Long
    Dim BanData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_ADMIN Then Exit Sub

    n = buffer.ReadLong

    If n < 1 Or n > MAX_BANS Then Exit Sub

    ' Update the Ban
    BanSize = LenB(Ban(n))
    ReDim BanData(BanSize - 1)
    BanData = buffer.ReadBytes(BanSize)
    CopyMemory ByVal VarPtr(Ban(n)), ByVal VarPtr(BanData(0)), BanSize
    Set buffer = Nothing
    
    ' Save it
    Call SaveBan(n)
    Call AddLog(GetPlayerName(index) & " saved Ban #" & n & ".", "Staff")
End Sub

Sub HandleBanEditor(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_ADMIN Then Exit Sub
    
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SBanEditor
    Call SendDataTo(index, buffer.ToArray())
    Set buffer = Nothing
End Sub

Sub HandleSetTitle(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim TitleNum As Byte
   
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
   
    TitleNum = buffer.ReadByte
    
    ' Check for an invalid title
    If TitleNum < 0 Or TitleNum > MAX_TITLES Then Exit Sub
    
    ' Make sure they have the title
    If Not TitleNum = 0 Then
        If Account(index).Chars(GetPlayerChar(index)).Title(TitleNum) = 0 Then Exit Sub
    End If
    
    ' Set the current title
    Account(index).Chars(GetPlayerChar(index)).CurrentTitle = TitleNum

    ' Send updated title to map
    Call SendPlayerTitles(index)
End Sub

' ::::::::::::::::::::::
' :: Save Title packet ::
' ::::::::::::::::::::::
Sub HandleSaveTitle(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim TitleSize As Long
    Dim TitleData() As Byte
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    n = buffer.ReadLong

    If n < 1 Or n > MAX_TITLES Then Exit Sub

    ' Update the Title
    TitleSize = LenB(Title(n))
    ReDim TitleData(TitleSize - 1)
    TitleData = buffer.ReadBytes(TitleSize)
    CopyMemory ByVal VarPtr(Title(n)), ByVal VarPtr(TitleData(0)), TitleSize
    Set buffer = Nothing
    
    ' Save it
    Call SaveTitle(n)
    Call SendUpdateTitleToAll(n)
    Call AddLog(GetPlayerName(index) & " saved Title #" & n & ".", "Staff")
End Sub

Sub HandleTitleEditor(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub
    
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong STitleEditor
    Call SendDataTo(index, buffer.ToArray())
    Set buffer = Nothing
End Sub

Sub HandleChangeStatus(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As New clsBuffer
    Dim Status As String
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    Status = buffer.ReadString
    
    If Trim$(Account(index).Chars(GetPlayerChar(index)).Status) = "Muted" Then Exit Sub
    
    Account(index).Chars(GetPlayerChar(index)).Status = Status
    Call SendPlayerStatus(index)
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Moral packet ::
' ::::::::::::::::::::::
Sub HandleSaveMoral(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim MoralSize As Long
    Dim MoralData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    n = buffer.ReadLong

    If n < 1 Or n > MAX_MORALS Then Exit Sub

    ' Update the Moral
    MoralSize = LenB(Moral(n))
    ReDim MoralData(MoralSize - 1)
    MoralData = buffer.ReadBytes(MoralSize)
    CopyMemory ByVal VarPtr(Moral(n)), ByVal VarPtr(MoralData(0)), MoralSize
    Set buffer = Nothing
    
    ' Save it
    Call SaveMoral(n)
    Call SendUpdateMoralToAll(n)
    Call AddLog(GetPlayerName(index) & " saved Moral #" & n & ".", "Staff")
End Sub

Sub HandleMoralEditor(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub
    
    Dim buffer As clsBuffer
    Set buffer = New clsBuffer
    
    buffer.WriteLong SMoralEditor
    Call SendDataTo(index, buffer.ToArray())
    Set buffer = Nothing
End Sub

' ::::::::::::::::::::::
' :: Save Class packet ::
' ::::::::::::::::::::::
Sub HandleSaveClass(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long, i As Long
    Dim buffer As clsBuffer
    Dim Classesize As Long
    Dim ClassData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    n = buffer.ReadLong

    If n < 1 Or n > MAX_CLASSES Then Exit Sub

    ' Update the Class
    Classesize = LenB(Class(n))
    ReDim ClassData(Classesize - 1)
    ClassData = buffer.ReadBytes(Classesize)
    CopyMemory ByVal VarPtr(Class(n)), ByVal VarPtr(ClassData(0)), Classesize
    Set buffer = Nothing
    
    ' Save it
    Call SaveClass(n)
    
    For i = 1 To Player_HighIndex
        If IsConnected(i) Then
            If Len(Trim$(Class(n).Name)) > 0 Then
                Call SendUpdateClassTo(i, n)
            End If
        End If
    Next
    
    Call AddLog(GetPlayerName(index) & " saved Class #" & n & ".", "Staff")
End Sub

Sub HandleClassEditor(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
        
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SClassEditor
    Call SendDataTo(index, buffer.ToArray())
    Set buffer = Nothing
End Sub

Sub HandleDestroyItem(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim InvNum As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    InvNum = buffer.ReadInteger

    ' Prevent subscript out of range
    If InvNum < 1 Or InvNum > MAX_INV Then Exit Sub
    
    Call TakeInvSlot(index, InvNum, 1, True)
End Sub

' :::::::::::::::::::::::::
' :: Save Emoticon packet ::
' :::::::::::::::::::::::::
Sub HandleSaveEmoticon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long, i As Long
    Dim buffer As clsBuffer
    Dim EmoticonSize As Long
    Dim EmoticonData() As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub

    n = buffer.ReadLong

    If n < 1 Or n > MAX_EMOTICONS Then Exit Sub

    ' Update the Emoticon
    EmoticonSize = LenB(Emoticon(n))
    ReDim EmoticonData(EmoticonSize - 1)
    EmoticonData = buffer.ReadBytes(EmoticonSize)
    CopyMemory ByVal VarPtr(Emoticon(n)), ByVal VarPtr(EmoticonData(0)), EmoticonSize
    Set buffer = Nothing
    
    ' Save it
    Call SaveEmoticon(n)
    Call SendUpdateEmoticonToAll(n)
    Call AddLog(GetPlayerName(index) & " saved Emoticon #" & n & ".", "Staff")
End Sub

Sub HandleEmoticonEditor(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
        
    ' Prevent hacking
    If GetPlayerAccess(index) < STAFF_DEVELOPER Then Exit Sub
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong SEmoticonEditor
    Call SendDataTo(index, buffer.ToArray())
    Set buffer = Nothing
End Sub

Private Sub HandleCheckEmoticon(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, EmoticonNum As Byte
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    EmoticonNum = buffer.ReadLong
    
    ' Subscript out of range
    If EmoticonNum < 1 Or EmoticonNum > MAX_EMOTICONS Then Exit Sub
    
    SendCheckEmoticon index, GetPlayerMap(index), EmoticonNum
End Sub

Sub HandleEventChatReply(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim EventID As Long, PageID As Long, reply As Long, i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    EventID = buffer.ReadLong
    PageID = buffer.ReadLong
    reply = buffer.ReadLong
    
    If TempPlayer(index).EventProcessingCount > 0 Then
        For i = 1 To TempPlayer(index).EventProcessingCount
            If TempPlayer(index).EventProcessing(i).EventID = EventID And TempPlayer(index).EventProcessing(i).PageID = PageID Then
                If TempPlayer(index).EventProcessing(i).WaitingForResponse = 1 Then
                    If reply = 0 Then
                        If Map(GetPlayerMap(index)).Events(EventID).Pages(PageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).index = EventType.evShowText Then
                            TempPlayer(index).EventProcessing(i).WaitingForResponse = 0
                        End If
                    ElseIf reply > 0 Then
                        If Map(GetPlayerMap(index)).Events(EventID).Pages(PageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).index = EventType.evShowChoices Then
                            Select Case reply
                                Case 1
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(EventID).Pages(PageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data1
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                                Case 2
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(EventID).Pages(PageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data2
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                                Case 3
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(EventID).Pages(PageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data3
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                                Case 4
                                    TempPlayer(index).EventProcessing(i).ListLeftOff(TempPlayer(index).EventProcessing(i).CurList) = TempPlayer(index).EventProcessing(i).CurSlot
                                    TempPlayer(index).EventProcessing(i).CurList = Map(GetPlayerMap(index)).Events(EventID).Pages(PageID).CommandList(TempPlayer(index).EventProcessing(i).CurList).Commands(TempPlayer(index).EventProcessing(i).CurSlot - 1).Data4
                                    TempPlayer(index).EventProcessing(i).CurSlot = 1
                            End Select
                        End If
                        TempPlayer(index).EventProcessing(i).WaitingForResponse = 0
                    End If
                End If
            End If
        Next
    End If
    Set buffer = Nothing
End Sub

Sub HandleEvent(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim n As Long
    Dim Damage As Long
    Dim TempIndex As Long
    Dim X As Long, Y As Long, BeginEventProcessing As Boolean, z As Long, buffer As clsBuffer

    ' Check tradeskills
    Select Case GetPlayerDir(index)
        Case DIR_UP

            If GetPlayerY(index) = 0 Then Exit Sub
            X = GetPlayerX(index)
            Y = GetPlayerY(index) - 1
        Case DIR_DOWN

            If GetPlayerY(index) = Map(GetPlayerMap(index)).MaxY Then Exit Sub
            X = GetPlayerX(index)
            Y = GetPlayerY(index) + 1
        Case DIR_LEFT

            If GetPlayerX(index) = 0 Then Exit Sub
            X = GetPlayerX(index) - 1
            Y = GetPlayerY(index)
        Case DIR_RIGHT

            If GetPlayerX(index) = Map(GetPlayerMap(index)).MaxX Then Exit Sub
            X = GetPlayerX(index) + 1
            Y = GetPlayerY(index)
    End Select
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data
    i = buffer.ReadLong
    Set buffer = Nothing
    
    If TempPlayer(index).EventMap.CurrentEvents > 0 Then
        For z = 1 To TempPlayer(index).EventMap.CurrentEvents
            ' Don't process events that are already processing
            If TempPlayer(index).EventProcessingCount > 0 Then
                If TempPlayer(index).EventProcessing(z).EventID = i Then Exit Sub
            End If
            
            If TempPlayer(index).EventMap.EventPages(z).EventID = i Then
                i = z
                BeginEventProcessing = True
                Exit For
            End If
        Next
    End If
    
    If BeginEventProcessing = True Then
        If Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).EventID).Pages(TempPlayer(index).EventMap.EventPages(i).PageID).CommandListCount > 0 Then
            ' Process this event, it is action button and everything checks out
            TempPlayer(index).EventProcessingCount = TempPlayer(index).EventProcessingCount + 1
            ReDim Preserve TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount)
            
            With TempPlayer(index).EventProcessing(TempPlayer(index).EventProcessingCount)
                .ActionTimer = timeGetTime
                .CurList = 1
                .CurSlot = 1
                .EventID = TempPlayer(index).EventMap.EventPages(i).EventID
                .PageID = TempPlayer(index).EventMap.EventPages(i).PageID
                .WaitingForResponse = 0
                ReDim .ListLeftOff(0 To Map(GetPlayerMap(index)).Events(TempPlayer(index).EventMap.EventPages(i).EventID).Pages(TempPlayer(index).EventMap.EventPages(i).PageID).CommandListCount)
            End With
        End If
        BeginEventProcessing = False
    End If
End Sub

Sub HandleRequestSwitchesAndVariables(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    SendSwitchesAndVariables (index)
End Sub

Sub HandleSwitchesAndVariables(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, i As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = buffer.ReadString
    Next
    
    SaveSwitches
    SaveVariables
    
    Set buffer = Nothing
    
    SendSwitchesAndVariables 0, True
End Sub

 ' Character Editor
Sub HandleRequestAllCharacters(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    If GetPlayerAccess(index) >= STAFF_ADMIN Then
        SendAllCharacters index
    End If
End Sub

Sub HandleRequestExtendedPlayerData(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, i As Long
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    SendExtendedPlayerData index, buffer.ReadString
    
    Set buffer = Nothing
End Sub

Sub HandleCharacterUpdate(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    Dim PlayerSize As Long, testSize As Long
    Dim PlayerData() As Byte
    Dim updatedPlayer As PlayerEditableRec
    
    PlayerSize = LenB(updatedPlayer)
    ReDim plaData(PlayerSize - 1)
    PlayerData = buffer.ReadBytes(PlayerSize)
    CopyMemory ByVal VarPtr(updatedPlayer), ByVal VarPtr(PlayerData(0)), PlayerSize
    Set buffer = Nothing
    
    ' Check if He is Online
    Dim tempSize As Long
    Dim i As Long, j As Long
    For i = 1 To MAX_PLAYERS
        For j = 1 To MAX_CHARS
            If Account(i).Login = "" Then GoTo use_offline_player
            If Trim(Account(i).Chars(j).Name) = Trim(updatedPlayer.Name) Then
                GoTo use_online_player
            End If
        Next
    Next
    
use_offline_player:
    ' Find associated Account Name
    Dim F As Long
    Dim s As String
    Dim charLogin() As String
    F = FreeFile
    
    Open App.path & "\data\accounts\charlist.txt" For Input As #F
        Do While Not EOF(F)
            Input #F, s
            charLogin = Split(s, ":")
            If charLogin(0) = Trim(updatedPlayer.Name) Then Exit Do
        Loop
    Close #F
    
    ' Load Character into temp variable - charLogin(0) -> Character Name | charLogin(1) -> Account/Login Name
    Dim tempAccount As AccountRec
    Dim filename As String
    
    filename = App.path & "\data\accounts\" & charLogin(1) & "\data.bin"
    
    F = FreeFile
    
    Open filename For Binary As #F
        Get #F, , tempAccount
    Close #F
    
    ' Get Character info, that we are requesting -> playerName
    Dim requestedClientPlayer As PlayerEditableRec
    For i = 1 To MAX_CHARS
        If Trim$(tempAccount.Chars(i).Name) = Trim(updatedPlayer.Name) Then
            Exit For
        End If
    Next
    
    tempAccount.Chars(i).Level = updatedPlayer.Level
    tempAccount.Chars(i).Exp = updatedPlayer.Exp
    tempAccount.Chars(i).Points = updatedPlayer.Points
    tempAccount.Chars(i).Sprite = updatedPlayer.Sprite
    tempAccount.Chars(i).Access = updatedPlayer.Access
    tempSize = LenB(tempAccount.Chars(i).Stat(1)) * UBound(tempAccount.Chars(i).Stat)
    CopyMemory ByVal VarPtr(tempAccount.Chars(i).Stat(1)), ByVal VarPtr(updatedPlayer.Stat(1)), tempSize
    tempSize = LenB(tempAccount.Chars(i).Vital(1)) * UBound(tempAccount.Chars(i).Vital)
    CopyMemory ByVal VarPtr(tempAccount.Chars(i).Vital(1)), ByVal VarPtr(updatedPlayer.Vital(1)), tempSize
    
    ' Save the account
    Call ChkDir(App.path & "\data\accounts\", Trim(tempAccount.Login))
    filename = App.path & "\data\accounts\" & Trim(tempAccount.Login) & "\data.bin"
    F = FreeFile
    
    Open filename For Binary As #F
        Put #F, , tempAccount
    Close #F
    Exit Sub
    
use_online_player:
    ' Copy over data
    Account(i).Chars(j).Level = updatedPlayer.Level
    Account(i).Chars(j).Exp = updatedPlayer.Exp
    Account(i).Chars(j).Points = updatedPlayer.Points
    Account(i).Chars(j).Sprite = updatedPlayer.Sprite
    Account(i).Chars(j).Access = updatedPlayer.Access
    tempSize = LenB(Account(i).Chars(j).Stat(1)) * UBound(Account(i).Chars(j).Stat)
    CopyMemory ByVal VarPtr(Account(i).Chars(j).Stat(1)), ByVal VarPtr(updatedPlayer.Stat(1)), tempSize
    tempSize = LenB(Account(i).Chars(j).Vital(1)) * UBound(Account(i).Chars(j).Vital)
    CopyMemory ByVal VarPtr(Account(i).Chars(j).Vital(1)), ByVal VarPtr(updatedPlayer.Vital(1)), tempSize
    Call SendPlayerData(i)
End Sub

' :::::::::::::::::::
' :: Search packet ::
' :::::::::::::::::::
Sub HandleTarget(ByVal index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim buffer As clsBuffer, Target As Long, TargetType As Long

    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()
    
    Target = buffer.ReadLong
    TargetType = buffer.ReadLong
    
    Set buffer = Nothing
    
    ' set player's target - no need to send, it's client side
    TempPlayer(index).Target = Target
    TempPlayer(index).TargetType = TargetType
End Sub
