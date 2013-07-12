Attribute VB_Name = "modHandleData"
Option Explicit

' ******************************************
' ** Parses and handles String packets    **
' ******************************************
Public Function GetAddress(FunAddr As Long) As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    GetAddress = FunAddr
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetAddress", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub InitMessages()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    HandleDataSub(SAlertMsg) = GetAddress(AddressOf HandleAlertMsg)
    HandleDataSub(SLoginOk) = GetAddress(AddressOf HandleLoginOk)
    HandleDataSub(SNewCharClasses) = GetAddress(AddressOf HandleNewCharClasses)
    HandleDataSub(SInGame) = GetAddress(AddressOf HandleInGame)
    HandleDataSub(SPlayerInv) = GetAddress(AddressOf HandlePlayerInv)
    HandleDataSub(SPlayerInvUpdate) = GetAddress(AddressOf HandlePlayerInvUpdate)
    HandleDataSub(SPlayerWornEq) = GetAddress(AddressOf HandlePlayerWornEq)
    HandleDataSub(SPlayerHP) = GetAddress(AddressOf HandlePlayerHP)
    HandleDataSub(SPlayerMP) = GetAddress(AddressOf HandlePlayerMP)
    HandleDataSub(SPlayerStats) = GetAddress(AddressOf HandlePlayerStats)
    HandleDataSub(SPlayerPoints) = GetAddress(AddressOf HandlePlayerPoints)
    HandleDataSub(SPlayerLevel) = GetAddress(AddressOf HandlePlayerLevel)
    HandleDataSub(SPlayerGuild) = GetAddress(AddressOf HandlePlayerGuild)
    HandleDataSub(SPlayerSprite) = GetAddress(AddressOf HandlePlayerSprite)
    HandleDataSub(SPlayerTitles) = GetAddress(AddressOf HandlePlayerTitles)
    HandleDataSub(SPlayerStatus) = GetAddress(AddressOf HandlePlayerStatus)
    HandleDataSub(SPlayerPK) = GetAddress(AddressOf HandlePlayerPK)
    HandleDataSub(SPlayerData) = GetAddress(AddressOf HandlePlayerData)
    HandleDataSub(SPlayerMove) = GetAddress(AddressOf HandlePlayerMove)
    HandleDataSub(SPlayerWarp) = GetAddress(AddressOf HandlePlayerWarp)
    HandleDataSub(SNpcMove) = GetAddress(AddressOf HandleNpcMove)
    HandleDataSub(SPlayerDir) = GetAddress(AddressOf HandlePlayerDir)
    HandleDataSub(SNpcDir) = GetAddress(AddressOf HandleNpcDir)
    HandleDataSub(SAttack) = GetAddress(AddressOf HandleAttack)
    HandleDataSub(SNpcAttack) = GetAddress(AddressOf HandleNpcAttack)
    HandleDataSub(SCheckForMap) = GetAddress(AddressOf HandleCheckForMap)
    HandleDataSub(SMapData) = GetAddress(AddressOf HandleMapData)
    HandleDataSub(SMapItemData) = GetAddress(AddressOf HandleMapItemData)
    HandleDataSub(SMapNpcData) = GetAddress(AddressOf HandleMapNpcData)
    HandleDataSub(SMapNpcTarget) = GetAddress(AddressOf HandleMapNpcTarget)
    HandleDataSub(SMapDone) = GetAddress(AddressOf HandleMapDone)
    HandleDataSub(SGlobalMsg) = GetAddress(AddressOf HandleGlobalMsg)
    HandleDataSub(SAdminMsg) = GetAddress(AddressOf HandleAdminMsg)
    HandleDataSub(SPlayerMsg) = GetAddress(AddressOf HandlePlayerMsg)
    HandleDataSub(SMapMsg) = GetAddress(AddressOf HandleMapMsg)
    HandleDataSub(SSpawnItem) = GetAddress(AddressOf HandleSpawnItem)
    HandleDataSub(SItemEditor) = GetAddress(AddressOf HandleItemEditor)
    HandleDataSub(SUpdateItem) = GetAddress(AddressOf HandleUpdateItem)
    HandleDataSub(SSpawnNpc) = GetAddress(AddressOf HandleSpawnNpc)
    HandleDataSub(SNpcDead) = GetAddress(AddressOf HandleNPCDead)
    HandleDataSub(SNpcEditor) = GetAddress(AddressOf HandleNPCEditor)
    HandleDataSub(SUpdateNpc) = GetAddress(AddressOf HandleUpdateNPC)
    HandleDataSub(SEditMap) = GetAddress(AddressOf HandleEditMap)
    HandleDataSub(SShopEditor) = GetAddress(AddressOf HandleShopEditor)
    HandleDataSub(SUpdateShop) = GetAddress(AddressOf HandleUpdateShop)
    HandleDataSub(SSpellEditor) = GetAddress(AddressOf HandleSpellEditor)
    HandleDataSub(SUpdateSpell) = GetAddress(AddressOf HandleUpdateSpell)
    HandleDataSub(SSpells) = GetAddress(AddressOf HandleSpells)
    HandleDataSub(SSpell) = GetAddress(AddressOf HandleSpell)
    HandleDataSub(SLeft) = GetAddress(AddressOf HandleLeft)
    HandleDataSub(SResourceCache) = GetAddress(AddressOf HandleResourceCache)
    HandleDataSub(SResourceEditor) = GetAddress(AddressOf HandleResourceEditor)
    HandleDataSub(SUpdateResource) = GetAddress(AddressOf HandleUpdateResource)
    HandleDataSub(SSendPing) = GetAddress(AddressOf HandleSendPing)
    HandleDataSub(SActionMsg) = GetAddress(AddressOf HandleActionMsg)
    HandleDataSub(SPlayerEXP) = GetAddress(AddressOf HandlePlayerExp)
    HandleDataSub(SBlood) = GetAddress(AddressOf HandleBlood)
    HandleDataSub(SAnimationEditor) = GetAddress(AddressOf HandleAnimationEditor)
    HandleDataSub(SUpdateAnimation) = GetAddress(AddressOf HandleUpdateAnimation)
    HandleDataSub(SAnimation) = GetAddress(AddressOf HandleAnimation)
    HandleDataSub(SMapNpcVitals) = GetAddress(AddressOf HandleMapNPCVitals)
    HandleDataSub(SSpellCooldown) = GetAddress(AddressOf HandleSpellCooldown)
    HandleDataSub(SClearSpellBuffer) = GetAddress(AddressOf HandleClearSpellBuffer)
    HandleDataSub(SSayMsg) = GetAddress(AddressOf HandleSayMsg)
    HandleDataSub(SOpenShop) = GetAddress(AddressOf HandleOpenShop)
    HandleDataSub(SResetShopAction) = GetAddress(AddressOf HandleResetShopAction)
    HandleDataSub(SStunned) = GetAddress(AddressOf HandleStunned)
    HandleDataSub(SMapWornEq) = GetAddress(AddressOf HandleMapWornEq)
    HandleDataSub(sbank) = GetAddress(AddressOf HandleBank)
    HandleDataSub(STrade) = GetAddress(AddressOf HandleOpenTrade)
    HandleDataSub(SCloseTrade) = GetAddress(AddressOf HandleCloseTrade)
    HandleDataSub(STradeUpdate) = GetAddress(AddressOf HandleTradeUpdate)
    HandleDataSub(STradeStatus) = GetAddress(AddressOf HandleTradeStatus)
    HandleDataSub(STradeRequest) = GetAddress(AddressOf HandleTradeRequest)
    HandleDataSub(SPartyInvite) = GetAddress(AddressOf HandlePartyInvite)
    HandleDataSub(SPartyUpdate) = GetAddress(AddressOf HandlePartyUpdate)
    HandleDataSub(SPartyVitals) = GetAddress(AddressOf HandlePartyVitals)
    HandleDataSub(SChatUpdate) = GetAddress(AddressOf HandleChatUpdate)
    HandleDataSub(STarget) = GetAddress(AddressOf HandlePlayerTarget)
    HandleDataSub(SHotbar) = GetAddress(AddressOf HandleHotbar)
    HandleDataSub(SGuildMembers) = GetAddress(AddressOf HandleGuildMembers)
    HandleDataSub(SGuildInvite) = GetAddress(AddressOf HandleGuildInvite)
    HandleDataSub(SMapReport) = GetAddress(AddressOf HandleMapReport)
    HandleDataSub(SNpcSpellBuffer) = GetAddress(AddressOf HandleNpcSpellBuffer)
    HandleDataSub(SCheckpoint) = GetAddress(AddressOf HandleCheckpoint)
    HandleDataSub(SUpdateLogs) = GetAddress(AddressOf HandleUpdateLogs)
    HandleDataSub(SFriendsList) = GetAddress(AddressOf HandleFriendsList)
    HandleDataSub(SFoesList) = GetAddress(AddressOf HandleFoesList)
    HandleDataSub(SHighIndex) = GetAddress(AddressOf HandlePlayer_HighIndex)
    HandleDataSub(SEntitySound) = GetAddress(AddressOf HandleEntitySound)
    HandleDataSub(SSendNews) = GetAddress(AddressOf HandleNews)
    HandleDataSub(SSound) = GetAddress(AddressOf HandleSound)
    HandleDataSub(SBanEditor) = GetAddress(AddressOf HandleBanEditor)
    HandleDataSub(SUpdateBan) = GetAddress(AddressOf HandleUpdateBan)
    HandleDataSub(STitleEditor) = GetAddress(AddressOf HandleTitleEditor)
    HandleDataSub(SUpdateTitle) = GetAddress(AddressOf HandleUpdateTitle)
    HandleDataSub(SMoralEditor) = GetAddress(AddressOf HandleMoralEditor)
    HandleDataSub(SUpdateMoral) = GetAddress(AddressOf HandleUpdateMoral)
    HandleDataSub(SClassEditor) = GetAddress(AddressOf HandleClassEditor)
    HandleDataSub(SUpdateClass) = GetAddress(AddressOf HandleUpdateClass)
    HandleDataSub(SCloseClient) = GetAddress(AddressOf HandleCloseClient)
    HandleDataSub(SLeaveGame) = GetAddress(AddressOf HandleLeaveGame)
    HandleDataSub(SEmoticonEditor) = GetAddress(AddressOf HandleEmoticonEditor)
    HandleDataSub(SUpdateEmoticon) = GetAddress(AddressOf HandleUpdateEmoticon)
    HandleDataSub(SCheckEmoticon) = GetAddress(AddressOf HandleCheckEmoticon)
    
    HandleDataSub(SSpawnEvent) = GetAddress(AddressOf HandleSpawnEventPage)
    HandleDataSub(SEventMove) = GetAddress(AddressOf HandleEventMove)
    HandleDataSub(SEventDir) = GetAddress(AddressOf HandleEventDir)
    HandleDataSub(SEventChat) = GetAddress(AddressOf HandleEventChat)
    
    HandleDataSub(SEventStart) = GetAddress(AddressOf HandleEventStart)
    HandleDataSub(SEventEnd) = GetAddress(AddressOf HandleEventEnd)
    
    HandleDataSub(SPlayBGM) = GetAddress(AddressOf HandlePlayBGM)
    HandleDataSub(SPlaySound) = GetAddress(AddressOf HandlePlaySound)
    HandleDataSub(SFadeoutBGM) = GetAddress(AddressOf HandleFadeoutBGM)
    HandleDataSub(SStopSound) = GetAddress(AddressOf HandleStopSound)
    HandleDataSub(SSwitchesAndVariables) = GetAddress(AddressOf HandleSwitchesAndVariables)
    
    HandleDataSub(SMapEventData) = GetAddress(AddressOf HandleMapEventData)
    
    HandleDataSub(SChatBubble) = GetAddress(AddressOf HandleChatBubble)
    
    HandleDataSub(SSpecialEffect) = GetAddress(AddressOf HandleSpecialEffect)
    'Character Editor
    HandleDataSub(SPlayersOnline) = GetAddress(AddressOf HandlePlayersOnline)
    HandleDataSub(SAllCharacters) = GetAddress(AddressOf HandleAllCharacters)
    HandleDataSub(SExtendedPlayerData) = GetAddress(AddressOf HandleExtendedPlayerData)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "InitMessages", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandleData(ByRef Data() As Byte)
    Dim buffer As clsBuffer
    Dim MsgType As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    MsgType = buffer.ReadLong

    If MsgType < 0 Then DestroyGame

    If MsgType >= SMSG_COUNT Then DestroyGame
    
    CallWindowProc HandleDataSub(MsgType), 1, buffer.ReadBytes(buffer.length), 0, 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandleAlertMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim Msg As String
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    Msg = buffer.ReadString
    Set buffer = Nothing
    
    frmLoad.Visible = False
    
    If InGame Then
        frmMain.Visible = False
    End If
    
    Call AlertMsg(Msg)
    
    If InGame Then
        LogoutGame
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleAlertMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandleLoginOk(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Now we can receive game data
    MyIndex = buffer.ReadLong
    
    ' Player high Index
    Player_HighIndex = buffer.ReadLong
    
    Set buffer = Nothing
    frmLoad.Visible = True
    Call SetStatus("Receiving game data...")
    
' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleLoginOk", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Sub HandleNewCharClasses(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    UpdateCharacterMenu
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleNewCharClasses", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandleInGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InGame = True
    Call GameInit
    Call GameLoop
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleInGame", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandlePlayerInv(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    n = 1

    For i = 1 To MAX_INV
        Call SetPlayerInvItemNum(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemValue(MyIndex, i, buffer.ReadLong)
        Call SetPlayerInvItemDur(MyIndex, i, buffer.ReadInteger)
        Call SetPlayerInvItemBind(MyIndex, i, buffer.ReadByte)
        n = n + 2
    Next
    
    ' Changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    TmpCurrencyItem = 0
    CurrencyMenu = 0 ' Clear

    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerInv", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandlePlayerInvUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    n = buffer.ReadByte
    
    SetGameFocus
    
    Call SetPlayerInvItemNum(MyIndex, n, buffer.ReadLong)
    Call SetPlayerInvItemValue(MyIndex, n, buffer.ReadLong)
    Call SetPlayerInvItemDur(MyIndex, n, buffer.ReadInteger)
    Call SetPlayerInvItemBind(MyIndex, n, buffer.ReadByte)
    Set buffer = Nothing
    
    ' Changes, clear drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    TmpCurrencyItem = 0
    CurrencyMenu = 0 ' Clear
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerInvUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandlePlayerWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Item number
    For i = 1 To Equipment.Equipment_Count - 1
        Call SetPlayerEquipment(MyIndex, buffer.ReadLong, i)
    Next
    
    ' Durability
    For i = 1 To Equipment.Equipment_Count - 1
        Call SetPlayerEquipmentDur(MyIndex, buffer.ReadInteger, i)
    Next
    Set buffer = Nothing
   
    If Index = MyIndex Then
        ' Changes to inventory, need to clear any drop menu
        frmMain.picCurrency.Visible = False
        frmMain.txtCurrency.text = vbNullString
        TmpCurrencyItem = 0
        CurrencyMenu = 0 ' Clear
        
        Call UpdateEquipPanel
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandleMapWornEq(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim PlayerNum As Long
    Dim i As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    PlayerNum = buffer.ReadLong
    
    ' Item number
    For i = 1 To Equipment.Equipment_Count - 1
        Call SetPlayerEquipment(MyIndex, buffer.ReadLong, i)
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleMapWornEq", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerHP(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Index = buffer.ReadLong
    player(Index).MaxVital(Vitals.HP) = buffer.ReadLong
    
    Call SetPlayerVital(Index, Vitals.HP, buffer.ReadLong)

    If Index = MyIndex Then
        If HPBarInit = False Then
            frmMain.imgHPBar.Width = ((GetPlayerVital(MyIndex, Vitals.HP) / HPBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / HPBar_Width)) * HPBar_Width
            HPBarInit = True
        End If
        OldHPBarWidth = frmMain.imgHPBar.Width
        NewHPBarWidth = ((GetPlayerVital(MyIndex, Vitals.HP) / HPBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.HP) / HPBar_Width)) * HPBar_Width
        
        If GetPlayerMaxVital(MyIndex, Vitals.HP) > 0 Then
            frmMain.lblHP.Caption = GetPlayerVital(MyIndex, Vitals.HP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.HP)
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerHP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerMP(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Index = buffer.ReadLong
    player(Index).MaxVital(Vitals.MP) = buffer.ReadLong
    
    Call SetPlayerVital(Index, Vitals.MP, buffer.ReadLong)

    If Index = MyIndex Then
        If MPBarInit = False Then
            frmMain.imgMPBar.Width = ((GetPlayerVital(MyIndex, Vitals.MP) / MPBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / MPBar_Width)) * MPBar_Width
            MPBarInit = True
        End If
        OldMPBarWidth = frmMain.imgMPBar.Width
        NewMPBarWidth = ((GetPlayerVital(MyIndex, Vitals.MP) / MPBar_Width) / (GetPlayerMaxVital(MyIndex, Vitals.MP) / MPBar_Width)) * MPBar_Width
        
        If GetPlayerMaxVital(MyIndex, Vitals.MP) > 0 Then
            frmMain.lblMP.Caption = GetPlayerVital(MyIndex, Vitals.MP) & "/" & GetPlayerMaxVital(MyIndex, Vitals.MP)
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerMP", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerStats(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Index = buffer.ReadLong

    For i = 1 To Stats.Stat_Count - 1
        SetPlayerStat Index, i, buffer.ReadInteger
        If Index = MyIndex Then
            If GetPlayerStat(Index, i) < 1 Then
                frmMain.lblCharStat(i).Caption = 0
            Else
                frmMain.lblCharStat(i).Caption = GetPlayerStat(Index, i)
            End If
        End If
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerStats", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerPoints(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Index = buffer.ReadLong
    SetPlayerPOINTS Index, buffer.ReadInteger
    
    If Index = MyIndex Then
        frmMain.lblPoints.Caption = GetPlayerPOINTS(Index)
        If GetPlayerPOINTS(Index) > 0 Then
            For i = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(Index, i) < MAX_STAT Then
                    frmMain.lblTrainStat(i).Visible = True
                Else
                    frmMain.lblTrainStat(i).Visible = False
                End If
            Next
        Else
            For i = 1 To Stats.Stat_Count - 1
                frmMain.lblTrainStat(i).Visible = False
            Next
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerPoints", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerLevel(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Index = buffer.ReadLong
    SetPlayerLevel Index, buffer.ReadInteger
    
    ' Set the character windows
    If Index = MyIndex Then
        frmMain.lblCharName = GetPlayerName(MyIndex)
        frmMain.lblCharLevel = "Lv: " & GetPlayerLevel(MyIndex)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerLevel", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
    
Private Sub HandlePlayerData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, X As Long
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    i = buffer.ReadLong
    
    player(i).Face = buffer.ReadInteger
    Call SetPlayerName(i, buffer.ReadString)
    Call SetPlayerGender(i, buffer.ReadByte)
    Call SetPlayerClass(i, buffer.ReadByte)
    Call SetPlayerLevel(i, buffer.ReadInteger)
    Call SetPlayerPOINTS(i, buffer.ReadInteger)
    Call SetPlayerSprite(i, buffer.ReadInteger)
    Call SetPlayerMap(i, buffer.ReadInteger)
    Call SetPlayerX(i, buffer.ReadByte)
    Call SetPlayerY(i, buffer.ReadByte)
    Call SetPlayerDir(i, buffer.ReadByte)
    Call SetPlayerAccess(i, buffer.ReadByte)
    Call SetPlayerPK(i, buffer.ReadByte)
    
    ' Guild
    player(i).Guild = buffer.ReadString
    player(i).GuildAcc = buffer.ReadByte

    ' Stats
    For X = 1 To Stats.Stat_Count - 1
        SetPlayerStat i, X, buffer.ReadInteger
    Next
    
    ' Amount of titles
    player(i).AmountOfTitles = buffer.ReadByte
    
    ' Player titles
    For X = 1 To player(i).AmountOfTitles
        player(i).title(X) = buffer.ReadByte
    Next
    
    ' Current title
    player(i).CurTitle = buffer.ReadByte
    
    ' Status
    player(i).Status = buffer.ReadString
    
    ' Check if the player is the client player
    If i = MyIndex Then
        ' Update the form guild boxes
        UpdateGuildPanel
        
        ' Reset directions
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
        
        ' Set the character windows
        frmMain.lblCharName = GetPlayerName(MyIndex)
        frmMain.lblCharLevel = "Lv: " & GetPlayerLevel(MyIndex)
        frmMain.lblPoints = GetPlayerPOINTS(MyIndex)
        
        ' Set stats on form
        For X = 1 To Stats.Stat_Count - 1
            If GetPlayerStat(MyIndex, X) < 1 Then
                frmMain.lblCharStat(X).Caption = 0
            Else
                frmMain.lblCharStat(X).Caption = GetPlayerStat(MyIndex, X)
            End If
        Next
        
        ' Set points on form
        frmMain.lblPoints.Caption = GetPlayerPOINTS(MyIndex)
        
        ' Set training label visiblity depending on points
        If GetPlayerPOINTS(MyIndex) > 0 Then
            For X = 1 To Stats.Stat_Count - 1
                If GetPlayerRawStat(MyIndex, X) < MAX_STAT Then
                    frmMain.lblTrainStat(X).Visible = True
                Else
                    frmMain.lblTrainStat(X).Visible = False
                End If
            Next
        Else
            For X = 1 To Stats.Stat_Count - 1
                frmMain.lblTrainStat(X).Visible = False
            Next
        End If
        
        ' Hide admin panel if visible and access is 0
        If frmAdmin.Visible = True Then
            If player(MyIndex).Access < STAFF_MODERATOR Then
                Unload frmAdmin
            End If
        End If
        UpdatePlayerTitles
    End If
    
    ' Make sure they aren't walking
    With tempPlayer(i)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerGuild(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Index = buffer.ReadLong
    
    ' Guild
    player(Index).Guild = buffer.ReadString
    player(Index).GuildAcc = buffer.ReadByte
    
    If Index = MyIndex Then
        ' Update the form guild boxes
        UpdateGuildPanel
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerGuild", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerSprite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Index = buffer.ReadLong
    player(Index).Sprite = buffer.ReadInteger
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerSprite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerTitles(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Index = buffer.ReadLong
    
    ' Amount of titles
    player(Index).AmountOfTitles = buffer.ReadByte
    
    ' Player titles
    For i = 1 To player(Index).AmountOfTitles
        player(Index).title(i) = buffer.ReadByte
    Next
    
    ' Current title
    player(Index).CurTitle = buffer.ReadByte
    
    Set buffer = Nothing
    
    If Index = MyIndex Then
        UpdatePlayerTitles
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerTitles", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerStatus(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Index = buffer.ReadLong
    
    player(Index).Status = buffer.ReadString
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerStatus", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerPK(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Index = buffer.ReadLong
    
    player(Index).PK = buffer.ReadByte
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerPK", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim X As Byte
    Dim Y As Byte
    Dim oldx As Long, oldy As Long
    Dim distanceX As Long, distanceY As Long
    Dim Dir As Byte
    Dim n As Byte
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    i = buffer.ReadLong
    X = buffer.ReadByte
    Y = buffer.ReadByte
    oldx = GetPlayerX(i)
    oldy = GetPlayerY(i)
    Dir = buffer.ReadByte
    n = buffer.ReadByte
    
    distanceX = (X - oldx)
    distanceY = (Y - oldy)
    
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, Dir)
    
    ' Let the client know we can move.
    IsWaitingForMove = False
    
    tempPlayer(i).Moving = n
    tempPlayer(i).xOffset = distanceX * PIC_X * -1
    tempPlayer(i).yOffset = distanceY * PIC_Y * -1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerWarp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim X As Byte
    Dim Y As Byte
    Dim Dir As Byte
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    i = buffer.ReadLong
    X = buffer.ReadByte
    Y = buffer.ReadByte
    Dir = buffer.ReadByte
    
    Call SetPlayerX(i, X)
    Call SetPlayerY(i, Y)
    Call SetPlayerDir(i, Dir)

    ' Make sure they aren't walking
    With tempPlayer(i)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With
    
    ' Let the client know we can move.
    IsWaitingForMove = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerWarp", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleNpcMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim MapNPCNum As Long
    Dim X As Byte
    Dim Y As Byte
    Dim Dir As Byte
    Dim Movement As Byte
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MapNPCNum = buffer.ReadLong
    X = buffer.ReadByte
    Y = buffer.ReadByte
    Dir = buffer.ReadByte
    Movement = buffer.ReadByte

    With MapNPC(MapNPCNum)
        .X = X
        .Y = Y
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = Movement

        Select Case .Dir
            Case DIR_UP
                .yOffset = PIC_Y
            Case DIR_DOWN
                .yOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleNpcMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Dir As Byte
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    i = buffer.ReadLong
    Dir = buffer.ReadByte
    
    Call SetPlayerDir(i, Dir)

    With tempPlayer(i)
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleNpcDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Dir As Byte
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    i = buffer.ReadLong
    Dir = buffer.ReadByte

    With MapNPC(i)
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleNpcDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    i = buffer.ReadLong
    
    ' Set player to attacking
    tempPlayer(i).Attacking = 1
    tempPlayer(i).AttackTimer = timeGetTime
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleNpcAttack(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    i = buffer.ReadLong
    
    ' Set player to attacking
    MapNPC(i).Attacking = 1
    MapNPC(i).AttackTimer = timeGetTime
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleNpcAttack", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleCheckForMap(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Integer
    Dim Y As Integer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Erase all players except self
    For i = 1 To Player_HighIndex
        If Not i = MyIndex Then
            Call SetPlayerMap(i, 0)
        End If
    Next

    ' Erase all temporary tile values
    Call ClearMapNpcs
    Call ClearMapItems
    Call ClearMap
    
    ' Clear the action messages
    For i = 1 To Action_HighIndex
        Call ClearActionMsg(i, False)
    Next
    
    Action_HighIndex = 0
    
    ' Clear the blood
    For i = 1 To Blood_HighIndex
        Call ClearBlood(i, False)
    Next
    
    Blood_HighIndex = 0
    
    ' Clear the chat bubbles
    For i = 1 To ChatBubble_HighIndex
        Call ClearChatBubble(i, False)
    Next
    
    ChatBubble_HighIndex = 0
    
    Map.CurrentEvents = 0
    ReDim Map.MapEvents(0)
    
    ' Get MapNum
    X = buffer.ReadInteger
    
    ' Get revision
    Y = buffer.ReadInteger
    
    Audio.CacheNewMapSounds
    InitAutotiles
    
    ' Request new map
    SendNeedMap

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        Unload frmEditor_Map
        ClearAttributeFrames

        If frmEditor_MapProperties.Visible Then
            Unload frmEditor_MapProperties
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleCheckForMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandleMapData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim X As Long
    Dim Y As Long
    Dim i As Long, Z As Long, w As Long
    Dim buffer As clsBuffer
    Dim MapNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteBytes Data()

    MapNum = buffer.ReadLong
    Map.name = buffer.ReadString
    Map.Music = buffer.ReadString
    Map.BGS = buffer.ReadString
    Map.Revision = buffer.ReadLong
    Map.Moral = buffer.ReadByte
    Map.Up = buffer.ReadLong
    Map.Down = buffer.ReadLong
    Map.Left = buffer.ReadLong
    Map.Right = buffer.ReadLong
    Map.BootMap = buffer.ReadLong
    Map.BootX = buffer.ReadByte
    Map.BootY = buffer.ReadByte
    
    Map.Weather = buffer.ReadLong
    Map.WeatherIntensity = buffer.ReadLong
    
    Map.Fog = buffer.ReadLong
    Map.FogSpeed = buffer.ReadLong
    Map.FogOpacity = buffer.ReadLong
    
    Map.Panorama = buffer.ReadLong
    
    Map.Red = buffer.ReadLong
    Map.Green = buffer.ReadLong
    Map.Blue = buffer.ReadLong
    Map.Alpha = buffer.ReadLong
    
    Map.MaxX = buffer.ReadByte
    Map.MaxY = buffer.ReadByte
    
    ReDim Map.Tile(0 To Map.MaxX, 0 To Map.MaxY)

    Map.Npc_HighIndex = buffer.ReadByte
    
    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            For i = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Layer(i).X = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Y = buffer.ReadLong
                Map.Tile(X, Y).Layer(i).Tileset = buffer.ReadLong
            Next
            
            For Z = 1 To MapLayer.Layer_Count - 1
                Map.Tile(X, Y).Autotile(Z) = buffer.ReadLong
            Next
            
            Map.Tile(X, Y).Type = buffer.ReadByte
            Map.Tile(X, Y).Data1 = buffer.ReadLong
            Map.Tile(X, Y).Data2 = buffer.ReadLong
            Map.Tile(X, Y).Data3 = buffer.ReadLong
            Map.Tile(X, Y).Data4 = buffer.ReadString
            Map.Tile(X, Y).DirBlock = buffer.ReadByte
        Next
    Next

    For X = 1 To MAX_MAP_NPCS
        Map.NPC(X) = buffer.ReadLong
        Map.NpcSpawnType(X) = buffer.ReadLong
        n = n + 1
    Next
    
    InitAutotiles
    
    Set buffer = Nothing

    ' Check if we get a map from someone else and if we were editing a map cancel it out
    If InMapEditor Then
        InMapEditor = False
        frmEditor_Map.Visible = False
        
        ClearAttributeFrames

        If frmEditor_MapProperties.Visible Then
            frmEditor_MapProperties.Visible = False
        End If
    End If
    
    Audio.CacheNewMapSounds
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleMapData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleMapItemData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
Dim i As Long
Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_ITEMS
        With MapItem(i)
            .playerName = buffer.ReadString
            .Num = buffer.ReadLong
            .Value = buffer.ReadLong
            .Durability = buffer.ReadInteger
            .X = buffer.ReadByte
            .Y = buffer.ReadByte
        End With
    Next
    
    ' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleMapItemData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleMapNpcData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long, X As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    For i = 1 To MAX_MAP_NPCS
        With MapNPC(i)
            .Num = buffer.ReadLong
            .X = buffer.ReadLong
            .Y = buffer.ReadLong
            .Dir = buffer.ReadLong
            
            For X = 1 To Vitals.Vital_Count - 1
                .Vital(X) = buffer.ReadLong
            Next
        End With
    Next
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleMapNpcTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim MapNPCNum As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MapNPCNum = buffer.ReadByte

    With MapNPC(MapNPCNum)
        .Target = buffer.ReadByte
        .TargetType = buffer.ReadByte
    End With
    
    ' Check for battle music
    Call CheckForBattleMusic(MapNPCNum)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleMapNpcData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleMapDone()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Load tilesets that we need
    LoadTilesets
    
    ' Play the music
    PlayMapMusic
    
    ' Re-position the map name
    Call UpdateDrawMapName

    For i = 1 To MAX_BYTE
        Call ClearAnimInstance(i)
    Next
    
    InitAutotiles
    
    CurrentWeather = Map.Weather
    CurrentWeatherIntensity = Map.WeatherIntensity
    CurrentFog = Map.Fog
    CurrentFogSpeed = Map.FogSpeed
    CurrentFogOpacity = Map.FogOpacity
    CurrentTintR = Map.Red
    CurrentTintG = Map.Green
    CurrentTintB = Map.Blue
    CurrentTintA = Map.Alpha

    GettingMap = False
    CanMoveNow = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleMapDone", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleGlobalMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadLong
    Call AddText(CheckMessage(Msg), Color)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleGlobalMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadByte
    
    ' Prevent ascii characters
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
    
    Call AddText(CheckMessage(Msg), Color)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayerMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleMapMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadByte
    Call AddText(CheckMessage(Msg), Color)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleMapMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleAdminMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Msg As String
    Dim Color As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Msg = buffer.ReadString
    Color = buffer.ReadByte
    Call AddText(CheckMessage(Msg), Color)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleAdminMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleSpawnItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong

    With MapItem(n)
        .playerName = buffer.ReadString
        .Num = buffer.ReadLong
        .Value = buffer.ReadLong
        .Durability = buffer.ReadInteger
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSpawnItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleItemEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Item
        Editor = EDITOR_ITEM
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ITEMS
            .lstIndex.AddItem i & ": " & Trim$(Item(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ItemEditorInit
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleItemEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleAnimationEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Animation
        Editor = EDITOR_ANIMATION
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_ANIMATIONS
            .lstIndex.AddItem i & ": " & Trim$(Animation(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        AnimationEditorInit
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleAnimationEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleUpdateItem(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim ItemSize As Long
    Dim ItemData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    
    ' Update the item
    ItemSize = LenB(Item(n))
    ReDim ItemData(ItemSize - 1)
    ItemData = buffer.ReadBytes(ItemSize)
    CopyMemory ByVal VarPtr(Item(n)), ByVal VarPtr(ItemData(0)), ItemSize
    Set buffer = Nothing

    ' Changes to inventory, need to clear any drop menu
    frmMain.picCurrency.Visible = False
    frmMain.txtCurrency.text = vbNullString
    TmpCurrencyItem = 0
    CurrencyMenu = 0 ' Clear
    
    If frmAdmin.Visible = True Then
        If n = frmAdmin.scrlAItem Then
            UpdateAdminScrollBar
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleUpdateItem", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleUpdateAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim AnimationSize As Long
    Dim AnimationData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    ' Update the Animation
    AnimationSize = LenB(Animation(n))
    ReDim AnimationData(AnimationSize - 1)
    AnimationData = buffer.ReadBytes(AnimationSize)
    CopyMemory ByVal VarPtr(Animation(n)), ByVal VarPtr(AnimationData(0)), AnimationSize
    Set buffer = Nothing
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "HandleUpdateAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleSpawnNpc(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong

    With MapNPC(n)
        .Num = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .Dir = buffer.ReadLong

        ' Client use only
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSpawnNpc", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleNPCDead(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    Call ClearMapNpc(n)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleNPCDead", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleNPCEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_NPC
        Editor = EDITOR_NPC
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_NPCS
            .lstIndex.AddItem i & ": " & Trim$(NPC(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        NpcEditorInit
    End With
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "HandleNPCEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleUpdateNPC(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim NpcSize As Long
    Dim NpcData() As Byte
    Dim i As Long
    Dim name As String
    Dim Letter As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    n = buffer.ReadLong
    
    NpcSize = LenB(NPC(n))
    ReDim NpcData(NpcSize - 1)
    NpcData = buffer.ReadBytes(NpcSize)
    CopyMemory ByVal VarPtr(NPC(n)), ByVal VarPtr(NpcData(0)), NpcSize
    
    Set buffer = Nothing
    
    If n = MAX_NPCS Then
        For i = 1 To MAX_NPCS
            If LenB(NPC(i)) = 0 Then
                NPC(i).name = vbNullString
                NPC(i).title = vbNullString
                NPC(i).AttackSay = vbNullString
            End If
        Next
    End If
    
    ' Update face
    Call DrawNPCFace(n)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleUpdateNPC", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleResourceEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Resource
        Editor = EDITOR_RESOURCE
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_RESOURCES
            .lstIndex.AddItem i & ": " & Trim$(Resource(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ResourceEditorInit
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleResourceEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleUpdateResource(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim ResourceNum As Long
    Dim buffer As clsBuffer
    Dim ResourceSize As Long
    Dim ResourceData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ResourceNum = buffer.ReadLong
    
    ResourceSize = LenB(Resource(ResourceNum))
    ReDim ResourceData(ResourceSize - 1)
    ResourceData = buffer.ReadBytes(ResourceSize)
    
    ClearResource ResourceNum
    
    CopyMemory ByVal VarPtr(Resource(ResourceNum)), ByVal VarPtr(ResourceData(0)), ResourceSize
    
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleUpdateResource", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleEditMap()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' We're in the map editor
    InMapEditor = True
    
    Call ToggleGUI(False)
    Call frmMain.ToggleChatLock(True)
    
    ' Show the form
    frmEditor_Map.Visible = True

    Call MapEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleEditMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleShopEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Shop
        Editor = EDITOR_SHOP
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SHOPS
            .lstIndex.AddItem i & ": " & Trim$(Shop(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ShopEditorInit
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleShopEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleUpdateShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim ShopSize As Long
    Dim ShopData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    n = buffer.ReadLong
    ShopSize = LenB(Shop(n))
    ReDim ShopData(ShopSize - 1)
    ShopData = buffer.ReadBytes(ShopSize)
    CopyMemory ByVal VarPtr(Shop(n)), ByVal VarPtr(ShopData(0)), ShopSize
    Set buffer = Nothing
    
    ' Update the shop combo box in the map editor if they are in it
    If InMapEditor Then
        MapEditorInitShop
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleUpdateShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleSpellEditor()
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With frmEditor_Spell
        Editor = EDITOR_SPELL
        .lstIndex.Clear

        ' Add the names
        For i = 1 To MAX_SPELLS
            .lstIndex.AddItem i & ": " & Trim$(Spell(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        SpellEditorInit
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSpellEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleUpdateSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim SpellNum As Long
    Dim buffer As clsBuffer
    Dim SpellSize As Long
    Dim SpellData() As Byte
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    SpellNum = buffer.ReadLong
    
    SpellSize = LenB(Spell(SpellNum))
    ReDim SpellData(SpellSize - 1)
    SpellData = buffer.ReadBytes(SpellSize)
    CopyMemory ByVal VarPtr(Spell(SpellNum)), ByVal VarPtr(SpellData(0)), SpellSize
    Set buffer = Nothing
    
    ' Refresh the spell
    Call UpdateSpells
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleUpdateSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandleSpells(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For i = 1 To MAX_PLAYER_SPELLS
        PlayerSpells(i) = buffer.ReadLong
    Next
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandleSpell(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim buffer As clsBuffer
    Dim SpellSlot As Byte, SpellNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    SpellSlot = buffer.ReadByte
    SpellNum = buffer.ReadLong
    
    PlayerSpells(SpellSlot) = SpellNum
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSpells", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleLeft(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Call ClearPlayer(buffer.ReadLong)
    Set buffer = Nothing
    
' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleLeft", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleResourceCache(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' If in map editor, we cache it ourselves
    If InMapEditor Then Exit Sub
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    Resource_Index = buffer.ReadLong
    Resources_Init = False

    If Resource_Index > 0 Then
        ReDim Preserve MapResource(0 To Resource_Index)

        For i = 0 To Resource_Index
            MapResource(i).ResourceState = buffer.ReadByte
            MapResource(i).X = buffer.ReadInteger
            MapResource(i).Y = buffer.ReadInteger
        Next

        Resources_Init = True
    Else
        ReDim MapResource(0 To 1)
    End If

    Set buffer = Nothing
    Exit Sub
        
' Error handler
ErrorHandler:
    HandleError "HandleResourceCache", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleSendPing(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    PingEnd = timeGetTime
    Ping = PingEnd - PingStart
    Call SetPing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSendPing", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleActionMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long, Message As String, Color As Long, TmpType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Message = buffer.ReadString
    Color = buffer.ReadLong
    TmpType = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong

    Set buffer = Nothing
    
    CreateActionMsg Message, Color, TmpType, X, Y
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleActionMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerExp(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim TNL As Long
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Index = buffer.ReadLong
    
    ' Make sure their not the max level
    If Not GetPlayerLevel(Index) = MAX_LEVEL Then
        Call SetPlayerExp(Index, buffer.ReadLong)
    ElseIf GetPlayerExp(Index) > 0 And GetPlayerLevel(Index) = MAX_LEVEL Then
        Call SetPlayerExp(Index, 0)
    End If
    
    TNL = buffer.ReadLong
    
    If Index = MyIndex Then
        ' Exp bar
        If Not GetPlayerLevel(MyIndex) = MAX_LEVEL Then
            If EXPBarInit = False Then
                frmMain.imgEXPBar.Width = ((GetPlayerExp(MyIndex) / EXPBar_Width) / (TNL / EXPBar_Width)) * EXPBar_Width
                EXPBarInit = True
            End If
            
            OldEXPBarWidth = frmMain.imgEXPBar.Width
            NewEXPBarWidth = ((GetPlayerExp(MyIndex) / EXPBar_Width) / (TNL / EXPBar_Width)) * EXPBar_Width
            frmMain.lblEXP.Visible = True
            frmMain.lblEXP.Caption = GetPlayerExp(Index) & "/" & TNL
        Else
            frmMain.imgEXPBar.Width = EXPBar_Width
            frmMain.lblEXP.Visible = False
            frmMain.lblEXP.Caption = ""
        End If
    End If
End Sub

Private Sub HandleBlood(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Set buffer = Nothing
    
    Call CreateBlood(X, Y)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleBlood", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleAnimation(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    AnimationIndex = AnimationIndex + 1
    If AnimationIndex >= MAX_ANIMATIONS Then AnimationIndex = 1
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    With AnimInstance(AnimationIndex)
        .Animation = buffer.ReadLong
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .LockType = buffer.ReadByte
        .lockIndex = buffer.ReadLong
        .Used(0) = True
        .Used(1) = True
        
        If Not .Animation = 1 Then
            ' Play the sound if we've got one
            PlaySoundEntity .X, .Y, SoundEntity.seAnimation, .Animation
        End If
    End With
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleAnimation", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleMapNPCVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim MapNPCNum As Byte, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MapNPCNum = buffer.ReadByte
    For i = 1 To Vitals.Vital_Count - 1
        MapNPC(MapNPCNum).Vital(i) = buffer.ReadLong
    Next
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleMapNPCVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleSpellCooldown(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Slot As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Slot = buffer.ReadByte
    SpellCD(Slot) = buffer.ReadLong
    Set buffer = Nothing
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "HandleSpellCooldown", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleClearSpellBuffer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    SpellBuffer = 0
    SpellBufferTimer = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleClearSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleSayMsg(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Access As Long
    Dim name As String
    Dim Message As String
    Dim Color As Long
    Dim Header As String
    Dim PK As Long
    Dim SayColor As Long
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    name = buffer.ReadString
    Access = buffer.ReadLong
    PK = buffer.ReadLong
    Message = CheckMessage(buffer.ReadString)
    Header = buffer.ReadString
    SayColor = buffer.ReadLong
    
    ' Prevent ascii characters
    For i = 1 To Len(Message)
        ' limit the ASCII
        If AscW(Mid$(Message, i, 1)) < 32 Or AscW(Mid$(Message, i, 1)) > 126 Then
            ' limit the extended ASCII
            If AscW(Mid$(Message, i, 1)) < 128 Or AscW(Mid$(Message, i, 1)) > 168 Then
                ' limit the extended ASCII
                If AscW(Mid$(Message, i, 1)) < 224 Or AscW(Mid$(Message, i, 1)) > 253 Then
                    Mid$(Message, i, 1) = ""
                End If
            End If
        End If
    Next
    
    ' Check access level
    If PK = NO Then
        Select Case Access
            Case 0
                Color = QBColor(15)
            Case 1
                Color = QBColor(3)
            Case 2
                Color = QBColor(2)
            Case 3
                Color = QBColor(BrightBlue)
            Case 4
                Color = QBColor(Yellow)
            Case 5
                Color = RGB(255, 165, 0)
        End Select
    Else
        Color = QBColor(BrightRed)
    End If
    
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text)
    frmMain.txtChat.SelColor = Color
    frmMain.txtChat.SelText = vbNewLine & Header & name & ": "
    
    If SayColor < Orange Then
        frmMain.txtChat.SelColor = QBColor(SayColor)
    Else
        frmMain.txtChat.SelColor = SayColor
    End If
    
    frmMain.txtChat.SelText = Message
    frmMain.txtChat.SelStart = Len(frmMain.txtChat.text) - 1
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSayMsg", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleOpenShop(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim ShopNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ShopNum = buffer.ReadLong
    
    Set buffer = Nothing
    
    OpenShop ShopNum
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleOpenShop", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleResetShopAction(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ShopAction = 0
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleResetShopAction", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleStunned(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    StunDuration = buffer.ReadLong
    
    Set buffer = Nothing
    
' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleStunned", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleBank(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For i = 1 To MAX_BANK
        bank.Item(i).Num = buffer.ReadLong
        bank.Item(i).Value = buffer.ReadLong
    Next
    
    InBank = True
    frmMain.picBank.Visible = True
    frmMain.picBank.ZOrder (0)
    
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleBank", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleOpenTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    InTrade = buffer.ReadLong
    frmMain.picTrade.Visible = True
    
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleOpenTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleCloseTrade(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InTrade = 0
    frmMain.picTrade.Visible = False
    frmMain.lblTradeStatus.Caption = vbNullString
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleCloseTrade", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleTradeUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim DataType As Byte
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    DataType = buffer.ReadByte
    
    If DataType = 0 Then ' Ours
        For i = 1 To MAX_INV
            TradeYourOffer(i).Num = buffer.ReadLong
            TradeYourOffer(i).Value = buffer.ReadLong
        Next
        
        frmMain.lblYourWorth.Caption = buffer.ReadLong & " " & Trim$(Item(1).name)
    ElseIf DataType = 1 Then ' Theirs
        For i = 1 To MAX_INV
            TradeTheirOffer(i).Num = buffer.ReadLong
            TradeTheirOffer(i).Value = buffer.ReadLong
        Next
        
        frmMain.lblTheirWorth.Caption = buffer.ReadLong & " " & Trim$(Item(1).name)
    End If
    
    Set buffer = Nothing
    
' Error handler
    Exit Sub
ErrorHandler:
    HandleError "HandleTradeUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Sub HandleTradeStatus(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim TradeStatus As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    TradeStatus = buffer.ReadByte
    
    Set buffer = Nothing
    
    Select Case TradeStatus
        Case 0 ' Clear
            frmMain.lblTradeStatus.Caption = vbNullString
        Case 1 ' They've accepted
            frmMain.lblTradeStatus.Caption = "Other player has accepted."
        Case 2 ' You've accepted
            frmMain.lblTradeStatus.Caption = "Waiting for other player to accept."
    End Select
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleTradeStatus", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleMapReport(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    With frmMapReport
        .lstMaps.Clear
        
        For i = 1 To MAX_MAPS
            .lstMaps.AddItem i & ": " & buffer.ReadString
        Next
        
        .lstMaps.ListIndex = 0
        .Show
    End With
   
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleMapReport", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleChatUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    Dim NpcNum As Long
    Dim mT As String
    Dim o1 As String
    Dim o2 As String
    Dim o3 As String
    Dim o4 As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    NpcNum = buffer.ReadLong
    mT = buffer.ReadString
    o1 = buffer.ReadString
    o2 = buffer.ReadString
    o3 = buffer.ReadString
    o4 = buffer.ReadString
    
    ' If NPCNum is out of range then exit the chat system
    If NpcNum < 1 Or NpcNum > MAX_NPCS Then
        InChat = False
        frmMain.picEventChat.Visible = False
        Exit Sub
    End If
    
    ' Setup the main window
    frmMain.lblEventChat.Caption = mT
    frmMain.lblChoices(1).Caption = o1
    frmMain.lblChoices(2).Caption = o2
    frmMain.lblChoices(3).Caption = o3
    frmMain.lblChoices(4).Caption = o4
    
    ' Render the face
    DrawNPCFace NpcNum
    
    ' Make it visible
    frmMain.picEventChat.Visible = True
    frmMain.picEventChat.ZOrder (0)
    
    ' Can't move
    InChat = True
    Set buffer = Nothing
    Exit Sub

' Error handler
ErrorHandler:
    HandleError "HandleChatUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayerTarget(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MyTarget = buffer.ReadByte
    MyTargetType = buffer.ReadByte
    Set buffer = Nothing
End Sub

Private Sub HandleHotbar(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
        
    For i = 1 To MAX_HOTBAR
        Hotbar(i).Slot = buffer.ReadByte
        Hotbar(i).SType = buffer.ReadByte
    Next
    
    frmMain.picHotbar.Refresh
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleHotbar", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleTradeRequest(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    name = buffer.ReadString
    
    Dialogue "Trade Request", name & " has requested to trade with you. Would you like to accept?", DIALOGUE_TYPE_TRADE, True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleTradeRequest", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePartyInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim name As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    name = buffer.ReadString
    
    Dialogue "Party Invitation", name & " has invited you to a party. Would you like to join?", DIALOGUE_TYPE_PARTY, True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePartyInvite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleGuildInvite(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim name As String, Guild As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    name = buffer.ReadString
    Guild = buffer.ReadString
    
    Dialogue "Guild Invitation", name & " has invited you to join the guild " & Guild & "!", DIALOGUE_TYPE_GUILD, True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleGuildInvite", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePartyUpdate(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, i As Long, InParty As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()

    InParty = buffer.ReadByte
    
    ' Exit out if we're not in a party
    If InParty = 0 Then
        Call ZeroMemory(ByVal VarPtr(Party), LenB(Party))
        ' Reset the labels
        For i = 1 To MAX_PARTY_MEMBERS
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
            Party.Num = 0
        Next
        Exit Sub
    End If
    
    ' Carry on otherwise
    For i = 1 To MAX_PARTY_MEMBERS
        Party.Member(i) = buffer.ReadLong
        If Party.Member(i) > 0 Then
            frmMain.lblPartyMember(i).Caption = Trim$(GetPlayerName(Party.Member(i)))
            frmMain.imgPartyHealth(i).Visible = True
            frmMain.imgPartySpirit(i).Visible = True
        Else
            frmMain.lblPartyMember(i).Caption = vbNullString
            frmMain.imgPartyHealth(i).Visible = False
            frmMain.imgPartySpirit(i).Visible = False
        End If
    Next
    Party.MemberCount = buffer.ReadLong
    Party.Num = buffer.ReadLong
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePartyUpdate", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePartyVitals(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim PlayerNum As Long, PartyIndex As Long
    Dim buffer As clsBuffer, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Which player?
    PlayerNum = buffer.ReadLong
    
    ' Set vitals
    For i = 1 To Vitals.Vital_Count - 1
        player(PlayerNum).MaxVital(i) = buffer.ReadLong
        player(PlayerNum).Vital(i) = buffer.ReadLong
    Next
    
    ' Find the party number
    For i = 1 To MAX_PARTY_MEMBERS
        If Party.Member(i) = PlayerNum Then
            PartyIndex = i
        End If
    Next
    
    ' Exit out if wrong data
    If PartyIndex <= 0 Or PartyIndex > MAX_PARTY_MEMBERS Then Exit Sub
    
    ' Hp bar
    frmMain.imgPartyHealth(PartyIndex).Width = ((GetPlayerVital(PlayerNum, Vitals.HP) / Party_HPWidth) / (GetPlayerMaxVital(PlayerNum, Vitals.HP) / Party_HPWidth)) * Party_HPWidth
    
    ' Mana bar
    frmMain.imgPartySpirit(PartyIndex).Width = ((GetPlayerVital(PlayerNum, Vitals.MP) / Party_MPWidth) / (GetPlayerMaxVital(PlayerNum, Vitals.MP) / Party_MPWidth)) * Party_MPWidth
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePartyVitals", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleSpawnEventPage(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim id As Long, i As Long, Z As Long, X As Long, Y As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    id = buffer.ReadLong
    
    If id > Map.CurrentEvents Then
        Map.CurrentEvents = id
        ReDim Preserve Map.MapEvents(Map.CurrentEvents)
    End If

    With Map.MapEvents(id)
        .name = buffer.ReadString
        .Dir = buffer.ReadLong
        .ShowDir = .Dir
        .GraphicNum = buffer.ReadLong
        .GraphicType = buffer.ReadLong
        .GraphicX = buffer.ReadLong
        .GraphicX2 = buffer.ReadLong
        .GraphicY = buffer.ReadLong
        .GraphicY2 = buffer.ReadLong
        .MovementSpeed = buffer.ReadLong
        .Moving = 0
        .X = buffer.ReadLong
        .Y = buffer.ReadLong
        .xOffset = 0
        .yOffset = 0
        .Position = buffer.ReadLong
        .Visible = buffer.ReadLong
        .WalkAnim = buffer.ReadLong
        .DirFix = buffer.ReadLong
        .WalkThrough = buffer.ReadLong
        .ShowName = buffer.ReadLong
    End With
    
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSpawnEventPage", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleEventMove(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim id As Long
    Dim X As Long
    Dim Y As Long
    Dim Dir As Long, ShowDir As Long
    Dim Movement As Long, MovementSpeed As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    id = buffer.ReadLong
    X = buffer.ReadLong
    Y = buffer.ReadLong
    Dir = buffer.ReadLong
    ShowDir = buffer.ReadLong
    MovementSpeed = buffer.ReadLong
    
    If id > Map.CurrentEvents Then Exit Sub

    With Map.MapEvents(id)
        .X = X
        .Y = Y
        .Dir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 1
        .ShowDir = ShowDir
        .MovementSpeed = MovementSpeed
        
        Select Case Dir
            Case DIR_UP
                .yOffset = PIC_Y
            Case DIR_DOWN
                .yOffset = PIC_Y * -1
            Case DIR_LEFT
                .xOffset = PIC_X
            Case DIR_RIGHT
                .xOffset = PIC_X * -1
        End Select

    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleEventMove", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleEventDir(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Dir As Byte
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    i = buffer.ReadLong
    Dir = buffer.ReadLong
    
    If i > Map.CurrentEvents Then Exit Sub

    With Map.MapEvents(i)
        .Dir = Dir
        .ShowDir = Dir
        .xOffset = 0
        .yOffset = 0
        .Moving = 0
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleEventDir", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleEventChat(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    Dim Dir As Byte
    Dim buffer As clsBuffer
    Dim Choices As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    frmMain.picEventChat.Visible = True
    frmMain.picEventChat.ZOrder (0)
    EventReplyID = buffer.ReadLong
    EventReplyPage = buffer.ReadLong
    frmMain.lblEventChat.Caption = buffer.ReadString
    frmMain.picEventChat.Visible = True
    frmMain.lblEventChat.Visible = True
    Choices = buffer.ReadLong
    
    InEvent = True
    
    For i = 1 To 4
        frmMain.lblChoices(i).Visible = False
    Next
    
    frmMain.lblEventChatContinue.Visible = False
    
    If Choices = 0 Then
        frmMain.lblEventChatContinue.Visible = True
    Else
        For i = 1 To Choices
            frmMain.lblChoices(i).Visible = True
            frmMain.lblChoices(i).Caption = buffer.ReadString
        Next
    End If
    
    AnotherChat = buffer.ReadLong
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleEventChat", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleEventStart(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InEvent = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleEventStart", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleEventEnd(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InEvent = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleEventEnd", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayBGM(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    str = buffer.ReadString
    
    Audio.StopMusic
    Audio.PlayMusic str
    
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlaySound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    str = buffer.ReadString

    Audio.PlaySound str, -1, -1
    
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlaySound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleFadeoutBGM(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim str As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Audio.StopMusic
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleStopSound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim str As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Audio.StopMapSounds
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleFadeoutBGM", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleSwitchesAndVariables(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim str As String, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    For i = 1 To MAX_SWITCHES
        Switches(i) = buffer.ReadString
    Next
    
    For i = 1 To MAX_VARIABLES
        Variables(i) = buffer.ReadString
    Next
    
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSwitchesAndVariables", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleMapEventData(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim str As String, i As Long, X As Long, Y As Long, Z As Long, w As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ' Event data
    Map.EventCount = buffer.ReadLong
        
    If Map.EventCount > 0 Then
        ReDim Map.events(0 To Map.EventCount)
        For i = 1 To Map.EventCount
            With Map.events(i)
                .name = buffer.ReadString
                .Global = buffer.ReadLong
                .X = buffer.ReadLong
                .Y = buffer.ReadLong
                .pageCount = buffer.ReadLong
            End With
            If Map.events(i).pageCount > 0 Then
                ReDim Map.events(i).Pages(0 To Map.events(i).pageCount)
                For X = 1 To Map.events(i).pageCount
                    With Map.events(i).Pages(X)
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
                            ReDim Map.events(i).Pages(X).MoveRoute(0 To .MoveRouteCount)
                            For Y = 1 To .MoveRouteCount
                                .MoveRoute(Y).Index = buffer.ReadLong
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
                        
                    If Map.events(i).Pages(X).CommandListCount > 0 Then
                        ReDim Map.events(i).Pages(X).CommandList(0 To Map.events(i).Pages(X).CommandListCount)
                        For Y = 1 To Map.events(i).Pages(X).CommandListCount
                            Map.events(i).Pages(X).CommandList(Y).CommandCount = buffer.ReadLong
                            Map.events(i).Pages(X).CommandList(Y).ParentList = buffer.ReadLong
                            If Map.events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                ReDim Map.events(i).Pages(X).CommandList(Y).Commands(1 To Map.events(i).Pages(X).CommandList(Y).CommandCount)
                                For Z = 1 To Map.events(i).Pages(X).CommandList(Y).CommandCount
                                    With Map.events(i).Pages(X).CommandList(Y).Commands(Z)
                                        .Index = buffer.ReadLong
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
                                                .MoveRoute(w).Index = buffer.ReadLong
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
    
    
    ' End event data
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleMapEventData", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleChatBubble(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, TargetType As Long, Target As Long, Message As String, colour As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Target = buffer.ReadLong
    TargetType = buffer.ReadLong
    Message = buffer.ReadString
    colour = buffer.ReadLong
    
    CreateChatBubble Target, TargetType, Message, colour
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleChatBubble", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
'Character Editor
Private Sub HandlePlayersOnline(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, playersOnline As String, length As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()
    
    playersOnline = buffer.ReadString
    
    length = UBound(charList)
    For i = 0 To length
        If InStr(playersOnline, charList(i, 0)) Then
            charList(i, 1) = "Online"
        End If
    Next
    frmCharEditor.ResetCharList
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSpecialEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
'Character Editor
Private Sub HandleAllCharacters(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, temp As String, allCharacters() As String, length As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    temp = buffer.ReadString
    allCharacters = Split(temp, ",")
    
    length = UBound(allCharacters)
    ReDim charList(0 To length, 1)
    For i = 0 To length
            charList(i, 0) = allCharacters(i)
            charList(i, 1) = "Offline"
    Next
    frmCharEditor.Visible = True
    SendRequestPlayersOnline
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSpecialEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
'Character Editor
Private Sub HandleExtendedPlayerData(ByVal Index As Long, ByRef data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes data()

    Dim PlayerSize As Long, testSize As Long
    Dim playerData() As Byte
    
    PlayerSize = LenB(requestedPlayer)
    ReDim plaData(PlayerSize - 1)
    playerData = buffer.ReadBytes(PlayerSize)
    CopyMemory ByVal VarPtr(requestedPlayer), ByVal VarPtr(playerData(0)), PlayerSize
    frmCharEditor.fetchPlayerData
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSpecialEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
Private Sub HandleSpecialEffect(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer, effectType As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    effectType = buffer.ReadLong
    
    Select Case effectType
        Case EFFECT_TYPE_FADEIN
            FadeType = 1
            FadeAmount = 0
        Case EFFECT_TYPE_FADEOUT
            FadeType = 0
            FadeAmount = 255
        Case EFFECT_TYPE_FLASH
            FlashTimer = timeGetTime + 150
        Case EFFECT_TYPE_FOG
            CurrentFog = buffer.ReadLong
            CurrentFogSpeed = buffer.ReadLong
            CurrentFogOpacity = buffer.ReadLong
        Case EFFECT_TYPE_WEATHER
            CurrentWeather = buffer.ReadLong
            CurrentWeatherIntensity = buffer.ReadLong
        Case EFFECT_TYPE_TINT
            CurrentTintR = buffer.ReadLong
            CurrentTintG = buffer.ReadLong
            CurrentTintB = buffer.ReadLong
            CurrentTintA = buffer.ReadLong
    End Select
    
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSpecialEffect", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleGuildMembers(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim MemberName() As String
    Dim i As Long
    Dim PlayerCount As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    PlayerCount = buffer.ReadLong
    frmMain.lstGuild.Clear
    
    If PlayerCount > 0 Then
        ReDim MemberName(1 To PlayerCount)
    
        For i = 1 To PlayerCount
            MemberName(i) = buffer.ReadString
        Next
        
        For i = 1 To UBound(MemberName)
            frmMain.lstGuild.AddItem MemberName(i)
        Next
    End If
    
    If frmMain.lstGuild.ListCount > 0 Then
        frmMain.lstGuild.Enabled = True
    Else
        frmMain.lstGuild.Enabled = False
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleGuildMembers", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleNpcSpellBuffer(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim MapNPCNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    MapNPCNum = buffer.ReadLong
    MapNPC(MapNPCNum).SpellBuffer = buffer.ReadLong
    MapNPC(MapNPCNum).SpellBufferTimer = timeGetTime
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleNpcSpellBuffer", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleCheckpoint(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    player(MyIndex).name = buffer.ReadLong
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleCheckpoint", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub HandleUpdateLogs(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim LogSize As Long
    Dim LogData() As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    LogSize = LenB(Log)
    ReDim LogData(LogSize - 1)
    LogData = buffer.ReadBytes(LogSize)
    CopyMemory ByVal VarPtr(Log), ByVal VarPtr(LogData(0)), LogSize
    Set buffer = Nothing
    
    ' Update the Log
    Call AddLog(Trim$(Log.Msg), Trim$(Log.file))
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleUpdateLogs", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandleFriendsList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim FriendsName As String
    Dim AmountOfFriends As Byte
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    AmountOfFriends = buffer.ReadByte
    
    ' Clear lstbox so it can be updated correctly.
    frmMain.lstFriends.Clear
    
    ' Prevents an error and clears your friends list when you have no friends
    If AmountOfFriends = 0 Then Exit Sub
   
    ' Adds Friends Name to the List
    For i = 1 To AmountOfFriends
        FriendsName = buffer.ReadString
        frmMain.lstFriends.AddItem buffer.ReadString
        player(MyIndex).Friends(i).name = FriendsName
    Next
    
    If frmMain.lstFriends.ListCount > 0 Then
        frmMain.lstFriends.Enabled = True
    Else
        frmMain.lstFriends.Enabled = False
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleFriendsList", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub HandleFoesList(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim FoesName As String
    Dim AmountOfFoes As Byte
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    AmountOfFoes = buffer.ReadByte
    
    ' Clear lstbox so it can be updated correctly.
    frmMain.lstFoes.Clear
    
    ' Prevents an error and clears your Foes list when you have no Foes
    If AmountOfFoes = 0 Then Exit Sub
   
    ' Adds Foes Name to the List
    For i = 1 To AmountOfFoes
        FoesName = buffer.ReadString
        frmMain.lstFoes.AddItem buffer.ReadString
        player(MyIndex).Foes(i).name = FoesName
    Next
    
    If frmMain.lstFoes.ListCount > 0 Then
        frmMain.lstFoes.Enabled = True
    Else
        frmMain.lstFoes.Enabled = True
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleFoesList", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandlePlayer_HighIndex(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Player_HighIndex = buffer.ReadLong
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandlePlayer_HighIndex", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleEntitySound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim X As Long, Y As Long, EntityType As Long, EntityNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    X = buffer.ReadLong
    Y = buffer.ReadLong
    EntityType = buffer.ReadLong
    EntityNum = buffer.ReadLong

    PlaySoundEntity X, Y, EntityType, EntityNum
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleEntitySound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleNews(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    frmMenu.lblNews.Caption = buffer.ReadString
    Set buffer = Nothing
    StopTimer = False
    
    frmMenu.lblServerStatus.Caption = "Online"
    frmMenu.lblServerStatus.ForeColor = vbGreen
    frmMenu.lblServerStatus.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleNews", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleSound(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim Sound As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    Sound = buffer.ReadString
    Audio.PlaySound Sound
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleSound", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleBanEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    With frmEditor_Ban
        Editor = EDITOR_BAN
        .lstIndex.Clear
        .ZOrder (0)
        
        ' Add the names
        For i = 1 To MAX_BANS
            .lstIndex.AddItem i & ": " & Trim$(Ban(i).playerName)
        Next

        .Show
        .lstIndex.ListIndex = 0
        BanEditorInit
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleBanEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleUpdateBan(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim BanSize As Long
    Dim BanData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    
    ' Update the Ban
    BanSize = LenB(Ban(n))
    ReDim BanData(BanSize - 1)
    BanData = buffer.ReadBytes(BanSize)
    CopyMemory ByVal VarPtr(Ban(n)), ByVal VarPtr(BanData(0)), BanSize
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleUpdateBan", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleUpdateTitle(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim TitleSize As Long
    Dim TitleData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    
    ' Update the Title
    TitleSize = LenB(title(n))
    ReDim TitleData(TitleSize - 1)
    TitleData = buffer.ReadBytes(TitleSize)
    CopyMemory ByVal VarPtr(title(n)), ByVal VarPtr(TitleData(0)), TitleSize
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleUpdateTitle", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleTitleEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    With frmEditor_Title
        Editor = EDITOR_TITLE
        .lstIndex.Clear
        .ZOrder (0)
        
        ' Add the names
        For i = 1 To MAX_TITLES
            .lstIndex.AddItem i & ": " & Trim$(title(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        TitleEditorInit
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleTitleEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleUpdateMoral(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim MoralSize As Long
    Dim MoralData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    
    ' Update the Moral
    MoralSize = LenB(Moral(n))
    ReDim MoralData(MoralSize - 1)
    MoralData = buffer.ReadBytes(MoralSize)
    CopyMemory ByVal VarPtr(Moral(n)), ByVal VarPtr(MoralData(0)), MoralSize
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleUpdateMoral", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleMoralEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    With frmEditor_Moral
        Editor = EDITOR_MORAL
        .lstIndex.Clear
        .ZOrder (0)
        
        ' Add the names
        For i = 1 To MAX_MORALS
            .lstIndex.AddItem i & ": " & Trim$(Moral(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        MoralEditorInit
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleMoralEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleClassEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    With frmEditor_Class
        Editor = EDITOR_CLASS
        .lstIndex.Clear
        .ZOrder (0)
        
        ' Add the names
        For i = 1 To MAX_CLASSES
            .lstIndex.AddItem i & ": " & Trim$(Class(i).name)
        Next

        .Show
        .lstIndex.ListIndex = 0
        ClassEditorInit
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleClassEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleCloseClient(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    AlertMsg ("You can't have multiple clients open at once!")
    DestroyGame
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleCloseClient", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleUpdateClass(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim ClassSize As Long
    Dim ClassData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    
    ' Update the Class
    ClassSize = LenB(Class(n))
    ReDim ClassData(ClassSize - 1)
    ClassData = buffer.ReadBytes(ClassSize)
    CopyMemory ByVal VarPtr(Class(n)), ByVal VarPtr(ClassData(0)), ClassSize
    Set buffer = Nothing
    
    ' Update lists
    If frmEditor_Spell.Visible Then
        SpellClassListInit
    End If
    
    If frmMenu.picCharacter.Visible Then
        UpdateCharacterMenu
    End If
    
    If frmEditor_Item.Visible Then
        ItemClassReqListInit
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleUpdateClass", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleLeaveGame(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim ThePlayer As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    ThePlayer = buffer.ReadLong
    Set buffer = Nothing
    
    Call ClearPlayer(ThePlayer)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleLeaveGame", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub SendCastSpell(ByVal SpellSlot As Byte)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong CCastSpell
    buffer.WriteLong SpellSlot
    SendData buffer.ToArray()
    Set buffer = Nothing
    
    SpellBuffer = SpellSlot
    SpellBufferTimer = timeGetTime
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SendCastSpell", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub SendNeedMap()
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong CNeedMap
    SendData buffer.ToArray()
    Set buffer = Nothing

    InitBattleMusic = False
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SendNeedMap", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub UpdateCharacterMenu()
    Dim i As Long, n As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Used for if the player is creating a new character
    frmMenu.Visible = True
    frmLoad.Visible = False
    ClearMenuPictures
    
    frmMenu.picCharacter.Visible = True
    
    ' Reset menu buttons
    CurButton_Menu = 2
    MenuButton(2).State = 2
    Call RenderButton_Menu(2)
    ResetMenuButtons
    
    frmMenu.txtCUser.SetFocus
    frmMenu.cmbClass.Clear
    
    n = 1
    
    For i = 1 To MAX_CLASSES
        If Class(i).Locked = 0 And Not Trim$(Class(i).name) = vbNullString Then
            frmMenu.cmbClass.AddItem Trim$(Class(i).name)
            ClassSelection(n) = i
            n = n + 1
        End If
    Next
    
    If frmMenu.cmbClass.ListCount = 0 Then
        frmMenu.cmbClass.AddItem "None"
        ClassSelection(n) = 1
    End If
    
    frmMenu.cmbClass.ListIndex = 0
    NewCharacterDrawSprite
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "UpdateCharacterMenu", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleEmoticonEditor(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    With frmEditor_Emoticon
        Editor = EDITOR_EMOTICON
        .lstIndex.Clear
        .ZOrder (0)
        
        ' Add the names
        For i = 1 To MAX_EMOTICONS
            .lstIndex.AddItem i & ": " & Trim$(Emoticon(i).Command)
        Next

        .Show
        .lstIndex.ListIndex = 0
        EmoticonEditorInit
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleEmoticonEditor", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleUpdateEmoticon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim n As Long
    Dim buffer As clsBuffer
    Dim EmoticonSize As Long
    Dim EmoticonData() As Byte

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    n = buffer.ReadLong
    
    ' Update the Emoticon
    EmoticonSize = LenB(Emoticon(n))
    ReDim EmoticonData(EmoticonSize - 1)
    EmoticonData = buffer.ReadBytes(EmoticonSize)
    CopyMemory ByVal VarPtr(Emoticon(n)), ByVal VarPtr(EmoticonData(0)), EmoticonSize
    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleUpdateEmoticon", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub HandleCheckEmoticon(ByVal Index As Long, ByRef Data() As Byte, ByVal StartAddr As Long, ByVal ExtraVar As Long)
    Dim buffer As clsBuffer
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    Set buffer = New clsBuffer
    buffer.WriteBytes Data()
    
    i = buffer.ReadLong
        
    tempPlayer(i).EmoticonNum = Emoticon(buffer.ReadLong).Pic
    tempPlayer(i).EmoticonTimer = timeGetTime + 2000
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "HandleCheckEmoticon", "modHandleData", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
