Attribute VB_Name = "modGameLogic"
Option Explicit

Public Sub GameLoop()
    Dim FrameTime As Long
    Dim tick As Long
    Dim TickFPS As Long
    Dim FPS As Long
    Dim i As Long
    Dim WalkTimer As Long
    Dim tmr25 As Long, tmr100 As Long, tmr250 As Long, tmr10000 As Long
    Dim X As Long, Y As Long
    Dim tmr500, Fadetmr As Long
    Dim Fogtmr As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' *** Start GameLoop ***
    Do While InGame
        tick = timeGetTime ' Set the inital tick
        ElapsedTime = tick - FrameTime ' Set the time difference for time-based movement
        FrameTime = tick ' Set the time second loop time to the first.

        ' * Check surface timers *
        ' Sprites
        If tmr10000 < tick Then
            ' Check Ping
            Call CheckPing
            Call SetPing
            tmr10000 = tick + 10000
        End If
        
        ' Animate the Autotiles
        If tmr250 < tick Then
            If AutoAnim < 3 Then
                AutoAnim = AutoAnim + 1
            Else
                AutoAnim = 0
            End If
            
            tmr250 = timeGetTime + 250
        End If
         
        If tmr25 < tick Then
            InGame = IsConnected
            Call CheckKeys ' Check to make sure they aren't trying to auto do anything

            ' Mute everything but still keep everything playing
            If frmMain.WindowState = vbMinimized Then
                If Not Audio.IsMuted Then Audio.MuteVolume
            Else
                If Audio.IsMuted Then Audio.UpdateVolume
            End If
         
            If GetForegroundWindow() = frmMain.hWnd Or GetForegroundWindow() = frmEditor_Events.hWnd Then
                Call CheckInputKeys ' Check which keys were pressed
            End If
            
            ' Check if we need to end the CD icon
            If NumSpellIcons > 0 Then
                For i = 1 To MAX_PLAYER_SPELLS
                    If PlayerSpells(i) > 0 Then
                        If SpellCD(i) > 0 Then
                            If SpellCD(i) + (Spell(PlayerSpells(i)).CDTime * 1000) < tick Then
                                SpellCD(i) = 0
                            End If
                        End If
                    End If
                Next
            End If

            ' Check if we need to unlock the player's spell casting restriction
            If SpellBuffer > 0 Then
                If PlayerSpells(SpellBuffer) > 0 Then
                    If SpellBufferTimer + (Spell(PlayerSpells(SpellBuffer)).CastTime * 1000) < tick Then
                        SpellBuffer = 0
                        SpellBufferTimer = 0
                    End If
                End If
            End If
            
            If CanMoveNow And GetForegroundWindow() = frmMain.hWnd Then
                Call CheckMovement ' Check if player is trying to move
                Call CheckAttack ' Check to see if player is trying to attack
            End If
            
            If tmr100 < tick Then
                ' Update item animations
                If NumItems > 0 Then DrawAnimatedItems
                
                tmr100 = tick + 100
                Call FindNearestTarget
            End If
            
            For i = 1 To MAX_ANIMATIONS
                CheckAnimInstance i
            Next
            
            ' Resize bars if vitals were changed
            ResizeHPBar
            ResizeMPBar
            ResizeExpBar
            
            tmr25 = tick + 25
        End If

        If tick > EventChatTimer Then
            If frmMain.lblEventChat.Visible = False Then
                If frmMain.picEventChat.Visible Then
                    frmMain.picEventChat.Visible = False
                End If
            End If
        End If
        
        ' Process input before rendering, otherwise input will be behind by 1 frame
        If WalkTimer < tick Then
            ' Process player movements (actually move them)
            For i = 1 To Player_HighIndex
                If IsPlaying(i) Then
                    Call ProcessPlayerMovement(i)
                End If
            Next

            ' Process npc movements (actually move them)
            For i = 1 To Map.NPC_HighIndex
                If Map.NPC(i) > 0 Then
                    Call ProcessNPCMovement(i)
                End If
            Next
            
            ' Process events movements (actually move them)
            If Map.CurrentEvents > 0 Then
                For i = 1 To Map.CurrentEvents
                    Call ProcessEventMovement(i)
                Next
            End If

            WalkTimer = tick + 15 ' Edit this Value to change WalkTimer
        End If
        
          ' Fog scrolling
        If Fogtmr < tick Then
            If CurrentFogSpeed > 0 Then
                ' Move
                fogOffsetX = fogOffsetX - 1
                fogOffsetY = fogOffsetY - 1
                
                ' Reset
                If fogOffsetX < -256 Then fogOffsetX = 0
                If fogOffsetY < -256 Then fogOffsetY = 0
                Fogtmr = tick + 255 - CurrentFogSpeed
            End If
        End If
        
        If tmr500 < tick Then
            ' Animate waterfalls
            Select Case waterfallFrame
                Case 0
                    waterfallFrame = 1
                Case 1
                    waterfallFrame = 2
                Case 2
                    waterfallFrame = 0
            End Select
            
            ' Animate autotiles
            Select Case autoTileFrame
                Case 0
                    autoTileFrame = 1
                Case 1
                    autoTileFrame = 2
                Case 2
                    autoTileFrame = 0
            End Select
            tmr500 = tick + 500
            redrawMapCache = True
        End If
        
        ProcessWeather
        
        If Fadetmr < tick Then
            If FadeType <> 2 Then
                If FadeType = 1 Then
                    If FadeAmount = 255 Then
                        
                    Else
                        FadeAmount = FadeAmount + 5
                    End If
                ElseIf FadeType = 0 Then
                    If FadeAmount = 0 Then
                    
                    Else
                        FadeAmount = FadeAmount - 5
                    End If
                End If
            End If
            
            Fadetmr = tick + 30
        End If

        ' *********************
        ' ** Render Graphics **
        ' *********************
        Call Render_Graphics
        
        Call Audio.UpdateSounds
        Call Audio.UpdateMapSounds
        DoEvents

        ' Lock fps
        Do While timeGetTime < tick + 15
            DoEvents
            Sleep 1
        Loop
        
        ' Calculate FPS
        If TickFPS < tick Then
            GameFPS = FPS
            TickFPS = tick + 1000
            FPS = 0
        Else
            FPS = FPS + 1
        End If
    Loop

    frmMain.Visible = False
    frmMenu.Visible = True
    If IsLogging Then
        IsLogging = False
    Else
        AlertMsg "Connection to server lost.", True
    End If
    GettingMap = True
    
    Call Audio.StopMusic
    Call Audio.PlayMusic(Options.MenuMusic)
    Call Audio.StopMapSounds
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "GameLoop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ProcessPlayerMovement(ByVal Index As Long)
    Dim MovementSpeed As Long

   ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Check if player is walking, and if so process moving them over
    Select Case TempPlayer(Index).Moving
        Case MOVING_WALKING: MovementSpeed = ((ElapsedTime / 1000) * ((MOVEMENT_SPEED / 2) * SIZE_X))
        Case MOVING_RUNNING: MovementSpeed = ((ElapsedTime / 1000) * (MOVEMENT_SPEED * SIZE_X))
        Case Else: Exit Sub
    End Select
    
    Select Case GetPlayerDir(Index)
        Case DIR_UP
            TempPlayer(Index).yOffset = TempPlayer(Index).yOffset - MovementSpeed
            If TempPlayer(Index).yOffset < 0 Then TempPlayer(Index).yOffset = 0
        Case DIR_DOWN
            TempPlayer(Index).yOffset = TempPlayer(Index).yOffset + MovementSpeed
            If TempPlayer(Index).yOffset > 0 Then TempPlayer(Index).yOffset = 0
        Case DIR_LEFT
            TempPlayer(Index).xOffset = TempPlayer(Index).xOffset - MovementSpeed
            If TempPlayer(Index).xOffset < 0 Then TempPlayer(Index).xOffset = 0
        Case DIR_RIGHT
            TempPlayer(Index).xOffset = TempPlayer(Index).xOffset + MovementSpeed
            If TempPlayer(Index).xOffset > 0 Then TempPlayer(Index).xOffset = 0
        Case DIR_UPLEFT
            TempPlayer(Index).yOffset = TempPlayer(Index).yOffset - MovementSpeed
            If TempPlayer(Index).yOffset < 0 Then TempPlayer(Index).yOffset = 0
            TempPlayer(Index).xOffset = TempPlayer(Index).xOffset - MovementSpeed
            If TempPlayer(Index).xOffset < 0 Then TempPlayer(Index).xOffset = 0
        Case DIR_UPRIGHT
            TempPlayer(Index).yOffset = TempPlayer(Index).yOffset - MovementSpeed
            If TempPlayer(Index).yOffset < 0 Then TempPlayer(Index).yOffset = 0
            TempPlayer(Index).xOffset = TempPlayer(Index).xOffset + MovementSpeed
            If TempPlayer(Index).xOffset > 0 Then TempPlayer(Index).xOffset = 0
        Case DIR_DOWNLEFT
            TempPlayer(Index).yOffset = TempPlayer(Index).yOffset + MovementSpeed
            If TempPlayer(Index).yOffset > 0 Then TempPlayer(Index).yOffset = 0
            TempPlayer(Index).xOffset = TempPlayer(Index).xOffset - MovementSpeed
            If TempPlayer(Index).xOffset < 0 Then TempPlayer(Index).xOffset = 0
        Case DIR_DOWNRIGHT
            TempPlayer(Index).yOffset = TempPlayer(Index).yOffset + MovementSpeed
            If TempPlayer(Index).yOffset > 0 Then TempPlayer(Index).yOffset = 0
            TempPlayer(Index).xOffset = TempPlayer(Index).xOffset + MovementSpeed
            If TempPlayer(Index).xOffset > 0 Then TempPlayer(Index).xOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If TempPlayer(Index).Moving > 0 Then
        If (TempPlayer(Index).xOffset = 0) And (TempPlayer(Index).yOffset = 0) Then
            TempPlayer(Index).Moving = 0
            If TempPlayer(Index).Step = 1 Then
                TempPlayer(Index).Step = 3
            Else
                TempPlayer(Index).Step = 1
            End If
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ProcessPlayerMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ProcessNPCMovement(ByVal MapNPCNum As Long)
    Dim MovementSpeed As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Check if NPC is walking, and if so process moving them over
    Select Case MapNPC(MapNPCNum).Moving
        Case MOVING_WALKING: MovementSpeed = MOVEMENT_SPEED / 2
        Case MOVING_RUNNING: MovementSpeed = MOVEMENT_SPEED
        Case Else: Exit Sub
    End Select
    
    Select Case MapNPC(MapNPCNum).Dir
        Case DIR_UP
            MapNPC(MapNPCNum).yOffset = MapNPC(MapNPCNum).yOffset - ((ElapsedTime / 1000) * (MovementSpeed * SIZE_Y))
            If MapNPC(MapNPCNum).yOffset < 0 Then MapNPC(MapNPCNum).yOffset = 0
            
        Case DIR_DOWN
            MapNPC(MapNPCNum).yOffset = MapNPC(MapNPCNum).yOffset + ((ElapsedTime / 1000) * (MovementSpeed * SIZE_Y))
            If MapNPC(MapNPCNum).yOffset > 0 Then MapNPC(MapNPCNum).yOffset = 0
            
        Case DIR_LEFT
            MapNPC(MapNPCNum).xOffset = MapNPC(MapNPCNum).xOffset - ((ElapsedTime / 1000) * (MovementSpeed * SIZE_X))
            If MapNPC(MapNPCNum).xOffset < 0 Then MapNPC(MapNPCNum).xOffset = 0
            
        Case DIR_RIGHT
            MapNPC(MapNPCNum).xOffset = MapNPC(MapNPCNum).xOffset + ((ElapsedTime / 1000) * (MovementSpeed * SIZE_X))
            If MapNPC(MapNPCNum).xOffset > 0 Then MapNPC(MapNPCNum).xOffset = 0
        
        Case DIR_UPLEFT
            MapNPC(MapNPCNum).yOffset = MapNPC(MapNPCNum).yOffset - ((ElapsedTime / 1000) * (MovementSpeed * SIZE_Y))
            If MapNPC(MapNPCNum).yOffset < 0 Then MapNPC(MapNPCNum).yOffset = 0
            MapNPC(MapNPCNum).xOffset = MapNPC(MapNPCNum).xOffset - ((ElapsedTime / 1000) * (MovementSpeed * SIZE_X))
            If MapNPC(MapNPCNum).xOffset < 0 Then MapNPC(MapNPCNum).xOffset = 0
            
        Case DIR_UPRIGHT
            MapNPC(MapNPCNum).yOffset = MapNPC(MapNPCNum).yOffset - ((ElapsedTime / 1000) * (MovementSpeed * SIZE_Y))
            If MapNPC(MapNPCNum).yOffset < 0 Then MapNPC(MapNPCNum).yOffset = 0
            MapNPC(MapNPCNum).xOffset = MapNPC(MapNPCNum).xOffset + ((ElapsedTime / 1000) * (MovementSpeed * SIZE_X))
            If MapNPC(MapNPCNum).xOffset > 0 Then MapNPC(MapNPCNum).xOffset = 0
            
        Case DIR_DOWNLEFT
            MapNPC(MapNPCNum).xOffset = MapNPC(MapNPCNum).xOffset - ((ElapsedTime / 1000) * (MovementSpeed * SIZE_X))
            If MapNPC(MapNPCNum).xOffset < 0 Then MapNPC(MapNPCNum).xOffset = 0
            MapNPC(MapNPCNum).yOffset = MapNPC(MapNPCNum).yOffset + ((ElapsedTime / 1000) * (MovementSpeed * SIZE_Y))
            If MapNPC(MapNPCNum).yOffset > 0 Then MapNPC(MapNPCNum).yOffset = 0
            
        Case DIR_DOWNRIGHT
            MapNPC(MapNPCNum).xOffset = MapNPC(MapNPCNum).xOffset + ((ElapsedTime / 1000) * (MovementSpeed * SIZE_X))
            If MapNPC(MapNPCNum).xOffset > 0 Then MapNPC(MapNPCNum).xOffset = 0
            MapNPC(MapNPCNum).yOffset = MapNPC(MapNPCNum).yOffset + ((ElapsedTime / 1000) * (MovementSpeed * SIZE_Y))
            If MapNPC(MapNPCNum).yOffset > 0 Then MapNPC(MapNPCNum).yOffset = 0
    End Select

    ' Check if completed walking over to the next tile
    If MapNPC(MapNPCNum).Moving > 0 Then
        If (MapNPC(MapNPCNum).xOffset = 0) And (MapNPC(MapNPCNum).yOffset = 0) Then
            MapNPC(MapNPCNum).Moving = 0
            If MapNPC(MapNPCNum).Step = 1 Then
                MapNPC(MapNPCNum).Step = 3
            Else
                MapNPC(MapNPCNum).Step = 1
            End If
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ProcessNPCMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub CheckMapGetItem()
    Dim buffer As New clsBuffer
    Dim i As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Set buffer = New clsBuffer

    If timeGetTime > TempPlayer(MyIndex).MapGetTimer + 250 Then
        For i = 1 To MAX_MAP_ITEMS
            If MapItem(i).num > 0 Then
                If MapItem(i).X = GetPlayerX(MyIndex) And MapItem(i).Y = GetPlayerY(MyIndex) Then
                    TempPlayer(MyIndex).MapGetTimer = timeGetTime
                    buffer.WriteLong CMapGetItem
                    buffer.WriteByte i
                    SendData buffer.ToArray()
                    Exit For
                End If
            End If
        Next
    End If

    Set buffer = Nothing
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckMapGetItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckAttack()
    Dim buffer As clsBuffer
    Dim AttackSpeed As Long, i As Long, X As Long, Y As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If ControlDown Then
        If InEvent Then Exit Sub ' in an event chat, fucking get outta here!
        If SpellBuffer > 0 Then Exit Sub ' currently casting a spell, can't attack
        If StunDuration > 0 Then Exit Sub ' stunned, can't attack
        If InBank Then Exit Sub
        If InShop > 0 Then Exit Sub
        If InTrade > 0 Then Exit Sub

        ' Speed from weapon
        If GetPlayerEquipment(MyIndex, Weapon) > 0 Then
            AttackSpeed = Item(GetPlayerEquipment(MyIndex, Weapon)).WeaponSpeed
        Else
            AttackSpeed = 1000
        End If

        If TempPlayer(MyIndex).AttackTimer + AttackSpeed < timeGetTime Then
            If TempPlayer(MyIndex).Attacking = 0 Then

                With TempPlayer(MyIndex)
                    .Attacking = 1
                    .AttackTimer = timeGetTime
                End With

                Set buffer = New clsBuffer
                buffer.WriteLong CAttack
                SendData buffer.ToArray()
                Set buffer = Nothing
            End If
        End If
        
        Select Case Player(MyIndex).Dir
            Case DIR_UP
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) - 1
            Case DIR_DOWN
                X = GetPlayerX(MyIndex)
                Y = GetPlayerY(MyIndex) + 1
            Case DIR_LEFT
                X = GetPlayerX(MyIndex) - 1
                Y = GetPlayerY(MyIndex)
            Case DIR_RIGHT
                X = GetPlayerX(MyIndex) + 1
                Y = GetPlayerY(MyIndex)
        End Select
        
        If timeGetTime > TempPlayer(MyIndex).EventTimer Then
            For i = 1 To Map.CurrentEvents
                If Map.MapEvents(i).Visible = 1 And Map.MapEvents(i).Trigger = 0 Then
                    If Map.MapEvents(i).X = X And Map.MapEvents(i).Y = Y Then
                        Set buffer = New clsBuffer
                        buffer.WriteLong CEvent
                        buffer.WriteLong i
                        SendData buffer.ToArray()
                        Set buffer = Nothing
                        TempPlayer(MyIndex).EventTimer = timeGetTime + 1000
                    End If
                End If
            Next
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckAttack", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function IsTryingToMove() As Boolean
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If MouseX = -1 And MouseY = -1 Then
        DirUp = False
        DirDown = False
        DirLeft = False
        DirRight = False
    End If
    
    If MyIndex > 0 Then
        If TempPlayer(MyIndex).Moving = 0 Then
            If ChatLocked Then
                If GetAsyncKeyState(VK_W) < 0 Or GetAsyncKeyState(VK_UP) < 0 Then
                    DirUp = True
                End If
                
                If GetAsyncKeyState(VK_S) < 0 Or GetAsyncKeyState(VK_DOWN) < 0 Then
                    DirDown = True
                End If
                
                If GetAsyncKeyState(VK_A) < 0 Or GetAsyncKeyState(VK_LEFT) < 0 Then
                    DirLeft = True
                End If
                
                If GetAsyncKeyState(VK_D) < 0 Or GetAsyncKeyState(VK_RIGHT) < 0 Then
                    DirRight = True
                End If

            Else
                MouseX = -1
                MouseY = -1
                DirUp = False
                DirDown = False
                DirLeft = False
                DirRight = False
            End If

        Else
            MouseX = -1
            MouseY = -1
            DirUp = False
            DirDown = False
            DirLeft = False
            DirRight = False
        End If
    End If
    
    If DirUp Or DirDown Or DirLeft Or DirRight Then
        IsTryingToMove = True

        If SpellBuffer > 0 Then
            SpellBuffer = 0
            SpellBufferTimer = 0
            Set buffer = New clsBuffer
            buffer.WriteLong CBreakSpell
            SendData buffer.ToArray
        End If
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsTryingToMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Function CanMove() As Boolean
    Dim d As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    CanMove = True

    ' Make sure they aren't trying to move when they are already moving
    If TempPlayer(MyIndex).Moving <> 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' Don't let them move if in an event
    If InEvent Then
        CanMove = False
        Exit Function
    End If

    ' Make sure they haven't just casted a spell
    If SpellBuffer > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' Make sure they're not stunned
    If StunDuration > 0 Then
        CanMove = False
        Exit Function
    End If
    
    ' Close open interfaces
    CloseInterfaces

    d = GetPlayerDir(MyIndex)
    
   '*********************
    '****MOVE UP LEFT*****
    '*********************
    If DirUp And DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_UPLEFT)
        CheckForNewMap

        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 And GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_UPLEFT) Then
                CanMove = False

                Exit Function
                
            Else
                
                Exit Function

            End If

        Else
            CanMove = False

            Exit Function

        End If
    End If

    '*********************
    '****MOVE UP RIGHT****
    '*********************
    If DirUp And DirRight Then
        Call SetPlayerDir(MyIndex, DIR_UPRIGHT)
        CheckForNewMap

        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 And GetPlayerX(MyIndex) < Map.MaxX Then
            If CheckDirection(DIR_UPRIGHT) Then
                CanMove = False

                Exit Function
                
            Else
                
                Exit Function
                
            End If

        Else
            CanMove = False

            Exit Function

        End If

    End If

    '*********************
    '***MOVE DOWN LEFT****
    '*********************
    If DirDown And DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_DOWNLEFT)
        CheckForNewMap

        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MaxY And GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_DOWNLEFT) Then
                CanMove = False

                Exit Function

            Else
                
                Exit Function
                
            End If

        Else
            CanMove = False

            Exit Function

        End If
    End If

    '*********************
    '***MOVE DOWN RIGHT***
    '*********************
    If DirDown And DirRight Then
        Call SetPlayerDir(MyIndex, DIR_DOWNRIGHT)
        CheckForNewMap

        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MaxY And GetPlayerX(MyIndex) < Map.MaxX Then
            If CheckDirection(DIR_DOWNRIGHT) Then
                CanMove = False
                
                Exit Function

            Else
                
                Exit Function
                
            End If

        Else
            CanMove = False

            Exit Function

        End If
    End If

    '*********************
    '******MOVE UP*******
    '*********************
    If DirUp Then
        Call SetPlayerDir(MyIndex, DIR_UP)
        CheckForNewMap
        
        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If

        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) > 0 Then
            If CheckDirection(DIR_UP) Then
                CanMove = False

                Exit Function

            Else
                
                Exit Function
                
            End If

        Else
            CanMove = False

            Exit Function

        End If
    End If

    '*********************
    '******MOVE DOWN******
    '*********************
    If DirDown Then
        Call SetPlayerDir(MyIndex, DIR_DOWN)
        CheckForNewMap

        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerY(MyIndex) < Map.MaxY Then
            If CheckDirection(DIR_DOWN) Then
                CanMove = False
                
                Exit Function

            Else
                
                Exit Function
                
            End If

        Else
            CanMove = False

            Exit Function

        End If
    End If

    '*********************
    '******MOVE LEFT******
    '*********************
    If DirLeft Then
        Call SetPlayerDir(MyIndex, DIR_LEFT)
        CheckForNewMap

        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If
        
        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) > 0 Then
            If CheckDirection(DIR_LEFT) Then
                CanMove = False
                
                Exit Function

            Else
                
                Exit Function
                
            End If

        Else
            CanMove = False

            Exit Function

        End If
    End If

    '*********************
    '*****MOVE RIGHT******
    '*********************
    If DirRight Then
        Call SetPlayerDir(MyIndex, DIR_RIGHT)
        CheckForNewMap

        If Last_Dir <> GetPlayerDir(MyIndex) Then
            Call SendPlayerDir
            Last_Dir = GetPlayerDir(MyIndex)
        End If

        ' Check to see if they are trying to go out of bounds
        If GetPlayerX(MyIndex) < Map.MaxX Then
            If CheckDirection(DIR_RIGHT) Then
                CanMove = False
                
                Exit Function

            Else
                
                Exit Function
                
            End If

        Else
            CanMove = False

            Exit Function

        End If
    End If
   
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "CanMove", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub CheckForNewMap()
    
    ' If debug mode then handle error
    If Options.Debug = 1 And App.LogMode = 1 Then On Error GoTo ErrorHandler

    If ((GetPlayerDir(MyIndex) = DIR_UP Or GetPlayerDir(MyIndex) = DIR_UPLEFT Or GetPlayerDir(MyIndex) = DIR_UPRIGHT) And GetPlayerY(MyIndex) = 0) Or ((GetPlayerDir(MyIndex) = DIR_DOWN Or GetPlayerDir(MyIndex) = DIR_DOWNLEFT Or GetPlayerDir(MyIndex) = DIR_DOWNRIGHT) And GetPlayerY(MyIndex) = Map.MaxY) Or ((GetPlayerDir(MyIndex) = DIR_LEFT Or GetPlayerDir(MyIndex) = DIR_DOWNLEFT Or GetPlayerDir(MyIndex) = DIR_UPLEFT) And GetPlayerX(MyIndex) = 0) Or ((GetPlayerDir(MyIndex) = DIR_RIGHT Or GetPlayerDir(MyIndex) = DIR_DOWNRIGHT Or GetPlayerDir(MyIndex) = DIR_UPRIGHT) And GetPlayerX(MyIndex) = Map.MaxX) Then
        Call SendPlayerRequestNewMap
    End If
   
    ' Error Handler
    Exit Sub

ErrorHandler:
    HandleError "CheckForNewMap", "modGameLogic", Err.Number, Err.Desciption, Err.Source, Err.HelpContext
    Err.Clear

    Exit Sub

End Sub

Function CheckDirection(ByVal Direction As Byte) As Boolean
    Dim X As Long
    Dim Y As Long
    Dim i As Long
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    CheckDirection = False
    
    ' Check directional blocking
    If Direction < DIR_RIGHT Then
        If IsDirBlocked(Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).DirBlock, Direction + 1) Then
            CheckDirection = True
            Exit Function
        End If
    End If

    Select Case Direction
        Case DIR_UP
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWN
            X = GetPlayerX(MyIndex)
            Y = GetPlayerY(MyIndex) + 1
        Case DIR_LEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex)
        Case DIR_RIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex)
        Case DIR_UPLEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex) - 1
        Case DIR_UPRIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex) - 1
        Case DIR_DOWNLEFT
            X = GetPlayerX(MyIndex) - 1
            Y = GetPlayerY(MyIndex) + 1
        Case DIR_DOWNRIGHT
            X = GetPlayerX(MyIndex) + 1
            Y = GetPlayerY(MyIndex) + 1
    End Select

    ' Check to see if the map tile is blocked or not
    If Map.Tile(X, Y).Type = TILE_TYPE_BLOCKED Then
        CheckDirection = True
        Exit Function
    End If

    ' Check to see if the map tile is a resource or not
    If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
        CheckDirection = True
        Exit Function
    End If
    
    ' Check if event is touched
    If timeGetTime = TempPlayer(MyIndex).EventTimer Then
        For i = 1 To Map.CurrentEvents
            If Map.MapEvents(i).Visible = 1 And Map.MapEvents(i).Trigger = 1 Then
                If Map.MapEvents(i).X = X And Map.MapEvents(i).Y = Y Then
                    Set buffer = New clsBuffer
                    buffer.WriteLong CEvent
                    buffer.WriteLong i
                    SendData buffer.ToArray()
                    Set buffer = Nothing
                    TempPlayer(MyIndex).EventTimer = timeGetTime + 1000
                End If
            End If
        Next
    End If
    
    ' Check to see if a player is already on that tile
    For i = 1 To Player_HighIndex
        If IsPlaying(i) And GetPlayerMap(i) = GetPlayerMap(MyIndex) Then
            If Moral(Map.Moral).PlayerBlocked = 1 Then
                If GetPlayerX(i) = X Then
                    If GetPlayerY(i) = Y Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next

    ' Check to see if a NPC is already on that tile
    For i = 1 To Map.NPC_HighIndex
        If MapNPC(i).num > 0 Then
            If MapNPC(i).X = X Then
                If MapNPC(i).Y = Y Then
                    CheckDirection = True
                    Exit Function
                End If
            End If
        End If
    Next
    
    ' Check to see if an event is already on that tile
    For i = 1 To Map.CurrentEvents
        If Map.MapEvents(i).Visible = 1 Then
            If Map.MapEvents(i).X = X Then
                If Map.MapEvents(i).Y = Y Then
                    If Map.MapEvents(i).WalkThrough = 0 Then
                        CheckDirection = True
                        Exit Function
                    End If
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "CheckDirection", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub CheckMovement()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If IsTryingToMove Then
        If CanMove Then
            SendPlayerMove
            If TempPlayer(MyIndex).xOffset = 0 Then
                If TempPlayer(MyIndex).yOffset = 0 Then
                    If Map.Tile(GetPlayerX(MyIndex), GetPlayerY(MyIndex)).Type = TILE_TYPE_WARP Then
                        GettingMap = True
                    End If
                End If
            End If
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CheckMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function IsInBounds()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If (CurX >= 0) Then
        If (CurX <= Map.MaxX) Then
            If (CurY >= 0) Then
                If (CurY <= Map.MaxY) Then
                    IsInBounds = True
                End If
            End If
        End If
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsInBounds", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub UpdateDrawMapName()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    DrawMapNameX = 12
    
    If GUIVisible Then
        DrawMapNameY = 88
    Else
        DrawMapNameY = 8
    End If
    
    DrawMapNameColor = Moral(Map.Moral).Color
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "UpdateDrawMapName", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub UseItem()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for subscript out of range
    If InventoryItemSelected < 1 Or InventoryItemSelected > MAX_INV Then Exit Sub
   
    Call SendUseItem(InventoryItemSelected)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "UseItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ForgetSpell(ByVal SpellSlot As Byte)
    Dim buffer As clsBuffer

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for subscript out of range
    If SpellSlot < 1 Or SpellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    ' Don't let them forget a spell which is in CD
    If SpellCD(SpellSlot) > 0 Then
        AddText "You can't forget a spell which is on cooldown!", BrightRed
        Exit Sub
    End If
    
    ' Don't let them forget a spell which is buffered
    If SpellBuffer = SpellSlot Then
        AddText "You can't forget a spell which you are casting!", BrightRed
        Exit Sub
    End If
    
    If PlayerSpells(SpellSlot) > 0 Then
        Set buffer = New clsBuffer
        buffer.WriteLong CForgetSpell
        buffer.WriteLong SpellSlot
        SendData buffer.ToArray()
        Set buffer = Nothing
    Else
        AddText "There is no spell here, report this to a staff member!", BrightRed
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ForgetSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CastSpell(ByVal SpellSlot As Byte)
    Dim X As Long, Y As Long, SpellCastType As Byte

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for subscript out of range
    If SpellSlot < 1 Or SpellSlot > MAX_PLAYER_SPELLS Then Exit Sub
    
    If SpellBuffer > 0 Then
        AddText "You can't cast another spell until the current spell has been completed.", BrightRed
        Exit Sub
    End If
    
    If SpellCD(SpellSlot) > 0 Then
        AddText "Spell has not cooled down yet!", BrightRed
        Exit Sub
    End If
    
    ' Don't allow them to cast if stunned
    If StunDuration > 0 Then Exit Sub

    If PlayerSpells(SpellSlot) > 0 Then
        ' Check if player has enough MP
        If GetPlayerVital(MyIndex, Vitals.MP) < Spell(PlayerSpells(SpellSlot)).MPCost Then
            Call AddText("Not enough mana to cast " & Trim$(Spell(PlayerSpells(SpellSlot)).Name) & ".", BrightRed)
            Exit Sub
        End If
        
        ' Find out what kind of spell it is (Self cast, Target or AOE)
        If Spell(PlayerSpells(SpellSlot)).Range > 0 Then
            ' Ranged attack, single target or aoe?
            If Not Spell(PlayerSpells(SpellSlot)).IsAoe Then
                SpellCastType = 2 ' targetted
            Else
                SpellCastType = 3 ' targetted aoe
            End If
        Else
            If Not Spell(PlayerSpells(SpellSlot)).IsAoe Then
                SpellCastType = 0 ' Self-cast
            Else
                SpellCastType = 1 ' Self-cast AoE
            End If
        End If
        
        Select Case SpellCastType
            Case 2, 3 ' targeted & targeted AOE
            ' Check if have target
            If MyTarget = 0 Then
                AddText "You do not have a target!", BrightRed
                Exit Sub
            End If
        End Select
                
        If MyTargetType = TARGET_TYPE_NPC Then
            ' Check if they have a target if spell is not self cast
            If Spell(PlayerSpells(SpellSlot)).Range > 0 Then
                ' Set the X and Y used for function below
                X = MapNPC(MyTarget).X
                Y = MapNPC(MyTarget).Y
                    
                ' Check if there in range
                If Not IsInRange(Spell(PlayerSpells(SpellSlot)).Range, GetPlayerX(MyIndex), GetPlayerY(MyIndex), X, Y) And Spell(PlayerSpells(SpellSlot)).CastTime = 0 Then
                    AddText "Target is not in range!", BrightRed
                    Exit Sub
                End If
            End If
        ElseIf MyTargetType = TARGET_TYPE_PLAYER Then
            ' Check if they have a target if spell is not self cast
            If Spell(PlayerSpells(SpellSlot)).Range > 0 Then
                ' Set the X and Y used for function below
                X = GetPlayerX(MyTarget)
                Y = GetPlayerY(MyTarget)
 
                ' Make sure we can only cast specific spells on ourselves
                If MyTargetType = TARGET_TYPE_PLAYER And MyTarget = MyIndex Then
                    If Spell(PlayerSpells(SpellSlot)).Type = SPELL_TYPE_DAMAGEHP Or Spell(PlayerSpells(SpellSlot)).Type = SPELL_TYPE_DAMAGEMP Or Spell(PlayerSpells(SpellSlot)).Type = SPELL_TYPE_WARPTOTARGET Then
                        AddText "You can't use this type of spell on yourself!", BrightRed
                        Exit Sub
                    End If
                End If
                
                ' Check if there in range
                If Not IsInRange(Spell(PlayerSpells(SpellSlot)).Range, GetPlayerX(MyIndex), GetPlayerY(MyIndex), X, Y) And Spell(PlayerSpells(SpellSlot)).CastTime = 0 Then
                    AddText "Target is not in range!", BrightRed
                    Exit Sub
                End If
            End If
        End If
        
        ' Can't use items while in a map that doesn't allow it
        If Moral(Map.Moral).CanCast = 0 Then
            AddText "You can't use spells in this area!", BrightRed
            Exit Sub
        End If
            
        If TempPlayer(MyIndex).Moving = 0 Then
            Call SendCastSpell(SpellSlot)
        Else
            Call AddText("Cannot cast while moving!", BrightRed)
        End If
    Else
        Call AddText("There is no spell here, report this to a Staff member!", BrightRed)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CastSpell", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function TwipsToPixels(ByVal Twip_Val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If XorY = 0 Then
        TwipsToPixels = Twip_Val / Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        TwipsToPixels = Twip_Val / Screen.TwipsPerPixelY
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "TwipsToPixels", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function PixelsToTwips(ByVal Pixel_Val As Long, ByVal XorY As Byte) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If XorY = 0 Then
        PixelsToTwips = Pixel_Val * Screen.TwipsPerPixelX
    ElseIf XorY = 1 Then
        PixelsToTwips = Pixel_Val * Screen.TwipsPerPixelY
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "PixelsToTwips", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function ConvertCurrency(ByVal Amount As Long) As String
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Int(Amount) < 1000 Then
        ConvertCurrency = Amount
    ElseIf Int(Amount) < 999999 Then
        ConvertCurrency = Int(Amount / 1000) & "k"
    ElseIf Int(Amount) < 999999999 Then
        ConvertCurrency = Int(Amount / 1000000) & "m"
    Else
        ConvertCurrency = Int(Amount / 1000000000) & "b"
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "ConvertCurrency", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub UpdateSpellDescWindow(ByVal SpellNum As Long, ByVal X As Long, ByVal Y As Long)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Check for off-screen
    If Y + frmMain.picSpellDesc.Height > frmMain.ScaleHeight Then
        Y = frmMain.ScaleHeight - frmMain.picSpellDesc.Height
    End If
    
    With frmMain
        .picSpellDesc.Top = Y
        .picSpellDesc.Left = X
        .picSpellDesc.Visible = True
        .picSpellDesc.ZOrder (0)
        
        If LastSpellDesc = SpellNum Then Exit Sub

        .lblSpellName.Caption = Trim$(Spell(SpellNum).Name)
        .lblSpellDesc.Caption = Trim$(Spell(SpellNum).Desc)
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "UpdteSpellWindow", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub UpdateItemDescWindow(ByVal ItemNum As Long, ByVal X As Long, ByVal Y As Long, Optional ByVal IsShopWindow As Boolean = False, Optional ByVal ShopValue As Long = 0, Optional ByVal ShopItem As Long)
    Dim i As Long
    Dim FirstLetter As String * 1
    Dim Name As String
    Dim Multiplier As Single
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    FirstLetter = LCase$(Left$(Trim$(Item(ItemNum).Name), 1))
   
    If FirstLetter = "$" Then
        Name = (Mid$(Trim$(Item(ItemNum).Name), 2, Len(Trim$(Item(ItemNum).Name)) - 1))
    Else
        Name = Trim$(Item(ItemNum).Name)
    End If
    
    ' Check for off-screen
    If Y + frmMain.picItemDesc.Height > frmMain.ScaleHeight Then
        Y = frmMain.ScaleHeight - frmMain.picItemDesc.Height
    End If
    
    ' Set z-order
    frmMain.picItemDesc.ZOrder (0)

    With frmMain
        .picItemDesc.Top = Y
        .picItemDesc.Left = X
        .picItemDesc.Visible = True
        
        If LastItemDesc = ItemNum Then Exit Sub ' Exit out after setting X + Y so we don't reset values
    
        ' Set the Name
        Select Case Item(ItemNum).Rarity
            Case 0 ' Grey
                .lblItemName.ForeColor = Grey
            Case 1 ' White
                .lblItemName.ForeColor = RGB(255, 255, 255)
            Case 2 ' Green
                .lblItemName.ForeColor = RGB(117, 198, 92)
            Case 3 ' Blue
                .lblItemName.ForeColor = RGB(103, 140, 224)
            Case 4 ' r
                .lblItemName.ForeColor = RGB(205, 34, 0)
            Case 5 ' Purple
                .lblItemName.ForeColor = RGB(193, 104, 204)
            Case 6 ' Orange
                .lblItemName.ForeColor = RGB(217, 150, 64)
        End Select
        
        ' Set captions
        .lblItemName.Caption = Name
        .lblItemDesc.Caption = Trim$(Item(ItemNum).Desc)
        .lblItemDesc = .lblItemDesc & vbNewLine
        
        LastItemDesc = 0
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "UpdateItemDescWindow", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CacheResources()
    Dim X As Long, Y As Long, Resource_Count As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Resource_Count = 0

    For X = 0 To Map.MaxX
        For Y = 0 To Map.MaxY
            If Map.Tile(X, Y).Type = TILE_TYPE_RESOURCE Then
                Resource_Count = Resource_Count + 1
                ReDim Preserve MapResource(0 To Resource_Count)
                MapResource(Resource_Count).X = X
                MapResource(Resource_Count).Y = Y
            End If
        Next
    Next

    Resource_Index = Resource_Count
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CacheResources", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CreateActionMsg(ByVal Message As String, ByVal Color As Long, ByVal msgType As Byte, ByVal X As Long, ByVal Y As Long)
    Dim i As Long '

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ActionMsgIndex = 0
    
    ' Carry on with the set
    For i = 1 To MAX_BYTE
        If ActionMsg(i).Timer = 0 Then
            ActionMsgIndex = i
            Exit For
        End If
    Next

    If ActionMsgIndex = 0 Then
        Call ClearActionMsg(1)
        ActionMsgIndex = 1
    End If
    
    With ActionMsg(ActionMsgIndex)
        .Message = Message
        .Color = Color
        .Type = msgType
        .Timer = timeGetTime
        .Scroll = 1
        .X = X
        .Y = Y
        .Alpha = 255
    End With

    If ActionMsg(ActionMsgIndex).Type = ACTIONMSG_SCROLL Then
        ActionMsg(ActionMsgIndex).Y = ActionMsg(ActionMsgIndex).Y + Random(-2, 6)
        ActionMsg(ActionMsgIndex).X = ActionMsg(ActionMsgIndex).X + Random(-8, 8)
    End If
    
    SetActionHighIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CreateActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CreateBlood(ByVal X As Long, ByVal Y As Long)
    Dim i As Long, Sprite As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    BloodIndex = 0
    
    ' Randomize sprite
    Sprite = Random(1, BloodCount)
    
    ' Make sure tile doesn't already have blood
    For i = 1 To Blood_HighIndex
        ' Already have blood
        If Blood(i).X = X And Blood(i).Y = Y Then
            ' Refresh the timer
            Blood(i).Timer = timeGetTime
            Exit Sub
        End If
    Next
    
    ' Carry on with the set
    For i = 1 To MAX_BYTE
        If Blood(i).Timer = 0 Then
            BloodIndex = i
            Exit For
        End If
    Next

    If BloodIndex = 0 Then
        Call ClearBlood(1)
        BloodIndex = 1
    End If
    
    ' Set the blood up
    With Blood(BloodIndex)
          .X = X
          .Y = Y
          .Sprite = Sprite
          .Timer = timeGetTime
          .Alpha = 255
      End With
      
    SetBloodHighIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CreateBlood", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CreateChatBubble(ByVal Target As Long, ByVal TargetType As Byte, ByVal Msg As String, ByVal Color As Long)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Loop through and see if that player/npc already has a chat bubble
    For i = 1 To MAX_BYTE
        If ChatBubble(i).TargetType = TargetType Then
            If ChatBubble(i).Target = Target Then
                ' Clear it out
                Call ClearChatBubble(i)
                Exit For
            End If
        End If
    Next
    
    ' Carry on with the set
    For i = 1 To MAX_BYTE
        If ChatBubble(i).Timer = 0 Then
            ChatBubbleIndex = i
            Exit For
        End If
    Next

    If ChatBubbleIndex = 0 Then
        Call ClearChatBubble(1)
        ChatBubbleIndex = 1
    End If
    
    ' Set the bubble up
    With ChatBubble(ChatBubbleIndex)
        .Target = Target
        .TargetType = TargetType
        .Msg = Msg
        .Color = Color
        .Timer = timeGetTime
        .active = True
        .Alpha = 255
    End With
    
    SetChatBubbleHighIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CreateChatBubble", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearActionMsg(ByVal Index As Byte, Optional ByVal SetHighIndex As Boolean = True)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With ActionMsg(Index)
        .Message = vbNullString
        .Timer = 0
        .Type = 0
        .Color = 0
        .Scroll = 0
        .X = 0
        .Y = 0
        .Alpha = 0
    End With
    
    If SetHighIndex Then
        SetActionHighIndex
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearActionMsg", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearBlood(ByVal Index As Long, Optional ByVal SetHighIndex As Boolean = True)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With Blood(Index)
        .X = 0
        .Y = 0
        .Sprite = 0
        .Timer = 0
        .Alpha = 0
    End With
    
    If SetHighIndex Then
        SetBloodHighIndex
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearBlood", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ClearChatBubble(ByVal Index As Long, Optional ByVal SetHighIndex As Boolean = True)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    With ChatBubble(Index)
        .Msg = vbNullString
        .Color = 0
        .Target = 0
        .TargetType = 0
        .Timer = 0
        .active = False
        .Alpha = 0
    End With
    
    If SetHighIndex Then
        SetChatBubbleHighIndex
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ClearChatBubble", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckAnimInstance(ByVal Index As Long)
    Dim looptime As Long
    Dim Layer As Long
    Dim FrameCount As Long
    Dim LockIndex As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' If doesn't exist then exit sub
    If AnimInstance(Index).Animation < 1 Then Exit Sub
    If AnimInstance(Index).Animation > MAX_ANIMATIONS Then Exit Sub
    
    For Layer = 0 To 1
        If AnimInstance(Index).Used(Layer) Then
            looptime = Animation(AnimInstance(Index).Animation).looptime(Layer)
            FrameCount = Animation(AnimInstance(Index).Animation).Frames(Layer)
            
            ' If zero'd then set so we don't have extra loop and/or frame
            If AnimInstance(Index).frameIndex(Layer) = 0 Then AnimInstance(Index).frameIndex(Layer) = 1
            If AnimInstance(Index).LoopIndex(Layer) = 0 Then AnimInstance(Index).LoopIndex(Layer) = 1
            
            ' Check if frame timer is set, and needs to have a frame change
            If AnimInstance(Index).Timer(Layer) + looptime <= timeGetTime Then
                ' Check if out of range
                If AnimInstance(Index).frameIndex(Layer) >= FrameCount Then
                    AnimInstance(Index).LoopIndex(Layer) = AnimInstance(Index).LoopIndex(Layer) + 1
                    If AnimInstance(Index).LoopIndex(Layer) > Animation(AnimInstance(Index).Animation).LoopCount(Layer) Then
                        AnimInstance(Index).Used(Layer) = False
                    Else
                        AnimInstance(Index).frameIndex(Layer) = 1
                    End If
                Else
                    AnimInstance(Index).frameIndex(Layer) = AnimInstance(Index).frameIndex(Layer) + 1
                End If
                AnimInstance(Index).Timer(Layer) = timeGetTime
            End If
        End If
    Next
    
    ' If neither layer is used, clear
    If AnimInstance(Index).Used(0) = False And AnimInstance(Index).Used(1) = False Then Call ClearAnimInstance(Index)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "checkAnimInstance", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub OpenShop(ByVal ShopNum As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    InShop = ShopNum
    TryingToFixItem = False
    
    If Shop(InShop).CanFix = 1 Then
        frmMain.ImgFix.Visible = True
    Else
        frmMain.ImgFix.Visible = False
    End If
    
    frmMain.picShop.Visible = True
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "OpenShop", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function GetBankItemNum(ByVal BankSlot As Byte) As Integer
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If BankSlot = 0 Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    If BankSlot > MAX_BANK Then
        GetBankItemNum = 0
        Exit Function
    End If
    
    GetBankItemNum = Bank.Item(BankSlot).num
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub SetBankItemNum(ByVal BankSlot As Byte, ByVal ItemNum As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Bank.Item(BankSlot).num = ItemNum
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetBankItemNum", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function GetBankItemValue(ByVal BankSlot As Byte) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    GetBankItemValue = Bank.Item(BankSlot).Value
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub SetBankItemValue(ByVal BankSlot As Byte, ByVal ItemValue As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Bank.Item(BankSlot).Value = ItemValue
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetBankItemValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Function GetPlayerBankItemDurValue(ByVal Index As Long, ByVal BankSlot As Byte) As Long
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Index < 1 Or Index > MAX_PLAYERS Then Exit Function
    GetPlayerBankItemDurValue = Bank.Item(Index).Durability
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "GetPlayerBankItemDurValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Sub SetPlayerBankItemDurValue(ByVal Index As Long, ByVal BankSlot As Byte, ByVal DurValue As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Bank.Item(Index).Durability = DurValue
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "SetPlayerBankItemDurValue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' BitWise Operators for directional blocking
Public Sub SetDirBlock(ByRef BlockVar As Byte, ByRef Dir As Byte, ByVal Block As Boolean)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Block Then
        BlockVar = BlockVar Or (2 ^ Dir)
    Else
        BlockVar = BlockVar And Not (2 ^ Dir)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "setDirBlock", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Function IsDirBlocked(ByRef BlockVar As Byte, ByRef Dir As Byte) As Boolean
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not BlockVar And (2 ^ Dir) Then
        IsDirBlocked = False
    Else
        IsDirBlocked = True
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "isDirBlocked", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function CheckGrammar(ByVal Word As String, Optional ByVal Caps As Byte = 0) As String
    Dim FirstLetter As String * 1
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
   
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
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "CheckGrammar", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Function CanPlayerPickupItem(ByVal Index As Long, ByVal MapItemNum As Integer)
    Dim MapNum As Integer

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    MapNum = GetPlayerMap(Index)
    
    ' Check for subscript out of range
    If MapNum < 1 Or MapNum > MAX_MAPS Then Exit Function
    
    If Moral(Map.Moral).CanPickupItem = 1 Then
        ' No lock or locked to player?
        If Trim$(MapItem(MapItemNum).PlayerName) = vbNullString Or Trim$(MapItem(MapItemNum).PlayerName) = GetPlayerName(Index) Then
            CanPlayerPickupItem = True
            Exit Function
        End If
    End If
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "CanPlayerPickupItem", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Function IsHotbarSlot(ByVal X As Single, ByVal Y As Single, Optional ByVal sType As Byte) As Long
    Dim Top As Long, Left As Long
    Dim i As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    For i = 1 To MAX_HOTBAR
        Top = HotbarTop
        Left = HotbarLeft + ((HotbarOffsetX + 32) * (((i - 1) Mod MAX_HOTBAR)))
        If X >= Left And X <= Left + PIC_X Then
            If Y >= Top And Y <= Top + PIC_Y Then
                If sType > 0 Then
                    If Not Hotbar(i).sType = sType Then Exit Function
                End If
                IsHotbarSlot = i
                Exit Function
            End If
        End If
    Next
    Exit Function
    
' Error handler
ErrorHandler:
    HandleError "IsHotbarSlot", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Public Sub PlaySoundEntity(ByVal X As Long, ByVal Y As Long, ByVal EntityType As Long, ByVal EntityNum As Long, Optional ByVal LockIndex As Long, Optional ByVal LockType As Byte)
    Dim SoundName As String

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If EntityNum <= 0 Then Exit Sub
    
    ' Find the sound
    Select Case EntityType
        ' Animations
        Case SoundEntity.seAnimation
            If EntityNum > MAX_ANIMATIONS Then Exit Sub
            SoundName = Trim$(Animation(EntityNum).Sound)
            
        ' Items
        Case SoundEntity.seItem
            If EntityNum > MAX_ITEMS Then Exit Sub
            SoundName = Trim$(Item(EntityNum).Sound)
        
        ' NPCs
        Case SoundEntity.seNPC
            If EntityNum > MAX_NPCS Then Exit Sub
            SoundName = Trim$(NPC(EntityNum).Sound)
        
        ' Resources
        Case SoundEntity.seResource
            If EntityNum > MAX_RESOURCES Then Exit Sub
            SoundName = Trim$(Resource(EntityNum).Sound)
        
        ' Spells
        Case SoundEntity.seSpell
            If EntityNum > MAX_SPELLS Then Exit Sub
            SoundName = Trim$(Spell(EntityNum).Sound)
        
        ' Other
        Case Else
            Exit Sub
    End Select
    
    ' Exit out if it's not set
    If Trim$(SoundName) = vbNullString Then Exit Sub

    ' Play the sound
    If LockType > 0 And LockIndex > 0 Then
      If LockType = TARGET_TYPE_PLAYER Then
         Audio.PlaySound SoundName, Player(LockIndex).X, Player(LockIndex).Y
      ElseIf LockType = TARGET_TYPE_NPC Then
         Audio.PlaySound SoundName, MapNPC(LockIndex).X, MapNPC(LockIndex).Y
      Else
         ' BUG
      End If
    Else
      Audio.PlaySound SoundName, X, Y
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "PlayMusic", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub Dialogue(ByVal DiTitle As String, ByVal DiText As String, ByVal DiIndex As Long, Optional ByVal IsYesNo As Boolean = False, Optional ByVal Data1 As Long = 0)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Exit out if we've already got a dialogue open
    If DialogueIndex > 0 Then Exit Sub
    
    ' Set global dialogue Index
    DialogueIndex = DiIndex
    
    ' Set the global dialogue data
    DialogueData1 = Data1

    ' Bring to front
    frmMain.picDialogue.ZOrder (0)

    ' Set the captions
    frmMain.lblDialogue_Title.Caption = DiTitle
    frmMain.lblDialogue_Text.Caption = DiText
    
    ' Show/hide buttons
    If Not IsYesNo Then
        frmMain.lblDialogue_Button(1).Visible = True ' Okay button
        frmMain.lblDialogue_Button(2).Visible = False ' Yes button
        frmMain.lblDialogue_Button(3).Visible = False ' No button
    Else
        frmMain.lblDialogue_Button(1).Visible = False ' Okay button
        frmMain.lblDialogue_Button(2).Visible = True ' Yes button
        frmMain.lblDialogue_Button(3).Visible = True ' No button
    End If
    
    ' Show txtDialogue if it is friend and set labels
    If DialogueIndex = DIALOGUE_TYPE_ADDFRIEND Or DialogueIndex = DIALOGUE_TYPE_REMOVEFRIEND Or DialogueIndex = DIALOGUE_TYPE_ADDFOE Or DialogueIndex = DIALOGUE_TYPE_REMOVEFOE Or DialogueIndex = DIALOGUE_TYPE_CHANGEGUILDACCESS Or DialogueIndex = DIALOGUE_TYPE_PARTYINVITE Or DialogueIndex = DIALOGUE_TYPE_GUILDINVITE Or DialogueIndex = DIALOGUE_TYPE_GUILDREMOVE Then
        frmMain.txtDialogue.Visible = True
        frmMain.lblDialogue_Button.Item(2).Caption = "Accept"
        frmMain.lblDialogue_Button.Item(3).Caption = "Cancel"
    Else
        frmMain.txtDialogue.Visible = False
        frmMain.lblDialogue_Button.Item(2).Caption = "Yes"
        frmMain.lblDialogue_Button.Item(3).Caption = "No"
    End If

    ' Show the dialogue box
    frmMain.picDialogue.Visible = True
    
    ' Set focus if it's visible
    If frmMain.txtDialogue.Visible = True Then
        frmMain.txtDialogue.SetFocus
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Dialogue", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DialogueHandler(ByVal Index As Long)
    ' Find out which button
    If Index = 1 Then ' Okay button
        ' Dialogue Index
        Select Case DialogueIndex
        
        End Select
    ElseIf Index = 2 Then ' Yes button
        ' Dialogue Index
        Select Case DialogueIndex
            Case DIALOGUE_TYPE_TRADE
                Call SendAcceptTradeRequest
            Case DIALOGUE_TYPE_FORGET
                Call ForgetSpell(DialogueData1)
            Case DIALOGUE_TYPE_PARTY
                Call SendAcceptParty
            Case DIALOGUE_TYPE_RESETSTATS
                Call SendUseItem(DialogueData1)
            Case DIALOGUE_TYPE_ADDFRIEND
                Call SendAddFriend(frmMain.txtDialogue.text)
            Case DIALOGUE_TYPE_REMOVEFRIEND
                Call SendRemoveFriend(frmMain.txtDialogue.text)
            Case DIALOGUE_TYPE_ADDFOE
                Call SendAddFoe(frmMain.txtDialogue.text)
            Case DIALOGUE_TYPE_REMOVEFOE
                Call SendRemoveFoe(frmMain.txtDialogue.text)
            Case DIALOGUE_TYPE_GUILD
                Call SendGuildAccept
            Case DIALOGUE_TYPE_GUILDDISBAND
                Call SendGuildDisband
            Case DIALOGUE_TYPE_DESTROYITEM
                Call SendDestroyItem(DialogueData1)
            Case DIALOGUE_TYPE_CHANGEGUILDACCESS
                If Not frmMain.lstGuild.ListIndex = -1 Then
                    Call SendGuildChangeAccess(frmMain.lstGuild.List(frmMain.lstGuild.ListIndex), frmMain.txtDialogue.text)
                End If
            Case DIALOGUE_TYPE_GUILDINVITE
                Call SendGuildInvite(frmMain.txtDialogue.text)
            Case DIALOGUE_TYPE_GUILDREMOVE
                Call SendGuildRemove(frmMain.txtDialogue.text)
            Case DIALOGUE_TYPE_PARTYINVITE
                Call SendPartyRequest(frmMain.txtDialogue.text)
        End Select
    ElseIf Index = 3 Then ' No button
        ' Dialogue Index
        Select Case DialogueIndex
            Case DIALOGUE_TYPE_TRADE
                Call SendDeclineTradeRequest
            Case DIALOGUE_TYPE_PARTY
                Call SendDeclineParty
            Case DIALOGUE_TYPE_GUILD
                Call SendGuildDecline
        End Select
    End If
End Sub

' Used to resize the game screen
Public Sub ResizeScreen(ByVal XWide As Long, ByVal YTall As Long)
    ' Set Min Map X and Y
    MIN_MAPX = XWide
    MIN_MAPY = YTall

    frmMain.Width = PixelsToTwips(MIN_MAPX * PIC_X, 0)
    
    Do While (frmMain.ScaleWidth < MIN_MAPX * PIC_X)
        frmMain.Width = frmMain.Width + Screen.TwipsPerPixelX
        frmMain.Left = (Screen.Width / 2) - (frmMain.Width / 2)
    Loop
    
    frmMain.Height = PixelsToTwips(MIN_MAPY * PIC_Y, 1)
    
    Do While (frmMain.ScaleHeight < MIN_MAPY * PIC_Y)
        frmMain.Height = frmMain.Height + Screen.TwipsPerPixelY
        frmMain.Top = (Screen.Height / 2) - (frmMain.Height / 2)
    Loop

    ' Resize and position the screen
    frmMain.picScreen.Width = MIN_MAPX * PIC_X
    frmMain.picScreen.Height = MIN_MAPY * PIC_Y
    frmMain.picForm.Width = MIN_MAPX * PIC_X
    frmMain.picForm.Height = MIN_MAPY * PIC_Y
    
    ' Recalculate the other variables
    HalfX = ((MIN_MAPX + 1) / 2) * PIC_X
    HalfY = ((MIN_MAPY + 1) / 2) * PIC_Y
    ScreenX = (MIN_MAPX) * PIC_X
    ScreenY = (MIN_MAPY) * PIC_Y
    StartXValue = ((MIN_MAPX + 1) / 2)
    StartYValue = ((MIN_MAPY + 1) / 2)
    EndXValue = (MIN_MAPX + 1)
    EndYValue = (MIN_MAPY + 1)
    CameraEndXValue = EndXValue + 1
    CameraEndYValue = EndYValue + 1
    
    frmMain.picScreen.Top = 0
    frmMain.picScreen.Left = 0
    frmMain.picForm.Top = 0
    frmMain.picForm.Left = 0
    frmMain.picMapEditor.Top = 0
    frmMain.picMapEditor.Left = 0
End Sub

Function IsInRange(ByVal Range As Byte, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Boolean
    Dim nVal As Long
    
    IsInRange = False
    
    nVal = Sqr((X1 - X2) ^ 2 + (Y1 - Y2) ^ 2)
    
    If nVal <= Range Then IsInRange = True
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

Public Function GetStatName(ByVal StatNum As Stats) As String
    Select Case StatNum
        Case Stats.Agility: GetStatName = "Agility"
        Case Stats.Endurance: GetStatName = "Endurance"
        Case Stats.Intelligence: GetStatName = "Intelligence"
        Case Stats.Strength: GetStatName = "Strength"
        Case Stats.Spirit: GetStatName = "Spirit"
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

Public Function GetColorName(ByVal ColorNum As Byte) As String
    Select Case ColorNum
        Case 0: GetColorName = "Black"
        Case 1: GetColorName = "Blue"
        Case 2: GetColorName = "Green"
        Case 3: GetColorName = "Cyan"
        Case 4: GetColorName = "Red"
        Case 5: GetColorName = "Magenta"
        Case 6: GetColorName = "Brown"
        Case 7: GetColorName = "Grey"
        Case 8: GetColorName = "Dark Grey"
        Case 9: GetColorName = "Bright Blue"
        Case 10: GetColorName = "Bright Green"
        Case 11: GetColorName = "Bright Cyan"
        Case 12: GetColorName = "Bright Red"
        Case 13: GetColorName = "Pink"
        Case 14: GetColorName = "Yellow"
        Case 15: GetColorName = "White"
        Case 16: GetColorName = "Dark Brown"
        Case 17: GetColorName = "Orange"
    End Select
End Function

Public Function GetCombatTreeName(ByVal CombatNum As Byte) As String
    Select Case CombatNum
        Case 1:
            GetCombatTreeName = "Melee"
        Case 2:
            GetCombatTreeName = "Range"
        Case 3:
            GetCombatTreeName = "Magic"
    End Select
End Function

Public Sub UpdatePlayerTitles()
    Dim i As Long, n As Long
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Clear the list
    frmMain.lstTitles.Clear
    frmMain.lstTitles.AddItem "None"
    
    ' Build the combo list
    For i = 1 To Player(MyIndex).AmountOfTitles
        If Player(MyIndex).title(i) > 0 Then
            frmMain.lstTitles.AddItem Trim$(title(Player(MyIndex).title(i)).Name)
        End If
    Next

    With frmMain
        If Player(MyIndex).CurTitle > 0 Then
            For i = 1 To MAX_TITLES
                If Player(MyIndex).CurTitle = Player(MyIndex).title(i) Then
                    frmMain.lblDesc.Caption = Trim$(title(Player(MyIndex).title(i)).Desc)
                    frmMain.lstTitles.ListIndex = i
                    Exit For
                End If
            Next
        Else
            .lblDesc.Caption = "None."
            frmMain.lstTitles.ListIndex = 0
        End If
        
        If .lstTitles.ListCount > 0 Then
            .lstTitles.Enabled = True
        Else
            .lstTitles.Enabled = False
        End If
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "UpdatePlayerTitles", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ToggleButtons(ByVal Visible As Boolean)
    Dim i As Byte
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ' Erase and reset buttons
    CurButton_Main = 0
    Call frmMain.ResetMainButtons
    
    If Visible Then
        For i = 1 To MAX_MAINBUTTONS
            If Not i = 14 And Not i = 15 Then
                frmMain.picButton(i).Visible = True
            End If
        Next
    Else
        For i = 1 To MAX_MAINBUTTONS
            If Not i = 14 And Not i = 15 Then
                frmMain.picButton(i).Visible = False
            End If
        Next
        Call frmMain.CloseAllPanels
    End If
    
    ButtonsVisible = Visible
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ToggleButtons", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ToggleGUI(Visible As Boolean)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Visible Then
        frmMain.picGUI_Vitals_Base.Visible = True
        frmMain.picChatbox.Visible = True
        frmMain.picChatbox.ZOrder (0)
        
        frmMain.picHotbar.Visible = True
    Else
        frmMain.picGUI_Vitals_Base.Visible = False
        frmMain.picChatbox.Visible = False
        frmMain.picHotbar.Visible = False
    End If
    
    GUIVisible = Visible
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ToggleGUI", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckForBattleMusic(ByVal MapNPCNum As Byte)
    Dim n As Byte
    
    ' Exit if invalid
    If MapNPCNum < 1 Or MapNPCNum > MAX_MAP_NPCS Then Exit Sub
    If MapNPC(MapNPCNum).TargetType = TARGET_TYPE_NPC Then Exit Sub
    
    ' Reset the old values
    If ActiveNPCTarget = MapNPCNum Then ActiveNPCTarget = 0
    
    If Options.Music = 1 And Options.BattleMusic = 1 Then
        If MapNPC(MapNPCNum).num > 0 Then
            If MapNPC(MapNPCNum).TargetType = TARGET_TYPE_PLAYER Then
                If MapNPC(MapNPCNum).Target = MyIndex And MapNPC(MapNPCNum).Target > 0 Then
                    If Len(Trim$(NPC(MapNPC(MapNPCNum).num).Music)) > 0 Then
                        ActiveNPCTarget = MapNPCNum
                    End If
                End If
                
                ' Check if party members are being targeted
                If Party.num > 0 Then
                    For n = 1 To MAX_PARTY_MEMBERS
                        If GetPlayerMap(MyIndex) = GetPlayerMap(Party.Member(n)) Then
                            If MapNPC(MapNPCNum).Target = Party.Member(n) And MapNPC(MapNPCNum).Target > 0 Then
                                If Len(Trim$(NPC(MapNPC(MapNPCNum).num).Music)) > 0 Then
                                    ActiveNPCTarget = MapNPCNum
                                End If
                            End If
                        End If
                    Next
                End If
            End If
            
            If InitBattleMusic = False Then Exit Sub
            
            If ActiveNPCTarget > 0 Then
                If MapNPC(ActiveNPCTarget).num > 0 Then
                    Call Audio.PlayMusic(Trim$(NPC(MapNPC(ActiveNPCTarget).num).Music))
                    BattleMusicActive = True
                    Exit Sub
                End If
            End If
        End If
    End If
    
    If InitBattleMusic = False Then Exit Sub
    If BattleMusicActive = False Then Exit Sub
    
    ' No battle music just reset it
    If Not Trim$(Map.Music) = vbNullString Then
        Call Audio.PlayMusic(Trim$(Map.Music))
        BattleMusicActive = False
    Else
        Call Audio.StopMusic
        BattleMusicActive = False
    End If
End Sub

Public Sub UpdateGuildPanel()
    If Not Trim$(Player(MyIndex).Guild) = vbNullString Then
        frmMain.lblGuildName.Caption = Player(MyIndex).Guild
    Else
        frmMain.lblGuildName.Caption = "Not in a Guild"
    End If
    
    If frmMain.picGuild_No.Visible Then
        If Not Trim$(Player(MyIndex).Guild) = vbNullString Then
            frmMain.picGuild.Visible = True
            frmMain.picGuild_No.Visible = False
            frmMain.picGuild.ZOrder (0)
        End If
    End If
    
    If frmMain.picGuild.Visible Then
        If Trim$(Player(MyIndex).Guild) = vbNullString Then
            frmMain.picGuild.Visible = False
        End If
    End If
End Sub

Public Sub PlayMapMusic()
    Dim i As Long
    Dim MusicFile As String
    
    BattleMusicActive = False
    ActiveNPCTarget = 0
    
    For i = 1 To Map.NPC_HighIndex - 1
        Call CheckForBattleMusic(i)
    Next
    
    InitBattleMusic = True
    
    Call CheckForBattleMusic(Map.NPC_HighIndex)
    
    ' Set the music to the music in the map properties
    If Options.BattleMusic = 0 Or Map.NPC_HighIndex = 0 Or BattleMusicActive = False Then
        MusicFile = Trim$(Map.Music)
        
        If MusicFile = vbNullString Then
            Call Audio.StopMusic
        ElseIf Not CurrentMusic = MusicFile And Not MusicFile = vbNullString Then
            Call Audio.PlayMusic(MusicFile)
        End If
    End If
End Sub

Public Sub SetActionHighIndex()
    Dim i As Long
    
    Action_HighIndex = 0
    
    ' Find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ActionMsg(i).Timer > 0 Then
            Action_HighIndex = i
            Exit For
        End If
    Next
End Sub

Public Sub SetBloodHighIndex()
    Dim i As Long
    
    Blood_HighIndex = 0
    
    ' Find the new high index
    For i = MAX_BYTE To 1 Step -1
        If Blood(i).Timer > 0 Then
            Blood_HighIndex = i
            Exit For
        End If
    Next
End Sub

Public Sub SetChatBubbleHighIndex()
    Dim i As Long
    
    ChatBubble_HighIndex = 0
    
    ' Find the new high index
    For i = MAX_BYTE To 1 Step -1
        If ChatBubble(i).Timer > 0 Then
            ChatBubble_HighIndex = i
            Exit For
        End If
    Next
End Sub

Public Sub RequestGuildResign()
    If GetPlayerGuild(MyIndex) = vbNullString Then
        AddText "You are not in a guild!", BrightRed
        Exit Sub
    End If
                    
    If GetPlayerGuildAccess(MyIndex) = MAX_GUILDACCESS Then
        Dialogue "Guild Disband", "Would you like to disband the guild " & GetPlayerGuild(MyIndex) & "?", DIALOGUE_TYPE_GUILDDISBAND, True
        Exit Sub
    End If
    
    If Not GetPlayerGuild(MyIndex) = vbNullString Then
        SendGuildDisband
    End If
End Sub

Sub ProcessEventMovement(ByVal id As Long)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' Check if NPC is walking, and if so process moving them over
    If Map.MapEvents(id).Moving = MOVING_WALKING Then
        Select Case Map.MapEvents(id).Dir
            Case DIR_UP
                Map.MapEvents(id).yOffset = Map.MapEvents(id).yOffset - ((ElapsedTime / 1000) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).yOffset < 0 Then Map.MapEvents(id).yOffset = 0
                
            Case DIR_DOWN
                Map.MapEvents(id).yOffset = Map.MapEvents(id).yOffset + ((ElapsedTime / 1000) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).yOffset > 0 Then Map.MapEvents(id).yOffset = 0
                
            Case DIR_LEFT
                Map.MapEvents(id).xOffset = Map.MapEvents(id).xOffset - ((ElapsedTime / 1000) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).xOffset < 0 Then Map.MapEvents(id).xOffset = 0
                
            Case DIR_RIGHT
                Map.MapEvents(id).xOffset = Map.MapEvents(id).xOffset + ((ElapsedTime / 1000) * (Map.MapEvents(id).MovementSpeed * SIZE_X))
                If Map.MapEvents(id).xOffset > 0 Then Map.MapEvents(id).xOffset = 0
                
        End Select
    
        ' Check if completed walking over to the next tile
        If (Map.MapEvents(id).xOffset = 0) And (Map.MapEvents(id).yOffset = 0) Then
            Map.MapEvents(id).Moving = 0
            If Map.MapEvents(id).Step = 1 Then
                Map.MapEvents(id).Step = 3
            Else
                Map.MapEvents(id).Step = 1
            End If
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ProcessEventMovement", "modGameLogic", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ProcessWeather()
    Dim i As Long
    
    If CurrentWeather > 0 Then
        i = Random(1, 101 - CurrentWeatherIntensity)
        
        If i = 1 Then
            'Add a new particle
            For i = 1 To MAX_WEATHER_PARTICLES
                If WeatherParticle(i).InUse = False Then
                    If Random(1, 2) = 1 Then
                        WeatherParticle(i).InUse = True
                        WeatherParticle(i).Type = CurrentWeather
                        WeatherParticle(i).Velocity = Random(8, 14)
                        WeatherParticle(i).X = (TileView.Left * 32) - 32
                        WeatherParticle(i).Y = (TileView.Top * 32) + Random(-32, frmMain.picScreen.ScaleHeight)
                    Else
                        WeatherParticle(i).InUse = True
                        WeatherParticle(i).Type = CurrentWeather
                        WeatherParticle(i).Velocity = Random(10, 15)
                        WeatherParticle(i).X = (TileView.Left * 32) + Random(-32, frmMain.picScreen.ScaleWidth)
                        WeatherParticle(i).Y = (TileView.Top * 32) - 32
                    End If
                    Exit For
                End If
            Next
        End If
    End If
    
    If CurrentWeather = WEATHER_TYPE_STORM Then
        i = Random(1, 400 - CurrentWeatherIntensity)
        If i = 1 Then
            ' Draw Thunder
            DrawThunder = Random(15, 22)
            Audio.PlaySound Sound_Thunder
        End If
    End If
    
    For i = 1 To MAX_WEATHER_PARTICLES
        If WeatherParticle(i).InUse Then
            If WeatherParticle(i).X > TileView.Right * 32 Or WeatherParticle(i).Y > TileView.Bottom * 32 Then
                WeatherParticle(i).InUse = False
            Else
                WeatherParticle(i).X = WeatherParticle(i).X + WeatherParticle(i).Velocity
                WeatherParticle(i).Y = WeatherParticle(i).Y + WeatherParticle(i).Velocity
            End If
        End If
    Next
End Sub

Function IsOdd(Number As Long) As Boolean
    If Number Mod 2 Then
        IsOdd = True
    Else
        IsOdd = False
    End If
End Function

Function IsEven(Number As Long) As Boolean
    If Number Mod 2 Then
        IsEven = False
    Else
        IsEven = True
    End If
End Function
