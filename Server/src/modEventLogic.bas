Attribute VB_Name = "modEventLogic"
Option Explicit

Public Sub RemoveDeadEvents()
Dim i As Long, MapNum As Long, Buffer As clsBuffer, X As Long, ID As Long, page As Long
    ' See if we should remove any events....
    For i = 1 To Player_HighIndex
        If TempPlayer(i).EventMap.CurrentEvents > 0 Then
            MapNum = GetPlayerMap(i)
            For X = 1 To TempPlayer(i).EventMap.CurrentEvents
                ID = TempPlayer(i).EventMap.EventPages(X).eventID
                page = TempPlayer(i).EventMap.EventPages(X).PageID
                If Map(MapNum).Events(ID).PageCount >= page Then
                
                    ' See if there is any reason to delete this event....
                    ' In other words, go back through conditions and make sure they all check up.
                    If TempPlayer(i).EventMap.EventPages(X).Visible = 1 Then
                        If Map(MapNum).Events(ID).Pages(page).chkHasItem = 1 Then
                            If HasItem(i, Map(MapNum).Events(ID).Pages(page).HasItemIndex) = 0 Then
                                TempPlayer(i).EventMap.EventPages(X).Visible = 0
                            End If
                        End If
                        
                        If Map(MapNum).Events(ID).Pages(page).chkSelfSwitch = 1 Then
                            If Map(MapNum).Events(ID).Pages(page).SelfSwitchCompare = 0 Then
                                If Map(MapNum).Events(ID).SelfSwitches(Map(MapNum).Events(ID).Pages(page).SelfSwitchIndex) = 0 Then
                                    TempPlayer(i).EventMap.EventPages(X).Visible = 0
                                End If
                            Else
                                If Map(MapNum).Events(ID).SelfSwitches(Map(MapNum).Events(ID).Pages(page).SelfSwitchIndex) = 1 Then
                                    TempPlayer(i).EventMap.EventPages(X).Visible = 0
                                End If
                            End If
                        End If
                        
                        If Map(MapNum).Events(ID).Pages(page).chkVariable = 1 Then
                            Select Case Map(MapNum).Events(ID).Pages(page).VariableCompare
                                Case 0
                                    If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(page).VariableIndex) <> Map(MapNum).Events(ID).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(X).Visible = 0
                                    End If
                                Case 1
                                    If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(page).VariableIndex) < Map(MapNum).Events(ID).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(X).Visible = 0
                                    End If
                                Case 2
                                    If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(page).VariableIndex) > Map(MapNum).Events(ID).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(X).Visible = 0
                                    End If
                                Case 3
                                    If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(page).VariableIndex) <= Map(MapNum).Events(ID).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(X).Visible = 0
                                    End If
                                Case 4
                                    If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(page).VariableIndex) >= Map(MapNum).Events(ID).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(X).Visible = 0
                                    End If
                                Case 5
                                    If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(page).VariableIndex) = Map(MapNum).Events(ID).Pages(page).VariableCondition Then
                                        TempPlayer(i).EventMap.EventPages(X).Visible = 0
                                    End If
                            End Select
                        End If
                        
                        If Map(MapNum).Events(ID).Pages(page).chkSwitch = 1 Then
                            If Map(MapNum).Events(ID).Pages(page).SwitchCompare = 1 Then
                                If Account(i).Chars(GetPlayerChar(i)).Switches(Map(MapNum).Events(ID).Pages(page).SwitchIndex) = 1 Then
                                    TempPlayer(i).EventMap.EventPages(X).Visible = 0
                                End If
                            Else
                                If Account(i).Chars(GetPlayerChar(i)).Switches(Map(MapNum).Events(ID).Pages(page).SwitchIndex) = 0 Then
                                    TempPlayer(i).EventMap.EventPages(X).Visible = 0
                                End If
                            End If
                        End If
                        
                        If TempPlayer(i).EventMap.EventPages(X).Visible = 0 Then
                            Set Buffer = New clsBuffer
                            Buffer.WriteLong SSpawnEvent
                            Buffer.WriteLong ID
                            With TempPlayer(i).EventMap.EventPages(X)
                                Buffer.WriteString Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(X).eventID).Name
                                Buffer.WriteLong .Dir
                                Buffer.WriteLong .GraphicNum
                                Buffer.WriteLong .GraphicType
                                Buffer.WriteLong .GraphicX
                                Buffer.WriteLong .GraphicX2
                                Buffer.WriteLong .GraphicY
                                Buffer.WriteLong .GraphicY2
                                Buffer.WriteLong .MovementSpeed
                                Buffer.WriteLong .X
                                Buffer.WriteLong .Y
                                Buffer.WriteLong .Position
                                Buffer.WriteLong .Visible
                                Buffer.WriteLong Map(MapNum).Events(ID).Pages(page).WalkAnim
                                Buffer.WriteLong Map(MapNum).Events(ID).Pages(page).DirFix
                                Buffer.WriteLong Map(MapNum).Events(ID).Pages(page).WalkThrough
                                Buffer.WriteLong Map(MapNum).Events(ID).Pages(page).ShowName
                                Buffer.WriteByte Map(MapNum).Events(ID).Pages(page).Trigger
                            End With
                            SendDataTo i, Buffer.ToArray
                            Set Buffer = Nothing
                        End If
                    End If
                End If
            Next
        End If
    Next
    
End Sub

Public Sub SpawnNewEvents()
    Dim Buffer As clsBuffer, PageID As Long, ID As Long, Compare As Long, i As Long, MapNum As Long, X As Long, z As Long, SpawnEvent As Boolean, p As Long
    
    ' That was only removing events... now we gotta worry about spawning them again, luckily, it is almost the same exact thing, but backwards!
    For i = 1 To Player_HighIndex
        If TempPlayer(i).EventMap.CurrentEvents > 0 Then
            MapNum = GetPlayerMap(i)
            
            For X = 1 To TempPlayer(i).EventMap.CurrentEvents
                ID = TempPlayer(i).EventMap.EventPages(X).eventID
                PageID = TempPlayer(i).EventMap.EventPages(X).PageID
                
                ' See if there is any reason to delete this event...
                ' In other words, go back through conditions and make sure they all check up
                For z = Map(MapNum).Events(ID).PageCount To 1 Step -1
                    SpawnEvent = True
                        
                    If Map(MapNum).Events(ID).Pages(z).chkHasItem = 1 Then
                        If HasItem(i, Map(MapNum).Events(ID).Pages(z).HasItemIndex) = 0 Then
                            SpawnEvent = False
                        End If
                    End If
                        
                    If Map(MapNum).Events(ID).Pages(z).chkSelfSwitch = 1 Then
                        If Map(MapNum).Events(ID).Pages(z).SelfSwitchCompare = 0 Then
                            Compare = 1
                        Else
                            Compare = 0
                        End If
                        
                        If Map(MapNum).Events(ID).Global = 1 Then
                            If Map(MapNum).Events(ID).SelfSwitches(Map(MapNum).Events(ID).Pages(z).SelfSwitchIndex) <> Compare Then
                                SpawnEvent = False
                            End If
                        Else
                            If TempPlayer(i).EventMap.EventPages(ID).SelfSwitches(Map(MapNum).Events(ID).Pages(z).SelfSwitchIndex) <> Compare Then
                                SpawnEvent = False
                            End If
                        End If
                    End If
                        
                    If Map(MapNum).Events(ID).Pages(z).chkVariable = 1 Then
                        Select Case Map(MapNum).Events(ID).Pages(z).VariableCompare
                            Case 0
                                If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(z).VariableIndex) <> Map(MapNum).Events(ID).Pages(z).VariableCondition Then
                                    SpawnEvent = False
                                End If
                            Case 1
                                If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(z).VariableIndex) < Map(MapNum).Events(ID).Pages(z).VariableCondition Then
                                    SpawnEvent = False
                                End If
                            Case 2
                                If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(z).VariableIndex) > Map(MapNum).Events(ID).Pages(z).VariableCondition Then
                                    SpawnEvent = False
                                End If
                            Case 3
                                If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(z).VariableIndex) <= Map(MapNum).Events(ID).Pages(z).VariableCondition Then
                                    SpawnEvent = False
                                End If
                            Case 4
                                If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(z).VariableIndex) >= Map(MapNum).Events(ID).Pages(z).VariableCondition Then
                                    SpawnEvent = False
                                End If
                            Case 5
                                If Account(i).Chars(GetPlayerChar(i)).Variables(Map(MapNum).Events(ID).Pages(z).VariableIndex) = Map(MapNum).Events(ID).Pages(z).VariableCondition Then
                                    SpawnEvent = False
                                End If
                        End Select
                    End If
                        
                    If Map(MapNum).Events(ID).Pages(z).chkSwitch = 1 Then
                        If Map(MapNum).Events(ID).Pages(z).SwitchCompare = 0 Then
                            If Account(i).Chars(GetPlayerChar(i)).Switches(Map(MapNum).Events(ID).Pages(z).SwitchIndex) = 0 Then
                                SpawnEvent = False
                            End If
                        Else
                            If Account(i).Chars(GetPlayerChar(i)).Switches(Map(MapNum).Events(ID).Pages(z).SwitchIndex) = 1 Then
                                SpawnEvent = False
                            End If
                        End If
                    End If
                        
                    If SpawnEvent = True Then
                        If TempPlayer(i).EventMap.EventPages(X).Visible = 1 Then
                            If z <= PageID Then
                                SpawnEvent = False
                            End If
                        End If
                    End If
                        
                    If SpawnEvent = True Then
                        With TempPlayer(i).EventMap.EventPages(X)
                            If Map(MapNum).Events(ID).Pages(z).GraphicType = 1 Then
                                Select Case Map(MapNum).Events(ID).Pages(z).GraphicY
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
                            
                            .GraphicNum = Map(MapNum).Events(ID).Pages(z).Graphic
                            .GraphicType = Map(MapNum).Events(ID).Pages(z).GraphicType
                            .GraphicX = Map(MapNum).Events(ID).Pages(z).GraphicX
                            .GraphicY = Map(MapNum).Events(ID).Pages(z).GraphicY
                            .GraphicX2 = Map(MapNum).Events(ID).Pages(z).GraphicX2
                            .GraphicY2 = Map(MapNum).Events(ID).Pages(z).GraphicY2
                            
                            Select Case Map(MapNum).Events(ID).Pages(z).MoveSpeed
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
                            
                            .Position = Map(MapNum).Events(ID).Pages(z).Position
                            .eventID = ID
                            .PageID = z
                            .Visible = 1
                                
                            .MoveType = Map(MapNum).Events(ID).Pages(z).MoveType
                            
                            If .MoveType = 2 Then
                                .MoveRouteCount = Map(MapNum).Events(ID).Pages(z).MoveRouteCount
                                ReDim .MoveRoute(0 To Map(MapNum).Events(ID).Pages(z).MoveRouteCount)
                                For p = 0 To Map(MapNum).Events(ID).Pages(z).MoveRouteCount
                                    .MoveRoute(p) = Map(MapNum).Events(ID).Pages(z).MoveRoute(p)
                                Next
                            End If
                                
                            .RepeatMoveRoute = Map(MapNum).Events(ID).Pages(z).RepeatMoveRoute
                            .IgnoreIfCannotMove = Map(MapNum).Events(ID).Pages(z).IgnoreMoveRoute
                                
                            .MoveFreq = Map(MapNum).Events(ID).Pages(z).MoveFreq
                            .MoveSpeed = Map(MapNum).Events(ID).Pages(z).MoveSpeed
                                
                            .WalkThrough = Map(MapNum).Events(ID).Pages(z).WalkThrough
                            .ShowName = Map(MapNum).Events(ID).Pages(z).ShowName
                            .WalkingAnim = Map(MapNum).Events(ID).Pages(z).WalkAnim
                            .FixedDir = Map(MapNum).Events(ID).Pages(z).DirFix
                            .Trigger = Map(MapNum).Events(ID).Pages(z).Trigger
                        End With
                         
                        Set Buffer = New clsBuffer
                        Buffer.WriteLong SSpawnEvent
                        Buffer.WriteLong ID
                        
                        With TempPlayer(i).EventMap.EventPages(X)
                            Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Name
                            Buffer.WriteLong .Dir
                            Buffer.WriteLong .GraphicNum
                            Buffer.WriteLong .GraphicType
                            Buffer.WriteLong .GraphicX
                            Buffer.WriteLong .GraphicX2
                            Buffer.WriteLong .GraphicY
                            Buffer.WriteLong .GraphicY2
                            Buffer.WriteLong .MovementSpeed
                            Buffer.WriteLong .X
                            Buffer.WriteLong .Y
                            Buffer.WriteLong .Position
                            Buffer.WriteLong .Visible
                            Buffer.WriteLong Map(MapNum).Events(ID).Pages(z).WalkAnim
                            Buffer.WriteLong Map(MapNum).Events(ID).Pages(z).DirFix
                            Buffer.WriteLong Map(MapNum).Events(ID).Pages(z).WalkThrough
                            Buffer.WriteLong Map(MapNum).Events(ID).Pages(z).ShowName
                            Buffer.WriteByte Map(MapNum).Events(ID).Pages(z).Trigger
                        End With
                        
                        SendDataTo i, Buffer.ToArray
                        Set Buffer = Nothing
                        GoTo nextevent
                    End If
                Next
nextevent:
            Next
        End If
    Next
    
End Sub

Public Sub ProcessEventMovement()
    Dim rand As Long, X As Long, i As Long, playerID As Long, eventID As Long, WalkThrough As Long, isglobal As Boolean, MapNum As Long, actualmovespeed As Long, Buffer As clsBuffer, z As Long, SendUpdate As Boolean
    
    ' Process Movement if needed for each player/each map/each event....
    For i = 1 To MAX_MAPS
        If PlayersOnMap(i) Then
            ' Manage Global Events First, then all the others.....
            If TempEventMap(i).EventCount > 0 Then
                For X = 1 To TempEventMap(i).EventCount
                    If TempEventMap(i).Events(X).Active = 1 Then
                        If TempEventMap(i).Events(X).MoveTimer <= timeGetTime Then
                            ' Real event - let's process it
                            Select Case TempEventMap(i).Events(X).MoveType
                                Case 0
                                    ' Nothing, fixed position
                                Case 1 ' Random, move randomly if possible...
                                    rand = Random(0, 3)
                                    
                                    If CanEventMove(0, i, TempEventMap(i).Events(X).X, TempEventMap(i).Events(X).Y, X, TempEventMap(i).Events(X).WalkThrough, rand, True) Then
                                        Select Case TempEventMap(i).Events(X).MoveSpeed
                                            Case 0
                                                EventMove 0, i, X, rand, 2, True
                                            Case 1
                                                EventMove 0, i, X, rand, 3, True
                                            Case 2
                                                EventMove 0, i, X, rand, 4, True
                                            Case 3
                                                EventMove 0, i, X, rand, 6, True
                                            Case 4
                                                EventMove 0, i, X, rand, 12, True
                                            Case 5
                                                EventMove 0, i, X, rand, 24, True
                                        End Select
                                    Else
                                        EventDir 0, i, X, rand, True
                                    End If
                                Case 2 ' Move Route - later
                                    With TempEventMap(i).Events(X)
                                        isglobal = True
                                        MapNum = i
                                        playerID = 0
                                        eventID = X
                                        WalkThrough = TempEventMap(i).Events(X).WalkThrough
                                        
                                        If .MoveRouteCount > 0 Then
                                            If .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 1 Then
                                                .MoveRouteStep = 0
                                            ElseIf .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 0 Then
                                                GoTo donotprocessmoveroute
                                            End If
                                            
                                            .MoveRouteStep = .MoveRouteStep + 1
                                            
                                            Select Case .MoveSpeed
                                                Case 0
                                                    actualmovespeed = 2
                                                Case 1
                                                    actualmovespeed = 3
                                                Case 2
                                                    actualmovespeed = 4
                                                Case 3
                                                    actualmovespeed = 6
                                                Case 4
                                                    actualmovespeed = 12
                                                Case 5
                                                    actualmovespeed = 24
                                            End Select
                                            
                                            Select Case .MoveRoute(.MoveRouteStep).Index
                                                Case 1
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, DIR_UP, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, DIR_UP, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 2
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, DIR_DOWN, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, DIR_DOWN, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 3
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, DIR_LEFT, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, DIR_LEFT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 4
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, DIR_RIGHT, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, DIR_RIGHT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 5
                                                    z = Random(0, 3)
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 6
                                                    If isglobal = False Then
                                                        If IsOneBlockAway(.X, .Y, GetPlayerX(playerID), GetPlayerY(playerID)) = True Then
                                                            EventDir playerID, GetPlayerMap(playerID), eventID, GetDirToPlayer(playerID, GetPlayerMap(playerID), eventID), False
                                                            If .IgnoreIfCannotMove = 0 Then
                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                            End If
                                                        Else
                                                            z = CanEventMoveTowardsPlayer(playerID, MapNum, eventID)
                                                            If z >= 4 Then
                                                                ' No
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            Else
                                                                If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, z, isglobal) Then
                                                                    EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                                Else
                                                                    If .IgnoreIfCannotMove = 0 Then
                                                                        .MoveRouteStep = .MoveRouteStep - 1
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 7
                                                    If isglobal = False Then
                                                        z = CanEventMoveAwayFromPlayer(playerID, MapNum, eventID)
                                                        If z >= 5 Then
                                                            ' No
                                                        Else
                                                            If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, z, isglobal) Then
                                                                EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 8
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, .Dir, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, .Dir, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 9
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                    End Select
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 10
                                                    .MoveTimer = timeGetTime + 100
                                                Case 11
                                                    .MoveTimer = timeGetTime + 500
                                                Case 12
                                                    .MoveTimer = timeGetTime + 1000
                                                Case 13
                                                    EventDir playerID, MapNum, eventID, DIR_UP, isglobal
                                                Case 14
                                                    EventDir playerID, MapNum, eventID, DIR_DOWN, isglobal
                                                Case 15
                                                    EventDir playerID, MapNum, eventID, DIR_LEFT, isglobal
                                                Case 16
                                                    EventDir playerID, MapNum, eventID, DIR_RIGHT, isglobal
                                                Case 17
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_DOWN
                                                        Case DIR_LEFT
                                                            z = DIR_UP
                                                        Case DIR_DOWN
                                                            z = DIR_LEFT
                                                    End Select
                                                    EventDir playerID, MapNum, eventID, z, isglobal
                                                Case 18
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_LEFT
                                                        Case DIR_RIGHT
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_RIGHT
                                                    End Select
                                                    EventDir playerID, MapNum, eventID, z, isglobal
                                                Case 19
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                    End Select
                                                    EventDir playerID, MapNum, eventID, z, isglobal
                                                Case 20
                                                    z = Random(0, 3)
                                                    EventDir playerID, MapNum, eventID, z, isglobal
                                                Case 21
                                                    If isglobal = False Then
                                                        z = GetDirToPlayer(playerID, MapNum, eventID)
                                                        EventDir playerID, MapNum, eventID, z, isglobal
                                                    End If
                                                Case 22
                                                    If isglobal = False Then
                                                        z = GetDirAwayFromPlayer(playerID, MapNum, eventID)
                                                        EventDir playerID, MapNum, eventID, z, isglobal
                                                    End If
                                                Case 23
                                                    .MoveSpeed = 0
                                                Case 24
                                                    .MoveSpeed = 1
                                                Case 25
                                                    .MoveSpeed = 2
                                                Case 26
                                                    .MoveSpeed = 3
                                                Case 27
                                                    .MoveSpeed = 4
                                                Case 28
                                                    .MoveSpeed = 5
                                                Case 29
                                                    .MoveFreq = 0
                                                Case 30
                                                    .MoveFreq = 1
                                                Case 31
                                                    .MoveFreq = 2
                                                Case 32
                                                    .MoveFreq = 3
                                                Case 33
                                                    .MoveFreq = 4
                                                Case 34
                                                    .WalkingAnim = 1
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 35
                                                    .WalkingAnim = 0
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 36
                                                    .FixedDir = 1
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 37
                                                    .FixedDir = 0
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 38
                                                    .WalkThrough = 1
                                                Case 39
                                                    .WalkThrough = 0
                                                Case 40
                                                    .Position = 0
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 41
                                                    .Position = 1
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 42
                                                    .Position = 2
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 43
                                                    .GraphicType = .MoveRoute(.MoveRouteStep).Data1
                                                    .GraphicNum = .MoveRoute(.MoveRouteStep).Data2
                                                    .GraphicX = .MoveRoute(.MoveRouteStep).Data3
                                                    .GraphicX2 = .MoveRoute(.MoveRouteStep).Data4
                                                    .GraphicY = .MoveRoute(.MoveRouteStep).Data5
                                                    .GraphicY2 = .MoveRoute(.MoveRouteStep).Data6
                                                    If .GraphicType = 1 Then
                                                        Select Case .GraphicY
                                                            Case 0
                                                                .Dir = DIR_DOWN
                                                            Case 1
                                                                .Dir = DIR_LEFT
                                                            Case 2
                                                                .Dir = DIR_RIGHT
                                                            Case 3
                                                                .Dir = DIR_UP
                                                        End Select
                                                    End If
                                                    'Need to Send Update to client
                                                    SendUpdate = True
                                            End Select
                                            
                                            If SendUpdate Then
                                                Set Buffer = New clsBuffer
                                                Buffer.WriteLong SSpawnEvent
                                                Buffer.WriteLong eventID
                                                With TempEventMap(i).Events(X)
                                                    Buffer.WriteString Map(GetPlayerMap(i)).Events(eventID).Name
                                                    Buffer.WriteLong .Dir
                                                    Buffer.WriteLong .GraphicNum
                                                    Buffer.WriteLong .GraphicType
                                                    Buffer.WriteLong .GraphicX
                                                    Buffer.WriteLong .GraphicX2
                                                    Buffer.WriteLong .GraphicY
                                                    Buffer.WriteLong .GraphicY2
                                                    Buffer.WriteLong .MoveSpeed
                                                    Buffer.WriteLong .X
                                                    Buffer.WriteLong .Y
                                                    Buffer.WriteLong .Position
                                                    Buffer.WriteLong .Active
                                                    Buffer.WriteLong .WalkingAnim
                                                    Buffer.WriteLong .FixedDir
                                                    Buffer.WriteLong .WalkThrough
                                                    Buffer.WriteLong .ShowName
                                                    Buffer.WriteByte .Trigger
                                                End With
                                                SendDataToMap i, Buffer.ToArray
                                                Set Buffer = Nothing
                                            End If
donotprocessmoveroute:
                                        End If
                                    End With
                            End Select
                            
                            Select Case TempEventMap(i).Events(X).MoveFreq
                                Case 0
                                    TempEventMap(i).Events(X).MoveTimer = timeGetTime + 4000
                                Case 1
                                    TempEventMap(i).Events(X).MoveTimer = timeGetTime + 2000
                                Case 2
                                    TempEventMap(i).Events(X).MoveTimer = timeGetTime + 1000
                                Case 3
                                    TempEventMap(i).Events(X).MoveTimer = timeGetTime + 500
                                Case 4
                                    TempEventMap(i).Events(X).MoveTimer = timeGetTime + 250
                            End Select
                        End If
                    End If
                Next
            End If
            ' HOPEFULLY this will not reduce CPS too much!
        End If
        DoEvents
    Next
End Sub

Public Sub ProcessLocalEventMovement()
Dim rand As Long, X As Long, i As Long, playerID As Long, eventID As Long, WalkThrough As Long, isglobal As Boolean, MapNum As Long, actualmovespeed As Long, Buffer As clsBuffer, z As Long, SendUpdate As Boolean
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            playerID = i
            If TempPlayer(i).EventMap.CurrentEvents > 0 Then
                For X = 1 To TempPlayer(i).EventMap.CurrentEvents
                    If Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(X).eventID).Global = 1 Then GoTo nextevent1
                    If TempPlayer(i).EventMap.EventPages(X).Visible = 1 Then
                        If TempPlayer(i).EventMap.EventPages(X).MoveTimer <= timeGetTime Then
                            ' Real event! Lets process it!
                            Select Case TempPlayer(i).EventMap.EventPages(X).MoveType
                                Case 0
                                    'Nothing, fixed position
                                Case 1 ' Random, move randomly if possible...
                                    rand = Random(0, 3)
                                    playerID = i
                                    If CanEventMove(i, GetPlayerMap(i), TempPlayer(i).EventMap.EventPages(X).X, TempPlayer(i).EventMap.EventPages(X).Y, X, TempPlayer(i).EventMap.EventPages(X).WalkThrough, rand, False) Then
                                        Select Case TempPlayer(i).EventMap.EventPages(X).MoveSpeed
                                            Case 0
                                                EventMove i, GetPlayerMap(i), X, rand, 2, False
                                            Case 1
                                                EventMove i, GetPlayerMap(i), X, rand, 3, False
                                            Case 2
                                                EventMove i, GetPlayerMap(i), X, rand, 4, False
                                            Case 3
                                                EventMove i, GetPlayerMap(i), X, rand, 6, False
                                            Case 4
                                                EventMove i, GetPlayerMap(i), X, rand, 12, False
                                            Case 5
                                                EventMove i, GetPlayerMap(i), X, rand, 24, False
                                        End Select
                                    Else
                                        EventDir 0, GetPlayerMap(i), X, rand, True
                                    End If
                                Case 2 'Move Route - later!
                                    With TempPlayer(i).EventMap.EventPages(X)
                                        isglobal = False
                                        MapNum = GetPlayerMap(i)
                                        playerID = i
                                        eventID = .eventID
                                        WalkThrough = .WalkThrough
                                        If .MoveRouteCount > 0 Then
                                            If .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 1 Then
                                                .MoveRouteStep = 0
                                            ElseIf .MoveRouteStep >= .MoveRouteCount And .RepeatMoveRoute = 0 Then
                                                GoTo donotprocessmoveroute1
                                            End If
                                            .MoveRouteStep = .MoveRouteStep + 1
                                            Select Case .MoveSpeed
                                                Case 0
                                                    actualmovespeed = 2
                                                Case 1
                                                    actualmovespeed = 3
                                                Case 2
                                                    actualmovespeed = 4
                                                Case 3
                                                    actualmovespeed = 6
                                                Case 4
                                                    actualmovespeed = 12
                                                Case 5
                                                    actualmovespeed = 24
                                            End Select
                                            Select Case .MoveRoute(.MoveRouteStep).Index
                                                Case 1
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, DIR_UP, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, DIR_UP, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 2
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, DIR_DOWN, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, DIR_DOWN, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 3
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, DIR_LEFT, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, DIR_LEFT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 4
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, DIR_RIGHT, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, DIR_RIGHT, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 5
                                                    z = Random(0, 3)
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 6
                                                    If isglobal = False Then
                                                        If IsOneBlockAway(.X, .Y, GetPlayerX(playerID), GetPlayerY(playerID)) = True Then
                                                            EventDir playerID, GetPlayerMap(playerID), eventID, GetDirToPlayer(playerID, GetPlayerMap(playerID), eventID), False
                                                            If .IgnoreIfCannotMove = 0 Then
                                                                .MoveRouteStep = .MoveRouteStep - 1
                                                            End If
                                                        Else
                                                            z = CanEventMoveTowardsPlayer(playerID, MapNum, eventID)
                                                            If z >= 4 Then
                                                                'No
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            Else
                                                                ' I is the direct, lets go...
                                                                If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, z, isglobal) Then
                                                                    EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                                Else
                                                                    If .IgnoreIfCannotMove = 0 Then
                                                                        .MoveRouteStep = .MoveRouteStep - 1
                                                                    End If
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 7
                                                    If isglobal = False Then
                                                        z = CanEventMoveAwayFromPlayer(playerID, MapNum, eventID)
                                                        If z >= 5 Then
                                                            'No
                                                        Else
                                                            ' I is the direct, lets go...
                                                            If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, z, isglobal) Then
                                                                EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                            Else
                                                                If .IgnoreIfCannotMove = 0 Then
                                                                    .MoveRouteStep = .MoveRouteStep - 1
                                                                End If
                                                            End If
                                                        End If
                                                    End If
                                                Case 8
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, .Dir, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, .Dir, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 9
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                    End Select
                                                    If CanEventMove(playerID, MapNum, .X, .Y, eventID, WalkThrough, z, isglobal) Then
                                                        EventMove playerID, MapNum, eventID, z, actualmovespeed, isglobal
                                                    Else
                                                        If .IgnoreIfCannotMove = 0 Then
                                                            .MoveRouteStep = .MoveRouteStep - 1
                                                        End If
                                                    End If
                                                Case 10
                                                    .MoveTimer = timeGetTime + 100
                                                Case 11
                                                    .MoveTimer = timeGetTime + 500
                                                Case 12
                                                    .MoveTimer = timeGetTime + 1000
                                                Case 13
                                                    EventDir playerID, MapNum, eventID, DIR_UP, isglobal
                                                Case 14
                                                    EventDir playerID, MapNum, eventID, DIR_DOWN, isglobal
                                                Case 15
                                                    EventDir playerID, MapNum, eventID, DIR_LEFT, isglobal
                                                Case 16
                                                    EventDir playerID, MapNum, eventID, DIR_RIGHT, isglobal
                                                Case 17
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_RIGHT
                                                        Case DIR_RIGHT
                                                            z = DIR_DOWN
                                                        Case DIR_LEFT
                                                            z = DIR_UP
                                                        Case DIR_DOWN
                                                            z = DIR_LEFT
                                                    End Select
                                                    EventDir playerID, MapNum, eventID, z, isglobal
                                                Case 18
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_LEFT
                                                        Case DIR_RIGHT
                                                            z = DIR_UP
                                                        Case DIR_LEFT
                                                            z = DIR_DOWN
                                                        Case DIR_DOWN
                                                            z = DIR_RIGHT
                                                    End Select
                                                    EventDir playerID, MapNum, eventID, z, isglobal
                                                Case 19
                                                    Select Case .Dir
                                                        Case DIR_UP
                                                            z = DIR_DOWN
                                                        Case DIR_RIGHT
                                                            z = DIR_LEFT
                                                        Case DIR_LEFT
                                                            z = DIR_RIGHT
                                                        Case DIR_DOWN
                                                            z = DIR_UP
                                                    End Select
                                                    EventDir playerID, MapNum, eventID, z, isglobal
                                                Case 20
                                                    z = Random(0, 3)
                                                    EventDir playerID, MapNum, eventID, z, isglobal
                                                Case 21
                                                    If isglobal = False Then
                                                        z = GetDirToPlayer(playerID, MapNum, eventID)
                                                        EventDir playerID, MapNum, eventID, z, isglobal
                                                    End If
                                                Case 22
                                                    If isglobal = False Then
                                                        z = GetDirAwayFromPlayer(playerID, MapNum, eventID)
                                                        EventDir playerID, MapNum, eventID, z, isglobal
                                                    End If
                                                Case 23
                                                    .MoveSpeed = 0
                                                Case 24
                                                    .MoveSpeed = 1
                                                Case 25
                                                    .MoveSpeed = 2
                                                Case 26
                                                    .MoveSpeed = 3
                                                Case 27
                                                    .MoveSpeed = 4
                                                Case 28
                                                    .MoveSpeed = 5
                                                Case 29
                                                    .MoveFreq = 0
                                                Case 30
                                                    .MoveFreq = 1
                                                Case 31
                                                    .MoveFreq = 2
                                                Case 32
                                                    .MoveFreq = 3
                                                Case 33
                                                    .MoveFreq = 4
                                                Case 34
                                                    .WalkingAnim = 1
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 35
                                                    .WalkingAnim = 0
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 36
                                                    .FixedDir = 1
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 37
                                                    .FixedDir = 0
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 38
                                                    .WalkThrough = 1
                                                Case 39
                                                    .WalkThrough = 0
                                                Case 40
                                                    .Position = 0
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 41
                                                    .Position = 1
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 42
                                                    .Position = 2
                                                    'Need to send update to client
                                                    SendUpdate = True
                                                Case 43
                                                    .GraphicType = .MoveRoute(.MoveRouteStep).Data1
                                                    .GraphicNum = .MoveRoute(.MoveRouteStep).Data2
                                                    .GraphicX = .MoveRoute(.MoveRouteStep).Data3
                                                    .GraphicX2 = .MoveRoute(.MoveRouteStep).Data4
                                                    .GraphicY = .MoveRoute(.MoveRouteStep).Data5
                                                    .GraphicY2 = .MoveRoute(.MoveRouteStep).Data6
                                                    If .GraphicType = 1 Then
                                                        Select Case .GraphicY
                                                            Case 0
                                                                .Dir = DIR_DOWN
                                                            Case 1
                                                                .Dir = DIR_LEFT
                                                            Case 2
                                                                .Dir = DIR_RIGHT
                                                            Case 3
                                                                .Dir = DIR_UP
                                                        End Select
                                                    End If
                                                    'Need to Send Update to client
                                                    SendUpdate = True
                                            End Select
                                            
                                            If SendUpdate Then
                                                Set Buffer = New clsBuffer
                                                Buffer.WriteLong SSpawnEvent
                                                Buffer.WriteLong eventID
                                                With TempPlayer(playerID).EventMap.EventPages(eventID)
                                                    Buffer.WriteString Map(GetPlayerMap(playerID)).Events(eventID).Name
                                                    Buffer.WriteLong .Dir
                                                    Buffer.WriteLong .GraphicNum
                                                    Buffer.WriteLong .GraphicType
                                                    Buffer.WriteLong .GraphicX
                                                    Buffer.WriteLong .GraphicX2
                                                    Buffer.WriteLong .GraphicY
                                                    Buffer.WriteLong .GraphicY2
                                                    Buffer.WriteLong .MoveSpeed
                                                    Buffer.WriteLong .X
                                                    Buffer.WriteLong .Y
                                                    Buffer.WriteLong .Position
                                                    Buffer.WriteLong .Visible
                                                    Buffer.WriteLong .WalkingAnim
                                                    Buffer.WriteLong .FixedDir
                                                    Buffer.WriteLong .WalkThrough
                                                    Buffer.WriteLong .ShowName
                                                    Buffer.WriteByte .Trigger
                                                End With
                                                SendDataTo playerID, Buffer.ToArray
                                                Set Buffer = Nothing
                                            End If
                                        End If
                                    End With
                            End Select

donotprocessmoveroute1:
                            Select Case TempPlayer(playerID).EventMap.EventPages(X).MoveFreq
                                Case 0
                                    TempPlayer(playerID).EventMap.EventPages(X).MoveTimer = timeGetTime + 4000
                                Case 1
                                    TempPlayer(playerID).EventMap.EventPages(X).MoveTimer = timeGetTime + 2000
                                Case 2
                                    TempPlayer(playerID).EventMap.EventPages(X).MoveTimer = timeGetTime + 1000
                                Case 3
                                    TempPlayer(playerID).EventMap.EventPages(X).MoveTimer = timeGetTime + 500
                                Case 4
                                    TempPlayer(playerID).EventMap.EventPages(X).MoveTimer = timeGetTime + 250
                            End Select
                        End If
                    End If
nextevent1:
                Next
            End If
        End If
        DoEvents
    Next
End Sub

Public Sub ProcessEventCommands()
    Dim Buffer As clsBuffer, i As Long, X As Long, z As Long, removeEventProcess As Boolean, w As Long, v As Long, p As Long
    
    ' Now, we process the damn things for commands
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            For X = 1 To TempPlayer(i).EventMap.CurrentEvents
                If TempPlayer(i).EventMap.EventPages(X).Visible Then
                    If Map(Account(i).Chars(GetPlayerChar(i)).Map).Events(TempPlayer(i).EventMap.EventPages(X).eventID).Pages(TempPlayer(i).EventMap.EventPages(X).PageID).Trigger = 2 Then 'Parallel Process baby!
                        If TempPlayer(i).EventProcessingCount > 0 Then
                            For z = 1 To TempPlayer(i).EventProcessingCount
                                If TempPlayer(i).EventProcessing(z).eventID = TempPlayer(i).EventMap.EventPages(X).eventID And TempPlayer(i).EventMap.EventPages(X).PageID = TempPlayer(i).EventProcessing(z).PageID Then
                                    ' Exit For
                                Else
                                    If z = TempPlayer(i).EventProcessingCount Then
                                        If Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(X).eventID).Pages(TempPlayer(i).EventMap.EventPages(X).PageID).CommandListCount > 0 Then
                                            ' Start new event processing
                                            TempPlayer(i).EventProcessingCount = TempPlayer(i).EventProcessingCount + 1
                                            ReDim Preserve TempPlayer(i).EventProcessing(TempPlayer(i).EventProcessingCount)
                                            With TempPlayer(i).EventProcessing(TempPlayer(i).EventProcessingCount)
                                                .ActionTimer = timeGetTime
                                                .CurList = 1
                                                .CurSlot = 1
                                                .eventID = TempPlayer(i).EventMap.EventPages(X).eventID
                                                .PageID = TempPlayer(i).EventMap.EventPages(X).PageID
                                                .WaitingForResponse = 0
                                                ReDim .ListLeftOff(0 To Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(X).eventID).Pages(TempPlayer(i).EventMap.EventPages(X).PageID).CommandListCount)
                                            End With
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next
                        Else
                            If Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(X).eventID).Pages(TempPlayer(i).EventMap.EventPages(X).PageID).CommandListCount > 0 Then
                                'Clearly need to start it!
                                TempPlayer(i).EventProcessingCount = 1
                                ReDim Preserve TempPlayer(i).EventProcessing(TempPlayer(i).EventProcessingCount)
                                With TempPlayer(i).EventProcessing(TempPlayer(i).EventProcessingCount)
                                    .ActionTimer = timeGetTime
                                    .CurList = 1
                                    .CurSlot = 1
                                    .eventID = TempPlayer(i).EventMap.EventPages(X).eventID
                                    .PageID = TempPlayer(i).EventMap.EventPages(X).PageID
                                    .WaitingForResponse = 0
                                    ReDim .ListLeftOff(0 To Map(GetPlayerMap(i)).Events(TempPlayer(i).EventMap.EventPages(X).eventID).Pages(TempPlayer(i).EventMap.EventPages(X).PageID).CommandListCount)
                                End With
                            End If
                        End If
                    End If
                End If
            Next
        End If
    Next
    
    ' That is it for starting parallel processes :D now we just have to make the code that actually processes the events to their fullest
    For i = 1 To Player_HighIndex
        If IsPlaying(i) Then
            If TempPlayer(i).EventProcessingCount > 0 Then
restartloop:
                For X = 1 To TempPlayer(i).EventProcessingCount
                    With TempPlayer(i).EventProcessing(X)
                        If TempPlayer(i).EventProcessingCount = 0 Then Exit Sub
                        removeEventProcess = False
                        If .WaitingForResponse = 2 Then
                            If TempPlayer(i).InShop = 0 Then
                                .WaitingForResponse = 0
                            End If
                        End If
                        If .WaitingForResponse = 3 Then
                            If TempPlayer(i).InBank = False Then
                                .WaitingForResponse = 0
                            End If
                        End If
                        If .WaitingForResponse = 0 Then
                            If .ActionTimer <= timeGetTime Then
restartlist:
                                If .ListLeftOff(.CurList) > 0 Then
                                    .CurSlot = .ListLeftOff(.CurList) + 1
                                End If
                                If .CurList > Map(Account(i).Chars(GetPlayerChar(i)).Map).Events(.eventID).Pages(.PageID).CommandListCount Then
                                    ' Get rid of this event, it is bad
                                    removeEventProcess = True
                                    GoTo endprocess
                                End If
                                If .CurSlot > Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).CommandCount Then
                                    If .CurList = 1 Then
                                        ' Get rid of this event, it is bad
                                        removeEventProcess = True
                                        GoTo endprocess
                                    Else
                                        .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).ParentList
                                        .CurSlot = 1
                                        GoTo restartlist
                                    End If
                                End If
                                ' If we are still here, then we are good to process shit :D
                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Index
                                    Case EventType.evAddText
                                        Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                            Case 0
                                                PlayerMsg i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                            Case 1
                                                MapMsg GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                            Case 2
                                                GlobalMsg Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        End Select
                                    Case EventType.evShowText
                                        Set Buffer = New clsBuffer
                                        Buffer.WriteLong SEventChat
                                        Buffer.WriteLong .eventID
                                        Buffer.WriteLong .PageID
                                        Buffer.WriteString ParseEventText(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text1)
                                        Buffer.WriteLong 0
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).CommandCount > .CurSlot Then
                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evShowText Or Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evShowChoices Then
                                                Buffer.WriteLong 1
                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evCondition Then
                                                Buffer.WriteLong 2
                                            Else
                                                Buffer.WriteLong 0
                                            End If
                                        Else
                                            Buffer.WriteLong 2
                                        End If
                                        Buffer.WriteLong Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data5
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                        .WaitingForResponse = 1
                                    Case EventType.evShowChoices
                                         Set Buffer = New clsBuffer
                                        Buffer.WriteLong SEventChat
                                        Buffer.WriteLong .eventID
                                        Buffer.WriteLong .PageID
                                        Buffer.WriteString ParseEventText(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text1)
                                        If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text2)) > 0 Then
                                            w = 1
                                            If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text3)) > 0 Then
                                                w = 2
                                                If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text4)) > 0 Then
                                                    w = 3
                                                    If Len(Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text5)) > 0 Then
                                                        w = 4
                                                    End If
                                                End If
                                            End If
                                        End If
                                        Buffer.WriteLong w
                                        For v = 1 To w
                                            Select Case v
                                                Case 1
                                                    Buffer.WriteString ParseEventText(i, Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text2))
                                                Case 2
                                                    Buffer.WriteString ParseEventText(i, Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text3))
                                                Case 3
                                                    Buffer.WriteString ParseEventText(i, Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text4))
                                                Case 4
                                                    Buffer.WriteString ParseEventText(i, Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text5))
                                            End Select
                                        Next
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).CommandCount > .CurSlot Then
                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evShowText Or Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evShowChoices Then
                                                Buffer.WriteLong 1
                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot + 1).Index = EventType.evCondition Then
                                                Buffer.WriteLong 2
                                            Else
                                                Buffer.WriteLong 0
                                            End If
                                        Else
                                            Buffer.WriteLong 2
                                        End If
                                        Buffer.WriteLong Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data5
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                        .WaitingForResponse = 1
                                    Case EventType.evPlayerVar
                                        Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                            Case 0
                                                Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                            Case 1
                                                Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) + Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                            Case 2
                                                Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) - Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                            Case 3
                                                Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) = Random(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data4)
                                        End Select
                                    Case EventType.evPlayerSwitch
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            Account(i).Chars(GetPlayerChar(i)).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) = 1
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            Account(i).Chars(GetPlayerChar(i)).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) = 0
                                        End If
                                    Case EventType.evSelfSwitch
                                        If Map(GetPlayerMap(i)).Events(.eventID).Global = 1 Then
                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 1
                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 0
                                            End If
                                        Else
                                            If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                                TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 1
                                            ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                                TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 + 1) = 0
                                            End If
                                        End If
                                    Case EventType.evCondition
                                        Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Condition
                                            Case 0
                                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                    Case 0
                                                        If Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 1
                                                        If Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) >= Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 2
                                                        If Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) <= Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 3
                                                        If Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) > Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 4
                                                        If Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) < Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 5
                                                        If Account(i).Chars(GetPlayerChar(i)).Variables(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) <> Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data3 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                End Select
                                            Case 1
                                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                    Case 0
                                                        If Account(i).Chars(GetPlayerChar(i)).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = 1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 1
                                                        If Account(i).Chars(GetPlayerChar(i)).Switches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = 0 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                End Select
                                            Case 2
                                                If HasItem(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) > 0 Then
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                    .CurSlot = 1
                                                Else
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                    .CurSlot = 1
                                                End If
                                            Case 3
                                                If Account(i).Chars(GetPlayerChar(i)).Class = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                    .CurSlot = 1
                                                Else
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                    .CurSlot = 1
                                                End If
                                            Case 4
                                                If HasSpell(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1) = True Then
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                    .CurSlot = 1
                                                Else
                                                    .ListLeftOff(.CurList) = .CurSlot
                                                    .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                    .CurSlot = 1
                                                End If
                                            Case 5
                                                Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                    Case 0
                                                        If GetPlayerLevel(i) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 1
                                                        If GetPlayerLevel(i) >= Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 2
                                                        If GetPlayerLevel(i) <= Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 3
                                                        If GetPlayerLevel(i) > Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 4
                                                        If GetPlayerLevel(i) < Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                    Case 5
                                                        If GetPlayerLevel(i) <> Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 Then
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                            .CurSlot = 1
                                                        Else
                                                            .ListLeftOff(.CurList) = .CurSlot
                                                            .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                            .CurSlot = 1
                                                        End If
                                                End Select
                                            Case 6
                                                If Map(GetPlayerMap(i)).Events(.eventID).Global = 1 Then
                                                    Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                        Case 0 ' Self Switch is true
                                                            If Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 1 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                        Case 1  ' Self switch is false
                                                            If Map(GetPlayerMap(i)).Events(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 0 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                    End Select
                                                Else
                                                    Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data2
                                                        Case 0 ' Self Switch is true
                                                            If TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 1 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                        Case 1  ' Self switch is false
                                                            If TempPlayer(i).EventMap.EventPages(.eventID).SelfSwitches(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.Data1 + 1) = 0 Then
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.CommandList
                                                                .CurSlot = 1
                                                            Else
                                                                .ListLeftOff(.CurList) = .CurSlot
                                                                .CurList = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).ConditionalBranch.ElseCommandList
                                                                .CurSlot = 1
                                                            End If
                                                    End Select
                                                End If
                                        End Select
                                        GoTo endprocess
                                    Case EventType.evExitProcess
                                        removeEventProcess = True
                                        GoTo endprocess
                                    Case EventType.evChangeItems
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            If HasItem(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) > 0 Then
                                                Call SetPlayerInvItemValue(i, HasItem(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1), Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3)
                                            End If
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            GiveInvItem i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 2 Then
                                            TakeInvItem i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                        End If
                                        SendInventory i
                                    Case EventType.evRestoreHP
                                        SetPlayerVital i, HP, GetPlayerMaxVital(i, HP)
                                        SendVital i, HP
                                    Case EventType.evRestoreMP
                                        SetPlayerVital i, MP, GetPlayerMaxVital(i, MP)
                                        SendVital i, MP
                                    Case EventType.evLevelUp
                                        SetPlayerExp i, GetPlayerNextLevel(i)
                                        CheckPlayerLevelUp i
                                    Case EventType.evChangeLevel
                                        SetPlayerLevel i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SetPlayerExp i, 0
                                        SendPlayerLevel i
                                        SendPlayerExp i
                                    Case EventType.evChangeSkills
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            If FindOpenSpellSlot(i) > 0 Then
                                                If HasSpell(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) = False Then
                                                    SetPlayerSpell i, FindOpenSpellSlot(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                Else
                                                    ' Error, already knows spell
                                                End If
                                            Else
                                                ' Error, no room for spells
                                            End If
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            If HasSpell(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) = True Then
                                                For p = 1 To MAX_PLAYER_SPELLS
                                                    If Account(i).Chars(GetPlayerChar(i)).Spell(p) = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 Then
                                                        SetPlayerSpell i, p, 0
                                                    End If
                                                Next
                                            End If
                                        End If
                                        SendPlayerSpells i
                                    Case EventType.evChangeClass
                                        ' Reset stats
                                        For z = 1 To Stats.Stat_count - 1
                                            Call SetPlayerPoints(i, GetPlayerPoints(i) + GetPlayerStat(i, z) - Class(GetPlayerClass(i)).Stat(z))
                                            Call SetPlayerStat(i, z, Class(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).Stat(z))
                                        Next
                                        Account(i).Chars(GetPlayerChar(i)).Class = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SendPlayerData i
                                    Case EventType.evChangeSprite
                                        Account(i).Chars(GetPlayerChar(i)).Sprite = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SendPlayerSprite i
                                    Case EventType.evChangeGender
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 0 Then
                                            Account(i).Chars(GetPlayerChar(i)).Gender = GENDER_MALE
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 1 Then
                                            Account(i).Chars(GetPlayerChar(i)).Gender = GENDER_FEMALE
                                        End If
                                        SendPlayerData i
                                    Case EventType.evChangePK
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 0 Then
                                            Account(i).Chars(GetPlayerChar(i)).PK = NO
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 = 1 Then
                                            Account(i).Chars(GetPlayerChar(i)).PK = YES
                                        End If
                                        SendPlayerPK i
                                    Case EventType.evWarpPlayer
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data4 = 0 Then
                                            PlayerWarp i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                        Else
                                            Account(i).Chars(GetPlayerChar(i)).Dir = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data4 - 1
                                            PlayerWarp i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                        End If
                                    Case EventType.evSetMoveRoute
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 <= Map(GetPlayerMap(i)).EventCount Then
                                            If Map(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).Global = 1 Then
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveType = 2
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).IgnoreIfCannotMove = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).RepeatMoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteCount = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).MoveRouteCount
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).MoveRoute
                                                TempEventMap(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteStep = 0
                                            Else
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveType = 2
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).IgnoreIfCannotMove = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).RepeatMoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteCount = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).MoveRouteCount
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRoute = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).MoveRoute
                                                TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).MoveRouteStep = 0
                                            End If
                                        End If
                                    Case EventType.evPlayAnimation
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 0 Then
                                            SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, GetPlayerX(i), GetPlayerY(i), TARGET_TYPE_PLAYER, i, i
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 1 Then
                                            If Map(GetPlayerMap(i)).Events(.eventID).Global = 1 Then
                                                SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).X, Map(GetPlayerMap(i)).Events(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3).Y
                                            Else
                                                SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3).X, TempPlayer(i).EventMap.EventPages(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3).Y, 0, 0, i
                                            End If
                                        ElseIf Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2 = 2 Then
                                            SendAnimation GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data4, 0, 0, i
                                        End If
                                    Case EventType.evCustomScript
                                        ' Runs Through Cases for a script
                                        Call CustomScript(i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1)
                                    Case EventType.evPlayBGM
                                        Set Buffer = New clsBuffer
                                        Buffer.WriteLong SPlayBGM
                                        Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text1
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                    Case EventType.evFadeoutBGM
                                        Set Buffer = New clsBuffer
                                        Buffer.WriteLong SFadeoutBGM
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                    Case EventType.evPlaySound
                                        Set Buffer = New clsBuffer
                                        Buffer.WriteLong SPlaySound
                                        Buffer.WriteString Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text1
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                    Case EventType.evStopSound
                                        Set Buffer = New clsBuffer
                                        Buffer.WriteLong SStopSound
                                        SendDataTo i, Buffer.ToArray
                                        Set Buffer = Nothing
                                    Case EventType.evSetAccess
                                        Account(i).Chars(GetPlayerChar(i)).Access = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                        SendPlayerData i
                                    Case EventType.evOpenShop
                                        If Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 > 0 Then ' shop exists?
                                            If Len(Trim$(Shop(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1).Name)) > 0 Then ' name exists?
                                                SendOpenShop i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                                TempPlayer(i).InShop = Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1 ' stops movement and the like
                                                .WaitingForResponse = 2
                                            End If
                                        End If
                                    Case EventType.evOpenBank
                                        SendBank i
                                        TempPlayer(i).InBank = True
                                        .WaitingForResponse = 3
                                    Case EventType.evGiveExp
                                        GivePlayerEXP i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                    Case EventType.evShowChatBubble
                                        Select Case Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                            Case TARGET_TYPE_PLAYER
                                                SendChatBubble GetPlayerMap(i), i, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text1, DarkBrown
                                            Case TARGET_TYPE_NPC
                                                SendChatBubble GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text1, DarkBrown
                                            Case TARGET_TYPE_EVENT
                                                SendChatBubble GetPlayerMap(i), Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text1, DarkBrown
                                        End Select
                                    Case EventType.evLabel
                                        ' Do nothing, just a label
                                    Case EventType.evGotoLabel
                                        ' Find the label's list of commands and slot
                                        FindEventLabel Trim$(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Text1), GetPlayerMap(i), .eventID, .PageID, .CurSlot, .CurList, .ListLeftOff
                                    Case EventType.evSpawnNpc
                                        If Map(GetPlayerMap(i)).NPC(Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1) > 0 Then
                                            SpawnNpc Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, GetPlayerMap(i), True
                                        End If
                                    Case EventType.evFadeIn
                                        SendSpecialEffect i, EFFECT_TYPE_FADEIN
                                    Case EventType.evFadeOut
                                        SendSpecialEffect i, EFFECT_TYPE_FADEOUT
                                    Case EventType.evFlashWhite
                                        SendSpecialEffect i, EFFECT_TYPE_FLASH
                                    Case EventType.evSetFog
                                        SendSpecialEffect i, EFFECT_TYPE_FOG, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3
                                    Case EventType.evSetweather
                                        SendSpecialEffect i, EFFECT_TYPE_WEATHER, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2
                                    Case EventType.evSetTint
                                        SendSpecialEffect i, EFFECT_TYPE_TINT, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data2, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data3, Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data4
                                    Case EventType.evWait
                                        .ActionTimer = timeGetTime + Map(GetPlayerMap(i)).Events(.eventID).Pages(.PageID).CommandList(.CurList).Commands(.CurSlot).Data1
                                End Select
                                .CurSlot = .CurSlot + 1
                            End If
                        End If
endprocess:
                    End With
                    If removeEventProcess = True Then
                        TempPlayer(i).EventProcessingCount = TempPlayer(i).EventProcessingCount - 1
                        If TempPlayer(i).EventProcessingCount <= 0 Then
                            ReDim TempPlayer(i).EventProcessing(0)
                            GoTo restartloop:
                        Else
                            For z = X To TempPlayer(i).EventProcessingCount - 1
                                TempPlayer(i).EventProcessing(X) = TempPlayer(i).EventProcessing(X + 1)
                            Next
                            ReDim Preserve TempPlayer(i).EventProcessing(TempPlayer(i).EventProcessingCount)
                            GoTo restartloop
                        End If
                    End If
                Next
            End If
        End If
    Next
End Sub

Public Sub UpdateEventLogic()
    Dim i As Long, X As Long, Y As Long, z As Long, MapNum As Long, ID As Long
    Dim page As Long, Buffer As clsBuffer, SpawnEvent As Boolean, p As Long, rand As Long, isglobal As Boolean, actualmovespeed As Long, playerID As Long, WalkThrough As Long, eventID As Long, SendUpdate As Boolean, removeEventProcess As Boolean, w As Long, v As Long
    
    ' Check Removing and Adding of Events (Did switches change or something?)
    RemoveDeadEvents
    SpawnNewEvents
    ProcessEventMovement
    ProcessLocalEventMovement
    ProcessEventCommands
End Sub

Sub SendSwitchesAndVariables(Index As Long, Optional everyone As Boolean = False)
    Dim Buffer As clsBuffer, i As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SSwitchesAndVariables
    
    For i = 1 To MAX_SWITCHES
        Buffer.WriteString Switches(i)
    Next
    
    For i = 1 To MAX_VARIABLES
        Buffer.WriteString Variables(i)
    Next
    
    If everyone Then
        SendDataToAll Buffer.ToArray
    Else
        SendDataTo Index, Buffer.ToArray
    End If

    Set Buffer = Nothing
End Sub

Sub SendMapEventData(Index As Long)
    Dim Buffer As clsBuffer, i As Long, X As Long, Y As Long, z As Long, MapNum As Long, w As Long
    
    Set Buffer = New clsBuffer
    Buffer.WriteLong SMapEventData
    MapNum = GetPlayerMap(Index)
    
    Buffer.WriteLong Map(MapNum).EventCount
        
    If Map(MapNum).EventCount > 0 Then
        For i = 1 To Map(MapNum).EventCount
            With Map(MapNum).Events(i)
                Buffer.WriteString .Name
                Buffer.WriteLong .Global
                Buffer.WriteLong .X
                Buffer.WriteLong .Y
                Buffer.WriteLong .PageCount
            End With
            
            If Map(MapNum).Events(i).PageCount > 0 Then
                For X = 1 To Map(MapNum).Events(i).PageCount
                    With Map(MapNum).Events(i).Pages(X)
                        Buffer.WriteLong .chkVariable
                        Buffer.WriteLong .VariableIndex
                        Buffer.WriteLong .VariableCondition
                        Buffer.WriteLong .VariableCompare
                            
                        Buffer.WriteLong .chkSwitch
                        Buffer.WriteLong .SwitchIndex
                        Buffer.WriteLong .SwitchCompare
                        
                        Buffer.WriteLong .chkHasItem
                        Buffer.WriteLong .HasItemIndex
                            
                        Buffer.WriteLong .chkSelfSwitch
                        Buffer.WriteLong .SelfSwitchIndex
                        Buffer.WriteLong .SelfSwitchCompare
                            
                        Buffer.WriteLong .GraphicType
                        Buffer.WriteLong .Graphic
                        Buffer.WriteLong .GraphicX
                        Buffer.WriteLong .GraphicY
                        Buffer.WriteLong .GraphicX2
                        Buffer.WriteLong .GraphicY2
                        
                        Buffer.WriteLong .MoveType
                        Buffer.WriteLong .MoveSpeed
                        Buffer.WriteLong .MoveFreq
                        Buffer.WriteLong .MoveRouteCount
                        
                        Buffer.WriteLong .IgnoreMoveRoute
                        Buffer.WriteLong .RepeatMoveRoute
                            
                        If .MoveRouteCount > 0 Then
                            For Y = 1 To .MoveRouteCount
                                Buffer.WriteLong .MoveRoute(Y).Index
                                Buffer.WriteLong .MoveRoute(Y).Data1
                                Buffer.WriteLong .MoveRoute(Y).Data2
                                Buffer.WriteLong .MoveRoute(Y).Data3
                                Buffer.WriteLong .MoveRoute(Y).Data4
                                Buffer.WriteLong .MoveRoute(Y).Data5
                                Buffer.WriteLong .MoveRoute(Y).Data6
                            Next
                        End If
                            
                        Buffer.WriteLong .WalkAnim
                        Buffer.WriteLong .DirFix
                        Buffer.WriteLong .WalkThrough
                        Buffer.WriteLong .ShowName
                        Buffer.WriteLong .Trigger
                        Buffer.WriteLong .CommandListCount
                        
                        Buffer.WriteLong .Position
                    End With
                        
                    If Map(MapNum).Events(i).Pages(X).CommandListCount > 0 Then
                        For Y = 1 To Map(MapNum).Events(i).Pages(X).CommandListCount
                            Buffer.WriteLong Map(MapNum).Events(i).Pages(X).CommandList(Y).CommandCount
                            Buffer.WriteLong Map(MapNum).Events(i).Pages(X).CommandList(Y).ParentList
                            If Map(MapNum).Events(i).Pages(X).CommandList(Y).CommandCount > 0 Then
                                For z = 1 To Map(MapNum).Events(i).Pages(X).CommandList(Y).CommandCount
                                    With Map(MapNum).Events(i).Pages(X).CommandList(Y).Commands(z)
                                        Buffer.WriteLong .Index
                                        Buffer.WriteString .Text1
                                        Buffer.WriteString .Text2
                                        Buffer.WriteString .Text3
                                        Buffer.WriteString .Text4
                                        Buffer.WriteString .Text5
                                        Buffer.WriteLong .Data1
                                        Buffer.WriteLong .Data2
                                        Buffer.WriteLong .Data3
                                        Buffer.WriteLong .Data4
                                        Buffer.WriteLong .Data5
                                        Buffer.WriteLong .Data6
                                        Buffer.WriteLong .ConditionalBranch.CommandList
                                        Buffer.WriteLong .ConditionalBranch.Condition
                                        Buffer.WriteLong .ConditionalBranch.Data1
                                        Buffer.WriteLong .ConditionalBranch.Data2
                                        Buffer.WriteLong .ConditionalBranch.Data3
                                        Buffer.WriteLong .ConditionalBranch.ElseCommandList
                                        Buffer.WriteLong .MoveRouteCount
                                        
                                        If .MoveRouteCount > 0 Then
                                            For w = 1 To .MoveRouteCount
                                                Buffer.WriteLong .MoveRoute(w).Index
                                                Buffer.WriteLong .MoveRoute(w).Data1
                                                Buffer.WriteLong .MoveRoute(w).Data2
                                                Buffer.WriteLong .MoveRoute(w).Data3
                                                Buffer.WriteLong .MoveRoute(w).Data4
                                                Buffer.WriteLong .MoveRoute(w).Data5
                                                Buffer.WriteLong .MoveRoute(w).Data6
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
    
    SendDataTo Index, Buffer.ToArray
    Set Buffer = Nothing
    Call SendSwitchesAndVariables(Index)
End Sub

Function ParseEventText(ByVal Index As Long, ByVal txt As String) As String
    Dim i As Long, X As Long, newtxt As String, parsestring As String, z As Long
    
    txt = Replace(txt, "/name", Trim$(GetPlayerName(Index)))
    txt = Replace(txt, "/p", Trim$(GetPlayerName(Index)))
    
    Do While InStr(1, txt, "/v") > 0
        X = InStr(1, txt, "/v")
        If X > 0 Then
            i = 0
            
            Do Until IsNumeric(Mid(txt, X + 2 + i, 1)) = False
                i = i + 1
            Loop
            
            newtxt = Mid(txt, 1, X - 1)
            parsestring = Mid(txt, X + 2, i)
            z = Account(Index).Chars(GetPlayerChar(Index)).Variables(Val(parsestring))
            newtxt = newtxt & CStr(z)
            newtxt = newtxt & Mid(txt, X + 2 + i, Len(txt) - (X + i))
            txt = newtxt
        End If
    Loop
    
    ParseEventText = txt
End Function

Function FindEventLabel(ByVal Label As String, MapNum As Long, eventID As Long, PageID As Long, CurSlot As Long, CurList As Long, ListLeftOff() As Long)
    Dim Stack() As Long, StackCount As Long, tmpCurSlot As Long, tmpCurList As Long, CurrentListOption() As Long
    Dim removeEventProcess As Boolean, tmpListLeftOff() As Long, restartlist As Boolean, w As Long

    tmpCurSlot = CurSlot
    tmpCurList = CurList
    tmpListLeftOff = ListLeftOff
    
    ReDim ListLeftOff(Map(MapNum).Events(eventID).Pages(PageID).CommandListCount)
    ReDim CurrentListOption(Map(MapNum).Events(eventID).Pages(PageID).CommandListCount)
    CurList = 1
    CurSlot = 1
    
    Do Until removeEventProcess = True
        If ListLeftOff(CurList) > 0 Then
            CurSlot = ListLeftOff(CurList)
            ListLeftOff(CurList) = 0
        End If
        
        If CurList > Map(MapNum).Events(eventID).Pages(PageID).CommandListCount Then
            ' Get rid of this event, it is bad
            removeEventProcess = True
        End If
        
        If CurSlot > Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).CommandCount Then
            If CurList = 1 Then
                removeEventProcess = True
            Else
                CurList = Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).ParentList
                CurSlot = 1
                restartlist = True
            End If
        End If
        
        If restartlist = False Then
            If removeEventProcess = False Then
                ' If we are still here, then we are good to process shit :D
                Select Case Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).Index
                    Case EventType.evShowChoices
                        If Len(Trim$(Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).Text2)) > 0 Then
                            w = 1
                            If Len(Trim$(Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).Text3)) > 0 Then
                                w = 2
                                If Len(Trim$(Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).Text4)) > 0 Then
                                    w = 3
                                    If Len(Trim$(Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).Text5)) > 0 Then
                                        w = 4
                                    End If
                                End If
                            End If
                        End If
                        
                        If w > 0 Then
                            If CurrentListOption(CurList) < w Then
                                CurrentListOption(CurList) = CurrentListOption(CurList) + 1
                                ' Process
                                ListLeftOff(CurList) = CurSlot
                                Select Case CurrentListOption(CurList)
                                    Case 1
                                        CurList = Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).Data1
                                    Case 2
                                        CurList = Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).Data2
                                    Case 3
                                        CurList = Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).Data3
                                    Case 4
                                        CurList = Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).Data4
                                End Select
                                CurSlot = 0
                            Else
                                CurrentListOption(CurList) = 0
                                ' Continue on
                            End If
                        End If
                        
                        w = 0
                        
                    Case EventType.evCondition
                        If CurrentListOption(CurList) = 0 Then
                            CurrentListOption(CurList) = 1
                            ListLeftOff(CurList) = CurSlot
                            CurList = Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).ConditionalBranch.CommandList
                            CurSlot = 0
                        ElseIf CurrentListOption(CurList) = 1 Then
                            CurrentListOption(CurList) = 2
                            ListLeftOff(CurList) = CurSlot
                            CurList = Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).ConditionalBranch.ElseCommandList
                            CurSlot = 0
                        ElseIf CurrentListOption(CurList) = 2 Then
                            CurrentListOption(CurList) = 0
                        End If
                    Case EventType.evLabel
                        ' Do nothing, just a label
                        If Trim$(Map(MapNum).Events(eventID).Pages(PageID).CommandList(CurList).Commands(CurSlot).Text1) = Trim$(Label) Then
                            Exit Function
                        End If
                End Select
                CurSlot = CurSlot + 1
            End If
        End If
        restartlist = False
    Loop
    
    ListLeftOff = tmpListLeftOff
    CurList = tmpCurList
    CurSlot = tmpCurSlot
End Function
