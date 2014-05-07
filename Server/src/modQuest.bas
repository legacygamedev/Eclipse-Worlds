Attribute VB_Name = "modQuest"
Public Const QUESTICON_LENGTH = 5
Public Const QUESTNAME_LENGTH = 40
Public Const QUESTDESC_LENGTH = 300
Public Quest(1 To MAX_QUESTS) As QuestRec

'Constants to use for tasks
Public Const TASK_KILL As Byte = 1
Public Const TASK_GATHER As Byte = 2
Public Const TASK_MEET As Byte = 3
Public Const TASK_GETSKILL As Byte = 4
Public Const ACTION_GIVE_ITEM As Byte = 5
Public Const ACTION_TAKE_ITEM As Byte = 6
Public Const ACTION_SHOWMSG As Byte = 7
Public Const ACTION_ADJUST_LVL As Byte = 8
Public Const ACTION_ADJUST_EXP As Byte = 9

Private Type RequirementsRec
    AccessReq As Long
    LevelReq As Long
    GenderReq As Long
    ClassReq As Long
    SkillReq As Long
    SkillLevelReq As Long
    Stat_Req(1 To Stats.Stat_count - 1) As Long
End Type

Private Type ActionRec
    TextHolder As String * QUESTDESC_LENGTH
    ActionID As Byte
    Amount As Long
    MainData As Long
    SecondaryData As Long
    TertiaryData As Long
    QuadData As Long
End Type

Private Type CLIRec
    ItemIndex As Long
    isNPC As Long
    Max_Actions As Long
    Action() As ActionRec
End Type

Private Type QuestRec
    Name As String * QUESTNAME_LENGTH
    Description As String * QUESTDESC_LENGTH
    Icon_Start As String * QUESTICON_LENGTH
    Icon_Progress As String * QUESTICON_LENGTH
    CanBeRetaken As Byte
    
    'Maxes
    Max_CLI As Long
    
    'Main data
    CLI() As CLIRec
    Requirements As RequirementsRec
End Type

Private Type TempQuestRec
    CLI() As CLIRec
End Type

Public Type FindQuestRec
    QuestIndex As Long
    CLIIndex As Long
    ActionIndex As Long
End Type

'/////////////////////////////////////////////////////////
'/////////////////QUEST SUBS AND FUNCTIONS////////////////
'/////////////////////////////////////////////////////////

Function GetPlayerQuestCLIID(ByVal Index As Long, ByVal QuestID As Long)
    If Index < 1 Or Index > Player_HighIndex Then Exit Function
    GetPlayerQuestCLIID = Account(Index).Chars(GetPlayerChar(Index)).QuestCLIID(QuestID)
End Function

Function GetPlayerQuestTaskID(ByVal Index As Long, ByVal QuestID As Long)
    GetPlayerQuestTaskID = Account(Index).Chars(GetPlayerChar(Index)).QuestTaskID(QuestID)
End Function

Function GetPlayerQuestAmount(ByVal Index As Long, ByVal QuestID As Long)
    GetPlayerQuestAmount = Account(Index).Chars(GetPlayerChar(Index)).QuestAmount(QuestID)
End Function

Sub SetPlayerQuestCLIID(ByVal Index As Long, ByVal QuestID As Long, Value As Long)
    Account(Index).Chars(GetPlayerChar(Index)).QuestCLIID(QuestID) = Value
End Sub

Sub SetPlayerQuestTaskID(ByVal Index As Long, ByVal QuestID As Long, Value As Long)
    Account(Index).Chars(GetPlayerChar(Index)).QuestTaskID(QuestID) = Value
End Sub

Sub SetPlayerQuestAmount(ByVal Index As Long, ByVal QuestID As Long, Value As Long, Optional ByVal PlusVal As Boolean = False)
    If PlusVal Then
        Account(Index).Chars(GetPlayerChar(Index)).QuestAmount(QuestID) = Account(Index).Chars(GetPlayerChar(Index)).QuestAmount(QuestID) + Value
    Else
        Account(Index).Chars(GetPlayerChar(Index)).QuestAmount(QuestID) = Value
    End If
End Sub

Public Function IsQuestCLI(ByVal Index As Long, ByVal NPCIndex As Long) As FindQuestRec
Dim I As Long, II As Long, III As Long
Dim temp As FindQuestRec

    'Dynamically find the correct quest item
    For I = 1 To MAX_QUESTS
        With Quest(I)
            For II = 1 To .Max_CLI
                'See if this npc is within a started quest first.
                If .CLI(II).ItemIndex = NPCIndex Then 'found a matching quest cli item, this npc is part of a quest
                    If IsInQuest(Index, I) Then
                        temp.QuestIndex = I
                        temp.CLIIndex = II
                        IsQuestCLI = temp
                        Exit Function
                    End If
                End If
            Next II
        End With
    Next I
    
    For I = 1 To MAX_QUESTS
        With Quest(I)
            For II = 1 To .Max_CLI
                'It's not within a started quest, so see if it's a start to a new quest
                If .CLI(II).ItemIndex = NPCIndex Then 'found a matching quest cli item, this npc is part of a quest
                    If II = 1 Then
                        temp.QuestIndex = I
                        temp.CLIIndex = II
                        IsQuestCLI = temp
                        Exit Function
                    End If
                End If
            Next II
        End With
    Next I
End Function

Public Sub CheckQuest(ByVal Index As Long, QuestIndex As Long, CLIIndex As Long, TaskIndex As Long)
Dim I As Long
    'Is the player on this quest?  If not, cancel out.
    If IsInQuest(Index, QuestIndex) Then
        'Is the player currently on this Chronological list item?
        If GetPlayerQuestCLIID(Index, QuestIndex) = CLIIndex Then
            Call HandleQuestTask(Index, QuestIndex, CLIIndex, GetPlayerQuestTaskID(Index, QuestIndex))
        Else
            'Dynamically show message from last known cli
            If GetPlayerQuestCLIID(Index, QuestIndex) - 1 > 0 Then
                For I = Quest(QuestIndex).CLI(GetPlayerQuestCLIID(Index, QuestIndex) - 1).Max_Actions To 1 Step -1
                    With Quest(QuestIndex).CLI(GetPlayerQuestCLIID(Index, QuestIndex) - 1).Action(I)
                        'quit early if we run into a task.  Means we don't have a msg to display
                        If .ActionID > 0 And .ActionID < 4 Then Exit For
                        
                        If .ActionID = ACTION_SHOWMSG Then
                            Call PlayerMsg(Index, Trim$(.TextHolder), White, True, QuestIndex)
                            Exit For
                        End If
                    End With
                Next I
            End If
        End If
    Else
        'lets start this quest if the CLI is the first greeter
        If CLIIndex = 1 Then
            ' see if the player has taken it all ready and if it can be retaken
            If IsQuestCompleted(Index, QuestIndex) Then
                If Quest(QuestIndex).CanBeRetaken = False Then
                    Call PlayerMsg(Index, "This quest cannot be retaken.", Green, True, QuestIndex)
                    Exit Sub
                End If
            End If
        
            'not in a quest, check the requirements
            With Quest(QuestIndex).Requirements
                'check level
                If .LevelReq > 0 Then
                    If Not GetPlayerLevel(Index) >= .LevelReq Then
                        Call PlayerMsg(Index, "Your level does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                        Exit Sub
                    End If
                End If
                'check class
                If .ClassReq > 0 Then
                    If Not GetPlayerClass(Index) = .ClassReq Then
                        Call PlayerMsg(Index, "Your class does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                        Exit Sub
                    End If
                End If
                'check gender
                If .GenderReq > 0 Then
                    If Not GetPlayerGender(Index) = .GenderReq Then
                        Call PlayerMsg(Index, "Your gender does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                        Exit Sub
                    End If
                End If
                'check access
                If Not GetPlayerAccess(Index) >= .AccessReq Then
                    Call PlayerMsg(Index, "Your administrative access level does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                    Exit Sub
                End If
                'check skill
                If .SkillReq > 0 Then
                    If Not GetPlayerSkill(Index, .SkillReq) >= .SkillLevelReq Then
                        Call PlayerMsg(Index, "Your " & GetSkillName(.SkillLevelReq) & " level does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                        Exit Sub
                    End If
                End If
                'check stats
                For I = 1 To Stats.Stat_count - 1
                    If Not GetPlayerStat(Index, I) >= .Stat_Req(I) Then
                        Call PlayerMsg(Index, "Your stats do not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                        Exit Sub
                    End If
                Next I
            End With
            
            'send the request to the player
            Call SendPlayerQuestRequest(Index, QuestIndex)
        End If
    End If
End Sub

Public Sub HandleQuestTask(ByVal Index As Long, ByVal QuestID As Long, ByVal CLIID As Long, ByVal TaskID As Long, Optional ByVal ShowRebuttal As Boolean = True)
Dim I As Long, tmp As Long
    'Manage the current task the player is on and move player forward through the quest.
    If QuestID < 1 Or QuestID > MAX_QUESTS Then Exit Sub
    If CLIID < 1 Then Exit Sub
    
    With Quest(QuestID).CLI(CLIID)
        ' Figure out what we need to do.
        Select Case .Action(TaskID).ActionID
            Case TASK_GATHER
                'check if the player gathered enough of the item
                If HasItem(Index, .Action(TaskID).MainData) >= .Action(TaskID).Amount Then
                    'player has the right amount.  move forward.
                    Call SetPlayerQuestAmount(Index, QuestID, 0)
                    Call CheckNextTask(Index, QuestID, CLIID, TaskID)
                    
                    If .Action(TaskID).SecondaryData = vbChecked Then 'take the item
                        Call TakeInvItem(Index, .Action(TaskID).MainData, .Action(TaskID).Amount, True)
                    End If
                Else
                    'we don't have the required amount, see if we need to say a rebuttal msg
                    If ShowRebuttal Then Call CheckRebuttal(Index, QuestID, CLIID, TaskID)
                End If
                Exit Sub
                
            Case TASK_KILL
                'check if the player killed enough
                If GetPlayerQuestAmount(Index, QuestID) >= .Action(TaskID).Amount Then
                    'player has the right amount.  move forward.
                    Call SetPlayerQuestAmount(Index, QuestID, 0)
                    Call CheckNextTask(Index, QuestID, CLIID, TaskID)
                Else
                    'we don't have the required amount, see if we need to say a rebuttal msg
                    If ShowRebuttal Then Call CheckRebuttal(Index, QuestID, CLIID, TaskID)
                End If
                
            Case TASK_GETSKILL
                'check if the player gained the right skill amount
                If GetPlayerSkill(Index, .Action(TaskID).MainData) >= .Action(TaskID).Amount Then
                    Call SetPlayerQuestAmount(Index, QuestID, 0)
                    Call CheckNextTask(Index, QuestID, CLIID, TaskID)
                Else
                    'we don't have the required amount, see if we need to say a rebuttal msg
                    If ShowRebuttal Then Call CheckRebuttal(Index, QuestID, CLIID, TaskID)
                End If
                
            Case ACTION_GIVE_ITEM
                'give the player so many of a certain item
                If Item(.Action(TaskID).MainData).Stackable > 0 Then
                    tmp = GiveInvItem(Index, .Action(TaskID).MainData, .Action(TaskID).Amount, , , True)
                Else
                    For I = 1 To .Action(TaskID).Amount
                        tmp = GiveInvItem(Index, .Action(TaskID).MainData, 1, , , True)
                    Next I
                End If
                
                If tmp < 1 Or tmp > MAX_INV Then
                    Call PlayerMsg(Index, "Not enough space in your inventory.  Please come back when you can hold everything I have to give you.", BrightRed, True, QuestID)
                    Exit Sub
                End If
                Call CheckNextTask(Index, QuestID, CLIID, TaskID)
                
            Case ACTION_TAKE_ITEM
                'take the player's item
                Call TakeInvItem(Index, .Action(TaskID).MainData, .Action(TaskID).Amount, True)
                Call CheckNextTask(Index, QuestID, CLIID, TaskID)
                
            Case ACTION_SHOWMSG
                'show the player a message
                Call PlayerMsg(Index, ModifyTxt(Index, QuestID, Trim$(.Action(TaskID).TextHolder)), .Action(TaskID).TertiaryData, True, QuestID)
                Call CheckNextTask(Index, QuestID, CLIID, TaskID)
                
            Case ACTION_ADJUST_LVL
                Call SetPlayerLevel(Index, .Action(TaskID).Amount, True)
                Call SendPlayerLevel(Index)
                Call CheckNextTask(Index, QuestID, CLIID, TaskID)
            
            Case ACTION_ADJUST_EXP
                Call SetPlayerExp(Index, .Action(TaskID).Amount, True)
                Call SendPlayerExp(Index)
                Call CheckNextTask(Index, QuestID, CLIID, TaskID)
                
        End Select
    End With
End Sub

Public Sub CheckNextTask(ByVal Index As Long, QuestID As Long, CLIID As Long, TaskID As Long)
    With Quest(QuestID).CLI(CLIID)
        ' move on to next task if there is one
        If TaskID + 1 > .Max_Actions Then GoTo NextCLI
        
        'check if the next task is a rebuttal, if so, skip it
        If .Action(TaskID + 1).ActionID = ACTION_SHOWMSG Then
            If .Action(TaskID + 1).SecondaryData = vbChecked Then
                If Not TaskID + 2 > .Max_Actions Then
                    'skip this rebuttal task
                    Call SetPlayerQuestTaskID(Index, QuestID, TaskID + 2)
                Else
                    GoTo NextCLI
                End If
            Else
                Call SetPlayerQuestTaskID(Index, QuestID, TaskID + 1)
            End If
        Else
            Call SetPlayerQuestTaskID(Index, QuestID, TaskID + 1)
        End If
        
        Call SendPlayerData(Index)
        Call HandleQuestTask(Index, QuestID, GetPlayerQuestCLIID(Index, QuestID), GetPlayerQuestTaskID(Index, QuestID), False)
        Exit Sub
        
NextCLI:
        'move on to next cli item if there is one
        If Not CLIID + 1 > Quest(QuestID).Max_CLI Then
            Call SetPlayerQuestCLIID(Index, QuestID, CLIID + 1)
            Call SetPlayerQuestTaskID(Index, QuestID, 1)
            Call SendPlayerData(Index)
            'We don't want to move straight for the next task here.  The player has to talk to them to start it.
        Else
            'quest completed
            Call MarkQuestCompleted(Index, QuestID)
            Call SetPlayerQuestCLIID(Index, QuestID, 0)
            Call SetPlayerQuestTaskID(Index, QuestID, 0)
            Call SetPlayerQuestAmount(Index, QuestID, 0)
            Call PlayerMsg(Index, "Congratulations!  You completed the " & Trim$(Quest(QuestID).Name) & " quest.", BrightGreen, True, QuestID)
            Call SendPlayerData(Index)
        End If
    End With
End Sub

Public Sub CheckRebuttal(ByVal Index As Long, QuestID As Long, CLIID As Long, TaskID As Long)
Dim I As Long
    With Quest(QuestID).CLI(CLIID)
        For I = TaskID To .Max_Actions
            If .Action(I).ActionID = ACTION_SHOWMSG Then
                If .Action(I).SecondaryData = vbChecked Then
                    'send the msg
                    Call PlayerMsg(Index, ModifyTxt(Index, QuestID, Trim$(.Action(I).TextHolder)), .Action(I).TertiaryData, True, QuestID)
                    Exit Sub
                End If
            End If
        Next I
    End With
End Sub

Public Function ModifyTxt(ByVal Index, ByVal QuestID As Long, ByVal Msg As String) As String
Dim nMsg As String
    nMsg = Replace$(Msg, "#kills#", GetPlayerQuestAmount(Index, QuestID)) 'replace with player kill amount
    ModifyTxt = nMsg
End Function

Public Sub CheckIfQuestKill(ByVal Index As Long, ByVal NPCNum As Long)
Dim I As Long, II As Long, III As Long
Dim Kills As Long, Needed As Long
    'Cycle through all the quests the player could be in
    For I = 1 To MAX_QUESTS
        II = GetPlayerQuestCLIID(Index, I)
        III = GetPlayerQuestTaskID(Index, I)
        If II > 0 Then
            If III > 0 Then
                'Make sure the player's current task for this quest is to kill enemies
                If Quest(I).CLI(II).Action(III).ActionID = TASK_KILL Then
                    'Make sure this is the NPC we're supposed to kill for this quest
                    If Quest(I).CLI(II).Action(III).MainData = NPCNum Then
                
                        Call SetPlayerQuestAmount(Index, I, 1, True)
                        Kills = GetPlayerQuestAmount(Index, I)
                        Needed = Quest(I).CLI(II).Action(III).Amount
                        
                        If Not Kills >= Needed Then
                            Call PlayerMsg(Index, "Quest Kills: " & Kills & " / " & Needed, White)
                            Call SendActionMsg(GetPlayerMap(Index), Kills & "/" & Needed & " kills", Green, 1, (GetPlayerX(Index) * 32), (GetPlayerY(Index) * 37))
                        Else
                            Call PlayerMsg(Index, "Quest Task Completed!  Kills: " & GetPlayerQuestAmount(Index, I) & " / " & Quest(I).CLI(II).Action(III).Amount & _
                            "  Go back and speak with " & Trim$(NPC(Quest(I).CLI(II).ItemIndex).Name) & " to continue.", BrightGreen, True, I)
                        End If
                    End If
                End If
            End If
        End If
    Next I
End Sub

Public Function IsQuestCompleted(ByVal Index As Long, ByVal QuestID As Long) As Boolean
Dim I As Long
    IsQuestCompleted = False
    If Not QuestID > 0 Then Exit Function
    
    If Account(Index).Chars(GetPlayerChar(Index)).QuestCompleted(QuestID) = True Then
        IsQuestCompleted = True
    End If
End Function

Public Sub MarkQuestCompleted(ByVal Index As Long, ByVal QuestID As Long)
Dim I As Long
    If Not QuestID > 0 Then Exit Sub
    
    Account(Index).Chars(GetPlayerChar(Index)).QuestCompleted(QuestID) = True
End Sub

Private Function IsInQuest(ByVal Index As Long, ByVal QuestID As Long) As Boolean
    If Not QuestID > 0 Then Exit Function
    
    If GetPlayerQuestCLIID(Index, QuestID) > 0 Then IsInQuest = True
End Function

Private Sub SendPlayerQuestRequest(ByVal Index As Long, ByVal QuestID As Long)
Dim buffer As clsBuffer

    If Index < 1 Or Index > Player_HighIndex Then Exit Sub
    If QuestID < 1 Or QuestID > MAX_QUESTS Then Exit Sub
    
    Set buffer = New clsBuffer
        buffer.WriteLong SQuestRequest
        buffer.WriteLong QuestID
        Call SendDataTo(Index, buffer.ToArray())
    Set buffer = Nothing
End Sub


