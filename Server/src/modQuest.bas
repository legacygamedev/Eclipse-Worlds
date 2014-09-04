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
Public Const ACTION_WARP As Byte = 10
Public Const ACTION_ADJUST_STAT_LVL As Byte = 11
Public Const ACTION_ADJUST_SKILL_LVL As Byte = 12
Public Const ACTION_ADJUST_SKILL_EXP As Byte = 13
Public Const ACTION_ADJUST_STAT_POINTS As Byte = 14

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

Function GetPlayerQuestCLIID(ByVal index As Long, ByVal QuestID As Long)
    If index < 1 Or index > Player_HighIndex Then Exit Function
    GetPlayerQuestCLIID = Account(index).Chars(GetPlayerChar(index)).QuestCLIID(QuestID)
End Function

Function GetPlayerQuestTaskID(ByVal index As Long, ByVal QuestID As Long)
    GetPlayerQuestTaskID = Account(index).Chars(GetPlayerChar(index)).QuestTaskID(QuestID)
End Function

Function GetPlayerQuestAmount(ByVal index As Long, ByVal QuestID As Long)
    GetPlayerQuestAmount = Account(index).Chars(GetPlayerChar(index)).QuestAmount(QuestID)
End Function

Sub SetPlayerQuestCLIID(ByVal index As Long, ByVal QuestID As Long, Value As Long)
    Account(index).Chars(GetPlayerChar(index)).QuestCLIID(QuestID) = Value
End Sub

Sub SetPlayerQuestTaskID(ByVal index As Long, ByVal QuestID As Long, Value As Long)
    Account(index).Chars(GetPlayerChar(index)).QuestTaskID(QuestID) = Value
End Sub

Sub SetPlayerQuestAmount(ByVal index As Long, ByVal QuestID As Long, Value As Long, Optional ByVal PlusVal As Boolean = False)
    If PlusVal Then
        Account(index).Chars(GetPlayerChar(index)).QuestAmount(QuestID) = Account(index).Chars(GetPlayerChar(index)).QuestAmount(QuestID) + Value
    Else
        Account(index).Chars(GetPlayerChar(index)).QuestAmount(QuestID) = Value
    End If
End Sub

Public Function IsQuestCLI(ByVal index As Long, ByVal NPCIndex As Long) As FindQuestRec
Dim i As Long, II As Long, III As Long
Dim temp As FindQuestRec

    'Dynamically find the correct quest item
    For i = 1 To MAX_QUESTS
        With Quest(i)
            For II = 1 To .Max_CLI
                'See if this npc is within a started quest first.
                If .CLI(II).ItemIndex = NPCIndex Then 'found a matching quest cli item, this npc is part of a quest
                    If IsInQuest(index, i) Then
                        temp.QuestIndex = i
                        temp.CLIIndex = II
                        IsQuestCLI = temp
                        Exit Function
                    End If
                End If
            Next II
        End With
    Next i
    
    For i = 1 To MAX_QUESTS
        With Quest(i)
            For II = 1 To .Max_CLI
                'It's not within a started quest, so see if it's a start to a new quest
                If .CLI(II).ItemIndex = NPCIndex Then 'found a matching quest cli item, this npc is part of a quest
                    If II = 1 Then
                        temp.QuestIndex = i
                        temp.CLIIndex = II
                        IsQuestCLI = temp
                        Exit Function
                    End If
                End If
            Next II
        End With
    Next i
End Function

Public Sub CheckQuest(ByVal index As Long, QuestIndex As Long, CLIIndex As Long, TaskIndex As Long)
Dim i As Long
    'Is the player on this quest?  If not, cancel out.
    If IsInQuest(index, QuestIndex) Then
        'Is the player currently on this Chronological list item?
        If GetPlayerQuestCLIID(index, QuestIndex) = CLIIndex Then
            Call HandleQuestTask(index, QuestIndex, CLIIndex, GetPlayerQuestTaskID(index, QuestIndex))
        Else
            'Dynamically show message from last known cli
            If GetPlayerQuestCLIID(index, QuestIndex) - 1 > 0 Then
                For i = Quest(QuestIndex).CLI(GetPlayerQuestCLIID(index, QuestIndex) - 1).Max_Actions To 1 Step -1
                    With Quest(QuestIndex).CLI(GetPlayerQuestCLIID(index, QuestIndex) - 1).Action(i)
                        'quit early if we run into a task.  Means we don't have a msg to display
                        If .ActionID > 0 And .ActionID < 4 Then Exit For
                        
                        If .ActionID = ACTION_SHOWMSG Then
                            Call PlayerMsg(index, Trim$(.TextHolder), White, True, QuestIndex)
                            Exit For
                        End If
                    End With
                Next i
            End If
        End If
    Else
        'lets start this quest if the CLI is the first greeter
        If CLIIndex = 1 Then
            ' see if the player has taken it all ready and if it can be retaken
            If IsQuestCompleted(index, QuestIndex) Then
                If Quest(QuestIndex).CanBeRetaken = False Then
                    Call PlayerMsg(index, "This quest cannot be retaken.", Green, True, QuestIndex)
                    Exit Sub
                End If
            End If
        
            'not in a quest, check the requirements
            With Quest(QuestIndex).Requirements
                'check level
                If .LevelReq > 0 Then
                    If Not GetPlayerLevel(index) >= .LevelReq Then
                        Call PlayerMsg(index, "Your level does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                        Exit Sub
                    End If
                End If
                'check class
                If .ClassReq > 0 Then
                    If Not GetPlayerClass(index) = .ClassReq Then
                        Call PlayerMsg(index, "Your class does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                        Exit Sub
                    End If
                End If
                'check gender
                If .GenderReq > 0 Then
                    If Not GetPlayerGender(index) = .GenderReq Then
                        Call PlayerMsg(index, "Your gender does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                        Exit Sub
                    End If
                End If
                'check access
                If Not GetPlayerAccess(index) >= .AccessReq Then
                    Call PlayerMsg(index, "Your administrative access level does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                    Exit Sub
                End If
                'check skill
                If .SkillReq > 0 Then
                    If Not GetPlayerSkill(index, .SkillReq) >= .SkillLevelReq Then
                        Call PlayerMsg(index, "Your " & GetSkillName(.SkillLevelReq) & " level does not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                        Exit Sub
                    End If
                End If
                'check stats
                For i = 1 To Stats.Stat_count - 1
                    If Not GetPlayerStat(index, i) >= .Stat_Req(i) Then
                        Call PlayerMsg(index, "Your stats do not meet the requirements to start this quest.", BrightRed, True, QuestIndex)
                        Exit Sub
                    End If
                Next i
            End With
            
            'send the request to the player
            Call SendPlayerQuestRequest(index, QuestIndex)
        End If
    End If
End Sub

Public Sub HandleQuestTask(ByVal index As Long, ByVal QuestID As Long, ByVal CLIID As Long, ByVal TaskID As Long, Optional ByVal ShowRebuttal As Boolean = True)
Dim i As Long, tmp As Long
Dim NPCNum As Long
    'Manage the current task the player is on and move player forward through the quest.
    If QuestID < 1 Or QuestID > MAX_QUESTS Then Exit Sub
    If CLIID < 1 Then Exit Sub
    
    With Quest(QuestID).CLI(CLIID)
        NPCNum = .ItemIndex
        ' Figure out what we need to do.
        Select Case .Action(TaskID).ActionID
            Case TASK_GATHER
                'check if the player gathered enough of the item
                If HasItem(index, .Action(TaskID).MainData) >= .Action(TaskID).Amount Then
                    'player has the right amount.  move forward.
                    Call SetPlayerQuestAmount(index, QuestID, 0)
                    Call CheckNextTask(index, QuestID, CLIID, TaskID)
                    
                    If .Action(TaskID).SecondaryData = vbChecked Then 'take the item
                        Call TakeInvItem(index, .Action(TaskID).MainData, .Action(TaskID).Amount, True)
                    End If
                Else
                    'we don't have the required amount, see if we need to say a rebuttal msg
                    If ShowRebuttal Then Call CheckRebuttal(index, QuestID, CLIID, TaskID)
                End If
                Exit Sub
                
            Case TASK_KILL
                'check if the player killed enough
                If GetPlayerQuestAmount(index, QuestID) >= .Action(TaskID).Amount Then
                    'player has the right amount.  move forward.
                    Call SetPlayerQuestAmount(index, QuestID, 0)
                    Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Else
                    'we don't have the required amount, see if we need to say a rebuttal msg
                    If ShowRebuttal Then Call CheckRebuttal(index, QuestID, CLIID, TaskID)
                End If
                
            Case TASK_GETSKILL
                'check if the player gained the right skill amount
                If GetPlayerSkill(index, .Action(TaskID).MainData) >= .Action(TaskID).Amount Then
                    Call SetPlayerQuestAmount(index, QuestID, 0)
                    Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Else
                    'we don't have the required amount, see if we need to say a rebuttal msg
                    If ShowRebuttal Then Call CheckRebuttal(index, QuestID, CLIID, TaskID)
                End If
                
            Case ACTION_GIVE_ITEM
                'give the player so many of a certain item
                If Item(.Action(TaskID).MainData).Stackable > 0 Then
                    tmp = GiveInvItem(index, .Action(TaskID).MainData, .Action(TaskID).Amount, , , True)
                Else
                    For i = 1 To .Action(TaskID).Amount
                        tmp = GiveInvItem(index, .Action(TaskID).MainData, 1, , , True)
                    Next i
                End If
                
                If tmp < 1 Or tmp > MAX_INV Then
                    Call PlayerMsg(index, "Not enough space in your inventory.  Please come back when you can hold everything I have to give you.", BrightRed, True, QuestID)
                    Exit Sub
                End If
                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
                
            Case ACTION_TAKE_ITEM
                'take the player's item
                Call TakeInvItem(index, .Action(TaskID).MainData, .Action(TaskID).Amount, True)
                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
                
            Case ACTION_SHOWMSG
                'show the player a message
                Call PlayerMsg(index, ModifyTxt(index, QuestID, Trim$(.Action(TaskID).TextHolder)), .Action(TaskID).TertiaryData, True, QuestID)
                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
                
            Case ACTION_ADJUST_LVL
                Call SetPlayerLevel(index, .Action(TaskID).Amount, .Action(TaskID).MainData)
                Call SendPlayerLevel(index)
                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
            
            Case ACTION_ADJUST_EXP
                Call SetPlayerExp(index, .Action(TaskID).Amount, .Action(TaskID).MainData)
                Call SendPlayerExp(index)
                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
            
            Case ACTION_ADJUST_STAT_LVL
                Call SetPlayerStat(index, .Action(TaskID).SecondaryData, .Action(TaskID).Amount, .Action(TaskID).MainData)
                Call SendPlayerStats(index)
                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
                
            Case ACTION_ADJUST_SKILL_LVL
                Call SetPlayerSkill(index, .Action(TaskID).Amount, .Action(TaskID).SecondaryData, .Action(TaskID).MainData)
                Call SendPlayerData(index)
                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
            
            Case ACTION_ADJUST_SKILL_LVL
                Call SetPlayerSkill(index, .Action(TaskID).Amount, .Action(TaskID).SecondaryData, .Action(TaskID).MainData)
                Call SendPlayerData(index)
                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
                
            Case ACTION_ADJUST_SKILL_EXP
                Call SetPlayerSkillExp(index, .Action(TaskID).Amount, .Action(TaskID).SecondaryData, .Action(TaskID).MainData)
                Call SendPlayerPoints(index)
                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
            
            Case ACTION_WARP
                Call PlayerWarp(index, .Action(TaskID).Amount, .Action(TaskID).MainData, .Action(TaskID).SecondaryData, , DIR_DOWN)
                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
                
            Case Else
                'continue on in case we missed something.  This will make it harder to find bugs, but will run smoother for the user
                Call CheckNextTask(index, QuestID, CLIID, TaskID)
                Call SendShowTaskCompleteOnNPC(index, NPCNum, False)
        End Select
    End With
End Sub

Public Sub CheckNextTask(ByVal index As Long, QuestID As Long, CLIID As Long, TaskID As Long)
    With Quest(QuestID).CLI(CLIID)
        ' move on to next task if there is one
        If TaskID + 1 > .Max_Actions Then GoTo NextCLI
        
        'check if the next task is a rebuttal, if so, skip it
        If .Action(TaskID + 1).ActionID = ACTION_SHOWMSG Then
            If .Action(TaskID + 1).SecondaryData = vbChecked Then
                If Not TaskID + 2 > .Max_Actions Then
                    'skip this rebuttal task
                    Call SetPlayerQuestTaskID(index, QuestID, TaskID + 2)
                Else
                    GoTo NextCLI
                End If
            Else
                Call SetPlayerQuestTaskID(index, QuestID, TaskID + 1)
            End If
        Else
            Call SetPlayerQuestTaskID(index, QuestID, TaskID + 1)
        End If
        
        Call SendPlayerData(index)
        Call HandleQuestTask(index, QuestID, GetPlayerQuestCLIID(index, QuestID), GetPlayerQuestTaskID(index, QuestID), False)
        Exit Sub
        
NextCLI:
        'move on to next cli item if there is one
        If Not CLIID + 1 > Quest(QuestID).Max_CLI Then
            Call SetPlayerQuestCLIID(index, QuestID, CLIID + 1)
            Call SetPlayerQuestTaskID(index, QuestID, 1)
            Call SendPlayerData(index)
            'We don't want to move straight for the next task here.  The player has to talk to them to start it.
        Else
            'quest completed
            Call MarkQuestCompleted(index, QuestID)
            Call SetPlayerQuestCLIID(index, QuestID, 0)
            Call SetPlayerQuestTaskID(index, QuestID, 0)
            Call SetPlayerQuestAmount(index, QuestID, 0)
            Call PlayerMsg(index, "Congratulations!  You completed the " & Trim$(Quest(QuestID).Name) & " quest.", BrightGreen, True, QuestID)
            Call SendPlayerData(index)
        End If
    End With
End Sub

Public Sub CheckRebuttal(ByVal index As Long, QuestID As Long, CLIID As Long, TaskID As Long)
Dim i As Long
    With Quest(QuestID).CLI(CLIID)
        For i = TaskID To .Max_Actions
            If .Action(i).ActionID = ACTION_SHOWMSG Then
                If .Action(i).SecondaryData = vbChecked Then
                    'send the msg
                    Call PlayerMsg(index, ModifyTxt(index, QuestID, Trim$(.Action(i).TextHolder)), .Action(i).TertiaryData, True, QuestID)
                    Exit Sub
                End If
            End If
        Next i
    End With
End Sub

Public Function ModifyTxt(ByVal index, ByVal QuestID As Long, ByVal Msg As String) As String
Dim nMsg As String
    nMsg = Replace$(Msg, "#kills#", GetPlayerQuestAmount(index, QuestID)) 'replace with player kill amount
    ModifyTxt = nMsg
End Function

Public Sub CheckIfQuestKill(ByVal index As Long, ByVal NPCNum As Long)
Dim i As Long, II As Long, III As Long
Dim Kills As Long, Needed As Long
    'Cycle through all the quests the player could be in
    For i = 1 To MAX_QUESTS
        II = GetPlayerQuestCLIID(index, i)
        III = GetPlayerQuestTaskID(index, i)
        If II > 0 Then
            If III > 0 Then
                'Make sure the player's current task for this quest is to kill enemies
                If Quest(i).CLI(II).Action(III).ActionID = TASK_KILL Then
                    'Make sure this is the NPC we're supposed to kill for this quest
                    If Quest(i).CLI(II).Action(III).MainData = NPCNum Then
                
                        Call SetPlayerQuestAmount(index, i, 1, True)
                        Kills = GetPlayerQuestAmount(index, i)
                        Needed = Quest(i).CLI(II).Action(III).Amount
                        
                        If Not Kills >= Needed Then
                            Call PlayerMsg(index, "Quest Kills: " & Kills & " / " & Needed, White)
                            Call SendActionMsg(GetPlayerMap(index), Kills & "/" & Needed & " kills", Green, 1, (GetPlayerX(index) * 32), (GetPlayerY(index) * 37))
                        Else
                            Call PlayerMsg(index, "Quest Task Completed!  Kills: " & GetPlayerQuestAmount(index, i) & " / " & Quest(i).CLI(II).Action(III).Amount & _
                            "  Go back and speak with " & Trim$(NPC(Quest(i).CLI(II).ItemIndex).Name) & " to continue.", BrightGreen, True, i)
                            Call SendShowTaskCompleteOnNPC(index, Quest(i).CLI(II).ItemIndex, True)
                        End If
                    End If
                End If
            End If
        End If
    Next i
End Sub

Public Function IsQuestCompleted(ByVal index As Long, ByVal QuestID As Long) As Boolean
Dim i As Long
    IsQuestCompleted = False
    If Not QuestID > 0 Then Exit Function
    
    If Account(index).Chars(GetPlayerChar(index)).QuestCompleted(QuestID) = True Then
        IsQuestCompleted = True
    End If
End Function

Public Sub MarkQuestCompleted(ByVal index As Long, ByVal QuestID As Long)
Dim i As Long
    If Not QuestID > 0 Then Exit Sub
    
    Account(index).Chars(GetPlayerChar(index)).QuestCompleted(QuestID) = True
End Sub

Private Function IsInQuest(ByVal index As Long, ByVal QuestID As Long) As Boolean
    If Not QuestID > 0 Then Exit Function
    
    If GetPlayerQuestCLIID(index, QuestID) > 0 Then IsInQuest = True
End Function

Private Sub SendPlayerQuestRequest(ByVal index As Long, ByVal QuestID As Long)
Dim Buffer As clsBuffer

    If index < 1 Or index > Player_HighIndex Then Exit Sub
    If QuestID < 1 Or QuestID > MAX_QUESTS Then Exit Sub
    
    Set Buffer = New clsBuffer
        Buffer.WriteLong SQuestRequest
        Buffer.WriteLong QuestID
        Call SendDataTo(index, Buffer.ToArray())
    Set Buffer = Nothing
End Sub

Public Function HasQuestItems(ByVal index As Long, QuestID As Long, Optional ByVal ReturnIfNot As Boolean = False) As String
Dim i As Long, CLIIndex As Long, TaskIndex As Long
    CLIIndex = GetPlayerQuestCLIID(index, QuestID)
    TaskIndex = GetPlayerQuestTaskID(index, QuestID)
    
    HasQuestItems = 0
    
    If CLIIndex > 0 Then
        If TaskIndex > 0 Then
            For i = TaskIndex To 1 Step -1
                If Quest(QuestID).CLI(CLIIndex).Action(i).ActionID = TASK_GATHER Then
                    If HasItem(index, Quest(QuestID).CLI(CLIIndex).Action(i).MainData) >= Quest(QuestID).CLI(CLIIndex).Action(i).Amount Then
                        HasQuestItems = Quest(QuestID).CLI(CLIIndex).ItemIndex 'return the npc number
                        Exit Function
                    Else
                        If ReturnIfNot Then
                            'return a value meant to be parsed so we can distinguish the returned value
                            HasQuestItems = Quest(QuestID).CLI(CLIIndex).ItemIndex & "|" & "Can't be empty... lol"
                            Exit Function
                        End If
                    End If
                End If
            Next i
        End If
    End If
End Function

Public Function HasQuestSkill(ByVal index As Long, QuestID As Long, Optional ByVal ReturnIfNot As Boolean = False) As Long
Dim i As Long, CLIIndex As Long, TaskIndex As Long
    CLIIndex = GetPlayerQuestCLIID(index, QuestID)
    TaskIndex = GetPlayerQuestTaskID(index, QuestID)
    
    HasQuestSkill = 0
    
    If CLIIndex > 0 Then
        If TaskIndex > 0 Then
            For i = TaskIndex To 1 Step -1
                If Quest(QuestID).CLI(CLIIndex).Action(i).ActionID = TASK_GETSKILL Then
                    If GetPlayerSkill(index, Quest(QuestID).CLI(CLIIndex).Action(i).MainData) >= Quest(QuestID).CLI(CLIIndex).Action(i).Amount Then
                        HasQuestSkill = Quest(QuestID).CLI(CLIIndex).ItemIndex 'return the npc number
                        Exit Function
                    Else
                        If ReturnIfNot Then
                            'return a value meant to be parsed so we can distinguish the returned value
                            HasQuestSkill = Quest(QuestID).CLI(CLIIndex).ItemIndex & "|" & "Can't be empty... lol"
                            Exit Function
                        End If
                    End If
                End If
            Next i
        End If
    End If
End Function

Public Sub SendRefreshCharEditor(ByVal index As Long)
Dim Buffer As clsBuffer

    Set Buffer = New clsBuffer
    Buffer.WriteLong SRefreshCharEditor
    Call SendDataTo(index, Buffer.ToArray())
    Set Buffer = Nothing
End Sub

