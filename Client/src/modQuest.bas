Attribute VB_Name = "modQuest"
Public Const QUESTICON_LENGTH = 5
Public Const QUESTNAME_LENGTH = 40
Public Const QUESTDESC_LENGTH = 300
Public Quest(1 To MAX_QUESTS) As QuestRec
Public QuestRequest As Long

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

Public Const Quest_Icon_Start As Byte = 1
Public Const Quest_Icon_Finished As Byte = 2
Public Const Quest_Icon_Progress As Byte = 3
Public Const Quest_Icon_Completed As Byte = 4

'Constants for list mover
Public Const LIST_CLI As Byte = 1
Public Const LIST_TASK As Byte = 2

'Variable for icon drawing - client only

Private Type RequirementsRec
    AccessReq As Long
    LevelReq As Long
    GenderReq As Long
    ClassReq As Long
    SkillReq As Long
    SkillLevelReq As Long
    Stat_Req(1 To Stats.Stat_Count - 1) As Long
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

'////////////////////////////////////////
'////////QUEST SUBS AND FUNCTIONS////////
'////////////////////////////////////////
Public Sub MoveListItem(ByVal ListID As Byte, ByVal Index As Long, ByVal CLIIndex As Long, ByVal ArrayID As Long, ByVal Dir As Integer)
Dim I As Long
Dim TempQuest As TempQuestRec

    'If the dir is -1, it means we're backtracking a slot to move the item up
    'If the dir is 1, it means we're moving ahead a slot to move the item down
    
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Select Case ListID
        Case LIST_CLI
            ReDim TempQuest.CLI(1 To 2)

            'Copy the slots
            TempQuest.CLI(1) = Quest(Index).CLI(ArrayID + Dir)
            TempQuest.CLI(2) = Quest(Index).CLI(ArrayID)
            
            'Paste the slots
            Quest(Index).CLI(ArrayID + Dir) = TempQuest.CLI(2)
            Quest(Index).CLI(ArrayID) = TempQuest.CLI(1)
            
        Case LIST_TASK
            ReDim TempQuest.CLI(1 To 2)
            ReDim Preserve TempQuest.CLI(1).Action(1 To 1)
            ReDim Preserve TempQuest.CLI(2).Action(1 To 1)
        
            'Copy the slots
            TempQuest.CLI(1).Action(1) = Quest(Index).CLI(CLIIndex).Action(ArrayID + Dir)
            TempQuest.CLI(2).Action(1) = Quest(Index).CLI(CLIIndex).Action(ArrayID)
            
            'Paste the slots
            Quest(Index).CLI(CLIIndex).Action(ArrayID + Dir) = TempQuest.CLI(2).Action(1)
            Quest(Index).CLI(CLIIndex).Action(ArrayID) = TempQuest.CLI(1).Action(1)
        Case Else
            Exit Sub
    End Select
    
    Call QuestEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "MoveListItem", "modQuest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' //////////////////
' // Quest Editor //
' //////////////////
Public Sub QuestEditorInit()
    Dim I As Long
    Dim SoundSet As Boolean
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    EditorIndex = frmEditor_Quest.lstIndex.ListIndex + 1
    Quest_Changed(EditorIndex) = True

    With Quest(EditorIndex)
        
        ' Basic data
        frmEditor_Quest.txtName.text = Trim$(.Name)
        frmEditor_Quest.txtDesc.text = Trim$(.Description)
        frmEditor_Quest.chkRetake.Value = .CanBeRetaken
        
        ' Gender requirement
        frmEditor_Quest.cmbGenderReq.ListIndex = .Requirements.GenderReq
        
        ' Skill requirement
        frmEditor_Quest.cmbSkillReq.ListIndex = .Requirements.SkillReq
        frmEditor_Quest.scrlSkill.Value = .Requirements.SkillLevelReq
        
        ' Basic requirements
        frmEditor_Quest.scrlAccessReq.Value = .Requirements.AccessReq
        frmEditor_Quest.scrlLevelReq.Value = .Requirements.LevelReq
        
        ' Class requirements
        frmEditor_Quest.cmbClassReq.ListIndex = .Requirements.ClassReq
        
        ' Loop for stats
        For I = 1 To Stats.Stat_Count - 1
            frmEditor_Quest.scrlStatReq(I).Value = .Requirements.Stat_Req(I)
        Next
        
        'Loop for CLI's
        frmEditor_Quest.lstTasks.Clear
        frmEditor_Quest.CLI.Clear
        For I = 1 To .Max_CLI
            If .CLI(I).ItemIndex > 0 Then
                If .CLI(I).isNPC Then
                    frmEditor_Quest.CLI.AddItem "Meet with: " & Trim$(NPC(.CLI(I).ItemIndex).Name)
                Else
                    'frmEditor_Quest.CLI.AddItem I & ": " & Trim$(tmpEvent(.CLI(I).ItemIndex).Name)
                End If
            End If
        Next I
    End With
    
    Call CheckStartMsg
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "QuestEditorInit", "modQuest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' ///////////////////////////
' // Quest Editor CLI List //
' ///////////////////////////
Public Sub QuestEditorInitCLI()
    Dim I As Long
    Dim Index As Long
    Dim Tmp As String
    Dim Indent As String
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    Index = frmEditor_Quest.CLI.ListIndex + 1
    If Index < 1 Then Exit Sub
    With Quest(EditorIndex).CLI(Index)
        
        'Loop for Actions
        frmEditor_Quest.lstTasks.Clear
        
        For I = 1 To .Max_Actions
            If .Action(I).ActionID > 0 Then
                'Little more in depth because it turns code into a description easily readable by the user
                Select Case .Action(I).ActionID
                    
                    Case TASK_KILL 'gather items
                        frmEditor_Quest.lstTasks.AddItem "Kill " & .Action(I).Amount & " " & Trim$(NPC(.Action(I).MainData).Name) & "('s)."
                    
                    Case TASK_GATHER 'gather items
                        If .Action(I).SecondaryData = 1 Then
                            frmEditor_Quest.lstTasks.AddItem "Gather and handover " & .Action(I).Amount & " " & Trim$(Item(.Action(I).MainData).Name) & "('s)."
                        Else
                            frmEditor_Quest.lstTasks.AddItem "Gather " & .Action(I).Amount & " " & Trim$(Item(.Action(I).MainData).Name) & "('s)."
                        End If
                        
                    Case TASK_GETSKILL
                        frmEditor_Quest.lstTasks.AddItem "Obtain level " & .Action(I).Amount & " for the " & GetSkillName(.Action(I).MainData) & " skill."
                    
                    Case ACTION_GIVE_ITEM 'give player item(s)
                        frmEditor_Quest.lstTasks.AddItem "----Give " & .Action(I).Amount & " " & Trim$(Item(.Action(I).MainData).Name) & "('s) to player."
                    
                    Case ACTION_TAKE_ITEM 'take player item(s)
                        frmEditor_Quest.lstTasks.AddItem "----Take " & .Action(I).Amount & " " & Trim$(Item(.Action(I).MainData).Name) & "('s) from player."
                    
                    Case ACTION_SHOWMSG 'show the player a message
                        If .Action(I).MainData = vbChecked Then
                            frmEditor_Quest.lstTasks.AddItem "START: (" & GetColorName(.Action(I).TertiaryData) & ") Show start msg: " & """" & Trim$(.Action(I).TextHolder) & """"
                        Else
                            If .Action(I).SecondaryData = vbChecked Then
                                frmEditor_Quest.lstTasks.AddItem "---- (" & GetColorName(.Action(I).TertiaryData) & ") Show msg if last task is incomplete: " & """" & Trim$(.Action(I).TextHolder) & """"
                            Else
                                frmEditor_Quest.lstTasks.AddItem "---- (" & GetColorName(.Action(I).TertiaryData) & ") Show msg: " & """" & Trim$(.Action(I).TextHolder) & """"
                            End If
                        End If
                    
                    Case ACTION_ADJUST_EXP 'adjust the player's experience
                        If .Action(I).MainData = 0 Then
                            If .Action(I).Amount > 0 Then Tmp = "+"
                            frmEditor_Quest.lstTasks.AddItem "----Modify Player EXP by " & Tmp & .Action(I).Amount
                        Else
                            frmEditor_Quest.lstTasks.AddItem "----Set Player EXP to " & .Action(I).Amount
                        End If
                    
                    Case ACTION_ADJUST_LVL 'adjust the player's level
                        If .Action(I).MainData = 0 Then
                            If .Action(I).Amount > 0 Then Tmp = "+"
                            frmEditor_Quest.lstTasks.AddItem "---- Modify Player Level by " & Tmp & .Action(I).Amount
                        Else
                            frmEditor_Quest.lstTasks.AddItem "---- Set Player Level to " & .Action(I).Amount
                        End If
                    
                    Case ACTION_ADJUST_STAT_LVL 'adjust the player's stat level
                        If .Action(I).MainData = 0 Then
                            If .Action(I).Amount > 0 Then Tmp = "+"
                            frmEditor_Quest.lstTasks.AddItem "---- Modify Player's " & GetStatName(.Action(I).SecondaryData) & " Level by " & Tmp & .Action(I).Amount
                        Else
                            frmEditor_Quest.lstTasks.AddItem "---- Set Player's " & GetStatName(.Action(I).SecondaryData) & " Level to " & .Action(I).Amount
                        End If
                        
                    Case ACTION_ADJUST_SKILL_LVL 'adjust the player's skill level
                        If .Action(I).MainData = 0 Then
                            If .Action(I).Amount > 0 Then Tmp = "+"
                            frmEditor_Quest.lstTasks.AddItem "---- Modify Player's " & GetSkillName(.Action(I).SecondaryData) & " level by " & Tmp & .Action(I).Amount
                        Else
                            frmEditor_Quest.lstTasks.AddItem "---- Set Player's " & GetSkillName(.Action(I).SecondaryData) & " level to " & .Action(I).Amount
                        End If
                        
                    Case ACTION_ADJUST_SKILL_EXP 'adjust the player's skill exp
                        If .Action(I).MainData = 0 Then
                            If .Action(I).Amount > 0 Then Tmp = "+"
                            frmEditor_Quest.lstTasks.AddItem "---- Modify Player's " & GetSkillName(.Action(I).SecondaryData) & " EXP by " & Tmp & .Action(I).Amount
                        Else
                            frmEditor_Quest.lstTasks.AddItem "---- Set Player's " & GetSkillName(.Action(I).SecondaryData) & " EXP to " & .Action(I).Amount
                        End If
                        
                    Case ACTION_ADJUST_STAT_POINTS 'adjust the player's stat points
                        If .Action(I).MainData = 0 Then
                            If .Action(I).Amount > 0 Then Tmp = "+"
                            frmEditor_Quest.lstTasks.AddItem "---- Modify Player's stat points by " & Tmp & .Action(I).Amount
                        Else
                            frmEditor_Quest.lstTasks.AddItem "---- Set Player's stat points to " & .Action(I).Amount
                        End If
                        
                    Case ACTION_WARP
                        frmEditor_Quest.lstTasks.AddItem "---- Warp player to Map: " & .Action(I).Amount & " (X" & .Action(I).MainData & ", Y" & .Action(I).SecondaryData & ")"
                    
                    Case Else
                        Exit Sub
                
                End Select
            End If
            
            Tmp = vbNullString
        Next I
    End With
    
    Call CheckStartMsg
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "QuestEditorInit", "modQuest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DeleteAction(ByVal QuestID As Long, ByVal Index As Long, ByVal ActionID As Long)
Dim I As Long
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo ErrorHandler
    With Quest(QuestID).CLI(Index)
        For I = 1 To .Max_Actions
            If I >= ActionID Then
                Call ZeroMemory(ByVal VarPtr(.Action(I)), LenB(.Action(I)))
                
                'start swaping the following slots
                If I + 1 <= .Max_Actions Then
                    .Action(I) = .Action(I + 1)
                End If
            End If
        Next I
        
        Call ZeroMemory(ByVal VarPtr(.Action(.Max_Actions)), LenB(.Action(.Max_Actions)))
        .Max_Actions = .Max_Actions - 1
        If .Max_Actions > 0 Then
            ReDim Preserve .Action(1 To .Max_Actions)
        Else
            ReDim .Action(0 To 0)
        End If
    End With
    Exit Sub
    
    ' Error handler
ErrorHandler:
    HandleError "DeleteAction", "modQuest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub DeleteCLI(ByVal QuestID As Long, ByVal Index As Long)
Dim I As Long
' If debug mode, handle error then exit out
If Options.Debug = 1 Then On Error GoTo ErrorHandler
    With Quest(QuestID)
        For I = 1 To .Max_CLI
            If I >= Index Then
                Call ZeroMemory(ByVal VarPtr(.CLI(I)), LenB(.CLI(I)))
                
                'start swaping the following slots
                If I + 1 <= .Max_CLI Then
                    .CLI(I) = .CLI(I + 1)
                End If
            End If
        Next I
        
        Call ZeroMemory(ByVal VarPtr(.CLI(.Max_CLI)), LenB(.CLI(.Max_CLI)))
        .Max_CLI = .Max_CLI - 1
        If .Max_CLI > 0 Then
            ReDim Preserve .CLI(1 To .Max_CLI)
        Else
            ReDim .CLI(0 To 0)
        End If
    End With
    
    Call QuestEditorInit
    Exit Sub
    
    ' Error handler
ErrorHandler:
    HandleError "DeleteCLI", "modQuest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CheckStartMsg()
Dim I As Long, II As Long, III As Long

For I = 1 To MAX_QUESTS
    For II = 1 To Quest(I).Max_CLI
        For III = 1 To Quest(I).CLI(II).Max_Actions
            If Quest(I).CLI(II).Action(III).ActionID = ACTION_SHOWMSG Then
                If Quest(I).CLI(II).Action(III).MainData = vbChecked Then
                    frmEditor_Quest.chkStart.Value = vbUnchecked
                    Exit Sub
                End If
            End If
        Next III
    Next II
Next I
End Sub

Public Sub CheckResponseMsg(ByVal QuestID As Long, ByVal CLIIndex As Long, ByVal SearchStopIndex As Long)
Dim I As Long

frmEditor_Quest.chkRes.Enabled = False
If QuestID < 1 Or QuestID > MAX_QUESTS Then Exit Sub
If CLIIndex < 1 Then Exit Sub
    
    For I = 1 To SearchStopIndex
        If Quest(QuestID).CLI(CLIIndex).Action(I).ActionID > 0 And Quest(QuestID).CLI(CLIIndex).Action(I).ActionID < 5 Then 'find a task
            frmEditor_Quest.chkRes.Enabled = True 'found a task, allow access to the checkbox
            Exit Sub
        End If
    Next I
End Sub

Public Function IsNPCInAnotherQuest(ByVal NPCIndex As Long, CurQuest As Long) As Boolean
Dim I As Long
    IsNPCInAnotherQuest = False
    
    For I = 1 To MAX_QUESTS
        If I <> CurQuest Then
            If Quest(I).Max_CLI > 0 Then 'prevent subscript out of range within the following prcedure
                If Quest(I).CLI(1).ItemIndex = NPCIndex Then
                    IsNPCInAnotherQuest = True
                    Exit Function
                End If
            End If
        End If
    Next I
End Function

Public Function DoesNPCStartQuest(ByVal NPCNum As Long) As Long
Dim I As Long
    DoesNPCStartQuest = 0
    
    For I = 1 To MAX_QUESTS
        If Quest(I).Max_CLI > 0 Then
            If Quest(I).CLI(1).ItemIndex = NPCNum Then
                DoesNPCStartQuest = I 'return the quest number
                Exit Function
            End If
        End If
    Next I
End Function

