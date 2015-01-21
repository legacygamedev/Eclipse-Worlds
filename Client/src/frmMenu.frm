VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form frmMenu 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12270
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMenu.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   0  'User
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSWinsockLib.Winsock Socket 
      Left            =   0
      Top             =   1155
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Timer tmrUpdateNews 
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmbClass_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    'NewCharClass = ClassSelection(cmbClass.ListIndex + 1)
    Menu_DrawCharacter
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmbClass_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
    Dim TmpTxt As String, TmpArray() As String, i As Long

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    ' General menu things
    Me.Caption = ""
    
    ' Set the maxes for the username and password
    txtLUser.MaxLength = NAME_LENGTH
    txtLPass.MaxLength = NAME_LENGTH
    txtCUser.MaxLength = NAME_LENGTH
    txtRUser.MaxLength = NAME_LENGTH
    txtRPass.MaxLength = NAME_LENGTH
    txtRPass2.MaxLength = NAME_LENGTH
    
    ' Load the username
    If Options.SaveUsername = 1 Then
        txtLUser.text = Trim$(Options.UserName)
        txtLUser.SelStart = Len(Trim$(Options.UserName))
        chkUsername.Value = Options.SaveUsername
    End If
    
    ' Load the password
    If Options.SavePassword = 1 Then
        txtLPass.text = Trim$(Options.Password)
        txtLPass.SelStart = Len(Trim$(Options.Password))
        chkPass.Value = Options.SavePassword
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    LastButton_Menu = 0
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If Not EnteringGame And Not gameDestroyed Then DestroyGame
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgButton_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not CurButton_Menu = Index Then
        Call Audio.PlaySound(ButtonClick)
        CurButton_Menu = Index
        MenuButton(Index).State = 2
        Call RenderButton_Menu(Index)
        Call ResetMenuButtons
            
        Select Case Index
            Case 1
                If Not picLogin.Visible Then
                    picCredits.Visible = False
                    picLogin.Visible = True
                    picRegister.Visible = False
                    picCharacter.Visible = False
                    picMain.Visible = False
                    txtLUser.SetFocus
                End If
            Case 2
                If Not picRegister.Visible Then
                    picCredits.Visible = False
                    picLogin.Visible = False
                    picRegister.Visible = True
                    picCharacter.Visible = False
                    picMain.Visible = False
                    txtRUser.SetFocus
                End If
            Case 3
                If Not picCredits.Visible Then
                    picCredits.Visible = True
                    picLogin.Visible = False
                    picRegister.Visible = False
                    picCharacter.Visible = False
                    picMain.Visible = False
                End If
            Case 4
                Call DestroyGame
        End Select
    
        ' Reset all buttons
        ResetMenuButtons
        CurButton_Menu = Index
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ImgButton_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Not LastButton_Menu = Index And Not CurButton_Menu = Index Then
        ResetMenuButtons
        MenuButton(Index).State = 1
        Call RenderButton_Menu(Index)
        Call Audio.PlaySound(ButtonHover)
        LastButton_Menu = Index
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "ImgButton_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblLAccept_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If IsLoginLegal(txtLUser.text, txtLPass.text) Then
        Call MenuState(MENU_STATE_LOGIN)
    End If
    
    ' Save options
    Options.SaveUsername = frmMenu.chkUsername.Value
    Options.SavePassword = frmMenu.chkPass.Value
    Options.UserName = Trim$(frmMenu.txtLUser.text)
    
    If frmMenu.chkUsername.Value = 0 Then
        Options.UserName = vbNullString
    Else
        Options.UserName = Trim$(frmMenu.txtLUser.text)
    End If
    
    If frmMenu.chkPass.Value = 0 Then
        Options.Password = vbNullString
    Else
        Options.Password = Trim$(frmMenu.txtLPass.text)
    End If
    SaveOptions
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblLAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optFemale_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    NewCharClass = ClassSelection(cmbClass.ListIndex + 1)
    Menu_DrawCharacter
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optFemale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub optFemale_GotFocus()
    txtCUser.SetFocus
    
    If optFemale.Value = False Then
        optFemale.Value = True
    End If
End Sub

Private Sub optMale_GotFocus()
    txtCUser.SetFocus
    
    If optMale.Value = False Then
        optMale.Value = True
    End If
End Sub

Private Sub optMale_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    NewCharClass = ClassSelection(cmbClass.ListIndex + 1)
    Menu_DrawCharacter
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "optMale_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picCharacter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picCharacter_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picCredits_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picCredits_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picLogin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picLogin_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picMain_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picMain_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picRegister_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    ResetMenuButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "picRegister_MouseMove", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Socket_DataArrival", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub tmrUpdateNews_Timer()
    ' Don't update if menu is not visible
    If frmMenu.Visible = False Then Exit Sub
    
    ' If we're connected we don't need to update anything
    If IsConnected Then Exit Sub
    
    ' If the the timer is paused but we're not connected clear it
    If StopTimer = True And IsConnected = False Then
        StopTimer = False
    End If
    
    ' Check if the timer is disabled
    If StopTimer Then Exit Sub
    
    If ConnectToServer(1) Then
        Call UpdateData
        StopTimer = True
    End If
    
    If IsConnected = False Then
        frmMenu.lblServerStatus.Caption = "Offline"
        frmMenu.lblServerStatus.ForeColor = vbRed
        frmMenu.lblNews.Caption = "The server appears to be offline. Please try connecting again later."
        frmMenu.lblServerStatus.Visible = True
    End If
End Sub

' Register
Private Sub txtRAccept_Click()
    Dim Name As String
    Dim Password As String
    Dim PasswordAgain As String, RndCharacters As String
    
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Name = Trim$(txtRUser.text)
    Password = Trim$(txtRPass.text)
    PasswordAgain = Trim$(txtRPass2.text)

    If IsLoginLegal(Name, Password) Then
        If IsLoginLegal(Name, PasswordAgain) Then
            If Not Password = PasswordAgain Then
                Call AlertMsg("Passwords don't match.")
                Exit Sub
            End If
    
            If Not IsStringLegal(Name) Then Exit Sub
    
            Call MenuState(MENU_STATE_NEWACCOUNT)
        End If
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtRAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' New Character
Private Sub lblCAccept_Click()
    ' If debug mode, handle error then exit out
    If App.LogMode = 1 And Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If IsNewCharLegal(txtCUser) Then
        Call MenuState(MENU_STATE_ADDCHAR)
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lblCAccept_Click", "frmMenu", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    
    If KeyAscii = vbKeyReturn Then
        If picLogin.Visible = True And lblLAccept.Enabled Then
            lblLAccept_Click
        ElseIf picCharacter.Visible And lblCAccept.Enabled Then
            lblCAccept_Click
        ElseIf picRegister.Visible And txtRAccept.Enabled Then
            txtRAccept_Click
        ElseIf picMain.Visible Then
            Call ImgButton_Click(1)
        End If
        KeyAscii = 0
    End If
    
    If KeyAscii = vbKeyEscape Then
        DestroyGame
        KeyAscii = 0
    End If
End Sub
