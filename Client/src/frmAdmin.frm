VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdmin 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Panel"
   ClientHeight    =   8625
   ClientLeft      =   810
   ClientTop       =   330
   ClientWidth     =   2850
   Icon            =   "frmAdmin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   190
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8640
      Left            =   30
      ScaleHeight     =   576
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   2835
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   1650
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   40
         Top             =   2310
         Width           =   510
      End
      Begin VB.TextBox txtSprite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1590
         TabIndex        =   39
         Text            =   "0"
         Top             =   3090
         Width           =   600
      End
      Begin VB.ComboBox cmbAccess 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         Enabled         =   0   'False
         ForeColor       =   &H80000002&
         Height          =   315
         ItemData        =   "frmAdmin.frx":038A
         Left            =   480
         List            =   "frmAdmin.frx":039D
         TabIndex        =   38
         Text            =   "Player's Access"
         Top             =   810
         Width           =   1695
      End
      Begin VB.ComboBox cmbPlayersOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   120
         TabIndex        =   37
         Text            =   "Choose Player"
         Top             =   390
         Width           =   2055
      End
      Begin VB.PictureBox picRefresh 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2280
         ScaleHeight     =   285
         ScaleWidth      =   345
         TabIndex        =   36
         Top             =   390
         Width           =   375
      End
      Begin VB.CommandButton cmdCharEditor 
         Caption         =   "Character Editor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1050
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1725
      End
      Begin VB.CommandButton cmdAEmoticon 
         Caption         =   "Emoticon"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   6780
         Width           =   1215
      End
      Begin VB.CommandButton cmdAClass 
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   6780
         Width           =   1200
      End
      Begin VB.CommandButton cmdAMute 
         Caption         =   "Mute"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   1500
         Width           =   855
      End
      Begin VB.CommandButton cmdATitle 
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   6480
         Width           =   1215
      End
      Begin VB.CommandButton cmdABanE 
         Caption         =   "Ban"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   6180
         Width           =   1215
      End
      Begin VB.CommandButton cmdLevelUp 
         Caption         =   "Level Up"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2100
         Width           =   855
      End
      Begin VB.CommandButton cmdAAnim 
         Caption         =   "Animation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   6180
         Width           =   1215
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "Respawn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   4650
         Width           =   2535
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Spawn Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   750
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   7980
         Width           =   1215
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         Left            =   1530
         Min             =   1
         TabIndex        =   2
         Top             =   7680
         Value           =   1
         Width           =   1155
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   1830
         Min             =   1
         TabIndex        =   1
         Top             =   7410
         Value           =   1
         Width           =   825
      End
      Begin VB.CommandButton cmdASpell 
         Caption         =   "Spell"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdAShop 
         Caption         =   "Shop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdAResource 
         Caption         =   "Resource"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   5580
         Width           =   1215
      End
      Begin VB.CommandButton cmdANpc 
         Caption         =   "NPC"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   5580
         Width           =   1215
      End
      Begin VB.CommandButton cmdAMap 
         Caption         =   "Map"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdAItem 
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   5280
         Width           =   1215
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "Map Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   4350
         Width           =   1215
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4350
         Width           =   1215
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   4050
         Width           =   1215
      End
      Begin VB.TextBox txtAMap 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   0
         Top             =   3990
         Width           =   465
      End
      Begin VB.CommandButton cmdAWarpMeTo 
         Caption         =   "Admin To Player"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1050
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1725
      End
      Begin VB.CommandButton cmdAWarpToMe 
         Caption         =   "Summon Player"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1050
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1725
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   855
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1800
         Width           =   855
      End
      Begin VB.CommandButton cmdAMoral 
         Caption         =   "Moral"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   6480
         Width           =   1215
      End
      Begin MSComCtl2.UpDown upSprite 
         Height          =   555
         Left            =   2250
         TabIndex        =   41
         Top             =   2430
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   979
         _Version        =   393216
         BuddyControl    =   "txtSprite"
         BuddyDispid     =   196611
         OrigLeft        =   3990
         OrigTop         =   1770
         OrigRight       =   4245
         OrigBottom      =   2265
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   30
         TabIndex        =   35
         Top             =   8250
         Visible         =   0   'False
         Width           =   2760
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMap 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Map"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   34
         Top             =   3600
         Width           =   2385
      End
      Begin VB.Label lblEditors 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Editors"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   33
         Top             =   4920
         Width           =   2385
      End
      Begin VB.Label lblSpawning 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Spawning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   32
         Top             =   7050
         Width           =   2445
      End
      Begin VB.Label lblPlayers 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Players"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   31
         Top             =   0
         Width           =   2505
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   8
         X2              =   180
         Y1              =   18
         Y2              =   18
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         Top             =   7740
         Width           =   1335
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Item: None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   27
         Top             =   7440
         Width           =   1875
      End
      Begin VB.Line Line4 
         BorderWidth     =   3
         X1              =   8
         X2              =   176
         Y1              =   488
         Y2              =   488
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   8
         X2              =   176
         Y1              =   344
         Y2              =   344
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   26
         Top             =   4020
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   10
         X2              =   178
         Y1              =   258
         Y2              =   258
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim refreshDown As Boolean
Dim autoAccess As Boolean, autoSprite As Boolean
Dim currentSprite As Long
Private Sub cmbAccess_Click()
    If autoAccess Then
        autoAccess = False
    Else
        cmbAccess.Enabled = False
        cmbPlayersOnline.Enabled = False
        SendSetAccess cmbPlayersOnline.text, cmbAccess.ListIndex
    End If
End Sub

Public Sub VerifyAccess(PlayerName As String, Success As Byte, Message As String, CurrentAccess As Byte)
    If PlayerName = cmbPlayersOnline.text Then
        If Success = 0 Then
            For i = 0 To UBound(g_playersOnline)
                If InStr(1, g_playersOnline(i), PlayerName) Then
                    Mid(g_playersOnline(i), InStr(1, g_playersOnline(i), ":"), 2) = ":" & CurrentAccess
                    setAdminAccessLevel
                    
                    DisplayStatus Message, Status.Error
                End If
            Next i
        ElseIf Success = 1 Then
            Mid(g_playersOnline(i), InStr(1, g_playersOnline(i), ":"), 2) = ":" & CurrentAccess
            setAdminAccessLevel
            
            DisplayStatus Message, Status.Correct
        End If
    End If
    cmbPlayersOnline.Enabled = True
End Sub

Public Sub DisplayStatus(ByVal msg As String, msgType As Status)
    Select Case msgType
        Case Status.Error:
            lblStatus.BackColor = &H8080FF
            lblStatus.Caption = msg
        Case Status.Correct:
            lblStatus.BackColor = &H80FF80
            lblStatus.Caption = msg
        Case Status.Neutral:
            lblStatus.BackColor = &H80FFFF
            lblStatus.Caption = msg
        Case Status.Info_:
            lblStatus.BackColor = &H8000000F
            lblStatus.Caption = msg
    End Select
    lblStatus.Visible = True
End Sub

Private Sub cmbPlayersOnline_Click()
    Dim i As Long, Length As Long
    
    Length = UBound(ignoreIndexes)
    For i = 0 To Length
        If cmbPlayersOnline.ListIndex = ignoreIndexes(i) Then
            cmbPlayersOnline.ListIndex = ignoreIndexes(i) + 1
            cmbPlayersOnline.text = cmbPlayersOnline.List(cmbPlayersOnline.ListIndex)
            Exit Sub
        End If
    Next
    autoAccess = True
    autoSprite = True
    For i = 0 To UBound(g_playersOnline)
            If InStr(1, g_playersOnline(i), cmbPlayersOnline.text) Then
                txtSprite.text = Split(g_playersOnline(i), ":")(2)
            End If
    Next i
    If Player(MyIndex).Access < 4 Then
        txtSprite.Enabled = False
        upSprite.Enabled = False
    Else
        txtSprite.Enabled = True
        upSprite.Enabled = True
    End If
    setAdminAccessLevel

    
End Sub

Private Sub setAdminAccessLevel()
    Dim accessLvl As String, tempTxt As String
    
    ' Set Access Level
    For i = 0 To UBound(g_playersOnline)
        If InStr(1, g_playersOnline(i), cmbPlayersOnline.List(cmbPlayersOnline.ListIndex)) Then
            accessLvl = Split(g_playersOnline(i), ":")(1)
            txtSprite.text = Split(g_playersOnline(i), ":")(2)
            
            If accessLvl = "5" Then
                accessLvl = "4"
                tempTxt = "Owner"

            Else
                tempTxt = cmbAccess.List(CLng(accessLvl))

            End If
            
            If Player(MyIndex).Access > CLng(accessLvl) And Player(MyIndex).Access >= 4 And Trim(Player(MyIndex).name) <> cmbPlayersOnline.text Then
                cmbAccess.Enabled = True
            Else
                cmbAccess.Enabled = False
            End If
            If Player(MyIndex).Access < 4 Then
                txtSprite.Enabled = False
                upSprite.Enabled = False
            Else
                txtSprite.Enabled = True
                upSprite.Enabled = True
            End If
            cmbAccess.ListIndex = accessLvl
            cmbAccess.text = tempTxt
        End If
    Next i
End Sub

Private Sub cmdAEmoticon_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    SendRequestEditEmoticon
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAEmoticon_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdASetPlayerSprite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetPlayerAccess(MyIndex) < STAFF_ADMIN Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    ' Subscript out of range
    If Len(Trim$(cmbPlayersOnline.text)) < 1 Then Exit Sub
    If IsNumeric(Trim$(cmbPlayersOnline.text)) Then Exit Sub
    If Not IsNumeric(Trim$(txtASprite.text)) Then Exit Sub
    If Int(txtASprite.text) > NumCharacters Or Int(txtASprite.text) < 1 Then Exit Sub

    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdASetPlayerSprite_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAAnim_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    SendRequestEditAnimation
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAAnim_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
'Character Editor
Private Sub cmdCharEditor_Click()
    ' Send request for character names
    Tex_CharSprite.Texture = 0
    SendRequestAllCharacters
End Sub

Private Sub cmdLevelUp_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    SendRequestLevelUp
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdLevelUp_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdALoc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    BLoc = Not BLoc
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdALoc_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAMap_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    SendRequestEditMap
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAMap_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAWarpToMe_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(cmbPlayersOnline.text)) < 1 Then Exit Sub
    If IsNumeric(Trim$(cmbPlayersOnline.text)) Then Exit Sub

    WarpToMe Trim$(cmbPlayersOnline.text)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAWarpToMe_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAWarpMeTo_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    ' Subscript out of range
    If Len(Trim$(cmbPlayersOnline.text)) < 1 Then Exit Sub
    If IsNumeric(Trim$(cmbPlayersOnline.text)) Then Exit Sub

    WarpMeTo Trim$(cmbPlayersOnline.text)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAWarpMeTo_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAWarp_Click()
    Dim n As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then Exit Sub
    If Not IsNumeric(Trim$(txtAMap.text)) Then Exit Sub
    
    n = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If n > 0 And n <= MAX_MAPS Then
        Call WarpTo(n)
    Else
        Call AddText("Invalid map number.", Red)
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAWarp_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdASprite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(txtASprite.text)) < 1 Then Exit Sub
    If Not IsNumeric(Trim$(txtASprite.text)) Then Exit Sub

    SendSetSprite CLng(Trim$(txtASprite.text))
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdASprite_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAMapReport_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    SendMapReport
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAMapReport_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdARespawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    SendMapRespawn
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdARespawn_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAKick_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_MODERATOR Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(cmbPlayersOnline.text)) < 1 Then Exit Sub

    SendKick Trim$(cmbPlayersOnline.text)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAKick_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdABan_Click()
    Dim StrInput As String

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_ADMIN Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(cmbPlayersOnline.text)) < 1 Then Exit Sub

    StrInput = InputBox("Reason: ", "Ban")

    SendBan Trim$(cmbPlayersOnline.text), Trim$(StrInput)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdABan_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    SendRequestEditItem
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAItem_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdANpc_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    SendRequestEditNPC
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdANpc_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAResource_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    SendRequestEditResource
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAResource_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    SendRequestEditShop
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAShop_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdASpell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    SendRequestEditSpell
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdASpell_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub


Private Sub cmdASpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    SendSpawnItem scrlAItem.Value, scrlAAmount.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdASpawn_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdABanE_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetPlayerAccess(MyIndex) < STAFF_ADMIN Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    SendRequestEditBan
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdABanE_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAMute_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_MODERATOR Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(cmbPlayersOnline.text)) < 1 Then Exit Sub

    SendMute Trim$(cmbPlayersOnline.text)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAMute_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAClass_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    SendRequestEditClass
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAClass_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdATitle_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    SendRequestEditTitle
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdATitle_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdAMoral_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    SendRequestEditMoral
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "cmdAMoral_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyInsert
            If Player(MyIndex).Access >= STAFF_MODERATOR Then
                If frmAdmin.Visible And GetForegroundWindow = frmAdmin.hwnd Then
                    Unload frmAdmin
                End If
            End If
    End Select
End Sub

Private Sub picAdmin_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picRefresh.Picture = LoadResPicture("REFRESH_UP", vbResBitmap)
End Sub

Private Sub picAdmin_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    refreshDown = False
End Sub

Private Sub picRefresh_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    refreshDown = True
    picRefresh.Picture = LoadResPicture("REFRESH_DOWN", vbResBitmap)
End Sub

Private Sub picRefresh_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Not refreshDown Then
        picRefresh.Picture = LoadResPicture("REFRESH_OVER", vbResBitmap)
    End If
End Sub

Private Sub picRefresh_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    refreshDown = False
    refreshingAdminList = True
    SendRequestPlayersOnline
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    UpdateAdminScrollBar
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlAItem_Change", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAAmount.Caption = "Amount: " & scrlAAmount.Value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlAAmount_Change", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub UpdatePlayersOnline()
    Dim players() As String, Staff() As String, tempTxt As String, temp() As String, Length As Long, i As Long, currentIgnore As Long
    Dim stuffCounter As Long, playersCounter As Long, overallCounter As Long, foundStuff As Boolean, foundPlayer As Boolean
    
    tempTxt = cmbPlayersOnline.text
    cmbPlayersOnline.Clear
    cmbPlayersOnline.text = tempTxt
    
    ' Get Stuff
    For i = 0 To UBound(g_playersOnline)
        If CByte(Split(g_playersOnline(i), ":")(1)) > 0 Then
            foundStuff = True
            ReDim Preserve Staff(stuffCounter)
            Staff(stuffCounter) = Split(g_playersOnline(i), ":")(0)
            stuffCounter = stuffCounter + 1
        End If
    Next
    
    'Get Players
    For i = 0 To UBound(g_playersOnline)
        If CByte(Split(g_playersOnline(i), ":")(1)) = 0 Then
            foundPlayer = True
            ReDim Preserve players(playersCounter)
            players(playersCounter) = Split(g_playersOnline(i), ":")(0)
            playersCounter = playersCounter + 1
        End If
    Next
    
    If foundStuff Then
        cmbPlayersOnline.AddItem ("----Staff: " & stuffCounter & "-----")
        
            ReDim Preserve ignoreIndexes(0)
            ignoreIndexes(0) = currentIgnore
            currentIgnore = currentIgnore + 1
            
        For i = 0 To UBound(Staff)
            cmbPlayersOnline.AddItem (Trim(Staff(i)))
            currentIgnore = currentIgnore + 1
        Next
        overallCounter = overallCounter + stuffCounter
    End If

    If foundPlayer Then
        cmbPlayersOnline.AddItem ("----Players: " & playersCounter & "----")
        
            ReDim Preserve ignoreIndexes(1)
            ignoreIndexes(1) = currentIgnore
            currentIgnore = currentIgnore + 1
        For i = 0 To UBound(players)
            cmbPlayersOnline.AddItem (Trim(players(i)))
            currentIgnore = currentIgnore + 1
        Next
        overallCounter = overallCounter + playersCounter
    End If
    
    lblPlayers.Caption = "Players: " & overallCounter
End Sub
Public Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    frmAdmin.picRefresh.BorderStyle = 0
    'Me.Move frmMain.Left + frmMain.Width, frmMain.Top
    If Trim(cmbPlayersOnline.text) = "Choose Player" Then
        txtSprite.Enabled = False
        upSprite.Enabled = False
    End If
    upSprite.max = NumCharacters
    upSprite.min = 0
    
    LastAdminSpriteTimer = timeGetTime

    scrlAItem.max = MAX_ITEMS
    UpdateAdminScrollBar
    picRefresh.Picture = LoadResPicture("REFRESH_UP", vbResBitmap)
    refreshingAdminList = True
    SendRequestPlayersOnline
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_Load", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtASprite_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtASprite.SelStart = Len(txtASprite)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtASprite_GotFocus", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtAMap_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtAMap.SelStart = Len(txtAMap)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtAMap_GotFocus", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtAName_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtAName.SelStart = Len(txtAName)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtAName_GotFocus", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
Private Function correctValue(ByRef textBox As textBox, ByRef valueToChange, min As Long, max As Long, Optional defaultVal As Long = 0) As Boolean
    Dim test As textBox
    
    If textBox.text = "" Then
        textBox.text = CStr(defaultVal)
        valueToChange = defaultVal
        correctValue = True
    End If

    If Len(textBox.text) = 1 And InStr(1, textBox.text, "-") = 1 Then
        correctValue = True
        Exit Function
    ElseIf Len(textBox.text) = 1 And IsNumeric(textBox.text) Then
        If verifyValue(textBox, min, max) Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
            correctValue = False
        End If
    ElseIf Len(textBox.text) > 1 And InStr(1, textBox.text, "-") = 0 And InStrRev(textBox.text, "-") = 0 And IsNumeric(textBox.text) Then

        If verifyValue(textBox, min, max) Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
            correctValue = False
        End If

    ElseIf Len(textBox.text) > 1 And InStr(1, textBox.text, "-") = 1 And InStrRev(textBox.text, "-") = 1 And IsNumeric(textBox.text) Then

        If verifyValue(textBox, min, max) Then
            TempValue = textBox.text
            valueToChange = TempValue
            correctValue = True
        Else
            textBox.text = CStr(valueToChange)
            textBox.SelStart = Len(textBox.text)
        correctValue = False
        End If
        
    Else
        textBox.text = CStr(valueToChange)
        textBox.SelStart = Len(textBox.text)
        correctValue = False
    End If
End Function

Private Sub reviseValue(ByRef textBox As textBox, ByRef valueToChange)
    If Not IsNumeric(textBox.text) Then
        textBox.text = CStr(valueToChange)
        displayFieldStatus textBox, " field accepts only Numbers!" & vbCrLf & "Reverting to last correct value...", Status.Correct
    Else
        textBox.text = CStr(valueToChange)
        displayFieldStatus textBox, " field is correct. Saving...", Status.Correct
    End If
End Sub

Private Function verifyValue(txtBox As textBox, min As Long, max As Long)
    Dim msg As String
    
    If (CLng(txtBox.text) >= min And CLng(txtBox.text) <= max) Then
        verifyValue = True
    Else
        msg = " field accepts only values: " & CStr(min) & " < value < " & CStr(max) & "." & vbCrLf & "Reverting value..."
        displayFieldStatus txtBox, msg, Status.Error
        verifyValue = False
    End If
End Function
Public Sub displayFieldStatus(ByVal txtBox As textBox, ByVal msg As String, msgType As Status)
    lblStatus.Visible = True
    Select Case msgType

        Case Status.Error:
            lblStatus.BackColor = &H8080FF
            lblStatus.Caption = Replace(txtBox.name, "txt", "") & msg
        Case Status.Correct:
            lblStatus.BackColor = &H80FF80
            lblStatus.Caption = Replace(txtBox.name, "txt", "") & msg
        Case Status.Neutral:
            lblStatus.BackColor = &H80FFFF
            lblStatus.Caption = Replace(txtBox.name, "txt", "") & msg
        Case Status.Info_:
            lblStatus.BackColor = &H8000000F
            lblStatus.Caption = Replace(txtBox.name, "txt", "") & msg
    End Select
End Sub
Private Sub selectValue(ByRef textBox As textBox)
    textBox.SelStart = 0
    textBox.SelLength = Len(textBox.text)
End Sub

Private Sub txtSprite_Change()
    If autoSprite Then
        autoSprite = False
        Exit Sub
    End If
    
     If correctValue(txtSprite, currentSprite, 0, NumCharacters) Then
        If txtSprite.text = 0 Then picSprite.Picture = Nothing
        If GetPlayerAccess(MyIndex) < STAFF_ADMIN Then
            AddText "You have insufficent access to do this!", BrightRed
            Exit Sub
        ElseIf txtSprite.text > 0 Then
            For i = 0 To UBound(g_playersOnline)
                If InStr(1, g_playersOnline(i), cmbPlayersOnline.text) Then
                    Mid(g_playersOnline(i), InStr(InStr(1, g_playersOnline(i), ":") + 1, g_playersOnline(i), ":"), Len(txtSprite.text) + 1) = ":" & txtSprite.text
                End If
            Next i

            SendSetPlayerSprite Trim$(cmbPlayersOnline.text), currentSprite
        End If


     End If

End Sub

Private Sub txtSprite_Click()
     selectValue txtSprite
End Sub

Private Sub txtSprite_GotFocus()
    selectValue txtSprite
End Sub

Private Sub txtSprite_LostFocus()
    reviseValue txtSprite, currentSprite
End Sub
