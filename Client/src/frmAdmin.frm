VERSION 5.00
Begin VB.Form frmAdmin 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Panel"
   ClientHeight    =   8625
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   2820
   Icon            =   "frmAdmin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   575
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   188
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox picAdmin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000002&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8,25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   8640
      Left            =   0
      ScaleHeight     =   576
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   189
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   2835
      Begin VB.ComboBox cmbPlayersOnline 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H80000002&
         Height          =   315
         Left            =   120
         TabIndex        =   44
         Text            =   "Choose Player"
         Top             =   390
         Width           =   2055
      End
      Begin VB.PictureBox picRefresh 
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         Height          =   315
         Left            =   2280
         ScaleHeight     =   315
         ScaleWidth      =   375
         TabIndex        =   43
         Top             =   390
         Width           =   375
      End
      Begin VB.CommandButton cmdCharEditor 
         Caption         =   "Character's Editor"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   2100
         Width           =   2535
      End
      Begin VB.CommandButton cmdAEmoticon 
         Caption         =   "Emoticon"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   6660
         Width           =   1215
      End
      Begin VB.CommandButton cmdAClass 
         Caption         =   "Class"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   6660
         Width           =   1215
      End
      Begin VB.CommandButton cmdAMute 
         Caption         =   "Mute"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1215
      End
      Begin VB.CommandButton cmdATitle 
         Caption         =   "Title"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   6300
         Width           =   1215
      End
      Begin VB.CommandButton cmdABanE 
         Caption         =   "Ban"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   5940
         Width           =   1215
      End
      Begin VB.CommandButton cmdASetPlayerSprite 
         Caption         =   "Set Player Sprite"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   2400
         Width           =   2535
      End
      Begin VB.CommandButton cmdLevelUp 
         Caption         =   "Level Up"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdAAnim 
         Caption         =   "Animation"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
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
         Top             =   5940
         Width           =   1215
      End
      Begin VB.CommandButton cmdAAccess 
         Caption         =   "Set Access"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
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
         Top             =   1800
         Width           =   1215
      End
      Begin VB.TextBox txtAAccess 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1410
         TabIndex        =   0
         Top             =   870
         Width           =   1215
      End
      Begin VB.TextBox txtASprite 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   2730
         Width           =   465
      End
      Begin VB.CommandButton cmdARespawn 
         Caption         =   "Respawn"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4140
         Width           =   2535
      End
      Begin VB.CommandButton cmdASprite 
         Caption         =   "Set Sprite"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   2760
         Width           =   1215
      End
      Begin VB.CommandButton cmdASpawn 
         Caption         =   "Spawn Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   750
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   8070
         Width           =   1215
      End
      Begin VB.HScrollBar scrlAAmount 
         Height          =   255
         Left            =   2040
         Min             =   1
         TabIndex        =   4
         Top             =   7740
         Value           =   1
         Width           =   615
      End
      Begin VB.HScrollBar scrlAItem 
         Height          =   255
         Left            =   2070
         Min             =   1
         TabIndex        =   3
         Top             =   7410
         Value           =   1
         Width           =   585
      End
      Begin VB.CommandButton cmdASpell 
         Caption         =   "Spell"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
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
         Top             =   5580
         Width           =   1215
      End
      Begin VB.CommandButton cmdAShop 
         Caption         =   "Shop"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   5580
         Width           =   1215
      End
      Begin VB.CommandButton cmdAResource 
         Caption         =   "Resource"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
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
         Top             =   5220
         Width           =   1215
      End
      Begin VB.CommandButton cmdANpc 
         Caption         =   "NPC"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
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
         Top             =   5220
         Width           =   1215
      End
      Begin VB.CommandButton cmdAMap 
         Caption         =   "Map"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
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
         Top             =   4860
         Width           =   1215
      End
      Begin VB.CommandButton cmdAItem 
         Caption         =   "Item"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
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
         Top             =   4860
         Width           =   1215
      End
      Begin VB.CommandButton cmdAMapReport 
         Caption         =   "Map Report"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
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
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdALoc 
         Caption         =   "Location"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
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
         Top             =   3840
         Width           =   1215
      End
      Begin VB.CommandButton cmdAWarp 
         Caption         =   "Warp To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   3540
         Width           =   1215
      End
      Begin VB.TextBox txtAMap 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   3480
         Width           =   465
      End
      Begin VB.CommandButton cmdAWarpMeTo 
         Caption         =   "Warp Me To"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
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
         Top             =   1500
         Width           =   1215
      End
      Begin VB.CommandButton cmdAWarpToMe 
         Caption         =   "Warp To Me"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdABan 
         Caption         =   "Ban"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   1200
         Width           =   1215
      End
      Begin VB.CommandButton cmdAKick 
         Caption         =   "Kick"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   900
         Width           =   1215
      End
      Begin VB.CommandButton cmdAMoral 
         Caption         =   "Moral"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1440
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   6300
         Width           =   1215
      End
      Begin VB.Label lblStatus 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H0000FF00&
         Caption         =   "Status"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   30
         TabIndex        =   42
         Top             =   8370
         Visible         =   0   'False
         Width           =   2760
      End
      Begin VB.Label lblMap 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Map"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8,25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   41
         Top             =   3090
         Width           =   2385
      End
      Begin VB.Label lblEditors 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Editors"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8,25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   40
         Top             =   4470
         Width           =   2385
      End
      Begin VB.Label lblSpawning 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Spawning"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8,25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   150
         TabIndex        =   39
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
            Size            =   8,25
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   38
         Top             =   0
         Width           =   2445
      End
      Begin VB.Line Line3 
         BorderWidth     =   3
         X1              =   4
         X2              =   176
         Y1              =   18
         Y2              =   18
      End
      Begin VB.Label Label33 
         BackStyle       =   0  'Transparent
         Caption         =   "Access:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Left            =   1440
         TabIndex        =   35
         Top             =   690
         Width           =   645
      End
      Begin VB.Label Label31 
         BackStyle       =   0  'Transparent
         Caption         =   "Sprite #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   150
         TabIndex        =   34
         Top             =   2760
         Width           =   855
      End
      Begin VB.Label lblAAmount 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 1"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   33
         Top             =   7740
         Width           =   975
      End
      Begin VB.Label lblAItem 
         BackStyle       =   0  'Transparent
         Caption         =   "Item: None"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   32
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
         Y1              =   316
         Y2              =   316
      End
      Begin VB.Label Label30 
         BackStyle       =   0  'Transparent
         Caption         =   "Map #:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8,25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   210
         TabIndex        =   31
         Top             =   3510
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   10
         X2              =   178
         Y1              =   224
         Y2              =   224
      End
   End
End
Attribute VB_Name = "frmAdmin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim refreshDown As Boolean

Private Sub cmbPlayersOnline_Click()
    Dim i As Long, length As Long

    length = UBound(ignoreIndexes)
    For i = 0 To length
        If cmbPlayersOnline.ListIndex = ignoreIndexes(i) Then
            cmbPlayersOnline.ListIndex = ignoreIndexes(i) + 1
            cmbPlayersOnline.text = cmbPlayersOnline.List(cmbPlayersOnline.ListIndex)
            Exit Sub
        End If
    Next
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

    SendSetPlayerSprite Trim$(cmbPlayersOnline.text), Trim$(txtASprite.text)
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
    Dim N As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_MAPPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    If Len(Trim$(txtAMap.text)) < 1 Then Exit Sub
    If Not IsNumeric(Trim$(txtAMap.text)) Then Exit Sub
    
    N = CLng(Trim$(txtAMap.text))

    ' Check to make sure its a valid map #
    If N > 0 And N <= MAX_MAPS Then
        Call WarpTo(N)
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

Private Sub cmdAAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_ADMIN Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If

    ' Subscript out of range
    If Len(Trim$(cmbPlayersOnline.text)) < 2 Then Exit Sub
    If IsNumeric(Trim$(cmbPlayersOnline.text)) Or Not IsNumeric(Trim$(txtAAccess.text)) Then Exit Sub

    SendSetAccess Trim$(cmbPlayersOnline.text), CLng(Trim$(txtAAccess.text))
    Exit Sub
     
' Error handler
errorhandler:
    HandleError "cmdAAccess_Click", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdASpawn_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerAccess(MyIndex) < STAFF_DEVELOPER Then
        AddText "You have insufficent access to do this!", BrightRed
        Exit Sub
    End If
    
    SendSpawnItem scrlAItem.value, scrlAAmount.value
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
                If frmAdmin.Visible = True And GetForegroundWindow = frmAdmin.hwnd Then
                    Unload frmAdmin
                End If
            End If
    End Select
End Sub


Private Sub picAdmin_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picRefresh.Picture = LoadResPicture("REFRESH_UP", vbResBitmap)
End Sub

Private Sub picAdmin_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    refreshDown = False
End Sub

Private Sub picRefresh_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    refreshDown = True
    picRefresh.Picture = LoadResPicture("REFRESH_DOWN", vbResBitmap)
End Sub

Private Sub picRefresh_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Not refreshDown Then
        picRefresh.Picture = LoadResPicture("REFRESH_OVER", vbResBitmap)
    End If
End Sub

Private Sub picRefresh_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    refreshDown = False
    refreshingAdminList = True
    SendRequestPlayersOnline
End Sub

Private Sub scrlAItem_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    UpdateAdminScrollBar
    scrlAAmount.Enabled = False
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlAItem_Change", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlAAmount_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    lblAAmount.Caption = "Amount: " & scrlAAmount.value
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "scrlAAmount_Change", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub
Public Sub UpdatePlayersOnline()
    Dim players() As String, length As Long, i As Long, currentIgnore As Long
    Dim temp As String
    temp = cmbPlayersOnline.text
    cmbPlayersOnline.Clear
    cmbPlayersOnline.text = temp
    cmbPlayersOnline.AddItem ("----Staff: 1-----")
    
        ReDim Preserve ignoreIndexes(0)
        ignoreIndexes(0) = currentIgnore
        currentIgnore = currentIgnore + 1
    
    cmbPlayersOnline.AddItem (Trim(Player(MyIndex).name))
    
        currentIgnore = currentIgnore + 1
        
    players = Split(Trim(g_playersOnline), ",")
    length = UBound(players)
    If length >= 1 Then
        cmbPlayersOnline.AddItem ("----Players: " & length & "----")
        
            ReDim Preserve ignoreIndexes(1)
            ignoreIndexes(1) = currentIgnore
            currentIgnore = currentIgnore + 1
        For i = 0 To length
            If Trim(Player(MyIndex).name) = Trim(players(i)) Then
                GoTo end_loop
            End If
            cmbPlayersOnline.AddItem (Trim(players(i)))
            currentIgnore = currentIgnore + 1
end_loop:
        Next
    End If
    lblPlayers.Caption = "Players: " & length + 1 ' +1 counting ourselfs
End Sub
Private Sub Form_Load()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
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

Private Sub txtAAccess_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    txtAAccess.SelStart = Len(txtAAccess)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtAAccess_GotFocus", "frmAdmin", Err.Number, Err.Description, Err.Source, Err.HelpContext
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
