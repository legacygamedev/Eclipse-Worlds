VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmAdmin 
   BackColor       =   &H80000002&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Admin Panel"
   ClientHeight    =   8715
   ClientLeft      =   810
   ClientTop       =   330
   ClientWidth     =   2985
   Icon            =   "frmAdmin.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   581
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   199
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picPanel 
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
      Height          =   8655
      Left            =   60
      ScaleHeight     =   577
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   191
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   15
      Width           =   2865
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   9
         Left            =   540
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   53
         ToolTipText     =   "Change sprites via dbl click."
         Top             =   6225
         Width           =   420
      End
      Begin VB.PictureBox picRecentItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000001&
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   195
         ScaleHeight     =   450
         ScaleWidth      =   450
         TabIndex        =   52
         Top             =   6870
         Width           =   480
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   8
         Left            =   990
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   50
         ToolTipText     =   "Auto Life ?? "
         Top             =   5790
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   7
         Left            =   540
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Reset Scrolls"
         Top             =   5790
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   6
         Left            =   90
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Teleport scrolls"
         Top             =   5790
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   5
         Left            =   990
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "Spells, scrolls, magic."
         Top             =   5355
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   4
         Left            =   540
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   46
         ToolTipText     =   "No idea???"
         Top             =   5355
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   3
         Left            =   90
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   45
         ToolTipText     =   "Potions, elixirs, food."
         Top             =   5355
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   2
         Left            =   990
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   44
         ToolTipText     =   "Things you can wear"
         Top             =   4920
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   1
         Left            =   540
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   43
         ToolTipText     =   "Things without a type"
         Top             =   4920
         Width           =   420
      End
      Begin VB.OptionButton optCat 
         Height          =   420
         Index           =   0
         Left            =   90
         MaskColor       =   &H80000001&
         Style           =   1  'Graphical
         TabIndex        =   42
         ToolTipText     =   "Recently spawned items"
         Top             =   4920
         Width           =   420
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   255
         Left            =   765
         TabIndex        =   40
         Top             =   6990
         Width           =   480
         _ExtentX        =   847
         _ExtentY        =   450
         _Version        =   393216
         Orientation     =   1
         Enabled         =   0   'False
      End
      Begin VB.CommandButton cmdSpawnLast 
         Caption         =   "Spawn Recent"
         Enabled         =   0   'False
         Height          =   255
         Left            =   30
         TabIndex        =   39
         Top             =   7935
         Width           =   1380
      End
      Begin VB.TextBox txtLastAmount 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   30
         TabIndex        =   38
         Text            =   "Recent Amount"
         Top             =   7620
         Width           =   1350
      End
      Begin VB.PictureBox picSprite 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         ForeColor       =   &H80000008&
         Height          =   720
         Left            =   1830
         ScaleHeight     =   46
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   35
         Top             =   2160
         Width           =   510
      End
      Begin VB.TextBox txtSprite 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1800
         TabIndex        =   34
         Text            =   "0"
         Top             =   2880
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
         TabIndex        =   33
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
         TabIndex        =   32
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
         TabIndex        =   31
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
         TabIndex        =   25
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
         Left            =   1590
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   4500
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
         Left            =   1590
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   4230
         Width           =   1215
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
         TabIndex        =   21
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
         Left            =   1590
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   6645
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
         Left            =   1590
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3960
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
         TabIndex        =   18
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
         Left            =   1560
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   3720
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
         Left            =   165
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   4065
         Width           =   1155
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
         Left            =   1590
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   6390
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
         Left            =   1590
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   6120
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
         Left            =   1590
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   5850
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
         Left            =   1590
         TabIndex        =   12
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
         Left            =   1590
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   5040
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
         Left            =   1590
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   4770
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
         Left            =   180
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   3495
         Width           =   1140
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
         Left            =   165
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3780
         Width           =   1155
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
         Left            =   195
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   3210
         Width           =   1125
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
         Left            =   825
         TabIndex        =   0
         Top             =   2850
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
         TabIndex        =   6
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
         TabIndex        =   5
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
         TabIndex        =   4
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
         TabIndex        =   3
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
         Left            =   1590
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   5310
         Width           =   1215
      End
      Begin MSComCtl2.UpDown upSprite 
         Height          =   555
         Left            =   2430
         TabIndex        =   36
         Top             =   2280
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   979
         _Version        =   393216
         BuddyControl    =   "txtSprite"
         BuddyDispid     =   196615
         OrigLeft        =   3990
         OrigTop         =   1770
         OrigRight       =   4245
         OrigBottom      =   2265
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lblRecent 
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Recent"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   315
         Left            =   435
         TabIndex        =   51
         Top             =   6600
         Width           =   555
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item Name: None"
         Enabled         =   0   'False
         ForeColor       =   &H00800080&
         Height          =   195
         Left            =   -45
         TabIndex        =   41
         Top             =   7380
         Width           =   1470
      End
      Begin VB.Label lblCat 
         Alignment       =   2  'Center
         BackColor       =   &H80000002&
         BackStyle       =   0  'Transparent
         Caption         =   "Categories"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   238
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800080&
         Height          =   255
         Left            =   105
         TabIndex        =   37
         Top             =   4635
         Width           =   1260
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
         Left            =   45
         TabIndex        =   30
         Top             =   8235
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
         Left            =   270
         TabIndex        =   29
         Top             =   2475
         Width           =   975
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
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   1425
         TabIndex        =   28
         Top             =   3345
         Width           =   1410
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
         ForeColor       =   &H00800080&
         Height          =   285
         Left            =   150
         TabIndex        =   27
         Top             =   4380
         Width           =   1140
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
         TabIndex        =   26
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
      Begin VB.Line Line4 
         BorderColor     =   &H00800080&
         BorderWidth     =   3
         X1              =   8
         X2              =   89
         Y1              =   311
         Y2              =   311
      End
      Begin VB.Line Line2 
         BorderColor     =   &H000000C0&
         BorderWidth     =   3
         X1              =   103
         X2              =   186
         Y1              =   240
         Y2              =   240
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
         Left            =   195
         TabIndex        =   23
         Top             =   2880
         Width           =   735
      End
      Begin VB.Line Line1 
         BorderWidth     =   3
         X1              =   10
         X2              =   90
         Y1              =   183
         Y2              =   183
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
Private catSub As Boolean
Public lastIndex As Integer
Public currentCategory As String

Private Const WM_ChangeUIState As Long = &H127
Private Const UIS_HideRectangle As Integer = &H1
Private Const UIS_ShowRectangle As Integer = &H2
Private Const UISF_FocusRectangle As Integer = &H1


 
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, lparam As Any) As Long

 
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" ( _
'    ByVal hWnd As Long, ByVal wMsg As Long, _
'    ByVal wParam As Long, lParam As Any) As Long
'
'Private Function MakeLong(ByVal wLow As Integer, _
'    ByVal wHigh As Integer) As Long
'    MakeLong = wHigh * &H10000 + wLow
'End Function



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
    Dim i As Long
    If PlayerName = cmbPlayersOnline.text Then
        If Success = 0 Then
            For i = 0 To UBound(g_playersOnline)
                If InStr(1, g_playersOnline(i), PlayerName) Then
                    Mid(g_playersOnline(i), InStr(1, g_playersOnline(i), ":"), 2) = ":" & CurrentAccess
                    setAdminAccessLevel
                    
                    DisplayStatus Message, status.Error
                End If
            Next i
        ElseIf Success = 1 Then
            Mid(g_playersOnline(i), InStr(1, g_playersOnline(i), ":"), 2) = ":" & CurrentAccess
            setAdminAccessLevel
            
            DisplayStatus Message, status.Correct
        End If
    End If
    cmbPlayersOnline.Enabled = True
End Sub

Public Sub DisplayStatus(ByVal Msg As String, msgType As status)
    Select Case msgType
        Case status.Error:
            lblStatus.BackColor = &H8080FF
            lblStatus.Caption = Msg
        Case status.Correct:
            lblStatus.BackColor = &H80FF80
            lblStatus.Caption = Msg
        Case status.Neutral:
            lblStatus.BackColor = &H80FFFF
            lblStatus.Caption = Msg
        Case status.Info_:
            lblStatus.BackColor = &H8000000F
            lblStatus.Caption = Msg
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
    Dim accessLvl As String, tempTxt As String, i As Long
    
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

Private Sub Command1_Click()
    frmItemSpawner.Visible = True
End Sub

Private Sub cmdRecent_Click()
    frmItemSpawner.Visible = True
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case vbKeyInsert
            If Player(MyIndex).Access >= STAFF_MODERATOR Then
                If frmAdmin.Visible And GetForegroundWindow = frmAdmin.hWnd Then
                    Unload frmAdmin
                End If
            End If
    End Select
End Sub

Private Sub optCat_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Select Case Index
    
        Case 0
            lblCat.Caption = "Recent"
        Case 1
            lblCat.Caption = "None"
        Case 2
            lblCat.Caption = "Equipment"
        Case 3
            lblCat.Caption = "Consumable"
        Case 4
            lblCat.Caption = "Title"
        Case 5
            lblCat.Caption = "Spell"
        Case 6
            lblCat.Caption = "Teleport"
        Case 7
            lblCat.Caption = "Reset Stats"
        Case 8
            lblCat.Caption = "Auto Life"
        Case 9
            lblCat.Caption = "Change Sprite"
    End Select
End Sub


Public Sub optCat_MouseUp(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    If optCat(Index).Value = False Then
        optCat(Index).Picture = LoadResPicture(100 + Index, vbResBitmap)
        
    Else

        optCat(Index).Picture = LoadResPicture(110 + Index, vbResBitmap)

    Select Case Index
    
        Case 0
            currentCategory = "Recent"
        Case 1
            currentCategory = "None"
        Case 2
            currentCategory = "Equipment"
        Case 3
            currentCategory = "Consumable"
        Case 4
            currentCategory = "Title"
        Case 5
            currentCategory = "Spell"
        Case 6
            currentCategory = "Teleport"
        Case 7
            currentCategory = "Reset Stats"
        Case 8
            currentCategory = "Auto Life"
        Case 9
            currentCategory = "Change Sprite"
    End Select
        If lastIndex <> -1 Then
            If optCat(lastIndex).Value = False Then
                optCat(lastIndex).Picture = LoadResPicture(100 + lastIndex, vbResBitmap)
            End If
        End If
        
        If Button <> 0 Then
            If lastIndex = Index Then
                frmItemSpawner.Visible = False
                optCat(Index).Value = False
                optCat(Index).Picture = LoadResPicture(100 + lastIndex, vbResBitmap)
                lastIndex = -1
                Exit Sub
            Else
                frmItemSpawner.Visible = True
                frmItemSpawner.tabItems.Tabs(Index + 1).Selected = True
                BringWindowToTop (frmItemSpawner.hWnd)
            End If
        End If

        
        lastIndex = Index
    End If
End Sub

Private Sub picPanel_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    picRefresh.Picture = LoadResPicture("REFRESH_UP", vbResBitmap)
    lblCat.Caption = currentCategory
    
End Sub

Private Sub picPanel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
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
Public Sub styleButtons()
Dim i As Long, temp1 As Long, temp2 As Long
    For i = 0 To optCat.UBound
        optCat(i).Value = False
        optCat(i).Picture = LoadResPicture(100 + i, vbResBitmap)
    Next
    If Not catSub Then
        catSub = True
        For i = 0 To optCat.UBound
            SubClassHwnd optCat(i).hWnd
        Next
    End If
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
    
    lastIndex = -1
    styleButtons
        
    upSprite.max = NumCharacters
    upSprite.min = 0
    
    currentCategory = lblCat.Caption
    
    LastAdminSpriteTimer = timeGetTime

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

Private Function correctValue(ByRef textBox As textBox, ByRef valueToChange, min As Long, max As Long, Optional defaultVal As Long = 0) As Boolean
    Dim test As textBox, TempValue As String
    
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
        displayFieldStatus textBox, " field accepts only Numbers!" & vbCrLf & "Reverting to last correct value...", status.Correct
    Else
        textBox.text = CStr(valueToChange)
        displayFieldStatus textBox, " field is correct. Saving...", status.Correct
    End If
End Sub

Private Function verifyValue(txtBox As textBox, min As Long, max As Long)
    Dim Msg As String
    
    If (CLng(txtBox.text) >= min And CLng(txtBox.text) <= max) Then
        verifyValue = True
    Else
        Msg = " field accepts only values: " & CStr(min) & " < value < " & CStr(max) & "." & vbCrLf & "Reverting value..."
        displayFieldStatus txtBox, Msg, status.Error
        verifyValue = False
    End If
End Function
Public Sub displayFieldStatus(ByVal txtBox As textBox, ByVal Msg As String, msgType As status)
    lblStatus.Visible = True
    Select Case msgType

        Case status.Error:
            lblStatus.BackColor = &H8080FF
            lblStatus.Caption = Replace(txtBox.name, "txt", "") & Msg
        Case status.Correct:
            lblStatus.BackColor = &H80FF80
            lblStatus.Caption = Replace(txtBox.name, "txt", "") & Msg
        Case status.Neutral:
            lblStatus.BackColor = &H80FFFF
            lblStatus.Caption = Replace(txtBox.name, "txt", "") & Msg
        Case status.Info_:
            lblStatus.BackColor = &H8000000F
            lblStatus.Caption = Replace(txtBox.name, "txt", "") & Msg
    End Select
End Sub
Private Sub selectValue(ByRef textBox As textBox)
    textBox.SelStart = 0
    textBox.SelLength = Len(textBox.text)
End Sub

Private Sub txtSprite_Change()
    Dim i As Long
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
