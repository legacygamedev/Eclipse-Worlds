VERSION 5.00
Begin VB.Form frmEditor_Quest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   $"frmEditor_Quest.frx":0000
   ClientHeight    =   10170
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   18975
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10170
   ScaleWidth      =   18975
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmeTask 
      Caption         =   "Add a new action/task to complete for this Greeter"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4335
      Left            =   6480
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   7215
      Begin VB.CommandButton btnWarp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Warp the player"
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   100
         ToolTipText     =   "Show a message to the player."
         Top             =   2880
         Width           =   2775
      End
      Begin VB.CommandButton btnMsgPlayer 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Show message."
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   84
         ToolTipText     =   "Show a message to the player."
         Top             =   2400
         Width           =   2775
      End
      Begin VB.CommandButton btnAdjustStat 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Adjust Player Stat"
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   81
         ToolTipText     =   "Give or take stat values from the player such as Str/End/Exp/ect."
         Top             =   1920
         Width           =   2775
      End
      Begin VB.CommandButton btnProtect 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Spawn and protect ally."
         Enabled         =   0   'False
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   68
         ToolTipText     =   "Spawn an NPC to follow the player for so long, making all NPC's want to attack it, but the player must protect it for the quest."
         Top             =   2760
         Width           =   3015
      End
      Begin VB.CommandButton btnSkillLvl 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Obtain a skill level."
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   67
         ToolTipText     =   "Require that the player obtain a skill level between the progress of this quest"
         Top             =   2160
         Width           =   3015
      End
      Begin VB.CommandButton btnTaskCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Close"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   3600
         Width           =   6735
      End
      Begin VB.CommandButton btnTask_Kill 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Kill enemie(s)"
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   20
         ToolTipText     =   "Select an NPC the player needs to kill and an amount of times to kill it"
         Top             =   1560
         Width           =   3015
      End
      Begin VB.CommandButton btnTask_Gather 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Gather items"
         Height          =   495
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   19
         ToolTipText     =   "Select an item for the player to gather for the quest"
         Top             =   960
         Width           =   3015
      End
      Begin VB.CommandButton btnTakeItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Take an item."
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   82
         ToolTipText     =   "Take an item from the player"
         Top             =   1320
         Width           =   2775
      End
      Begin VB.CommandButton btnGiveItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Give an item."
         Height          =   375
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   83
         ToolTipText     =   "Give the player an item"
         Top             =   960
         Width           =   2775
      End
      Begin VB.Line Line9 
         X1              =   240
         X2              =   6960
         Y1              =   3480
         Y2              =   3480
      End
      Begin VB.Line Line8 
         X1              =   3720
         X2              =   3720
         Y1              =   240
         Y2              =   3480
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Actions"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   86
         Top             =   480
         Width           =   2775
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Tasks"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   480
         TabIndex        =   85
         Top             =   480
         Width           =   2775
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   6960
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Frame fmeModify 
      Caption         =   "Adjust Player Stats"
      Height          =   3855
      Left            =   4560
      TabIndex        =   69
      Top             =   6600
      Visible         =   0   'False
      Width           =   4575
      Begin VB.ComboBox cboItem 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":0089
         Left            =   240
         List            =   "frmEditor_Quest.frx":008B
         Style           =   2  'Dropdown List
         TabIndex        =   115
         Top             =   2160
         Width           =   4215
      End
      Begin VB.OptionButton opSkillEXP 
         Caption         =   "Skill EXP"
         Height          =   255
         Left            =   240
         TabIndex        =   114
         ToolTipText     =   "Select to modify the player's Level"
         Top             =   1800
         Width           =   4095
      End
      Begin VB.OptionButton opSkill 
         Caption         =   "Skill Level"
         Height          =   255
         Left            =   240
         TabIndex        =   113
         ToolTipText     =   "Select to modify the player's Level"
         Top             =   1560
         Width           =   4095
      End
      Begin VB.OptionButton opStatP 
         Caption         =   "Stat Points"
         Height          =   255
         Left            =   240
         TabIndex        =   112
         ToolTipText     =   "Select to modify the player's Level"
         Top             =   1320
         Width           =   4095
      End
      Begin VB.OptionButton opStatEXP 
         Caption         =   "Stat EXP"
         Height          =   255
         Left            =   240
         TabIndex        =   111
         ToolTipText     =   "Select to modify the player's Level"
         Top             =   1080
         Width           =   4095
      End
      Begin VB.OptionButton opStat 
         Caption         =   "Stat"
         Height          =   255
         Left            =   240
         TabIndex        =   110
         ToolTipText     =   "Select to modify the player's Level"
         Top             =   840
         Width           =   4095
      End
      Begin VB.CommandButton btnModCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   76
         Top             =   3360
         Width           =   1455
      End
      Begin VB.CommandButton btnModAccept 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   75
         Top             =   3360
         Width           =   2775
      End
      Begin VB.CheckBox chkSet 
         Caption         =   "Set the value instead of adding/subtracting"
         Height          =   255
         Left            =   240
         TabIndex        =   74
         ToolTipText     =   "This option will decide whether we set the amount or add/subtract to the current amount."
         Top             =   2520
         Width           =   4215
      End
      Begin VB.HScrollBar scrlModify 
         Height          =   255
         LargeChange     =   25
         Left            =   120
         Min             =   -32767
         TabIndex        =   73
         Top             =   3000
         Width           =   4335
      End
      Begin VB.OptionButton opLvl 
         Caption         =   "Level"
         Height          =   255
         Left            =   240
         TabIndex        =   71
         ToolTipText     =   "Select to modify the player's Level"
         Top             =   600
         Width           =   4095
      End
      Begin VB.OptionButton opEXP 
         Caption         =   "EXP"
         Height          =   255
         Left            =   240
         TabIndex        =   70
         ToolTipText     =   "Select to modify the player's EXP"
         Top             =   360
         Width           =   4095
      End
      Begin VB.Label lblModify 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Amount to modify: 0"
         Height          =   255
         Left            =   240
         TabIndex        =   72
         Top             =   2760
         Width           =   4095
      End
   End
   Begin VB.Frame fmeWarp 
      Caption         =   "Select map and location to warp to"
      Height          =   3015
      Left            =   15000
      TabIndex        =   101
      Top             =   120
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton btnCloseWarp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   109
         Top             =   2520
         Width           =   1335
      End
      Begin VB.CommandButton btnAddWarp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   108
         Top             =   2520
         Width           =   1335
      End
      Begin VB.HScrollBar scrlMapY 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   107
         Top             =   2040
         Width           =   3495
      End
      Begin VB.HScrollBar scrlMapX 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         TabIndex        =   105
         Top             =   1440
         Width           =   3495
      End
      Begin VB.HScrollBar scrlMap 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         Max             =   255
         Min             =   1
         TabIndex        =   103
         Top             =   600
         Value           =   1
         Width           =   3495
      End
      Begin VB.Label lblMapY 
         Caption         =   "Y: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   106
         Top             =   1800
         Width           =   3495
      End
      Begin VB.Label lblMapX 
         Caption         =   "X: 0"
         Height          =   255
         Left            =   120
         TabIndex        =   104
         Top             =   1200
         Width           =   3495
      End
      Begin VB.Label lblMap 
         Alignment       =   2  'Center
         Caption         =   "Map: 1"
         Height          =   255
         Left            =   120
         TabIndex        =   102
         Top             =   360
         Width           =   3495
      End
   End
   Begin VB.Frame fmeSelectItem 
      Caption         =   "Select Item To Give"
      Height          =   2655
      Left            =   15000
      TabIndex        =   51
      Top             =   3120
      Visible         =   0   'False
      Width           =   3615
      Begin VB.CheckBox chkTake 
         Caption         =   "Take the item the player gathers?"
         Height          =   255
         Left            =   120
         TabIndex        =   78
         Top             =   1800
         Visible         =   0   'False
         Width           =   3375
      End
      Begin VB.HScrollBar scrlItemAmount 
         Height          =   255
         LargeChange     =   5
         Left            =   120
         TabIndex        =   57
         Top             =   1440
         Width           =   3375
      End
      Begin VB.CommandButton btnAddItem 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   55
         Top             =   2160
         Width           =   1335
      End
      Begin VB.CommandButton btnAItemCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   2160
         Width           =   1335
      End
      Begin VB.HScrollBar scrlItem 
         Height          =   255
         Left            =   120
         TabIndex        =   53
         Top             =   840
         Width           =   3375
      End
      Begin VB.Label lblAmount 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   1200
         Width           =   3360
      End
      Begin VB.Label lblItem 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Item: 0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         TabIndex        =   52
         Top             =   240
         Width           =   3360
      End
   End
   Begin VB.Frame fmeCLI 
      Caption         =   "Add a new NPC/Event the player will need to meet with"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   15600
      TabIndex        =   14
      Top             =   2760
      Visible         =   0   'False
      Width           =   5055
      Begin VB.ComboBox cmbNPC 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   95
         Top             =   1440
         Width           =   4815
      End
      Begin VB.HScrollBar scrlKillAmnt 
         Height          =   255
         LargeChange     =   10
         Left            =   120
         TabIndex        =   79
         Top             =   2760
         Width           =   4815
      End
      Begin VB.CommandButton btnCLICancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   45
         Top             =   3120
         Width           =   2175
      End
      Begin VB.OptionButton opEvent 
         Caption         =   "Events"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2280
         Width           =   4455
      End
      Begin VB.OptionButton opNPC 
         Caption         =   "NPC's"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1920
         Value           =   -1  'True
         Width           =   4455
      End
      Begin VB.CommandButton btnAddCLI 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   3120
         Width           =   2175
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   $"frmEditor_Quest.frx":008D
         Height          =   855
         Left            =   120
         TabIndex        =   96
         Top             =   480
         Width           =   4815
      End
      Begin VB.Label lblKillAmnt 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Amount: 0"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   80
         Top             =   2520
         Width           =   4800
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   4920
         Y1              =   240
         Y2              =   240
      End
   End
   Begin VB.Frame fmeShowMsg 
      Caption         =   "Show player a message"
      Height          =   2775
      Left            =   15000
      TabIndex        =   61
      Top             =   6480
      Visible         =   0   'False
      Width           =   4455
      Begin VB.ComboBox cmbColor 
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":017F
         Left            =   720
         List            =   "frmEditor_Quest.frx":01B9
         Style           =   2  'Dropdown List
         TabIndex        =   97
         Top             =   1800
         Width           =   3615
      End
      Begin VB.CommandButton btnMsgCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2280
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CommandButton btnMsgAccept 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   2160
         Width           =   2055
      End
      Begin VB.CheckBox chkRes 
         Caption         =   "This is the response if the last task is not done."
         Height          =   255
         Left            =   120
         TabIndex        =   94
         ToolTipText     =   "Check this box if this message will be shown to the player if the first task before this message isn't completed yet."
         Top             =   1560
         Width           =   4215
      End
      Begin VB.CheckBox chkStart 
         BackColor       =   &H00E0E0E0&
         Caption         =   "This is just a placeholder"
         Enabled         =   0   'False
         Height          =   255
         Left            =   120
         TabIndex        =   77
         ToolTipText     =   "Check this box if this message will be the first message the player sees before starting the quest."
         Top             =   2280
         Width           =   4215
      End
      Begin VB.TextBox txtMsg 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   62
         ToolTipText     =   "Enter a message that will be shown to the player."
         Top             =   240
         Width           =   4215
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Color:"
         Height          =   255
         Left            =   120
         TabIndex        =   98
         Top             =   1850
         Width           =   495
      End
   End
   Begin VB.Frame fmeReq 
      Caption         =   "Quest Requirements"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   6480
      TabIndex        =   23
      Top             =   4200
      Visible         =   0   'False
      Width           =   8295
      Begin VB.HScrollBar scrlSkill 
         Height          =   255
         Left            =   5280
         Max             =   255
         TabIndex        =   66
         Top             =   600
         Width           =   1575
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   5
         LargeChange     =   10
         Left            =   6360
         TabIndex        =   34
         Top             =   2160
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   4
         LargeChange     =   10
         Left            =   3720
         TabIndex        =   33
         Top             =   2160
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   3
         LargeChange     =   10
         Left            =   1200
         TabIndex        =   32
         Top             =   2160
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   2
         LargeChange     =   10
         Left            =   6360
         TabIndex        =   31
         Top             =   1680
         Width           =   1455
      End
      Begin VB.HScrollBar scrlStatReq 
         Height          =   255
         Index           =   1
         LargeChange     =   10
         Left            =   3720
         TabIndex        =   30
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ComboBox cmbClassReq 
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":025A
         Left            =   6360
         List            =   "frmEditor_Quest.frx":025C
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   1200
         Width           =   1695
      End
      Begin VB.HScrollBar scrlAccessReq 
         Height          =   255
         Left            =   1200
         Max             =   5
         TabIndex        =   28
         Top             =   1200
         Width           =   1455
      End
      Begin VB.ComboBox cmbGenderReq 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":025E
         Left            =   3720
         List            =   "frmEditor_Quest.frx":026B
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   1200
         Width           =   1455
      End
      Begin VB.HScrollBar scrlLevelReq 
         Height          =   255
         LargeChange     =   10
         Left            =   1200
         Max             =   255
         TabIndex        =   26
         Top             =   1680
         Width           =   1455
      End
      Begin VB.ComboBox cmbSkillReq 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":0283
         Left            =   1200
         List            =   "frmEditor_Quest.frx":0285
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   2415
      End
      Begin VB.CommandButton btnReqOk 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   2760
         Width           =   8055
      End
      Begin VB.Label lblSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill Level: 0"
         Height          =   195
         Left            =   3960
         TabIndex        =   65
         Top             =   600
         Width           =   900
      End
      Begin VB.Line Line6 
         X1              =   120
         X2              =   8160
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Spi: 0"
         Height          =   195
         Index           =   5
         Left            =   5400
         TabIndex        =   44
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Agi: 0"
         Height          =   195
         Index           =   4
         Left            =   2880
         TabIndex        =   43
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   405
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Int: 0"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   42
         Top             =   2160
         UseMnemonic     =   0   'False
         Width           =   360
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "End: 0"
         Height          =   195
         Index           =   2
         Left            =   5400
         TabIndex        =   41
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   465
      End
      Begin VB.Label lblStatReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Str: 0"
         Height          =   195
         Index           =   1
         Left            =   2880
         TabIndex        =   40
         Top             =   1680
         UseMnemonic     =   0   'False
         Width           =   375
      End
      Begin VB.Label lblClassReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Class:"
         Height          =   195
         Left            =   5400
         TabIndex        =   39
         Top             =   1200
         Width           =   420
      End
      Begin VB.Label lblAccessReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Access: 0"
         Height          =   195
         Left            =   240
         TabIndex        =   38
         Top             =   1200
         Width           =   705
      End
      Begin VB.Label lblGenderReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gender:"
         Height          =   195
         Left            =   2880
         TabIndex        =   37
         Top             =   1200
         Width           =   570
      End
      Begin VB.Label lblLevelReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Level: 0"
         Height          =   195
         Left            =   240
         TabIndex        =   36
         Top             =   1680
         Width           =   570
      End
      Begin VB.Label lblSkillReq 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill:"
         Height          =   195
         Left            =   240
         TabIndex        =   35
         Top             =   600
         Width           =   330
      End
   End
   Begin VB.Frame fmeObtainSKill 
      Caption         =   "Select a skill and skill level for the player to obtain."
      Height          =   1935
      Left            =   16200
      TabIndex        =   87
      Top             =   0
      Visible         =   0   'False
      Width           =   3855
      Begin VB.CommandButton btnObAccept 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Accept"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   93
         Top             =   1320
         Width           =   1695
      End
      Begin VB.CommandButton btnObCancel 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   2040
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   1320
         Width           =   1695
      End
      Begin VB.HScrollBar scrlObSkill 
         Height          =   255
         LargeChange     =   5
         Left            =   1560
         Max             =   255
         TabIndex        =   90
         Top             =   840
         Width           =   2055
      End
      Begin VB.ComboBox cmbObSKill 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmEditor_Quest.frx":0287
         Left            =   1200
         List            =   "frmEditor_Quest.frx":0289
         Style           =   2  'Dropdown List
         TabIndex        =   88
         ToolTipText     =   "Select the specific skill the player will need to reach a level for."
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label lblObSkill 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill Level:"
         Height          =   195
         Left            =   240
         TabIndex        =   91
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Skill:"
         Height          =   195
         Left            =   360
         TabIndex        =   89
         Top             =   360
         Width           =   330
      End
   End
   Begin VB.Frame fmeMoveItem 
      Caption         =   "Move List Items"
      Height          =   2415
      Left            =   1560
      TabIndex        =   47
      Top             =   4320
      Visible         =   0   'False
      Width           =   1815
      Begin VB.CommandButton btnDeleteAction 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Delete"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   117
         ToolTipText     =   "Move the currently selected list item down."
         Top             =   1920
         Width           =   1575
      End
      Begin VB.CommandButton btnEditAction 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   116
         ToolTipText     =   "Move the currently selected list item down."
         Top             =   1440
         Width           =   1575
      End
      Begin VB.CommandButton btnHide 
         BackColor       =   &H00FFFFFF&
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1560
         Style           =   1  'Graphical
         TabIndex        =   50
         Top             =   0
         Width           =   255
      End
      Begin VB.CommandButton btnDown 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Move Item Down"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   49
         ToolTipText     =   "Move the currently selected list item down."
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton btnUp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Move Item Up"
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "Move the currently selected list item up."
         Top             =   360
         Width           =   1575
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   1680
         Y1              =   1320
         Y2              =   1320
      End
   End
   Begin VB.CommandButton btnDelete 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   60
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton btnCancel 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   12120
      Style           =   1  'Graphical
      TabIndex        =   59
      Top             =   8280
      Width           =   2535
   End
   Begin VB.CommandButton btnSave 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   58
      Top             =   8280
      Width           =   5775
   End
   Begin VB.Frame fraNPC 
      Caption         =   "Quest List"
      Height          =   8655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3135
      Begin VB.ListBox lstIndex 
         Height          =   7665
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   2895
      End
      Begin VB.CommandButton cmdCopy 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Copy"
         Height          =   315
         Left            =   1680
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   615
      End
      Begin VB.TextBox txtSearch 
         CausesValidation=   0   'False
         Height          =   270
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
      Begin VB.CommandButton cmdPaste 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Paste"
         Height          =   315
         Left            =   2400
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Quest Details"
      Height          =   2415
      Left            =   3240
      TabIndex        =   5
      Top             =   0
      Width           =   4455
      Begin VB.TextBox txtDesc 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1305
         Left            =   120
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   9
         ToolTipText     =   "The description of the quest.  Will be seen in the player's quest viewer."
         Top             =   960
         Width           =   4215
      End
      Begin VB.TextBox txtName 
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
         Left            =   720
         MaxLength       =   40
         ScrollBars      =   1  'Horizontal
         TabIndex        =   6
         ToolTipText     =   "The name of the quest.  Will be see in the player's quest viewer."
         Top             =   360
         Width           =   3615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Description:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   720
         UseMnemonic     =   0   'False
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   360
         UseMnemonic     =   0   'False
         Width           =   465
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Quest Editing"
      Height          =   5775
      Left            =   3240
      TabIndex        =   11
      Top             =   2400
      Width           =   11535
      Begin VB.ListBox CLI 
         Height          =   5325
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "List of Greeters AKA NPC's the player will need to meet with throughout the quest.  Right click for more options."
         Top             =   360
         Width           =   2895
      End
      Begin VB.ListBox lstTasks 
         Height          =   5325
         Left            =   3240
         TabIndex        =   13
         ToolTipText     =   "List of all the actions and tasks that will be completed for the selected Greeter.   Right click for more options."
         Top             =   360
         Width           =   8175
      End
      Begin VB.Line Line2 
         X1              =   3120
         X2              =   3120
         Y1              =   600
         Y2              =   5520
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Editing Features"
      Height          =   2415
      Left            =   7800
      TabIndex        =   10
      Top             =   0
      Width           =   6975
      Begin VB.CheckBox chkRetake 
         Caption         =   "Can this quest be retaken?"
         Height          =   255
         Left            =   120
         TabIndex        =   99
         ToolTipText     =   "This sets whether or not the quest can be retaken again and again after completion."
         Top             =   1320
         Width           =   4095
      End
      Begin VB.Timer tmrMsg 
         Interval        =   3000
         Left            =   1800
         Top             =   240
      End
      Begin VB.CommandButton btnReq 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Edit Requirements"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   120
         Style           =   1  'Graphical
         TabIndex        =   22
         ToolTipText     =   "Edit the requirements for the player to be able to start the quest."
         Top             =   840
         Width           =   6735
      End
      Begin VB.Label lblMsg 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Right-Click a list below for additional options."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   11.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808000&
         Height          =   315
         Left            =   120
         TabIndex        =   46
         Top             =   360
         Width           =   6675
      End
   End
   Begin VB.Menu mnuCLI 
      Caption         =   "CLIMenu"
      Visible         =   0   'False
      Begin VB.Menu mnuEdit 
         Caption         =   "Edit Item"
      End
      Begin VB.Menu mnuACLI 
         Caption         =   "Add Greeter"
      End
      Begin VB.Menu mnuRCLI 
         Caption         =   "Remove Greeter"
      End
      Begin VB.Menu mnuAAction 
         Caption         =   "Add Action/Task"
      End
      Begin VB.Menu mnuRTask 
         Caption         =   "Remove Action/Task"
      End
   End
End
Attribute VB_Name = "frmEditor_Quest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private CLIHasFocus As Boolean
Private Editing_CLI As Boolean
Private Editing_CLI_Index As Long
Private Editing_Task As Boolean
Private Editing_Task_Index As Long
Private Gather As Boolean
Private GiveItem As Boolean
Private TakeItem As Boolean
Private KillNPC As Boolean
Private TmpIndex As Long

Private Sub btnAddCLI_Click()
Dim Index As Long, I As Long, tmpID As Long, TempStr As String, NPCIndex As Long
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    tmpStr = cmbNPC.List(cmbNPC.ListIndex)
    NPCIndex = 0
    For I = 1 To MAX_NPCS
        If Trim$(NPC(I).Name) = tmpStr Then
            NPCIndex = I
            Exit For
        End If
    Next I
    If Not NPCIndex > 0 Then Exit Sub
    
    With Quest(EditorIndex)
        If Not KillNPC Then
            If Not Editing_CLI Then
                tmpID = Editing_CLI_Index
                
                .Max_CLI = .Max_CLI + 1
                Index = .Max_CLI
                ReDim Preserve .CLI(1 To Index)
                
                .CLI(Index).ItemIndex = NPCIndex
                .CLI(Index).isNPC = opNPC.Value
                
                'add in a start message automatically for the first CLI element created.
                If Index = 1 Then
                    .CLI(Index).Max_Actions = .CLI(Index).Max_Actions + 1
                    I = .CLI(Index).Max_Actions
                    ReDim Preserve .CLI(Index).Action(1 To I)
                    .CLI(Index).Action(I).ActionID = ACTION_SHOWMSG
                    .CLI(Index).Action(I).MainData = vbChecked
                    .CLI(Index).Action(I).TextHolder = "Double click to edit the start message."
                End If
            Else
                .CLI(Editing_CLI_Index).ItemIndex = NPCIndex
                .CLI(Editing_CLI_Index).isNPC = opNPC.Value
                Editing_CLI_Index = 0
                Editing_CLI = False
            End If
        Else
            Index = CLI.ListIndex + 1
            If Index < 1 Then Exit Sub
            If Editing_Task Then
                tmpID = Editing_Task_Index
            Else
                .CLI(Index).Max_Actions = .CLI(Index).Max_Actions + 1
                ReDim Preserve .CLI(Index).Action(1 To .CLI(Index).Max_Actions)
                tmpID = .CLI(Index).Max_Actions
            End If
            
            .CLI(Index).Action(tmpID).ActionID = TASK_KILL
            .CLI(Index).Action(tmpID).MainData = NPCIndex
            .CLI(Index).Action(tmpID).SecondaryData = opNPC.Value
            .CLI(Index).Action(tmpID).Amount = scrlKillAmnt.Value
            Editing_Task_Index = 0
            Editing_Task = False
        End If
    End With
    
    QuestEditorInit
    If Not Editing_CLI Then CLI.ListIndex = CLI.ListCount - 1 Else CLI.ListIndex = Editing_CLI_Index
    fmeCLI.Visible = False
    ResetEditButtons
    Editing_CLI_Index = 0
    Editing_CLI = False
    Call ResetEditButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "btnAddCLI_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub btnAddItem_Click()
Dim Index As Long, Amnt As Long, Itm As Long, id As Long, I As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Index = CLI.ListIndex + 1
    Itm = scrlItem.Value
    Amnt = scrlItemAmount.Value
    If Index < 1 Then Exit Sub
    If Itm < 1 Then Exit Sub
    If Amnt < 1 Then Exit Sub
    
    If Gather Then
        id = TASK_GATHER
    ElseIf GiveItem Then
        id = ACTION_GIVE_ITEM
    ElseIf TakeItem Then
        id = ACTION_TAKE_ITEM
    End If
    
    'add the item to the list
    
    With Quest(EditorIndex).CLI(Index)
        If Editing_Task Then
            I = Editing_Task_Index
        Else
            .Max_Actions = .Max_Actions + 1
            ReDim Preserve .Action(1 To .Max_Actions)
            I = .Max_Actions
        End If
        
        .Action(I).ActionID = id
        .Action(I).MainData = Itm
        .Action(I).Amount = Amnt
        .Action(I).SecondaryData = chkTake.Value
        Editing_Task_Index = 0
        Editing_Task = False
        
        Call QuestEditorInitCLI
        Gather = False
        TakeItem = False
        GiveItem = False
        CLI.ListIndex = Index - 1
        fmeSelectItem.Visible = False
        Call ResetEditButtons
    End With
    
' Error handler
ErrorHandler:
    HandleError "btnAddItem_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub btnAddWarp_Click()
Dim Index As Long, X As Long, Y As Long, MapNum As Long, I As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Index = CLI.ListIndex + 1
    MapNum = scrlMap.Value
    X = scrlMapX.Value
    Y = scrlMapY.Value
    If Index < 1 Then Exit Sub
    If MapNum < 1 Then Exit Sub
    
    'add the item to the list
    
    With Quest(EditorIndex).CLI(Index)
        If Editing_Task Then
            I = Editing_Task_Index
        Else
            .Max_Actions = .Max_Actions + 1
            ReDim Preserve .Action(1 To .Max_Actions)
            I = .Max_Actions
        End If
        
        .Action(I).ActionID = ACTION_WARP
        .Action(I).Amount = MapNum
        .Action(I).MainData = X
        .Action(I).SecondaryData = Y
        Editing_Task_Index = 0
        Editing_Task = False
        
        Call QuestEditorInitCLI
        CLI.ListIndex = Index - 1
        fmeWarp.Visible = False
        Call ResetEditButtons
    End With
    
' Error handler
ErrorHandler:
    HandleError "btnAddWarp_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub btnAdjustStat_Click()
Dim Index As Long
    Index = CLI.ListIndex + 1
    If Index < 1 Then Exit Sub
    
    'open the panel to give an item.
    fmeModify.Visible = True
    Call BTF(fmeModify)
    cboItem.Clear
    cboItem.AddItem "None"
    cboItem.ListIndex = 0
End Sub

Private Sub btnAItemCancel_Click()
    scrlItem.Value = 0
    scrlItemAmount.Value = 0
    chkTake.Value = vbUnchecked
    fmeSelectItem.Visible = False
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
End Sub

Private Sub btnCancel_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_ITEMS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Unload frmEditor_Quest
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "btnCancel_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub btnCLICancel_Click()
    fmeCLI.Visible = False
    lblKillAmnt.Visible = False
    scrlKillAmnt.Value = 0
    scrlKillAmnt.Visible = False
    Editing_CLI = False
    Editing_CLI_Index = 0
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
End Sub

Private Sub btnCloseWarp_Click()
    scrlMap.Value = 1
    scrlMapX.Value = 1
    scrlMapY.Value = 1
    fmeWarp.Visible = False
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
End Sub

Private Sub btnDelete_Click()
Dim TmpIndex As Long
    
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    ClearQuest EditorIndex
    
    TmpIndex = lstIndex.ListIndex
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Trim$(Quest(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = TmpIndex

    QuestEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "btnDelete_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub btnDeleteAction_Click()
    Call mnuRTask_Click
    fmeMoveItem.Visible = False
End Sub

Private Sub btnDown_Click()
Dim tempSel As Long, tempSel2 As Long
    If CLIHasFocus Then
        'move item up within the CLI
        tempSel = CLI.ListIndex
        If tempSel < 0 Or tempSel > CLI.ListCount - 1 Then
            btnDown.Enabled = False
            Exit Sub
        End If
        If Not CLI.ListCount > 1 Then Exit Sub
        
        Call MoveListItem(LIST_CLI, EditorIndex, 0, tempSel + 1, 1)
        CLI.ListIndex = tempSel + 1
    Else
        'move item up within the Task List
        tempSel = lstTasks.ListIndex
        If tempSel < 0 Or tempSel > lstTasks.ListCount - 1 Then
            btnDown.Enabled = False
            Exit Sub
        End If
        If CLI.ListCount < 0 Then Exit Sub
        tempSel2 = CLI.ListIndex
        If tempSel2 < 0 Then Exit Sub
        
        'tempsel/2 is +1 because the array for the data starts at 1 whereas the listbox starts at 0
        Call MoveListItem(LIST_TASK, EditorIndex, tempSel2 + 1, tempSel + 1, 1)
        CLI.ListIndex = tempSel2
        lstTasks.ListIndex = tempSel + 1
    End If
End Sub

Private Sub btnEditAction_Click()
    Call mnuEdit_Click
    fmeMoveItem.Visible = False
End Sub

Private Sub btnGiveItem_Click()
Dim Index As Long
    Index = CLI.ListIndex + 1
    If Index < 1 Then Exit Sub
    
    'open the panel to give an item.
    GiveItem = True
    TakeItem = False
    Gather = False
    fmeSelectItem.Caption = "Select an item to give."
    chkTake.Visible = False
    fmeSelectItem.Visible = True
    Call BTF(fmeSelectItem)
End Sub

Private Sub btnHide_Click()
    fmeMoveItem.Visible = False
End Sub

Private Sub btnModAccept_Click()
Dim Index As Long, id As Long
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
        
        Index = CLI.ListIndex + 1
        If Index < 1 Then
            QMsg ("Please be sure a greeter is selected.")
            Exit Sub
        End If
        If chkSet.Value = vbChecked And scrlModify.Value < 1 Then Exit Sub
        If cboItem.ListIndex < 0 Then Exit Sub
        
        With Quest(EditorIndex).CLI(Index)
            If Editing_Task Then
                id = Editing_Task_Index
            Else
                .Max_Actions = .Max_Actions + 1
                ReDim Preserve .Action(1 To .Max_Actions)
                id = .Max_Actions
            End If
                
            If opEXP.Value = True Then
                .Action(id).ActionID = ACTION_ADJUST_EXP
            ElseIf opLvl.Value = True Then
                .Action(id).ActionID = ACTION_ADJUST_LVL
            ElseIf opStat.Value = True Then
                .Action(id).ActionID = ACTION_ADJUST_STAT_LVL
            ElseIf opSkill.Value = True Then
                .Action(id).ActionID = ACTION_ADJUST_SKILL_LVL
            ElseIf opSkillEXP.Value = True Then
                .Action(id).ActionID = ACTION_ADJUST_SKILL_EXP
            ElseIf opStatP.Value = True Then
                .Action(id).ActionID = ACTION_ADJUST_STAT_POINTS
            End If
            
            .Action(id).Amount = scrlModify.Value
            .Action(id).MainData = chkSet.Value
            .Action(id).SecondaryData = cboItem.ListIndex
            Editing_Task_Index = 0
            Editing_Task = False
        End With
        
        chkSet.Value = vbUnchecked
        scrlModify.Value = 0
        opEXP.Value = True
        opEXP.Value = False
        Call QuestEditorInitCLI
        CLI.ListIndex = Index - 1
        fmeModify.Visible = False
        Call ResetEditButtons
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "btnModAccept_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub btnModCancel_Click()
    chkSet.Value = vbUnchecked
    opEXP.Value = False
    opLvl.Value = False
    scrlModify.Value = 0
    fmeModify.Visible = False
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
End Sub

Private Sub btnMsgAccept_Click()
Dim Index As Long, Msg As String
Dim I As Long, II As Long, III As Long, id As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Index = CLI.ListIndex + 1
    Msg = txtMsg.text
    If Index < 1 Then Exit Sub
    If Len(Msg) < 1 Then
        Call QMsg("Please type a message to show the player.")
        Exit Sub
    End If
    If cmbColor.ListIndex < 0 Then
        Call QMsg("Please select a message color.")
        Exit Sub
    End If
    txtMsg.text = vbNullString
    
    'add the item to the list
    With Quest(EditorIndex).CLI(Index)
        If Editing_Task Then
            id = Editing_Task_Index
        Else
            .Max_Actions = .Max_Actions + 1
            ReDim Preserve .Action(1 To .Max_Actions)
            id = .Max_Actions
        End If
        
        .Action(id).ActionID = ACTION_SHOWMSG
        .Action(id).MainData = chkStart.Value
        .Action(id).SecondaryData = chkRes.Value
        .Action(id).TertiaryData = cmbColor.ListIndex
        .Action(id).TextHolder = Msg
        Editing_Task_Index = 0
        Editing_Task = False
            
        CLI.ListIndex = Index - 1
        fmeShowMsg.Visible = False
        chkStart.Value = vbUnchecked
        chkRes.Value = vbUnchecked
        Call QuestEditorInitCLI
        Call ResetEditButtons
    End With
    
' Error handler
ErrorHandler:
    HandleError "btnMsgAccept_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub btnMsgCancel_Click()
    chkStart.Value = vbUnchecked
    chkRes.Value = vbUnchecked
    txtMsg.text = vbNullString
    fmeShowMsg.Visible = False
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
End Sub

Private Sub btnMsgPlayer_Click()
Dim Index As Long
    Index = CLI.ListIndex + 1
    If Index < 1 Then Exit Sub
    
    'open the panel to give an item.
    Call CheckResponseMsg(EditorIndex, Index, Quest(EditorIndex).CLI(Index).Max_Actions)
    fmeShowMsg.Visible = True
    Call BTF(fmeShowMsg)
End Sub

Private Sub btnObAccept_Click()
Dim Index As Long, Amnt As Long, id As Long, SkillID As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Index = CLI.ListIndex + 1
    SkillID = cmbObSKill.ListIndex + 1
    Amnt = scrlObSkill.Value
    If Index < 1 Then Exit Sub
    If SkillID < 1 Then Exit Sub
    If Amnt < 1 Then Exit Sub
    
    'add the item to the list
    
    With Quest(EditorIndex).CLI(Index)
        If Editing_Task Then
            id = Editing_Task_Index
        Else
            .Max_Actions = .Max_Actions + 1
            ReDim Preserve .Action(1 To .Max_Actions)
            id = .Max_Actions
        End If
        
        .Action(id).ActionID = TASK_GETSKILL
        .Action(id).MainData = SkillID
        .Action(id).Amount = Amnt
        Editing_Task_Index = 0
        Editing_Task = False
        
        Call QuestEditorInitCLI
        CLI.ListIndex = Index - 1
        fmeObtainSKill.Visible = False
        Call ResetEditButtons
    End With
    
' Error handler
ErrorHandler:
    HandleError "btnAddItem_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub btnObCancel_Click()
    cmbObSKill.ListIndex = 0
    scrlObSkill.Value = 0
    fmeObtainSKill.Visible = False
    Editing_Task = False
    Editing_Task_Index = 0
    Call ResetEditButtons
End Sub

Private Sub btnReq_Click()
    fmeReq.Visible = True
    Call BTF(fmeReq)
    fmeMoveItem.Visible = False
    DisableEditButtons
End Sub

Private Sub btnReqOk_Click()
    fmeReq.Visible = False
    ResetEditButtons
End Sub

Private Sub btnSave_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    EditorSave = True
    Call QuestEditorSave
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "btnSave_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub btnSkillLvl_Click()
    fmeObtainSKill.Visible = True
    Call BTF(fmeObtainSKill)
End Sub

Private Sub btnTakeItem_Click()
Dim Index As Long
    Index = CLI.ListIndex + 1
    If Index < 1 Then Exit Sub
    
    'open the panel to give an item.
    TakeItem = True
    GiveItem = False
    Gather = False
    fmeSelectItem.Caption = "Select an item to take."
    chkTake.Visible = False
    fmeSelectItem.Visible = True
    Call BTF(fmeSelectItem)
End Sub

Private Sub btnTask_Gather_Click()
Dim Index As Long
    Index = CLI.ListIndex + 1
    If Index < 1 Then Exit Sub
    
    'open the panel to give an item.
    Gather = True
    GiveItem = False
    TakeItem = False
    fmeSelectItem.Caption = "Select an item to gather."
    chkTake.Visible = True
    fmeSelectItem.Visible = True
    Call BTF(fmeSelectItem)
End Sub

Private Sub btnTask_Kill_Click()
    On Error Resume Next
    KillNPC = True
    fmeCLI.Caption = "Select the enemy the player will need to kill"
    lblKillAmnt.Visible = True
    scrlKillAmnt.Visible = True
    cmbNPC.ListIndex = -1
    Call SetNPCBox(True, CLI.ListIndex + 1)
    fmeCLI.Visible = True
    Call BTF(fmeCLI)
End Sub

Private Sub btnTaskCancel_Click()
    fmeTask.Visible = False
    ResetEditButtons
End Sub

Private Sub btnUp_Click()
Dim tempSel As Long, tempSel2 As Long
    If CLIHasFocus Then
        'move item up within the CLI
        tempSel = CLI.ListIndex
        If tempSel <= 0 Then
            btnUp.Enabled = False
            Exit Sub
        End If
        If Not CLI.ListCount > 1 Then Exit Sub
        
        Call MoveListItem(LIST_CLI, EditorIndex, 0, tempSel + 1, -1)
        CLI.ListIndex = tempSel - 1
    Else
        'move item up within the Task List
        tempSel = lstTasks.ListIndex
        If tempSel <= 0 Then
            btnUp.Enabled = False
            Exit Sub
        End If
        If Not lstTasks.ListCount > 1 Then Exit Sub
        tempSel2 = CLI.ListIndex
        If tempSel2 < 0 Then Exit Sub
        
        'tempsel/2 is +1 because the array for the data starts at 1 whereas the listbox starts at 0
        Call MoveListItem(LIST_TASK, EditorIndex, tempSel2 + 1, tempSel + 1, -1)
        CLI.ListIndex = tempSel2
        lstTasks.ListIndex = tempSel - 1
    End If
End Sub

Private Sub btnWarp_Click()
Dim Index As Long
    Index = CLI.ListIndex + 1
    If Index < 1 Then Exit Sub
    
    'open the panel to warp the player.
    fmeWarp.Visible = True
    Call BTF(fmeWarp)
End Sub

Private Sub chkRetake_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    fmeMoveItem.Visible = False
    Quest(EditorIndex).CanBeRetaken = chkRetake.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "chkRetake_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub chkSet_Click()
    If chkSet.Value = vbChecked Then
        scrlModify.Value = 0
        scrlModify.min = 0
    Else
        scrlModify.Value = 0
        scrlModify.min = -32767
    End If
End Sub

Private Sub CLI_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    CLIHasFocus = True
    QuestEditorInitCLI
    If CLI.ListCount > 1 Then
        fmeMoveItem.Visible = True
        fmeMoveItem.Left = 1560
        If CLI.ListIndex = 0 Then btnUp.Enabled = False Else btnUp.Enabled = True
        If CLI.ListIndex = CLI.ListCount - 1 Then btnDown.Enabled = False Else btnDown.Enabled = True
    Else
        fmeMoveItem.Visible = False
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CLI_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub CLI_DblClick()
Dim Index As Long, I As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    'we're gonna edit this list item instead of creating one.
    CLIHasFocus = True
    Index = CLI.ListIndex + 1
    If Index < 1 Then
        Call QMsg("Please select a greeter to edit first.")
        Exit Sub
    End If
    Editing_CLI = True
    fmeCLI.Visible = True
    Editing_CLI_Index = Index
    Call SetNPCBox(False, Index)
    
    For I = 0 To cmbNPC.ListCount - 1
        If cmbNPC.List(I) = Trim$(NPC(Quest(EditorIndex).CLI(Index).ItemIndex).Name) Then
            cmbNPC.ListIndex = I
            Exit Sub
        End If
    Next I
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "CLI_DblClick", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub CLI_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    If Button = vbRightButton Then
        CLIHasFocus = True
        mnuACLI.Visible = True
        mnuRCLI.Visible = True
        mnuRTask.Visible = False
        If CLI.ListCount > 0 Then mnuEdit.Visible = True Else mnuEdit.Visible = False
        PopupMenu mnuCLI
    End If
    Exit Sub
' Error handler
ErrorHandler:
    HandleError "CLI_MouseDown", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbClassReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Quest(EditorIndex).Requirements.ClassReq = cmbClassReq.ListIndex
    Exit Sub
' Error handler
ErrorHandler:
    HandleError "cmbClassReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbGenderReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Quest(EditorIndex).Requirements.GenderReq = cmbGenderReq.ListIndex
    Exit Sub
' Error handler
ErrorHandler:
    HandleError "cmbGenderReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmbSkillReq_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Quest(EditorIndex).Requirements.SkillReq = cmbSkillReq.ListIndex
    Exit Sub
' Error handler
ErrorHandler:
    HandleError "cmbSkillReq_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdCopy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    TmpIndex = lstIndex.ListIndex
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdCopy_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub cmdPaste_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    lstIndex.RemoveItem EditorIndex - 1
    Call CopyMemory(ByVal VarPtr(Quest(EditorIndex)), ByVal VarPtr(Quest(TmpIndex + 1)), LenB(Quest(TmpIndex + 1)))
    lstIndex.AddItem EditorIndex & ": " & Trim$(Quest(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = EditorIndex - 1
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "cmdPaste_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Load()
Dim I As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmEditor_Quest.Width = 14895
    frmEditor_Quest.Height = 9195
    frmEditor_Quest.Caption = "Quest Editor"
    frmMain.SubDaFocus Me.hWnd
    frmEditor_Quest.scrlLevelReq.max = MAX_LEVEL
    txtName.MaxLength = QUESTNAME_LENGTH
    txtSearch.MaxLength = NAME_LENGTH
    txtDesc.MaxLength = QUESTDESC_LENGTH
    
    cmbSkillReq.Clear
    cmbClassReq.Clear
    cmbSkillReq.AddItem "None"
    cmbClassReq.AddItem "None"
    
    For I = 1 To Skill_Count - 1
        cmbSkillReq.AddItem GetSkillName(I)
        cmbObSKill.AddItem GetSkillName(I)
    Next I
    
    For I = 1 To MAX_CLASSES
        If Len(Trim$(Class(I).Name)) > 0 Then
            cmbClassReq.AddItem Trim$(Class(I).Name)
        End If
    Next I
    
    Call PositionFrames
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Load", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    frmMain.UnsubDaFocus Me.hWnd
    If EditorSave = False Then
        'QuestEditorCancel
    Else
        EditorSave = False
    End If
    frmAdmin.chkEditor(EDITOR_QUESTS).Value = False
    BringWindowToTop (frmAdmin.hWnd)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "Form_Unload", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblCredit_Click()
    fmeMoveItem.Visible = False
End Sub

Private Sub lstIndex_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    fmeMoveItem.Visible = False
    QuestEditorInit
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstIndex_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstTasks_Click()
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    CLIHasFocus = False
    
    If lstTasks.ListCount > 1 Then
        If Quest(EditorIndex).CLI(CLI.ListIndex + 1).Action(lstTasks.ListIndex + 1).ActionID = ACTION_SHOWMSG Then
            If Quest(EditorIndex).CLI(CLI.ListIndex + 1).Action(lstTasks.ListIndex + 1).MainData = vbChecked Then
                fmeMoveItem.Visible = False
                Exit Sub
            End If
        End If
        
        fmeMoveItem.Visible = True
        fmeMoveItem.Left = 4680
        If lstTasks.ListIndex = 0 Then btnUp.Enabled = False Else btnUp.Enabled = True
        If lstTasks.ListIndex = lstTasks.ListCount - 1 Then btnDown.Enabled = False Else btnDown.Enabled = True
        
        If lstTasks.ListIndex = 1 Then
            If Quest(EditorIndex).CLI(CLI.ListIndex + 1).Action(lstTasks.ListIndex).ActionID = ACTION_SHOWMSG Then
                If Quest(EditorIndex).CLI(CLI.ListIndex + 1).Action(lstTasks.ListIndex).MainData = vbChecked Then
                    btnUp.Enabled = False
                End If
            End If
        End If
    Else
        fmeMoveItem.Visible = False
    End If
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstTasks_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstTasks_DblClick()
Dim Index As Long, I As Long, II As Long
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    CLIHasFocus = False
    'edit the selected list item instead of creating a new one
    Index = CLI.ListIndex + 1
    I = lstTasks.ListIndex + 1
    If Index < 1 Then
        Call QMsg("Please select a greeter first.")
        Exit Sub
    End If
    If I < 1 Then
        Call QMsg("Please select a task to edit first.")
        Exit Sub
    End If
    
    Editing_Task = True
    Editing_Task_Index = I
    
    Call DisableEditButtons
    
    With Quest(EditorIndex).CLI(Index).Action(I)
        Select Case .ActionID
            Case TASK_GATHER
                Gather = True
                GiveItem = False
                TakeItem = False
                chkTake.Value = .SecondaryData
                fmeSelectItem.Visible = True
                Call BTF(fmeSelectItem)
                chkTake.Visible = True
                scrlItem.Value = .MainData
                scrlItemAmount.Value = .Amount
            Case TASK_KILL
                KillNPC = True
                fmeCLI.Visible = True
                Call BTF(fmeCLI)
                lblKillAmnt.Visible = True
                scrlKillAmnt.Visible = True
                Call SetNPCBox(True, CLI.ListIndex + 1)
                For I = 0 To cmbNPC.ListCount - 1
                    If cmbNPC.List(I) = Trim$(NPC(.MainData).Name) Then
                        cmbNPC.ListIndex = I
                        Exit For
                    End If
                Next I
                scrlKillAmnt.Value = .Amount
            Case TASK_GETSKILL
                fmeObtainSKill.Visible = True
                Call BTF(fmeObtainSKill)
                cmbObSKill.ListIndex = .MainData - 1
                scrlObSkill.Value = .Amount
            Case ACTION_SHOWMSG
                Call CheckResponseMsg(EditorIndex, Index, I - 1)
                fmeShowMsg.Visible = True
                Call BTF(fmeShowMsg)
                txtMsg.text = Trim$(.TextHolder)
                chkStart.Value = .MainData
                chkRes.Value = .SecondaryData
                cmbColor.ListIndex = .TertiaryData
                If .MainData = vbChecked Then chkStart.Enabled = True
            Case ACTION_ADJUST_EXP
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opEXP.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .Amount
            Case ACTION_ADJUST_LVL
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opLvl.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .Amount
            Case ACTION_ADJUST_STAT_LVL
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opStat.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .Amount
                cboItem.ListIndex = .SecondaryData
            Case ACTION_ADJUST_STAT_EXP
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opStatEXP.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .Amount
                cboItem.ListIndex = .SecondaryData
            Case ACTION_ADJUST_STAT_POINTS
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opStatP.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .Amount
                cboItem.ListIndex = .SecondaryData
            Case ACTION_ADJUST_SKILL_LVL
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opSkill.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .Amount
                cboItem.ListIndex = .SecondaryData
            Case ACTION_ADJUST_SKILL_EXP
                fmeModify.Visible = True
                Call BTF(fmeModify)
                opSkillEXP.Value = True
                chkSet.Value = .MainData
                scrlModify.Value = .Amount
                cboItem.ListIndex = .SecondaryData
            Case ACTION_GIVE_ITEM
                Gather = False
                GiveItem = True
                TakeItem = False
                fmeSelectItem.Visible = True
                Call BTF(fmeSelectItem)
                scrlItem.Value = .MainData
                scrlItemAmount.Value = .Amount
            Case ACTION_TAKE_ITEM
                Gather = False
                GiveItem = False
                TakeItem = True
                fmeSelectItem.Visible = True
                Call BTF(fmeSelectItem)
                scrlItem.Value = .MainData
                scrlItemAmount.Value = .Amount
            Case ACTION_WARP
                fmeWarp.Visible = True
                Call BTF(fmeWarp)
                scrlMap.Value = .Amount
                scrlMapX.Value = .MainData
                scrlMapY.Value = .SecondaryData
            Case Else
                Exit Sub
        End Select
    End With
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "lstTasks_DblClick", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstTasks_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = vbRightButton Then
        CLIHasFocus = False
        If lstTasks.ListCount > 0 Then mnuEdit.Visible = True Else mnuEdit.Visible = False
        mnuACLI.Visible = False
        mnuRCLI.Visible = False
        mnuRTask.Visible = True
        PopupMenu mnuCLI
    End If
End Sub

Private Sub mnuAAction_Click()
    If CLI.ListIndex < 0 Then
        Call QMsg("Please select one of the NPC's to meet with from the list below.")
        Exit Sub
    End If
    fmeMoveItem.Visible = False
    fmeTask.Visible = True
    Call BTF(fmeTask)
    DisableEditButtons
End Sub

Private Sub mnuACLI_Click()
    KillNPC = False
    fmeCLI.Caption = "Add a new NPC/Event the player will need to meet with"
    fmeCLI.Visible = True
    lblKillAmnt.Visible = False
    scrlKillAmnt.Visible = False
    Call SetNPCBox(False, CLI.ListIndex + 1, True)
    Call BTF(fmeCLI)
    DisableEditButtons
End Sub

Private Sub mnuEdit_Click()
    If CLIHasFocus Then Call CLI_DblClick Else Call lstTasks_DblClick
End Sub

Private Sub mnuRCLI_Click()
Dim Index As Long
Dim Res As VbMsgBoxResult
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Index = CLI.ListIndex + 1
    If Index < 1 Then
        Call QMsg("Please select the Greeter you would like to delete.")
        Exit Sub
    End If
    fmeMoveItem.Visible = False
    
    'lets delete the selected action/task
    Res = MsgBox("Are you sure you want to delete this Greeter?", vbYesNo, "VERIFICATION")
    If Res = vbYes Then Call DeleteCLI(EditorIndex, Index)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "mnuRCLI_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub mnuRTask_Click()
Dim Index As Long, TaskID As Long
Dim Res As VbMsgBoxResult
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Index = CLI.ListIndex + 1
    TaskID = lstTasks.ListIndex + 1
    If Index < 1 Then
        Call QMsg("Please select a Greeter first, and then an Action/Task to be removed..")
        Exit Sub
    End If
    If TaskID < 1 Then
        Call QMsg("Please select an Action/Task to remove from the list.")
        Exit Sub
    End If
    fmeMoveItem.Visible = False
    
    'lets delete the selected action/task
    Res = MsgBox("Are you sure you want to delete this Action/Task?", vbYesNo, "VERIFICATION")
    If Res = vbYes Then Call DeleteAction(EditorIndex, Index, TaskID)
    CLI.ListIndex = Index - 1
    Call QuestEditorInitCLI
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "mnuRTask_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub opEvent_Click()
Dim I As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    cmbNPC.Clear
    For I = 1 To MAX_NPCS
        If Len(Trim$(NPC(I).Name)) > 0 Then
            If NPC(I).Behavior = NPC_BEHAVIOR_QUEST Then
                cmbNPC.AddItem Trim$(NPC(I).Name)
            End If
        End If
    Next I
    
' Error handler
ErrorHandler:
    HandleError "opEvent_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub opEXP_Click()
    cboItem.Enabled = False
    cboItem.Clear
    cboItem.AddItem "None"
    cboItem.ListIndex = 0
End Sub

Private Sub opLvl_Click()
    cboItem.Enabled = False
    cboItem.Clear
    cboItem.AddItem "None"
    cboItem.ListIndex = 0
End Sub

Private Sub opNPC_Click()
Dim I As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    cmbNPC.Clear
    cmbNPC.AddItem "Select An NPC"
    For I = 1 To MAX_NPCS
        If Len(Trim$(NPC(I).Name)) > 0 Then
            If NPC(I).Behavior = NPC_BEHAVIOR_QUEST Then
                cmbNPC.AddItem Trim$(NPC(I).Name)
            End If
        End If
    Next I
    
' Error handler
ErrorHandler:
    HandleError "opNPC_Click", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub opSkill_Click()
Dim I As Long
    cboItem.Enabled = True
    cboItem.Clear
    cboItem.AddItem "None"
    For I = 1 To Skills.Skill_Count - 1
        cboItem.AddItem GetSkillName(I)
    Next I
    cboItem.ListIndex = 0
End Sub

Private Sub opSkillEXP_Click()
Dim I As Long
    cboItem.Enabled = True
    cboItem.Clear
    cboItem.AddItem "None"
    For I = 1 To Skills.Skill_Count - 1
        cboItem.AddItem GetSkillName(I)
    Next I
    cboItem.ListIndex = 0
End Sub

Private Sub opStat_Click()
Dim I As Long
    cboItem.Enabled = True
    cboItem.Clear
    cboItem.AddItem "None"
    For I = 1 To Stats.Stat_Count - 1
        Select Case I
            Case Stats.Agility
                cboItem.AddItem "Agility"
            Case Stats.Endurance
                cboItem.AddItem "Endurance"
            Case Stats.Intelligence
                cboItem.AddItem "Intelligence"
            Case Stats.Spirit
                cboItem.AddItem "Spirit"
            Case Stats.Strength
                cboItem.AddItem "Strength"
        End Select
    Next I
    cboItem.ListIndex = 0
End Sub

Private Sub opStatEXP_Click()
Dim I As Long
    cboItem.Enabled = True
    cboItem.Clear
    cboItem.AddItem "None"
    For I = 1 To Stats.Stat_Count - 1
        Select Case I
            Case Stats.Agility
                cboItem.AddItem "Agility"
            Case Stats.Endurance
                cboItem.AddItem "Endurance"
            Case Stats.Intelligence
                cboItem.AddItem "Intelligence"
            Case Stats.Spirit
                cboItem.AddItem "Spirit"
            Case Stats.Strength
                cboItem.AddItem "Strength"
        End Select
    Next I
    cboItem.ListIndex = 0
End Sub

Private Sub opStatP_Click()
    cboItem.Enabled = False
    cboItem.Clear
    cboItem.AddItem "None"
    cboItem.ListIndex = 0
End Sub

Private Sub scrlAccessReq_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblAccessReq.Caption = "Access: " & scrlAccessReq.Value
    Quest(EditorIndex).Requirements.AccessReq = scrlAccessReq.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlAccessReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlItem_Change()
    scrlItemAmount.Value = 0
    lblItem.Caption = "Item: " & scrlItem.Value

    If scrlItem.Value > 0 And scrlItem.Value < MAX_ITEMS Then
        If Len(Trim$(Item(scrlItem.Value).Name)) > 0 Then
            lblItem.Caption = "Item: " & Trim$(Item(scrlItem.Value).Name)
            If Gather Then
                If Not Item(scrlItem.Value).stackable > 0 Then
                    scrlItemAmount.max = MAX_INV
                Else
                    scrlItemAmount.max = 32767
                End If
            End If
        End If
    End If
End Sub

Private Sub scrlItemAmount_Change()
    lblAmount.Caption = "Amount: " & scrlItemAmount.Value
End Sub

Private Sub scrlKillAmnt_Change()
    lblKillAmnt.Caption = "Amount: " & scrlKillAmnt.Value
End Sub

Private Sub scrlLevelReq_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblLevelReq.Caption = "Level: " & scrlLevelReq.Value
    Quest(EditorIndex).Requirements.LevelReq = scrlLevelReq.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlLevelReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlMap_Change()
    lblMap.Caption = "Map: " & scrlMap.Value
End Sub

Private Sub scrlMapX_Change()
    lblMapX.Caption = "X: " & scrlMapX.Value
End Sub

Private Sub scrlMapY_Change()
    lblMapY.Caption = "Y: " & scrlMapY.Value
End Sub

Private Sub scrlModify_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblModify.Caption = "Amount to modify: " & scrlModify.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlModify_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ResetEditButtons()
    mnuACLI.Enabled = True
    mnuRCLI.Enabled = True
    mnuAAction.Enabled = True
    mnuRTask.Enabled = True
    mnuEdit.Enabled = True
    btnReq.Enabled = True
    CLI.Enabled = True
End Sub

Private Sub DisableEditButtons()
    mnuACLI.Enabled = False
    mnuRCLI.Enabled = False
    mnuAAction.Enabled = False
    mnuRTask.Enabled = False
    mnuEdit.Enabled = False
    btnReq.Enabled = False
    CLI.Enabled = False
End Sub

Private Sub scrlObSkill_Change()
    lblObSkill.Caption = "Skill Level: " & scrlObSkill.Value
End Sub

Private Sub scrlSkill_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler

    lblSkill.Caption = "Skill Level: " & scrlSkill.Value
    Quest(EditorIndex).Requirements.SkillLevelReq = scrlSkill.Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlSkill_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub scrlStatReq_Change(Index As Integer)
Dim tmpStr As String
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    Select Case Index
        Case 1
            tmpStr = "Str: "
        Case 2
            tmpStr = "End: "
        Case 3
            tmpStr = "Int: "
        Case 4
            tmpStr = "Agi: "
        Case 5
            tmpStr = "Spi: "
        Case Else
            Exit Sub
    End Select

    lblStatReq(Index).Caption = tmpStr & scrlStatReq(Index).Value
    Quest(EditorIndex).Requirements.Stat_Req(Index) = scrlStatReq(Index).Value
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "scrlStatReq_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub tmrMsg_Timer()
    lblMsg.Caption = vbNullString
End Sub

Private Sub txtDesc_Change()
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    Quest(EditorIndex).Description = Trim$(txtDesc.text)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtDesc_Change", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtDesc_Click()
    fmeMoveItem.Visible = False
End Sub

Private Sub txtName_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    fmeMoveItem.Visible = False
    txtName.SelStart = Len(txtName)
    Exit Sub
    
' Error handler
ErrorHandler:
    HandleError "txtName_GotFocus", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtName_Validate(Cancel As Boolean)
Dim I As Long
Dim TmpIndex As Long
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo ErrorHandler
    
    'be sure we don't use the same name twice.
    For I = 1 To MAX_QUESTS
        If LCase(Trim(txtName.text)) = LCase(Trim(Quest(I).Name)) Then
            txtName.text = vbNullString
            Call MsgBox("Duplicate quest name found.  Quest names must be unique.  Please change it.", vbOKOnly, "Duplicate Quest Name")
            Exit Sub
        End If
    Next I
    
    If EditorIndex < 1 Or EditorIndex > MAX_QUESTS Then Exit Sub
    
    TmpIndex = lstIndex.ListIndex
    Quest(EditorIndex).Name = Trim$(txtName.text)
    lstIndex.RemoveItem EditorIndex - 1
    lstIndex.AddItem EditorIndex & ": " & Trim$(Quest(EditorIndex).Name), EditorIndex - 1
    lstIndex.ListIndex = TmpIndex
    Exit Sub
    
' Error handlerin
ErrorHandler:
    HandleError "txtName_Validate", "frmEditor_Quest", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub QMsg(ByVal Msg As String)
    tmrMsg.Enabled = False
    lblMsg.Caption = Msg
    lblMsg.Visible = True
    tmrMsg.Enabled = True
End Sub

Private Sub txtSearch_Click()
    fmeMoveItem.Visible = False
End Sub

Private Sub PositionFrames()
Dim tTop As Integer, tLeft As Integer
Dim FME As Control
    tTop = 2640
    tLeft = 6480
    
    For Each FME In frmEditor_Quest.Controls
        If (TypeOf FME Is Frame) Then
            If FME.Name <> "Frame1" And FME.Name <> "Frame2" And FME.Name <> "Frame3" And FME.Name <> "fraNPC" And FME.Name <> "fmeMoveItem" Then
                FME.Top = tTop
                FME.Left = tLeft
            End If
        End If
    Next
End Sub

Private Sub BTF(ByVal FrameID As Frame)
    Call FrameID.ZOrder(vbBringToFront)
End Sub

Private Sub SetNPCBox(ByVal All As Boolean, Optional ByVal CurCLI As Long = 0, Optional ByVal Adding As Boolean = False)
Dim I As Long, ShowItem As Boolean
    cmbNPC.Clear
    cmbNPC.AddItem "Select An NPC"
    
    For I = 1 To MAX_NPCS
        If Len(Trim$(NPC(I).Name)) > 0 Then
            If All Then
                If Not NPC(I).Behavior = NPC_BEHAVIOR_QUEST Then
                    cmbNPC.AddItem Trim$(NPC(I).Name)
                End If
            Else
                If NPC(I).Behavior = NPC_BEHAVIOR_QUEST Then
                    ShowItem = True
                    
                    If Quest(EditorIndex).Max_CLI = 0 Then
                        'Don't allow the same NPC to be used for the beginning of more than one quest
                        If IsNPCInAnotherQuest(I, EditorIndex) Then ShowItem = False
                    Else
                        'Don't show the NPC if it used in the previous slot
                        If Adding Then
                            If Quest(EditorIndex).CLI(Quest(EditorIndex).Max_CLI).ItemIndex = I Then ShowItem = False
                        Else
                            If CurCLI > 1 Then
                                If Quest(EditorIndex).CLI(CurCLI - 1).ItemIndex = I Then ShowItem = False
                            ElseIf CurCLI = 1 Then
                                'Don't allow the same NPC to be used for the beginning of more than one quest
                                If IsNPCInAnotherQuest(I, EditorIndex) Then ShowItem = False
                            End If
                        End If
                    End If
                    
                    If ShowItem Then cmbNPC.AddItem Trim$(NPC(I).Name)
                End If
            End If
        End If
    Next I
    
    cmbNPC.ListIndex = 0
End Sub
