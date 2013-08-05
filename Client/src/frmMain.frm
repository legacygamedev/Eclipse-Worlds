VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   ClientHeight    =   13470
   ClientLeft      =   120
   ClientTop       =   795
   ClientWidth     =   12000
   BeginProperty Font 
      Name            =   "Georgia"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   898
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   StartUpPosition =   2  'CenterScreen
   Tag             =   " "
   Visible         =   0   'False
   Begin VB.ListBox lstDropDownBox 
      Height          =   270
      Left            =   0
      TabIndex        =   142
      Top             =   0
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.PictureBox picSpellDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   3120
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   138
      Top             =   8760
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picSpellDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   139
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblSpellName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   141
         Top             =   210
         Width           =   2805
      End
      Begin VB.Label lblSpellDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         ForeColor       =   &H00FFFFFF&
         Height          =   1530
         Left            =   240
         TabIndex        =   140
         Top             =   1800
         Width           =   2640
      End
   End
   Begin VB.PictureBox picItemDesc 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   -120
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   210
      TabIndex        =   134
      Top             =   8760
      Visible         =   0   'False
      Width           =   3150
      Begin VB.PictureBox picItemDescPic 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   1095
         ScaleHeight     =   64
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   64
         TabIndex        =   135
         Top             =   600
         Width           =   960
      End
      Begin VB.Label lblItemName 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "N/A"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Left            =   150
         TabIndex        =   137
         Top             =   210
         Width           =   2805
      End
      Begin VB.Label lblItemDesc 
         BackStyle       =   0  'Transparent
         Caption         =   """This is an example of an item's description. It  can be quite big, so we have to keep it at a decent size."""
         ForeColor       =   &H00FFFFFF&
         Height          =   1530
         Left            =   240
         TabIndex        =   136
         Top             =   1800
         Width           =   2640
      End
   End
   Begin VB.PictureBox picOptionSwearFilter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8100
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   106
      TabStop         =   0   'False
      Top             =   9615
      Width           =   735
   End
   Begin VB.PictureBox picOptionWeather 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8100
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   105
      TabStop         =   0   'False
      Top             =   9975
      Width           =   735
   End
   Begin VB.PictureBox picOptionAutoTile 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8100
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   104
      TabStop         =   0   'False
      Top             =   10335
      Width           =   735
   End
   Begin VB.PictureBox picOptionDebug 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8100
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   103
      TabStop         =   0   'False
      Top             =   11055
      Width           =   735
   End
   Begin VB.PictureBox picOptionBlood 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   8100
      ScaleHeight     =   285
      ScaleWidth      =   735
      TabIndex        =   102
      TabStop         =   0   'False
      Top             =   10695
      Width           =   735
   End
   Begin VB.PictureBox picSSMap 
      AutoRedraw      =   -1  'True
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
      Height          =   210
      Left            =   12000
      ScaleHeight     =   11.529
      ScaleMode       =   0  'User
      ScaleWidth      =   17
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   8520
      Width           =   255
   End
   Begin VB.PictureBox picTempSpell 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   7560
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   8760
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   6960
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   8760
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picTempInv 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   540
      Left            =   6360
      ScaleHeight     =   36
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   36
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   8760
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picShop 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00E0E0E0&
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
      Height          =   5115
      Left            =   3960
      ScaleHeight     =   341
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   275
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   1800
      Visible         =   0   'False
      Width           =   4125
      Begin VB.PictureBox picShopItems 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3165
         Left            =   615
         ScaleHeight     =   211
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   630
         Width           =   2895
      End
      Begin VB.Image imgShopBuy 
         Height          =   435
         Left            =   360
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imgShopSell 
         Height          =   435
         Left            =   1530
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image imgLeaveShop 
         Height          =   435
         Left            =   2700
         Top             =   4320
         Width           =   1035
      End
      Begin VB.Image ImgFix 
         Height          =   315
         Left            =   1890
         Top             =   3840
         Width           =   375
      End
   End
   Begin VB.PictureBox picForm 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   238
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
      ScaleWidth      =   800
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   12000
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   9
         Left            =   11400
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   7080
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   10
         Left            =   9600
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   89
         TabStop         =   0   'False
         Top             =   8040
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   16
         Left            =   9600
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   85
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   12
         Left            =   9000
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   83
         TabStop         =   0   'False
         Top             =   8040
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   15
         Left            =   11400
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   8040
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   14
         Left            =   10800
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   81
         TabStop         =   0   'False
         Top             =   8040
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   13
         Left            =   10800
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   80
         TabStop         =   0   'False
         Top             =   7080
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   11
         Left            =   10200
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   79
         TabStop         =   0   'False
         Top             =   8040
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   1
         Left            =   7800
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   78
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   3
         Left            =   9000
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   4
         Left            =   11400
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   76
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   5
         Left            =   7800
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   75
         TabStop         =   0   'False
         Top             =   8040
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   6
         Left            =   10200
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   74
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   7
         Left            =   8400
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   8040
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   8
         Left            =   10800
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   72
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picButton 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   435
         Index           =   2
         Left            =   8400
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   32
         TabIndex        =   71
         TabStop         =   0   'False
         Top             =   7560
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.PictureBox picHotbar 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   450
         Left            =   4800
         ScaleHeight     =   30
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   476
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   120
         Width           =   7140
      End
      Begin VB.PictureBox picGUI_Vitals_Base 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1125
         Left            =   120
         ScaleHeight     =   75
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   254
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   120
         Width           =   3810
         Begin VB.Label lblHP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100/100"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1815
            TabIndex        =   22
            Top             =   135
            Width           =   1845
         End
         Begin VB.Label lblMP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100/100"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1815
            TabIndex        =   21
            Top             =   465
            Width           =   1845
         End
         Begin VB.Label lblEXP 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "100/100"
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1815
            TabIndex        =   20
            Top             =   795
            Width           =   1845
         End
         Begin VB.Image imgHPBar 
            Height          =   240
            Left            =   105
            Top             =   135
            Width           =   3615
         End
         Begin VB.Image imgMPBar 
            Height          =   240
            Left            =   105
            Top             =   465
            Width           =   3615
         End
         Begin VB.Image imgEXPBar 
            Height          =   240
            Left            =   120
            Top             =   795
            Width           =   3615
         End
      End
      Begin MSWinsockLib.Winsock Socket 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   393216
      End
      Begin VB.PictureBox picChatbox 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2175
         Left            =   120
         ScaleHeight     =   145
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   484
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   6360
         Width           =   7260
         Begin VB.TextBox txtMyChat 
            Appearance      =   0  'Flat
            BackColor       =   &H00080D10&
            BorderStyle     =   0  'None
            Enabled         =   0   'False
            ForeColor       =   &H00FFFFFF&
            Height          =   225
            Left            =   600
            MaxLength       =   512
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   1890
            Width           =   6585
         End
         Begin RichTextLib.RichTextBox txtChat 
            Height          =   1755
            Left            =   60
            TabIndex        =   24
            TabStop         =   0   'False
            Top             =   60
            Width           =   7170
            _ExtentX        =   12647
            _ExtentY        =   3096
            _Version        =   393217
            BackColor       =   527632
            BorderStyle     =   0
            ReadOnly        =   -1  'True
            ScrollBars      =   2
            Appearance      =   0
            TextRTF         =   $"frmMain.frx":038A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
      Begin VB.PictureBox picScreen 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00181C21&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   5760
         Left            =   0
         ScaleHeight     =   384
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   512
         TabIndex        =   101
         TabStop         =   0   'False
         Top             =   0
         Width           =   7680
      End
      Begin VB.PictureBox picTitles 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9000
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   194
         TabIndex        =   95
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   2910
         Begin VB.ListBox lstTitles 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2130
            Left            =   240
            TabIndex        =   96
            TabStop         =   0   'False
            Top             =   600
            Width           =   2460
         End
         Begin VB.Label lblDesc 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "None."
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   735
            Left            =   240
            TabIndex        =   98
            Top             =   3240
            Width           =   2535
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Description:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   855
            TabIndex        =   97
            Top             =   3000
            Width           =   1215
         End
      End
      Begin VB.PictureBox picSpells 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9000
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   194
         TabIndex        =   133
         TabStop         =   0   'False
         Top             =   2880
         Visible         =   0   'False
         Width           =   2910
      End
      Begin VB.PictureBox picParty 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3750
         Left            =   9000
         ScaleHeight     =   250
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   195
         TabIndex        =   48
         TabStop         =   0   'False
         Top             =   3210
         Visible         =   0   'False
         Width           =   2925
         Begin VB.Image imgPartySpirit 
            Height          =   135
            Index           =   4
            Left            =   105
            Top             =   2760
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartyHealth 
            Height          =   135
            Index           =   4
            Left            =   105
            Top             =   2625
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartySpirit 
            Height          =   135
            Index           =   3
            Left            =   105
            Top             =   2025
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartyHealth 
            Height          =   135
            Index           =   3
            Left            =   105
            Top             =   1890
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartySpirit 
            Height          =   135
            Index           =   2
            Left            =   105
            Top             =   1320
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartyHealth 
            Height          =   135
            Index           =   2
            Left            =   105
            Top             =   1170
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartySpirit 
            Height          =   135
            Index           =   1
            Left            =   120
            Top             =   555
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Image imgPartyHealth 
            Height          =   135
            Index           =   1
            Left            =   105
            Top             =   420
            Visible         =   0   'False
            Width           =   2730
         End
         Begin VB.Label lblPartyLeave 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   1575
            TabIndex        =   54
            Top             =   3165
            Width           =   1095
         End
         Begin VB.Label lblPartyInvite 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   375
            Left            =   375
            TabIndex        =   53
            Top             =   3165
            Width           =   1095
         End
         Begin VB.Label lblPartyMember 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   4
            Left            =   255
            TabIndex        =   52
            Top             =   2355
            Width           =   2415
         End
         Begin VB.Label lblPartyMember 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   3
            Left            =   255
            TabIndex        =   51
            Top             =   1620
            Width           =   2415
         End
         Begin VB.Label lblPartyMember 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   2
            Left            =   255
            TabIndex        =   50
            Top             =   885
            Width           =   2415
         End
         Begin VB.Label lblPartyMember 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Index           =   1
            Left            =   255
            TabIndex        =   49
            Top             =   150
            Width           =   2415
         End
      End
      Begin VB.PictureBox picFriends 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9000
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   194
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   2910
         Begin VB.ListBox lstFriends 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2550
            Left            =   240
            TabIndex        =   56
            TabStop         =   0   'False
            Top             =   600
            Width           =   2460
         End
         Begin VB.Label lblRemoveFriend 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remove Friend"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   60
            TabIndex        =   58
            Top             =   3600
            Width           =   2805
         End
         Begin VB.Label lblAddFriend 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Add Friend"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   165
            Left            =   0
            TabIndex        =   57
            Top             =   3360
            Width           =   2925
         End
      End
      Begin VB.PictureBox picEventChat 
         Appearance      =   0  'Flat
         BackColor       =   &H000C0E10&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2790
         Left            =   120
         ScaleHeight     =   186
         ScaleMode       =   0  'User
         ScaleWidth      =   482
         TabIndex        =   143
         TabStop         =   0   'False
         Top             =   5760
         Visible         =   0   'False
         Width           =   7230
         Begin VB.PictureBox picChatFace 
            AutoRedraw      =   -1  'True
            BorderStyle     =   0  'None
            Height          =   1500
            Left            =   120
            ScaleHeight     =   100
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   100
            TabIndex        =   144
            TabStop         =   0   'False
            Top             =   120
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.Label lblChoices 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "[Option 1]"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   210
            Index           =   1
            Left            =   240
            TabIndex        =   150
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblChoices 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "[Option 2]"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   210
            Index           =   2
            Left            =   240
            TabIndex        =   149
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblChoices 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "[Option 3]"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   210
            Index           =   3
            Left            =   3240
            TabIndex        =   148
            Top             =   1920
            Width           =   975
         End
         Begin VB.Label lblChoices 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "[Option 4]"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   210
            Index           =   4
            Left            =   3240
            TabIndex        =   147
            Top             =   2400
            Width           =   975
         End
         Begin VB.Label lblEventChatContinue 
            AutoSize        =   -1  'True
            BackColor       =   &H00000000&
            BackStyle       =   0  'Transparent
            Caption         =   "Continue..."
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000003&
            Height          =   210
            Left            =   6000
            TabIndex        =   146
            Top             =   2400
            Width           =   1095
         End
         Begin VB.Label lblEventChat 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "[Text]"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   1335
            Left            =   180
            TabIndex        =   145
            Top             =   120
            Width           =   6915
         End
      End
      Begin VB.PictureBox picEquipment 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   2400
         Left            =   9000
         ScaleHeight     =   160
         ScaleMode       =   0  'User
         ScaleWidth      =   195
         TabIndex        =   84
         TabStop         =   0   'False
         Top             =   4560
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.PictureBox picOptions 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9000
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   194
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   2910
         Begin VB.PictureBox picOptionMusic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   122
            TabStop         =   0   'False
            Top             =   240
            Width           =   735
         End
         Begin VB.PictureBox picOptionSound 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   121
            TabStop         =   0   'False
            Top             =   600
            Width           =   735
         End
         Begin VB.PictureBox picOptionLevel 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   120
            TabStop         =   0   'False
            Top             =   960
            Width           =   735
         End
         Begin VB.PictureBox picOptionGuild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   1320
            Width           =   735
         End
         Begin VB.PictureBox picOptionTitle 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   118
            TabStop         =   0   'False
            Top             =   1680
            Width           =   735
         End
         Begin VB.PictureBox picOptionWASD 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   117
            TabStop         =   0   'False
            Top             =   2040
            Width           =   735
         End
         Begin VB.PictureBox picOptionMouse 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   116
            TabStop         =   0   'False
            Top             =   2400
            Width           =   735
         End
         Begin VB.PictureBox picOptionBattleMusic 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   115
            TabStop         =   0   'False
            Top             =   2760
            Width           =   735
         End
         Begin VB.PictureBox picOptionNpcVitals 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   114
            TabStop         =   0   'False
            Top             =   3480
            Width           =   735
         End
         Begin VB.PictureBox picOptionPlayerVitals 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2040
            ScaleHeight     =   285
            ScaleWidth      =   735
            TabIndex        =   113
            TabStop         =   0   'False
            Top             =   3120
            Width           =   735
         End
         Begin VB.Label lblMusic 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Music"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   132
            Top             =   240
            Width           =   555
         End
         Begin VB.Label lblSound 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Sound"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   131
            Top             =   600
            Width           =   600
         End
         Begin VB.Label lblGuilds 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Guilds"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   130
            Top             =   1320
            Width           =   630
         End
         Begin VB.Label lblLevels 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Levels"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   129
            Top             =   960
            Width           =   615
         End
         Begin VB.Label lblWASD 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "WASD"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   128
            Top             =   2040
            Width           =   585
         End
         Begin VB.Label lblMouse 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Mouse"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   127
            Top             =   2400
            Width           =   600
         End
         Begin VB.Label lblTitles 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Titles"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   126
            Top             =   1680
            Width           =   570
         End
         Begin VB.Label lblNpcVitals 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Npc Vitals"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   125
            Top             =   3480
            Width           =   1005
         End
         Begin VB.Label lblBattleMusic 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Battle Music"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   124
            Top             =   2760
            Width           =   1200
         End
         Begin VB.Label lblPlayerVitals 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Player Vitals"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   123
            Top             =   3120
            Width           =   1290
         End
      End
      Begin VB.PictureBox picGuild_No 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9000
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   194
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   2910
         Begin VB.Label lblNoGuild 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "You are not in a guild!"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   330
            TabIndex        =   87
            Top             =   2040
            Width           =   2205
         End
      End
      Begin VB.PictureBox picGuild 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9000
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   194
         TabIndex        =   66
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   2910
         Begin VB.ListBox lstGuild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000007&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2130
            Left            =   240
            TabIndex        =   67
            TabStop         =   0   'False
            Top             =   600
            Width           =   2460
         End
         Begin VB.Label lblGuildName 
            BackStyle       =   0  'Transparent
            Caption         =   "Guild Name"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   240
            TabIndex        =   100
            Top             =   240
            Width           =   2415
         End
         Begin VB.Label lblGuildRemove 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remove"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1080
            TabIndex        =   99
            Top             =   3480
            Width           =   795
         End
         Begin VB.Label lblGuildInvite 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Invite"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   70
            Top             =   3240
            Width           =   2715
         End
         Begin VB.Label lblResign 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Resign"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   69
            Top             =   3720
            Width           =   2715
         End
         Begin VB.Label lblChangeAccess 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Change Access"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   120
            TabIndex        =   68
            Top             =   3000
            Width           =   2715
         End
      End
      Begin VB.PictureBox picInventory 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9000
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   195
         TabIndex        =   27
         TabStop         =   0   'False
         Top             =   2880
         Visible         =   0   'False
         Width           =   2925
      End
      Begin VB.PictureBox picFoes 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4050
         Left            =   9000
         ScaleHeight     =   270
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   194
         TabIndex        =   91
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   2910
         Begin VB.ListBox lstFoes 
            Appearance      =   0  'Flat
            BackColor       =   &H80000006&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   2550
            Left            =   240
            TabIndex        =   92
            TabStop         =   0   'False
            Top             =   600
            Width           =   2460
         End
         Begin VB.Label lblAddFoe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Add Foe"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1065
            TabIndex        =   94
            Top             =   3360
            Width           =   795
         End
         Begin VB.Label lblRemoveFoe 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Remove Foe"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   885
            TabIndex        =   93
            Top             =   3600
            Width           =   1155
         End
      End
      Begin VB.PictureBox picCharacter 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   4215
         Left            =   9000
         ScaleHeight     =   281
         ScaleMode       =   0  'User
         ScaleWidth      =   195
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   2760
         Visible         =   0   'False
         Width           =   2925
         Begin VB.PictureBox picFace 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H00000000&
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
            Height          =   1440
            Left            =   735
            ScaleHeight     =   96
            ScaleMode       =   3  'Pixel
            ScaleWidth      =   96
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   570
            Width           =   1440
         End
         Begin VB.Label lblCharLevel 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Lv: 1"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1200
            TabIndex        =   88
            Top             =   2160
            Width           =   465
         End
         Begin VB.Label lblCharName 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Empty"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1170
            TabIndex        =   47
            Top             =   150
            Width           =   660
         End
         Begin VB.Label lblCharStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   720
            TabIndex        =   46
            Top             =   2880
            Width           =   315
         End
         Begin VB.Label lblCharStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   4
            Left            =   2040
            TabIndex        =   45
            Top             =   2880
            Width           =   315
         End
         Begin VB.Label lblCharStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   2
            Left            =   720
            TabIndex        =   44
            Top             =   3120
            Width           =   315
         End
         Begin VB.Label lblCharStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   5
            Left            =   2040
            TabIndex        =   43
            Top             =   3120
            Width           =   315
         End
         Begin VB.Label lblCharStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   3
            Left            =   720
            TabIndex        =   42
            Top             =   3360
            Width           =   315
         End
         Begin VB.Label lblTrainStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   1
            Left            =   1200
            TabIndex        =   41
            Top             =   2880
            Width           =   120
         End
         Begin VB.Label lblTrainStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   4
            Left            =   2520
            TabIndex        =   40
            Top             =   2880
            Width           =   120
         End
         Begin VB.Label lblTrainStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   2
            Left            =   1200
            TabIndex        =   39
            Top             =   3120
            Width           =   120
         End
         Begin VB.Label lblTrainStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   5
            Left            =   2520
            TabIndex        =   38
            Top             =   3120
            Width           =   120
         End
         Begin VB.Label lblTrainStat 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "+"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Index           =   3
            Left            =   1200
            TabIndex        =   37
            Top             =   3360
            Width           =   120
         End
         Begin VB.Label lblPoints 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "999"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   2280
            TabIndex        =   36
            Top             =   3360
            Width           =   315
         End
         Begin VB.Label lblStr 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Str:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   35
            Top             =   2880
            Width           =   360
         End
         Begin VB.Label lblEnd 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "End:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   34
            Top             =   3120
            Width           =   435
         End
         Begin VB.Label lblInt 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Int:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   240
            TabIndex        =   33
            Top             =   3360
            Width           =   360
         End
         Begin VB.Label lblSpi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Spi:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1560
            TabIndex        =   32
            Top             =   3120
            Width           =   360
         End
         Begin VB.Label lblAgi 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Agi:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1560
            TabIndex        =   31
            Top             =   2880
            Width           =   390
         End
         Begin VB.Label lblPoint 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Points:"
            BeginProperty Font 
               Name            =   "Georgia"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   210
            Left            =   1560
            TabIndex        =   30
            Top             =   3360
            Width           =   675
         End
      End
   End
   Begin VB.PictureBox picBank 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   150
      ScaleHeight     =   384
      ScaleMode       =   0  'User
      ScaleWidth      =   480
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   2760
      Visible         =   0   'False
      Width           =   7200
   End
   Begin VB.PictureBox picTrade 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5760
      Left            =   1800
      ScaleHeight     =   384
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   480
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   1440
      Visible         =   0   'False
      Width           =   7200
      Begin VB.PictureBox picYourTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   435
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   480
         Width           =   2895
      End
      Begin VB.PictureBox picTheirTrade 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   238
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   3705
         Left            =   3840
         ScaleHeight     =   247
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   193
         TabIndex        =   14
         TabStop         =   0   'False
         Top             =   480
         Width           =   2895
      End
      Begin VB.Label lblYourWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   1680
         TabIndex        =   18
         Top             =   4500
         Width           =   1815
      End
      Begin VB.Label lblTheirWorth 
         BackStyle       =   0  'Transparent
         Caption         =   "1234567890"
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5160
         TabIndex        =   17
         Top             =   4500
         Width           =   1815
      End
      Begin VB.Label lblTradeStatus 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   600
         TabIndex        =   16
         Top             =   5520
         Width           =   5895
      End
      Begin VB.Image imgAcceptTrade 
         Height          =   435
         Left            =   2475
         Top             =   5040
         Width           =   1035
      End
      Begin VB.Image imgDeclineTrade 
         Height          =   435
         Left            =   3675
         Top             =   5040
         Width           =   1035
      End
   End
   Begin VB.PictureBox picCurrency 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
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
      Height          =   2175
      Left            =   120
      ScaleHeight     =   145
      ScaleMode       =   0  'User
      ScaleWidth      =   484
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   7260
      Begin VB.TextBox txtCurrency 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2160
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   840
         Width           =   2775
      End
      Begin VB.Label lblCurrencyCancel 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   6
         Top             =   1440
         Width           =   2775
      End
      Begin VB.Label lblCurrencyOk 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2160
         TabIndex        =   5
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label lblCurrency 
         Alignment       =   2  'Center
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "How many do you want to drop?"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   30
         TabIndex        =   3
         Top             =   480
         Width           =   7155
      End
   End
   Begin VB.PictureBox picDialogue 
      Appearance      =   0  'Flat
      BackColor       =   &H000C0E10&
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
      Height          =   2175
      Left            =   120
      ScaleHeight     =   2175
      ScaleWidth      =   7260
      TabIndex        =   59
      TabStop         =   0   'False
      Top             =   6360
      Visible         =   0   'False
      Width           =   7260
      Begin VB.TextBox txtDialogue 
         Height          =   315
         Left            =   2520
         TabIndex        =   60
         TabStop         =   0   'False
         Top             =   960
         Visible         =   0   'False
         Width           =   2295
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "No"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   3
         Left            =   3405
         TabIndex        =   65
         Top             =   1665
         Width           =   285
      End
      Begin VB.Label lblDialogue_Title 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Trade Request"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   63
         Top             =   360
         Width           =   7215
      End
      Begin VB.Label lblDialogue_Text 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Player has requested a trade. Would you like to accept?"
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   0
         TabIndex        =   62
         Top             =   600
         Width           =   7215
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Ok"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   1
         Left            =   3405
         TabIndex        =   61
         Top             =   1560
         Width           =   285
      End
      Begin VB.Label lblDialogue_Button 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BackStyle       =   0  'Transparent
         Caption         =   "Yes"
         BeginProperty Font 
            Name            =   "Georgia"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   210
         Index           =   2
         Left            =   3360
         TabIndex        =   64
         Top             =   1440
         Width           =   345
      End
   End
   Begin VB.PictureBox picMapEditor 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   0
      ScaleHeight     =   62
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   798
      TabIndex        =   151
      Top             =   12510
      Width           =   12000
      Begin VB.CheckBox mapPreviewSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H0080C0FF&
         ForeColor       =   &H80000008&
         Height          =   540
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   153
         ToolTipText     =   "Map Preview - Docked"
         Top             =   315
         Width           =   540
      End
      Begin VB.Label lblTitle 
         BackColor       =   &H0080C0FF&
         BackStyle       =   0  'Transparent
         Caption         =   "UBER Map Editor"
         BeginProperty Font 
            Name            =   "Arial Black"
            Size            =   9
            Charset         =   0
            Weight          =   900
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Left            =   45
         TabIndex        =   152
         Top             =   30
         Width           =   1695
      End
   End
   Begin VB.Label lblSwearFilter 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Swear Filter"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6300
      TabIndex        =   111
      Top             =   9615
      Width           =   1200
   End
   Begin VB.Label lblWeather 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Weather"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6300
      TabIndex        =   110
      Top             =   9975
      Width           =   855
   End
   Begin VB.Label lblAutoTile 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Auto Tile"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6300
      TabIndex        =   109
      Top             =   10335
      Width           =   915
   End
   Begin VB.Label lblDebug 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Debug"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6300
      TabIndex        =   108
      Top             =   11055
      Width           =   600
   End
   Begin VB.Label lblBlood 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Blood"
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   6300
      TabIndex        =   107
      Top             =   10695
      Width           =   525
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ************
' ** Events **
' ************
Private MoveForm As Boolean
Private LastX As Long
Private LastY As Long
Private Declare Function GetForegroundWindow Lib "user32.dll" () As Long

Private Const WM_MOUSEMOVE      As Long = &H200
Private Const WM_LBUTTONDOWN    As Long = &H201
Private Const WM_LBUTTONUP      As Long = &H202
Private Const WM_CAPTURECHANGED As Long = &H215
Private Const WM_GETMINMAXINFO  As Long = &H24
Private Const WM_ACTIVATEAPP    As Long = &H1C
Private Const WM_SETFOCUS       As Long = &H7

Private WithEvents cSubclasserHooker As cSelfSubHookCallback
Attribute cSubclasserHooker.VB_VarHelpID = -1
Private taskBarClick As Boolean

Private Sub Form_Activate()
    hwndLastActiveWnd = hWnd
    If FormVisible("frmAdmin") And adminMin Then
        frmAdmin.centerMiniVert Width, Height, Left, Top
    End If
End Sub

Private Sub Form_Initialize()
    Set cSubclasserHooker = New cSelfSubHookCallback
    If cSubclasserHooker.ssc_Subclass(Me.hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg Me.hWnd, eMsgWhen.MSG_BEFORE, WM_ACTIVATEAPP, WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_CAPTURECHANGED, WM_GETMINMAXINFO
    End If
    
    If cSubclasserHooker.ssc_Subclass(Me.picMapEditor.hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg Me.picMapEditor.hWnd, eMsgWhen.MSG_BEFORE, WM_ACTIVATEAPP, WM_MOUSEMOVE, WM_LBUTTONDOWN, WM_LBUTTONUP, WM_CAPTURECHANGED, WM_GETMINMAXINFO
    End If
    
    If cSubclasserHooker.ssc_Subclass(Me.mapPreviewSwitch.hWnd, ByVal 1, 1, Me) Then
        cSubclasserHooker.ssc_AddMsg Me.mapPreviewSwitch.hWnd, eMsgWhen.MSG_BEFORE, WM_SETFOCUS
    End If
End Sub

Private Sub Form_Paint()
    If FormVisible("frmCharEditor") Then
        frmCharEditor.Show
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Cancel = True
    LogoutGame
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_Unload", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    Call ClearChatButton(0)
    ClearButtons
    
    ' Reset all buttons
    Call ResetMainButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgAcceptTrade_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picCurrency.Visible = False
    TmpCurrencyItem = 0
    CurrencyMenu = 0 ' Clear
    AcceptTrade
    Exit Sub
     
' Error handler
errorhandler:
    HandleError "imgAcceptTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgFix_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InShop = 0 Then Exit Sub
    If Shop(InShop).CanFix = 0 Then Exit Sub
    
    TryingToFixItem = True
    
    AddText "Double-click on the item in your inventory you wish to fix.", BrightGreen
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ImgFix_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub imgShopBuy_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ShopAction = 1 Then Exit Sub
    ShopAction = 1 ' buying an item
    AddText "Click on the item in the shop you wish to buy.", White
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "imgShopBuy_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub imgShopSell_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If ShopAction = 2 Then Exit Sub
    ShopAction = 2 ' selling an item
    
    AddText "Double-click on the item in your inventory you wish to sell.", BrightGreen
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "imgShopSell_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblAddFriend_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Dialogue "Add Friend", "Who do you want to add as a friend?", DIALOGUE_TYPE_ADDFRIEND, True
    
    If MyTargetType = TARGET_TYPE_PLAYER Then
        If MyTarget > 0 And MyTarget <= MAX_PLAYERS Then
            If Not MyTarget = MyIndex Then
                If IsPlaying(MyTarget) Then
                    txtDialogue.text = GetPlayerName(MyTarget)
                End If
            End If
        End If
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblAddFriend_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblAddFoe_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Dialogue "Add Foe", "Who do you want to add as a foe?", DIALOGUE_TYPE_ADDFOE, True
    
    If MyTargetType = TARGET_TYPE_PLAYER Then
        If MyTarget > 0 And MyTarget <= MAX_PLAYERS Then
            If Not MyTarget = MyIndex Then
                If IsPlaying(MyTarget) Then
                    txtDialogue.text = GetPlayerName(MyTarget)
                End If
            End If
        End If
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblAddFoe_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblChoices_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    
    buffer.WriteLong CEventChatReply
    buffer.WriteLong EventReplyID
    buffer.WriteLong EventReplyPage
    buffer.WriteLong Index
    SendData buffer.ToArray
    Set buffer = Nothing
    ClearEventChat
    
    Call ClearChatButton(Index)
    InEvent = False
    Audio.PlaySound ButtonClick
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblChoices_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblChoices_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ClearChatButton(Index)
    If frmMain.lblChoices(Index).Visible = False Then Exit Sub
    If frmMain.lblChoices.item(Index).ForeColor = vbYellow Then Exit Sub
    frmMain.lblChoices.item(Index).ForeColor = vbYellow
    Audio.PlaySound ButtonHover
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblChoices_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ClearChatButton(Index As Integer)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To 4
        If frmMain.lblChoices.item(i).ForeColor = vbYellow And Not Index = i Then
            frmMain.lblChoices.item(i).ForeColor = &H80000003
        End If
    Next
    
    frmMain.lblEventChatContinue.ForeColor = &H80000003
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ClearChatButton", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ClearButtons()
    LastButton_Main = 0
    ResetOptionButtons
    Call ResetMainButtons
End Sub

Private Sub lblEventChatContinue_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If frmMain.lblEventChatContinue.Visible = False Then Exit Sub
    If frmMain.lblEventChatContinue.ForeColor = vbYellow Then Exit Sub
    frmMain.lblEventChatContinue.ForeColor = vbYellow
    Audio.PlaySound ButtonHover
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblEventChatContinue_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub lblEventChatContinue_Click()
    Dim buffer As clsBuffer
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Set buffer = New clsBuffer
    buffer.WriteLong CEventChatReply
    buffer.WriteLong EventReplyID
    buffer.WriteLong EventReplyPage
    buffer.WriteLong 0
    SendData buffer.ToArray
    Set buffer = Nothing
    ClearEventChat
    InEvent = False
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblEventChatContinue_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Sub ClearEventChat()
    Dim i As Long
    
    If AnotherChat = 1 Then
        For i = 1 To 4
            frmMain.lblChoices(i).Visible = False
        Next
        
        frmMain.lblEventChat.Caption = ""
        frmMain.lblEventChatContinue.Visible = False
    ElseIf AnotherChat = 2 Then
        For i = 1 To 4
            frmMain.lblChoices(i).Visible = False
        Next
        
        frmMain.lblEventChat.Visible = False
        frmMain.lblEventChatContinue.Visible = False
        EventChatTimer = timeGetTime + 100
    Else
        frmMain.picEventChat.Visible = False
    End If
End Sub

Private Sub lblEquipCharName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picItemDesc.Visible = False
    LastItemDesc = 0 ' No Item was last loaded
End Sub

Private Sub lblGuildRemove_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Dialogue "Guild Remove", "Who do you want to remove from the guild?", DIALOGUE_TYPE_GUILDREMOVE, True
    
    If MyTargetType = TARGET_TYPE_PLAYER Then
        If MyTarget > 0 And MyTarget <= MAX_PLAYERS Then
            If Not MyTarget = MyIndex Then
                If IsPlaying(MyTarget) Then
                    txtDialogue.text = GetPlayerName(MyTarget)
                End If
            End If
        End If
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblGuildRemove_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblItemDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picItemDesc.Visible = False
    LastItemDesc = 0 ' No Item was last loaded
End Sub

Private Sub lblItemName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picItemDesc.Visible = False
    LastItemDesc = 0 ' No Item was last loaded
End Sub

Private Sub lblRemoveFriend_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Dialogue "Remove Friend", "What friend do you want to remove?", DIALOGUE_TYPE_REMOVEFRIEND, True
    
    If (lstFriends.ListIndex + 1) > 0 And lstFriends.ListIndex + 1 <= MAX_PEOPLE Then
        txtDialogue.text = Trim$(Player(MyIndex).Friends(lstFriends.ListIndex + 1).name)
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblRemoveFriend_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblRemoveFoe_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Dialogue "Remove Foe", "What foe do you want to remove?", DIALOGUE_TYPE_REMOVEFOE, True
    
    If (lstFoes.ListIndex + 1) > 0 And lstFoes.ListIndex + 1 <= MAX_PEOPLE Then
        txtDialogue.text = Trim$(Player(MyIndex).Foes(lstFoes.ListIndex + 1).name)
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblRemoveFoe_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblChangeAccess_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Dialogue "Change Guild Access", "What access would you like to change this user to?", DIALOGUE_TYPE_CHANGEGUILDACCESS, True
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblChangeAccess_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblCurrencyCancel_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picCurrency.Visible = False
    txtCurrency.text = vbNullString
    TmpCurrencyItem = 0
    CurrencyMenu = 0 ' Clear
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblCurrencyCancel_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgDeclineTrade_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    CloseTrade
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ImgDeclineTrade_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblLeaveBank_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    CloseBank
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblLeaveBank_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub ImgLeaveShop_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InShop = 0 Then Exit Sub
    CloseShop
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ImgLeaveShop_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblCurrencyOk_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsNumeric(txtCurrency.text) Then
        Select Case CurrencyMenu
            Case 1 ' Drop item
                SendDropItem TmpCurrencyItem, Val(txtCurrency.text)
            Case 2 ' Deposit item
                DepositItem TmpCurrencyItem, Val(txtCurrency.text)
            Case 3 ' withdraw item
                WithdrawItem TmpCurrencyItem, Val(txtCurrency.text)
            Case 4 ' Offer trade item
                TradeItem TmpCurrencyItem, Val(txtCurrency.text)
        End Select
    Else
        AddText "Please enter a valid amount.", BrightRed
        Exit Sub
    End If
    
    picCurrency.Visible = False
    TmpCurrencyItem = 0
    txtCurrency.text = vbNullString
    CurrencyMenu = 0 ' Clear
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblCurrencyOk_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblDialogue_Button_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Call the handler
    DialogueHandler Index
    
    txtDialogue.text = vbNullString
    picDialogue.Visible = False
    DialogueIndex = 0
    SetGameFocus
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblDialogue_Button_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblPartyInvite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Dialogue "Party Invite", "Who do you want to invite to the party?", DIALOGUE_TYPE_PARTYINVITE, True
    
    If MyTargetType = TARGET_TYPE_PLAYER Then
        If MyTarget > 0 And MyTarget <= MAX_PLAYERS Then
            If Not MyTarget = MyIndex Then
                If IsPlaying(MyTarget) Then
                    txtDialogue.text = GetPlayerName(MyTarget)
                End If
            End If
        End If
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblPartyLeave_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Party.num > 0 Then
        SendPartyLeave
    Else
        AddText "You are not in a party.", BrightRed
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblPartyInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblGuildInvite_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Dialogue "Guild Invite", "Who do you want to invite to the guild?", DIALOGUE_TYPE_GUILDINVITE, True
    
    If MyTargetType = TARGET_TYPE_PLAYER Then
        If MyTarget > 0 And MyTarget <= MAX_PLAYERS Then
            If Not MyTarget = MyIndex Then
                If IsPlaying(MyTarget) Then
                    txtDialogue.text = GetPlayerName(MyTarget)
                End If
            End If
        End If
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblGuildInvite_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblResign_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    RequestGuildResign
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblResign_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lblSpellDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
End Sub

Private Sub lblSpellName_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
End Sub

Private Sub lblTrainStat_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GetPlayerPOINTS(MyIndex) = 0 Then Exit Sub
    SendTrainStat Index
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblTrainStat_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub lstDropDownBox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (TmpInvNum < 1 Or TmpInvNum > MAX_INV) And (TmpSpellSlot < 1 Or TmpSpellSlot > MAX_PLAYER_SPELLS) Then Exit Sub
    If lstDropDownBox.ListIndex = -1 Then Exit Sub
    
    If TmpInvNum > 0 Then
        If lstDropDownBox.List(lstDropDownBox.ListIndex) = "Use" Or lstDropDownBox.List(lstDropDownBox.ListIndex) = "Equip" Then
            Call SendUseItem(TmpInvNum)
        ElseIf lstDropDownBox.List(lstDropDownBox.ListIndex) = "Drop" Then
            Call DropItem(TmpInvNum)
        'ElseIf lstDropDownBox.List(lstDropDownBox.ListIndex) = "Examine" Then
        '    If Trim$(Item(GetPlayerInvItemNum(MyIndex, TmpInvNum)).Desc) = vbNullString Then
        '        Call AddText("This item does not have a description, report this to a Staff member!", BrightRed)
        '    Else
        '        Call AddText(Trim$(Item(GetPlayerInvItemNum(MyIndex, TmpInvNum)).Desc), Yellow)
        '    End If
        End If
    Else
        If lstDropDownBox.List(lstDropDownBox.ListIndex) = "Cast" Then
            Call SendCastSpell(TmpSpellSlot)
        ElseIf lstDropDownBox.List(lstDropDownBox.ListIndex) = "Forget" Then
            Dialogue "Forget Spell", "Are you sure you want to forget how to cast " & Trim$(Spell(PlayerSpells(TmpSpellSlot)).name) & "?", DIALOGUE_TYPE_FORGET, True, TmpSpellSlot
        'ElseIf lstDropDownBox.List(lstDropDownBox.ListIndex) = "Examine" Then
        '    If Trim$(Spell(PlayerSpells(TmpSpellSlot)).Desc) = vbNullString Then
        '        Call AddText("This spell does not have a description, report this to a Staff member!", BrightRed)
        '    Else
        '        Call AddText(Trim$(Spell(PlayerSpells(TmpSpellSlot)).Desc), Yellow)
        '    End If
        End If
    End If
    Call Audio.PlaySound(ButtonClick)
    lstDropDownBox.Visible = False
End Sub

Private Sub lstTitles_Click()
    Dim i As Byte
    
    ' Check if we're setting it to one we already have as our current title
    If lstTitles.ListIndex = Player(MyIndex).CurTitle Then Exit Sub
        
    If Not lstTitles.ListIndex = 0 Then
        For i = 1 To MAX_TITLES
            If Not Player(MyIndex).CurTitle = i Then
                If lstTitles.List(lstTitles.ListIndex) = Trim$(title(i).name) Then
                    lblDesc.Caption = Trim$(title(i).Desc)
                    Call SendSetTitle(i)
                    Exit For
                End If
            End If
        Next
    Else
        lblDesc.Caption = "None."
        Call SendSetTitle(0)
    End If
End Sub

Private Sub lstTitles_GotFocus()
    SetGameFocus
End Sub

Private Sub lstFoes_GotFocus()
    SetGameFocus
End Sub

Private Sub lstFriends_GotFocus()
    SetGameFocus
End Sub

Private Sub lstGuild_GotFocus()
    SetGameFocus
End Sub

Private Sub mapPreviewSwitch_Click()
    If mapPreviewSwitch.Value Then
        mapPreviewSwitch.Picture = LoadResPicture("MAP_DOWN", vbResBitmap)
        frmMapPreview.Show
    Else
        mapPreviewSwitch.Picture = LoadResPicture("MAP_UP", vbResBitmap)
        Unload frmMapPreview
    End If
End Sub

Private Sub picChatbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ClearButtons
    ResetOptionButtons
End Sub

Private Sub picEquipFace_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    picItemDesc.Visible = False
    LastItemDesc = 0 ' No Item was last loaded
End Sub

Private Sub picOptionBlood_Click()
    If Options.Blood = 0 Then
        Options.Blood = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Blood = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionBlood, OptionButtons.Opt_Blood, Options.Blood)
End Sub

Private Sub picOptionDebug_Click()
    If Options.Debug = 0 Then
        Options.Debug = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Debug = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionDebug, OptionButtons.Opt_Debug, Options.Debug)
End Sub

Private Sub picOptionSwearFilter_Click()
    If Options.SwearFilter = 0 Then
        Options.SwearFilter = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.SwearFilter = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionSwearFilter, OptionButtons.Opt_SwearFilter, Options.SwearFilter)
End Sub

Private Sub picOptionSound_Click()
    If Options.Sound = 0 Then
        Options.Sound = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Sound = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionSound, OptionButtons.Opt_Sound, Options.Sound)
End Sub

Private Sub picOptionMouse_Click()
    If Options.Mouse = 0 Then
        Options.Mouse = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Mouse = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    MouseX = -1
    MouseY = -1
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionMouse, OptionButtons.Opt_Mouse, Options.Mouse)
End Sub

Private Sub picOptionMusic_Click()
    If Options.Music = 0 Then
        Options.Music = 1
        Call Audio.PlaySound(ButtonClick)
        
        ' Start playing music
        PlayMapMusic
    Else
        Options.Music = 0
        Call Audio.PlaySound(ButtonBuzzer)
        
        ' Stop playing music
        Audio.StopMusic
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionMusic, OptionButtons.Opt_Music, Options.Music)
End Sub

Private Sub picOptionWeather_Click()
    If Options.Weather = 0 Then
        Options.Weather = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Weather = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionWeather, OptionButtons.Opt_Weather, Options.Weather)
End Sub

Private Sub picOptionAutoTile_Click()
    If Options.Autotile = 0 Then
        Options.Autotile = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Autotile = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionAutoTile, OptionButtons.Opt_AutoTile, Options.Autotile)
End Sub

Private Sub picOptionBattleMusic_Click()
    If Options.BattleMusic = 0 Then
        Options.BattleMusic = 1
        Call Audio.PlaySound(ButtonClick)
        
        ' Start playing music
        PlayMapMusic
    Else
        Options.BattleMusic = 0
        Call Audio.PlaySound(ButtonBuzzer)
        If Trim(Map.Music) = vbNullString Then
            Call Audio.StopMusic
        Else
            Call Audio.PlayMusic(Trim$(Map.Music))
        End If
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionBattleMusic, OptionButtons.Opt_BattleMusic, Options.BattleMusic)
End Sub

Private Sub picOptionTitle_Click()
    If Options.Titles = 0 Then
        Options.Titles = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Titles = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionTitle, OptionButtons.Opt_Title, Options.Titles)
End Sub

Private Sub picOptionPlayerVitals_Click()
    If Options.PlayerVitals = 0 Then
        Options.PlayerVitals = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.PlayerVitals = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionPlayerVitals, OptionButtons.Opt_PlayerVitals, Options.PlayerVitals)
End Sub

Private Sub picOptionNpcVitals_Click()
    If Options.NpcVitals = 0 Then
        Options.NpcVitals = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.NpcVitals = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionNpcVitals, OptionButtons.Opt_NpcVitals, Options.NpcVitals)
End Sub

Private Sub picOptionLevel_Click()
    If Options.Levels = 0 Then
        Options.Levels = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Levels = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionLevel, OptionButtons.Opt_Level, Options.Levels)
End Sub

Private Sub picOptionGuild_Click()
    If Options.Guilds = 0 Then
        Options.Guilds = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.Guilds = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionGuild, OptionButtons.Opt_Guilds, Options.Guilds)
End Sub

Private Sub picOptionWASD_Click()
    If Options.WASD = 0 Then
        Options.WASD = 1
        Call Audio.PlaySound(ButtonClick)
    Else
        Options.WASD = 0
        Call Audio.PlaySound(ButtonBuzzer)
    End If
    SaveOptions
    SetGameFocus
    
    Call RenderOptionButton(picOptionWASD, OptionButtons.Opt_WASD, Options.WASD)
End Sub

Private Sub picOptionBlood_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Blood)
    If OptionButton(OptionButtons.Opt_Blood).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionBlood, OptionButtons.Opt_Blood, 2 + Options.Blood)
End Sub

Private Sub picOptionDebug_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Debug)
    If OptionButton(OptionButtons.Opt_Debug).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionDebug, OptionButtons.Opt_Debug, 2 + Options.Debug)
End Sub

Private Sub picOptionSwearFilter_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_SwearFilter)
    If OptionButton(OptionButtons.Opt_SwearFilter).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionSwearFilter, OptionButtons.Opt_SwearFilter, 2 + Options.SwearFilter)
End Sub

Private Sub picOptionSound_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Sound)
    If OptionButton(OptionButtons.Opt_Sound).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionSound, OptionButtons.Opt_Sound, 2 + Options.Sound)
End Sub

Private Sub picOptionMouse_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Mouse)
    If OptionButton(OptionButtons.Opt_Mouse).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionMouse, OptionButtons.Opt_Mouse, 2 + Options.Mouse)
End Sub

Private Sub picOptionMusic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Music)
    If OptionButton(OptionButtons.Opt_Music).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionMusic, OptionButtons.Opt_Music, 2 + Options.Music)
End Sub

Private Sub picOptionWeather_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Weather)
    If OptionButton(OptionButtons.Opt_Weather).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionWeather, OptionButtons.Opt_Weather, 2 + Options.Weather)
End Sub

Private Sub picOptionBattleMusic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_BattleMusic)
    If OptionButton(OptionButtons.Opt_BattleMusic).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionBattleMusic, OptionButtons.Opt_BattleMusic, 2 + Options.BattleMusic)
End Sub

Private Sub picOptionTitle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Title)
    If OptionButton(OptionButtons.Opt_Title).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionTitle, OptionButtons.Opt_Title, 2 + Options.Titles)
End Sub

Private Sub picOptionPlayerVitals_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_PlayerVitals)
    If OptionButton(OptionButtons.Opt_PlayerVitals).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionPlayerVitals, OptionButtons.Opt_PlayerVitals, 2 + Options.PlayerVitals)
End Sub

Private Sub picOptionNpcVitals_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_NpcVitals)
    If OptionButton(OptionButtons.Opt_NpcVitals).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionNpcVitals, OptionButtons.Opt_NpcVitals, 2 + Options.NpcVitals)
End Sub

Private Sub picOptionLevel_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Level)
    If OptionButton(OptionButtons.Opt_Level).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionLevel, OptionButtons.Opt_Level, 2 + Options.Levels)
End Sub

Private Sub picOptionGuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_Guilds)
    If OptionButton(OptionButtons.Opt_Guilds).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionGuild, OptionButtons.Opt_Guilds, 2 + Options.Guilds)
End Sub

Private Sub picOptionWASD_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_WASD)
    If OptionButton(OptionButtons.Opt_WASD).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionWASD, OptionButtons.Opt_WASD, 2 + Options.WASD)
End Sub

Private Sub picOptionAutoTile_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetOptionButtons(OptionButtons.Opt_AutoTile)
    If OptionButton(OptionButtons.Opt_AutoTile).State > 1 Then Exit Sub
    Call Audio.PlaySound(ButtonHover)
    Call RenderOptionButton(picOptionAutoTile, OptionButtons.Opt_AutoTile, 2 + Options.Autotile)
End Sub

Private Sub picEventChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ClearChatButton(0)
    ClearButtons
    ResetOptionButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picEventChat_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ToggleChatLock(Optional ByVal ForceLock As Boolean, Optional ByVal SoundEffect As Boolean = True)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If ForceLock Then
        ChatLocked = True
    Else
        ChatLocked = Not ChatLocked
    End If
    
    If ChatLocked Then
        If SoundEffect Then Call Audio.PlaySound(ButtonBuzzer)
        frmMain.txtMyChat.text = vbNullString
        frmMain.txtMyChat.Enabled = False
        Exit Sub
    Else
        If SoundEffect Then Call Audio.PlaySound(ButtonClick)
        frmMain.txtMyChat.Enabled = True
    End If
    
    Call SetGameFocus
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ToggleChatLock", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picButton_Click(Index As Integer)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not CurButton_Main = Index Then
        Call Audio.PlaySound(ButtonClick)
        
        ' Don't set it if it's the trade/GUI adjusting button
        If Not Index = 5 And Not Index = 14 And Not Index = 15 Then
            CurButton_Main = Index
            picButton(Index).Picture = LoadPicture(App.Path & GFX_PATH & "gui\main\buttons\" & MainButton(Index).FileName & "_click.jpg")
            Call ResetMainButtons
        End If
        
        Call TogglePanel(Index)
    Else ' Hide the panel, if it is the open one
        CurButton_Main = 0
        LastButton_Main = 0
        Call ResetMainButtons
        Call Audio.PlaySound(ButtonClick)
        Call TogglePanel(0)
    End If
    SetGameFocus
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picButton_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picButton_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If Not LastButton_Main = Index And Not CurButton_Main = Index Then
        Call ResetMainButtons
        picButton(Index).Picture = LoadPicture(App.Path & GFX_PATH & "gui\main\buttons\" & MainButton(Index).FileName & "_hover.jpg")
        Call Audio.PlaySound(ButtonHover)
        LastButton_Main = Index
    End If
    Call ClearChatButton(0)
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picButton_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub TogglePanel(ByVal PanelNum As Long)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Don't close panels if it's the trade button
    If Not PanelNum = 5 Then
        Call CloseAllPanels
    End If
    
    Select Case PanelNum
        Case 1
            picInventory.Visible = True
            picInventory.ZOrder (0)
        Case 2
            picSpells.Visible = True
            picSpells.ZOrder (0)
        Case 3
            picCharacter.Visible = True
            picCharacter.ZOrder (0)
        Case 4
            picOptions.Visible = True
            picOptions.ZOrder (0)
        Case 5
            If MyTargetType = TARGET_TYPE_PLAYER And Not MyTarget = MyIndex Then
                SendTradeRequest
            Else
                AddText "Invalid trade target.", BrightRed
            End If
        Case 6
            picParty.Visible = True
            picParty.ZOrder (0)
        Case 7
            picFriends.Visible = True
            picFriends.ZOrder (0)
        Case 8
            If GetPlayerGuild(MyIndex) = vbNullString Then
                picGuild_No.Visible = True
                picGuild_No.ZOrder (0)
            Else
                picGuild.Visible = True
                picGuild.ZOrder (0)
            End If
        Case 10
            picTitles.Visible = True
            picTitles.ZOrder (0)
        Case 12
            picFoes.Visible = True
            picFoes.ZOrder (0)
        Case 14
            ButtonsVisible = Not ButtonsVisible
            If ButtonsVisible Then
                MainButton(14).FileName = "btn_hidepanels"
            Else
                MainButton(14).FileName = "btn_showpanels"
            End If
            Call ResetMainButtons
            Call ToggleButtons(ButtonsVisible)
        Case 15
            GUIVisible = Not GUIVisible
            If GUIVisible Then
                MainButton(15).FileName = "btn_hidegui"
            Else
                MainButton(15).FileName = "btn_showgui"
            End If
            Call ResetMainButtons
            Call ToggleGUI(GUIVisible)
        Case 16
            picEquipment.Visible = True
            picEquipment.ZOrder (0)
    End Select
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "TogglePanel", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub ResetMainButtons()
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To MAX_MAINBUTTONS
        If Not CurButton_Main = i Then
            picButton(i).Picture = LoadPicture(App.Path & GFX_PATH & "gui\main\buttons\" & MainButton(i).FileName & "_norm.jpg")
        End If
    Next
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "ResetMainButtons", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picForm_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    lstDropDownBox.Visible = False
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picForm_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picForm_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call SetGameFocus
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picForm_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picFriends_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picFriends_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picGuild_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picGuild_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picHotbar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long, rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ' Hotbar
    For i = 1 To MAX_HOTBAR
        With rec_pos
            .Top = picHotbar.Top - picHotbar.Top
            .Left = picHotbar.Left - picHotbar.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
            .Right = .Left + 32
            .Bottom = picHotbar.Top - picHotbar.Top + 32
        End With
        
        If X >= rec_pos.Left And X <= rec_pos.Right Then
            If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                SendSwapHotbarSlots DragHotbarSlot, i
            End If
        End If
    Next
    
    DragHotbarSlot = 0
    picTempInv.Visible = False
    picTempSpell.Visible = False
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picHotbar_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picHotbar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SlotNum As Long, i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SlotNum = IsHotbarSlot(X, Y)

    If SlotNum > 0 Then
        If Button = 1 Then
            If ShiftDown Then
                DragHotbarSlot = SlotNum
                
                For i = 1 To MAX_PLAYER_SPELLS
                    If Hotbar(DragHotbarSlot).Slot = PlayerSpells(i) Then
                        DragHotbarSpell = i
                    End If
                Next
            Else
                SendHotbarUse SlotNum
            End If
        ElseIf Button = 2 Then
            SendHotbarChange 0, 0, SlotNum
        End If
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picHotbar_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picHotbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SlotNum As Long, i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DragHotbarSlot > 0 Then
        If Hotbar(DragHotbarSlot).sType = 1 Then
            Call DrawDraggedItem(X + picHotbar.Left - 16, Y + picHotbar.Top - 16, True)
        Else
            Call DrawDraggedSpell(X + picHotbar.Left - 16, Y + picHotbar.Top - 16, True)
        End If
        picSpellDesc.Visible = False
        picItemDesc.Visible = False
        LastSpellDesc = 0 ' No spell was last loaded
        LastItemDesc = 0 ' No item was last loaded
        Exit Sub
    Else
        SlotNum = IsHotbarSlot(X, Y)
        
        If SlotNum <> 0 Then
              If Hotbar(SlotNum).sType = 1 Then ' item
                X = X + picHotbar.Left - picItemDesc.Width - 1
                Y = Y + picHotbar.Top
                UpdateItemDescWindow Hotbar(SlotNum).Slot, X, Y
                LastItemDesc = Hotbar(SlotNum).Slot ' Set it so you don't re-set values
                Exit Sub
              ElseIf Hotbar(SlotNum).sType = 2 Then ' spell
                X = X + picHotbar.Left - picSpellDesc.Width - 1
                Y = Y + picHotbar.Top
                UpdateSpellDescWindow Hotbar(SlotNum).Slot, X, Y
                LastSpellDesc = Hotbar(SlotNum).Slot

                For i = 1 To MAX_PLAYER_SPELLS
                    If Hotbar(SlotNum).Slot = PlayerSpells(i) Then
                        LastSpellSlotDesc = i
                    End If
                Next
                Exit Sub
              End If
          End If
    End If
    
    Call ClearChatButton(0)
    ClearButtons
    picSpellDesc.Visible = False
    picItemDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    picTempInv.Visible = False
    picTempSpell.Visible = False
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picHotbar_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picOptions_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picOptions_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picParty_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picParty_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picPet_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picPet_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picScreen_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    If InMapEditor Then
        If frmEditor_Map.chkEyeDropper.Value = 1 Then
            Call MapEditorEyeDropper
        Else
            If ControlDown And Button = 1 Then
                MapEditorFillSelection
                Exit Sub

            ElseIf ControlDown And Button = 2 Then
                MapEditorClearSelection
                Exit Sub

            ElseIf ShiftDown And Button = 1 Then
                MapEditorEyeDropper
                Exit Sub

            ElseIf Button = vbRightButton Then
                If ShiftDown Then
                    ' Admin warp if we're pressing shift and right clicking
                    If GetPlayerAccess(MyIndex) >= STAFF_MAPPER Then
                        If CanMoveNow Then
                            AdminWarp CurX, CurY
                        End If
                    End If
                End If
            End If
            
            Call MapEditorMouseDown(Button, X, Y, False)
            redrawMapCache = True
        End If
    Else
        ' Left click
        If Button = vbLeftButton Then
            ' Targetting
            Call PlayerSearch(CurX, CurY)
            ' Right click
        ElseIf Button = vbRightButton Then
            If ShiftDown Then
                ' Admin warp if we're pressing shift and right clicking
                If GetPlayerAccess(MyIndex) >= STAFF_MAPPER Then
                    If CanMoveNow Then
                        AdminWarp CurX, CurY
                    End If
                End If
            ElseIf InMapEditor Then
                DeleteEvent CurX, CurY
            End If
        End If
    End If

    Call SetGameFocus
    frmMain.picSpellDesc.Visible = False
    frmMain.picItemDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    Call ClearChatButton(0)
    ClearButtons
    ResetOptionButtons
    Exit Sub
    
    ' Error handler
errorhandler:
    HandleError "picScreen_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picScreen_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    CurX = TileView.Left + ((X + Camera.Left) \ PIC_X)
    CurY = TileView.Top + ((Y + Camera.Top) \ PIC_Y)
    
    If InMapEditor Then
        Call MapEditorMouseDown(Button, X, Y, False)
        If (LastX <> CurX Or LastY <> CurY) And frmEditor_Map.chkRandom.Value = 0 And Button >= 1 Then
            redrawMapCache = True
        End If
    ElseIf Button = vbLeftButton And Options.Mouse = 1 Then
        ' Mouse
        If CurX = GetPlayerX(MyIndex) And CurY = GetPlayerY(MyIndex) Then
            Call CheckMapGetItem
        Else
            MouseX = CurX
            MouseY = CurY
        End If
    End If
    
    LastX = CurX
    LastY = CurY
    
    ' Set the description windows off
    lstDropDownBox.Visible = False
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    Call ClearChatButton(0)
    LastButton_Main = 0
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picScreen_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Function IsShopItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_TRADES
        If Shop(InShop).TradeItem(i).item > 0 And Shop(InShop).TradeItem(i).item <= MAX_ITEMS Then
            With TempRec
                .Top = ShopTop + ((ShopOffsetY + PIC_Y) * ((i - 1) \ ShopColumns))
                .Bottom = .Top + PIC_Y
                .Left = ShopLeft + ((ShopOffsetX + PIC_X) * (((i - 1) Mod ShopColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsShopItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
errorhandler:
    HandleError "IsShopItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Sub picShop_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    LastItemDesc = 0 ' No item was last loaded
    
    ' Reset all buttons
    Call ResetMainButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picShop_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picShopItems_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ShopItem As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ShopItem = IsShopItem(X, Y)
    
    If ShopItem > 0 Then
        Select Case ShopAction
            Case 0 ' no action, give cost
                With Shop(InShop).TradeItem(ShopItem)
                    If .CostItem > 0 And .CostItem2 = 0 Then
                        AddText "You can buy this item for " & .CostValue & " " & Trim$(item(.CostItem).name) & ".", BrightGreen
                    ElseIf .CostItem2 > 0 And .CostItem = 0 Then
                        AddText "You can buy this item for " & .CostValue & " " & Trim$(item(.CostItem).name) & ".", BrightGreen
                    ElseIf .CostItem > 0 And .CostItem2 > 0 Then
                        AddText "You can buy this item for " & .CostValue & " " & Trim$(item(.CostItem).name) & " and " & .CostValue2 & " " & Trim$(item(.CostItem2).name) & ".", BrightGreen
                    Else
                        Exit Sub
                    End If
                End With
            Case 1 ' buy item
                ' buy item code
                BuyItem ShopItem
        End Select
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picShopItems_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picShopItems_dblClick()
    Dim ShopItem As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ShopItem = IsShopItem(ShopX, ShopY)
    
    If ShopItem > 0 Then
        BuyItem ShopItem
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picShopItems_dblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picShopItems_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim ShopSlot As Long
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    ShopX = X
    ShopY = Y
    
    ShopSlot = IsShopItem(X, Y)

    If ShopSlot <> 0 Then
        X2 = X + picShop.Left + picShopItems.Left + 1
        Y2 = Y + picShop.Top + picShopItems.Top + 1
        UpdateItemDescWindow Shop(InShop).TradeItem(ShopSlot).item, X2, Y2
        LastItemDesc = Shop(InShop).TradeItem(ShopSlot).item
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picShopItems_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picSpellDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picSpellDesc.Visible = False
    LastSpellDesc = 0 ' No spell was last loaded
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picSpellDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picSpells_DblClick()
    Dim SpellNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InTrade > 0 Or InBank Or InShop > 0 Or InChat Then Exit Sub

    SpellNum = IsPlayerSpell(SpellX, SpellY)

    If SpellNum > 0 And SpellNum <= MAX_SPELLS Then
        Call CastSpell(SpellNum)
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picSpells_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picSpells_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SpellSlot As Byte
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SpellX = X
    SpellY = Y
    
    SpellSlot = IsPlayerSpell(X, Y)
    
    If DragSpellSlot > 0 Then
        Call DrawDraggedSpell(X + picSpells.Left - 16, Y + picSpells.Top - 16)
    Else
        If SpellSlot <> 0 Then
            X2 = picSpells.Left - picSpellDesc.Width - 4
            Y2 = picSpells.Top
            UpdateSpellDescWindow PlayerSpells(SpellSlot), X2, Y2
            LastSpellDesc = PlayerSpells(SpellSlot)
            LastSpellSlotDesc = SpellSlot
            Exit Sub
        End If
    End If
    
    lstDropDownBox.Visible = False
    picSpellDesc.Visible = False
    LastSpellDesc = 0
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picSpells_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picSpells_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim rec_pos As RECT

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DragSpellSlot > 0 Then
        ' Drag and Drop
        For i = 1 To MAX_PLAYER_SPELLS
            With rec_pos
                .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + PIC_X) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    If Not DragSpellSlot = i Then
                        If Not DialogueIndex = DIALOGUE_TYPE_FORGET Then
                            SendChangeSpellSlots DragSpellSlot, i
                        End If
                        Exit For
                    End If
                End If
            End If
        Next
        
        ' Hotbar
        For i = 1 To MAX_HOTBAR
            With rec_pos
                .Top = picHotbar.Top - picSpells.Top
                .Left = picHotbar.Left - picSpells.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.Top - picSpells.Top + 32
            End With
            
            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 2, DragSpellSlot, i
                    DragSpellSlot = 0
                    picTempSpell.Visible = False
                    Exit Sub
                End If
            End If
        Next
    End If

    DragSpellSlot = 0
    picTempSpell.Visible = False
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picSpells_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picSpells_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SpellNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picSpellDesc.Visible = False
    LastSpellDesc = 0
    SpellNum = IsPlayerSpell(SpellX, SpellY)
    
    If Button = 1 Then ' left click
        If SpellNum <> 0 Then
            DragSpellSlot = SpellNum
            Exit Sub
        End If
    ElseIf Button = 2 Then ' Right click
        If SpellNum > 0 And SpellNum <= MAX_PLAYER_SPELLS Then
            X = X + picSpells.Left
            Y = Y + picSpells.Top
            lstDropDownBox.Top = Y
            lstDropDownBox.Left = X
            
            ' If the original height is stored then set the height and clear it
            If TmplstDropDownBoxHeight > 0 Then
                lstDropDownBox.Height = TmplstDropDownBoxHeight
                TmplstDropDownBoxHeight = 0
            End If
            
            ' Set the height
            TmplstDropDownBoxHeight = lstDropDownBox.Height
            
            ' Clear the list
            lstDropDownBox.Clear
            
            ' Build the list
            lstDropDownBox.AddItem "Cast"
            lstDropDownBox.AddItem "Forget"
            'lstDropDownBox.AddItem "Examine"
            
            ' Set the new height
            lstDropDownBox.Height = lstDropDownBox.Height * lstDropDownBox.ListCount
            
            ' Other stuff
            lstDropDownBox.ListIndex = -1
            lstDropDownBox.ZOrder (0)
            lstDropDownBox.Visible = True
            
            ' Store the spell number for future use
            TmpSpellSlot = SpellNum
            TmpInvNum = 0
        End If
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picSpells_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picToggleButtons_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ResetMainButtons
End Sub

Private Sub picTitles_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call ClearChatButton(0)
    ClearButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picTitles_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Hide the descriptions
    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picYourTrade_DblClick()
Dim TradeNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(TradeX, TradeY, True)

    If TradeNum <> 0 Then
        UntradeItem TradeNum
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picYourTrade_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picYourTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeX = X
    TradeY = Y
    
    TradeNum = IsTradeItem(X, Y, True)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picYourTrade.Left + 4
        Y = Y + picTrade.Top + picYourTrade.Top + 4
        UpdateItemDescWindow GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num), X, Y
        LastItemDesc = GetPlayerInvItemNum(MyIndex, TradeYourOffer(TradeNum).num) ' Set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picYourTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picTheirTrade_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim TradeNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    TradeNum = IsTradeItem(X, Y, False)
    
    If TradeNum <> 0 Then
        X = X + picTrade.Left + picTheirTrade.Left + 4
        Y = Y + picTrade.Top + picTheirTrade.Top + 4
        UpdateItemDescWindow TradeTheirOffer(TradeNum).num, X, Y
        LastItemDesc = TradeTheirOffer(TradeNum).num ' Set it so you don't re-set values
        Exit Sub
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0 ' no item was last loaded
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picTheirTrade_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' Winsock event
Private Sub Socket_DataArrival(ByVal bytesTotal As Long)

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If IsConnected Then
        Call IncomingData(bytesTotal)
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Socket_DataArrival", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Dim i As Long
    Dim NpcDistanceX(1 To MAX_MAP_NPCS) As Long
    Dim NpcDistanceY(1 To MAX_MAP_NPCS) As Long
    Dim PlayerDistanceX(1 To MAX_PLAYERS) As Long
    Dim PlayerDistanceY(1 To MAX_PLAYERS) As Long
    Dim LowestDistance As Long
    Dim PlayerTarget As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If GettingMap Then Exit Sub

    ' Set focus if making it visible
    If KeyAscii = vbKeyReturn Then
        If picEventChat.Visible Then
            If frmMain.lblEventChatContinue.Visible Then
                frmMain.lblEventChatContinue_Click
                KeyAscii = 0
                Exit Sub
            End If
        End If
        
        If picChatbox.Visible Then
            If txtMyChat.text = vbNullString Then
                If picCurrency.Visible = False Then
                    If picDialogue.Visible = False Then
                        Call ToggleChatLock
                        KeyAscii = 0
                        Exit Sub
                    End If
                End If
            End If
        End If
    End If

    If KeyAscii = vbKeyTab And ShiftDown = False Then
        ' Set the NPC distance for all the NPCs on the map
        For i = 1 To Map.Npc_HighIndex
            NpcDistanceX(i) = MapNPC(i).X - GetPlayerX(MyIndex)
            NpcDistanceY(i) = MapNPC(i).Y - GetPlayerY(MyIndex)
    
            ' Make sure we get a positive Value
            If NpcDistanceX(i) < 0 Then NpcDistanceX(i) = NpcDistanceX(i) * -1
            If NpcDistanceY(i) < 0 Then NpcDistanceY(i) = NpcDistanceY(i) * -1
        Next
        
        ' Clear the target constant
        PlayerTarget = 0
        
        ' Find the closest NPC target
        For i = 1 To Map.Npc_HighIndex
            If i = 1 Then
                LowestDistance = NpcDistanceX(i) + NpcDistanceY(i)
                PlayerTarget = i
            ElseIf NpcDistanceX(i) + NpcDistanceY(i) < LowestDistance Then
                LowestDistance = NpcDistanceX(i) + NpcDistanceY(i)
                PlayerTarget = i
            End If
        Next
        
        ' Set the target
        If PlayerTarget > 0 Then
            If Not MyTarget = PlayerTarget Then
                Call PlayerSearch(MapNPC(PlayerTarget).X, MapNPC(PlayerTarget).Y)
            End If
        End If
    End If
    
    If KeyAscii = vbKeyTab And ShiftDown Then
        ' Set the Player distance for all the Players on the map
        For i = 1 To Player_HighIndex
            PlayerDistanceX(i) = Player(i).X - GetPlayerX(MyIndex)
            PlayerDistanceY(i) = Player(i).Y - GetPlayerY(MyIndex)
    
            ' Make sure we get a positive Value
            If PlayerDistanceX(i) < 0 Then PlayerDistanceX(i) = PlayerDistanceX(i) * -1
            If PlayerDistanceY(i) < 0 Then PlayerDistanceY(i) = PlayerDistanceY(i) * -1
        Next
        
        ' Clear the target constant
        PlayerTarget = 0
        
        ' Find the closest Player target
        For i = 1 To Player_HighIndex
            If Not i = MyIndex Then
                If i = 1 Then
                    LowestDistance = PlayerDistanceX(i) + PlayerDistanceY(i)
                    PlayerTarget = i
                ElseIf PlayerDistanceX(i) + PlayerDistanceY(i) < LowestDistance Then
                    LowestDistance = PlayerDistanceX(i) + PlayerDistanceY(i)
                    PlayerTarget = i
                End If
            End If
        Next
        
        ' Set the target
        If PlayerTarget > 0 Then
            If Not MyTarget = PlayerTarget Then
                Call PlayerSearch(Player(PlayerTarget).X, Player(PlayerTarget).Y)
            End If
        End If
    End If
    
    Call HandleKeyPresses(KeyAscii)

    ' Check if we need to call a label
    If frmMain.picCurrency.Visible Then
        If KeyAscii = vbKeyReturn = True Then Call lblCurrencyOk_Click
        If KeyAscii = vbKeyEscape = True Then Call lblCurrencyCancel_Click
    End If
    
    If frmMain.picDialogue.Visible Then
        If lblDialogue_Button(1).Visible Then
            If KeyAscii = vbKeyReturn Then Call lblDialogue_Button_Click(1)
        Else
            If KeyAscii = vbKeyReturn Then Call lblDialogue_Button_Click(2)
            If KeyAscii = vbKeyEscape Then Call lblDialogue_Button_Click(3)
        End If
    End If
    
    ' Prevents textbox on error ding soundnly be assigned to
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyEscape Or ControlDown Then KeyAscii = 0
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_KeyPress", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    ' Make sure they can't press keys until they are in the game
    If InGame = False Then Exit Sub

    Select Case KeyCode
        Case vbKeyInsert
            If Player(MyIndex).Access >= STAFF_MODERATOR Then
                If FormVisible("frmAdmin") Then
                    If GetForegroundWindow = frmAdmin.hWnd Then
                        Unload frmAdmin
                    ElseIf GetForegroundWindow <> frmAdmin.hWnd Then
                        BringWindowToTop (frmAdmin.hWnd)
                        'InitAdminPanel
                    End If
                Else
                    InitAdminPanel
                End If
            End If
        
        Case vbKeyUp
            If ChatLocked Then
                If TempPlayer(MyIndex).Moving = NO And Options.WASD = 1 Then
                    Call SetPlayerDir(MyIndex, DIR_UP)
                    Call SendPlayerDir
                    MouseX = -1
                    MouseY = -1
                    Exit Sub
                End If
            End If

        Case vbKeyDown
            If ChatLocked Then
                If TempPlayer(MyIndex).Moving = NO And Options.WASD = 1 Then
                    Call SetPlayerDir(MyIndex, DIR_DOWN)
                    Call SendPlayerDir
                    MouseX = -1
                    MouseY = -1
                    Exit Sub
                End If
            End If

        Case vbKeyLeft
            If ChatLocked Then
                If TempPlayer(MyIndex).Moving = NO And Options.WASD = 1 Then
                    Call SetPlayerDir(MyIndex, DIR_LEFT)
                    Call SendPlayerDir
                    MouseX = -1
                    MouseY = -1
                    Exit Sub
                End If
            End If

        Case vbKeyRight
            If ChatLocked Then
                If TempPlayer(MyIndex).Moving = NO And Options.WASD = 1 Then
                    Call SetPlayerDir(MyIndex, DIR_RIGHT)
                    Call SendPlayerDir
                    MouseX = -1
                    MouseY = -1
                    Exit Sub
                End If
            End If
        
        Case vbKeyEnd
            If ChatLocked Then
                If TempPlayer(MyIndex).Moving = NO Then
                    If GetPlayerDir(MyIndex) = 0 Then
                        Call SetPlayerDir(MyIndex, GetPlayerDir(MyIndex) + 3)
                    ElseIf GetPlayerDir(MyIndex) = 1 Then
                        Call SetPlayerDir(MyIndex, GetPlayerDir(MyIndex) + 1)
                    ElseIf GetPlayerDir(MyIndex) = 2 Then
                        Call SetPlayerDir(MyIndex, GetPlayerDir(MyIndex) - 2)
                    ElseIf GetPlayerDir(MyIndex) = 3 Then
                        Call SetPlayerDir(MyIndex, GetPlayerDir(MyIndex) - 2)
                    End If
                    Call SendPlayerDir
                    MouseX = -1
                    MouseY = -1
                    Exit Sub
                End If
            End If

    End Select
    
    ' Handles delete events
    If KeyCode = vbKeyDelete Then
        If InMapEditor Then DeleteEvent CurX, CurY
    End If
    
    ' Handles copy + pasting events
    If KeyCode = vbKeyC Then
        If ControlDown Then
            If InMapEditor Then
                CopyEvent_Map CurX, CurY
            End If
        End If
    End If
    
    If KeyCode = vbKeyV Then
        If ControlDown Then
            If InMapEditor Then
                PasteEvent_Map CurX, CurY
            End If
        End If
    End If
    
     ' Hotbar
    If Options.WASD = 1 Then
        For i = 1 To MAX_HOTBAR - 3 '
            If KeyCode = 48 + i Then
                SendHotbarUse i
            End If
        Next
        ' Hot bar button 0
        If KeyCode = 48 Then SendHotbarUse 10
        
        ' Hot bar button -
        If KeyCode = 189 Then SendHotbarUse 11
        
        ' Hot bar button +
        If KeyCode = 187 Then SendHotbarUse 12
        Exit Sub
    Else
        For i = 1 To MAX_HOTBAR
            If KeyCode = 111 + i Then
                SendHotbarUse i
            End If
        Next
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "Form_KeyUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtDialogue_Change()
    If DialogueIndex = DIALOGUE_TYPE_CHANGEGUILDACCESS Then
        If Not txtDialogue.text = vbNullString Then
            If Not IsNumeric(txtDialogue.text) Then txtDialogue.text = 1
            If txtDialogue.text < 1 Then txtDialogue.text = 1
            If txtDialogue.text > MAX_GUILDACCESS Then txtDialogue.text = MAX_GUILDACCESS
        End If
    End If
End Sub

Private Sub txtMyChat_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    MyText = txtMyChat
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtMyChat_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtChat_GotFocus()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    SetGameFocus
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtChat_GotFocus", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub txtChat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picItemDesc.Visible = False
    picSpellDesc.Visible = False
    ClearButtons
    ResetOptionButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtChat_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' ***************
' ** Inventory **
' ***************
Private Sub lblUseItem_Click()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    Call UseItem
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "lblUseItem_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picInventory_DblClick()
    Dim InvNum As Long
    Dim Value As Long
    Dim Multiplier As Double
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    InvNum = IsInvItem(InvX, InvY)

    If InvNum <> 0 Then
        ' Are we in a shop
        If InShop > 0 Then
            If Not TryingToFixItem Then
                SellItem InvNum
            Else
                FixItem InvNum
                TryingToFixItem = False
            End If
            Exit Sub
        End If
        
        ' In Bank
        If InBank Then
            If item(GetPlayerInvItemNum(MyIndex, InvNum)).stackable = 1 Then
                If GetPlayerInvItemValue(MyIndex, InvNum) > 1 Then
                    CurrencyMenu = 2 ' Deposit
                    lblCurrency.Caption = "How many do you want to deposit?"
                    TmpCurrencyItem = InvNum
                    txtCurrency.text = vbNullString
                    picCurrency.Visible = True
                    picCurrency.ZOrder (0)
                    txtCurrency.SetFocus
                Else
                    Call DepositItem(InvNum, 1)
                End If
            Else
                Call DepositItem(InvNum, 0)
            End If
            Exit Sub
        End If
        
        ' In trade
        If InTrade > 0 Then
            ' Exit out if we're offering that item
            For i = 1 To MAX_INV
                If TradeYourOffer(i).num = InvNum Then
                    ' Is currency?
                    If item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).stackable = 1 Then
                        ' Only exit out if we're offering all of it
                        If TradeYourOffer(i).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then Exit Sub
                    Else
                        Exit Sub
                    End If
                End If
            Next
            
            If item(GetPlayerInvItemNum(MyIndex, InvNum)).stackable = 1 Then
                If GetPlayerInvItemValue(MyIndex, InvNum) > 1 Then
                    CurrencyMenu = 4 ' Offer in trade
                    lblCurrency.Caption = "How many do you want to trade?"
                    TmpCurrencyItem = InvNum
                    txtCurrency.text = vbNullString
                    picCurrency.Visible = True
                    picCurrency.ZOrder (0)
                    txtCurrency.SetFocus
                Else
                    Call TradeItem(InvNum, 1)
                End If
            Else
                Call TradeItem(InvNum, 0)
            End If
            Exit Sub
        End If
        
        ' Don't use an item if it is None or Auto Life
        If item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_NONE Or item(GetPlayerInvItemNum(MyIndex, InvNum)).stackable = 1 Or item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_AUTOLIFE Then
            AddText "You can't use this type of item!", BrightRed
            Exit Sub
        End If
        
        ' Reset Stat Points
        If item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_RESETSTATS Then
            Dialogue "Reset Stats", "Are you sure you wish to reset your stats?", DIALOGUE_TYPE_RESETSTATS, True, InvNum
            Exit Sub
        End If
        
        ' Use item if not doing anything else
        Call SendUseItem(InvNum)
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picInventory_DblClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
    Exit Sub
End Sub

Private Function IsEqItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To Equipment.Equipment_Count - 1
        If GetPlayerEquipment(MyIndex, i) > 0 And GetPlayerEquipment(MyIndex, i) <= MAX_ITEMS Then
            With TempRec
                .Top = EquipSlotTop(i)
                .Bottom = .Top + PIC_Y
                .Left = EquipSlotLeft(i)
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsEqItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
errorhandler:
    HandleError "IsEqItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Function IsInvItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_INV
        If GetPlayerInvItemNum(MyIndex, i) > 0 And GetPlayerInvItemNum(MyIndex, i) <= MAX_ITEMS Then
            With TempRec
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + PIC_X) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsInvItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
errorhandler:
    HandleError "IsInvItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Function IsPlayerSpell(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    For i = 1 To MAX_PLAYER_SPELLS
        If PlayerSpells(i) > 0 And PlayerSpells(i) <= MAX_PLAYER_SPELLS Then
            With TempRec
                .Top = SpellTop + ((SpellOffsetY + 32) * ((i - 1) \ SpellColumns))
                .Bottom = .Top + PIC_Y
                .Left = SpellLeft + ((SpellOffsetX + PIC_X) * (((i - 1) Mod SpellColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsPlayerSpell = i
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
errorhandler:
    HandleError "IsPlayerSpell", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Function IsTradeItem(ByVal X As Single, ByVal Y As Single, ByVal Yours As Boolean) As Long
    Dim TempRec As RECT
    Dim i As Long
    Dim ItemNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    For i = 1 To MAX_INV
        If Yours Then
            ItemNum = GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)
        Else
            ItemNum = TradeTheirOffer(i).num
        End If

        If ItemNum > 0 And ItemNum <= MAX_ITEMS Then
            With TempRec
                .Top = InvTop - 12 + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + PIC_X) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsTradeItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
errorhandler:
    HandleError "IsTradeItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Sub picInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Byte
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InTrade > 0 Then Exit Sub
    
    InvNum = IsInvItem(X, Y)
    
    If Button = 1 Then
        If InvNum > 0 And InvNum <= MAX_INV Then
            DragInvSlot = InvNum
            Exit Sub
        End If
    ElseIf Button = 2 Then
        If InvNum > 0 And InvNum <= MAX_INV Then
            X = X + picInventory.Left
            Y = Y + picInventory.Top
            lstDropDownBox.Top = Y
            lstDropDownBox.Left = X
            
            ' If the original height is stored then set the height and clear it
            If TmplstDropDownBoxHeight > 0 Then
                lstDropDownBox.Height = TmplstDropDownBoxHeight
                TmplstDropDownBoxHeight = 0
            End If
            
            ' Set the height
            TmplstDropDownBoxHeight = lstDropDownBox.Height
            
            ' Clear the list
            lstDropDownBox.Clear
            
            ' Build the list
            If Not item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_NONE And Not item(GetPlayerInvItemNum(MyIndex, InvNum)).stackable = 1 And Not item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_AUTOLIFE Then
                lstDropDownBox.AddItem "Use"
            ElseIf item(GetPlayerInvItemNum(MyIndex, InvNum)).Type = ITEM_TYPE_EQUIPMENT Then
                lstDropDownBox.AddItem "Equip"
            End If
            
            lstDropDownBox.AddItem "Drop"
            'lstDropDownBox.AddItem "Examine"
            
            ' Set the new height
            lstDropDownBox.Height = lstDropDownBox.Height * lstDropDownBox.ListCount
            
            ' Other stuff
            lstDropDownBox.ListIndex = -1
            lstDropDownBox.ZOrder (0)
            lstDropDownBox.Visible = True
            
            ' Store the inventory number for future use
            TmpInvNum = InvNum
            TmpSpellSlot = 0
        End If
    End If

    SetGameFocus
    picItemDesc.Visible = False
    LastItemDesc = 0
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picInventory_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picInventory_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim InvNum As Byte
    Dim i As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    InvX = X
    InvY = Y

    If DragInvSlot > 0 Then
        If InTrade > 0 Then Exit Sub
        Call DrawDraggedItem(X + picInventory.Left - 16, Y + picInventory.Top - 16)
    Else
        InvNum = IsInvItem(X, Y)

        If Not InvNum = 0 Then
            ' Exit out if we're offering that item
            If InTrade > 0 Then
                For i = 1 To MAX_INV
                    If TradeYourOffer(i).num = InvNum Then
                        ' is currency?
                        If item(GetPlayerInvItemNum(MyIndex, TradeYourOffer(i).num)).stackable = 1 Then
                            ' Only exit out if we're offering all of it
                            If TradeYourOffer(i).Value = GetPlayerInvItemValue(MyIndex, TradeYourOffer(i).num) Then Exit Sub
                        Else
                            Exit Sub
                        End If
                    End If
                Next
            End If
            
            X = picInventory.Left - picItemDesc.Width - 4
            Y = picInventory.Top
            UpdateItemDescWindow GetPlayerInvItemNum(MyIndex, InvNum), X, Y
            LastItemDesc = GetPlayerInvItemNum(MyIndex, InvNum) ' Set it so you don't re-set values
            Exit Sub
        End If
    End If

    picItemDesc.Visible = False
    LastItemDesc = 0 ' No item was last loaded
    lstDropDownBox.Visible = False
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picInventory_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picInventory_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If InTrade > 0 Then Exit Sub
    
    If DragInvSlot > 0 Then
        ' Drag and Drop
        For i = 1 To MAX_INV
            With rec_pos
                .Top = InvTop + ((InvOffsetY + 32) * ((i - 1) \ InvColumns))
                .Bottom = .Top + PIC_Y
                .Left = InvLeft + ((InvOffsetX + PIC_X) * (((i - 1) Mod InvColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then '
                    If Not DragInvSlot = i Then
                        SendChangeInvSlots DragInvSlot, i
                        Exit For
                    End If
                End If
            End If
        Next
        
        ' Hotbar
        For i = 1 To MAX_HOTBAR
            With rec_pos
                .Top = picHotbar.Top - picInventory.Top
                .Left = picHotbar.Left - picInventory.Left + (HotbarOffsetX * (i - 1)) + (32 * (i - 1))
                .Right = .Left + 32
                .Bottom = picHotbar.Top - picInventory.Top + 32
            End With
            
            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    SendHotbarChange 1, DragInvSlot, i
                    DragInvSlot = 0
                    picTempInv.Visible = False
                    Exit Sub
                End If
            End If
        Next
    End If

    DragInvSlot = 0
    picTempInv.Visible = False
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picInventory_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picItemDesc_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    picItemDesc.Visible = False
    LastItemDesc = 0 ' No item was last loaded
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picItemDesc_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' *****************
' ** Char Window **
' *****************
Private Sub picEquipment_Click()
    Dim EqNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EqNum = IsEqItem(EqX, EqY)

    If Not EqNum = 0 Then
        SendUnequip EqNum
    End If
    
    picItemDesc.Visible = False
    LastItemDesc = 0
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picEquipment_Click", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picEquipment_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim EqNum As Long
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    EqX = X
    EqY = Y
    EqNum = IsEqItem(X, Y)

    If Not EqNum = 0 Then
        X = X + picEquipment.Left - picItemDesc.Width - 1
        Y = Y + picEquipment.Top - picItemDesc.Height
        UpdateItemDescWindow GetPlayerEquipment(MyIndex, EqNum), X, Y
        LastItemDesc = GetPlayerEquipment(MyIndex, EqNum) ' Set it so you don't re-set values
        Exit Sub
    End If
    
    Call ClearChatButton(0)
    picItemDesc.Visible = False
    LastItemDesc = 0 ' No item was last loaded
    ClearButtons
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picEquipment_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

' Bank
Private Sub picBank_DblClick()
    Dim BankNum As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    DragBankSlot = 0

    BankNum = IsBankItem(BankX, BankY)
    
    If Not BankNum = 0 Then
        If item(GetBankItemNum(BankNum)).stackable = 1 Then
            If GetBankItemValue(BankNum) > 1 Then
                CurrencyMenu = 3 ' Withdraw
                lblCurrency.Caption = "How many do you want to withdraw?"
                TmpCurrencyItem = BankNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                picCurrency.ZOrder (0)
                txtCurrency.SetFocus
                Exit Sub
            Else
                WithdrawItem BankNum, 1
                Exit Sub
            End If
        Else
            WithdrawItem BankNum, 1
            Exit Sub
        End If
        WithdrawItem BankNum, 0
        Exit Sub
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picBank_DlbClick", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picBank_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim BankNum As Long
                        
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    BankNum = IsBankItem(X, Y)
    
    If Not BankNum = 0 Then
        If Button = 1 Then
            DragBankSlot = BankNum
        End If
    End If
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picBank_MouseDown", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picBank_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    Dim rec_pos As RECT
    
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    If DragBankSlot > 0 Then
        For i = 1 To MAX_BANK
            With rec_pos
                .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + PIC_X) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With

            If X >= rec_pos.Left And X <= rec_pos.Right Then
                If Y >= rec_pos.Top And Y <= rec_pos.Bottom Then
                    If DragBankSlot <> i Then
                        SwapBankSlots DragBankSlot, i
                        Exit For
                    End If
                End If
            End If
        Next
    End If

    DragBankSlot = 0
    picTempBank.Visible = False
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picBank_MouseUp", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub picBank_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim BankNum As Long, ItemNum As Long, ItemType As Long
    Dim X2 As Long, Y2 As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    BankX = X
    BankY = Y
    
    If DragBankSlot > 0 Then
        Call DrawBankItem(X + picBank.Left, Y + picBank.Top)
    Else
        BankNum = IsBankItem(X, Y)
        
        If BankNum <> 0 Then
            X2 = X + picBank.Left + 1
            Y2 = Y + picBank.Top + 1
            UpdateItemDescWindow bank.item(BankNum).num, X2, Y2
            LastItemDesc = bank.item(BankNum).num
            Exit Sub
        End If
    End If
    
    frmMain.picItemDesc.Visible = False
    LastBankDesc = 0
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "picBank_MouseMove", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Function IsBankItem(ByVal X As Single, ByVal Y As Single) As Long
    Dim TempRec As RECT
    Dim i As Long

    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler

    IsBankItem = 0
    
    For i = 1 To MAX_BANK
        If GetBankItemNum(i) > 0 And GetBankItemNum(i) <= MAX_ITEMS Then
            With TempRec
                .Top = BankTop + ((BankOffsetY + 32) * ((i - 1) \ BankColumns))
                .Bottom = .Top + PIC_Y
                .Left = BankLeft + ((BankOffsetX + PIC_X) * (((i - 1) Mod BankColumns)))
                .Right = .Left + PIC_X
            End With
            
            If X >= TempRec.Left And X <= TempRec.Right Then
                If Y >= TempRec.Top And Y <= TempRec.Bottom Then
                    IsBankItem = i
                    Exit Function
                End If
            End If
        End If
    Next
    Exit Function
    
' Error handler
errorhandler:
    HandleError "IsBankItem", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Function

Private Sub txtTransChat_Change()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    Call SetGameFocus
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "txtTransChat_Change", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Public Sub CloseAllPanels()
    ' If debug mode, handle error then exit out
    If Options.Debug = 1 Then On Error GoTo errorhandler
    
    picInventory.Visible = False
    picSpells.Visible = False
    picCharacter.Visible = False
    picOptions.Visible = False
    picGuild.Visible = False
    picGuild_No.Visible = False
    picFriends.Visible = False
    picParty.Visible = False
    picEquipment.Visible = False
    picFoes.Visible = False
    picTitles.Visible = False
    Exit Sub
    
' Error handler
errorhandler:
    HandleError "CloseAllPanels", "frmMain", Err.Number, Err.Description, Err.Source, Err.HelpContext
    Err.Clear
End Sub

Private Sub DropItem(ByVal InvNum As Byte)
    If InvNum > 0 And InvNum <= MAX_INV Then
        If item(GetPlayerInvItemNum(MyIndex, InvNum)).stackable = 1 Then
            If GetPlayerInvItemValue(MyIndex, InvNum) > 0 Then
                CurrencyMenu = 1 ' drop
                lblCurrency.Caption = "How many do you want to drop?"
                TmpCurrencyItem = InvNum
                txtCurrency.text = vbNullString
                picCurrency.Visible = True
                picCurrency.ZOrder (0)
                txtCurrency.SetFocus
                Exit Sub
            End If
        Else
            Call SendDropItem(InvNum, 0)
        End If
    End If
End Sub
Private Sub myWndProc(ByVal bBefore As Boolean, _
                      ByRef bHandled As Boolean, _
                      ByRef lReturn As Long, _
                      ByVal lng_hWnd As Long, _
                      ByVal uMsg As Long, _
                      ByVal wParam As Long, _
                      ByVal lParam As Long, _
                      ByRef lParamUser As Long)
    Select Case uMsg
        Case WM_ACTIVATEAPP
            taskBarClick = True
        Case WM_LBUTTONDOWN
            MainLButtonDown lng_hWnd
        Case WM_LBUTTONUP
            MainLButtonUp lng_hWnd
        Case WM_CAPTURECHANGED
            MainCaptureChanged lng_hWnd, lParam
        Case WM_MOUSEMOVE
            MainMouseMove lng_hWnd
        Case WM_GETMINMAXINFO 'Prevent Resizing, so we can keep nice frame when turning off CAPTION.
            If Not taskBarClick Then
                MainPreventResizing Me.hWnd, (Me.Width \ Screen.TwipsPerPixelX), (Me.Height \ Screen.TwipsPerPixelY), lParam
            Else
                taskBarClick = False
            End If
        Case WM_SETFOCUS
            If lng_hWnd = mapPreviewSwitch.hWnd Then
                bHandled = True
                lReturn = 1
            End If
    End Select

' *************************************************************
' C A U T I O N   C A U T I O N   C A U T I O N   C A U T I O N
' -------------------------------------------------------------
' DO NOT ADD ANY OTHER CODE BELOW THE "END SUB" STATEMENT BELOW
' *************************************************************
End Sub
