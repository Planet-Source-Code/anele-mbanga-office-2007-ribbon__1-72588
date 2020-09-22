VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   6075
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12360
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   12360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command10 
      Caption         =   "Get Date"
      Height          =   375
      Left            =   9600
      TabIndex        =   12
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Set Date"
      Height          =   375
      Left            =   7680
      TabIndex        =   11
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Timer AnimateLogo 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   4560
      Top             =   4560
   End
   Begin VB.CommandButton Command8 
      Caption         =   "Animate Logo"
      Height          =   375
      Left            =   7680
      TabIndex        =   10
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Update Fonts"
      Height          =   375
      Left            =   5760
      TabIndex        =   9
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Set Windows"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Read Windows"
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Read Names"
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   5040
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Rename Saver"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton cmdRenameTab 
      Caption         =   "Rename Tab 2"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   5520
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Show Progress"
      Height          =   375
      Left            =   1920
      TabIndex        =   3
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Timer Downloading 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   3840
      Top             =   4560
   End
   Begin VB.TextBox Text1 
      Height          =   2175
      Left            =   0
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2160
      Width           =   7815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load Ribbon"
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   5040
      Width           =   1815
   End
   Begin TheRibbon.ACPRibbon ACPRibbon1 
      Height          =   2130
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12360
      _extentx        =   21802
      _extenty        =   3757
      imagesize       =   0
      font            =   "Form1.frx":0000
      usepermissions  =   0   'False
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   230
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0028
            Key             =   "save"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":05C2
            Key             =   "commentw"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0AC1
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":105B
            Key             =   "find"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15F5
            Key             =   "opened"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B8F
            Key             =   "report"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2129
            Key             =   "npo"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":26C3
            Key             =   "empty"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2C5D
            Key             =   "full"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":31F7
            Key             =   "restore"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3549
            Key             =   "isazi"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3AE3
            Key             =   "inbox"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":407D
            Key             =   "experts"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4617
            Key             =   "runsql2"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4A69
            Key             =   "survey"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5003
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":515D
            Key             =   "xx"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5D2F
            Key             =   "clock"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":6181
            Key             =   "excel"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BDA3
            Key             =   "refresh"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BEFD
            Key             =   "table"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C497
            Key             =   "ie"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CA31
            Key             =   "sum"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CB93
            Key             =   "key1"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CFE5
            Key             =   "module"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D13F
            Key             =   "stats"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D459
            Key             =   "new"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D803
            Key             =   "print"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DC01
            Key             =   "taskt"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13B9B
            Key             =   "attacht"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13FED
            Key             =   "verify"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14307
            Key             =   "defer"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1471E
            Key             =   "discuss"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14B34
            Key             =   "maybe"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":14F4B
            Key             =   "move"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15366
            Key             =   "risk"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15779
            Key             =   "yes"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15B8F
            Key             =   "high"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":15F5F
            Key             =   "normal"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1634D
            Key             =   "low"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1673C
            Key             =   "furious"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16B60
            Key             =   "happy"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":16F92
            Key             =   "neutral"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":173C3
            Key             =   "upsat"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":177EC
            Key             =   "sad"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17C18
            Key             =   "task25"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":17FEE
            Key             =   "task50"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":183A2
            Key             =   "task75"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18736
            Key             =   "task100"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18B33
            Key             =   "task0"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":18F23
            Key             =   "email"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":194BD
            Key             =   "hight"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19895
            Key             =   "lowt"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":19C8C
            Key             =   "normalt"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A082
            Key             =   "furioust"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A4AE
            Key             =   "happyt"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1A8E8
            Key             =   "neutralt"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1AD21
            Key             =   "upsatt"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B152
            Key             =   "sadt"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B583
            Key             =   "defert"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1B9A2
            Key             =   "discusst"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1BDC0
            Key             =   "maybet"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C1DF
            Key             =   "movet"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1C602
            Key             =   "riskt"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1CA1D
            Key             =   "yest"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1CE3B
            Key             =   "task25t"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D219
            Key             =   "task50t"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D5D5
            Key             =   "task75t"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1D971
            Key             =   "task100t"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1DD76
            Key             =   "task0t"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E16E
            Key             =   "green"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1E5C0
            Key             =   "red"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EA12
            Key             =   "organization"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1EE64
            Key             =   "region"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":1F2B6
            Key             =   "department"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":222D8
            Key             =   "owner"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2312A
            Key             =   "resources"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":293C4
            Key             =   "target1"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2995E
            Key             =   "date"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":29D81
            Key             =   "perspective"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A1D3
            Key             =   "duedate"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2A76D
            Key             =   "complete"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":2AD07
            Key             =   "expected"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":30929
            Key             =   "taborder"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":30EC3
            Key             =   "link"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3109D
            Key             =   "column"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":31B67
            Key             =   "runsql"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":32090
            Key             =   "taskx"
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":32420
            Key             =   "attach"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":327FE
            Key             =   "info"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":32C50
            Key             =   "develop"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":41A9B
            Key             =   "mindmanager"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":437A5
            Key             =   "suite"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4502F
            Key             =   "star"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":462B1
            Key             =   "sync"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4C507
            Key             =   "offline"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4D291
            Key             =   "highr"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4E01B
            Key             =   "lowr"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4EDA5
            Key             =   "mediumr"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":4FB2F
            Key             =   "wss"
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":51EB1
            Key             =   "wssdoc"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":529FB
            Key             =   "toolicon"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":551AD
            Key             =   "useraccount"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":55546
            Key             =   "calender"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":55860
            Key             =   "chart"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":55AC7
            Key             =   "customer"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":55E2E
            Key             =   "list"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":56093
            Key             =   "newsomething"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":563A0
            Key             =   "iconopen"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":566DC
            Key             =   "profile"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":56A0A
            Key             =   "project"
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":56D6C
            Key             =   "resources1"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":570A1
            Key             =   "reports"
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":575FE
            Key             =   "info1"
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":57D50
            Key             =   "warn"
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5806A
            Key             =   "traffic"
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":584BC
            Key             =   "target"
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":5890E
            Key             =   "doclibrary"
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":59458
            Key             =   "live1"
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":63D9A
            Key             =   "stop"
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":641EC
            Key             =   "calc"
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":69E0E
            Key             =   "exportproject"
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":73A5C
            Key             =   "importmpp"
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":817FD
            Key             =   "x"
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8190F
            Key             =   "calendar2"
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":878A9
            Key             =   "decrement"
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":87CFB
            Key             =   "increment"
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8814D
            Key             =   "collaborate"
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":8DD6F
            Key             =   "review2"
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":95271
            Key             =   "progress"
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":95B14
            Key             =   "yellowr"
         EndProperty
         BeginProperty ListImage132 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":95EFC
            Key             =   "greenr"
         EndProperty
         BeginProperty ListImage133 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":962EE
            Key             =   "projectplan"
         EndProperty
         BeginProperty ListImage134 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":96D40
            Key             =   "redr"
         EndProperty
         BeginProperty ListImage135 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":971EF
            Key             =   "people"
         EndProperty
         BeginProperty ListImage136 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":98DC6
            Key             =   "bundle"
         EndProperty
         BeginProperty ListImage137 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":99191
            Key             =   "running"
         EndProperty
         BeginProperty ListImage138 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9E983
            Key             =   "stopped"
         EndProperty
         BeginProperty ListImage139 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9F395
            Key             =   "right"
         EndProperty
         BeginProperty ListImage140 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9F92F
            Key             =   "left"
         EndProperty
         BeginProperty ListImage141 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":9FEC9
            Key             =   "deletex"
         EndProperty
         BeginProperty ListImage142 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A0993
            Key             =   "editx"
         EndProperty
         BeginProperty ListImage143 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A145D
            Key             =   "check"
         EndProperty
         BeginProperty ListImage144 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A1777
            Key             =   "group1"
         EndProperty
         BeginProperty ListImage145 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A2349
            Key             =   "none"
         EndProperty
         BeginProperty ListImage146 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A2760
            Key             =   "bluer"
         EndProperty
         BeginProperty ListImage147 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A33BF
            Key             =   "purpler"
         EndProperty
         BeginProperty ListImage148 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A4120
            Key             =   "task"
         EndProperty
         BeginProperty ListImage149 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A4DDC
            Key             =   "note"
         EndProperty
         BeginProperty ListImage150 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A51BC
            Key             =   "money"
         EndProperty
         BeginProperty ListImage151 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A5624
            Key             =   "warn1"
         EndProperty
         BeginProperty ListImage152 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A630B
            Key             =   "question"
         EndProperty
         BeginProperty ListImage153 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A6C32
            Key             =   "change2"
         EndProperty
         BeginProperty ListImage154 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A7148
            Key             =   "excel2"
         EndProperty
         BeginProperty ListImage155 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A86D0
            Key             =   "chart1"
         EndProperty
         BeginProperty ListImage156 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":A9E3A
            Key             =   "pdf1"
         EndProperty
         BeginProperty ListImage157 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AA5DB
            Key             =   "robot1"
         EndProperty
         BeginProperty ListImage158 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AAAD8
            Key             =   "wssw1"
         EndProperty
         BeginProperty ListImage159 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AB498
            Key             =   "resource"
         EndProperty
         BeginProperty ListImage160 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ABCED
            Key             =   "day"
         EndProperty
         BeginProperty ListImage161 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AC2A6
            Key             =   "wssw"
         EndProperty
         BeginProperty ListImage162 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ACF7B
            Key             =   "group"
         EndProperty
         BeginProperty ListImage163 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ADD0E
            Key             =   "robot"
         EndProperty
         BeginProperty ListImage164 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AEF2D
            Key             =   "calendar1"
         EndProperty
         BeginProperty ListImage165 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AFCFD
            Key             =   "actionw"
         EndProperty
         BeginProperty ListImage166 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B016B
            Key             =   "action1"
         EndProperty
         BeginProperty ListImage167 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B05DC
            Key             =   "action"
         EndProperty
         BeginProperty ListImage168 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B0A36
            Key             =   "powerpoint"
         EndProperty
         BeginProperty ListImage169 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B1D43
            Key             =   "pie"
         EndProperty
         BeginProperty ListImage170 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B6287
            Key             =   "shake"
         EndProperty
         BeginProperty ListImage171 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B6A37
            Key             =   "newx"
         EndProperty
         BeginProperty ListImage172 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B74A8
            Key             =   "refreshmeeting"
         EndProperty
         BeginProperty ListImage173 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B818A
            Key             =   "discuss1"
         EndProperty
         BeginProperty ListImage174 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B9291
            Key             =   "write"
         EndProperty
         BeginProperty ListImage175 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B9856
            Key             =   "action2"
         EndProperty
         BeginProperty ListImage176 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":B9CA8
            Key             =   "company"
         EndProperty
         BeginProperty ListImage177 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BA1E4
            Key             =   "redfolder"
         EndProperty
         BeginProperty ListImage178 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BB0F4
            Key             =   "greenfolder"
         EndProperty
         BeginProperty ListImage179 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BBF8F
            Key             =   "construct"
         EndProperty
         BeginProperty ListImage180 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BC6BD
            Key             =   "camera"
         EndProperty
         BeginProperty ListImage181 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BD127
            Key             =   "expand"
         EndProperty
         BeginProperty ListImage182 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BDF7D
            Key             =   "live"
         EndProperty
         BeginProperty ListImage183 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BE489
            Key             =   "change"
         EndProperty
         BeginProperty ListImage184 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BE9D7
            Key             =   "documents"
         EndProperty
         BeginProperty ListImage185 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BEDE0
            Key             =   "docs"
         EndProperty
         BeginProperty ListImage186 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BF1F1
            Key             =   "docsfolder"
         EndProperty
         BeginProperty ListImage187 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BF5EE
            Key             =   "sitevisit"
         EndProperty
         BeginProperty ListImage188 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":BFD2A
            Key             =   "photo"
         EndProperty
         BeginProperty ListImage189 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C079F
            Key             =   "tracking"
         EndProperty
         BeginProperty ListImage190 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C0C74
            Key             =   "report1"
         EndProperty
         BeginProperty ListImage191 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C115A
            Key             =   "recommendationw"
         EndProperty
         BeginProperty ListImage192 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C166F
            Key             =   "recommendationt"
         EndProperty
         BeginProperty ListImage193 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C1B8D
            Key             =   "commentt"
         EndProperty
         BeginProperty ListImage194 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C209B
            Key             =   "wizard1"
         EndProperty
         BeginProperty ListImage195 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C3E31
            Key             =   "milestone"
         EndProperty
         BeginProperty ListImage196 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C41EA
            Key             =   "view"
         EndProperty
         BeginProperty ListImage197 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C45D7
            Key             =   "wizard"
         EndProperty
         BeginProperty ListImage198 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C4C8C
            Key             =   "saveall"
         EndProperty
         BeginProperty ListImage199 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C51F3
            Key             =   "checkmark"
         EndProperty
         BeginProperty ListImage200 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C579B
            Key             =   "xmark"
         EndProperty
         BeginProperty ListImage201 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C5E16
            Key             =   "calendar"
         EndProperty
         BeginProperty ListImage202 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":C6792
            Key             =   "project.show"
         EndProperty
         BeginProperty ListImage203 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CCFF4
            Key             =   "executivet"
         EndProperty
         BeginProperty ListImage204 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CE407
            Key             =   "executivew"
         EndProperty
         BeginProperty ListImage205 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CF875
            Key             =   "key"
         EndProperty
         BeginProperty ListImage206 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":CFC89
            Key             =   "keyt"
         EndProperty
         BeginProperty ListImage207 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D009C
            Key             =   "reviewt"
         EndProperty
         BeginProperty ListImage208 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D0636
            Key             =   "revieww"
         EndProperty
         BeginProperty ListImage209 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D09B1
            Key             =   "increaseform"
         EndProperty
         BeginProperty ListImage210 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D102C
            Key             =   "notenlarge"
         EndProperty
         BeginProperty ListImage211 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D1473
            Key             =   "ts"
         EndProperty
         BeginProperty ListImage212 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D2C49
            Key             =   "blue"
         EndProperty
         BeginProperty ListImage213 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D39A0
            Key             =   "brown"
         EndProperty
         BeginProperty ListImage214 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D4882
            Key             =   "bluet"
         EndProperty
         BeginProperty ListImage215 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D55E7
            Key             =   "ambert"
         EndProperty
         BeginProperty ListImage216 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D59F9
            Key             =   "greent"
         EndProperty
         BeginProperty ListImage217 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D5E15
            Key             =   "redt"
         EndProperty
         BeginProperty ListImage218 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D6223
            Key             =   "offlinef"
         EndProperty
         BeginProperty ListImage219 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D6630
            Key             =   "onlinef"
         EndProperty
         BeginProperty ListImage220 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D6A57
            Key             =   "closedw"
         EndProperty
         BeginProperty ListImage221 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D6E0C
            Key             =   "openedw"
         EndProperty
         BeginProperty ListImage222 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D71BF
            Key             =   "synchronize"
         EndProperty
         BeginProperty ListImage223 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D85F6
            Key             =   "moon1"
         EndProperty
         BeginProperty ListImage224 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D8A48
            Key             =   "moon2"
         EndProperty
         BeginProperty ListImage225 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D8E9A
            Key             =   "moon3"
         EndProperty
         BeginProperty ListImage226 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D92EC
            Key             =   "moon4"
         EndProperty
         BeginProperty ListImage227 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D973E
            Key             =   "moon5"
         EndProperty
         BeginProperty ListImage228 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D9B90
            Key             =   "moon6"
         EndProperty
         BeginProperty ListImage229 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":D9FE2
            Key             =   "moon7"
         EndProperty
         BeginProperty ListImage230 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":DA434
            Key             =   "moon8"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private lngCounter As Long
Private animation As Integer
Private Sub ACPRibbon1_ButtonClick(ByVal Id As String, ByVal Caption As String)
    On Error Resume Next
    Text1.Text = Id & ", " & Caption
    Select Case Id
    Case "mnuExit"
        MsgBox "Are you sure that you want to exit this application?", vbYesNo + vbQuestion + vbApplicationModal, "Confirm Exit"
    Case "3"
    Case "godisgood"
        MsgBox "God is good all the time"
    End Select
    Err.Clear
End Sub
Private Sub ACPRibbon1_ComboClick(ByVal ComboName As String, ByVal Text As String)
    On Error Resume Next
    Text1.Text = ComboName & ", " & Text
    Err.Clear
End Sub
Private Sub ACPRibbon1_DatePickClick(ByVal DatePickName As String, ByVal DatePicked As String)
    On Error Resume Next
    Text1.Text = DatePickName & ", " & DatePicked
    Err.Clear
End Sub
Private Sub ACPRibbon1_MenuClick(ByVal Id As String, ByVal Caption As String)
    On Error Resume Next
    Select Case Id
    Case "offline"
        ACPRibbon1.EditTopButton "offline", "online", "Work online", "onlinef", "Work online"
        MsgBox "You are now working online."
    Case "online"
        ACPRibbon1.EditTopButton "online", "offline", "Work offline", "offlinef", "Work offline"
        MsgBox "You are now working offline."
    End Select
    Err.Clear
End Sub
Private Sub AnimateLogo_Timer()
    On Error Resume Next
    animation = animation + 1
    If animation > 8 Then animation = 1
    ACPRibbon1.Icon = "moon" & CStr(animation)
    Err.Clear
End Sub
Private Sub cmdRenameTab_Click()
    On Error Resume Next
    ACPRibbon1.RenameTab "2", "Anele Mbanga"
    Err.Clear
End Sub
Public Sub Command1_Click()
    On Error Resume Next
    animation = 0
    Command8.Caption = "Animate Logo"
    AnimateLogo.Enabled = False
    ACPRibbon1.UsePermissions = False
    ACPRibbon1.Clear
    ACPRibbon1.ImageList = imgIcons
    ACPRibbon1.ImageSize = Size320
    ACPRibbon1.Top = 0
    ACPRibbon1.Left = 0
    ACPRibbon1.Icon = "ie"
    ACPRibbon1.ResizeLogo 480
    ACPRibbon1.AddTopButton "print", "Printer", "print", "Print"
    ACPRibbon1.AddTopButton "offline", "Offline", "offlinef", "Work Offline"
    ACPRibbon1.AddTab "1", "Example", True
    ACPRibbon1.AddCat "1", "1", "Cat 1", False, ""
    ACPRibbon1.AddButton "1", "1", "Search", "find", True, "", False
    ACPRibbon1.AddComboBox "5", "1", "Names", "", "cboNames", 2000
    ACPRibbon1.AddComboBoxItem "cboNames", "Anele Mbanga"
    ACPRibbon1.AddComboBoxItem "cboNames", "Sikelela Mbanga"
    ACPRibbon1.AddComboBoxItem "cboNames", "Usibabale Mbanga"
    ACPRibbon1.AddComboBoxItem "cboNames", "Olothando Mbanga"
    ACPRibbon1.AddComboBox "2", "1", "Fonts", "", "cboFonts", 3000
    ACPRibbon1.AddComboBoxItem "cboFonts", "Arial"
    ACPRibbon1.AddComboBoxItem "cboFonts", "Tahoma"
    ACPRibbon1.AddComboBoxItem "cboFonts", "Gothica"
    ACPRibbon1.AddTextBox "3", "1", "Windows", "Testing text boxes", "txtWindows", 1500
    ACPRibbon1.AddDatePicker "4", "1", "My Date", "Select date of birth", "DOB", 1355
    ACPRibbon1.AddProgressBar "5", "1", "Progress", "Show my progress bar", "progNote", 2000, 0, 500
    ACPRibbon1.AddLabel "6", "1", "So Far", False, "How far we are at", False, "0"
    ACPRibbon1.AddTab "2", "Tab 2", True
    ACPRibbon1.AddCat "2", "2", "Group 1", False, ""
    ACPRibbon1.AddButton "3", "2", "Search", "save", False, "", False
    ACPRibbon1.AddButton "saver", "1", "Saver", "Save", True, "Saving data", False
    ''''''''''''''''''
    ACPRibbon1.AddButtonMenu "1", "mnuFile", "File", True
    ACPRibbon1.AddButtonMenu "1", "mnuFile\mnuOpen", "Open", True
    ACPRibbon1.AddButtonMenu "1", "mnuFile\mnuSave", "Save"
    ACPRibbon1.AddButtonMenu "1", "mnuFile\mnuDelete", "Delete"
    ''''
    ACPRibbon1.AddButtonMenu "1", "mnuSearch", "Search Database", True
    ACPRibbon1.AddButtonMenu "1", "mnuSearch\mnuNames", "Names"
    ACPRibbon1.AddButtonMenu "1", "mnuSearch\mnuDates", "Dates"
    ACPRibbon1.AddButtonMenu "1", "mnuCompress", "Compress Database", True
    ACPRibbon1.AddButtonMenu "1", "mnuSearch\mnuTables", "Tables"
    ACPRibbon1.AddButtonMenu "1", "-", "-"
    ACPRibbon1.AddButtonMenu "1", "mnuExit", "Exit"
    ACPRibbon1.AddButtonMenu "1", "mnuCompress\mnuMakeMDE", "Make MDE"
    ACPRibbon1.AddButtonMenu "1", "mnuCompress\mnuRepair", "Repair"
    'ACPRibbon1.Refresh
    Err.Clear
End Sub
Private Sub Command10_Click()
    On Error Resume Next
    MsgBox ACPRibbon1.DatePickerGetDate("dob")
    Err.Clear
End Sub
Private Sub Command2_Click()
    On Error Resume Next
    ACPRibbon1.ProgressBarReset "progNote", 100
    lngCounter = 0
    Downloading.Enabled = True
    Err.Clear
End Sub
Private Sub Command3_Click()
    On Error Resume Next
    ACPRibbon1.EditButton "saver", "God is Good", "synchronize", False, "God is good all the time.", False, "", "godisgood"
    Err.Clear
End Sub
Private Sub Command4_Click()
    On Error Resume Next
    MsgBox ACPRibbon1.ComboBoxGetText("cboNames")
    Err.Clear
End Sub
Private Sub Command5_Click()
    On Error Resume Next
    MsgBox ACPRibbon1.TextBoxGetText("txtWindows")
    Err.Clear
End Sub
Private Sub Command6_Click()
    On Error Resume Next
    ACPRibbon1.TextBoxSetText "txtWindows", "My Window"
    Err.Clear
End Sub
Private Sub Command7_Click()
    On Error Resume Next
    Dim fontTot As Long
    Dim fontCnt As Long
    Dim myF As Variant
    ACPRibbon1.ComboBoxClear "cboFonts"
    For fontCnt = 0 To Screen.FontCount - 1
        ACPRibbon1.AddComboBoxItem "cboFonts", Screen.Fonts(fontCnt)
    Next
    ACPRibbon1.ComboBoxRefresh
    Err.Clear
End Sub
Private Sub Command8_Click()
    On Error Resume Next
    Select Case Command8.Caption
    Case "Animate Logo"
        AnimateLogo.Enabled = True
        Command8.Caption = "Stop Animation"
    Case Else
        animation = 0
        Command8.Caption = "Animate Logo"
        AnimateLogo.Enabled = False
        ACPRibbon1.Icon = "ie"
    End Select
    Err.Clear
End Sub
Private Sub Command9_Click()
    On Error Resume Next
    ACPRibbon1.DatePickerSetDate "DOB", "01/01/2009"
    Err.Clear
End Sub
Private Sub Downloading_Timer()
    On Error Resume Next
    Downloading.Enabled = False
    lngCounter = lngCounter + 1
    ACPRibbon1.ProgressBarUpdate "progNote", lngCounter
    ACPRibbon1.LabelUpdate "6", CStr(lngCounter)
    Text1.Text = lngCounter
    Downloading.Enabled = True
    If lngCounter > 100 Then
        lngCounter = 0
        ACPRibbon1.LabelUpdate "6", "0"
        Downloading.Enabled = False
    End If
    Err.Clear
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Command1_Click
    Text1.Text = ""
    Set ACPRibbon1.ParentForm = Me
    Err.Clear
End Sub
Private Sub Form_Resize()
    On Error Resume Next
    Set ACPRibbon1.ParentForm = Me
    Err.Clear
End Sub
