VERSION 5.00
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Begin VB.Form InsulinDM 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Insulin Dosing"
   ClientHeight    =   5025
   ClientLeft      =   60
   ClientTop       =   750
   ClientWidth     =   8385
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "InsulinDM.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5025
   ScaleWidth      =   8385
   StartUpPosition =   2  'CenterScreen
   Begin Project1.AutoResize Resize 
      Left            =   390
      Tag             =   "NO"
      Top             =   4545
      _ExtentX        =   714
      _ExtentY        =   714
      KeepAspectRatio =   -1  'True
      AspectRatioValue=   0.660714268684387
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4950
      Left            =   0
      ScaleHeight     =   4950
      ScaleWidth      =   8340
      TabIndex        =   0
      Top             =   0
      Width           =   8340
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000009&
         Caption         =   "a"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   1
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   89
         Top             =   255
         Width           =   270
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000009&
         Caption         =   "`"
         BeginProperty Font 
            Name            =   "Wingdings 3"
            Size            =   9.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Index           =   0
         Left            =   105
         Style           =   1  'Graphical
         TabIndex        =   88
         Top             =   255
         Width           =   270
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   5
         Left            =   7500
         TabIndex        =   12
         Top             =   3465
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   5
         Left            =   5475
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         ToolTipText     =   "Input gms CHO,Fat,Prot, kcal, GI and symptoms"
         Top             =   2940
         Width           =   2700
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   4
         Left            =   4785
         TabIndex        =   10
         Top             =   3465
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   4
         Left            =   2760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   9
         ToolTipText     =   "Input gms CHO,Fat,Prot, kcal, GI and symptoms"
         Top             =   2940
         Width           =   2700
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   3
         Left            =   2085
         TabIndex        =   8
         Top             =   3465
         Width           =   645
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Index           =   3
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         ToolTipText     =   "Input gms CHO,Fat,Prot, kcal, GI and symptoms"
         Top             =   2940
         Width           =   2685
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   2
         Left            =   7500
         TabIndex        =   6
         Top             =   2430
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Index           =   2
         Left            =   5490
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         ToolTipText     =   "Input gms CHO,Fat,Prot, kcal, GI and symptoms"
         Top             =   1125
         Width           =   2685
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   1
         Left            =   4785
         TabIndex        =   4
         Top             =   2430
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Index           =   1
         Left            =   2760
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   3
         ToolTipText     =   "Input gms CHO,Fat,Prot, kcal, GI and symptoms"
         Top             =   1125
         Width           =   2700
      End
      Begin VB.TextBox Text2 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Index           =   0
         Left            =   2070
         TabIndex        =   2
         Top             =   2430
         Width           =   660
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1290
         Index           =   0
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   1
         ToolTipText     =   "Input gms CHO,Fat,Prot, kcal, GI and symptoms"
         Top             =   1125
         Width           =   2685
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "Graphics"
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   6
         Left            =   7515
         TabIndex        =   87
         ToolTipText     =   "Click to set Time and Date"
         Top             =   4680
         Width           =   750
      End
      Begin VB.Image Image4 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Index           =   0
         Left            =   7545
         Picture         =   "InsulinDM.frx":08CA
         Stretch         =   -1  'True
         Top             =   3915
         Width           =   690
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "CHO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   31
         Left            =   4725
         TabIndex        =   84
         Top             =   4695
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "CHO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   30
         Left            =   4725
         TabIndex        =   83
         Top             =   4680
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   29
         Left            =   6195
         TabIndex        =   82
         Top             =   4695
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   28
         Left            =   6210
         TabIndex        =   81
         Top             =   4695
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   27
         Left            =   4080
         TabIndex        =   80
         Top             =   4665
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   26
         Left            =   4095
         TabIndex        =   79
         Top             =   4665
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   25
         Left            =   5565
         TabIndex        =   78
         Top             =   4665
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   24
         Left            =   5550
         TabIndex        =   77
         Top             =   4680
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   23
         Left            =   3375
         TabIndex        =   76
         Top             =   4680
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   22
         Left            =   3390
         TabIndex        =   75
         Top             =   4680
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "CHO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   21
         Left            =   2625
         TabIndex        =   74
         Top             =   4680
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "CHO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   19
         Left            =   2640
         TabIndex        =   73
         Top             =   4680
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   18
         Left            =   1890
         TabIndex        =   72
         Top             =   4665
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   17
         Left            =   1905
         TabIndex        =   71
         Top             =   4665
         Width           =   360
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   16
         Left            =   1185
         TabIndex        =   70
         Top             =   4650
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   15
         Left            =   1200
         TabIndex        =   69
         Top             =   4650
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Snack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   14
         Left            =   6090
         TabIndex        =   68
         Top             =   4245
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Snack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   13
         Left            =   6105
         TabIndex        =   67
         Top             =   4245
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Snack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   12
         Left            =   3900
         TabIndex        =   66
         Top             =   4275
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Snack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   5
         Left            =   3915
         TabIndex        =   65
         Top             =   4260
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Snack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   4
         Left            =   1830
         TabIndex        =   64
         Top             =   4260
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Snack"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   0
         Left            =   1845
         TabIndex        =   63
         Top             =   4260
         Width           =   465
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Lunch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   9
         Left            =   3915
         TabIndex        =   60
         Top             =   3840
         Width           =   525
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Lunch"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   6
         Left            =   3930
         TabIndex        =   59
         Top             =   3855
         Width           =   525
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Snacks"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   3
         Left            =   -105
         TabIndex        =   52
         Top             =   2700
         Width           =   885
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Dinner"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   5985
         TabIndex        =   51
         Top             =   450
         Width           =   1245
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Lunch"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   3360
         TabIndex        =   50
         Top             =   450
         Width           =   1245
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Breakfast"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   675
         TabIndex        =   49
         Top             =   450
         Width           =   1245
      End
      Begin VB.Shape Shape5 
         Height          =   540
         Index           =   2
         Left            =   45
         Top             =   2925
         Width           =   8145
      End
      Begin VB.Shape Shape5 
         Height          =   270
         Index           =   0
         Left            =   45
         Top             =   2415
         Width           =   8145
      End
      Begin VB.Shape Shape4 
         Height          =   375
         Left            =   45
         Top             =   750
         Width           =   8145
      End
      Begin VB.Shape Shape3 
         Height          =   3345
         Left            =   45
         Top             =   390
         Width           =   8145
      End
      Begin VB.Label Br 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   1140
         TabIndex        =   46
         Top             =   4065
         Width           =   675
      End
      Begin VB.Label Br 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   1
         Left            =   1815
         TabIndex        =   45
         Top             =   4065
         Width           =   675
      End
      Begin VB.Label L 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   3270
         TabIndex        =   44
         Top             =   4065
         Width           =   675
      End
      Begin VB.Label L 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   1
         Left            =   3930
         TabIndex        =   43
         Top             =   4050
         Width           =   675
      End
      Begin VB.Label D 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   5430
         TabIndex        =   42
         Top             =   4065
         Width           =   675
      End
      Begin VB.Label D 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   1
         Left            =   6090
         TabIndex        =   41
         Top             =   4065
         Width           =   675
      End
      Begin VB.Label BS 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   1140
         TabIndex        =   40
         Top             =   4425
         Width           =   675
      End
      Begin VB.Label BS 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   1
         Left            =   1815
         TabIndex        =   39
         Top             =   4425
         Width           =   675
      End
      Begin VB.Label LS 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   3270
         TabIndex        =   38
         Top             =   4425
         Width           =   675
      End
      Begin VB.Label LS 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   1
         Left            =   3930
         TabIndex        =   37
         Top             =   4425
         Width           =   675
      End
      Begin VB.Label DS 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   5430
         TabIndex        =   36
         Top             =   4425
         Width           =   675
      End
      Begin VB.Label DS 
         Alignment       =   2  'Center
         BackColor       =   &H80000012&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Index           =   1
         Left            =   6090
         TabIndex        =   35
         Top             =   4425
         Width           =   675
      End
      Begin VB.Label BS 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   2490
         TabIndex        =   34
         Top             =   4425
         Width           =   675
      End
      Begin VB.Label Br 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   2490
         TabIndex        =   33
         Top             =   4065
         Width           =   675
      End
      Begin VB.Label LS 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   4605
         TabIndex        =   32
         Top             =   4425
         Width           =   675
      End
      Begin VB.Label L 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   4605
         TabIndex        =   31
         Top             =   4065
         Width           =   675
      End
      Begin VB.Label DS 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   6765
         TabIndex        =   30
         Top             =   4425
         Width           =   675
      End
      Begin VB.Label D 
         Alignment       =   2  'Center
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   2
         Left            =   6765
         TabIndex        =   29
         Top             =   4065
         Width           =   675
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   3
         Left            =   2190
         Picture         =   "InsulinDM.frx":1194
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   2655
         Width           =   270
      End
      Begin VB.Image Image1 
         Height          =   390
         Index           =   2
         Left            =   7755
         Picture         =   "InsulinDM.frx":1A5E
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   360
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   390
         Index           =   1
         Left            =   5010
         Picture         =   "InsulinDM.frx":2328
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   390
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   390
         Index           =   0
         Left            =   2220
         Picture         =   "InsulinDM.frx":2BF2
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   375
         Width           =   405
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Input Name"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   7215
         TabIndex        =   28
         ToolTipText     =   "Click to set Date"
         Top             =   45
         Width           =   1005
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Blood Glucose after 90 min:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   5
         Left            =   5475
         TabIndex        =   27
         Top             =   3465
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Blood Glucose after 90 min:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   4
         Left            =   2760
         TabIndex        =   26
         Top             =   3465
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Blood Glucose after 90 min:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   3
         Left            =   60
         TabIndex        =   25
         Top             =   3465
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Blood Glucose after 90 min:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   2
         Left            =   5475
         TabIndex        =   24
         Top             =   2430
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   360
         Index           =   2
         Left            =   7245
         Top             =   765
         Width           =   915
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   360
         Index           =   1
         Left            =   4515
         Top             =   765
         Width           =   915
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Blood Glucose after 90 min:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   1
         Left            =   2760
         TabIndex        =   23
         Top             =   2430
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000009&
         Caption         =   "Blood Glucose after 90 min:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   22
         Top             =   2430
         Width           =   1995
      End
      Begin VB.Shape Shape1 
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   360
         Index           =   0
         Left            =   1815
         Top             =   765
         Width           =   915
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Meal Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   5
         Left            =   6315
         TabIndex        =   21
         ToolTipText     =   "Click to set Time and Date"
         Top             =   2715
         Width           =   765
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Meal Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   4
         Left            =   3630
         TabIndex        =   20
         ToolTipText     =   "Click to set Time and Date"
         Top             =   2715
         Width           =   765
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Meal Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   3
         Left            =   1185
         TabIndex        =   19
         ToolTipText     =   "Click to set Time and Date"
         Top             =   2715
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Meal Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   2
         Left            =   6210
         TabIndex        =   18
         ToolTipText     =   "Click to set Time and Date"
         Top             =   840
         Width           =   765
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Meal Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   1
         Left            =   3540
         TabIndex        =   17
         ToolTipText     =   "Click to set Time and Date"
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Set Date"
         ForeColor       =   &H00FF0000&
         Height          =   225
         Left            =   30
         TabIndex        =   16
         ToolTipText     =   "Click to set Date"
         Top             =   45
         Width           =   780
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Meal Time"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   0
         Left            =   900
         TabIndex        =   15
         ToolTipText     =   "Click to set Time and Date"
         Top             =   840
         Width           =   780
      End
      Begin VB.Image Image1 
         Height          =   390
         Index           =   11
         Left            =   2280
         Picture         =   "InsulinDM.frx":34BC
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   390
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   390
         Index           =   10
         Left            =   5070
         Picture         =   "InsulinDM.frx":3D86
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   405
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   390
         Index           =   9
         Left            =   7815
         Picture         =   "InsulinDM.frx":4650
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   375
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   8
         Left            =   2250
         Picture         =   "InsulinDM.frx":4F1A
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   2670
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   5
         Left            =   7890
         Picture         =   "InsulinDM.frx":57E4
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   2670
         Width           =   270
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   4
         Left            =   4875
         Picture         =   "InsulinDM.frx":60AE
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   2670
         Width           =   270
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   7
         Left            =   4935
         Picture         =   "InsulinDM.frx":6978
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   2685
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Image Image1 
         Height          =   255
         Index           =   6
         Left            =   7950
         Picture         =   "InsulinDM.frx":7242
         Stretch         =   -1  'True
         ToolTipText     =   "Insulin Dose, Type and Time"
         Top             =   2685
         Visible         =   0   'False
         Width           =   270
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Glycemic Responses"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   345
         Index           =   0
         Left            =   2460
         TabIndex        =   14
         ToolTipText     =   "Click to set Date"
         Top             =   0
         Width           =   3285
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Glycemic Responses"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Index           =   1
         Left            =   2475
         TabIndex        =   13
         ToolTipText     =   "Click to set Date"
         Top             =   15
         Width           =   3285
      End
      Begin VB.Shape Shape5 
         Height          =   285
         Index           =   1
         Left            =   45
         Top             =   2655
         Width           =   8145
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   5460
         X2              =   5460
         Y1              =   420
         Y2              =   3735
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   2745
         X2              =   2745
         Y1              =   405
         Y2              =   3720
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Breakfast"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   3
         Left            =   1650
         TabIndex        =   47
         Top             =   3840
         Width           =   840
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Breakfast"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   2
         Left            =   1665
         TabIndex        =   48
         Top             =   3855
         Width           =   840
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   0
         Left            =   540
         Picture         =   "InsulinDM.frx":7B0C
         Top             =   3930
         Width           =   480
      End
      Begin VB.Image Image3 
         Height          =   480
         Index           =   1
         Left            =   600
         Picture         =   "InsulinDM.frx":83D6
         Top             =   3975
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   0
         Left            =   60
         Picture         =   "InsulinDM.frx":8CA0
         Top             =   3930
         Width           =   480
      End
      Begin VB.Image Image2 
         Height          =   480
         Index           =   1
         Left            =   90
         Picture         =   "InsulinDM.frx":956A
         Top             =   3960
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Dinner"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   225
         Index           =   11
         Left            =   6225
         TabIndex        =   62
         Top             =   3825
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "Dinner"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   10
         Left            =   6225
         TabIndex        =   61
         Top             =   3840
         Width           =   555
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "CHO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   195
         Index           =   33
         Left            =   6900
         TabIndex        =   86
         Top             =   4680
         Width           =   345
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H80000009&
         BackStyle       =   0  'Transparent
         Caption         =   "CHO"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   195
         Index           =   32
         Left            =   6930
         TabIndex        =   85
         Top             =   4695
         Width           =   345
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00808080&
         BackStyle       =   1  'Opaque
         FillColor       =   &H00C0C0C0&
         FillStyle       =   6  'Cross
         Height          =   1050
         Left            =   1125
         Top             =   3870
         Width           =   6360
      End
      Begin VB.Image Image4 
         BorderStyle     =   1  'Fixed Single
         Height          =   735
         Index           =   1
         Left            =   7590
         Picture         =   "InsulinDM.frx":9E34
         Stretch         =   -1  'True
         Top             =   3945
         Visible         =   0   'False
         Width           =   690
      End
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      ForeColor       =   &H80000008&
      Height          =   2235
      Left            =   2745
      ScaleHeight     =   2205
      ScaleWidth      =   2655
      TabIndex        =   53
      Top             =   735
      Visible         =   0   'False
      Width           =   2685
      Begin VB.CommandButton Command1 
         BackColor       =   &H8000000E&
         Caption         =   "ok"
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   54
         Top             =   120
         Width           =   390
      End
      Begin MSACAL.Calendar Calendar1 
         Height          =   1890
         Left            =   0
         TabIndex        =   55
         Top             =   405
         Width           =   2670
         _Version        =   524288
         _ExtentX        =   4710
         _ExtentY        =   3334
         _StockProps     =   1
         BackColor       =   12648447
         Year            =   2007
         Month           =   10
         Day             =   15
         DayLength       =   1
         MonthLength     =   2
         DayFontColor    =   16711680
         FirstDay        =   7
         GridCellEffect  =   1
         GridFontColor   =   16711680
         GridLinesColor  =   -2147483632
         ShowDateSelectors=   -1  'True
         ShowDays        =   -1  'True
         ShowHorizontalGrid=   -1  'True
         ShowTitle       =   0   'False
         ShowVerticalGrid=   -1  'True
         TitleFontColor  =   0
         ValueIsNull     =   0   'False
         BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Comic Sans MS"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin Project1.BBTime BBTime1 
         Height          =   315
         Left            =   435
         TabIndex        =   56
         Top             =   60
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   556
         BackColor       =   -2147483642
         ForeColor       =   -2147483643
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "x"
         ForeColor       =   &H000000FF&
         Height          =   240
         Index           =   0
         Left            =   2445
         TabIndex        =   58
         Top             =   45
         Width           =   165
      End
      Begin VB.Label Label61 
         BackStyle       =   0  'Transparent
         Caption         =   "N"
         BeginProperty Font 
            Name            =   "Wingdings 2"
            Size            =   21.75
            Charset         =   2
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   315
         Left            =   -60
         TabIndex        =   57
         Top             =   -15
         Width           =   450
      End
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuNewPt 
         Caption         =   "New Patient"
      End
      Begin VB.Menu mnuClear 
         Caption         =   "Clear"
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "Open"
      End
      Begin VB.Menu mnuSave 
         Caption         =   "Save"
      End
      Begin VB.Menu mnuSaveAs 
         Caption         =   "Save As"
      End
      Begin VB.Menu eyedg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuDisclaimer 
         Caption         =   "Disclaimer"
      End
      Begin VB.Menu mnuDoc 
         Caption         =   "Documentation"
      End
      Begin VB.Menu sdfdf 
         Caption         =   "-"
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
      End
   End
   Begin VB.Menu mnuDate 
      Caption         =   "Date"
      Visible         =   0   'False
      Begin VB.Menu mnuInput 
         Caption         =   "Input Date"
      End
      Begin VB.Menu mnuToday 
         Caption         =   "Set Date as Today"
      End
   End
   Begin VB.Menu mnuTime 
      Caption         =   "Time"
      Visible         =   0   'False
      Begin VB.Menu mnuInputTime 
         Caption         =   "Input Time"
      End
      Begin VB.Menu mnuTimeNow 
         Caption         =   "Set time as Now"
      End
   End
   Begin VB.Menu mnuFood 
      Caption         =   "Food"
      Visible         =   0   'False
      Begin VB.Menu mnuGmCHO 
         Caption         =   "Grams CHO"
      End
      Begin VB.Menu mnuGI 
         Caption         =   "Glycemic Index"
      End
      Begin VB.Menu mnuKcals 
         Caption         =   "Total Calories"
      End
   End
   Begin VB.Menu mnuInsulin 
      Caption         =   "Insulin"
      Visible         =   0   'False
      Begin VB.Menu mnuTimeGiven 
         Caption         =   "Time Given"
      End
      Begin VB.Menu mnuUnits 
         Caption         =   "Units"
      End
      Begin VB.Menu mnuGmCHOo 
         Caption         =   "gm CHO"
      End
   End
   Begin VB.Menu mnuBreak 
      Caption         =   "Break"
      Visible         =   0   'False
      Begin VB.Menu mnubrUnits 
         Caption         =   "Units"
      End
      Begin VB.Menu mnuBrTime 
         Caption         =   "Time"
      End
      Begin VB.Menu mnuBrGms 
         Caption         =   "gms CHO"
      End
   End
   Begin VB.Menu mnuGraphx 
      Caption         =   "Graphx"
      Visible         =   0   'False
      Begin VB.Menu mnugToday 
         Caption         =   "Graph Today"
         Index           =   0
      End
      Begin VB.Menu mnugToday 
         Caption         =   "Graph This Week"
         Index           =   1
      End
      Begin VB.Menu mnugToday 
         Caption         =   "Graph Past 2 Weeks"
         Index           =   2
      End
      Begin VB.Menu mnugToday 
         Caption         =   "Graph Past Month"
         Index           =   3
      End
      Begin VB.Menu mnugToday 
         Caption         =   "Graph Past 2 Month"
         Index           =   4
      End
      Begin VB.Menu mnugToday 
         Caption         =   "Graph Past 6 Month"
         Index           =   5
      End
      Begin VB.Menu mnugToday 
         Caption         =   "Graph Past Year"
         Index           =   6
      End
   End
End
Attribute VB_Name = "InsulinDM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 0

Dim TimeIndex As Integer, InsulinFlag As Boolean, InsulinIndex As Integer

Private Sub HbDate_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Message, Title, Default, MyValue
Message = "MMDDYYYY"
Title = "Date of Hba1c Assay"
Default = HbDate.Caption
HbDate.Caption = InputBox(Message, Title, Default, Me.Left, Me.Top)
If HbDate.Caption = "" Then HbDate.Caption = "Date"

End Sub

Private Sub Command1_Click()
    Picture2.Visible = False
    BBTime1.Visible = True
    
If Picture2.Height = 2250 Then
    Label1.Caption = Calendar1.Month & "/" & Calendar1.Day & "/" & Calendar1.Year
    If Label1.Caption = "" Then Label1.Caption = "Set Date"
Else
    If InsulinFlag = True Then
        InsulinFlag = False
        Select Case InsulinIndex
            Case 0
                Br(0).Caption = BBTime1.Text
            Case 1
                L(0).Caption = BBTime1.Text
            Case 2
                D(0).Caption = BBTime1.Text
            Case 3
                BS(0).Caption = BBTime1.Text
            Case 4
                LS(0).Caption = BBTime1.Text
            Case 5
                DS(0).Caption = BBTime1.Text
        End Select
    Else
        Label2(TimeIndex).Caption = BBTime1.Text
        If Label2(TimeIndex).Caption = "" Then Label2(TimeIndex).Caption = "Meal Time"
    End If
End If
Picture2.Height = 2250

End Sub

Private Sub Command2_Click(Index As Integer)
    Select Case Index
        Case 0
            PresentDay = PresentDay - 1
            If PresentDay < 0 Then PresentDay = 0
            Filling PresentDay
        Case 1
            PresentDay = PresentDay + 1
            If PresentDay > j - 1 Then PresentDay = j - 1
            Filling PresentDay
    End Select
End Sub

Private Sub Form_Load()
Dim NN As String
    Picture2.Height = 2250
    Picture2.ZOrder 0
    BBTime1.Text = Time
    Calendar1.Month = Month(Date)
    Calendar1.Day = Day(Date)
    Calendar1.Year = Year(Date)
    If Dir(App.Path & "\PatientName") <> "" Then
        Open App.Path & "\PatientName" For Input As #1
            Line Input #1, NN
            Label5.Caption = NN
        Close #1
    End If
    Image2_MouseUp 0, 0, 0, 0, 0
    If Dir(App.Path & "\Firsttime") = "" Then
        Open App.Path & "\Firsttime" For Output As #1
        Close #1
        Load Disclaimer
        Disclaimer.Show
    End If
End Sub

Private Sub Image1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
    Case 0
        Image1(0).Visible = False
        Image1(11).Visible = True
    Case 1
        Image1(1).Visible = False
        Image1(10).Visible = True
    Case 2
        Image1(2).Visible = False
        Image1(9).Visible = True
    Case 3
        Image1(3).Visible = False
        Image1(8).Visible = True
    Case 4
        Image1(4).Visible = False
        Image1(7).Visible = True
    Case 5
        Image1(5).Visible = False
        Image1(6).Visible = True

End Select

End Sub

Private Sub Image1_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
InsulinIndex = Index
Select Case Index
    Case 0
        Image1(0).Visible = True
        Image1(11).Visible = False
    Case 1
        Image1(1).Visible = True
        Image1(10).Visible = False
    Case 2
        Image1(2).Visible = True
        Image1(9).Visible = False
    Case 3
        Image1(3).Visible = True
        Image1(8).Visible = False
    Case 4
        Image1(4).Visible = True
        Image1(7).Visible = False
    Case 5
        Image1(5).Visible = True
        Image1(6).Visible = False
End Select
PopupMenu mnuInsulin
End Sub

Private Sub Image2_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image2(0).Visible = False
    Image2(1).Visible = True
End Sub

Private Sub Image2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Dim L As Integer, m As Integer, n As Integer, HoldingCo As String
    
    Image2(0).Visible = True
    Image2(1).Visible = False
    Open App.Path & "\Data.dat" For Input As #1
        For i = 0 To 36
            Line Input #1, Holder(i)
        Next
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(0), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, Dateof(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(1), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, BreakfastTime(j)
            j = j + 1
        Loop
    Close #1
    
    
    
    j = 0
    HoldingCo = Replace(Holder(2), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, LunchTime(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(3), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, DinnerTime(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(4), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, BSnackTime(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(5), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, LSnackTime(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(6), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, DSnackTime(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(7), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, BreakfastMealText(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(8), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, LunchMealText(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(9), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, DinnerMealText(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(10), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, BSnackMealText(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(11), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, LSnackMealText(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(12), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, DSnackMealText(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(13), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, BSBr(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(14), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, BSLu(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(15), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, BSDin(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(16), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, BSBs(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(17), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, BSLs(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(18), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, BSDs(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(19), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, UnitsBr(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(20), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, UnitsLu(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(21), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, UnitsDin(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(22), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, UnitsBs(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(23), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, UnitsLs(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(24), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, UnitsDs(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(25), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, CHOBr(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(26), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, CHOLu(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(27), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, CHODin(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(28), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, CHOBs(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(29), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, CHOLs(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(30), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, CHODs(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(31), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, TimeBr(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(32), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, TimeLu(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(33), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, TimeDin(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(34), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, TimeBs(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(35), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, TimeLs(j)
            j = j + 1
        Loop
    Close #1
    
    j = 0
    HoldingCo = Replace(Holder(36), ",", vbCrLf)
    Open App.Path & "\HoldingCo" For Output As #1
        Print #1, HoldingCo
    Close #1
    Open App.Path & "\HoldingCo" For Input As #1
        Do While Not EOF(1)
            Line Input #1, TimeDs(j)
            j = j + 1
        Loop
    Close #1
    Open App.Path & "\Dataaa.dat" For Output As #1
        For i = 0 To 36
            Print #1, Holder(i)
        Next
    Close #1
    
    PresentDay = j - 1
 Fillin
 
End Sub


Sub Fillin()

Label1.Caption = Dateof(j - 1)
'Label5.Caption = Name
Label2(0).Caption = BreakfastTime(j - 1)
Label2(1).Caption = LunchTime(j - 1)
Label2(2).Caption = DinnerTime(j - 1)
Label2(3).Caption = BSnackTime(j - 1)
Label2(4).Caption = LSnackTime(j - 1)
Label2(5).Caption = DSnackTime(j - 1)
Text1(0).Text = BreakfastMealText(j - 1)
Text1(1).Text = LunchMealText(j - 1)
Text1(2).Text = DinnerMealText(j - 1)
Text1(3).Text = BSnackMealText(j - 1)
Text1(4).Text = LSnackMealText(j - 1)
Text1(5).Text = DSnackMealText(j - 1)
Text2(0).Text = BSBr(j - 1)
Text2(1).Text = BSLu(j - 1)
Text2(2).Text = BSDin(j - 1)
Text2(3).Text = BSBs(j - 1)
Text2(4).Text = BSLs(j - 1)
Text2(5).Text = BSDs(j - 1)
Br(1).Caption = UnitsBr(j - 1)
L(1).Caption = UnitsLu(j - 1)
D(1).Caption = UnitsDin(j - 1)
BS(1).Caption = UnitsBs(j - 1)
LS(1).Caption = UnitsLs(j - 1)
DS(1).Caption = UnitsDs(j - 1)
Br(2).Caption = CHOBr(j - 1)
L(2).Caption = CHOLu(j - 1)
D(2).Caption = CHODin(j - 1)
BS(2).Caption = CHOBs(j - 1)
LS(2).Caption = CHOLs(j - 1)
DS(2).Caption = CHODs(j - 1)
Br(0).Caption = TimeBr(j - 1)
L(0).Caption = TimeLu(j - 1)
D(0).Caption = TimeDin(j - 1)
BS(0).Caption = TimeBs(j - 1)
LS(0).Caption = TimeLs(j - 1)
DS(0).Caption = TimeDs(j - 1)
End Sub
Public Sub Filling(Temp As Long)
On Error Resume Next
Label1.Caption = Dateof(Temp)
'Label5.Caption = Name
Label2(0).Caption = BreakfastTime(Temp)
Label2(1).Caption = LunchTime(Temp)
Label2(2).Caption = DinnerTime(Temp)
Label2(3).Caption = BSnackTime(Temp)
Label2(4).Caption = LSnackTime(Temp)
Label2(5).Caption = DSnackTime(Temp)
Text1(0).Text = BreakfastMealText(Temp)
Text1(1).Text = LunchMealText(Temp)
Text1(2).Text = DinnerMealText(Temp)
Text1(3).Text = BSnackMealText(Temp)
Text1(4).Text = LSnackMealText(Temp)
Text1(5).Text = DSnackMealText(Temp)
Text2(0).Text = BSBr(Temp)
Text2(1).Text = BSLu(Temp)
Text2(2).Text = BSDin(Temp)
Text2(3).Text = BSBs(Temp)
Text2(4).Text = BSLs(Temp)
Text2(5).Text = BSDs(Temp)
Br(1).Caption = UnitsBr(Temp)
L(1).Caption = UnitsLu(Temp)
D(1).Caption = UnitsDin(Temp)
BS(1).Caption = UnitsBs(Temp)
LS(1).Caption = UnitsLs(Temp)
DS(1).Caption = UnitsDs(Temp)
Br(2).Caption = CHOBr(Temp)
L(2).Caption = CHOLu(Temp)
D(2).Caption = CHODin(Temp)
BS(2).Caption = CHOBs(Temp)
LS(2).Caption = CHOLs(Temp)
DS(2).Caption = CHODs(Temp)
Br(0).Caption = TimeBr(Temp)
L(0).Caption = TimeLu(Temp)
D(0).Caption = TimeDin(Temp)
BS(0).Caption = TimeBs(Temp)
LS(0).Caption = TimeLs(Temp)
DS(0).Caption = TimeDs(Temp)
End Sub

Private Sub Image3_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image3(0).Visible = False
    Image3(1).Visible = True
End Sub

Private Sub Image3_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
'On Error Resume Next
Dim Hold, Hold2
    For i = 0 To 36: Holder(i) = "": Next
    Image3(0).Visible = True
    Image3(1).Visible = False
    Dateof(j) = Label1.Caption
    BreakfastTime(j) = Label2(0).Caption
    LunchTime(j) = Label2(1).Caption
    DinnerTime(j) = Label2(2).Caption
    BSnackTime(j) = Label2(3).Caption
    LSnackTime(j) = Label2(4).Caption
    DSnackTime(j) = Label2(5).Caption
    BreakfastMealText(j) = Text1(0).Text
    LunchMealText(j) = Text1(1).Text
    DinnerMealText(j) = Text1(2).Text
    BSnackMealText(j) = Text1(3).Text
    LSnackMealText(j) = Text1(4).Text
    DSnackMealText(j) = Text1(5).Text
    BSBr(j) = Text2(0).Text
    BSLu(j) = Text2(1).Text
    BSDin(j) = Text2(2).Text
    BSBs(j) = Text2(3).Text
    BSLs(j) = Text2(4).Text
    BSDs(j) = Text2(5).Text
    UnitsBr(j) = Br(1).Caption
    UnitsLu(j) = L(1).Caption
    UnitsDin(j) = D(1).Caption
    UnitsBs(j) = BS(1).Caption
    UnitsLs(j) = LS(1).Caption
    UnitsDs(j) = DS(1).Caption
    CHOBr(j) = Br(2).Caption
    CHOLu(j) = L(2).Caption
    CHODin(j) = D(2).Caption
    CHOBs(j) = BS(2).Caption
    CHOLs(j) = LS(2).Caption
    CHODs(j) = DS(2).Caption
    TimeBr(j) = Br(0).Caption
    TimeLu(j) = L(0).Caption
    TimeDin(j) = D(0).Caption
    TimeBs(j) = BS(0).Caption
    TimeLs(j) = LS(0).Caption
    TimeDs(j) = DS(0).Caption
    For i = 0 To j
            Holder(0) = Holder(0) & Dateof(i) & ","
            Holder(1) = Holder(1) & BreakfastTime(i) & ","
            Holder(2) = Holder(2) & LunchTime(i) & ","
            Holder(3) = Holder(3) & DinnerTime(i) & ","
            Holder(4) = Holder(4) & BSnackTime(i) & ","
            Holder(5) = Holder(5) & LSnackTime(i) & ","
            Holder(6) = Holder(6) & DSnackTime(i) & ","
            Holder(7) = Holder(7) & BreakfastMealText(i) & ","
            Holder(8) = Holder(8) & LunchMealText(i) & ","
            Holder(9) = Holder(9) & DinnerMealText(i) & ","
            Holder(10) = Holder(10) & BSnackMealText(i) & ","
            Holder(11) = Holder(11) & LSnackMealText(i) & ","
            Holder(12) = Holder(12) & DSnackMealText(i) & ","
            Holder(13) = Holder(13) & BSBr(i) & ","
            Holder(14) = Holder(14) & BSLu(i) & ","
            Holder(15) = Holder(15) & BSDin(i) & ","
            Holder(16) = Holder(16) & BSBs(i) & ","
            Holder(17) = Holder(17) & BSLs(i) & ","
            Holder(18) = Holder(18) & BSDs(i) & ","
            Holder(19) = Holder(19) & UnitsBr(i) & ","
            Holder(20) = Holder(20) & UnitsLu(i) & ","
            Holder(21) = Holder(21) & UnitsDin(i) & ","
            Holder(22) = Holder(22) & UnitsBs(i) & ","
            Holder(23) = Holder(23) & UnitsLs(i) & ","
            Holder(24) = Holder(24) & UnitsDs(i) & ","
            Holder(25) = Holder(25) & CHOBr(i) & ","
            Holder(26) = Holder(26) & CHOLu(i) & ","
            Holder(27) = Holder(27) & CHODin(i) & ","
            Holder(28) = Holder(28) & CHOBs(i) & ","
            Holder(29) = Holder(29) & CHOLs(i) & ","
            Holder(30) = Holder(30) & CHODs(i) & ","
            Holder(31) = Holder(31) & TimeBr(i) & ","
            Holder(32) = Holder(32) & TimeLu(i) & ","
            Holder(33) = Holder(33) & TimeDin(i) & ","
            Holder(34) = Holder(34) & TimeBs(i) & ","
            Holder(35) = Holder(35) & TimeLs(i) & ","
            Holder(36) = Holder(36) & TimeDs(i) & ","
    Next
    Open App.Path & "\Data.dat" For Output As #1
        For i = 0 To 36
            Print #1, Left(Holder(i), Len(Holder(i)) - 1)
        Next
    Close #1
    PresentDay = j
    j = j + 1

End Sub

Private Sub Image4_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4(0).Visible = False
    Image4(1).Visible = True
End Sub

Private Sub Image4_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Image4(0).Visible = True
    Image4(1).Visible = False
    PopupMenu mnuGraphx
    'Load Chart
    'Chart.Show
End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    PopupMenu mnuDate
End Sub

Private Sub Label11_Click(Index As Integer)

End Sub

Private Sub Label2_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Picture2.Height = 2250
    BBTime1.Visible = True
    TimeIndex = Index
    PopupMenu mnuTime
    
End Sub

Private Sub Label5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Message, Title, Default, MyValue
Message = "Lastname, First, MI"
Title = "Patient Name"
Default = Label5.Caption
Label5.Caption = InputBox(Message, Title, Default, Me.Left, Me.Top)
If Label5.Caption = "" Then
    Label5.Caption = "Input Name"
Else
    Open App.Path & "\PatientName" For Output As #2
        Print #2, Label5.Caption
    Close #2
End If
End Sub


Private Sub Label6_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Picture2.Visible = False
End Sub

Private Sub Label61_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Picture2.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub

Private Sub mnuClear_Click()
ClearTextBoxes Me

End Sub

Private Sub mnuDisclaimer_Click()
    Load Disclaimer
    Disclaimer.Show
End Sub

Private Sub mnuGmCHOo_Click()
Dim Message, Title, Default, MyValue
Message = "How many grams of CHO were covered?"
Title = "Insulin Administration"
MyValue = InputBox(Message, Title, Default, Me.Left, Me.Top)
    Select Case InsulinIndex
        Case 0
            Br(2).Caption = MyValue
        Case 1
            L(2).Caption = MyValue
        Case 2
            D(2).Caption = MyValue
        Case 3
            BS(2).Caption = MyValue
        Case 4
            LS(2).Caption = MyValue
        Case 5
            DS(2).Caption = MyValue
    End Select

End Sub

Private Sub mnugPastYear_Click()
    Load Chart
    Chart.Show

End Sub

Private Sub mnugToday_Click(Index As Integer)
Select Case Index
    Case 0  'today
        Itt = j - 1
        
        If Itt < 0 Then Itt = 0
        Load Chart: Chart.Show
        Chart.HScroll1.Value = 1390
    Case 1  'past week
        Itt = j - 7
        If Itt < 0 Then Itt = 0
        Load Chart: Chart.Show
        Chart.HScroll1.Value = 224
    Case 2  'past 2 week
        Itt = j - 14
        If Itt < 0 Then Itt = 0
        Load Chart: Chart.Show
        Chart.HScroll1.Value = 116
    Case 3  'past month
        Itt = j - 30
        If Itt < 0 Then Itt = 0
        Load Chart: Chart.Show
        Chart.HScroll1.Value = 55
    Case 4  'past 2 month
        Itt = j - 60
        If Itt < 0 Then Itt = 0
        Load Chart: Chart.Show
        Chart.HScroll1.Value = 28
    Case 5  'past 6 month
        Itt = j - 180
        If Itt < 0 Then Itt = 0
        Load Chart: Chart.Show
        Chart.HScroll1.Value = 22
    Case 6  'past year
        Itt = j - 360
        If Itt < 0 Then Itt = 0
        Chart.HScroll1.Value = 10
        Load Chart: Chart.Show
End Select
End Sub

Private Sub mnuInput_Click()
    BBTime1.Visible = False
    Picture2.Visible = True
    Picture2.Height = 2250
End Sub

Private Sub mnuInputTime_Click()
    Picture2.Visible = True
    Picture2.Height = 420

End Sub



Private Sub mnuSave_Click()
    Image3_MouseUp 0, 0, 0, 0, 0
End Sub

Private Sub mnuTimeGiven_Click()
    Picture2.Visible = True
    Picture2.Height = 420
    InsulinFlag = True
End Sub

Private Sub mnuTimeNow_Click()
    Label2(TimeIndex).Caption = Time
End Sub

Private Sub mnuToday_Click()
Label1.Caption = Date

End Sub

Private Sub mnuUnits_Click()
Dim Message, Title, Default, MyValue
Message = "How many Units of Insulin were given?"
Title = "Insulin Administration"
MyValue = InputBox(Message, Title, Default, Me.Left, Me.Top)
    Select Case InsulinIndex
        Case 0
            Br(1).Caption = MyValue
        Case 1
            L(1).Caption = MyValue
        Case 2
            D(1).Caption = MyValue
        Case 3
            BS(1).Caption = MyValue
        Case 4
            LS(1).Caption = MyValue
        Case 5
            DS(1).Caption = MyValue
    End Select

End Sub
Public Sub ClearTextBoxes(frmClearMe As Form)

 Dim txt As Control

'clear the text boxes
 For Each txt In frmClearMe

  If TypeOf txt Is TextBox Then txt.Text = ""

 Next

End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
PopupMenu mnuFile
End Sub

Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Picture2.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If

End Sub
