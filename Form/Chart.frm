VERSION 5.00
Begin VB.Form Chart 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   Caption         =   " "
   ClientHeight    =   6450
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9315
   DrawMode        =   10  'Mask Pen
   FillStyle       =   6  'Cross
   Icon            =   "Chart.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6450
   ScaleWidth      =   9315
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000007&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1485
      Left            =   8040
      ScaleHeight     =   1485
      ScaleWidth      =   1275
      TabIndex        =   42
      Top             =   120
      Width           =   1275
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF0000&
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
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   0
         Left            =   0
         TabIndex        =   48
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H000000FF&
         Caption         =   "Snack 1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   1
         Left            =   0
         TabIndex        =   47
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FF00FF&
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
         ForeColor       =   &H0000FFFF&
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   46
         Top             =   480
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         Caption         =   "Snack 2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Index           =   3
         Left            =   0
         TabIndex        =   45
         Top             =   720
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H0000FFFF&
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   4
         Left            =   0
         TabIndex        =   44
         Top             =   960
         Width           =   1245
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         BackColor       =   &H0000FF00&
         Caption         =   "Snack 3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   43
         Top             =   1200
         Width           =   1245
      End
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   135
      LargeChange     =   233
      Left            =   885
      Max             =   1390
      Min             =   1
      TabIndex        =   41
      Top             =   6060
      Value           =   60
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000006&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   6075
      Left            =   0
      ScaleHeight     =   6075
      ScaleWidth      =   900
      TabIndex        =   4
      Top             =   -15
      Width           =   900
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   4
         Left            =   0
         TabIndex        =   33
         Top             =   15
         Width           =   210
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   6
         Left            =   435
         TabIndex        =   32
         Top             =   30
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   5
         Left            =   0
         TabIndex        =   30
         Top             =   945
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Index           =   15
         Left            =   0
         TabIndex        =   28
         Top             =   4965
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Index           =   16
         Left            =   0
         TabIndex        =   26
         Top             =   4095
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "175"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   2
         Left            =   0
         TabIndex        =   24
         Top             =   2910
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "350 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   0
         Left            =   0
         TabIndex        =   22
         Top             =   2010
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mg/dl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   0
         Left            =   420
         TabIndex        =   16
         Top             =   2025
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   3885
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   195
         Index           =   3
         Left            =   0
         TabIndex        =   14
         Top             =   1740
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mg/dl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   1
         Left            =   405
         TabIndex        =   12
         Top             =   2880
         Width           =   420
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mg/dl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   225
         Index           =   2
         Left            =   390
         TabIndex        =   11
         Top             =   3885
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Index           =   17
         Left            =   0
         TabIndex        =   10
         Top             =   5790
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   7
         Left            =   360
         TabIndex        =   9
         Top             =   945
         Width           =   420
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   225
         Index           =   8
         Left            =   420
         TabIndex        =   8
         Top             =   1695
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "gm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Index           =   18
         Left            =   465
         TabIndex        =   7
         Top             =   4110
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "gm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Index           =   20
         Left            =   510
         TabIndex        =   6
         Top             =   4935
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "gm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   195
         Index           =   22
         Left            =   540
         TabIndex        =   5
         Top             =   5820
         Width           =   240
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   10
         Left            =   420
         TabIndex        =   40
         Top             =   30
         Width           =   420
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   9
         Left            =   345
         TabIndex        =   39
         Top             =   930
         Width           =   420
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Units"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00E0E0E0&
         Height          =   225
         Index           =   11
         Left            =   420
         TabIndex        =   38
         Top             =   1680
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "gm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   21
         Left            =   510
         TabIndex        =   37
         Top             =   4920
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "gm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   23
         Left            =   540
         TabIndex        =   36
         Top             =   5805
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "gm"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   19
         Left            =   465
         TabIndex        =   35
         Top             =   4095
         Width           =   240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H80000012&
         BackStyle       =   0  'Transparent
         Caption         =   "20"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   10
         Left            =   0
         TabIndex        =   34
         Top             =   0
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "10"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   9
         Left            =   0
         TabIndex        =   31
         Top             =   960
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   8
         Left            =   0
         TabIndex        =   29
         Top             =   4950
         Visible         =   0   'False
         Width           =   210
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "100"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   7
         Left            =   0
         TabIndex        =   27
         Top             =   4080
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "175"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   12
         Left            =   0
         TabIndex        =   25
         Top             =   2895
         Width           =   315
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "350 "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   14
         Left            =   0
         TabIndex        =   23
         Top             =   1995
         Width           =   375
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mg/dl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   3
         Left            =   375
         TabIndex        =   21
         Top             =   3870
         Width           =   420
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mg/dl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   4
         Left            =   420
         TabIndex        =   20
         Top             =   2895
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   13
         Left            =   0
         TabIndex        =   19
         Top             =   3870
         Width           =   105
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "mg/dl"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Index           =   5
         Left            =   405
         TabIndex        =   18
         Top             =   2010
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   11
         Left            =   0
         TabIndex        =   17
         Top             =   1725
         Width           =   105
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   6
         Left            =   0
         TabIndex        =   13
         Top             =   5775
         Width           =   105
      End
   End
   Begin VB.Line Line2 
      BorderColor     =   &H80000005&
      BorderWidth     =   4
      X1              =   915
      X2              =   20000
      Y1              =   6000
      Y2              =   6000
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H80000008&
      Caption         =   " Time"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   480
      Left            =   915
      TabIndex        =   0
      Top             =   6045
      Width           =   8475
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6840
      TabIndex        =   2
      Top             =   1785
      Width           =   75
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6795
      TabIndex        =   1
      Top             =   3795
      Width           =   75
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   4
      Index           =   1
      X1              =   945
      X2              =   20000
      Y1              =   4005
      Y2              =   4005
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackColor       =   &H00C0FFFF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   6840
      TabIndex        =   3
      Top             =   5730
      Width           =   75
   End
   Begin VB.Line Line3 
      BorderColor     =   &H0000FFFF&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   5490
      X2              =   5490
      Y1              =   -7000
      Y2              =   15000
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFF00&
      BorderWidth     =   2
      Visible         =   0   'False
      X1              =   5300
      X2              =   5600
      Y1              =   1950
      Y2              =   1950
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   367
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   366
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   365
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   364
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   363
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   362
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   361
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   360
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   359
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   358
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   357
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   356
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   355
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   354
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   353
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   352
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   351
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   350
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   349
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   348
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   347
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   346
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   345
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   344
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   343
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   342
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   341
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   340
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   339
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   338
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   337
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   336
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   335
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   334
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   333
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   332
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   331
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   330
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   329
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   328
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   327
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   326
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   325
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   324
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   323
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   322
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   321
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   320
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   319
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   318
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   317
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   316
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   315
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   314
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   313
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   312
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   311
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   310
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   309
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   308
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   307
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   306
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   305
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   304
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   303
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   302
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   301
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   300
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   299
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   298
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   297
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   296
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   295
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   294
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   293
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   292
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   291
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   290
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   289
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   288
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   287
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   286
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   285
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   284
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   283
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   282
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   281
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   280
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   279
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   278
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   277
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   276
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   275
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   274
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   273
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   272
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   271
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   270
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   269
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   268
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   267
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   266
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   265
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   264
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   263
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   262
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   261
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   260
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   259
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   258
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   257
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   256
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   255
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   254
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   253
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   252
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   251
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   250
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   249
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   248
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   247
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   246
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   245
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   244
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   243
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   242
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   241
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   240
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   239
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   238
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   237
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   236
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   235
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   234
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   233
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   232
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   231
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   230
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   229
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   228
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   227
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   226
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   225
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   224
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   223
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   222
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   221
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   220
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   219
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   218
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   217
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   216
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   215
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   214
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   213
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   212
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   211
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   210
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   209
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   208
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   207
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   206
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   205
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   204
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   203
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   202
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   201
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   200
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   199
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   198
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   197
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   196
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   195
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   194
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   193
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   192
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   191
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   190
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   189
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   188
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   187
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   186
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   185
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   184
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   183
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   182
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   181
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   180
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   179
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   178
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   177
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   176
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   175
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   174
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   173
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   172
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   171
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   170
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   169
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   168
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   167
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   166
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   165
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   164
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   163
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   162
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   161
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   160
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   159
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   158
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   157
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   156
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   155
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   154
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   153
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   152
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   151
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   150
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   149
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   148
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   147
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   146
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   145
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   144
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   143
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   142
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   141
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   140
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   139
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   138
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   137
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   136
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   135
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   134
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   133
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   132
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   131
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   130
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   129
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   128
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   127
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   126
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   125
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   124
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   123
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   122
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   121
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   120
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   119
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   118
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   117
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   116
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   115
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   114
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   113
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   112
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   111
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   110
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   109
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   108
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   107
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   106
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   105
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   104
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   103
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   102
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   101
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   100
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   99
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   98
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   97
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   96
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   95
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   94
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   93
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   92
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   91
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   90
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   89
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   88
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   87
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   86
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   85
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   84
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   83
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   82
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   81
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   80
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   79
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   78
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   77
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   76
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   75
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   74
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   73
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   72
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   71
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   70
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   69
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   68
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   67
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   66
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   65
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   64
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   63
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   62
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   61
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   60
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   59
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   58
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   57
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   56
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   55
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   54
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   53
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   52
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   51
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   50
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   49
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   48
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   47
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   46
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   45
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   44
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   43
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   42
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   41
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   40
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   39
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   38
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   37
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   36
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   35
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   34
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   33
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   32
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   31
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   30
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   29
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   28
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   27
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   26
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   25
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   24
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   23
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   22
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   21
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   20
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   19
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   18
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   17
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   16
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   15
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   14
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   13
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   12
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   11
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   10
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   9
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   8
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   7
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   6
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   5
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   4
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   3
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   2
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnCHO 
      BorderColor     =   &H00FF00FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   5070
      X2              =   5170
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   367
      X1              =   4935
      X2              =   5035
      Y1              =   6225
      Y2              =   6225
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   366
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   365
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   364
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   363
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   362
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   361
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   360
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   359
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   358
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   357
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   356
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   355
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   354
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   353
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   352
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   351
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   350
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   349
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   348
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   347
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   346
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   345
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   344
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   343
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   342
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   341
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   340
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   339
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   338
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   337
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   336
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   335
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   334
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   333
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   332
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   331
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   330
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   329
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   328
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   327
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   326
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   325
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   324
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   323
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   322
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   321
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   320
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   319
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   318
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   317
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   316
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   315
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   314
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   313
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   312
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   311
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   310
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   309
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   308
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   307
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   306
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   305
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   304
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   303
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   302
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   301
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   300
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   299
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   298
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   297
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   296
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   295
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   294
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   293
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   292
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   291
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   290
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   289
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   288
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   287
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   286
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   285
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   284
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   283
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   282
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   281
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   280
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   279
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   278
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   277
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   276
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   275
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   274
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   273
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   272
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   271
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   270
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   269
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   268
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   267
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   266
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   265
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   264
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   263
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   262
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   261
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   260
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   259
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   258
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   257
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   256
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   255
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   254
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   253
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   252
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   251
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   250
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   249
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   248
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   247
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   246
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   245
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   244
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   243
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   242
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   241
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   240
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   239
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   238
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   237
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   236
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   235
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   234
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   233
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   232
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   231
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   230
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   229
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   228
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   227
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   226
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   225
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   224
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   223
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   222
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   221
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   220
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   219
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   218
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   217
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   216
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   215
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   214
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   213
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   212
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   211
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   210
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   209
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   208
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   207
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   206
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   205
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   204
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   203
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   202
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   201
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   200
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   199
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   198
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   197
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   196
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   195
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   194
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   193
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   192
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   191
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   190
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   189
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   188
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   187
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   186
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   185
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   184
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   183
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   182
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   181
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   180
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   179
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   178
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   177
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   176
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   175
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   174
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   173
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   172
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   171
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   170
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   169
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   168
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   167
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   166
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   165
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   164
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   163
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   162
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   161
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   160
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   159
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   158
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   157
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   156
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   155
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   154
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   153
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   152
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   151
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   150
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   149
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   148
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   147
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   146
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   145
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   144
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   143
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   142
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   141
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   140
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   139
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   138
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   137
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   136
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   135
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   134
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   133
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   132
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   131
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   130
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   129
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   128
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   127
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   126
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   125
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   124
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   123
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   122
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   121
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   120
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   119
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   118
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   117
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   116
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   115
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   114
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   113
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   112
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   111
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   110
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   109
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   108
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   107
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   106
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   105
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   104
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   103
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   102
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   101
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   100
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   99
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   98
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   97
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   96
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   95
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   94
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   93
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   92
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   91
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   90
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   89
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   88
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   87
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   86
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   85
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   84
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   83
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   82
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   81
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   80
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   79
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   78
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   77
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   76
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   75
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   74
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   73
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   72
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   71
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   70
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   69
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   68
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   67
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   66
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   65
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   64
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   63
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   62
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   61
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   60
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   59
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   58
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   57
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   56
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   55
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   54
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   53
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   52
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   51
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   50
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   49
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   48
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   47
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   46
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   45
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   44
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   43
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   42
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   41
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   40
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   39
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   38
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   37
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   36
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   35
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   34
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   33
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   32
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   31
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   30
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   29
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   28
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   27
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   26
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   25
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   24
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   23
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   22
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   21
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   20
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   19
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   18
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   17
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   16
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   15
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   14
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   13
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   12
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   11
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   10
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   9
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   8
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   7
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   6
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   5
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   4
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   3
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   2
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   1
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnGlucose 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Index           =   0
      X1              =   5025
      X2              =   5125
      Y1              =   6210
      Y2              =   6210
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   367
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   366
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   365
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   364
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   363
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   362
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   361
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   360
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   359
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   358
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   357
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   356
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   355
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   354
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   353
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   352
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   351
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   350
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   349
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   348
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   347
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   346
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   345
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   344
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   343
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   342
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   341
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   340
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   339
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   338
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   337
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   336
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   335
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   334
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   333
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   332
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   331
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   330
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   329
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   328
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   327
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   326
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   325
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   324
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   323
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   322
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   321
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   320
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   319
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   318
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   317
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   316
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   315
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   314
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   313
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   312
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   311
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   310
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   309
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   308
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   307
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   306
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   305
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   304
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   303
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   302
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   301
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   300
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   299
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   298
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   297
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   296
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   295
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   294
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   293
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   292
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   291
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   290
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   289
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   288
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   287
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   286
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   285
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   284
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   283
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   282
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   281
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   280
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   279
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   278
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   277
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   276
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   275
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   274
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   273
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   272
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   271
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   270
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   269
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   268
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   267
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   266
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   265
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   264
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   263
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   262
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   261
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   260
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   259
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   258
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   257
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   256
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   255
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   254
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   253
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   252
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   251
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   250
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   249
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   248
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   247
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   246
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   245
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   244
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   243
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   242
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   241
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   240
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   239
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   238
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   237
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   236
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   235
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   234
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   233
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   232
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   231
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   230
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   229
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   228
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   227
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   226
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   225
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   224
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   223
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   222
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   221
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   220
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   219
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   218
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   217
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   216
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   215
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   214
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   213
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   212
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   211
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   210
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   209
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   208
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   207
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   206
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   205
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   204
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   203
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   202
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   201
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   200
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   199
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   198
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   197
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   196
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   195
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   194
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   193
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   192
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   191
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   190
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   189
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   188
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   187
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   186
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   185
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   184
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   183
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   182
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   181
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   180
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   179
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   178
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   177
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   176
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   175
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   174
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   173
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   172
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   171
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   170
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   169
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   168
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   167
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   166
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   165
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   164
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   163
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   162
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   161
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   160
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   159
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   158
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   157
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   156
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   155
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   154
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   153
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   152
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   151
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   150
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   149
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   148
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   147
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   146
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   145
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   144
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   143
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   142
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   141
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   140
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   139
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   138
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   137
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   136
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   135
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   134
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   133
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   132
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   131
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   130
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   129
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   128
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   127
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   126
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   125
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   124
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   123
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   122
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   121
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   120
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   119
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   118
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   117
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   116
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   115
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   114
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   113
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   112
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   111
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   110
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   109
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   108
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   107
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   106
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   105
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   104
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   103
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   102
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   101
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   100
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   99
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   98
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   97
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   96
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   95
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   94
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   93
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   92
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   91
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   90
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   89
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   88
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   87
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   86
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   85
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   84
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   83
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   82
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   81
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   80
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   79
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   78
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   77
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   76
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   75
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   74
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   73
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   72
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   71
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   70
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   69
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   68
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   67
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   66
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   65
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   64
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   63
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   62
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   61
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   60
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   59
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   58
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   57
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   56
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   55
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   54
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   53
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   52
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   51
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   50
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   49
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   48
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   47
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   46
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   45
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   44
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   43
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   42
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   41
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   40
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   39
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   38
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   37
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   36
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   35
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   34
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   33
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   32
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   31
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   30
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   29
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   28
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   27
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   26
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   25
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   24
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   23
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   22
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   21
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   20
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   19
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   18
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   17
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   16
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   15
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   14
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   13
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   12
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   11
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   10
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   9
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   8
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   7
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   6
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   5
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   4
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   3
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   2
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   1
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line LnInsulin 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   2
      Index           =   0
      X1              =   5025
      X2              =   5115
      Y1              =   6240
      Y2              =   6240
   End
   Begin VB.Line Line1 
      BorderColor     =   &H80000005&
      BorderWidth     =   4
      Index           =   0
      X1              =   945
      X2              =   20000
      Y1              =   1995
      Y2              =   1995
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H80000006&
      BorderColor     =   &H80000005&
      BorderWidth     =   2
      FillColor       =   &H00FFFFFF&
      FillStyle       =   6  'Cross
      Height          =   6000
      Left            =   915
      Top             =   0
      Width           =   20000
   End
End
Attribute VB_Name = "Chart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

    'Dim NormalWindowStyle As Long
    'Dim HWD As Long
    'NormalWindowStyle = GetWindowLong(HWD, GWL_EXSTYLE)
    'SetWindowLong Me.hwnd, GWL_EXSTYLE, NormalWindowStyle Or WS_EX_LAYERED
    'SetLayeredWindowAttributes Me.hwnd, 0, 215, LWA_ALPHA
    LineWidth = 60
    Plott Itt
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
InsulinDM.Filling Int((X - 945) / (LineWidth * 5))
'Me.Caption = Int((X - 945) / (LineWidth * 5))
InsulinDM.ZOrder 0
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label4.Caption = Int((4000 - Y) / BSmax) & " mg/dl" 'X & "  " & Y
Label5.Caption = Int((2000 - Y) / Umax) & " Units" 'X & "  " & Y
Label6.Caption = Int((6000 - Y) / GMmax) & " gms CHO" 'X & "  " & Y
If Y <= 2000 Then
    Label5.Visible = True
    Label5.Left = X + 150
Else
    Label5.Visible = False
End If
If Y > 2000 And Y < 4000 Then
    Label4.Visible = True
    Label4.Left = X + 150
Else
    Label4.Visible = False
End If
If Y >= 4000 Then
    Label6.Visible = True
    Label6.Left = X + 150
Else
    Label6.Visible = False
End If
If X > Shape1.Left And Y < Shape1.Height Then
    Line3.Visible = True
    Line4.Visible = True
    Line3.X1 = X
    Line3.X2 = X
    Line4.X1 = 0
    Line4.X2 = Me.Width
    Line4.Y1 = Y
    Line4.Y2 = Y
Else
    Label4.Visible = False
    Label5.Visible = False
    Label6.Visible = False
    Line3.Visible = False
    Line4.Visible = False
End If
End Sub

Private Sub Form_Resize()
    If Me.Height > 6960 Then Me.Height = 6960
    If Me.Height < 6960 Then Me.Height = 6960
End Sub

Private Sub HScroll1_Change()
    LineWidth = HScroll1.Value
    Plott Itt
    Me.Caption = HScroll1.Value
End Sub

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim Message, Title, Default, MyValue
Message = "Input New Maximum Insulin Dosage"
Title = "Adjust Dialog"
Default = Label5.Caption
Label1(4).Caption = InputBox(Message, Title, Default, Me.Left, Me.Top)
If Label1(4).Caption = "" Then Label1(4).Caption = Default
Label1(10).Caption = Label1(4).Caption
End Sub
Sub Plott(aG As Integer)
On Error Resume Next
    Umax = 2000 / Label1(4).Caption
    BSmax = 2000 / Label1(0).Caption
    GMmax = 2000 / Label1(16).Caption
    LnGlucose(0).BorderColor = &HFF0000
    LnGlucose(0).X1 = 945
    LnGlucose(0).X2 = LnGlucose(0).X1 + LineWidth
    LnGlucose(0).Y1 = 4000 - BSBr(aG) * BSmax
    LnGlucose(0).Y2 = 4000 - BSBr(aG) * BSmax
    LnGlucose(1).BorderColor = &HFF&
    LnGlucose(1).X1 = LnGlucose(0).X2
    LnGlucose(1).X2 = LnGlucose(1).X1 + LineWidth
    LnGlucose(1).Y1 = LnGlucose(0).Y2
    LnGlucose(1).Y2 = 4000 - BSBs(aG) * BSmax
    LnGlucose(2).BorderColor = &HFF00FF
    LnGlucose(2).X1 = LnGlucose(1).X2
    LnGlucose(2).X2 = LnGlucose(2).X1 + LineWidth
    LnGlucose(2).Y1 = LnGlucose(1).Y2
    LnGlucose(2).Y2 = 4000 - BSLu(aG) * BSmax
    LnGlucose(3).BorderColor = &H80000005
    LnGlucose(3).X1 = LnGlucose(2).X2
    LnGlucose(3).X2 = LnGlucose(3).X1 + LineWidth
    LnGlucose(3).Y1 = LnGlucose(2).Y2
    LnGlucose(3).Y2 = 4000 - BSLs(aG) * BSmax
    LnGlucose(4).BorderColor = &HFFFF&
    LnGlucose(4).X1 = LnGlucose(3).X2
    LnGlucose(4).X2 = LnGlucose(4).X1 + LineWidth
    LnGlucose(4).Y1 = LnGlucose(3).Y2
    LnGlucose(4).Y2 = 4000 - BSDin(aG) * BSmax
    LnGlucose(5).BorderColor = &HFF00&
    LnGlucose(5).X1 = LnGlucose(4).X2
    LnGlucose(5).X2 = LnGlucose(5).X1 + LineWidth
    LnGlucose(5).Y1 = LnGlucose(4).Y2
    LnGlucose(5).Y2 = 4000 - BSDs(aG) * BSmax
    k = 6
    For i = aG + 1 To j - 1
        LnGlucose(k).BorderColor = &HFF0000
        LnGlucose(k).X1 = LnGlucose(k - 1).X2
        LnGlucose(k).X2 = LnGlucose(k).X1 + LineWidth
        LnGlucose(k).Y1 = LnGlucose(k - 1).Y2
        LnGlucose(k).Y2 = 4000 - BSBr(i) * BSmax
        k = k + 1
        LnGlucose(k).BorderColor = &HFF&
        LnGlucose(k).X1 = LnGlucose(k - 1).X2
        LnGlucose(k).X2 = LnGlucose(k).X1 + LineWidth
        LnGlucose(k).Y1 = LnGlucose(k - 1).Y2
        LnGlucose(k).Y2 = 4000 - BSBs(i) * BSmax
        k = k + 1
        LnGlucose(k).BorderColor = &HFF00FF
        LnGlucose(k).X1 = LnGlucose(k - 1).X2
        LnGlucose(k).X2 = LnGlucose(k).X1 + LineWidth
        LnGlucose(k).Y1 = LnGlucose(k - 1).Y2
        LnGlucose(k).Y2 = 4000 - BSLu(i) * BSmax
        k = k + 1
        LnGlucose(k).BorderColor = &H80000005
        LnGlucose(k).X1 = LnGlucose(k - 1).X2
        LnGlucose(k).X2 = LnGlucose(k).X1 + LineWidth
        LnGlucose(k).Y1 = LnGlucose(k - 1).Y2
        LnGlucose(k).Y2 = 4000 - BSLs(i) * BSmax
        k = k + 1
        LnGlucose(k).BorderColor = &HFFFF&
        LnGlucose(k).X1 = LnGlucose(k - 1).X2
        LnGlucose(k).X2 = LnGlucose(k).X1 + LineWidth
        LnGlucose(k).Y1 = LnGlucose(k - 1).Y2
        LnGlucose(k).Y2 = 4000 - BSDin(i) * BSmax
        k = k + 1
        LnGlucose(k).BorderColor = &HFF00&
        LnGlucose(k).X1 = LnGlucose(k - 1).X2
        LnGlucose(k).X2 = LnGlucose(k).X1 + LineWidth
        LnGlucose(k).Y1 = LnGlucose(k - 1).Y2
        LnGlucose(k).Y2 = 4000 - BSDs(i) * BSmax
    Next

    LnInsulin(0).BorderColor = &HFF0000
    LnInsulin(0).X1 = 945
    LnInsulin(0).X2 = LnInsulin(0).X1 + LineWidth
    LnInsulin(0).Y1 = 2000 - UnitsBr(aG) * Umax
    LnInsulin(0).Y2 = 2000 - UnitsBr(aG) * Umax
    LnInsulin(1).BorderColor = &HFF&
    LnInsulin(1).X1 = LnInsulin(0).X2
    LnInsulin(1).X2 = LnInsulin(1).X1 + LineWidth
    LnInsulin(1).Y1 = LnInsulin(0).Y2
    LnInsulin(1).Y2 = 2000 - UnitsBs(aG) * Umax
    LnInsulin(2).BorderColor = &HFF00FF
    LnInsulin(2).X1 = LnInsulin(1).X2
    LnInsulin(2).X2 = LnInsulin(2).X1 + LineWidth
    LnInsulin(2).Y1 = LnInsulin(1).Y2
    LnInsulin(2).Y2 = 2000 - UnitsLu(aG) * Umax
    LnInsulin(3).BorderColor = &H80000005
    LnInsulin(3).X1 = LnInsulin(2).X2
    LnInsulin(3).X2 = LnInsulin(3).X1 + LineWidth
    LnInsulin(3).Y1 = LnInsulin(2).Y2
    LnInsulin(3).Y2 = 2000 - UnitsLs(aG) * Umax
    LnInsulin(4).BorderColor = &HFFFF&
    LnInsulin(4).X1 = LnInsulin(3).X2
    LnInsulin(4).X2 = LnInsulin(4).X1 + LineWidth
    LnInsulin(4).Y1 = LnInsulin(3).Y2
    LnInsulin(4).Y2 = 2000 - UnitsDin(aG) * Umax
    LnInsulin(5).BorderColor = &HFF00&
    LnInsulin(5).X1 = LnInsulin(4).X2
    LnInsulin(5).X2 = LnInsulin(5).X1 + LineWidth
    LnInsulin(5).Y1 = LnInsulin(4).Y2
    LnInsulin(5).Y2 = 2000 - UnitsDs(aG) * Umax
    k = 6
    For i = aG + 1 To j - 1
        LnInsulin(k).BorderColor = &HFF0000
        LnInsulin(k).X1 = LnInsulin(k - 1).X2
        LnInsulin(k).X2 = LnInsulin(k).X1 + LineWidth
        LnInsulin(k).Y1 = LnInsulin(k - 1).Y2
        LnInsulin(k).Y2 = 2000 - UnitsBr(i) * Umax
        k = k + 1
        LnInsulin(k).BorderColor = &HFF&
        LnInsulin(k).X1 = LnInsulin(k - 1).X2
        LnInsulin(k).X2 = LnInsulin(k).X1 + LineWidth
        LnInsulin(k).Y1 = LnInsulin(k - 1).Y2
        LnInsulin(k).Y2 = 2000 - UnitsBs(i) * Umax
        k = k + 1
        LnInsulin(k).BorderColor = &HFF00FF
        LnInsulin(k).X1 = LnInsulin(k - 1).X2
        LnInsulin(k).X2 = LnInsulin(k).X1 + LineWidth
        LnInsulin(k).Y1 = LnInsulin(k - 1).Y2
        LnInsulin(k).Y2 = 2000 - UnitsLu(i) * Umax
        k = k + 1
        LnInsulin(k).BorderColor = &H80000005
        LnInsulin(k).X1 = LnInsulin(k - 1).X2
        LnInsulin(k).X2 = LnInsulin(k).X1 + LineWidth
        LnInsulin(k).Y1 = LnInsulin(k - 1).Y2
        LnInsulin(k).Y2 = 2000 - UnitsLs(i) * Umax
        k = k + 1
        LnInsulin(k).BorderColor = &HFFFF&
        LnInsulin(k).X1 = LnInsulin(k - 1).X2
        LnInsulin(k).X2 = LnInsulin(k).X1 + LineWidth
        LnInsulin(k).Y1 = LnInsulin(k - 1).Y2
        LnInsulin(k).Y2 = 2000 - UnitsDin(i) * Umax
        k = k + 1
        LnInsulin(k).BorderColor = &HFF00&
        LnInsulin(k).X1 = LnInsulin(k - 1).X2
        LnInsulin(k).X2 = LnInsulin(k).X1 + LineWidth
        LnInsulin(k).Y1 = LnInsulin(k - 1).Y2
        LnInsulin(k).Y2 = 2000 - UnitsDs(i) * Umax
    Next



    LnCHO(0).BorderColor = &HFF0000
    LnCHO(0).X1 = 945
    LnCHO(0).X2 = LnCHO(0).X1 + LineWidth
    LnCHO(0).Y1 = 6000 - CHOBr(aG) * GMmax
    LnCHO(0).Y2 = 6000 - CHOBr(aG) * GMmax
    LnCHO(1).BorderColor = &HFF&
    LnCHO(1).X1 = LnCHO(0).X2
    LnCHO(1).X2 = LnCHO(1).X1 + LineWidth
    LnCHO(1).Y1 = LnCHO(0).Y2
    LnCHO(1).Y2 = 6000 - CHOBs(aG) * GMmax
    LnCHO(2).BorderColor = &HFF00FF
    LnCHO(2).X1 = LnCHO(1).X2
    LnCHO(2).X2 = LnCHO(2).X1 + LineWidth
    LnCHO(2).Y1 = LnCHO(1).Y2
    LnCHO(2).Y2 = 6000 - CHOLu(aG) * GMmax
    LnCHO(3).BorderColor = &H80000005
    LnCHO(3).X1 = LnCHO(2).X2
    LnCHO(3).X2 = LnCHO(3).X1 + LineWidth
    LnCHO(3).Y1 = LnCHO(2).Y2
    LnCHO(3).Y2 = 6000 - CHOLs(aG) * GMmax
    LnCHO(4).BorderColor = &HFFFF&
    LnCHO(4).X1 = LnCHO(3).X2
    LnCHO(4).X2 = LnCHO(4).X1 + LineWidth
    LnCHO(4).Y1 = LnCHO(3).Y2
    LnCHO(4).Y2 = 6000 - CHODin(aG) * GMmax
    LnCHO(5).BorderColor = &HFF00&
    LnCHO(5).X1 = LnCHO(4).X2
    LnCHO(5).X2 = LnCHO(5).X1 + LineWidth
    LnCHO(5).Y1 = LnCHO(4).Y2
    LnCHO(5).Y2 = 6000 - CHODs(aG) * GMmax
    k = 6
    For i = aG + 1 To j - 1
        LnCHO(k).BorderColor = &HFF0000
        LnCHO(k).X1 = LnCHO(k - 1).X2
        LnCHO(k).X2 = LnCHO(k).X1 + LineWidth
        LnCHO(k).Y1 = LnCHO(k - 1).Y2
        LnCHO(k).Y2 = 6000 - CHOBr(i) * GMmax
        k = k + 1
        LnCHO(k).BorderColor = &HFF&
        LnCHO(k).X1 = LnCHO(k - 1).X2
        LnCHO(k).X2 = LnCHO(k).X1 + LineWidth
        LnCHO(k).Y1 = LnCHO(k - 1).Y2
        LnCHO(k).Y2 = 6000 - CHOBs(i) * GMmax
        k = k + 1
        LnCHO(k).BorderColor = &HFF00FF
        LnCHO(k).X1 = LnCHO(k - 1).X2
        LnCHO(k).X2 = LnCHO(k).X1 + LineWidth
        LnCHO(k).Y1 = LnCHO(k - 1).Y2
        LnCHO(k).Y2 = 6000 - CHOLu(i) * GMmax
        k = k + 1
        LnCHO(k).BorderColor = &H80000005
        LnCHO(k).X1 = LnCHO(k - 1).X2
        LnCHO(k).X2 = LnCHO(k).X1 + LineWidth
        LnCHO(k).Y1 = LnCHO(k - 1).Y2
        LnCHO(k).Y2 = 6000 - CHOLs(i) * GMmax
        k = k + 1
        LnCHO(k).BorderColor = &HFFFF&
        LnCHO(k).X1 = LnCHO(k - 1).X2
        LnCHO(k).X2 = LnCHO(k).X1 + LineWidth
        LnCHO(k).Y1 = LnCHO(k - 1).Y2
        LnCHO(k).Y2 = 6000 - CHODin(i) * GMmax
        k = k + 1
        LnCHO(k).BorderColor = &HFF00&
        LnCHO(k).X1 = LnCHO(k - 1).X2
        LnCHO(k).X2 = LnCHO(k).X1 + LineWidth
        LnCHO(k).Y1 = LnCHO(k - 1).Y2
        LnCHO(k).Y2 = 6000 - CHODs(i) * GMmax
    Next



End Sub

Private Sub Label7_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbLeftButton Then
  ReleaseCapture
  SendMessage Picture2.hwnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&
End If
End Sub
