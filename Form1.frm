VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00D8E9EC&
   Caption         =   "Form1"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   10260
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.Frame Frame2 
      Caption         =   "BarStyle= 0£­Progressbar"
      Height          =   6255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   9135
      Begin VB.Frame Frame5 
         Caption         =   "3-prgDown"
         Height          =   5895
         Left            =   7920
         TabIndex        =   7
         Top             =   240
         Width           =   1095
         Begin HzxYProgressBarTest.HzxYProgressBar HzxYProgressBar4 
            Height          =   5415
            Left            =   360
            TabIndex        =   8
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   9551
            Bar_Pic         =   "Form1.frx":0000
            BarFillDirection=   3
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "2-prgUp"
         Height          =   5895
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   855
         Begin HzxYProgressBarTest.HzxYProgressBar HzxYProgressBar3 
            Height          =   5415
            Left            =   240
            TabIndex        =   6
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   9551
            Bar_Pic         =   "Form1.frx":0152
            BarFillDirection=   2
         End
      End
      Begin VB.Frame Frame3 
         Height          =   5775
         Left            =   1080
         TabIndex        =   4
         Top             =   360
         Width           =   6735
         Begin VB.Frame Frame10 
            Caption         =   "BarFillDirection= 2-prgLeft"
            Height          =   975
            Left            =   120
            TabIndex        =   17
            Top             =   4560
            Width           =   6375
            Begin HzxYProgressBarTest.HzxYProgressBar HzxYProgressBar8 
               Height          =   615
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   6135
               _ExtentX        =   10821
               _ExtentY        =   1085
               BarColorSet     =   0
               Bar_Pic         =   "Form1.frx":02A4
               BarFillDirection=   0
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "BarFillDirection= 1-prgRight, BarColorSet= 1-XP_Default"
            Height          =   1335
            Left            =   120
            TabIndex        =   15
            Top             =   120
            Width           =   6375
            Begin HzxYProgressBarTest.HzxYProgressBar HzxYProgressBar2 
               Height          =   855
               Left            =   120
               TabIndex        =   16
               Top             =   360
               Width           =   6135
               _ExtentX        =   10821
               _ExtentY        =   1508
               Bar_Pic         =   "Form1.frx":05BE
            End
         End
         Begin VB.Frame Frame8 
            Caption         =   "BarColorSet= 0-Custom, and Set ""Bar_Pic"" By yourself "
            Height          =   1215
            Left            =   120
            TabIndex        =   13
            Top             =   3240
            Width           =   6375
            Begin HzxYProgressBarTest.HzxYProgressBar HzxYProgressBar7 
               Height          =   615
               Left            =   120
               TabIndex        =   14
               Top             =   360
               Width           =   6135
               _ExtentX        =   10821
               _ExtentY        =   1085
               BarColorSet     =   0
               Bar_Pic         =   "Form1.frx":0700
               BarSpaceBetweenImages=   0
            End
         End
         Begin VB.Frame Frame7 
            Caption         =   "BarSpaceBetweenImages=5,BorderColor=FF,BackColor=00FFFF"
            Height          =   735
            Left            =   120
            TabIndex        =   11
            Top             =   2400
            Width           =   6375
            Begin HzxYProgressBarTest.HzxYProgressBar HzxYProgressBar6 
               Height          =   255
               Left            =   120
               TabIndex        =   12
               Top             =   360
               Width           =   6135
               _ExtentX        =   10821
               _ExtentY        =   450
               BarColorSet     =   8
               Bar_Pic         =   "Form1.frx":13DA
               BorderColor     =   255
               BackColor       =   12648447
               BarSpaceBetweenImages=   5
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "BarborderStyle= 0-prgNone, BarColorSet= 3-XP_DarkBlue"
            Height          =   975
            Left            =   120
            TabIndex        =   9
            Top             =   1440
            Width           =   6375
            Begin HzxYProgressBarTest.HzxYProgressBar HzxYProgressBar5 
               Height          =   375
               Left            =   120
               TabIndex        =   10
               Top             =   360
               Width           =   6135
               _ExtentX        =   10821
               _ExtentY        =   661
               BarColorSet     =   3
               Bar_Pic         =   "Form1.frx":151C
               BarBorderStyle  =   0
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "BarStyle= 1£­SearchBar"
      Height          =   855
      Left            =   2160
      TabIndex        =   1
      Top             =   120
      Width           =   5895
      Begin HzxYProgressBarTest.HzxYProgressBar HzxYProgressBar1 
         Height          =   375
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   5415
         _ExtentX        =   9551
         _ExtentY        =   661
         BarStyle        =   1
         Bar_Pic         =   "Form1.frx":165E
         BarBorderStyle  =   0
      End
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8760
      Top             =   0
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Go!"
      Height          =   375
      Left            =   720
      TabIndex        =   0
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prgSign As Integer
Private Sub Command1_Click()
    Timer1.Enabled = Not Timer1.Enabled
End Sub

Private Sub Form_Load()
    prgSign = 1
End Sub

Private Sub Timer1_Timer()
    HzxYProgressBar1.Value = HzxYProgressBar1.Value + prgSign * 1
    HzxYProgressBar2.Value = HzxYProgressBar2.Value + prgSign * 1
    HzxYProgressBar3.Value = HzxYProgressBar3.Value + prgSign * 1
    HzxYProgressBar4.Value = HzxYProgressBar4.Value + prgSign * 1
    HzxYProgressBar5.Value = HzxYProgressBar5.Value + prgSign * 1
    HzxYProgressBar6.Value = HzxYProgressBar6.Value + prgSign * 1
    HzxYProgressBar7.Value = HzxYProgressBar7.Value + prgSign * 1
    HzxYProgressBar8.Value = HzxYProgressBar8.Value + prgSign * 1
    Debug.Print HzxYProgressBar1.Value
    If HzxYProgressBar1.Value >= HzxYProgressBar1.Max Then prgSign = -1
    If HzxYProgressBar1.Value <= HzxYProgressBar1.Min Then prgSign = 1
End Sub
