VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MorphRangeRoamer Demo - Matthew R. Usner"
   ClientHeight    =   4740
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   ScaleHeight     =   4740
   ScaleWidth      =   8760
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Caption         =   "Color Schemes"
      Height          =   2775
      Left            =   1680
      TabIndex        =   13
      Top             =   120
      Width           =   3015
      Begin VB.CheckBox Check2 
         Caption         =   "RW_GenerateEvent"
         Height          =   255
         Left            =   360
         TabIndex        =   29
         Top             =   2280
         Width           =   1815
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Purple People Eater"
         Height          =   255
         Left            =   1080
         TabIndex        =   22
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Blue Moon"
         Height          =   255
         Left            =   1080
         TabIndex        =   21
         Top             =   1680
         Width           =   1575
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Cyan Eyed"
         Height          =   255
         Left            =   1080
         TabIndex        =   20
         Top             =   1440
         Width           =   1575
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Penny Wise"
         Height          =   255
         Left            =   1080
         TabIndex        =   19
         Top             =   1200
         Width           =   1575
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Red Rum"
         Height          =   255
         Left            =   1080
         TabIndex        =   18
         Top             =   960
         Width           =   1575
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Golden Goose"
         Height          =   255
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Green With Envy"
         Height          =   255
         Left            =   1080
         TabIndex        =   16
         Top             =   480
         Width           =   1575
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Gunmetal Grey"
         Height          =   255
         Left            =   1080
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1575
      End
      Begin prjRangeRoamer.MorphRangeRoamer MorphRangeRoamer1 
         Height          =   1080
         Left            =   240
         TabIndex        =   14
         Top             =   600
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   1905
         RW_GenerateEvent=   0   'False
         UD_IncrementInterval=   100
         Value           =   1
         ValueMax        =   762342
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "RW_GenerateEvent"
      Height          =   4575
      Left            =   4800
      TabIndex        =   3
      Top             =   120
      Width           =   3855
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   120
         ScaleHeight     =   1665
         ScaleWidth      =   1095
         TabIndex        =   4
         Top             =   240
         Width           =   1125
         Begin VB.PictureBox picRed 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            ScaleHeight     =   345
            ScaleWidth      =   345
            TabIndex        =   8
            Top             =   0
            Width           =   375
         End
         Begin VB.PictureBox picGreen 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   360
            ScaleHeight     =   345
            ScaleWidth      =   345
            TabIndex        =   7
            Top             =   0
            Width           =   375
         End
         Begin VB.PictureBox picBlue 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   720
            ScaleHeight     =   345
            ScaleWidth      =   345
            TabIndex        =   6
            Top             =   0
            Width           =   375
         End
         Begin VB.PictureBox picRGB 
            Appearance      =   0  'Flat
            BackColor       =   &H00000000&
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   0
            ScaleHeight     =   345
            ScaleWidth      =   1065
            TabIndex        =   5
            Top             =   1080
            Width           =   1095
         End
         Begin prjRangeRoamer.MorphRangeRoamer mrrBlue 
            Height          =   720
            Left            =   720
            TabIndex        =   9
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   1270
            RW_BackColor1   =   4194304
            RW_BackColor2   =   12582912
            RW_BorderColor1 =   4194304
            RW_BorderColor2 =   16744576
            RW_LED_BurnInColor=   8388608
            RW_LED_DigitColor=   16744576
            RW_PBarColor1   =   4194304
            RW_PBarColor2   =   16744576
            Theme           =   3
            UD_ArrowColor   =   16761024
            UD_BorderColor1 =   8388608
            UD_BorderColor2 =   16744576
            UD_BorderWidth  =   4
            UD_ButtonColor1 =   4194304
            UD_ButtonColor2 =   16744576
            UD_FocusBorderColor1=   4194304
            UD_FocusBorderColor2=   16711680
            UD_IncrementInterval=   100
            ValueMax        =   255
            ValueMin        =   0
            Wrap            =   -1  'True
         End
         Begin prjRangeRoamer.MorphRangeRoamer mrrGreen 
            Height          =   720
            Left            =   360
            TabIndex        =   10
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   1270
            RW_BackColor1   =   16384
            RW_BackColor2   =   40960
            RW_BorderColor1 =   16384
            RW_BorderColor2 =   8454016
            RW_LED_BurnInColor=   32768
            RW_LED_DigitColor=   8454016
            RW_PBarColor1   =   16384
            RW_PBarColor2   =   8454016
            Theme           =   5
            UD_ArrowColor   =   12648384
            UD_BorderColor1 =   16384
            UD_BorderColor2 =   8454016
            UD_BorderWidth  =   4
            UD_ButtonColor1 =   16384
            UD_ButtonColor2 =   65280
            UD_FocusBorderColor1=   16384
            UD_FocusBorderColor2=   49152
            UD_IncrementInterval=   100
            ValueMax        =   255
            ValueMin        =   0
         End
         Begin prjRangeRoamer.MorphRangeRoamer mrrRed 
            Height          =   720
            Left            =   0
            TabIndex        =   11
            Top             =   360
            Width           =   375
            _ExtentX        =   661
            _ExtentY        =   1270
            RW_BackColor1   =   64
            RW_BackColor2   =   192
            RW_BorderColor1 =   64
            RW_BorderColor2 =   8421631
            RW_LED_BurnInColor=   128
            RW_LED_DigitColor=   8421631
            RW_PBarColor1   =   64
            RW_PBarColor2   =   8421631
            Theme           =   4
            UD_ArrowColor   =   12632319
            UD_BorderColor1 =   64
            UD_BorderColor2 =   8421631
            UD_BorderWidth  =   4
            UD_ButtonColor1 =   64
            UD_ButtonColor2 =   8421631
            UD_FocusBorderColor1=   64
            UD_FocusBorderColor2=   4210943
            UD_IncrementInterval=   100
            ValueMax        =   255
            ValueMin        =   0
         End
         Begin VB.Label lblHex 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "H0"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            TabIndex        =   12
            Top             =   1440
            Width           =   1095
         End
      End
      Begin VB.Label Label3 
         Caption         =   $"Form1.frx":0000
         Height          =   2415
         Left            =   120
         TabIndex        =   24
         Top             =   1995
         Width           =   3615
      End
      Begin VB.Label Label2 
         Caption         =   $"Form1.frx":0227
         Height          =   1815
         Left            =   1440
         TabIndex        =   23
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Enabled Demo"
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1455
      Begin prjRangeRoamer.MorphRangeRoamer MorphRangeRoamer3 
         Height          =   855
         Left            =   480
         TabIndex        =   27
         Top             =   960
         Width           =   495
         _ExtentX        =   873
         _ExtentY        =   1508
         UD_BorderWidth  =   6
         ValueMax        =   10000
         ValueMin        =   -10000
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   720
         Value           =   1  'Checked
         Width           =   975
      End
      Begin prjRangeRoamer.MorphRangeRoamer MorphRangeRoamer2 
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Width           =   930
         _ExtentX        =   1640
         _ExtentY        =   873
         UD_BorderWidth  =   6
         UD_ButtonDownAngle=   180
         UD_ButtonUpAngle=   180
         UD_Orientation  =   1
         ValueIncrCtrl   =   1
         ValueIncrShift  =   5
         ValueIncrShiftCtrl=   10
         ValueMax        =   3
         Wrap            =   -1  'True
      End
      Begin VB.Label Label5 
         Caption         =   "Both Horizontal and Vertical UpDown layouts are supported."
         Height          =   870
         Left            =   120
         TabIndex        =   28
         Top             =   1800
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      Caption         =   $"Form1.frx":031F
      Height          =   1455
      Left            =   120
      TabIndex        =   26
      Top             =   3120
      Width           =   4695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Check1_Click()
   If Check1.Value = vbChecked Then
      MorphRangeRoamer2.Enabled = True
      MorphRangeRoamer3.Enabled = True
   Else
      MorphRangeRoamer2.Enabled = False
      MorphRangeRoamer3.Enabled = False
   End If
End Sub

Private Sub Check2_Click()
   If Check2.Value = vbChecked Then
      MorphRangeRoamer1.RW_GenerateEvent = True
   Else
      MorphRangeRoamer1.RW_GenerateEvent = False
   End If
End Sub

Private Sub Form_Load()
   Form2.Show
End Sub

Private Sub MorphRangeRoamer1_Change()
   Label1.Caption = MorphRangeRoamer1.Value
End Sub

' color picker code
Private Sub mrrRed_Change()
   picRed.BackColor = mrrRed.Value
   picRGB.BackColor = RGB(mrrRed.Value, mrrGreen.Value, mrrBlue.Value)
   lblHex.Caption = "H" & Hex(picRGB.BackColor)
End Sub
Private Sub mrrGreen_Change()
   picGreen.BackColor = 256 * mrrGreen.Value
   picRGB.BackColor = RGB(mrrRed.Value, mrrGreen.Value, mrrBlue.Value)
   lblHex.Caption = "H" & Hex(picRGB.BackColor)
End Sub
Private Sub mrrBlue_Change()
   picBlue.BackColor = 65536 * mrrBlue.Value
   picRGB.BackColor = RGB(mrrRed.Value, mrrGreen.Value, mrrBlue.Value)
   lblHex.Caption = "H" & Hex(picRGB.BackColor)
End Sub

' themes
Private Sub Option1_Click()
   MorphRangeRoamer1.Theme = [Gunmetal Grey]
End Sub

Private Sub Option2_Click()
   MorphRangeRoamer1.Theme = [Green With Envy]
End Sub

Private Sub Option3_Click()
   MorphRangeRoamer1.Theme = [Golden Goose]
End Sub

Private Sub Option4_Click()
   MorphRangeRoamer1.Theme = [Red Rum]
End Sub

Private Sub Option5_Click()
   MorphRangeRoamer1.Theme = [Penny Wise]

End Sub

Private Sub Option6_Click()
   MorphRangeRoamer1.Theme = [Cyan Eyed]
End Sub

Private Sub Option7_Click()
   MorphRangeRoamer1.Theme = [Blue Moon]
End Sub

Private Sub Option8_Click()
   MorphRangeRoamer1.Theme = [Purple People Eater]
End Sub
