VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "RW_ShowLED Property"
   ClientHeight    =   1965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   1965
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin prjRangeRoamer.MorphRangeRoamer MorphRangeRoamer1 
      Height          =   1080
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1905
      RW_GenerateEvent=   0   'False
      UD_IncrementInterval=   100
      Value           =   1
      ValueMax        =   762342
   End
   Begin prjRangeRoamer.MorphRangeRoamer MorphRangeRoamer2 
      Height          =   1080
      Left            =   3960
      TabIndex        =   1
      Top             =   240
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   1905
      RW_ShowLED      =   0   'False
      UD_IncrementInterval=   100
      Value           =   1
      ValueMax        =   762342
   End
   Begin VB.Label Label2 
      Caption         =   $"Form2.frx":0000
      Height          =   1935
      Left            =   960
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "1"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   3960
      TabIndex        =   2
      Top             =   1440
      Width           =   615
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MorphRangeRoamer2_Change()
   Label1.Caption = MorphRangeRoamer2.Value
End Sub

