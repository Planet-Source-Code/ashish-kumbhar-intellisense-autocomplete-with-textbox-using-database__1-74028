VERSION 5.00
Begin VB.Form IntelliSense 
   BackColor       =   &H00FFFFFF&
   Caption         =   "IntelliSense With Textbox with Database"
   ClientHeight    =   3105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9060
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3105
   ScaleWidth      =   9060
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtMonths 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   2
      Top             =   720
      Width           =   2655
   End
   Begin VB.TextBox txtCountry 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   $"IntelliSense.frx":0000
      BeginProperty Font 
         Name            =   "Georgia"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808000&
      Height          =   1620
      Left            =   360
      TabIndex        =   4
      Top             =   1200
      Width           =   8295
      WordWrap        =   -1  'True
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Months"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   2520
      TabIndex        =   3
      Top             =   780
      Width           =   720
   End
   Begin VB.Label lblCountry 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   2520
      TabIndex        =   1
      Top             =   300
      Width           =   780
   End
End
Attribute VB_Name = "IntelliSense"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Call cnn
End Sub

Private Sub txtCountry_Change()
    IntelliSense txtCountry, "Country", "Name"
End Sub

Private Sub txtCountry_KeyDown(KeyCode As Integer, Shift As Integer)
    CheckKey KeyCode, txtCountry
End Sub

Private Sub txtMonths_Change()
    IntelliSense txtMonths, "Country", "Months"
End Sub

Private Sub txtMonths_KeyDown(KeyCode As Integer, Shift As Integer)
    CheckKey KeyCode, txtMonths
End Sub
