VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Height Converter"
   ClientHeight    =   4230
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   3540
   LinkTopic       =   "Form5"
   ScaleHeight     =   4230
   ScaleWidth      =   3540
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   1200
      TabIndex        =   2
      Top             =   720
      Width           =   495
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "CONVERT"
      Height          =   615
      Left            =   2160
      TabIndex        =   0
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "M"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2520
      TabIndex        =   5
      Top             =   2160
      Width           =   735
   End
   Begin VB.Label Label2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   2160
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "    FT           IN"
      Height          =   375
      Left            =   480
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim f, i, m As Double
f = Val(Text1.Text)
i = Val(Text2.Text)

i = i * 0.08333
f = f + i
m = f * 0.3048
Label2.Caption = Math.Round(m, 4)
Text1.Text = ""
Text2.Text = ""
End Sub
