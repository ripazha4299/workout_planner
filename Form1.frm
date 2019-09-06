VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form1 
   Caption         =   "Login"
   ClientHeight    =   4980
   ClientLeft      =   3735
   ClientTop       =   2490
   ClientWidth     =   7440
   LinkTopic       =   "Form1"
   Picture         =   "Form1.frx":0000
   ScaleHeight     =   4980
   ScaleWidth      =   7440
   Begin VB.CommandButton Command4 
      Caption         =   "Height Converter"
      Height          =   615
      Left            =   4320
      TabIndex        =   16
      Top             =   3840
      Width           =   1215
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   840
      Top             =   3960
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   1
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   $"Form1.frx":C929
      OLEDBString     =   $"Form1.frx":C9C8
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "select *from UserD "
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H8000000C&
      Caption         =   "Enter User ID"
      Height          =   2535
      Left            =   3720
      TabIndex        =   1
      Top             =   1080
      Width           =   3135
      Begin VB.CommandButton Command3 
         Caption         =   "OK"
         Height          =   375
         Left            =   1800
         TabIndex        =   14
         Top             =   360
         Width           =   615
      End
      Begin VB.TextBox Text5 
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label6 
         Height          =   855
         Left            =   240
         TabIndex        =   15
         Top             =   960
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000016&
      Caption         =   "New User"
      Height          =   3615
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   3015
      Begin VB.CommandButton Command2 
         Caption         =   "Save"
         Height          =   495
         Left            =   720
         TabIndex        =   11
         Top             =   2760
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "BMI"
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   1920
         Width           =   615
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   1080
         TabIndex        =   9
         Top             =   1440
         Width           =   615
      End
      Begin VB.TextBox Text3 
         Height          =   285
         Left            =   1080
         TabIndex        =   8
         Top             =   1080
         Width           =   615
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   1080
         TabIndex        =   7
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1080
         TabIndex        =   6
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label5 
         Height          =   375
         Left            =   1200
         TabIndex        =   12
         Top             =   1920
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Weight (kg)"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Height (m)"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Unique ID"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim h, b, w As Double
h = Val(Text3.Text)
w = Val(Text4.Text)
b = w / (h * h)
Label5.Caption = Math.Round(b, 4)
Label5.Visible = True

End Sub

Private Sub Command2_Click()
Dim h, b, w As Double
Dim c As String
c = Text2.Text
h = Val(Text3.Text)
w = Val(Text4.Text)
b = w / (h * h)
Label5.Caption = Math.Round(b, 4)
    Adodc1.Refresh
    Adodc1.Recordset.AddNew
    Adodc1.Recordset.Fields("UID") = c
    Adodc1.Recordset.Fields("UNAME") = Text1.Text
    Adodc1.Recordset.Fields("HEIGHT") = h
    Adodc1.Recordset.Fields("WEIGHT") = w
    Adodc1.Recordset.Fields("BMI") = Label5.Caption
        MsgBox "Data Added"
         Form2.Show
    Form1.Hide

End Sub


Private Sub Command3_Click()

Adodc1.RecordSource = "select *from UserD where UID = '" + Text5.Text + "'"
Adodc1.Refresh

If (Adodc1.Recordset.EOF = False) Then
   Label6.Caption = Adodc1.Recordset.Fields("UNAME") & "   BMI=" & Adodc1.Recordset.Fields("BMI")
   MsgBox "Welcome"
        Form3.Label5.Caption = Text5.Text
      Form3.Show
    Form1.Hide
Else
    MsgBox "Invalid UserID!", vbCritical, "Logi"
    Text5.Text = ""
    Label6.Caption = ""
    End If
    
 
    
End Sub

Private Sub Command4_Click()
Form5.Show
End Sub
