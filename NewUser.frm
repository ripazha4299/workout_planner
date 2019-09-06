VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "NewUser"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9450
   LinkTopic       =   "Form2"
   Picture         =   "NewUser.frx":0000
   ScaleHeight     =   6255
   ScaleWidth      =   9450
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Height Converter"
      Height          =   495
      Left            =   8040
      TabIndex        =   17
      Top             =   840
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "UserID"
      Height          =   615
      Left            =   8040
      TabIndex        =   15
      Top             =   120
      Width           =   975
      Begin VB.Label Label2 
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   495
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update injuries"
      Height          =   735
      Left            =   7560
      TabIndex        =   14
      Top             =   4320
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Neck"
      Height          =   255
      Left            =   6240
      TabIndex        =   13
      Top             =   2520
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Proceed"
      Height          =   495
      Left            =   6360
      TabIndex        =   12
      Top             =   240
      Width           =   1575
   End
   Begin VB.CheckBox Check9 
      Caption         =   "Lower Back"
      Height          =   255
      Left            =   6240
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Upper Back"
      Height          =   255
      Left            =   6240
      TabIndex        =   10
      Top             =   5160
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Legs"
      Height          =   255
      Left            =   6240
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Abdomen"
      Height          =   255
      Left            =   6240
      TabIndex        =   8
      Top             =   4440
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Triceps"
      Height          =   255
      Left            =   6240
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Biceps"
      Height          =   375
      Left            =   6240
      TabIndex        =   6
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Chest"
      Height          =   255
      Left            =   6240
      TabIndex        =   5
      Top             =   3240
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Shoulders"
      Height          =   255
      Left            =   6240
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select INJURIES(if any)"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   5460
      Left            =   120
      Picture         =   "NewUser.frx":101A2
      ScaleHeight     =   5400
      ScaleWidth      =   4440
      TabIndex        =   2
      Top             =   840
      Visible         =   0   'False
      Width           =   4500
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   615
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.Label Label1 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l, c, b, a, wb As Integer

Dim nam, u, bm As String
Private Sub Command1_Click()
Picture1.Visible = True
Check1.Visible = True
Check2.Visible = True
Check3.Visible = True
Check4.Visible = True
Check5.Visible = True
Check6.Visible = True
Check7.Visible = True
Check8.Visible = True
Check9.Visible = True
Command3.Visible = True

End Sub

Private Sub Command2_Click()
    Form1.Adodc1.Recordset.Fields("BASIC") = wb
    Form1.Adodc1.Recordset.Fields("LEGS") = l
    Form1.Adodc1.Recordset.Fields("CHEST") = c
    Form1.Adodc1.Recordset.Fields("BACK") = b
    Form1.Adodc1.Recordset.Fields("ARMS") = a
    Form1.Adodc1.Recordset.Update
    MsgBox "Injuries Updated"
    Form4.Label5.Caption = Label2.Caption
    Form4.Label6.Caption = nam

    Form4.Show
    Form2.Hide
    
    End Sub

Private Sub Command3_Click()


If Check7.Value = 1 Or Check9.Value = 1 Then
l = 0
End If
If Check2.Value = 1 Or Check3.Value = 1 Or Check4.Value = 1 Or Check5.Value = 1 Or Check6.Value = 1 Or Check8.Value = 1 Then
c = 0
End If
If Check1.Value = 1 Or Check2.Value = 1 Or Check3.Value = 1 Or Check4.Value = 1 Or Check5.Value = 1 Or Check6.Value = 1 Or Check8.Value = 1 Or Check9.Value = 1 Then
b = 0
End If
If Check2.Value = 1 Or Check4.Value = 1 Or Check5.Value = 1 Then
a = 0

End If

End Sub

Private Sub Command4_Click()
Form5.Show

End Sub

Private Sub Form_Load()
l = 1
b = 1
c = 1
a = 1
wb = 1
nam = Form1.Text1.Text
u = Form1.Text2.Text
Label2.Caption = u
bm = Form1.Label5.Caption
Frame1.Caption = nam
If Val(bm) < 18 Then
bm = "Underweight"
wb = 0
ElseIf Val(bm) > 25 Then
bm = "Overweight"
wb = 2
Else
bm = "Normal"
End If
Label1.Caption = bm
End Sub

