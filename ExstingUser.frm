VERSION 5.00
Begin VB.Form Form3 
   Caption         =   "ExstingUser"
   ClientHeight    =   6405
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9360
   LinkTopic       =   "Form2"
   Picture         =   "ExstingUser.frx":0000
   ScaleHeight     =   6405
   ScaleWidth      =   9360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command4 
      Caption         =   "Height Converter"
      Height          =   495
      Left            =   8280
      TabIndex        =   22
      Top             =   600
      Width           =   855
   End
   Begin VB.Frame Frame2 
      Height          =   1095
      Left            =   4920
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   3975
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   2520
         TabIndex        =   17
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "HEIGHT IN METRES"
         Height          =   255
         Left            =   2280
         TabIndex        =   19
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "WEIGHT IN KG"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   720
         Width           =   1215
      End
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Update"
      Height          =   735
      Left            =   6840
      TabIndex        =   14
      Top             =   3720
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Neck"
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   1440
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
      Left            =   5280
      TabIndex        =   11
      Top             =   4440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CheckBox Check8 
      Caption         =   "Upper Back"
      Height          =   255
      Left            =   5280
      TabIndex        =   10
      Top             =   4080
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CheckBox Check7 
      Caption         =   "Legs"
      Height          =   255
      Left            =   5280
      TabIndex        =   9
      Top             =   3720
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check6 
      Caption         =   "Abdomen"
      Height          =   255
      Left            =   5280
      TabIndex        =   8
      Top             =   3360
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check5 
      Caption         =   "Triceps"
      Height          =   255
      Left            =   5280
      TabIndex        =   7
      Top             =   3000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CheckBox Check4 
      Caption         =   "Biceps"
      Height          =   375
      Left            =   5280
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Chest"
      Height          =   255
      Left            =   5280
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Shoulders"
      Height          =   255
      Left            =   5280
      TabIndex        =   4
      Top             =   1800
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Update Details"
      Height          =   495
      Left            =   2640
      TabIndex        =   3
      Top             =   240
      Width           =   2655
   End
   Begin VB.PictureBox Picture1 
      Height          =   5460
      Left            =   120
      Picture         =   "ExstingUser.frx":10983
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
   Begin VB.Label Label5 
      Height          =   255
      Left            =   8040
      TabIndex        =   21
      Top             =   240
      Width           =   1215
   End
   Begin VB.Label Label4 
      BackColor       =   &H80000002&
      Caption         =   "INJURIES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   20
      Top             =   960
      Visible         =   0   'False
      Width           =   1335
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l, c, b, a, wb As Integer
Dim fh, fb, fw As Double
Dim nam, u, bm As String
Private Sub Command1_Click()
Picture1.Visible = True
Label4.Visible = True
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
Frame2.Visible = True
End Sub

Private Sub Command2_Click()

If (Command3.Visible = True) Then
 Form1.Adodc1.Recordset.Fields("WEIGHT") = fw
 Form1.Adodc1.Recordset.Fields("HEIGHT") = fh
 Form1.Adodc1.Recordset.Fields("BMI") = fb
 Form1.Adodc1.Recordset.Fields("BASIC") = wb
 Form1.Adodc1.Recordset.Fields("LEGS") = l
 Form1.Adodc1.Recordset.Fields("CHEST") = c
 Form1.Adodc1.Recordset.Fields("BACK") = b
 Form1.Adodc1.Recordset.Fields("ARMS") = a
 Form1.Adodc1.Recordset.Update
 MsgBox "Click OK to proceed"
   End If
  Form4.Label5.Caption = Label5.Caption
  
    Form4.Label6.Caption = nam
 Form4.Show
 Form3.Hide
     
    End Sub

Private Sub Command3_Click()
If Text1.Text = "" And Text2.Text = "" Then
fh = Form1.Adodc1.Recordset.Fields("HEIGHT")
fw = Form1.Adodc1.Recordset.Fields("WEIGHT")
fb = Form1.Adodc1.Recordset.Fields("BMI")
Else
fw = Val(Text1.Text)
fh = Val(Text2.Text)
fb = Val(fw / (fh * fh))
End If
l = 1
b = 1
c = 1
a = 1


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
If Val(fb) < 18 Then
wb = 0
ElseIf Val(fb) > 25 Then
wb = 2
Else
wb = 1
End If
MsgBox "Details Updated"
End Sub

Private Sub Command4_Click()
Form5.Show

End Sub

Private Sub Form_Load()
u = Label5.Caption
Form1.Adodc1.RecordSource = "select *from UserD where UID = '" + u + "'"
If (Form1.Adodc1.Recordset.EOF = False) Then
wb = Form1.Adodc1.Recordset.Fields("BASIC")
l = Form1.Adodc1.Recordset.Fields("LEGS")
c = Form1.Adodc1.Recordset.Fields("CHEST")
b = Form1.Adodc1.Recordset.Fields("BACK")
a = Form1.Adodc1.Recordset.Fields("ARMS")
nam = Form1.Adodc1.Recordset.Fields("UNAME")
bm = Form1.Adodc1.Recordset.Fields("BMI")
End If
Frame1.Caption = nam
If Val(bm) < 18 Then
bm = "Underweight"
wb = 0
ElseIf Val(bm) > 25 Then
bm = "Overweight"
wb = 2
Else
bm = "Normal"
wb = 1
End If
Label1.Caption = bm
End Sub

