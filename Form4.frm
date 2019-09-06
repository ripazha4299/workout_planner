VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form Form4 
   Caption         =   "Workout"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14475
   LinkTopic       =   "Form4"
   Picture         =   "Form4.frx":0000
   ScaleHeight     =   7995
   ScaleWidth      =   14475
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "User NAME"
      Height          =   615
      Left            =   2280
      TabIndex        =   7
      Top             =   120
      Width           =   3015
      Begin VB.Label Label6 
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "USER ID"
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   2295
      Begin VB.Label Label5 
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Arial Rounded MT Bold"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4350
      Left            =   7680
      TabIndex        =   4
      Top             =   1440
      Width           =   4815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Exercises"
      Height          =   855
      Left            =   6000
      TabIndex        =   3
      Top             =   1680
      Width           =   1335
   End
   Begin MSAdodcLib.Adodc exe 
      Height          =   375
      Left            =   9120
      Top             =   240
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   2
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
      Connect         =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\antik\Desktop\Workout Planner\workoutchart.mdb;Persist Security Info=False"
      OLEDBString     =   "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\Users\antik\Desktop\Workout Planner\workoutchart.mdb;Persist Security Info=False"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "workout"
      Caption         =   "Exercises"
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
   Begin VB.Label Label7 
      Caption         =   "initial"
      Height          =   975
      Left            =   1080
      TabIndex        =   10
      Top             =   1320
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      DataField       =   "wName"
      DataSource      =   "exe"
      Height          =   135
      Left            =   7560
      TabIndex        =   9
      Top             =   1200
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "set"
      Height          =   495
      Left            =   9720
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "Reps"
      Height          =   375
      Left            =   9120
      TabIndex        =   1
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "Exercise"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   7800
      TabIndex        =   0
      Top             =   960
      Width           =   1335
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim l, c, b, a, wb, count1 As Integer
Dim u, fl, fb, fc, fa, s(3) As String


Private Sub Command2_Click()
u = Label5.Caption
count1 = 0
Form1.Adodc1.RecordSource = "select *from UserD where UID = '" + u + "'"

If (Form1.Adodc1.Recordset.EOF = False) Then
wb = Form1.Adodc1.Recordset.Fields("BASIC")
l = Form1.Adodc1.Recordset.Fields("LEGS")
c = Form1.Adodc1.Recordset.Fields("CHEST")
b = Form1.Adodc1.Recordset.Fields("BACK")
a = Form1.Adodc1.Recordset.Fields("ARMS")

If l = 0 Then
s(count1) = "l"
count1 = count1 + 1
End If

If a = 0 Then
s(count1) = "a"
count1 = count1 + 1
End If

If b = 0 Then
s(count1) = "b"
count1 = count1 + 1
End If

If c = 0 Then
s(count1) = "c"
count1 = count1 + 1
End If

End If



While exe.Recordset.EOF = False

If CStr(exe.Recordset.Fields("chart")) = CStr(wb) Then

  If (count1 = 0) Then
List1.AddItem (exe.Recordset.Fields("wName") & "  " & exe.Recordset.Fields("reps") & "  " & exe.Recordset.Fields("set"))

ElseIf count1 = 1 Then
    If exe.Recordset.Fields(s(0)) = 0 Then
    List1.AddItem (exe.Recordset.Fields("wName") & "  " & exe.Recordset.Fields("reps") & "  " & exe.Recordset.Fields("set"))
    End If
    
ElseIf count1 = 2 Then
    If ((exe.Recordset.Fields(s(0)) = 0) And (exe.Recordset.Fields(s(1)) = 0)) Then
    List1.AddItem (exe.Recordset.Fields("wName") & "  " & exe.Recordset.Fields("reps") & "  " & exe.Recordset.Fields("set"))
    End If

ElseIf count1 = 3 Then
    If ((exe.Recordset.Fields(s(0)) = 0) And (exe.Recordset.Fields(s(1)) = 0) And (exe.Recordset.Fields(s(2)) = 0)) Then
    List1.AddItem (exe.Recordset.Fields("wName") & "  " & exe.Recordset.Fields("reps") & "  " & exe.Recordset.Fields("set"))
        End If
Else
List1.AddItem (" NO WORKOUT POSSIBLE.")
List1.AddItem ("Wish you a speedy recovery ")
GoTo x
End If

End If
exe.Recordset.MoveNext
Wend
x:
Label7.Caption = u

Command2.Visible = False

End Sub
