VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Restore Last Deleted Record"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6555
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command5 
      Caption         =   "Blank"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      TabIndex        =   17
      Top             =   2760
      Width           =   1215
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1980
      Left            =   360
      TabIndex        =   16
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   15
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Restore Deleted"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   14
      Top             =   2760
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   13
      Top             =   2760
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1680
      TabIndex        =   12
      Top             =   2760
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   11
      Top             =   2280
      Width           =   2775
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   10
      Top             =   1920
      Width           =   2775
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   9
      Top             =   1560
      Width           =   2775
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   8
      Top             =   1200
      Width           =   2775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Top             =   840
      Width           =   2775
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   3360
      TabIndex        =   6
      Top             =   480
      Width           =   2775
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   2280
      TabIndex        =   5
      Top             =   2280
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mob"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   2280
      TabIndex        =   3
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   2280
      TabIndex        =   2
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2280
      TabIndex        =   1
      Top             =   840
      Width           =   1935
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "First Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2280
      TabIndex        =   0
      Top             =   480
      Width           =   1935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'developer: Shyam Singh Chandel < shyamschandel@rediffmail.com >

Private Sub Command1_Click()
SQL = "SELECT * FROM info"
Set RS = New ADODB.Recordset
 RS.Open SQL, CN, adOpenStatic, adLockOptimistic
  RS.AddNew
    RS!FName = Text1.Text
    RS!sName = Text2.Text
    RS!Address = Text3.Text
    RS!mob = Text4.Text
    RS!email = Text5.Text
    RS!city = Text6.Text
  RS.Update
  MsgBox "Record has been saved"
  Command5_Click
   RS.Close
   List1.Clear
   load
End Sub

Private Sub Command2_Click()
SQL = "SELECT * FROM info where Fname='" & List1.Text & "'"
Set RS = New ADODB.Recordset
RS.Open SQL, CN, adOpenStatic, adLockOptimistic
 Backup
RS.Delete
MsgBox "Record has been deleted"
Command5_Click
RS.Close
List1.Clear
load
End Sub

Private Sub Command3_Click()

SQL = "SELECT * FROM info where Fname='" & List1.Text & "'"
Set RS = New ADODB.Recordset
RS.Open SQL, CN, adOpenStatic, adLockOptimistic
RestoreBackup
If Text1.Text = "" Then
MsgBox "NO RECORD FOR RESTORE"
Exit Sub
End If
RS.AddNew
    RS!FName = Text1.Text
    RS!sName = Text2.Text
    RS!Address = Text3.Text
    RS!mob = Text4.Text
    RS!email = Text5.Text
    RS!city = Text6.Text
  RS.Update
  MsgBox "Record has been Restored"
  Command5_Click
  RS.Close
  List1.Clear
  load
  DeleteBackup
End Sub

Private Sub Command4_Click()
End
End Sub

Private Sub Command5_Click()
BlankFields
End Sub

Private Sub Form_Load()
connectDB
load
End Sub
Private Sub load()
SQL = "SELECT * FROM info"
Set RS = New ADODB.Recordset
 RS.Open SQL, CN, adOpenStatic, adLockOptimistic
   Do While Not RS.EOF
      List1.AddItem RS!FName
    RS.MoveNext
    Loop
  RS.Close
End Sub

Private Sub List1_Click()
SQL = "SELECT * FROM info where Fname='" & List1.Text & "'"
Set RS = New ADODB.Recordset
 RS.Open SQL, CN, adOpenStatic, adLockOptimistic
    AddFields
 RS.Close
End Sub
Private Sub DeleteBackup()
Call SaveSetting("Restore", "Restore Data", "Fname", "")
Call SaveSetting("Restore", "Restore Data", "Sname", "")
Call SaveSetting("Restore", "Restore Data", "Address", "")
Call SaveSetting("Restore", "Restore Data", "Mob", "")
Call SaveSetting("Restore", "Restore Data", "Email", "")
Call SaveSetting("Restore", "Restore Data", "city", "")
End Sub
Sub Backup()
Call SaveSetting("Restore", "Restore Data", "Fname", Text1.Text)
Call SaveSetting("Restore", "Restore Data", "Sname", Text2.Text)
Call SaveSetting("Restore", "Restore Data", "Address", Text3.Text)
Call SaveSetting("Restore", "Restore Data", "Mob", Text4.Text)
Call SaveSetting("Restore", "Restore Data", "Email", Text5.Text)
Call SaveSetting("Restore", "Restore Data", "city", Text6.Text)
End Sub
Sub RestoreBackup()
Text1.Text = GetSetting("Restore", "Restore Data", "Fname")
Text2.Text = GetSetting("Restore", "Restore Data", "Sname")
Text3.Text = GetSetting("Restore", "Restore Data", "Address")
Text4.Text = GetSetting("Restore", "Restore Data", "Mob")
Text5.Text = GetSetting("Restore", "Restore Data", "Email")
Text6.Text = GetSetting("Restore", "Restore Data", "city")
End Sub
Sub AddFields()
    Text1 = RS!FName
    Text2 = RS!sName
    Text3 = RS!Address
    Text4 = RS!mob
    Text5 = RS!email
    Text6 = RS!city
End Sub
Sub BlankFields()
  Text1 = ""
  Text2 = ""
  Text3 = ""
  Text4 = ""
  Text5 = ""
  Text6 = ""
End Sub
