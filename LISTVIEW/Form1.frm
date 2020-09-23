VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "Listview Basics By Garz0r7"
   ClientHeight    =   9000
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   9000
   ScaleWidth      =   8400
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command5 
      Caption         =   "Clear List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   21
      Top             =   8280
      Width           =   3975
   End
   Begin VB.TextBox rmv 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   525
      Left            =   6480
      TabIndex        =   17
      Text            =   "2"
      Top             =   7680
      Width           =   1815
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Remove A Row"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   16
      Top             =   7680
      Width           =   3975
   End
   Begin VB.TextBox col 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5880
      TabIndex        =   15
      Text            =   "1"
      Top             =   7080
      Width           =   495
   End
   Begin VB.TextBox rrow 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      TabIndex        =   14
      Text            =   "1"
      Top             =   7080
      Width           =   495
   End
   Begin VB.TextBox txt 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   11
      Text            =   "Your Text"
      Top             =   7080
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Add into row , column a text "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   7080
      Width           =   3975
   End
   Begin VB.TextBox newtext 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   9
      Text            =   "Add Text"
      Top             =   6480
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add something into the first column"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   7
      Top             =   6480
      Width           =   3975
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   7560
      Top             =   8280
   End
   Begin VB.TextBox txtremove 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   5
      Text            =   "1"
      Top             =   5880
      Width           =   1815
   End
   Begin VB.CommandButton Remo 
      Caption         =   "Remove A column"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   5880
      Width           =   3975
   End
   Begin VB.TextBox colname 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      TabIndex        =   3
      Text            =   "col1"
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add New Header(Add New Column)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   5280
      Width           =   3975
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   8705
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label items 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6480
      TabIndex        =   20
      Top             =   8280
      Width           =   855
   End
   Begin VB.Label Label6 
      Caption         =   "Items Number:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   19
      Top             =   8280
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Row's index to remove:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4200
      TabIndex        =   18
      Top             =   7680
      Width           =   2175
   End
   Begin VB.Label Label4 
      Caption         =   "Column:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5280
      TabIndex        =   13
      Top             =   7080
      Width           =   735
   End
   Begin VB.Label Label3 
      Caption         =   "Row:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4200
      TabIndex        =   12
      Top             =   7080
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Give A Text :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   8
      Top             =   6480
      Width           =   2055
   End
   Begin VB.Label lab 
      Caption         =   "Give Column's Index :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4320
      TabIndex        =   6
      Top             =   5880
      Width           =   2055
   End
   Begin VB.Label Label1 
      Caption         =   "Give Columns Name :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   161
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   2
      Top             =   5280
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

'ADD NEW COLUMN

ListView1.ColumnHeaders.Add , , colname

End Sub

Private Sub Command2_Click()
ListView1.ListItems.Add , , newtext
End Sub

Private Sub Command3_Click()
On Error GoTo er

ListView1.ListItems(CInt(rrow)).SubItems(CInt(col)) = txt

er:

'handles bad number for row-column
End Sub

Private Sub Command4_Click()
On Error GoTo er

ListView1.ListItems.Remove (CInt(rmv))

er:

'handles bad number for row
End Sub

Private Sub Command5_Click()

'Clear the list

ListView1.ListItems.Clear

End Sub

Private Sub Form_Load()

' Set the View

ListView1.View = lvwReport

End Sub

Private Sub Form_Unload(Cancel As Integer)
MsgBox "DON'T FORGET TO VOTE !", , "Garz0r7"

End Sub

Private Sub Timer1_Timer()

'for the column remaning

colname = "col" + Str(ListView1.ColumnHeaders.Count + 1)

items = ListView1.ListItems.Count

End Sub

Private Sub Remo_Click()
On Error GoTo er

ListView1.ColumnHeaders.Remove (CInt(txtremove))

er:
'handles if we select a column number
'which not exists.
End Sub
