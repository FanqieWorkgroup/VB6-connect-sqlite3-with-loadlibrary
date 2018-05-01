VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.TextBox Text2 
      Height          =   855
      Left            =   240
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   1455
      Left            =   600
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   0
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1215
      Left            =   3120
      TabIndex        =   0
      Top             =   1080
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long


Public odb As Object
Private Sub Form_Load()
Dim hdll As Long
hdll = LoadLibrary("C:\Users\835078903@qq.com\Desktop\sss\sqlite3.dll")
 Text2.Text = hdll
    Set odb = CreateObject("LiteX.LiteConnection")
    Text1.Text = odb.Version
    odb.Open ("E:\sss\virnuscenter.db")
End Sub

Private Sub command1_Click()
   Set rs = odb.Prepare("select * from md5 where words like '%8d95%'")
 rs.Step
   Print rs.RowCount
End Sub

