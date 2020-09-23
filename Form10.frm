VERSION 5.00
Begin VB.Form Form10 
   BackColor       =   &H8000000A&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "search"
   ClientHeight    =   1305
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3930
   Icon            =   "FORM10.frx":0000
   LinkTopic       =   "Form10"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1305
   ScaleWidth      =   3930
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "search"
      Height          =   855
      Left            =   3000
      TabIndex        =   3
      Top             =   240
      Width           =   735
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.ComboBox Combo1 
      BackColor       =   &H80000006&
      ForeColor       =   &H0000FF00&
      Height          =   315
      ItemData        =   "FORM10.frx":1782
      Left            =   600
      List            =   "FORM10.frx":179B
      TabIndex        =   0
      Top             =   240
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "type"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   240
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "name"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
End
Attribute VB_Name = "Form10"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
On Error Resume Next
Unload Form9
GoTo d
d:
Call Load(Form9)
Form9.Show
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Unload Form9
End Sub
