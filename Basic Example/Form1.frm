VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Basic File Indexer Example"
   ClientHeight    =   1095
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1095
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "Get Data"
      Height          =   855
      Left            =   2280
      TabIndex        =   1
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Put Data"
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim FI As New clsFileIndexer

Private Sub Command1_Click()
FI.OpenFileIndex App.Path & "\test.fi"
    FI.AddToFileIndex "Hello World! This is string #0!"
    FI.AddToFileIndex "This is string #1! *gasp*"
    FI.AddToFileIndex "If you cant count, this is #2! hehehe"
FI.CloseFileIndex
End Sub

Private Sub Command2_Click()
FI.OpenFileIndex App.Path & "\test.fi"
    MsgBox FI.GetFromFileIndex(0)
    MsgBox FI.GetFromFileIndex(1)
    MsgBox FI.GetFromFileIndex(2)
FI.CloseFileIndex
End Sub
