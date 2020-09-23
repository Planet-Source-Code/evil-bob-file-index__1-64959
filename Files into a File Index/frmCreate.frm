VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmCreate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Create File Index"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   7695
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   6
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   6360
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.ListBox lstFiles 
      Height          =   4155
      Left            =   1440
      TabIndex        =   4
      Top             =   120
      Width           =   4815
   End
   Begin VB.CommandButton cmdRemoveAll 
      Caption         =   "Remove All"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdRemove 
      Caption         =   "Remove"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton cmdAddFolder 
      Caption         =   "Add Folder"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add File"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   240
      Top             =   3360
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdAdd_Click()
On Error GoTo No_Save

CMDialog1.Filter = "All Files (*.*)|*.*|"
CMDialog1.DialogTitle = "Add File"
CMDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
CMDialog1.ShowOpen

If CMDialog1.FileName = "" Then Exit Sub

lstFiles.AddItem CMDialog1.FileName

CMDialog1.FileName = ""

No_Save:
Exit Sub
End Sub

Private Sub cmdAddFolder_Click()
Dim Path1 As String
Path1 = BrowseFolder(Me.hwnd, "Select Folder Containing Files")

If Path1 = "" Then Exit Sub

GetFiles Path1, lstFiles
End Sub

Private Sub cmdCancel_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdCreate_Click()
CMDialog1.Filter = "File Index (*.fi)|*.fi|"
CMDialog1.DialogTitle = "Create File Index"
CMDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
On Error GoTo No_Save
CMDialog1.ShowSave

FI.OpenFileIndex CMDialog1.FileName
    For i = 0 To lstFiles.ListCount - 1
        FI.AddFileToFileIndex lstFiles.List(i), 1000
        DoEvents
        FI.AddToFileIndex FixFileFormat(lstFiles.List(i))
        DoEvents
    Next i
FI.CloseFileIndex

MsgBox "File Index created!", vbInformation, "Yaaay..."

No_Save:
Exit Sub
Unload Me
End Sub

Private Sub cmdRemove_Click()
On Error Resume Next
lstFiles.RemoveItem lstFiles.ListIndex
End Sub

Private Sub cmdRemoveAll_Click()
lstFiles.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub
