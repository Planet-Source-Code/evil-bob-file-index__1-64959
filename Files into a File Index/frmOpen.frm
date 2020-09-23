VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmOpen 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Open File Index"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   6375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExtractAll 
      Caption         =   "Extract All"
      Height          =   375
      Left            =   5160
      TabIndex        =   4
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdExtract 
      Caption         =   "Extract"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "Open"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.ListBox lstFiles 
      Height          =   4350
      Left            =   1320
      TabIndex        =   0
      Top             =   120
      Width           =   3735
   End
   Begin MSComDlg.CommonDialog CMDialog1 
      Left            =   120
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
Unload Me
frmMain.Show
End Sub

Private Sub cmdExtract_Click()
CMDialog1.Filter = "All Files (*.*)|*.*|"
CMDialog1.DialogTitle = "Create File Index"
CMDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
On Error GoTo No_Save
CMDialog1.ShowSave

FI.GetFileFromFileIndex lstFiles.ListIndex * 2, CMDialog1.FileName, 1000

No_Save:
Exit Sub
Unload Me
End Sub

Private Sub cmdExtractAll_Click()
Dim Path1 As String
Path1 = BrowseFolder(Me.hwnd, "Select Folder To Extract Files")

If Path1 = "" Then Exit Sub

For i = 0 To lstFiles.ListCount - 1
    FI.GetFileFromFileIndex i * 2, Path1 & "\" & lstFiles.List(i), 1000
    DoEvents
Next i

MsgBox "Extracted!", vbInformation, "w00t"
End Sub

Private Sub cmdOpen_Click()
Dim Count As Long, Current As String

CMDialog1.Filter = "File Index (*.fi)|*.fi|"
CMDialog1.DialogTitle = "Create File Index"
CMDialog1.Flags = cdlOFNOverwritePrompt + cdlOFNPathMustExist
On Error GoTo No_Open
CMDialog1.ShowOpen

FI.CloseFileIndex

FI.OpenFileIndex CMDialog1.FileName

Count = FI.GetFileIndexCount()

    For i = 1 To Count Step 2
        If Count <= i Then GoTo No_Open
        
        Current = FI.GetFromFileIndex(i)
        lstFiles.AddItem Current
        DoEvents
    Next i

No_Open:
CMDialog1.FileName = ""
Exit Sub
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMain.Show
End Sub
