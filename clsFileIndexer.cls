VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFileIndexer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

'Title Name:      File Indexer
'Version:         1.0
'Created Date:    4/6/2006
'Last Update:     4/6/2006
'Author:          Skyler Lyon
'Description:     Forms an Index file containing multiple pieces of data,
'                 like an array within a file. And it allows
'                 fast and memory efficient file input into the index.

'                 The Header is composed of two things: the Datalength and the
'                 TermChar. The Datalength is a number that determines how long
'                 the piece of data is; and the TermChar is the character that
'                 tells it that the data begins here.

Private Opened As Long
Private CurrIndex As Long

Public Function OpenFileIndex(FileName As String) As Boolean
On Error GoTo NotOpenable

Opened = FreeFile

Open FileName For Binary As Opened
CurrIndex = 0
OpenFileIndex = True
Exit Function

NotOpenable:
OpenFileIndex = False
End Function

Public Function CloseFileIndex() As Boolean
On Error GoTo NotClosable

Close Opened
CloseFileIndex = True
Exit Function

NotClosable:
CloseFileIndex = False
End Function

Public Function AddToFileIndex(Data As String)
Put Opened, , CStr(Len(Data)) & "F"
Put Opened, , Data
End Function

Public Function AddFileToFileIndex(FileName As String, ChunkSize As Long)
Dim Data1 As String
Dim Data2 As String
Dim i As Long
Dim Jesus As Long

If FileName = "" Then Exit Function

Jesus = FreeFile

i = 1

'Set the amount of data to grab equal in size to the ChunkSize
For e = 1 To ChunkSize
    Data2 = Data2 & " "
Next e

Open FileName For Binary As Jesus
    'Write header
    Put Opened, , LOF(Jesus) & "F"
    
    While i <= LOF(Jesus)
        'See if its in the chunksize range
        If LOF(Jesus) - i >= ChunkSize Then
            'If it is then get that much
            Data1 = Data2
            Get Jesus, i, Data1
    
            Put Opened, , Data1
        Else
            'If its not grab the last bit of the file
            For g = i To LOF(Jesus)
                Data1 = " "
                Get Jesus, g, Data1
                
                Put Opened, , Data1
            Next g
        End If
    
    i = i + ChunkSize
    Wend
Close Jesus
End Function

Public Function GetFromFileIndex(ByVal Index2 As Integer) As String
Dim X As Long, X2 As Long
Dim CurrData As String, Num As String

'See if we can start at a further position in the File Index or not
'The reason I did this was to make accessing more faster
If CurrIndex > Index2 Then
    X = 0
Else
    X = CurrIndex
End If
X2 = 1
 
While X <= Index2
    
    CurrData = " "
    Get Opened, X2, CurrData
    
    'Check if we're at the term char
    If CurrData = "F" Then
        'Check if we have reached the Index
        If X = Index2 Then
        
            'Go through the File Index grabbing the data
            X2 = X2 + 1
            For R = X2 To (X2 - 1) + Num
                CurrData = " "
                Get Opened, R, CurrData
                GetFromFileIndex = GetFromFileIndex & CurrData
            Next R
            Exit Function
            
        Else
            'If it isnt the index lets go on
            X2 = X2 + Num
            Num = vbNullString
        End If
        'Add 1 to the current index
        X = X + 1
    Else
        'If we arent at the term char add data to the Datalength
        Num = Num & CurrData
    End If
    
    'Add 1 to the file position
    X2 = X2 + 1
Wend
End Function

Public Function GetFileFromFileIndex(Index As Integer, FileName As String, ChunkSize As Long) As String
Dim X As Long, X2 As Long, R As Long
Dim CurrData As String, Num As String, CurrData2 As String
Dim Jesus As Long

If FileName = "" Then Exit Function

Jesus = FreeFile

'Set the amount of data to grab equal in size to the ChunkSize
For e = 1 To ChunkSize
    CurrData2 = CurrData2 & " "
Next e

'See if we can start at a further position in the File Index or not
'The reason I did this was to make accessing more faster
If CurrIndex > Index Then
    X = 0
Else
    X = CurrIndex
End If
X2 = 1

While X <= Index
    
    CurrData = " "
    Get Opened, X2, CurrData
    
    'Check if we're at the term char
    If CurrData = "F" Then
        'Check if we have reached the Index
        If X = Index Then
        
            'Go through the File Index grabbing the data
            Open FileName For Binary As Jesus
                i = X2 + 1
                R = X2 + Num
                While i <= R
                    'See if its in the chunksize range
                    If R - i >= ChunkSize Then
                        CurrData = CurrData2
                        Get Opened, i, CurrData
                        Put Jesus, , CurrData
                    Else
                        'Get the rest because its less than the chunksize
                        For g = i To R
                            CurrData = " "
                            Get Opened, g, CurrData
                            Put Jesus, , CurrData
                        Next g
                    End If
                i = i + ChunkSize
                Wend
            Close Jesus
            Exit Function
            
        Else
            'If it isnt the index lets go on
            X2 = X2 + Num
            Num = vbNullString
        End If
        'Add 1 to the current index
        X = X + 1
    Else
        'If we arent at the term char add data to the Datalength
        Num = Num & CurrData
    End If
    
    'Add 1 to the file position
    X2 = X2 + 1
Wend
End Function

Public Function GetFileIndexCount() As Long
Dim X As Long, X2 As Long
Dim CurrData As String, Num As String

X = 0
X2 = 1

While X2 <= LOF(Opened)
    
    CurrData = " "
    Get Opened, X2, CurrData
    
    'Check if we're at the term char
    If CurrData = "F" Then
        X2 = X2 + Num
        Num = vbNullString
        'Add 1 to the count
        X = X + 1
    Else
        'If we arent at the term char add data to the Datalength
        Num = Num & CurrData
    End If
    
    'Add 1 to the file position
    X2 = X2 + 1
Wend

GetFileIndexCount = X
End Function
