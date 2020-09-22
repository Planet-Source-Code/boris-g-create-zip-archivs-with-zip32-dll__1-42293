VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2100
   ClientLeft      =   60
   ClientTop       =   375
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   ScaleHeight     =   2100
   ScaleWidth      =   6600
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Add File"
      Height          =   975
      Left            =   3600
      TabIndex        =   1
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Unzip"
      Height          =   975
      Left            =   4920
      TabIndex        =   0
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
addfile App.Path & "\thezip.zip", "c:\cdcopy\cdcopy.exe"
End Sub

Public Function addfile(archivpath As String, fileadd As String)
Dim gzip As CGZipFiles

    Set gzip = New CGZipFiles
    
    If Dir(fileadd) = "" Then MsgBox "change sourcecode": Exit Function
    If Dir(archivpath) = "" Then MsgBox "change sourcecode": Exit Function
    
    With gzip
        .ZipFileName = archivpath
        .UpdatingZip = False
        .addfile fileadd
        If .MakeZipFile <> 0 Then MsgBox "error": End
    End With
    Set ozip = Nothing
End Function

Public Function extARchive(aPath As String, extPath As String)
Dim bzip As CGUnzipFiles
Set bzip = New CGUnzipFiles

With bzip
    .Unzip aPath, extPath
End With
End Function


Private Sub Command2_Click()
extARchive App.Path & "\thezip.zip", App.Path & "\"

End Sub

Private Sub Command3_Click()

End Sub
