VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "File Properties"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   4365
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkSystem 
      Caption         =   "System"
      Height          =   255
      Left            =   2640
      TabIndex        =   7
      Top             =   3960
      Width           =   855
   End
   Begin VB.CheckBox chkArchive 
      Caption         =   "Archive"
      Height          =   375
      Left            =   1200
      TabIndex        =   6
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CheckBox chkHidden 
      Caption         =   "Hidden"
      Height          =   255
      Left            =   2640
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.CheckBox chkRead 
      Caption         =   "Read-only"
      Height          =   255
      Left            =   1200
      TabIndex        =   4
      Top             =   3480
      Width           =   1095
   End
   Begin VB.CommandButton cmdProperties 
      Caption         =   "Properties"
      Height          =   405
      Left            =   3120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
   Begin VB.TextBox txtFile 
      Height          =   405
      Left            =   120
      TabIndex        =   0
      Text            =   "C:\My Documents\anilrag\shared\a.txt"
      Top             =   120
      Width           =   2895
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   195
      Left            =   120
      TabIndex        =   16
      Top             =   720
      Width           =   405
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Location:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Top             =   1140
      Width           =   660
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Size:"
      Height          =   195
      Left            =   120
      TabIndex        =   14
      Top             =   1560
      Width           =   345
   End
   Begin VB.Label lblFileType 
      AutoSize        =   -1  'True
      Caption         =   "Type:"
      Height          =   195
      Left            =   960
      TabIndex        =   13
      Top             =   720
      Width           =   405
   End
   Begin VB.Label lblLocation 
      AutoSize        =   -1  'True
      Caption         =   "Location"
      Height          =   195
      Left            =   960
      TabIndex        =   12
      Top             =   1140
      Width           =   615
   End
   Begin VB.Label lblSize 
      AutoSize        =   -1  'True
      Caption         =   "Size"
      Height          =   195
      Left            =   960
      TabIndex        =   11
      Top             =   1560
      Width           =   300
   End
   Begin VB.Label lblModified 
      AutoSize        =   -1  'True
      Caption         =   "Modified"
      Height          =   195
      Left            =   1680
      TabIndex        =   10
      Top             =   2640
      Width           =   600
   End
   Begin VB.Label lblMSDosName 
      AutoSize        =   -1  'True
      Caption         =   "MSDos Name"
      Height          =   195
      Left            =   1680
      TabIndex        =   9
      Top             =   2280
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Modified:"
      Height          =   195
      Left            =   240
      TabIndex        =   8
      Top             =   2640
      Width           =   645
   End
   Begin VB.Line Line4 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4200
      Y1              =   2060
      Y2              =   2060
   End
   Begin VB.Line Line3 
      BorderColor     =   &H80000005&
      X1              =   120
      X2              =   4200
      Y1              =   3135
      Y2              =   3135
   End
   Begin VB.Line Line2 
      X1              =   120
      X2              =   4200
      Y1              =   3120
      Y2              =   3120
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Attributes"
      Height          =   195
      Left            =   240
      TabIndex        =   3
      Top             =   3480
      Width           =   660
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "MS-DOS Name:"
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   2280
      Width           =   1140
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   4200
      Y1              =   2040
      Y2              =   2040
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Fname, Fpath As String

Private Sub cmdProperties_Click()
On Error Resume Next
    attribvalue = GetAttr(txtFile.Text)
    FileAtrributes
End Sub

Private Sub GetLocationAndFileName(ABPath As String)
On Error Resume Next
For i = Len(ABPath) To 0 Step -1
    If Mid(ABPath, i, 1) = "\" Then
        'Extract the filename...assuming the complete path of the file is typed
        'for Eg. C:\My Documents\anilrag\shared\a.txt
        'With this path, Fname will hold "a.txt" and Fpath will hold "C:\My Documents\anilrag\shared"
        Fname = Mid(ABPath, i + 1, Len(ABPath) - i)
        Fpath = Left(ABPath, Len(ABPath) - Len(Fname))
        Exit For
    End If
Next i
End Sub
Private Sub FileAtrributes()
On Error Resume Next
    'Call the Sub to Extract Filename and the Path
    GetLocationAndFileName (Trim(txtFile.Text))
    
    'Type of file....depending on the extension
    lblFileType.Caption = UCase(Right(txtFile.Text, 3)) & " File"
    
    lblLocation.Caption = Fpath
    'FileLen returns filesize in Bytes
    FLen = FileLen(txtFile.Text)
    If (FLen > 1024) Then
        lblSize.Caption = Int(FLen / 1024) & "KB, " & FLen & " Bytes"
    Else
        lblSize.Caption = FLen & " Bytes"
    End If
    
    lblMSDosName.Caption = Fname
    'FileDateTime() returns the the date time file was last modified
    lblModified.Caption = FileDateTime(txtFile.Text)
    'GetAttr() Returns an Integer representing the attributes of a file, directory, or folder.
    'vbNormal       0   Normal.
'    vbReadOnly     1   Read-only.
'    vbHidden       2   Hidden.
'    vbSystem       4   System file.
'    vbDirectory    16  Directory or folder.
'    vbArchive      32  File has changed since last backup.
    attribvalue = GetAttr(txtFile.Text)
    Select Case attribvalue
        Case 1:     chkRead.Value = 1
        Case 2:     chkHidden.Value = 1
        Case 4:     chkSystem.Value = 1
                   
        Case 32:    chkArchive.Value = 1
        Case 33:    chkArchive.Value = 1
                    chkRead.Value = 1
        Case 34:    chkArchive.Value = 1
                    chkHidden.Value = 1
        Case 35:    chkArchive.Value = 1
                    chkRead.Value = 1
                    chkHidden.Value = 1
        Case 36:    chkArchive.Value = 1
                    chkSystem.Value = 1
        Case 37:    chkArchive.Value = 1
                    chkRead.Value = 1
                    chkSystem.Value = 1
        Case 38:    chkArchive.Value = 1
                    chkSystem.Value = 1
                    chkHidden.Value = 1
        Case 39:    chkArchive.Value = 1
                    chkRead.Value = 1
                    chkHidden.Value = 1
                    chkSystem.Value = 1
        End Select
        'Similarly u can use the SetAttr statement to Set attribute information for a file.
End Sub

