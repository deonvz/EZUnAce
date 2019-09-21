VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ez Unace"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "DRIVE2~1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   6060
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "&New Folder"
      Height          =   255
      Left            =   3480
      TabIndex        =   15
      Top             =   4440
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   3480
      TabIndex        =   14
      Text            =   "Type new folder name here"
      Top             =   3960
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   3000
      Top             =   3600
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00000000&
      Caption         =   "View only *.ace file in Directory"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   120
      TabIndex        =   12
      Top             =   4200
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Extract"
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   5040
      Width           =   975
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00000000&
      Caption         =   "All ace files"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   8
      Top             =   5520
      Value           =   -1  'True
      Width           =   1575
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00000000&
      Caption         =   "Selected Files"
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   240
      MaskColor       =   &H00FFFFFF&
      TabIndex        =   7
      Top             =   5040
      Width           =   1335
   End
   Begin VB.CommandButton exit 
      BackColor       =   &H00C0C0C0&
      Caption         =   "E&xit"
      Height          =   375
      Left            =   4800
      MaskColor       =   &H00400000&
      TabIndex        =   5
      Top             =   5520
      Width           =   975
   End
   Begin VB.DirListBox Dir2 
      BackColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   3480
      TabIndex        =   4
      Top             =   720
      Width           =   2415
   End
   Begin VB.DriveListBox Drive2 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   3480
      TabIndex        =   3
      Top             =   360
      Width           =   2415
   End
   Begin VB.FileListBox File1 
      BackColor       =   &H00FFFFFF&
      Height          =   870
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   2
      Top             =   3360
      Width           =   2415
   End
   Begin VB.DirListBox Dir1 
      BackColor       =   &H00FFFFFF&
      Height          =   2565
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   2415
   End
   Begin VB.DriveListBox Drive1 
      BackColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   2415
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   16
      Top             =   3360
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   "Dos instruction view"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   4560
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   1095
      Left            =   120
      Shape           =   4  'Rounded Rectangle
      Top             =   4920
      Width           =   5775
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   3000
      X2              =   3000
      Y1              =   -600
      Y2              =   4920
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "[ Extract to ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "[ From ]"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   480
      TabIndex        =   9
      Top             =   0
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Brought to you by Rico.Select the ace file you want to extract and also if all files or selected file must be extracted."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1800
      TabIndex        =   6
      Top             =   5040
      Width           =   2895
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Dim acename As String
Dim b As String

Private Sub Check1_Click()

    If Check1.Value = 1 Then
    File1.Pattern = "*.ace"
    Else
    File1.Pattern = "*.*"
    End If

End Sub

Private Sub Command1_Click()

    'Write instruction

    Label5.Caption = "ace" & " x" & " -r " & acename & " " & b & " " & Dir2.Path
    
    'Write as batch file
    
    Dim fileno As Integer
    fileno = FreeFile
    Open "ezunace.bat" For Append As fileno
    Print #fileno, Label5.Caption
    Print #fileno, "cls"
    Print #fileno, "@Echo Unace complete....Ez Unace brought to you by Rico visit Black Sun at www.angelfire.com/sc/blacksun"
    Print #fileno, "del ezunace.bat"
    Print #fileno, "exit"
    Close fileno
   
    'Run batch file and unace
    
    Shell "ezunace.bat"
    
    'Finished with everything
    
    MsgBox "You can close this program now but let the batchfile run to completion.", vbInformation

End Sub

Private Sub Command3_Click()
    If Text1.Text = "Type new folder name here" Then
    MsgBox "Enter new directory name.", vbExclamation
    ElseIf Dir2.Path = Drive2.Drive + "\" Then
    MkDir Dir2.Path & Text1.Text
    Dir2.Refresh
    Dir2.Path = Dir2.Path & Text1.Text
    Text1.Text = ""
    Else
     MkDir Dir2.Path & "\" & Text1.Text
     Dir2.Refresh
     Dir2.Path = Dir2.Path & "\" & Text1.Text
    Text1.Text = ""
    End If
    
End Sub

Private Sub Dir1_Change()
 
    File1.Path = Dir1.Path
    File1.Refresh
    
End Sub


    Private Sub Drive1_Change()
   
    On Error GoTo DriveError
     Dir1.Path = Drive1.Drive
     Exit Sub
     
DriveError:
     MsgBox "Device Not Ready!", vbExclamation, "Error"
     Drive1.Drive = Dir1.Path
     Exit Sub
     
     Dir1.Refresh
End Sub
    
    

Private Sub Drive2_Change()

    Dim freespace As String
     
    On Error GoTo DriveError
     
     'Check free disk Space
    drivename = Drive2.Drive
    x = GetDiskFreeSpace(Drive2.Drive & "\", secperclus, byteperclus, freeclus, totalclus)
    freespace = secperclus * byteperclus * freeclus
    Label4.Caption = "Free space:" & freespace & "bytes"
   Dir2.Path = Drive2.Drive
    'Do a Test of Drive2`s Text
    
    Debug.Print Drive2.Drive
    
    'Covert Free Space From Bytes to KB ,MB ,Gig ect.
    
     If freespace >= 1000000000 Then
    freespace = Val(freespace) / 1000000000
    Label4.Caption = "Free space:" & Int(freespace) & "Gig" & " available on " & Drive2.Drive
    
    ElseIf freespace >= 1000000 Then
    freespace = Val(freespace) / 1000000
    Label4.Caption = "Free space:" & Int(freespace) & "Mb" & " available on " & Drive2.Drive
   
    ElseIf freespace >= 1000 Then
    freespace = Val(freespace) / 1000
    Label4.Caption = "Free space:" & Int(freespace) & "Kb" & " available on " & Drive2.Drive
    
    Else
    Label4.Caption = "Free space:" & Int(freespace) & "Bytes" & " available on " & Drive2.Drive
     
    End If
     
     Exit Sub
     
     
DriveError:
     MsgBox "Device Not Ready!", vbExclamation, "Error"
     Drive2.Drive = Dir2.Path
     Exit Sub
     
    Dir2.Refresh
    
    

End Sub

Private Sub exit_Click()

    End
    
End Sub

Private Sub File1_Click()

'Make short filename and give value to acename
'===============================================

Dim x As String
Dim buffer As String * 255
x = GetShortPathName(Dir1.Path & "\" & File1.FileName, buffer, 255)
Debug.Print Left(buffer, x)  ' Could be C:\MYDOCU~1\RICHTE~1.rtf
acename = Left(buffer, x)

End Sub

Private Sub Form_Load()

Form1.Caption = ""

End Sub

Private Sub Option1_Click()

If Option1.Value = True Then
b = ""
Else
End If

End Sub

Private Sub Option2_Click()

If Option2.Value = True Then
b = "*.*"
Else
End If

End Sub

Private Sub Timer1_Timer()

    If Form1.Caption = "" Then
    Form1.Caption = "E"
    ElseIf Form1.Caption = "E" Then
    Form1.Caption = "Ez"
    ElseIf Form1.Caption = "Ez" Then
    Form1.Caption = "Ez U"
    ElseIf Form1.Caption = "Ez U" Then
    Form1.Caption = "Ez Un"
    ElseIf Form1.Caption = "Ez Un" Then
    Form1.Caption = "Ez Una"
    ElseIf Form1.Caption = "Ez Una" Then
    Form1.Caption = "Ez Unac"
    ElseIf Form1.Caption = "Ez Unac" Then
    Form1.Caption = "Ez Unace"
    ElseIf Form1.Caption = "Ez Unace" Then
    Form1.Caption = ""
    End If

End Sub
