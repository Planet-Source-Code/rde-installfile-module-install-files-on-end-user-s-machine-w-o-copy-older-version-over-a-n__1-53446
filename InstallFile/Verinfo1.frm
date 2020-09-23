VERSION 5.00
Begin VB.Form frmNavFiles 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "VerInfo Demo"
   ClientHeight    =   4695
   ClientLeft      =   5055
   ClientTop       =   3135
   ClientWidth     =   5640
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H80000008&
   LinkMode        =   1  'Source
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4695
   ScaleWidth      =   5640
   Begin VB.CommandButton cmdShowInfo 
      Caption         =   "&Show Version Info"
      Default         =   -1  'True
      Height          =   315
      Left            =   3000
      TabIndex        =   8
      Top             =   4170
      Width           =   2445
   End
   Begin VB.DriveListBox Drive1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   3000
      TabIndex        =   7
      Top             =   3735
      Width           =   2455
   End
   Begin VB.DirListBox Dir1 
      Appearance      =   0  'Flat
      Height          =   2340
      Left            =   3000
      TabIndex        =   5
      Top             =   1035
      Width           =   2430
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   3735
      Left            =   210
      TabIndex        =   1
      Top             =   765
      Width           =   2535
   End
   Begin VB.TextBox txtFileExt 
      Appearance      =   0  'Flat
      Height          =   288
      Left            =   3000
      TabIndex        =   3
      Text            =   "*.*"
      Top             =   360
      Width           =   2430
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Path:"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   5
      Left            =   210
      TabIndex        =   10
      Top             =   105
      Width           =   510
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Dri&ves:"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   4
      Left            =   3000
      TabIndex        =   6
      Top             =   3450
      Width           =   1680
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Directories:"
      ForeColor       =   &H80000008&
      Height          =   195
      Index           =   3
      Left            =   3000
      TabIndex        =   4
      Top             =   735
      Width           =   2340
   End
   Begin VB.Label lblPath 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "c:\"
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   765
      TabIndex        =   9
      Top             =   105
      Width           =   4830
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "&Files:"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   0
      Left            =   210
      TabIndex        =   0
      Top             =   405
      Width           =   615
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "File &Ext:"
      ForeColor       =   &H80000008&
      Height          =   210
      Index           =   1
      Left            =   2160
      TabIndex        =   2
      Top             =   405
      Width           =   750
   End
End
Attribute VB_Name = "frmNavFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Declare Function GetSystemDirectory Lib "kernel32" Alias _
    "GetSystemDirectoryA" (ByVal szPath As String, ByVal cbBytes As Long) As Long

Private sPath As String

Private Sub cmdShowInfo_Click()
   sPath = lblPath.Caption
   If (Right$(sPath, 1) <> "\") Then sPath = sPath & "\"
   DisplayInfo sPath & File1.List(File1.ListIndex)
End Sub

Private Sub Dir1_Change()
   File1.Path = Dir1.Path
   lblPath.Caption = File1.Path
End Sub

Private Sub Drive1_Change()
   Dir1.Path = Drive1.Drive
   File1.Path = Dir1.Path
   lblPath.Caption = File1.Path
End Sub

Private Sub File1_DblClick()
   cmdShowInfo_Click
End Sub

Private Sub txtFileExt_Change()
   File1.Pattern = txtFileExt.Text
End Sub

Private Sub Form_Load()
   Dim sBuffer As String
   Dim sSysPath As String
   Dim rc As Long

   ' Set default directory to Windows System
   sBuffer = Space$(256)
   rc = GetSystemDirectory(sBuffer, 256)

   sSysPath = LCase$(Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1))

   Dir1.Path = sSysPath
   File1.Path = sSysPath
   Drive1.Drive = Left$(sSysPath, 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   End
End Sub

Public Sub DisplayInfo(sFileSpec As String)
   If (sFileSpec = "") Then Err.Raise 5
   If (Dir(sFileSpec) = "") Then Err.Raise 53

   Dim fvi As FILEVERINFO

   If GetVersionInfo(sFileSpec, fvi) Then
        frmVerInfo.Show

        frmVerInfo.CurrentX = 2
        frmVerInfo.CurrentY = 1
        frmVerInfo.Print "File: "; sFileSpec
        
        frmVerInfo.CurrentX = 2
        frmVerInfo.CurrentY = 3
        frmVerInfo.Print "Language:"
        
        frmVerInfo.CurrentX = 2
        frmVerInfo.Print "Company:"
        
        frmVerInfo.CurrentX = 2
        frmVerInfo.Print "Copyright:"

        frmVerInfo.CurrentX = 2
        frmVerInfo.Print "Internal Name:"
        
        frmVerInfo.CurrentX = 2
        frmVerInfo.Print "Original Name:"

        frmVerInfo.CurrentX = 2
        frmVerInfo.Print "File Version:"
        
        frmVerInfo.CurrentX = 2
        frmVerInfo.Print "Product Version:"
        
        frmVerInfo.CurrentX = 2
        frmVerInfo.Print "File Desc:"
        
        frmVerInfo.CurrentX = 2
        frmVerInfo.Print "File Flags:"
        
        frmVerInfo.CurrentX = 2
        frmVerInfo.Print "File OS:"
        
        frmVerInfo.CurrentX = 2
        frmVerInfo.Print "File Type:"
        
        frmVerInfo.CurrentX = 2
        frmVerInfo.Print "File Sub-type:"
        

        frmVerInfo.CurrentX = 18
        frmVerInfo.CurrentY = 3
        frmVerInfo.Print fvi.Language
        
        frmVerInfo.CurrentX = 18
        frmVerInfo.Print fvi.Company
        
        frmVerInfo.CurrentX = 18
        frmVerInfo.Print fvi.Copyright

        frmVerInfo.CurrentX = 18
        frmVerInfo.Print fvi.InternalName
        
        frmVerInfo.CurrentX = 18
        frmVerInfo.Print fvi.OriginalName

        frmVerInfo.CurrentX = 18
        frmVerInfo.Print fvi.FileVer
        
        frmVerInfo.CurrentX = 18
        frmVerInfo.Print fvi.ProdVer
        
        frmVerInfo.CurrentX = 18
        frmVerInfo.Print fvi.FileDesc
        
        frmVerInfo.CurrentX = 18
        frmVerInfo.Print fvi.FileFlags
        
        frmVerInfo.CurrentX = 18
        frmVerInfo.Print fvi.FileOS
        
        frmVerInfo.CurrentX = 18
        frmVerInfo.Print fvi.FileType
        
        frmVerInfo.CurrentX = 18
        frmVerInfo.Print fvi.FileSubtype
   Else
        MsgBox "No Version Info available."
   End If
End Sub
