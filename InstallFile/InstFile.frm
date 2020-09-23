VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   2910
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   2910
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Install the file"
      Height          =   645
      Left            =   765
      TabIndex        =   0
      Top             =   1350
      Width           =   3150
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Go on, click that button. Don't be affraid"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   60
      TabIndex        =   1
      Top             =   750
      Width           =   4545
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Important - any project that makes use of InstFile.bas
'             must also include VerInfo.bas.

' If you don't know what to do from this simple demo then
' you should read the header comments in InstFile.bas and
' the accompanying InstallFile.htm.

Private Sub Command1_Click()

   Dim rc As eInstallFile
   Dim sFile As String
   Dim sRetMsg As String

   sFile = "C:\WINDOWS\Desktop\Test.exe"

   rc = InstallFile(sFile, CurDir)

   If rc <= 0 Then

      Select Case rc
         Case rc = SUCCESS_Did_Not_Exist
            sRetMsg = sFile & vbCrLf & "was installed successfully, it did not exist"
            
         Case rc = SUCCESS_Was_Updated
            sRetMsg = sFile & vbCrLf & "was installed successfully, it was updated"
            
         Case rc = SUCCESS_Already_Newer
            sRetMsg = sFile & vbCrLf & "was not installed, a newer version already exists"
      End Select

   Else
      sRetMsg = "Error " & rc & " occured"
      
      If (rc And VIF_TEMPFILE) = VIF_TEMPFILE Then sRetMsg = sRetMsg & vbCrLf & _
         "The temporary copy of the new file is in the destination directory. The cause of failure is reflected in other flags. If this flag is returned, the sFileSpec parameter of InstallFile is modified to specify the temporary file."
      
      If (rc And VIF_CANNOTREADSRC) = VIF_CANNOTREADSRC Then sRetMsg = sRetMsg & vbCrLf & _
         "The function cannot read the source file. This could mean that the path was not specified properly."
      
      If (rc And VIF_CANNOTREADDST) = VIF_CANNOTREADDST Then sRetMsg = sRetMsg & vbCrLf & _
         "The function cannot read the destination (existing) file. This prevents the function from examining the file's attributes."
      
      If (rc And VIF_CANNOTCREATE) = VIF_CANNOTCREATE Then sRetMsg = sRetMsg & vbCrLf & _
         "The function cannot create the temporary file. The specific error may be described by another flag."
      
      If (rc And VIF_CANNOTRENAME) = VIF_CANNOTRENAME Then sRetMsg = sRetMsg & vbCrLf & _
         "The function cannot rename the temporary file, but already deleted the destination file."
      
      If (rc And VIF_SRCOLD) = VIF_SRCOLD Then sRetMsg = sRetMsg & vbCrLf & _
         "The file to install is older than the preexisting file. This error can be overridden by calling VerInstallFile again with the VIFF_FORCEINSTALL flag set."
      
      If (rc And VIF_WRITEPROT) = VIF_WRITEPROT Then sRetMsg = sRetMsg & vbCrLf & _
         "The preexisting file is write-protected. This error can be overridden by calling VerInstallFile again with the VIFF_FORCEINSTALL flag set."
      
      If (rc And VIF_FILEINUSE) = VIF_FILEINUSE Then sRetMsg = sRetMsg & vbCrLf & _
         "The pre-existing file is in use by the system and cannot be deleted."
      
      If (rc And VIF_OUTOFSPACE) = VIF_OUTOFSPACE Then sRetMsg = sRetMsg & vbCrLf & _
         "The function cannot create the temporary file due to insufficient disk space on the destination drive."
      
      If (rc And VIF_ACCESSVIOLATION) = VIF_ACCESSVIOLATION Then sRetMsg = sRetMsg & vbCrLf & _
         "A read, create, delete, or rename operation failed due to an access violation."
      
      If (rc And VIF_SHARINGVIOLATION) = VIF_SHARINGVIOLATION Then sRetMsg = sRetMsg & vbCrLf & _
         "A read, create, delete, or rename operation failed due to a sharing violation."
      
      If (rc And VIF_OUTOFMEMORY) = VIF_OUTOFMEMORY Then sRetMsg = sRetMsg & vbCrLf & _
         "The function cannot complete the requested operation due to insufficient memory. Generally, this means the application ran out of memory attempting to expand a compressed file."
      
      If (rc And VIF_CANNOTDELETE) = VIF_CANNOTDELETE Then sRetMsg = sRetMsg & vbCrLf & _
         "The function cannot delete the destination file, or cannot delete the existing version of the file located in another directory. If the VIF_TEMPFILE bit is set, the installation failed, and the destination file probably cannot be deleted."
      
      If (rc And VIF_CANNOTDELETECUR) = VIF_CANNOTDELETECUR Then sRetMsg = sRetMsg & vbCrLf & _
         "The existing version of the file could not be deleted and VIFF_DONTDELETEOLD was not specified."
      
      If (rc And VIF_MISMATCH) = VIF_MISMATCH Then sRetMsg = sRetMsg & vbCrLf & _
         "The new and preexisting files differ in one or more attributes. This error can be overridden by calling VerInstallFile again with the VIFF_FORCEINSTALL flag set."
      
      If (rc And VIF_DIFFLANG) = VIF_DIFFLANG Then sRetMsg = sRetMsg & vbCrLf & _
         "The new and preexisting files have different language or code-page values. This error can be overridden by calling VerInstallFile again with the VIFF_FORCEINSTALL flag set."
      
      If (rc And VIF_DIFFCODEPG) = VIF_DIFFCODEPG Then sRetMsg = sRetMsg & vbCrLf & _
         "The new file requires a code page that cannot be displayed by the version of the system currently running. This error can be overridden by calling VerInstallFile with the VIFF_FORCEINSTALL flag set."
      
      If (rc And VIF_DIFFTYPE) = VIF_DIFFTYPE Then sRetMsg = sRetMsg & vbCrLf & _
         "The new file has a different type, subtype, or operating system from the preexisting file. This error can be overridden by calling VerInstallFile again with the VIFF_FORCEINSTALL flag set."
      
   End If
   
   Debug.Print sRetMsg

End Sub


