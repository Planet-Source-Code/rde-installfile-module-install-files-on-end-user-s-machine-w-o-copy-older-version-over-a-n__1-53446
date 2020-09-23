Attribute VB_Name = "InstFile"
Option Explicit

' Important - any project that makes use of this module must
'             also include the VerInfo.bas module.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' Use error handlers in all procedures that call InstallFile.

' All effort has been made to eliminate errors. Therefore, this
' function should operate reliably and without any unexpected
' runtime exceptions, so long as you do not pass invalid args.

' If InstallFile does receive invalid arguments then it will
' RAISE ERRORS to let you know.

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

' When installing a file on the user's machine, you should not
' copy an older version of the file over a new version.

' The InstallFile function in this module uses the VerFindFile
' and VerInstallFile API functions to copy files to the user's
' machine, and will not overwrite an existing file with an older
' version.

' The InstallFile function's return values are defined as the
' eInstallFile enumeration and details are as follows:

'SUCCESS_Did_Not_Exist (-2) The file was installed, it did not exist
'SUCCESS_Was_Updated   (-1) The file was installed, it was updated
'SUCCESS_Already_Newer  (0) The file was not installed, as a newer
'                           version of this file already exists

' Return values >0 indicate that an error occured, and the file
' was not installed.

' You can get further information about errors from a positive
' return value which is a bitmask that indicates exceptions.

' It can be one or more of the following values:

'VIF_TEMPFILE         The temporary copy of the new file is in the destination directory. The cause of failure is reflected in other flags. _
                      If this flag is returned, the sFileSpec parameter of InstallFile is modified to specify the temporary file.

'VIF_CANNOTREADSRC    The function cannot read the source file. This could mean that the path was not specified properly.
'VIF_CANNOTREADDST    The function cannot read the destination (existing) file. This prevents the function from examining the file's attributes.

'VIF_CANNOTCREATE     The function cannot create the temporary file. The specific error may be described by another flag.
'VIF_CANNOTRENAME     The function cannot rename the temporary file, but already deleted the destination file.

'VIF_SRCOLD           The file to install is older than the preexisting file. This error can be overridden by calling VerInstallFile again with the VIFF_FORCEINSTALL flag set.
'VIF_WRITEPROT        The preexisting file is write-protected. This error can be overridden by calling VerInstallFile again with the VIFF_FORCEINSTALL flag set.

'VIF_FILEINUSE        The pre-existing file is in use by the system and cannot be deleted.
'VIF_OUTOFSPACE       The function cannot create the temporary file due to insufficient disk space on the destination drive.
'VIF_ACCESSVIOLATION  A read, create, delete, or rename operation failed due to an access violation.
'VIF_SHARINGVIOLATION A read, create, delete, or rename operation failed due to a sharing violation.
'VIF_OUTOFMEMORY      The function cannot complete the requested operation due to insufficient memory. Generally, this means the application ran out of memory attempting to expand a compressed file.
'VIF_CANNOTDELETE     The function cannot delete the destination file, or cannot delete the existing version of the file located in another directory. If the VIF_TEMPFILE bit is set, the installation failed, and the destination file probably cannot be deleted.
'VIF_CANNOTDELETECUR  The existing version of the file could not be deleted and VIFF_DONTDELETEOLD was not specified.

'VIF_MISMATCH         The new and preexisting files differ in one or more attributes. This error can be overridden by calling VerInstallFile again with the VIFF_FORCEINSTALL flag set.
'VIF_DIFFLANG         The new and preexisting files have different language or code-page values. This error can be overridden by calling VerInstallFile again with the VIFF_FORCEINSTALL flag set.
'VIF_DIFFCODEPG       The new file requires a code page that cannot be displayed by the version of the system currently running. This error can be overridden by calling VerInstallFile with the VIFF_FORCEINSTALL flag set.
'VIF_DIFFTYPE         The new file has a different type, subtype, or operating system from the preexisting file. This error can be overridden by calling VerInstallFile again with the VIFF_FORCEINSTALL flag set.

Private Declare Function VerFindFile Lib "Version.dll" Alias "VerFindFileA" (ByVal dwFlags As Long, ByVal szFileName As String, ByVal szWinDir As String, ByVal szAppDir As String, ByVal szCurDir As String, lpuCurDirLen As Long, ByVal szDestDir As String, lpuDestDirLen As Long) As Long
Private Declare Function VerInstallFile Lib "Version.dll" Alias "VerInstallFileA" (ByVal dwFlags As Long, ByVal szSrcFileName As String, ByVal szDestFileName As String, ByVal szSrcDir As String, ByVal szDestDir As String, ByVal szCurDir As String, ByVal szTmpFile As String, lpuTmpFileLen As Long) As Long

'  ----- InstallFile() return values -----
Public Enum eInstallFile
    SUCCESS_Did_Not_Exist = &HFFFFFFFE ' -2
    SUCCESS_Was_Updated = &HFFFFFFFF   ' -1
    SUCCESS_Already_Newer = &H0&       '  0
    ' VerInstallFile() error codes
    VIF_TEMPFILE = &H1&
    VIF_MISMATCH = &H2&
    VIF_SRCOLD = &H4&
    VIF_DIFFLANG = &H8&
    VIF_DIFFCODEPG = &H10&
    VIF_DIFFTYPE = &H20&
    VIF_WRITEPROT = &H40&
    VIF_FILEINUSE = &H80&
    VIF_OUTOFSPACE = &H100&
    VIF_ACCESSVIOLATION = &H200&
    VIF_SHARINGVIOLATION = &H400&
    VIF_CANNOTCREATE = &H800&
    VIF_CANNOTDELETE = &H1000&
    VIF_CANNOTRENAME = &H2000&
    VIF_CANNOTDELETECUR = &H4000&
    VIF_OUTOFMEMORY = &H8000&
    VIF_CANNOTREADSRC = &H10000
    VIF_CANNOTREADDST = &H20000
End Enum
'  ----- VerInstallFile() error code -----
Private Const VIF_BUFFTOOSMALL = &H40000 ' Handled by InstallFile internally

'  ----- VerInstallFile() flags -----
Private Const VIFF_FORCEINSTALL = &H1
Private Const VIFF_DONTDELETEOLD = &H2

'  ----- VerFindFile() flags -----
Private Const VFFF_ISSHAREDFILE = &H1

'  ----- VerFindFile() errors -----
Private Const VFF_CURNEDEST = &H1
Private Const VFF_FILEINUSE = &H2
Private Const VFF_BUFFTOOSMALL = &H4

Private Declare Function GetWinDir Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetAttribs Lib "kernel32" Alias "GetFileAttributesA" (ByVal lpSpec As String) As Long
Private Declare Function SetAttribs Lib "kernel32" Alias "SetFileAttributesA" (ByVal lpSpec As String, ByVal dwAttributes As Long) As Long
Private Declare Function ShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long
Private Declare Function LongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

Private Const MAX_PATH As Long = 260
Private Const DIR_SEP As String = "\"

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' **** Public InstallFile Function ****

Public Function InstallFile(sFileSpec As String, sAppPath As String, Optional sDestFileName As String, Optional fIsSharedFile As Boolean, Optional fForceInstall As Boolean, Optional fDeleteOtherCopies As Boolean) As eInstallFile
    If Len(sFileSpec) = 0 Then Err.Raise 5

    On Error GoTo Fail

    If Not FileExists(sFileSpec) Then
        InstallFile = VIF_CANNOTREADSRC
        Exit Function
    End If

    Dim lRetVal As Long, lShareFlag As Long
    Dim sSrcName As String, sSrcPath As String
    Dim sShortSrcName As String, sShortSrcPath As String
    Dim sCurDir As String, sRecDestDir As String
    Dim lCurBufLen As Long, lDestBufLen As Long

    SplitFileSpec sFileSpec, sSrcName, sSrcPath
    SplitFileSpec sFileSpec, sShortSrcName, sShortSrcPath, True
    If Len(sDestFileName) = 0 Then sDestFileName = sSrcName

    If (fIsSharedFile) Then lShareFlag = VFFF_ISSHAREDFILE

    lCurBufLen = MAX_PATH
    lDestBufLen = MAX_PATH

TryFindAgain:
    sCurDir = String$(lCurBufLen, vbNullChar)
    sRecDestDir = String$(lDestBufLen, vbNullChar)

    lRetVal = VerFindFile(lShareFlag, sDestFileName, WinDir, sAppPath, sCurDir, lCurBufLen, sRecDestDir, lDestBufLen)

    If (lRetVal And VFF_BUFFTOOSMALL) Then GoTo TryFindAgain
    If (lRetVal And VFF_FILEINUSE) Then
        InstallFile = VIF_FILEINUSE ' File is in use
        Exit Function
    End If

    sCurDir = TrimNull(sCurDir)

    Dim sDestDir As String, fDummy As Boolean
    sDestDir = AddBackslash(GetLongPathName(TrimNull(sRecDestDir)))

    If FileExists(sDestDir & sDestFileName) Then
        ' Requires VerInfo.bas module
        Dim ffiSrc As FIXEDFILEINFO, ffiDest As FIXEDFILEINFO
        GetVersionInfoStruct sSrcPath & sSrcName, ffiSrc
        GetVersionInfoStruct sDestDir & sDestFileName, ffiDest
        If Not IsNewerVersion(ffiSrc, ffiDest) Then
            InstallFile = 0 ' Existing file is equal or newer
            Exit Function
        End If
    Else
        ' If the destination file does not already exist, we
        ' create a dummy with the correct (long) filename so
        ' that we can get its short filename for VerInstallFile.
        Open sDestDir & sDestFileName For Output Access Write As #1
        fDummy = True
        Close #1
    End If

    ' VerInstallFile under Windows 95 does not handle long
    ' filenames, so we must give it the short versions.
    Dim sShortDestDir As String, sShortDestName As String

    SplitFileSpec sDestDir & sDestFileName, sShortDestName, sShortDestDir, True

    If fDummy Then Kill sDestDir & sDestFileName

    Dim sTempFile As String, lTempBufLen As Long
    lTempBufLen = MAX_PATH

    Dim flags As Long
    If fDeleteOtherCopies = False Then flags = VIFF_DONTDELETEOLD
    If fForceInstall Then flags = flags Or VIFF_FORCEINSTALL

TryInstAgain:
    sTempFile = String$(lTempBufLen, vbNullChar)

    lRetVal = VerInstallFile(flags, sShortSrcName, sShortDestName, sShortSrcPath, sShortDestDir, sCurDir, sTempFile, lTempBufLen)

    If (lRetVal And VIF_BUFFTOOSMALL) Then GoTo TryInstAgain
    Dim lAttribs As Long
    If (lRetVal = VIF_WRITEPROT) Then
        lAttribs = GetAttribs(sShortDestDir & sShortDestName)
        SetAttribs sShortDestDir & sShortDestName, vbNormal
        GoTo TryInstAgain
    End If

    ' The InstallFile function's return values are:
    '   -2 : The file was installed successfully, it did not exist
    '   -1 : The file was installed successfully, it was updated
    '    0 : The file was not installed, a newer version exists
    '   >0 : An error occured, the file was not installed

    If (lRetVal = 0) Then
        If fDummy Then
            InstallFile = SUCCESS_Did_Not_Exist
        Else
            InstallFile = SUCCESS_Was_Updated
        End If
    ElseIf (lRetVal = VIF_SRCOLD) Then
        InstallFile = SUCCESS_Already_Newer
    Else
        If (lRetVal And VIF_TEMPFILE) Then
            sFileSpec = sDestDir & TrimNull(sTempFile)
        End If
        InstallFile = lRetVal
    End If

    ' One more kludge for long filenames: VerInstallFile may have renamed
    ' the file to its short version if it went through with the copy.
    ' If so, we simply rename it back to what it should be.
    If FileExists(sShortDestDir & sShortDestName) Then
        Name sShortDestDir & sShortDestName As sDestDir & sDestFileName
        If (lAttribs) Then SetAttribs sDestDir & sDestFileName, lAttribs
    End If

    Exit Function
Fail:
    ' Abort if the version or file expansion DLLs failed
    InstallFile = VIF_CANNOTCREATE
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++

' **** Private Support Functions ****

Private Function TrimNull(sNullTerm As String) As String
    Dim Idx As Integer: Idx = InStr(sNullTerm, vbNullChar)
    If (Idx <> 0) Then
        TrimNull = Left$(sNullTerm, Idx - 1)
    Else
        TrimNull = sNullTerm
    End If
End Function

Private Function WinDir() As String
    Dim sTemp As String: sTemp = String$(MAX_PATH, vbNullChar)
    If (GetWinDir(sTemp, MAX_PATH)) Then WinDir = TrimNull(sTemp)
End Function

Private Sub SplitFileSpec(sFileSpec As String, sOutFileName As String, sOutPath As String, Optional fGetShort As Boolean)
    Dim Idx As Integer, sFile As String
    If fGetShort Then
        sFile = GetShortPathName(sFileSpec)
    Else
        sFile = GetLongPathName(sFileSpec)
    End If
    For Idx = Len(sFile) To 1 Step -1
        If (StrComp(Mid$(sFile, Idx, 1), DIR_SEP) = 0) Then
            sOutFileName = Mid$(sFile, Idx + 1)
            sOutPath = Left$(sFile, Idx)
            Exit Sub
        End If
    Next
End Sub

Private Function AddBackslash(sSpec As String) As String
    If Right$(sSpec, 1) <> DIR_SEP Then
        AddBackslash = sSpec & DIR_SEP
    Else
        AddBackslash = sSpec
    End If
End Function

Private Function FileExists(sFileSpec As String) As Boolean
    Dim Attribs As Long: Attribs = GetAttribs(sFileSpec)
    If (Attribs <> -1) Then
        FileExists = ((Attribs And vbDirectory) <> vbDirectory)
    End If
End Function

Private Function GetLongPathName(sShortPath As String) As String
    GetLongPathName = sShortPath
    On Error GoTo GetFailed
    Dim sPath As String, lResult As Long
    sPath = String$(MAX_PATH, vbNullChar)
    lResult = LongPathName(sShortPath, sPath, MAX_PATH)
    If (lResult) Then GetLongPathName = TrimNull(sPath)
GetFailed:
End Function

Private Function GetShortPathName(sLongPath As String) As String
    GetShortPathName = sLongPath
    On Error GoTo GetFailed
    Dim sPath As String, lResult As Long
    sPath = String$(MAX_PATH, vbNullChar)
    lResult = ShortPathName(sLongPath, sPath, MAX_PATH)
    If (lResult) Then GetShortPathName = TrimNull(sPath)
GetFailed:
End Function

' +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
