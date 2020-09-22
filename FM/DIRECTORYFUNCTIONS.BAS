Attribute VB_Name = "DirectoryFunctions"
Option Explicit

Private Declare Function PathMatchSpec Lib "shlwapi" _
   Alias "PathMatchSpecW" _
  (ByVal pszFileParam As Long, _
   ByVal pszSpec As Long) As Long
   
Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias _
    "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
    ByVal lpBuffer As String) As Long
    
Private Declare Function FindExecutable Lib _
    "shell32.dll" Alias "FindExecutableA" _
    (ByVal lpFile As String, ByVal lpDirectory _
    As String, ByVal lpResult As String) As Long

Function ScanTree(ByVal Path As String, Exclude As Boolean, ParamArray Mask()) As Variant

    'This function recursively traverses all files and folders within and below
    'the path specified and returns an array of all files (full pathname) which
    'match the mask supplied (if any). 04/13/02 - dgk

    On Error GoTo ERRORHANDLER
    
    Dim aCurrentFolderContents() As String
    Dim aFiles() As String
    Dim strCurrentFile As String
    Dim intFileCount As Long
    Dim intCurFile As Long
    Dim intParmCount As Integer
    Dim i As Integer
    Dim aMasks As Variant
    Dim bMatchesMask As Boolean
    
    'make sure the path ends with a "\"
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    
    'because we are calling ScanTree recursively, recursive calls must pass
    'the masks as a single parameter as an array.  This also adds the
    'flexibility of passing masks in an array on the initial call
    intParmCount = UBound(Mask)
    If intParmCount <> -1 Then
        If intParmCount = 0 Then
            If IsArray(Mask(0)) Then
                aMasks = Mask(0)
            Else
                aMasks = Array()
                aAdd aMasks, Mask(0)
            End If
        Else
            aMasks = Array()
            For i = 0 To intParmCount
                aAdd aMasks, Mask(i)
            Next i
        End If
    End If
    
    'we have to create an array with the the contents of the current folder
    'because the Dir function does not retain state between recursive calls
    
    'get all the subfolders
    strCurrentFile = GetFile(Path & "*.", vbDirectory)
    Do While strCurrentFile <> ""
        aAdd aCurrentFolderContents, strCurrentFile
        strCurrentFile = GetFile
        DoEvents
    Loop
    
    'get all the files that match the mask(s)
    'if the list of masks is exclusionary then get all files
    'we'll have to match the mask for each file individually
    If Exclude Then
        strCurrentFile = GetFile(Path & "*.*", vbArchive + vbHidden + vbNormal + vbReadOnly)
        Do While strCurrentFile <> ""
            aAdd aCurrentFolderContents, strCurrentFile
            strCurrentFile = GetFile
            DoEvents
        Loop
    Else
        intParmCount = aLen(aMasks) - 1
        For i = 0 To intParmCount
            strCurrentFile = GetFile(Path & aMasks(i), vbArchive + vbHidden + vbNormal + vbReadOnly)
            Do While strCurrentFile <> ""
                aAdd aCurrentFolderContents, strCurrentFile
                strCurrentFile = GetFile
                DoEvents
            Loop
        Next i
    End If
    
    'adjust file count for zero-based array
    intFileCount = aLen(aCurrentFolderContents) - 1
    
    'process folder contents
    For intCurFile = 0 To intFileCount
        If UCase(aCurrentFolderContents(intCurFile)) <> "PAGEFILE.SYS" Then 'NT doesn't set the system attribute for pagefile.sys
            If GetAttr(Path & aCurrentFolderContents(intCurFile)) And vbDirectory Then
                If aCurrentFolderContents(intCurFile) <> "." And aCurrentFolderContents(intCurFile) <> ".." Then
                    'recurse subfolders
                    aConcat aFiles, ScanTree(Path & aCurrentFolderContents(intCurFile), Exclude, aMasks)
                End If
            Else
            
                'additional mask matching must be done if the mask parameters
                'are exclusions
                bMatchesMask = True
                If Exclude Then
                    For i = 0 To intParmCount
                        If Not MatchesMask(Path & aCurrentFolderContents(intCurFile), CStr(aMasks(i)), Exclude) Then
                            bMatchesMask = False
                            Exit For
                        End If
                    Next i
                End If
            
                If bMatchesMask Then
                    aAdd aFiles, Path & aCurrentFolderContents(intCurFile)
                End If
                
            End If
        End If
        DoEvents
    Next
    
    ScanTree = aFiles
    
    Exit Function
    
ERRORHANDLER:
   
    Err.Raise vbObjectError + 1234, "ScanTree"
   
End Function

Private Function MatchesMask(sFile As String, sSpec As String, Optional Exclude As Boolean) As Boolean

   MatchesMask = CBool(PathMatchSpec(StrPtr(sFile), StrPtr(sSpec))) = Not Exclude
   
End Function

' Check whether a given drive exists
' Note that this returns True even if the drive isn't currently ready
' (e.g. a diskette isn't in drive A:)
Function DriveExists(ByVal sDrive As String) As Boolean

    Dim buffer As String
    buffer = Space(64)
    ' return False if invalid argument
    If Len(sDrive) = 0 Then Exit Function
    'get the string that contains all drives
    GetLogicalDriveStrings Len(buffer), buffer
    ' check that the letter we're looking for is there
    DriveExists = InStr(1, buffer, Left$(sDrive, 1), vbTextCompare)
    
End Function

' Test whether a drive is ready
Function IsDriveReady(sDrive As String) As Boolean

   Dim fso As New FileSystemObject
   IsDriveReady = fso.GetDrive(sDrive).IsReady
   Set fso = Nothing
   
End Function

Function GetFile(Optional ByVal strPath, Optional lngType As FileAttribute) As String

    'my solution to Dir - doesn't croak on drive that doesn't exist or is not ready
    Dim strDrive As String
    
    If Not IsMissing(strPath) Then
    
        If InStr(strPath, ":") = 2 Then
            strDrive = Left(strPath, 1)
        Else
            strDrive = Left(App.Path, 1)
        End If
        
        If Not DriveExists(strDrive) Then Exit Function
        
        If Not IsDriveReady(strDrive) Then Exit Function
        
        GetFile = Dir(strPath, lngType)
        
    Else
    
        GetFile = Dir()
        
    End If
    
End Function

Function GetFileList(strSourceDir As String, bIncludeSubs As Boolean, Optional strMask As String) As Variant

    Dim strFileName     As String
    Dim aFileList       As Variant
    
    Dim t1 As Single
    Dim totaltime As Single


    If Right(strSourceDir, 1) <> "\" Then strSourceDir = strSourceDir & "\"
    If strMask = "" Then strMask = "*.*"

    t1 = Timer
    If bIncludeSubs Then
        aFileList = ScanTree(strSourceDir, False, strMask)
    Else
        aFileList = Array()
        strFileName = GetFile(strSourceDir & strMask, vbNormal + vbReadOnly + vbHidden)
        Do While strFileName <> ""
            aAdd aFileList, strSourceDir & strFileName
            strFileName = GetFile
            DoEvents
        Loop
    End If
    
    GetFileList = aFileList
    
    totaltime = Timer - t1

End Function

Public Function GetAssociatedExecutable(ByVal FileName As String) As String

    Const MAX_PATH As Long = 260
    Dim Path As String

    'Get name of associated EXE
    Path = String$(MAX_PATH, 0)

    Call FindExecutable(FileName, _
        vbNullString, Path)
        
    GetAssociatedExecutable = Left$( _
        Path, InStr(Path, vbNullChar) - 1)

End Function


