Attribute VB_Name = "QRS_LibDOS"
Option Explicit

' Module : QRS_LibDOS
' Project: any
' Purpose: Disk Operating System functions
'          possibly without the use of the
'          FileSystemObject library
' By     : QRS, Roger Strebel
' Date   : 13.03.2018                  FileExists, FileDelete added
'          13.02.2019                  PathFile handles empty filename
'          17.05.2019                  PathIsAbs added
'          19.05.2019                  PathLoc1Dr added for onedrive paths
' --- The public interface
'     DrivExists                       Check if a disk volume exists 13.03.2018
'     FileDelete                       Delete a file if exists       13.03.2018
'     FileExists                       Check if a file exists        13.03.2018
'     FileList                         List of files in folder
'     PathFile                         Append file name to path      13.02.2019
'     PathIsAbs                        Is path absolute?g
'     PathLoc1Dr                       Convert one drive to local    19.05.2019
' --- The private sphere

Public Function FileExists(sFile As String) As Boolean

' checks if a file exists in a folder using the VBA Dir function
' Returns true if the specified file exists
' Returns true if a file exists matching a wildcard pattern
' Use of the VBA Dir function:
' Dir("path\file") returns "file" if the file exists
'                  returns "" if the file exists not
' Booby trap: A path without a file ("path\") will also return
'             a non-empty string. This case must be handled apart

   Dim sPS As String

   If sFile = "" Then Exit Function    ' --- no name, no check

   sPS = Application.PathSeparator     ' --- no file, no check
   If Right(sFile, Len(sPS)) = sPS Then Exit Function

   FileExists = Not Dir(sFile) = ""    ' --- file found

End Function

Public Function FileDelete(sFile As String) As Boolean

' deletes a file if it exists

   If FileExists(sFile) Then Kill sFile

End Function

Public Function DrivExists(sDrv As String) As Boolean

' Checks if a volume (or logical drive) exists
' Returns true if the specified volume exists
' Use of VBA Dir function:
' Dir(Drive) returns a non-empty string
' if the specified volume exists in the system

   If sDrv = "" Then Exit Function

   DrivExists = Not Dir(sDrv, vbVolume) = ""

End Function

Public Function PathFile(sPath As String, sFile As String) As String

' Concatenates a file name to a path, handling the path separator

   Dim sFull As String, sPS As String

   If sPath = "" Then                  ' --- No path
      sFull = sFile                    '     just file
   Else                                ' --- Path given
      sPS = Application.PathSeparator  '     ends with separator
      If Right(sPath, Len(sPS)) = sPS Then
         sFull = sPath & sFile
      Else                             '     no ending separator
         If sFile = "" Then            '     file name  empty
            sFull = sPath              '     -> use just path
         Else                          '     file name not empty
            sFull = sPath & sPS & sFile '    -> insert
         End If
      End If
   End If
   PathFile = sFull

End Function

Public Function FileList(sFile As String, sList() As String) As Long

' Returns the number of files with names matching the search pattern
' The list of file names is returned in sList()
' The list is not ordered

   Const CsDel As String = ";"

   Dim sDir As String, sLst As String
   Dim lN As Long
   Dim b1 As Boolean

   b1 = True
   sDir = Dir(sFile, vbDirectory)
   While Not sDir = ""
      lN = lN + 1
      If b1 Then sLst = sDir Else sLst = sLst & CsDel & sDir
      b1 = False
      sDir = Dir()
   Wend
   If lN > 0 Then
      QRS_LibStr.LstStrLst sLst, sList(), CsDel
   End If
   FileList = lN

End Function

Public Function PathIsAbs(sPath As String) As Boolean

' Returns true if a path specified is absolute
' Absolute path begins with the path delimiter
' or a drive letter followed by column or UNC \\

   Static sSep As String
   Static lSep As Long
   Dim bAbs As Boolean

   If lSep = 0 Then
      sSep = Application.PathSeparator
      lSep = Len(sSep)
   End If

   bAbs = Left(sPath, lSep) = sSep
   If Not bAbs Then bAbs = Mid(sPath, 2, 1) = ":"
   If Not bAbs Then bAbs = Left(sPath, 2) = "\\"
   If Not bAbs Then bAbs = Left(sPath, 2) = "//"
   
   PathIsAbs = bAbs

End Function

Private Function PathLoc1Dr(sFP As String) As String

' return the local path for doc, which is either already local or on OneDrive
' Credits: social.msdn

   Const CsPreFix1Dr As String = "https://d.docs.live.net/"

   Dim sLoc As String, sSep As String
   Dim lLen As Long

   lLen = Len(CsPreFix1Dr)             ' --- yep, Path is on OneDrive
   If LCase(Left(sFP, lLen)) = CsPreFix1Dr Then
                                       ' --- locate end of "remote" part
      lLen = InStr(lLen + 1, sFP, "/")
      sLoc = Environ("OneDrive")       ' --- Get local designation for OneDrive
      sLoc = sLoc & Mid(sFP, lLen)     ' --- Get "local" part
      sSep = Application.PathSeparator
      sLoc = Replace(sLoc, "/", sSep)  ' --- All slashes back
      sLoc = Replace(sLoc, "%20", " ") ' --- White space for VBA
   Else
      sLoc = sFP
   End If

   PathLoc1Dr = sLoc
 
End Function
