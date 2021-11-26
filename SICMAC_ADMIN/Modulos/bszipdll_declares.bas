Attribute VB_Name = "BigSpeed_Zip_DLL_API_Declares"
'Originated: Richard Tainsh, Linhoff March Ltd, 17/2/99
Option Explicit
DefInt A-Z  'Use it to allow faster running application.
            'be aware hovewer of what it does, by understanding
            'his function to avoid compilation problem
'
'-----VB declarations for BigSpeed Zip DLL.  Written for VB5/6
'

'-----Windows API functions needed for StringFromPointer routine
Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal StringPointer As Long) As Long
Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, ByVal StringPointer As Long) As Long

'-----BigSpeed zip dll API functions
'     Note: For arguments, True = 1 , False = 0.

'.....This function opens zip file and creates a list of items, corresponding to
'     contents of zip file. Must be called before any other function to get access
'     to zip file. ZipFileName is the path and full name of zip file,
'     for example: 'C:\TEST.ZIP'. Returns zero on success.
Declare Function zOpenZipFile Lib "bszip.dll" (ByVal zipname As String) As Long

'.....Closes zip file, opened with zOpenZipFile. Returns zero on success.
Declare Function zCloseZipFile Lib "bszip.dll" () As Long

'.....Returns the total number of compressed files within zip file, i.e. the number
'     of items in the list. First item has index = 0, last item has '
'     Index = zGetTotalFiles - 1
Declare Function zGetTotalFiles Lib "bszip.dll" () As Long


'.....Returns the sum of uncompressed file size of all files within zip file.
Declare Function zGetTotalBytes Lib "bszip.dll" () As Long

'.....Returns the number of currently selected files.
Declare Function zGetSelectedFiles Lib "bszip.dll" () As Long

'.....Returns the sum of uncompressed file size of selected files within zip file.
Declare Function zGetSelectedBytes Lib "bszip.dll" () As Long

'.....If some error occurs in the DLL, this function will return a pointer to a string
'     that describes the error.
'     MUST be wrapped with StringFromPointer to obtain the VB String
Declare Function zGetLastErrorAsText Lib "bszip.dll" () As Long

'.....After extracting operation, this function will return the number of skipped
'     files - usually zero.
Declare Function zGetSkippedFiles Lib "bszip.dll" () As Long

'.....This function returns the number of currently processed files and bytes.
'     This is necessary to calculate position of progress indicator.
'     Returns 1 on success.
Declare Function zGetRunTimeInfo Lib "bszip.dll" (ByRef ProcessedFiles As Long, ByRef ProcessedBytes As Long) As Byte

'.....To cancel extract operation, call this function.
Declare Function zCancelOperation Lib "bszip.dll" () As Byte

'.....Extracts only one of the file in zip file. This is useful for viewing/launching files.
Declare Function zExtractOne Lib "bszip.dll" (ByVal ItemNo As Long, _
   ByVal ExtractDirectory As String, ByVal Password As String, _
   ByVal OverwriteExisting As Byte, ByVal SkipOlder As Byte, _
   ByVal UseFolders As Byte, ByVal TestOnly As Byte, _
   ByVal RTInfoFunc As Long) As Long
'     Parameters are:
'     ItemNo -  Item number of the file in the list.
'          0<= ItemNo <= zGetTotalFile-1
'     ExtractDirectory - Directory where to extract files. If empty, current directory is used.
'     Password - password to be used, if file is encrypted.
'     OverwriteExisting - if 1, any existing file will be overwritten, otherwise will be skipped.
'     UseFolders - if 1, the relative path information stored in zip file will be used to determine
'          location of the extracting file.
'     TestOnly - if 1, files are only tested, not saved to disk.
'     RTInfoFunc - pointer to the application function, which will be called
'          periodically from library to show runtime information,
'          for example some kind of progress bar. Within this function, application can
'          call zGetRunTimeInfo or zCancelOperation. If nil, no runtime information will
'          be available.
'          In VB5+, this may be called by using AddressOf in the call, in VB4 set to nil.
'
'     Return value is zero on success, otherwise application can call zGetLastErrorAsText to get
'     type of error.


'.....Extracts only files in the list, previously selected by zSelectFile.
'     Parameters are the same as in zExtractOne. Returns zero on success.
Declare Function zExtractSelected Lib "bszip.dll" ( _
    ByVal ExtractDirectory As String, ByVal Password As String, _
   ByVal OverwriteExisting As Byte, ByVal SkipOlder As Byte, _
   ByVal UseFolders As Byte, ByVal TestOnly As Byte, _
   ByVal RTInfoFunc As Long) As Long


'.....Extracts all files. Parameters are the same as in zExtractOne.
'     Returns zero on success.
Declare Function zExtractAll Lib "bszip.dll" ( _
    ByVal ExtractDirectory As String, ByVal Password As String, _
   ByVal OverwriteExisting As Byte, ByVal SkipOlder As Byte, _
   ByVal UseFolders As Byte, ByVal TestOnly As Byte, _
   ByVal RTInfoFunc As Long) As Long


'.....Returns a pointer to the name of the file in the list with index i.
'     MUST be wrapped with StringFromPointer to obtain the VB String
Declare Function zGetFileName Lib "bszip.dll" (ByVal i As Long) As Long


'.....Returns a pointer to the extension of the file in the list with index i.
'     MUST be wrapped with StringFromPointer to obtain the VB String
Declare Function zGetFileExt Lib "bszip.dll" (ByVal i As Long) As Long

'.....Returns a pointer to the stored path of the file in the list with index i.
'     MUST be wrapped with StringFromPointer to obtain the VB String
Declare Function zGetFilePath Lib "bszip.dll" (ByVal i As Long) As Long

'.....Returns the MSDOS date of the file in the list with index i.
Declare Function zGetFileDate Lib "bszip.dll" (ByVal i As Long) As Long

'.....Returns the MSDOS time of the file in the list with index i.
Declare Function zGetFileTime Lib "bszip.dll" (ByVal i As Long) As Long

'.....Returns the uncompressed size of the file in the list with index i.
Declare Function zGetFileSize Lib "bszip.dll" (ByVal i As Long) As Long

'.....Returns the compressed size of the file in the list with index i.
Declare Function zGetCompressedFileSize Lib "bszip.dll" (ByVal i As Long) As Long

'.....Returns 1, if file in the list with index i is encrypted.
Declare Function zFileIsEncrypted Lib "bszip.dll" (ByVal i As Long) As Byte

'.....Returns a pointer to the result of the last operation on the file in the list with index i.
'     Usually 'Ok.'
'     MUST be wrapped with StringFromPointer to obtain the VB String
Declare Function zGetLastOperResult Lib "bszip.dll" (ByVal i As Long) As Long

'.....Returns 1, if file in the list with index i is selected.
Declare Function zFileIsSelected Lib "bszip.dll" (ByVal i As Long) As Byte

'.....Select/Unselect file in the list with index i. If how is 1, the file will be selected,
'     otherwise will be unselected.
Declare Function zSelectFile Lib "bszip.dll" (ByVal i As Long, ByVal how As Byte) As Byte

'.....To unselect all files in the list call this function
Declare Function zUnselectAll Lib "bszip.dll" () As Byte

'.....Selects files in the zip archive using wildcard pattern
Declare Function zSelectByWildcards Lib "bszip.dll" (ByVal FileMask As String, _
ByVal IncludeSubfolders As Byte) As Long
'     Parameters are:
'     FileMask- Wildcard pattern to be used, for example: "C:\TEST\*.TXT; *.DOC"
'     IncludeSubfolders- If True, the subfolders will be searched too for matched files.




'.....Creates new empty zip file witn name ZipFileName, for example: "C:\temp\test.zip".
'     Returns zero on success.
Declare Function zCreateNewZip Lib "bszip.dll" (ByVal zipname As String) As Long

'.....Call this function after zCreateNewZip or zOpenZipFile for every file, which must be included
'     in the Zip file. The function checks if the file already exists in the archive, date and time
'     of the file depending of the UpdateMode, and if the conditions are true, adds the file to the
'     list of "to be compressed files".
Declare Function zOrderFile Lib "bszip.dll" (ByVal FileName As String, ByVal StoredName As String, _
    ByVal UpdateMode As Long) As Long
'     Parameters are:
'     FileName - the full path and name of the file, for example: "C:\TEMP\TEST.TXT"
'     StoredName - this is the path and the name, under which the file must be stored in zip file,
'          for example: "TEMP\TEST.TXT".
'     UpdateMode - points if the file must be added if it is already exists in the archive. There
'          are 3 possible values:
'          0 - "Add (and Replace) file", if file does not exist in the archive, it will be added, if exists, it will be replaced.
'          1 - "Freshen file", file will be compressed only if it already exists in the archive and is newer.
'          2 - "Update file", file will be compressed if does not exists in the archive or if it exists, but is newer.
'     The return value is zero if the file is succesfuly added to "Add List".




'.....Selects files to be compressed into zip archive using wildcard pattern
Declare Function zOrderByWildcards Lib "bszip.dll" (ByVal FileMask As String, _
ByVal BasePath As String, _
ByVal IncludeSubfolders As Byte, _
ByVal CheckArchiveAttribute As Byte, _
ByVal UpdateMode As Long) As Long
'     Parameters are:
'  FileMask - The wildcard pattern to be used, for example: "C:\TEST\*.TXT; *.DOC"
'  BasePath - This is the shared part of the path to be excluded from all selected files, when they will be stored in the archive, for example: "C:\TEST\".
'  IncludeSubfolders - If True, the subfolders will be searched too for matched files.
'  CheckArchiveAttribute - If True, only those of the files will be added, which have set "Archive" attribute.
'  UpdateMode - Points if the file must be added if it already exists in the archive. There are 3 possible values:
'   0 - "Add (and Replace) file", if the file does not exist in the archive, it will be added, if exists, it will be replaced.
'   1 - "Freshen file", the file will be compressed only if it already exists in the archive and is newer.
'   2 - "Update file", the file will be compressed if it does not exist in the archive or if it exists, but is newer.
'Return: Zero on success, otherwise an error code.




'.....Call this function to actually compress the files, previously requested with zOrderFile/zOrderByWildcards.
Declare Function zCompressFiles Lib "bszip.dll" _
       (ByVal TempDir As String, _
        ByVal Password As String, _
        ByVal CompressionMethod As Long, _
        ByVal ResetArchiveAttribute As Byte, _
        ByVal SpanSize As Long, _
        ByVal Comment As String, _
        ByVal RTInfoFunc As Long) As Long
'     Parameters are:
'     TempDir - the temporary directory to be used
'     Password - if is not empty, the files will be encrypted with this password.
'     CompressionMethod - there 4 possible values:
'          0 : Stored, no compression
'          1:     Fast compression
'          2:     Normal compression
'          3:     Max compression
'     ResetArchiveAttribute - if it is 1, the Archive attribute of the file will be reset after the
'          compressing of the file.
'
'     SpanSize - There are 3 cases here:
'          < 0 : Disabled - spanning is not allowed.
'          = 0 : Automatic - automatically prompt for another diskette when the current one is full. Applicable only for removable drives.
'          > 0 : Custom - multi-volume zip archive will be created and this value will be the maximum size of each volume. Applicable only for non-removable drives.
'
'     Comment - if is not empty, will be stored as a comment in the zip file.
'
'     RTInfoFunc - pointer to the application function, which will be called periodically from
'          library to show runtime information, for example some kind of progress bar.
'          Within this function, application can call zGetRunTimeInfo or zCancelOperation.
'          If nil, no runtime information will be available.
'          In VB5+, this may be called by using AddressOf in the call, in VB4 set to nil.
'
'     Return value is zero on success, otherwise application can call zGetLastErrorAsText
'     to get type of error.
'
'     Do not forget to close it with zCloseZipFile.



'.....Removes the selected files from the zip archive. Returns zero on success, otherwise an error code.
Declare Function zDeleteFiles Lib "bszip.dll" () As Long


'.....Returns the number of files requested for compression with functions zOrderFile/zOrderByWildcards.
Declare Function zGetOrderedFiles Lib "bszip.dll" () As Long


'.....Returns the total size of files requested for compression with functions zOrderFile/zOrderByWildcards.
Declare Function zGetOrderedBytes Lib "bszip.dll" () As Long


'.....Returns a pointer to the stored comment in the zip file
'     MUST be wrapped with StringFromPointer to obtain the VB String
Declare Function zGetComment Lib "bszip.dll" () As Long


'.....To check if the opened zip file is spanned on many disks/volumes, call this function.
Declare Function zIsSpanned Lib "bszip.dll" () As Byte















Public Function StringFromPointer(ByVal Pointer As Long) As String
  '
  '-----Returns a VB String given a pointer to an LPSTR
  '
  Dim StrLen As Long
  Dim NewString As String
  
  '.....Get the length of the LPSTR
  StrLen = lstrlen(Pointer)
  '.....Allocate the NewString to the right size
  NewString = String(StrLen, " ")
  
  '.....Copy the LPSTR to the VB string
  Call lstrcpy(NewString, Pointer)
  
  StringFromPointer = NewString
End Function

Public Sub DoEndProgram()
  'This sub Was added by Patrice Villeneuve 14/11/2000
  'See VbZipCreate Demo
  
  'Call this sub to terminate the Program.
  'This Unload All form and then terminate the program.
  Dim xForm As Form
  
  'Unload all Forms
  For Each xForm In Forms
    Unload xForm
    Set xForm = Nothing
  Next xForm
End Sub

