Attribute VB_Name = "gcUsuariosPDC"
Option Explicit

Public Type JoinLong
   x As Long
   Dummy As Integer
End Type

Public Type JoinInt
   Bottom As Integer
   Top As Integer
   Dummy As Integer
End Type

Public Type NETRESOURCE
        dwScope As Long
        dwType As Long
        dwDisplayType As Long
        dwUsage As Long
        lpLocalName As Long
        lpRemoteName As Long
        lpComment As Long
        lpProvider As Long
End Type

Private Declare Function NetQueryDisplayInformation Lib "netapi32.dll" (ByRef servername As Byte, ByVal level As Long, ByVal Index As Long, ByVal EntriesRequested As Long, ByVal PreferredMaximumLength As Long, ByRef ReturnedEntryCount As Long, ByRef SortedBuffer As Long) As Long
Private Declare Function PtrToInt Lib "kernel32" Alias "lstrcpynW" (RetVal As Any, ByVal Ptr As Long, ByVal nCharCount As Long) As Long
Private Declare Function PtrToStr Lib "kernel32" Alias "lstrcpyW" (RetVal As Byte, ByVal Ptr As Long) As Long
Private Declare Function StrLen Lib "kernel32" Alias "lstrlenW" (ByVal Ptr As Long) As Long
Private Declare Function NetAPIBufferFree Lib "netapi32.dll" Alias "NetApiBufferFree" (ByVal Ptr As Long) As Long

Public Users$()
Public typRootResourses() As NETRESOURCE
Public typDomainResourses() As NETRESOURCE

Public Sub GetUsers(ByVal Domain As String)
    On Error Resume Next
    Dim SNArray() As Byte, level&, Index&, EntriesRequested&, _
    PreferredMaximumLength&, ReturnedEntryCount&, SortedBuffer&, _
    APIResult As Long, StrArray(500) As Byte, I&, TempPtr As JoinLong, _
    TempStr As JoinInt, data$(), Result&, Size%
    Let level = 1
    Let SNArray = Domain & vbNullChar
    Let Index = 0
    Let EntriesRequested = 500
    Let PreferredMaximumLength = 6000
    Do
        DoEvents
        APIResult = NetQueryDisplayInformation(SNArray(0), level, Index, _
        EntriesRequested, PreferredMaximumLength, ReturnedEntryCount, _
        SortedBuffer)
        If ReturnedEntryCount = 0 Then Exit Do
        For I = 1 To ReturnedEntryCount
            Let Size = Size + 1
            APIResult = PtrToInt(TempStr.Bottom, SortedBuffer + (I - 1) * 24, 2)
            APIResult = PtrToInt(TempStr.Top, SortedBuffer + (I - 1) * 24 + 2, 2)
            LSet TempPtr = TempStr
            APIResult = PtrToStr(StrArray(0), TempPtr.x)
            ReDim Preserve Users$(1 To Size)
            Users(Size) = Left(StrArray, StrLen(TempPtr.x))
            APIResult = PtrToInt(TempStr.Bottom, SortedBuffer + (I - 1) * 24 + 20, 2)
            APIResult = PtrToInt(TempStr.Top, SortedBuffer + (I - 1) * 24 + 22, 2)
            LSet TempPtr = TempStr
            Index = TempPtr.x
            DoEvents
        Next I
        Result = NetAPIBufferFree(SortedBuffer)
    Loop Until APIResult = 0
End Sub

Public Function NombreCompleto(ByVal sUsuario As String, ByVal psDominio As String)
    'Dim User As IADsUser
    'Dim container As IADsContainer
    'Set container = GetObject(gcWINNT & psDominio)
    'Set User = GetObject(gcWINNT & psDominio & "/" & sUsuario & ",user")
    'NombreCompleto = Trim(User.FullName)
    'Set container = Nothing
    'Set User = Nothing
End Function

'Public Function EnumUsers(ByVal Server As String) As Variant
'  Dim Users() As String
'  Dim Result As Long
'  Dim BufPtr As Long
'  Dim EntriesRead As Long
'  Dim TotalEntries As Long
'  Dim ResumeHandle As Long
'  Dim BufLen As Long
'  Dim SNArray() As Byte
'  Dim GNArray() As Byte
'  Dim UNArray(99) As Byte
'  Dim UName As String
'  Dim I As Integer
'  Dim UNPtr As Lon
'  Dim TempPtr As MungeLong
'  Dim TempStr As MungeInt
'  Dim Pass As Long
'
'  If Server <> "" And Left(Server, 2) <> "\\" Then Server = "\\" & Server
'
'  SNArray = Server & vbNullChar       ' Move to byte array
'  BufLen = 255                       ' Buffer size
'  ResumeHandle = 0                   ' Start with the first entry
'
'  Pass = 0
'  Do
'    Result = NetUserEnum(SNArray(0), 0, FILTER_NORMAL_ACCOUNT, BufPtr, BufLen, EntriesRead, TotalEntries, ResumeHandle)
'
'    If Result <> 0 And Result <> 234 Then    ' 234 means multiple reads required
'      Err.Raise Result, "EnumUsers", "Error enumerating user " & EntriesRead & " of " & TotalEntries
'      Exit Function
'    End If
'    For I = 1 To EntriesRead
'      ' Get pointer to string from beginning of buffer
'      ' Copy 4-byte block of memory in 2 steps
'      Result = PtrToInt(TempStr.XLo, BufPtr + (I - 1) * 4, 2)
'      Result = PtrToInt(TempStr.XHi, BufPtr + (I - 1) * 4 + 2, 2)
'      LSet TempPtr = TempStr ' munge 2 integers into a Long
'      ' Copy string to array
'      Result = PtrToStr(UNArray(0), TempPtr.x)
'      UName = Left(UNArray, StrLenW(TempPtr.x))
'      ReDim Preserve Users(0 To Pass)
'      Users(Pass) = UName
'      Pass = Pass + 1
'    Next I
'  Loop Until EntriesRead = TotalEntries
'  Result = NetAPIBufferFree(BufPtr)         ' Don't leak memory
'  EnumUsers = Users()
'End Function
'
