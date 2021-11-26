VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CdlgExplorer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Option Explicit

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

' ------------- Dialog calling functions
Private Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long

' --------------- Enum Flags
Public Enum CdlgExt_Flags
  OFNAllowMultiselect = &H200
  OFNCreatePrompt = &H2000
  OFNexplorer = &H80000
  OFNEnableHook = &H20
  OFNExtensionDifferent = &H400
  OFNFileMustExist = &H1000
  OFNHelpButton = &H10
  OFNHideReadOnly = &H4
  OFNLongNames = &H200000
  OFNNoChangeDir = &H8
  OFNNoDereferenceLinks = &H100000
  OFNNoLongNames = &H40000
  OFNNoReadOnlyReturn = &H8000
  OFNNoValidate = &H100
  OFNOverwritePrompt = &H2
  OFNPathMustExist = &H800
  OFNReadOnly = &H1
  OFNShareAware = &H4000
End Enum

Private RetValue As Long
Private OFN As OPENFILENAME

Private Const MAX_PATH = 260
Private Const MAX_FILE = 260

'Inner variables for properties
Private mFileName As String
Private mFileTitle As String
Private mhOwner As Long
Private mDialogTitle As String
Private mFilter As String
Private mInitDir As String
Private mSelDir As String
Private mDefaultExt As String
Private mFilterIndex As Long
Private mFlags As CdlgExt_Flags
Private mCancelError As Boolean
Private mLeft As Long
Private mTop As Long
Private mWidth As Long
Private mHeight As Long
Private mTipoVentana As Integer
Public Property Let TipoVentana(ByVal vData As Integer)
   mTipoVentana = vData
End Property

Public Property Get TipoVentana() As Integer
  TipoVentana = mTipoVentana
End Property

Public Property Let Height(ByVal vData As Long)
   mHeight = vData
End Property

Public Property Get Height() As Long
  Height = mHeight
End Property

Public Property Let Width(ByVal vData As Long)
   mWidth = vData
End Property

Public Property Get Width() As Long
  Width = mWidth
End Property

' Let/Get Properties
Public Property Let CancelError(ByVal vData As Boolean)
   mCancelError = vData
End Property

Public Property Get CancelError() As Boolean
  CancelError = mCancelError
End Property

Public Property Get hOwner() As Long
    hOwner = mhOwner
End Property

Public Property Let hOwner(ByVal New_hOwner As Long)
    mhOwner = New_hOwner
End Property
'Utilizado por la Inicializacion de InitOFN
Public Property Get flags() As CdlgExt_Flags
    flags = mFlags
End Property
'Utilizado por la Inicializacion de InitOFN
Public Property Let flags(ByVal New_Flags As CdlgExt_Flags)
    mFlags = New_Flags
End Property

Public Property Get DialogTitle() As String
   DialogTitle = mDialogTitle
End Property

Public Property Let DialogTitle(sTitle As String)
   mDialogTitle = sTitle
End Property

Public Property Get DefaultExt() As String
   DefaultExt = mDefaultExt
End Property

Public Property Let DefaultExt(sDefExt As String)
   mDefaultExt = DefaultExt
End Property

Public Property Get FileName() As String
   FileName = mFileName
End Property

Public Property Let FileName(sFileName As String)
   mFileName = sFileName
End Property

Public Property Get FileTitle() As String
   FileTitle = mFileTitle
End Property

Public Property Let FileTitle(sTitle As String)
   mFileTitle = sTitle
End Property

Public Property Get Filter() As String
   Filter = mFilter
End Property

Public Property Let Filter(sFilter As String)
   mFilter = sFilter
End Property

Public Property Get FilterIndex() As Long
   FilterIndex = mFilterIndex
End Property

Public Property Let FilterIndex(lIndex As Long)
    mFilterIndex = lIndex
End Property

Public Property Get InitDir() As String
   InitDir = mInitDir
End Property

Public Property Let InitDir(sDir As String)
    mInitDir = sDir
End Property
'Utilizado por MoveDlg
Public Property Get Left() As Long
   Left = mLeft
End Property

Public Property Let Left(lLeft As Long)
    mLeft = lLeft
End Property
'Utilizado por MoveDlg
Public Property Get Top() As Long
   Top = mTop
End Property

Public Property Let Top(lTop As Long)
    mTop = lTop
End Property

' Standard  Dialogs
Public Sub ShowOpen()
  Dim iDelim As Integer
  InitOFN
  RetValue = GetOpenFileName(OFN)
  If RetValue > 0 Then
     iDelim = InStr(OFN.lpstrFileTitle, vbNullChar)
     If iDelim Then mFileTitle = Left$(OFN.lpstrFileTitle, iDelim - 1)
     iDelim = InStr(OFN.lpstrFile, vbNullChar)
     If iDelim Then mFileName = Left$(OFN.lpstrFile, iDelim - 1)
  Else
     If mCancelError Then Err.Raise 0
  End If
  
End Sub

Private Sub InitOFN()
'  Dim sTemp As String, i As Integer
'  Dim uFlag As Long
'  CDlg.flags = OFNHelpButton
'  uFlag = mFlags Or OFNEnableHook Or OFNexplorer Or OFNFileMustExist Or OFNHideReadOnly Or OFNPathMustExist
'  With OFN
'       .lStructSize = Len(OFN)
'       .hwndOwner = mhOwner
'       .lpfnHook = lHookAddress(AddressOf DialogHookFunction)
'       .flags = uFlag
'       .lpstrDefExt = mDefaultExt
'       sTemp = mInitDir
'       If sTemp = "" Then sTemp = App.path
'       .lpstrInitialDir = sTemp
'       sTemp = mFileName
'       .lpstrFile = sTemp & String$(MAX_PATH - Len(sTemp), 0)
'       .nMaxFile = MAX_FILE
'       .lpstrFileTitle = String$(MAX_PATH, 0)
'       .nMaxFileTitle = MAX_FILE
'        sTemp = mFilter
'        For i = 1 To Len(sTemp)
'            If Mid(sTemp, i, 1) = "|" Then
'               Mid(sTemp, i, 1) = vbNullChar
'            End If
'        Next
'        sTemp = sTemp & String$(2, 0)
'        .lpstrFilter = sTemp
'        .nFilterIndex = mFilterIndex
'        .lpstrTitle = mDialogTitle
'        .hInstance = App.hInstance
' End With
End Sub

