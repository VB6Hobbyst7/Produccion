VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmHerActualizaSicmact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualiza Sicmact"
   ClientHeight    =   2595
   ClientLeft      =   1695
   ClientTop       =   1515
   ClientWidth     =   7320
   Icon            =   "frmHerActualizaSicmact.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   2595
   ScaleWidth      =   7320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraActualiza 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   30
      TabIndex        =   0
      Top             =   -30
      Width           =   7260
      Begin VB.CommandButton Browsedestination 
         Caption         =   "..."
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6825
         TabIndex        =   4
         Top             =   1095
         Width           =   315
      End
      Begin VB.TextBox Destinationpath 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   1080
         Width           =   7020
      End
      Begin VB.CommandButton Browsefile 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6825
         TabIndex        =   2
         Top             =   495
         Width           =   315
      End
      Begin VB.TextBox Filepath 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   7020
      End
      Begin MSComDlg.CommonDialog Dialog 
         Left            =   105
         Top             =   1455
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   6148
      End
      Begin VB.Label Label2 
         Caption         =   "Destino del Ejecutable"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   3555
      End
      Begin VB.Label Label1 
         Caption         =   "Origen Ultima  Actualizacion"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3165
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   1110
      Left            =   30
      TabIndex        =   7
      Top             =   1425
      Width           =   7260
      Begin MSComctlLib.ProgressBar ProgressBar 
         Height          =   360
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   635
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "&Actualizar"
         Height          =   375
         Left            =   2490
         TabIndex        =   9
         Top             =   645
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancela 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   3645
         TabIndex        =   8
         Top             =   645
         Width           =   1095
      End
      Begin VB.Label lblMensaje 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   3975
      End
   End
End
Attribute VB_Name = "frmHerActualizaSicmact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Private Declare Function SHGetPathFromIDList Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Declare Function SHGetSpecialFolderLocation Lib "shell32.dll" (ByVal hwndOwner As Long, ByVal nFolder As Long, pidl As ITEMIDLIST) As Long
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (lpBrowseInfo As BROWSEINFO) As Long

Private Type SHITEMID
     cb As Long
     abID As Byte
End Type

Private Type ITEMIDLIST
     mkid As SHITEMID
End Type

Private Type BROWSEINFO
     hOwner As Long
     pidlRoot As Long
     pszDisplayName As String
     lpszTitle As String
     ulFlags As Long
     lpfn As Long
     lParam As Long
     iImage As Long
End Type

Private Const NOERROR = 0

Private Const BIF_RETURNONLYFSDIRS = &H1
Private Const BIF_DONTGOBELOWDOMAIN = &H2
Private Const BIF_STATUSTEXT = &H4
Private Const BIF_RETURNFSANCESTORS = &H8
Private Const BIF_BROWSEFORCOMPUTER = &H1000
Private Const BIF_BROWSEFORPRINTER = &H2000

Dim fsFechaHora As String
Dim fsRutaUltActualiz As String
Dim fsRutaSICMACT As String
Public lbActivate As Boolean


'Realiza el copiado
Private Function CopyFile(Src As String, Dst As String) As Long
Static Buf$
Dim BTest!, FSize!
Dim Chunk%, F1%, F2%
Const BUFSIZE = 1024 '  set the buffer size

'If Len(Dir(Dst)) Then ' Verifica si existe el archivo
   'Response = MsgBox(Dst + oImpresora.gPrnSaltoLinea  + oImpresora.gPrnSaltoLinea  + "File already exists. Do you want to overwrite it?", vbYesNo + vbQuestion) 'prompt the user with a message box
   'If Response = vbNo Then 'if the "No" button was clicked
   '   Exit Function 'exit the procedure
   'Else             'otherwise
   '   Kill Dst      'delete the already found file, and carryon with the code
   'End If
'End If
On Error GoTo FileCopyError
F1 = FreeFile                'returns file number available
Open Src For Binary As F1    'open the source file
F2 = FreeFile                'returns file number available
Open Dst For Binary As F2    'open the destination file
FSize = LOF(F1)
BTest = FSize - LOF(F2)
Do
    DoEvents
    If BTest < BUFSIZE Then
       Chunk = BTest
    Else
       Chunk = BUFSIZE
    End If
    Buf = String(Chunk, " ")
    Get F1, , Buf
    Put F2, , Buf
    BTest = FSize - LOF(F2)
    ProgressBar.value = (100 - Int(100 * BTest / FSize)) 'advance the progress bar as the file is copied
Loop Until BTest = 0
Close F1
Close F2
CopyFile = 0
ProgressBar.value = 0  ' returns the progress bar to zero
Exit Function

FileCopyError:
    MsgBox "Error en la Actualizacion, Consulte con el Area de Sistemas ..."
    Close F1        'closes the source file
    Close F2        'closes the destination file
    Exit Function   'exit the procedure
End Function

'This code is used to extract the filename provided by the user from the
'Source text box. The filename is extracted and passed to the string
'SpecOut. Once the filename is extraced from the text box, it is then added
'to the destination path provided by the user.
Public Function ExtractName(SpecIn As String) As String
Dim i As Integer 'declare the needed variables
Dim SpecOut As String
On Error Resume Next 'ignore any errors
For i = Len(SpecIn) To 1 Step -1 ' assume what follows the last backslash is the file to be extracted
If Mid(SpecIn, i, 1) = "\" Then
   SpecOut = Mid(SpecIn, i + 1) 'extract the filename from the path provided
   Exit For
End If
Next i
ExtractName = SpecOut 'returns the extracted filename from the path
End Function

Private Sub Browsedestination_Click()
Dim bi As BROWSEINFO 'declare the needed variables
Dim rtn&, pidl&, path$, Pos%
Dim T As Long
Dim SpecIn As String, SpecOut As String
bi.hOwner = Me.hwnd 'centres the dialog on the screen
bi.lpszTitle = "Browse for Destination..." 'set the title text
bi.ulFlags = BIF_RETURNONLYFSDIRS 'the type of folder(s) to return
pidl& = SHBrowseForFolder(bi) 'show the dialog box
path = Space(512) 'sets the maximum characters
T = SHGetPathFromIDList(ByVal pidl&, ByVal path) 'gets the selected path
Pos% = InStr(path$, Chr$(0)) 'extracts the path from the string
SpecIn = Left(path$, Pos - 1) 'sets the extracted path to SpecIn
If Right$(SpecIn, 1) = "\" Then 'makes sure that "\" is at the end of the path
   SpecOut = SpecIn             'if so then, do nothing
Else                            'otherwise
   SpecOut = SpecIn + "\"       'add the "\" to the end of the path
End If
Destinationpath.Text = SpecOut + ExtractName(Filepath.Text) 'merges both the destination path and the source filename into one string
End Sub

Private Sub Browsefile_Click()
Dialog.DialogTitle = "Browse for source..." 'set the dialog title
Dialog.ShowOpen 'show the dialog box
Filepath.Text = Dialog.FileName 'set the target text box to the file chosen
End Sub

Private Sub cmdActualiza_Click()
    Call ACTUALIZA
End Sub

Public Sub ACTUALIZA()
Dim lsNombre As String
Dim lsNuevoNombre As String
Dim lnAplicacion As Integer
cmdActualiza.Enabled = False
fraActualiza.Enabled = False

Dim fs As Scripting.FileSystemObject
Dim fCurrent As Scripting.Folder
Dim fi As Scripting.File
Dim fd As Scripting.File

Dim lsRutaUltActualiz As String
Dim lsRutaSICMACT As String
Dim lsFecUltModifLOCAL As String
Dim lsFecUltModifORIGEN As String
Dim ActualizaUltVersionEXE As Boolean
    
On Error GoTo ErrActualiza
    Set fs = New Scripting.FileSystemObject
    Set fCurrent = fs.GetFolder(fsRutaUltActualiz)
    For Each fi In fCurrent.Files
          If Right(UCase(fi.Name), 3) = "EXE" Or Right(UCase(fi.Name), 3) = "INI" Or Right(UCase(fi.Name), 3) = "DLL" Then
             lsFecUltModifORIGEN = Format(fi.DateLastModified, "yyyy/mm/dd hh:mm:ss")
             ActualizaUltVersionEXE = False
             If Dir(fsRutaSICMACT & fi.Name) <> "" Then
                Set fd = fs.GetFile(fsRutaSICMACT & fi.Name)
                lsFecUltModifLOCAL = Format(fd.DateLastModified, "yyyy/mm/dd hh:mm:ss")
                If lsFecUltModifLOCAL < lsFecUltModifORIGEN And lsFecUltModifORIGEN <> "" Then ' ACTUALIZA
                    ActualizaUltVersionEXE = True
                End If
             Else
                ActualizaUltVersionEXE = True
             End If
             If ActualizaUltVersionEXE = True Then
                'Borra SicmactANT.Exe si lo encuentra
                Destinationpath.Text = fsRutaSICMACT & fi.Name
                Filepath.Text = fsRutaUltActualiz & fi.Name
                
                lsNuevoNombre = Mid(Destinationpath.Text, 1, Len(Trim(Destinationpath.Text)) - 4) + "ANT" + Right(fi.Name, 4)
                If Len(Dir(lsNuevoNombre)) Then ' Verifica si existe el archivo
                    Kill lsNuevoNombre
                End If
                'Renonbra el Sicmact.Exe que existe
                If Len(Dir(Destinationpath.Text)) Then ' Verifica si existe el archivo
                    lsNombre = Destinationpath.Text
                    lsNuevoNombre = Mid(Destinationpath.Text, 1, Len(Trim(Destinationpath.Text)) - 4) + "ANT" + Right(fi.Name, 4)
                    Name lsNombre As lsNuevoNombre
                End If
                
                lblMensaje = "Actualizando " & fi.Name & " ..."
                '*** Copia el Archivo
                ProgressBar.value = CopyFile(Filepath.Text, Destinationpath.Text)
                '***
                lblMensaje = ""
             End If
          End If
          If Right(UCase(fi.Name), 3) = "DLL" Then
             lsFecUltModifORIGEN = Format(fi.DateLastModified, "yyyy/mm/dd hh:mm:ss")
             If Dir(fsRutaSICMACT & fi.Name) <> "" Then
                Set fd = fs.GetFile(fsRutaSICMACT & "\" & fi.Name)
                lsFecUltModifLOCAL = Format(fi.DateLastModified, "yyyy/mm/dd hh:mm:ss")
                If lsFecUltModifLOCAL < lsFecUltModifORIGEN And lsFecUltModifORIGEN <> "" Then ' ACTUALIZA
                    ActualizaUltVersionEXE = True
                End If
             Else
                ActualizaUltVersionEXE = True
             End If
             If ActualizaUltVersionEXE = True Then
                ' Registra DLL
                Dim retval
                '*** Copia el Archivo
                Filepath.Text = fsRutaUltActualiz & fi.Name
                Destinationpath.Text = fsRutaSICMACT & fi.Name
                ProgressBar.value = CopyFile(Filepath.Text, Destinationpath.Text)
                '**** Lo registra
                retval = Shell("regsvr32 " & fi.Name)
                '**************
             End If
          End If
    Next

Unload Me
MsgBox " SE HA ACTUALIZADO EL SICMACT " & oImpresora.gPrnSaltoLinea & " Debe volver a Ingresar al Sicmact", vbInformation, "Aviso"
End
Exit Sub
ErrActualiza:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub
Private Sub cmdCancela_Click()
    Unload Me
End Sub

Private Sub filepath_Change()
    Destinationpath.Enabled = True
    Browsedestination.Enabled = True
    'Destinationpath.SetFocus
End Sub

Private Sub Form_Activate()
If lbActivate Then
    cmdActualiza_Click
End If
End Sub

Private Sub Form_Load()
On Error GoTo ErrLoad
    'fsFechaHora = Mid(gdFecSis, 7, 4) & Mid(gdFecSis, 4, 2) & Mid(gdFecSis, 1, 2) & Mid(Time(), 1, 2) & Mid(Time(), 4, 2)
    Dim oCons As NConstSistemas
    Set oCons = New NConstSistemas
    
    CentraForm Me
    fsRutaUltActualiz = oCons.GetRutaAcceso(gsCodAge)
    fsRutaSICMACT = App.path & "\"
    Filepath.Text = fsRutaUltActualiz
    Destinationpath.Text = fsRutaSICMACT
    Exit Sub
ErrLoad:
    MsgBox Err.Description, vbInformation, "!Aviso!"
End Sub

Public Sub IniciaVariables(ByVal pbejecutaActualiza As Boolean)
    lbActivate = pbejecutaActualiza
End Sub
