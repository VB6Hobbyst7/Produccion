VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCCECargaOficinas 
   BackColor       =   &H80000016&
   Caption         =   "Carga de Oficinas de la CCE"
   ClientHeight    =   5805
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12630
   Icon            =   "frmCCECargaOficinas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5805
   ScaleWidth      =   12630
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11400
      TabIndex        =   11
      Top             =   5400
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   9960
      TabIndex        =   10
      Top             =   5400
      Width           =   1230
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "Cargar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   8640
      TabIndex        =   9
      Top             =   5400
      Width           =   1230
   End
   Begin VB.Frame fraArchivoLOG 
      BackColor       =   &H80000016&
      Caption         =   "Cámara de Compensación Electrónica"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4785
      Left            =   4920
      TabIndex        =   3
      Top             =   360
      Width           =   7575
      Begin VB.TextBox txtResultado 
         BackColor       =   &H00FFFFFF&
         DragIcon        =   "frmCCECargaOficinas.frx":030A
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Lucida Sans"
            Size            =   9
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   3495
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   1080
         Width           =   7095
      End
      Begin VB.TextBox txtCarpeta 
         Height          =   255
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   240
         Width           =   5280
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   285
         Left            =   6720
         MaskColor       =   &H00E0E0E0&
         TabIndex        =   5
         Top             =   240
         Width           =   465
      End
      Begin MSComctlLib.ProgressBar PbCCE 
         Height          =   255
         Left            =   1200
         TabIndex        =   4
         Top             =   600
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label lblRura 
         BackColor       =   &H80000016&
         Caption         =   "Ubicación :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000016&
         Caption         =   "Estado     :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         Left            =   120
         TabIndex        =   7
         Top             =   600
         Width           =   975
      End
   End
   Begin VB.Frame fraUbicacion 
      BackColor       =   &H80000016&
      Caption         =   "Elija ubicación de lectura de archivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4815
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   4455
      Begin VB.DirListBox dirCarpetas 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3690
         Left            =   240
         TabIndex        =   2
         Top             =   720
         Width           =   3495
      End
      Begin VB.DriveListBox drvUnidades 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   3495
      End
   End
   Begin MSComDlg.CommonDialog dlgArchivo 
      Left            =   5520
      Top             =   5280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmCCECargaOficinas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*******************************************************************************************************************
'** Nombre : frmCCECargaOficinas
'** Descripción : Para la Carga de Oficinas de la CCE   , Proyecto: Implementacion del Servicio de Compensaciòn Electrónica Diferido de Instrumentos Compensables CCE
'** Creación : VAPA, 20170621
'*******************************************************************************************************************
Option Explicit
Dim oCCE As COMNCajaGeneral.NCOMCCE
Dim oCCER As COMNCajaGeneral.NCOMCCE
Dim oCamara As COMNCajaGeneral.NCOMCCE
Dim rsConstante As ADODB.Recordset
Dim Data() As String
Dim NBdata As Integer
Public Function LeerCarpeta(ByVal Carp As String) As Integer
Dim obj, f, F1, Chem As String
Dim CarpP As Folder
Dim i As Integer, Ext As String
Dim Ruta As String
Dim T As Double

    Set obj = CreateObject("Scripting.FileSystemObject")
    Set CarpP = obj.GetFolder(Carp)
    Set f = CarpP.Files
    GoSub RellenarData
    
Exit Function
RellenarData:
    For Each F1 In f
        Ext = LCase(Right(F1.Name, 3))
        If Ext = "txt" Then
            ReDim Preserve Data(1, NBdata)
            Data(0, NBdata) = F1.Name
            NBdata = NBdata + 1
        End If
    Next F1
    
Return
End Function
Private Sub cmdBuscar_Click()
Dim i As Integer
Dim sTexto, lsRuta As String
Dim Linea, Dato, DataB() As String
Dim valor As Boolean
Dim FechaCreacion As Date
      txtResultado.Text = ""
      
    If txtCarpeta.Text = " " Then
    MsgBox "Debe Elegir una Ubicación Desde donde Cargar el Archivo de Oficinas", vbInformation, "Aviso"
    Exit Sub
    End If
   
    
Erase Data

    If Len(txtCarpeta) > 1 Then
        LeerCarpeta (txtCarpeta)
        i = 0
        
        If NBdata = 0 Then
        MsgBox "No Existe nigun archivo TXT en esta ubicación ", vbInformation, "Aviso"
        Exit Sub
        End If
         txtResultado.Text = txtResultado.Text & "El Archivo encontrado para cargar las Oficinas es : " & vbCrLf
         txtResultado.Text = txtResultado.Text & vbCrLf
            Do Until i = NBdata
                sTexto = Data(0, i)
                lsRuta = txtCarpeta.Text & "\" & sTexto
                i = i + 1
                If Not sTexto = "" Then
                
                If InStr(1, sTexto, "Ofic", vbTextCompare) Then
                            FechaCreacion = FileSystem.FileDateTime(lsRuta)
                            txtResultado.Text = txtResultado.Text & sTexto & " - Archivo de Oficinas de Cámara de Compensación Electrónica " & FechaCreacion & vbCrLf
                            cmdCargar.Enabled = True
                Else
                FechaCreacion = FileSystem.FileDateTime(lsRuta)
                txtResultado.Text = txtResultado.Text & sTexto & " !NO ES VALIDO! " & FechaCreacion & vbCrLf
                MsgBox "Debe ubicar solo el TXT de Oficinas de la CCE ", vbInformation, "Aviso"
                End If

                End If
                
            Loop
    Exit Sub
    Else
        txtResultado.Text = txtResultado.Text & _
        "No se encontraron archivos en la carpeta: " & vbCrLf
    Exit Sub
    End If
 End Sub
Private Sub LeerTramaRecibido()
Dim i As Integer
    If Len(txtCarpeta.Text) = 0 Then
        txtResultado.Text = txtResultado.Text & "No se seleccionó ninguna carpeta..." & vbCrLf
        Exit Sub
    Else
           Do Until i = NBdata
''Lectura de Oficinas*************************************************************************
        If InStr(1, Data(0, i), "Ofic", vbTextCompare) <> 0 Then
                LeerOficinasCCE (txtCarpeta.Text & "\" & (Data(0, i)))
        End If
                 i = i + 1
        Loop
    End If
End Sub
Private Sub LeerOficinasCCE(Optional ByRef psFile As String = "")
    Dim oFSO As Scripting.FileSystemObject
    Dim oStream As Scripting.TextStream
    Dim mvarRuta, sNombreArchivo, sRecibido, sId, sFechaArchivo, sCodAplicacion, sNumArchivo, sUsuario, sMoneda, Linea, Data, lsCabDataMN, lsCabDataME As String
    Dim nId, NumeroArchivo As Integer
    Dim lbCuentaSoles As Boolean
    Dim lvConfirmado() As String
    Dim lvDevuelto() As String
    Dim lnIDRef As Long
    Dim lsEfinCodBCR, lsOficCod, lsOficCodMismaPlaza, lsOficUbigPlazaComp, lsOficUbigPlazaCompTr, lsOficNombre, lsOficDomicilio, lsOficLocalidad As String
    Dim lsOficCodPostalDist, lsOficNombrePlaza, lsOficTipo, lsOficUbigINEI, lsOficPlazaExc, lsOficMotivoActualiza, lsFechaActualiza As String
    Dim lbArch As Boolean
    Dim sFechaMB As String
    Dim lnTotal As Integer
    Dim z As Long

On Error GoTo ErrorPreTransf
oCCE.BeginTrans
lbArch = True


        
    If Len(psFile) <> 0 Then mvarRuta = psFile
    If Len(mvarRuta) = 0 Then Exit Sub: oCCE.CommitTrans
    
    Set oFSO = New Scripting.FileSystemObject
    Set oStream = oFSO.OpenTextFile(mvarRuta, ForReading)
    sNombreArchivo = oFSO.GetFileName(mvarRuta)
    
    
    NumeroArchivo = FreeFile
   
   
    Do Until oStream.AtEndOfStream
                lnTotal = lnTotal + 1
                oStream.SkipLine
    Loop
    
      Me.PbCCE.Max = lnTotal
      Me.PbCCE.Min = 0
      Me.PbCCE.value = 0
  
  
    'Detalle ***********************************
    
    If MsgBox("¿Esta seguro de actualizar las Oficinas ?", vbYesNo + vbInformation, "Aviso") = vbNo Then Exit Sub
    
    oCCE.CCE_EliminaOficinas
    
    Open mvarRuta For Input As #NumeroArchivo
    'Replace(vCadena, "á", Chr(32), , , vbTextCompare) o"cus
                Do While Not EOF(NumeroArchivo)
                 Line Input #NumeroArchivo, Linea
                   Data = Linea
                   lsEfinCodBCR = Mid(Data, 3, 3)
                   lsOficCod = Mid(Data, 6, 3)
                   lsOficCodMismaPlaza = Mid(Data, 9, 4)
                   lsOficUbigPlazaComp = Mid(Data, 13, 6)
                   lsOficUbigPlazaCompTr = Mid(Data, 19, 6)
                   lsOficNombre = Mid(Data, 48, 35)
                   lsOficDomicilio = Mid(Data, 83, 80)
                   lsOficLocalidad = Mid(Data, 163, 35)
                   lsOficCodPostalDist = Mid(Data, 198, 4)
                   lsOficNombrePlaza = Mid(Data, 202, 35)
                   lsOficTipo = Mid(Data, 237, 1)
                   lsOficUbigINEI = Mid(Data, 238, 6)
                   lsOficPlazaExc = Mid(Data, 244, 1)
                   lsOficMotivoActualiza = Mid(Data, 287, 1)
                   lsFechaActualiza = Mid(Data, 288, 8)
                   
                   lsEfinCodBCR = Replace(lsEfinCodBCR, "'", Chr(32))
                   lsOficCod = Replace(lsOficCod, "'", Chr(32))
                   lsOficCodMismaPlaza = Replace(lsOficCodMismaPlaza, "'", Chr(32))
                   lsOficUbigPlazaComp = Replace(lsOficUbigPlazaComp, "'", Chr(32))
                   lsOficUbigPlazaCompTr = Replace(lsOficUbigPlazaCompTr, "'", Chr(32))
                   lsOficNombre = Replace(lsOficNombre, "'", Chr(32))
                   lsOficDomicilio = Replace(lsOficDomicilio, "'", Chr(32))
                   lsOficLocalidad = Replace(lsOficLocalidad, "'", Chr(32))
                   lsOficCodPostalDist = Replace(lsOficCodPostalDist, "'", Chr(32))
                   lsOficNombrePlaza = Replace(lsOficNombrePlaza, "'", Chr(32))
                   lsOficTipo = Replace(lsOficTipo, "'", Chr(32))
                   lsOficUbigINEI = Replace(lsOficUbigINEI, "'", Chr(32))
                   lsOficPlazaExc = Replace(lsOficPlazaExc, "'", Chr(32))
                   lsOficMotivoActualiza = Replace(lsOficMotivoActualiza, "'", Chr(32))
                   lsFechaActualiza = Replace(lsFechaActualiza, "'", Chr(32))
                   
                   
                   
                   oCCE.CCE_InsertaOficinas lsEfinCodBCR, lsOficCod, lsOficCodMismaPlaza, lsOficUbigPlazaComp, lsOficUbigPlazaCompTr, _
                    lsOficNombre, lsOficDomicilio, lsOficLocalidad, lsOficCodPostalDist, lsOficNombrePlaza, lsOficTipo, lsOficUbigINEI, _
                   lsOficPlazaExc, lsOficMotivoActualiza, lsFechaActualiza
                   
                   Me.PbCCE.value = Me.PbCCE.value + 1
                  
                  
                
                Loop
    Close #NumeroArchivo

  
      txtResultado.Text = txtResultado.Text & _
                        "Archivo Cargado: " & sNombreArchivo & vbCrLf & _
                        "Registros Cargados: " & lnTotal & vbCrLf
    cmdCargar.Enabled = False
    oCCE.CommitTrans
    lbArch = False
    Exit Sub

ErrorPreTransf:
    If lbArch Then
        oCCE.RollbackTrans
        Set oCCE = Nothing
    End If
    Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Private Sub cmdCargar_Click()
  LeerTramaRecibido
End Sub
Private Sub cmdLimpiar_Click()
    PbCCE.value = 0
    txtResultado.Refresh
    txtResultado.Text = ""
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub dirCarpetas_Change()
    txtCarpeta.Text = dirCarpetas.Path
End Sub
Private Sub drvUnidades_Change()
    dirCarpetas.Path = drvUnidades.Drive
End Sub
Private Sub Form_Load()
Dim i, sTexto
        txtResultado.Text = ""
            Set oCCE = New COMNCajaGeneral.NCOMCCE
            Set oCCER = New COMNCajaGeneral.NCOMCCE
            txtCarpeta.Text = " "
            cmdCargar.Enabled = False
End Sub


