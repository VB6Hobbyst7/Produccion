VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmBancoPagadorConfirmacionAbonos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Confirmación de Abonos"
   ClientHeight    =   7170
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10650
   Icon            =   "frmBancoPagadorConfirmacionAbonos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7170
   ScaleWidth      =   10650
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   765
      Left            =   120
      TabIndex        =   24
      Top             =   6360
      Width           =   10455
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   8040
         TabIndex        =   27
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   9240
         TabIndex        =   26
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdConfirmarAbonos 
         Caption         =   "&Confirmar Abonos"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Visible         =   0   'False
         Width           =   1545
      End
   End
   Begin VB.TextBox txtAbonosRealizados 
      BackColor       =   &H80000000&
      Height          =   300
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   17
      Top             =   5520
      Width           =   855
   End
   Begin VB.TextBox txtMTSolesRealizados 
      BackColor       =   &H80000000&
      Height          =   300
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   16
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtMTDolaresRealizados 
      BackColor       =   &H80000000&
      Height          =   300
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   15
      Top             =   5520
      Width           =   1455
   End
   Begin VB.TextBox txtAbonosRechazados 
      BackColor       =   &H80000000&
      Height          =   300
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   14
      Top             =   6000
      Width           =   855
   End
   Begin VB.TextBox txtMTSolesRechazados 
      BackColor       =   &H80000000&
      Height          =   300
      Left            =   5160
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   6000
      Width           =   1455
   End
   Begin VB.TextBox txtMTDolaresRechazados 
      BackColor       =   &H80000000&
      Height          =   300
      Left            =   9120
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   6000
      Width           =   1455
   End
   Begin VB.ComboBox cboArchivo 
      Height          =   315
      Left            =   1080
      Style           =   2  'Dropdown List
      TabIndex        =   11
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton cmdCargar 
      Caption         =   "&Cargar"
      Height          =   375
      Left            =   3000
      TabIndex        =   5
      Top             =   720
      Width           =   975
   End
   Begin VB.TextBox txtUsuarioConfirma 
      Height          =   300
      Left            =   9480
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtFecha 
      Height          =   300
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   120
      Width           =   1575
   End
   Begin TabDlg.SSTab SSTAB 
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   1320
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   7223
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Abonos Soles"
      TabPicture(0)   =   "frmBancoPagadorConfirmacionAbonos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "flxAbonosRealizados"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Abonos Dolares"
      TabPicture(1)   =   "frmBancoPagadorConfirmacionAbonos.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "flxAbonosRechazados"
      Tab(1).ControlCount=   1
      Begin Sicmact.FlexEdit flxAbonosRealizados 
         Height          =   3615
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   6376
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Cuenta-Moneda-Imp. Neto-Saldo Disp.-Saldo Cont.-Estado Cta.-Resultado"
         EncabezadosAnchos=   "300-1900-1000-1400-1400-1400-1400-2500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Mshbco1 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   7
         Top             =   600
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6588
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin Sicmact.FlexEdit flxAbonosRechazados 
         Height          =   3615
         Left            =   -74880
         TabIndex        =   29
         Top             =   360
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   6376
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Cuenta-Moneda-Imp. Neto-Saldo Disp.-Saldo Cont.-Estado Cta.-Resultado"
         EncabezadosAnchos=   "300-1900-1000-1400-1400-1400-1400-2500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "F. Proceso:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -69480
         TabIndex        =   10
         Top             =   4680
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Importe :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -66720
         TabIndex        =   9
         Top             =   4680
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "No procesados :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72120
         TabIndex        =   8
         Top             =   4680
         Visible         =   0   'False
         Width           =   1410
      End
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "Abonos Realizados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   23
      Top             =   5580
      Width           =   1695
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "M. T. Soles Realizados :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2880
      TabIndex        =   22
      Top             =   5580
      Width           =   2100
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "M. T. Dolares Realizados :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6705
      TabIndex        =   21
      Top             =   5580
      Width           =   2280
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "Abonos Rechazados:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   120
      TabIndex        =   20
      Top             =   6060
      Width           =   1815
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "M. T. Soles Rechazados :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   2880
      TabIndex        =   19
      Top             =   6060
      Width           =   2220
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "M. T. Dolares Rechazados :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   6705
      TabIndex        =   18
      Top             =   6060
      Width           =   2400
   End
   Begin VB.Label Label3 
      Caption         =   "Archivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   720
      Width           =   735
   End
   Begin VB.Label Label2 
      Caption         =   "Usuario:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   8760
      TabIndex        =   2
      Top             =   120
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha y Hora:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmBancoPagadorConfirmacionAbonos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnIdProc As Long	' ANGC20211020 INT A LONG

Private Sub cmdCancelar_Click()
    txtAbonosRealizados.Text = ""
    txtAbonosRechazados.Text = ""
    txtMTSolesRealizados.Text = ""
    txtMTSolesRechazados.Text = ""
    txtMTDolaresRealizados.Text = ""
    txtMTDolaresRechazados.Text = ""
    cmdConfirmarAbonos.Visible = False
    Call LimpiaFlex(flxAbonosRealizados)
    Call LimpiaFlex(flxAbonosRechazados)
    Call InicializarDatos
End Sub

Private Sub cmdCargar_Click()
If cboArchivo.Text = "" Then
    MsgBox "Necesita seleccionar un archivo", vbInformation, "Aviso"
    Exit Sub
End If
    Call CargaDatosFinProcesoInicial(CLng(Right(cboArchivo.Text, 7))) 'GIPO 20180620
    cmdConfirmarAbonos.Visible = True
End Sub

Public Sub InicializarDatos()
Dim oBancoPag As New DBancoPagador
Dim rsArchiPro As ADODB.Recordset
Set rsArchiPro = oBancoPag.DevuelveArchivosProcesados()

Call CargaCombo(cboArchivo, rsArchiPro)
txtUsuarioConfirma.Text = gsCodUser
txtFecha.Text = Format(gdFecSis, "DD-MM-YYYY") & " " & Format(Time(), "HH:MM:SS")
End Sub

Public Sub CargaDatosFinProcesoInicial(ByVal pnIdProc As Long)	' ANGC20211020 INT A LONG
    Dim oBancoPag As New DBancoPagador
    Dim rsValFinProcIni As New ADODB.Recordset
    Set rsValFinProcIni = oBancoPag.DevuelveValoresFinProcesoInicial(pnIdProc)
    lnIdProc = pnIdProc
    
    
    If Not rsValFinProcIni.EOF Then
        txtAbonosRealizados.Text = rsValFinProcIni!nCantRealizados
        txtAbonosRechazados.Text = rsValFinProcIni!nCantRechazados
        txtMTSolesRealizados.Text = Format(rsValFinProcIni!nMTSolesRealizados, "#0.00")
        txtMTSolesRechazados.Text = Format(rsValFinProcIni!nMTSolesRechazados, "#0.00")
        txtMTDolaresRealizados.Text = Format(rsValFinProcIni!nMTDolaresRealizados, "#0.00")
        txtMTDolaresRechazados.Text = Format(rsValFinProcIni!nMTDolaresRechazados, "#0.00")
    End If
    
    Call CargaDatosSoles(pnIdProc, 2, 1) '' 2 Procesado
    Call CargaDatosDolares(pnIdProc, 2, 2)
End Sub


Public Sub CargaDatosSoles(ByVal pnIdProc As Long, ByVal pnEstado As Integer, ByVal pnMoneda As Integer)	' ANGC20211020 INT A LONG
    Dim nNumFila As Integer
    Dim oBancoPag As New DBancoPagador
    Dim rsRegReali As ADODB.Recordset
    Set rsRegReali = New ADODB.Recordset
    Set rsRegReali = oBancoPag.DevuelveRegistrosxMoneda(pnIdProc, pnEstado, pnMoneda) '''Realizados = Soles
    
    Call LimpiaFlex(flxAbonosRealizados)
    
    If RSVacio(rsRegReali) Then
       Exit Sub
    End If
    
    If rsRegReali.RecordCount > 0 Then
        Me.flxAbonosRealizados.Clear
        Me.flxAbonosRealizados.Rows = 2
        Me.flxAbonosRealizados.FormaCabecera
        
        Do While Not rsRegReali.EOF
            
            If flxAbonosRealizados.TextMatrix(1, 1) <> "" Then
                flxAbonosRealizados.AdicionaFila
            End If
            
            nNumFila = flxAbonosRealizados.Rows - 1
            
            flxAbonosRealizados.TextMatrix(nNumFila, 0) = nNumFila
            flxAbonosRealizados.TextMatrix(nNumFila, 1) = rsRegReali!cCodCnta_8
            flxAbonosRealizados.TextMatrix(nNumFila, 2) = rsRegReali!cMonedaCnta_9
            flxAbonosRealizados.TextMatrix(nNumFila, 3) = Format(rsRegReali!nImporteNeto_12, "##,##0.00")
            flxAbonosRealizados.TextMatrix(nNumFila, 4) = Format(rsRegReali!nSaldoDisp, "##,##0.00")
            flxAbonosRealizados.TextMatrix(nNumFila, 5) = Format(rsRegReali!nSaldoCont, "##,##0.00")
            flxAbonosRealizados.TextMatrix(nNumFila, 6) = rsRegReali!cEstadoCta
            flxAbonosRealizados.TextMatrix(nNumFila, 7) = rsRegReali!cDescripcion
            rsRegReali.MoveNext
        Loop
    Else
        MsgBox "No se encontraron datos", vbInformation, "Mensaje del Sistema"
    End If
End Sub

Public Sub CargaDatosDolares(ByVal pnIdProc As Long, ByVal pnEstado As Integer, ByVal pnMoneda As Integer)		' ANGC20211020 INT A LONG
    Dim nNumFila As Integer
    Dim oBancoPag As New DBancoPagador
    Dim rsRegRecha As ADODB.Recordset
    Set rsRegRecha = New ADODB.Recordset
    Set rsRegRecha = oBancoPag.DevuelveRegistrosxMoneda(pnIdProc, pnEstado, pnMoneda) '''Rechazados = Dolares
    
    Call LimpiaFlex(flxAbonosRechazados)
    
    If RSVacio(rsRegRecha) Then
       Exit Sub
    End If
    
        If rsRegRecha.RecordCount > 0 Then
            Me.flxAbonosRechazados.Clear
            Me.flxAbonosRechazados.Rows = 2
            Me.flxAbonosRechazados.FormaCabecera
            
            Do While Not rsRegRecha.EOF
                
                If flxAbonosRechazados.TextMatrix(1, 1) <> "" Then
                    flxAbonosRechazados.AdicionaFila
                End If
                
                nNumFila = flxAbonosRechazados.Rows - 1
                
                flxAbonosRechazados.TextMatrix(nNumFila, 0) = nNumFila
                flxAbonosRechazados.TextMatrix(nNumFila, 1) = rsRegRecha!cCodCnta_8
                flxAbonosRechazados.TextMatrix(nNumFila, 2) = rsRegRecha!cMonedaCnta_9
                flxAbonosRechazados.TextMatrix(nNumFila, 3) = Format(rsRegRecha!nImporteNeto_12, "##,##0.00")
                flxAbonosRechazados.TextMatrix(nNumFila, 4) = Format(rsRegRecha!nSaldoDisp, "##,##0.00")
                flxAbonosRechazados.TextMatrix(nNumFila, 5) = Format(rsRegRecha!nSaldoCont, "##,##0.00")
                flxAbonosRechazados.TextMatrix(nNumFila, 6) = rsRegRecha!cEstadoCta
                flxAbonosRechazados.TextMatrix(nNumFila, 7) = rsRegRecha!cDescripcion
                rsRegRecha.MoveNext
            Loop
        Else
            MsgBox "No se encontraron datos", vbInformation, "Mensaje del Sistema"
        End If
End Sub

Private Sub cmdConfirmarAbonos_Click()
Dim oBancoPag As New DBancoPagador
Dim bTransac As Boolean

    If MsgBox("¿Esta seguro de realizar la confirmación de abonos?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If

On Error GoTo ErrAbonoBancPag
     Set oBancoPag = New DBancoPagador
     bTransac = False
     Call oBancoPag.dBeginTrans
     bTransac = True

     If (oBancoPag.InsertaAbonoBancoPagadorConfirmacion(lnIdProc, gsCodUser, gAhoBancPagIni, gAhoITFCargoCta, gsCodAge, Format(gdFecSis, "yyyy-mm-dd hh:mm:ss"))) Then
         oBancoPag.dCommitTrans
         bTransac = False
         MsgBox "Se ha realizado la confirmación de abono a las cuentas con éxito", vbInformation, "Aviso"
         cmdConfirmarAbonos.Enabled = False
     Else
         MsgBox "Se ha generado un error en la confirmación de abono. Comunicar a TI", vbInformation, "Aviso"
         Call oBancoPag.dRollbackTrans
     End If

    Set oBancoPag = Nothing
        Call CargaDatosSoles(lnIdProc, 4, 1) '' 2 Procesado
        Call CargaDatosDolares(lnIdProc, 4, 2)
    Exit Sub
ErrAbonoBancPag:
    If bTransac Then
        oBancoPag.dRollbackTrans
        Set oBancoPag = Nothing
    End If
    Err.Raise Err.Number, "Error En Confirmacion de Abonos", Err.Description
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
 Call InicializarDatos
End Sub
