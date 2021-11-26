VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCapServicioPagoBajaArchivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Servicio - Baja de Trama de Pagos"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11805
   Icon            =   "frmCapServicioPagoBajaArchivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   11805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5295
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   11655
      _ExtentX        =   20558
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Selección de Convenio"
      TabPicture(0)   =   "frmCapServicioPagoBajaArchivo.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "rtfCartas"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FRConvenio"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FRTrama"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame FRTrama 
         Caption         =   "Tramas recibidas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   3255
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   11415
         Begin VB.CommandButton cmdSalir 
            Caption         =   "&Salir"
            Height          =   375
            Left            =   120
            TabIndex        =   13
            Top             =   2760
            Width           =   1095
         End
         Begin VB.CommandButton cmdBaja 
            Caption         =   "&Dar Baja"
            Height          =   375
            Left            =   10200
            TabIndex        =   12
            Top             =   2760
            Width           =   1095
         End
         Begin SICMACT.FlexEdit FETramas 
            Height          =   2415
            Left            =   120
            TabIndex        =   3
            Top             =   240
            Width           =   11175
            _ExtentX        =   19711
            _ExtentY        =   4260
            Cols0           =   10
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Fecha-Referencia-N° Usuarios-Saldo S/.-Baja-Detalle-cCtaCod-cReferencia-cArchivo"
            EncabezadosAnchos=   "500-1200-4200-1200-1200-1200-1200-0-0-0"
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
            ColumnasAEditar =   "X-X-X-X-X-5-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-4-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-R-R-C-C-C-C-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   6
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame FRConvenio 
         Caption         =   "Busqueda de convenio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1455
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   11415
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "..."
            Height          =   375
            Left            =   3000
            TabIndex        =   11
            Top             =   210
            Width           =   375
         End
         Begin VB.TextBox txtCodigoConvenio 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1030
            Locked          =   -1  'True
            TabIndex        =   9
            Top             =   240
            Width           =   1935
         End
         Begin VB.TextBox txtCodigoEmpresa 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1030
            Locked          =   -1  'True
            TabIndex        =   8
            Top             =   960
            Width           =   1935
         End
         Begin VB.TextBox txtConvenio 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1035
            Locked          =   -1  'True
            TabIndex        =   6
            Top             =   600
            Width           =   6015
         End
         Begin VB.TextBox txtEmpresa 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   3010
            Locked          =   -1  'True
            TabIndex        =   4
            Top             =   960
            Width           =   4040
         End
         Begin VB.Label Label1 
            Caption         =   "Código:"
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
            Left            =   120
            TabIndex        =   10
            Top             =   255
            Width           =   735
         End
         Begin VB.Label lblConvenio 
            Caption         =   "Convenio:"
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
            Left            =   120
            TabIndex        =   7
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lblEmpresa 
            Caption         =   "Empresa:"
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
            Left            =   120
            TabIndex        =   5
            Top             =   960
            Width           =   855
         End
      End
      Begin RichTextLib.RichTextBox rtfCartas 
         Height          =   330
         Left            =   10920
         TabIndex        =   14
         Top             =   0
         Visible         =   0   'False
         Width           =   720
         _ExtentX        =   1270
         _ExtentY        =   582
         _Version        =   393217
         ScrollBars      =   3
         TextRTF         =   $"frmCapServicioPagoBajaArchivo.frx":0326
      End
   End
End
Attribute VB_Name = "frmCapServicioPagoBajaArchivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapServicioPagoBajaArchivo
'*** Descripción : Formulario para dar de baja una trama para un convenio.
'*** Creación : ELRO el 20130705 09:52:24 AM, según RFC1306270002
'********************************************************************
Option Explicit

Dim fnIdSerPag As Long
Dim fsNomSerPag As String
Dim fsPersCod As String
Dim fsPersNombre As String
Dim fsCodSerPag As String

Private Sub generarTramaBajaServicioPago(ByVal pnFila As Long)
Dim lsNombreArchivo As String
Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
Dim rsBeneficiarios As ADODB.Recordset
Set rsBeneficiarios = New ADODB.Recordset
Dim lnFilaFin As Integer
Dim lsArchivo1, lsArchivo2 As String

lsNombreArchivo = FETramas.TextMatrix(pnFila, 9)

Set rsBeneficiarios = oNCOMCaptaGenerales.obtenerBeneficiariosNoCobroTramaBajaConvenioServicioPago(CLng(FETramas.TextMatrix(pnFila, 8)))

If InStr(Trim(UCase(lsNombreArchivo)), ".XLS") <> 0 Or InStr(Trim(UCase(lsNombreArchivo)), ".XLSX") <> 0 Then
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsExtension As String
    Dim lnPosicion As Integer
    
    Set xlAplicacion = New Excel.Application
    
    lnPosicion = InStr(Trim(UCase(lsNombreArchivo)), ".")
    
    lsExtension = Right(lsNombreArchivo, Len(lsNombreArchivo) - lnPosicion)
    
    If UCase(lsExtension) = "XLS" Then
        lsArchivo1 = Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS")
        lsArchivo2 = "\SPOOLER\TramaConvenio_" & lsArchivo1 & ".xls"
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\FormatoCarta\TramaConvenio.xls")
        
    ElseIf UCase(lsExtension) = "XLSX" Then
        lsArchivo1 = Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS")
        lsArchivo2 = "\SPOOLER\TramaConvenio_" & lsArchivo1 & ".xlsx"

        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\FormatoCarta\TramaConvenio.xlsx")
    End If

    If Not (rsBeneficiarios.BOF And rsBeneficiarios.EOF) Then
        lnFilaFin = 10
        For Each xlHoja1 In xlLibro.Worksheets
            xlHoja1.Cells(3, 2) = UCase(txtEmpresa)
            xlHoja1.Cells(5, 2) = UCase(txtCodigoConvenio)
            xlHoja1.Cells(7, 2) = UCase(txtConvenio)
            Do While Not rsBeneficiarios.EOF
                xlHoja1.Cells(lnFilaFin, 1) = UCase(rsBeneficiarios!cPersIDnro)
                xlHoja1.Cells(lnFilaFin, 2) = UCase(rsBeneficiarios!cPersNombre)
                xlHoja1.Cells(lnFilaFin, 3) = UCase(rsBeneficiarios!nMonto)
                rsBeneficiarios.MoveNext
                lnFilaFin = lnFilaFin + 1
            Loop
        Next
    End If

    xlLibro.SaveAs (App.path & lsArchivo2)
    xlLibro.Close
    xlAplicacion.Quit
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing
ElseIf InStr(Trim(UCase(lsNombreArchivo)), ".TXT") Then
    Dim psArchivoAGrabar As String
    Dim ArcSal As Integer
    
    lsArchivo1 = Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS")
    lsArchivo2 = "\SPOOLER\TramaConvenio_" & lsArchivo1 & ".txt"
    psArchivoAGrabar = App.path & lsArchivo2
    
    ArcSal = FreeFile
    Open psArchivoAGrabar For Output As ArcSal
        If Not (rsBeneficiarios.BOF And rsBeneficiarios.EOF) Then
            Print #ArcSal, UCase(Trim(fsCodSerPag)) & "|" & UCase(txtEmpresa) & "|0"
            Do While Not rsBeneficiarios.EOF
                Print #ArcSal, UCase(Trim(rsBeneficiarios!cPersIDnro)) & "|" & UCase(Trim(rsBeneficiarios!cPersNombre)) & "|" & UCase(Trim(rsBeneficiarios!nMonto))
                rsBeneficiarios.MoveNext
            Loop
    End If
    Close ArcSal
    
End If
Set rsBeneficiarios = Nothing
Set oNCOMCaptaGenerales = Nothing
End Sub

Private Sub cargarTramaConvenioServicioPago()
Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
Dim rsTramas As ADODB.Recordset
Set rsTramas = New ADODB.Recordset

 LimpiaFlex FETramas
Set rsTramas = oNCOMCaptaGenerales.obtenerTramaConvenioServicioPago(fnIdSerPag)
Do While Not rsTramas.EOF
    FETramas.SetFocus
    FETramas.lbEditarFlex = True
    FETramas.AdicionaFila
    FETramas.TextMatrix(FETramas.row, 1) = rsTramas!cFecha
    FETramas.TextMatrix(FETramas.row, 2) = rsTramas!cRefArc
    FETramas.TextMatrix(FETramas.row, 3) = rsTramas!nNroBeneficiario
    FETramas.TextMatrix(FETramas.row, 4) = Format$(rsTramas!nSaldo, "##,##0.00")
    FETramas.TextMatrix(FETramas.row, 5) = rsTramas!bBaja
    FETramas.TextMatrix(FETramas.row, 6) = rsTramas!cDetalle
    FETramas.TextMatrix(FETramas.row, 7) = rsTramas!cCtaCod
    FETramas.TextMatrix(FETramas.row, 8) = rsTramas!Id_SerPagArc
    FETramas.TextMatrix(FETramas.row, 9) = rsTramas!cNomArc
    rsTramas.MoveNext
Loop
Set oNCOMCaptaGenerales = Nothing
Set rsTramas = Nothing
End Sub

Private Sub cmdBaja_Click()

If Trim(FETramas.TextMatrix(1, 0)) = "" Then Exit Sub

If MsgBox("¿Esta seguro que desea dar de Baja la trama?", vbYesNo, "Aviso") = vbYes Then
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim oNCOMContFunciones As NCOMContFunciones
    Set oNCOMContFunciones = New NCOMContFunciones
    Dim fs As Scripting.FileSystemObject
    Set fs = New Scripting.FileSystemObject
    Dim lsMovNro, lsNombreArchivo As String
    Dim lnConfirmar As Long
    Dim i As Long

    lsNombreArchivo = FETramas.TextMatrix(FETramas.row, 9)
    
    If InStr(Trim(UCase(lsNombreArchivo)), ".XLS") <> 0 Or InStr(Trim(UCase(lsNombreArchivo)), ".XLSX") <> 0 Then
        If InStr(Trim(UCase(lsNombreArchivo)), ".XLS") <> 0 Then
            If Not fs.FileExists(App.path & "\FormatoCarta\TramaConvenio.xls") Then
                MsgBox "No existe la plantilla TramaConvenio.xls en la carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
                Exit Sub
            End If
        
        ElseIf InStr(Trim(UCase(lsNombreArchivo)), ".XLSX") <> 0 Then
            If Not fs.FileExists(App.path & "\FormatoCarta\TramaConvenio.xlsx") Then
                MsgBox "No existe la plantilla TramaConvenio.xlsx en la carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
                Exit Sub
            End If
        End If
    End If
    
    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    For i = 1 To FETramas.Rows - 1
        If Trim(FETramas.TextMatrix(i, 5)) = "." Then
            lnConfirmar = oNCOMCaptaGenerales.darBajaTramaConvenioServicioPago(CLng(FETramas.TextMatrix(i, 8)), _
                                                                               lsMovNro, FETramas.TextMatrix(i, 7), _
                                                                               CCur(FETramas.TextMatrix(i, 4)))
            If lnConfirmar > 0 Then
                generarTramaBajaServicioPago (i)
            Else
                MsgBox "No se dio de baja la trama.", vbCritical, "Aviso"
                Exit Sub
            End If
        End If
    Next i
    cargarTramaConvenioServicioPago
    Set oNCOMCaptaGenerales = Nothing
    Set oNCOMContFunciones = Nothing
    
End If
End Sub

Private Sub CmdBuscar_Click()
fnIdSerPag = 0
fsNomSerPag = ""
fsPersCod = ""
fsPersNombre = ""
fsCodSerPag = ""
frmCapServicioPagoBusqueda.IniciarBusqueda fnIdSerPag, fsNomSerPag, fsPersCod, fsPersNombre, fsCodSerPag
txtCodigoConvenio = fsCodSerPag
txtConvenio = fsNomSerPag
txtCodigoEmpresa = fsPersCod
txtEmpresa = fsPersNombre
cargarTramaConvenioServicioPago
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub mostrarTramaBajaServicioPago()
Dim lnFila  As Integer

lnFila = FETramas.row

If Trim(FETramas.TextMatrix(lnFila, 6)) = "Ver" Then
    Dim oclsPrevioCredito As clsprevio 'PrevioCredito.clsPrevioCredito
    Dim oNCOMCaptaGenerales As COMNCaptaGenerales.NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim rsBeneficiarios As New ADODB.Recordset
    Dim fs As Scripting.FileSystemObject
    Set fs = New Scripting.FileSystemObject
    Dim lsConvenio As String
    Dim lnlinea  As Integer
    
    If Not fs.FileExists(App.path & "\FormatoCarta\TramaConvenio.txt") Then
        MsgBox "No existe la plantilla TramaConvenio.xls en la carpeta FormatoCarta, Consulte con el Area de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If
      
    rtfCartas.FileName = App.path & "\FormatoCarta\TramaConvenio.txt"
    Set rsBeneficiarios = oNCOMCaptaGenerales.obtenerBeneficiariosNoCobroTramaBajaConvenioServicioPago(CLng(FETramas.TextMatrix(lnFila, 8)))
    
    lsConvenio = rtfCartas.Text
    
    lsConvenio = Replace(lsConvenio, "<<empresa>>", txtEmpresa)
    lsConvenio = Replace(lsConvenio, "<<codigo>>", txtCodigoConvenio)
    lsConvenio = Replace(lsConvenio, "<<nombre>>", txtConvenio)

    If Not (rsBeneficiarios.BOF And rsBeneficiarios.EOF) Then
        
        Do While Not rsBeneficiarios.EOF
            lsConvenio = lsConvenio & ImpreFormat(UCase(Trim(rsBeneficiarios!cPersIDnro)), 8) & ImpreFormat(UCase(Trim(rsBeneficiarios!cPersNombre)), 66) & Space(4) & ImpreFormat(UCase(Format$(rsBeneficiarios!nMonto, "##,##0.00")), 13, 2, True) & Chr(10)
            rsBeneficiarios.MoveNext
        Loop
    End If

    Set oNCOMCaptaGenerales = Nothing
    Set oclsPrevioCredito = New clsprevio 'PrevioCredito.clsPrevioCredito

    oclsPrevioCredito.Show lsConvenio, "Benefeciarios que no cobraron", True, , gImpresora
    Set oclsPrevioCredito = Nothing
End If

End Sub

Private Sub FETramas_Click()
    If FETramas.Col = 6 Then
        mostrarTramaBajaServicioPago
    End If
End Sub

Private Sub FETramas_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim lsColumnas() As String
lsColumnas = Split(FETramas.ColumnasAEditar, "-")

If lsColumnas(FETramas.Col) = "X" Then
    Cancel = False
    SendKeys "{Tab}", True
    Exit Sub
End If
End Sub

Private Sub Form_Load()
fnIdSerPag = 0
fsNomSerPag = ""
fsPersCod = ""
fsPersNombre = ""
fsCodSerPag = ""
txtCodigoConvenio = ""
txtConvenio = ""
txtCodigoEmpresa = ""
txtEmpresa = ""
End Sub
