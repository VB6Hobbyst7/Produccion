VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form frmComentarioCalificacion 
   Caption         =   "Reporte de Calificación de Créditos - Auditoria"
   ClientHeight    =   3630
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10050
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmComentarioCalificacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3630
   ScaleWidth      =   10050
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   855
      Left            =   120
      TabIndex        =   26
      Top             =   2640
      Width           =   9855
      Begin VB.CommandButton Command1 
         Caption         =   "&Cargar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8760
         TabIndex        =   27
         Top             =   240
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Búsqueda"
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9855
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   8520
         TabIndex        =   24
         Top             =   1440
         Width           =   1215
      End
      Begin VB.CheckBox chkTodos 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "&Todos"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6120
         TabIndex        =   18
         Top             =   720
         Width           =   885
      End
      Begin VB.TextBox txtAMonto 
         Height          =   315
         Left            =   3480
         TabIndex        =   14
         Top             =   1080
         Width           =   1455
      End
      Begin VB.TextBox txtDeMonto 
         Height          =   315
         Left            =   1680
         TabIndex        =   13
         Top             =   1080
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Height          =   735
         Left            =   240
         TabIndex        =   8
         Top             =   1680
         Width           =   4095
         Begin MSComCtl2.DTPicker DTPicker1 
            Height          =   300
            Left            =   480
            TabIndex        =   9
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   71958529
            CurrentDate     =   39743
         End
         Begin MSComCtl2.DTPicker DTPicker2 
            Height          =   300
            Left            =   2640
            TabIndex        =   10
            Top             =   240
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   529
            _Version        =   393216
            Format          =   71958529
            CurrentDate     =   39743
         End
         Begin VB.Label Label5 
            Caption         =   "Hasta:"
            Height          =   255
            Left            =   1920
            TabIndex        =   12
            Top             =   240
            Width           =   615
         End
         Begin VB.Label Label4 
            Caption         =   "De:"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   240
            Width           =   375
         End
      End
      Begin VB.CheckBox chkFDesembolso 
         Alignment       =   1  'Right Justify
         Caption         =   "F. Desembolso:"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
      End
      Begin VB.TextBox txtTCambio 
         Height          =   285
         Left            =   8640
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin MSMask.MaskEdBox mskPeriodo1Del 
         Height          =   315
         Left            =   1680
         TabIndex        =   1
         Top             =   360
         Width           =   1500
         _ExtentX        =   2646
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSDataListLib.DataCombo dcInstituciones 
         Height          =   315
         Left            =   1680
         TabIndex        =   5
         Top             =   720
         Width           =   3255
         _ExtentX        =   5741
         _ExtentY        =   556
         _Version        =   393216
         BackColor       =   -2147483628
         ForeColor       =   -2147483640
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin SICMACT.TxtBuscar TxtAgencia 
         Height          =   285
         Left            =   7080
         TabIndex        =   19
         Top             =   720
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin MSDataListLib.DataCombo dcAnalista 
         Height          =   315
         Left            =   6240
         TabIndex        =   22
         Top             =   1080
         Width           =   3495
         _ExtentX        =   6165
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
      End
      Begin MSDataListLib.DataCombo dcTipoCredito 
         Height          =   315
         Left            =   7200
         TabIndex        =   29
         Top             =   1800
         Visible         =   0   'False
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   556
         _Version        =   393216
         Text            =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label9 
         Caption         =   "Producto:"
         Height          =   255
         Left            =   5280
         TabIndex        =   28
         Top             =   1920
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   5280
         TabIndex        =   25
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Analista:"
         Height          =   255
         Left            =   5280
         TabIndex        =   23
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7920
         TabIndex        =   21
         Top             =   720
         Width           =   1845
      End
      Begin VB.Label Label6 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   5280
         TabIndex        =   20
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Monto:"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   1080
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "De:"
         Height          =   255
         Left            =   1200
         TabIndex        =   16
         Top             =   1080
         Width           =   300
      End
      Begin VB.Label Label12 
         Caption         =   "A:"
         Height          =   255
         Left            =   3240
         TabIndex        =   15
         Top             =   1080
         Width           =   240
      End
      Begin VB.Label Label2 
         Caption         =   "Instituciones:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   720
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "F. Cierre:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio:"
         Height          =   195
         Left            =   5280
         TabIndex        =   3
         Top             =   360
         Width           =   1155
      End
   End
End
Attribute VB_Name = "frmComentarioCalificacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oGen As COMDConstSistema.DCOMGeneral

Private Sub chkFDesembolso_Click()
    If chkFDesembolso.value = 1 Then
        Frame2.Visible = True
        Label4.Visible = True
        Label5.Visible = True
        DTPicker1.Visible = True
        DTPicker2.Visible = True
    Else
        Frame2.Visible = False
        Label4.Visible = False
        Label5.Visible = False
        DTPicker1.Visible = False
        DTPicker2.Visible = False
    End If
End Sub

Private Sub chkTodos_Click()
    If chkTodos.value = 1 Then
        TxtAgencia.Text = ""
        lblAgencia.Caption = ""
        chkTodos.value = 1
        CargarAnalistas TxtAgencia.Text
    Else
        chkTodos.value = 0
    End If
End Sub

Private Sub Command1_Click()
    Call MostrarReporteCalificacionCartera(mskPeriodo1Del, txtTCambio.Text, dcInstituciones.BoundText)
End Sub

Public Sub MostrarReporteCalificacionCartera(pdFechaProc As Date, pnTipCamb As Double, ByVal cCodInst As String)
Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
Dim R As ADODB.Recordset
Dim lMatCabecera As Variant
Dim lsNombreArchivo As String

    lsNombreArchivo = "ReporteCalifCartAuditoria"
    
    ReDim lMatCabecera(47, 2)

    lMatCabecera(0, 0) = "cCtaCod": lMatCabecera(0, 1) = ""
    lMatCabecera(1, 0) = "Agencia": lMatCabecera(1, 1) = ""
    lMatCabecera(2, 0) = "Destino": lMatCabecera(2, 1) = ""
    lMatCabecera(3, 0) = "Codigo Cliente": lMatCabecera(3, 1) = ""
    lMatCabecera(4, 0) = "Cliente": lMatCabecera(4, 1) = ""
    lMatCabecera(5, 0) = "Documento": lMatCabecera(5, 1) = ""
    lMatCabecera(6, 0) = "Monto Aprob.": lMatCabecera(6, 1) = "N"
    lMatCabecera(7, 0) = "Estado": lMatCabecera(7, 1) = ""
    lMatCabecera(8, 0) = "Cuotas": lMatCabecera(8, 1) = "N"
    lMatCabecera(9, 0) = "Dia Fijo": lMatCabecera(9, 1) = "N"
    lMatCabecera(10, 0) = "Analista": lMatCabecera(10, 1) = ""
    lMatCabecera(11, 0) = "Tasa": lMatCabecera(11, 1) = "N"
    lMatCabecera(12, 0) = "Linea": lMatCabecera(12, 1) = ""
    lMatCabecera(13, 0) = "Fec. Desemb": lMatCabecera(13, 1) = "D"
    lMatCabecera(14, 0) = "Saldo Cap.": lMatCabecera(14, 1) = "N"
    lMatCabecera(15, 0) = "Cuota Actual": lMatCabecera(15, 1) = "N"
    lMatCabecera(16, 0) = "Tipo Personeria": lMatCabecera(16, 1) = ""
    lMatCabecera(17, 0) = "CIIU": lMatCabecera(17, 1) = ""
    lMatCabecera(18, 0) = "Direccion": lMatCabecera(18, 1) = ""
    lMatCabecera(19, 0) = "Calif. Anterior": lMatCabecera(19, 1) = ""
    lMatCabecera(20, 0) = "Calif. Actual": lMatCabecera(20, 1) = ""
    lMatCabecera(21, 0) = "Dias Atraso": lMatCabecera(21, 1) = "N"
    lMatCabecera(22, 0) = "Linea Credito": lMatCabecera(22, 1) = ""
    lMatCabecera(23, 0) = "Plazo": lMatCabecera(23, 1) = ""
    lMatCabecera(24, 0) = "Tipo Prod.": lMatCabecera(24, 1) = ""
    lMatCabecera(25, 0) = "Moneda": lMatCabecera(25, 1) = ""
    lMatCabecera(26, 0) = "Fec. Vcto": lMatCabecera(26, 1) = "D"
    lMatCabecera(27, 0) = "Int. Deveng.": lMatCabecera(27, 1) = "N"
    lMatCabecera(28, 0) = "Int. Suspen.": lMatCabecera(28, 1) = "N"
    lMatCabecera(29, 0) = "Por. Prov.": lMatCabecera(29, 1) = "N"
    lMatCabecera(30, 0) = "Prov. Con RCC": lMatCabecera(30, 1) = "N"
    lMatCabecera(31, 0) = "Prov. Sin RCC": lMatCabecera(31, 1) = "N"
    lMatCabecera(32, 0) = "Prov.Ant.Sin RCC": lMatCabecera(32, 1) = "N"
    lMatCabecera(33, 0) = "Prov.Ant Con RCC": lMatCabecera(33, 1) = "N"
    lMatCabecera(34, 0) = "Saldo Deudor": lMatCabecera(34, 1) = "N"
    lMatCabecera(35, 0) = "GAR. PREF": lMatCabecera(35, 1) = "N"
    lMatCabecera(36, 0) = "GAR.NO PREF": lMatCabecera(36, 1) = "N"
    lMatCabecera(37, 0) = "GAR. AUTOL": lMatCabecera(37, 1) = "N"
    lMatCabecera(38, 0) = "TIPO GAR CALIF": lMatCabecera(38, 1) = "N"
    lMatCabecera(39, 0) = "ALINEADO": lMatCabecera(39, 1) = ""
    lMatCabecera(40, 0) = "Condicion": lMatCabecera(40, 1) = ""
    lMatCabecera(41, 0) = "Calif.Sin Alin.": lMatCabecera(41, 1) = ""
    lMatCabecera(42, 0) = "Calif.Sist.F.": lMatCabecera(42, 1) = ""
    lMatCabecera(43, 0) = "Prov.Sin Alin.": lMatCabecera(43, 1) = "N"
    lMatCabecera(44, 0) = "Prov.Sist.F.": lMatCabecera(44, 1) = "N"
    lMatCabecera(45, 0) = "Cliente Unico CMACM": lMatCabecera(45, 1) = ""
    lMatCabecera(46, 0) = "Cod_SBS": lMatCabecera(46, 1) = ""
    
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Set R = objCOMNAuditoria.ObtenerDatosCalificacionComentario(pdFechaProc, pnTipCamb, cCodInst, IIf(chkFDesembolso.value = 0, "", Format(DTPicker1.value, "yyyymmdd")), IIf(chkFDesembolso.value = 0, "", Format(DTPicker2.value, "yyyymmdd")), txtDeMonto.Text, txtAMonto.Text, IIf(chkTodos.value = 1, "", TxtAgencia.Text), IIf(dcAnalista.BoundText = "0", "0", (Mid(dcAnalista.Text, Len(dcAnalista.Text) - 3, Len(dcAnalista.Text)))), IIf(cboMoneda.ListIndex = 0 Or cboMoneda.ListIndex = -1, "", cboMoneda.ListIndex), IIf(dcTipoCredito.BoundText = "705", dcTipoCredito.BoundText, Mid(dcTipoCredito.BoundText, 1, 1)))
    
    Set objCOMNAuditoria = Nothing
    If R.RecordCount <> 0 Then
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Comentarios - Calificación de Cartera", "", lsNombreArchivo, lMatCabecera, R, 2, , , True)
    Else
    MsgBox ("No Existen Registros de Datos"), vbCritical, "Aviso"
    End If
End Sub

Public Sub GeneraReporteEnArchivoExcel(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal psTitulo As String, ByVal psSubTitulo As String, _
                                    ByVal psNomArchivo As String, ByVal pMatCabeceras As Variant, ByVal prRegistros As ADODB.Recordset, _
                                    Optional pnNumDecimales As Integer, Optional Visible As Boolean = False, Optional psNomHoja As String = "", _
                                    Optional pbSinFormatDeReg As Boolean = False, _
                                    Optional pbUsarCabecerasDeRS As Boolean = False)
    Dim rs As ADODB.Recordset
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim liLineas As Integer, i As Integer
    Dim fs As Scripting.FileSystemObject
    Dim lnNumColumns As Integer

    If Not (prRegistros.EOF And prRegistros.BOF) Then
        If pbUsarCabecerasDeRS = True Then
            lnNumColumns = prRegistros.Fields.Count
        Else
            lnNumColumns = UBound(pMatCabeceras)
            lnNumColumns = IIf(prRegistros.Fields.Count < lnNumColumns, prRegistros.Fields.Count, prRegistros.Fields.Count)
        End If

        If psNomHoja = "" Then psNomHoja = psNomArchivo
        psNomArchivo = psNomArchivo & "_" & psCodUser & ".xls"

        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.path & "\Spooler\" & psNomArchivo) Then
            fs.DeleteFile (App.path & "\Spooler\" & psNomArchivo)
        End If
        Set xlLibro = xlAplicacion.Workbooks.Add
        Set xlHoja1 = xlLibro.Worksheets.Add

        xlHoja1.Name = psNomHoja
        xlHoja1.Cells.Select
        'xlHoja1.Cells.NumberFormat = "@"

        'Cabeceras
        xlHoja1.Cells(1, 1) = psNomCmac
        xlHoja1.Cells(1, lnNumColumns) = Trim(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"))
        xlHoja1.Cells(2, 1) = psNomAge
        xlHoja1.Cells(2, lnNumColumns) = psCodUser
        xlHoja1.Cells(4, 1) = psTitulo
        xlHoja1.Cells(5, 1) = psSubTitulo
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, lnNumColumns)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, lnNumColumns)).HorizontalAlignment = xlCenter

        liLineas = 6
        If pbUsarCabecerasDeRS = True Then
            For i = 0 To prRegistros.Fields.Count - 1
                xlHoja1.Cells(liLineas, i + 1) = prRegistros.Fields(i).Name
            Next i
        Else
            For i = 0 To lnNumColumns - 1
                If (i + 1) > UBound(pMatCabeceras) Then
                    xlHoja1.Cells(liLineas, i + 1) = prRegistros.Fields(i).Name
                Else
                    xlHoja1.Cells(liLineas, i + 1) = pMatCabeceras(i, 0)
                End If
            Next i
        End If
        
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).HorizontalAlignment = xlCenter

        If pbSinFormatDeReg = False Then
            liLineas = liLineas + 1
            While Not prRegistros.EOF
                For i = 0 To lnNumColumns - 1
                    If pMatCabeceras(i, 1) = "" Then  'Verificamos si tiene tipo
                        xlHoja1.Cells(liLineas, i + 1) = prRegistros(i)
                    Else
                        Select Case pMatCabeceras(i, 1)
                            Case "S"
                                xlHoja1.Cells(liLineas, i + 1) = prRegistros(i)
                            Case "N"
                                xlHoja1.Cells(liLineas, i + 1) = Format(prRegistros(i), "#0.00")
                            Case "D"
                                xlHoja1.Cells(liLineas, i + 1) = IIf(Format(prRegistros(i), "yyyymmdd") = "19000101", "", Format(prRegistros(i), "dd/mm/yyyy"))
                        End Select
                    End If
                Next i
                liLineas = liLineas + 1
                prRegistros.MoveNext
            Wend
        Else
            xlHoja1.Range("A7").CopyFromRecordset prRegistros 'Copia el contenido del recordset a excel
        End If

        xlHoja1.SaveAs App.path & "\Spooler\" & psNomArchivo
        MsgBox "Se ha generado el Archivo en " & App.path & "\Spooler\" & psNomArchivo

        If Visible Then
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
        End If

            xlLibro.Close
            xlAplicacion.Quit
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing

    End If
End Sub

Private Sub Form_Load()
    Dim oCons As COMDConstantes.DCOMConstantes
    Set oCons = New COMDConstantes.DCOMConstantes
    Me.TxtAgencia.rs = oCons.getAgencias(, , True)
    chkTodos.value = 1
    CargarDatos
    CargarInstituciones
    CargarTipoCredito
    CargarAnalistas TxtAgencia.Text
    DTPicker1.Visible = False
    DTPicker1.value = Date
    DTPicker2.value = Date
    DTPicker2.Visible = False
    Frame2.Visible = False
    Label4.Visible = False
    Label5.Visible = False
    CargarMoneda
    cboMoneda.SelText = "Todos"
End Sub

Private Sub CargarInstituciones()
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Dim R As ADODB.Recordset
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Set R = objCOMNAuditoria.ObtenerInstituciones
    dcInstituciones.BoundColumn = "cPersCod"
    dcInstituciones.DataField = "cPersCod"
    Set dcInstituciones.RowSource = R
    dcInstituciones.ListField = "cPersNombre"
    dcInstituciones.BoundText = 0
End Sub

Private Sub CargarTipoCredito()
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Dim rsProducto As ADODB.Recordset
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Set rsProducto = objCOMNAuditoria.ObtenerTipoCredito
    Set dcTipoCredito.RowSource = rsProducto
    dcTipoCredito.BoundColumn = "nConsValor"
    dcTipoCredito.ListField = "cConsDescripcion"
    Set objCOMNAuditoria = Nothing
    Set rsProducto = Nothing
    dcTipoCredito.BoundText = 0
End Sub

Private Sub CargarMoneda()
    cboMoneda.AddItem "Todos", 0
    cboMoneda.AddItem "SOLES", gMonedaNacional
    cboMoneda.AddItem "DOLARES", gMonedaExtranjera
End Sub

Private Sub CargarAnalistas(ByVal sAgencia As String)
    Dim rsAnalista As New ADODB.Recordset
    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Set rsAnalista.DataSource = objCOMNAuditoria.DarAnalista(sAgencia)
    dcAnalista.BoundColumn = "cPersCod"
    dcAnalista.DataField = "cPersCod"
    Set dcAnalista.RowSource = rsAnalista
    dcAnalista.ListField = "cPersNombre"
    dcAnalista.BoundText = 0
End Sub

Public Sub CargarDatos()
    Dim oTipCambio As nTipoCambio
    Dim FechaFinMes As String
    Me.mskPeriodo1Del = gdFecData
    FechaFinMes = gdFecData
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Set oTipCambio = New nTipoCambio
        txtTCambio.Text = Format(oTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "#0.000")
    Set oTipCambio = Nothing
End Sub

Private Sub TxtAgencia_EmiteDatos()
    Set oGen = New COMDConstSistema.DCOMGeneral
    Me.lblAgencia.Caption = TxtAgencia.psDescripcion
    chkTodos.value = 0
    CargarAnalistas TxtAgencia.Text
End Sub

Private Sub txtDeMonto_LostFocus()
    FormatoMoneda
End Sub

Private Sub txtAMonto_LostFocus()
    FormatoMoneda
End Sub

Sub FormatoMoneda()
    If Len(txtDeMonto.Text) > 0 Then
    txtDeMonto.Text = Format(txtDeMonto.Text, "#,##0.00")
    End If
    If Len(txtAMonto.Text) > 0 Then
    txtAMonto.Text = Format(txtAMonto.Text, "#,##0.00")
    End If
End Sub
