VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRegCompraGastos 
   Caption         =   "Registro de Compras y Gastos"
   ClientHeight    =   7785
   ClientLeft      =   1125
   ClientTop       =   2535
   ClientWidth     =   4905
   Icon            =   "frmRegCompraGastos.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7785
   ScaleWidth      =   4905
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox chkDetallarCuentas 
      Caption         =   "Detallar Cuentas Contables"
      Height          =   255
      Left            =   1560
      TabIndex        =   20
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cliente"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   16
      Top             =   1800
      Width           =   4605
      Begin Sicmact.TxtBuscar TxtBCodPers 
         Height          =   375
         Left            =   1680
         TabIndex        =   18
         Top             =   240
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   661
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   2
      End
      Begin VB.CheckBox chkCliente 
         Caption         =   "Activar con"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblNombrePersonaJur 
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
         Left            =   120
         TabIndex        =   19
         Top             =   840
         Width           =   4335
      End
   End
   Begin VB.Frame frame3 
      Caption         =   "Tipo Documento"
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
      Height          =   855
      Left            =   120
      TabIndex        =   14
      Top             =   960
      Width           =   4605
      Begin VB.ComboBox cmbDocumento 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmRegCompraGastos.frx":030A
         Left            =   240
         List            =   "frmRegCompraGastos.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   4215
      End
   End
   Begin VB.CheckBox chktodas 
      Caption         =   "Todos"
      Height          =   255
      Left            =   165
      TabIndex        =   13
      Top             =   3480
      Value           =   1  'Checked
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdConsolidado 
      Cancel          =   -1  'True
      Caption         =   "Consolidado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   2160
      TabIndex        =   12
      Top             =   7290
      Width           =   1305
   End
   Begin VB.CommandButton cmdGenerarResumen 
      Caption         =   "Re&sumen"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   1470
      TabIndex        =   11
      Top             =   7290
      Visible         =   0   'False
      Width           =   1185
   End
   Begin MSComctlLib.ListView lv 
      Height          =   2700
      Left            =   105
      TabIndex        =   10
      Top             =   3780
      Visible         =   0   'False
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   4763
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Cuenta"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Descripción"
         Object.Width           =   6174
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fechas"
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
      Height          =   795
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   4605
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   300
         Left            =   810
         TabIndex        =   0
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin MSMask.MaskEdBox txtFecha2 
         Height          =   315
         Left            =   3150
         TabIndex        =   1
         Top             =   300
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   556
         _Version        =   393216
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   345
         Width           =   555
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2460
         TabIndex        =   8
         Top             =   330
         Width           =   510
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3510
      TabIndex        =   5
      Top             =   7275
      Width           =   1185
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   135
      TabIndex        =   4
      Top             =   7275
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Caption         =   "Presentación"
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
      Height          =   615
      Left            =   120
      TabIndex        =   6
      Top             =   6540
      Width           =   4605
      Begin VB.OptionButton optPrinter 
         Caption         =   "Formato &Texto"
         Enabled         =   0   'False
         Height          =   300
         Index           =   0
         Left            =   450
         TabIndex        =   2
         Top             =   210
         Width           =   1800
      End
      Begin VB.OptionButton optPrinter 
         Caption         =   "Formato &Excel"
         Height          =   195
         Index           =   1
         Left            =   2640
         TabIndex        =   3
         Top             =   263
         Value           =   -1  'True
         Width           =   1470
      End
   End
End
Attribute VB_Name = "frmRegCompraGastos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Variables de Excel
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lsArchivo As String
Dim oPersona As DPersona

Dim sSql As String, sTexto As String
Dim rs  As New ADODB.Recordset
Dim rsTDoc  As New ADODB.Recordset
Dim nImporte As Currency, nImporteME As Currency
Dim nTotal As Currency, nTotalH As Currency

Dim sCtaCod As String
Dim sCtaIGV As String, sCtaISC As String
Dim sCtaIGVND As String 'Add by GITU 11-12-2008
Dim sCtaRetrac As String
Dim sCtaImpCuarta As String

Dim lbLibroOpen As Boolean
Dim oBarra As New clsProgressBar
'-Para el Reg Comp Gitu
Dim i As Long
Dim SQLRegCom As String
Dim rsRegCom As ADODB.Recordset
'Dim lsAnio As String
'Dim lsMes As String
'Dim lsPeriodo As String
Dim dbRegCom  As ADODB.Connection
Dim oRegCom   As NContImpreReg
'Dim oBarra As clsProgressBar
Dim sConRegCom As String

Dim lbPorDocumento As Boolean
Dim nlogico As Integer

Public Sub PorDocumento(pbPorDocumento As Boolean)
    lbPorDocumento = pbPorDocumento
End Sub

Public Function ImprimeRegCompraGastosExcel(psCtaCod As String, pdFecha As Date, pdFecha2 As Date, pnTipo As Integer) As Boolean

Dim rsCta As New ADODB.Recordset
Dim rsMay As New ADODB.Recordset
Dim N As Integer
Dim sBillete As String, sPersona As String
Dim sDoc As String
Dim nImporte As Currency
Dim nVV As Currency
Dim nTotVV As Currency, nTotIgv As Currency, nTotISC As Currency, nTotPV As Currency
'GITU 29012009***********************************************************************
'Dim nItem As Integer, nLin As Integer, P As Integer
Dim nItem As Integer, nLin As Long, P As Integer
'************************************************************
Dim nIGV As Currency, nISC As Currency
Dim sDocs As String
Dim sCTAS As String
Dim cOpeTpoDes As String

Dim sImpre As String, sImpreAge As String
Dim lsRepTitulo    As String
Dim lsHoja         As String
'On Error GoTo ErrImprime

lsRepTitulo = "REGISTRO DE COMPRAS Y GASTOS"

oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = 1
'oBarra.Progress 0, lsRepTitulo, "Obteniendo Cuentas", ""
Dim oOpe As New DOperacion
'sCTAS = oOpe.CargaListaCuentasOperacion(gsOpeCod, psCtaCod)
'oBarra.Progress 1, lsRepTitulo, "Obteniendo Cuentas", ""

'Documentos que se consideran en Reporte
oBarra.Progress 0, lsRepTitulo, "Obteniendo Documentos", ""
sDocs = oOpe.CargaListaDocsOperacion(gsOpeCod)
Set oOpe = Nothing
oBarra.Progress 1, lsRepTitulo, "Obteniendo Documentos", ""

'Se tienen que capturar de Operación
Dim oConst As New NConstSistemas
sCtaIGV = oConst.LeeConstSistema(gConstCtaIGV)
sCtaISC = "21140209"
sCtaRetrac = "29180799"

Set oConst = Nothing

sCTAS = psCtaCod
oBarra.Progress 0, lsRepTitulo, "Obteniendo Movimientos", ""
Dim oReg As New NContImpreReg
Set rsCta = oReg.GetMovCompraGastos(sCTAS, sDocs, pdFecha, pdFecha2, sCtaIGV, sCtaISC)
If rsCta.EOF Then
   MsgBox "No se registraron movimientos...!", vbInformation, "Resultado"
   oBarra.CloseForm Me
   ImprimeRegCompraGastosExcel = False
   Exit Function
End If
oBarra.Progress 1, lsRepTitulo, "Obteniendo Movimientos", ""

If pnTipo = 0 Then 'tipo para impresion general (0 General , 1 Consolidada)
    oBarra.Progress 0, lsRepTitulo, "Mayorizando Movimientos", ""
    Set rsMay = oReg.GetMovCompraGastosMayor(sCTAS, sDocs, pdFecha, pdFecha2, sCtaIGV, sCtaISC, , , sCtaRetrac)
    nItem = 0: nLin = gnLinPage: P = 0
    oBarra.Progress 1, lsRepTitulo, "Mayorizando Movimientos", ""
End If
   nLin = 9
   lsHoja = "C_" & Mid(psCtaCod, 2, 2)
   ExcelAddHoja lsHoja, xlLibro, xlHoja1
   ImprimeRegCompraGastosExcelCab pdFecha, pdFecha2, gsNomCmac

oBarra.Max = rsCta.RecordCount
Do While Not rsCta.EOF
   oBarra.Progress rsCta.Bookmark, lsRepTitulo, "", "Generando Reporte... ", vbBlue
   sCtaCod = rsCta!cCtaContCod
   
   'Imprimimos mayorización
If pnTipo = 0 Then ' tipo para impresion general  (0 General , 1 Consolidada)
   Do While Not rsMay.EOF
      If Mid(sCtaCod, 1, Len(rsMay!cCtaContCod)) <> rsMay!cCtaContCod Then
         Exit Do
      End If
      '***Modificado por ELRO el 20111010, según Acta 278-2011/TI-D
      LineaVaciaExcel nLin
      '***Fin Modificado por ELRO**********************************
      ImpreMayorExcel rsMay!cCtaContCod, rsMay!cCtaContDesc, PrnVal(rsMay!nPV - rsMay!nIGV - rsMay!nISC, 14, 2), PrnVal(rsMay!nIGV, 14, 2), PrnVal(rsMay!nISC, 14, 2), PrnVal(rsMay!nPV, 14, 2), nLin, rsMay!Acumulado
      '***Modificado por ELRO el 20111010, según Acta 278-2011/TI-D
      If nLin > 14 Then
        LineaVaciaExcel nLin
      End If
      '***Fin Modificado por ELRO**********************************
      rsMay.MoveNext
   Loop
End If
   If rsCta.EOF Then
      Exit Do
   End If
   N = 0
   nTotVV = 0
   nTotPV = 0
   nTotIgv = 0
   nTotISC = 0
   LineaVaciaExcel nLin

   Do While rsCta!cCtaContCod = sCtaCod
      N = N + 1
      If Mid(rsCta!cOpeCod, 3, 1) = "2" And rsCta!nIGV <> 0 Then
         nIGV = Round(rsCta!nPV - (rsCta!nPV / (1 + gnIGVValor)), 2)
      Else
         nIGV = rsCta!nIGV
      End If
      nISC = rsCta!nISC
      nVV = rsCta!nPV - nIGV - nISC
      sDoc = IIf(IsNull(rsCta!nDocTpo), space(36), Format(rsCta!dDocFecha, "dd/mm/yyyy") & " " & Mid(rsCta!nDocTpo & space(3), 1, 3) & " " & Mid(rsCta!cDocNro & space(20), 1, 20) & " ")
      cOpeTpoDes = oOpe.GetOperacionDesc(Mid(rsCta!cOpeCod, 1, 4) & "00")
      ImpreDetalleExcel Mid(rsCta!cMovNro, 1, 8) & "-" & Mid(rsCta!cMovNro, 9, 6) & "-" & Right(Trim(rsCta!cMovNro), 4), _
                        IIf(IsNull(rsCta!nDocTpo), space(10), rsCta!dDocFecha), _
                        IIf(IsNull(rsCta!nDocTpo), space(3), Mid(rsCta!nDocTpo & space(3), 1, 3)), _
                        IIf(IsNull(rsCta!nDocTpo), space(20), Mid(rsCta!cDocNro & space(20), 1, 20)), _
                        IIf(IsNull(rsCta!cRuc), space(8), rsCta!cRuc), _
                        IIf(IsNull(rsCta!cPersNombre), "", rsCta!cPersNombre), _
                        Replace(Replace(rsCta!cMovDesc, Chr(13), " "), oImpresora.gPrnSaltoLinea, ""), _
                        PrnVal(nVV, 14, 2), PrnVal(nIGV, 14, 2), PrnVal(nISC, 14, 2), PrnVal(rsCta!nPV, 14, 2), rsCta!cDocNroEmi, rsCta!dDocFechaEmi, _
                        nLin, cOpeTpoDes
      nTotVV = nTotVV + nVV
      nTotPV = nTotPV + rsCta!nPV
      nTotIgv = nTotIgv + nIGV
      nTotISC = nTotISC + nISC
      rsCta.MoveNext
      If rsCta.EOF Then
         Exit Do
      End If
   Loop
   LineaVaciaExcel nLin
Loop
RSClose rsCta
RSClose rsMay
RSClose rs
oBarra.CloseForm Me
ImprimeRegCompraGastosExcel = True
Exit Function
ErrImprime:
 MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
'   If lbLibroOpen Then
'      xlLibro.Close
'      xlAplicacion.Quit
'   End If
'   Set xlAplicacion = Nothing
'   Set xlLibro = Nothing
'   Set xlHoja1 = Nothing
End Function
'GITU 29012009*******************************
'Private Sub ImpreMayorExcel(sCta As String, sDes As String, nVV As Currency, nIGV As Currency, nISC As Currency, nPV As Currency, ByRef nLin As Integer)
'***Modificado por ELRO el 20111011, según Acta 278-2011/TI-D
'Private Sub ImpreMayorExcel(sCta As String, sDes As String, nVV As Currency, nIGV As Currency, nISC As Currency, nPV As Currency, ByRef nLin As Long)
Private Sub ImpreMayorExcel(sCta As String, sDes As String, nVV As Currency, nIGV As Currency, nISC As Currency, nPV As Currency, ByRef nLin As Long, Optional ByVal nAcumulado As Currency = 0)
'***Fin Modificado por ELRO**********************************
'*******************************************
xlHoja1.Cells(nLin, 1) = ""
xlHoja1.Cells(nLin, 2) = "'" & sCta
xlHoja1.Cells(nLin, 3) = ""
xlHoja1.Cells(nLin, 4) = sDes
xlHoja1.Cells(nLin, 5) = ""
xlHoja1.Cells(nLin, 6) = ""
xlHoja1.Cells(nLin, 7) = ""
xlHoja1.Cells(nLin, 8) = nVV
xlHoja1.Cells(nLin, 9) = nIGV
xlHoja1.Cells(nLin, 10) = nISC
xlHoja1.Cells(nLin, 11) = nPV


'***Modificado por ELRO el 20111011, según Acta 278-2011/TI-D
'If nLin > 14 Then
xlHoja1.Cells(nLin - 1, 6) = "Saldo Acumulado"
xlHoja1.Cells(nLin, 6) = "Saldo en Consulta"
xlHoja1.Cells(nLin + 1, 6) = "Saldo Total"
xlHoja1.Cells(nLin - 1, 7) = nAcumulado
xlHoja1.Cells(nLin, 7) = nPV
xlHoja1.Cells(nLin + 1, 7).Interior.Color = RGB(255, 255, 0)
xlHoja1.Cells(nLin + 1, 7).Formula = "=SUM(G" & nLin - 1 & ":G" & nLin & ")"
nLin = nLin + 1
'End If
'***Fin Modificado por ELRO**********************************

nLin = nLin + 1
End Sub
'GITU 29012009*************************************
'Private Sub ImpreDetalleExcel(sFec As String, sDocFecha As String, sDocTpo As String, sDocNro, sRUC As String, sPer As String, sCon As String, nVV As Currency, nIGV As Currency, nISC As Currency, nPV As Currency, sDocNroEmi As String, dDocFechaEmi As Date, nLin As Integer, sOpeTpoDesc As String)
Private Sub ImpreDetalleExcel(sFec As String, sDocFecha As String, sDocTpo As String, sDocNro, sRUC As String, sPer As String, sCon As String, nVV As Currency, nIGV As Currency, nISC As Currency, nPV As Currency, sDocNroEmi As String, dDocFechaEmi As Date, nLin As Long, sOpeTpoDesc As String)
'*************************************************
xlHoja1.Cells(nLin, 1) = sFec
If Trim(sDocFecha) <> "" Then
   xlHoja1.Range(xlHoja1.Cells(nLin, 2), xlHoja1.Cells(nLin, 2)).NumberFormat = "dd/mm/yyyy"
   xlHoja1.Cells(nLin, 2) = CDate(sDocFecha)
End If
xlHoja1.Cells(nLin, 3) = sDocTpo
xlHoja1.Cells(nLin, 4) = sDocNro
xlHoja1.Cells(nLin, 5) = sRUC
xlHoja1.Cells(nLin, 6) = sPer
xlHoja1.Cells(nLin, 7) = sCon
xlHoja1.Cells(nLin, 8) = nVV
xlHoja1.Cells(nLin, 9) = nIGV
xlHoja1.Cells(nLin, 10) = nISC
xlHoja1.Cells(nLin, 11) = nPV
xlHoja1.Cells(nLin, 12) = sDocNroEmi
xlHoja1.Cells(nLin, 13) = Format(dDocFechaEmi, gsFormatoFechaView)
xlHoja1.Cells(nLin, 14) = sOpeTpoDesc
nLin = nLin + 1
End Sub
'GITU 29012009****************************************
'Private Sub ImpreDetalleExcelResumen(sDocFecha As String, sDocTpo As String, sDocNro, sFec As String, sRUC As String, sPer As String, sCon As String, nVV As Currency, nIGV As Currency, nISC As Currency, nPV As Currency, sDocNroEmi As String, dDocFechaEmi As Date, nLin As Integer, sOpeTpoDesc As String)
Private Sub ImpreDetalleExcelResumen(sDocFecha As String, sDocTpo As String, sDocNro, sFec As String, sRUC As String, sPer As String, sCon As String, nVV As Currency, nIGV As Currency, nISC As Currency, nPV As Currency, sDocNroEmi As String, dDocFechaEmi As Date, nLin As Long, sOpeTpoDesc As String)
'*****************************************************

If Trim(sDocFecha) <> "" Then
   xlHoja1.Range(xlHoja1.Cells(nLin, 1), xlHoja1.Cells(nLin, 1)).NumberFormat = "dd/mm/yyyy"
   xlHoja1.Cells(nLin, 1) = CDate(sDocFecha)
End If
xlHoja1.Cells(nLin, 2) = sDocTpo
xlHoja1.Cells(nLin, 3) = sDocNro
xlHoja1.Cells(nLin, 4) = sFec
xlHoja1.Cells(nLin, 5) = sRUC
xlHoja1.Cells(nLin, 6) = sPer
xlHoja1.Cells(nLin, 7) = sCon
xlHoja1.Cells(nLin, 8) = nVV
xlHoja1.Cells(nLin, 9) = nIGV
xlHoja1.Cells(nLin, 10) = nISC
xlHoja1.Cells(nLin, 11) = nPV
xlHoja1.Cells(nLin, 12) = sDocNroEmi
xlHoja1.Cells(nLin, 13) = Format(dDocFechaEmi, gsFormatoFechaView)
xlHoja1.Cells(nLin, 14) = sOpeTpoDesc
nLin = nLin + 1
End Sub
'GITU 29012009****************************************
'Private Sub LineaVaciaExcel(nLin As Integer)
Private Sub LineaVaciaExcel(nLin As Long)
'*******************************
Dim N As Integer
For N = 1 To 12
   xlHoja1.Cells(nLin, N) = ""
Next
nLin = nLin + 1
End Sub
'********Modificado por ALPA
'********29/02/2008
Private Sub chkCliente_Click()
If chkCliente.value = 1 Then
    TxtBCodPers.Enabled = True
Else
    TxtBCodPers.Enabled = False
End If
End Sub

'***Agregado por ELRO el 20111011, según Acta 278-2011/TI-D
Private Sub chkDetallarCuentas_Click()
    Dim K As Integer
    Dim bDetallarCuentas As Boolean

    bDetallarCuentas = False

    For K = 1 To lv.ListItems.Count
        If lv.ListItems(K).Checked = True And lv.ListItems(K).Text = "45" Then
             bDetallarCuentas = True
        End If
    Next

    If bDetallarCuentas = False Then
        chkDetallarCuentas.value = Unchecked
    End If

End Sub
'***Fin Agregado por ELRO**********************************

'**********************
Private Sub chktodas_Click()
Dim K As Integer
If chktodas.value = 1 Then
    For K = 1 To lv.ListItems.Count
        lv.ListItems(K).Checked = True
    Next
    '***Modificado por ELRO el 20111011, según Acta 278-2011/TI-D
    chkDetallarCuentas.Visible = False
    '***Fin Modificado por ELRO**********************************
Else
    For K = 1 To lv.ListItems.Count
        lv.ListItems(K).Checked = False
    Next
    '***Modificado por ELRO el 20111011, según Acta 278-2011/TI-D
    chkDetallarCuentas.Visible = True
    '***Fin Modificado por ELRO**********************************
End If
End Sub
'********Modificado por ALPA
'********29/02/2008
Private Sub Validar()
    If cmbDocumento.Text = "" Then
        cmbDocumento.Text = Mid("TODOS" & space(100), 1, 100) & "0000"
    End If
    If chkCliente.value = 1 Then
        If TxtBCodPers.Text = "" Then
            MsgBox "Deberia realizar busqueda del proveedor.", vbInformation, "Aviso"
            TxtBCodPers.SetFocus
            nlogico = 1
            Exit Sub
        End If
        If TxtBCodPers.Text = "00000000000" Then
            lblNombrePersonaJur.Caption = ""
            TxtBCodPers.SetFocus
            nlogico = 1
            MsgBox "Proveedor no se encuentra registrado.", vbInformation, "Aviso"
            Exit Sub
        End If
     End If
End Sub
'***************************
Private Sub cmdConsolidado_Click()
Dim K      As Integer
Dim sImpre As String
Dim lsCtaS As String
Dim oOpe As New DOperacion

sImpre = ""
If ValidaFecha(txtFecha) <> "" Or ValidaFecha(txtFecha2) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Aviso"
   Exit Sub
End If

If CDate(txtFecha) > CDate(txtFecha2) Then
   MsgBox "Fecha Inicial debe ser menor o igual que fecha final.", vbInformation, "Aviso"
   Exit Sub
End If

lsArchivo = App.path & "\SPOOLER\RCConsolidado_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & gsCodUser & ".XLS"

lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
If Not lbLibroOpen Then
   oBarra.CloseForm Me
   Exit Sub
End If

oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = lv.ListItems.Count
'Se quito la condicion de cuentas contables porque no estaba cuadrando
'con el registro de compras detallado By GITU 14-01-2009
If gsOpeCod = 760200 Then
    For K = 1 To lv.ListItems.Count
        lsCtaS = lsCtaS & "," & oOpe.CargaListaCuentasOperacion(gsOpeCod, lv.ListItems(K).Text)
        oBarra.Progress K, "REGISTRO DE COMPRAS Y GASTOS", "Obteniendo Cuentas", ""
    Next
Else
    oBarra.Progress K, "REGISTRO DE COMPRAS Y GASTOS", "Procesando", ""
    oBarra.CloseForm Me
End If

'If lsCtaS = "" Then
'   MsgBox "Debe seleccionar una cuenta", vbInformation, "Aviso"
'Else
        'If lsCtaS <> "" Or lbPorDocumento Then
        If lbPorDocumento Then
           'lsCtaS = Mid(lsCtaS, 2, Len(lsCtaS))
           
           If optPrinter(1) Then
              '***********Comentado por ALPA
              '***********29/02/2008
              'Comentado por GITU 14-01-2009
              'ImprimeConsolidado lsCtaS, CDate(txtFecha), CDate(txtFecha2), 0
              ImprimeConsolidado CDate(txtFecha), CDate(txtFecha2), 0
              'EnD Gitu
              '*****************************
           End If
           
        End If
        
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        CargaArchivo lsArchivo, App.path & "\SPOOLER\"
                
        If optPrinter(0) Then
           EnviaPrevio sImpre, "REGISTRO DE COMPRAS Y GASTOS - CONSOLIDADO", gnLinPage, False
        End If
'End If



End Sub

Private Sub cmdGenerar_Click()
Dim K      As Integer
Dim sImpre As String
Dim lsCtaS As String
Dim oOpe As New DOperacion
'***Modificado por ELRO el 20111011, según Acta 278-2011/TI-D
Dim oForm As frmCuentasContables
Dim lsCuentasContables As String
'***Fin Modificado por ELRO**********************************
sImpre = ""

If chkCliente.value = 1 Then
Else
TxtBCodPers.Text = ""
End If
If ValidaFecha(txtFecha) <> "" Or ValidaFecha(txtFecha2) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Aviso"
   Exit Sub
End If

If CDate(txtFecha) > CDate(txtFecha2) Then
   MsgBox "Fecha Inicial debe ser menor o igual que fecha final.", vbInformation, "Aviso"
   Exit Sub
End If
'*************Modificado ALPA
'*************28/02/2008
nlogico = 0
Call Validar
If nlogico = 0 Then
'***************************
'lsArchivo = App.path & "\SPOOLER\RCG_" & Format(txtFecha2 & " " & Time, "mmyyyy") & gsCodUser & gsOpeCod & ".XLS"
If gsOpeCod = 760200 Then
    lsArchivo = App.path & "\SPOOLER\RCG_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
Else
    lsArchivo = App.path & "\SPOOLER\RCG1_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
End If

If gsOpeCod = 760201 Then
    dbRegCom.Execute "DELETE FROM RegCom"
End If

lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
If Not lbLibroOpen Then
    oBarra.CloseForm Me
    Exit Sub
End If

'***Modificado por ELRO el 20111011, según Acta 278-2011/TI-D
If Me.chkDetallarCuentas.value = Checked Then
    Set oForm = New frmCuentasContables
    lsCuentasContables = oForm.Inicio(gsOpeCod)
End If
'***Fin Modificado por ELRO**********************************

oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = lv.ListItems.Count

If gsOpeCod = 760200 Then
    '***Modificado por ELRO el 20111011, según Acta 278-2011/TI-D
      '      For k = 1 To lv.ListItems.Count
      '        If lv.ListItems(k).Checked Then
      '            lsCtaS = lsCtaS & "," & oOpe.CargaListaCuentasOperacion(gsOpeCod, lv.ListItems(k).Text)
      '            oBarra.Progress k, "REGISTRO DE COMPRAS Y GASTOS", "Obteniendo Cuentas", ""
      '            'lv.ListItems(K).ForeColor = vbRed
      '        End If
      '    Next
      '    oBarra.CloseForm Me
  
    If Me.chkDetallarCuentas.value = Unchecked Then
    
        For K = 1 To lv.ListItems.Count
            If lv.ListItems(K).Checked Then
                lsCtaS = lsCtaS & "," & oOpe.CargaListaCuentasOperacion(gsOpeCod, lv.ListItems(K).Text)
                oBarra.Progress K, "REGISTRO DE COMPRAS Y GASTOS", "Obteniendo Cuentas", ""
                'lv.ListItems(K).ForeColor = vbRed
            End If
        Next
        
        oBarra.CloseForm Me
        
    Else
    
        If lsCuentasContables <> "" Then
            lsCtaS = lsCuentasContables
            oBarra.Progress 100, "REGISTRO DE COMPRAS Y GASTOS", "Obteniendo Cuentas", ""
            oBarra.CloseForm Me
            
        Else
            oBarra.CloseForm Me
        End If
    
    End If
Else
    oBarra.Progress K, "REGISTRO DE COMPRAS Y GASTOS", "Procesando", ""
    oBarra.CloseForm Me
End If


If lsCtaS = "" And gsOpeCod = 760200 Then ' Se agrego la condicion del cOpecod gitu
   MsgBox "Debe seleccionar una cuenta", vbInformation, "Aviso"
Else
    If lsCtaS <> "" Or lbPorDocumento Then
        lsCtaS = Mid(lsCtaS, 2, Len(lsCtaS))
       If optPrinter(1) Then
          'If Not lbPorDocumento Or gsOpeCod = 700200 Then
          If gsOpeCod = 760200 Then
            If ImprimeRegCompraGastosExcel(lsCtaS, CDate(txtFecha), CDate(txtFecha2), 0) = False Then
                Exit Sub
            End If
          Else
            '*************Modificado ALPA
            '*************28/02/2008
            ImprimeRegCompraGastosExcelLima lsCtaS, CDate(txtFecha), CDate(txtFecha2), Trim(Right(Trim(cmbDocumento.Text), 4)), TxtBCodPers.Text
            '****************************
          End If
       Else
    '      sImpre = sImpre & ImprimeRegCompraGastos(lv.ListItems(K).Text, lv.ListItems(K).SubItems(1), CDate(txtFecha), CDate(txtFecha2))
       End If
    End If
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
    CargaArchivo lsArchivo, App.path & "\SPOOLER\"
    
    If optPrinter(0) Then
       EnviaPrevio sImpre, "REGISTRO DE COMPRAS Y GASTOS", gnLinPage, False
    End If
    
End If
End If
End Sub

Private Sub cmdGenerarResumen_Click()
Dim K      As Integer
Dim sImpre As String
Dim lsCtaS As String
Dim oOpe As New DOperacion

sImpre = ""
If ValidaFecha(txtFecha) <> "" Or ValidaFecha(txtFecha2) <> "" Then
   MsgBox "Fecha no válida...!", vbInformation, "Aviso"
   Exit Sub
End If

If CDate(txtFecha) > CDate(txtFecha2) Then
   MsgBox "Fecha Inicial debe ser menor o igual que fecha final.", vbInformation, "Aviso"
   Exit Sub
End If

lsArchivo = App.path & "\SPOOLER\RCGR_" & Format(txtFecha2 & " " & Time, "mmyyyy") & gsCodUser & ".XLS"
lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
If Not lbLibroOpen Then
   oBarra.CloseForm Me
   Exit Sub
End If

oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = lv.ListItems.Count

For K = 1 To lv.ListItems.Count
   If lv.ListItems(K).Checked Then
      lsCtaS = lsCtaS & "," & oOpe.CargaListaCuentasOperacion(gsOpeCod, lv.ListItems(K).Text)
      oBarra.Progress K, "REGISTRO DE COMPRAS Y GASTOS", "Obteniendo Cuentas", ""

   End If
Next
oBarra.CloseForm Me

If lsCtaS = "" Then
   MsgBox "Debe seleccionar una cuenta", vbInformation, "Aviso"
Else
     
        If lsCtaS <> "" Or lbPorDocumento Then
            lsCtaS = Mid(lsCtaS, 2, Len(lsCtaS))
           If optPrinter(1) Then
              
              If ImprimeRegCompraGastosExcelResumen(lsCtaS, CDate(txtFecha), CDate(txtFecha2), 1) = False Then
                Exit Sub
              End If
              
           Else
        
           End If
        End If
        
        
        ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        CargaArchivo lsArchivo, App.path & "\SPOOLER\"
                
        If optPrinter(0) Then
           EnviaPrevio sImpre, "REGISTRO DE COMPRAS Y GASTOS", gnLinPage, False
        End If
End If
End Sub

Public Function ImprimeRegCompraGastosExcelResumen(psCtaCod As String, pdFecha As Date, pdFecha2 As Date, pnTipo As Integer) As Boolean
Dim rsCta As New ADODB.Recordset
Dim N As Integer
Dim sBillete As String, sPersona As String
Dim sDoc As String
Dim nImporte As Currency
Dim nVV As Currency
Dim nTotVV As Currency, nTotIgv As Currency, nTotISC As Currency, nTotPV As Currency
'GITU 29012009****************************************
'Dim nItem As Integer, nLin As Integer, P As Integer
Dim nItem As Integer, nLin As Long, P As Integer
Dim nIGV As Currency, nISC As Currency
Dim sDocs As String
Dim sCTAS As String
Dim cOpeTpoDes As String

Dim sImpre As String, sImpreAge As String
Dim lsRepTitulo    As String
Dim lsHoja         As String

lsRepTitulo = "REGISTRO DE COMPRAS Y GASTOS CONSOLIDADO"

oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = 1

Dim oOpe As New DOperacion

'Documentos que se consideran en Reporte
oBarra.Progress 0, lsRepTitulo, "Obteniendo Documentos", ""
sDocs = oOpe.CargaListaDocsOperacion(gsOpeCod)
Set oOpe = Nothing
oBarra.Progress 1, lsRepTitulo, "Obteniendo Documentos", ""

'Se tienen que capturar de Operación
Dim oConst As New NConstSistemas
sCtaIGV = oConst.LeeConstSistema(gConstCtaIGV)
sCtaISC = "21140209"
sCtaRetrac = "29180799"

Set oConst = Nothing

sCTAS = psCtaCod
oBarra.Progress 0, lsRepTitulo, "Obteniendo Movimientos", ""
Dim oReg As New NContImpreReg
Set rsCta = oReg.GetMovCompraGastos(sCTAS, sDocs, pdFecha, pdFecha2, sCtaIGV, sCtaISC, , , 1)
If rsCta.EOF Then
   MsgBox "No se registraron movimientos...!", vbInformation, "Resultado"
   oBarra.CloseForm Me
   ImprimeRegCompraGastosExcelResumen = False
   Exit Function
End If
oBarra.Progress 1, lsRepTitulo, "Obteniendo Movimientos", ""

   nLin = 9
   lsHoja = "C_" & Mid(psCtaCod, 2, 2)
   ExcelAddHoja lsHoja, xlLibro, xlHoja1
   
  'cambio de formato
   ImprimeRegCompraGastosExcelCabResumen pdFecha, pdFecha2, gsNomCmac

oBarra.Max = rsCta.RecordCount
Do While Not rsCta.EOF
   oBarra.Progress rsCta.Bookmark, lsRepTitulo, "", "Generando Reporte... ", vbBlue
   sCtaCod = rsCta!cCtaContCod
   

   If rsCta.EOF Then
      Exit Do
   End If
   N = 0
   nTotVV = 0
   nTotPV = 0
   nTotIgv = 0
   nTotISC = 0
   'LineaVaciaExcel nLin

   Do While rsCta!cCtaContCod = sCtaCod
      N = N + 1
      If Mid(rsCta!cOpeCod, 3, 1) = "2" And rsCta!nIGV <> 0 Then
         nIGV = Round(rsCta!nPV - (rsCta!nPV / (1 + gnIGVValor)), 2)
      Else
         nIGV = rsCta!nIGV
      End If
      nISC = rsCta!nISC
      nVV = rsCta!nPV - nIGV - nISC
      sDoc = IIf(IsNull(rsCta!nDocTpo), space(36), Format(rsCta!dDocFecha, "dd/mm/yyyy") & " " & Mid(rsCta!nDocTpo & space(3), 1, 3) & " " & Mid(rsCta!cDocNro & space(20), 1, 20) & " ")
      cOpeTpoDes = oOpe.GetOperacionDesc(Mid(rsCta!cOpeCod, 1, 4) & "00")
      ImpreDetalleExcelResumen IIf(IsNull(rsCta!nDocTpo), space(10), rsCta!dDocFecha), _
                        IIf(IsNull(rsCta!nDocTpo), space(3), Mid(rsCta!nDocTpo & space(3), 1, 3)), _
                        IIf(IsNull(rsCta!nDocTpo), space(20), Mid(rsCta!cDocNro & space(20), 1, 20)), _
                        Mid(rsCta!cMovNro, 1, 8) & "-" & Mid(rsCta!cMovNro, 9, 6) & "-" & Right(Trim(rsCta!cMovNro), 4), _
                        IIf(IsNull(rsCta!cRuc), space(8), rsCta!cRuc), _
                        IIf(IsNull(rsCta!cPersNombre), "", rsCta!cPersNombre), _
                        Replace(Replace(rsCta!cMovDesc, Chr(13), " "), oImpresora.gPrnSaltoLinea, ""), _
                        PrnVal(nVV, 14, 2), PrnVal(nIGV, 14, 2), PrnVal(nISC, 14, 2), PrnVal(rsCta!nPV, 14, 2), rsCta!cDocNroEmi, rsCta!dDocFechaEmi, _
                        nLin, cOpeTpoDes
      nTotVV = nTotVV + nVV
      nTotPV = nTotPV + rsCta!nPV
      nTotIgv = nTotIgv + nIGV
      nTotISC = nTotISC + nISC
      rsCta.MoveNext
      If rsCta.EOF Then
         Exit Do
      End If
   Loop
   'LineaVaciaExcel nLin
Loop
RSClose rsCta
'RSClose rsMay
RSClose rs
oBarra.CloseForm Me
ImprimeRegCompraGastosExcelResumen = True

Exit Function
ErrImprime:
 MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Function

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Initialize()
    lbPorDocumento = False
End Sub

Private Sub Form_Load()
Dim oOpe As New DOperacion
Dim lvItem As ListItem
Dim oConConsol As DConecta
Dim sql As String
Set oConConsol = New DConecta
oConConsol.AbreConexion
CentraForm Me
frmReportes.Enabled = False
TxtBCodPers.Enabled = False
'*************Modificado ALPA
'*************28/02/2008
sql = "stp_sel_Documentos '" & gsOpeCod & "'"
Set rsTDoc = oConConsol.CargaRecordSet(sql)
Set rsTDoc.ActiveConnection = Nothing

If gsOpeCod = "760200" Then
    lv.Visible = True
    chktodas.Visible = True
End If

Do While Not rsTDoc.EOF
   cmbDocumento.AddItem Mid(rsTDoc!cDocDesc & space(100), 1, 100) & rsTDoc!nDocTpo
   rsTDoc.MoveNext
Loop
'***************************
Set rs = oOpe.CargaOpeCta(gsOpeCod)
Do While Not rs.EOF
   Set lvItem = lv.ListItems.Add
   lvItem.Text = rs!cCtaContCod
   lvItem.SubItems(1) = rs!cCtaContDesc
   rs.MoveNext
Loop
If gsOpeCod = 760200 Then
   cmdGenerarResumen.Visible = True
   cmdConsolidado.Visible = False
End If
txtFecha = gdFecSis
txtFecha2 = gdFecSis
lbLibroOpen = False
chktodas_Click

' para el registrod e compras
On Error GoTo ConexionErr
 sConRegCom = "DSN=DSNRegCom;"
 Set dbRegCom = New ADODB.Connection
 'Set oCoa = New NContImpreReg
 'CentraForm Me
    dbRegCom.CommandTimeout = 30
    dbRegCom.ConnectionTimeout = 30
    dbRegCom.Open sConRegCom
    'txtAnio.Text = Str(Year(gdFecSis))
    'cboMes.ListIndex = Month(gdFecSis) - 1
    'RotateText 90, Picture1, "Times New Roman", 15, 25, 1500, "C O A"
   Exit Sub
ConexionErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmReportes.Enabled = True
'dbRegCom.Close 'TORE:Comentado por enviar error al cerrar el formulario.
Set dbRegCom = Nothing
End Sub

'***Agregado por ELRO el 20111011, según Acta 278-2011/TI-D
Private Sub lv_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim bDetallarCuentas As Boolean
    
    bDetallarCuentas = True
    
    If Item.Checked = False And Item.Text = "45" Then
        bDetallarCuentas = False
    End If
    
    If bDetallarCuentas = False Then
        chkDetallarCuentas.value = Unchecked
    End If
    
End Sub
'***Fin Agregado por ELRO**********************************

Private Sub optPrinter_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdGenerar.SetFocus
End If
End Sub

Private Sub TxtBCodPers_EmiteDatos()
    Dim oPersona As DConecta
    Set oPersona = New DConecta
    lblNombrePersonaJur.Caption = ""
    oPersona.AbreConexion
    If Trim(TxtBCodPers.Text) = "" Then
        Exit Sub
    End If
    '*************Modificado ALPA
    '*************28/02/2008
    On Error GoTo ArchivoErr
    If TxtBCodPers.Text = "00000000000" Then
    oPersona.CierraConexion
    Else
    Dim TipoDocJ As String
    TipoDocJ = 2
    
    sSql = "stp_sel_DevolverPersonaJPorTipoyNroDoc '" & TxtBCodPers.Text & "', '" & TipoDocJ & "'"
    Set rs = oPersona.CargaRecordSet(sSql)
    Set rs.ActiveConnection = Nothing
    If rs!Nombre = "" Then
        MsgBox "No se pudo encontrar los datos de la Persona," & Chr(10) & " Verifique que la Persona exista", vbInformation, "Aviso"
        Exit Sub
    Else
        lblNombrePersonaJur.Caption = rs!Nombre
    End If
    oPersona.CierraConexion
    End If
    Exit Sub
ArchivoErr:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
    oPersona.CierraConexion
    '***************************
End Sub
Private Sub CargaDatos()
Dim i As Integer
     
'    bEstadoCargando = True
'    SSTDatosGen.Tab = 0
'    SSTabs.Tab = 0
'    'Carga Personeria
'    cmbPersPersoneria.ListIndex = IndiceListaCombo(cmbPersPersoneria, Trim(Str(oPersona.Personeria)))
'
'    'Habilita o Deshabilita Ficha de Persona Juridica
'
'    Call HabilitaFichaPersonaJur(False)
'    Call HabilitaFichaPersonaNat(False)
'
'    'Carga Ubicacion Georgrafica
'    If Len(Trim(oPersona.UbicacionGeografica)) = 12 Then
'        cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & "04028")
'        cmbPersUbiGeo(1).ListIndex = IndiceListaCombo(cmbPersUbiGeo(1), Space(30) & "1" & Mid(oPersona.UbicacionGeografica, 2, 2) & String(9, "0"))
'        cmbPersUbiGeo(2).ListIndex = IndiceListaCombo(cmbPersUbiGeo(2), Space(30) & "2" & Mid(oPersona.UbicacionGeografica, 2, 4) & String(7, "0"))
'        cmbPersUbiGeo(3).ListIndex = IndiceListaCombo(cmbPersUbiGeo(3), Space(30) & "3" & Mid(oPersona.UbicacionGeografica, 2, 6) & String(5, "0"))
'        cmbPersUbiGeo(4).ListIndex = IndiceListaCombo(cmbPersUbiGeo(4), Space(30) & oPersona.UbicacionGeografica)
'    Else
'        cmbPersUbiGeo(0).ListIndex = IndiceListaCombo(cmbPersUbiGeo(0), Space(30) & oPersona.UbicacionGeografica)
'        cmbPersUbiGeo(1).Clear
'        cmbPersUbiGeo(1).AddItem cmbPersUbiGeo(0).Text
'        cmbPersUbiGeo(1).ListIndex = 0
'        cmbPersUbiGeo(2).Clear
'        cmbPersUbiGeo(2).AddItem cmbPersUbiGeo(0).Text
'        cmbPersUbiGeo(2).ListIndex = 0
'        cmbPersUbiGeo(3).Clear
'        cmbPersUbiGeo(3).AddItem cmbPersUbiGeo(0).Text
'        cmbPersUbiGeo(3).ListIndex = 0
'        cmbPersUbiGeo(4).Clear
'        cmbPersUbiGeo(4).AddItem cmbPersUbiGeo(0).Text
'        cmbPersUbiGeo(4).ListIndex = 0
'    End If
'
'    'Carga Direccion
'    txtPersDireccDomicilio.Text = oPersona.Domicilio
'
'    'Selecciona la Condicion del Doicilio
'    cmbPersDireccCondicion.ListIndex = IndiceListaCombo(cmbPersDireccCondicion, oPersona.CondicionDomicilio)
'
'    txtValComercial.Text = Format(oPersona.ValComDomicilio, "#####,###.00")
'
'    'Tipo de sangre
'    If CboTipoSangre.ListCount > 0 Then
'        CboTipoSangre.ListIndex = IndiceListaCombo(CboTipoSangre, oPersona.TipoSangre)
'    End If
'
'    'Carga Ficha 1
'    If oPersona.Personeria = gPersonaNat Then
'        txtPersNombreAP.Text = oPersona.ApellidoPaterno
'        txtPersNombreAM.Text = oPersona.ApellidoMaterno
'        txtPersNombreN.Text = oPersona.Nombres
'    Else
'        txtPersNombreRS.Text = oPersona.NombreCompleto
'    End If
'    TxtTalla.Text = Format(oPersona.Talla, "#0.00")
'    TxtPeso.Text = Format(oPersona.Peso, "#0.00")
'    TxtEmail.Text = oPersona.Email
'
'    txtPersTelefono2.Text = oPersona.Telefonos2
'    If oPersona.Sexo = "F" Then
'        TxtApellidoCasada.Text = oPersona.ApellidoCasada
'        cmbPersNatSexo.ListIndex = 0
'        Call DistribuyeApellidos(True)
'    Else
'        cmbPersNatSexo.ListIndex = 1
'        Call DistribuyeApellidos(False)
'    End If
'    cmbPersNatEstCiv.ListIndex = IndiceListaCombo(cmbPersNatEstCiv, oPersona.EstadoCivil)
'    txtPersNatHijos.Text = Trim(Str(oPersona.Hijos))
'    cmbNacionalidad.ListIndex = IndiceListaCombo(cmbNacionalidad, Space(30) & oPersona.Nacionalidad)
'    chkResidente.value = oPersona.Residencia
'
'    'Carga Datos Generales
'    txtPersNacCreac.Text = Format(oPersona.FechaNacimiento, "dd/mm/yyyy")
'    txtPersTelefono.Text = oPersona.Telefonos
'    CboPersCiiu.ListIndex = IndiceListaCombo(CboPersCiiu, oPersona.CIIU)
'
'    Call CargaControlEstadoPersona(oPersona.Personeria)
'    cmbPersEstado.ListIndex = IndiceListaCombo(cmbPersEstado, Right("0" & oPersona.Estado, 2))
'
'    TxtSiglas.Text = oPersona.Siglas     'Carga Razon Social
'
'    TxtSbs.Text = oPersona.PersCodSbs    'Carga Codigo SBS
'
'    'Selecciona el Tipo de Persona Juridica
'    cmbPersJurTpo.ListIndex = IndiceListaCombo(cmbPersJurTpo, Trim(Str(IIf(oPersona.TipoPersonaJur = "", -1, oPersona.TipoPersonaJur))))
'
'    'Selecciona la relacion Con la Persona
'    CmbRela.ListIndex = IndiceListaCombo(CmbRela, oPersona.PersRelInst)
'
'    'Selecciona la magnitud Empresarial
'    cmbPersJurMagnitud.ListIndex = IndiceListaCombo(cmbPersJurMagnitud, Trim(oPersona.MagnitudEmpresarial))
'
'    'Carga Numero de Empleados
'    txtPersJurEmpleados.Text = Trim(Str(oPersona.NumerosEmpleados))
'
'    Call CargaDocumentos                    'Carga Los Documentos de la Personas
'
'    Call CargaRelacionesPersonas            'Carga las Relaciones de las Personas
'
'    Call CargaFuentesIngreso                'Carga las Fuentes de Ingresos de las Personas
'
'    Call CargaRefComerciales                'Carga las Referencias Comerciales
'    lnNumRefCom = oPersona.MaxRefComercial  'Carga el max Ref Comercial
'
'    Call CargaRefBancarias                  'Carga las Referencias Bancarias
'
'    Call CargaPatVehicular                  'Carga el Patrimonio Vehicular
'    lnNumPatVeh = oPersona.MaxPatVehicular  'Carga el max Pat Vehicular
'
'    'Carga Firma
'    'Call IDBFirma.CargarFirma(oPersona.RFirma)
'
'    If Not oPersona Is Nothing Then
'    If oPersona.Personeria = gPersonaNat Then
'        SSTabs.Tab = 0
'    Else
'        SSTabs.Tab = 1
'    End If
'    End If
'    bEstadoCargando = False
'    CmdPersFteConsultar.Enabled = True
End Sub
Private Sub HabilitaControlesPersonaFtesIngreso(ByVal pbBloqueo As Boolean)
    lv.Enabled = pbBloqueo
    cmbDocumento.Enabled = pbBloqueo
    cmdGenerar.Enabled = pbBloqueo
    cmdGenerarResumen.Enabled = pbBloqueo
    cmdConsolidado.Enabled = pbBloqueo
    cmdSalir.Enabled = pbBloqueo
End Sub
Private Sub txtFecha_GotFocus()
   txtFecha.SelStart = 0
   txtFecha.SelLength = Len(txtFecha.Text)
End Sub

Private Sub txtFecha_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   If ValidaFecha(txtFecha) = "" Then
      txtFecha2.SetFocus
   Else
      MsgBox "Fecha no Valida...", vbInformation, "Error!"
      txtFecha.SelStart = 0
      txtFecha.SelLength = Len(txtFecha.Text)
   End If
End If
End Sub

Private Sub txtFecha2_GotFocus()
   txtFecha2.SelStart = 0
   txtFecha2.SelLength = Len(txtFecha2.Text)
End Sub

Private Sub txtFecha2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   If ValidaFecha(txtFecha2) = "" Then
      Me.optPrinter(1).SetFocus
   Else
      MsgBox "Fecha no Valida...", vbInformation, "Error!"
      txtFecha.SelStart = 0
      txtFecha.SelLength = Len(txtFecha.Text)
   End If
End If
End Sub

Private Sub ImprimeRegCompraGastosExcelCab(pdFecha As Date, pdFecha2 As Date, psEmpresa As String)
xlHoja1.Range("A1:K1").EntireColumn.Font.FontStyle = "Arial"
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 75
xlHoja1.PageSetup.TopMargin = 2
xlHoja1.Range("A9:M1").EntireColumn.Font.Size = 9
xlHoja1.Range("B1").EntireColumn.HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("H1:K1").EntireColumn.NumberFormat = "#,##0.00;-#,##0.00"
xlHoja1.Range("A1:A1").RowHeight = 17
xlHoja1.Range("A1:A1").ColumnWidth = 16
xlHoja1.Range("B1:B1").ColumnWidth = 10
xlHoja1.Range("C1:C1").ColumnWidth = 3
xlHoja1.Range("D1:D1").ColumnWidth = 14
xlHoja1.Range("E1:E1").ColumnWidth = 12
xlHoja1.Range("F1:G1").ColumnWidth = 34
xlHoja1.Range("H1:K1").ColumnWidth = 12
xlHoja1.Range("L1:M1").ColumnWidth = 13
xlHoja1.Range("N1:N1").ColumnWidth = 38
xlHoja1.Range("B1:B1").Font.Size = 12
xlHoja1.Range("A2:B4").Font.Size = 10
xlHoja1.Range("A1:B4").Font.Bold = True
xlHoja1.Cells(1, 2) = "R E G I S T R O   D E   C O M P R A S / G A S T O S"
xlHoja1.Cells(2, 2) = "( DEL " & pdFecha & " AL " & pdFecha2 & " )"
xlHoja1.Cells(4, 1) = "INSTITUCION : " & psEmpresa
xlHoja1.Range("A1:K2").Merge True
xlHoja1.Range("A1:K2").HorizontalAlignment = xlHAlignCenter

xlHoja1.Cells(6, 1) = "Fecha"
xlHoja1.Cells(7, 1) = "Registro"
xlHoja1.Cells(6, 2) = "  D O C U M E N T O  "
xlHoja1.Cells(7, 2) = "Fecha"
xlHoja1.Cells(7, 3) = "Tpo"
xlHoja1.Cells(7, 4) = "Número"
xlHoja1.Cells(6, 5) = "P R O V E E D O R"
xlHoja1.Cells(7, 5) = "R.U.C."
xlHoja1.Cells(7, 6) = "Apellidos y Nombres / Razón Social"
xlHoja1.Cells(6, 7) = "C O N C E P T O"
xlHoja1.Cells(6, 8) = "VALOR"
xlHoja1.Cells(7, 8) = "VENTA"
xlHoja1.Cells(6, 9) = "I.G.V."
xlHoja1.Cells(6, 10) = "OTROS IMP."
xlHoja1.Cells(6, 11) = "PRECIO"
xlHoja1.Cells(7, 11) = "VENTA"
xlHoja1.Cells(6, 12) = "Constancia de Depósito SPOT"
xlHoja1.Cells(7, 12) = "Número"
xlHoja1.Cells(7, 13) = "Fecha"
xlHoja1.Cells(6, 14) = "Operacion"

xlHoja1.Range("B6:D6").Merge False
xlHoja1.Range("E6:F6").Merge False
xlHoja1.Range("L6:M6").Merge False

xlHoja1.Range("A6:N7").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("A6:N7").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
xlHoja1.Range("A6:N7").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A6:N7").Borders(xlInsideVertical).Color = vbBlack
xlHoja1.Range("B6:D6").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("B6:D7").Borders(xlEdgeBottom).Color = vbBlack
xlHoja1.Range("E6:F6").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("L6:M6").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("E6:F7").Borders(xlEdgeBottom).Color = vbBlack

End Sub

Private Sub ImprimeRegCompraGastosExcelCabResumen(pdFecha As Date, pdFecha2 As Date, psEmpresa As String)
xlHoja1.Range("A1:K1").EntireColumn.Font.FontStyle = "Arial"
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 75
xlHoja1.PageSetup.TopMargin = 2
xlHoja1.Range("A9:M1").EntireColumn.Font.Size = 9
xlHoja1.Range("B1").EntireColumn.HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("H1:K1").EntireColumn.NumberFormat = "#,##0.00;-#,##0.00"
xlHoja1.Range("A1:A1").RowHeight = 16
xlHoja1.Range("A1:A1").ColumnWidth = 10
xlHoja1.Range("B1:B1").ColumnWidth = 3
xlHoja1.Range("C1:C1").ColumnWidth = 17
xlHoja1.Range("D1:D1").ColumnWidth = 14
xlHoja1.Range("E1:E1").ColumnWidth = 12
xlHoja1.Range("F1:G1").ColumnWidth = 34
xlHoja1.Range("H1:K1").ColumnWidth = 12
xlHoja1.Range("L1:M1").ColumnWidth = 13
xlHoja1.Range("N1:N1").ColumnWidth = 38
xlHoja1.Range("B1:B1").Font.Size = 12
xlHoja1.Range("A2:B4").Font.Size = 10
xlHoja1.Range("A1:B4").Font.Bold = True
xlHoja1.Cells(1, 2) = "R E G I S T R O   D E   C O M P R A S / G A S T O S"
xlHoja1.Cells(2, 2) = "( DEL " & pdFecha & " AL " & pdFecha2 & " )"
xlHoja1.Cells(4, 1) = "INSTITUCION : " & psEmpresa
xlHoja1.Range("A1:K2").Merge True
xlHoja1.Range("A1:K2").HorizontalAlignment = xlHAlignCenter


xlHoja1.Cells(6, 1) = "  D O C U M E N T O  "
xlHoja1.Cells(7, 1) = "Fecha"
xlHoja1.Cells(7, 2) = "Tpo"
xlHoja1.Cells(7, 3) = "Número"

xlHoja1.Cells(6, 4) = "Fecha"
xlHoja1.Cells(7, 4) = "Registro"

xlHoja1.Cells(6, 5) = "P R O V E E D O R"
xlHoja1.Cells(7, 5) = "R.U.C."
xlHoja1.Cells(7, 6) = "Apellidos y Nombres / Razón Social"
xlHoja1.Cells(6, 7) = "C O N C E P T O"
xlHoja1.Cells(6, 8) = "VALOR"
xlHoja1.Cells(7, 8) = "VENTA"
xlHoja1.Cells(6, 9) = "I.G.V."
xlHoja1.Cells(6, 10) = "OTROS IMP."
xlHoja1.Cells(6, 11) = "PRECIO"
xlHoja1.Cells(7, 11) = "VENTA"
xlHoja1.Cells(6, 12) = "Constancia de Depósito SPOT"
xlHoja1.Cells(7, 12) = "Número"
xlHoja1.Cells(7, 13) = "Fecha"
xlHoja1.Cells(6, 14) = "Operacion"

xlHoja1.Range("A6:C6").Merge False
xlHoja1.Range("E6:F6").Merge False
xlHoja1.Range("L6:M6").Merge False

xlHoja1.Range("A6:N7").HorizontalAlignment = xlHAlignCenter
xlHoja1.Range("A6:N7").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
xlHoja1.Range("A6:N7").Borders(xlInsideVertical).LineStyle = xlContinuous
xlHoja1.Range("A6:N7").Borders(xlInsideVertical).Color = vbBlack
xlHoja1.Range("A6:C6").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("A6:C7").Borders(xlEdgeBottom).Color = vbBlack
xlHoja1.Range("E6:F6").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("L6:M6").Borders(xlEdgeBottom).LineStyle = xlContinuous
xlHoja1.Range("E6:F7").Borders(xlEdgeBottom).Color = vbBlack

End Sub

Public Sub ImprimeRegCompraGastosExcelLima(psCtaCod As String, pdFecha As Date, pdFecha2 As Date, Optional nTipoDoc As Integer = 0, Optional psRuc As String)
    Dim rsCta As New ADODB.Recordset
    Dim rsMay As New ADODB.Recordset
    Dim N As Integer
    Dim sBillete As String, sPersona As String
    Dim sDoc As String
    Dim nImporte As Currency
    Dim nVV As Currency
    Dim nTotVV As Currency, nTotIgv As Currency, nTotISC As Currency, nTotPV As Currency
    Dim nItem As Integer, P As Integer
    Dim nLin As Long
    Dim nIGV As Currency, nISC As Currency, nRetrac As Currency
    Dim nOtroImpuesto As Currency
    Dim sDocs As String
    Dim sCTAS As String
    Dim lnTipoC As Currency
    Dim dfechaRef As String
    Dim lsPersNombre As String
    'ALPA 20090304************************************
    Dim lsPersRUC As String
    '*************************************************
    Dim lsDescripcion As String
    Dim nVenta As Currency
     
    Dim sImpre As String, sImpreAge As String
    Dim lsRepTitulo    As String
    Dim lsHoja         As String
    
    Dim DocNro As Long
    Dim DocNroTemp As Long
    Dim nPosIni As Long
    Dim nPosFin As Long
    Dim w As Long
    Dim lnNoAfecto As Currency '***Agregado por ELRO el 20130604, según TI-ERS064-2013
    'Suma Total
    Dim H As String
    Dim i As String
    Dim j As String
    Dim K As String
    Dim l As String
    Dim X As String
    Dim Y As String
    Dim O As String
    Dim Q As String
    Dim nSumH As Double
    Dim nSumI As Double
    Dim nSumJ As Double
    Dim nSumK As Double
    Dim nSumL As Double
    Dim nSumX As Double
    Dim nSumY As Double
    Dim nSumO As Double
    Dim nSumQ As Double
    nSumH = 0
    nSumI = 0
    nSumJ = 0
    nSumK = 0
    nSumL = 0
    nSumX = 0
    nSumY = 0
    nSumO = 0
    nSumQ = 0
    w = 0
    'On Error GoTo ErrImprime
    
    lsRepTitulo = "REGISTRO DE COMPRAS Y GASTOS"
    
    oBarra.ShowForm Me
    oBarra.CaptionSyle = eCap_CaptionPercent
    oBarra.Max = 1
    'oBarra.Progress 0, lsRepTitulo, "Obteniendo Cuentas", ""
    Dim oOpe As New DOperacion
    'sCTAS = oOpe.CargaListaCuentasOperacion(gsOpeCod, psCtaCod)
    'oBarra.Progress 1, lsRepTitulo, "Obteniendo Cuentas", ""
    
    'Documentos que se consideran en Reporte
    oBarra.Progress 0, lsRepTitulo, "Obteniendo Documentos", ""
    sDocs = oOpe.CargaListaDocsOperacion(gsOpeCod, nTipoDoc)
    Set oOpe = Nothing
    oBarra.Progress 1, lsRepTitulo, "Obteniendo Documentos", ""
    
    'Se tienen que capturar de Operación
    Dim oConst As New NConstSistemas
    sCtaIGV = oConst.LeeConstSistema(gConstCtaIGV)
    sCtaIGVND = oConst.LeeConstSistema(337) 'Add by GITU 10-12-2008
    sCtaISC = oConst.LeeConstSistema(167)
    sCtaRetrac = oConst.LeeConstSistema(166) 'cta detraccion
    sCtaImpCuarta = oConst.LeeConstSistema(169) 'cta Renta 4 categoria
    Set oConst = Nothing
    
    sCTAS = psCtaCod
    oBarra.Progress 0, lsRepTitulo, "Obteniendo Movimientos", ""
    Dim oReg As New NContImpreReg
'    Set rsCta = oReg.GetMovCompraGastos(sCTAS, sDocs, pdFecha, pdFecha2, sCtaIGV, sCtaISC, , True)
'    If rsCta.EOF Then
'       MsgBox "No se registraron movimientos...!", vbInformation, "Resultado"
'       oBarra.CloseForm Me
'       Exit Sub
'    End If
'    oBarra.Progress 1, lsRepTitulo, "Obteniendo Movimientos", ""
    
    oBarra.Progress 0, lsRepTitulo, "Mayorizando Movimientos", ""
    'Se agrego la variable de cta. contable de igv para los'
    'para los proveedores no domiciliados GITU 11-12-2008  '
    Set rsMay = oReg.GetMovCompraGastosMayor(sCTAS, sDocs, pdFecha, pdFecha2, sCtaIGV, sCtaISC, , True, sCtaRetrac, sCtaImpCuarta, psRuc, gsOpeCod, sCtaIGVND)
    'End Gitu'
    nItem = 0: nLin = gnLinPage: P = 0
    oBarra.Progress 1, lsRepTitulo, "Mayorizando Movimientos", ""
    
       nLin = 9
       lsHoja = "C_" & Mid(psCtaCod, 2, 2)
       ExcelAddHoja lsHoja, xlLibro, xlHoja1
       'JEOM  -- Formato texto a una Columna
       xlHoja1.Range(xlHoja1.Cells(1, 6), xlHoja1.Cells(1500, 6)).NumberFormat = "@"
       'FIN JEOM
       'ImprimeRegCompraGastosExcelCabLima pdFecha, pdFecha2, gsNomCmac
      
       oBarra.Max = rsMay.RecordCount
       Do While Not rsMay.EOF
          N = N + 1
          w = w + 1
          nOtroImpuesto = 0
          oBarra.Progress rsMay.Bookmark, lsRepTitulo, "Generando Excel", ""
        'John ----------------------------
          
          If rsMay.Bookmark = 1 Then
             DocNroTemp = rsMay!nDocTpo
             xlHoja1.Cells(5, 1) = IIf(Len(rsMay!nDocTpo) = 1, "0" & rsMay!nDocTpo & space(4) & rsMay!cDocDesc, rsMay!nDocTpo & space(4) & rsMay!cDocDesc)
             ImprimeRegCompraGastosExcelCabLima pdFecha, pdFecha2, gsNomCmac
             xlHoja1.Range("A" & 5 & ":Z" & 8).Font.Bold = True
             nPosIni = 8
          End If
          
          DocNro = rsMay!nDocTpo
       
          If DocNro <> DocNroTemp Then
             xlHoja1.Cells(4 + nPosFin, 1) = IIf(Len(rsMay!nDocTpo) = 1, "0" & rsMay!nDocTpo & space(4) & rsMay!cDocDesc, rsMay!nDocTpo & space(4) & rsMay!cDocDesc)
             ImprimeRegCompraCabeceraRepetida 5 + nPosFin
             xlHoja1.Range("A" & 4 + nPosFin & ":Z" & 7 + nPosFin).Font.Bold = True
             
          End If
          '---------------------------------

        '*** PEAC 20110315
'          If Mid(rsMay!cOpeCod, 3, 1) = "2" And rsMay!nIGV <> 0 Then
'             nIGV = Round(rsMay!nPV - (rsMay!nPV / (1 + gnIGVValor)), 2)
'          Else
'             nIGV = rsMay!nIGV
'          End If
          nIGV = rsMay!nIGV
          '*** FIN PEAC
          
          lnTipoC = rsMay!nMovTpoCambio
          
          '---------------
          Select Case rsMay!nDocTpo
                Case "1" '*** PEAC 20110309
                      
                      nISC = 0
                      nRetrac = rsMay!nRetracc
                      '***Modificado por ELRO el 20130607, según TI-ERS064-2013****
                      'nVV = rsMay!nPV - nIGV
                      lnNoAfecto = rsMay!nISC + rsMay!nNoAfecto
                      nVV = rsMay!nPV - nIGV - lnNoAfecto
                      '***Fin Modificado por ELRO el 20130607, según TI-ERS064-2013
                      nVenta = PrnVal(rsMay!nPV, 14, 2)
                
                Case "2"
                      nIGV = nIGV * -1
                      nISC = rsMay!nISC * -1
                      nRetrac = rsMay!nRetracc
                      nVV = rsMay!nPV - nIGV - nISC
                       
                      nOtroImpuesto = rsMay!nRentaCuarta
                      nVenta = PrnVal(rsMay!nPV, 14, 2) - nOtroImpuesto
                      
                Case "95"
                      nIGV = nIGV * -1
                      nISC = rsMay!nISC * -1
                      nRetrac = rsMay!nRetracc
                      nVV = rsMay!nPV - nIGV - nISC
                      nOtroImpuesto = rsMay!nRentaCuarta
                      nVenta = PrnVal(rsMay!nPV, 14, 2) - nOtroImpuesto
                Case "7"
                      nIGV = nIGV * -1
                      nISC = rsMay!nISC * -1
                      nRetrac = rsMay!nRetracc
                      
                      nVenta = PrnVal(rsMay!nPV * -1, 14, 2)
                
                    'nVV = (rsMay!nPV - nIGV - nISC) * -1 '*** PEAC 20110513
                    nVV = (nVenta - nIGV - nISC)
                
                Case Else
                    '***Modificado por ELRO el 20130607, según TI-ERS064-2013****
                    'nISC = rsMay!nISC
                    'nRetrac = rsMay!nRetracc
                    'nVV = rsMay!nPV - nIGV - nISC
                    'nVenta = PrnVal(rsMay!nPV, 14, 2)
                    If nIGV <> 0 Then
                        nISC = 0
                        nRetrac = rsMay!nRetracc
                        lnNoAfecto = rsMay!nISC + rsMay!nNoAfecto
                        nVV = rsMay!nPV - nIGV - lnNoAfecto
                        nVenta = PrnVal(rsMay!nPV, 14, 2)
                    Else
                        nISC = rsMay!nISC
                        nRetrac = rsMay!nRetracc
                        nVV = rsMay!nPV - nIGV - nISC
                        nVenta = PrnVal(rsMay!nPV, 14, 2)
                    End If
                    '***Fin Modificado por ELRO el 20130607, según TI-ERS064-2013
          End Select
          
'          If rsMay!nDocTpo = 7 Then
'            nIGV = nIGV * -1
'            nISC = rsMay!nISC * -1
'            nRetrac = rsMay!nRetracc
'            nVV = (rsMay!nPV - nIGV - nISC) * -1
'          Else
'            nISC = rsMay!nISC
'            nRetrac = rsMay!nRetracc
'            nVV = rsMay!nPV - nIGV - nISC
'          End If
          
          
          
          '-------------------
          sDoc = IIf(IsNull(rsMay!nDocTpo), space(36), Format(rsMay!dDocFecha, "dd/mm/yyyy") & " " & Mid(rsMay!nDocTpo & space(3), 1, 3) & " " & Mid(rsMay!cDocNro & space(20), 1, 20) & " ")
          If IsNull(rsMay!FechaRef) Then
            dfechaRef = ""
          Else
            dfechaRef = rsMay!FechaRef
          End If
          
          If rsMay!cPersNombreRef = "" Then
             lsPersNombre = IIf(IsNull(rsMay!cPersNombre), "", rsMay!cPersNombre)
             'ALPA 20090304********************************
             lsPersRUC = rsMay!cRuc
             '*********************************************
          Else
             lsPersNombre = IIf(IsNull(rsMay!cPersNombreRef), "", rsMay!cPersNombreRef)
             'ALPA 20090304********************************
             lsPersRUC = rsMay!cRUCRef
             '*********************************************
          End If
          
          lsDescripcion = rsMay!cMovDesc
          
          'JEOM   Obtener Fecha y Nro de Constancia del Pago de las detracciones
          '------------------------------------------------------------------
          Dim rsRef As ADODB.Recordset
          Dim lnMovNro As Long
          Dim lsFechaPagoDetra As String
          Dim lsNroDocPagoDetra As String
          
          Set rsRef = New ADODB.Recordset
          
          lsFechaPagoDetra = ""
          lsNroDocPagoDetra = ""
          lnMovNro = rsMay!nMovNro
          
          Set rsRef = oReg.GetMovCompraGastosMayorRef(lnMovNro)
          
          If Not rsRef.EOF Or Not rsRef.BOF Then
             lsFechaPagoDetra = rsRef!dDocFecha
             lsNroDocPagoDetra = rsRef!cDocNro
          End If
          
          Dim lsRuc As String
          
       
          
          '-------------------------------------------------------------------
          ' FIN JEOM
  
          If DocNro <> DocNroTemp Then
               Dim nVal1 As Long
            
               nVal1 = nPosFin + 8
            
'ALPA 20090304***********************************************************
'             ImpreDetalleExcelLima Mid(rsMay!cMovNro, 1, 8) & "-" & Mid(rsMay!cMovNro, 9, 6) & "-" & Right(Trim(rsMay!cMovNro), 4), _
'                                    IIf(IsNull(rsMay!nDocTpo), Space(10), rsMay!dDocFecha), _
'                                    IIf(IsNull(rsMay!nDocTpo), Space(3), Mid(rsMay!nDocTpo & Space(3), 1, 3)), _
'                                    IIf(IsNull(rsMay!nDocTpo), Space(30), Mid(rsMay!cDocDesc & Space(30), 1, 30)), _
'                                    IIf(IsNull(rsMay!nDocTpo), Space(20), Mid(rsMay!cDocNro & Space(20), 1, 20)), _
'                                    IIf(IsNull(rsMay!cRuc), Space(8), rsMay!cRuc), _
'                                    lsPersNombre, _
'                                    Replace(Replace(rsMay!cMovDesc, Chr(13), " "), oImpresora.gPrnSaltoLinea, ""), _
'                                    PrnVal(nVV, 14, 2), PrnVal(nIGV, 14, 2), PrnVal(nISC, 14, 2), nVenta, _
'                                    nVal1, rsMay!cOC, IIf(IsNull(rsMay!nMovTpoCambio), 0, rsMay!nMovTpoCambio), Trim(rsMay!Tipo), Trim(rsMay!OFIS), PrnVal(nRetrac, 14, 2), CStr(rsMay!nMovNro), w, lnTipoC, dfechaRef, lsDescripcion, nOtroImpuesto, lsFechaPagoDetra, lsNroDocPagoDetra
            ImpreDetalleExcelLima Mid(rsMay!cMovNro, 1, 8) & "-" & Mid(rsMay!cMovNro, 9, 6) & "-" & Right(Trim(rsMay!cMovNro), 4), _
                                    IIf(IsNull(rsMay!nDocTpo), space(10), rsMay!dDocFecha), _
                                    IIf(IsNull(rsMay!nDocTpo), space(3), Mid(rsMay!nDocTpo & space(3), 1, 3)), _
                                    IIf(IsNull(rsMay!nDocTpo), space(30), Mid(rsMay!cDocDesc & space(30), 1, 30)), _
                                    IIf(IsNull(rsMay!nDocTpo), space(20), Mid(rsMay!cDocNro & space(20), 1, 20)), _
                                    IIf(IsNull(lsPersRUC), space(8), lsPersRUC), _
                                    lsPersNombre, _
                                    Replace(Replace(rsMay!cMovDesc, Chr(13), " "), oImpresora.gPrnSaltoLinea, ""), _
                                    PrnVal(nVV, 14, 2), PrnVal(nIGV, 14, 2), PrnVal(nISC, 14, 2), nVenta, _
                                    nVal1, rsMay!cOC, IIf(IsNull(rsMay!nMovTpoCambio), 0, rsMay!nMovTpoCambio), Trim(rsMay!Tipo), Trim(rsMay!OFIS), PrnVal(nRetrac, 14, 2), CStr(rsMay!nMovNro), w, lnTipoC, dfechaRef, lsDescripcion, nOtroImpuesto, lsFechaPagoDetra, lsNroDocPagoDetra, lnNoAfecto
'****************************************************************************
             DocNroTemp = DocNro
             
                          
           xlHoja1.Range("A" & nPosFin + 1 & ":U" & nPosFin + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
           
           xlHoja1.Range("H" & nPosFin + 1 & ":H" & nPosFin + 1).Formula = "=SUM(H" & nPosIni & ":H" & nPosFin & ")"
           xlHoja1.Range("I" & nPosFin + 1 & ":I" & nPosFin + 1).Formula = "=SUM(I" & nPosIni & ":I" & nPosFin & ")"
           xlHoja1.Range("J" & nPosFin + 1 & ":J" & nPosFin + 1).Formula = "=SUM(J" & nPosIni & ":J" & nPosFin & ")"
           xlHoja1.Range("K" & nPosFin + 1 & ":K" & nPosFin + 1).Formula = "=SUM(K" & nPosIni & ":K" & nPosFin & ")"
           xlHoja1.Range("L" & nPosFin + 1 & ":L" & nPosFin + 1).Formula = "=SUM(L" & nPosIni & ":L" & nPosFin & ")"
           xlHoja1.Range("M" & nPosFin + 1 & ":M" & nPosFin + 1).Formula = "=SUM(M" & nPosIni & ":M" & nPosFin & ")"
           xlHoja1.Range("N" & nPosFin + 1 & ":N" & nPosFin + 1).Formula = "=SUM(N" & nPosIni & ":N" & nPosFin & ")"
           'xlHoja1.Range("O" & nPosFin + 1 & ":O" & nPosFin + 1).Formula = "=SUM(O" & nPosIni & ":O" & nPosFin & ")"
           
           xlHoja1.Range("Q" & nPosFin + 1 & ":Q" & nPosFin + 1).Formula = "=SUM(Q" & nPosIni & ":Q" & nPosFin & ")"
    
           xlHoja1.Range("I" & nPosFin + 1 & ":Q" & nPosFin + 1).NumberFormat = "#,##0.00"
           
           ' SUMA TOTAL FINAL

           
'           H = H & " +" & " SUM(H" & nPosIni & ":H" & nPosFin & ")"
'           H = H & " +" & " H" & (nPosIni + 1)
'           I = I & " +" & " SUM(I" & nPosIni & ":I" & nPosFin & ")"
'           J = J & " +" & " SUM(J" & nPosIni & ":J" & nPosFin & ")"
'           K = K & " +" & " SUM(K" & nPosIni & ":K" & nPosFin & ")"
'           L = L & " +" & " SUM(L" & nPosIni & ":L" & nPosFin & ")"
'           X = X & " +" & " SUM(M" & nPosIni & ":M" & nPosFin & ")"
'           Y = Y & " +" & " SUM(N" & nPosIni & ":N" & nPosFin & ")"
'           'O = O & " +" & " SUM(O" & nPosIni & ":O" & nPosFin & ")"
'           Q = Q & " +" & " SUM(Q" & nPosIni & ":Q" & nPosFin & ")"
           
            '               ALPA 17/03/2008
             If PrnVal(nIGV, 14, 2) <> 0 Then
             If rsMay!cOC <> "" Then
             nSumI = nSumI + PrnVal(nVV, 14, 2) '9
             nSumL = nSumL + PrnVal(nIGV, 14, 2) '12
             Else
             nSumH = nSumH + PrnVal(nVV, 14, 2) '8
             nSumK = nSumK + PrnVal(nIGV, 14, 2) '11
             End If
             Else
             nSumJ = nSumJ + PrnVal(nVV, 14, 2) '10
             End If
                             
             If IIf(IsNull(rsMay!nDocTpo), space(3), Mid(rsMay!nDocTpo & space(3), 1, 3)) = 2 Or IIf(IsNull(rsMay!nDocTpo), space(3), Mid(rsMay!nDocTpo & space(3), 1, 3)) = 95 Then
              nSumX = nSumX + nOtroImpuesto '13
             Else
              nSumX = nSumX + PrnVal(nISC, 14, 2) '13
             End If
             nSumY = nSumY + nVenta '14
             nSumQ = nSumQ + PrnVal(nRetrac, 14, 2) '17

           nPosIni = nPosFin + 8
           nPosFin = nPosIni
                    
          Else
             Dim nVal As Long
             
            If rsMay.Bookmark = 1 Then
               nVal = nLin + N - 1
            Else
               nVal = nPosFin + 1
            End If
'ALPA 20090304*******************************************************************
'             ImpreDetalleExcelLima Mid(rsMay!cMovNro, 1, 8) & "-" & Mid(rsMay!cMovNro, 9, 6) & "-" & Right(Trim(rsMay!cMovNro), 4), _
'                                    IIf(IsNull(rsMay!nDocTpo), Space(10), rsMay!dDocFecha), _
'                                    IIf(IsNull(rsMay!nDocTpo), Space(3), Mid(rsMay!nDocTpo & Space(3), 1, 3)), _
'                                    IIf(IsNull(rsMay!nDocTpo), Space(30), Mid(rsMay!cDocDesc & Space(30), 1, 30)), _
'                                    IIf(IsNull(rsMay!nDocTpo), Space(20), Mid(rsMay!cDocNro & Space(20), 1, 20)), _
'                                    IIf(IsNull(rsMay!cRuc), Space(8), rsMay!cRuc), _
'                                     lsPersNombre, _
'                                    Replace(Replace(rsMay!cMovDesc, Chr(13), " "), oImpresora.gPrnSaltoLinea, ""), _
'                                    PrnVal(nVV, 14, 2), PrnVal(nIGV, 14, 2), PrnVal(nISC, 14, 2), nVenta, _
'                                    nVal, rsMay!cOC, IIf(IsNull(rsMay!nMovTpoCambio), 0, rsMay!nMovTpoCambio), Trim(rsMay!Tipo), Trim(rsMay!OFIS), PrnVal(nRetrac, 14, 2), CStr(rsMay!nMovNro), w, lnTipoC, dfechaRef, lsDescripcion, nOtroImpuesto, lsFechaPagoDetra, lsNroDocPagoDetra
                ImpreDetalleExcelLima Mid(rsMay!cMovNro, 1, 8) & "-" & Mid(rsMay!cMovNro, 9, 6) & "-" & Right(Trim(rsMay!cMovNro), 4), _
                                    IIf(IsNull(rsMay!nDocTpo), space(10), rsMay!dDocFecha), _
                                    IIf(IsNull(rsMay!nDocTpo), space(3), Mid(rsMay!nDocTpo & space(3), 1, 3)), _
                                    IIf(IsNull(rsMay!nDocTpo), space(30), Mid(rsMay!cDocDesc & space(30), 1, 30)), _
                                    IIf(IsNull(rsMay!nDocTpo), space(20), Mid(rsMay!cDocNro & space(20), 1, 20)), _
                                    IIf(IsNull(lsPersRUC), space(8), lsPersRUC), _
                                     lsPersNombre, _
                                    Replace(Replace(rsMay!cMovDesc, Chr(13), " "), oImpresora.gPrnSaltoLinea, ""), _
                                    PrnVal(nVV, 14, 2), PrnVal(nIGV, 14, 2), PrnVal(nISC, 14, 2), nVenta, _
                                    nVal, rsMay!cOC, IIf(IsNull(rsMay!nMovTpoCambio), 0, rsMay!nMovTpoCambio), Trim(rsMay!Tipo), Trim(rsMay!OFIS), PrnVal(nRetrac, 14, 2), CStr(rsMay!nMovNro), w, lnTipoC, dfechaRef, lsDescripcion, nOtroImpuesto, lsFechaPagoDetra, lsNroDocPagoDetra, lnNoAfecto
'*********************************************************************************
            If rsMay.Bookmark = 1 Then
               nPosFin = rsMay.Bookmark + 8
            Else
               nPosFin = nPosFin + 1
            End If
            
        If PrnVal(nIGV, 14, 2) <> 0 Then
             If rsMay!cOC <> "" Then
             nSumI = nSumI + PrnVal(nVV, 14, 2) '9
             nSumL = nSumL + PrnVal(nIGV, 14, 2) '12
             Else
             nSumH = nSumH + PrnVal(nVV, 14, 2) '8
             nSumK = nSumK + PrnVal(nIGV, 14, 2) '11
             End If
             Else
             nSumJ = nSumJ + PrnVal(nVV, 14, 2) '10
             End If
                             
             If IIf(IsNull(rsMay!nDocTpo), space(3), Mid(rsMay!nDocTpo & space(3), 1, 3)) = 2 Or IIf(IsNull(rsMay!nDocTpo), space(3), Mid(rsMay!nDocTpo & space(3), 1, 3)) = 95 Then
              nSumX = nSumX + nOtroImpuesto '13
             Else
              nSumX = nSumX + PrnVal(nISC, 14, 2) '13
             End If
              nSumY = nSumY + nVenta '14
             nSumQ = nSumQ + PrnVal(nRetrac, 14, 2) '15
             
          End If
          
          nTotVV = nTotVV + nVV
          nTotPV = nTotPV + rsMay!nPV
          nTotIgv = nTotIgv + nIGV
          nTotISC = nTotISC + nISC
          
          
          
          
         If rsMay.Bookmark = rsMay.RecordCount Then
                xlHoja1.Range("A" & nPosFin + 1 & ":U" & nPosFin + 1).Borders(xlEdgeTop).LineStyle = xlContinuous
       
               'posicion inicia y final
               xlHoja1.Range("H" & nPosFin + 1 & ":H" & nPosFin + 1).Formula = "=SUM(H" & nPosIni & ":H" & nPosFin & ")"
               xlHoja1.Range("I" & nPosFin + 1 & ":I" & nPosFin + 1).Formula = "=SUM(I" & nPosIni & ":I" & nPosFin & ")"
               xlHoja1.Range("J" & nPosFin + 1 & ":J" & nPosFin + 1).Formula = "=SUM(J" & nPosIni & ":J" & nPosFin & ")"
               xlHoja1.Range("K" & nPosFin + 1 & ":K" & nPosFin + 1).Formula = "=SUM(K" & nPosIni & ":K" & nPosFin & ")"
               xlHoja1.Range("L" & nPosFin + 1 & ":L" & nPosFin + 1).Formula = "=SUM(L" & nPosIni & ":L" & nPosFin & ")"
               xlHoja1.Range("M" & nPosFin + 1 & ":M" & nPosFin + 1).Formula = "=SUM(M" & nPosIni & ":M" & nPosFin & ")"
               xlHoja1.Range("N" & nPosFin + 1 & ":N" & nPosFin + 1).Formula = "=SUM(N" & nPosIni & ":N" & nPosFin & ")"
               'xlHoja1.Range("O" & nPosFin + 1 & ":O" & nPosFin + 1).Formula = "=SUM(O" & nPosIni & ":O" & nPosFin & ")"
               
               xlHoja1.Range("Q" & nPosFin + 1 & ":Q" & nPosFin + 1).Formula = "=SUM(Q" & nPosIni & ":Q" & nPosFin & ")"
        
               xlHoja1.Range("I" & nPosFin + 1 & ":Q" & nPosFin + 1).NumberFormat = "#,##0.00"
               
               
                'H = H & " +" & " SUM(H" & nPosIni & ":H" & nPosFin & ")"
'                I = I & " +" & " SUM(I" & nPosIni & ":I" & nPosFin & ")"
'                J = J & " +" & " SUM(J" & nPosIni & ":J" & nPosFin & ")"
'                K = K & " +" & " SUM(K" & nPosIni & ":K" & nPosFin & ")"
'                L = L & " +" & " SUM(L" & nPosIni & ":L" & nPosFin & ")"
'                X = X & " +" & " SUM(M" & nPosIni & ":M" & nPosFin & ")"
'                Y = Y & " +" & " SUM(N" & nPosIni & ":N" & nPosFin & ")"
'               ' O = O & " +" & " SUM(O" & nPosIni & ":O" & nPosFin & ")"
'                Q = Q & " +" & " SUM(Q" & nPosIni & ":Q" & nPosFin & ")"
               
               
               xlHoja1.Range("A" & nPosFin + 4 & ":U" & nPosFin + 4).Borders(xlEdgeTop).LineStyle = xlContinuous
               xlHoja1.Range("A" & nPosFin + 5 & ":U" & nPosFin + 5).Borders(xlEdgeTop).LineStyle = xlContinuous
               
               xlHoja1.Range("A" & nPosFin + 5 & ":U" & nPosFin + 4).Font.Size = 9
               xlHoja1.Range("A" & nPosFin + 5 & ":U" & nPosFin + 4).Font.Bold = True
               
               xlHoja1.Cells(nPosFin + 3, 1) = "TOTAL"
               xlHoja1.Range("A" & nPosFin + 3 & ":A" & nPosFin + 3).Font.Bold = True
               
               'xlHoja1.Range("H" & nPosFin + 4 & ":H" & nPosFin + 4).Formula = "=" & H
               xlHoja1.Cells(nPosFin + 4, 8) = nSumH
               xlHoja1.Cells(nPosFin + 4, 9) = nSumI
               xlHoja1.Cells(nPosFin + 4, 10) = nSumJ
               xlHoja1.Cells(nPosFin + 4, 11) = nSumK
               xlHoja1.Cells(nPosFin + 4, 12) = nSumL
               xlHoja1.Cells(nPosFin + 4, 13) = nSumX
               xlHoja1.Cells(nPosFin + 4, 14) = nSumY
               xlHoja1.Cells(nPosFin + 4, 17) = nSumQ
'               xlHoja1.Range("I" & nPosFin + 4 & ":I" & nPosFin + 4).Formula = "=" & I
'               xlHoja1.Range("J" & nPosFin + 4 & ":J" & nPosFin + 4).Formula = "=" & J
'               xlHoja1.Range("K" & nPosFin + 4 & ":K" & nPosFin + 4).Formula = "=" & K
'               xlHoja1.Range("L" & nPosFin + 4 & ":L" & nPosFin + 4).Formula = "=" & L
'               xlHoja1.Range("M" & nPosFin + 4 & ":M" & nPosFin + 4).Formula = "=" & X
'               xlHoja1.Range("N" & nPosFin + 4 & ":N" & nPosFin + 4).Formula = "=" & Y
'               'xlHoja1.Range("O" & nPosFin + 4 & ":O" & nPosFin + 4).Formula = "=" & O
'               xlHoja1.Range("Q" & nPosFin + 4 & ":Q" & nPosFin + 4).Formula = "=" & Q
               
         End If
           
          
          rsMay.MoveNext
          If rsMay.EOF Then
             Exit Do
          End If
       Loop

    
    RSClose rsCta
    RSClose rsMay
    RSClose rs
    oBarra.CloseForm Me
    
    Exit Sub
ErrImprime:
     MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
       If lbLibroOpen Then
          xlLibro.Close
          xlAplicacion.Quit
       End If
       Set xlAplicacion = Nothing
       Set xlLibro = Nothing
       Set xlHoja1 = Nothing
End Sub

'Private Sub ImprimeRegCompraGastosExcelCabLima(pdFecha As Date, pdFecha2 As Date, psEmpresa As String)
'    xlHoja1.Range("A1:S1").EntireColumn.Font.FontStyle = "Arial"
'    xlHoja1.PageSetup.Orientation = xlLandscape
'    xlHoja1.PageSetup.CenterHorizontally = True
'    xlHoja1.PageSetup.Zoom = 75
'    xlHoja1.PageSetup.TopMargin = 2
'    xlHoja1.Range("A9:K1").EntireColumn.Font.Size = 9
'    xlHoja1.Range("B1").EntireColumn.HorizontalAlignment = xlHAlignCenter
'    'xlHoja1.Range("G1:G1").EntireColumn.NumberFormat = "#,##0.00;-#,##0.00"
'    'xlHoja1.Range("I1:N1").EntireColumn.NumberFormat = "#,##0.00;-#,##0.0000"
'    xlHoja1.Range("A1:A1").RowHeight = 17
'    xlHoja1.Range("A1:A1").ColumnWidth = 16
'    xlHoja1.Range("B1:B1").ColumnWidth = 10
'    xlHoja1.Range("C1:C1").ColumnWidth = 14
'    xlHoja1.Range("D1:D1").ColumnWidth = 12
'    xlHoja1.Range("E1:E1").ColumnWidth = 10
'    xlHoja1.Range("F1:F1").ColumnWidth = 14
'    xlHoja1.Range("G1:G1").ColumnWidth = 12
'    xlHoja1.Range("H1:H1").ColumnWidth = 34
'    xlHoja1.Range("I1:J1").ColumnWidth = 20
'    xlHoja1.Range("K1:K1").ColumnWidth = 18
'    xlHoja1.Range("L1:L1").ColumnWidth = 20
'    xlHoja1.Range("M1:M1").ColumnWidth = 20
'    xlHoja1.Range("N1:N1").ColumnWidth = 16
'    xlHoja1.Range("O1:O1").ColumnWidth = 20
'    xlHoja1.Range("P1:P1").ColumnWidth = 18
'    xlHoja1.Range("Q1:Q1").ColumnWidth = 16
'    xlHoja1.Range("R1:R1").ColumnWidth = 16
'    xlHoja1.Range("S1:S1").ColumnWidth = 16
'    xlHoja1.Range("T1:T1").ColumnWidth = 16
'    xlHoja1.Range("U1:U1").ColumnWidth = 16
'    xlHoja1.Range("V1:V1").ColumnWidth = 16
'
'    xlHoja1.Range("B1:B1").Font.Size = 12
'    xlHoja1.Range("A2:B4").Font.Size = 10
'    xlHoja1.Range("A1:B4").Font.Bold = True
'    xlHoja1.Cells(1, 2) = "R E G I S T R O   D E   C O M P R A S"
'    xlHoja1.Cells(2, 2) = "( DEL " & pdFecha & " AL " & pdFecha2 & " )"
'    xlHoja1.Cells(4, 1) = "INSTITUCION : " & psEmpresa
'    xlHoja1.Range("A1:N2").Merge True
'    xlHoja1.Range("A1:N2").HorizontalAlignment = xlHAlignCenter
'
'    xlHoja1.Cells(6, 1) = "FECHA DE"
'    xlHoja1.Cells(7, 1) = "EMISION"
'
'    xlHoja1.Cells(6, 2) = "FECHA DE "
'    xlHoja1.Cells(7, 2) = "VENCIMIENT."
'    xlHoja1.Cells(8, 2) = "SERVICIO"
'
'    xlHoja1.Cells(6, 3) = "FECHA DE "
'    xlHoja1.Cells(7, 3) = "PAGO DEL"
'    xlHoja1.Cells(8, 3) = "SERVICIO"
'
'    xlHoja1.Cells(6, 4) = "TIPO DE"
'    xlHoja1.Cells(7, 4) = "COMPROB"
'
'    xlHoja1.Cells(6, 5) = "COMPROBANTE"
'    xlHoja1.Cells(7, 5) = "SERIE"
'    xlHoja1.Cells(7, 6) = "NUMERO"
'
'    xlHoja1.Cells(6, 7) = "RUC DEL"
'    xlHoja1.Cells(7, 7) = "PROVEEDOR"
'
'    xlHoja1.Cells(6, 8) = "RAZON SOCIAL Y/O"
'    xlHoja1.Cells(7, 8) = "APELLIDOS Y"
'    xlHoja1.Cells(6, 8) = "NOMBRES"
'
'
'    xlHoja1.Cells(6, 9) = "BASE IMPONIBLE"
'    xlHoja1.Cells(7, 9) = "DEST.A VENTAS"
'    xlHoja1.Cells(8, 9) = "GRAVAD.EXCLUSIV."
'
'    xlHoja1.Cells(6, 10) = "BASE IMPONIBLE"
'    xlHoja1.Cells(7, 10) = "DEST.A VENTAS NO"
'    xlHoja1.Cells(8, 10) = "GRAVAD.EXCLUSIV."
'
'    xlHoja1.Cells(6, 11) = "VALOR DE"
'    xlHoja1.Cells(7, 11) = "ADQUISICION"
'    xlHoja1.Cells(8, 11) = "NO GRABADAS"
'
'    xlHoja1.Cells(6, 12) = "IGV."
'    xlHoja1.Cells(7, 12) = "DEST.A VENTAS"
'    xlHoja1.Cells(8, 12) = "GRAVAD.EXCLUS"
'
'    xlHoja1.Cells(6, 13) = "IGV"
'    xlHoja1.Cells(7, 13) = "DEST.A VENTAS NO"
'    xlHoja1.Cells(8, 13) = "GRAVAD.EXCLUS"
'
'    xlHoja1.Cells(6, 14) = "OTROS TRIBUTO"
'    xlHoja1.Cells(7, 14) = "O CARGOS NO"
'    xlHoja1.Cells(8, 14) = "IMPONIBLES"
'
'    xlHoja1.Cells(6, 15) = "TOTAL"
'    xlHoja1.Cells(7, 15) = "ADQUISICION"
'
'    xlHoja1.Cells(6, 16) = "DETRACCIÓN"
'
'    xlHoja1.Cells(6, 17) = "N°de CONST."
'    xlHoja1.Cells(7, 17) = "DE DEPOSITO"
'    xlHoja1.Cells(8, 17) = "DETRACCION"
'
'    xlHoja1.Cells(6, 18) = "FECHA DE"
'    xlHoja1.Cells(7, 18) = "DEPOSITO"
'    xlHoja1.Cells(8, 18) = "DETRACCION"
'
'    xlHoja1.Cells(6, 19) = "REFERENCIA"
'    xlHoja1.Cells(7, 19) = "N/ DEBITO"
'    xlHoja1.Cells(8, 19) = "N/ CREDITO"
'
'    'xlHoja1.Range("B6:D6").Merge False
'    xlHoja1.Range("E6:F6").Merge True
'    'xlHoja1.Range("G6:J6").Merge False
'
'    xlHoja1.Range("L6:N7").HorizontalAlignment = xlHAlignCenter
'
'    xlHoja1.Range("A6:S8").HorizontalAlignment = xlHAlignCenter
'    xlHoja1.Range("A6:S8").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
'    xlHoja1.Range("A6:S8").Borders(xlInsideVertical).LineStyle = xlContinuous
'    xlHoja1.Range("A6:S8").Borders(xlInsideVertical).Color = vbBlack
'    'xlHoja1.Range("B6:D6").Borders(xlEdgeBottom).LineStyle = xlContinuous
'    'xlHoja1.Range("B6:D7").Borders(xlEdgeBottom).Color = vbBlack
'    xlHoja1.Range("E6:G6").Borders(xlEdgeBottom).LineStyle = xlContinuous
'    'xlHoja1.Range("E6:J7").Borders(xlEdgeBottom).Color = vbBlack
'
'    With xlHoja1.PageSetup
'        .LeftHeader = ""
'        .CenterHeader = ""
'        .RightHeader = ""
'        .LeftFooter = ""
'        .CenterFooter = ""
'        .RightFooter = ""
''        .LeftMargin = Application.InchesToPoints(0)
''        .RightMargin = Application.InchesToPoints(0)
''        .TopMargin = Application.InchesToPoints(2.77777777777778E-02)
''        .BottomMargin = Application.InchesToPoints(0)
''        .HeaderMargin = Application.InchesToPoints(0)
''        .FooterMargin = Application.InchesToPoints(0)
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .CenterHorizontally = True
'        .CenterVertically = False
'        .Orientation = xlLandscape
'        .Draft = False
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False
'        .Zoom = 55
'    End With
'End Sub

Private Sub ImprimeRegCompraGastosExcelCabLima(pdFecha As Date, pdFecha2 As Date, psEmpresa As String)
    xlHoja1.Range("A1:S1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    xlHoja1.Range("A9:Z1").EntireColumn.Font.Size = 7
    xlHoja1.Range("B1").EntireColumn.HorizontalAlignment = xlHAlignCenter
    
    xlHoja1.Range("A1:A1").RowHeight = 17
    xlHoja1.Range("A1:A1").ColumnWidth = 8
    xlHoja1.Range("B1:B1").ColumnWidth = 10
    xlHoja1.Range("C1:C1").ColumnWidth = 14
    xlHoja1.Range("D1:D1").ColumnWidth = 12
    xlHoja1.Range("E1:E1").ColumnWidth = 10
    xlHoja1.Range("F1:F1").ColumnWidth = 14
    xlHoja1.Range("G1:G1").ColumnWidth = 38
    xlHoja1.Range("H1:H1").ColumnWidth = 20
    xlHoja1.Range("I1:J1").ColumnWidth = 18
    xlHoja1.Range("K1:K1").ColumnWidth = 20
    xlHoja1.Range("L1:L1").ColumnWidth = 20
    xlHoja1.Range("M1:M1").ColumnWidth = 16
    xlHoja1.Range("N1:N1").ColumnWidth = 20
    xlHoja1.Range("O1:O1").ColumnWidth = 18
    xlHoja1.Range("P1:P1").ColumnWidth = 25
    xlHoja1.Range("Q1:Q1").ColumnWidth = 16
    xlHoja1.Range("R1:R1").ColumnWidth = 20
    xlHoja1.Range("S1:S1").ColumnWidth = 16
    'xlHoja1.Range("T1:T1").ColumnWidth = 16 comentado por gitu 10/04/08
    xlHoja1.Range("U1:U1").ColumnWidth = 16
    xlHoja1.Range("V1:V1").ColumnWidth = 16
    xlHoja1.Range("W1:W1").ColumnWidth = 16
    xlHoja1.Range("X1:X1").ColumnWidth = 16
    
    xlHoja1.Range("B1:B1").Font.Size = 12
    xlHoja1.Range("A2:B4").Font.Size = 10
    xlHoja1.Range("A1:B4").Font.Bold = True
    xlHoja1.Cells(1, 2) = "R E G I S T R O   D E   C O M P R A S"
    xlHoja1.Cells(2, 2) = "( DEL " & pdFecha & " AL " & pdFecha2 & " )"
    xlHoja1.Cells(4, 1) = "INSTITUCION : " & psEmpresa
    xlHoja1.Range("A1:N2").Merge True
    xlHoja1.Range("A1:N2").HorizontalAlignment = xlHAlignCenter
    
    
    xlHoja1.Cells(6, 1) = "ITEM"
        
    xlHoja1.Cells(6, 2) = "FECHA DE"
    xlHoja1.Cells(7, 2) = "EMISION"
    
    xlHoja1.Cells(6, 3) = "FECHA DE VENCIM "
    xlHoja1.Cells(7, 3) = "DEL SERV.O FECHA"
    xlHoja1.Cells(8, 3) = "PAGO"
    
'    xlHoja1.Cells(6, 4) = "FECHA DE "
'    xlHoja1.Cells(7, 4) = "PAGO DEL"
'    xlHoja1.Cells(8, 4) = "SERVICIO"

    
    xlHoja1.Cells(6, 4) = "COMPROBANTE"
    xlHoja1.Cells(7, 4) = "SERIE"
    xlHoja1.Cells(7, 5) = "NUMERO"
    
    xlHoja1.Cells(6, 6) = "RUC DEL"
    xlHoja1.Cells(7, 6) = "PROVEEDOR"
    
    xlHoja1.Cells(6, 7) = "APELLIDOS Y NOMBRES,DENOMINACION "
    xlHoja1.Cells(7, 7) = "O RAZON SOCIAL"
              
    
    xlHoja1.Cells(6, 8) = "BASE IMPONIBLE"
    xlHoja1.Cells(7, 8) = "DEST.A OPERACIONES"
    xlHoja1.Cells(8, 8) = "GRAVADAS"
    
    xlHoja1.Cells(6, 9) = "BASE IMPONIBLE"
    xlHoja1.Cells(7, 9) = "DEST.A OPERACIONES NO"
    xlHoja1.Cells(8, 9) = "GRAVADAS"
    
    xlHoja1.Cells(6, 10) = "VALOR DE"
    xlHoja1.Cells(7, 10) = "ADQUISICION"
    xlHoja1.Cells(8, 10) = "NO GRABADAS"
    
    xlHoja1.Cells(6, 11) = "IGV."
    xlHoja1.Cells(7, 11) = "DEST.A OPERACIONES"
    xlHoja1.Cells(8, 11) = "GRAVADAS"
    
    xlHoja1.Cells(6, 12) = "IGV"
    xlHoja1.Cells(7, 12) = "DEST.A OPERACIONES NO"
    xlHoja1.Cells(8, 12) = "GRAVADAS"
    
    xlHoja1.Cells(6, 13) = "OTROS TRIBUTO"
    xlHoja1.Cells(7, 13) = "Y/O CARGOS"
        
    xlHoja1.Cells(6, 14) = "TOTAL"
    xlHoja1.Cells(7, 14) = "ADQUISICION"
    
'----
    xlHoja1.Cells(6, 15) = "TIPO"
    xlHoja1.Cells(7, 15) = "DE CAMBIO"
    
    xlHoja1.Cells(6, 16) = "DETALLE"
'-----

    xlHoja1.Cells(6, 17) = "CONSTANCIA DEPOSITO DETRACCION"
    xlHoja1.Cells(7, 17) = "DETRACCIÓN"
    
    xlHoja1.Cells(7, 18) = "N°de CONSTANCIA"
    xlHoja1.Cells(8, 18) = "DEPOSITO DETRACCION"
    
    
    xlHoja1.Cells(7, 19) = "FECHA DE DEPOSITO"
    xlHoja1.Cells(8, 19) = "DETRACCION"
    
    'xlHoja1.Cells(6, 20) = "REFERENCIA" comentado por gitu 10/04/08
    'xlHoja1.Cells(7, 20) = "N/ DEBITO"
    'xlHoja1.Cells(8, 20) = "N/ CREDITO"
    
    xlHoja1.Range("D6:E6").Merge True
    xlHoja1.Range("Q6:S6").Merge True
    
    xlHoja1.Range("L6:N7").HorizontalAlignment = xlHAlignCenter
        
    xlHoja1.Range("A6:T8").HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A6:T8").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A6:T8").Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A6:T8").Borders(xlInsideVertical).Color = vbBlack
    xlHoja1.Range("D6:E6").Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("Q6:S6").Borders(xlEdgeBottom).LineStyle = xlContinuous


    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""

        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
End Sub


Private Sub ImprimeRegCompraCabeceraRepetida(ByVal i As Long)
            
    xlHoja1.Cells(i, 1) = "ITEM"
        
    xlHoja1.Cells(i, 2) = "FECHA DE"
    xlHoja1.Cells(i + 1, 2) = "EMISION"
    
    xlHoja1.Cells(i, 3) = "FECHA DE VENCIM "
    xlHoja1.Cells(i + 1, 3) = "DEL SERV.O FECHA"
    xlHoja1.Cells(i + 2, 3) = "PAGO"
    
'    xlHoja1.Cells(I, 4) = "FECHA DE "
'    xlHoja1.Cells(I + 1, 4) = "PAGO DEL"
'    xlHoja1.Cells(I + 2, 4) = "SERVICIO"
    
    xlHoja1.Cells(i, 4) = "COMPROBANTE"
    xlHoja1.Cells(i + 1, 4) = "SERIE"
    xlHoja1.Cells(i + 1, 5) = "NUMERO"
    
    xlHoja1.Cells(i, 6) = "RUC DEL"
    xlHoja1.Cells(i + 1, 6) = "PROVEEDOR"
    
    xlHoja1.Cells(i, 7) = "APELLIDOS Y NOMBRES,DENOMINACION "
    xlHoja1.Cells(i + 1, 7) = "O RAZON SOCIAL"
              
    
    xlHoja1.Cells(i, 8) = "BASE IMPONIBLE"
    xlHoja1.Cells(i + 1, 8) = "DEST.A OPERACIONES"
    xlHoja1.Cells(i + 2, 8) = "GRAVADAS"
    
    xlHoja1.Cells(i, 9) = "BASE IMPONIBLE"
    xlHoja1.Cells(i + 1, 9) = "DEST.A OPERACIONES NO"
    xlHoja1.Cells(i + 2, 9) = "GRAVADAS"
    
    xlHoja1.Cells(i, 10) = "VALOR DE"
    xlHoja1.Cells(i + 1, 10) = "ADQUISICION"
    xlHoja1.Cells(i + 2, 10) = "NO GRABADAS"
    
    xlHoja1.Cells(i, 11) = "IGV."
    xlHoja1.Cells(i + 1, 11) = "DEST.A OPERACIONES"
    xlHoja1.Cells(i + 2, 11) = "GRAVADAS"
    
    xlHoja1.Cells(i, 12) = "IGV"
    xlHoja1.Cells(i + 1, 12) = "DEST.A OPERACIONES NO"
    xlHoja1.Cells(i + 2, 12) = "GRAVADAS"
    
    xlHoja1.Cells(i, 13) = "OTROS TRIBUTO"
    xlHoja1.Cells(i + 1, 13) = "Y/O CARGOS"
        
    xlHoja1.Cells(i, 14) = "TOTAL"
    xlHoja1.Cells(i + 1, 14) = "ADQUISICION"
    
'----
    xlHoja1.Cells(i, 15) = "TIPO"
    xlHoja1.Cells(i + 1, 15) = "DE CAMBIO"
    
    xlHoja1.Cells(i, 16) = "DETALLE"
'-----

    xlHoja1.Cells(i, 17) = "CONSTANCIA DEPOSITO DETRACCION"
    xlHoja1.Cells(i + 1, 17) = "DETRACCIÓN"
    
    xlHoja1.Cells(i + 1, 18) = "N°de CONSTANCIA"
    xlHoja1.Cells(i + 2, 18) = "DEPOSITO DETRACCION"
    
    
    xlHoja1.Cells(i + 1, 19) = "FECHA DE DEPOSITO"
    xlHoja1.Cells(i + 2, 19) = "DETRACCION"
    
'    xlHoja1.Cells(I, 20) = "REFERENCIA"
'    xlHoja1.Cells(I + 1, 20) = "N/ DEBITO"
'    xlHoja1.Cells(I + 2, 20) = "N/ CREDITO"
    
    
    
    
    xlHoja1.Range("D" & i & ":E" & i).Merge True
    xlHoja1.Range("Q" & i & ":S" & i).Merge True
    
    xlHoja1.Range("L" & i & ":N" & i).HorizontalAlignment = xlHAlignCenter
        
    xlHoja1.Range("A" & i & ":T" & i + 2).HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A" & i & ":T" & i + 2).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A" & i & ":T" & i + 2).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A" & i & ":T" & i + 2).Borders(xlInsideVertical).Color = vbBlack
    xlHoja1.Range("D" & i & ":E" & i).Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("Q" & i & ":S" & i).Borders(xlEdgeBottom).LineStyle = xlContinuous

End Sub




'Private Sub ImpreDetalleExcelLima(sFec As String, sDocFecha As String, sDocTpo As String, sDocNro, sRUC As String, sPer As String, sCon As String, nVV As Currency, nIGV As Currency, nISC As Currency, nPV As Currency, nLin As Integer, psOC As String, pnTpoCambio As String, _
'                                sDocAbrev As String, sOFIS As String, nRetrac As Currency, sMovNro As String)
'    If IsDate(sDocFecha) Then
'        xlHoja1.Cells(nLin, 1) = CDate(sDocFecha)
'    Else
'        xlHoja1.Cells(nLin, 1) = CDate(Mid(sFec, 7, 2) & "/" & Mid(sFec, 5, 2) & "/" & Left(sFec, 4))
'    End If
'
'    xlHoja1.Cells(nLin, 4) = Format(sDocTpo, "'00") 'sDocAbrev 'sDocTpo
'
'    If InStr(1, sDocNro, "-") <> 0 Then
'        xlHoja1.Cells(nLin, 5) = "'" & Format(Left(sDocNro, InStr(1, sDocNro, "-") - 1), "000")
'        xlHoja1.Cells(nLin, 6) = "'" & Format(Mid(sDocNro, InStr(1, sDocNro, "-") + 1), "000000000")
'    Else
'        xlHoja1.Cells(nLin, 6) = "'" & sDocNro
'    End If
'
'    xlHoja1.Cells(nLin, 7) = sRUC
'    xlHoja1.Cells(nLin, 8) = sPer
'
'
'    If nIGV <> 0 Then
'        If psOC <> "" Then 'Bienes
'            xlHoja1.Cells(nLin, 10) = nVV
'            xlHoja1.Cells(nLin, 13) = nIGV
'        Else
'            xlHoja1.Cells(nLin, 10) = nVV
'            xlHoja1.Cells(nLin, 13) = nIGV
'        End If
'    Else
'        xlHoja1.Cells(nLin, 11) = nVV
'    End If
'
'    xlHoja1.Cells(nLin, 14) = nISC
'    xlHoja1.Cells(nLin, 15) = nPV
'
'    'Retraccion
'    xlHoja1.Cells(nLin, 16) = nRetrac
'    '**************
'    nLin = nLin + 1
'End Sub


Private Sub ImpreDetalleExcelLima(sFec As String, sDocFecha As String, sDocTpo As String, sDocDesc As String, sDocNro, sRUC As String, sPer As String, sCon As String, nVV As Currency, nIGV As Currency, nISC As Currency, nPV As Currency, nLin As Long, psOC As String, pnTpoCambio As String, _
                                sDocAbrev As String, sOFIS As String, nRetrac As Currency, sMovNro As String, nFilas As Long, lnTipoC As Currency, lsFechaRef As String, lsDescripcion As String, OtroImpuesto As Currency, _
                                 psFechaDocDetra As String, psNroDocDetra As String, Optional pnNoAfecto As Currency = 0)
  Dim ldFechaDoc As String
  Dim ldFecRef As String
  Dim lsSerie As String
  Dim lsNumero As String
  Dim lsRuc As String
  Dim lsDescripcionVF As String
  Dim nBasImpGra As Double, nBasImpNoGra As Double
  Dim nValAdqGra As Double
  Dim nIgvgra As Double, nIgvNoGra As Double
  Dim nOtrosTrib As Double, nTotAdq As Double
  Dim nTipoCambio As Double
    
    xlHoja1.Cells(nLin, 1) = nFilas
    If IsDate(sDocFecha) Then
        xlHoja1.Cells(nLin, 2) = CDate(sDocFecha)
        ldFechaDoc = CDate(sDocFecha)
    Else
        xlHoja1.Cells(nLin, 2) = CDate(Mid(sFec, 7, 2) & "/" & Mid(sFec, 5, 2) & "/" & Left(sFec, 4))
        ldFechaDoc = CDate(Mid(sFec, 7, 2) & "/" & Mid(sFec, 5, 2) & "/" & Left(sFec, 4))
    End If
    
    
'    If Trim(sCon) = "Programa Seguridad Bancos - Enero 2011" Then
'        MsgBox "sdfsd"
'    End If
'
    
    If lsFechaRef <> "" Then xlHoja1.Cells(nLin, 3) = CDate(lsFechaRef)
    If lsFechaRef <> "" Then ldFecRef = CDate(lsFechaRef)
    
    'xlHoja1.Cells(nLin, 4) = Format(sDocTpo, "'00") 'sDocAbrev 'sDocTpo
    
    If InStr(1, sDocNro, "-") <> 0 Then
        xlHoja1.Cells(nLin, 4) = "'" & Format(Left(sDocNro, InStr(1, sDocNro, "-") - 1), "000")
        lsSerie = Format(Left(sDocNro, InStr(1, sDocNro, "-") - 1), "000")
        xlHoja1.Cells(nLin, 5) = "'" & Format(Mid(sDocNro, InStr(1, sDocNro, "-") + 1), "000000000")
        lsNumero = Format(Mid(sDocNro, InStr(1, sDocNro, "-") + 1), "000000000")
    Else
        xlHoja1.Cells(nLin, 5) = "'" & sDocNro
        lsNumero = sDocNro
    End If
    
    xlHoja1.Cells(nLin, 6) = "'" & sRUC
    lsRuc = sRUC
    
    xlHoja1.Cells(nLin, 7) = sPer

    If nIGV <> 0 Then
        If psOC <> "" Then 'Bienes
            xlHoja1.Cells(nLin, 9) = nVV
            nBasImpNoGra = nVV
            
            '*** PEAC 20101104
'            If nPV <> nVV Then
'                xlHoja1.Cells(nLin, 10) = nPV - nVV - nIGV
'                nValAdqGra = nPV - nVV - nIGV
'            End If
            '*** FIN PEAC
            
            xlHoja1.Cells(nLin, 12) = nIGV
            nIgvNoGra = nIGV
        Else
            xlHoja1.Cells(nLin, 8) = nVV
            nBasImpGra = nVV
            xlHoja1.Cells(nLin, 11) = nIGV
            nIgvgra = nIGV
        End If
        xlHoja1.Cells(nLin, 10) = pnNoAfecto '***Agregado por ELRO el 20130607, según TI-ERS064-2013
    Else
        If sDocTpo = 2 Or sDocTpo = 95 Then '*** PEAC 20110309
            xlHoja1.Cells(nLin, 10) = nVV
            nValAdqGra = nVV
        Else
            xlHoja1.Cells(nLin, 10) = nPV ''nVV
            nValAdqGra = nVV
        End If
    End If

    If sDocTpo = 2 Or sDocTpo = 95 Then
       xlHoja1.Cells(nLin, 13) = OtroImpuesto
       nOtrosTrib = OtroImpuesto
    Else
       xlHoja1.Cells(nLin, 13) = nISC
       nOtrosTrib = nISC
    End If
    
    xlHoja1.Cells(nLin, 14) = nPV
    nTotAdq = nPV
    xlHoja1.Cells(nLin, 15) = Format(lnTipoC, "0.000")
    nTipoCambio = Format(lnTipoC, "0.000")
    xlHoja1.Cells(nLin, 16) = lsDescripcion
    lsDescripcionVF = Mid(Trim(CHRTRAN(lsDescripcion, Chr(10) & "'", "")), 1, 100)
    xlHoja1.Cells(nLin, 17) = nRetrac
    xlHoja1.Cells(nLin, 18) = psNroDocDetra
    xlHoja1.Cells(nLin, 19) = psFechaDocDetra

    dbRegCom.BeginTrans
    SQLRegCom = "INSERT INTO RegCom(nItem,dFecEmis,dFecVenc,cSerie,cNumero,cRuc,cRazon," _
         & " nBaseImpgr,nBaImpnogr,nVadqnogr,nIgvdespgr,nIgvDepnog,nOtrosTrib,nTotAdq,nTipocambi,cDetalle," _
         & " nDetraccio,cnroconst,dFecDepDet,cDocTpo, cDocDesc,crefdebcre) VALUES( " & nFilas & ", " _
         & " '" & ldFechaDoc & "', '" & ldFecRef & "', '" & lsSerie & "', '" & lsNumero & "', '" & lsRuc & "', '" & Trim(sPer) & "', " & nBasImpGra & "," _
         & " " & nBasImpNoGra & ", " & nValAdqGra & ", " & nIgvgra & ", " & nIgvNoGra & ", " _
         & " " & nOtrosTrib & ", " & nTotAdq & ", " & nTipoCambio & ", '" & lsDescripcionVF & "', " _
         & " " & nRetrac & ", '" & psNroDocDetra & "', '" & psFechaDocDetra & "', '" & sDocTpo & "', '" & sDocDesc & "','0')"
    dbRegCom.Execute SQLRegCom

    dbRegCom.CommitTrans
    nLin = nLin + 1
End Sub

'Se quito la condicion de cuentas contables psCtaCod As String By gitu 14-01-2009
Public Function ImprimeConsolidado(pdFecha As Date, pdFecha2 As Date, nTipoDoc As Integer)
    Dim rsCta As New ADODB.Recordset
    Dim rsMay As New ADODB.Recordset
    Dim N As Integer
    Dim sBillete As String, sPersona As String
    Dim sDoc As String
    Dim nImporte As Currency
    Dim nCantidad As Long
    Dim nTipo  As Long
    Dim nDesc As String
    
    Dim nItem As Integer, P As Integer
    Dim nLin As Long
    Dim sDocs As String
    Dim sCTAS As String
    Dim lnTipoC As Currency
    
    Dim sImpre As String, sImpreAge As String
    Dim lsRepTitulo    As String
    Dim lsHoja         As String
    Dim nOtroImpuesto As Currency
    
        
    lsRepTitulo = "REGISTRO DE COMPRAS Y GASTOS - CONSOLIDADO"
    
    oBarra.ShowForm Me
    oBarra.CaptionSyle = eCap_CaptionPercent
    oBarra.Max = 1
    
    Dim oOpe As New DOperacion
           
    oBarra.Progress 0, lsRepTitulo, "Obteniendo Documentos", ""
    sDocs = oOpe.CargaListaDocsOperacion(gsOpeCod)
    Set oOpe = Nothing
    oBarra.Progress 1, lsRepTitulo, "Obteniendo Documentos", ""
    
    'Se tienen que capturar de Operación
    Dim oConst As New NConstSistemas
    sCtaIGV = oConst.LeeConstSistema(gConstCtaIGV)
    sCtaISC = oConst.LeeConstSistema(167)
    sCtaRetrac = oConst.LeeConstSistema(166)
    sCtaImpCuarta = oConst.LeeConstSistema(169)
    Set oConst = Nothing
    'comentado by gitu 14-01-2009
    'sCTAS = psCtaCod
    oBarra.Progress 0, lsRepTitulo, "Obteniendo Movimientos", ""
    Dim oReg As New NContImpreReg

    oBarra.Progress 0, lsRepTitulo, "Mayorizando Movimientos", ""
    'comentado by gitu 14-01-2009
    'Set rsMay = oReg.GetCompraGastosConsolidado(sCTAS, sDocs, pdFecha, pdFecha2, sCtaIGV, sCtaISC, , True, sCtaRetrac, sCtaImpCuarta)
    Set rsMay = oReg.GetCompraGastosConsolidado(sDocs, pdFecha, pdFecha2, sCtaIGV, sCtaISC, , True, sCtaRetrac, sCtaImpCuarta)
    'End Gitu
    nItem = 0: nLin = gnLinPage: P = 0
    oBarra.Progress 1, lsRepTitulo, "Mayorizando Movimientos", ""
    
       nLin = 7
       lsHoja = "RESUMEN " '& Mid(psCtaCod, 2, 2)
       ExcelAddHoja lsHoja, xlLibro, xlHoja1
       
       ImprimeCabeceraRegCompraConsolidado pdFecha, pdFecha2, gsNomCmac
      
       Do While Not rsMay.EOF
          
          nTipo = rsMay!Tipo
          nDesc = rsMay!Descripcion
          nImporte = rsMay!Total
          nCantidad = rsMay!Cantidad
          
          Select Case rsMay!Tipo
                Case "2"
                      nOtroImpuesto = rsMay!Renta
                      nImporte = PrnVal(rsMay!Total, 14, 2) - nOtroImpuesto
                Case "95"
                      nOtroImpuesto = rsMay!Renta
                      nImporte = PrnVal(rsMay!Total, 14, 2) - nOtroImpuesto
                Case Else
                      nImporte = rsMay!Total
                      
          End Select
          
          
          
          ImpreDetalleRegCompraConsolidado nCantidad, nTipo & space(5) & nDesc, nImporte, nLin
                      
          rsMay.MoveNext
          If rsMay.EOF Then
             Exit Do
          End If
       Loop
       
           xlHoja1.Range("A" & nLin & ":C" & nLin).Borders(xlEdgeTop).LineStyle = xlContinuous
           xlHoja1.Range("C" & nLin & ":C" & nLin).Formula = "=SUM(C7:C" & nLin - 1 & ")"
           xlHoja1.Range("C" & nLin & ":C" & nLin).NumberFormat = "#,##0.00"
           xlHoja1.Range("C" & nLin & ":C" & nLin).Font.Bold = True
          

    
    RSClose rsCta
    RSClose rsMay
    RSClose rs
    oBarra.CloseForm Me
    
    Exit Function
ErrImprime:
     MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
       If lbLibroOpen Then
          xlLibro.Close
          xlAplicacion.Quit
       End If
       Set xlAplicacion = Nothing
       Set xlLibro = Nothing
       Set xlHoja1 = Nothing
End Function



Private Sub ImprimeCabeceraRegCompraConsolidado(pdFecha As Date, pdFecha2 As Date, psEmpresa As String)
    xlHoja1.Range("A1:C1").EntireColumn.Font.FontStyle = "Arial"
    xlHoja1.PageSetup.Orientation = xlLandscape
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 75
    xlHoja1.PageSetup.TopMargin = 2
    xlHoja1.Range("A9:C1").EntireColumn.Font.Size = 7
    xlHoja1.Range("A7").EntireColumn.HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("B1").EntireColumn.HorizontalAlignment = xlHAlignLeft
    
    xlHoja1.Range("A1:A1").RowHeight = 17
    xlHoja1.Range("A1:A1").ColumnWidth = 8
    xlHoja1.Range("B1:B1").ColumnWidth = 35
    xlHoja1.Range("C1:C1").ColumnWidth = 10

    
    xlHoja1.Range("B1:B1").Font.Size = 12
    xlHoja1.Range("A2:B4").Font.Size = 10
    xlHoja1.Range("A1:C4").Font.Bold = True
    xlHoja1.Cells(1, 2) = "R E G I S T R O   D E   C O M P R A S   C O N S O L I D A D O"
    xlHoja1.Cells(2, 2) = "( DEL " & pdFecha & " AL " & pdFecha2 & " )"
    xlHoja1.Cells(4, 2) = psEmpresa
    xlHoja1.Range("A1:F2").Merge True
    xlHoja1.Range("A1:F2").HorizontalAlignment = xlHAlignCenter
    
    
    xlHoja1.Cells(6, 1) = "Nro."
    xlHoja1.Cells(6, 2) = "DOCUMENTO"
    xlHoja1.Cells(6, 3) = "TOTAL "
        
                
    
    xlHoja1.Range("A1:C6").Font.Bold = True
    xlHoja1.Range("A6:A6").HorizontalAlignment = xlHAlignCenter
    xlHoja1.Range("A6:C6").BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
    xlHoja1.Range("A6:C6").Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range("A6:C6").Borders(xlInsideVertical).Color = vbBlack
    xlHoja1.Range("A6:C6").Borders(xlEdgeBottom).LineStyle = xlContinuous
    xlHoja1.Range("A6:C6").Borders(xlEdgeBottom).LineStyle = xlContinuous


    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = ""
        .RightHeader = ""
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""

        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = True
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 55
    End With
End Sub

Private Sub ImpreDetalleRegCompraConsolidado(pnCantidad As Long, psDesc As String, pnMonto As Currency, nLin As Long)
    
    xlHoja1.Cells(nLin, 1) = pnCantidad
    xlHoja1.Cells(nLin, 2) = psDesc
    xlHoja1.Cells(nLin, 3) = Format(pnMonto, "0.00")
    

    nLin = nLin + 1
End Sub

