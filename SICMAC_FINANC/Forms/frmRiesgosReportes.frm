VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRiesgosReportes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Riesgos : Reportes "
   ClientHeight    =   7455
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "frmRiesgosReportes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7455
   ScaleWidth      =   4275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fra1 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      ForeColor       =   &H80000008&
      Height          =   2025
      Left            =   45
      TabIndex        =   22
      Top             =   5280
      Width           =   4035
      Begin VB.CheckBox chkuntitular 
         Caption         =   "Check1"
         Height          =   195
         Left            =   960
         TabIndex        =   29
         Top             =   1080
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.ComboBox cmbAgencia 
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   720
         Width           =   1695
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
         Height          =   375
         Left            =   480
         TabIndex        =   4
         Top             =   1320
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   2040
         TabIndex        =   5
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Frame fraTCambio 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   2010
         TabIndex        =   25
         Top             =   0
         Visible         =   0   'False
         Width           =   2025
         Begin VB.TextBox TxtTipoC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1140
            TabIndex        =   3
            Top             =   195
            Width           =   765
         End
         Begin VB.Label Label35 
            Caption         =   "Tipo Cambio"
            Height          =   210
            Left            =   120
            TabIndex        =   26
            Top             =   262
            Width           =   900
         End
      End
      Begin VB.Frame fraNroClientes 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   615
         Left            =   0
         TabIndex        =   23
         Top             =   0
         Visible         =   0   'False
         Width           =   2025
         Begin VB.TextBox txtNroClientes 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   345
            Left            =   1140
            MaxLength       =   10
            TabIndex        =   2
            Top             =   173
            Width           =   750
         End
         Begin VB.Label Label31 
            Caption         =   "Nro Clientes"
            Height          =   210
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Label lbluntitular 
         Caption         =   "Solo un Titular por Cuenta"
         Height          =   255
         Left            =   1320
         TabIndex        =   30
         Top             =   1080
         Visible         =   0   'False
         Width           =   2055
      End
      Begin VB.Label Agencia 
         Caption         =   "Agencias"
         Height          =   255
         Left            =   240
         TabIndex        =   28
         Top             =   720
         Width           =   855
      End
   End
   Begin VB.Frame fraProductos 
      Caption         =   "Productos"
      ForeColor       =   &H00800000&
      Height          =   5055
      Left            =   105
      TabIndex        =   21
      Top             =   120
      Visible         =   0   'False
      Width           =   4050
      Begin MSComctlLib.TreeView TreeView1 
         Height          =   4740
         Left            =   75
         TabIndex        =   1
         Top             =   210
         Width           =   3900
         _ExtentX        =   6879
         _ExtentY        =   8361
         _Version        =   393217
         LabelEdit       =   1
         Style           =   7
         Checkboxes      =   -1  'True
         ImageList       =   "imglstFiguras"
         Appearance      =   1
      End
   End
   Begin VB.Frame fraProductos1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Linea Credito"
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
      Height          =   4065
      Left            =   4440
      TabIndex        =   0
      Top             =   405
      Width           =   1815
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Pignoraticio"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   11
         Left            =   360
         TabIndex        =   20
         Tag             =   "305"
         Top             =   3720
         Width           =   1365
      End
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "CTS"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   8
         Left            =   360
         TabIndex        =   19
         Tag             =   "303"
         Top             =   2871
         Width           =   1245
      End
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Plazo Fijo"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   7
         Left            =   360
         TabIndex        =   18
         Tag             =   "302"
         Top             =   2658
         Width           =   1230
      End
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Dscto x Planilla"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   6
         Left            =   360
         TabIndex        =   17
         Tag             =   "301"
         Top             =   2430
         Width           =   1425
      End
      Begin VB.CheckBox chkCredConsumo1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Consumo"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   2160
         Width           =   1530
      End
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Pesquero"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   5
         Left            =   375
         TabIndex        =   15
         Tag             =   "103"
         Top             =   1890
         Width           =   1260
      End
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Agropecuario"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   4
         Left            =   375
         TabIndex        =   14
         Tag             =   "102"
         Top             =   1680
         Width           =   1260
      End
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Empresarial"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   225
         Index           =   3
         Left            =   375
         TabIndex        =   13
         Tag             =   "101"
         Top             =   1470
         Width           =   1260
      End
      Begin VB.CheckBox chkCredComercial1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Comercial"
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   90
         TabIndex        =   12
         Top             =   1200
         Width           =   1245
      End
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Pesquero"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   375
         TabIndex        =   11
         Tag             =   "203"
         Top             =   1005
         Width           =   1290
      End
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Agropecuario"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   375
         TabIndex        =   10
         Tag             =   "202"
         Top             =   780
         Width           =   1290
      End
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Empresarial"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   375
         TabIndex        =   9
         Tag             =   "201"
         Top             =   540
         Width           =   1290
      End
      Begin VB.CheckBox chkCredMES1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "MicroEmpresa "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   90
         TabIndex        =   8
         Top             =   240
         Width           =   1485
      End
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Usos Diversos"
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   9
         Left            =   360
         TabIndex        =   7
         Tag             =   "304"
         Top             =   3099
         Width           =   1365
      End
      Begin VB.CheckBox chkCred1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Administ. Trab."
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   210
         Index           =   10
         Left            =   360
         TabIndex        =   6
         Tag             =   "320"
         Top             =   3330
         Width           =   1365
      End
   End
   Begin MSComctlLib.ImageList imglstFiguras 
      Left            =   5010
      Top             =   5070
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRiesgosReportes.frx":030A
            Key             =   "Padre"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRiesgosReportes.frx":065C
            Key             =   "Bebe"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRiesgosReportes.frx":09AE
            Key             =   "Hijo"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRiesgosReportes.frx":0D00
            Key             =   "Hijito"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmRiesgosReportes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim fsCodReport As String
Dim fdFechaDataRep As Date
Dim sservidorconsolidada As String
Dim cVigente As String
Dim cPigno As String

Public Sub Inicio(ByVal psCodReport As String, ByVal pdFechaDataRep As Date)
    fsCodReport = psCodReport
    fdFechaDataRep = pdFechaDataRep
End Sub

Private Sub cmdProcesar_Click()
Dim lsArchivo   As String
Dim lbLibroOpen As Boolean
Dim N           As Integer
Dim lsFechaRep  As String
Dim lsFec As String
Dim lsFecIniMes As String

If fraNroClientes.Visible = True Then
    If Val(txtNroClientes.Text) = 0 Then
        MsgBox "Ingrese un número de clientes válido", vbExclamation, "Aviso"
        txtNroClientes.SetFocus
        Exit Sub
    End If
End If
If fraTCambio.Visible = True Then
    If Val(TxtTipoC.Text) = 0 Then
        MsgBox "Ingrese un Tipo de Cambio Válido", vbExclamation, "Aviso"
        TxtTipoC.SetFocus
        Exit Sub
    End If
End If

lsFechaRep = Format(DateAdd("d", gdFecSis, -1 * Day(gdFecSis)), "mm/dd/yyyy")
lsFec = Format(lsFechaRep, "yyyymmdd")

Select Case fsCodReport
    Case gRiesgoCalfCarCred  ' Calificacion Cartera Credit x Analista
        lsArchivo = App.path & "\Spooler\CalifxAnalista" & lsFec & ".xls"
        lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
        If lbLibroOpen Then
            
            Call GeneraRepCalifCarteraxAnalista(lsFechaRep)
            ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
            CargaArchivo lsArchivo, App.path & "\Spooler"
        End If

    Case gRiesgoCalfAltoRiesgo  ' Cartera de Alto Riesgo
        lsArchivo = App.path & "\Spooler\CartAltoRiesgo" & lsFec & ".xls"
        lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
        If lbLibroOpen Then
            Call GeneraRepCartAltoRiesgo(lsFechaRep)
            ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
            CargaArchivo lsArchivo, App.path & "\Spooler"
        End If
        
    Case gRiesgoConceCarCred  ' Concentracion Cartera Crediticia
                 
        
        lsArchivo = App.path & "\Spooler\ConcentracionCred" & lsFec & ".xls"
        lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
        If lbLibroOpen Then
            lsFecIniMes = Format(DateAdd("d", -1 * Day(Format(lsFechaRep, "mm/dd/yyyy")), Format(lsFechaRep, "mm/dd/yyyy")), "mm/dd/yyyy")
            Call GeneraRepConcentracionCred(CDbl(TxtTipoC.Text), lsFecIniMes, Format(lsFechaRep, "mm/dd/yyyy"))
            ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
            CargaArchivo lsArchivo, App.path & "\Spooler"
        End If
'


    Case gRiesgoEstratDepPlazo  ' Estratos de Depositos a Plazos
        lsArchivo = App.path & "\Spooler\EstratosDepositos" & lsFec & ".xls"
        lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
        If lbLibroOpen Then
            'ExcelAddHoja "EstratosDeposit", xlLibro, xlHoja1
            Call GeneraRepEstratosDepositos(lsFechaRep)
            ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
            CargaArchivo lsArchivo, App.path & "\Spooler"
        End If

    Case gRiesgoPrincipClientesAhorros  ' Principales Clientes Captaciones (Ahorros)
        lsArchivo = App.path & "\Spooler\PrinClientesCap" & lsFec & ".xls"
        lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
        If lbLibroOpen Then
            Call GeneraRepCapMejoresClientes(Val(txtNroClientes), Val(TxtTipoC), lsFechaRep, Trim(Right(cmbAgencia.Text, 8)))
            ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
            CargaArchivo lsArchivo, App.path & "\Spooler"
        End If

    Case gRiesgoPrincipClientesCreditos  ' Principales Clientes Colocaciones (Creditos)
        lsArchivo = App.path & "\Spooler\PrinClientesCol" & lsFec & ".xls"
        lbLibroOpen = ExcelBegin(lsArchivo, xlAplicacion, xlLibro)
        If lbLibroOpen Then
            Call GeneraRepColMejoresClientes(Val(txtNroClientes), Val(TxtTipoC), lsFechaRep, Trim(Right(cmbAgencia.Text, 8)))
            ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
            CargaArchivo lsArchivo, App.path & "\Spooler"
        End If

    
End Select

End Sub


Private Sub GeneraRepCalifCarteraxAnalista(ByVal pdFecha As Date)
    Dim lnFila As Integer
    Dim lsConec As String
    Dim lsSql As String
    Dim lrAge As New ADODB.Recordset
    Dim lrReg As New ADODB.Recordset
    Dim I As Integer, lnIIni As Integer
    Dim lsCadCond As String, lsCadCondDesc As String
    Dim lsCadTotal() As String, lnContAgencia As Integer
    Dim lsCadTotales As String, J As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta

    If gbBitCentral = False Then
        oCon.AbreConexion 'Remota "07", , , "03"
    Else
        oCon.AbreConexion
    End If
    
   lsCadCond = GetProdsMarcados
   lsCadCondDesc = GetProdsMarcadosDesc
    
   ExcelAddHoja "CalifxAnalista", xlLibro, xlHoja1
   Call GeneraRepCalifCarteraxAnalistaCab(pdFecha, lsCadCondDesc)
  
    If gbBitCentral = False Then
        lsSql = " Select cCodTab, cValor, cNomTab From dbcomunes..Tablacod Where cCodTab like '47%'" _
              & " And cValor like '112%' Order By cValor "
    Else
        lsSql = "select cAgeCod as cCodTab, cAgeCod as cValor, cAgeDescripcion as cNomTab from agencias Order by cAgeCod"
    End If
   lrAge.CursorLocation = adUseClient
   Set lrAge = oCon.CargaRecordSet(lsSql)
   Set lrAge.ActiveConnection = Nothing
   ReDim lsCadTotal(lrAge.RecordCount)
   lnContAgencia = 1
   I = 10
   Do While Not lrAge.EOF
        
        'Imprime Cabecera
        xlHoja1.Cells(I, 1) = Mid(lrAge!cValor, 4, 2) & " " & Trim(lrAge!cNomtab)
        xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 1)).Font.Bold = True
        I = I + 1
   
   
        If gbBitCentral = True Then
             
            'Pepe
            lsSql = " SELECT c.cCodAnalista, " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '1' THEN c.cctacod END )  As NumSoles, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '1' THEN c.nSaldoCap END ), 0 )  As SKSoles, " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '1' And c.cCalGen = '0' THEN c.cctacod END )  As NumCal0Sol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '1' And c.cCalGen = '0' THEN c.nSaldoCap END ), 0 )  As SKCal0Sol, " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '1' And c.cCalGen = '1' THEN c.cctacod END )  As NumCal1Sol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '1' And c.cCalGen = '1' THEN c.nSaldoCap END ), 0 )  As SKCal1Sol, " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '1' And c.cCalGen = '2' THEN c.cctacod END )  As NumCal2Sol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '1' And c.cCalGen = '2' THEN c.nSaldoCap END ), 0 )  As SKCal2Sol, " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '1' And c.cCalGen = '3' THEN c.cctacod END )  As NumCal3Sol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '1' And c.cCalGen = '3' THEN c.nSaldoCap END ), 0 )  As SKCal3Sol, " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '1' And c.cCalGen = '4' THEN c.cctacod END )  As NumCal4Sol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '1' And c.cCalGen = '4' THEN c.nSaldoCap END ), 0 )  As SKCal4Sol, " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '2' THEN c.cctacod END )  As NumDolar, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '2' THEN c.nSaldoCap END ), 0 )  As SKDolar, " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '2' And c.cCalGen = '0' THEN c.cctacod END )  As NumCal0Dol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '2' And c.cCalGen = '0' THEN c.nSaldoCap END ), 0 )  As SKCal0Dol,  " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '2' And c.cCalGen = '1' THEN c.cctacod END )  As NumCal1Dol,  " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '2' And c.cCalGen = '1' THEN c.nSaldoCap END ), 0 )  As SKCal1Dol, " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '2' And c.cCalGen = '2' THEN c.cctacod END )  As NumCal2Dol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '2' And c.cCalGen = '2' THEN c.nSaldoCap END ), 0 )  As SKCal2Dol, " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '2' And c.cCalGen = '3' THEN c.cctacod END )  As NumCal3Dol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '2' And c.cCalGen = '3' THEN c.nSaldoCap END ), 0 )  As SKCal3Dol, " _
                & " COUNT( CASE WHEN Substring(c.cctacod,9,1) = '2' And c.cCalGen = '4' THEN c.cctacod END )  As NumCal4Dol,  " _
                & " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) = '2' And c.cCalGen = '4' THEN c.nSaldoCap END ), 0 )  As SKCal4Dol "
            lsSql = lsSql & " " _
                & " FROM  ColocCalifProv c " _
                & " WHERE c.nPrdEstado in (" & cVigente & ", " & cPigno & ") And Substring(c.cctacod,4,2) = '" & Trim(lrAge!cValor) & "' " _
                & " And nSaldoCap > 0 " & lsCadCond _
                & " Group by c.cCodAnalista " _
                & " Order by c.cCodAnalista "
            
        Else
   
        '=================== SALDO ACTUAL =====================
            lsSql = " SELECT c.cCodAnalista, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '1' THEN c.cCodCta END )  As NumSoles, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '1' THEN c.nSaldoCap END ), 0 )  As SKSoles, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '1' And c.cCalGen = '0' THEN c.cCodCta END )  As NumCal0Sol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '1' And c.cCalGen = '0' THEN c.nSaldoCap END ), 0 )  As SKCal0Sol, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '1' And c.cCalGen = '1' THEN c.cCodCta END )  As NumCal1Sol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '1' And c.cCalGen = '1' THEN c.nSaldoCap END ), 0 )  As SKCal1Sol, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '1' And c.cCalGen = '2' THEN c.cCodCta END )  As NumCal2Sol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '1' And c.cCalGen = '2' THEN c.nSaldoCap END ), 0 )  As SKCal2Sol, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '1' And c.cCalGen = '3' THEN c.cCodCta END )  As NumCal3Sol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '1' And c.cCalGen = '3' THEN c.nSaldoCap END ), 0 )  As SKCal3Sol, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '1' And c.cCalGen = '4' THEN c.cCodCta END )  As NumCal4Sol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '1' And c.cCalGen = '4' THEN c.nSaldoCap END ), 0 )  As SKCal4Sol, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '2' THEN c.cCodCta END )  As NumDolar, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '2' THEN c.nSaldoCap END ), 0 )  As SKDolar, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '2' And c.cCalGen = '0' THEN c.cCodCta END )  As NumCal0Dol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '2' And c.cCalGen = '0' THEN c.nSaldoCap END ), 0 )  As SKCal0Dol,  " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '2' And c.cCalGen = '1' THEN c.cCodCta END )  As NumCal1Dol,  " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '2' And c.cCalGen = '1' THEN c.nSaldoCap END ), 0 )  As SKCal1Dol, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '2' And c.cCalGen = '2' THEN c.cCodCta END )  As NumCal2Dol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '2' And c.cCalGen = '2' THEN c.nSaldoCap END ), 0 )  As SKCal2Dol, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '2' And c.cCalGen = '3' THEN c.cCodCta END )  As NumCal3Dol, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '2' And c.cCalGen = '3' THEN c.nSaldoCap END ), 0 )  As SKCal3Dol, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '2' And c.cCalGen = '4' THEN c.cCodCta END )  As NumCal4Dol,  " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '2' And c.cCalGen = '4' THEN c.nSaldoCap END ), 0 )  As SKCal4Dol "
            lsSql = lsSql & " " _
                & " FROM CreditoAudi c " _
                & " WHERE c.cestado in ('F','V','1','4','6','7') And Substring(c.cCodCta,1,2) = '" & Mid(lrAge!cValor, 4, 2) & "' " _
                & " And nSaldoCap > 0 " & lsCadCond _
                & " Group by c.cCodAnalista " _
                & " Order by c.cCodAnalista "
        End If
        
        lrReg.CursorLocation = adUseClient
        Set lrReg = oCon.CargaRecordSet(lsSql)
        Set lrReg.ActiveConnection = Nothing
        
        lnIIni = I
        Do While Not lrReg.EOF
        
            xlHoja1.Cells(I, 1) = lrReg!cCodAnalista
            xlHoja1.Cells(I, 2) = lrReg!NumSoles
            xlHoja1.Cells(I, 3) = lrReg!skSoles
            xlHoja1.Cells(I, 4) = lrReg!NumCal0Sol
            xlHoja1.Cells(I, 5) = lrReg!SKCal0Sol
            xlHoja1.Cells(I, 6) = lrReg!NumCal1Sol
            xlHoja1.Cells(I, 7) = lrReg!SKCal1Sol
            xlHoja1.Cells(I, 8) = lrReg!NumCal2Sol
            xlHoja1.Cells(I, 9) = lrReg!SKCal2Sol
            xlHoja1.Cells(I, 10) = lrReg!NumCal3Sol
            xlHoja1.Cells(I, 11) = lrReg!SKCal3Sol
            xlHoja1.Cells(I, 12) = lrReg!NumCal4Sol
            xlHoja1.Cells(I, 13) = lrReg!SKCal4Sol
            
            xlHoja1.Cells(I, 14) = lrReg!NumDolar
            xlHoja1.Cells(I, 15) = lrReg!skDolar
            xlHoja1.Cells(I, 16) = lrReg!NumCal0Dol
            xlHoja1.Cells(I, 17) = lrReg!SKCal0Dol
            xlHoja1.Cells(I, 18) = lrReg!NumCal1Dol
            xlHoja1.Cells(I, 19) = lrReg!SKCal1Dol
            xlHoja1.Cells(I, 20) = lrReg!NumCal2Dol
            xlHoja1.Cells(I, 21) = lrReg!SKCal2Dol
            xlHoja1.Cells(I, 22) = lrReg!NumCal3Dol
            xlHoja1.Cells(I, 23) = lrReg!SKCal3Dol
            xlHoja1.Cells(I, 24) = lrReg!NumCal4Dol
            xlHoja1.Cells(I, 25) = lrReg!SKCal4Dol
            
            I = I + 1
            lrReg.MoveNext
        Loop
        lrReg.Close
        If I <> lnIIni Then
             '***** Suma Totales Agencia
            xlHoja1.Range("A" & I & ":A" & I) = "Total Agencia"
            xlHoja1.Range("B" & I & ":B" & I).Formula = "=SUM(B" & lnIIni & ":B" & I - 1 & ")"
            xlHoja1.Range("C" & I & ":C" & I).Formula = "=SUM(C" & lnIIni & ":C" & I - 1 & ")"
            xlHoja1.Range("D" & I & ":D" & I).Formula = "=SUM(D" & lnIIni & ":D" & I - 1 & ")"
            xlHoja1.Range("E" & I & ":E" & I).Formula = "=SUM(E" & lnIIni & ":E" & I - 1 & ")"
            xlHoja1.Range("F" & I & ":F" & I).Formula = "=SUM(F" & lnIIni & ":F" & I - 1 & ")"
            xlHoja1.Range("G" & I & ":G" & I).Formula = "=SUM(G" & lnIIni & ":G" & I - 1 & ")"
            xlHoja1.Range("H" & I & ":H" & I).Formula = "=SUM(H" & lnIIni & ":H" & I - 1 & ")"
            xlHoja1.Range("I" & I & ":I" & I).Formula = "=SUM(I" & lnIIni & ":I" & I - 1 & ")"
            xlHoja1.Range("J" & I & ":J" & I).Formula = "=SUM(J" & lnIIni & ":J" & I - 1 & ")"
            xlHoja1.Range("K" & I & ":K" & I).Formula = "=SUM(K" & lnIIni & ":K" & I - 1 & ")"
            xlHoja1.Range("L" & I & ":L" & I).Formula = "=SUM(L" & lnIIni & ":L" & I - 1 & ")"
            xlHoja1.Range("M" & I & ":M" & I).Formula = "=SUM(M" & lnIIni & ":M" & I - 1 & ")"
            xlHoja1.Range("N" & I & ":N" & I).Formula = "=SUM(N" & lnIIni & ":N" & I - 1 & ")"
            xlHoja1.Range("O" & I & ":O" & I).Formula = "=SUM(O" & lnIIni & ":O" & I - 1 & ")"
            xlHoja1.Range("P" & I & ":P" & I).Formula = "=SUM(P" & lnIIni & ":P" & I - 1 & ")"
            xlHoja1.Range("Q" & I & ":Q" & I).Formula = "=SUM(Q" & lnIIni & ":Q" & I - 1 & ")"
            xlHoja1.Range("R" & I & ":R" & I).Formula = "=SUM(R" & lnIIni & ":R" & I - 1 & ")"
            xlHoja1.Range("S" & I & ":S" & I).Formula = "=SUM(S" & lnIIni & ":S" & I - 1 & ")"
            xlHoja1.Range("T" & I & ":T" & I).Formula = "=SUM(T" & lnIIni & ":T" & I - 1 & ")"
            xlHoja1.Range("U" & I & ":U" & I).Formula = "=SUM(U" & lnIIni & ":U" & I - 1 & ")"
            xlHoja1.Range("V" & I & ":V" & I).Formula = "=SUM(V" & lnIIni & ":V" & I - 1 & ")"
            xlHoja1.Range("W" & I & ":W" & I).Formula = "=SUM(W" & lnIIni & ":W" & I - 1 & ")"
            xlHoja1.Range("X" & I & ":X" & I).Formula = "=SUM(X" & lnIIni & ":X" & I - 1 & ")"
            xlHoja1.Range("Y" & I & ":Y" & I).Formula = "=SUM(Y" & lnIIni & ":Y" & I - 1 & ")"
            xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 25)).Font.Bold = True
        End If
        lsCadTotal(lnContAgencia) = I
        lnContAgencia = lnContAgencia + 1
        
        I = I + 1
        
        lrAge.MoveNext
    
   Loop
   I = I + 1
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 25)).Font.Bold = True
   xlHoja1.Range("A" & I & ":A" & I) = " TOTAL"
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "B" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("B" & I & ":B" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "C" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("C" & I & ":C" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "C" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("C" & I & ":C" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "D" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("D" & I & ":D" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "E" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("E" & I & ":E" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "F" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("F" & I & ":F" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "G" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("G" & I & ":G" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "H" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("H" & I & ":H" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "I" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("I" & I & ":I" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "J" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("J" & I & ":J" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "K" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("K" & I & ":K" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "L" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("L" & I & ":L" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "M" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("M" & I & ":M" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "N" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("N" & I & ":N" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "O" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("O" & I & ":O" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "P" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("P" & I & ":P" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "Q" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("Q" & I & ":Q" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "R" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("R" & I & ":R" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "S" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("S" & I & ":S" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "T" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("T" & I & ":T" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "U" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("U" & I & ":U" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "V" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("V" & I & ":V" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "W" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("W" & I & ":W" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "X" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("X" & I & ":X" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "Y" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("Y" & I & ":Y" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   
   xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(I + 2, 25)).Borders(xlInsideVertical).LineStyle = xlContinuous
   xlHoja1.Range("A8:Y" & I + 2).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
   
   oCon.CierraConexion
   
   MsgBox "Reporte Generado Satisfactoriamente"
End Sub

Private Sub GeneraRepCalifCarteraxAnalistaCab(ByVal pdFecha As Date, ByVal psCondiciones As String)
xlHoja1.PageSetup.Orientation = xlLandscape
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 55
xlHoja1.Cells(1, 1) = gsNomCmac
xlHoja1.Cells(2, 1) = "CALIFICACION DE CARTERA POR ANALISTA"
xlHoja1.Cells(3, 1) = "Al " & Format(pdFecha, "dd/mm/yyyy")

xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 24)).Font.Bold = True
xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 24)).Merge True
xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 24)).HorizontalAlignment = xlCenter

xlHoja1.Cells(4, 1) = psCondiciones

xlHoja1.Cells(6, 1) = "Analista"
xlHoja1.Cells(7, 1) = "Creditos"
xlHoja1.Cells(5, 2) = "Moneda Nacional"
xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 13)).Merge True
xlHoja1.Cells(6, 2) = "Cartera"
xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 3)).Merge True
xlHoja1.Cells(7, 2) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(150, 2)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 3) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 3), xlHoja1.Cells(150, 3)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 4) = "Calif. (0)"
xlHoja1.Range(xlHoja1.Cells(6, 4), xlHoja1.Cells(6, 5)).Merge True
xlHoja1.Cells(7, 4) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 4), xlHoja1.Cells(250, 4)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 5) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 5), xlHoja1.Cells(250, 5)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 6) = "Calif. (1)"
xlHoja1.Range(xlHoja1.Cells(6, 6), xlHoja1.Cells(6, 7)).Merge True
xlHoja1.Cells(7, 6) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 6), xlHoja1.Cells(250, 6)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 7) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 7), xlHoja1.Cells(250, 7)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 8) = "Calif. (2)"
xlHoja1.Range(xlHoja1.Cells(6, 8), xlHoja1.Cells(6, 9)).Merge True
xlHoja1.Cells(7, 8) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 8), xlHoja1.Cells(250, 8)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 9) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 9), xlHoja1.Cells(250, 9)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 10) = "Calif. (3)"
xlHoja1.Range(xlHoja1.Cells(6, 10), xlHoja1.Cells(6, 11)).Merge True
xlHoja1.Cells(7, 10) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 10), xlHoja1.Cells(250, 10)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 11) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 11), xlHoja1.Cells(250, 11)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 12) = "Calif. (4)"
xlHoja1.Range(xlHoja1.Cells(6, 12), xlHoja1.Cells(6, 13)).Merge True
xlHoja1.Cells(7, 12) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 12), xlHoja1.Cells(250, 12)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 13) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 13), xlHoja1.Cells(250, 13)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(5, 14) = "Moneda Extranjera"
xlHoja1.Range(xlHoja1.Cells(5, 14), xlHoja1.Cells(5, 25)).Merge True
xlHoja1.Cells(6, 14) = "Cartera"
xlHoja1.Range(xlHoja1.Cells(6, 14), xlHoja1.Cells(6, 15)).Merge True
xlHoja1.Cells(7, 14) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 14), xlHoja1.Cells(250, 14)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 15) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 15), xlHoja1.Cells(250, 15)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 16) = "Calif. (0)"
xlHoja1.Range(xlHoja1.Cells(6, 16), xlHoja1.Cells(6, 17)).Merge True
xlHoja1.Cells(7, 16) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 16), xlHoja1.Cells(250, 16)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 17) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 17), xlHoja1.Cells(250, 17)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 18) = "Calif. (1)"
xlHoja1.Range(xlHoja1.Cells(6, 18), xlHoja1.Cells(6, 19)).Merge True
xlHoja1.Cells(7, 18) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 18), xlHoja1.Cells(250, 18)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 19) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 19), xlHoja1.Cells(250, 19)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 20) = "Calif. (2)"
xlHoja1.Range(xlHoja1.Cells(6, 20), xlHoja1.Cells(6, 20)).Merge True
xlHoja1.Cells(7, 20) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 20), xlHoja1.Cells(250, 20)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 21) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 21), xlHoja1.Cells(250, 21)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 22) = "Calif. (3)"
xlHoja1.Range(xlHoja1.Cells(6, 22), xlHoja1.Cells(6, 23)).Merge True
xlHoja1.Cells(7, 22) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 22), xlHoja1.Cells(250, 22)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 23) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 23), xlHoja1.Cells(250, 23)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 24) = "Calif. (4)"
xlHoja1.Range(xlHoja1.Cells(6, 24), xlHoja1.Cells(6, 25)).Merge True
xlHoja1.Cells(7, 24) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 24), xlHoja1.Cells(250, 24)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 25) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 25), xlHoja1.Cells(250, 25)).NumberFormat = "#,##0.00;#,##0.00"

xlHoja1.Range("A1:A150").ColumnWidth = 30
xlHoja1.Range("B1:B150").ColumnWidth = 10
xlHoja1.Range("C1:C150").ColumnWidth = 14
xlHoja1.Range("D1:D150").ColumnWidth = 8
xlHoja1.Range("E1:E150").ColumnWidth = 14
xlHoja1.Range("F1:F150").ColumnWidth = 8
xlHoja1.Range("G1:G150").ColumnWidth = 14
xlHoja1.Range("H1:H150").ColumnWidth = 8
xlHoja1.Range("I1:I150").ColumnWidth = 14
xlHoja1.Range("J1:J150").ColumnWidth = 8
xlHoja1.Range("K1:K150").ColumnWidth = 14
xlHoja1.Range("L1:L150").ColumnWidth = 8
xlHoja1.Range("M1:M150").ColumnWidth = 14
xlHoja1.Range("N1:N150").ColumnWidth = 8
xlHoja1.Range("O1:O150").ColumnWidth = 14
xlHoja1.Range("P1:P150").ColumnWidth = 8
xlHoja1.Range("Q1:Q150").ColumnWidth = 14
xlHoja1.Range("R1:R150").ColumnWidth = 8
xlHoja1.Range("S1:S150").ColumnWidth = 14
xlHoja1.Range("T1:T150").ColumnWidth = 8
xlHoja1.Range("U1:U150").ColumnWidth = 14
xlHoja1.Range("V1:V150").ColumnWidth = 8
xlHoja1.Range("W1:W150").ColumnWidth = 14
xlHoja1.Range("X1:X150").ColumnWidth = 8
xlHoja1.Range("Y1:Y150").ColumnWidth = 14



xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(7, 25)).HorizontalAlignment = xlCenter
xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(7, 13)).Interior.Color = &HC0E0FF
xlHoja1.Range(xlHoja1.Cells(5, 14), xlHoja1.Cells(7, 25)).Interior.Color = &HB0B0B0
xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(7, 25)).Font.Bold = True
xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(7, 25)).Cells.Borders.LineStyle = xlOutside
xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(7, 25)).Cells.Borders.LineStyle = xlInside


End Sub


Private Sub GeneraRepCartAltoRiesgo(ByVal pdFecha As Date)
Dim lnFila As Integer
Dim lsConec As String
Dim lsSql As String
Dim lrAge As New ADODB.Recordset
Dim lrReg As New ADODB.Recordset
Dim I As Integer, lnIIni As Integer
Dim lsCadCond As String, lsCadCondDesc As String
Dim lsCadTotal() As String, lnContAgencia As Integer
Dim lsCadTotales As String, J As Integer
    Dim oCon As DConecta
    Set oCon = New DConecta

    If gbBitCentral = True Then
        oCon.AbreConexion
    Else
        oCon.AbreConexion 'Remota "07", , , "03"
    End If
   
   lsCadCond = GetProdsMarcados
   lsCadCondDesc = GetProdsMarcadosDesc
   
   ExcelAddHoja "CartAltoRiesgo", xlLibro, xlHoja1
   Call GeneraRepCartAltoRiesgoCab(pdFecha, lsCadCondDesc)
  
    If gbBitCentral = False Then
        lsSql = " Select cCodTab, cValor, cNomTab From dbcomunes..Tablacod Where cCodTab like '47%'" _
               & " And cValor like '112%'  Order By cValor "
    Else
        lsSql = "select cAgeCod as cCodTab, cAgeCod as cValor, cAgeDescripcion as cNomTab from agencias Order by cAgeCod"
    End If
    
   lrAge.CursorLocation = adUseClient
   Set lrAge = oCon.CargaRecordSet(lsSql)
   Set lrAge.ActiveConnection = Nothing
   
   ReDim lsCadTotal(lrAge.RecordCount)
   lnContAgencia = 1
   I = 10
   Do While Not lrAge.EOF
        
        'Imprime Cabecera
        xlHoja1.Cells(I, 1) = Mid(lrAge!cValor, 4, 2) & " " & Trim(lrAge!cNomtab)
        xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 1)).Font.Bold = True
        I = I + 1
   
        '=======================================
        
        If gbBitCentral = True Then
            lsSql = " Select c.cCodAnalista, " _
                & " COUNT( CASE WHEN Substring(c.cCtaCod,9,1) = '1' THEN c.cCtaCod END )  As NumSoles, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCtaCod,9,1) = '1' THEN c.nSaldoCap END ), 0 )  As SKSoles, " _
                & " COUNT( CASE WHEN Substring(c.cCtaCod,9,1) = '1' And " _
                & "        ( ( Substring(c.cCtaCod,6,1) = '1' and c.nDiasAtraso > 15 and c.cRefinan = 'N') " _
                & "         or ( Substring(c.cCtaCod,6,1) in('2','3') and c.nDiasAtraso > 30 and c.cRefinan = 'N') " _
                & "         or (c.cRefinan ='R') or  (c.nPrdEstado ='" & gColocEstRecVigJud & "') ) THEN c.cCtaCod END )  As NumCARSol , " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCtaCod,9,1) = '1' And " _
                & "         ( ( Substring(c.cCtaCod,6,1) = '1' and c.nDiasAtraso > 15 and c.cRefinan = 'N') " _
                & "         or ( Substring(c.cCtaCod,6,1) in('2','3') and c.nDiasAtraso > 30 and c.cRefinan = 'N') " _
                & "         or (c.cRefinan ='R') or  (c.nPrdEstado ='" & gColocEstRecVigJud & "') ) THEN c.nSaldoCap END ), 0 )  As SKCARSol , " _
                & " COUNT( CASE WHEN Substring(c.cCtaCod,9,1) = '2' THEN c.cCtaCod END )  As NumDolar, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCtaCod,9,1) = '2' THEN c.nSaldoCap END ), 0 )  As SKDolar,  " _
                & " COUNT( CASE WHEN Substring(c.cCtaCod,9,1) = '2' And " _
                & "         ( ( Substring(c.cCtaCod,6,1) = '2' and c.nDiasAtraso > 15 and c.cRefinan = 'N') " _
                & "         or ( Substring(c.cCtaCod,6,1) in('2','3') and c.nDiasAtraso > 30 and c.cRefinan = 'N') " _
                & "         or (c.cRefinan ='R') or  (c.nPrdEstado ='" & gColocEstRecVigJud & "') ) THEN c.cCtaCod END )  As NumCARDol , " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCtaCod,9,1) = '2' And " _
                & "         ( ( Substring(c.cCtaCod,6,1) = '1' and c.nDiasAtraso > 15 and c.cRefinan  = 'N') " _
                & "         or ( Substring(c.cCtaCod,6,1) in('2','3') and c.nDiasAtraso > 30 and c.cRefinan = 'N') " _
                & "         or (c.cRefinan ='R') or  (c.nPrdEstado ='" & gColocEstRecVigJud & "')) THEN c.nSaldoCap END ), 0 )  As SKCARDol " _
                & " from " & sservidorconsolidada & "creditoConsol c " _
                & " where (( c.nPrdEstado in (" & cVigente & ", " & cPigno & ")) or ( c.nPrdEstado in ('" & gColocEstRecVigJud & "'))) and c.nSaldoCap > 0 " _
                & " And Substring(c.cCtaCod,4,2) = '" & Trim(lrAge!cValor) & "' " & lsCadCond _
                & " Group by c.cCodAnalista  Order by c.cCodAnalista "
        Else
        
            lsSql = " Select c.cCodAnalista, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '1' THEN c.cCodCta END )  As NumSoles, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '1' THEN c.nSaldoCap END ), 0 )  As SKSoles, " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '1' And " _
                & "        ( ( Substring(c.cCodcta,3,1) = '1' and c.nDiasAtraso > 15 and c.cRefinan = 'N') " _
                & "         or ( Substring(c.cCodcta,3,1) in('2','3') and c.nDiasAtraso > 30 and c.cRefinan = 'N') " _
                & "         or (c.cRefinan ='R') or  (c.cEstado = 'V') ) THEN c.cCodCta END )  As NumCARSol , " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '1' And " _
                & "         ( ( Substring(c.cCodcta,3,1) = '1' and c.nDiasAtraso > 15 and c.cRefinan = 'N') " _
                & "         or ( Substring(c.cCodcta,3,1) in('2','3') and c.nDiasAtraso > 30 and c.cRefinan = 'N') " _
                & "         or (c.cRefinan ='R') or  (c.cEstado = 'V') ) THEN c.nSaldoCap END ), 0 )  As SKCARSol , " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '2' THEN c.cCodCta END )  As NumDolar, " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '2' THEN c.nSaldoCap END ), 0 )  As SKDolar,  " _
                & " COUNT( CASE WHEN Substring(c.cCodCta,6,1) = '2' And " _
                & "         ( ( Substring(c.cCodcta,3,1) = '2' and c.nDiasAtraso > 15 and c.cRefinan = 'N') " _
                & "         or ( Substring(c.cCodcta,3,1) in('2','3') and c.nDiasAtraso > 30 and c.cRefinan = 'N') " _
                & "         or (c.cRefinan ='R') or  (c.cEstado = 'V') ) THEN c.cCodCta END )  As NumCARDol , " _
                & " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) = '2' And " _
                & "         ( ( Substring(c.cCodcta,3,1) = '1' and c.nDiasAtraso > 15 and c.cRefinan  = 'N') " _
                & "         or ( Substring(c.cCodcta,3,1) in('2','3') and c.nDiasAtraso > 30 and c.cRefinan = 'N') " _
                & "         or (c.cRefinan ='R') or  (c.cEstado = 'V') ) THEN c.nSaldoCap END ), 0 )  As SKCARDol " _
                & " from creditoConsol c " _
                & " where (( c.cEstado in ('F','1','4','6','7')) or ( c.cEstado in ('V') and c.cCondCre ='J' ))  and c.nSaldoCap > 0 " _
                & " And Substring(c.cCodCta,1,2) = '" & Mid(lrAge!cValor, 4, 2) & "' " & lsCadCond _
                & " Group by c.cCodAnalista  Order by c.cCodAnalista "
        End If
        
        lrReg.CursorLocation = adUseClient
        Set lrReg = oCon.CargaRecordSet(lsSql)
        Set lrReg.ActiveConnection = Nothing
        
        lnIIni = I
        Do While Not lrReg.EOF
        
            xlHoja1.Cells(I, 1) = lrReg!cCodAnalista
            xlHoja1.Cells(I, 2) = lrReg!NumSoles
            xlHoja1.Cells(I, 3) = lrReg!skSoles
            xlHoja1.Cells(I, 4) = lrReg!NumCARSol
            xlHoja1.Cells(I, 5) = lrReg!skCARSol
            If lrReg!skSoles > 0 Then
                xlHoja1.Cells(I, 6) = 100 * lrReg!skCARSol / lrReg!skSoles
            Else
                xlHoja1.Cells(I, 6) = 0
            End If
            
            xlHoja1.Cells(I, 7) = lrReg!NumDolar
            xlHoja1.Cells(I, 8) = lrReg!skDolar
            xlHoja1.Cells(I, 9) = lrReg!NumCARDol
            xlHoja1.Cells(I, 10) = lrReg!skCARDol
            If lrReg!skDolar > 0 Then
                xlHoja1.Cells(I, 11) = 100 * lrReg!skCARDol / lrReg!skDolar
            Else
                xlHoja1.Cells(I, 11) = 0
            End If
            
            I = I + 1
            lrReg.MoveNext
        Loop
        lrReg.Close
        If I <> lnIIni Then
             '***** Suma Totales Agencia
            xlHoja1.Range("A" & I & ":A" & I) = "Total Agencia"
            xlHoja1.Range("B" & I & ":B" & I).Formula = "=SUM(B" & lnIIni & ":B" & I - 1 & ")"
            xlHoja1.Range("C" & I & ":C" & I).Formula = "=SUM(C" & lnIIni & ":C" & I - 1 & ")"
            xlHoja1.Range("D" & I & ":D" & I).Formula = "=SUM(D" & lnIIni & ":D" & I - 1 & ")"
            xlHoja1.Range("E" & I & ":E" & I).Formula = "=SUM(E" & lnIIni & ":E" & I - 1 & ")"
            xlHoja1.Range("F" & I & ":F" & I).Formula = "=SUM(F" & lnIIni & ":F" & I - 1 & ")"
            xlHoja1.Range("G" & I & ":G" & I).Formula = "=SUM(G" & lnIIni & ":G" & I - 1 & ")"
            xlHoja1.Range("H" & I & ":H" & I).Formula = "=SUM(H" & lnIIni & ":H" & I - 1 & ")"
            xlHoja1.Range("I" & I & ":I" & I).Formula = "=SUM(I" & lnIIni & ":I" & I - 1 & ")"
            xlHoja1.Range("J" & I & ":J" & I).Formula = "=SUM(J" & lnIIni & ":J" & I - 1 & ")"
            xlHoja1.Range("K" & I & ":K" & I).Formula = "=SUM(K" & lnIIni & ":K" & I - 1 & ")"
            
        End If
        lsCadTotal(lnContAgencia) = I
        lnContAgencia = lnContAgencia + 1
        I = I + 1
        
        lrAge.MoveNext
    
   Loop
   
   I = I + 1
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 25)).Font.Bold = True
   xlHoja1.Range("A" & I & ":A" & I) = " TOTAL"
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "B" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("B" & I & ":B" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "C" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("C" & I & ":C" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "D" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("D" & I & ":D" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "E" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("E" & I & ":E" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "F" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("F" & I & ":F" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "G" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("G" & I & ":G" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "H" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("H" & I & ":H" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "I" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("I" & I & ":I" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "J" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("J" & I & ":J" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   lsCadTotales = "="
   For J = 1 To lrAge.RecordCount
        lsCadTotales = lsCadTotales + "K" & lsCadTotal(J) & "+"
   Next
   xlHoja1.Range("K" & I & ":K" & I).Formula = Mid(lsCadTotales, 1, Len(lsCadTotales) - 1)
   
   xlHoja1.Range(xlHoja1.Cells(8, 1), xlHoja1.Cells(I + 2, 11)).Borders(xlInsideVertical).LineStyle = xlContinuous
   xlHoja1.Range("A8:K" & I + 2).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
   
   oCon.CierraConexion
   
   MsgBox "Reporte Generado Satisfactoriamente"
End Sub

Private Sub GeneraRepCartAltoRiesgoCab(ByVal pdFecha As Date, ByVal psCondiciones As String)
xlHoja1.PageSetup.Orientation = xlPortrait
xlHoja1.PageSetup.CenterHorizontally = True
xlHoja1.PageSetup.Zoom = 60
xlHoja1.Cells(1, 1) = gsNomCmac
xlHoja1.Cells(2, 1) = "CARTERA DE ALTO RIESGO POR ANALISTA"
xlHoja1.Cells(3, 1) = "Al " & Format(pdFecha, "dd/mm/yyyy")

xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 11)).Font.Bold = True
xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 11)).Merge True
xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 11)).HorizontalAlignment = xlCenter

xlHoja1.Cells(4, 1) = psCondiciones

xlHoja1.Cells(6, 1) = "Analista"
xlHoja1.Cells(7, 1) = "Creditos"
xlHoja1.Cells(5, 2) = "Moneda Nacional"
xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(5, 6)).Merge True
xlHoja1.Cells(6, 2) = "Total Cartera"
xlHoja1.Range(xlHoja1.Cells(6, 2), xlHoja1.Cells(6, 3)).Merge True
xlHoja1.Cells(7, 2) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(150, 2)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 3) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 3), xlHoja1.Cells(150, 3)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 4) = "Cartera Alto Riesgo"
xlHoja1.Range(xlHoja1.Cells(6, 4), xlHoja1.Cells(6, 5)).Merge True
xlHoja1.Cells(7, 4) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 4), xlHoja1.Cells(150, 4)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 5) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 5), xlHoja1.Cells(150, 5)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 6) = "C.A.R."
xlHoja1.Cells(7, 6) = " % "
xlHoja1.Range(xlHoja1.Cells(7, 6), xlHoja1.Cells(150, 6)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(5, 7) = "Moneda Extranjera"
xlHoja1.Range(xlHoja1.Cells(5, 7), xlHoja1.Cells(5, 11)).Merge True
xlHoja1.Cells(6, 7) = "Total Cartera"
xlHoja1.Range(xlHoja1.Cells(6, 7), xlHoja1.Cells(6, 8)).Merge True
xlHoja1.Cells(7, 7) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 7), xlHoja1.Cells(150, 7)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 8) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 8), xlHoja1.Cells(150, 8)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 9) = "Cartera Alto Riesgo"
xlHoja1.Range(xlHoja1.Cells(6, 9), xlHoja1.Cells(6, 10)).Merge True
xlHoja1.Cells(7, 9) = "Nro"
xlHoja1.Range(xlHoja1.Cells(7, 9), xlHoja1.Cells(150, 9)).NumberFormat = "#,#0;#,#0"
xlHoja1.Cells(7, 10) = "S.K."
xlHoja1.Range(xlHoja1.Cells(7, 10), xlHoja1.Cells(150, 10)).NumberFormat = "#,##0.00;#,##0.00"
xlHoja1.Cells(6, 11) = "C.A.R."
xlHoja1.Cells(7, 11) = " % "
xlHoja1.Range(xlHoja1.Cells(7, 11), xlHoja1.Cells(150, 11)).NumberFormat = "#,##0.00;#,##0.00"

xlHoja1.Range("A1:A150").ColumnWidth = 30
xlHoja1.Range("B1:B150").ColumnWidth = 10
xlHoja1.Range("C1:C150").ColumnWidth = 14
xlHoja1.Range("D1:D150").ColumnWidth = 8
xlHoja1.Range("E1:E150").ColumnWidth = 14
xlHoja1.Range("F1:F150").ColumnWidth = 10
xlHoja1.Range("G1:G150").ColumnWidth = 10
xlHoja1.Range("H1:H150").ColumnWidth = 14
xlHoja1.Range("I1:I150").ColumnWidth = 8
xlHoja1.Range("J1:J150").ColumnWidth = 14
xlHoja1.Range("K1:K150").ColumnWidth = 10

xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(7, 12)).HorizontalAlignment = xlCenter
xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(7, 6)).Interior.Color = &HC0E0FF
xlHoja1.Range(xlHoja1.Cells(5, 7), xlHoja1.Cells(7, 11)).Interior.Color = &HB0B0B0
xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(7, 11)).Font.Bold = True
xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(7, 11)).Cells.Borders.LineStyle = xlOutside
xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(7, 11)).Cells.Borders.LineStyle = xlInside

End Sub

Private Sub GeneraRepEstratosDepositos(ByVal pdFecha As Date)
    Dim lsConec As String
    Dim lsSql As String
    Dim lrReg As New ADODB.Recordset
    Dim nTemp As Integer
    Dim I As Integer, lnIIni As Integer
    Dim lnMonedas As Integer
    Dim oCon As DConecta
    
    Dim regTemp As New ADODB.Recordset
    
    Set oCon = New DConecta

    If gbBitCentral = True Then
        oCon.AbreConexion
    Else
        oCon.AbreConexion 'Remota "07", , , "03"
    End If
   
   
'*********  Creditos
   ExcelAddHoja "TasasIntCred", xlLibro, xlHoja1
   xlHoja1.PageSetup.Orientation = xlPortrait
   xlHoja1.PageSetup.CenterHorizontally = True
   xlHoja1.PageSetup.Zoom = 55
   xlHoja1.Cells(1, 1) = gsNomCmac
   xlHoja1.Cells(2, 1) = "CREDITOS POR TASAS DE INTERES"
   xlHoja1.Cells(3, 1) = "AL " & Format(pdFecha, "dd/mm/yyyy")
   
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 4)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 4)).Merge True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 4)).HorizontalAlignment = xlCenter

   xlHoja1.Range("A1:A150").ColumnWidth = 30
   xlHoja1.Range("B1:B150").ColumnWidth = 10
   xlHoja1.Range("C1:C150").ColumnWidth = 10
   xlHoja1.Range("D1:D150").ColumnWidth = 15
   xlHoja1.Range(xlHoja1.Cells(7, 2), xlHoja1.Cells(150, 2)).NumberFormat = "#,##0.00;#,##0.00"
   xlHoja1.Range(xlHoja1.Cells(7, 3), xlHoja1.Cells(150, 3)).NumberFormat = "#,#0;#,#0"
   xlHoja1.Range(xlHoja1.Cells(7, 4), xlHoja1.Cells(150, 4)).NumberFormat = "#,##0.00;#,##0.00"
   
   I = 5
   
   For lnMonedas = 1 To 2
      I = I + 1
      xlHoja1.Cells(I, 1) = IIf(lnMonedas = "1", "SOLES", "DOLARES")
      xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 1)).Interior.Color = &HC0E0FF
      xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 1)).Font.Bold = True
      I = I + 1
      xlHoja1.Cells(I, 1) = "Producto"
      xlHoja1.Cells(I, 2) = "Tasa Int"
      xlHoja1.Cells(I, 3) = "Nro"
      xlHoja1.Cells(I, 4) = "Saldo"
      xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 4)).HorizontalAlignment = xlCenter
      xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 4)).Cells.Borders.LineStyle = xlOutside
      xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 4)).Cells.Borders.LineStyle = xlInside
      I = I + 1
      xlHoja1.Cells(I, 1) = "Tipo de"
      xlHoja1.Cells(I + 1, 1) = "Credito"
      xlHoja1.Cells(I, 2) = " Tasa"
      xlHoja1.Cells(I + 1, 2) = "Interes"
      xlHoja1.Cells(I, 3) = "Numero"
      xlHoja1.Cells(I, 4) = "Saldo"
   
        '=======================================
        
        If gbBitCentral = True Then
            lsSql = " Select Substring(cctacod,9,1) Moneda, Substring(cctacod,6,3) Tipo, " _
                & " nTasaint Tasa, count(cctacod) Nro , sum(nsaldocap) Saldo " _
                & " From " & sservidorconsolidada & "creditoconsol Where nPrdestado IN (" & cVigente & ") And Substring(cctacod,9,1) ='" & lnMonedas & "' " _
                & " Group by Substring(cctacod,9,1), Substring(cctacod,6,3), nTasaint " _
                & " Order by Moneda, Tipo, Tasa "
        Else
            lsSql = " Select Substring(ccodcta,6,1) Moneda, Substring(ccodcta,3,3) Tipo, " _
                & " nTasaint Tasa, count(ccodcta) Nro , sum(nsaldocap) Saldo " _
                & " From creditoconsol Where cestado = 'F' And Substring(cCodcta,6,1) ='" & lnMonedas & "' " _
                & " Group by Substring(ccodcta,6,1), Substring(ccodcta,3,3), nTasaint " _
                & " Order by Moneda, Tipo, Tasa "
        End If
        
        lrReg.CursorLocation = adUseClient
        Set lrReg = oCon.CargaRecordSet(lsSql)
        Set lrReg.ActiveConnection = Nothing
        
        lnIIni = I
        
        If gbBitCentral = True Then
            lsSql = "select cProdCod as cValor, cProdDesc as cProducto "
            lsSql = lsSql & " from  RepAgruProdDet where nAgruCod=1 "
            lsSql = lsSql & " order by cvalor"
        Else
            lsSql = " Select cGrupo=nConsCod, cValor=nconsvalor, cProducto= cconsdescripcion "
            lsSql = lsSql & " From constante C where C.nConsCod='1001' and nconsvalor not in(" & Producto.gCapAhorros & ", " & Producto.gCapPlazoFijo & ", " & Producto.gCapCTS & ", " & Producto.gColConsuPrendario & ") "
            lsSql = lsSql & " and (case when nconsvalor in(select min(nconsvalor) from constante K where K.nconscod=C.nConscod AND substring(convert(varchar(3), K.nconsvalor),1,1) = substring(convert(varchar(3), C.nconsvalor),1,1)) Then 1 Else 2 End)= 1"
            lsSql = lsSql & " Order by nconsvalor "
        End If
                    
        If gbBitCentral = False Then
            oCon.AbreConexion
        End If
        
        Set regTemp = oCon.CargaRecordSet(lsSql)
        
        If gbBitCentral = False Then
            oCon.AbreConexion 'Remota "07", , , "03"
        End If
        
        Do While Not lrReg.EOF
          
            regTemp.MoveFirst
            Do While Not regTemp.EOF
               If lrReg!Tipo = regTemp!cValor Then
                   xlHoja1.Cells(I, 1) = regTemp!cProducto
                   Exit Do
               End If
               regTemp.MoveNext
            Loop
                 
            
            xlHoja1.Cells(I, 2) = lrReg!TASA
            xlHoja1.Cells(I, 3) = lrReg!nro
            xlHoja1.Cells(I, 4) = lrReg!Saldo
            
            I = I + 1
            lrReg.MoveNext
        Loop
        lrReg.Close
        regTemp.Close
        xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 4)).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 4)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
        
   Next
   
'*********  Depositos a Plazos
   ExcelAddHoja "DepositPlazo", xlLibro, xlHoja1
   xlHoja1.PageSetup.Orientation = xlPortrait
   xlHoja1.PageSetup.CenterHorizontally = True
   xlHoja1.PageSetup.Zoom = 55
   xlHoja1.Cells(1, 1) = gsNomCmac
   xlHoja1.Cells(2, 1) = "DEPOSITOS A PLAZOS "
   xlHoja1.Cells(3, 1) = "AL " & Format(pdFecha, "dd/mm/yyyy")
   
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 5)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 5)).Merge True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 5)).HorizontalAlignment = xlCenter

   xlHoja1.Range("A1:A150").ColumnWidth = 30
   xlHoja1.Range("B1:B150").ColumnWidth = 10
   xlHoja1.Range("C1:C150").ColumnWidth = 14
   xlHoja1.Range("D1:D150").ColumnWidth = 10
   xlHoja1.Range("E1:E150").ColumnWidth = 14
   
   xlHoja1.Range(xlHoja1.Cells(5, 2), xlHoja1.Cells(150, 2)).NumberFormat = "#,#0;#,#0"
   xlHoja1.Range(xlHoja1.Cells(5, 3), xlHoja1.Cells(150, 3)).NumberFormat = "#,##0.00;#,##0.00"
   xlHoja1.Range(xlHoja1.Cells(5, 4), xlHoja1.Cells(150, 4)).NumberFormat = "#,#0;#,#0"
   xlHoja1.Range(xlHoja1.Cells(5, 5), xlHoja1.Cells(150, 5)).NumberFormat = "#,##0.00;#,##0.00"
   
   xlHoja1.Cells(4, 1) = "Ahorros - Plazos Fijos"
   
   I = 5
   
   xlHoja1.Cells(I, 1) = "Rango"
   xlHoja1.Cells(I, 2) = "Soles"
   xlHoja1.Range(xlHoja1.Cells(I, 2), xlHoja1.Cells(I, 3)).Merge True
   xlHoja1.Cells(I + 1, 2) = "Nro"
   xlHoja1.Cells(I + 1, 3) = "Saldo"
   xlHoja1.Cells(I, 4) = "Dolares"
   xlHoja1.Range(xlHoja1.Cells(I, 4), xlHoja1.Cells(I, 5)).Merge True
   xlHoja1.Cells(I + 1, 4) = "Nro"
   xlHoja1.Cells(I + 1, 5) = "Saldo"
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 1, 5)).HorizontalAlignment = xlCenter
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 1, 5)).Cells.Borders.LineStyle = xlOutside
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 1, 5)).Cells.Borders.LineStyle = xlInside

   I = I + 2
    '=======================================
    If gbBitCentral = True Then
        lsSql = " Select Moneda, Rango, Count(Cuenta) Num, SUM(Saldo) Monto FROM (" _
            & " Select Substring(cCtaCod,9,1) Moneda, Rango = CASE " _
            & " WHEN nPlazo <= 30 THEN '01. Menor a 30' " _
            & " WHEN nPlazo > 30 And nPlazo <= 89 THEN '02. 31 a 89' " _
            & " WHEN nPlazo > 89 And nPlazo <= 179 THEN '03. 90 a 179' " _
            & " WHEN nPlazo > 179 And nPlazo <= 359 THEN '04. 180 a 359' " _
            & " WHEN nPlazo > 359 And nPlazo <= 539 THEN '05. 360 a 539' " _
            & " WHEN nPlazo > 539 And nPlazo <= 719 THEN '06. 540 a 719' " _
            & " WHEN nPlazo > 719 THEN '07. Mayor a 720' END, " _
            & " cCtaCod Cuenta, nSaldCntPF Saldo from " & sservidorconsolidada & "PlazoFijoConsol where " _
            & " nEstCtaPF not in (1400,1300) " _
            & " Union " _
            & " Select Substring(cCtaCod,6,1) Moneda, Rango = '07. Mayor a 720', " _
            & " cCtaCod Cuenta, nSaldCntCTS Saldo from " & sservidorconsolidada & "CTSConsol where " _
            & " nEstCtaCTS not in (1400,1300) " _
            & " ) T Group by Moneda, Rango Order by Rango, Moneda "
    
    Else
        lsSql = " Select Moneda, Rango, Count(Cuenta) Num, SUM(Saldo) Monto FROM (" _
            & " Select Substring(cCodCta,6,1) Moneda, Rango = CASE " _
            & " WHEN nPlazo <= 30 THEN '01. Menor a 30' " _
            & " WHEN nPlazo > 30 And nPlazo <= 89 THEN '02. 31 a 89' " _
            & " WHEN nPlazo > 89 And nPlazo <= 179 THEN '03. 90 a 179' " _
            & " WHEN nPlazo > 179 And nPlazo <= 359 THEN '04. 180 a 359' " _
            & " WHEN nPlazo > 359 And nPlazo <= 539 THEN '05. 360 a 539' " _
            & " WHEN nPlazo > 539 And nPlazo <= 719 THEN '06. 540 a 719' " _
            & " WHEN nPlazo > 719 THEN '07. Mayor a 720' END, " _
            & " cCodCta Cuenta, nSaldCntPF Saldo from PlazoFijoConsol where " _
            & " cEstCtaPF not in ('C','U') " _
            & " Union " _
            & " Select Substring(cCodCta,6,1) Moneda, Rango = '07. Mayor a 720', " _
            & " cCodCta Cuenta, nSaldCntCTS Saldo from CTSConsol where " _
            & " cEstCtaCTS not in ('C','U') " _
            & " ) T Group by Moneda, Rango Order by Rango, Moneda "
    End If
    
    lrReg.CursorLocation = adUseClient
    Set lrReg = oCon.CargaRecordSet(lsSql)
    Set lrReg.ActiveConnection = Nothing
    
    lnIIni = I
    Do While Not lrReg.EOF
        xlHoja1.Cells(I, 1) = lrReg!Rango
        xlHoja1.Cells(I, 2) = lrReg!Num
        xlHoja1.Cells(I, 3) = lrReg!Monto
        lrReg.MoveNext
        If lrReg.EOF Then
            xlHoja1.Cells(I, 4) = 0
            xlHoja1.Cells(I, 5) = 0
        Else
            xlHoja1.Cells(I, 4) = lrReg!Num
            xlHoja1.Cells(I, 5) = lrReg!Monto
            lrReg.MoveNext
        End If
        I = I + 1
    Loop
    lrReg.Close
    
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 5)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic

    
   I = I + 2
   xlHoja1.Cells(I, 1) = "Ahorros - Ordenes de Pago"
   I = I + 1
   xlHoja1.Cells(I, 2) = "Soles"
   xlHoja1.Range(xlHoja1.Cells(I, 2), xlHoja1.Cells(I, 3)).Merge True
   xlHoja1.Cells(I + 1, 2) = "Nro"
   xlHoja1.Cells(I + 1, 3) = "Saldo"
   xlHoja1.Cells(I, 4) = "Dolares"
   xlHoja1.Range(xlHoja1.Cells(I, 4), xlHoja1.Cells(I, 5)).Merge True
   xlHoja1.Cells(I + 1, 4) = "Nro"
   xlHoja1.Cells(I + 1, 5) = "Saldo"
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 1, 5)).HorizontalAlignment = xlCenter
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 1, 5)).Cells.Borders.LineStyle = xlOutside
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 1, 5)).Cells.Borders.LineStyle = xlInside
   I = I + 2
    '=======================================
    If gbBitCentral = True Then
        lsSql = " SELECT nOrdPag, Substring(cCtaCod,9,1) moneda, Count(*) Num, Sum(nSaldCntAC) Monto " _
              & " From " & sservidorconsolidada & "AhorrocConsol where nEstCtaAC not in (1400,1300) " _
              & " Group by nOrdPag, substring(cCtaCod,9,1) " _
              & " Order by nOrdPag, substring(cCtaCod,9,1) "
    Else
        lsSql = " SELECT cOrdPag, Substring(cCodCta,6,1) moneda, Count(*) Num, Sum(nSaldCntAC) Monto " _
              & " From AhorrocConsol where cEstCtaAC not in ('C','U') " _
              & " Group by cOrdPag, substring(ccodcta,6,1) " _
              & " Order by cOrdPag, substring(ccodcta,6,1) "
    End If
    
    lrReg.CursorLocation = adUseClient
    Set lrReg = oCon.CargaRecordSet(lsSql)
    Set lrReg.ActiveConnection = Nothing
    
    lnIIni = I
    Do While Not lrReg.EOF
        
        If gbBitCentral = True Then
            xlHoja1.Cells(I, 1) = IIf(lrReg!nOrdPag = 0, "Sin Ordenes", "Con Ordenes")
        Else
            xlHoja1.Cells(I, 1) = IIf(lrReg!cOrdPag = "N", "Sin Ordenes", "Con Ordenes")
        End If
        
        xlHoja1.Cells(I, 2) = lrReg!Num
        xlHoja1.Cells(I, 3) = lrReg!Monto
        lrReg.MoveNext
        xlHoja1.Cells(I, 4) = lrReg!Num
        xlHoja1.Cells(I, 5) = lrReg!Monto
        lrReg.MoveNext
        I = I + 1
    Loop
    lrReg.Close
   
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 5)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 5)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
   
   oCon.CierraConexion
   
   MsgBox "Reporte Generado Satisfactoriamente"
End Sub

Private Sub GeneraRepCapMejoresClientes(ByVal pnNumero As Long, ByVal pnTipCambio As Double, ByVal pdFecha As Date, ByVal psAgencia As String)
    Dim lsConec As String
    Dim lsSql As String
    Dim lrReg As New ADODB.Recordset
    Dim I As Integer, lnIIni As Integer
    Dim lnContador As Long
    Dim lsQuery As String
    
    Dim oCon As DConecta
    Set oCon = New DConecta

    If gbBitCentral = True Then
        oCon.AbreConexion
    Else
        oCon.AbreConexion 'Remota "07", , , "03"
    End If
    
   ExcelAddHoja "PrinClientesCap", xlLibro, xlHoja1
   xlHoja1.PageSetup.Orientation = xlPortrait
   xlHoja1.PageSetup.CenterHorizontally = True
   xlHoja1.PageSetup.Zoom = 55
   xlHoja1.Cells(1, 1) = gsNomCmac
   xlHoja1.Cells(2, 1) = "PRINCIPALES CLIENTES DE CAPTACIONES"
   xlHoja1.Cells(3, 1) = "AL " & Format(pdFecha, "dd/mm/yyyy")
   
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 8)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 8)).Merge True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 8)).HorizontalAlignment = xlCenter

   xlHoja1.Range("A1:A150").ColumnWidth = 5
   xlHoja1.Range("B1:B150").ColumnWidth = 10
   xlHoja1.Range("C1:C150").ColumnWidth = 40
   xlHoja1.Range("D1:D150").ColumnWidth = 30
   xlHoja1.Range("E1:E150").ColumnWidth = 12
   xlHoja1.Range("F1:F150").ColumnWidth = 12
   xlHoja1.Range("G1:G150").ColumnWidth = 12
   xlHoja1.Range("H1:H150").ColumnWidth = 15
   xlHoja1.Range("I1:I150").ColumnWidth = 30
   xlHoja1.Range("J1:J150").ColumnWidth = 25
   
   xlHoja1.Range(xlHoja1.Cells(6, 5), xlHoja1.Cells(150, 5)).NumberFormat = "#,##0.00;#,##0.00"
   xlHoja1.Range(xlHoja1.Cells(6, 7), xlHoja1.Cells(150, 7)).NumberFormat = "dd/mm/yyyy"
   
   I = 5
    
   xlHoja1.Cells(I, 1) = "Nro"
   xlHoja1.Cells(I, 2) = "Codigo"
   xlHoja1.Cells(I, 3) = "Cliente"
   xlHoja1.Cells(I, 4) = "Direccion"
   xlHoja1.Cells(I, 5) = "Saldo"
   xlHoja1.Cells(I, 6) = "Telefono"
   xlHoja1.Cells(I, 7) = "Fec.Nac."
   xlHoja1.Cells(I, 8) = "Zona"
   xlHoja1.Cells(I, 9) = "Personeria"
   xlHoja1.Cells(I, 10) = "Codigo Personeria"
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 10)).HorizontalAlignment = xlCenter
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 10)).Cells.Borders.LineStyle = xlOutside
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 10)).Cells.Borders.LineStyle = xlInside
   I = I + 1
    
    'JEOM
    lsQuery = ""
    If psAgencia = "Todos" Then
       lsQuery = ""
    Else
       lsQuery = " and ag.cAgecod ='" & psAgencia & "'"
    End If
    
    
    'By Capi Setiembre 2007 Req PRO0709065636 (Planeamiento)
    If chkuntitular = 1 Then
    If gbBitCentral = True Then
        lsSql = "  Select TOP " & pnNumero & " " _
            & " TA.cpersCod, P.cPersNombre, TA.nSaldo, P.cPersDireccDomicilio, P.cPersTelefono, " _
            & " P.dPersNacCreac, ISNULL(Z.cDesZon,'') Zona,c.cConsDescripcion ,p.nPersPersoneria Personeria    " _
            & " FROM ( Select T.cpersCod, SUM(T.nSaldo) nSaldo FROM ( " _
            & " Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) " _
            & "  WHEN '1' THEN A.nSaldCntAC WHEN '2' THEN A.nSaldCntAC *" & pnTipCambio & "  END " _
            & "  FROM " & sservidorconsolidada & " AhorroCConsol A INNER JOIN (Select cCtaCod,Min(cPersCod) cPersCod,Min(nPrdPersRelac) nPrdPersRelac FROM " & sservidorconsolidada & " ProductoPersonaConsol Where nPrdPersRelac=" & gCapRelPersTitular & " Group By cCtaCod) PC ON A.cctacod = PC.cctacod " _
            & "  INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) WHERE A.nEstCtaAC " _
            & "     NOT IN (1400,1300) AND PC.nPrdPersRelac = " & gCapRelPersTitular & lsQuery & " Union " _
            & " Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) " _
            & "        WHEN '1' THEN A.nSaldCntPF WHEN '2' THEN A.nSaldCntPF *" & pnTipCambio & "  END " _
            & " FROM " & sservidorconsolidada & "PlazoFijoConsol A Inner Join (Select cCtaCod,Min(cPersCod) cPersCod,Min(nPrdPersRelac) nPrdPersRelac FROM " & sservidorconsolidada & " ProductoPersonaConsol Where nPrdPersRelac=" & gCapRelPersTitular & " Group By cCtaCod) PC ON A.cctacod = PC.cctacod " _
            & " INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) Where " _
            & "    A.nEstCtaPF not in (1400,1300) and PC.nPrdPersRelac = " & gCapRelPersTitular & lsQuery & " Union " _
            & " Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) " _
            & "       WHEN '1' THEN A.nSaldCntCTS WHEN '2' THEN A.nSaldCntCTS *" & pnTipCambio & " END " _
            & "       FROM " & sservidorconsolidada & "CTSConsol A Inner Join (Select cCtaCod,Min(cPersCod) cPersCod,Min(nPrdPersRelac) nPrdPersRelac FROM " & sservidorconsolidada & " ProductoPersonaConsol Where nPrdPersRelac=" & gCapRelPersTitular & " Group By cCtaCod) PC ON A.cctacod = PC.cctacod " _
            & " INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) Where " _
            & " A.nEstCtaCTS not in (1400,1300) and PC.nPrdPersRelac = " & gCapRelPersTitular _
            & lsQuery _
            & " ) T GROUP BY T.cpersCod " _
            & " ) TA INNER JOIN Persona P  INNER JOIN Constante c ON c.nConsValor =p.nPersPersoneria and nConsCod='1002' LEFT JOIN " & sservidorconsolidada & "Zonas Z ON P.cPersDireccUbiGeo = Z.cCodZon " _
            & " ON TA.cpersCod = P.cpersCod ORDER BY TA.nSaldo DESC "
        
    Else
        'Cuando se trabaja con data descentralizada no aplicable a la institucion
        lsSql = "  Select TOP " & pnNumero & " " _
            & " TA.cCodPers, P.cNomPers, TA.nSaldo, P.cDirPers, P.cTelPers, " _
            & " P.dFecNac, ISNULL(Z.cDesZon,'') Zona " _
            & " FROM ( Select T.cCodPers, SUM(T.nSaldo) nSaldo FROM ( " _
            & " Select PC.cCodPers, A.cCodCta, nSaldo = CASE SUBSTRING(A.cCodCta,6,1) " _
            & "     WHEN '1' THEN A.nSaldCntAC WHEN '2' THEN A.nSaldCntAC *" & pnTipCambio & "  END " _
            & "     FROM AhorroCConsol A INNER JOIN PersCuentaConsol PC ON A.cCodCta = PC.cCodCta WHERE A.cEstCtaAC " _
            & "     NOT IN ('C','U') AND PC.cRelaCta = 'TI' " _
            & " Union " _
            & " Select PC.cCodPers, A.cCodCta, nSaldo = CASE SUBSTRING(A.cCodCta,6,1) " _
            & "        WHEN '1' THEN A.nSaldCntPF WHEN '2' THEN A.nSaldCntPF *" & pnTipCambio & "  END " _
            & " FROM PlazoFijoConsol A Inner Join PersCuentaConsol PC on A.cCodCta = PC.cCodCta Where " _
            & "    A.cEstCtaPF not in ('C','U') and PC.cRelaCta = 'TI' " _
            & " Union " _
            & " Select PC.cCodPers, A.cCodCta, nSaldo = CASE SUBSTRING(A.cCodCta,6,1) " _
            & "       WHEN '1' THEN A.nSaldCntCTS WHEN '2' THEN A.nSaldCntCTS *" & pnTipCambio & " END " _
            & "       FROM CTSConsol A Inner Join PersCuentaConsol PC on A.cCodCta = PC.cCodCta Where " _
            & " A.cEstCtaCTS not in ('C','U') and PC.cRelaCta = 'TI' " _
            & " ) T GROUP BY T.cCodPers " _
            & " ) TA INNER JOIN dbPersona..Persona P LEFT JOIN Zonas Z ON P.cCodZon = Z.cCodZon " _
            & " ON TA.cCodPers = P.cCodPers ORDER BY TA.nSaldo DESC "
    End If
    'End By
    Else
    
    
    If gbBitCentral = True Then
        lsSql = "  Select TOP " & pnNumero & " " _
            & " TA.cpersCod, P.cPersNombre, TA.nSaldo, P.cPersDireccDomicilio, P.cPersTelefono, " _
            & " P.dPersNacCreac, ISNULL(Z.cDesZon,'') Zona,c.cConsDescripcion ,p.nPersPersoneria Personeria    " _
           & " FROM ( Select T.cpersCod, SUM(T.nSaldo) nSaldo FROM ( " _
            & " Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) " _
           & "     WHEN '1' THEN A.nSaldCntAC WHEN '2' THEN A.nSaldCntAC *" & pnTipCambio & "  END " _
          & "     FROM " & sservidorconsolidada & "AhorroCConsol A INNER JOIN " & sservidorconsolidada & "ProductoPersonaConsol PC ON A.cctacod = PC.cctacod INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) WHERE A.nEstCtaAC " _
         & "     NOT IN (1400,1300) AND PC.nPrdPersRelac = " & gCapRelPersTitular _
        & lsQuery _
       & " Union " _
      & " Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) " _
     & "        WHEN '1' THEN A.nSaldCntPF WHEN '2' THEN A.nSaldCntPF *" & pnTipCambio & "  END " _
    & " FROM " & sservidorconsolidada & "PlazoFijoConsol A Inner Join " & sservidorconsolidada & "ProductoPersonaConsol PC on A.cctacod = PC.cctacod INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) Where " _
            & "    A.nEstCtaPF not in (1400,1300) and PC.nPrdPersRelac = " & gCapRelPersTitular _
            & lsQuery _
            & " Union " _
            & " Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) " _
            & "       WHEN '1' THEN A.nSaldCntCTS WHEN '2' THEN A.nSaldCntCTS *" & pnTipCambio & " END " _
            & "       FROM " & sservidorconsolidada & "CTSConsol A Inner Join " & sservidorconsolidada & "ProductoPersonaConsol PC on A.cctacod = PC.cctacod INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) Where " _
            & " A.nEstCtaCTS not in (1400,1300) and PC.nPrdPersRelac = " & gCapRelPersTitular _
            & lsQuery _
            & " ) T GROUP BY T.cpersCod " _
            & " ) TA INNER JOIN Persona P  INNER JOIN Constante c ON c.nConsValor =p.nPersPersoneria and nConsCod='1002' LEFT JOIN " & sservidorconsolidada & "Zonas Z ON P.cPersDireccUbiGeo = Z.cCodZon " _
            & " ON TA.cpersCod = P.cpersCod ORDER BY TA.nSaldo DESC "
    Else
        lsSql = "  Select TOP " & pnNumero & " " _
            & " TA.cCodPers, P.cNomPers, TA.nSaldo, P.cDirPers, P.cTelPers, " _
            & " P.dFecNac, ISNULL(Z.cDesZon,'') Zona " _
            & " FROM ( Select T.cCodPers, SUM(T.nSaldo) nSaldo FROM ( " _
            & " Select PC.cCodPers, A.cCodCta, nSaldo = CASE SUBSTRING(A.cCodCta,6,1) " _
            & "     WHEN '1' THEN A.nSaldCntAC WHEN '2' THEN A.nSaldCntAC *" & pnTipCambio & "  END " _
            & "     FROM AhorroCConsol A INNER JOIN PersCuentaConsol PC ON A.cCodCta = PC.cCodCta WHERE A.cEstCtaAC " _
            & "     NOT IN ('C','U') AND PC.cRelaCta = 'TI' " _
            & " Union " _
            & " Select PC.cCodPers, A.cCodCta, nSaldo = CASE SUBSTRING(A.cCodCta,6,1) " _
            & "        WHEN '1' THEN A.nSaldCntPF WHEN '2' THEN A.nSaldCntPF *" & pnTipCambio & "  END " _
            & " FROM PlazoFijoConsol A Inner Join PersCuentaConsol PC on A.cCodCta = PC.cCodCta Where " _
            & "    A.cEstCtaPF not in ('C','U') and PC.cRelaCta = 'TI' " _
            & " Union " _
            & " Select PC.cCodPers, A.cCodCta, nSaldo = CASE SUBSTRING(A.cCodCta,6,1) " _
            & "       WHEN '1' THEN A.nSaldCntCTS WHEN '2' THEN A.nSaldCntCTS *" & pnTipCambio & " END " _
            & "       FROM CTSConsol A Inner Join PersCuentaConsol PC on A.cCodCta = PC.cCodCta Where " _
            & " A.cEstCtaCTS not in ('C','U') and PC.cRelaCta = 'TI' " _
            & " ) T GROUP BY T.cCodPers " _
            & " ) TA INNER JOIN dbPersona..Persona P LEFT JOIN Zonas Z ON P.cCodZon = Z.cCodZon " _
            & " ON TA.cCodPers = P.cCodPers ORDER BY TA.nSaldo DESC "
    End If
    End If
    
    
    lrReg.CursorLocation = adUseClient
    Set lrReg = oCon.CargaRecordSet(lsSql)
    Set lrReg.ActiveConnection = Nothing
    lnContador = 1
    lnIIni = I
    Do While Not lrReg.EOF
        If gbBitCentral = True Then
            xlHoja1.Cells(I, 1) = lnContador
            xlHoja1.Cells(I, 2) = lrReg!cPersCod
            xlHoja1.Cells(I, 3) = lrReg!cPersNombre
            xlHoja1.Cells(I, 4) = lrReg!cPersDireccDomicilio
            xlHoja1.Cells(I, 5) = lrReg!nSaldo
            xlHoja1.Cells(I, 6) = lrReg!cPersTelefono
            xlHoja1.Cells(I, 7) = Format(lrReg!dPersNacCreac, "dd/mm/yyyy")
            xlHoja1.Cells(I, 8) = lrReg!Zona
            xlHoja1.Cells(I, 9) = lrReg!cConsDescripcion
            xlHoja1.Cells(I, 10) = lrReg!Personeria
        Else
            xlHoja1.Cells(I, 1) = lnContador
            xlHoja1.Cells(I, 2) = lrReg!cCodPers
            xlHoja1.Cells(I, 3) = lrReg!cNomPers
            xlHoja1.Cells(I, 4) = lrReg!cDirPers
            xlHoja1.Cells(I, 5) = lrReg!nSaldo
            xlHoja1.Cells(I, 6) = lrReg!cTelPers
            xlHoja1.Cells(I, 7) = Format(lrReg!dFecNac, "dd/mm/yyyy")
            xlHoja1.Cells(I, 8) = lrReg!Zona
        End If
        
        lnContador = lnContador + 1
        I = I + 1
        lrReg.MoveNext
    Loop
    lrReg.Close
    
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 10)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
    
    oCon.CierraConexion
    
   MsgBox "Reporte Generado Satisfactoriamente"
End Sub

Private Sub GeneraRepColMejoresClientes(ByVal pnNumero As Long, ByVal pnTipCambio As Double, ByVal pdFecha As Date, ByVal psAgencia As String)
    Dim lsConec As String
    Dim lsSql As String
    Dim lrReg As New ADODB.Recordset
    Dim I As Integer, lnIIni As Integer
    Dim lnContador As Long
    Dim lsQuery As String
    Dim cVigenteTemp As String
    Dim bnBitCentral As Integer
    cVigenteTemp = cVigente & ",'2101','2104','2107','2201','2205'"
    cVigenteTemp = Replace(cVigenteTemp, "'", "")
    cVigenteTemp = "'" & cVigenteTemp & "'"
    
    Dim oCon As DConecta
    Set oCon = New DConecta

    If gbBitCentral = True Then
        oCon.AbreConexion
    Else
        oCon.AbreConexion ' Remota "07", , , "03"
    End If
    
   ExcelAddHoja "PrinClientesCol", xlLibro, xlHoja1
   xlHoja1.PageSetup.Orientation = xlPortrait
   xlHoja1.PageSetup.CenterHorizontally = True
   xlHoja1.PageSetup.Zoom = 55
   xlHoja1.Cells(1, 1) = gsNomCmac
   xlHoja1.Cells(2, 1) = "PRINCIPALES CLIENTES DE COLOCACIONES"
   xlHoja1.Cells(3, 1) = "AL " & Format(pdFecha, "dd/mm/yyyy")
   
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 8)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 8)).Merge True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 8)).HorizontalAlignment = xlCenter

   xlHoja1.Range("A1:A150").ColumnWidth = 5
   xlHoja1.Range("B1:B150").ColumnWidth = 10
   xlHoja1.Range("C1:C150").ColumnWidth = 40
   xlHoja1.Range("D1:D150").ColumnWidth = 30
   xlHoja1.Range("E1:E150").ColumnWidth = 12
   xlHoja1.Range("F1:F150").ColumnWidth = 12
   xlHoja1.Range("G1:G150").ColumnWidth = 12
   xlHoja1.Range("H1:H150").ColumnWidth = 30
   xlHoja1.Range("I1:I150").ColumnWidth = 30
   xlHoja1.Range("J1:J150").ColumnWidth = 20
   
   xlHoja1.Range(xlHoja1.Cells(6, 5), xlHoja1.Cells(150, 5)).NumberFormat = "#,##0.00;#,##0.00"
   xlHoja1.Range(xlHoja1.Cells(6, 7), xlHoja1.Cells(150, 7)).NumberFormat = "dd/mm/yyyy"
   
   I = 5
    
   xlHoja1.Cells(I, 1) = "Nro"
   xlHoja1.Cells(I, 2) = "Codigo"
   xlHoja1.Cells(I, 3) = "Cliente"
   xlHoja1.Cells(I, 4) = "Direccion"
   xlHoja1.Cells(I, 5) = "Saldo"
   xlHoja1.Cells(I, 6) = "Telefono"
   xlHoja1.Cells(I, 7) = "Fec.Nac."
   xlHoja1.Cells(I, 8) = "Zona"
   xlHoja1.Cells(I, 9) = "Personeria"
   xlHoja1.Cells(I, 10) = "Codigo Personeria"
   If gbBitCentral = True Then
   xlHoja1.Cells(I, 11) = "Nro Credito"
   xlHoja1.Cells(I, 12) = "Fecha de desembolso"
   xlHoja1.Cells(I, 13) = "Monto desembolsado"
   'xlHoja1.Cells(I, 14) = "Saldo deudor por cliente"
   xlHoja1.Cells(I, 14) = "Tipo de Crédito"
   xlHoja1.Cells(I, 15) = "Calificación del Cliente"
   End If
   If gbBitCentral = True Then
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 16)).HorizontalAlignment = xlCenter
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 16)).Cells.Borders.LineStyle = xlOutside
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 16)).Cells.Borders.LineStyle = xlInside
   Else
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 10)).HorizontalAlignment = xlCenter
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 10)).Cells.Borders.LineStyle = xlOutside
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 10)).Cells.Borders.LineStyle = xlInside
   
   End If
   I = I + 1
    
    
    'JEOM
    ''ALPA**09/07/2008
''    lsQuery = ""
''    If psAgencia = "Todos" Then
''       lsQuery = ""
''    Else
''       lsQuery = " and ag.cAgecod ='" & psAgencia & "'"
''    End If
    If gbBitCentral = True Then
    bnBitCentral = 1
    Else
    bnBitCentral = 0
    End If
''    If gbBitCentral = True Then
''        lsSql = " Select TOP " & pnNumero & " "
''        lsSql = lsSql & " TA.cPersCod, P.cPersNombre, TA.nSaldo, P.cPersDireccDomicilio, P.cPersTelefono,  "
''        lsSql = lsSql & " P.dPersNacCreac, ISNULL(Z.cDesZon,'') Zona ,c.cConsDescripcion ,p.nPersPersoneria Personeria "
''        '****ALPA******21/05/2008***************************************************************************************
''        lsSql = lsSql & " ,max(Pr.cCtaCod) cCtaCod,(select dVigencia from Colocaciones where cCtaCod=max(Pr.cCtaCod)) dVigencia ,(select nMontoCol from Colocaciones where cCtaCod=max(Pr.cCtaCod)) nMontoCol ,"
''        'lsSql = lsSql & " SDeudT,"
''        lsSql = lsSql & " (select cConsDescripcion from constante where substring(max(Pr.cCtaCod),6,3)=NConsValor and nConsCod=1001) cConsDescripcion, "
''        lsSql = lsSql & " cCalifActual= (select case when cCalGen='0' then '0 NORMAL' "
''        lsSql = lsSql & "       when cCalGen='1' then '1 CPP' "
''        lsSql = lsSql & "       when cCalGen='2' then '2 DEFICIENTE' "
''        lsSql = lsSql & "      when cCalGen='3' then '3 DUDOSO' "
''        lsSql = lsSql & "      when cCalGen='4' then '4 PERDIDA' end from ColocCalifProv where cCtaCod=max(Pr.cCtaCod)) " 'and datediff(day,dFecha, " & Format(pdFecha, "yyyy/mm/dd") & ")=0
''        '****END******21/05/2008*******d************6********************************************************************
''        lsSql = lsSql & " FROM ( Select T.cpersCod, SUM(T.nSaldo) nSaldo FROM "
''        lsSql = lsSql & "        (  Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) "
''        lsSql = lsSql & "            WHEN '1' THEN A.nSaldoCap WHEN '2' THEN A.nSaldoCap *" & pnTipCambio & "  END  "
''        lsSql = lsSql & "            FROM " & sservidorconsolidada & "CreditoConsol A INNER JOIN " & sservidorconsolidada & "productopersonaconsol PC "
''        lsSql = lsSql & "            ON A.cctacod = PC.cctacod INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) WHERE A.nPrdEstado IN (" & cVigenteTemp & ") AND PC.nPrdPersRelac =" & gColRelPersTitular
''        lsSql = lsSql & "            and substring(A.cctacod,6,3) not in (121,221) "
''        lsSql = lsSql & lsQuery
''        lsSql = lsSql & "            union "
''        lsSql = lsSql & "           Select PC.cpersCod, A.cctacod, nSaldo = CASE SUBSTRING(A.cctacod,9,1) "
''        lsSql = lsSql & "            WHEN '1' THEN A.nMontoApr WHEN '2' THEN A.nMontoApr *" & pnTipCambio & "  END  "
''        lsSql = lsSql & "            FROM " & sservidorconsolidada & "cartafianzaconsol A INNER JOIN productopersona PC "
''        lsSql = lsSql & "            ON A.cctacod = PC.cctacod INNER JOIN Agencias ag ON ag.cAgeCod = substring(a.cCtaCod,4,2) WHERE A.nPrdEstado IN (" & cVigenteTemp & ") AND PC.nPrdPersRelac =" & gColRelPersTitular
''        lsSql = lsSql & "            and substring(A.cctacod,6,3) in (121,221) "
''        lsSql = lsSql & lsQuery
''        lsSql = lsSql & "  ) T GROUP BY T.cPersCod  "
''        lsSql = lsSql & " ) TA INNER JOIN Persona P  "
''        lsSql = lsSql & " INNER JOIN Constante c ON c.nConsValor =p.nPersPersoneria and nConsCod='1002' "
''        lsSql = lsSql & "  LEFT JOIN " & sservidorconsolidada & "Zonas Z ON P.cPersDireccUbiGeo = Z.cCodZon "
''        lsSql = lsSql & " ON TA.cPersCod = P.cPersCod "
''        '****ALPA******21/05/2008***************************************************************************************
''        lsSql = lsSql & "inner join productoPersona PrP on (P.cPersCod = PrP.cPersCod) "
''        lsSql = lsSql & "inner join producto Pr on (Pr.cCtaCod=Prp.cCtaCod) and Prp.nPrdPersRelac=20 "
''        lsSql = lsSql & "inner join Colocaciones Coloc on (Pr.cCtaCod=Coloc.cCtaCod)"
''        lsSql = lsSql & "where Pr.nPrdEstado in (" & cVigenteTemp & ") "
''        lsSql = lsSql & "group by TA.cPersCod, P.cPersNombre, TA.nSaldo, P.cPersDireccDomicilio, "
''        lsSql = lsSql & "P.cPersTelefono,P.dPersNacCreac, Z.cDesZon, "
''        lsSql = lsSql & "c.cConsDescripcion ,p.nPersPersoneria "
''        '****END******21/05/2008***************************************************************************************
''        lsSql = lsSql & " ORDER BY TA.nSaldo DESC  "
''
''
''    Else
''        lsSql = " Select TOP " & pnNumero & " " _
''              & " TA.cCodPers, P.cNomPers, TA.nSaldo, P.cDirPers, P.cTelPers,  " _
''              & " P.dFecNac, ISNULL(Z.cDesZon,'') Zona " _
''              & " FROM ( Select T.cCodPers, SUM(T.nSaldo) nSaldo FROM " _
''              & "        (  Select PC.cCodPers, A.cCodCta, nSaldo = CASE SUBSTRING(A.cCodCta,6,1) " _
''              & "            WHEN '1' THEN A.nSaldoCap WHEN '2' THEN A.nSaldoCap *" & pnTipCambio & "  END  " _
''              & "            FROM CreditoConsol A INNER JOIN PersCreditoConsol PC " _
''              & "            ON A.cCodCta = PC.cCodCta WHERE A.cEstado IN ('F') AND PC.cRelaCta = 'TI' " _
''              & "  ) T GROUP BY T.cCodPers  " _
''              & " ) TA INNER JOIN dbPersona..Persona P  " _
''              & "  LEFT JOIN Zonas Z ON P.cCodZon = Z.cCodZon " _
''              & " ON TA.cCodPers = P.cCodPers " _
''              & " ORDER BY TA.nSaldo DESC  "
''    End If
    
    lrReg.CursorLocation = adUseClient
    lsSql = ""
    lsSql = "exec  stp_sel_ReporteColMejoresClientes " & bnBitCentral & "," & pnNumero & ",'" & Format(pdFecha, "yyyy/mm/dd") & "'," & pnTipCambio & "," & cVigenteTemp & ", '" & psAgencia & "'"
    Set lrReg = oCon.CargaRecordSet(lsSql)
    Set lrReg.ActiveConnection = Nothing
    lnContador = 1
    lnIIni = I
    Do While Not lrReg.EOF
        xlHoja1.Cells(I, 1) = lnContador
        
        If gbBitCentral = True Then
            xlHoja1.Cells(I, 2) = lrReg!cPersCod
            xlHoja1.Cells(I, 3) = lrReg!cPersNombre
            xlHoja1.Cells(I, 4) = lrReg!cPersDireccDomicilio
            xlHoja1.Cells(I, 5) = lrReg!nSaldo
            xlHoja1.Cells(I, 6) = lrReg!cPersTelefono
            xlHoja1.Cells(I, 7) = Format(lrReg!dPersNacCreac, "dd/mm/yyyy")
            xlHoja1.Cells(I, 8) = lrReg!Zona
            xlHoja1.Cells(I, 9) = lrReg!cConsDescripcion
            xlHoja1.Cells(I, 10) = lrReg!Personeria
             xlHoja1.Cells(I, 10) = lrReg!Personeria
             '****ALPA**21/05/2008**********
            xlHoja1.Cells(I, 11) = lrReg!cCtaCod
            xlHoja1.Cells(I, 12) = Format(lrReg!dVigencia, "YYYY/MM/DD")
            xlHoja1.Cells(I, 13) = Format(lrReg!nMontoCol, "###,####,####.##")
           ' xlHoja1.Cells(I, 14) = Format(lrReg!SDeudT, "###,####,####.##")
            xlHoja1.Cells(I, 14) = lrReg!cConsDescripcion
            xlHoja1.Cells(I, 15) = lrReg!cCalifActual
            '****END***************
        Else
            xlHoja1.Cells(I, 2) = lrReg!cCodPers
            xlHoja1.Cells(I, 3) = lrReg!cNomPers
            xlHoja1.Cells(I, 4) = lrReg!cDirPers
            xlHoja1.Cells(I, 5) = lrReg!nSaldo
            xlHoja1.Cells(I, 6) = lrReg!cTelPers
            xlHoja1.Cells(I, 7) = Format(lrReg!dFecNac, "dd/mm/yyyy")
            xlHoja1.Cells(I, 8) = lrReg!Zona
        End If
        
        lnContador = lnContador + 1
        I = I + 1
        lrReg.MoveNext
    Loop
    lrReg.Close
 If gbBitCentral = True Then
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 16)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 16)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
 Else
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 10)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 10)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
 End If
   
   oCon.CierraConexion
   
   MsgBox "Reporte Generado Satisfactoriamente"
End Sub

Private Sub GeneraRepConcentracionCred(ByVal pnTipCambio As Double, ByVal pdFecIni As String, ByVal pdFecFin As String)
'** Info
    Dim lsConec As String
    Dim lsSql As String
    Dim lrReg As New ADODB.Recordset
    Dim I As Integer, lnIIni As Integer
    Dim lnContador As Long
    
    Dim oCon As DConecta
    Set oCon = New DConecta

    If gbBitCentral = True Then
        oCon.AbreConexion
    Else
        oCon.AbreConexion 'Remota "07", , , "03"
    End If
    
   ExcelAddHoja "Info1", xlLibro, xlHoja1
   xlHoja1.PageSetup.Orientation = xlPortrait
   xlHoja1.PageSetup.CenterHorizontally = True
   xlHoja1.PageSetup.Zoom = 75
   xlHoja1.Cells(1, 1) = gsNomCmac
   xlHoja1.Cells(1, 8) = "TCF " & Str(pnTipCambio)
   xlHoja1.Cells(2, 1) = "REPORTE INFO 1"
   xlHoja1.Cells(3, 1) = "Informacion del " & Format(pdFecIni, "dd/mm/yyyy") & "  al  " & Format(pdFecFin, "dd/mm/yyyy")
   
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 8)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 8)).Merge True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 8)).HorizontalAlignment = xlCenter

   xlHoja1.Range("A1:A150").ColumnWidth = 20
   xlHoja1.Range("B1:B150").ColumnWidth = 12
   xlHoja1.Range("C1:C150").ColumnWidth = 15
   xlHoja1.Range("D1:D150").ColumnWidth = 12
   xlHoja1.Range("E1:E150").ColumnWidth = 15
   xlHoja1.Range("F1:F150").ColumnWidth = 12
   xlHoja1.Range("G1:G150").ColumnWidth = 15
   xlHoja1.Range("H1:H150").ColumnWidth = 12
   xlHoja1.Range("I1:I150").ColumnWidth = 15
   
   xlHoja1.Range(xlHoja1.Cells(6, 3), xlHoja1.Cells(150, 3)).NumberFormat = "#,##0.00;#,##0.00"
   xlHoja1.Range(xlHoja1.Cells(6, 5), xlHoja1.Cells(150, 5)).NumberFormat = "#,##0.00;#,##0.00"
   xlHoja1.Range(xlHoja1.Cells(6, 7), xlHoja1.Cells(150, 7)).NumberFormat = "#,##0.00;#,##0.00"
   xlHoja1.Range(xlHoja1.Cells(6, 9), xlHoja1.Cells(150, 9)).NumberFormat = "#,##0.00;#,##0.00"
   I = 5
   '*** Creditos Otorgados
   xlHoja1.Cells(I, 2) = "CREDITOS OTORGADOS"
   I = I + 1
   xlHoja1.Cells(I, 2) = "Pyme"
   xlHoja1.Cells(I + 1, 2) = "Nro"
   xlHoja1.Cells(I + 1, 3) = "Cartera"
   xlHoja1.Cells(I, 4) = "Agricolas"
   xlHoja1.Cells(I + 1, 4) = "Nro"
   xlHoja1.Cells(I + 1, 5) = "Cartera"
   xlHoja1.Cells(I, 6) = "Personales"
   xlHoja1.Cells(I + 1, 6) = "Nro"
   xlHoja1.Cells(I + 1, 7) = "Cartera"
   xlHoja1.Cells(I, 8) = "Prendario"
   xlHoja1.Cells(I + 1, 9) = "Nro"
   xlHoja1.Cells(I + 1, 9) = "Cartera"
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 9)).HorizontalAlignment = xlCenter
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 9)).Cells.Borders.LineStyle = xlOutside
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 9)).Cells.Borders.LineStyle = xlInside
   I = I + 2
    
    
'    Moneda.gMonedaNacional
'    Moneda.gMonedaExtranjera
'    Producto.gColComercEmp = 101
'    Producto.gColPYMEEmp = 201
'    Producto.gColPYMEAgro = 202
'    Producto.gColConsuDctoPlan = 301
'    Producto.gColConsuPlazoFijo = 302
'    Producto.gColConsCTS = 303
'    Producto.gColConsuUsosDiv = 304
'    Producto.gColConsuPrendario = 305
'    Producto.gColConsuPrestAdm = 320
    
    
    
    If gbBitCentral = True Then
        lsSql = " Select " & _
             " Count (c.cctacod) NumTot, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKTot , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') THEN c.cctacod END )  As NumPyme, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr END ), 0 )  As SKPyme , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') THEN c.cctacod END )  As NumAgri, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKAgri , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') THEN c.cctacod END )  As NumCons, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKCons , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') THEN c.cctacod END )  As NumPign , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPign   , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.ncondcre in ('" & gColocCredCondRecurrente & "','" & gColocCredCondParalelo & "') AND C.cRefinan ='N' THEN c.cctacod END )  As NumPymeRep, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.ncondcre in ('" & gColocCredCondRecurrente & "','" & gColocCredCondParalelo & "') AND C.cRefinan ='N' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             " WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.ncondcre in ('" & gColocCredCondRecurrente & "','" & gColocCredCondParalelo & "') AND C.cRefinan ='N' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr END ), 0 )  As SKPymeRep , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.ncondcre in ('" & gColocCredCondRecurrente & "','" & gColocCredCondParalelo & "') AND C.cRefinan ='N' THEN c.cctacod END )  As NumAgriRep, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.ncondcre in ('" & gColocCredCondRecurrente & "','" & gColocCredCondParalelo & "') AND C.cRefinan ='N' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.ncondcre in ('" & gColocCredCondRecurrente & "','" & gColocCredCondParalelo & "') AND C.cRefinan ='N' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKAgriRep , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.ncondcre in ('" & gColocCredCondRecurrente & "','" & gColocCredCondParalelo & "') AND C.cRefinan ='N' THEN c.cctacod END )  As NumConsRep, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.ncondcre in ('" & gColocCredCondRecurrente & "','" & gColocCredCondParalelo & "') AND C.cRefinan ='N' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.ncondcre in ('" & gColocCredCondRecurrente & "','" & gColocCredCondParalelo & "') AND C.cRefinan ='N' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKConsRep , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov > 0  THEN c.cctacod END )  As NumPignRep, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov >0  and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov >0  and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPignRep , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.ncondcre in ('" & gColocCredCondNormal & "')  THEN c.cctacod END )  As NumPymeNue , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.ncondcre in ('" & gColocCredCondNormal & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.ncondcre in ('" & gColocCredCondNormal & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPymeNue , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.ncondcre in ('" & gColocCredCondNormal & "')  THEN c.cctacod END )  As NumAgriNue , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.ncondcre in ('" & gColocCredCondNormal & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.ncondcre in ('" & gColocCredCondNormal & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr END ), 0 )  As SKAgriNue , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.ncondcre in ('" & gColocCredCondNormal & "')  THEN c.cctacod END )  As NumConsNue , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.ncondcre in ('" & gColocCredCondNormal & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.ncondcre in ('" & gColocCredCondNormal & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKConsNue , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov = 0  THEN c.cctacod END )  As NumPignNue, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov =0  and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov =0  and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPignNue  , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And  C.cRefinan ='R' THEN c.cctacod END )  As NumPymeRef, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') AND C.cRefinan ='R' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') AND C.cRefinan ='R' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPymeRef , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And  C.cRefinan ='R' THEN c.cctacod END )  As NumAgriRef, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') AND C.cRefinan ='R' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') AND C.cRefinan ='R' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKAgriRef , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And  C.cRefinan ='R' THEN c.cctacod END )  As NumConsRef, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') AND C.cRefinan ='R' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') AND C.cRefinan ='R' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKConsRef , " & _
             " 0 as NumPignRef, 0 SKPignRef " & _
             " From " & sservidorconsolidada & "CreditoConsol c " & _
             " Where c.nPrdEstado in (" & cVigente & "," & cPigno & ") and c.dFecVig Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' and '" & Format(pdFecFin, "mm/dd/yyyy") & " 23:59' "

    Else
         lsSql = " Select " & _
             " Count (c.cCodCta) NumTot, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKTot , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') THEN c.cCodCta END )  As NumPyme, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr END ), 0 )  As SKPyme , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') THEN c.cCodCta END )  As NumAgri, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKAgri , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') THEN c.cCodCta END )  As NumCons, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKCons , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') THEN c.cCodCta END )  As NumPign , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPign   , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And C.cCondCre IN ('R','P') AND C.cRefinan ='N' THEN c.cCodCta END )  As NumPymeRep, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And C.cCondCre IN ('R','P') AND C.cRefinan ='N' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             " WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And C.cCondCre IN ('R','P') AND C.cRefinan ='N' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr END ), 0 )  As SKPymeRep , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And C.cCondCre IN ('R','P') AND C.cRefinan ='N' THEN c.cCodCta END )  As NumAgriRep, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And C.cCondCre IN ('R','P') AND C.cRefinan ='N' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And C.cCondCre IN ('R','P') AND C.cRefinan ='N' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKAgriRep , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And C.cCondCre IN ('R','P') AND C.cRefinan ='N' THEN c.cCodCta END )  As NumConsRep, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And C.cCondCre IN ('R','P') AND C.cRefinan ='N' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And C.cCondCre IN ('R','P') AND C.cRefinan ='N' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKConsRep , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov > 0  THEN c.cCodCta END )  As NumPignRep, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov >0  and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov >0  and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPignRep , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And C.cCondCre IN ('N','A')  THEN c.cCodCta END )  As NumPymeNue , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And C.cCondCre IN ('N','A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And C.cCondCre IN ('N','A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPymeNue , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And C.cCondCre IN ('N','A')  THEN c.cCodCta END )  As NumAgriNue , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And C.cCondCre IN ('N','A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And C.cCondCre IN ('N','A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr END ), 0 )  As SKAgriNue , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And C.cCondCre IN ('N','A')  THEN c.cCodCta END )  As NumConsNue , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And C.cCondCre IN ('N','A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And C.cCondCre IN ('N','A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKConsNue , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov = 0  THEN c.cCodCta END )  As NumPignNue, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov =0  and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And C.nNumRenov =0  and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPignNue  , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And  C.cRefinan ='R' THEN c.cCodCta END )  As NumPymeRef, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') AND C.cRefinan ='R' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') AND C.cRefinan ='R' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPymeRef , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And  C.cRefinan ='R' THEN c.cCodCta END )  As NumAgriRef, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') AND C.cRefinan ='R' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') AND C.cRefinan ='R' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKAgriRef , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And  C.cRefinan ='R' THEN c.cCodCta END )  As NumConsRef, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') AND C.cRefinan ='R' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') AND C.cRefinan ='R' and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKConsRef , " & _
             " 0 as NumPignRef, 0 SKPignRef " & _
             " From CreditoConsol c " & _
             " Where c.cEstado in ('F','1','4','6','7') and c.dFecVig Between '" & Format(pdFecIni, "mm/dd/yyyy") & "' and '" & Format(pdFecFin, "mm/dd/yyyy") & " 23:59' "
    End If
    
    lrReg.CursorLocation = adUseClient
    Set lrReg = oCon.CargaRecordSet(lsSql)
    Set lrReg.ActiveConnection = Nothing
    
    lnContador = 1
    lnIIni = I
    xlHoja1.Cells(I, 2) = "CREDITOS OTORGADOS"
    I = I + 1
    xlHoja1.Cells(I, 1) = "TOTAL"
    xlHoja1.Cells(I, 2) = lrReg!NumPyme:         xlHoja1.Cells(I, 3) = lrReg!SKPyme
    xlHoja1.Cells(I, 4) = lrReg!NumAgri:         xlHoja1.Cells(I, 5) = lrReg!SKAgri
    xlHoja1.Cells(I, 6) = lrReg!NumCons:         xlHoja1.Cells(I, 7) = lrReg!SKCons
    xlHoja1.Cells(I, 8) = lrReg!NumPign:         xlHoja1.Cells(I, 9) = lrReg!SKPign
    xlHoja1.Cells(I + 1, 1) = "Represtamos"
    xlHoja1.Cells(I + 1, 2) = lrReg!NumPymeRep:  xlHoja1.Cells(I + 1, 3) = lrReg!SKPymeRep
    xlHoja1.Cells(I + 1, 4) = lrReg!NumAgriRep:  xlHoja1.Cells(I + 1, 5) = lrReg!SKAgriRep
    xlHoja1.Cells(I + 1, 6) = lrReg!NumConsRep:  xlHoja1.Cells(I + 1, 7) = lrReg!SKConsRep
    xlHoja1.Cells(I + 1, 8) = lrReg!NumPignRep:  xlHoja1.Cells(I + 1, 9) = lrReg!SKPignRep
    xlHoja1.Cells(I + 2, 1) = "Nuevos"
    xlHoja1.Cells(I + 2, 2) = lrReg!NumPymeNue:  xlHoja1.Cells(I + 2, 3) = lrReg!SKPymeNue
    xlHoja1.Cells(I + 2, 4) = lrReg!NumAgriNue:  xlHoja1.Cells(I + 2, 5) = lrReg!SKAgriNue
    xlHoja1.Cells(I + 2, 6) = lrReg!NumConsNue:  xlHoja1.Cells(I + 2, 7) = lrReg!SKConsNue
    xlHoja1.Cells(I + 2, 8) = lrReg!NumPignNue:  xlHoja1.Cells(I + 2, 9) = lrReg!SKPignNue
    xlHoja1.Cells(I + 3, 1) = "Refinanciados"
    xlHoja1.Cells(I + 3, 2) = lrReg!NumPymeRef:  xlHoja1.Cells(I + 3, 3) = lrReg!SKPymeRef
    xlHoja1.Cells(I + 3, 4) = lrReg!NumAgriRef:  xlHoja1.Cells(I + 3, 5) = lrReg!SKAgriRef
    xlHoja1.Cells(I + 3, 6) = lrReg!NumConsRef:  xlHoja1.Cells(I + 3, 7) = lrReg!SKConsRef
    xlHoja1.Cells(I + 3, 8) = lrReg!NumPignRef:  xlHoja1.Cells(I + 3, 9) = lrReg!SKPignRef
    
    lnContador = lnContador + 1
    I = I + 4
    
    lrReg.Close
    
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 9)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
   
   '*** Creditos Vigentes
   I = I + 1
   xlHoja1.Cells(I, 2) = "CREDITOS VIGENTES"
   I = I + 1
   xlHoja1.Cells(I, 2) = "Pyme"
   xlHoja1.Cells(I + 1, 2) = "Nro"
   xlHoja1.Cells(I + 1, 3) = "Cartera"
   xlHoja1.Cells(I, 4) = "Agricolas"
   xlHoja1.Cells(I + 1, 4) = "Nro"
   xlHoja1.Cells(I + 1, 5) = "Cartera"
   xlHoja1.Cells(I, 6) = "Personales"
   xlHoja1.Cells(I + 1, 6) = "Nro"
   xlHoja1.Cells(I + 1, 7) = "Cartera"
   xlHoja1.Cells(I, 8) = "Prendario"
   xlHoja1.Cells(I + 1, 8) = "Nro"
   xlHoja1.Cells(I + 1, 9) = "Cartera"
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 1, 9)).HorizontalAlignment = xlCenter
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 1, 9)).Cells.Borders.LineStyle = xlOutside
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I + 1, 9)).Cells.Borders.LineStyle = xlInside
   I = I + 2
    If gbBitCentral = True Then
        lsSql = " Select " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') THEN c.cctacod END )  As NumPyme, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPyme , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') THEN c.cctacod END )  As NumAgri, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKAgri , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') THEN c.cctacod END )  As NumCons, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKCons , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') THEN c.cctacod END )  As NumPign , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPign   , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 1 and 7 THEN c.cctacod END )  As NumPymeMora1, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPymeMora1 , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 1 and 7 THEN c.cctacod END )  As NumAgriMora1, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKAgriMora1 , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 1 and 7 THEN c.cctacod END )  As NumConsMora1, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKConsMora1 , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 1 and 7 THEN c.cctacod END )  As NumPignMora1 , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPignMora1   , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 8 and 30 THEN c.cctacod END )  As NumPymeMora2, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPymeMora2 , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 8 and 30 THEN c.cctacod END )  As NumAgriMora2, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKAgriMora2 , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 8 and 30 THEN c.cctacod END )  As NumConsMora2, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKConsMora2 , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 8 and 30 THEN c.cctacod END )  As NumPignMora2 , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPignMora2   , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso > 30 THEN c.cctacod END )  As NumPymeMora3, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso > 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso > 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPymeMora3 , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso > 30 THEN c.cctacod END )  As NumAgriMora3, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso > 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso > 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKAgriMora3 , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso > 30 THEN c.cctacod END )  As NumConsMora3, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso > 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso > 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKConsMora3 , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso > 30 THEN c.cctacod END )  As NumPignMora3 , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso > 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso > 30 and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPignMora3    " & _
             " From " & sservidorconsolidada & "CreditoConsol c " & _
             " Where c.nPrdEstado in (" & cVigente & "," & cPigno & ")  "

    Else
         lsSql = " Select " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') THEN c.cCodCta END )  As NumPyme, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPyme , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') THEN c.cCodCta END )  As NumAgri, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKAgri , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') THEN c.cCodCta END )  As NumCons, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKCons , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') THEN c.cCodCta END )  As NumPign , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPign   , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 1 and 7 THEN c.cCodCta END )  As NumPymeMora1, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPymeMora1 , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 1 and 7 THEN c.cCodCta END )  As NumAgriMora1, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKAgriMora1 , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 1 and 7 THEN c.cCodCta END )  As NumConsMora1, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKConsMora1 , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 1 and 7 THEN c.cCodCta END )  As NumPignMora1 , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 1 and 7 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPignMora1   , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 8 and 30 THEN c.cCodCta END )  As NumPymeMora2, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPymeMora2 , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 8 and 30 THEN c.cCodCta END )  As NumAgriMora2, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKAgriMora2 , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 8 and 30 THEN c.cCodCta END )  As NumConsMora2, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKConsMora2 , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 8 and 30 THEN c.cCodCta END )  As NumPignMora2 , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso between 8 and 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPignMora2   , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso > 30 THEN c.cCodCta END )  As NumPymeMora3, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso > 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.nDiasAtraso > 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPymeMora3 , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso > 30 THEN c.cCodCta END )  As NumAgriMora3, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso > 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.nDiasAtraso > 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKAgriMora3 , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso > 30 THEN c.cCodCta END )  As NumConsMora3, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso > 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.nDiasAtraso > 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKConsMora3 , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso > 30 THEN c.cCodCta END )  As NumPignMora3 , " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso > 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr " & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuPrendario & "') And c.nDiasAtraso > 30 and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr *" & pnTipCambio & " END ), 0 )  As SKPignMora3    " & _
             " From CreditoConsol c " & _
             " Where c.cEstado in ('F','1','4','6','7')  "

    End If
    
    lrReg.CursorLocation = adUseClient
    Set lrReg = oCon.CargaRecordSet(lsSql)
    Set lrReg.ActiveConnection = Nothing
    
    lnContador = 1
    lnIIni = I
    xlHoja1.Cells(I, 1) = "TOTAL"
    xlHoja1.Cells(I, 2) = lrReg!NumPyme:         xlHoja1.Cells(I, 3) = lrReg!SKPyme
    xlHoja1.Cells(I, 4) = lrReg!NumAgri:         xlHoja1.Cells(I, 5) = lrReg!SKAgri
    xlHoja1.Cells(I, 6) = lrReg!NumCons:         xlHoja1.Cells(I, 7) = lrReg!SKCons
    xlHoja1.Cells(I, 8) = lrReg!NumPign:         xlHoja1.Cells(I, 9) = lrReg!SKPign
    xlHoja1.Cells(I + 1, 1) = "Mora 1 - 7 Dias"
    xlHoja1.Cells(I + 1, 2) = lrReg!NumPymeMora1:  xlHoja1.Cells(I + 1, 3) = lrReg!SKPymeMora1
    xlHoja1.Cells(I + 1, 4) = lrReg!NumAgriMora1:  xlHoja1.Cells(I + 1, 5) = lrReg!SKAgriMora1
    xlHoja1.Cells(I + 1, 6) = lrReg!NumConsMora1:  xlHoja1.Cells(I + 1, 7) = lrReg!SKConsMora1
    xlHoja1.Cells(I + 1, 8) = lrReg!NumPignMora1:  xlHoja1.Cells(I + 1, 9) = lrReg!SKPignMora1
    xlHoja1.Cells(I + 2, 1) = "Mora 8 - 30 Dias"
    xlHoja1.Cells(I + 2, 2) = lrReg!NumPymeMora2:  xlHoja1.Cells(I + 2, 3) = lrReg!SKPymeMora2
    xlHoja1.Cells(I + 2, 4) = lrReg!NumAgriMora2:  xlHoja1.Cells(I + 2, 5) = lrReg!SKAgriMora2
    xlHoja1.Cells(I + 2, 6) = lrReg!NumConsMora2:  xlHoja1.Cells(I + 2, 7) = lrReg!SKConsMora2
    xlHoja1.Cells(I + 2, 8) = lrReg!NumPignMora2:  xlHoja1.Cells(I + 2, 9) = lrReg!SKPignMora2
    xlHoja1.Cells(I + 3, 1) = "Mora > 30 Dias"
    xlHoja1.Cells(I + 3, 2) = lrReg!NumPymeMora3:  xlHoja1.Cells(I + 3, 3) = lrReg!SKPymeMora3
    xlHoja1.Cells(I + 3, 4) = lrReg!NumAgriMora3:  xlHoja1.Cells(I + 3, 5) = lrReg!SKAgriMora3
    xlHoja1.Cells(I + 3, 6) = lrReg!NumConsMora3:  xlHoja1.Cells(I + 3, 7) = lrReg!SKConsMora3
    xlHoja1.Cells(I + 3, 8) = lrReg!NumPignMora3:  xlHoja1.Cells(I + 3, 9) = lrReg!SKPignMora3
    
    lnContador = lnContador + 1
    I = I + 4
    lrReg.Close
    '** Creditos Recuperaciones
    
    If gbBitCentral = True Then
         lsSql = " Select " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and c.nprdestado='" & gColocEstRecVigJud & "' THEN c.cctacod END )  As NumPymeJud, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and c.nprdestado='" & gColocEstRecVigJud & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and c.nprdestado='" & gColocEstRecVigJud & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPymeJud , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') and c.nprdestado='" & gColocEstRecVigJud & "' THEN c.cctacod END )  As NumAgriJud, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') and c.nprdestado='" & gColocEstRecVigJud & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') and c.nprdestado='" & gColocEstRecVigJud & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKAgriJud , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and c.nprdestado='" & gColocEstRecVigJud & "' THEN c.cctacod END )  As NumConsJud, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and c.nprdestado='" & gColocEstRecVigJud & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and c.nprdestado='" & gColocEstRecVigJud & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr END ), 0 )  As SKConsJud , " & _
             " 0 As NumPignJud , 0 As SKPignJud   , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and c.nprdestado='" & gColocEstRecVigCast & "' THEN c.cctacod END )  As NumPymeCas, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and c.nprdestado='" & gColocEstRecVigCast & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') and c.nprdestado='" & gColocEstRecVigCast & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPymeCas , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') and c.nprdestado='" & gColocEstRecVigCast & "' THEN c.cctacod END )  As NumAgriCas, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') and c.nprdestado='" & gColocEstRecVigCast & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColPYMEAgro & "') and c.nprdestado='" & gColocEstRecVigCast & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKAgriCas , " & _
             " COUNT( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and c.nprdestado='" & gColocEstRecVigCast & "' THEN c.cctacod END )  As NumConsCas, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and c.nprdestado='" & gColocEstRecVigCast & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cctacod,6,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') and c.nprdestado='" & gColocEstRecVigCast & "' and Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKConsCas , " & _
             " 0   As NumPignCas , 0  As SKPignCas   " & _
             " From  " & sservidorconsolidada & "CreditoConsol c " & _
             " Where c.nprdEstado in (" & gColocEstRecVigJud & "," & gColocEstRecVigCast & ") And c.nSaldoCap > 0 "
    
    Else
         lsSql = " Select " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.cCondCre in ('J') THEN c.cCodCta END )  As NumPymeJud, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.cCondCre in ('J') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.cCondCre in ('J') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPymeJud , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.cCondCre in ('J') THEN c.cCodCta END )  As NumAgriJud, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.cCondCre in ('J') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.cCondCre in ('J') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKAgriJud , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.cCondCre in ('J') THEN c.cCodCta END )  As NumConsJud, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.cCondCre in ('J') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.cCondCre in ('J') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr END ), 0 )  As SKConsJud , " & _
             " 0 As NumPignJud , 0 As SKPignJud   , "
        lsSql = lsSql & " " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.cCondCre in ('A') THEN c.cCodCta END )  As NumPymeCas, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.cCondCre in ('A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEEmp & "', '" & Producto.gColComercEmp & "') And c.cCondCre in ('A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKPymeCas , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.cCondCre in ('A') THEN c.cCodCta END )  As NumAgriCas, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.cCondCre in ('A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColPYMEAgro & "') And c.cCondCre in ('A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKAgriCas , " & _
             " COUNT( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.cCondCre in ('A') THEN c.cCodCta END )  As NumConsCas, " & _
             " IsNull(SUM( CASE WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.cCondCre in ('A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
             "                  WHEN Substring(c.cCodCta,3,3) in('" & Producto.gColConsuDctoPlan & "', '" & Producto.gColConsuPlazoFijo & "', '" & Producto.gColConsCTS & "', '" & Producto.gColConsuUsosDiv & "', '" & Producto.gColConsuPrestAdm & "') And c.cCondCre in ('A') and Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKConsCas , " & _
             " 0   As NumPignCas , 0  As SKPignCas   " & _
             " From CreditoConsol c " & _
             " Where c.cEstado in ('V') And c.cCondCre in ('J','A') And c.nSaldoCap > 0 "
    End If
    
    lrReg.CursorLocation = adUseClient
    Set lrReg = oCon.CargaRecordSet(lsSql)
    Set lrReg.ActiveConnection = Nothing
    
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 9)).Borders(xlInsideVertical).LineStyle = xlContinuous
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I - 1, 9)).BorderAround xlContinuous, xlThick, xlColorIndexAutomatic
    lrReg.Close
    
' *********  INFO 3
   ExcelAddHoja "Info3", xlLibro, xlHoja1
   xlHoja1.PageSetup.Orientation = xlPortrait
   xlHoja1.PageSetup.CenterHorizontally = True
   xlHoja1.PageSetup.Zoom = 90
   xlHoja1.Cells(1, 1) = gsNomCmac
   xlHoja1.Cells(1, 3) = "TCF " & Str(pnTipCambio)
   xlHoja1.Cells(2, 1) = "REPORTE INFO 3"
   xlHoja1.Cells(3, 1) = "Informacion AL " & Format(pdFecFin, "dd/mm/yyyy")
   
   
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 3)).Font.Bold = True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 3)).Merge True
   xlHoja1.Range(xlHoja1.Cells(2, 1), xlHoja1.Cells(3, 3)).HorizontalAlignment = xlCenter

   xlHoja1.Range("A1:A150").ColumnWidth = 30
   xlHoja1.Range("B1:B150").ColumnWidth = 12
   xlHoja1.Range("C1:C150").ColumnWidth = 15
   xlHoja1.Range("D1:D150").ColumnWidth = 12
   xlHoja1.Range("E1:E150").ColumnWidth = 15
   xlHoja1.Range("F1:F150").ColumnWidth = 12
   xlHoja1.Range("G1:G150").ColumnWidth = 15
   xlHoja1.Range("H1:H150").ColumnWidth = 12
   xlHoja1.Range("I1:I150").ColumnWidth = 15
   
   xlHoja1.Range(xlHoja1.Cells(5, 3), xlHoja1.Cells(150, 3)).NumberFormat = "#,##0.00;#,##0.00"
   xlHoja1.Range(xlHoja1.Cells(5, 5), xlHoja1.Cells(150, 5)).NumberFormat = "#,##0.00;#,##0.00"
   xlHoja1.Range(xlHoja1.Cells(5, 7), xlHoja1.Cells(150, 7)).NumberFormat = "#,##0.00;#,##0.00"
   xlHoja1.Range(xlHoja1.Cells(5, 9), xlHoja1.Cells(150, 9)).NumberFormat = "#,##0.00;#,##0.00"
   I = 5
   '*** Estratificacion por montos
   xlHoja1.Cells(I, 1) = ""
   xlHoja1.Cells(I, 2) = "Nro"
   xlHoja1.Cells(I, 3) = "Cartera"
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).HorizontalAlignment = xlCenter
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Cells.Borders.LineStyle = xlOutside
   xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Cells.Borders.LineStyle = xlInside
   I = I + 1
   
   
   'sectores pepe
   If gbBitCentral = True Then
        lsSql = " select " & _
               " COUNT( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.cctacod " & _
               "             WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.cctacod  END ) As NumTot, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKTot  , " & _
               " COUNT( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr <= 500 /" & pnTipCambio & " THEN c.cctacod " & _
               "             WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr <= 500  THEN c.cctacod  END ) As NumRang1, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr <= 500 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr <= 500  THEN c.nMontoApr  END ), 0 )  As SKRang1  , " & _
               " COUNT( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 500 /" & pnTipCambio & " and nMontoApr <= 1000 /" & pnTipCambio & "  THEN c.cctacod " & _
               "             WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 500 and nMontoApr <=1000  THEN c.cctacod  END ) As NumRang2 , " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 500 /" & pnTipCambio & " and nMontoApr <= 1000 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 500 and nMontoApr <= 1000 THEN c.nMontoApr  END ), 0 )  As SKRang2 , " & _
               " COUNT( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 1000 /" & pnTipCambio & " and nMontoApr <= 2000 /" & pnTipCambio & " THEN c.cctacod " & _
               "             WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 1000 and nMontoApr <= 2000  THEN c.cctacod  END ) As NumRang3, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 1000 /" & pnTipCambio & " and nMontoApr <= 2000 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 1000 and nMontoApr <= 2000 THEN c.nMontoApr  END ), 0 )  As SKRang3 , " & _
               " COUNT( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 2000 /" & pnTipCambio & " and nMontoApr <= 5000 /" & pnTipCambio & "  THEN c.cctacod " & _
               "             WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 2000 and nMontoApr <= 5000  THEN c.cctacod  END ) As NumRang4, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 2000 /" & pnTipCambio & " and nMontoApr <= 5000 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 2000 and nMontoApr <= 5000 THEN c.nMontoApr  END ), 0 )  As SKRang4 , " & _
               " COUNT( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 5000 /" & pnTipCambio & " and nMontoApr <= 10000 /" & pnTipCambio & "  THEN c.cctacod " & _
               "             WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 5000 and nMontoApr <= 10000  THEN c.cctacod  END ) As NumRang5, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 5000 /" & pnTipCambio & " and nMontoApr <= 10000 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 5000 and nMontoApr <= 10000 THEN c.nMontoApr  END ), 0 )  As SKRang5 , "
        lsSql = lsSql & "  " & _
               " COUNT( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 10000 /" & pnTipCambio & " and nMontoApr <= 20000 /" & pnTipCambio & "  THEN c.cctacod " & _
               "             WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 10000 and nMontoApr <= 20000  THEN c.cctacod  END ) As NumRang6, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 10000 /" & pnTipCambio & " and nMontoApr <= 20000 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 10000 and nMontoApr <= 20000 THEN c.nMontoApr  END ), 0 )  As SKRang6 , " & _
               " COUNT( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 20000 /" & pnTipCambio & "   THEN c.cctacod " & _
               "             WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 20000  THEN c.cctacod  END ) As NumRang7, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 20000 /" & pnTipCambio & "  THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 20000  THEN c.nMontoApr  END ), 0 )  As SKRang7 , "
         lsSql = lsSql & " " & _
               " COUNT( case when  (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  < 120  then c.cctacod end) As NumPlazo1,  " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  < 120  then c.nMontoApr  / " & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  < 120  then c.nMontoApr  END ), 0 ) As SKPlazo1 , " & _
               " COUNT( case when  (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr) between 121 and 180  then c.cctacod end) As NumPlazo2,  " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 121 and 180  then c.nMontoApr  / " & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 121 and 180  then c.nMontoApr  END ), 0 ) As SKPlazo2 , " & _
               " COUNT( case when  (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr) between 181 and 360  then c.cctacod end) As NumPlazo3,  " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 181 and 360  then c.nMontoApr  / " & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 181 and 360  then c.nMontoApr  END ), 0 ) As SKPlazo3 , " & _
               " COUNT( case when  (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr) between 361 and 720  then c.cctacod end) As NumPlazo4,  " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 361 and 720  then c.nMontoApr  / " & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 361 and 720  then c.nMontoApr  END ), 0 ) As SKPlazo4 , " & _
               " COUNT( case when  (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  > 720  then c.cctacod end) As NumPlazo5,  " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  > 720  then c.nMontoApr  / " & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  > 720  then c.nMontoApr  END ), 0 ) As SKPlazo5 , "
        lsSql = lsSql & "  " & _
               " COUNT( CASE WHEN f.cSector in ('S', 'X', 'E')  THEN c.cctacod END )  As NumSect1, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and f.cSector in ('S', 'X', 'E') THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and f.cSector in ('S', 'X', 'E') THEN c.nMontoApr  END ), 0 )  As SKSect1 , " & _
               " COUNT( CASE WHEN f.cSector in ('C', 'R')  THEN c.cctacod END )  As NumSect2, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and f.cSector in ('C', 'R') THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and f.cSector in ('C', 'R') THEN c.nMontoApr  END ), 0 )  As SKSect2 , " & _
               " COUNT( CASE WHEN f.cSector in ('I', 'M', 'Q', 'P')  THEN c.cctacod END )  As NumSect3, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and f.cSector in ( 'I', 'M', 'Q', 'P') THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and f.cSector in ( 'I', 'M', 'Q', 'P') THEN c.nMontoApr  END ), 0 )  As SKSect3 , " & _
               " COUNT( CASE WHEN f.cSector in ('A')  THEN c.cctacod END )  As NumSect4, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and f.cSector in ('A') THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and f.cSector in ('A') THEN c.nMontoApr  END ), 0 )  As SKSect4 , " & _
               " COUNT( CASE WHEN c.nDestcre in (1)  THEN c.cctacod END )  As NumDest1, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and c.nDestcre in (1) THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and c.nDestcre in (1) THEN c.nMontoApr  END ), 0 )  As SKDest1 , " & _
               " COUNT( CASE WHEN c.nDestcre in (1) THEN c.cctacod END )  As NumDest2, " & _
               " IsNull(SUM( CASE WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaNacional & "' and c.nDestcre not in (1) THEN c.nMontoApr /" & pnTipCambio & _
               "                  WHEN Substring(c.cctacod,9,1) ='" & Moneda.gMonedaExtranjera & "' and c.nDestcre not in (1) THEN c.nMontoApr  END ), 0 )  As SKDest2  " & _
               " From  " & sservidorconsolidada & "CreditoConsol  c " & _
               " Left Join " & sservidorconsolidada & "FuenteIngresoConsol f On c.cNumFuente = f.cNumFuente " & _
               " Where nprdestado in (" & cVigente & ") " & _
               " And Substring(c.cctacod,6,3 ) in ('" & Producto.gColComercEmp & "', '" & Producto.gColPYMEEmp & "') "

   Else
        lsSql = " select " & _
                " COUNT( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.cCodCta " & _
                "             WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.cCodCta  END ) As NumTot, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' THEN c.nMontoApr  END ), 0 )  As SKTot  , " & _
                " COUNT( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr <= 500 /" & pnTipCambio & " THEN c.cCodCta " & _
                "             WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr <= 500  THEN c.cCodCta  END ) As NumRang1, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr <= 500 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr <= 500  THEN c.nMontoApr  END ), 0 )  As SKRang1  , " & _
                " COUNT( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 500 /" & pnTipCambio & " and nMontoApr <= 1000 /" & pnTipCambio & "  THEN c.cCodCta " & _
                "             WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 500 and nMontoApr <=1000  THEN c.cCodCta  END ) As NumRang2 , " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 500 /" & pnTipCambio & " and nMontoApr <= 1000 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 500 and nMontoApr <= 1000 THEN c.nMontoApr  END ), 0 )  As SKRang2 , " & _
                " COUNT( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 1000 /" & pnTipCambio & " and nMontoApr <= 2000 /" & pnTipCambio & " THEN c.cCodCta " & _
                "             WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 1000 and nMontoApr <= 2000  THEN c.cCodCta  END ) As NumRang3, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 1000 /" & pnTipCambio & " and nMontoApr <= 2000 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 1000 and nMontoApr <= 2000 THEN c.nMontoApr  END ), 0 )  As SKRang3 , " & _
                " COUNT( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 2000 /" & pnTipCambio & " and nMontoApr <= 5000 /" & pnTipCambio & "  THEN c.cCodCta " & _
                "             WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 2000 and nMontoApr <= 5000  THEN c.cCodCta  END ) As NumRang4, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 2000 /" & pnTipCambio & " and nMontoApr <= 5000 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 2000 and nMontoApr <= 5000 THEN c.nMontoApr  END ), 0 )  As SKRang4 , " & _
                " COUNT( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 5000 /" & pnTipCambio & " and nMontoApr <= 10000 /" & pnTipCambio & "  THEN c.cCodCta " & _
                "             WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 5000 and nMontoApr <= 10000  THEN c.cCodCta  END ) As NumRang5, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 5000 /" & pnTipCambio & " and nMontoApr <= 10000 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 5000 and nMontoApr <= 10000 THEN c.nMontoApr  END ), 0 )  As SKRang5 , "
         lsSql = lsSql & "  " & _
                " COUNT( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 10000 /" & pnTipCambio & " and nMontoApr <= 20000 /" & pnTipCambio & "  THEN c.cCodCta " & _
                "             WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 10000 and nMontoApr <= 20000  THEN c.cCodCta  END ) As NumRang6, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 10000 /" & pnTipCambio & " and nMontoApr <= 20000 /" & pnTipCambio & " THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 10000 and nMontoApr <= 20000 THEN c.nMontoApr  END ), 0 )  As SKRang6 , " & _
                " COUNT( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 20000 /" & pnTipCambio & "   THEN c.cCodCta " & _
                "             WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 20000  THEN c.cCodCta  END ) As NumRang7, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and nMontoApr > 20000 /" & pnTipCambio & "  THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and nMontoApr > 20000  THEN c.nMontoApr  END ), 0 )  As SKRang7 , "
          lsSql = lsSql & " " & _
                " COUNT( case when  (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  < 120  then c.ccodcta end) As NumPlazo1,  " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  < 120  then c.nMontoApr  / " & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  < 120  then c.nMontoApr  END ), 0 ) As SKPlazo1 , " & _
                " COUNT( case when  (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr) between 121 and 180  then c.ccodcta end) As NumPlazo2,  " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 121 and 180  then c.nMontoApr  / " & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 121 and 180  then c.nMontoApr  END ), 0 ) As SKPlazo2 , " & _
                " COUNT( case when  (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr) between 181 and 360  then c.ccodcta end) As NumPlazo3,  " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 181 and 360  then c.nMontoApr  / " & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 181 and 360  then c.nMontoApr  END ), 0 ) As SKPlazo3 , " & _
                " COUNT( case when  (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr) between 361 and 720  then c.ccodcta end) As NumPlazo4,  " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 361 and 720  then c.nMontoApr  / " & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  between 361 and 720  then c.nMontoApr  END ), 0 ) As SKPlazo4 , " & _
                " COUNT( case when  (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  > 720  then c.ccodcta end) As NumPlazo5,  " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  > 720  then c.nMontoApr  / " & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' And (case when nPlazoApr = 0 then 30 when nPlazoApr = null then 0 else nPlazoApr end *  nCuotasApr)  > 720  then c.nMontoApr  END ), 0 ) As SKPlazo5 , "
         lsSql = lsSql & "  " & _
                " COUNT( CASE WHEN f.cSector in ('S', 'X', 'E')  THEN c.cCodCta END )  As NumSect1, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and f.cSector in ('S', 'X', 'E') THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and f.cSector in ('S', 'X', 'E') THEN c.nMontoApr  END ), 0 )  As SKSect1 , " & _
                " COUNT( CASE WHEN f.cSector in ('C', 'R')  THEN c.cCodCta END )  As NumSect2, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and f.cSector in ('C', 'R') THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and f.cSector in ('C', 'R') THEN c.nMontoApr  END ), 0 )  As SKSect2 , " & _
                " COUNT( CASE WHEN f.cSector in ('I', 'M', 'Q', 'P')  THEN c.cCodCta END )  As NumSect3, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and f.cSector in ( 'I', 'M', 'Q', 'P') THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and f.cSector in ( 'I', 'M', 'Q', 'P') THEN c.nMontoApr  END ), 0 )  As SKSect3 , " & _
                " COUNT( CASE WHEN f.cSector in ('A')  THEN c.cCodCta END )  As NumSect4, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and f.cSector in ('A') THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and f.cSector in ('A') THEN c.nMontoApr  END ), 0 )  As SKSect4 , " & _
                " COUNT( CASE WHEN c.cDestcre in ('1')  THEN c.cCodCta END )  As NumDest1, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and c.cDestcre in ('1') THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and c.cDestcre in ('1') THEN c.nMontoApr  END ), 0 )  As SKDest1 , " & _
                " COUNT( CASE WHEN c.cDestcre not in ('1') THEN c.cCodCta END )  As NumDest2, " & _
                " IsNull(SUM( CASE WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaNacional & "' and c.cDestcre not in ('1') THEN c.nMontoApr /" & pnTipCambio & _
                "                  WHEN Substring(c.cCodCta,6,1) ='" & Moneda.gMonedaExtranjera & "' and c.cDestcre not in ('1') THEN c.nMontoApr  END ), 0 )  As SKDest2  " & _
                " From CreditoConsol  c " & _
                " Left Join FuenteIngresoConsol f On c.cNumFuente = f.cNumFuente " & _
                " Where cestado in ('F') " & _
                " And Substring(c.cCodCta,3,3 ) in ('" & Producto.gColComercEmp & "', '" & Producto.gColPYMEEmp & "') "
    End If
    lrReg.CursorLocation = adUseClient
    Set lrReg = oCon.CargaRecordSet(lsSql)
    Set lrReg.ActiveConnection = Nothing
    
    lnContador = 1
    I = I + 1
    xlHoja1.Cells(I, 1) = "ESTRATIFICACION POR MONTOS"
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).HorizontalAlignment = xlCenter
    I = I + 1
    lnIIni = I
    xlHoja1.Cells(I, 1) = "TOTAL"
    xlHoja1.Cells(I, 2) = lrReg!NumTot:         xlHoja1.Cells(I, 3) = lrReg!SKTot
    xlHoja1.Cells(I + 1, 1) = "Hasta 500"
    xlHoja1.Cells(I + 1, 2) = lrReg!NumRang1:     xlHoja1.Cells(I + 1, 3) = lrReg!SKRang1
    xlHoja1.Cells(I + 2, 1) = "De 501 A 1000"
    xlHoja1.Cells(I + 2, 2) = lrReg!NumRang2:     xlHoja1.Cells(I + 2, 3) = lrReg!SKRang2
    xlHoja1.Cells(I + 3, 1) = "De 1001 A 2000"
    xlHoja1.Cells(I + 3, 2) = lrReg!NumRang3:     xlHoja1.Cells(I + 3, 3) = lrReg!SKRang3
    xlHoja1.Cells(I + 4, 1) = "De 2001 A 5000"
    xlHoja1.Cells(I + 4, 2) = lrReg!NumRang4:     xlHoja1.Cells(I + 4, 3) = lrReg!SKRang4
    xlHoja1.Cells(I + 5, 1) = "De 5001 A 10000"
    xlHoja1.Cells(I + 5, 2) = lrReg!NumRang5:     xlHoja1.Cells(I + 5, 3) = lrReg!SKRang5
    xlHoja1.Cells(I + 6, 1) = "De 5001 A 10000"
    xlHoja1.Cells(I + 6, 2) = lrReg!NumRang6:     xlHoja1.Cells(I + 6, 3) = lrReg!SKRang6
    xlHoja1.Cells(I + 7, 1) = "De 10001 A 20000"
    xlHoja1.Cells(I + 7, 2) = lrReg!NumRang7:     xlHoja1.Cells(I + 7, 3) = lrReg!SKRang7
    
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I + 7, 3)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I + 7, 3)).Cells.Borders.LineStyle = xlInside

    I = I + 8
    xlHoja1.Cells(I, 1) = "ESTRATIFICACION POR PLAZOS"
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).HorizontalAlignment = xlCenter
    I = I + 1
    lnIIni = I
    xlHoja1.Cells(I, 1) = "TOTAL"
    xlHoja1.Cells(I, 2) = lrReg!NumTot:         xlHoja1.Cells(I, 3) = lrReg!SKTot
    xlHoja1.Cells(I + 1, 1) = "Hasta 3 Meses"
    xlHoja1.Cells(I + 1, 2) = lrReg!NumPlazo1:     xlHoja1.Cells(I + 1, 3) = lrReg!SKPlazo1
    xlHoja1.Cells(I + 2, 1) = "De 3 a 6 Meses"
    xlHoja1.Cells(I + 2, 2) = lrReg!NumPlazo2:     xlHoja1.Cells(I + 2, 3) = lrReg!SKPlazo2
    xlHoja1.Cells(I + 3, 1) = "De 6 a 12 Meses"
    xlHoja1.Cells(I + 3, 2) = lrReg!NumPlazo3:     xlHoja1.Cells(I + 3, 3) = lrReg!SKPlazo3
    xlHoja1.Cells(I + 4, 1) = "De 12 a 24 Meses"
    xlHoja1.Cells(I + 4, 2) = lrReg!NumPlazo4:     xlHoja1.Cells(I + 4, 3) = lrReg!SKPlazo4
    xlHoja1.Cells(I + 5, 1) = "Mas de 4 Meses"
    xlHoja1.Cells(I + 5, 2) = lrReg!NumPlazo5:     xlHoja1.Cells(I + 5, 3) = lrReg!SKPlazo5
    
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I + 5, 3)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I + 5, 3)).Cells.Borders.LineStyle = xlInside
    
    I = I + 6
    xlHoja1.Cells(I, 1) = "ESTRATIFICACION POR SECTORES"
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).HorizontalAlignment = xlCenter
    I = I + 1
    lnIIni = I
    xlHoja1.Cells(I, 1) = "TOTAL"
    xlHoja1.Cells(I, 2) = lrReg!NumTot:         xlHoja1.Cells(I, 3) = lrReg!SKTot
    xlHoja1.Cells(I + 1, 1) = "Servicio"
    xlHoja1.Cells(I + 1, 2) = lrReg!NumSect1:     xlHoja1.Cells(I + 1, 3) = lrReg!SKSect1
    xlHoja1.Cells(I + 2, 1) = "Comercio"
    xlHoja1.Cells(I + 2, 2) = lrReg!NumSect2:     xlHoja1.Cells(I + 2, 3) = lrReg!SKSect2
    xlHoja1.Cells(I + 3, 1) = "Produccion(No Agrop)"
    xlHoja1.Cells(I + 3, 2) = lrReg!NumSect3:     xlHoja1.Cells(I + 3, 3) = lrReg!SKSect3
    xlHoja1.Cells(I + 4, 1) = "Produccion Agropec."
    xlHoja1.Cells(I + 4, 2) = lrReg!NumSect4:     xlHoja1.Cells(I + 4, 3) = lrReg!SKSect4
    
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I + 4, 3)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I + 4, 3)).Cells.Borders.LineStyle = xlInside
    I = I + 5
    xlHoja1.Cells(I, 1) = "ESTRATIFICACION POR DESTINO"
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Font.Bold = True
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).Merge True
    xlHoja1.Range(xlHoja1.Cells(I, 1), xlHoja1.Cells(I, 3)).HorizontalAlignment = xlCenter
    I = I + 1
    lnIIni = I
    xlHoja1.Cells(I, 1) = "TOTAL"
    xlHoja1.Cells(I, 2) = lrReg!NumTot:         xlHoja1.Cells(I, 3) = lrReg!SKTot
    xlHoja1.Cells(I + 1, 1) = "Capital Trabajo"
    xlHoja1.Cells(I + 1, 2) = lrReg!NumDest1:     xlHoja1.Cells(I + 1, 3) = lrReg!SKDest1
    xlHoja1.Cells(I + 2, 1) = "Inversiones (Act.Fijo)"
    xlHoja1.Cells(I + 2, 2) = lrReg!NumDest2:     xlHoja1.Cells(I + 2, 3) = lrReg!SKDest2
    
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I + 2, 3)).Cells.Borders.LineStyle = xlOutside
    xlHoja1.Range(xlHoja1.Cells(lnIIni, 1), xlHoja1.Cells(I + 2, 3)).Cells.Borders.LineStyle = xlInside
    
    I = I + 3
    
    lrReg.Close
   
    oCon.CierraConexion
        
   MsgBox "Reporte Generado Satisfactoriamente"
End Sub
  
Private Sub cmdSalir_Click()
CierraConexion
Unload Me
End Sub


Private Sub Form_Load()
Dim rCargaRuta As New ADODB.Recordset
Dim rCargaAgencias As New ADODB.Recordset
Dim oCons As DConecta

    Set oCons = New DConecta
    oCons.AbreConexion
    Set rCargaRuta = oCons.CargaRecordSet("select nconssisvalor from constsistema where nconssiscod=" & gConstSistServCentralRiesgos)
    If rCargaRuta.BOF Then
    Else
        sservidorconsolidada = rCargaRuta!nConsSisValor
    End If
    Set rCargaRuta = Nothing
    
    
    cVigente = "'" & gColocEstVigNorm & "', '" & gColocEstVigVenc & "', '" & gColocEstVigMor & "', '" & gColocEstRefNorm & "', '" & gColocEstRefVenc & "', '" & gColocEstRefMor & "'"
    cPigno = "'" & gColPEstDesem & "', '" & gColPEstVenci & "', '" & gColPEstRemat & "', '" & gColPEstRenov & "'"
 
    Set rCargaAgencias = oCons.CargaRecordSet("Select cAgeCod,cAgeDescripcion from Agencias where nEstado=1")
    
   cmbAgencia.Clear
    Do While Not rCargaAgencias.EOF
        cmbAgencia.AddItem Trim(rCargaAgencias!cAgeDescripcion) & Space(100) & rCargaAgencias!cAgecod
        rCargaAgencias.MoveNext
    Loop
    Dim nPos As Integer
    
    nPos = cmbAgencia.ListIndex
    cmbAgencia.AddItem "Todos"
    
    cmbAgencia.ListIndex = -1
    cmbAgencia.ListIndex = nPos
    
   Set rCargaAgencias = Nothing
   oCons.CierraConexion
   
   CentraForm Me
   LlenaProductos
   oCons.AbreConexion
   'gcIntCentra = CentraSdi(Me)
   ' Habilita Controles
   Select Case fsCodReport
        Case gRiesgoCalfCarCred
            Call HabilitaControles(True, False, False)
        Case gRiesgoCalfAltoRiesgo
            Call HabilitaControles(True, False, False)
        Case gRiesgoConceCarCred
            Call HabilitaControles(False, True, False)
        Case gRiesgoEstratDepPlazo
            Call HabilitaControles(False, False, False)
        Case gRiesgoPrincipClientesAhorros
            Call HabilitaControles(False, True, True)
        Case gRiesgoPrincipClientesCreditos
            Call HabilitaControles(False, True, True)
        Case "816007"
            Call HabilitaControles(True, False, False)
        Case Else
            Call HabilitaControles(False, False, False)

   End Select
End Sub
 
Private Sub LlenaProductos()
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim sOpePadre As String
Dim sOpeHijo As String
Dim nodOpe As Node
Dim cSql As String

Set oCon = New DConecta
oCon.AbreConexion

Set rs = New ADODB.Recordset

If gbBitCentral = True Then
    cSql = " select nAgruCod as cGrupo, nAgruCab as cValor, cAgruDes as cProducto, 1 as cNivel from RepAgruProd where nAgruCod=1 " & _
           " Union " & _
           " select nAgruCod as cGrupo, cProdCod as cValor, cProdDesc as cProducto,2  as cNivel   from  RepAgruProdDet where nAgruCod=1 " & _
           " order by cvalor"
Else
    cSql = " Select cGrupo=nConsCod, cValor=nconsvalor, cProducto= cconsdescripcion, " & _
           " cNivel=case when nconsvalor in(select min(nconsvalor) from constante K where K.nconscod=C.nConscod AND substring(convert(varchar(3), K.nconsvalor),1,1) = substring(convert(varchar(3), C.nconsvalor),1,1)) " & _
           " Then 1 Else 2 End " & _
           " From constante C where C.nConsCod='1001' and nconsvalor not in(" & Producto.gCapAhorros & ", " & Producto.gCapPlazoFijo & ", " & Producto.gCapCTS & ", " & Producto.gColConsuPrendario & ") Order by nconsvalor"
End If



Set rs = oCon.CargaRecordSet(cSql)

TreeView1.Nodes.Clear

    Do While Not rs.EOF
        Select Case rs!cNivel
            Case "1"
                sOpePadre = "P" & rs!cValor
                Set nodOpe = TreeView1.Nodes.Add(, , sOpePadre, rs!cProducto, "Padre")
                nodOpe.Tag = rs!cValor
            Case "2"
                sOpeHijo = "H" & rs!cValor
                Set nodOpe = TreeView1.Nodes.Add(sOpePadre, tvwChild, sOpeHijo, rs!cProducto, "Hijo")
                nodOpe.Tag = rs!cValor
        
        End Select
        rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing

    oCon.CierraConexion

End Sub
Private Function GetProdsMarcados() As String
    Dim I As Integer
    Dim sCad As String
    sCad = ""
    For I = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(I).Checked = True Then
            If Mid(TreeView1.Nodes(I).Key, 1, 1) = "H" Then
                If Len(Trim(sCad)) = 0 Then
                    sCad = "'" & Mid(TreeView1.Nodes(I).Key, 2, 3)
                Else
                    sCad = sCad & "', '" & Mid(TreeView1.Nodes(I).Key, 2, 3)
                End If
            End If
        End If
    Next
    If Len(Trim(sCad)) > 0 Then
        sCad = "(" & sCad & "')"
    End If
                
    If Len(sCad) > 0 Then
        If gbBitCentral = True Then
            sCad = " AND substring(C.cCtaCod,6,3) IN " & sCad & " "
        Else
            sCad = " AND substring(C.cCodCta,3,3) IN " & sCad & " "
        End If
    Else
        sCad = ""
    End If
    
    GetProdsMarcados = sCad

End Function



Private Sub HabilitaControles(ByVal pbProductos As Boolean, ByVal pbTCambio As Boolean, ByVal pbNroClientes As Boolean)

fraProductos.Visible = pbProductos
fraProductos1.Visible = pbProductos
fraTCambio.Visible = pbTCambio
fraNroClientes.Visible = pbNroClientes

If fsCodReport = gRiesgoPrincipClientesAhorros Then
    chkuntitular.Visible = True
    lbluntitular.Visible = True
End If


If pbProductos = False Then
    fra1.Top = 120
    Me.Height = 2295
Else
    fra1.Top = 5265
    Me.Height = 6945
End If
End Sub
Private Function ProdSeleccionado() As String
Dim I As Integer
Dim lsCad As String

lsCad = ""
For I = 0 To chkCred1.Count - 1
    If chkCred1(I).value Then
        lsCad = lsCad & "'" & chkCred1(I).Tag & "',"
    End If
Next I

If Len(lsCad) > 0 Then
    lsCad = Mid(lsCad, 1, (Len(lsCad) - 1))
    ProdSeleccionado = " AND substring(C.cCodCta,3,3) IN (" & lsCad & ") "
Else
    ProdSeleccionado = ""
End If
        
End Function

Private Function ProdSeleccionadoDesc() As String
Dim lsProductos As String
Dim I As Integer
lsProductos = "PRODUCTOS : "
  For I = 0 To chkCred1.Count - 1
    If chkCred1(I).value Then
        If I < 3 Then
           lsProductos = lsProductos & "/MES-" & Mid(chkCred1(I).Caption, 1, 3)
        ElseIf I < 6 Then
            lsProductos = lsProductos & "/COM-" & Mid(chkCred1(I).Caption, 1, 3)
        Else
            lsProductos = lsProductos & "/CON-" & Mid(chkCred1(I).Caption, 1, 4)
        End If
    End If
  Next I
ProdSeleccionadoDesc = lsProductos
End Function


Private Function GetProdsMarcadosDesc() As String
    Dim I As Integer
    Dim sCad As String
    
    For I = 1 To TreeView1.Nodes.Count
        If TreeView1.Nodes(I).Checked = True Then
            If Mid(TreeView1.Nodes(I).Key, 1, 1) = "H" Then
                If Len(Trim(sCad)) = 0 Then
                    sCad = "Productos:" & Mid(TreeView1.Nodes(I).Text, 1, 15)
                Else
                    sCad = sCad & "/" & Mid(TreeView1.Nodes(I).Text, 1, 15)
                End If
            End If
        End If
    Next
    If Len(Trim(sCad)) > 0 Then
        sCad = "(" & sCad & "')"
    End If
                
    GetProdsMarcadosDesc = sCad

End Function

Private Sub TreeView1_NodeCheck(ByVal Node As MSComctlLib.Node)
Dim I As Integer
For I = 1 To TreeView1.Nodes.Count
    If Mid(TreeView1.Nodes(I).Key, 2, 1) = Mid(Node.Key, 2, 1) And Mid(Node.Key, 1, 1) = "P" Then
        TreeView1.Nodes(I).Checked = Node.Checked
    End If
Next
End Sub



Private Sub txtNroClientes_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       TxtTipoC.SetFocus
    End If
End Sub

Private Sub TxtTipoC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdProcesar.SetFocus
    End If
End Sub
