VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmLimitesFam 
   Caption         =   "Muestra Informacion de Limites Familiares"
   ClientHeight    =   7485
   ClientLeft      =   1200
   ClientTop       =   2670
   ClientWidth     =   12090
   Icon            =   "frmLimitesFamiliares.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7485
   ScaleWidth      =   12090
   Begin VB.OptionButton optsel 
      Caption         =   "Seleccionado"
      Height          =   255
      Index           =   1
      Left            =   7800
      TabIndex        =   17
      Top             =   7125
      Width           =   1320
   End
   Begin VB.OptionButton optsel 
      Caption         =   "Todos"
      Height          =   255
      Index           =   0
      Left            =   7785
      TabIndex        =   16
      Top             =   6855
      Value           =   -1  'True
      Width           =   1185
   End
   Begin VB.Frame Frame1 
      Height          =   690
      Left            =   165
      TabIndex        =   7
      Top             =   6135
      Width           =   11820
      Begin VB.Label lblTotalTC 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   345
         Left            =   9840
         TabIndex        =   13
         Top             =   195
         Width           =   1770
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Total al TC:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   240
         Left            =   8235
         TabIndex        =   12
         Top             =   255
         Width           =   1230
      End
      Begin VB.Label lblTotalME 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   345
         Left            =   5475
         TabIndex        =   11
         Top             =   195
         Width           =   1770
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Total Dolares:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00008000&
         Height          =   240
         Left            =   3870
         TabIndex        =   10
         Top             =   255
         Width           =   1500
      End
      Begin VB.Label lblTotSoles 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   1935
         TabIndex        =   9
         Top             =   195
         Width           =   1800
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Total Soles:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   240
         Left            =   375
         TabIndex        =   8
         Top             =   240
         Width           =   1275
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   10530
      TabIndex        =   6
      Top             =   6900
      Width           =   1410
   End
   Begin VB.CommandButton cmdExcel 
      Caption         =   "Migrar a Excel"
      Height          =   390
      Left            =   9135
      TabIndex        =   5
      Top             =   6900
      Width           =   1410
   End
   Begin MSComCtl2.DTPicker txtFecha 
      Height          =   360
      Left            =   1620
      TabIndex        =   3
      Top             =   165
      Width           =   1590
      _ExtentX        =   2805
      _ExtentY        =   635
      _Version        =   393216
      Format          =   83951617
      CurrentDate     =   38427
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   5415
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   11835
      _ExtentX        =   20876
      _ExtentY        =   9551
      _Version        =   393216
      BackColor       =   16777215
      RowHeightMin    =   350
      ForeColorFixed  =   8421376
      WordWrap        =   -1  'True
      GridLinesUnpopulated=   1
      AllowUserResizing=   1
      BandDisplay     =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   375
      Left            =   10185
      TabIndex        =   0
      Top             =   105
      Width           =   1710
   End
   Begin VB.TextBox txtTC 
      Alignment       =   1  'Right Justify
      Height          =   315
      Left            =   9015
      TabIndex        =   14
      Text            =   "0.00"
      Top             =   150
      Width           =   840
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo de Cambio:"
      Height          =   195
      Left            =   7755
      TabIndex        =   15
      Top             =   210
      Width           =   1155
   End
   Begin VB.Label lblMensaje 
      AutoSize        =   -1  'True
      Caption         =   "Informacion al:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   240
      Left            =   225
      TabIndex        =   4
      Top             =   7020
      Width           =   1530
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informacion al:"
      Height          =   195
      Left            =   465
      TabIndex        =   2
      Top             =   240
      Width           =   1035
   End
End
Attribute VB_Name = "frmLimitesFam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lsTablaTMP As String

Private Function CargaDatos() As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Set oCon = New DConecta
Dim lsSQL As String

lsSQL = _
    " SHAPE {SELECT  CGRUFAM AS GrupoFam, CPERSCODTRAB, CPERSNOMTRAB as Trabajador, " _
    & "                 SUM(NSALDOMN) AS TotalSaldoMN, SUM(NSALDOME) AS TotalSaldoME, SUM(NSALDOTC) AS TotalSaldoTC " _
    & "         From " & lsTablaTMP _
    & "         GROUP BY CGRUFAM, CPERSCODTRAB, CPERSNOMTRAB} " _
    & " APPEND ({SELECT CGRUFAM AS GrupoFam, CPERSCOD , CPERSNOMBRE AS FAMILIAR, CONVERT(decimal(10,2), NSALDOMN)  AS SALDOMN, CONVERT(decimal(10,2), NSALDOME) AS SALDOME, CONVERT(decimal(10,2), NSALDOTC) AS SALDOTC, CCTACOD AS [Nro Credito], CMONEDA as Moneda, CRELACION AS [Rel.Fam], DVIGENCIA  AS Vigencia, NDIASATRASO AS Atraso, nCalif as Calificacion, CPERSCODTRAB " _
    & "          FROM " & lsTablaTMP & " ORDER BY CPERSNOMBRE} AS DET RELATE CPERSCODTRAB TO CPERSCODTRAB)"

    oCon.AbreConexion
    
    Set MSHFlexGrid1.DataSource = Nothing
    
    Set rs = New ADODB.Recordset
    rs.Open lsSQL, Replace(oCon.CadenaConexion, "Provider=SQLOLEDB.1", "PROVIDER=MSDataShape;DATA PROVIDER=SQLOLEDB")
    
    MSHFlexGrid1.Clear
    MSHFlexGrid1.ClearStructure
    MSHFlexGrid1.Cols = 2
    MSHFlexGrid1.Rows = 2

    Set MSHFlexGrid1.DataSource = rs
    
    If Not rs.EOF And Not rs.BOF Then
        Call FormateaGrid
    End If
    oCon.CierraConexion
End Function
Sub FormateaGrid()
Dim s As String
    
    MSHFlexGrid1.BandDisplay = flexBandDisplayVertical
    MSHFlexGrid1.Redraw = False
    
    s$ = "|^Grupo|<cPersCod|<Trabajador|>SaldoMN|>SaldoME|>Total Saldo TC "
    MSHFlexGrid1.FillStyle = flexFillRepeat
    MSHFlexGrid1.FormatString = s$
    MSHFlexGrid1.ColAlignmentHeader(4) = flexAlignLeftCenter
    'MSHFlexGrid1.ForeColorFixed = vbRed
    MSHFlexGrid1.FontFixed.Bold = True
    MSHFlexGrid1.FontFixed.Size = 10
    
    MSHFlexGrid1.ColWidth(0) = 300
    MSHFlexGrid1.ColWidth(1) = 1000
    MSHFlexGrid1.ColWidth(2) = 0
    MSHFlexGrid1.ColWidth(3) = 4050
    MSHFlexGrid1.ColWidth(4) = 1500
    MSHFlexGrid1.ColWidth(5) = 1500
    MSHFlexGrid1.ColWidth(6) = 1800
    
    'empieza el detalle
    MSHFlexGrid1.ColWidth(0, 1) = 1000
    MSHFlexGrid1.ColWidth(1, 1) = 0
    MSHFlexGrid1.ColWidth(2, 1) = 4050
    MSHFlexGrid1.ColWidth(3, 1) = 1200
    MSHFlexGrid1.ColWidth(4, 1) = 1500
    MSHFlexGrid1.ColWidth(5, 1) = 1500
    MSHFlexGrid1.ColWidth(6, 1) = 1800 'nrocredito
    MSHFlexGrid1.ColWidth(7, 1) = 1200 'moneda
    MSHFlexGrid1.ColWidth(8, 1) = 1500 'relacio
    MSHFlexGrid1.ColWidth(9, 1) = 1500 'vigencia
    MSHFlexGrid1.ColWidth(10, 1) = 1500 'atraso
    MSHFlexGrid1.ColWidth(11, 1) = 1500 'calificacion
    MSHFlexGrid1.ColWidth(12, 1) = 0
    
    MSHFlexGrid1.ColAlignmentFixed(4) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentFixed(5) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentFixed(6) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentFixed(7) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentFixed(8) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentFixed(9) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentFixed(10) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentFixed(11) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentFixed(12) = flexAlignCenterCenter
    
    MSHFlexGrid1.ColAlignmentHeader(6, 1) = flexAlignLeftCenter
    MSHFlexGrid1.ColAlignmentHeader(6, 2) = flexAlignRightCenter
    MSHFlexGrid1.ColAlignmentHeader(7, 1) = flexAlignRightCenter
    MSHFlexGrid1.ColAlignmentHeader(7, 2) = flexAlignRightCenter
    MSHFlexGrid1.ColAlignmentHeader(8, 1) = flexAlignRightCenter
    MSHFlexGrid1.ColAlignmentHeader(8, 2) = flexAlignRightCenter
    MSHFlexGrid1.ColAlignmentHeader(10, 1) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentHeader(10, 2) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentHeader(12, 1) = flexAlignCenterCenter
    MSHFlexGrid1.ColAlignmentHeader(12, 2) = flexAlignCenterCenter
    
    MSHFlexGrid1.CollapseAll
    
    Dim i As Long
    Dim lnTotalMN As Currency
    Dim lnTotalME As Currency
    Dim lnTotalTC As Currency
    
    lnTotalTC = 0
    lnTotalMN = 0
    lnTotalME = 0
    For i = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.Col = 1
        MSHFlexGrid1.Row = i
        MSHFlexGrid1.CellForeColor = vbBlue
        MSHFlexGrid1.Text = Format(MSHFlexGrid1.Text, "##,#0.00")
        
        MSHFlexGrid1.Col = 3
        MSHFlexGrid1.Row = i
        MSHFlexGrid1.CellForeColor = vbBlue
        MSHFlexGrid1.Text = Format(MSHFlexGrid1.Text, "##,#0.00")
        
        MSHFlexGrid1.Col = 4
        MSHFlexGrid1.Row = i
        MSHFlexGrid1.CellAlignment = flexAlignRightCenter
        MSHFlexGrid1.CellForeColor = vbBlue
        MSHFlexGrid1.Text = Format(MSHFlexGrid1.Text, "##,#0.00")
        lnTotalMN = lnTotalMN + CCur(MSHFlexGrid1.Text)
        
        MSHFlexGrid1.Col = 5
        MSHFlexGrid1.Row = i
        MSHFlexGrid1.CellForeColor = vbBlue
        MSHFlexGrid1.CellAlignment = flexAlignRightCenter
        MSHFlexGrid1.Text = Format(MSHFlexGrid1.Text, "##,#0.00")
        lnTotalME = lnTotalME + CCur(MSHFlexGrid1.Text)
        
        MSHFlexGrid1.Col = 6
        MSHFlexGrid1.Row = i
        MSHFlexGrid1.CellFontBold = True
        MSHFlexGrid1.CellForeColor = vbBlue
        MSHFlexGrid1.CellBackColor = &H80FFFF
        MSHFlexGrid1.CellAlignment = flexAlignRightCenter
        MSHFlexGrid1.Text = Format(MSHFlexGrid1.Text, "##,#0.00")
        lnTotalTC = lnTotalTC + CCur(MSHFlexGrid1.Text)
    
    Next
    MSHFlexGrid1.Redraw = True
    
    lblTotalME = Format(lnTotalME, "##,#0.00")
    lblTotSoles = Format(lnTotalMN, "##,#0.00")
    lblTotalTC = Format(lnTotalTC, "##,#0.00")
End Sub

Private Sub cmdExcel_Click()
GenerareporteExcel
End Sub

Private Sub cmdProcesar_Click()
ProcesaInfoLimites txtFecha
CargaDatos
Me.lblMensaje = "Informacion procesada al :" & txtFecha
Me.lblMensaje.AutoSize = True
Me.lblMensaje.Refresh
End Sub

Private Sub cmdSalir_Click()
VerificaTablaTemporal
Unload Me
End Sub

Private Sub Command1_Click()
frmEstadVenMen.Show 1
End Sub

Private Sub Form_Load()
CentraForm Me
Me.txtFecha = gdFecSis
GetTipCambio (gdFecSis)
Me.txtTC = gnTipCambio
End Sub
Sub VerificaTablaTemporal()
Dim rs As ADODB.Recordset
Dim sql As String
Dim oCon As DConecta

lsTablaTMP = "TMPLIMITE" & gsCodUser

Set oCon = New DConecta
oCon.AbreConexion

Set rs = oCon.CargaRecordSet("select * from sysobjects where name like '%" & lsTablaTMP & "%'")
If Not rs.EOF And Not rs.BOF Then
    sql = "DROP TABLE " & lsTablaTMP
    oCon.Ejecutar sql
End If
rs.Close
Set rs = Nothing

End Sub
Sub ProcesaInfoLimites(ByVal pdFecha As Date)
Dim sql As String
Dim oCon As DConecta
Dim lsFiltroFech As String
Dim lsCampoSaldo As String
Dim lsTabla As String
Dim lnTC As Currency

lnTC = txtTC

If pdFecha = gdFecSis Then
    lsFiltroFech = ""
    lsCampoSaldo = " NSALDO "
    lsTabla = " PRODUCTO "
Else
    lsFiltroFech = " and convert(char(10),C.dFecha,112) ='" & Format(pdFecha, "yyyymmdd") & "' "
    lsCampoSaldo = " NSALDOCAP "
    lsTabla = " ColocacSaldo "
End If
Set oCon = New DConecta

VerificaTablaTemporal

sql = "SELECT   CTRAB.CGRUFAM, CTRAB.cPersCodTrab,CTRAB.cPersNomTrab, SAL.cPersCod, SAL.cPersNombre, SAL.CCTACOD, "
sql = sql + "   SAL.CMONEDA, 'TRABAJADOR' AS CRELACION, convert(varchar(10),SAL.dVigencia,103) AS dVigencia, SAL.nDiasAtraso,"
sql = sql + "   SAL.NSALDOMN, SAL.NSALDOME,"
sql = sql + "   ROUND(CASE WHEN SAL.CMONEDA='1' THEN SAL.NSALDOMN ELSE ROUND(SAL.NSALDOME*" & lnTC & ",2) END,2) NSALDOTC,"
sql = sql + "   SAL.nCalif , SAL.nProvi"
sql = sql + "   Into " & lsTablaTMP
sql = sql + "   From    "
sql = sql + "       (SELECT P.cPersCod, P.cPersNombre, C.CCTACOD, SUBSTRING(C.CCTACOD,9,1) AS CMONEDA,"
sql = sql + "               CASE WHEN SUBSTRING(C.CCTACOD,9,1) ='1' THEN " & lsCampoSaldo & " ELSE 0 END AS NSALDOMN,"
sql = sql + "               CASE WHEN SUBSTRING(C.CCTACOD,9,1) ='2' THEN " & lsCampoSaldo & " ELSE 0 END AS NSALDOME,"
sql = sql + "               dVigencia = (SELECT dVigencia FROM Colocaciones where cCtaCod = C.cCtaCod ),"
sql = sql + "               nDiasAtraso = (SELECT isnull(nDiasAtraso,0) from ColocacCred where cCtaCod = C.cCtaCod), "
sql = sql + "               nCalif = (Select isnull(cCalGen,'') FROM ColocCalifProv where cCtaCod = C.cCtaCod),"
sql = sql + "               nProvi = (Select isnull(nProvision,0) FROM ColocCalifProv where cCtaCod = C.cCtaCod)"
sql = sql + "        FROM   " & lsTabla & " C"
sql = sql + "               JOIN ProductoPersona R on R.cCtaCod = C.cCtaCod and R.nPrdPersRelac= 20"
sql = sql + "               JOIN Persona P ON P.cPersCod = R.cPersCod"
sql = sql + "       WHERE   nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205)  "
sql = sql + "                " & lsFiltroFech
sql = sql + "               AND C.cCtaCod NOT IN (SELECT CCTACOD FROM COLOCACCRED WHERE CRFA ='RFA')) AS SAL"
sql = sql + "         Join"
sql = sql + "         ( select PR.cPersCod AS cPersCodTrab, P.CPERSNOMBRE AS cPersNomTrab, PR.cGruFam"
sql = sql + "           from    PersRelaciones PR"
sql = sql + "                   JOIN RRHH RH ON RH.cPersCod = PR.cPersCod  and not nRHEstado like '[867]%'"
sql = sql + "                   JOIN PERSONA P ON P.CPERSCOD = PR.CPERSCOD"
sql = sql + "           Where cGruFam Is Not Null"
sql = sql + "           GROUP BY PR.cPersCod, P.CPERSNOMBRE, PR.cGruFam) AS CTRAB ON CTRAB.cPersCodTrab = SAL.CPERSCOD"
sql = sql + "           Union All"
sql = sql + "           SELECT  CTRAB.CGRUFAM, CTRAB.cPersCodTrab,CTRAB.cPersNomTrab, SAL.cPersCod, SAL.cPersNombre, SAL.CCTACOD,"
sql = sql + "                   SAL.CMONEDA, CTRAB.CRELACION,convert(varchar(10),SAL.dVigencia,103) as dVigencia, SAL.nDiasAtraso,"
sql = sql + "                   SAL.NSALDOMN, SAL.NSALDOME,"
sql = sql + "                   ROUND(CASE WHEN SAL.CMONEDA='1' THEN SAL.NSALDOMN ELSE ROUND(SAL.NSALDOME*" & lnTC & ",2) END,2) NSALDOTC,"
sql = sql + "                   SAL.nCalif , SAL.nProvi"
sql = sql + "            From"
sql = sql + "               (SELECT     P.cPersCod, P.cPersNombre, C.CCTACOD, SUBSTRING(C.CCTACOD,9,1) AS CMONEDA,"
sql = sql + "                           CASE WHEN SUBSTRING(C.CCTACOD,9,1) ='1' THEN " & lsCampoSaldo & " ELSE 0 END AS NSALDOMN,"
sql = sql + "                           CASE WHEN SUBSTRING(C.CCTACOD,9,1) ='2' THEN " & lsCampoSaldo & " ELSE 0 END AS NSALDOME,"
sql = sql + "                           dVigencia = (SELECT dVigencia FROM Colocaciones where cCtaCod = C.cCtaCod ),"
sql = sql + "                           nDiasAtraso = (SELECT isnull(nDiasAtraso,0) from ColocacCred where cCtaCod = C.cCtaCod),"
sql = sql + "                           nCalif = (Select isnull(cCalGen,'') FROM ColocCalifProv where cCtaCod = C.cCtaCod),"
sql = sql + "                           nProvi = (Select isnull(nProvision,0) FROM ColocCalifProv where cCtaCod = C.cCtaCod)"
sql = sql + "                 FROM      " & lsTabla & " C"
sql = sql + "                           JOIN ProductoPersona R on R.cCtaCod = C.cCtaCod and R.nPrdPersRelac= 20"
sql = sql + "                           JOIN Persona P ON P.cPersCod = R.cPersCod"
sql = sql + "                 WHERE     nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205) " & lsFiltroFech
sql = sql + "                           " & lsFiltroFech
sql = sql + "                           AND C.cCtaCod NOT IN (SELECT CCTACOD FROM COLOCACCRED WHERE CRFA ='RFA')) AS SAL"
sql = sql + "                  Join"
sql = sql + "                  (select PR.cPersCod as cPersCodTrab, P.cPersNombre as cPersNomTrab,"
sql = sql + "                           PR.cPersRelacPersCod, PR.nPersRelac, C.cConsDescripcion AS CRELACION, PR.cGruFam"
sql = sql + "                   from    PersRelaciones PR"
sql = sql + "                           JOIN RRHH RH ON RH.cPersCod = PR.cPersCod  and not nRHEstado like '[867]%'"
sql = sql + "                           JOIN CONSTANTE C ON C.nConsValor = PR.nPersRelac AND C.nConsCod=1006"
sql = sql + "                           JOIN PERSONA P ON P.cPersCod = PR.cPersCod"
sql = sql + "                    Where cGruFam Is Not Null"
sql = sql + "                           GROUP BY PR.cPersCod,P.cPersNombre, PR.cPersRelacPersCod, PR.nPersRelac,"
sql = sql + "                           C.cConsDescripcion, PR.cGruFam ) AS CTRAB ON CTRAB.cPersRelacPersCod = SAL.CPERSCOD"
sql = sql + "                           ORDER BY cGruFam, cPersCodTrab"

Me.lblMensaje = "Procesando Informacion de Limites...."
Me.lblMensaje.Refresh

oCon.AbreConexion
oCon.Ejecutar sql
oCon.CierraConexion

Set oCon = Nothing
Me.lblMensaje = "Procesando culminado...."
Me.lblMensaje.Refresh
End Sub

Sub GenerareporteExcel()
Dim vExcelObj As Excel.Application
Dim vNHC As String
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim lsGrupo As String
Set oCon = New DConecta

If MSHFlexGrid1.Rows < 3 Then
    Exit Sub
End If

If Me.optsel(0).value = True Then
    sql = "SELECT * FROM " & lsTablaTMP
Else
    Me.MSHFlexGrid1.Col = 1
    MSHFlexGrid1.Row = MSHFlexGrid1.Row
    lsGrupo = MSHFlexGrid1.Text

    sql = "SELECT * FROM " & lsTablaTMP & " where CGRUFAM='" & lsGrupo & "'"
End If

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(sql)
oCon.CierraConexion

If Not rs.EOF And Not rs.BOF Then

vNHC = App.path & "\spooler\LIMITES_AL" & Format(txtFecha, "yyyymmdd") & ".XLS"

Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
vExcelObj.DisplayAlerts = False

vExcelObj.Workbooks.Add
vExcelObj.Sheets("Hoja1").Select
vExcelObj.Sheets("Hoja1").name = "LIMITES"

vExcelObj.Range("A1:IV65536").Font.name = "Arial Narrow"
vExcelObj.Range("A1:IV65536").Font.Size = 8
vExcelObj.Columns("A:IV").Select
vExcelObj.Selection.VerticalAlignment = 3

vExcelObj.Columns("A").Select
vExcelObj.Selection.HorizontalAlignment = 1
vExcelObj.Columns("B:H").Select
vExcelObj.Selection.HorizontalAlignment = 1

vExcelObj.Range("A1").Select
vExcelObj.Range("A1").Font.Bold = True
vExcelObj.Range("A1").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = UCase(Trim(gsNomCmac))

vExcelObj.Range("M1").Select
vExcelObj.Range("M1").Font.Bold = True
vExcelObj.Range("M1").HorizontalAlignment = 1
vExcelObj.ActiveCell.value = "Informacion al:" & Format(Me.txtFecha, "dd/mm/yyyy")

vExcelObj.Range("E4").Select
vExcelObj.Range("E4").Font.Bold = True
vExcelObj.Range("E4").HorizontalAlignment = 1
vExcelObj.Range("E4").Font.Size = 12
vExcelObj.ActiveCell.value = "REPORTE DE CONTROL DE LIMITES DE TRABAJADORES DIRECTORES"

vExcelObj.Range("A3").Select
vExcelObj.Range("A3").Font.Bold = True
vExcelObj.Range("A3").HorizontalAlignment = xlLeft
vExcelObj.ActiveCell.value = "Tipo de Cambio : " & txtTC.Text

'vExcelObj.Range("A6:N6").AutoFilter

vExcelObj.Range("A6").Select
vExcelObj.Range("A6").Font.Bold = True
vExcelObj.Range("A6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "GRUPO"

vExcelObj.Range("B6").Select
vExcelObj.Range("B6").Font.Bold = True
vExcelObj.Range("B6").ColumnWidth = 20
vExcelObj.ActiveCell.value = "CODIGO TRAB."

vExcelObj.Range("C6").Select
vExcelObj.Range("C6").Font.Bold = True
vExcelObj.Range("C6").ColumnWidth = 30
vExcelObj.ActiveCell.value = "TRABAJADOR"

vExcelObj.Range("D6").Select
vExcelObj.Range("D6").Font.Bold = True
vExcelObj.Range("D6").ColumnWidth = 15
vExcelObj.ActiveCell.value = "COD.FAMILIAR"

vExcelObj.Range("E6").Select
vExcelObj.Range("E6").Font.Bold = True
vExcelObj.Range("E6").ColumnWidth = 50
vExcelObj.ActiveCell.value = "NOMBRE DEL FAMILIAR/TRABAJADOR"
                     
vExcelObj.Range("F6").Select
vExcelObj.Range("F6").Font.Bold = True
vExcelObj.Range("F6").ColumnWidth = 15
vExcelObj.ActiveCell.value = "CCODCTA"
                        
vExcelObj.Range("G6").Select
vExcelObj.Range("G6").Font.Bold = True
vExcelObj.Range("G6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "MONEDA"

vExcelObj.Range("H6").Select
vExcelObj.Range("H6").Font.Bold = True
vExcelObj.Range("H6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "RELACION"

vExcelObj.Range("I6").Select
vExcelObj.Range("I6").Font.Bold = True
vExcelObj.Range("I6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "VIGENCIA"

vExcelObj.Range("J6").Select
vExcelObj.Range("J6").Font.Bold = True
vExcelObj.Range("J6").ColumnWidth = 15
vExcelObj.ActiveCell.value = "ATRASO"

vExcelObj.Range("K6").Select
vExcelObj.Range("K6").Font.Bold = True
vExcelObj.Range("K6").ColumnWidth = 15
vExcelObj.ActiveCell.value = "SALDO MN"

vExcelObj.Range("L6").Select
vExcelObj.Range("L6").Font.Bold = True
vExcelObj.Range("L6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "SALDO MN"

vExcelObj.Range("L6").Select
vExcelObj.Range("L6").Font.Bold = True
vExcelObj.Range("L6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "SALDO ME"

vExcelObj.Range("M6").Select
vExcelObj.Range("M6").Font.Bold = True
vExcelObj.Range("M6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "SALDO TC"

vExcelObj.Range("N6").Select
vExcelObj.Range("N6").Font.Bold = True
vExcelObj.Range("N6").ColumnWidth = 10
vExcelObj.ActiveCell.value = "Calificacion"

lnTotaCapSICMAC = 0
lnTotaCapSIAFC = 0

lnTotIntSICMAC = 0
lnTotIntSIAFC = 0

lnTotaMoraSICMAC = 0
lnTotaMoraSIAFC = 0

lnTotaITFSICMAC = 0
lnTotaITFSIAFC = 0

rs.MoveFirst
Dim lsCodCta As String
vIni = 6
vItem = vIni
lnTotalCtasCMAC = 0
lblMensaje = "Migrando Informacion a hoja de Calculo por favor Espere..."
Do While Not rs.EOF
         vItem = vItem + 1
    
         vCel = "A" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
         vExcelObj.ActiveCell.value = "'" + rs!CGRUFAM
    
         vCel = "B" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!cPersCodTrab
    
         vCel = "C" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!cPersNomTrab
    
         vCel = "D" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!cPersCod
    
         vCel = "E" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!cPersNombre
    
         vCel = "F" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" & rs!cCtaCod
    
         vCel = "G" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
         vExcelObj.ActiveCell.value = "'" + rs!cmoneda
         
         vCel = "H" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!CRELACION
    
         vCel = "I" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!dVigencia
         
         vCel = "J" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
         vExcelObj.ActiveCell.value = rs!nDiasAtraso
         
         vCel = "K" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!NSALDOMN, "#,#0.00")
         
         vCel = "L" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!NSALDOME, "#,#0.00")
         
         vCel = "M" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!NSALDOTC, "#,#0.00")
         
         vCel = "N" + Trim(Str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
         vExcelObj.ActiveCell.value = "'" + rs!nCalif
         
    rs.MoveNext
Loop
vExcelObj.Range("A6:N6").Activate
vExcelObj.Range("A6").Subtotal 1, xlSum, Array(11, 12, 13), True

'If Dir(vNHC) <> "" Then
'   If MsgBox("Archivo Ya Existe ...  Desea Reemplazarlo ??", vbQuestion + vbYesNo + vbDefaultButton1, " Mensaje del Sistema ...") = vbNo Then
'      Exit Function
'   End If
'End If
vExcelObj.Range("A1").Select
vExcelObj.ActiveWorkbook.SaveAs (vNHC)
vExcelObj.ActiveWorkbook.Close
vExcelObj.Workbooks.Open (vNHC)
vExcelObj.Visible = True

Me.lblMensaje = "Proceso Culminado satisfactoriamente"

Set vExcelObj = Nothing
MsgBox "SE HA GENERADO CON ÉXITO EL ARCHIVO !!  ", vbInformation, " Mensaje del Sistema ..."
End If

rs.Close
Set rs = Nothing
Me.lblMensaje = "Informacion procesada al :" & txtFecha
End Sub

Private Sub MSHFlexGrid1_Collapse(Cancel As Boolean)
Dim i As Long
For i = 1 To MSHFlexGrid1.Cols - 1
    MSHFlexGrid1.Col = i
    MSHFlexGrid1.Row = MSHFlexGrid1.Row
    MSHFlexGrid1.CellForeColor = vbBlue
    MSHFlexGrid1.CellFontItalic = False
    If i <> 6 Then
        MSHFlexGrid1.CellFontBold = False
    End If
Next
End Sub

Private Sub MSHFlexGrid1_Expand(Cancel As Boolean)
Dim i As Long
For i = 1 To MSHFlexGrid1.Cols - 1
    MSHFlexGrid1.Col = i
    MSHFlexGrid1.Row = MSHFlexGrid1.Row
    MSHFlexGrid1.CellForeColor = vbBlue
    MSHFlexGrid1.CellFontItalic = True
    MSHFlexGrid1.CellFontBold = True
Next

End Sub

Private Sub txtTC_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTC, KeyAscii, 12, 4)
If KeyAscii = 13 Then
    Me.cmdProcesar.SetFocus
End If
End Sub
