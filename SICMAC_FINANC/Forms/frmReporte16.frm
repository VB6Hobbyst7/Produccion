VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmReporte16 
   Caption         =   "Reporte N° 16"
   ClientHeight    =   3045
   ClientLeft      =   5340
   ClientTop       =   2340
   ClientWidth     =   3075
   Icon            =   "frmReporte16.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   3045
   ScaleWidth      =   3075
   Begin MSComCtl2.DTPicker txtFecha 
      Height          =   330
      Left            =   960
      TabIndex        =   14
      Top             =   135
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   582
      _Version        =   393216
      Format          =   58851329
      CurrentDate     =   38411
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   1560
      TabIndex        =   12
      Top             =   2550
      Width           =   1290
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   390
      Left            =   285
      TabIndex        =   11
      Top             =   2550
      Width           =   1290
   End
   Begin VB.Frame Frame1 
      Height          =   1860
      Left            =   240
      TabIndex        =   0
      Top             =   600
      Width           =   2610
      Begin VB.OptionButton optRep16 
         Caption         =   "Anexo K2"
         Height          =   300
         Index           =   9
         Left            =   1395
         TabIndex        =   10
         Top             =   1380
         Width           =   1065
      End
      Begin VB.OptionButton optRep16 
         Caption         =   "Anexo K1"
         Height          =   300
         Index           =   8
         Left            =   150
         TabIndex        =   9
         Top             =   1395
         Width           =   1380
      End
      Begin VB.OptionButton optRep16 
         Caption         =   "Anexo G2"
         Height          =   300
         Index           =   7
         Left            =   1395
         TabIndex        =   8
         Top             =   1095
         Width           =   1095
      End
      Begin VB.OptionButton optRep16 
         Caption         =   "Anexo G1"
         Height          =   300
         Index           =   6
         Left            =   150
         TabIndex        =   7
         Top             =   1095
         Width           =   1380
      End
      Begin VB.OptionButton optRep16 
         Caption         =   "Anexo F2"
         Height          =   300
         Index           =   5
         Left            =   1395
         TabIndex        =   6
         Top             =   810
         Width           =   1125
      End
      Begin VB.OptionButton optRep16 
         Caption         =   "Anexo F1"
         Height          =   300
         Index           =   4
         Left            =   150
         TabIndex        =   5
         Top             =   795
         Width           =   1380
      End
      Begin VB.OptionButton optRep16 
         Caption         =   "Anexo E2"
         Height          =   300
         Index           =   3
         Left            =   1410
         TabIndex        =   4
         Top             =   495
         Width           =   1080
      End
      Begin VB.OptionButton optRep16 
         Caption         =   "Anexo E1"
         Height          =   300
         Index           =   2
         Left            =   150
         TabIndex        =   3
         Top             =   465
         Width           =   1380
      End
      Begin VB.OptionButton optRep16 
         Caption         =   "Anexo D2"
         Height          =   300
         Index           =   1
         Left            =   1395
         TabIndex        =   2
         Top             =   195
         Width           =   1125
      End
      Begin VB.OptionButton optRep16 
         Caption         =   "Anexo D1"
         Height          =   300
         Index           =   0
         Left            =   150
         TabIndex        =   1
         Top             =   180
         Width           =   1380
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha:"
      Height          =   195
      Left            =   360
      TabIndex        =   13
      Top             =   180
      Width           =   495
   End
End
Attribute VB_Name = "frmReporte16"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsRep16 As ADODB.Recordset

Private Sub cmdProcesar_Click()
Dim aProd() As String

If optRep16(0).value = True Then 'anexo D1
    ReDim Preserve aProd(2)
    aProd(0) = "304"
    aProd(1) = "320"
    DatosReporte16 1, 5000, 10000, 12, 12, aProd, "D1"
End If

If optRep16(1).value = True Then 'anexo D2
    ReDim Preserve aProd(2)
    aProd(0) = "304"
    aProd(1) = "320"
    DatosReporte16 2, 1500, 4000, 12, 12, aProd, "D2"
End If
If optRep16(2).value = True Then 'anexo E1
    ReDim Preserve aProd(4)
    aProd(0) = "304"
    aProd(1) = "320"
    aProd(2) = "302"
    aProd(3) = "303"
    
    DatosReporte16 1, 500, 1000, 12, 12, aProd, "E1"
End If
If optRep16(3).value = True Then 'anexo E2
    ReDim Preserve aProd(4)
    aProd(0) = "304"
    aProd(1) = "320"
    aProd(2) = "302"
    aProd(3) = "303"
    
    DatosReporte16 2, 150, 300, 12, 12, aProd, "E2"
End If
If optRep16(4).value = True Then 'anexo f1
    ReDim Preserve aProd(2)
    aProd(0) = "201"
    aProd(1) = "202"
    
    DatosReporte16 1, 15000, 30000, 9, 24, aProd, "F1"
End If
If optRep16(5).value = True Then 'anexo f2
    ReDim Preserve aProd(2)
    aProd(0) = "201"
    aProd(1) = "202"
    
    DatosReporte16 2, 3000, 6000, 9, 24, aProd, "F2"
End If
If optRep16(6).value = True Then 'anexo g1
    ReDim Preserve aProd(2)
    aProd(0) = "201"
    aProd(1) = "202"
    
    DatosReporte16 1, 2000, 5000, 9, 24, aProd, "G1"
End If
If optRep16(7).value = True Then 'anexo g2
    ReDim Preserve aProd(2)
    aProd(0) = "201"
    aProd(1) = "202"
    
    DatosReporte16 2, 500, 1000, 9, 24, aProd, "G2"
End If
If optRep16(8).value = True Then 'anexo k1
    ReDim Preserve aProd(1)
    aProd(0) = "423"
    DatosReporte16 1, 54000, 76500, 15 * 12, 15 * 12, aProd, "K1"
End If
If optRep16(9).value = True Then 'anexo k2
    ReDim Preserve aProd(1)
    aProd(0) = "423"
    DatosReporte16 2, 18000, 22500, 15 * 12, 15 * 12, aProd, "K2"
End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Me.txtFecha = gdFecSis
End Sub
Function InteresReal(ByVal pnTasInt As Double, ByVal pnDias As Long, ByVal pnPeriodo As Long) As Double
InteresReal = 0
InteresReal = ((1 + pnTasInt / 100) ^ (pnDias / pnPeriodo)) - 1
End Function
Function cuotaconstante(ByVal pnTasInt As Double, ByVal pnNroCuo As Long, _
        ByVal pnMontoReq As Currency, ByVal pnDias As Long, ByVal pnPeriodo As Long) As Currency

Dim nPotCuo As Double
Dim pnTasIntCal As Double
nPotCuo = 0
pnTasIntCal = InteresReal(pnTasInt, pnDias, pnPeriodo)
nPotCuo = (1 + pnTasIntCal) ^ pnNroCuo
cuotaconstante = Round(((nPotCuo * pnTasIntCal) / (nPotCuo - 1)) * pnMontoReq, 2)
End Function
Function GeneraPlanCal(ByVal lnMontoDes As Currency, ByVal lnTasaInt As Double, _
                        ByVal lnPlazo As Long, ByVal lnDias As Long, ByVal lnPeriodo As Long, _
                        ByVal lsTipoAnexo As String, Optional ByVal lnTasSegDes As Double = 0, Optional ByVal lnTasSegInc As Double = 0, _
                        Optional ByVal lnTasComCof As Double = 0) As Currency

Dim lnSaldoK As Currency
Dim lnTotalInt  As Currency
Dim ldFecPro As Date
Dim sql As String
Dim oCon As DConecta
Dim lnCuotaIni As Currency
Dim lnCuota As Currency
Dim I As Long
Dim lnMontoBruto As Currency
Dim lnSegDes As Currency
Dim lnSegInc As Currency
Dim lnComCof As Currency
Dim rs As ADODB.Recordset

Set oCon = New DConecta

lnSaldoK = 0
lnTotalInt = 0
'ldFecPro = {}
lnInteres = 0
lnCapital = 0

oCon.AbreConexion

sql = "select * from dbo.sysobjects where name ='PLANPAGOSTMP'"
Set rs = oCon.CargaRecordSet(sql)
If Not rs.BOF And Not rs.BOF Then
    sql = "DROP TABLE PLANPAGOSTMP"
    oCon.Ejecutar sql
End If
rs.Close
Set rs = Nothing

sql = "CREATE TABLE PLANPAGOSTMP( " _
    & " cNroCuo     CHAR(3)    , " _
    & " nCapital    money , " _
    & " nInteres    money , " _
    & " nSegDes     money , " _
    & " nSegInc     money , " _
    & " nTasCom     money , " _
    & " nSaldoCap   money , " _
    & " nCuota      money , " _
    & " nTasInt     money )"

oCon.Ejecutar sql

lnMontoBruto = lnMontoDes
If lsTipoAnexo = "K2" Or lsTipoAnexo = "K1" Then
    lnSaldoK = lnMontoBruto * 0.8
Else
    lnSaldoK = lnMontoBruto
End If
lnCuotaIni = cuotaconstante(lnTasaInt, lnPlazo, lnSaldoK, lnDias, lnPeriodo)

pnTasInt = 0
lnInteresTotal = 0
lnSegDes = 0
lnSegInc = 0
lnComCof = 0

For I = 1 To lnPlazo
    lnSegDes = 0
    lnSegInc = 0
    lnComCof = 0

    pnTasInt = InteresReal(lnTasaInt, lnDias, lnPeriodo)
    
    lnInteres = Round(pnTasInt * lnSaldoK, 2)
    lnCapital = lnCuotaIni - lnInteres
    lnInteresTotal = lnInteresTotal + lnInteres
    'para mi vivienda solamente
    lnSegDes = Round((lnTasSegDes / 100) * lnSaldoK, 2)
    lnSegInc = Round((((lnTasSegInc / 100) * lnMontoBruto) / 0.9), 2)
    lnComCof = Round((lnTasComCof / 100) * lnSaldoK, 2)
    '************************************************
    lnSaldoK = lnSaldoK - lnCapital
    
    lnCuota = lnCuotaIni + lnSegDes + lnSegInc + lnComCof
    
    sql = " INSERT INTO PLANPAGOSTMP (cNroCuo, nCapital, nInteres, nSaldoCap, nCuota, nTasInt ) " _
        & " VALUES('" & Format(I, "000") & "'," & lnCapital & "," & lnInteres & "," & lnSaldoK & "," & lnCuota & "," & lnTasaInt & ")"
    
    oCon.Ejecutar sql
Next
oCon.CierraConexion
Set oCon = Nothing
GeneraPlanCal = lnCuota
End Function

Function GetTir(ByVal lnTasaInt As Double, ByVal lnMontTotal As Currency) As Double
Dim lnTir As Double
'Dim lnVan As Currency
Dim oCon As DConecta
Dim sql As String
Dim rs As ADODB.Recordset
Dim I As Long
Dim lnVan As Double
Dim gnMenor As Double
Dim gnMayor As Double

Set oCon = New DConecta
Set rs = New ADODB.Recordset

oCon.AbreConexion

lnTir = 0
lnVan = 0

sql = " SELECT  CONVERT(INT,cnrocuo)*-1 as ntiempo, cnrocuo, nCuota " _
     & " FROM    PLANPAGOSTMP"
     
Set rs = oCon.CargaRecordSet(sql)
lnTir = lnTasaInt
lnVan = 99999

lnTotalItems = 0

lnTotalItems = rs.RecordCount
If lnTotalItems = 0 Then
    Return
End If
lnSuma = 0
I = 0
gnMenor = -0.00001
gnMayor = 0.00001
lnDif = 0
lnIterac = 0
lnMontoCred = lnMontTotal

Do While True
    lnIterac = lnIterac + 1
    lnSuma = 0
    rs.MoveFirst
    Do While Not rs.EOF
        lnSuma = lnSuma + rs!nCuota * ((1 + (lnTir / 100)) ^ (rs!nTiempo))
        'lnSuma = lnSuma + (rs!nCuota / ((1 + (lnTir / 100)) ^ rs!nTiempo))
        rs.MoveNext
    Loop
    lnVan = lnSuma - lnMontoCred
    If (lnVan >= -5 And lnVan <= 5) Then
        Exit Do
    End If
    
    lnDif = Abs((gnMayor - gnMenor + 0.0001) * Rnd(gnMenor))
    If lnVan < 0 Then
        lnTir = lnTir - lnDif
    End If
    If lnVan > 0 Then
        lnTir = lnTir + lnDif
    End If
Loop
GetTir = lnTir

End Function
Function GetDatosreporte16(ByVal pnMoneda As Moneda, pnMonto1 As Currency, pnMonto2 As Currency, pnMes1 As Long, pnMes2 As Long, MatProd As Variant) As ADODB.Recordset
Dim sql As String
Dim rs As ADODB.Recordset
Dim oCon As DConecta
Dim sCadProd As String
Dim I As Integer
Dim lnPlazo1 As Long
Dim lnPlazo2 As Long
  
    sCadProd = "('"
    For I = 0 To UBound(MatProd) - 1
        sCadProd = sCadProd & MatProd(I) & "','"
    Next I
    sCadProd = Mid(sCadProd, 1, Len(sCadProd) - 2) & ")"
    
    lnPlazo1 = IIf(pnMes1 <= 12, 1, 2) '1 corto plazo , 2 largo plazo.
    lnPlazo2 = IIf(pnMes2 <= 12, 1, 2) '1 corto plazo , 2 largo plazo.

sql = "SELECT * "
sql = sql + " FROM      dbo.RepColumna RC"
sql = sql + "           Left Join"
sql = sql + "           (SELECT ISNULL(NCOND1.NCOLUMNA,NCOND2.NCOLUMNA) AS NCOLUMNA,"
sql = sql + "                   ISNULL(NCOND1.nMoneda,NCOND2.nMoneda) AS NMONEDA,"
sql = sql + "                   ISNULL(NCOND1.nPlazo,NCOND2.nPlazo) AS nPlazo, "
sql = sql + "                   ISNULL(CASE WHEN ISNULL(NCOND1.NCOLUMNA,NCOND2.NCOLUMNA)=9 THEN ROUND(nTasaCol1_1*360,2) ELSE dbo.ConvierteTEMaTEA(nTasaCol1_1) END,0) as nTasAnual1_1, "
sql = sql + "                   nTasaCol1_1, ISNULL(nMontoCol1_1,0) as nMontoCol1_1,"
sql = sql + "                   ISNULL(CASE WHEN ISNULL(NCOND1.NCOLUMNA,NCOND2.NCOLUMNA)=9 THEN ROUND(nTasaCol1_2*360,2) ELSE dbo.ConvierteTEMaTEA(nTasaCol1_2) END,0) as nTasAnual1_2,"
sql = sql + "                   nTasaCol1_2, isnull(nMontoCol1_2,0) as nMontoCol1_2,"
sql = sql + "                   ISNULL(CASE WHEN ISNULL(NCOND1.NCOLUMNA,NCOND2.NCOLUMNA)=9 THEN ROUND(nTasaCol2_1*360,2) ELSE dbo.ConvierteTEMaTEA(nTasaCol2_1) END,0) as nTasAnual2_1,"
sql = sql + "                   nTasaCol2_1, isnull(nMontoCol2_1,0) as nMontoCol2_1,"
sql = sql + "                   ISNULL(CASE WHEN ISNULL(NCOND1.NCOLUMNA,NCOND2.NCOLUMNA)=9 THEN ROUND(nTasaCol2_2*360,2) ELSE dbo.ConvierteTEMaTEA(nTasaCol2_2) END,0) as nTasAnual2_2,"
sql = sql + "                   nTasaCol2_2, ISNULL(nMontoCol2_2,0) as nMontoCol2_2"
sql = sql + "             From "
sql = sql + "                   (SELECT CASE  nColocLinCredTasaTpo WHEN 1 THEN 1"
sql = sql + "                                                           WHEN 3 THEN 9"
sql = sql + "                                                               ELSE NULL END AS NCOLUMNA,"
sql = sql + "                           nMoneda, nPlazo, "
sql = sql + "                           MIN(nTasaFin) as nTasaCol1_1, 0 as nMontoCol1_1,"
sql = sql + "                           MAX(nTasaFin) as nTasaCol1_2, 0 as nMontoCol1_2"
sql = sql + "                    From"
sql = sql + "                       (select     L.cLineaCred, substring(L.cLineaCred,5,1) as nMoneda,"
sql = sql + "                                   substring(L.cLineaCred,6,1) as nPlazo,"
sql = sql + "                                   substring(L.cLineaCred,7,3) as nProducto,"
sql = sql + "                                   nPlazoMin, nPlazoMax,nMontoMax, nMontoMin,"
sql = sql + "                                   nColocLinCredTasaTpo , nTasaFin"
sql = sql + "                          FROM     coloclineacredito l "
sql = sql + "                                   JOIN  coloclineacreditoTasa lt on lt.cLineaCred = L.cLineaCred"
sql = sql + "                                   join  colocaciones c1 on c1.cLineaCred = l.cLineaCred"
sql = sql + "                                   join  producto p on p.cCtaCod = c1.cCtaCod "
sql = sql + "                         WHERE     p.nprdEstado in (2020,2021,2022,2030,2031,2031,2201,2205,2101,2101,2106,2107) "
sql = sql + "                                   AND nColocLinCredTasaTpo IN (1,3)"
sql = sql + "                         GROUP by l.cLineaCred,substring(L.cLineaCred,5,1), substring(L.cLineaCred,6,1) ,"
sql = sql + "                                   substring(L.cLineaCred,7,3), nPlazoMin, nPlazoMax,nPlazoMin, nMontoMax, nMontoMin,"
sql = sql + "                                   nColocLinCredTasaTpo, nTasaFin) as cLin"
sql = sql + "                       where   " & pnMonto1 & " between cLin.nMontoMin  and cLin.nMontoMax and nProducto in " & sCadProd & "  and nMoneda =" & pnMoneda & ""
sql = sql + "                               AND NPLAZO = " & lnPlazo1
sql = sql + "                       GROUP BY nMoneda, nPlazo, nColocLinCredTasaTpo )  AS NCOND1"
sql = sql + "                       FULL OUTER JOIN"
sql = sql + "                       (SELECT CASE  nColocLinCredTasaTpo WHEN 1 THEN 1"
sql = sql + "                                                               WHEN 3 THEN 9"
sql = sql + "                                                               ELSE NULL END AS NCOLUMNA,"
sql = sql + "                                   nMoneda, nPlazo, "
sql = sql + "                                   MIN(nTasaFin)  as nTasaCol2_1, 0 as nMontoCol2_1,"
sql = sql + "                                   MAX(nTasaFin) as nTasaCol2_2, 0 as nMontoCol2_2"
sql = sql + "                        From "
sql = sql + "                                   (SELECT  L.cLineaCred, substring(L.cLineaCred,5,1) as nMoneda,"
sql = sql + "                                            substring(L.cLineaCred,6,1) as nPlazo,"
sql = sql + "                                            substring(L.cLineaCred,7,3) as nProducto,"
sql = sql + "                                            nPlazoMin, nPlazoMax,nMontoMax, nMontoMin,"
sql = sql + "                                            nColocLinCredTasaTpo , nTasaFin"
sql = sql + "                                     from   coloclineacredito l"
sql = sql + "                                            JOIN  coloclineacreditoTasa lt on lt.cLineaCred = L.cLineaCred"
sql = sql + "                                            join  colocaciones c1 on c1.cLineaCred = l.cLineaCred"
sql = sql + "                                            join  producto p on p.cCtaCod = c1.cCtaCod"
sql = sql + "                                     where  p.nprdEstado in (2020,2021,2022,2030,2031,2031,2201,2205,2101,2101,2106,2107)"
sql = sql + "                                            AND nColocLinCredTasaTpo IN (1,3)"
sql = sql + "                                     group by l.cLineaCred,substring(L.cLineaCred,5,1), substring(L.cLineaCred,6,1) ,"
sql = sql + "                                               substring(L.cLineaCred,7,3), nPlazoMin, nPlazoMax,nPlazoMin, nMontoMax, nMontoMin,"
sql = sql + "                                               nColocLinCredTasaTpo, nTasaFin) as cLin"
sql = sql + "                         where   " & pnMonto2 & " between cLin.nMontoMin  and cLin.nMontoMax and nProducto in " & sCadProd & " and nMoneda =" & pnMoneda & ""
sql = sql + "                                   AND NPLAZO = " & lnPlazo2
sql = sql + "                         GROUP BY nMoneda, nPlazo, nColocLinCredTasaTpo ) AS NCOND2 ON NCOND1.NCOLUMNA = NCOND2.NCOLUMNA"
sql = sql + "                         Union"
sql = sql + "                         SELECT    NCOLUMNA, NMONEDA, '0' AS NPLAZO, "
sql = sql + "                                   SUM(NVALOR) AS nTasAnual1_1, SUM(NVALOR) AS nTasaCol1_1,   SUM(NMONTO) AS nMontoCol1_1,"
sql = sql + "                                   SUM(NVALOR) AS nTasAnual1_2, SUM(NVALOR) AS nTasaCol1_2,   SUM(NMONTO) AS nMontoCol1_2,"
sql = sql + "                                   SUM(NVALOR) AS nTasAnual2_1, SUM(NVALOR) AS nTasaCol2_1,   SUM(NMONTO) AS nMontoCol2_1,"
sql = sql + "                                   SUM(NVALOR) AS nTasAnual2_2, SUM(NVALOR) AS nTasaCol2_2,   SUM(NMONTO) AS nMontoCol2_2 "
sql = sql + "                           From"
sql = sql + "                                   (   SELECT  8 as NCOLUMNA, p.cDescripcion, " & pnMoneda & " AS nMoneda, 0 AS nValor, " & IIf(pnMoneda = 1, 2, 1) & " as nMonto "
sql = sql + "                                       FROM    ProductoConcepto p"
sql = sql + "                                       WHERE   p.nPrdConceptoCod in (124350)"
sql = sql + "                                       GROUP BY p.nPrdConceptoCod, p.cDescripcion, p.nMoneda, p.nValor) AS NCOM"
sql = sql + "                           GROUP BY NCOLUMNA, NMONEDA"
sql = sql + "                           Union"
sql = sql + "                           SELECT  NCOLUMNA, NMONEDA, '0' AS NPLAZO, "
sql = sql + "                                   SUM(NVALOR) AS nTasAnual1_1, SUM(NVALOR) AS nTasaCol1_1, 0 AS nMontoCol1_1,"
sql = sql + "                                   SUM(NVALOR) AS nTasAnual1_2, SUM(NVALOR) AS nTasaCol1_2, 0 AS nMontoCol1_2,"
sql = sql + "                                   SUM(NVALOR) AS nTasAnual2_1, SUM(NVALOR) AS nTasaCol2_1, 0 AS nMontoCol2_1,"
sql = sql + "                                   SUM(NVALOR) AS nTasAnual2_2, SUM(NVALOR) AS nTasaCol2_2, 0 AS nMontoCol2_2"
sql = sql + "                           From "
sql = sql + "                               (SELECT     CASE p.nPrdConceptoCod when 124300 then 2"
sql = sql + "                                                                           when 124200 then 3"
sql = sql + "                                                                                   when 1258 then 5"
sql = sql + "                                                                                           when 1204 then 5"
sql = sql + "                                                                                                   when 124350 then 8"
sql = sql + "                                                                                                       when 12434 then 4"
sql = sql + "                                                                                                           else null end as NCOLUMNA,"
sql = sql + "                                           p.cDescripcion, p.nMoneda, convert(money,nValor) AS nValor, 0 as nMonto"
sql = sql + "                                FROM       ProductoConcepto p"
sql = sql + "                                           left join ProductoConceptoFiltro pf on pf.nPrdConceptoCod = p.nPrdConceptoCod "
sql = sql + "                                WHERE   p.nPrdConceptoCod in (124350,1258,124200,1204,124300,12434) and p.nMoneda = " & pnMoneda & " and nProdCod in " & sCadProd & ""
sql = sql + "                                GROUP BY p.nPrdConceptoCod, p.cDescripcion, p.nMoneda, p.nValor"
sql = sql + "                                ) AS GASTOSMV"
sql = sql + "              GROUP BY NCOLUMNA, NMONEDA ) AS DATOS ON DATOS.NCOLUMNA = RC.nNroCol"
sql = sql + "              Where rc.cOpeCod = 780130"
sql = sql + "              order by RC.nNroCol"

Set oCon = New DConecta
oCon.AbreConexion
Set GetDatosreporte16 = oCon.CargaRecordSet(sql)
oCon.CierraConexion
Set oCon = Nothing
End Function
Sub CreRSReporte16()
Set rsRep16 = New ADODB.Recordset
rsRep16.Fields.Append "nColumna", adInteger
rsRep16.Fields.Append "cConcepto", adVarChar, 200
rsRep16.Fields.Append "nTasa1_1", adCurrency
rsRep16.Fields.Append "nTasaMen1_1", adCurrency
rsRep16.Fields.Append "nMonto1_1", adCurrency
rsRep16.Fields.Append "nTasa1_2", adCurrency
rsRep16.Fields.Append "nTasaMen1_2", adCurrency
rsRep16.Fields.Append "nMonto1_2", adCurrency
rsRep16.Fields.Append "nTasa2_1", adCurrency
rsRep16.Fields.Append "nTasaMen2_1", adCurrency
rsRep16.Fields.Append "nMonto2_1", adCurrency
rsRep16.Fields.Append "nTasa2_2", adCurrency
rsRep16.Fields.Append "nTasaMen2_2", adCurrency
rsRep16.Fields.Append "nMonto2_2", adCurrency
rsRep16.Open
End Sub

Sub DatosReporte16(ByVal pnMoneda As Moneda, ByVal pnMonto1 As Currency, ByVal pnMonto2 As Currency, ByVal pnMes1 As Long, ByVal pnMes2 As Long, ByVal MatProd As Variant, ByVal lsAnexo As String)
Dim rs As ADODB.Recordset
Dim lnSegDes As Double
Dim lnSegInc As Double
Dim lnComCof As Double

Dim lnCuotaMin1 As Currency
Dim lnCuotaMin2 As Currency

Dim lnCuotaMax1 As Currency
Dim lnCuotaMax2 As Currency

Dim lnTirAnualMin1 As Currency
Dim lnTirAnualMax1 As Currency

Dim lnTirAnualMin2 As Currency
Dim lnTirAnualMax2 As Currency

Dim lnTasaCal As Currency


Set rs = GetDatosreporte16(pnMoneda, pnMonto1, pnMonto2, pnMes1, pnMes2, MatProd)

lnSegDes = 0
lnSegInc = 0
lnComCof = 0

CreRSReporte16
Do While Not rs.EOF
    rsRep16.AddNew
    rsRep16("nColumna").value = rs!nNroCol
    rsRep16("cConcepto").value = rs!cDescCol
    
    rsRep16("nTasa1_1").value = IIf(IsNull(rs!nTasAnual1_1), 0, rs!nTasAnual1_1)
    rsRep16("nTasaMen1_1").value = IIf(IsNull(rs!nTasaCol1_1), 0, rs!nTasaCol1_1)
    rsRep16("nMonto1_1").value = IIf(IsNull(rs!nMontoCol1_1), 0, rs!nMontoCol1_1)
    rsRep16("nTasa1_2").value = IIf(IsNull(rs!nTasAnual1_2), 0, rs!nTasAnual1_2)
    rsRep16("nTasaMen1_2").value = IIf(IsNull(rs!nTasaCol1_2), 0, rs!nTasaCol1_2)
    rsRep16("nMonto1_2").value = IIf(IsNull(rs!nMontoCol1_2), 0, rs!nMontoCol1_2)
    rsRep16("nTasa2_1").value = IIf(IsNull(rs!nTasAnual2_1), 0, rs!nTasAnual2_1)
    rsRep16("nTasaMen2_1").value = IIf(IsNull(rs!nTasaCol2_1), 0, rs!nTasaCol2_1)
    rsRep16("nMonto2_1").value = IIf(IsNull(rs!nMontoCol2_1), 0, rs!nMontoCol2_1)
    rsRep16("nTasa2_2").value = IIf(IsNull(rs!nTasAnual2_2), 0, rs!nTasAnual2_2)
    rsRep16("nTasaMen2_2").value = IIf(IsNull(rs!nTasaCol2_2), 0, rs!nTasaCol2_2)
    rsRep16("nMonto2_2").value = IIf(IsNull(rs!nMontoCol2_2), 0, rs!nMontoCol2_2)
    
    Select Case rs!nNroCol
        Case 2 'seguro de desgravamet
            lnSegDes = IIf(IsNull(rs!nTasaCol1_1), 0, rs!nTasaCol1_1)
        Case 3 'seguro de incendio
            lnSegInc = IIf(IsNull(rs!nTasaCol1_1), 0, rs!nTasaCol1_1)
        Case 4 'portes mi vivienda
            lnComCof = IIf(IsNull(rs!nTasaCol1_1), 0, rs!nTasaCol1_1)
    End Select
    rs.MoveNext
Loop
Dim lnTirMin1 As Currency
Dim lnTirMax1 As Currency
Dim lnTirMin2 As Currency
Dim lnTirMax2 As Currency


If Not (rsRep16.EOF And rsRep16.BOF) Then rsRep16.MoveFirst
Do While Not rsRep16.EOF
    Select Case rsRep16!nColumna
        Case 1
            '***************************************************************************************
            If rsRep16!nTasaMen1_1 > 0 Then
                lnCuotaMin1 = GeneraPlanCal(pnMonto1, rsRep16!nTasaMen1_1, pnMes1, 30, 30, lsAnexo, lnSegDes, lnSegInc, lnComCof)
                lnTirMin1 = GetTir(rsRep16!nTasaMen1_1, pnMonto1)
                'lnTirAnualMin1 = ((1 + (lnTirMin1 / 100)) ^ 12 - 1) * 100
                lnTirAnualMin1 = lnTirMin1
                lnCuotaMin1 = cuotaconstante(lnTirMin1, pnMes1, pnMonto1, 30, 30)
            End If
            '***************************************************************************************
            If rsRep16!nTasaMen1_2 > 0 Then
                lnCuotaMax1 = GeneraPlanCal(pnMonto1, rsRep16!nTasaMen1_2, pnMes1, 30, 30, lsAnexo, lnSegDes, lnSegInc, lnComCof)
                lnTirMax1 = GetTir(rsRep16!nTasaMen1_2, pnMonto1)
                'lnTirAnualMax1 = ((1 + lnTirMax1 / 100) ^ 12 - 1) * 100
                lnTirAnualMax1 = lnTirMax1
                lnCuotaMax1 = cuotaconstante(lnTirMax1, pnMes1, pnMonto1, 30, 30)
            End If
            '***************************************************************************************
            If rsRep16!nTasaMen2_1 > 0 Then
                lnCuotaMin2 = GeneraPlanCal(pnMonto2, rsRep16!nTasaMen2_1, pnMes2, 30, 30, lsAnexo, lnSegDes, lnSegInc, lnComCof)
                lnTirMin2 = GetTir(rsRep16!nTasaMen2_1, pnMonto2)
                'lnTirAnualMin2 = ((1 + lnTirMin2 / 100) ^ 12 - 1) * 100
                lnTirAnualMin2 = lnTirMin2
                lnCuotaMin2 = cuotaconstante(lnTirMin2, pnMes2, pnMonto2, 30, 30)
            End If
            '******************************************************************************************
            If rsRep16!nTasaMen2_2 > 0 Then
                lnCuotaMax2 = GeneraPlanCal(pnMonto2, rsRep16!nTasaMen2_2, pnMes2, 30, 30, lsAnexo, lnSegDes, lnSegInc, lnComCof)
                lnTirMax2 = GetTir(rsRep16!nTasaMen2_2, pnMonto2)
                'lnTirAnualMax2 = ((1 + lnTirMax2 / 100) ^ 12 - 1) * 100
                lnTirAnualMax2 = lnTirMax2
                lnCuotaMax2 = cuotaconstante(lnTirMax2, pnMes2, pnMonto2, 30, 30)
            End If
        Case 2
            
        Case 3
            
        Case 4
            
        Case 5
            rsRep16("nMonto1_1").value = rsRep16("nTasa1_1").value / 100 * pnMonto1
            rsRep16("nMonto1_2").value = rsRep16("nTasa1_2").value / 100 * pnMonto1
            
            rsRep16("nMonto2_1").value = rsRep16("nTasa2_1").value / 100 * pnMonto2
            rsRep16("nMonto2_2").value = rsRep16("nTasa2_2").value / 100 * pnMonto2
        Case 6 'TIR ANUAL
            rsRep16("nTasa1_1").value = lnTirAnualMin1
            rsRep16("nTasa1_2").value = lnTirAnualMax1
            
            rsRep16("nTasa2_1").value = lnTirAnualMin2
            rsRep16("nTasa2_2").value = lnTirAnualMax2
        Case 7 'CUOTA MENSUAL
            rsRep16("nMonto1_1").value = lnCuotaMin1
            rsRep16("nMonto1_2").value = lnCuotaMax1
            
            rsRep16("nMonto2_1").value = lnCuotaMin2
            rsRep16("nMonto2_2").value = lnCuotaMax2
    End Select
    
    rsRep16.MoveNext
Loop

If Not (rsRep16.EOF And rsRep16.BOF) Then rsRep16.MoveFirst
If lsAnexo = "K1" Or lsAnexo = "K2" Then
    GeneraReporte16K rsRep16, pnMonto1, pnMonto2, pnMes1, pnMes2, pnMoneda, lsAnexo
Else
    GeneraReporte16 rsRep16, pnMonto1, pnMonto2, pnMes1, pnMes2, pnMoneda, lsAnexo
End If

End Sub
Sub GeneraReporte16K(ByVal rs As ADODB.Recordset, pnMonto1 As Currency, pnMonto2 As Currency, pnMes1 As Long, pnMes2 As Long, ByVal pnMoneda As Currency, ByVal lsAnexo As String)
Dim vNHC As String
Dim vExcelObj As Excel.Application
Dim I As Long

vNHC = App.path & "\spooler\REPORTE16_" & lsAnexo & Format(txtFecha, "yyyymmdd") & ".XLS"

Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
vExcelObj.DisplayAlerts = False

vExcelObj.Workbooks.Add
vExcelObj.Sheets("Hoja1").Select
vExcelObj.Sheets("Hoja1").Name = "ANEXO" & lsAnexo

vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
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

vExcelObj.Range("E1").Select
vExcelObj.Range("E1").Font.Bold = True
vExcelObj.Range("E1").HorizontalAlignment = xlHAlignCenter
vExcelObj.ActiveCell.value = Format(gdFecSis, "dd/mm/yyyy")

vExcelObj.Range("B2").Select
vExcelObj.Range("B2").Font.Bold = True
vExcelObj.Range("B2").HorizontalAlignment = xlHAlignCenter
vExcelObj.ActiveCell.value = "REPORTE 16 ANEXO " & lsAnexo

vExcelObj.Range("B4").Select
vExcelObj.Range("B4").Font.Bold = True
vExcelObj.Range("B4").HorizontalAlignment = xlHAlignCenter
vExcelObj.ActiveCell.value = "PRESTAMO HIPOTECARIO MI VIVIENDA EN " & IIf(pnMoneda = 1, "MONEDA NACIONAL + VAC", "MONEDA EXTRANJERA")

Dim vCel As String
Dim lsCeldaIni As String
Dim lsCeldaFin As String

vIni = 7
vItem = vIni
lnTotalCtasCMAC = 0
lsCeldaIni = ""
lsCeldaFin = ""
Do While Not rs.EOF
    Select Case rs!nColumna
        Case 1
            vCel = "B" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "DE S/.", "DE US.") & pnMonto1 & " A " & Format(pnMes1 / 12)
            
            vCel = "B" + Trim(Str(vItem - 1)) & ":C" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Merge
            
            vCel = "D" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "DE S/.", "DE US.") & pnMonto2 & " A " & Format(pnMes2 / 12)
            
            vCel = "D" + Trim(Str(vItem - 1)) & ":E" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Merge
            
            
            vCel = "A" + Trim(Str(vItem))
            lsCeldaIni = "A" + Trim(Str(vItem - 1))
            
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 30
            vExcelObj.ActiveCell.value = "CONCEPTOS"
            
            vCel = "A" + Trim(Str(vItem - 1)) & ":A" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Merge
                        
            vCel = "B" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "TASA"
            
            vCel = "B" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "(%)"
            
            vCel = "C" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "MONTO"
            
            vCel = "C" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "(DOLARES)"
            
            vCel = "D" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "Tasa"
            
            vCel = "D" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "(%)"
            
            vCel = "E" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "MONTO"
            
            vCel = "E" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "(DOLARES)"
                        
            vItem = vItem + 1
        Case 8
            'parte 2 del reporte
            vItem = vItem + 4
            vCel = "B" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "DE S/.", "DE US.") & pnMonto1 & " A " & Format(pnMes1 / 12)
            
            vCel = "B" + Trim(Str(vItem - 1)) & ":C" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Merge
            
            vCel = "D" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "DE S/.", "DE US.") & pnMonto2 & " A " & Format(pnMes2 / 12)
            
            vCel = "D" + Trim(Str(vItem - 1)) & ":E" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Merge
            
            vCel = "A" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "OTROS GASTOS"
            
            lsCeldaIni = "A" + Trim(Str(vItem - 1))
            
            vCel = "A" + Trim(Str(vItem - 1)) & ":A" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Merge
            
            vCel = "B" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "TASA"
            
            vCel = "B" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "(%)"
            
            vCel = "C" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "MONTO"
            
            vCel = "C" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "(DOLARES)"
            
            vCel = "D" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "Tasa"
            
            vCel = "D" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "(%)"
            
            
            vCel = "E" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "MONTO"
            
            vCel = "E" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "(DOLARES)"
            
            vItem = vItem + 1
            
            Dim lsTexto As String
            For I = 1 To 5
                vItem = vItem + 1
                Select Case I
                    Case 1
                        lsTexto = "Tasacion"
                    Case 2
                        lsTexto = "Estudios de Titulos"
                    Case 3
                        lsTexto = "Gastos Notariales"
                    Case 4
                        lsTexto = "Gastos Registrales"
                    Case 5
                        lsTexto = "Otros"
                End Select
                vCel = "A" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = lsTexto
            Next
            
            vCel = "E" + Trim(Str(vItem))
            lsCeldaFin = vCel
            
            vCel = lsCeldaIni & ":" & lsCeldaFin
            vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin
            vExcelObj.Range(vCel).Borders(xlInsideVertical).LineStyle = xlContinuous
            vExcelObj.Range(vCel).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
            '********************************************************************************
            vItem = vItem + 2
            vCel = "A" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin
            vExcelObj.ActiveCell.value = "Es Obligatorio Realizar estos Trámites del Banco ??"
            
            vCel = "B" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin
            vExcelObj.ActiveCell.value = "2"
            
            '*************************************************************************************************+
            'parte 3 del reporte
            vItem = vItem + 4
            vCel = "B" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "DE S/.", "DE US.") & pnMonto1 & " A " & Format(pnMes1 / 12)
            
            vCel = "B" + Trim(Str(vItem - 1)) & ":C" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Merge
            
            vCel = "D" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "DE S/.", "DE US.") & pnMonto2 & " A " & Format(pnMes2 / 12)
            
            vCel = "D" + Trim(Str(vItem - 1)) & ":E" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Merge
            
            vCel = "A" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "INFORMACION DE MORAS Y ATRASOS"
            
            lsCeldaIni = "A" + Trim(Str(vItem - 1))
            
            vCel = "A" + Trim(Str(vItem - 1)) & ":A" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Merge
            
            vCel = "B" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "TASA"
            
            vCel = "B" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "(%)"
            
            vCel = "C" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "MONTO"
            
            vCel = "C" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "(DOLARES)"
            
            vCel = "D" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "Tasa"
            
            vCel = "D" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "(%)"
            
            vCel = "E" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "MONTO"
            
            vCel = "E" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "(DOLARES)"
            
            lsCeldaFin = vCel
            
            vCel = lsCeldaIni & ":" & lsCeldaFin
            vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin
            vExcelObj.Range(vCel).Borders(xlInsideVertical).LineStyle = xlContinuous
            vExcelObj.Range(vCel).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
            vItem = vItem + 1
    End Select
    
    vItem = vItem + 1
    vCel = "A" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = Trim(rs!cConcepto)
    
    vCel = "B" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = rs!nTasa1_2
    
    vCel = "C" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = rs!nMonto1_2
    
    vCel = "D" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = rs!nTasa2_2
    
    vCel = "E" + Trim(Str(vItem))
    vExcelObj.Range(vCel).Select
    vExcelObj.ActiveCell.value = rs!nMonto2_2
    
         
    Select Case rs!nColumna
        Case 1, 6, 9
            'vCel = "C" + Trim(Str(vItem))
            'vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, vbYellow
            
            'vCel = "D" + Trim(Str(vItem))
            'vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, vbYellow
            
        Case 10
            'vCel = "B" + Trim(Str(vItem))
            'vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, vbYellow
            
            'vCel = "D" + Trim(Str(vItem))
            'vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, vbYellow
            
        Case 11
            vCel = "B" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = "5"
        Case 12
            
            vCel = "B" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = "1"
            
            lsCeldaFin = "E" + Trim(Str(vItem))
            
            vCel = lsCeldaIni & ":" & lsCeldaFin
            vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin
            vExcelObj.Range(vCel).Borders(xlInsideVertical).LineStyle = xlContinuous
            vExcelObj.Range(vCel).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
        Case 7
            
            'vCel = "B" + Trim(Str(vItem))
            'vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, vbYellow
            
            'vCel = "D" + Trim(Str(vItem))
            'vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, vbYellow
            
            
            lsCeldaFin = vCel
            
            vCel = lsCeldaIni & ":" & lsCeldaFin
            vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin
            vExcelObj.Range(vCel).Borders(xlInsideVertical).LineStyle = xlContinuous
            vExcelObj.Range(vCel).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            
            vItem = vItem + 2
            vCel = "A" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin
            vExcelObj.ActiveCell.value = "Cobra Obligatoriamente Seguro por la Vivienda?"
            
            vCel = "B" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin
            vExcelObj.ActiveCell.value = "1"
            
    End Select
    rs.MoveNext
Loop
                    
vExcelObj.Range("A1").Select
vExcelObj.ActiveWorkbook.SaveAs (vNHC)
vExcelObj.ActiveWorkbook.Close
vExcelObj.Workbooks.Open (vNHC)
vExcelObj.Visible = True

Set vExcelObj = Nothing
MsgBox "SE HA GENERADO CON ÉXITO EL ARCHIVO !!  ", vbInformation, " Mensaje del Sistema ..."
                        
End Sub

Sub GeneraReporte16(ByVal rs As ADODB.Recordset, pnMonto1 As Currency, pnMonto2 As Currency, pnMes1 As Long, pnMes2 As Long, ByVal pnMoneda As Currency, ByVal lsAnexo As String)
Dim vNHC As String
Dim vExcelObj As Excel.Application
Dim I As Long

vNHC = App.path & "\spooler\REPORTE16_" & lsAnexo & Format(txtFecha, "yyyymmdd") & ".XLS"
If VerArchivoCargado(vNHC) = True Then
    MsgBox "Archivo se encuentra abierto por favor cerrarlo antes de continuar", vbInformation, "Aviso"
    Exit Sub
End If
Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
vExcelObj.DisplayAlerts = False

vExcelObj.Workbooks.Add
vExcelObj.Sheets("Hoja1").Select
vExcelObj.Sheets("Hoja1").Name = "ANEXO" & lsAnexo

vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
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

vExcelObj.Range("E1").Select
vExcelObj.Range("E1").Font.Bold = True
vExcelObj.Range("E1").HorizontalAlignment = xlHAlignCenter
vExcelObj.ActiveCell.value = Format(gdFecSis, "dd/mm/yyyy")

vExcelObj.Range("B2").Select
vExcelObj.Range("B2").Font.Bold = True
vExcelObj.Range("B2").HorizontalAlignment = xlHAlignCenter
vExcelObj.ActiveCell.value = "REPORTE 16 ANEXO " & lsAnexo

vExcelObj.Range("B4").Select
vExcelObj.Range("B4").Font.Bold = True
vExcelObj.Range("B4").HorizontalAlignment = xlHAlignCenter
Select Case lsAnexo
    Case "D1", "D2"
        vExcelObj.ActiveCell.value = "PRESTAMO PERSONALES EN " & IIf(pnMoneda = 1, "MONEDA NACIONAL", "MONEDA EXTRANJERA")
    Case "E1", "E2"
        vExcelObj.ActiveCell.value = "PRESTAMO PERSONALES EN " & IIf(pnMoneda = 1, "MONEDA NACIONAL", "MONEDA EXTRANJERA") & " CON INGRESO FAMILIAR MENOR S/. 800"
    Case "F1", "F2"
        vExcelObj.ActiveCell.value = "PRESTAMO EN " & IIf(pnMoneda = 1, "MONEDA NACIONAL", "MONEDA EXTRANJERA") & " A MICROEMPRESA CON INGRESOS MENSUALES ENTRE S/. 15 000 Y S/.30 000 "
    Case "G1", "G2"
        vExcelObj.ActiveCell.value = "PRESTAMO EN " & IIf(pnMoneda = 1, "MONEDA NACIONAL", "MONEDA EXTRANJERA") & " A MICROEMPRESA CON INGRESOS MENSUALES ENTRE S/. 2 000 Y S/.5 000 "
End Select

Dim vCel As String
Dim lsCeldaIni As String
Dim lsCeldaFin As String

vIni = 7
vItem = vIni
lnTotalCtasCMAC = 0
lsCeldaIni = ""
lsCeldaFin = ""
Do While Not rs.EOF
    Select Case rs!nColumna
        Case 1, 8
            If rs!nColumna = 8 Then
                vItem = vItem + 4
            End If
            vCel = "B" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "DE S/.", "DE US.") & pnMonto1 & " A " & Format(pnMes1) & " meses"
            
            vCel = "B" + Trim(Str(vItem - 1)) & ":E" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Merge
            
            vCel = "F" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "DE S/.", "DE US.") & pnMonto2 & " A " & Format(pnMes2) & " meses"
            
            vCel = "F" + Trim(Str(vItem - 1)) & ":I" + Trim(Str(vItem - 1))
            vExcelObj.Range(vCel).Merge
            
            vCel = "A" + Trim(Str(vItem))
            lsCeldaIni = "A" + Trim(Str(vItem - 1))
            
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 30
            vExcelObj.ActiveCell.value = IIf(rs!nColumna = 8, "INFORMACION DE MORAS Y ATRASOS", "CONCEPTOS")
            
            vCel = "A" + Trim(Str(vItem - 1)) & ":A" + Trim(Str(vItem + 2))
            vExcelObj.Range(vCel).Merge
            '*********************************************************************************
            'vItem = vItem + 1
            vCel = "B" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "Valores Mínimos"
            
            vCel = "D" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "Valores Máximos"
            
            vCel = "B" + Trim(Str(vItem)) & ":C" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Merge
            
            vCel = "D" + Trim(Str(vItem)) & ":E" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Merge
            
            vCel = "F" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "Valores Mínimos"
            
            vCel = "H" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "Valores Máximos"
            
            vCel = "F" + Trim(Str(vItem)) & ":G" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Merge
            
            vCel = "H" + Trim(Str(vItem)) & ":I" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Merge
            
            '*********************************************************************************
            vItem = vItem + 1
            vCel = "B" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.ActiveCell.value = "TASA"
            
            vCel = "B" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "(%)"
            
            vCel = "C" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "MONTO"
            
            vCel = "C" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "(SOLES)", "(DOLARES)")
            
            vCel = "D" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "Tasa"
            
            vCel = "D" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "(%)"
            
            vCel = "E" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "MONTO"
            
            vCel = "E" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "(SOLES)", "(DOLARES)")
            
            vCel = "F" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "Tasa"
            
            vCel = "F" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "(%)"
            
            vCel = "G" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "MONTO"
            
            vCel = "G" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "(SOLES)", "(DOLARES)")
            
            vCel = "H" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "Tasa"
            
            vCel = "H" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "(%)"
            
            vCel = "I" + Trim(Str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = "MONTO"
            
            vCel = "I" + Trim(Str(vItem + 1))
            vExcelObj.Range(vCel).Select
            vExcelObj.Range(vCel).Font.Bold = True
            vExcelObj.Range(vCel).HorizontalAlignment = xlHAlignCenter
            vExcelObj.Range(vCel).ColumnWidth = 10
            vExcelObj.ActiveCell.value = IIf(pnMoneda = 1, "(SOLES)", "(DOLARES)")
            vItem = vItem + 1
    End Select
    If rs!nColumna <> 2 Then
        vItem = vItem + 1
        vCel = "A" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = Trim(rs!cConcepto)
        
        vCel = "B" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = rs!nTasa1_1
        
        vCel = "C" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = rs!nMonto1_1
        
        vCel = "D" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = rs!nTasa1_2
        
        vCel = "E" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = rs!nMonto1_2
        
        
        vCel = "F" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = rs!nTasa2_1
        
        vCel = "G" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = rs!nMonto2_1
        
        vCel = "H" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = rs!nTasa2_2
        
        vCel = "I" + Trim(Str(vItem))
        vExcelObj.Range(vCel).Select
        vExcelObj.ActiveCell.value = rs!nMonto2_2
             
        Select Case rs!nColumna
            Case 7
                lsCeldaFin = vCel
                
                vCel = lsCeldaIni & ":" & lsCeldaFin
                vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin
                vExcelObj.Range(vCel).Borders(xlInsideVertical).LineStyle = xlContinuous
                vExcelObj.Range(vCel).Borders(xlInsideHorizontal).LineStyle = xlContinuous
            Case 11
                vCel = "B" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = "5"
            Case 12
                lsCeldaFin = vCel
                
                vCel = "B" + Trim(Str(vItem))
                vExcelObj.Range(vCel).Select
                vExcelObj.ActiveCell.value = "1"
                
                vCel = lsCeldaIni & ":" & lsCeldaFin
                vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin
                vExcelObj.Range(vCel).Borders(xlInsideVertical).LineStyle = xlContinuous
                vExcelObj.Range(vCel).Borders(xlInsideHorizontal).LineStyle = xlContinuous
        End Select
    End If
    rs.MoveNext
Loop
vExcelObj.Range("A1").Select
vExcelObj.ActiveWorkbook.SaveAs (vNHC)
vExcelObj.ActiveWorkbook.Close
MsgBox "SE HA GENERADO CON ÉXITO EL ARCHIVO !!  ", vbInformation, " Mensaje del Sistema ..."
vExcelObj.Workbooks.Open (vNHC)
vExcelObj.Visible = True
Set vExcelObj = Nothing

                        

End Sub

