VERSION 5.00
Begin VB.Form frmEstadVenMen 
   Caption         =   "Estadisticas Mensuales"
   ClientHeight    =   2670
   ClientLeft      =   4365
   ClientTop       =   3435
   ClientWidth     =   5325
   Icon            =   "frmEstadVenMen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   2670
   ScaleWidth      =   5325
   Begin VB.TextBox txtTC 
      Alignment       =   1  'Right Justify
      Height          =   345
      Left            =   4125
      TabIndex        =   15
      Text            =   "0.00"
      Top             =   75
      Width           =   945
   End
   Begin VB.CommandButton cmdsalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   4035
      TabIndex        =   11
      Top             =   2145
      Width           =   1155
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   2865
      TabIndex        =   10
      Top             =   2145
      Width           =   1170
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   360
      Left            =   255
      TabIndex        =   9
      Top             =   2145
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Height          =   1590
      Left            =   225
      TabIndex        =   0
      Top             =   435
      Width           =   4935
      Begin VB.OptionButton optRep 
         Caption         =   "Por Analista"
         Height          =   330
         Index           =   7
         Left            =   2355
         TabIndex        =   8
         Top             =   1170
         Width           =   2445
      End
      Begin VB.OptionButton optRep 
         Caption         =   "Por Financiamiento"
         Height          =   330
         Index           =   6
         Left            =   2340
         TabIndex        =   7
         Top             =   840
         Width           =   2445
      End
      Begin VB.OptionButton optRep 
         Caption         =   "Por Oficinas."
         Height          =   330
         Index           =   5
         Left            =   2355
         TabIndex        =   6
         Top             =   510
         Width           =   2445
      End
      Begin VB.OptionButton optRep 
         Caption         =   "Por Rango de Desembolsos"
         Height          =   330
         Index           =   4
         Left            =   2355
         TabIndex        =   5
         Top             =   165
         Width           =   2445
      End
      Begin VB.OptionButton optRep 
         Caption         =   "Por Vinculacion"
         Height          =   330
         Index           =   3
         Left            =   135
         TabIndex        =   4
         Top             =   1170
         Width           =   2055
      End
      Begin VB.OptionButton optRep 
         Caption         =   "Sectores Económicos"
         Height          =   330
         Index           =   2
         Left            =   135
         TabIndex        =   3
         Top             =   840
         Width           =   2055
      End
      Begin VB.OptionButton optRep 
         Caption         =   "Por Producto Detallado"
         Height          =   330
         Index           =   1
         Left            =   135
         TabIndex        =   2
         Top             =   510
         Width           =   2055
      End
      Begin VB.OptionButton optRep 
         Caption         =   "Por Producto"
         Height          =   330
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   165
         Value           =   -1  'True
         Width           =   1515
      End
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Tipo Cambio Mes:"
      Height          =   195
      Left            =   2685
      TabIndex        =   14
      Top             =   150
      Width           =   1275
   End
   Begin VB.Label lblFecha 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "dd/mm/yyyy"
      Height          =   300
      Left            =   1335
      TabIndex        =   13
      Top             =   90
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Informacion al:"
      Height          =   195
      Left            =   210
      TabIndex        =   12
      Top             =   120
      Width           =   1035
   End
End
Attribute VB_Name = "frmEstadVenMen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsTablaTMP As String
Dim lnTC As Currency
Public Sub Inicio(ByVal pdFecFinMes As Date, ByVal pnTCFijMes As Currency)
lnTC = pnTCFijMes
lblFecha = pdFecFinMes
Me.Show 1
End Sub

Private Sub cmdImprimir_Click()
'If optRep(5).value = True Then
    EstadxOficina
'End If
End Sub

Private Sub cmdProcesar_Click()
lnTC = Format(txtTC, "#0.000")
Procesar
End Sub

Private Sub cmdSalir_Click()
VerificaTablaTemporal
Unload Me
End Sub

Private Sub Form_Load()
Dim loConstS As NConstSistemas
Dim loTipCambio As nTipoCambio

CentraForm Me
Me.cmdImprimir.Enabled = False
Me.Frame1.Enabled = False
Me.cmdProcesar.Enabled = True
'Me.lblFecha = "28/02/2005"
'lnTC = 3.257

Set loConstS = New NConstSistemas
    lblFecha = CDate(loConstS.LeeConstSistema(gConstSistCierreMesNegocio))
Set loConstS = Nothing

Set loTipCambio = New nTipoCambio
    lnTC = Format(loTipCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "0.###")
    txtTC = lnTC
Set loTipCambio = Nothing
    
End Sub
Sub Procesar()
Dim SQL As String
Dim oCon As DConecta

VerificaTablaTemporal

SQL = "SELECT * "
SQL = SQL & " INTO " & lsTablaTMP
SQL = SQL & " From "
SQL = SQL & "       (SELECT cCtaCod, nPrdEstado, cRefinan, nSaldoCap, nCapVencido, nDiasAtraso, cMoneda, "
SQL = SQL & "               cCodAnalista, cRecurso, cCodOfi, ISNULL(cRFA,'') AS cRFA,"
SQL = SQL & "               CASE when SUBSTRING(CCTACOD,6,1)='3' THEN 'CONSU'"
SQL = SQL & "                    WHEN SUBSTRING(CCTACOD,6,1) = '4' THEN 'HIPO'"
SQL = SQL & "                    WHEN cActEcon ='0000' AND SUBSTRING(CCTACOD,6,3) IN ('202','102') THEN '0199'"
SQL = SQL & "                    WHEN cActEcon ='0000' AND SUBSTRING(CCTACOD,6,3) NOT IN ('202','102') THEN '3600'"
SQL = SQL & "                    ELSE CASE WHEN cActEcon='' THEN '9999' ELSE ISNULL(cActEcon,'9999') END END as cActEcon,"
SQL = SQL & "               CASE    WHEN SUBSTRING(CCTACOD,6,1)='1' THEN 100 "
SQL = SQL & "                       WHEN SUBSTRING(CCTACOD,6,1)='2' THEN 200"
SQL = SQL & "                       WHEN SUBSTRING(CCTACOD,6,1)='3' AND SUBSTRING(CCTACOD,6,3)<>'305' THEN 300"
SQL = SQL & "                       WHEN SUBSTRING(CCTACOD,6,3) = '305' THEN 305"
SQL = SQL & "                       WHEN SUBSTRING(CCTACOD,6,1)='4' THEN 400 ELSE 9999 END AS CPRODTOT,"
SQL = SQL & "               CASE    WHEN nMontoDesemb BETWEEN 0 AND 500.99 THEN '1 RANGO 0 - 500'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 501 AND 1000.99 THEN '2 RANGO 500 - 1000'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 1001 AND 1500.99 THEN '3 RANGO 1000 - 1500'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 1501 AND 2500.99 THEN '4 RANGO 1500 - 2000'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 2501 AND 3500.99 THEN '5 RANGO 2000 - 3500'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 3501 AND 5000.99 THEN '6 RANGO 3500 - 5000'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 5001 AND 8000.99 THEN '7 RANGO 5000 - 8000'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 8001 AND 12000.99 THEN '8 RANGO 8000 - 12000'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 12001 AND 16000.99 THEN '9 RANGO 12000 - 16000'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 16000 AND 20000.99 THEN 'A RANGO 16000 - 20000'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 20001 AND 30000.99 THEN 'B RANGO 20000 - 30000'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 30001 AND 50000.99 THEN 'C RANGO 30000 - 50000'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 50001 AND 100000.99 THEN 'D RANGO 50000 - 100000'"
SQL = SQL & "                       WHEN nMontoDesemb BETWEEN 100001 AND 200000.99 THEN 'E RANGO 100000 - 200000'"
SQL = SQL & "                       WHEN nMontoDesemb>200001 THEN 'F RANGO >200000' ELSE '***********' END AS CRANGODESEM,"
SQL = SQL & "                nMontoDesemb,"
SQL = SQL & "                case when cmoneda ='1' THEN nSaldoCap ELSE 0 END AS nSaldoMN,"
SQL = SQL & "                case when cmoneda ='2' THEN ROUND(nSaldoCap*" & lnTC & ",2) ELSE 0 END AS nSaldoME,"
SQL = SQL & "                case when cmoneda ='1' THEN VIGCom+VIGMes+VIGCons+VIGHip + VIGCons31 +VIGHip31 ELSE 0 END AS NVIGMN,"
SQL = SQL & "                case when cmoneda ='2' THEN ROUND((VIGCom+VIGMes+VIGCons+VIGHip + VIGCons31 +VIGHip31)*" & lnTC & ",2) ELSE 0 END AS NVIGME,"
SQL = SQL & "                case when cmoneda ='1' THEN VencCOM + VencMES + VencCons + VencHIP + VencCons31 + VencHip31 + VencPrend ELSE 0 END AS NVENCMN,"
SQL = SQL & "                case when cmoneda ='2' THEN ROUND((VencCOM + VencMES + VencCons + VencHIP + VencCons31 + VencHip31 + VencPrend)*" & lnTC & ",2) ELSE 0 END AS NVENCME,"
SQL = SQL & "                case when cmoneda ='1' THEN REFCom+REFMes+REFCons+REFHip + REFCons31 +REFHip31 ELSE 0 END AS NREFMN,"
SQL = SQL & "                case when cmoneda ='2' THEN ROUND((REFCom+REFMes+REFCons+REFHip+REFCons31+REFHip31)*" & lnTC & ",2) ELSE 0 END AS NREFME,"
SQL = SQL & "                case when cmoneda ='1' THEN (JUDCOM + JUDMES + JUDCONS + JUDHIP) ELSE 0 END AS NJUDMN,"
SQL = SQL & "                case when cmoneda ='2' THEN ROUND((JUDCOM + JUDMES + JUDCONS + JUDHIP)*" & lnTC & ",2) ELSE 0 END AS NJUDME,"
SQL = SQL & "                numVigCom + numVigMes+ numVigCons+ numVigHip+ numVigCons31+ numVigHip31 as nNumVig,"
SQL = SQL & "                numVencCOM+numVencMES+ numVencCons+ numVencHIP+numVencCons31+numVencHip31+numVencPrend as nNumVenc,"
SQL = SQL & "                numREFCom +numREFMes + numREFCons + numREFHip + numREFCons31 + numREFHip31 as nNumRef,"
SQL = SQL & "                numJUDCOM +numJUDMES+numJUDCONS+numJUDHIP  as nNumJUD"
SQL = SQL & "          From"
SQL = SQL & "                (SELECT    cCtaCod, nPrdEstado, cRefinan , nSaldoCap,"
SQL = SQL & "                           nCapVencido, nDiasAtraso, SUBSTRING(CCTACOD,9,1) AS cMoneda,"
SQL = SQL & "                           cCodAnalista,  LEFT(cLineaCred,2) as cRecurso, SUBSTRING(CCTACOD,4,2) as cCodOfi, nMontoDesemb,"
SQL = SQL & "                           cActEcon = (    select TOP 1 FI.cActEcon"
SQL = SQL & "                                           from    DBConsolidada..CreditoConsolTotal CT"
SQL = SQL & "                                                   JOIN DBConsolidada..FuenteIngresoConsol FI ON FI.cNumFuente = CT.cNumFuente"
SQL = SQL & "                                           where   CT.cCtaCod = C.cCtaCod),"
SQL = SQL & "                           cRFA = (SELECT CRFA FROM ColocacCred WHERE cCtaCod = C.cCtaCod),"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso<=15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigCom,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigMes,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigCons,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap else 0 end as VigHip,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap-nCapVencido else 0 end as VigCons31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then nSaldoCap-nCapVencido else 0 end as VigHip31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso>15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencCOM,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso>30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencMES,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND SUBSTRING(CCTACOD,6,3)<>'305' AND nDiasAtraso>90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencCons,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso>90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencHIP,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND SUBSTRING(CCTACOD,6,3)<>'305' AND nDiasAtraso BETWEEN 31 AND 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nCapVencido else 0 end as VencCons31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso BETWEEN 31 AND 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nCapVencido else 0 end as VencHip31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,3)='305' AND nDiasAtraso>30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then nSaldoCap else 0 end as VencPrend,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso<=15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFCom,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFMes,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFCons,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap else 0 end as REFHip,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap-nCapVencido else 0 end as REFCons31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then nSaldoCap-nCapVencido else 0 end as REFHip31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDCOM,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDMES,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDCONS,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nPrdEstado in (2201,2205) then nSaldoCap else 0 end as JUDHIP,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso<=15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then 1 else 0 end as numVigCom,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then 1 else 0 end as numVigMes,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then 1 else 0 end as numVigCons,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then 1 else 0 end as numVigHip,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then 1 else 0 end as numVigCons31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='N' then 1 else 0 end as numVigHip31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso>15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then 1 else 0 end as numVencCOM,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso>30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then 1 else 0 end as numVencMES,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND SUBSTRING(CCTACOD,6,3)<>'305' AND nDiasAtraso>90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then 1 else 0 end as numVencCons,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso>90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then 1 else 0 end as numVencHIP,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND SUBSTRING(CCTACOD,6,3)<>'305' AND nDiasAtraso BETWEEN 31 AND 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then 1 else 0 end as numVencCons31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso BETWEEN 31 AND 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then 1 else 0 end as numVencHip31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,3)='305' AND nDiasAtraso>30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) then 1 else 0 end as numVencPrend,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nDiasAtraso<=15 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then 1 else 0 end as numREFCom,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then 1 else 0 end as numREFMes,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then 1 else 0 end as numREFCons,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso<=30 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then 1 else 0 end as numREFHip,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then 1 else 0 end as numREFCons31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nDiasAtraso between 31 and 90 and nPrdEstado in (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107) and cRefinan ='R' then 1 else 0 end as numREFHip31,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='1' AND nPrdEstado in (2201,2205) then 1 else 0 end as numJUDCOM,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='2' AND nPrdEstado in (2201,2205) then 1 else 0 end as numJUDMES,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='3' AND nPrdEstado in (2201,2205) then 1 else 0 end as numJUDCONS,"
SQL = SQL & "                           CASE WHEN SUBSTRING(CCTACOD,6,1)='4' AND nPrdEstado in (2201,2205) then 1 else 0 end as numJUDHIP"
SQL = SQL & "  FROM    DBConsolidada..creditoconsol  C"
SQL = SQL & "  WHERE   nPrdEstado IN (2020,2021,2022,2030,2031,2032,2101,2104,2106,2107,2201,2205)"
SQL = SQL & "          AND cCtaCod NOT IN (SELECT CCTACOD FROM COLOCACCRED WHERE CRFA ='RFA')) AS DATA) AS DATCONSOL"

Set oCon = New DConecta
oCon.AbreConexion
'oCon.Ejecutar SQL
Call oCon.ConexionActiva.Execute(SQL)   'ARCV 30-07-2006
oCon.CierraConexion

Me.cmdImprimir.Enabled = True
Me.Frame1.Enabled = True
Me.cmdProcesar.Enabled = False
Me.txtTC.Enabled = False

MsgBox "Proceso finalizado correctamente ahora podrá realizar las impresiones", vbInformation, "aviso"
End Sub

Sub VerificaTablaTemporal()
Dim rs As ADODB.Recordset
Dim SQL As String
Dim oCon As DConecta

lsTablaTMP = "TMPVENREF" & gsCodUser

Set oCon = New DConecta
oCon.AbreConexion

Set rs = oCon.CargaRecordSet("select * from sysobjects where name like '%" & lsTablaTMP & "%'")
If Not rs.EOF And Not rs.BOF Then
    SQL = "DROP TABLE " & lsTablaTMP
    oCon.Ejecutar SQL
End If
rs.Close
Set rs = Nothing


End Sub

Sub EstadxOficina()
Dim SQL As String
Dim oCon As DConecta
Dim rs As ADODB.Recordset
Dim vExcelObj As Excel.Application
Dim lnRep As Long
Dim lTitulo As String
Dim lsNomHoja As String
Dim vNHC As String
Dim lsTitulo As String
If optRep(0).value = True Then 'por producto
    lTitulo = "MOROSIDAD Y REFINANCIADOS POR TIPO DE CREDITO AL " & lblFecha
    lsNomHoja = "Producto"
    lnRep = 0
    
    SQL = "         SELECT   CPRODTOT, cConsDescripcion as cDescripcion,"
    SQL = SQL & "       NNUMVEN, NVENCMN, NVENCME, NTOTVEN, ROUND((NTOTVEN/NTOTALSAL)*100,4) AS [% MORA.SUBTOTAL], ROUND((NTOTVEN/NTOTALCART)*100,4) AS [% MORA.CAR.TOTAL],"
    SQL = SQL & "       NNUMREF, NREFMN, NREFME, NTOTREF, ROUND((NTOTREF/NTOTALSAL)*100,4) AS [% REF.SUBTOTAL], ROUND((NTOTREF/NTOTALCART)*100,4) AS [% REF.CAR.TOTAL],"
    SQL = SQL & "       NNUMCRED , NSALDOMN, NSALDOME, NTOTALSAL"
    SQL = SQL & "       From"
    SQL = SQL & "           (SELECT     CPRODTOT ,"
    SQL = SQL & "                       SUM(nNumVenc + nNumJUD) AS NNUMVEN,"
    SQL = SQL & "                       SUM(NVENCMN+NJUDMN) AS NVENCMN, SUM(NVENCME+NJUDME) AS NVENCME, SUM(NVENCMN+NJUDMN+ NVENCME+NJUDME) AS NTOTVEN,"
    SQL = SQL & "                       SUM(nNumRef) AS NNUMREF, SUM(NREFMN) AS NREFMN, SUM(NREFME) AS NREFME, SUM(NREFMN+ NREFME) AS NTOTREF,"
    SQL = SQL & "                       COUNT(CCTACOD) AS NNUMCRED, SUM(NSALDOMN) AS NSALDOMN, SUM(NSALDOME) AS NSALDOME,"
    SQL = SQL & "                       SUM(NSALDOMN + NSALDOME) AS NTOTALSAL,"
    SQL = SQL & "                       NTOTALCART = (SELECT SUM(NSALDOMN+NSALDOME) FROM  " & lsTablaTMP & ")"
    SQL = SQL & "                       From " & lsTablaTMP
    SQL = SQL & "                       GROUP BY CPRODTOT) AS X"
    SQL = SQL & "                       JOIN CONSTANTE A ON A.NCONSVALOR = X.CPRODTOT AND NCONSCOD =1001"
    SQL = SQL & "                       ORDER BY CPRODTOT"
End If

If optRep(1).value = True Then 'por producto detallado
    lTitulo = "MOROSIDAD Y REFINANCIADOS POR PRODUCTO DETALLADO" & lblFecha
    lsNomHoja = "Prod_Det"
    lnRep = 1
    
    SQL = "        SELECT   CCODPROD, cConsDescripcion AS cDescripcion, "
    SQL = SQL & "           NNUMVEN, NVENCMN, NVENCME, NTOTVEN, ROUND((NTOTVEN/NTOTALSAL)*100,4) AS [% MORA.SUBTOTAL], ROUND((NTOTVEN/NTOTALCART)*100,4) AS [% MORA.CAR.TOTAL],"
    SQL = SQL & "           NNUMREF, NREFMN, NREFME, NTOTREF, ROUND((NTOTREF/NTOTALSAL)*100,4) AS [% REF.SUBTOTAL], ROUND((NTOTREF/NTOTALCART)*100,4) AS [% REF.CAR.TOTAL],"
    SQL = SQL & "           NNUMCRED , NSALDOMN, NSALDOME, NTOTALSAL"
    SQL = SQL & "   From "
    SQL = SQL & "       (SELECT SUBSTRING(CCTACOD,6,3) AS CCODPROD ,"
    SQL = SQL & "               SUM(nNumVenc + nNumJUD) AS NNUMVEN,"
    SQL = SQL & "               SUM(NVENCMN+NJUDMN) AS NVENCMN, SUM(NVENCME+NJUDME) AS NVENCME, SUM(NVENCMN+NJUDMN+ NVENCME+NJUDME) AS NTOTVEN,"
    SQL = SQL & "               SUM(nNumRef) AS NNUMREF, SUM(NREFMN) AS NREFMN, SUM(NREFME) AS NREFME, SUM(NREFMN+ NREFME) AS NTOTREF,"
    SQL = SQL & "               COUNT(CCTACOD) AS NNUMCRED, SUM(NSALDOMN) AS NSALDOMN, SUM(NSALDOME) AS NSALDOME,"
    SQL = SQL & "               SUM(NSALDOMN + NSALDOME) AS NTOTALSAL,"
    SQL = SQL & "               NTOTALCART = (SELECT SUM(NSALDOMN+NSALDOME) FROM " & lsTablaTMP & ")"
    SQL = SQL & "           From " & lsTablaTMP
    SQL = SQL & "       GROUP BY SUBSTRING(CCTACOD,6,3)) AS X"
    SQL = SQL & "       JOIN CONSTANTE A ON A.NCONSVALOR = X.CCODPROD AND NCONSCOD =1001"
    SQL = SQL & "   ORDER BY CCODPROD"

End If

If optRep(2).value = True Then 'por sectores económicos
    lTitulo = "MOROSIDAD Y REFINANCIADOS POR SECTORES ECONOMICOS" & lblFecha
    lsNomHoja = "SECTORES"
    lnRep = 2
    '--POR SECTORES EXCONOMICOS.
    SQL = "         SELECT  cCodRango, cDesCrip AS cDescripcion,cDivision,"
    SQL = SQL & "           NREFMN, NREFME, NTOTREF,"
    SQL = SQL & "           ROUND((NTOTREF/NTOTALCART)*100,4) AS [% REF.CAR.TOTAL],"
    SQL = SQL & "           NVENCMN , NVENCME, NTOTVEN,"
    SQL = SQL & "           ROUND((NTOTVEN/NTOTALCART)*100,4) AS [% MORA.CAR.TOTAL],"
    SQL = SQL & "           NSALDOMN , NSALDOME, NTOTALSAL"
    SQL = SQL & "   From"
    SQL = SQL & "           (SELECT cCodRango, cDesCrip,  ISNULL(cDivision,'') AS cDivision ,"
    SQL = SQL & "                   ISNULL(SUM(NREFMN),0) AS NREFMN, ISNULL(SUM(NREFME),0) AS NREFME, ISNULL(SUM(NREFMN+ NREFME),0) AS NTOTREF,"
    SQL = SQL & "                   ISNULL(SUM(NVENCMN+NJUDMN),0) AS NVENCMN, ISNULL(SUM(NVENCME+NJUDME),0) AS NVENCME, ISNULL(SUM(NVENCMN+NJUDMN+ NVENCME+NJUDME),0) AS NTOTVEN,"
    SQL = SQL & "                   ISNULL(SUM(NSALDOMN),0) AS NSALDOMN, ISNULL(SUM(NSALDOME),0) AS NSALDOME,"
    SQL = SQL & "                   ISNULL(SUM(NSALDOMN + NSALDOME),0) AS NTOTALSAL,"
    SQL = SQL & "                   NTOTALCART = (SELECT SUM(NSALDOMN+NSALDOME) FROM TMPMEMO8)"
    SQL = SQL & "             From"
    SQL = SQL & "               (SELECT CASE WHEN LEFT(T.cActEcon,2)='CO' THEN 'CO'"
    SQL = SQL & "                       WHEN LEFT(T.cActEcon,2) = 'HI' THEN 'HI' ELSE cCodRango END AS cCodRango,"
    SQL = SQL & "                       CASE WHEN LEFT(T.cActEcon,2)='CO' THEN 'CONSUMO'"
    SQL = SQL & "                       WHEN LEFT(T.cActEcon,2) = 'HI' THEN 'HIPOTECARIO' ELSE cDesCrip END AS cDesCrip,"
    SQL = SQL & "                       RTRIM(RAN.cDesde) + '-' + RTRIM(RAN.cHasta) as cDivision,"
    SQL = SQL & "                       CCTACOD,nNumRef, NREFMN, NREFME, nNumVenc, nNumJUD, NVENCMN, NJUDMN, NVENCME, NJUDME,"
    SQL = SQL & "                       NSALDOMN , NSALDOME"
    SQL = SQL & "                FROM (SELECT *,"
    SQL = SQL & "                       REPLICATE ('0', 2- LEN(CONVERT(CHAR(2),NDESDE)))+ CONVERT(CHAR(2),NDESDE)  AS CDESDE,"
    SQL = SQL & "                       REPLICATE ('0', 2- LEN(CONVERT(CHAR(2),NHASTA)))+ CONVERT(CHAR(2),NHASTA)  AS CHASTA"
    SQL = SQL & "                       From anxriesgosrango"
    SQL = SQL & "                       where   copecod='770030') AS RAN"
    SQL = SQL & "                       FULL OUTER JOIN " & lsTablaTMP & " T ON  LEFT(T.cActEcon,2) BETWEEN CDESDE AND CHASTA ) AS DATA"
    SQL = SQL & "           GROUP BY cCodRango, cDesCrip,cDivision )  AS X"
    SQL = SQL & "           ORDER BY cCodRango, cDesCrip"

End If

If optRep(3).value = True Then 'por vinculacion
End If

If optRep(4).value = True Then 'por rango  de desembolso
    lTitulo = "MOROSIDAD Y REFINANCIADOS POR RANGO DE DESEMBOLSOS AL " & lblFecha
    lsNomHoja = "SECTORES"
    lnRep = 4
    '--- POR RANGO DE DESEMBOLSO
    SQL = "     SELECT  CRANGODESEM AS cDescripcion,"
    SQL = SQL & "       NNUMVEN, NVENCMN, NVENCME, NTOTVEN, ROUND((NTOTVEN/NTOTALCART)*100,4) AS [% MORA.CAR.TOTAL],"
    SQL = SQL & "       NNUMREF, NREFMN, NREFME, NTOTREF, ROUND((NTOTREF/NTOTALCART)*100,4) AS [% REF.CAR.TOTAL],"
    SQL = SQL & "       NNUMCRED , NSALDOMN, NSALDOME, NTOTALSAL"
    SQL = SQL & " From"
    SQL = SQL & "    ("
    SQL = SQL & "     SELECT    CRANGODESEM ,"
    SQL = SQL & "               SUM(nNumVenc + nNumJUD) AS NNUMVEN,"
    SQL = SQL & "               SUM(NVENCMN+NJUDMN) AS NVENCMN, SUM(NVENCME+NJUDME) AS NVENCME, SUM(NVENCMN+NJUDMN+ NVENCME+NJUDME) AS NTOTVEN,"
    SQL = SQL & "               SUM(nNumRef) AS NNUMREF, SUM(NREFMN) AS NREFMN, SUM(NREFME) AS NREFME, SUM(NREFMN+ NREFME) AS NTOTREF,"
    SQL = SQL & "               COUNT(CCTACOD) AS NNUMCRED, SUM(NSALDOMN) AS NSALDOMN, SUM(NSALDOME) AS NSALDOME,"
    SQL = SQL & "               SUM(NSALDOMN + NSALDOME) AS NTOTALSAL,"
    SQL = SQL & "               NTOTALCART = (SELECT SUM(NSALDOMN+NSALDOME) FROM " & lsTablaTMP & ")"
    SQL = SQL & "           From " & lsTablaTMP
    SQL = SQL & "           GROUP BY CRANGODESEM"
    SQL = SQL & "           ) AS X"
    SQL = SQL & "           ORDER BY CRANGODESEM"
End If
If optRep(5).value = True Then
    lTitulo = "MOROSIDAD Y REFINANCIADOS POR OFICINAS AL :" & lblFecha
    lsNomHoja = "Oficinas"
    lnRep = 5
    
    SQL = "SELECT   CCODOFI, cAgeDescripcion as cDescripcion,"
    SQL = SQL & "   NNUMVEN, NVENCMN, NVENCME, NTOTVEN, ROUND((NTOTVEN/NTOTALSAL)*100,4) AS [% MORA.SUBTOTAL], ROUND((NTOTVEN/NTOTALCART)*100,4) AS [% MORA.CAR.TOTAL],"
    SQL = SQL & "   NNUMREF, NREFMN, NREFME, NTOTREF, ROUND((NTOTREF/NTOTALSAL)*100,4) AS [% REF.SUBTOTAL], ROUND((NTOTREF/NTOTALCART)*100,4) AS [% REF.CAR.TOTAL],"
    SQL = SQL & "   NNUMCRED , NSALDOMN, NSALDOME, NTOTALSAL"
    SQL = SQL & "   From"
    SQL = SQL & "   (SELECT CCODOFI,"
    SQL = SQL & "           SUM(nNumVenc + nNumJUD) AS NNUMVEN,"
    SQL = SQL & "           SUM(NVENCMN+NJUDMN) AS NVENCMN, SUM(NVENCME+NJUDME) AS NVENCME, SUM(NVENCMN+NJUDMN+ NVENCME+NJUDME) AS NTOTVEN,"
    SQL = SQL & "           SUM(nNumRef) AS NNUMREF, SUM(NREFMN) AS NREFMN, SUM(NREFME) AS NREFME, SUM(NREFMN+ NREFME) AS NTOTREF,"
    SQL = SQL & "           COUNT(CCTACOD) AS NNUMCRED, SUM(NSALDOMN) AS NSALDOMN, SUM(NSALDOME) AS NSALDOME,"
    SQL = SQL & "           SUM(NSALDOMN + NSALDOME) AS NTOTALSAL,"
    SQL = SQL & "           NTOTALCART = (SELECT SUM(NSALDOMN+NSALDOME) FROM " & lsTablaTMP & ")"
    SQL = SQL & "     From " & lsTablaTMP
    SQL = SQL & "     GROUP BY CCODOFI) AS X"
    SQL = SQL & "     JOIN AGENCIAS A ON A.cAgeCod = X.CCODOFI"
    SQL = SQL & "  ORDER BY CCODOFI"

End If

If optRep(6).value = True Then 'por lineas de financiamiento
    lTitulo = "MOROSIDAD Y REFINANCIADOS POR FUENTES DE FINANCIAMIENTO :" & lblFecha
    lsNomHoja = "RECURSOS"
    lnRep = 6
    
    '--- POR FUENTE DE FINANCIAMIENTO
    SQL = "SELECT  cRecurso, "
    SQL = SQL & "  cDescripcion = (    SELECT  DISTINCT P.CPERSNOMBRE"
    SQL = SQL & "                       FROM    COLOCLINEACREDITO C"
    SQL = SQL & "                               JOIN PERSONA P ON P.CPERSCOD = C.CPERSCOD"
    SQL = SQL & "                        WHERE   LEN(C.CLINEACRED) =2 and LEFT(C.CLINEACRED,2)= X.cRecurso ),"
    SQL = SQL & "       NNUMVEN, NVENCMN, NVENCME, NTOTVEN, ROUND((NTOTVEN/NTOTALSAL)*100,4) AS [% MORA.SUBTOTAL], ROUND((NTOTVEN/NTOTALCART)*100,4) AS [% MORA.CAR.TOTAL], "
    SQL = SQL & "       NNUMREF, NREFMN, NREFME, NTOTREF, ROUND((NTOTREF/NTOTALSAL)*100,4) AS [% REF.SUBTOTAL], ROUND((NTOTREF/NTOTALCART)*100,4) AS [% REF.CAR.TOTAL], "
    SQL = SQL & "       NNUMCRED , NSALDOMN, NSALDOME, NTOTALSAL"
    SQL = SQL & "   From "
    SQL = SQL & "           ("
    SQL = SQL & "               SELECT  cRecurso ,"
    SQL = SQL & "                       SUM(nNumVenc + nNumJUD) AS NNUMVEN,"
    SQL = SQL & "                       SUM(NVENCMN+NJUDMN) AS NVENCMN, SUM(NVENCME+NJUDME) AS NVENCME, SUM(NVENCMN+NJUDMN+ NVENCME+NJUDME) AS NTOTVEN,"
    SQL = SQL & "                       SUM(nNumRef) AS NNUMREF, SUM(NREFMN) AS NREFMN, SUM(NREFME) AS NREFME, SUM(NREFMN+ NREFME) AS NTOTREF,"
    SQL = SQL & "                       COUNT(CCTACOD) AS NNUMCRED, SUM(NSALDOMN) AS NSALDOMN, SUM(NSALDOME) AS NSALDOME,"
    SQL = SQL & "                       SUM(NSALDOMN + NSALDOME) AS NTOTALSAL,"
    SQL = SQL & "                       NTOTALCART = (SELECT SUM(NSALDOMN+NSALDOME) FROM " & lsTablaTMP & ")"
    SQL = SQL & "               From " & lsTablaTMP
    SQL = SQL & "               GROUP BY cRecurso"
    SQL = SQL & "           ) AS X"
    SQL = SQL & "           ORDER BY cRecurso"

End If

If optRep(7).value = True Then 'por analista
    lTitulo = "MOROSIDAD Y REFINANCIADOS POR ANALISTA :" & lblFecha
    lsNomHoja = "ANALISTA"
    lnRep = 7

    SQL = "         SELECT  cCodAnalista AS cDescripcion, "
    SQL = SQL & "           NNUMREF, NREFMN, NREFME, NTOTREF,"
    SQL = SQL & "           CASE WHEN NTOTALSAL>0 THEN ROUND((NTOTREF/NTOTALSAL)*100,4) ELSE 0 END AS [% REF.SUBTOTAL],"
    SQL = SQL & "           ROUND((NTOTREF/NTOTALCART)*100,4) AS [% REF.CAR.TOTAL],"
    SQL = SQL & "           NNUMVEN, NVENCMN, NVENCME, NTOTVEN,"
    SQL = SQL & "           CASE WHEN NTOTALSAL > 0 THEN ROUND((NTOTVEN/NTOTALSAL)*100,4) ELSE 0 END AS [% MORA.SUBTOTAL],"
    SQL = SQL & "           ROUND((NTOTVEN/NTOTALCART)*100,4) AS [% MORA.CAR.TOTAL],"
    SQL = SQL & "           NNUMCRED , NSALDOMN, NSALDOME, NTOTALSAL"
    SQL = SQL & "   From"
    SQL = SQL & "       ("
    SQL = SQL & "           SELECT  cCodAnalista ,"
    SQL = SQL & "                   SUM(nNumVenc + nNumJUD) AS NNUMVEN,"
    SQL = SQL & "                   SUM(NVENCMN+NJUDMN) AS NVENCMN, SUM(NVENCME+NJUDME) AS NVENCME, SUM(NVENCMN+NJUDMN+ NVENCME+NJUDME) AS NTOTVEN,"
    SQL = SQL & "                   SUM(nNumRef) AS NNUMREF, SUM(NREFMN) AS NREFMN, SUM(NREFME) AS NREFME, SUM(NREFMN+ NREFME) AS NTOTREF,"
    SQL = SQL & "                   COUNT(CCTACOD) AS NNUMCRED, SUM(NSALDOMN) AS NSALDOMN, SUM(NSALDOME) AS NSALDOME,"
    SQL = SQL & "                   SUM(NSALDOMN + NSALDOME) AS NTOTALSAL,"
    SQL = SQL & "                   NTOTALCART = (SELECT SUM(NSALDOMN+NSALDOME) FROM " & lsTablaTMP & ")"
    SQL = SQL & "           From " & lsTablaTMP
    SQL = SQL & "           GROUP BY cCodAnalista "
    SQL = SQL & "   ) AS X "
    SQL = SQL & " ORDER BY cCodAnalista"

End If


Set oCon = New DConecta

oCon.AbreConexion
Set rs = oCon.CargaRecordSet(SQL)
oCon.CierraConexion

If Not rs.EOF And Not rs.BOF Then
    vNHC = App.Path & "\spooler\EstadFinMes_" & Format(lblFecha, "yyyymmdd") & ".XLS"

    Set vExcelObj = New Excel.Application  '   = CreateObject("Excel.Application")
    vExcelObj.DisplayAlerts = False
    
    vExcelObj.Workbooks.Add
    vExcelObj.Sheets("Hoja1").Select
    vExcelObj.Sheets("Hoja1").Name = lsNomHoja
    
    vExcelObj.Range("A1:IV65536").Font.Name = "Arial Narrow"
    vExcelObj.Range("A1:IV65536").Font.Size = 8
    vExcelObj.Columns("A:IV").Select
    vExcelObj.Selection.VerticalAlignment = 3
    
    vExcelObj.Columns("A").Select
    vExcelObj.Selection.HorizontalAlignment = 1
    vExcelObj.Columns("B:H").Select
    vExcelObj.Selection.HorizontalAlignment = 1
    
    vExcelObj.Range("A2").Select
    vExcelObj.Range("A2").Font.Bold = True
    vExcelObj.Range("A2").HorizontalAlignment = 1
    vExcelObj.ActiveCell.value = UCase(Trim(gsNomCmac))
    
    vExcelObj.Range("M1").Select
    vExcelObj.Range("M1").Font.Bold = True
    vExcelObj.Range("M1").HorizontalAlignment = 1
    vExcelObj.ActiveCell.value = "Informacion al:" & Format(Me.lblFecha, "dd/mm/yyyy")
    
    vExcelObj.Range("A4").Select
    vExcelObj.Range("A4").Font.Bold = True
    vExcelObj.Range("A4").HorizontalAlignment = 1
    vExcelObj.Range("A4").Font.Size = 12
    vExcelObj.ActiveCell.value = lsTitulo
    
    vExcelObj.Range("M4").Select
    vExcelObj.Range("M4").Font.Bold = True
    vExcelObj.Range("M4").HorizontalAlignment = xlLeft
    vExcelObj.ActiveCell.value = "Tipo de Cambio : " & lnTC
    
    vExcelObj.Range("A6").Select
    vExcelObj.Range("A6").value = lsNomHoja
    
    vExcelObj.Range("A6:A7").Select
    vExcelObj.Range("A6:A7").Merge
    vExcelObj.Range("A6:A7").HorizontalAlignment = xlCenter
    vExcelObj.Range("A6:A7").Font.Bold = True
    
    vExcelObj.Range("B6").Select
    vExcelObj.Range("B6").value = "CREDITOS EN MORA"
    
    vExcelObj.Range("B6:E6").Merge
    vExcelObj.Range("B6:E6").HorizontalAlignment = xlCenter
    vExcelObj.Range("B6:E6").Font.Bold = True
    
    vExcelObj.Range("B7").Select
    vExcelObj.Range("B7").value = "CANT"
    vExcelObj.Range("B7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("B7").Font.Bold = True
    
    vExcelObj.Range("C7").Select
    vExcelObj.Range("C7").value = "M.N."
    vExcelObj.Range("C7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("C7").Font.Bold = True
    
    vExcelObj.Range("D7").Select
    vExcelObj.Range("D7").value = "M.E."
    vExcelObj.Range("D7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("D7").Font.Bold = True
    
    vExcelObj.Range("E7").Select
    vExcelObj.Range("E7").value = "TOTAL"
    vExcelObj.Range("E7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("E7").Font.Bold = True
    
    vExcelObj.Range("F6").Select
    vExcelObj.Range("F6").value = "%.MORA.SUB.TOTAL"
    
    vExcelObj.Range("F6:F7").Merge
    vExcelObj.Range("F6:F7").HorizontalAlignment = xlCenter
    vExcelObj.Range("F6:F7").Font.Bold = True
    
    vExcelObj.Range("G6").Select
    vExcelObj.Range("G6").value = "% MORA.TOTAL.CART."
    
    vExcelObj.Range("G6:G7").Merge
    vExcelObj.Range("G6:G7").HorizontalAlignment = xlCenter
    vExcelObj.Range("G6:G7").Font.Bold = True
    
    vExcelObj.Range("H6").Select
    vExcelObj.Range("H6").value = "CREDITOS REFINANCIADOS"
    
    vExcelObj.Range("H6:K6").Merge
    vExcelObj.Range("H6:K6").HorizontalAlignment = xlCenter
    vExcelObj.Range("H6:K6").Font.Bold = True
    
    vExcelObj.Range("H7").Select
    vExcelObj.Range("H7").value = "CANT"
    vExcelObj.Range("H7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("H7").Font.Bold = True
    
    vExcelObj.Range("I7").Select
    vExcelObj.Range("I7").value = "M.N."
    vExcelObj.Range("I7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("I7").Font.Bold = True
    
    vExcelObj.Range("J7").Select
    vExcelObj.Range("J7").value = "M.E."
    vExcelObj.Range("J7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("J7").Font.Bold = True
    
    vExcelObj.Range("K7").Select
    vExcelObj.Range("K7").value = "TOTAL"
    vExcelObj.Range("K7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("K7").Font.Bold = True
    
    vExcelObj.Range("L6").Select
    vExcelObj.Range("L6").value = "% REF"
    
    vExcelObj.Range("L6:L7").Merge
    vExcelObj.Range("L6:L7").HorizontalAlignment = xlCenter
    vExcelObj.Range("L6:L7").Font.Bold = True
    
    vExcelObj.Range("M6").Select
    vExcelObj.Range("M6").value = "% REF CARTERA TOTAL"
    
    vExcelObj.Range("M6:M7").Merge
    vExcelObj.Range("M6:M7").HorizontalAlignment = xlCenter
    vExcelObj.Range("M6:M7").Font.Bold = True
    
    
    vExcelObj.Range("N6").Select
    vExcelObj.Range("N6").value = "TOTAL CARTERA"
    
    vExcelObj.Range("N6:Q6").Merge
    vExcelObj.Range("N6:Q6").HorizontalAlignment = xlCenter
    vExcelObj.Range("N6:Q6").Font.Bold = True
    
    vExcelObj.Range("N7").Select
    vExcelObj.Range("N7").value = "CANT"
    vExcelObj.Range("N7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("N7").Font.Bold = True
    
    vExcelObj.Range("O7").Select
    vExcelObj.Range("O7").value = "M.N."
    vExcelObj.Range("O7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("O7").Font.Bold = True
    
    vExcelObj.Range("P7").Select
    vExcelObj.Range("P7").value = "M.E."
    vExcelObj.Range("P7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("P7").Font.Bold = True
    
    vExcelObj.Range("Q7").Select
    vExcelObj.Range("Q7").value = "TOTAL"
    vExcelObj.Range("Q7").HorizontalAlignment = xlHAlignCenter
    vExcelObj.Range("Q7").Font.Bold = True
    
    Dim vCel As String
    
    Dim vItem As Long
    Dim vIni As Long
    vIni = 7
    vItem = vIni
    Do While Not rs.EOF
         vItem = vItem + 1
    
         vCel = "A" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = "'" + rs!cDescripcion
         
         If lnRep <> 2 And lnRep <> 4 Then
            vCel = "B" + Trim(str(vItem))
            vExcelObj.Range(vCel).Select
            vExcelObj.ActiveCell.value = Format(rs!NNUMVEN, "#,#0")
        End If
         
         vCel = "C" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!NVENCMN, "#,#0.00")
         
         vCel = "D" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!NVENCME, "#,#0.00")
         
         vCel = "E" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.Range(vCel).Formula = "=SUM(C" & vItem & ":D" & vItem & ")"
         
         If lnRep <> 2 And lnRep <> 4 Then
         vCel = "F" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.Range(vCel).Formula = Format(rs("% MORA.SUBTOTAL"), "#,#0.0000")
         End If
         
         vCel = "G" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.Range(vCel).Formula = Format(rs("% MORA.CAR.TOTAL"), "#,#0.0000")
         
         If lnRep <> 2 And lnRep <> 4 Then
         vCel = "H" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!NNUMREF, "#,#0")
         End If
         vCel = "I" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!NREFMN, "#,#0.00")
         
         vCel = "J" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!NREFME, "#,#0.00")
         
         vCel = "K" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.Range(vCel).Formula = "=SUM(I" & vItem & ":J" & vItem & ")"
         If lnRep <> 2 And lnRep <> 4 Then
         vCel = "L" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.Range(vCel).Formula = Format(rs("% REF.SUBTOTAL"), "#,#0.0000")
         End If
         vCel = "M" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.Range(vCel).Formula = Format(rs("% REF.CAR.TOTAL"), "#,#0.0000")
         
         If lnRep <> 2 And lnRep <> 4 Then
         vCel = "N" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!NNUMCRED, "#,#0")
         End If
         vCel = "O" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!NSALDOMN, "#,#0.00")
         
         vCel = "P" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.ActiveCell.value = Format(rs!NSALDOME, "#,#0.00")
         
         vCel = "Q" + Trim(str(vItem))
         vExcelObj.Range(vCel).Select
         vExcelObj.Range(vCel).Formula = "=SUM(O" & vItem & ":P" & vItem & ")"
         
        rs.MoveNext
    Loop
    
    vCel = "A" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.ActiveCell.value = "TOTAL"
    
    vCel = "B" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(B" & vIni + 1 & ":B" & vItem & ")"
    
    vCel = "C" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(C" & vIni + 1 & ":C" & vItem & ")"
    
    vCel = "D" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(D" & vIni + 1 & ":D" & vItem & ")"
    
    vCel = "E" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(E" & vIni + 1 & ":E" & vItem & ")"
    
    vCel = "F" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(F" & vIni + 1 & ":F" & vItem & ")"
    
    vCel = "G" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(G" & vIni + 1 & ":G" & vItem & ")"
    
    vCel = "H" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(H" & vIni + 1 & ":H" & vItem & ")"
    
    vCel = "I" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(I" & vIni + 1 & ":I" & vItem & ")"
    
    vCel = "J" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(J" & vIni + 1 & ":J" & vItem & ")"
    
    vCel = "K" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(K" & vIni + 1 & ":K" & vItem & ")"
    
    vCel = "L" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(L" & vIni + 1 & ":L" & vItem & ")"
    
    vCel = "M" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(M" & vIni + 1 & ":M" & vItem & ")"
    
    vCel = "N" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(N" & vIni + 1 & ":N" & vItem & ")"
    
    vCel = "O" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(O" & vIni + 1 & ":O" & vItem & ")"
    
    vCel = "P" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(P" & vIni + 1 & ":P" & vItem & ")"
    
    vCel = "Q" + Trim(str(vItem + 1))
    vExcelObj.Range(vCel).Select
    vExcelObj.Range(vCel).Font.Bold = True
    vExcelObj.Range(vCel).Formula = "=SUM(Q" & vIni + 1 & ":Q" & vItem & ")"
    
    vCel = "A6:Q" & vItem + 1  'lsCeldaIni & ":" & lsCeldaFin
    vExcelObj.Range(vCel).BorderAround xlContinuous, xlThin
    vExcelObj.Range(vCel).Borders(xlInsideVertical).LineStyle = xlContinuous
    vExcelObj.Range(vCel).Borders(xlInsideHorizontal).LineStyle = xlContinuous
    
    vExcelObj.Range("A1").Select
    vExcelObj.ActiveWorkbook.SaveAs (vNHC)
    vExcelObj.ActiveWorkbook.Close
    MsgBox "SE HA GENERADO CON ÉXITO EL ARCHIVO !!  ", vbInformation, " Mensaje del Sistema ..."
    vExcelObj.Workbooks.Open (vNHC)
    vExcelObj.Visible = True
    
    Set vExcelObj = Nothing
    
End If

rs.Close
Set rs = Nothing

End Sub


Private Sub txtTC_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTC, KeyAscii, 7, 3)
If KeyAscii = 13 Then
    Me.cmdProcesar.SetFocus
End If
End Sub

