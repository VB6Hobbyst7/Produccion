VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Object = "{2E37A9B8-9906-4590-9F3A-E67DB58122F5}#7.0#0"; "OcxLabelX.ocx"
Begin VB.Form frmCredPagoConvenioBcoNac 
   Caption         =   "Pago por Corresponsalia"
   ClientHeight    =   6180
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10560
   Icon            =   "frmCredPagoConvenioBcoNac.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6180
   ScaleWidth      =   10560
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   765
      Left            =   120
      TabIndex        =   1
      Top             =   5400
      Width           =   10380
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   32
         Top             =   240
         Width           =   975
      End
      Begin VB.CommandButton CmdCargaArch 
         Caption         =   "&Subir Archivo"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton cmdLlenar 
         Caption         =   "&Mostrar Datos"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   22
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   9120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7920
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir Pago"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6600
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   30
      Left            =   150
      TabIndex        =   0
      Top             =   200
      Visible         =   0   'False
      Width           =   700
      _ExtentX        =   1244
      _ExtentY        =   53
      Filtro          =   "Archivos de Texto (*.pagos)|*.pagos|Archivos de Texto (*.cobros)|*.cobros"
      Altura          =   0
   End
   Begin TabDlg.SSTab SSTAB 
      Height          =   5175
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Pagados"
      TabPicture(0)   =   "frmCredPagoConvenioBcoNac.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl4"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPdolar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lblPSoles"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbltotal"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Mshbco"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbldolar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblsoles"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lblfec"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "lblreg"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Datos No Pagados"
      TabPicture(1)   =   "frmCredPagoConvenioBcoNac.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1"
      Tab(1).Control(1)=   "Label2"
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(3)=   "Label4"
      Tab(1).Control(4)=   "Mshbco1"
      Tab(1).Control(5)=   "lblimpdx"
      Tab(1).Control(6)=   "lblimpsx"
      Tab(1).Control(7)=   "lblfecx"
      Tab(1).Control(8)=   "lblnumx"
      Tab(1).ControlCount=   9
      Begin OcxLabelX.LabelX lblreg 
         Height          =   495
         Left            =   4200
         TabIndex        =   7
         Top             =   4560
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblfec 
         Height          =   495
         Left            =   3360
         TabIndex        =   8
         Top             =   360
         Visible         =   0   'False
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblsoles 
         Height          =   495
         Left            =   6240
         TabIndex        =   9
         Top             =   360
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lbldolar 
         Height          =   495
         Left            =   8760
         TabIndex        =   10
         Top             =   360
         Visible         =   0   'False
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblnumx 
         Height          =   495
         Left            =   -73200
         TabIndex        =   15
         Top             =   4560
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblfecx 
         Height          =   495
         Left            =   -71040
         TabIndex        =   16
         Top             =   4560
         Visible         =   0   'False
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblimpsx 
         Height          =   495
         Left            =   -68640
         TabIndex        =   17
         Top             =   4560
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblimpdx 
         Height          =   495
         Left            =   -66120
         TabIndex        =   18
         Top             =   4560
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Mshbco 
         Height          =   3495
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   10095
         _ExtentX        =   17806
         _ExtentY        =   6165
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Mshbco1 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   25
         Top             =   600
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6588
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin OcxLabelX.LabelX lbltotal 
         Height          =   495
         Left            =   1080
         TabIndex        =   26
         Top             =   360
         Visible         =   0   'False
         Width           =   945
         _ExtentX        =   1667
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblPSoles 
         Height          =   495
         Left            =   6480
         TabIndex        =   28
         Top             =   4560
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblPdolar 
         Height          =   495
         Left            =   9000
         TabIndex        =   29
         Top             =   4560
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   873
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Imp Dolares :"
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
         Height          =   195
         Left            =   7800
         TabIndex        =   31
         Top             =   4680
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Imp Soles :"
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
         Height          =   195
         Left            =   5400
         TabIndex        =   30
         Top             =   4680
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Num. Reg :"
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
         Height          =   195
         Left            =   120
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label4 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -72120
         TabIndex        =   21
         Top             =   4680
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Soles :"
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
         Height          =   195
         Left            =   -69360
         TabIndex        =   20
         Top             =   4680
         Visible         =   0   'False
         Width           =   600
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Dolares :"
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
         Height          =   195
         Left            =   -67200
         TabIndex        =   19
         Top             =   4680
         Visible         =   0   'False
         Width           =   780
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   -74760
         TabIndex        =   14
         Top             =   4680
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label lbl2 
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
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   2280
         TabIndex        =   13
         Top             =   480
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "Imp Soles :"
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
         Height          =   195
         Left            =   5160
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   960
      End
      Begin VB.Label lbl4 
         AutoSize        =   -1  'True
         Caption         =   "Imp Dolares :"
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
         Height          =   195
         Left            =   7560
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   1140
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Procesados :"
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
         Height          =   195
         Left            =   2880
         TabIndex        =   6
         Top             =   4680
         Visible         =   0   'False
         Width           =   1125
      End
   End
End
Attribute VB_Name = "frmCredPagoConvenioBcoNac"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim oBarra  As clsProgressBar
Dim fsNomFile As String
Dim fsPathFile As String
Dim fsruta As String
Dim Cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rsDocs As ADODB.Recordset

'Variables para utilizar la estructura anterior
Dim Datos As cobros54
Dim Datos1 As cobros60
Dim Datos4 As cobros234
Dim Datos6 As cobros236
'Variables para el archivo de los datos de contacto y temporal
Dim FileFree As Integer
Dim FileTemp As Integer
'Variables para la posición del primer y último registro
Dim RegActual As Long
Dim RegUltimo As Long
Dim RegCan As Long
Dim procesados As Integer
Dim n_procesados As Integer
'Variable para la posición Temporal del registro
Dim RegActualTemp As Long
Dim Pos As Integer, P As Integer
Dim cod_det As String

Private nNroTransac As Long
Private bCalenDinamic As Boolean
Private bCalenCuotaLibre As Boolean
Private bRecepcionCmact As Boolean
Private sPersCmac As String
Private vnIntPendiente As Double
Private vnIntPendientePagado As Double
Dim nCalPago As Integer
Dim bDistrib As Boolean
Dim bPrepago As Integer
Dim nCalendDinamTipo As Integer
Dim nMivivienda As Integer
Dim MatDatos As Variant
Dim sOperacion As String
Dim sPersCod As String
Dim nInteresDesagio As Double
Dim psCtaCod As String
Dim nCalendDinamico As Integer

Dim nMontoPago As Double
Dim lblMetLiq As String
Dim nITF As Double
Dim LblMoneda As String
Dim LblDiasAtraso As Integer
Dim dProxFec As Date
Dim LblProxfec As String

Dim sCadImpre As String
Dim bExoneradaLavado As Boolean
Dim sPerscodLav As String
Dim sNombreLav As String
Dim sDireccionLav As String
Dim sDocIdLav As String

Private bOperacionEfectivo As Boolean
Private nMontoLavDinero As Double
Private nTC As Double

Dim LblMonCalDin As Double
Dim tipoArch As Integer

Dim nPrestamo As Double
Dim bCuotaCom As Integer

'Archivo cabecera
 Dim f As Integer
 Dim fd As String
 Dim p1 As String
 Dim p3 As String
 Dim p2 As Integer
 Dim lineas As Long
 Dim str_Linea As String

Dim bRFA As Boolean
Dim nDiasAtras As Integer

Private oCredito As COMNCredito.NCOMCredito

Private Type cobros54
    cCodBco As String * 3
    cCodCli As String * 4
    nNumReg As String * 7
    nImpToTS As String * 15
    nImpToTD As String * 15
    fProceso As String * 8
    tRegistro As String * 2
End Type

Private Type cobros60
    nNumReg As String * 4
    nNumPag As String * 13
    nNumDec As String * 2
    fProceso As String * 8
End Type

Private Type cobros234
    cCodCta As String * 12
    nNroCta As String * 4
    cSituacion As String * 2
    nmoneda As String * 1
    cPatCliente As String * 20
    cMatCliente As String * 20
    cNomCliente As String * 20
    nImporteCuota As String * 15
    dCaducidad As String * 8
    nIndTasa As String * 1
    nfactorMora As String * 15
    nFactorCompensacion As String * 15
    nImporteGastos As String * 15
    cCuentaCliente As String * 11
    cOrdenCobro As String * 12
    nMora As String * 15
    nCompensacion As String * 15
    nImporteCobrado As String * 15
    CAgencia As String * 4
    dFechaCobro As String * 8
    dHoraCobro As String * 6
End Type

Private Type cobros236
    cCodCta As String * 12
    nNroCuota As String * 4
    nmoneda As String * 3
    dProceso As String * 8
    nNumPag As String * 13
    nNumDec As String * 2
    cNomCliente As String * 30
End Type
'EJVG20131128 ***
Private Enum TipoPagoBcoNac
    Corresponsalia = 1
    Cobros = 2
End Enum
Dim fnCobrosPorcMN As Currency
Dim fnCobrosPorcME As Currency
Dim fnCobrosMontoMinMN As Currency
Dim fnCobrosMontoMinME As Currency
Dim fnCobrosMontoMaxMN As Currency
Dim fnCobrosMontoMaxME As Currency
'END EJVG *******

Private Sub cmdCancelar_Click()
ConfigurarMShComite
ConfigurarMShComite1
sCadImpre = ""
If Me.SSTAB.TabVisible(1) Then
    activar_lbl1 (False)
    Me.lblimpdx.Caption = ""
    Me.lblimpsx.Caption = ""
    Me.lblfecx.Caption = ""
    Me.lblnumx.Caption = ""
    Me.SSTAB.TabVisible(1) = False
End If
procesados = 0
n_procesados = 0
Me.CmdCargaArch.Enabled = True
Me.CmdImprimir.Enabled = False
Me.cmdLlenar.Enabled = False
Me.cmdReporte.Enabled = False
'Labeles procesados
activar_lbl (False)
Me.lblsoles.Caption = ""
        Me.lbldolar.Caption = ""
        lblfec.Caption = ""
        Me.lblreg.Caption = ""
        Me.lblTotal.Caption = ""
        Me.lblPSoles.Caption = ""
        Me.lblPdolar.Caption = ""
End Sub

Private Sub CmdCargaArch_Click()
CdlgFile.nHwd = Me.hwnd
CdlgFile.Filtro = "Archivos Convenio (*.cobros)|*.cobros|Archivos Corresponsalia (*.pagos)|*.pagos"
Me.CdlgFile.altura = 400
CdlgFile.Show

fsPathFile = CdlgFile.Ruta
fsruta = fsPathFile
        If fsPathFile <> Empty Then
            For i = Len(fsPathFile) - 1 To 1 Step -1
                    If Mid(fsPathFile, i, 1) = "\" Then
                        fsPathFile = Mid(CdlgFile.Ruta, 1, i)
                        fsNomFile = Mid(CdlgFile.Ruta, i + 1, Len(CdlgFile.Ruta) - i)
                        Exit For
                    End If
             Next i
          Screen.MousePointer = 11
          
          If Right(fsNomFile, 6) = "COBROS" Then
            Leer_Lineas (fsruta)
            tipoArch = 1
          Else
            Leer_Lineas1 (fsruta)
            tipoArch = 2
          End If
          
        Else
           MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
           Exit Sub
        End If
    Screen.MousePointer = 0
End Sub

Sub cargar_datos()
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsX = oGen.DevolverDatos_Cabecera_108121(cod_det)
    Set oGen = Nothing

     If Not rsX.EOF And Not rsX.BOF Then
        activar_lbl (True)
        Me.lblsoles.Caption = rsX("nImpToTS")
        Me.lbldolar.Caption = rsX("nImpToTD")
        lblfec.Caption = rsX("fProceso")
        Me.lblreg.Caption = procesados
        Me.lblTotal.Caption = rsX("nNumReg")
        Me.lblPSoles.Caption = rsX("nImpToTS1")
        Me.lblPdolar.Caption = rsX("nImpToTD1")
     End If
    
    rsX.Close
    Set rsX = Nothing
End Sub

Sub cargar_datos1()
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsX = oGen.DevolverDatos_Cabecera_108121_NoProcesados(cod_det)
    Set oGen = Nothing

     If Not rsX.EOF And Not rsX.BOF Then
        activar_lbl1 (True)
        Me.lblimpdx.Caption = rsX("nImpToTD")
        Me.lblimpsx.Caption = rsX("nImpToTS")
        Me.lblfecx.Caption = rsX("fProceso")
        Me.lblnumx.Caption = n_procesados
     End If

    rsX.Close
    Set rsX = Nothing
End Sub

 Sub Cargar_grilla()
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim i As Integer
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsX = oGen.DevolverDatos_Detalle_108121(cod_det)
    Set oGen = Nothing
    i = 0

    ConfigurarMShComite
      Do Until rsX.EOF
        With Me.Mshbco
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsX!cCodCtafull
            .TextMatrix(.Rows - 2, 2) = rsX!cNombreCliente
            .TextMatrix(.Rows - 2, 3) = rsX!nNroCuota
            .TextMatrix(.Rows - 2, 4) = rsX!nImporteCuota
            .TextMatrix(.Rows - 2, 5) = rsX!nImporteCobrado
        End With
        rsX.MoveNext
    Loop
  Me.CmdImprimir.Enabled = True
  Me.cmdReporte.Enabled = True
  
  rsX.Close
  Set rsX = Nothing
End Sub

Sub Cargar_grilla1()
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim i As Integer
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsX = oGen.DevolverDatos_Detalle_108121_NoPagados(cod_det)
    Set oGen = Nothing
    i = 0

    ConfigurarMShComite1
      Do Until rsX.EOF
        With Me.Mshbco1
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsX!cCodCtafull
            .TextMatrix(.Rows - 2, 2) = rsX!cNombreCliente
            .TextMatrix(.Rows - 2, 3) = rsX!nNroCuota
            .TextMatrix(.Rows - 2, 4) = rsX!nImporteCuota
            .TextMatrix(.Rows - 2, 5) = rsX!nImporteCobrado
        End With
        rsX.MoveNext
    Loop
  Me.CmdImprimir.Enabled = True
  Me.cmdReporte.Enabled = True
  
  rsX.Close
  Set rsX = Nothing
End Sub

'Arreglar
Sub ConfigurarMShComite()
 Mshbco.Clear
    Mshbco.Cols = 6
    Mshbco.Rows = 2

    With Mshbco
        .TextMatrix(0, 0) = "Id"
        .TextMatrix(0, 1) = "Cuenta"
        .TextMatrix(0, 2) = "Nombre Cliente"
        .TextMatrix(0, 3) = "Cuota"
        .TextMatrix(0, 4) = "Imp.Cuota"
        .TextMatrix(0, 5) = "Imp.Cobrado"

        .ColWidth(0) = 400
        .ColWidth(1) = 1800
        .ColWidth(2) = 4500
        .ColWidth(3) = 700
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100

    End With
End Sub

Sub ConfigurarMShComite1()
 Mshbco1.Clear
    Mshbco1.Cols = 6
    Mshbco1.Rows = 2

    With Mshbco1
        .TextMatrix(0, 0) = "Id"
        .TextMatrix(0, 1) = "Cuenta"
        .TextMatrix(0, 2) = "Nombre Cliente"
        .TextMatrix(0, 3) = "Cuota"
        .TextMatrix(0, 4) = "Imp.Cuota"
        .TextMatrix(0, 5) = "Imp.Cobrado"

        .ColWidth(0) = 400
        .ColWidth(1) = 1800
        .ColWidth(2) = 4500
        .ColWidth(3) = 700
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100

    End With
End Sub

Sub activar_lbl(ByVal xb As Boolean)
    Me.lbl1.Visible = xb
    Me.Lbl2.Visible = xb
    Me.lbl3.Visible = xb
    Me.Label5.Visible = xb
    Me.Label6.Visible = xb
    Me.Label7.Visible = xb
    Me.lbl4.Visible = xb
    Me.lbldolar.Visible = xb
    Me.lblsoles.Visible = xb
    Me.lblPSoles.Visible = xb
    Me.lblPdolar.Visible = xb
    Me.lblTotal.Visible = xb
    Me.lblimpdx.Visible = xb
    Me.lblimpsx.Visible = xb
    Me.lblfec.Visible = xb
    Me.lblreg.Visible = xb
End Sub

Sub activar_lbl1(ByVal xb As Boolean)
    Me.SSTAB.TabVisible(1) = True
    Me.Label1.Visible = xb
    Me.Label2.Visible = xb
    Me.Label3.Visible = xb
    Me.Label4.Visible = xb
    Me.lblfecx.Visible = xb
    Me.lblimpdx.Visible = xb
    Me.lblimpsx.Visible = xb
    Me.lblnumx.Visible = xb
End Sub

Public Function LlenaRecordSet_Datos54() As Boolean
    Dim oCredD As COMDCredito.DCOMCreditos
    Dim fd1 As String
    Dim per As Boolean
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos
    Set rsDocs = New ADODB.Recordset
    Set oCredD = New COMDCredito.DCOMCreditos
    per = True

With rsDocs
    .Fields.Append "cCodBco", adVarChar, 3
    .Fields.Append "cCodCli", adVarChar, 4
    .Fields.Append "nNumReg", adVarChar, 7
    .Fields.Append "nImpToTS", adCurrency
    .Fields.Append "nImpToTD", adCurrency
    .Fields.Append "fProceso", adVarChar, 8
    .Fields.Append "tRegistro", adVarChar, 2
    .Open
        .AddNew
        .Fields("cCodBco") = Datos.cCodBco
        .Fields("cCodCli") = Datos.cCodCli
        .Fields("nNumReg") = Datos.nNumReg
        .Fields("nImpToTS") = Format(Mid(Datos.nImpToTS, 1, Len(Datos.nImpToTS) - 2), ".") & Right(Datos.nImpToTS, 2) 'Datos.nImpToTS
        .Fields("nImpToTD") = Format(Mid(Datos.nImpToTD, 1, Len(Datos.nImpToTD) - 2), ".") & Right(Datos.nImpToTD, 2) 'Datos.nImpToTD
        .Fields("fProceso") = Mid(Datos.fProceso, 5, 4) & Mid(Datos.fProceso, 3, 2) & Mid(Datos.fProceso, 1, 2)
        .Fields("tRegistro") = Datos.tRegistro
End With
           fd = rsDocs("fProceso")
           p1 = rsDocs("nImpToTS")
           p2 = rsDocs("nNumReg")
           p3 = rsDocs("nImpToTD")
           Set rsX = oGen.CargaDatos_Cabecera_108121(p1, p3, p2, fd)
           Set oGen = Nothing
        
           fd1 = Mid(fd, 7, 2) & "/" & Mid(fd, 5, 2) & "/" & Mid(fd, 1, 4)

           If Not CDate(fd1) = gdFecSis Then
                If CDate(fd1) > gdFecSis Then
                    per = False
                    MsgBox "No puede procesar el Archivo debido a la Fecha Posterior, Verifique!!!", vbInformation, "Aviso"
                ElseIf Not MsgBox("Seguro de Procesar un Archivo de Fecha Anterior", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                    per = False
                End If
           End If

           If per Then
                If rsX.EOF And rsX.BOF Then
                   nDiasAtras = DateDiff("d", CDate(fd1), gdFecSis)
                   Call oCredD.Insertar_Datos_Cabecera_BN(rsDocs, nDiasAtras)
                   LlenaRecordSet_Datos54 = True
                   procesados = 0
                   n_procesados = 0
                Else
                  MsgBox "El archivo ya fue Procesado con Anterioridad, Verifique!!!", vbInformation, "Aviso"
                  LlenaRecordSet_Datos54 = False
                End If
            Else
                   LlenaRecordSet_Datos54 = False
                   Exit Function
           End If
        rsX.Close
        Set rsX = Nothing
End Function

Public Function LlenaRecordSet_Datos60() As Boolean
    Dim oCredD As COMDCredito.DCOMCreditos
    Dim fd1 As String
    Dim per As Boolean
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos
    Set rsDocs = New ADODB.Recordset
    Set oCredD = New COMDCredito.DCOMCreditos
    per = True

With rsDocs
    .Fields.Append "nNumReg", adVarChar, 4
    .Fields.Append "nImpToTS", adCurrency
    .Fields.Append "fProceso", adVarChar, 8
    .Open
        .AddNew
        .Fields("nNumReg") = Datos1.nNumReg
        .Fields("nImpToTS") = CDbl(Datos1.nNumPag & "." & Datos1.nNumDec)
        .Fields("fProceso") = Datos1.fProceso
End With
           fd = rsDocs("fProceso")
           p1 = rsDocs("nImpToTS")
           p2 = rsDocs("nNumReg")
           Set rsX = oGen.CargaDatos_Cabecera_corresponsalia(p1, p2, fd)
          
           fd1 = Mid(fd, 7, 2) & "/" & Mid(fd, 5, 2) & "/" & Mid(fd, 1, 4)

           If Not CDate(fd1) = gdFecSis Then
                If CDate(fd1) > gdFecSis Then
                    per = False
                    MsgBox "El archivo No se puede procesar debido a la Fecha, Verifique!!!", vbInformation, "Aviso"
                ElseIf Not MsgBox("Esta Seguro de Procesar un Archivo de Fecha Anterior", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                    per = False
                End If
           End If

           If per Then
                If rsX.EOF And rsX.BOF Then
                   nDiasAtras = DateDiff("d", CDate(fd1), gdFecSis)
                   Call oCredD.Insertar_Datos_Cabecera_corresponsaliaBN(rsDocs, nDiasAtras)
                   LlenaRecordSet_Datos60 = True
                   Set rsX = oGen.CargaDatos_Cabecera_corresponsalia(p1, p2, fd)
                   cod_det = rsX("id") 'obtener el id de la cabecera
                   procesados = 0
                   n_procesados = 0
                Else
                  MsgBox "El archivo ya fue Procesado con Anterioridad, Verifique!!!", vbInformation, "Aviso"
                  LlenaRecordSet_Datos60 = False
                End If
            Else
                   LlenaRecordSet_Datos60 = False
                   Exit Function
           End If
        
        Set oGen = Nothing
        rsX.Close
        Set rsX = Nothing
End Function

Sub LlenaRecordSet_Datos234()
Dim oCredD As COMDCredito.DCOMCreditos
Set rsDocs = New ADODB.Recordset
Set oCredD = New COMDCredito.DCOMCreditos
With rsDocs
    'Crear RecordSet
    .Fields.Append "id_cab", adVarChar, 4
    .Fields.Append "cCodCta", adVarChar, 12
    .Fields.Append "cCodCtaF", adVarChar, 18
    .Fields.Append "nNroCta", adVarChar, 4
    .Fields.Append "cSituacion", adVarChar, 2
    .Fields.Append "nmoneda", adVarChar, 1
    .Fields.Append "cNomCliente", adVarChar, 60
    .Fields.Append "nImporteCuota", adCurrency 'adVarChar, 15
    .Fields.Append "dCaducidad", adVarChar, 8
    .Fields.Append "nIndTasa", adVarChar, 1
    .Fields.Append "nfactorMora", adCurrency 'adVarChar, 15
    .Fields.Append "nFactorCompensacion", adCurrency 'adVarChar, 15
    .Fields.Append "nImporteGastos", adCurrency 'adVarChar, 15
    .Fields.Append "cCuentaCliente", adVarChar, 11
    .Fields.Append "cOrdenCobro", adVarChar, 12
    .Fields.Append "nMora", adCurrency 'adVarChar, 15
    .Fields.Append "nCompensacion", adCurrency 'adVarChar, 15
    .Fields.Append "nImporteCobrado", adCurrency 'adVarChar, 15
    .Fields.Append "cAgencia", adVarChar, 4
    .Fields.Append "dFechaCobro", adVarChar, 8
    .Fields.Append "dHoraCobro", adVarChar, 6
    .Open

    'Llenar Recordset
        .AddNew
        .Fields("id_cab") = cod_det
        .Fields("cCodCta") = Datos4.cCodCta
        .Fields("nNroCta") = Datos4.nNroCta
        .Fields("cSituacion") = Datos4.cSituacion
        .Fields("nmoneda") = Datos4.nmoneda
        .Fields("nImporteCuota") = Format(Mid(Datos4.nImporteCuota, 1, Len(Datos4.nImporteCuota) - 2), ".") + Right(Datos4.nImporteCuota, 2)
        .Fields("dCaducidad") = Datos4.dCaducidad
        .Fields("nIndTasa") = Datos4.nIndTasa
        .Fields("nfactorMora") = Format(Mid(Datos4.nImporteCuota, 1, Len(Datos4.nImporteCuota) - 8), ".") + Right(Datos4.nfactorMora, 8)
        .Fields("nFactorCompensacion") = Format(Mid(Datos4.nFactorCompensacion, 1, Len(Datos4.nFactorCompensacion) - 8), ".") + Right(Datos4.nFactorCompensacion, 8)
        .Fields("nImporteGastos") = Format(Mid(Datos4.nImporteGastos, 1, Len(Datos4.nImporteGastos) - 2), ".") + Right(Datos4.nImporteGastos, 2)
        .Fields("cCuentaCliente") = Datos4.cCuentaCliente
        .Fields("cOrdenCobro") = Datos4.cOrdenCobro
        .Fields("nMora") = Format(Mid(Datos4.nMora, 1, Len(Datos4.nMora) - 2), ".") + Right(Datos4.nMora, 2)
        .Fields("nCompensacion") = Format(Mid(Datos4.nCompensacion, 1, Len(Datos4.nCompensacion) - 2), ".") + Right(Datos4.nCompensacion, 2)
        .Fields("nImporteCobrado") = Format(Mid(Datos4.nImporteCobrado, 1, Len(Datos4.nImporteCobrado) - 2), ".") + Right(Datos4.nImporteCobrado, 2)
        .Fields("cAgencia") = Datos4.CAgencia
        .Fields("dFechaCobro") = Datos4.dFechaCobro 'Mid(Datos4.dFechaCobro, 5, 4) & Mid(Datos4.dFechaCobro, 3, 2) & Mid(Datos4.dFechaCobro, 1, 2)
        .Fields("dHoraCobro") = Datos4.dHoraCobro
End With
        'completar ctacod
        Dim numctafull As String
        Dim nomcli As String
        Dim bCobraComision As Boolean
        nomcli = Trim(Datos4.cPatCliente) & "%" & Trim(Datos4.cMatCliente) & "%"
        numctafull = Datos4.cCodCta

        Dim rsX As ADODB.Recordset
        Dim rsy As New ADODB.Recordset
        Dim oGen As COMDCredito.DCOMCreditos

        Set rsX = New ADODB.Recordset
        Set rsy = New ADODB.Recordset
        Set oGen = New COMDCredito.DCOMCreditos
        
        Set rsX = oGen.DevolverDatos_previoCabecera_108121(numctafull, nomcli) 'obtengo los datos del cliente

       If Not rsX.EOF And Not rsX.BOF Then
            rsDocs.Fields("cCodCtaF") = rsX("cCtaCod")
            psCtaCod = rsX("cCtaCod")
            rsDocs.Fields("cNomCliente") = Trim(Mid(rsX("cPersNombre"), 1, 59))
      
            Call oCredD.Insertar_Datos_Detalle_BN(cod_det, rsDocs) 'inserto en detalle
            'para actualizar los Mov por cada operacion
            Set rsy = oGen.DevolverIDDatos_Detalle(cod_det, CInt(Datos4.nNroCta), psCtaCod, Datos4.dFechaCobro)
            If Not rsy.EOF And Not rsy.BOF Then
                Dim cod_dety As Double
                cod_dety = CDbl(rsy("Id"))
            End If
            bCobraComision = IIf(rsX("nING") = 1, True, False)
            'se cambio por Importe Cobrado, debido a pagos parciales
            If Not CargaDatos(cod_dety, psCtaCod, CDbl(rsDocs.Fields("nImporteCobrado")), Datos4.dFechaCobro, CInt(Datos4.nNroCta), bCobraComision, Cobros) Then
                   Dim ultm As Integer
                   n_procesados = n_procesados + 1
                   Call oCredD.Actualizar_DetallexEst_BN(cod_det, psCtaCod, Datos4.nNroCta, (Datos4.dFechaCobro), 1)
                   ultm = n_procesados + procesados
                      If n_procesados = p2 Then
                          MsgBox "El Pago no se Realizó" & vbCrLf & "Comuníquese con el Area de TI"
                      ElseIf ultm = p2 Then
                          MsgBox "El Proceso de Pago Finalizó correctamente pero hubieron registros NO procesados, Revise el Reporte!!"
                          Me.cmdLlenar.Enabled = True
                          Me.CmdCargaArch.Enabled = False
                      End If
             End If
       Else
             MsgBox "No se pudo leer los datos de : " & nomcli & " Este registro no será procesado automaticamente"
             n_procesados = n_procesados + 1
             ultm = n_procesados + procesados
             If n_procesados = p2 Then
                MsgBox "No se realizó el Pago de Ningún Registro" & vbCrLf & "Comuniquese con el Area de TI"
             ElseIf ultm = p2 Then
                MsgBox "El Proceso de Pago Finalizó Correctamente, pero hubieron registros NO procesados"
                Me.cmdLlenar.Enabled = True
                Me.CmdCargaArch.Enabled = False
            End If
       End If
     Set oGen = Nothing
 End Sub

'Private Function CargaDatos(ByVal cod_det As Double, ByVal psCtaCod As String, pnImporteCobrado As Double, pdFechaCobro As String, pnNroCta As Integer, Optional valor As Boolean = False) As Boolean
Private Function CargaDatos(ByVal cod_det As Double, ByVal psCtaCod As String, pnImporteCobrado As Double, pdFechaCobro As String, pnNroCta As Integer, Optional valor As Boolean = False, Optional ByVal pTpoPago As TipoPagoBcoNac = Corresponsalia) As Boolean
Dim oCredD As COMDCredito.DCOMCreditos
Set rsDocs = New ADODB.Recordset
Set oCredD = New COMDCredito.DCOMCreditos
Dim rsPers As ADODB.Recordset
Dim rsCredVig As ADODB.Recordset
Dim sAgencia As String
Dim nGastos As Double
Dim nMonPago As Double
Dim nMora As Double
Dim nCuotasMora As Integer
Dim nTotalDeuda As Currency
Dim nInteresDesagio As Double
Dim nMonCalDin As Double
Dim sMensaje As String
Dim cNomCliente As String
'ARCV
Dim nNewSalCap As Double
Dim nNewCPend As Integer
Dim dProxFec As Date
Dim sEstado As String
Dim bRFA As Boolean
'ARCV
Dim nCuotaPendiente As Integer
Dim nMoraCalculada As Double
Dim dFechaVencimiento As Date
Dim sImpreBoleta As String
Dim bValorProceso As Boolean
Dim nMonIntGra As Double

Dim oITF As COMDConstSistema.FCOMITF
Dim dif As Double
Dim nNroCta As Integer
Dim FechaCobro As String
Dim FechaCobro1 As Date
Dim NroMov As Long
Dim est As Boolean
Dim nTempITF As Double
Dim MontoPagosinITF As Double
'Corresponsalia
Dim Cuota As Double
Dim comision As Double
Dim bInstFinanc As Boolean 'JUEZ 20140411

    On Error GoTo ErrorCargaDatos
    nInteresDesagio = 0
    MontoPagosinITF = 0
    est = False
    Set oCredito = New COMNCredito.NCOMCredito
    If nDiasAtras > 0 Then
        Perdonar_Mora psCtaCod, nDiasAtras, CDate(Mid(pdFechaCobro, 1, 4) & "-" & Mid(pdFechaCobro, 5, 2) & "-" & Mid(pdFechaCobro, 7, 2))
    End If
    Call oCredito.CargaDatosPagoCuotas(psCtaCod, gdFecSis, bPrepago, gsCodAge, rsCredVig, sAgencia, nCalendDinamico, bCalenDinamic, bCalenCuotaLibre, _
                                    nMivivienda, nCalPago, nGastos, nMonPago, nMora, nCuotasMora, nTotalDeuda, nInteresDesagio, _
                                    nMonCalDin, sMensaje, sPersCod, sOperacion, bExoneradaLavado, bRFA, rsPers, bOperacionEfectivo, nMontoLavDinero, nTC, _
                                    nMontoPago, nITF, vnIntPendientePagado, nNewSalCap, nNewCPend, dProxFec, sEstado, nCuotaPendiente, nMoraCalculada, dFechaVencimiento)

    If Not rsCredVig.BOF And Not rsCredVig.EOF Then
        lblMetLiq = Trim(rsCredVig!cMetLiquidacion)
        nCalendDinamTipo = rsCredVig!nCalendDinamTipo
        cNomCliente = PstaNombre(rsCredVig!cPersNombre)
        dif = 0
        CargaDatos = True
        vnIntPendiente = IIf(IsNull(rsCredVig!nintPend), 0, rsCredVig!nintPend)
        vnIntPendientePagado = 0
        nNroTransac = IIf(IsNull(rsCredVig!nTransacc), 0, rsCredVig!nTransacc)
        nPrestamo = Format(rsCredVig!nMontoCol, "#0.00")
        bCuotaCom = IIf(IsNull(rsCredVig!bCuotaCom), 0, rsCredVig!bCuotaCom)
        LblMoneda = Trim(rsCredVig!cmoneda)
        LblDiasAtraso = CInt(rsCredVig!nDiasAtraso)
        LblMonCalDin = nMonCalDin
        
        'JUEZ 20140411 **************************************
        Dim oDInstFinan As COMDPersona.DCOMInstFinac
        Set oDInstFinan = New COMDPersona.DCOMInstFinac
        bInstFinanc = oDInstFinan.VerificaEsInstFinanc(rsCredVig!cPersCod)
        Set oDInstFinan = Nothing
        If bInstFinanc Then nITF = 0
        'END JUEZ *******************************************
        
        'INICIO ORCR20140714----------------------------------------------------------
        Dim pnImporteCobradoDesc As Double
        pnImporteCobradoDesc = pnImporteCobrado
        'FIN ORCR20140714-------------------------------------------------------------

        If dProxFec <> 0 Then LblProxfec = dProxFec
    
        If valor Then
            If pTpoPago = Corresponsalia Then
                If LblMoneda = "SOLES" Then
                    Cuota = pnImporteCobrado / 1.01
                    comision = Round(pnImporteCobrado - Cuota, 2)
                        If (comision < 2.5) Then
                            comision = 2.5
                        End If
                        If comision > 100 Then
                            comision = 100
                        End If
                Else
                    Cuota = pnImporteCobrado / 1.015
                    comision = Round(pnImporteCobrado - Cuota, 2)
                    If (comision < 2) Then
                        comision = 2
                    End If
                End If
            ElseIf pTpoPago = Cobros Then
                If LblMoneda = "SOLES" Then
                    Cuota = pnImporteCobrado / fnCobrosPorcMN
                    comision = Round(pnImporteCobrado - Cuota, 2)
                        If (comision < fnCobrosMontoMinMN) Then
                            comision = fnCobrosMontoMinMN
                        End If
                        If comision > fnCobrosMontoMaxMN Then
                            comision = fnCobrosMontoMaxMN
                        End If
                Else
                    Cuota = pnImporteCobrado / fnCobrosPorcME
                    comision = Round(pnImporteCobrado - Cuota, 2)
                    If (comision < fnCobrosMontoMinME) Then
                        comision = fnCobrosMontoMinME
                    End If
                    If comision > fnCobrosMontoMaxME Then
                        comision = fnCobrosMontoMaxME
                    End If
                End If
            End If
        Cuota = pnImporteCobrado - comision
        pnImporteCobrado = Cuota
        End If

        'verificar que el monto a pagar sea menor que el monto total de la deuda
         If Round(pnImporteCobrado, 2) > Round(nTotalDeuda, 2) Then
                dif = pnImporteCobrado - nTotalDeuda
                FechaCobro = pdFechaCobro
                nNroCta = pnNroCta
                est = True
                fd1 = Mid(FechaCobro, 7, 2) & "/" & Mid(FechaCobro, 5, 2) & "/" & Mid(FechaCobro, 1, 4)
                FechaCobro1 = CDate(fd1)
                pnImporteCobrado = nTotalDeuda
        End If
        
        If Round(nMontoPago + nITF, 2) <> Round(pnImporteCobrado, 2) Then
            Set oITF = New COMDConstSistema.FCOMITF
            oITF.fgITFParametros

            Dim pnITF As Double
            Dim nRedondeoITF As Double 'BRGO 20110914
            If Mid(psCtaCod, 6, 3) = "423" Then
                pnITF = 0
                nITF = 0
            Else
                Dim lnValor As Double
                lnValor = (pnImporteCobrado * oITF.gnITFPorcent / (100 + oITF.gnITFPorcent)) * 100
                lnValor = oITF.CortaDosITF(lnValor)
                pnITF = lnValor
                nITF = pnITF
            End If
            Set oITF = Nothing
        End If
        
        If bInstFinanc Then nITF = 0 'JUEZ 20140411
        
         '*** MADM 20120123 ************************************************
        nRedondeoITF = fgDiferenciaRedondeoITF(nITF)
        If nRedondeoITF > 0 Then
            nITF = Format(CCur(Format(nITF, "#0.00")) - nRedondeoITF, "#,##0.00")
        End If
        '*** MADM 20120123 ************************************************
        
        MontoPagosinITF = pnImporteCobrado - nITF
        
        If est Then
            MontoPagosinITF = pnImporteCobrado
        End If
        nTempITF = nITF
        
        'EJVG20130805 *** Exactitud decimales x Cancelación
        MontoPagosinITF = Round(MontoPagosinITF, 2)
        nTotalDeuda = Round(nTotalDeuda, 2)
        'END EJVG *******
        bValorProceso = oCredito.ActualizaMontoPago(CDbl(MontoPagosinITF), CDbl(nTotalDeuda), psCtaCod, gdFecSis, lblMetLiq, vnIntPendiente, vnIntPendientePagado, _
                                        bCalenCuotaLibre, bCalenDinamic, bPrepago, CDbl(MontoPagosinITF), LblMonCalDin, sMensaje, nITF, _
                                        nInteresDesagio, nNewSalCap, nNewCPend, dProxFec, sEstado, nMonIntGra)
        
        nITF = nTempITF
        nITF = Format(nITF, "#0.00")
        
        If bValorProceso Then
               Call oCredito.GrabarPagoCuotas(psCtaCod, nMivivienda, nCalPago, CDbl(MontoPagosinITF), _
                                gdFecSis, lblMetLiq, 1, gsCodAge, gsCodUser, gsCodCMAC, "", _
                                bRecepcionCmact, sPersCmac, vnIntPendiente, vnIntPendientePagado, bPrepago, "", nITF, _
                                nInteresDesagio, nTotalDeuda, bCalenDinamic, CDbl(LblMonCalDin), nCalendDinamTipo, gsNomAge, CInt(Mid(psCtaCod, 6, 3)), _
                                cNomCliente, LblMoneda, nNroTransac, LblProxfec, sLpt, gsInstCmac, False, "", _
                                "", "", sImpreBoleta, LblDiasAtraso, gsProyectoActual, gbImpTMU, , "", "", "", "", NroMov, , , , , , , nDiasAtras)
                                'Parametro nDiasAtras agregado por AMDO20130402

               sCadImpre = sCadImpre & sImpreBoleta
                            
               'Operacion de egreso
                Dim clsMov1 As COMDCaptaGenerales.DCOMCaptaMovimiento
                Dim ClsMov As COMNContabilidad.NCOMContFunciones
                Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
                Dim sMovNro As String
                Dim nMovNroR As Long
                Dim nMovNroC As Long
                Dim cod_ageContable As String
                Dim oTipoCambio As nTipoCambio
                Dim lsMovNro As String 'FRHU 20150318
                
                Set ClsMov = New COMNContabilidad.NCOMContFunciones
                sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Set ClsMov = Nothing 'FRHU 20150318
                
                Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
                clsServ.OtrasOpeDepCtaBanco sMovNro, "01.1090100824631.0110001", psCtaCod, MontoPagosinITF + nITF, cNomCliente, gsNomAge, IIf(LblMoneda = "SOLES", 1, 2), 300593
                Set clsServ = Nothing
                
               Set clsMov1 = New COMDCaptaGenerales.DCOMCaptaMovimiento
               nMovNroR = clsMov1.GetnMovNro(sMovNro)

              'Parte contable -----------------------------------------------------------------
               Set ClsMov = New COMNContabilidad.NCOMContFunciones 'FRHU 20150318
               lsMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
               Set ClsMov = Nothing


              'INICIO ORCR20140714-------------------------------------------------------------
              Dim psMovDesc As String
              'psMovDesc = psCtaCod & " " & cNomCliente & " " & pnNroCta & " - " & (MontoPagosinITF + nITF)
              'psMovDesc = Right(psCtaCod, 5) & "|" & pnNroCta & "|" & (MontoPagosinITF + nITF) & "|" & cNomCliente
              psMovDesc = Right(psCtaCod, 5) & "|" & pnNroCta & "|" & pnImporteCobradoDesc & "|" & cNomCliente
              
              '--------------------------------------------------------------------------------
              'If valor Then
              If pTpoPago = Corresponsalia Then 'EJVG20131128
                If LblMoneda = "SOLES" Then
                    cod_ageContable = "2918070101" & gsCodAge
                     'oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "401583", "Deposito Banco", "1113010201",
                     oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "401583", psMovDesc, "1113010201", _
                    cod_ageContable, CCur(MontoPagosinITF + nITF), 0, "", pdFechaCobro, "", _
                    1, "01.1090100824631.0110004", True, , , , , 0, , , , , , , , True, , , , , nMovNroR
                Else
                      cod_ageContable = "2928070101" & gsCodAge
                     'oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "402583", "Deposito Banco", "1123010201",
                     oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "402583", psMovDesc, "1123010201", _
                    cod_ageContable, CCur(MontoPagosinITF + nITF), 0, "", pdFechaCobro, "", _
                    1, "01.1090100824631.0120001", True, , , , , 0, , , , , , , , True, , , , , nMovNroR
                End If
              ElseIf pTpoPago = Cobros Then
                If LblMoneda = "SOLES" Then
                    cod_ageContable = "2918070101" & gsCodAge
                    'oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "401583", "Deposito Banco", "1113010201",
                    oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "401583", psMovDesc, "1113010201", _
                    cod_ageContable, CCur(MontoPagosinITF + nITF), 0, "", pdFechaCobro, "", _
                    1, "01.1090100824631.0110002", True, , , , , 0, , , , , , , , True, , , , , nMovNroR
                Else
                     cod_ageContable = "2928070101" & gsCodAge
                    'oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "402583", "Deposito Banco", "1123010201",
                    oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "402583", psMovDesc, "1123010201", _
                    cod_ageContable, CCur(MontoPagosinITF + nITF), 0, "", pdFechaCobro, "", _
                    1, "01.1090100824631.0120002", True, , , , , 0, , , , , , , , True, , , , , nMovNroR
                End If
              End If
              'FIN ORCR20140714--------------------------------------------------------------
              nMovNroC = clsMov1.GetnMovNro(lsMovNro)
               Set clsMov1 = Nothing

              'Guardar Movs (pago, reversion y contable) --------------------------------------
                If NroMov > 0 And nMovNroR > 0 Then
                    'If valor Then
                    If pTpoPago = Corresponsalia Then 'EJVG20131128
                        Call oCredD.Actualizar_DetallexEst_MovCorresponsalia(cod_det, NroMov, nMovNroR, nMovNroC, comision)
                    ElseIf pTpoPago = Cobros Then
                        Call oCredD.Actualizar_DetallexEst_Mov(cod_det, NroMov, nMovNroR, nMovNroC)
                    End If
                    If est = True Then
                        If (dif - nITF) > 0 Then
                            'If valor Then
                            If pTpoPago = Corresponsalia Then 'EJVG20131128
                                Call oCredD.Insertar_Datos_SobrantesCorresponsalia(psCtaCod, nNroCta, dif - nITF, FechaCobro1, cNomCliente, NroMov)
                            ElseIf pTpoPago = Cobros Then
                                Call oCredD.Insertar_Datos_Sobrantes(psCtaCod, nNroCta, dif - nITF, FechaCobro1, cNomCliente, NroMov)
                            End If
                        End If
                    End If
                End If
              '----------------------------------------------------------------------------------
               procesados = procesados + 1
               ultm = n_procesados + procesados
               If procesados = p2 Then
                    MsgBox "El Pago Finalizó correctamente" ' & vbCrLf & "Mensaje"
                    Me.cmdLlenar.Enabled = True
                    Me.CmdCargaArch.Enabled = False
              ElseIf ultm = p2 And n_procesados > 0 Then
                    MsgBox "El Pago Finalizó correctamente pero hubieron registros NO procesados, Revise el Reporte!!"
                    Me.cmdLlenar.Enabled = True
                    Me.CmdCargaArch.Enabled = False
               ElseIf ultm = p2 Then
                    MsgBox "El Proceso de Pago Finalizó correctamente" ' & vbCrLf & "Mensaje"
                    Me.cmdLlenar.Enabled = True
                    Me.CmdCargaArch.Enabled = False
               End If
               Set oCredito = Nothing
               Set oCredD = Nothing
        Else
             CargaDatos = False
        End If
    Else
        CargaDatos = False
    End If
    Exit Function

ErrorCargaDatos:
    MsgBox err.Description, vbCritical, "Aviso"
End Function

Public Function Contar_Lineas(ByVal strTextFile As String) As Long
    Dim f As Integer
    Dim lineas As Long
    
    f = FreeFile
    Open strTextFile For Input As #f
    Do
        Line Input #f, str_Linea
        lineas = lineas + 1
    Loop While Not EOF(f)
    Close #f
    Contar_Lineas = lineas
End Function

Public Sub Leer_Lineas(strTextFile As String)
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim nCerrar As Integer
    Set oBarra = New clsProgressBar
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    lineas = 0
    f = FreeFile
    nCerrar = 0
    Open strTextFile For Input As #f
    Do 'grabar por cada item el cuerpo
        Line Input #f, str_Linea
        lineas = lineas + 1

        If lineas = 1 And Len(str_Linea) <> 54 Then
                MsgBox "El archivo adjunto NO tiene la cabecera correcta", vbInformation, "Aviso"
                Close #f
                nCerrar = 0
                Exit Sub
        ElseIf lineas = 1 And Len(str_Linea) = 54 Then
           Datos.cCodBco = Mid(str_Linea, 1, 3)
           Datos.cCodCli = Mid(str_Linea, 4, 4)
           Datos.nNumReg = Mid(str_Linea, 8, 7)
           Datos.nImpToTS = Mid(str_Linea, 15, 15)
           Datos.nImpToTD = Mid(str_Linea, 30, 15)
           Datos.fProceso = Mid(str_Linea, 45, 8)
           Datos.tRegistro = Mid(str_Linea, 53, 2)
           nCerrar = 0
             'valida el numero de registros a procesar
             If Not Contar_Lineas(strTextFile) = CInt(Datos.nNumReg) + 1 Then
                 MsgBox "El archivo a procesar NO tiene Completos los registros, Solicite un nuevo archivo al Area de TI", vbInformation, "Aviso"
                  Set rsX = Nothing
                    nCerrar = 0
                    Close #f
                    Exit Do
                    Exit Sub
             End If
             
             If LlenaRecordSet_Datos54 = False Then
                    Set rsX = Nothing
                    nCerrar = 0
                    Close #f
                    Exit Do
                    Exit Sub
             End If
                   'ya lleno el recordset y consulto con el
                   fd = rsDocs("fProceso")
                   p1 = rsDocs("nImpToTS")
                   p2 = rsDocs("nNumReg")
                   p3 = rsDocs("nImpToTD")
                   RegCan = p2
                   nCerrar = 1
                   Set rsX = oGen.CargaDatos_Cabecera_108121(p1, p3, p2, fd)
                   Set oGen = Nothing

                   If Not rsX.EOF And Not rsX.BOF Then
                       cod_det = rsX("id") 'obtener el id de la cabecera
                       MsgBox "Se van a procesar " & RegCan & " registros, Esta operación puede demorar unos minutos ...", vbInformation, "Aviso"
                   Else
                        MsgBox "El archivo no tiene la estructura correcta", vbInformation, "Aviso"
                        Close #f
                        Exit Sub
                   End If

        Else
            If Len(str_Linea) = 234 Then
                If lineas = 2 Then
                    oBarra.ShowForm Me
                    oBarra.CaptionSyle = eCap_CaptionPercent
                    oBarra.Max = RegCan
                    oBarra.Progress 0, "Proceso de Pago Convenio BN : ", "Preparando Pago...", "Preparando", vbBlue
                Else
                   oBarra.Progress (lineas - 1), "Proceso de Pago Convenio BN : ", "Efectuando Pago...", "Efectuando", vbBlue
                End If

                With Datos4
                            .cCodCta = Mid(str_Linea, 1, 12)
                            .nNroCta = Mid(str_Linea, 13, 4)
                            .cSituacion = Mid(str_Linea, 17, 2)
                            .nmoneda = Mid(str_Linea, 19, 1)
                            .cPatCliente = Mid(str_Linea, 20, 20)
                            .cMatCliente = Mid(str_Linea, 40, 20)
                            .cNomCliente = Mid(str_Linea, 60, 20)
                            .nImporteCuota = Mid(str_Linea, 80, 15)
                            .dCaducidad = Mid(str_Linea, 95, 8)
                            .nIndTasa = Mid(str_Linea, 103, 1)
                            .nfactorMora = IIf(val(Mid(str_Linea, 104, 15)) = 0, "000000000000000", Mid(str_Linea, 104, 15))
                            .nFactorCompensacion = IIf(val(Mid(str_Linea, 119, 15)) = 0, "000000000000000", Mid(str_Linea, 119, 15))
                            .nImporteGastos = IIf(val(Mid(str_Linea, 134, 15)) = 0, "000000000000000", Mid(str_Linea, 134, 15))
                            .cCuentaCliente = Mid(str_Linea, 149, 11)
                            .cOrdenCobro = Mid(str_Linea, 160, 12)
                            .nMora = Mid(str_Linea, 172, 15)
                            .nCompensacion = Mid(str_Linea, 187, 15)
                            .nImporteCobrado = Mid(str_Linea, 202, 15)
                            .CAgencia = Mid(str_Linea, 217, 4)
                            .dFechaCobro = Mid(str_Linea, 221, 8)
                            .dHoraCobro = Mid(str_Linea, 229, 6)
                            Call LlenaRecordSet_Datos234 'cuerpo
                End With
            Else
                MsgBox "Los datos del archivo no tienen la estructura correcta", vbInformation, "Aviso"
                Close #f
                nCerrar = 0
                Exit Sub
            End If
        End If
    Loop While Not EOF(f)
    Close #f
    If nCerrar = 1 Then
        rsX.Close
    End If
    oBarra.CloseForm Me
    Set oBarra = Nothing
End Sub

Function validar_cola(numreg As Integer, filex As String) As Boolean
    Dim f As Integer
    Dim lineas As Long
    validar_cola = True
    
    f = FreeFile
    Open filex For Input As #f
    Do
        Line Input #f, str_Linea
        
        If lineas = (numreg - 1) Then
            
            If Trim(Mid(str_Linea, 13, 4)) = "" Or Trim(Mid(str_Linea, 29, 8)) = "" Or Trim(Mid(str_Linea, 59, 13)) = "" Or Trim(Mid(str_Linea, 72, 2)) = "" Then
                validar_cola = False
                Exit Function
            End If
            
            Datos1.nNumReg = Mid(str_Linea, 13, 4)
            Datos1.fProceso = Mid(str_Linea, 29, 8)
            Datos1.nNumPag = Mid(str_Linea, 59, 13)
            Datos1.nNumDec = Mid(str_Linea, 72, 2)
                        
            If Not LlenaRecordSet_Datos60 Then
                validar_cola = False
            End If
        End If
        lineas = lineas + 1
    Loop While Not EOF(f)
    Close #f
End Function

Public Sub Leer_Lineas1(strTextFile As String)
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim nCerrar As Integer
    Dim Tot_Corr As Integer
    Set oBarra = New clsProgressBar
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos
   
    lineas = 0
    f = FreeFile
    nCerrar = 0
    
    'valida la cola   'valida el numero de registros a procesar
    Tot_Corr = Contar_Lineas(strTextFile)
    
    If Not validar_cola(Tot_Corr, strTextFile) Then
         MsgBox "Los datos del archivo no tienen la estructura correcta", vbInformation, "Aviso"
         Close #f
         nCerrar = 0
         Exit Sub
    Else
        Open strTextFile For Input As #f
        Do 'grabar por cada item el cuerpo
            Line Input #f, str_Linea
            lineas = lineas + 1
            If lineas = 1 Then
               MsgBox "Se van a procesar " & Tot_Corr - 1 & " registros, Esta operación puede demorar unos minutos ...", vbInformation, "Aviso"
               oBarra.ShowForm Me
               oBarra.CaptionSyle = eCap_CaptionPercent
               oBarra.Max = Tot_Corr
               oBarra.Progress 0, "Proceso de Pago Corresponsalia BN : ", "Preparando Pago...", "Preparando", vbRed
           Else
               oBarra.Progress (lineas - 1), "Proceso de Pago Corresponsalia BN : ", "Efectuando Pago...", "Efectuando", vbRed
           End If
    
            If Not (Tot_Corr = lineas) Then
                If Mid(str_Linea, 1, 4) = "0085" And lineas <> Tot_Corr Then
                   Datos6.cCodCta = Mid(str_Linea, 5, 12)
                   Datos6.nNroCuota = Mid(str_Linea, 17, 4)
                   Datos6.dProceso = Mid(str_Linea, 29, 8)
                   Datos6.nmoneda = Mid(str_Linea, 57, 3)
                   Datos6.nNumPag = Mid(str_Linea, 60, 13)
                   Datos6.nNumDec = Mid(str_Linea, 73, 2)
                   Datos6.cNomCliente = Mid(str_Linea, 75, 30)
                   LlenaRecordSet_Datos236 'cuerpo
                Else
                    MsgBox "Los datos del archivo no tienen la estructura correcta", vbInformation, "Aviso"
                    Close #f
                    nCerrar = 0
                    Exit Sub
                End If
            End If
            Loop While Not EOF(f)
        Close #f
    End If
    If nCerrar = 1 Then
        rsX.Close
    End If
    oBarra.CloseForm Me
    Set oBarra = Nothing
End Sub

Sub LlenaRecordSet_Datos236()
Dim oCredD As COMDCredito.DCOMCreditos
Set rsDocs = New ADODB.Recordset
Set oCredD = New COMDCredito.DCOMCreditos

With rsDocs
    'Crear RecordSet
    .Fields.Append "id_cab", adVarChar, 4
    .Fields.Append "cCodCta", adVarChar, 12
    .Fields.Append "cCodCtaF", adVarChar, 18
    .Fields.Append "nNroCta", adInteger
    .Fields.Append "nmoneda", adInteger
    .Fields.Append "cNomCliente", adVarChar, 60
    .Fields.Append "nImporteCobrado", adCurrency
    .Fields.Append "dFechaCobro", adVarChar, 8
    .Open
        .AddNew
        .Fields("id_cab") = cod_det
        .Fields("cCodCta") = Datos6.cCodCta
        .Fields("nNroCta") = Datos6.nNroCuota
        .Fields("nmoneda") = IIf(Datos6.nmoneda = "SOL", 1, 2)
        .Fields("cNomCliente") = Trim(Datos6.cNomCliente)
        .Fields("nImporteCobrado") = CDbl((Datos6.nNumPag) + "." + (Datos6.nNumDec))
        .Fields("dFechaCobro") = Datos6.dProceso
End With
        'completar ctacod
        Dim numctafull As String
        Dim nomcli As String
        nomcli = Trim(Datos6.cNomCliente)
        numctafull = Datos6.cCodCta

        Dim rsX As ADODB.Recordset
        Dim rsy As New ADODB.Recordset
        Dim oGen As COMDCredito.DCOMCreditos

        Set rsX = New ADODB.Recordset
        Set rsy = New ADODB.Recordset
        Set oGen = New COMDCredito.DCOMCreditos
        
        Set rsX = oGen.DevolverDatos_Cliente_Corresponsalia(numctafull, nomcli) 'obtengo los datos del cliente

       If Not rsX.EOF And Not rsX.BOF Then
            rsDocs.Fields("cCodCtaF") = rsX("cCtaCod")
            psCtaCod = rsX("cCtaCod")
            rsDocs.Fields("cNomCliente") = Trim(Mid(rsX("cPersNombre"), 1, 59))
      
             Call oCredD.Insertar_Datos_Detalle_corresponsaliaBN(cod_det, rsDocs)  'inserto en detalle
            'para actualizar los Mov por cada operacion
            Set rsy = oGen.DevolverIDDatos_DetalleCorresponsalia(cod_det, CInt(Datos6.nNroCuota), psCtaCod, Datos6.dProceso)
            If Not rsy.EOF And Not rsy.BOF Then
                Dim cod_dety As Double
                cod_dety = CDbl(rsy("Id"))
            End If
            
            'se cambio por Importe Cobrado, debido a pagos parciales
            If Not CargaDatos(cod_dety, psCtaCod, CDbl((Datos6.nNumPag) + "." + (Datos6.nNumDec)), Datos6.dProceso, CInt(Datos6.nNroCuota), True, Corresponsalia) Then
                   Dim ultm As Integer
                   n_procesados = n_procesados + 1
                   Call oCredD.Actualizar_DetallexEst_BN(cod_det, psCtaCod, CInt(Datos6.nNroCuota), (Datos6.dProceso), 0, 0)
                   ultm = n_procesados + procesados
                      If n_procesados = p2 Then
                          MsgBox "El Proceso de Pago no se Realizó" & vbCrLf & " Comuníquese con el Area de TI"
                      ElseIf ultm = p2 Then
                          MsgBox "El Pago Finalizó correctamente pero hubieron registros NO procesados, Revise el Reporte!!"
                          Me.cmdLlenar.Enabled = True
                          Me.CmdCargaArch.Enabled = False
                      End If
             End If
       
       Else
             MsgBox "No se pudo leer los datos de : " & nomcli & " Este registro no será procesado automaticamente"
             n_procesados = n_procesados + 1
             ultm = n_procesados + procesados
             If n_procesados = p2 Then
                MsgBox "No se realizó el Pago de Ningún Registro" & vbCrLf & " Comuníquese con el Area de TI"
             ElseIf ultm = p2 Then
                MsgBox "El Pago Finalizó Correctamente, pero hubieron registros NO procesados"
                Me.cmdLlenar.Enabled = True
                Me.CmdCargaArch.Enabled = False
            End If
       End If
     Set oGen = Nothing
End Sub

Private Sub oImp_BarraClose()
    oBarra.CloseForm Me
End Sub

Private Sub oImp_BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
    oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub oImp_BarraShow(pnMax As Variant)
    oBarra.ShowForm Me
    oBarra.CaptionSyle = eCap_CaptionPercent
    oBarra.Max = pnMax
End Sub

Private Sub cmdImprimir_Click()
    Dim previo As previo.clsprevio
    Set previo = New previo.clsprevio
    previo.Show sCadImpre, "Pago Convenio BN", True
    Set previo = Nothing
End Sub

Private Sub cmdLlenar_Click()
If tipoArch <> 2 Then
    cargar_datos
    Cargar_grilla
        If n_procesados >= 1 Then
            cargar_datos1
            Cargar_grilla1
        End If
Else
    cargar_datosCorres
End If
End Sub

Sub cargar_datosCorres()
    Dim rs1d As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Set rs1d = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rs1d = oGen.DevolverDatos_Cabecera_reporte_CorresponsaliaPago(cod_det)
    Set oGen = Nothing

    If Not rs1d.EOF And Not rs1d.BOF Then
        est = 1
        activar_lbl (True)
        Me.lblTotal.Caption = rs1d("Num_Pro")
        Me.lblsoles.Caption = rs1d("nImpToTS")
        Me.lbldolar.Caption = rs1d("nImpToTD")
        lblfec.Caption = rs1d("fProceso")
        Me.lblPSoles.Caption = rs1d("nImpToTS1")
        Me.lblPdolar.Caption = rs1d("nImpToTD1")
        Me.lblreg.Caption = rs1d("Num_Pro")
        
        Cargar_grillaCorresponsalia (CDbl(rs1d("id")))
           If rs1d("Num_NoPro") > 0 Then
                cargar_datos1_Corresponsalia (CDbl(rs1d("id")))
                Cargar_grilla1_Corresponsalia (CDbl(rs1d("id")))
            End If
    Else
        est = 0
    End If
    rs1d.Close
    Set rs1d = Nothing
End Sub

Sub Cargar_grillaCorresponsalia(ByVal cod_det As Double)
    Dim rsx1 As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim i As Integer
    Set rsx1 = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsx1 = oGen.DevolverDatos_Detalle_108121_Corresponsalia(cod_det)
    Set oGen = Nothing
    i = 0

    ConfigurarMShComite
      Do Until rsx1.EOF
        With Me.Mshbco
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsx1!cCodCtafull
            .TextMatrix(.Rows - 2, 2) = rsx1!cNombreCliente
            .TextMatrix(.Rows - 2, 3) = rsx1!nNroCuota
            .TextMatrix(.Rows - 2, 4) = IIf(rsx1!nmoneda = 1, "Soles", "Dolares")
            .TextMatrix(.Rows - 2, 5) = rsx1!nImporteCobrado
        End With
        rsx1.MoveNext
    Loop
  Me.CmdImprimir.Enabled = True
  Me.cmdReporte.Enabled = True
  rsx1.Close
Set rsx1 = Nothing
End Sub

Sub cargar_datos1_Corresponsalia(ByVal cod_det As Double)
    Dim rs2c As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Set rs2c = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rs2c = oGen.DevolverDatos_Cabecera_108121_NoProcesados_Corresponsalia(cod_det)
    Set oGen = Nothing

     If Not rs2c.EOF And Not rs2c.BOF Then
        activar_lbl1 (True)
        Me.lblimpdx.Caption = rs2c("nImpToTD")
        Me.lblimpsx.Caption = rs2c("nImpToTS")
        Me.lblnumx.Caption = rs2c("nNumReg")
     End If

    rs2c.Close
    Set rs2c = Nothing
End Sub

Sub Cargar_grilla1_Corresponsalia(ByVal cod_det As Double)
    Dim rsy As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCreditos
    Dim i As Integer
    Set rsy = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCreditos

    Set rsy = oGen.DevolverDatos_Detalle_108121_NoPagados_Corresponsalia(cod_det)
    Set oGen = Nothing
    i = 0

    ConfigurarMShComite1
      Do Until rsy.EOF
        With Me.Mshbco1
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsy!cCodCtafull
            .TextMatrix(.Rows - 2, 2) = rsy!cNombreCliente
            .TextMatrix(.Rows - 2, 3) = rsy!nNroCuota
            .TextMatrix(.Rows - 2, 4) = IIf(rsy!nmoneda = 1, "Soles", "Dolares")
            .TextMatrix(.Rows - 2, 5) = rsy!nImporteCobrado
        End With
        rsy.MoveNext
    Loop
  rsy.Close
Set rsy = Nothing
End Sub

Private Sub cmdReporte_Click()
    Dim sCadImp As String
    Dim oPrev As previo.clsprevio
    Dim oGen As COMNCredito.NCOMCredito

    Set oPrev = New previo.clsprevio
    Set oGen = New COMNCredito.NCOMCredito

    If (Me.lblfec.Caption) <> "" Then
        If tipoArch = 2 Then
            sCadImp = oGen.ImprimeClientesCorresponsaliaBN(gsCodUser, gdFecSis, CDate(Me.lblfec.Caption), gsNomCmac)
        Else
            sCadImp = oGen.ImprimeClientesRecuperacionesBN(gsCodUser, gdFecSis, CDate(Me.lblfec.Caption), gsNomCmac)
        End If
    Else
        MsgBox "No se puede presentar el reporte de Pagos"
        Exit Sub
    End If
    
    Set oNPers = Nothing
    previo.Show sCadImp, "Registro de Archivo de Recuperaciones Cobradas", False
    Set oPrev = Nothing
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set Cn = New ADODB.Connection
    ConfigurarMShComite
    ConfigurarMShComite1
    sCadImpre = ""
    Me.SSTAB.TabVisible(1) = False
    procesados = 0
    n_procesados = 0
    'EJVG20131128 ***
    Dim obj As New COMDCredito.DCOMParametro
    fnCobrosPorcMN = (obj.RecuperaValorParametro(3082) / 100) + 1
    fnCobrosPorcME = (obj.RecuperaValorParametro(3083) / 100) + 1
    fnCobrosMontoMinMN = obj.RecuperaValorParametro(3084)
    fnCobrosMontoMinME = obj.RecuperaValorParametro(3086)
    fnCobrosMontoMaxMN = obj.RecuperaValorParametro(3085)
    fnCobrosMontoMaxME = obj.RecuperaValorParametro(3087)
    Set obj = Nothing
    'END EJVG *******
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    Set Cn = Nothing
End Sub

Sub Perdonar_Mora(cCtaCod As String, nNumPerdon As Integer, dFProceso As Date)
    Dim MatCalendMora() As String
    Dim ValorTotalMora As Currency
    Dim ValorTotalMoraDia As Currency
    Dim ValorTotalMoraDiaCalc As Currency
    Dim nDiasAtrasoSist As Integer
    Dim nNroCalen As Integer
    Dim nCuota As Integer
    Dim oCalend As COMDCredito.DCOMCalendario
    Dim oCale As COMDCredito.DCOMCredActBD
    Dim R1 As ADODB.Recordset
    Dim R2 As ADODB.Recordset
    'MAVM 20120418
    Dim R3 As ADODB.Recordset
    Dim R4 As ADODB.Recordset
    Dim nTasaInteresMora As Double
    Dim nCapital As Currency
    Dim nCapitalAnt As Currency
    '***
    ''MAVM 20130620 ***
    Dim R5 As ADODB.Recordset
    Dim R6 As ADODB.Recordset
    Dim R7 As ADODB.Recordset
    Dim R8 As ADODB.Recordset
    Dim sTpoProdCod As String
    Dim nTpoCamb As Double
    '***
    Dim nIntComp As Currency 'JUEZ 20131111
    Dim nIntGracia As Currency 'JUEZ 20131111
    
    If nNumPerdon > 0 Then
        Set oCalend = New COMDCredito.DCOMCalendario
        Set oCale = New COMDCredito.DCOMCredActBD
        ValorTotalMora = 0#
        ValorTotalMoraDia = 0#
        ValorTotalMoraDiaCalc = 0#
        nDiasAtrasoSist = 0
        nCuota = 0
        nNroCalen = 0
        sTpoProdCod = "" 'MAVM 20130620
        nTpoCamb = 0 'MAVM 20130620
        
        'MAVM 20120809 *** Verifica Pago Dia Anterior
        Set R4 = oCalend.RecuperaPagCredDiaAnt(psCtaCod)
        If R4.RecordCount > 0 Then
            nCapitalAnt = R4!nMonto
            R4.Close
        Else
            nCapitalAnt = 0
            Set R4 = Nothing
        End If
        '***
        
        Set R2 = oCalend.RecuperaCalendarioPagosPendiente(psCtaCod)
        Set oCalend = Nothing
        ReDim MatCalendMora(R2.RecordCount, 13)
        
        If R2.RecordCount > 0 Then
            MatCalendMora(R2.Bookmark - 1, 0) = Format(R2!dVenc, "dd/mm/yyyy")
            MatCalendMora(R2.Bookmark - 1, 1) = Trim(str(R2!nCuota))
            MatCalendMora(R2.Bookmark - 1, 6) = Format(IIf(IsNull(R2!nIntMor), 0, R2!nIntMor), "#0.00")
            'MAVM 20120418 ***
            MatCalendMora(R2.Bookmark - 1, 7) = Format(IIf(IsNull(R2!nCapital), 0, R2!nCapital), "#0.00")
            MatCalendMora(R2.Bookmark - 1, 8) = Format(IIf(IsNull(R2!nIntComp), 0, R2!nIntComp), "#0.00") 'JUEZ 20131111
            MatCalendMora(R2.Bookmark - 1, 10) = Format(IIf(IsNull(R2!nIntGracia), 0, R2!nIntGracia), "#0.00") 'JUEZ 20131111
            '***
            R2.Close
            Set R2 = Nothing
               
            ValorTotalMora = MatCalendMora(0, 6)
            nCuota = MatCalendMora(0, 1)
            'MAVM 20120418 ***
            nCapital = MatCalendMora(0, 7)
            '***
            nIntComp = MatCalendMora(0, 8) 'JUEZ 20131111
            nIntGracia = MatCalendMora(0, 10) 'JUEZ 20131111
        Else
            Set R2 = Nothing
            Exit Sub
        End If
    
        Dim oCreditoMora As COMDCredito.DCOMCredito
        Set oCreditoMora = New COMDCredito.DCOMCredito
        Set R1 = oCreditoMora.RecuperaColocacCred(psCtaCod)
        'MAVM 20120418 ***
        Set R3 = oCreditoMora.RecuperaProductoTasaInteres(psCtaCod, gColocLineaCredTasasIntMoratNormal)
        '***
        'MAVM 20130620 ***
        Set R5 = oCreditoMora.RecuperaTpoProd_AplicaPenalidad(431)
        If R5.RecordCount > 0 Then
            Set R6 = oCreditoMora.RecuperaTpoProd_XValor(psCtaCod, R5!nConsSisValor)
            If R6.RecordCount > 0 Then
                sTpoProdCod = R6!cTpoProdCod
            End If
        End If
        '***
        
        If R1.RecordCount > 0 Then
            nDiasAtrasoSist = R1!nDiasAtraso
            nNroCalen = R1!nNroCalen
            R1.Close
        Else
            Set R1 = Nothing
            Exit Sub
        End If
        
        'MAVM 20120418 ***
        'Set oCreditoMora = Nothing
        If R3.RecordCount > 0 Then
            nTasaInteresMora = R3!nTasaInteres
            R3.Close
        Else
            Set R3 = Nothing
            Exit Sub
        End If
        '***
        
        If nDiasAtrasoSist >= nNumPerdon Then
            If ValorTotalMora > 0 Then
                'If (nDiasAtrasoSist - nNumPerdon) <= 7 Then
                '    Call oCale.dUpdateColocalendDetMora(psCtaCod, nNroCalen, nCuota, 0, "1215")
                'End If
                    
                'MAVM 20120418 ***
                If oCreditoMora.bEsFeriadoDomingo(DateAdd("d", dFProceso, -1), "01") And (nDiasAtrasoSist - nNumPerdon) = 1 Then
                    Call oCale.dUpdateColocalendDetMora(psCtaCod, nNroCalen, nCuota, 0)
                Else
                    If sTpoProdCod = "" Then
                        If nCapitalAnt = 0 Then
                            'ValorTotalMoraDiaCalc = IIf(Round(Round((nTasaInteresMora / 100 * nCapital * 1), 4) * (nDiasAtrasoSist - nNumPerdon), 2) <= 0, 0, Round(Round((nTasaInteresMora / 100 * nCapital * 1), 4) * (nDiasAtrasoSist - nNumPerdon), 2))
                            ValorTotalMoraDiaCalc = oCreditoMora.CalculaMoraBN(nTasaInteresMora, nDiasAtrasoSist - nNumPerdon, nCapital, nIntComp, nIntGracia) 'JUEZ 20131111
                        Else
                            nCapital = nCapital + nCapitalAnt
                            'ValorTotalMoraDiaCalc = IIf(Round(Round((nTasaInteresMora / 100 * nCapital * 1), 4) * (nDiasAtrasoSist - nNumPerdon), 2) <= 0, 0, Round(Round((nTasaInteresMora / 100 * nCapital * 1), 4) * (nDiasAtrasoSist - nNumPerdon), 2))
                            ValorTotalMoraDiaCalc = oCreditoMora.CalculaMoraBN(nTasaInteresMora, nDiasAtrasoSist - nNumPerdon, nCapital, nIntComp, nIntGracia) 'JUEZ 20131111
                        End If
                    Else
                        If Mid(psCtaCod, 9, 1) <> "1" Then
                            Set R8 = oCreditoMora.DevolverTCMoneda(dFProceso - 1)
                            nTpoCamb = R8!nVenta
                        End If
                        Set R7 = oCreditoMora.CalculaMontoPenalidad(psCtaCod, nDiasAtrasoSist - nNumPerdon, nTpoCamb)
                        If R7.RecordCount > 0 Then
                            ValorTotalMoraDiaCalc = Format(R7!nImporte, "#0.00")
                        End If
                    End If
                End If
                '***
                If ValorTotalMoraDiaCalc >= 0 Then
                    Call oCale.dUpdateColocalendDetMora(psCtaCod, nNroCalen, nCuota, ValorTotalMoraDiaCalc)
                End If
            End If
        'Else
            'Call oCale.dUpdateColocalendDetMora(psCtaCod, nNroCalen, nCuota, 0)
        End If
    End If
End Sub
