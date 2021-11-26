VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2E37A9B8-9906-4590-9F3A-E67DB58122F5}#7.0#0"; "OcxLabelX.ocx"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmCredGeneraLecturaFilePagoVarios 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lectura de Pago de Servicios"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10800
   Icon            =   "frmCredGeneraLecturaFilePagoVarios.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   10800
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   765
      Left            =   120
      TabIndex        =   0
      Top             =   5880
      Width           =   10620
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7920
         TabIndex        =   5
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   9120
         TabIndex        =   4
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdLlenar 
         Caption         =   "&Mostrar Datos"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton CmdCargaArch 
         Caption         =   "&Subir Archivo"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6840
         TabIndex        =   1
         Top             =   240
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTAB 
      Height          =   5655
      Left            =   0
      TabIndex        =   6
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9975
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Datos Procesados"
      TabPicture(0)   =   "frmCredGeneraLecturaFilePagoVarios.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label7"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label5"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl2"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbl4"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl1"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblinsitucion"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label8"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label9"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblfec1"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblPdolar"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblPSoles"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbltotal"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Mshbco"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lbldolar"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lblsoles"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblfec"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lblreg"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Option1(0)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Option1(1)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "Datos No Procesados"
      TabPicture(1)   =   "frmCredGeneraLecturaFilePagoVarios.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblnumx"
      Tab(1).Control(1)=   "lblfecx"
      Tab(1).Control(2)=   "lblimpsx"
      Tab(1).Control(3)=   "lblimpdx"
      Tab(1).Control(4)=   "Mshbco1"
      Tab(1).Control(5)=   "Label1"
      Tab(1).Control(6)=   "Label2"
      Tab(1).Control(7)=   "Label3"
      Tab(1).Control(8)=   "Label4"
      Tab(1).ControlCount=   9
      Begin VB.OptionButton Option1 
         Caption         =   "Texto (*.caja)"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   35
         Top             =   480
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Excel"
         ForeColor       =   &H00000080&
         Height          =   195
         Index           =   0
         Left            =   1680
         TabIndex        =   34
         Top             =   480
         Visible         =   0   'False
         Width           =   1335
      End
      Begin OcxLabelX.LabelX lblreg 
         Height          =   450
         Left            =   5040
         TabIndex        =   7
         Top             =   1200
         Visible         =   0   'False
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   794
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblfec 
         Height          =   450
         Left            =   2760
         TabIndex        =   8
         Top             =   720
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   794
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblsoles 
         Height          =   450
         Left            =   7200
         TabIndex        =   9
         Top             =   720
         Visible         =   0   'False
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   794
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lbldolar 
         Height          =   450
         Left            =   9600
         TabIndex        =   10
         Top             =   720
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   794
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblnumx 
         Height          =   495
         Left            =   -73200
         TabIndex        =   11
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
         TabIndex        =   12
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
         TabIndex        =   13
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
         TabIndex        =   14
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
         Height          =   3375
         Left            =   120
         TabIndex        =   15
         Top             =   1680
         Width           =   10455
         _ExtentX        =   18441
         _ExtentY        =   5953
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Mshbco1 
         Height          =   3735
         Left            =   -74760
         TabIndex        =   16
         Top             =   600
         Width           =   9975
         _ExtentX        =   17595
         _ExtentY        =   6588
         _Version        =   393216
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin OcxLabelX.LabelX lbltotal 
         Height          =   450
         Left            =   1080
         TabIndex        =   17
         Top             =   720
         Visible         =   0   'False
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   794
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblPSoles 
         Height          =   450
         Left            =   7200
         TabIndex        =   18
         Top             =   1200
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   794
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblPdolar 
         Height          =   450
         Left            =   9600
         TabIndex        =   19
         Top             =   1200
         Visible         =   0   'False
         Width           =   1065
         _ExtentX        =   1879
         _ExtentY        =   794
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin OcxLabelX.LabelX lblfec1 
         Height          =   450
         Left            =   5040
         TabIndex        =   36
         Top             =   720
         Visible         =   0   'False
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   794
         FondoBlanco     =   0   'False
         Resalte         =   0
         Bold            =   0   'False
         Alignment       =   0
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "F. Caducidad:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   3960
         TabIndex        =   37
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Insitucion :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   240
         TabIndex        =   33
         Top             =   5280
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label lblinsitucion 
         AutoSize        =   -1  'True
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   240
         Left            =   1440
         TabIndex        =   32
         Top             =   5280
         Visible         =   0   'False
         Width           =   7635
      End
      Begin VB.Label lbl1 
         AutoSize        =   -1  'True
         Caption         =   "Procesados :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   3960
         TabIndex        =   30
         Top             =   1320
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.Label lbl4 
         AutoSize        =   -1  'True
         Caption         =   "Imp Dolares :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   8400
         TabIndex        =   29
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "Imp Soles :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   6240
         TabIndex        =   28
         Top             =   840
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label lbl2 
         AutoSize        =   -1  'True
         Caption         =   "F. Proceso:"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   1800
         TabIndex        =   27
         Top             =   840
         Visible         =   0   'False
         Width           =   915
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
         TabIndex        =   26
         Top             =   4680
         Visible         =   0   'False
         Width           =   1410
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
         TabIndex        =   25
         Top             =   4680
         Visible         =   0   'False
         Width           =   780
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
         TabIndex        =   24
         Top             =   4680
         Visible         =   0   'False
         Width           =   600
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
         TabIndex        =   23
         Top             =   4680
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Num. Reg :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   120
         TabIndex        =   22
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Imp Soles :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   6240
         TabIndex        =   21
         Top             =   1320
         Visible         =   0   'False
         Width           =   900
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Imp Dolares :"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   225
         Left            =   8400
         TabIndex        =   20
         Top             =   1320
         Visible         =   0   'False
         Width           =   1095
      End
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   30
      Left            =   0
      TabIndex        =   31
      Top             =   0
      Visible         =   0   'False
      Width           =   800
      _ExtentX        =   1402
      _ExtentY        =   53
      Filtro          =   "Archivos de Texto (*.pagos)|*.pagos|Archivos de Texto (*.cobros)|*.cobros"
      Altura          =   0
   End
End
Attribute VB_Name = "frmCredGeneraLecturaFilePagoVarios"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Excel
Dim modExcell() As String
Dim obj_Excel As Object, obj_Workbook As Object, obj_Worksheet As Object
Dim oBarra  As clsProgressBar
'-------
Dim fsNomFile As String
Dim fsPathFile As String
Dim fsruta As String
Dim Cn As ADODB.Connection
Dim rs As ADODB.Recordset
Dim rsDocs As ADODB.Recordset
Dim rsDocsE As ADODB.Recordset
'Variables para utilizar la estructura anterior
Dim Datos As CabServicios
Dim Datos4 As DetServicios
Dim DatosE As CabServiciosExcel
Dim Datos4E As DetServiciosExcel
'Variables para la posición del primer y último registro
Dim RegCan As Long
Dim procesados As Long
Dim n_procesados As Long
Dim cod_det As Long
Dim bterminoE As Boolean
Dim tFormaIngreso As Integer
'archivo cabecera
 Dim f As Integer
 Dim fd As String
 Dim fc As String
 Dim p1 As String
 Dim p3 As String
 Dim p2 As Long
 Dim lineas As Long
 Dim str_Linea As String
 Dim pCodInstTemp As String
 Dim ultm As Long
 Dim ultmE As Long
''''''Private oCredito As COMNCredito.NCOMCredito

Private Type CabServicios
    cCodIns As String * 13
    nNumReg As String * 10
    fProceso As String * 8
    fCaducidad As String * 8
    nImpToTS As String * 10
    nImpToTD As String * 10
End Type

'MADM 20110420 cCodCliE As String * 20
Private Type DetServicios
    cCodCli As String * 20
    cTpoDoc As String * 3
    cNumDoc As String * 12
    cApeCli As String * 30
    cNomCli As String * 30
    cMoneda As String * 3
    nImporte As String * 10
    cConcepto As String * 30
    cCodServicio As String * 25
    cPeriodo As String * 8
    cImprime As String * 25
End Type

Private Type CabServiciosExcel
    cCodInsE As String * 13
    nNumRegE As String * 10
    fProcesoE As String * 10
    fCaducidadE As String * 10
    nImpToTSE As String * 10
    nImpToTDE As String * 10
End Type

Private Type DetServiciosExcel
    cCodCliE As String * 20
    cTpoDocE As String * 3
    cNumDocE As String * 12
    cApeCliE As String * 30
    cNomCliE As String * 30
    cMonedaE As String * 3
    nImporteE As String * 10
    cConceptoE As String * 30
    cCodServicioE As String * 25
    cPeriodoE As String * 10
    cImprimeE As String * 25
End Type

'FORMA DE INGRESO tFormaIngreso = 3 :XLS tFormaIngreso = 2 : TXT tFormaIngreso = 1 MANUAL
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
    Me.cmdLlenar.Enabled = False
    Me.cmdReporte.Enabled = False
    'labeles procesados
    activar_lbl (False)
    Me.lblsoles.Caption = ""
    Me.lbldolar.Caption = ""
    lblfec.Caption = ""
    lblfec1.Caption = ""
    Me.lblreg.Caption = ""
    Me.lbltotal.Caption = ""
    Me.lblPSoles.Caption = ""
    Me.lblPdolar.Caption = ""
End Sub

Private Sub CmdCargaArch_Click()

If Option1(0).value = True Then
     MsgBox "Opción no Habilitada para iniciar proceso", vbInformation, "Aviso"
'    CdlgFile.nHwd = Me.hwnd
'    CdlgFile.Filtro = "Archivos Excel (*.xls)|*.xls"
'    Me.CdlgFile.Altura = 300
'    CdlgFile.Show
'
'    fsPathFile = CdlgFile.Ruta
'    fsruta = fsPathFile
'            If fsPathFile <> Empty Then
'                For i = Len(fsPathFile) - 1 To 1 Step -1
'                        If Mid(fsPathFile, i, 1) = "\" Then
'                            fsPathFile = Mid(CdlgFile.Ruta, 1, i)
'                            fsNomFile = Mid(CdlgFile.Ruta, i + 1, Len(CdlgFile.Ruta) - i)
'                            Exit For
'                        End If
'                 Next i
'              Screen.MousePointer = 11
'              tFormaIngreso = 3
'              Leer_LineasExcel (fsruta)
'            Else
'               MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
'               Exit Sub
'            End If
Else
    CdlgFile.nHwd = Me.hwnd
    CdlgFile.Filtro = "Archivos Servicios (*.caja)|*.caja"
    Me.CdlgFile.Altura = 300
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
              tFormaIngreso = 2
              Leer_Lineas (fsruta)
            Else
               MsgBox "No se ha Seleccionado ningún Archivo para Procesar", vbCritical, "Aviso"
               Exit Sub
            End If
End If
    Screen.MousePointer = 0
End Sub
Sub cargar_datos()
    Dim rsx As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCredito
    Set rsx = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCredito
'MADM 20110420
    Set rsx = oGen.DevolverDatosCabeceraPagoTotal(cod_det)
    Set oGen = Nothing

     If Not rsx.EOF And Not rsx.BOF Then
        activar_lbl (True)
        Me.lblsoles.Caption = rsx("nImpToTS")
        Me.lbldolar.Caption = rsx("nImpToTD")
        lblfec.Caption = rsx("fProceso")
        lblfec1.Caption = rsx("fCaducidad")
        Me.lblreg.Caption = procesados
        Me.lbltotal.Caption = rsx("nNumReg")
        Me.lblPSoles.Caption = rsx("nImpToTS1")
        Me.lblPdolar.Caption = rsx("nImpToTD1")
        lblinsitucion.Caption = rsx("cPersCodCli")
     End If
    rsx.Close
    Set rsx = Nothing
End Sub

 Sub Cargar_grilla()
    Dim rsx As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCredito
    Dim i As Integer
    Set rsx = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCredito

    Set rsx = oGen.DevolverDatosDetPagoTotal(cod_det)
    Set oGen = Nothing
    i = 0

    ConfigurarMShComite
      Do Until rsx.EOF
        With Me.Mshbco
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsx!cNombre
            .TextMatrix(.Rows - 2, 2) = rsx!cMoneda
            .TextMatrix(.Rows - 2, 3) = rsx!nImporteCuota
            .TextMatrix(.Rows - 2, 4) = rsx!cConcepto
            .TextMatrix(.Rows - 2, 5) = rsx!cCodServicio
            .TextMatrix(.Rows - 2, 6) = rsx!cPeriodo
        End With
        rsx.MoveNext
    Loop
  Me.cmdReporte.Enabled = True
  rsx.Close
  Set rsx = Nothing
End Sub

Sub ConfigurarMShComite()
 Mshbco.Clear
    Mshbco.Cols = 7
    Mshbco.Rows = 2
    With Mshbco
        .TextMatrix(0, 0) = "Id"
        .TextMatrix(0, 1) = "Nombre Cliente"
        .TextMatrix(0, 2) = "Moneda"
        .TextMatrix(0, 3) = "Imp.a Cobrar"
        .TextMatrix(0, 4) = "Concepto"
        .TextMatrix(0, 5) = "Cod Servicio"
        .TextMatrix(0, 6) = "Periodo"

        .ColWidth(0) = 400
        .ColWidth(1) = 3500
        .ColWidth(2) = 700
        .ColWidth(3) = 1500
        .ColWidth(4) = 3000
        .ColWidth(5) = 2000
        .ColWidth(6) = 1500
    End With
End Sub

Sub ConfigurarMShComite1()
 Mshbco1.Clear
    Mshbco1.Cols = 7
    Mshbco1.Rows = 2

    With Mshbco1
        .TextMatrix(0, 0) = "Id"
        .TextMatrix(0, 1) = "Nombre Cliente"
        .TextMatrix(0, 2) = "Moneda"
        .TextMatrix(0, 3) = "Imp.Cobrado"
        .TextMatrix(0, 4) = "Concepto"
        .TextMatrix(0, 5) = "Cod Servicio"
        .TextMatrix(0, 6) = "Periodo"

        .ColWidth(0) = 400
        .ColWidth(1) = 3500
        .ColWidth(2) = 700
        .ColWidth(3) = 1500
        .ColWidth(4) = 3000
        .ColWidth(5) = 2000
        .ColWidth(6) = 1500
    End With
End Sub

Sub activar_lbl(ByVal xb As Boolean)
Me.lbl1.Visible = xb
Me.lbl2.Visible = xb
Me.lbl3.Visible = xb
Me.Label5.Visible = xb
Me.Label6.Visible = xb
Me.Label7.Visible = xb
Me.Label8.Visible = xb
Me.Label9.Visible = xb
Me.lbl4.Visible = xb
Me.lbldolar.Visible = xb
Me.lblsoles.Visible = xb
Me.lblPSoles.Visible = xb
Me.lblPdolar.Visible = xb
Me.lbltotal.Visible = xb
Me.lblimpdx.Visible = xb
Me.lblimpsx.Visible = xb
Me.lblfec.Visible = xb
Me.lblfec1.Visible = xb
Me.lblreg.Visible = xb
lblinsitucion.Visible = xb
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

Public Function LlenaRecordSet_CabSer() As Boolean
    Dim fd1 As String
    Dim fc1 As String
    Dim per As Boolean
    Dim oCredD1 As COMDCredito.DCOMCredito
    Dim rsx As ADODB.Recordset
'    Dim nActualizaAnterior As Integer
    
    Set rsx = New ADODB.Recordset
    Set rsDocs = New ADODB.Recordset
    Set oCredD1 = New COMDCredito.DCOMCredito
    per = True
'    nActualizaAnterior = 0
    'MADM 20110601
        With rsDocs
            .Fields.Append "cCodIns", adVarChar, 13
            .Fields.Append "nNumReg", adVarChar, 10
            .Fields.Append "fProceso", adVarChar, 8
            .Fields.Append "fCaducidad", adVarChar, 8
            .Fields.Append "nImpToTS", adCurrency
            .Fields.Append "nImpToTD", adCurrency
            .Open
                .AddNew
                .Fields("cCodIns") = Datos.cCodIns
                .Fields("nNumReg") = Datos.nNumReg
                .Fields("fProceso") = Datos.fProceso
                .Fields("fCaducidad") = Datos.fCaducidad
                .Fields("nImpToTS") = Format(Mid(Datos.nImpToTS, 1, Len(Datos.nImpToTS) - 2), ".") & Right(Datos.nImpToTS, 2)
                .Fields("nImpToTD") = Format(Mid(Datos.nImpToTD, 1, Len(Datos.nImpToTD) - 2), ".") & Right(Datos.nImpToTD, 2)
        End With
           fd = rsDocs("fProceso")
           fc = rsDocs("fCaducidad")
           p1 = rsDocs("nImpToTS")
           p2 = rsDocs("nNumReg")
           p3 = rsDocs("cCodIns")
           'MADM 20110420 - tforma = 2 - texto
           Set rsx = oCredD1.CargaDatosCabeceraPagoServicio(p1, p3, p2, fd, fc, tFormaIngreso)
           'MADM 20110601 - FEC CADUCIDAD
           fd1 = Mid(fd, 7, 2) & "/" & Mid(fd, 5, 2) & "/" & Mid(fd, 1, 4)
           fc1 = Mid(fc, 7, 2) & "/" & Mid(fc, 5, 2) & "/" & Mid(fc, 1, 4)

           'MADM 20110602 - FECHA DE EMISION DISTINTA - NO PROCESA MAYOR
           If Not CDate(fd1) = gdFecSis Then
                If CDate(fd1) > gdFecSis Then
                    per = False
                    MsgBox "No puede procesar el Archivo debido a la Fecha Posterior", vbCritical, "Aviso"
                Else
                   If CDate(fc1) >= gdFecSis Then
                        If Not MsgBox("Seguro de Procesar un Archivo de Fecha Anterior", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                            per = False
                        End If
                   Else
                        per = False
                        MsgBox "No puede procesar el Archivo debido a la Fecha de Vencimiento", vbCritical, "Aviso"
                   End If
                End If
           End If
  
           If per Then
                If rsx.EOF And rsx.BOF Then
                   'MADM 20110712 - ActualizaDatosPagoVariosCabAnt
                    oCredD1.ActualizaDatosPagoVariosCabAnt CDate(fd1), p3
                   'MADM 20110420
                    Call oCredD1.InsertarDatosPagoVariosCab(rsDocs, tFormaIngreso)
                    LlenaRecordSet_CabSer = True
                    procesados = 0
                    n_procesados = 0
                Else
                  MsgBox "El archivo ya fue Procesado con Anterioridad, Verifique!!!", vbInformation, "Aviso"
                  LlenaRecordSet_CabSer = False
                End If
            Else
                   LlenaRecordSet_CabSer = False
                   Exit Function
           End If
        rsx.Close
        Set rsx = Nothing
        Set oCredD1 = Nothing
End Function

Public Function LlenaRecordSet_CabSerExcel() As Boolean
    Dim fd1 As String
    Dim per As Boolean
    Dim rsx As ADODB.Recordset
    Dim oCredD1 As COMDCredito.DCOMCredito

    Set rsx = New ADODB.Recordset
    Set rsDocsE = New ADODB.Recordset
    Set oCredD1 = New COMDCredito.DCOMCredito
    per = True
    'MADM 20110601
        With rsDocsE
            .Fields.Append "cCodIns", adVarChar, 13
            .Fields.Append "nNumReg", adVarChar, 10
            .Fields.Append "fProceso", adVarChar, 10
            .Fields.Append "fCaducidad", adVarChar, 10
            .Fields.Append "nImpToTS", adCurrency
            .Fields.Append "nImpToTD", adCurrency
            .Open
                .AddNew
                .Fields("cCodIns") = DatosE.cCodInsE
                .Fields("nNumReg") = DatosE.nNumRegE
                .Fields("fProceso") = DatosE.fProcesoE
                .Fields("fCaducidad") = DatosE.fCaducidadE
                .Fields("nImpToTS") = DatosE.nImpToTSE
                .Fields("nImpToTD") = DatosE.nImpToTDE
        End With

           fd = rsDocsE("fProceso")
           fc = rsDocsE("fCaducidad")
           p1 = rsDocsE("nImpToTS")
           p2 = rsDocsE("nNumReg")
           p3 = rsDocsE("cCodIns")
           'MADM 20110420
           Set rsx = oCredD1.CargaDatosCabeceraPagoServicio(p1, p3, p2, fd, fc, tFormaIngreso)

            If Not CDate(fd) = gdFecSis Then
                'NO PROCESA FECHA ANTERIOR
                If CDate(fd) > gdFecSis Then
                    per = Falses
                    MsgBox "No puede procesar el Archivo debido a la Fecha Posterior", vbCritical, "Aviso"
                Else
                   If CDate(fc) > gdFecSis Then
                        If Not MsgBox("Seguro de Procesar un Archivo de Fecha Anterior", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                            per = False
                        End If
                   Else
                        per = False
                        MsgBox "No puede procesar el Archivo debido a la Fecha de Vencimiento", vbCritical, "Aviso"
                   End If
                End If
           End If

           If per Then
                If rsx.EOF And rsx.BOF Then
                   Call oCredD1.InsertarDatosPagoVariosCab(rsDocsE, tFormaIngreso)
                   LlenaRecordSet_CabSerExcel = True
                   procesados = 0
                   n_procesados = 0
                Else
                  MsgBox "El archivo ya fue Procesado con Anterioridad, Verifique!!!", vbInformation, "Aviso"
                  LlenaRecordSet_CabSerExcel = False
                End If
            Else
                   LlenaRecordSet_CabSerExcel = False
                   Exit Function
           End If
        rsx.Close
        Set rsx = Nothing
        Set oCredD1 = Nothing
End Function
Sub LlenaRecordSet_DetSer()
Dim oCredD As COMDCredito.DCOMCredito
Set rsDocs = New ADODB.Recordset
Set oCredD = New COMDCredito.DCOMCredito
'MADM 20110602 - cImprime
With rsDocs
    .Fields.Append "cNumId", adDouble
    .Fields.Append "cCodCli", adVarChar, 20
    .Fields.Append "cTpoDoc", adVarChar, 3
    .Fields.Append "cNumDoc", adVarChar, 12
    .Fields.Append "cApeCli", adVarChar, 30
    .Fields.Append "cNomCli", adVarChar, 30
    .Fields.Append "cMoneda", adVarChar, 3
    .Fields.Append "nImporte", adCurrency
    .Fields.Append "cConcepto", adVarChar, 30
    .Fields.Append "cCodServicio", adVarChar, 25
    .Fields.Append "cPeriodo", adVarChar, 8
    .Fields.Append "cImprime", adVarChar, 25
    .Open

        .AddNew
        .Fields("cNumId") = cod_det
        .Fields("cCodCli") = Trim(Datos4.cCodCli)
        .Fields("cTpoDoc") = Datos4.cTpoDoc
        .Fields("cNumDoc") = IIf(Datos4.cTpoDoc = "DNI", Right(Datos4.cNumDoc, 8), Datos4.cNumDoc)
        .Fields("cApeCli") = Replace(Datos4.cApeCli, "'", " ")
        .Fields("cNomCli") = Replace(Datos4.cNomCli, "'", " ")
        .Fields("cMoneda") = Datos4.cMoneda
        .Fields("nImporte") = Format(Mid(Datos4.nImporte, 1, Len(Datos4.nImporte) - 2), ".") + Right(Datos4.nImporte, 2)
        .Fields("cConcepto") = Replace(Datos4.cConcepto, "'", " ")
        .Fields("cCodServicio") = Replace(Datos4.cCodServicio, "'", " ")
        .Fields("cPeriodo") = Datos4.cPeriodo
        .Fields("cImprime") = Datos4.cImprime
End With

        If oCredD.InsertarDatosPagoVariosDet(rsDocs) Then
           procesados = procesados + 1
           ultm = n_procesados + procesados
            If n_procesados = p2 Then
                MsgBox "No se pudieron registrar ningun registro " & vbCrLf & "Comuníquese con el Area de TI"
            ElseIf procesados = p2 Then
                MsgBox "El Proceso de Registro se Realizó Correctamente"
                Me.cmdLlenar.Enabled = True
                Me.CmdCargaArch.Enabled = False
            ElseIf ultm = p2 Then
                MsgBox "El Proceso de Registro Finalizó correctamente pero hubieron registros NO procesados, Revise el Reporte!!"
                Me.cmdLlenar.Enabled = True
                Me.CmdCargaArch.Enabled = False
            End If
        Else
            MsgBox "No se pudo leer los datos de : " & NomCli & " Este registro no será procesado automaticamente"
             n_procesados = n_procesados + 1
             ultm = n_procesados + procesados
             If n_procesados = p2 Then
                MsgBox "No se realizó el Proceso de Ningún Registro" & vbCrLf & "Comuniquese con el Area de TI"
             ElseIf ultm = p2 Then
                MsgBox "El Proceso de Registro terminó Correctamente, pero hubieron registros NO procesados"
                Me.cmdLlenar.Enabled = True
                Me.CmdCargaArch.Enabled = False
            End If

        End If
 End Sub
Sub LlenaRecordSet_DetSerExcel(pfin As Integer)
Dim oCredD As COMDCredito.DCOMCredito
Set rsDocsED = New ADODB.Recordset
Set oCredD = New COMDCredito.DCOMCredito

If pfin = 1 Then
    ultmE = n_procesados + procesados
        If n_procesados = p2 Then
            MsgBox "No se pudieron registrar ningun registro " & vbCrLf & "Comuníquese con el Area de TI"
            bterminoE = True
        ElseIf procesados = p2 Then
            MsgBox "El Proceso de Registro terminó Correctamente"
            Me.cmdLlenar.Enabled = True
            Me.CmdCargaArch.Enabled = False
            bterminoE = True
        ElseIf ultmE = p2 Then
            MsgBox "El Proceso de Registro Finalizó correctamente pero hubieron registros NO procesados, Revise el Reporte!!"
            Me.cmdLlenar.Enabled = True
            Me.CmdCargaArch.Enabled = False
            bterminoE = True
        End If
Else
    With rsDocsED
        .Fields.Append "cNumId", adDouble
        .Fields.Append "cCodCli", adVarChar, 20
        .Fields.Append "cTpoDoc", adVarChar, 3
        .Fields.Append "cNumDoc", adVarChar, 12
        .Fields.Append "cApeCli", adVarChar, 30
        .Fields.Append "cNomCli", adVarChar, 30
        .Fields.Append "cMoneda", adVarChar, 3
        .Fields.Append "nImporte", adCurrency
        .Fields.Append "cConcepto", adVarChar, 30
        .Fields.Append "cCodServicio", adVarChar, 25
        .Fields.Append "cPeriodo", adVarChar, 10
        .Fields.Append "cImprime", adVarChar, 25
        .Open

            .AddNew
            .Fields("cNumId") = cod_det
            .Fields("cCodCli") = Datos4E.cCodCliE
            .Fields("cTpoDoc") = Datos4E.cTpoDocE
            .Fields("cNumDoc") = Datos4E.cNumDocE
            .Fields("cApeCli") = Datos4E.cApeCliE
            .Fields("cNomCli") = Datos4E.cNomCliE
            .Fields("cMoneda") = Datos4E.cMonedaE
            .Fields("nImporte") = Datos4E.nImporteE
            .Fields("cConcepto") = Datos4E.cConceptoE
            .Fields("cCodServicio") = Datos4E.cCodServicioE
            .Fields("cPeriodo") = Datos4E.cPeriodoE
            .Fields("cImprime") = Datos4E.cImprimeE
    End With
            'MADM 20110420
            If oCredD.InsertarDatosPagoVariosDet(rsDocsED, tFormaIngreso) Then
               procesados = procesados + 1
               ultmE = n_procesados + procesados
                If n_procesados = p2 Then
                    MsgBox "No se pudieron registrar ningun registro " & vbCrLf & "Comuníquese con el Area de TI"
                    bterminoE = True
                ElseIf procesados = p2 Then
                    MsgBox "El Proceso de Registro se Realizó Correctamente", vbInformation, "Aviso"
                    Me.cmdLlenar.Enabled = True
                    Me.CmdCargaArch.Enabled = False
                    bterminoE = True
                ElseIf ultmE = p2 Then
                    MsgBox "El Proceso de Registro Finalizó correctamente pero hubieron registros NO procesados, Revise el Reporte!!"
                    Me.cmdLlenar.Enabled = True
                    Me.CmdCargaArch.Enabled = False
                    bterminoE = True
                End If
            Else
                MsgBox "No se pudo leer los datos de : " & NomCli & " Este registro no será procesado automaticamente"
                 n_procesados = n_procesados + 1
                 ultmE = n_procesados + procesados
                 If n_procesados = p2 Then
                    MsgBox "No se realizó el Proceso de Ningún Registro" & vbCrLf & "Comuniquese con el Area de TI"
                    bterminoE = True
                 ElseIf ultmE = p2 Then
                    MsgBox "El Proceso de Registro terminó Correctamente, pero hubieron registros NO procesados"
                    Me.cmdLlenar.Enabled = True
                    Me.CmdCargaArch.Enabled = False
                    bterminoE = True
                End If
            End If
End If
 End Sub
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
    Dim rsx As ADODB.Recordset
    Dim oGen1 As COMDCredito.DCOMCredito
    Dim nCerrar As Integer
    Set oBarra = New clsProgressBar
    Set rsx = New ADODB.Recordset
    Set oGen1 = New COMDCredito.DCOMCredito

    lineas = 0
    f = FreeFile
    nCerrar = 0
    Open strTextFile For Input As #f
    Do 'grabar por cada item el cuerpo
        Line Input #f, str_Linea
        lineas = lineas + 1
        'MADM 20110601 - 59
        If lineas = 1 And Len(str_Linea) <> 59 Then
                MsgBox "El Archivo adjunto NO tiene la Cabecera correcta", vbInformation, "Aviso"
                Close #f
                nCerrar = 0
                Exit Sub
        ElseIf lineas = 1 Then
        'MADM 20110420
           Datos.cCodIns = Mid(str_Linea, 1, 13)
           pCodInstTemp = Trim(Datos.cCodIns)
           'MADM 20110601
           Datos.nNumReg = Mid(str_Linea, 14, 10)
           Datos.fProceso = Mid(str_Linea, 24, 8)
           Datos.fCaducidad = Mid(str_Linea, 32, 8)
           Datos.nImpToTS = Trim(Mid(str_Linea, 40, 10))
           Datos.nImpToTD = Trim(Mid(str_Linea, 50, 10))

           nCerrar = 0
             If Not Contar_Lineas(strTextFile) = CDbl(Datos.nNumReg) + 1 Then
                 MsgBox "El archivo a procesar NO tiene Completos los registros, Solicite un nuevo archivo al Area de TI", vbInformation, "Aviso"
                  Set rsx = Nothing
                    nCerrar = 0
                    Close #f
                    Exit Do
                    Exit Sub
             End If

             If LlenaRecordSet_CabSer = False Then
                    Set rsx = Nothing
                    nCerrar = 0
                    Close #f
                    Exit Do
                    Exit Sub
             End If
                    fd = rsDocs("fProceso")
                    fc = rsDocs("fCaducidad")
                    p1 = rsDocs("nImpToTS")
                    p2 = rsDocs("nNumReg")
                    p3 = rsDocs("cCodIns")
                   RegCan = p2
                   nCerrar = 1

                   'MADM 20110602 - 20110420 - Obtener el numero de deuda
                   Set rsx = oGen1.CargaDatosCabeceraPagoServicio(p1, p3, p2, fd, fc, 2)
                   Set oGen = Nothing

                   If Not rsx.EOF And Not rsx.BOF Then
                       'obtiene numero de deuda
                       cod_det = rsx("nDeuNro")
                       MsgBox "Se van a procesar " & RegCan & " registros, Esta operación puede demorar unos minutos ...", vbInformation, "Aviso"
                   Else
                        MsgBox "El archivo no tiene la estructura correcta", vbInformation, "Aviso"
                        Close #f
                        Exit Sub
                   End If
        Else
            'MADM 20110330 - DEFINIR LONGITUDES CABECERA Y DETALLE
            'MADM 20110602 191 - 20110420 - 149
            If Len(str_Linea) = 196 Then
                If lineas = 2 Then
                    oBarra.ShowForm Me
                    oBarra.CaptionSyle = eCap_CaptionPercent
                    oBarra.Max = RegCan
                    oBarra.Progress 0, "Proceso de Servicios : ", "Preparando Pago...", "Preparando", vbBlue
                Else
                   oBarra.Progress (lineas - 1), "Proceso de Servicios : ", "Efectuando Pago...", "Efectuando", vbBlue
                End If

                With Datos4
                        .cCodCli = Mid(str_Linea, 1, 20)
                        .cTpoDoc = Mid(str_Linea, 21, 3)
                        .cNumDoc = Mid(str_Linea, 24, 12)
                        .cApeCli = Mid(str_Linea, 36, 30)
                        .cNomCli = Mid(str_Linea, 66, 30)
                        .cMoneda = Mid(str_Linea, 96, 3)
                        .nImporte = Mid(str_Linea, 99, 10)
                        .cConcepto = Mid(str_Linea, 109, 30)
                        .cCodServicio = Mid(str_Linea, 139, 25)
                        .cPeriodo = Mid(str_Linea, 164, 8)
                        .cImprime = Mid(str_Linea, 172, 25)
                        Call LlenaRecordSet_DetSer 'cuerpo
                End With
            Else
                n_procesados = n_procesados + 1
                ultm = n_procesados + procesados
            If n_procesados = p2 Then
                MsgBox "El Registro no se pudo Realizar" & vbCrLf & "Comuníquese con el Area de TI"
            ElseIf ultm = p2 Then
                MsgBox "El Proceso de Registro Finalizó correctamente pero hubieron registros NO procesados, Revise el Reporte!!"
                Me.cmdLlenar.Enabled = True
                Me.CmdCargaArch.Enabled = False
            End If
                nCerrar = 0
            End If
        End If
    Loop While Not EOF(f)
    Close #f
    If nCerrar = 1 Then
        rsx.Close
    End If
    oBarra.CloseForm Me
    Set oBarra = Nothing
End Sub

Public Sub Leer_LineasExcel(strTextFile As String)
    Dim rsx As ADODB.Recordset
    Dim oGen1 As COMDCredito.DCOMCredito
    Dim nCerrar As Integer
    Dim cCodigoInsE As String
    Dim i As Long
    Dim n As Long
    Dim NumColumnasExcel As Integer
    Set oBarra = New clsProgressBar
    Set rsx = New ADODB.Recordset
    Set oGen1 = New COMDCredito.DCOMCredito

    lineas = 0
    nCerrar = 0
    NumColumnasExcel = 10
    Set obj_Excel = CreateObject("Excel.Application")
    Set obj_Workbook = obj_Excel.Workbooks.Open(strTextFile)
    Set obj_Worksheet = obj_Workbook.ActiveSheet

           Do
               If Not bterminoE Then
                   lineas = lineas + 1
                        If lineas = 1 Then
                            'MADM 20110420 - EXTRAE CODIGO INSTITUCION
                            cCodigoInsE = IIf(Len(Trim(obj_Worksheet.Cells(i + 2, n + 2).value)) = 13, Trim(obj_Worksheet.Cells(i + 2, n + 2).value), "")
                            If cCodigoInsE = "" Then
                                MsgBox "El código de la Institución no tiene el formato Correcto", vbCritical, "Aviso"
                                nCerrar = 0
                                Exit Sub
                            Else
    '                            If Not oGen1.GetValConvenioInstitucionxAgenciaTipo(cCodigoInsE) Then
    '                                MsgBox "La Institución no está Registrada para Convenio de Pago de Servicios Total", vbCritical, "Aviso"
    '                                nCerrar = 0
    '                                Exit Sub
    '                            ElseIf Not oGen1.GetValConvenioInstitucionxFechaVenc(cCodigoInsE) Then
    '                                MsgBox "El Convenio con la Institución ha Finalizado, Verifique Fecha vencimiento !! ", vbCritical, "Aviso"
    '                                nCerrar = 0
    '                                Exit Sub
    '                            Else
    '                                DatosE.cCodInsE = cCodigoInsE
    '                                pCodInstTemp = DatosE.cCodInsE
    '                            End If

                                If Not oGen1.GetValConvenioInstitucionxFechaVenc(cCodigoInsE) Then
                                    MsgBox "El Convenio con la Institución ha Finalizado, Verifique Fecha vencimiento !! ", vbCritical, "Aviso"
                                    nCerrar = 0
                                    Exit Sub
                                Else
                                    DatosE.cCodInsE = cCodigoInsE
                                    pCodInstTemp = DatosE.cCodInsE
                                End If

                            End If
                       ElseIf lineas = 2 Then
                            DatosE.nNumRegE = obj_Worksheet.Cells(lineas + 1, 2).value
                       ElseIf lineas = 3 Then
                            DatosE.fProcesoE = obj_Worksheet.Cells(lineas + 1, 2).value
                       ElseIf lineas = 4 Then
                            DatosE.fCaducidadE = obj_Worksheet.Cells(lineas + 1, 2).value
                       ElseIf lineas = 5 Then
                            DatosE.nImpToTSE = obj_Worksheet.Cells(lineas + 1, 2).value
                       ElseIf lineas = 6 Then
                            DatosE.nImpToTDE = obj_Worksheet.Cells(lineas + 1, 2).value
                       ElseIf lineas = 7 Then
                            If LlenaRecordSet_CabSerExcel = False Then
                                nCerrar = 0
                                Exit Do
                                Exit Sub
                            End If

                            fd = rsDocsE("fProceso")
                            fc = rsDocsE("fCaducidad")
                            p1 = rsDocsE("nImpToTS")
                            p2 = rsDocsE("nNumReg")
                            p3 = rsDocsE("cCodIns")

                           RegCan = p2
                           nCerrar = 1

                           'MADM 20110420
                           Set rsx = oGen1.CargaDatosCabeceraPagoServicio(p1, p3, p2, fd, fc, 3)
                           Set oGen = Nothing

                           If Not rsx.EOF And Not rsx.BOF Then
                               cod_det = rsx("nDeuNro")
                               MsgBox "Se van a procesar registros, Esta operación puede demorar unos minutos ...", vbInformation, "Aviso"
                           Else
                                MsgBox "El archivo no tiene la estructura correcta", vbInformation, "Aviso"
                                Set obj_Excel = Nothing
                                Set obj_Workbook = Nothing
                                Set obj_Worksheet = Nothing
                                bterminoE = False
                                oBarra.CloseForm Me
                                Set oBarra = Nothing
                                If nCerrar = 1 Then
                                    rsx.Close
                                End If
                                Exit Sub
                           End If

                            oBarra.ShowForm Me
                            oBarra.CaptionSyle = eCap_CaptionPercent
                            oBarra.Max = RegCan + 8
                            oBarra.Progress 0, "Proceso de Servicios : ", "Preparando Registros de Pago de Servicios...", "Preparando Registro", vbBlue

                       ElseIf lineas > 7 Then
                            oBarra.Progress (lineas - 1), "Proceso de Servicios : ", "Efectuando Registro ...", "Efectuando", vbBlue

                            If lineas <> 8 And (obj_Worksheet.Cells(lineas + 2, 2).value = "" Or obj_Worksheet.Cells(lineas + 2, 3).value = "" Or obj_Worksheet.Cells(lineas + 2, 4).value = "" Or obj_Worksheet.Cells(lineas + 2, 5).value = "") Then
                               n_procesados = n_procesados + 1
                               LlenaRecordSet_DetSerExcel 1
                            Else
                                With Datos4E
                                'MADM 20110420
                                        .cCodCliE = obj_Worksheet.Cells(lineas + 2, 2).value
                                        .cTpoDocE = obj_Worksheet.Cells(lineas + 2, 2).value
                                        .cNumDocE = Trim(obj_Worksheet.Cells(lineas + 2, 3).value)
                                        .cApeCliE = Trim(obj_Worksheet.Cells(lineas + 2, 4).value)
                                        .cNomCliE = Trim(obj_Worksheet.Cells(lineas + 2, 5).value)
                                        .cMonedaE = Trim(obj_Worksheet.Cells(lineas + 2, 6).value)
                                        .nImporteE = Trim(obj_Worksheet.Cells(lineas + 2, 7).value)
                                        .cConceptoE = Trim(obj_Worksheet.Cells(lineas + 2, 8).value)
                                        .cCodServicioE = Trim(obj_Worksheet.Cells(lineas + 2, 9).value)
                                        .cPeriodoE = Trim(obj_Worksheet.Cells(lineas + 2, 10).value)
                                        .cImprimeE = Trim(obj_Worksheet.Cells(lineas + 2, 11).value)
                                        LlenaRecordSet_DetSerExcel 0
                                End With
                            End If
                        End If
                    Else
                        Set obj_Excel = Nothing
                        Set obj_Workbook = Nothing
                        Set obj_Worksheet = Nothing
                        bterminoE = False
                        oBarra.CloseForm Me
                        Set oBarra = Nothing
                        If nCerrar = 1 Then
                            rsx.Close
                        End If
                        Exit Sub
                   End If
               Loop


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

Private Sub cmdLlenar_Click()
' Dim ClsPersona As COMDPersona.DCOMPersonas
' Dim RIns As ADODB.Recordset
'
' Set ClsPersona = New COMDPersona.DCOMPersonas
' Set RIns = ClsPersona.BuscaClienteServicios(txtDocPer.Text, bInstit, BusquedaDocumento)
' Set ClsPersona = Nothing
' Trim (R!cNombre)
 cargar_datos
 Cargar_grilla
End Sub

Private Sub cmdReporte_Click()
    Dim sCadImp As String
    Dim oPrev As previo.clsprevio
    Dim oGen As COMNCredito.NCOMCredito

    Set oPrev = New previo.clsprevio
    Set oGen = New COMNCredito.NCOMCredito

    If (Me.lblfec.Caption) <> "" Then
        sCadImp = oGen.ImprimeReporteClientesPagoVarios(gsCodUser, gdFecSis, CDate(Me.lblfec.Caption), tFormaIngreso, gsNomCmac, lblinsitucion, cod_det)
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
CentraForm Me
    Me.Top = 0
    Me.Left = (Screen.Width - Me.Width) / 2
    Me.Icon = LoadPicture(App.path & gsRutaIcono)

ConfigurarMShComite
ConfigurarMShComite1
sCadImpre = ""
Me.SSTAB.TabVisible(1) = False
procesados = 0
n_procesados = 0
bterminoE = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set rs = Nothing
    Set Cn = Nothing
End Sub

'COMENTADO MADm 20110601
'Function validar_cola(numreg As Integer, filex As String) As Boolean
'    Dim f As Integer
'    Dim lineas As Long
'    validar_cola = True
'
'    f = FreeFile
'    Open filex For Input As #f
'    Do
'        Line Input #f, str_Linea
'
'        If lineas = (numreg - 1) Then
'            Datos1.nNumReg = Mid(str_Linea, 13, 4)
'            Datos1.fProceso = Mid(str_Linea, 29, 8)
'            Datos1.nNumPag = Mid(str_Linea, 59, 13)
'            Datos1.nNumDec = Mid(str_Linea, 72, 2)
'
'            If Not LlenaRecordSet_Datos60 Then
'                validar_cola = False
'            End If
'        End If
'        lineas = lineas + 1
'    Loop While Not EOF(f)
'    Close #f
'End Function




