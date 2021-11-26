VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{2E37A9B8-9906-4590-9F3A-E67DB58122F5}#7.0#0"; "OcxLabelX.ocx"
Object = "{160AE063-3670-11D5-8214-000103686C75}#6.0#0"; "PryOcxExplorer.ocx"
Begin VB.Form frmCredPagoConvenioBcoCred 
   Caption         =   "Proceso Archivos de Resultados - BCP"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10395
   Icon            =   "frmCredPagoConvenioBcoCred.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   6090
   ScaleWidth      =   10395
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   765
      Left            =   0
      TabIndex        =   19
      Top             =   5280
      Width           =   10380
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir Pago"
         Enabled         =   0   'False
         Height          =   375
         Left            =   6600
         TabIndex        =   25
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   7920
         TabIndex        =   24
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   9120
         TabIndex        =   23
         Top             =   240
         Width           =   1095
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
      Begin VB.CommandButton CmdCargaArch 
         Caption         =   "&Subir Archivo"
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   1425
      End
      Begin VB.CommandButton cmdReporte 
         Caption         =   "&Reporte"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5520
         TabIndex        =   20
         Top             =   240
         Width           =   975
      End
   End
   Begin TabDlg.SSTab SSTAB 
      Height          =   5175
      Left            =   0
      TabIndex        =   0
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
      TabPicture(0)   =   "frmCredPagoConvenioBcoCred.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbl2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl1"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label8"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lblMone"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lblPSoles"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbltotal"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Mshbco"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lblsoles"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblfec"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lblreg"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      TabCaption(1)   =   "Datos No Pagados"
      TabPicture(1)   =   "frmCredPagoConvenioBcoCred.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label4"
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(2)=   "Label1"
      Tab(1).Control(3)=   "Mshbco1"
      Tab(1).Control(4)=   "lblimpsx"
      Tab(1).Control(5)=   "lblfecx"
      Tab(1).Control(6)=   "lblnumx"
      Tab(1).ControlCount=   7
      Begin OcxLabelX.LabelX lblreg 
         Height          =   495
         Left            =   6720
         TabIndex        =   1
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
         TabIndex        =   2
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
         Left            =   8640
         TabIndex        =   3
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
      Begin OcxLabelX.LabelX lblnumx 
         Height          =   495
         Left            =   -70560
         TabIndex        =   4
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
         Left            =   -68400
         TabIndex        =   5
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
         Left            =   -65880
         TabIndex        =   6
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
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Mshbco 
         Height          =   3495
         Left            =   120
         TabIndex        =   7
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
         TabIndex        =   8
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
         TabIndex        =   9
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
         Left            =   9000
         TabIndex        =   10
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
      Begin OcxLabelX.LabelX lblMone 
         Height          =   495
         Left            =   6120
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
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5040
         TabIndex        =   28
         Top             =   480
         Visible         =   0   'False
         Width           =   750
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
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   5400
         TabIndex        =   18
         Top             =   4680
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.Label lbl3 
         AutoSize        =   -1  'True
         Caption         =   "Imp :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   7440
         TabIndex        =   17
         Top             =   480
         Visible         =   0   'False
         Width           =   435
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
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   2280
         TabIndex        =   16
         Top             =   480
         Visible         =   0   'False
         Width           =   990
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -72120
         TabIndex        =   15
         Top             =   4680
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.Label Label3 
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
         TabIndex        =   14
         Top             =   4680
         Visible         =   0   'False
         Width           =   765
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
         ForeColor       =   &H000040C0&
         Height          =   195
         Left            =   -69480
         TabIndex        =   13
         Top             =   4680
         Visible         =   0   'False
         Width           =   990
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
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Imp :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   7920
         TabIndex        =   11
         Top             =   4680
         Visible         =   0   'False
         Width           =   435
      End
   End
   Begin PryOcxExplorer.OcxCdlgExplorer CdlgFile 
      Height          =   30
      Left            =   0
      TabIndex        =   27
      Top             =   0
      Visible         =   0   'False
      Width           =   700
      _ExtentX        =   1244
      _ExtentY        =   53
      Filtro          =   "Archivos de Texto (*.pagos)|*.pagos|Archivos de Texto (*.cobros)|*.cobros"
      Altura          =   0
   End
End
Attribute VB_Name = "frmCredPagoConvenioBcoCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim DatosC As Cabecera
Dim DatosD As Detalle

Dim rsBCPC As New ADODB.Recordset
Dim lineas As Long
Dim str_Linea As String
Dim f As Integer

Dim RegActual As Long
Dim RegUltimo As Long
Dim RegCan As Long
Dim procesados As Integer
Dim n_procesados As Integer
Dim RegActual2 As Long
Dim nDiasAtras As Integer
Dim nDiasDiferente As Integer
Dim oBarra As clsProgressBar

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
Dim cod_det As Double

Private Type Cabecera
    c1 As String * 2
    c2 As String * 3
    c3 As String * 1
    c4 As String * 7
    c5 As String * 1
    c6 As String * 8
    c7 As String * 9
    c8 As String * 15
    c9 As String * 4
    c10 As String * 6
End Type

Private Type Detalle
    d1 As String * 2
    d2 As String * 3
    d3 As String * 1
    d4 As String * 7
    d5 As String * 14
    d6 As String * 30
    d7 As String * 8
    d8 As String * 8
    d9 As String * 15
    d10 As String * 15
    d11 As String * 15
    d12 As String * 6
    d13 As String * 6
    d14 As String * 22
    d15 As String * 4
    d16 As String * 12
    d17 As String * 6
    d18 As String * 10
    d19 As String * 2
    d20 As String * 10
    d21 As String * 1
End Type
Private oCredito As COMNCredito.NCOMCredito

Private Sub cmdCancelar_Click()
ConfigurarMShComite
sCadImpre = ""
If Me.SSTAB.TabVisible(1) Then
    activar_lbl1 (False)
    Me.lblimpsx.Caption = ""
    Me.lblfecx.Caption = ""
    Me.lblnumx.Caption = ""
    Me.SSTAB.TabVisible(1) = False
End If
procesados = 0
n_procesados = 0
Me.CmdCargaArch.Enabled = True
Me.cmdImprimir.Enabled = False
Me.cmdLlenar.Enabled = False
Me.cmdReporte.Enabled = False
activar_lbl (False)
Me.lblsoles.Caption = ""
lblfec.Caption = ""
Me.lblreg.Caption = ""
Me.lbltotal.Caption = ""
Me.lblPSoles.Caption = ""
End Sub

Private Sub CmdCargaArch_Click()
    CdlgFile.nHwd = Me.hwnd
    CdlgFile.Filtro = "Archivos BCP (*.TXT)|*.TXT"
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
        Leer_Lineas (fsruta)
    Else
       MsgBox "No se selecciono Archivo", vbInformation, "Aviso"
       Exit Sub
    End If
    
    Screen.MousePointer = 0
End Sub

'Public Sub Leer_Lineas(strTextFile As String)
'    Dim rsx As ADODB.Recordset
'    Dim oGen As COMDCredito.DCOMCredito
'    Dim nCerrar As Integer
'    Set oBarra = New clsProgressBar
'    Set rsx = New ADODB.Recordset
'    Set oGen = New COMDCredito.DCOMCredito
'
'    lineas = 0
'    f = FreeFile
'    nCerrar = 0
'    Open strTextFile For Input As #f
'    Do
'        Line Input #f, str_Linea
'        lineas = lineas + 1
'
'       If lineas = 1 Then
'           DatosC.c1 = Mid(str_Linea, 1, 2)
'           DatosC.c2 = Mid(str_Linea, 3, 3)
'           DatosC.c3 = Mid(str_Linea, 6, 1)
'           DatosC.c4 = Mid(str_Linea, 7, 7)
'           DatosC.c5 = Mid(str_Linea, 14, 1)
'           DatosC.c6 = Mid(str_Linea, 15, 8)
'           DatosC.c7 = Mid(str_Linea, 23, 9)
'           DatosC.c8 = Mid(str_Linea, 32, 15)
'           DatosC.c9 = Mid(str_Linea, 47, 4)
'           DatosC.c10 = Mid(str_Linea, 51, 6)
'           nCerrar = 0
'
'           If LeerCabecera = False Then
'                Set rsx = Nothing
'                nCerrar = 0
'                Close #f
'                Exit Do
'                Exit Sub
'           End If
'           nCerrar = 1
'
'            Set rsx = oGen.getDatosCabeceraBCP(rsBCPC("c8"), CInt(rsBCPC("c7")), rsBCPC("c6"))
'            Set oGen = Nothing
'
'            If Not rsx.EOF And Not rsx.BOF Then
'                cod_det = rsx("id")
'                MsgBox "Se van a procesar " & RegActual2 & " registros, Esta operación puede demorar unos minutos ...", vbInformation, "Aviso"
'            Else
'                MsgBox "El archivo no tiene la estructura correcta", vbInformation, "Aviso"
'                Close #f
'                Exit Sub
'            End If
'        Else
'            If lineas = 2 Then
'                oBarra.ShowForm Me
'                oBarra.CaptionSyle = eCap_CaptionPercent
'                oBarra.Max = RegActual2
'                oBarra.Progress 0, "Proceso de Pago Convenio BCP : ", "Preparando Pago...", "Preparando", vbBlue
'            Else
'               oBarra.Progress (lineas - 1), "Proceso de Pago Convenio BCP : ", "Efectuando Pago...", "Efectuando", vbBlue
'            End If
'
'                With DatosD
'                            .d1 = Mid(str_Linea, 1, 2)
'                            .d2 = Mid(str_Linea, 3, 3)
'                            .d3 = Mid(str_Linea, 6, 1)
'                            .d4 = Mid(str_Linea, 7, 7)
'                            .d5 = Mid(str_Linea, 14, 14)
'                            .d6 = Mid(str_Linea, 28, 30)
'                            .d7 = Mid(str_Linea, 58, 8)
'                            .d8 = Mid(str_Linea, 66, 8)
'                            .d9 = Mid(str_Linea, 74, 15)
'                            .d10 = Mid(str_Linea, 89, 15)
'                            .d11 = Mid(str_Linea, 104, 15)
'                            .d12 = Mid(str_Linea, 119, 6)
'                            .d13 = Mid(str_Linea, 125, 6)
'                            .d14 = Mid(str_Linea, 131, 22)
'                            .d15 = Mid(str_Linea, 153, 4)
'                            .d16 = Mid(str_Linea, 157, 12)
'                            .d17 = Mid(str_Linea, 169, 6)
'                            .d18 = Mid(str_Linea, 175, 10)
'                            .d19 = Mid(str_Linea, 185, 2)
'                            .d20 = Mid(str_Linea, 187, 10)
'                            .d21 = Mid(str_Linea, 197, 1)
'                            Call LeerDetalle
'                End With
'            End If
'    Loop While Not EOF(f)
'    Close #f
'    If nCerrar = 1 Then
'        rsx.Close
'    End If
'    oBarra.CloseForm Me
'    Set oBarra = Nothing
'End Sub
'EJVG20130823 *** Se adecuó el proceso para no tomar los pagos extornados en el BCP
Public Sub Leer_Lineas(strTextFile As String)
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCredito
    Dim nCerrar As Integer
    Set oBarra = New clsProgressBar
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCredito
    Dim MatCab() As String
    Dim MatDet() As String
    Dim MatDetExt() As String
    Dim iMat As Long
    Dim lnNroPagosExt As Integer
    Dim lnMontoPagosExt As Currency
    Dim lsMontoPagosCab As Currency
    
    ReDim MatCab(1 To 10, 0 To 0)
    ReDim MatDet(1 To 21, 0 To 0)
    ReDim MatDetExt(1 To 21, 0 To 0)

    lineas = 0
    f = FreeFile
    nCerrar = 0
    Open strTextFile For Input As #f
    Do
        Line Input #f, str_Linea
        lineas = lineas + 1
        
        If lineas = 1 Then
            ReDim Preserve MatCab(1 To 10, 0 To 1)
            MatCab(1, 1) = Mid(str_Linea, 1, 2)
            MatCab(2, 1) = Mid(str_Linea, 3, 3)
            MatCab(3, 1) = Mid(str_Linea, 6, 1)
            MatCab(4, 1) = Mid(str_Linea, 7, 7)
            MatCab(5, 1) = Mid(str_Linea, 14, 1)
            MatCab(6, 1) = Mid(str_Linea, 15, 8)
            MatCab(7, 1) = Mid(str_Linea, 23, 9)
            MatCab(8, 1) = Mid(str_Linea, 32, 15)
            MatCab(9, 1) = Mid(str_Linea, 47, 4)
            MatCab(10, 1) = Mid(str_Linea, 51, 6)
        Else
            If Mid(str_Linea, 197, 1) = "E" Then 'Pago Extornado en el BCP
                iMat = UBound(MatDetExt, 2) + 1
                ReDim Preserve MatDetExt(1 To 21, 0 To iMat)
                MatDetExt(1, iMat) = Mid(str_Linea, 1, 2)
                MatDetExt(2, iMat) = Mid(str_Linea, 3, 3)
                MatDetExt(3, iMat) = Mid(str_Linea, 6, 1)
                MatDetExt(4, iMat) = Mid(str_Linea, 7, 7)
                MatDetExt(5, iMat) = Mid(str_Linea, 14, 14)
                MatDetExt(6, iMat) = Mid(str_Linea, 28, 30)
                MatDetExt(7, iMat) = Mid(str_Linea, 58, 8)
                MatDetExt(8, iMat) = Mid(str_Linea, 66, 8)
                MatDetExt(9, iMat) = Mid(str_Linea, 74, 15)
                MatDetExt(10, iMat) = Mid(str_Linea, 89, 15)
                MatDetExt(11, iMat) = Mid(str_Linea, 104, 15)
                MatDetExt(12, iMat) = Mid(str_Linea, 119, 6)
                MatDetExt(13, iMat) = Mid(str_Linea, 125, 6)
                MatDetExt(14, iMat) = Mid(str_Linea, 131, 22)
                MatDetExt(15, iMat) = Mid(str_Linea, 153, 4)
                MatDetExt(16, iMat) = Mid(str_Linea, 157, 12)
                MatDetExt(17, iMat) = Mid(str_Linea, 169, 6)
                MatDetExt(18, iMat) = Mid(str_Linea, 175, 10)
                MatDetExt(19, iMat) = Mid(str_Linea, 185, 2)
                MatDetExt(20, iMat) = Mid(str_Linea, 187, 10)
                MatDetExt(21, iMat) = Mid(str_Linea, 197, 1)
            Else
                iMat = UBound(MatDet, 2) + 1
                ReDim Preserve MatDet(1 To 21, 0 To iMat)
                MatDet(1, iMat) = Mid(str_Linea, 1, 2)
                MatDet(2, iMat) = Mid(str_Linea, 3, 3)
                MatDet(3, iMat) = Mid(str_Linea, 6, 1)
                MatDet(4, iMat) = Mid(str_Linea, 7, 7)
                MatDet(5, iMat) = Mid(str_Linea, 14, 14)
                MatDet(6, iMat) = Mid(str_Linea, 28, 30)
                MatDet(7, iMat) = Mid(str_Linea, 58, 8)
                MatDet(8, iMat) = Mid(str_Linea, 66, 8)
                MatDet(9, iMat) = Mid(str_Linea, 74, 15)
                MatDet(10, iMat) = Mid(str_Linea, 89, 15)
                MatDet(11, iMat) = Mid(str_Linea, 104, 15)
                MatDet(12, iMat) = Mid(str_Linea, 119, 6)
                MatDet(13, iMat) = Mid(str_Linea, 125, 6)
                MatDet(14, iMat) = Mid(str_Linea, 131, 22)
                MatDet(15, iMat) = Mid(str_Linea, 153, 4)
                MatDet(16, iMat) = Mid(str_Linea, 157, 12)
                MatDet(17, iMat) = Mid(str_Linea, 169, 6)
                MatDet(18, iMat) = Mid(str_Linea, 175, 10)
                MatDet(19, iMat) = Mid(str_Linea, 185, 2)
                MatDet(20, iMat) = Mid(str_Linea, 187, 10)
                MatDet(21, iMat) = Mid(str_Linea, 197, 1)
            End If
        End If
    Loop While Not EOF(f)
    Close #f
    
    If UBound(MatCab, 2) = 0 Then
        MsgBox "No se ha podido realizar el proceso de Pago, verifique que sea el archivo correcto," & Chr(13) & "si el error persiste comuniquese con el Dpto. de TI", vbInformation, "Aviso"
        Exit Sub
    End If
    If UBound(MatDet, 2) = 0 Then
        MsgBox "No se realizará el proceso de Pago, ya que en el actual archivo" & Chr(13) & "no existe registros de Pagos Vigentes realizados en el BCP", vbInformation, "Aviso"
        Exit Sub
    End If

    DatosC.c1 = MatCab(1, 1)
    DatosC.c2 = MatCab(2, 1)
    DatosC.c3 = MatCab(3, 1)
    DatosC.c4 = MatCab(4, 1)
    DatosC.c5 = MatCab(5, 1)
    DatosC.c6 = MatCab(6, 1)
    DatosC.c7 = MatCab(7, 1)
    DatosC.c8 = MatCab(8, 1)
    DatosC.c9 = MatCab(9, 1)
    DatosC.c10 = MatCab(10, 1)
    nCerrar = 0
    
    lnNroPagosExt = UBound(MatDetExt, 2)
    For iMat = 1 To lnNroPagosExt
        lnMontoPagosExt = lnMontoPagosExt + CCur(Mid(MatDetExt(9, iMat), 1, Len(MatDetExt(9, iMat)) - 2) & "." & Right(MatDetExt(9, iMat), 2))
    Next
    
    DatosC.c7 = Format(CInt(DatosC.c7) - lnNroPagosExt, "000000000")
    lsMontoPagosCab = Replace(Format(CCur(Mid(DatosC.c8, 1, Len(DatosC.c8) - 2) & "." & Right(DatosC.c8, 2)) - lnMontoPagosExt, "#0.00"), ".", "")
    DatosC.c8 = Format(lsMontoPagosCab, "000000000000000")
    
    If LeerCabecera = False Then
        Set rsX = Nothing
        Exit Sub
    End If
    
    Set rsX = oGen.getDatosCabeceraBCP(rsBCPC("c8"), CInt(rsBCPC("c7")), rsBCPC("c6"))
    Set oGen = Nothing
    
    If Not rsX.EOF And Not rsX.BOF Then
        cod_det = rsX("id")
        MsgBox "Se van a procesar " & RegActual2 & " registros, Esta operación puede demorar unos minutos ...", vbInformation, "Aviso"
    Else
        MsgBox "El archivo no tiene la estructura correcta", vbInformation, "Aviso"
        Exit Sub
    End If
    
    For iMat = 1 To UBound(MatDet, 2)
        If iMat = 1 Then
            oBarra.ShowForm Me
            oBarra.CaptionSyle = eCap_CaptionPercent
            oBarra.Max = RegActual2
            oBarra.Progress 0, "Proceso de Pago Convenio BCP : ", "Preparando Pago...", "Preparando", vbBlue
        Else
            oBarra.Progress iMat, "Proceso de Pago Convenio BCP : ", "Efectuando Pago...", "Efectuando", vbBlue
        End If
        
        With DatosD
            .d1 = MatDet(1, iMat)
            .d2 = MatDet(2, iMat)
            .d3 = MatDet(3, iMat)
            .d4 = MatDet(4, iMat)
            .d5 = MatDet(5, iMat)
            .d6 = MatDet(6, iMat)
            .d7 = MatDet(7, iMat)
            .d8 = MatDet(8, iMat)
            .d9 = MatDet(9, iMat)
            .d10 = MatDet(10, iMat)
            .d11 = MatDet(11, iMat)
            .d12 = MatDet(12, iMat)
            .d13 = MatDet(13, iMat)
            .d14 = MatDet(14, iMat)
            .d15 = MatDet(15, iMat)
            .d16 = MatDet(16, iMat)
            .d17 = MatDet(17, iMat)
            .d18 = MatDet(18, iMat)
            .d19 = MatDet(19, iMat)
            .d20 = MatDet(20, iMat)
            .d21 = MatDet(21, iMat)
            Call LeerDetalle
        End With
    Next
    oBarra.CloseForm Me
    Set rsX = Nothing
    Set oBarra = Nothing
End Sub
'END EJVG *******
Public Function LeerCabecera() As Boolean
    Dim dFechaGrego As String
    Dim bValorProFecha As Boolean
    Dim rsX As ADODB.Recordset
    Dim fd As Date
    Set rsX = New ADODB.Recordset
        
    Dim oCredD As COMDCredito.DCOMCredito
    Set oCredD = New COMDCredito.DCOMCredito
    
    Set rsBCPC = New ADODB.Recordset
    bValorProFecha = True

    With rsBCPC
        .Fields.Append "c1", adVarChar, 2
        .Fields.Append "c2", adVarChar, 3
        .Fields.Append "c3", adVarChar, 1
        .Fields.Append "c4", adVarChar, 7
        .Fields.Append "c5", adVarChar, 1
        .Fields.Append "c6", adVarChar, 8
        .Fields.Append "c7", adVarChar, 9
        .Fields.Append "c8", adCurrency
        .Fields.Append "c9", adVarChar, 4
        .Fields.Append "c10", adVarChar, 6
        .Open
            .AddNew
            .Fields("c1") = DatosC.c1
            .Fields("c2") = DatosC.c2
            .Fields("c3") = DatosC.c3
            .Fields("c4") = DatosC.c4
            .Fields("c5") = DatosC.c5
            .Fields("c6") = DatosC.c6 'Mid(DatosC.c6, 5, 4) & Mid(DatosC.c6, 3, 2) & Mid(DatosC.c6, 1, 2)
            .Fields("c7") = DatosC.c7
            .Fields("c8") = Format(Mid(DatosC.c8, 1, Len(DatosC.c8) - 2), ".") & Right(DatosC.c8, 2)
            .Fields("c9") = DatosC.c9
            .Fields("c10") = DatosC.c10
            RegActual2 = rsBCPC("c7")
            Set rsX = oCredD.getDatosCabeceraBCP(rsBCPC("c8"), CInt(rsBCPC("c7")), rsBCPC("c6"))
            Set oGen = Nothing
    
    End With
              
       dFechaGrego = Mid(DatosC.c6, 7, 2) & "/" & Mid(DatosC.c6, 5, 2) & "/" & Mid(DatosC.c6, 1, 4)

       If Not CDate(dFechaGrego) = gdFecSis Then
            If CDate(dFechaGrego) > gdFecSis Then
                bValorProFecha = False
                MsgBox "No puede procesar el Archivo debido a la Fecha Posterior, Verifique!!!", vbInformation, "Aviso"
            ElseIf Not MsgBox("Seguro de Procesar un Archivo de Fecha Anterior", vbQuestion + vbYesNo, "Aviso") = vbYes Then
                bValorProFecha = False
            End If
       End If

       If bValorProFecha Then
            If rsX.EOF And rsX.BOF Then
               nDiasDiferente = DateDiff("d", CDate(dFechaGrego), gdFecSis)
               Call oCredD.InsertarDatosCabeceraBCP(rsBCPC, nDiasDiferente)
               LeerCabecera = True
               procesados = 0
               n_procesados = 0
            Else
              MsgBox "El archivo ya fue Procesado con Anterioridad, Verifique!!!", vbInformation, "Aviso"
              LeerCabecera = False
            End If
        Else
               LeerCabecera = False
               Exit Function
       End If
       
    rsX.Close
    Set rsX = Nothing
End Function

Sub LeerDetalle()
    Dim oCredD As COMDCredito.DCOMCredito
    Dim nomcli As String
    Dim rsBCPCD As New ADODB.Recordset
    
    Set rsDocs = New ADODB.Recordset
    Set oCredD = New COMDCredito.DCOMCredito
    With rsBCPCD
        .Fields.Append "id_cab", adInteger
        .Fields.Append "cCodCtaF", adVarChar, 18
        .Fields.Append "Titular", adVarChar, 50
        .Fields.Append "Cuota", adInteger
        .Fields.Append "c1", adVarChar, 2
        .Fields.Append "c2", adVarChar, 3
        .Fields.Append "c3", adVarChar, 1
        .Fields.Append "c4", adVarChar, 7
        .Fields.Append "c5", adVarChar, 14
        .Fields.Append "c6", adVarChar, 30
        .Fields.Append "c7", adVarChar, 8
        .Fields.Append "c8", adVarChar, 8
        .Fields.Append "c9", adCurrency
        .Fields.Append "c10", adCurrency
        .Fields.Append "c11", adCurrency
        .Fields.Append "c12", adVarChar, 6
        .Fields.Append "c13", adVarChar, 6
        .Fields.Append "c14", adVarChar, 22
        .Fields.Append "c15", adVarChar, 4
        .Fields.Append "c16", adVarChar, 12
        .Fields.Append "c17", adVarChar, 6
        .Fields.Append "c18", adVarChar, 10
        .Fields.Append "c19", adVarChar, 2
        .Fields.Append "c20", adVarChar, 10
        .Fields.Append "c21", adVarChar, 1
        .Open
    
            .AddNew
            .Fields("id_cab") = cod_det
            .Fields("c1") = DatosD.d1
            .Fields("c2") = DatosD.d2
            .Fields("c3") = DatosD.d3
            .Fields("c4") = DatosD.d4
            .Fields("c5") = DatosD.d5
            .Fields("c6") = DatosD.d6
            .Fields("c7") = DatosD.d7
            .Fields("c8") = DatosD.d8
            .Fields("c9") = Format(Mid(DatosD.d9, 1, Len(DatosD.d9) - 2), ".") + Right(DatosD.d9, 2)
            .Fields("c10") = Format(Mid(DatosD.d10, 1, Len(DatosD.d10) - 2), ".") + Right(DatosD.d10, 2)
            .Fields("c11") = Format(Mid(DatosD.d11, 1, Len(DatosD.d11) - 2), ".") + Right(DatosD.d11, 2)
            .Fields("c12") = DatosD.d12
            .Fields("c13") = DatosD.d13
            .Fields("c14") = DatosD.d14
            .Fields("c15") = DatosD.d15
            .Fields("c16") = DatosD.d16
            .Fields("c17") = DatosD.d17
            .Fields("c18") = DatosD.d18
            .Fields("c19") = DatosD.d19
            .Fields("c20") = DatosD.d20
            .Fields("c21") = DatosD.d21
    End With
    
        Dim rsX As ADODB.Recordset
        Dim rsy As New ADODB.Recordset
        Dim oGen As COMDCredito.DCOMCredito
        Dim NomTemporal As String
        NomTemporal = Trim(DatosD.d6)
        
        Set rsX = New ADODB.Recordset
        Set rsy = New ADODB.Recordset
        Set oGen = New COMDCredito.DCOMCredito
        
        Set rsX = oGen.getDatosCabeceraBCPCta(DatosD.d5, gdFecSis, NomTemporal)

       If Not rsX.EOF And Not rsX.BOF Then
            rsBCPCD.Fields("cCodCtaF") = rsX("cCtaCod")
            rsBCPCD.Fields("Titular") = Mid(rsX("Titular"), 1, 50)
            nomcli = rsBCPCD.Fields("Titular")
            rsBCPCD.Fields("Cuota") = rsX("nCuota")
            Call oCredD.InsertarDatosDetalleBCP(cod_det, rsBCPCD)
            Set oCredD = Nothing
            Set rsy = oGen.getDatosDetalleBcoBCPxId(cod_det, rsBCPCD.Fields("cCodCtaF"))
            If Not rsy.EOF And Not rsy.BOF Then
                Dim cod_dety As Double
                cod_dety = CDbl(rsy("Id"))
                rsy.Close
                Set rsy = Nothing
            End If
            
            If Not CargaDatos(cod_dety, rsBCPCD.Fields("cCodCtaF"), CDbl(rsBCPCD.Fields("c11")), CStr(Format(gdFecSis, "YYYYMMDD")), rsBCPCD.Fields("Cuota")) Then
                   Dim ultm As Integer
                   n_procesados = n_procesados + 1
                   Call oGen.ActualizarDetallexEstBCP(cod_det, rsBCPCD.Fields("cCodCtaF"), rsBCPCD.Fields("Cuota"), gdFecSis, 1)
                     rsBCPCD.Close
                     Set rsBCPCD = Nothing
                   
                   ultm = n_procesados + procesados
                      If n_procesados = RegActual2 Then
                          MsgBox "El Pago no se Realizó" & vbCrLf & "Comuníquese con el Area de TI"
                      ElseIf ultm = RegActual2 Then
                          MsgBox "El Proceso de Pago Finalizó correctamente pero hubieron registros NO procesados, Revise el Reporte!!"
                          Me.cmdLlenar.Enabled = True
                          Me.CmdCargaArch.Enabled = False
                      End If
             End If
       Else
             MsgBox "No se pudo leer los datos de : " & nomcli & " Este registro no será procesado automaticamente"
             n_procesados = n_procesados + 1
             ultm = n_procesados + procesados
             If n_procesados = RegActual2 Then
                MsgBox "No se realizó el Pago de Ningún Registro" & vbCrLf & "Comuniquese con el Area de TI"
                If cod_det > 0 Then
                    Eliminar_IDCabecera cod_det
                End If
             ElseIf ultm = RegActual2 Then
                MsgBox "El Proceso de Pago Finalizó Correctamente, pero hubieron registros NO procesados"
                Me.cmdLlenar.Enabled = True
                Me.CmdCargaArch.Enabled = False
            End If
       End If
     Set oGen = Nothing
 End Sub
Private Function CargaDatos(ByVal cod_det As Double, ByVal psCtaCod As String, ByVal pnImporteCobrado As Double, ByVal pdFechaCobro As String, ByVal pnNroCta As Integer) As Boolean

Dim oCredD As COMDCredito.DCOMCreditos
Dim oCredD1 As COMDCredito.DCOMCredito
Set rsDocs = New ADODB.Recordset
Set oCredD = New COMDCredito.DCOMCreditos
Set oCredD1 = New COMDCredito.DCOMCredito

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
'---------------
Dim oITF As COMDConstSistema.FCOMITF
Dim dif As Double
Dim nNroCta As Integer
Dim FechaCobro As String
Dim FechaCobro1 As Date
Dim NroMov As Long
Dim est As Boolean
Dim nTempITF As Double
Dim MontoPagosinITF As Double
'corresponsalia
Dim Cuota As Double
Dim comision As Double
Dim bInstFinanc As Boolean 'JUEZ 20140411

    On Error GoTo ErrorCargaDatos
    nInteresDesagio = 0
    MontoPagosinITF = 0
    est = False
    Set oCredito = New COMNCredito.NCOMCredito
    If nDiasAtras > 0 Then
        Perdonar_Mora psCtaCod, nDiasAtras
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
                            
               'operacion de egreso --------------------------------------------
                Dim clsMov1 As COMDCaptaGenerales.DCOMCaptaMovimiento
                Dim ClsMov As COMNContabilidad.NCOMContFunciones
                Dim clsServ As COMNCaptaServicios.NCOMCaptaServicios
                Dim sMovNro As String
                Dim nMovNroR As Long
                Dim nMovNroC As Long
                Dim cod_ageContable As String
                Dim oTipoCambio As nTipoCambio
                
                Set ClsMov = New COMNContabilidad.NCOMContFunciones
                sMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

                Set clsServ = New COMNCaptaServicios.NCOMCaptaServicios
                clsServ.OtrasOpeDepCtaBanco sMovNro, "01.1090100824640.0110006", psCtaCod, MontoPagosinITF + nITF, cNomCliente, gsNomAge, IIf(LblMoneda = "SOLES", 1, 2), 300593
                Set clsServ = Nothing
                
               Set clsMov1 = New COMDCaptaGenerales.DCOMCaptaMovimiento
               nMovNroR = clsMov1.GetnMovNro(sMovNro)

              'parte contable -----------------------------------------------------------------
               lsMovNro = ClsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
               Set ClsMov = Nothing
                        
                'INICIO ORCR20140714-------------------------------------------------------------
                Dim psMovDesc As String
                'psMovDesc = psCtaCod & " " & cNomCliente & " " & pnNroCta & " - " & (MontoPagosinITF + nITF)
                'psMovDesc = Right(psCtaCod, 5) & "|" & pnNroCta & "|" & (MontoPagosinITF + nITF) & "|" & cNomCliente
                psMovDesc = Right(psCtaCod, 5) & "|" & pnNroCta & "|" & pnImporteCobradoDesc & "|" & cNomCliente
                '--------------------------------------------------------------------------------
                If LblMoneda = "SOLES" Then
                   cod_ageContable = "2918070101" & gsCodAge
                   'oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "401583", "Deposito Banco", "1113010301",
                    oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "401583", psMovDesc, "1113010301", _
                    cod_ageContable, CCur(MontoPagosinITF + nITF), 0, "", pdFechaCobro, "", _
                    1, "01.1090100824640.0110006", True, , , , , 0, , , , , , , , True, , , , , nMovNroR
                Else
                    cod_ageContable = "2928070101" & gsCodAge
                    'oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "402583", "Deposito Banco", "1123010301",
                    oCredD.GrabaDeposiBancoNegocioFinan lsMovNro, "402583", psMovDesc, "1123010301", _
                    cod_ageContable, CCur(MontoPagosinITF + nITF), 0, "", pdFechaCobro, "", _
                    1, "01.1090100824640.0120004", True, , , , , 0, , , , , , , , True, , , , , nMovNroR
                End If
                'FIN ORCR20140714--------------------------------------------------------------
'
              nMovNroC = clsMov1.GetnMovNro(lsMovNro)
               Set clsMov1 = Nothing
                '------------------------------------------------------------------------------
              'guardar Movs (pago, reversion y contable) --------------------------------------
                If NroMov > 0 And nMovNroR > 0 Then
                   Call oCredD1.ActualizarDetallexEstMovBCP(cod_det, NroMov, nMovNroR, nMovNroC)
                        If (dif - nITF) > 0 Then
                            Call oCredD1.InsertarDatosBCPSobrantes(psCtaCod, nNroCta, dif - nITF, FechaCobro1, cNomCliente, NroMov)
                        End If
                End If
              '----------------------------------------------------------------------------------
               procesados = procesados + 1
               ultm = n_procesados + procesados
               If procesados = RegActual2 Then
                    MsgBox "El Pago Finalizó correctamente" ' & vbCrLf & "Mensaje"
                    Me.cmdLlenar.Enabled = True
                    Me.CmdCargaArch.Enabled = False
              ElseIf ultm = RegActual2 And n_procesados > 0 Then
                    MsgBox "El Pago Finalizó correctamente pero hubieron registros NO procesados, Revise el Reporte!!"
                    Me.cmdLlenar.Enabled = True
                    Me.CmdCargaArch.Enabled = False
               ElseIf ultm = RegActual2 Then
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

Sub cargar_datos()
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCredito
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCredito

    Set rsX = oGen.getCabeceraBcoBCPxID(cod_det)
    Set oGen = Nothing
'id,Tregistro,Tregistro,CSucursal,TMoneda,NCuenta,TValida,FProceso,CRegistros,CBCP,CCasilla,fOperacion,nPerdon,
'Total , nImpToTS, fProceso, tRegistro
     If Not rsX.EOF And Not rsX.BOF Then
        activar_lbl (True)
        Me.lblsoles.Caption = rsX("Total")
        Me.lblMone.Caption = rsX("TMoneda")
        lblfec.Caption = rsX("fProceso")
        Me.lblreg.Caption = procesados
        Me.lbltotal.Caption = rsX("nNumReg")
        Me.lblPSoles.Caption = rsX("nImpToTS1")
     End If
    rsX.Close
    Set rsX = Nothing
End Sub

Sub ConfigurarMShComite()
 Mshbco.Clear
    Mshbco.Cols = 6
    Mshbco.Rows = 2

    With Mshbco
        .TextMatrix(0, 0) = "Id"
        .TextMatrix(0, 1) = "Cuenta"
        .TextMatrix(0, 2) = "Nombre Cliente"
        .TextMatrix(0, 3) = "Cuota"
        .TextMatrix(0, 4) = "Imp.Pagado"
        .TextMatrix(0, 5) = "Imp.PagadoTotal"

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
        .TextMatrix(0, 4) = "Imp.Pagado"
        .TextMatrix(0, 5) = "Imp.PagadoTotal"

        .ColWidth(0) = 400
        .ColWidth(1) = 1800
        .ColWidth(2) = 4500
        .ColWidth(3) = 700
        .ColWidth(4) = 1100
        .ColWidth(5) = 1100

    End With
End Sub


 Sub Cargar_grilla()
    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCredito
    Dim i As Integer
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCredito

    Set rsX = oGen.getDetalleBcoBCPDetallexId(cod_det)
    Set oGen = Nothing
    i = 0
    ConfigurarMShComite
      Do Until rsX.EOF
        With Me.Mshbco
             i = i + 1
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 2, 0) = i 'rsX!id
            .TextMatrix(.Rows - 2, 1) = rsX!cCtaCod
            .TextMatrix(.Rows - 2, 2) = rsX!Titular
            .TextMatrix(.Rows - 2, 3) = rsX!Cuota
            .TextMatrix(.Rows - 2, 4) = rsX!Total
            .TextMatrix(.Rows - 2, 5) = rsX!MontoTotal
        End With
        rsX.MoveNext
    Loop
  Me.cmdImprimir.Enabled = True
  Me.cmdReporte.Enabled = True
  
  rsX.Close
  Set rsX = Nothing
End Sub

Private Sub cmdImprimir_Click()
    Dim previo As previo.clsprevio
    Set previo = New previo.clsprevio
    previo.Show sCadImpre, "Pago Convenio BCP", True
    Set previo = Nothing
End Sub

Private Sub cmdLlenar_Click()
    cargar_datos
    Cargar_grilla
    If n_procesados >= 1 Then
        cargar_datos1
        Cargar_grilla1
    End If
End Sub

Sub cargar_datos1()

    Dim rsX As ADODB.Recordset
    Dim oGen As COMDCredito.DCOMCredito
    Set rsX = New ADODB.Recordset
    Set oGen = New COMDCredito.DCOMCredito

    Set rsX = oGen.getDevolverDatosDatosReporteBCPNoProcesados(cod_det)
    Set oGen = Nothing

     If Not rsX.EOF And Not rsX.BOF Then
        activar_lbl1 (True)
'        Me.lblimpdx.Caption = rsx("nImpToTD")
        Me.lblimpsx.Caption = rsX("nImpToTS")
        Me.lblfecx.Caption = rsX("fProceso")
        Me.lblnumx.Caption = n_procesados
     End If

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
            .TextMatrix(.Rows - 2, 1) = rsX!cCtaCod
            .TextMatrix(.Rows - 2, 2) = rsX!Titular
            .TextMatrix(.Rows - 2, 3) = rsX!Cuota
            .TextMatrix(.Rows - 2, 4) = rsX!Total
            .TextMatrix(.Rows - 2, 5) = rsX!MontoTotal
        End With
        rsX.MoveNext
    Loop
  Me.cmdImprimir.Enabled = True
  Me.cmdReporte.Enabled = True
  
  rsX.Close
  Set rsX = Nothing
End Sub
Private Sub cmdReporte_Click()
 Dim sCadImp As String
    Dim oPrev As previo.clsprevio
    Dim oGen As COMNCredito.NCOMCredito

    Set oPrev = New previo.clsprevio
    Set oGen = New COMNCredito.NCOMCredito

    If (Me.lblfec.Caption) <> "" Then
        sCadImp = oGen.ImprimeClientesRecuperacionesCred(gsCodUser, gdFecSis, CDate(Me.lblfec.Caption), gsNomCmac, cod_det)
    Else
        MsgBox "La fecha para Generar el Reporte No es Válida", vbCritical, "Aviso"
        Exit Sub
    End If
    
    Set oNPers = Nothing
    previo.Show sCadImp, "Registro de Archivo de Retorno de Cobranzas", False
    Set oPrev = Nothing
End Sub

Sub activar_lbl(ByVal xb As Boolean)
Me.lbl1.Visible = xbs
Me.lbl2.Visible = xb
Me.lbl3.Visible = xb
Me.Label5.Visible = xb
Me.Label6.Visible = xb
'Me.Label7.Visible = xb
Me.Label8.Visible = xb
Me.lblsoles.Visible = xb
Me.lblPSoles.Visible = xb
Me.lbltotal.Visible = xb
'Me.lblimpdx.Visible = xb
Me.lblfec.Visible = xb
Me.lblreg.Visible = xb
Me.lblMone.Visible = xb
End Sub

Sub activar_lbl1(ByVal xb As Boolean)
Me.SSTAB.TabVisible(1) = True
Me.Label1.Visible = xb
'Me.Label2.Visible = xb
Me.Label3.Visible = xb
Me.Label4.Visible = xb
Me.lblfecx.Visible = xb
'Me.lblimpdx.Visible = xb
Me.lblimpsx.Visible = xb
Me.lblnumx.Visible = xb
End Sub

Sub Perdonar_Mora(cCtaCod As String, nNumPerdon As Integer)
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
    
    If nNumPerdon > 0 Then
        Set oCalend = New COMDCredito.DCOMCalendario
        Set oCale = New COMDCredito.DCOMCredActBD
        ValorTotalMora = 0#
        ValorTotalMoraDia = 0#
        ValorTotalMoraDiaCalc = 0#
        nDiasAtrasoSist = 0
        nCuota = 0
        nNroCalen = 0
        
        Set R2 = oCalend.RecuperaCalendarioPagosPendiente(psCtaCod)
        Set oCalend = Nothing
        ReDim MatCalendMora(R2.RecordCount, 13)
        
         If R2.RecordCount > 0 Then
                MatCalendMora(R2.Bookmark - 1, 0) = Format(R2!dVenc, "dd/mm/yyyy")
                MatCalendMora(R2.Bookmark - 1, 1) = Trim(Str(R2!nCuota))
                MatCalendMora(R2.Bookmark - 1, 6) = Format(IIf(IsNull(R2!nIntMor), 0, R2!nIntMor), "#0.00")
               R2.Close
               Set R2 = Nothing
               
               ValorTotalMora = MatCalendMora(0, 6)
               nCuota = MatCalendMora(0, 1)
        Else
            Set R2 = Nothing
            Exit Sub
        End If
    
        Dim oCreditoMora As COMDCredito.DCOMCredito
        Set oCreditoMora = New COMDCredito.DCOMCredito
        Set R1 = oCreditoMora.RecuperaColocacCred(psCtaCod)
        Set oCreditoMora = Nothing
        
        If R1.RecordCount > 0 Then
            nDiasAtrasoSist = R1!nDiasAtraso
            nNroCalen = R1!nNroCalen
            R1.Close
        Else
            Set R1 = Nothing
            Exit Sub
        End If
        
        If nDiasAtrasoSist > nNumPerdon Then
              If ValorTotalMora > 0 Then
'                    If (nDiasAtrasoSist - nNumPerdon) = 6 And (nNumPerdon) > 1 Then
'                       Call oCale.dUpdateColocalendDetMora(psCtaCod, nNroCalen, nCuota, 0, "1215")
'                    End If
                    
                    If (nDiasAtrasoSist - nNumPerdon) <= 7 Then
                       Call oCale.dUpdateColocalendDetMora(psCtaCod, nNroCalen, nCuota, 0, "1215")
                    End If
                    
                    ValorTotalMoraDia = Format(ValorTotalMora / nDiasAtrasoSist, "#0.00")
                    ValorTotalMoraDiaCalc = ValorTotalMora - Format(ValorTotalMoraDia * (nNumPerdon), "#0.00")
                    ValorTotalMoraDiaCalc = IIf(ValorTotalMoraDiaCalc <= 0, 0, ValorTotalMoraDiaCalc)
                     If ValorTotalMoraDiaCalc > 0 Then
                          Call oCale.dUpdateColocalendDetMora(psCtaCod, nNroCalen, nCuota, ValorTotalMoraDiaCalc)
                     End If
               End If
        Else
               Call oCale.dUpdateColocalendDetMora(psCtaCod, nNroCalen, nCuota, 0)
        End If
    End If
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
End Sub

Sub Eliminar_IDCabecera(ByVal id As Integer)
Dim oCredD As COMDCredito.DCOMCredito
Set oCredD = New COMDCredito.DCOMCredito

Call oCredD.EliminaCabceraBCP(id)

End Sub
