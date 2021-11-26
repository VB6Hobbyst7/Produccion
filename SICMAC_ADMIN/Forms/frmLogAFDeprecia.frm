VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmLogAFDeprecia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "LOAGISTICA:ACTIVO FIJO:DEPRECIACION DE ACTIVO FIJO"
   ClientHeight    =   6780
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13230
   Icon            =   "frmLogAFDeprecia.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   13230
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTipoActivo 
      Caption         =   "Configurar Tipo Activo"
      Height          =   345
      Left            =   10080
      TabIndex        =   14
      Top             =   6375
      Width           =   1935
   End
   Begin VB.CommandButton cmdNuevoAnio 
      Caption         =   "Nuevo Año"
      Height          =   345
      Left            =   5955
      TabIndex        =   12
      Top             =   7575
      Width           =   960
   End
   Begin VB.CommandButton cmdImpLimaAjustado 
      Caption         =   "&Imp Ajust."
      Height          =   345
      Left            =   6975
      TabIndex        =   11
      Top             =   7575
      Width           =   960
   End
   Begin VB.ComboBox cboTpo 
      Height          =   315
      Left            =   4695
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   0
      Width           =   4020
   End
   Begin VB.CommandButton cmdimpLima 
      Caption         =   "&Imp Hist."
      Height          =   345
      Left            =   7980
      TabIndex        =   8
      Top             =   7575
      Width           =   960
   End
   Begin TabDlg.SSTab Tab 
      Height          =   5865
      Left            =   45
      TabIndex        =   6
      Top             =   420
      Width           =   13170
      _ExtentX        =   23230
      _ExtentY        =   10345
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      BackColor       =   12632256
      TabCaption(0)   =   "Activo Fijo Contable"
      TabPicture(0)   =   "frmLogAFDeprecia.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Flex"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdGrabar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdDeprecia"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdImprimir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "chkSoloEstad"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdVerAsntoCnt"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdExtornaCnt"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdAjuste"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdExtAjuste"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdVerAjustes"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "Activo Fijo Tributario"
      TabPicture(1)   =   "frmLogAFDeprecia.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "cmdExtornaAFTrib"
      Tab(1).Control(1)=   "cmdImpAFTrib"
      Tab(1).Control(2)=   "cmdDepreAFTrib"
      Tab(1).Control(3)=   "cmdGrabarAFTrib"
      Tab(1).Control(4)=   "Flex2"
      Tab(1).ControlCount=   5
      Begin VB.CommandButton cmdVerAjustes 
         Caption         =   "Ver Ajustes"
         Height          =   345
         Left            =   4200
         TabIndex        =   30
         Top             =   5415
         Width           =   960
      End
      Begin VB.CommandButton cmdExtAjuste 
         Caption         =   "Ext.Ajuste"
         Height          =   345
         Left            =   5160
         TabIndex        =   29
         Top             =   5415
         Width           =   960
      End
      Begin VB.CommandButton cmdAjuste 
         Caption         =   "Ajuste"
         Height          =   345
         Left            =   6120
         TabIndex        =   28
         Top             =   5415
         Width           =   960
      End
      Begin VB.CommandButton cmdExtornaCnt 
         Caption         =   "Ext.Depre."
         Height          =   345
         Left            =   7440
         TabIndex        =   27
         Top             =   5415
         Width           =   960
      End
      Begin VB.CommandButton cmdVerAsntoCnt 
         Caption         =   "Ver Asnto.Cnt."
         Height          =   345
         Left            =   8400
         TabIndex        =   26
         Top             =   5415
         Width           =   1200
      End
      Begin VB.CommandButton cmdExtornaAFTrib 
         Caption         =   "Extornar"
         Height          =   345
         Left            =   -68400
         TabIndex        =   25
         Top             =   5415
         Width           =   960
      End
      Begin VB.CheckBox chkSoloEstad 
         Appearance      =   0  'Flat
         Caption         =   "Solo Estadistico (no genera asiento contable)"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   24
         Top             =   5400
         Width           =   3660
      End
      Begin VB.CommandButton cmdImpAFTrib 
         Caption         =   "&Imprimir"
         Height          =   345
         Left            =   -64455
         TabIndex        =   23
         Top             =   5415
         Width           =   720
      End
      Begin VB.CommandButton cmdDepreAFTrib 
         Caption         =   "&Calcular Depreciación"
         Height          =   345
         Left            =   -63795
         TabIndex        =   22
         Top             =   5415
         Width           =   1800
      End
      Begin VB.CommandButton cmdGrabarAFTrib 
         Caption         =   "&Grabar"
         Height          =   345
         Left            =   -65160
         TabIndex        =   21
         Top             =   5415
         Width           =   720
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         Height          =   345
         Left            =   10545
         TabIndex        =   20
         Top             =   5415
         Width           =   720
      End
      Begin VB.CommandButton cmdDeprecia 
         Caption         =   "&Calcular Depreciación"
         Height          =   345
         Left            =   11205
         TabIndex        =   19
         Top             =   5415
         Width           =   1800
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   345
         Left            =   9840
         TabIndex        =   18
         Top             =   5415
         Width           =   720
      End
      Begin MSDataGridLib.DataGrid Flex 
         Height          =   4815
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   8493
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin MSDataGridLib.DataGrid Flex2 
         Height          =   4815
         Left            =   -74760
         TabIndex        =   17
         Top             =   480
         Width           =   12735
         _ExtentX        =   22463
         _ExtentY        =   8493
         _Version        =   393216
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdAgregar 
      Caption         =   "&Agregar"
      Height          =   345
      Left            =   165
      TabIndex        =   5
      Top             =   6375
      Width           =   960
   End
   Begin VB.ComboBox cmbMes 
      Height          =   315
      ItemData        =   "frmLogAFDeprecia.frx":0342
      Left            =   2175
      List            =   "frmLogAFDeprecia.frx":0344
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   30
      Width           =   1980
   End
   Begin MSMask.MaskEdBox mskAnio 
      Height          =   300
      Left            =   600
      TabIndex        =   2
      Top             =   60
      Width           =   885
      _ExtentX        =   1561
      _ExtentY        =   529
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   4
      Mask            =   "####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   12135
      TabIndex        =   0
      Top             =   6375
      Width           =   960
   End
   Begin Sicmact.FlexEdit flexRes 
      Height          =   765
      Left            =   720
      TabIndex        =   16
      Top             =   7920
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   1349
      Cols0           =   11
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Año-V_Historico-V_Ajustado-V_H_Acum-V_A_Acum-V_H_Mes-V_A_Mes-Ajuste_Mes-V_H_Acum_A-V_A_Acum_A"
      EncabezadosAnchos=   "300-800-1200-1200-1000-1000-1000-1000-1000-1000-1000"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-R-R-R-R-R-R-R-R-R"
      FormatosEdit    =   "0-0-2-2-2-2-2-2-2-2-2"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      Appearance      =   0
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblEstadoMes 
      AutoSize        =   -1  'True
      BorderStyle     =   1  'Fixed Single
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
      Height          =   300
      Left            =   8760
      TabIndex        =   15
      Top             =   0
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Label lblTipo 
      AutoSize        =   -1  'True
      Caption         =   "Tipo :"
      Height          =   195
      Left            =   4215
      TabIndex        =   10
      Top             =   60
      Width           =   405
   End
   Begin VB.OLE OleExcel 
      Appearance      =   0  'Flat
      AutoActivate    =   3  'Automatic
      Enabled         =   0   'False
      Height          =   240
      Left            =   9000
      SizeMode        =   1  'Stretch
      TabIndex        =   7
      Top             =   7680
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.Label lblMes 
      AutoSize        =   -1  'True
      Caption         =   "Mes :"
      Height          =   195
      Left            =   1680
      TabIndex        =   4
      Top             =   75
      Width           =   390
   End
   Begin VB.Label lblAnio 
      Caption         =   "Año :"
      Height          =   210
      Left            =   135
      TabIndex        =   1
      Top             =   90
      Width           =   705
   End
End
Attribute VB_Name = "frmLogAFDeprecia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsCaption As String
Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet
Dim lRs As ADODB.Recordset
Dim lRsTr As ADODB.Recordset
Dim nTotal As Double
Dim nTotalTr As Double

Dim nMes As Integer
Dim nAnio As Integer
Dim dFecha As Date

Public Sub Ini(psCaption As String)
    lsCaption = psCaption
    Me.Show 1
End Sub


Private Sub cboTpo_Click()

    If Val(Right(Me.cboTpo.Text, 3)) = 10 Then
        Me.Tab.TabVisible(0) = False
        Me.Tab.TabVisible(1) = False
'        Me.Tab.TabVisible(2) = True
'        Me.FlJoyas.SetFocus
        Me.Refresh
    Else
        Me.Tab.TabVisible(0) = True
        Me.Tab.TabVisible(1) = True
'        Me.Tab.TabVisible(2) = True
    End If
End Sub

Private Sub cmdAgregar_Click()
    frmLogAFMant.Show 1
End Sub

Private Sub cmdAjuste_Click()
    Dim oDep As DMov
    Set oDep = New DMov
    
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim oconect As DConecta
    Set oconect = New DConecta
    
    Dim lnMovNro As Long
    Dim lsMovNro As String
    
    Dim lnMovNroDif As Long
    Dim lsMovNroDif As String
    
    Dim lnMovNroDifAjuste As Long
    Dim lsMovNroDifAjuste As String
    
    Dim lnMovNroR As Long
    Dim lsMovNroR As String
    Dim lnMovNroDifAjusteR As Long
    Dim lsMovNroDifAjusteR As String
    Dim lsTipo As String
    Dim lsFecha As String
    Dim I As Integer '*** PEAC 20090924
    Dim lnI As Long
    Dim lnContador As Long
    Dim lsCtaCont As String
    Dim oPrevio As clsPrevio
    Dim oAsiento As NContImprimir
    Set oPrevio = New clsPrevio
    Set oAsiento = New NContImprimir
    
    Dim oJDep As DLogBieSer '*** PEAC 20090923
    Set oJDep = New DLogBieSer '*** PEAC 20090923
    
    Dim lcTextoCntEstad As String
    
    Dim rs As ADODB.Recordset
    Dim lcDebe As String, lcHaber As String, lcCodBSJoyAdj As String '*** PEAC 20090923
    
    Dim ldFechaDepre As Date
    
    Dim lcSerieAjuste As String
    Dim lnMontoAjuste As String
    
    lcTextoCntEstad = 0


    lsTipo = Trim(Right(cboTpo.Text, 2))
    
    If lsTipo <> "10" And nTotal = 0 Then
        MsgBox "No existen datos para procesar", vbCritical, "Atención"
        Exit Sub
    End If

    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If

    If lRs.Fields(19) <= 0 Then 'And CDbl(lRs.Fields(17)) <> 0 Then
        MsgBox "Seleccione un bien.", vbCritical, "Atención"
        Exit Sub
    End If

    Set rs = New ADODB.Recordset

    Set rs = BuscaRegAjusteBSAF(gnDepAF, lsFecha, lsTipo, lRs.Fields(1))

    If Not rs.EOF Then
        MsgBox "Ajuste ya generado a este Bien.", vbCritical, "Aviso!"
        Exit Sub
    End If

    lnMontoAjuste = InputBox("Serie:" + lRs.Fields(1) + Chr(13) + _
                             "Descripción:" + lRs.Fields(3), "Ingreso de Monto de Ajuste")

    If lnMontoAjuste = "" Then Exit Sub
    

    If Not IsNumeric(lnMontoAjuste) Then

        MsgBox "Ingrese por favor solo numeros.", vbCritical, "Atención"
        Exit Sub

'    ElseIf CDbl(lnMontoAjuste) <= 0 Then
'        MsgBox "Ingrese un monto mayor a cero.", vbCritical, "Atención"
'        Exit Sub
    End If
    
    If MsgBox("¿Desea Grabar el Ajuste? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
    
    ldFechaDepre = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(Me.cmbMes.Text, 2) & "/" & Me.mskAnio.Text)))
    
    oDep.BeginTrans

        lsMovNro = oDep.GeneraMovNro(ldFechaDepre, Right(gsCodAge, 2), gsCodUser)
        oDep.InsertaMov lsMovNro, gnDepAF, "Ajuste Depreciación de Activo Fijo " & Format(ldFechaDepre, gsFormatoFechaView), 25
        
        lnMovNro = oDep.GetnMovNro(lsMovNro)
        
        If Round(lRs.Fields(19), 2) > 0 Then
            
                oDep.InsertaMovBSAF Me.mskAnio.Text, lRs.Fields(19), 1, lRs.Fields(20), lRs.Fields(1), lnMovNro, lsTipo
                lsCtaCont = oDep.GetOpeCtaCtaOtro(gnDepAF, Left(lRs.Fields(2), Len(lRs.Fields(2)) - 2) & "AG", "", False)
                lsCtaCont = Replace(lsCtaCont, "AG", Right(lRs.Fields(2), 2))
                oDep.InsertaMovCta lnMovNro, 1, lsCtaCont, Round(CDbl(lnMontoAjuste), 2)
            
        End If

        If Round(lRs.Fields(19), 2) > 0 Then
            
                lsCtaCont = oDep.GetOpeCtaCta(gnDepAF, Left(lRs.Fields(2), Len(lRs.Fields(2)) - 2) & "AG", "")

                lsCtaCont = Replace(lsCtaCont, "AG", Right(lRs.Fields(2), 2))
                
                If lsCtaCont = "18190701" Then
                    lsCtaCont = "1819070102"
                End If
                oDep.InsertaMovCta lnMovNro, 2, lsCtaCont, Round(CDbl(lnMontoAjuste), 2) * -1
                
        End If

    oDep.CommitTrans

    MsgBox "Ajuste Depreciación se generó OK.", vbInformation + vbOKOnly, "Atención"
    
    Call DepreContAF

End Sub

Private Sub cmdDepreAFTrib_Click()
    Call DepreTribuAf
End Sub

Private Sub DepreTribuAf()
    Dim oDep As DLogDeprecia
    Dim ldFecha As Date
    Set oDep = New DLogDeprecia

    Dim lsFecha, lsTipo As String
    Dim lnDepreciado As Integer
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    If Not IsNumeric(Me.mskAnio.Text) Then
        MsgBox "Debe Ingresar un año Valido.", vbInformation, "Aviso"
        Me.mskAnio.SetFocus
        Exit Sub
    ElseIf Me.cmbMes.Text = "" Then
        MsgBox "Debe Ingresar un mes Valido.", vbInformation, "Aviso"
        Me.cmbMes.SetFocus
        Exit Sub
    ElseIf Me.cboTpo.Text = "" Then
        MsgBox "Debe Ingresar un tipo de depreciacion Valido.", vbInformation, "Aviso"
        Me.cboTpo.SetFocus
        Exit Sub
    End If
    
    '----------------------------------------
    lsTipo = Trim(Right(cboTpo.Text, 2))
    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If
    
    Set rs = GetBSAFAsiento(gnDepTributAF, lsFecha, lsTipo)

    If Not rs.EOF Then
        Me.lblEstadoMes.Visible = True
        Me.lblEstadoMes.Caption = "Mes Ya fue Depreciado"
        Me.lblEstadoMes.ForeColor = &HFF&
        lnDepreciado = 1
    Else
        Me.lblEstadoMes.Visible = True
        Me.lblEstadoMes.Caption = "Mes sin Depreciar"
        Me.lblEstadoMes.ForeColor = &HFF0000
        lnDepreciado = 0
    End If
    '----------------------------------------
    
    ldFecha = CDate("01/" & Format(Trim(Right(Me.cmbMes.Text, 5)), "00") & "/" & Me.mskAnio.Text)

    nTotalTr = 0
    Set lRsTr = New ADODB.Recordset
    Set lRsTr = oDep.GetAFDeprecia(ldFecha, Right(Me.cboTpo.Text, 3), mskAnio.Text, , , 2, lnDepreciado)
    Set Me.Flex2.DataSource = lRsTr
    nTotalTr = lRsTr.RecordCount

    MsgBox "Se cargaron los datos de la Depreciación Tributaria OK.", vbInformation, "Aviso"

End Sub

Private Sub cmdDeprecia_Click()
    Call DepreContAF
End Sub

'*** PEAC 20120612
Private Sub DepreContAF()
    Dim oDep As DLogDeprecia
    Dim ldFecha As Date
    Set oDep = New DLogDeprecia

    Dim lsFecha, lsTipo As String
    Dim lnDepreciado As Integer
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

'    '*** PEAC 20090922
'    Dim oJDep As DLogBieSer
'    Set oJDep = New DLogBieSer
'    '*** FIN PEAC
    
    If Not IsNumeric(Me.mskAnio.Text) Then
        MsgBox "Debe Ingresar un año Valido.", vbInformation, "Aviso"
        Me.mskAnio.SetFocus
        Exit Sub
    ElseIf Me.cmbMes.Text = "" Then
        MsgBox "Debe Ingresar un mes Valido.", vbInformation, "Aviso"
        Me.cmbMes.SetFocus
        Exit Sub
    ElseIf Me.cboTpo.Text = "" Then
        MsgBox "Debe Ingresar un tipo de depreciacion Valido.", vbInformation, "Aviso"
        Me.cboTpo.SetFocus
        Exit Sub
    End If
    
    '----------------------------------------
    lsTipo = Trim(Right(cboTpo.Text, 2))
    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If
    
    Set rs = GetBSAFAsiento(gnDepAF, lsFecha, lsTipo)

    If Not rs.EOF Then
        Me.lblEstadoMes.Visible = True
        Me.lblEstadoMes.Caption = "Mes Depreciado" + IIf(rs!nEstadCnt = 1, "/Con Asnto Cnt.", "/Sin Asnto Cnt.(Estadistico)")
        Me.lblEstadoMes.ForeColor = &HFF&
        lnDepreciado = 1
    Else
        Me.lblEstadoMes.Visible = True
        Me.lblEstadoMes.Caption = "Mes Sin Depreciar"
        Me.lblEstadoMes.ForeColor = &HFF0000
        lnDepreciado = 0
    End If
    '----------------------------------------
    
    ldFecha = CDate("01/" & Format(Trim(Right(Me.cmbMes.Text, 5)), "00") & "/" & Me.mskAnio.Text)

    Me.Caption = "LOGISTICA: DEPRECIACION DE ACTIVOS FIJOS" + " *** ACTUALIZANDO DATOS ***"

    nTotal = 0
    Set lRs = New ADODB.Recordset
    Set lRs = oDep.GetAFDeprecia(ldFecha, Right(Me.cboTpo.Text, 3), mskAnio.Text, , , 1, lnDepreciado)
    Set Me.Flex.DataSource = lRs
    nTotal = lRs.RecordCount
    
    Me.Caption = "LOGISTICA: DEPRECIACION DE ACTIVOS FIJOS"
    
'    nTotalTr = 0
'    Set lRsTr = New ADODB.Recordset
'    Set lRsTr = oDep.GetAFDeprecia(ldFecha, Right(Me.cboTpo.Text, 3), mskAnio.Text, , , 2, lnDepreciado)
'    Set Me.Flex2.DataSource = lRsTr
'    nTotalTr = lRsTr.RecordCount

    MsgBox "Se cargó los datos de la Depreciación OK.", vbInformation, "Aviso"

End Sub

Private Sub cmdExtAjuste_Click()
    Dim oDep As DMov
    Set oDep = New DMov
    
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim oconect As DConecta
    Set oconect = New DConecta
    
    Dim lnMovNro As Long
    Dim lsMovNro As String
    
    Dim lnMovNroDif As Long
    Dim lsMovNroDif As String
    
    Dim lnMovNroDifAjuste As Long
    Dim lsMovNroDifAjuste As String
    
    Dim lnMovNroR As Long
    Dim lsMovNroR As String
    Dim lnMovNroDifAjusteR As Long
    Dim lsMovNroDifAjusteR As String
    Dim lsTipo As String
    Dim lsFecha As String
    Dim I As Integer '*** PEAC 20090924
    Dim lnI As Long
    Dim lnContador As Long
    Dim lsCtaCont As String
    Dim oPrevio As clsPrevio
    Dim oAsiento As NContImprimir
    Set oPrevio = New clsPrevio
    Set oAsiento = New NContImprimir
    
    Dim oJDep As DLogBieSer '*** PEAC 20090923
    Set oJDep = New DLogBieSer '*** PEAC 20090923
    
    Dim lcTextoCntEstad As String
    
    Dim rs As ADODB.Recordset
    Dim lcDebe As String, lcHaber As String, lcCodBSJoyAdj As String '*** PEAC 20090923
    
    Dim ldFechaDepre As Date
    
    Dim lcSerieAjuste As String
    Dim lnMontoAjuste As String
    
    'Me.Flex.SetFocus
    
    lcTextoCntEstad = 0
    
    lsTipo = Trim(Right(cboTpo.Text, 2))
    
    If lsTipo <> "10" And nTotal = 0 Then
        MsgBox "No existen datos para procesar", vbCritical, "Atención"
        Exit Sub
    End If

    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If

    If lRs.Fields(19) <= 0 Then 'And CDbl(lRs.Fields(17)) <> 0 Then
        MsgBox "Seleccione un bien.", vbCritical, "Atención"
        Exit Sub
    End If
    
    Set rs = New ADODB.Recordset

    Set rs = BuscaRegAjusteBSAF(gnDepAF, lsFecha, lsTipo, lRs.Fields(1))

    If (rs.EOF And rs.BOF) Then
        MsgBox "No se realizó Ajuste a este bien.", vbCritical, "Aviso!"
        Exit Sub
    End If

    If MsgBox("¿Desea Extornar el Ajuste de este bien? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
    oDep.EliminaMov rs!nMovNro
    MsgBox "Ajuste de este bien fue extornado OK.", vbInformation + vbOKOnly, "Atención"
    Call DepreContAF
    
End Sub

Private Sub cmdExtornaAFTrib_Click()
    
    Dim lsFecha As String
    Dim lsTipo As String
    Dim oDep As DMov
    Set oDep = New DMov
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    
    lsTipo = Trim(Right(cboTpo.Text, 2))
    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If
    
    
    Set rs = BuscaMovDepreAF(gnDepTributAF, lsFecha, lsTipo)
    
    If (rs.EOF And rs.BOF) Then
        MsgBox "Este periodo no tiene Depreciación Tributaria.", vbCritical, "Aviso!"
        Exit Sub
    End If
    
    If MsgBox("¿Desea extornar la Depreciación Tributaria de este periodo? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub

    oDep.EliminaMov rs!nMovNro
'    Set oDep = Nothing

    MsgBox "Se extornó la Depreciación Tributaria de este periodo.", vbInformation + vbOKOnly, "Atención"

    Call DepreTribuAf

End Sub

Private Sub cmdExtornaCnt_Click()
    
    Dim lsFecha As String
    Dim lsTipo As String
    Dim oDep As DMov
    Set oDep = New DMov
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    
    lsTipo = Trim(Right(cboTpo.Text, 2))
    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If lsTipo <> "10" And nTotal = 0 Then
        MsgBox "No existen datos para extornar.", vbCritical, "Atención"
        Exit Sub
    End If
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If

    Set rs = GetBSAFAsiento(gnDepAF, lsFecha, lsTipo)
    
    If (rs.EOF And rs.BOF) Then
        MsgBox "Este periodo no fue depreciado.", vbCritical, "Aviso!"
        Exit Sub
    Else
        If rs!nEstadCnt = 1 Then
            MsgBox "La Depreciación tiene asiento contable, debe ser extornado por el Financiero," & Chr(13) & "coordine con el area de Contabilidad.", vbInformation + vbOKOnly, "Atención"
            Exit Sub
        ElseIf rs!nAjuste = 1 Then
            MsgBox "Este Activo Fijo tiene Ajustes, verifique y extorne los ajustes para continuar.", vbInformation + vbOKOnly, "Atención"
            Exit Sub
        Else
            If MsgBox("¿Desea extornar la Depreciación Contable (solo estadistico) de este periodo? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
            oDep.EliminaMov rs!nMovNro
            MsgBox "Se extornó la Depreciación Contable (solo estadístico).", vbInformation + vbOKOnly, "Atención"
            Call DepreContAF
        End If
    End If
    
    Call DepreContAF

End Sub

'Private Sub CmdEliminar_Click()
'    frmLogAFMant.Ini 1, False, False, True, Me.Flex.TextMatrix(Me.Flex.Row, 1), Me.Flex.TextMatrix(Me.Flex.Row, 3), CDate(Me.Flex.TextMatrix(Me.Flex.Row, 18)), Me.Flex.TextMatrix(Me.Flex.Row, 4), 0, 0, Me.Flex.TextMatrix(Me.Flex.Row, 7), Me.Flex.TextMatrix(Me.Flex.Row, 8), Me.Flex.TextMatrix(Me.Flex.Row, 19) & Me.Flex.TextMatrix(Me.Flex.Row, 20), Me.Flex.TextMatrix(Me.Flex.Row, 21), Me.Flex.TextMatrix(Me.Flex.Row, 17)
'End Sub

Private Sub cmdGrabar_Click()
    Dim oDep As DMov
    Set oDep = New DMov
    
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim oconect As DConecta
    Set oconect = New DConecta
    
    Dim lnMovNro As Long
    Dim lsMovNro As String
    
    Dim lnMovNroDif As Long
    Dim lsMovNroDif As String
    
    Dim lnMovNroDifAjuste As Long
    Dim lsMovNroDifAjuste As String
    
    Dim lnMovNroR As Long
    Dim lsMovNroR As String
    Dim lnMovNroDifAjusteR As Long
    Dim lsMovNroDifAjusteR As String
    Dim lsTipo As String
    Dim lsFecha As String
    Dim I As Integer '*** PEAC 20090924
    Dim lnI As Long
    Dim lnContador As Long
    Dim lsCtaCont As String
    Dim oPrevio As clsPrevio
    Dim oAsiento As NContImprimir
    Set oPrevio = New clsPrevio
    Set oAsiento = New NContImprimir
    Dim oContFunc As NContFunciones
    Dim oJDep As DLogBieSer '*** PEAC 20090923
    Set oJDep = New DLogBieSer '*** PEAC 20090923
    
    Dim lcCtaInexistente, lcCtaUltNivel As String
    
    Dim lcTextoCntEstad As String
    
    Dim rs As ADODB.Recordset
    Dim lcDebe As String, lcHaber As String, lcCodBSJoyAdj As String '*** PEAC 20090923
    
    Dim ldFechaDepre As Date
    Dim sSql As String
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lcTextoValida As String
    
    Dim rsUlt As ADODB.Recordset
    Set rsUlt = New ADODB.Recordset
    
    Dim rsExis As ADODB.Recordset
    Set rsExis = New ADODB.Recordset

    lcTextoCntEstad = 0

'    '*** PEAC 20090923
    'lsTipo = Right(cboTpo.Text, 1)
    lsTipo = Trim(Right(cboTpo.Text, 2))
'    '*** FIN PEAC

'    '*** PEAC 20090923
    If lsTipo <> "10" And nTotal = 0 Then
        MsgBox "No existen datos para procesar", vbCritical, "Atención"
        Exit Sub
'    ElseIf Len(Me.FlJoyas.TextMatrix(2, 1)) = 0 Then
'        MsgBox "No existen datos para procesar", vbCritical, "Atención"
'        Exit Sub
    End If
'    '*** FIN PEAC
    
    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If

    Set rs = New ADODB.Recordset

''*** PEAC 20090923 -----------------------------------------------------------------------
'    If Me.Tab.TabVisible(2) Then
'
'        Set rs = oJDep.VerificaAsientoDepreJoyAdj(gnDepJoyAdj, lsFecha, lsTipo)
'
'        If Not rs.EOF Then
'            MsgBox "Asiento ya fue generado ", vbCritical, "Atención"
'            Exit Sub
'        End If
'
'        Set rs = Nothing
'        Set rs = oJDep.ObtieneCtasParaAsientoContaDepreJoyAdj()
'        If (rs.EOF And rs.BOF) Then
'            MsgBox "No se encontró el codigo de Bienes Servicios  - Bienes Adjudicados - Oro para la depreciación de joyas adjudicadas.", vbCritical, "Atención"
'            Exit Sub
'        End If
'        lcCodBSJoyAdj = Trim(rs!cBSCod)
'
'        Set rs = Nothing
'        Set rs = oJDep.ObtieneCtasParaAsientoContaDepreJoyAdj()
'        If (rs.EOF And rs.BOF) Then
'            MsgBox "No se encontró las cuentas del debe y haber para trabajar la depreciación de joyas adjudicadas.", vbCritical, "Atención"
'            Exit Sub
'        End If
'
'        If Len(Trim(rs!cdebe)) = 0 Or Len(Trim(rs!chaber)) = 0 Then
'            MsgBox "Una de las cuentas está vacio. Debe : " & Trim(rs!cdebe) & " Haber : " & Trim(rs!chaber), vbCritical, "Atención"
'            Exit Sub
'        End If
'
'        lcDebe = Trim(rs!cdebe): lcHaber = Trim(rs!chaber)
'        'debe 431402010101 haber 161902010102
'
'        If MsgBox("¿Desea Procesar el asiento de Depreciación de Joyas adjudicadas? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
'        ldFechaDepre = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(Me.cmbMes.Text, 2) & "/" & Me.mskAnio.Text)))
'
'        oDep.BeginTrans
'            lsMovNro = oDep.GeneraMovNro(ldFechaDepre, Right(gsCodAge, 2), gsCodUser)
'            oDep.InsertaMov lsMovNro, gnDepJoyAdj, "Depreciación Mensual de Joyas Adjudicadas" & Format(ldFechaDepre, gsFormatoFechaView), 10
'            lnMovNro = oDep.GetnMovNro(lsMovNro)
'            i = 0
'            For lnI = 1 To Me.FlJoyas.Rows - 2
'                i = i + 1
'                If Trim(Me.FlJoyas.TextMatrix(lnI, 2)) <> "TOTAL" Then
'                    oDep.InsertaMovBSAF Me.mskAnio.Text, Me.FlJoyas.TextMatrix(lnI, 14), i, lcCodBSJoyAdj, Me.FlJoyas.TextMatrix(lnI, 2), lnMovNro, CStr(lsTipo)
'                ElseIf Trim(Me.FlJoyas.TextMatrix(lnI, 2)) = "TOTAL" Then
'                    lsCtaCont = lcDebe + Trim(Me.FlJoyas.TextMatrix(lnI, 13))
'                    oDep.InsertaMovCta lnMovNro, i, lsCtaCont, Round(Me.FlJoyas.TextMatrix(lnI, 8), 2)
'                End If
'            Next lnI
'
'           lnContador = i
'
'            For lnI = 1 To Me.FlJoyas.Rows - 2
'                lnContador = lnContador + 1
'                If Trim(Me.FlJoyas.TextMatrix(lnI, 2)) = "TOTAL" Then
'                    lsCtaCont = lcHaber + Trim(Me.FlJoyas.TextMatrix(lnI, 13))
'                    oDep.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Round(Me.FlJoyas.TextMatrix(lnI, 8), 2) * -1
'                End If
'            Next lnI
'        oDep.CommitTrans
'        oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNro, 60, 80, "DEPRECIACION DE JOYAS ADJUDICADAS"), "DEPRECIACION DE JOYAS ADJUDICADAS", True
'        Exit Sub
'    End If
'
''*** FIN PEAC ------------------------------------------------------------------------
    
    Set rs = GetBSAFAsiento(gnDepAF, lsFecha, lsTipo)
    
    If Not rs.EOF Then
        MsgBox "Asiento ya generado ", vbCritical, "Aviso!"
        Exit Sub
    End If
        
'***-----------------------------------------------------------------------------------
If Me.chkSoloEstad.value = 0 Then
    nMes = Val(Trim(Right(cmbMes.Text, 2)))
    nAnio = Val(mskAnio.Text)
    dFecha = DateAdd("m", 1, "01/" & Format(nMes, "00") & "/" & Format(nAnio, "0000")) - 1
    
    Set oContFunc = New NContFunciones
    If Not oContFunc.PermiteModificarAsiento(Format(dFecha, gsFormatoMovFecha), False) Then
       Set oContFunc = Nothing
       MsgBox "Imposible grabar el asiento en un mes cerrado.", vbInformation, "Aviso"
       Exit Sub
    End If
End If
'**------------------------------------------------------------------------------------
    
    Set rs = GetBSAFCorrelativoAsiento(gnDepAF, lsFecha, lsTipo)
    
    If rs!cPeriodo <> "" And rs!cPeriodo <> lsFecha Then
        MsgBox "Para Depreciar el presente periodo debe estar Depreciado el anterior.", vbCritical, "Aviso!"
        Exit Sub
    ElseIf rs!cPeriodo = "" Then
        MsgBox "Este periodo es el primero en Depreciar por lo que en adelante se realizará correlativamente.", vbInformation + vbOKOnly, "Aviso!"
    End If
    
'**------------------------------------------------------------------------------------
    
    If MsgBox("¿Desea Procesar? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
    
    ldFechaDepre = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(Me.cmbMes.Text, 2) & "/" & Me.mskAnio.Text)))


    '--------------    VALIDACION ANTES DE GRABAR
    oCon.AbreConexion

    lcCtaInexistente = "": lcCtaUltNivel = ""
    lRs.MoveFirst
    For lnI = 0 To nTotal - 1
                
    Me.Caption = "LOGISTICA: DEPRECIACION DE ACTIVOS FIJOS" + " *** VALIDANDO *** " + IIf(lnI Mod 4 = 0, "|", IIf(lnI Mod 4 = 1, "/", IIf(lnI Mod 4 = 2, "-", IIf(lnI Mod 4 = 3, "\", "|"))))
                
        If Round(lRs.Fields(18), 2) > 0 And CDbl(lRs.Fields(19)) <> 0 Then
        
            lsCtaCont = oDep.GetOpeCtaCtaOtro(gnDepAF, Left(lRs.Fields(2), Len(lRs.Fields(2)) - 2) & "AG", "", False)
            If lsCtaCont = "" Then
                MsgBox "La Cta. Cont ''" & Left(lRs.Fields(2), Len(lRs.Fields(2)) - 2) & "AG" & "'' de la Operación ''" & gnDepAF & "'' del bien con serie Nº " & lRs.Fields(1) & Chr(10) & "No existe en la tabla ''OpeCtaCta'', Comunique a Sistemas."
                Exit Sub
            End If
            lsCtaCont = Replace(lsCtaCont, "AG", Right(lRs.Fields(2), 2))
            sSql = "exec stp_sel_ValidaSiCtaContExisteEnPlan '" & lsCtaCont & "'"
            Set rsExis = oCon.CargaRecordSet(sSql)
            If (rsExis.EOF And rsExis.BOF) Then
                lcCtaInexistente = lcCtaInexistente + lRs.Fields(1) + Space(10) + lsCtaCont + oImpresora.gPrnSaltoLinea
            Else
                sSql = "stp_sel_ValidaSiCtaContEsUltimoNivel '" & lsCtaCont & "'"
                Set rsUlt = oCon.CargaRecordSet(sSql)
                If Not (rsUlt.EOF And rsUlt.BOF) Then
                    lcCtaUltNivel = lcCtaUltNivel + lRs.Fields(1) + Space(10) + lsCtaCont + oImpresora.gPrnSaltoLinea
                End If
            End If
        End If
    lRs.MoveNext
    Next

    If Len(lcCtaUltNivel) + Len(lcCtaUltNivel) > 0 Then

        lcTextoValida = lcTextoValida & " VALIDACION DE CUENTAS CONTABLES" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        lcTextoValida = lcTextoValida & " CUENTAS INEXISTENTES EN EL PLAN CONTABLE" & oImpresora.gPrnSaltoLinea
        lcTextoValida = lcTextoValida & " ----------------------------------------" & oImpresora.gPrnSaltoLinea
        lcTextoValida = lcTextoValida & lcCtaInexistente & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        lcTextoValida = lcTextoValida & " CUENTAS QUE NO ESTAN EN EL ULTIMO NIVEL" & oImpresora.gPrnSaltoLinea
        lcTextoValida = lcTextoValida & " ---------------------------------------" & oImpresora.gPrnSaltoLinea
        lcTextoValida = lcTextoValida & lcCtaUltNivel & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea

        oPrevio.Show lcTextoValida, Caption, True

        Exit Sub
    Else
        MsgBox "Se validó existencia de Ctas. Contables en el Plan y Ctas. en ultimo nivel.", vbInformation + vbOKOnly, "Aviso"
    End If

    Me.Caption = "LOGISTICA: DEPRECIACION DE ACTIVOS FIJOS"
    '--------------    FIN VALIDACION
    
    
    oDep.BeginTrans

        lsMovNro = oDep.GeneraMovNro(ldFechaDepre, Right(gsCodAge, 2), gsCodUser)
        oDep.InsertaMov lsMovNro, gnDepAF, "Depreciación Mensual de Activo Fijo " & Format(ldFechaDepre, gsFormatoFechaView), 25
        lnMovNro = oDep.GetnMovNro(lsMovNro)
'        lRs.MoveFirst
        
        For lnI = 0 To nTotal - 1

            Me.Caption = "LOGISTICA: DEPRECIACION DE ACTIVOS FIJOS" + " *** GENERANDO ASNTO.CONT *** " + IIf(lnI Mod 4 = 0, "|", IIf(lnI Mod 4 = 1, "/", IIf(lnI Mod 4 = 2, "-", IIf(lnI Mod 4 = 3, "\", "|"))))

            If lnI = 0 Then
                lRs.MoveFirst
            End If

                If Round(lRs.Fields(18), 2) > 0 And CDbl(lRs.Fields(19)) <> 0 Then
                    
                        oDep.InsertaMovBSAF Me.mskAnio.Text, lRs.Fields(19), lnI, lRs.Fields(20), lRs.Fields(1), lnMovNro, lsTipo
                        
                        lsCtaCont = oDep.GetOpeCtaCtaOtro(gnDepAF, Left(lRs.Fields(2), Len(lRs.Fields(2)) - 2) & "AG", "", False)
                        
                        lsCtaCont = Replace(lsCtaCont, "AG", Right(lRs.Fields(2), 2))
                        
                        oDep.InsertaMovCta lnMovNro, lnI, lsCtaCont, Round(lRs.Fields(18), 2)
                    
                End If

            lRs.MoveNext
        Next lnI
        
       lnContador = lnI
        
        For lnI = 0 To nTotal - 1
        
        Me.Caption = "LOGISTICA: DEPRECIACION DE ACTIVOS FIJOS" + " *** GENERANDO ASNTO.CONT *** " + IIf(lnI Mod 4 = 0, "|", IIf(lnI Mod 4 = 1, "/", IIf(lnI Mod 4 = 2, "-", IIf(lnI Mod 4 = 3, "\", "|"))))
        
        If lnI = 0 Then
            lRs.MoveFirst
        End If

                If Round(lRs.Fields(18), 2) <> 0 And CDbl(lRs.Fields(19)) <> 0 Then
                    
                        lsCtaCont = oDep.GetOpeCtaCta(gnDepAF, Left(lRs.Fields(2), Len(lRs.Fields(2)) - 2) & "AG", "")

                        lsCtaCont = Replace(lsCtaCont, "AG", Right(lRs.Fields(2), 2))
                        
                        If lsCtaCont = "18190701" Then
                            lsCtaCont = "1819070102"
                        End If
                        oDep.InsertaMovCta lnMovNro, lnContador, lsCtaCont, Round(lRs.Fields(18), 2) * -1
                        
                        lnContador = lnContador + 1
                    'End If
                    
                End If

            lRs.MoveNext
        Next lnI
        

        If Me.chkSoloEstad.value = 0 Then
                lsMovNroR = oDep.GeneraMovNro(ldFechaDepre, Right(gsCodAge, 2), gsCodUser)
                oDep.InsertaMov lsMovNroR, gnDepAF, "Deprecación Mensual de Activo Fijo - Resumen " & Trim(Left(Me.cboTpo.Text, 30)) & " " & Format(ldFechaDepre, gsFormatoFechaView), 10
                
        '        lsMovNroDifAjusteR = oDep.GeneraMovNro(ldFechaDepre, Right(gsCodAge, 2), gsCodUser, lsMovNro)
        '        oDep.InsertaMov lsMovNroDifAjusteR, gnDepDifAjusteAF, "Deprecación Mensual de Ajuste de Activo Fijo - Resumen" & Format(ldFechaDepre, gsFormatoFechaView), 10
        '
                lnMovNroR = oDep.GetnMovNro(lsMovNroR)
        '        lnMovNroDifAjusteR = oDep.GetnMovNro(lsMovNroDifAjusteR)
                
                
                oDep.GeneraAsientoRes lnMovNroR, lnMovNro
                oDep.InsertaMovRef lnMovNroR, lnMovNro
        '        oDep.GeneraAsientoRes lnMovNroDifAjusteR, lnMovNroDifAjuste
        
                lcTextoCntEstad = 1
        End If

    oDep.CommitTrans
    
    Me.Caption = "LOGISTICA: DEPRECIACION DE ACTIVOS FIJOS"
    
    MsgBox "Depreciacion " + IIf(lcTextoCntEstad = 1, "Contable con asiento", "Contable solo estadistico") + " se generó OK.", vbInformation + vbOKOnly, "Atención"
    
     'oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNroR, 60, 80, Caption) & oImpresora.gPrnSaltoPagina & oAsiento.ImprimeAsientoContable(lsMovNroDifAjusteR, 60, 80, , Caption), Caption, True
    If lcTextoCntEstad = 1 Then
        oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNroR, 60, 80, Caption), Caption, True
    End If
    
    Call DepreContAF
    
'    gnDepAF = 581201
'    gnDepAjusteAF = 581202
End Sub

Private Sub cmdGrabarAFTrib_Click()
    Dim oDep As DMov
    Set oDep = New DMov
    
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim oconect As DConecta
    Set oconect = New DConecta
    
    Dim lnMovNro As Long
    Dim lsMovNro As String
    
    Dim lnMovNroDif As Long
    Dim lsMovNroDif As String
    
    Dim lnMovNroDifAjuste As Long
    Dim lsMovNroDifAjuste As String
    
    Dim lnMovNroR As Long
    Dim lsMovNroR As String
    Dim lnMovNroDifAjusteR As Long
    Dim lsMovNroDifAjusteR As String
    Dim lsTipo As String
    Dim lsFecha As String
    Dim I As Integer
    Dim lnI As Long
    Dim lnContador As Long
    Dim lsCtaCont As String
    Dim oPrevio As clsPrevio
    Dim oAsiento As NContImprimir
    Set oPrevio = New clsPrevio
    Set oAsiento = New NContImprimir
    Dim oContFunc As NContFunciones
    Dim oJDep As DLogBieSer
    Set oJDep = New DLogBieSer
    
    Dim rs As ADODB.Recordset
    Dim lcDebe As String, lcHaber As String, lcCodBSJoyAdj As String
    
    Dim ldFechaDepre As Date

    lsTipo = Trim(Right(cboTpo.Text, 2))

    If lsTipo <> "10" And nTotalTr = 0 Then
        MsgBox "No existen datos para procesar", vbCritical, "Atención"
        Exit Sub

    End If
    
    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If
    
    Set rs = New ADODB.Recordset
    
    '*** verifcar como busca ciando es tributario
    Set rs = GetBSAFAsiento(gnDepTributAF, lsFecha, lsTipo)
    
    If Not rs.EOF Then
        MsgBox "Depreciación Tributaria ya fue grabado.", vbCritical, "Aviso!"
        Exit Sub
    End If
    
    '-------------------------------------------------------------
    
    Set rs = GetBSAFTributCorrelativoAsiento(gnDepTributAF, lsFecha, lsTipo)
    
    If rs!cPeriodo <> "" And rs!cPeriodo <> lsFecha Then
        MsgBox "Para Depreciar el presente periodo debe estar Depreciado el anterior.", vbCritical, "Aviso!"
        Exit Sub
    ElseIf rs!cPeriodo = "" Then
        MsgBox "Este periodo es el primero en Depreciar por lo que en adelante se realizará correlativamente.", vbInformation + vbOKOnly, "Aviso!"
    End If
    
    '-------------------------------------------------------------
    
    If MsgBox("¿Desea Procesar? ", vbQuestion + vbYesNo, "Atención") = vbNo Then Exit Sub
    
    ldFechaDepre = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(Me.cmbMes.Text, 2) & "/" & Me.mskAnio.Text)))
    
    oDep.BeginTrans

        lsMovNro = oDep.GeneraMovNro(ldFechaDepre, Right(gsCodAge, 2), gsCodUser)
        oDep.InsertaMov lsMovNro, gnDepTributAF, "Depreciación Tribut. Mensual de Activo Fijo." & Format(ldFechaDepre, gsFormatoFechaView), 25
        
        lnMovNro = oDep.GetnMovNro(lsMovNro)
        
'        lRsTr.MoveFirst
        
        For lnI = 0 To nTotalTr - 1

            If lnI = 0 Then
                lRsTr.MoveFirst
            End If

                If Round(lRsTr.Fields(18), 2) > 0 And CDbl(lRsTr.Fields(19)) <> 0 Then
                        oDep.InsertaMovBSAF Me.mskAnio.Text, lRsTr.Fields(19), lnI, lRsTr.Fields(20), lRsTr.Fields(1), lnMovNro, lsTipo
                        oDep.InsertaMovDepreAFTrib lnMovNro, lnI, Round(lRsTr.Fields(18), 2)
                End If
            lRsTr.MoveNext
        Next lnI

    oDep.CommitTrans
    
    MsgBox "La Depreciación Tributaria se grabó satisfactoriamente.", vbExclamation + vbOKOnly, "Atención"
    
    Call DepreTribuAf

End Sub

Private Sub cmdImpAFTrib_Click()
   Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Set Me.Flex2.DataSource = lRsTr
    If Me.Flex2.Columns(0) = "" Then
        MsgBox "Debe Depreciar antes de imprimir.", vbInformation, "Aviso"
        Me.cmdDeprecia.SetFocus
        Exit Sub
    End If
    
    lsArchivoN = App.path & "\Spooler\" & Format(gdFecSis, "yyyymmdd" + Left(Replace(Time, ":", ""), 6)) & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       Call GeneraReporte(lRsTr)
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    MousePointer = 0

End Sub

Private Sub cmdimpLima_Click()
'    Dim rsE As ADODB.Recordset
'    Set rsE = New ADODB.Recordset
'    Dim lsArchivoN As String
'    Dim lbLibroOpen As Boolean
'
'    Dim oDep As DLogDeprecia
'    Dim ldFecha As Date
'    Set oDep = New DLogDeprecia
'
'    If Me.Flex.TextMatrix(1, 1) = "" Then
'        MsgBox "Debe Depreciar antes de imprimir.", vbInformation, "Aviso"
'        Me.cmdDeprecia.SetFocus
'        Exit Sub
'    End If
'
'    If Not IsNumeric(Me.mskAnio.Text) Then
'        MsgBox "Debe Ingresar un año Valido.", vbInformation, "Aviso"
'        Me.mskAnio.SetFocus
'        Exit Sub
'    ElseIf Me.cmbMes.Text = "" Then
'        MsgBox "Debe Ingresar un mes Valido.", vbInformation, "Aviso"
'        Me.cmbMes.SetFocus
'        Exit Sub
'    ElseIf Me.cboTpo.Text = "" Then
'        MsgBox "Debe Ingresar un tipo de depreciacion Valido.", vbInformation, "Aviso"
'        Me.cboTpo.SetFocus
'        Exit Sub
'    End If
'
'    ldFecha = CDate("01/" & Format(Trim(Right(Me.cmbMes.Text, 5)), "00") & "/" & Me.mskAnio.Text)
'
'    Me.Flex.rsFlex = oDep.GetAFDeprecia(ldFecha, Right(Me.cboTpo.Text, 3), mskAnio.Text, True)
'
'    lsArchivoN = App.path & "\Spooler\" & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & ".xls"
'    OleExcel.Class = "ExcelWorkSheet"
'    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
'    If lbLibroOpen Then
'       Set xlHoja1 = xlLibro.Worksheets(1)
'       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
'       Call GeneraReporteSaldoHistorico(Me.Flex.GetRsNew)
'       OleExcel.Class = "ExcelWorkSheet"
'       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
'       OleExcel.SourceDoc = lsArchivoN
'       OleExcel.Verb = 1
'       OleExcel.Action = 1
'       OleExcel.DoVerb -1
'    End If
'    MousePointer = 0
End Sub

Private Sub cmdImpLimaAjustado_Click()
'    Dim rsE As ADODB.Recordset
'    Set rsE = New ADODB.Recordset
'    Dim lsArchivoN As String
'    Dim lbLibroOpen As Boolean
'
'    Dim oDep As DLogDeprecia
'    Dim ldFecha As Date
'    Set oDep = New DLogDeprecia
'
'    If Me.Flex.TextMatrix(1, 1) = "" Then
'        MsgBox "Debe Depreciar antes de imprimir.", vbInformation, "Aviso"
'        Me.cmdDeprecia.SetFocus
'        Exit Sub
'    End If
'
'    If Not IsNumeric(Me.mskAnio.Text) Then
'        MsgBox "Debe Ingresar un año Valido.", vbInformation, "Aviso"
'        Me.mskAnio.SetFocus
'        Exit Sub
'    ElseIf Me.cmbMes.Text = "" Then
'        MsgBox "Debe Ingresar un mes Valido.", vbInformation, "Aviso"
'        Me.cmbMes.SetFocus
'        Exit Sub
'    ElseIf Me.cboTpo.Text = "" Then
'        MsgBox "Debe Ingresar un tipo de depreciacion Valido.", vbInformation, "Aviso"
'        Me.cboTpo.SetFocus
'        Exit Sub
'    End If
'
'    ldFecha = CDate("01/" & Format(Trim(Right(Me.cmbMes.Text, 5)), "00") & "/" & Me.mskAnio.Text)
'
'    Me.Flex.rsFlex = oDep.GetAFDeprecia(ldFecha, Right(Me.cboTpo.Text, 3), mskAnio.Text, True)
'
'    lsArchivoN = App.path & "\Spooler\" & Format(gdFecSis & " " & Time, "yyyymmddhhmmss") & ".xls"
'    OleExcel.Class = "ExcelWorkSheet"
'    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
'    If lbLibroOpen Then
'       Set xlHoja1 = xlLibro.Worksheets(1)
'       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
'       Call GeneraReporteSaldoAjustado(Flex.GetRsNew)
'       OleExcel.Class = "ExcelWorkSheet"
'       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
'       OleExcel.SourceDoc = lsArchivoN
'       OleExcel.Verb = 1
'       OleExcel.Action = 1
'       OleExcel.DoVerb -1
'    End If
'    MousePointer = 0
    
End Sub

Private Sub CmdImprimir_Click()
    Dim rsE As ADODB.Recordset
    Set rsE = New ADODB.Recordset
    Dim lsArchivoN As String
    Dim lbLibroOpen As Boolean
    Set Me.Flex.DataSource = lRs
    Dim I As Integer
    
    If Me.Flex.Columns(18) = "" Then
        MsgBox "Debe Depreciar antes de imprimir.", vbInformation, "Aviso"
        Me.cmdDeprecia.SetFocus
        Exit Sub
    End If
        
    lsArchivoN = App.path & "\Spooler\" & Format(gdFecSis, "yyyymmdd" + Left(Replace(Time, ":", ""), 6)) & ".xls"
    OleExcel.Class = "ExcelWorkSheet"
    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
    If lbLibroOpen Then
       Set xlHoja1 = xlLibro.Worksheets(1)
       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
       Call GeneraReporte(lRs)
       OleExcel.Class = "ExcelWorkSheet"
       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
       OleExcel.SourceDoc = lsArchivoN
       OleExcel.Verb = 1
       OleExcel.Action = 1
       OleExcel.DoVerb -1
    End If
    
    Me.Flex.Row = 1

End Sub

'Public Function ReporteAFCabeceraExcel(Optional xlHoja1 As Excel.Worksheet, Optional pbDepre As Boolean = True) As String
'    xlHoja1.PageSetup.LeftMargin = 1.5
'    xlHoja1.PageSetup.RightMargin = 0
'    xlHoja1.PageSetup.BottomMargin = 1
'    xlHoja1.PageSetup.TopMargin = 1
'    xlHoja1.PageSetup.Zoom = 70
'    xlHoja1.Cells.Font.Name = "Arial"
'    xlHoja1.Cells.Font.Size = 8
'
'    xlHoja1.Cells(2, 2) = "FORMATO 7.1 : ''REGISTRO DE ACTIVOS FIJOS - DETALLE DE LOS ACTIVOS FIJOS''"
'    xlHoja1.Cells(4, 2) = "PERIODO : " + txtPeriodo.Text
'    xlHoja1.Cells(5, 2) = "RUC : 20103845328"
'    xlHoja1.Cells(6, 2) = "APELLIDOS Y NOMBRES, DENOMINACION O RAZÓN SOCIAL : CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS SA."
'
'    xlHoja1.Range("B7:B8").MergeCells = True
'    xlHoja1.Range("C7:C8").MergeCells = True
'    xlHoja1.Range("D7:G7").MergeCells = True
'    xlHoja1.Range("H7:H8").MergeCells = True
'    xlHoja1.Range("I7:I8").MergeCells = True
'    xlHoja1.Range("J7:J8").MergeCells = True
'    xlHoja1.Range("K7:K8").MergeCells = True
'    xlHoja1.Range("L7:M7").MergeCells = True
'    xlHoja1.Range("N7:N8").MergeCells = True
'    xlHoja1.Range("O7:O8").MergeCells = True
'    xlHoja1.Range("P7:P8").MergeCells = True
'    xlHoja1.Range("Q7:Q8").MergeCells = True
'
'    xlHoja1.Cells(7, 2) = "CODIGO RELACIONADO CON EL ACTIVO FIJO"
'    xlHoja1.Cells(7, 3) = "CUENTA CONTA-BLE DEL ACTIVO FIJO"
'    xlHoja1.Cells(7, 4) = "DETALLE DE L ACTIVO FIJO"
'
'    xlHoja1.Cells(8, 4) = "DESCRIPCIÓN: MAQUINARIAS"
'    xlHoja1.Cells(8, 5) = "MARCA DEL ACTIVO FIJO"
'    xlHoja1.Cells(8, 6) = "MODELO DEL ACTIVO FIJO"
'    xlHoja1.Cells(8, 7) = "NUMERO DE SERIE Y/O PLACA DEL ACTIVO FIJO"
'    xlHoja1.Cells(7, 8) = "SALDO INICIAL"
'    xlHoja1.Cells(7, 9) = "ADQUISI-CIONES ADICIONES"
'    xlHoja1.Cells(7, 10) = "FECHA DE ADQUISI-CIÓN"
'    xlHoja1.Cells(7, 11) = "FECHA DE INICIO DEL USO DEL ACTIVO FIJO"
'    xlHoja1.Cells(7, 12) = "DEPRECIACION"
'    xlHoja1.Cells(8, 12) = "METODO APLICADO"
'    xlHoja1.Cells(8, 13) = "N° DE DOCUMEN-TO DE AUTORIZA-CIÓN"
'    xlHoja1.Cells(7, 14) = "PORCEN-TAJE DE DEPRE-CIACIÓN"
'    xlHoja1.Cells(7, 15) = "DEPRECIA-CIÓN ACUMULA-DA AL CIERRE DEL EJERCICIO ANTERIOR"
'    xlHoja1.Cells(7, 16) = "DEPRECIA-CIÓN DEL EJERCICIO"
'    xlHoja1.Cells(7, 17) = "DEPRECIA-CIÓN ACUMULADA HÍSTORICA"
'
'    xlHoja1.Rows("8:8").RowHeight = 77.25
'    xlHoja1.Columns("B:B").ColumnWidth = 16
'    xlHoja1.Columns("C:C").ColumnWidth = 14
'    xlHoja1.Columns("D:D").ColumnWidth = 38
'
'    xlHoja1.Range("B7:Q8").HorizontalAlignment = xlCenter
'    xlHoja1.Range("B7:Q8").VerticalAlignment = xlCenter
'    xlHoja1.Range("B2:Q8").Font.Bold = True
'    xlHoja1.Range("B7:Q8").WrapText = True
'    xlHoja1.Range("B7:Q8").Borders.LineStyle = 1 ''(xlDiagonalDown).LineStyle = xlNone
'
'
''    ApExcel.Cells.Select
''    ApExcel.Cells.EntireColumn.AutoFit
''    ApExcel.Columns("B:B").ColumnWidth = 6#
''    ApExcel.Range("B2").Select
'
'End Function

Private Sub cmdModificar_Click()
'    frmLogAFMant.Ini 1, False, True, False, Me.Flex.TextMatrix(Me.Flex.Row, 1), Me.Flex.TextMatrix(Me.Flex.Row, 3), CDate(Me.Flex.TextMatrix(Me.Flex.Row, 18)), Me.Flex.TextMatrix(Me.Flex.Row, 4), 0, 0, Me.Flex.TextMatrix(Me.Flex.Row, 7), Me.Flex.TextMatrix(Me.Flex.Row, 8), Me.Flex.TextMatrix(Me.Flex.Row, 19) & Me.Flex.TextMatrix(Me.Flex.Row, 20), Me.Flex.TextMatrix(Me.Flex.Row, 21), Me.Flex.TextMatrix(Me.Flex.Row, 17)
End Sub

Private Sub cmdNuevoAnio_Click()
'Dim oDep As DMov
'    Set oDep = New DMov
'
'    Dim oOpe As DOperacion
'    Set oOpe = New DOperacion
'
'    Dim lnMovNro As Long
'    Dim lsMovNro As String
'
'    Dim lnMovNroDif As Long
'    Dim lsMovNroDif As String
'
'    Dim lnMovNroDifAjuste As Long
'    Dim lsMovNroDifAjuste As String
'
'    Dim lnMovNroR As Long
'    Dim lsMovNroR As String
'    Dim lnMovNroDifAjusteR As Long
'    Dim lsMovNroDifAjusteR As String
'
'    Dim lnI As Long
'    Dim lnContador As Long
'    Dim lsCtaCont As String
'    Dim oPrevio As clsPrevio
'    Dim oAsiento As NContImprimir
'    Set oPrevio = New clsPrevio
'    Set oAsiento = New NContImprimir
'
'    Dim ldFechaDepre As Date
'
'
'    If MsgBox("Desea Procesar ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    ldFechaDepre = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(Me.cmbMes.Text, 2) & "/" & Me.mskAnio.Text)))
'    Me.mskAnio.Text = Trim(Str(CLng(Me.mskAnio.Text) + 1))
'    oDep.BeginTrans
'
'        oDep.IniActivoFijo Me.mskAnio.Text, Format(ldFechaDepre, "yyyy"), ldFechaDepre, Right(Me.cboTpo.Text, 3)
'
'        lsMovNro = oDep.GeneraMovNro(ldFechaDepre, Right(gsCodAge, 2), gsCodUser)
'        oDep.InsertaMov lsMovNro, gnDepAF, "Deprecación Mensual de Activo Fijo " & Format(ldFechaDepre, gsFormatoFechaView), 25
'        lsMovNroDif = oDep.GeneraMovNro(ldFechaDepre, Right(gsCodAge, 2), gsCodUser, lsMovNro)
'        oDep.InsertaMov lsMovNroDif, gnDepAjusteAF, "Deprecación Mensual de Ajuste de Activo Fijo " & Format(ldFechaDepre, gsFormatoFechaView), 25
'
'        lnMovNro = oDep.GetnMovNro(lsMovNro)
'        lnMovNroDif = oDep.GetnMovNro(lsMovNroDif)
'
'        For lnI = 1 To Me.Flex.Rows - 1
'            oDep.InsertaMovBSAF Val(Me.mskAnio.Text), Val(Me.Flex.TextMatrix(lnI, 12)), lnI, Me.Flex.TextMatrix(lnI, 1), Me.Flex.TextMatrix(lnI, 3), lnMovNro
'            lsCtaCont = oDep.GetOpeCtaCtaOtro(gnDepAF, Me.Flex.TextMatrix(lnI, 13), "", False)
'            If Me.Flex.TextMatrix(lnI, 10) <> "" Then
'                oDep.InsertaMovCta lnMovNro, lnI, lsCtaCont, Round(Me.Flex.TextMatrix(lnI, 10), 2)
'            Else
'                oDep.InsertaMovCta lnMovNro, lnI, lsCtaCont, Round(Me.Flex.TextMatrix(lnI, 9), 2)
'            End If
'        Next lnI
'
'        lsCtaCont = oOpe.EmiteOpeCta(gnDepAF, "D")
'
'        For lnI = 1 To Me.Flex.Rows - 1
'            oDep.InsertaMovBSAF Me.mskAnio.Text, Me.Flex.TextMatrix(lnI, 12), lnI, Me.Flex.TextMatrix(lnI, 1), Me.Flex.TextMatrix(lnI, 3), lnMovNroDif
'            If Me.Flex.TextMatrix(lnI, 10) <> "" Then
'                oDep.InsertaMovCta lnMovNroDif, lnI, lsCtaCont, Round(CCur(Me.Flex.TextMatrix(lnI, 10)), 2)
'            Else
'                oDep.InsertaMovCta lnMovNroDif, lnI, lsCtaCont, Round(CCur(Me.Flex.TextMatrix(lnI, 9)), 2)
'            End If
'        Next lnI
'    oDep.CommitTrans
End Sub

Private Sub cmdSalir_Click()
    Unload Me
   
End Sub

Private Sub cmdTipoActivo_Click()
    Call frmLogConfiguracionTipoActivo.Show(1)
End Sub

Private Sub Command1_Click()
'   Dim rsE As ADODB.Recordset
'    Set rsE = New ADODB.Recordset
'    Dim lsArchivoN As String
'    Dim lbLibroOpen As Boolean
'    Set Me.Flex2.DataSource = lRsTr
'    If Me.Flex2.Columns(0) = "" Then
'        MsgBox "Debe Depreciar antes de imprimir.", vbInformation, "Aviso"
'        Me.cmdDeprecia.SetFocus
'        Exit Sub
'    End If
'
'    lsArchivoN = App.path & "\Spooler\" & Format(gdFecSis, "yyyymmdd") & ".xls"
'    OleExcel.Class = "ExcelWorkSheet"
'    lbLibroOpen = ExcelBegin(lsArchivoN, xlAplicacion, xlLibro)
'    If lbLibroOpen Then
'       Set xlHoja1 = xlLibro.Worksheets(1)
'       ExcelAddHoja Format(gdFecSis, "yyyymmdd"), xlLibro, xlHoja1
'       Call GeneraReporte(lRsTr)
'       OleExcel.Class = "ExcelWorkSheet"
'       ExcelEnd lsArchivoN, xlAplicacion, xlLibro, xlHoja1
'       OleExcel.SourceDoc = lsArchivoN
'       OleExcel.Verb = 1
'       OleExcel.Action = 1
'       OleExcel.DoVerb -1
'    End If
'    MousePointer = 0
End Sub

Private Sub Command3_Click()

End Sub

Private Sub Command4_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub cmdVerAjustes_Click()
    Dim oPrevio As clsPrevio
    Dim oAsiento As NContImprimir
    Set oPrevio = New clsPrevio
    Set oAsiento = New NContImprimir

    Dim lsFecha As String
    Dim lsTipo As String
    Dim oDep As DMov
    Set oDep = New DMov
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim lsTextoImp As String
    Dim lin  As Integer

    lsTipo = Trim(Right(cboTpo.Text, 2))
    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If lsTipo <> "10" And nTotal = 0 Then
        MsgBox "No existen datos para procesar", vbCritical, "Atención"
        Exit Sub
    End If
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If
        
    Set rs = MuestraAjustes(gnDepAF, lsFecha, lsTipo)
    
    If (rs.EOF And rs.BOF) Then
        MsgBox "Este Activo Fijo no tiene Ajustes.", vbCritical, "Aviso!"
        Exit Sub
    Else
    
        lsTextoImp = lsTextoImp + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
        lsTextoImp = lsTextoImp + Space(5) + "CMAC MAYNAS S.A." + oImpresora.gPrnSaltoLinea
        lsTextoImp = lsTextoImp + Space(5) + Trim(CStr(Date)) + " " + CStr(Time()) + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
        lsTextoImp = lsTextoImp + Space(5) + "DETALLE DE BIENES CON AJUSTES DEL ACTIVO FIJO " + Left(Me.cboTpo.Text, Len(Me.cboTpo.Text) - 3) + oImpresora.gPrnSaltoLinea
        lsTextoImp = lsTextoImp + Space(5) + "============================================================================" + oImpresora.gPrnSaltoLinea
        lsTextoImp = lsTextoImp + oImpresora.gPrnSaltoLinea + oImpresora.gPrnSaltoLinea
        lsTextoImp = lsTextoImp + Space(5) + "Item  Serie           Cod.Bien.    Descripcion                                             Valor     Fec.Compra       Importe  Fecha" + oImpresora.gPrnSaltoLinea
        lsTextoImp = lsTextoImp + Space(5) + "-----------------------------------------------------------------------------------------------------------------------------------------" + oImpresora.gPrnSaltoLinea
        rs.MoveFirst
        lin = 0
        Do While Not rs.EOF
            lin = lin + 1
            lsTextoImp = lsTextoImp + Space(5) + Trim(CStr(lin)) + Space(4 - Len(Trim(CStr(lin)))) + Space(2) + Left(rs!cSerie, 20) + Space(2) + Left(rs!cBSCod, 15) + Space(2) + Trim(rs!cDescripcion) + Space(50 - Len(Trim(rs!cDescripcion))) + Space(2) + Space(12 - Len(Format(rs!nBSValor, "###,##0.00"))) & Format(rs!nBSValor, "###,##0.00") + Space(2) + Format(rs!dCompra, "dd/mm/yyyy") + Space(2) + Space(12 - Len(Format(rs!nMovImporte, "###,##0.00"))) & Format(rs!nMovImporte, "###,##0.00") + Space(2) + Format(rs!cFecha, "dd/mm/yyyy") + oImpresora.gPrnSaltoLinea
            '"pedro jose"+space(15-len("pedro jose"))
            If lin >= 48 Then
                lin = 0
                lsTextoImp = lsTextoImp + oImpresora.gPrnSaltoPagina
            End If
            rs.MoveNext
        Loop
        lsTextoImp = lsTextoImp + Space(5) + "-----------------------------------------------------------------------------------------------------------------------------------------" + oImpresora.gPrnSaltoLinea
        
        'oPrevio.Show oAsiento.ImprimeAsientoContable(rs!cMovNro, 60, 80, Caption), Caption, True
        oPrevio.Show lsTextoImp, Caption, True
        
        
    End If

End Sub

Private Sub cmdVerAsntoCnt_Click()
    Dim oPrevio As clsPrevio
    Dim oAsiento As NContImprimir
    Set oPrevio = New clsPrevio
    Set oAsiento = New NContImprimir

    Dim lsFecha As String
    Dim lsTipo As String
    Dim oDep As DMov
    Set oDep = New DMov
    
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
            
    lsTipo = Trim(Right(cboTpo.Text, 2))
    lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    
    If lsTipo <> "10" And nTotal = 0 Then
        MsgBox "No existen datos para procesar", vbCritical, "Atención"
        Exit Sub
    End If
    
    If Len(Trim(Right(cmbMes.Text, 2))) = 1 Then
        lsFecha = mskAnio.Text & "0" & Trim(Right(cmbMes.Text, 2))
    Else
        lsFecha = mskAnio.Text & Trim(Right(cmbMes.Text, 2))
    End If
        
    'Set rs = BuscaMovAsntoCntDepreAF(gnDepAF, lsFecha, lsTipo)
    Set rs = GetBSAFAsiento(gnDepAF, lsFecha, lsTipo)
    
    If (rs.EOF And rs.BOF) Then
        MsgBox "Este periodo no fue Depreciado.", vbCritical, "Aviso!"
        Exit Sub
    Else
        If rs!nEstadCnt = 0 Then
            MsgBox "La Depreciación no tiene asiento contable, ya que fue grabado como estadístico.", vbInformation + vbOKOnly, "Atención"
            Exit Sub
        Else
            oPrevio.Show oAsiento.ImprimeAsientoContable(rs!cMovNro, 60, 80, Caption), Caption, True
        End If
    End If

End Sub

'Private Sub Command1_Click()
'    Dim oDep As DMov
'    Set oDep = New DMov
'
'    Dim oOpe As DOperacion
'    Set oOpe = New DOperacion
'
'    Dim lnMovNro As Long
'    Dim lsMovNro As String
'
'    Dim lnMovNroDif As Long
'    Dim lsMovNroDif As String
'
'    Dim lnMovNroDifAjuste As Long
'    Dim lsMovNroDifAjuste As String
'
'    Dim lnMovNroR As Long
'    Dim lsMovNroR As String
'    Dim lnMovNroDifAjusteR As Long
'    Dim lsMovNroDifAjusteR As String
'
'    Dim lnI As Long
'    Dim lnContador As Long
'    Dim lsCtaCont As String
'    Dim oPrevio As clsPrevio
'    Dim oAsiento As NContImprimir
'    Set oPrevio = New clsPrevio
'    Set oAsiento = New NContImprimir
'
'    Dim ldFechaDepre As Date
'
'
'    If MsgBox("Desea Procesar ? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
'    Me.mskAnio.Text = "2002"
'    ldFechaDepre = DateAdd("d", -1, DateAdd("m", 1, CDate("01/" & Right(Me.cmbMes.Text, 2) & "/" & Me.mskAnio.Text)))
'    Me.mskAnio.Text = "2003"
'    oDep.BeginTrans
'        lsMovNro = oDep.GeneraMovNro(ldFechaDepre, Right(gsCodAge, 2), gsCodUser)
'        oDep.InsertaMov lsMovNro, gnDepAF, "Deprecación Mensual de Activo Fijo " & Format(ldFechaDepre, gsFormatoFechaView), 25
'        lsMovNroDif = oDep.GeneraMovNro(ldFechaDepre, Right(gsCodAge, 2), gsCodUser, lsMovNro)
'        oDep.InsertaMov lsMovNroDif, gnDepAjusteAF, "Deprecación Mensual de Ajuste de Activo Fijo " & Format(ldFechaDepre, gsFormatoFechaView), 25
'
'        lnMovNro = oDep.GetnMovNro(lsMovNro)
'        lnMovNroDif = oDep.GetnMovNro(lsMovNroDif)
'
'        For lnI = 1 To Me.Flex.Rows - 1
'            oDep.InsertaMovBSAF Me.mskAnio.Text, Me.Flex.TextMatrix(lnI, 12), lnI, Me.Flex.TextMatrix(lnI, 1), Me.Flex.TextMatrix(lnI, 3), lnMovNro
'            lsCtaCont = oDep.GetOpeCtaCtaOtro(gnDepAF, Me.Flex.TextMatrix(lnI, 13), "", False)
'            oDep.InsertaMovCta lnMovNro, lnI, lsCtaCont, Round(Me.Flex.TextMatrix(lnI, 10), 2)
'        Next lnI
'
'        lsCtaCont = oOpe.EmiteOpeCta(gnDepAF, "D")
'
'        For lnI = 1 To Me.Flex.Rows - 1
'            oDep.InsertaMovBSAF Me.mskAnio.Text, Me.Flex.TextMatrix(lnI, 12), lnI, Me.Flex.TextMatrix(lnI, 1), Me.Flex.TextMatrix(lnI, 3), lnMovNroDif
'            oDep.InsertaMovCta lnMovNroDif, lnI, lsCtaCont, Round(CCur(Me.Flex.TextMatrix(lnI, 10)), 2)
'        Next lnI
'    oDep.CommitTrans
'End Sub

Private Sub Form_Load()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral
    
    Set rs = oGen.GetConstante(1010)
    Me.cmbMes.Clear
    While Not rs.EOF
        cmbMes.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
    
    Set rs = oGen.GetConstante(5062, False)
    Me.cboTpo.Clear
    While Not rs.EOF
        cboTpo.AddItem rs.Fields(0) & Space(50) & rs.Fields(1)
        rs.MoveNext
    Wend
    
    Me.mskAnio.Text = Format(gdFecSis, "yyyy")
     
    'Me.Caption = lsCaption
    Me.Caption = "LOGISTICA: DEPRECIACION DE ACTIVOS FIJOS"
    
End Sub

Private Sub GeneraReporte(prRs As ADODB.Recordset)
    Dim I As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Variant 'Currency
    Dim lnSer As Currency
    
'    I = -1
'    prRs.MoveFirst
'    I = I + 1
'    For j = 0 To prRs.Fields.Count - 1
'        xlHoja1.Cells(I + 1, j + 1) = prRs.Fields(j).Name
'    Next j
    
    '*********************************************************
    '*** PEAC 20120326
    xlHoja1.PageSetup.LeftMargin = 1.5
    xlHoja1.PageSetup.RightMargin = 0
    xlHoja1.PageSetup.BottomMargin = 1
    xlHoja1.PageSetup.TopMargin = 1
    xlHoja1.PageSetup.Zoom = 70
    xlHoja1.Cells.Font.Name = "Arial"
    xlHoja1.Cells.Font.Size = 8

    '*** PEAC 20120326
    xlHoja1.Cells(2, 1) = "FORMATO 7.1 : ''REGISTRO DE ACTIVOS FIJOS - DETALLE DE LOS ACTIVOS FIJOS''"
    xlHoja1.Cells(4, 1) = "PERIODO : " + Me.mskAnio + " - " + Left(Me.cmbMes, Len(Me.cmbMes) - 3)
    xlHoja1.Cells(4, 7) = IIf(Me.Tab.Tab = 0, "DEPRECIACION CONTABLE", "DEPRECIACION TRIBUTARIA")
    xlHoja1.Cells(5, 1) = "RUC : 20103845328"
    xlHoja1.Cells(6, 1) = "APELLIDOS Y NOMBRES, DENOMINACION O RAZÓN SOCIAL : CAJA MUNICIPAL DE AHORRO Y CREDITO DE MAYNAS SA."

    xlHoja1.Range("A7:A8").MergeCells = True
    xlHoja1.Range("B7:B8").MergeCells = True
    xlHoja1.Range("C7:C8").MergeCells = True
    xlHoja1.Range("D7:G7").MergeCells = True
    xlHoja1.Range("H7:H8").MergeCells = True
    xlHoja1.Range("I7:I8").MergeCells = True
    xlHoja1.Range("J7:J8").MergeCells = True
    xlHoja1.Range("K7:K8").MergeCells = True
    xlHoja1.Range("L7:M7").MergeCells = True
    xlHoja1.Range("N7:N8").MergeCells = True
    xlHoja1.Range("O7:O8").MergeCells = True
    xlHoja1.Range("P7:P8").MergeCells = True
    xlHoja1.Range("Q7:Q8").MergeCells = True
    xlHoja1.Range("R7:R8").MergeCells = True
    xlHoja1.Range("S7:S8").MergeCells = True

    xlHoja1.Cells(7, 1) = "AGENCIA"
    xlHoja1.Cells(7, 2) = "CODIGO RELACIONADO CON EL ACTIVO FIJO"
    xlHoja1.Cells(7, 3) = "CUENTA CONTABLE DEL ACTIVO FIJO"
    xlHoja1.Cells(7, 4) = "DETALLE DE L ACTIVO FIJO"

    xlHoja1.Cells(8, 4) = "DESCRIPCIÓN:" + Trim(Left(Me.cboTpo.Text, Len(Me.cboTpo.Text) - 3))
    xlHoja1.Cells(8, 5) = "MARCA DEL ACTIVO FIJO"
    xlHoja1.Cells(8, 6) = "MODELO DEL ACTIVO FIJO"
    xlHoja1.Cells(8, 7) = "NUMERO DE SERIE Y/O PLACA DEL ACTIVO FIJO"
    xlHoja1.Cells(7, 8) = "SALDO INICIAL"
    xlHoja1.Cells(7, 9) = "ADQUISICIONES ADICIONES"
    xlHoja1.Cells(7, 10) = "FECHA DE ADQUISICIÓN"
    xlHoja1.Cells(7, 11) = "FECHA DE INICIO DEL USO DEL ACTIVO FIJO"
    xlHoja1.Cells(7, 12) = "DEPRECIACION"
    xlHoja1.Cells(8, 12) = "METODO APLICADO"
    xlHoja1.Cells(8, 13) = "N° DE DOCUMENTO DE AUTORIZACIÓN"
    xlHoja1.Cells(7, 14) = "PORCENTAJE DE DEPRECIACIÓN"
    xlHoja1.Cells(7, 15) = "DEPRECIACIÓN ACUMULADA AL CIERRE DEL EJERCICIO ANTERIOR"
    xlHoja1.Cells(7, 16) = "DEPRECIACIÓN DEL EJERCICIO"
    xlHoja1.Cells(7, 17) = "DEPRECIACIÓN ACUMULADA HÍSTORICA"
    xlHoja1.Cells(7, 18) = "NETO"
    xlHoja1.Cells(7, 19) = "DEPRECIACION MES"

'    ApExcel.Cells.Select
'    ApExcel.Cells.EntireColumn.AutoFit
'    ApExcel.Columns("B:B").ColumnWidth = 6#
'    ApExcel.Range("B2").Select

    '*********************************************************
    
    prRs.MoveFirst
    
    I = 7
    While Not prRs.EOF
        I = I + 1
        
        Me.Caption = "LOGISTICA: DEPRECIACION DE ACTIVOS FIJOS" + " *** GENERANDO IMPRESION *** " + IIf(I Mod 4 = 0, "|", IIf(I Mod 4 = 1, "/", IIf(I Mod 4 = 2, "-", IIf(I Mod 4 = 3, "\", "|"))))
        
        For j = 0 To prRs.Fields.Count - 3
            xlHoja1.Cells(I + 1, j + 1) = IIf(j = 14, "'" + Trim(CStr(IIf(IsNull(prRs.Fields(j)), 0, prRs.Fields(j)))), prRs.Fields(j))
        Next j
        prRs.MoveNext
    Wend
    
    Me.Caption = "LOGISTICA: DEPRECIACION DE ACTIVOS FIJOS"
        
    I = I + 1
    
    xlHoja1.Range("A1:A" & Trim(Str(I))).Font.Bold = True
'    xlHoja1.Range("B1:B" & Trim(Str(I))).Font.Bold = True
'    xlHoja1.Range("1:1").Font.Bold = True
    
    xlHoja1.Range("H1:H" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("I1:I" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("O1:O" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("P1:P" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("Q1:Q" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("R1:R" & Trim(Str(I))).NumberFormat = "#,##0.00"
    xlHoja1.Range("S1:S" & Trim(Str(I))).NumberFormat = "#,##0.00"

'    xlHoja1.Range("V1:V" & Trim(Str(I))).NumberFormat = "dd/mm/yyyy"
'    xlHoja1.Range("Z1:Z" & Trim(Str(I))).NumberFormat = "dd/mm/yyyy"
'    xlHoja1.Range("AB1:AB" & Trim(Str(I))).NumberFormat = "dd/mm/yyyy"
        
'    xlHoja1.Range("1:1").Font.Bold = True
    
    xlHoja1.Columns.AutoFit

    With xlHoja1.PageSetup
        .LeftHeader = ""
        .CenterHeader = "&""Arial,Negrita""&14Reporte de Activo Fijo Mes :  " & Trim(Left(Me.cmbMes, 15)) & " - " & Me.mskAnio.Text & " - " & Trim(Left(Me.cboTpo.Text, 15))
        .RightHeader = "&P"
        .LeftFooter = ""
        .CenterFooter = ""
        .RightFooter = ""
        .PrintHeadings = False
        .PrintGridlines = False
        .PrintComments = xlPrintNoComments
        .CenterHorizontally = False
        .CenterVertically = False
        .Orientation = xlLandscape
        .Draft = False
        .FirstPageNumber = xlAutomatic
        .Order = xlDownThenOver
        .BlackAndWhite = False
        .Zoom = 70
    End With

    With xlHoja1.Range("A1:S" & Trim(Str(I + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:S" & Trim(Str(I + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:S" & Trim(Str(I + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:S" & Trim(Str(I + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:S" & Trim(Str(I + 1))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A1:S" & Trim(Str(I + 1))).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
    xlHoja1.Rows("8:8").RowHeight = 77.25
    xlHoja1.Columns("A:A").ColumnWidth = 8
    xlHoja1.Columns("B:B").ColumnWidth = 16
    xlHoja1.Columns("C:C").ColumnWidth = 12
    xlHoja1.Columns("D:D").ColumnWidth = 38
    xlHoja1.Columns("E:E").ColumnWidth = 13
    xlHoja1.Columns("F:F").ColumnWidth = 13
    xlHoja1.Columns("G:G").ColumnWidth = 20
    xlHoja1.Columns("I:I").ColumnWidth = 14
    xlHoja1.Columns("J:J").ColumnWidth = 12
    xlHoja1.Columns("K:K").ColumnWidth = 12
    xlHoja1.Columns("L:L").ColumnWidth = 11
    xlHoja1.Columns("M:M").ColumnWidth = 12.5
    xlHoja1.Columns("N:N").ColumnWidth = 12
    xlHoja1.Columns("O:O").ColumnWidth = 13
    xlHoja1.Columns("P:P").ColumnWidth = 13
    xlHoja1.Columns("Q:Q").ColumnWidth = 13
    xlHoja1.Columns("R:R").ColumnWidth = 13
    xlHoja1.Columns("S:S").ColumnWidth = 13

    xlHoja1.Range("B7:S8").HorizontalAlignment = xlCenter
    xlHoja1.Range("B7:S8").VerticalAlignment = xlCenter
    xlHoja1.Range("B2:S8").Font.Bold = True
    xlHoja1.Range("B7:S8").WrapText = True
    xlHoja1.Range("B7:S8").Borders.LineStyle = 1 ''(xlDiagonalDown).LineStyle = xlNone
    
    
End Sub

Private Sub GeneraReporteSaldoHistorico(prRs As ADODB.Recordset)
    Dim I As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    I = 1
    
    With xlHoja1.Range("A1:N1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A1:N1").Merge
    xlHoja1.Range("A1:N1").FormulaR1C1 = " REPORTE DE SALDOS HISTORICOS "
    
    With xlHoja1.Range("A2:G2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A2:G2").Merge
    xlHoja1.Range("A2:G2").FormulaR1C1 = " COSTO DEL ACTIVO FIJO "
    
    With xlHoja1.Range("H2:N2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("H2:N2").Merge
    xlHoja1.Range("H2:N2").FormulaR1C1 = " DEPRECIACION DE ACTIVOS FIJOS "
    
    prRs.MoveFirst
    While Not prRs.EOF
        I = I + 1
        
        If I = 2 Then
            xlHoja1.Cells(I + 1, 1) = "COD CTIVO"
            xlHoja1.Cells(I + 1, 2) = "COD. PATRIM."
            xlHoja1.Cells(I + 1, 3) = "F. ADQUIS"
            xlHoja1.Cells(I + 1, 4) = "SALDO AÑO ANT."
            xlHoja1.Cells(I + 1, 5) = "COMPRAS AÑO"
            xlHoja1.Cells(I + 1, 6) = "RETIROS AÑO"
            xlHoja1.Cells(I + 1, 7) = "SALDO ACTUAL"
            xlHoja1.Cells(I + 1, 8) = "DEP ACUM EJER ANT"
            xlHoja1.Cells(I + 1, 9) = "DEP AL MES ANT"
            xlHoja1.Cells(I + 1, 10) = "DEP DEP MES"
            xlHoja1.Cells(I + 1, 11) = "TOT DEP DEL EJER"
            xlHoja1.Cells(I + 1, 12) = "DEP ACUM DE RET"
            xlHoja1.Cells(I + 1, 13) = "DEP ACUM TOTAL"
            xlHoja1.Cells(I + 1, 14) = "VALOR EN LIBROS"
            
            I = I + 1
            
            xlHoja1.Cells(I + 1, 1) = prRs!codigo & "-" & prRs!Serie
            xlHoja1.Cells(I + 1, 2) = Format(I - 2, "00000")
            xlHoja1.Cells(I + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(I + 1, 4) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            xlHoja1.Cells(I + 1, 5) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Then
                xlHoja1.Cells(I + 1, 6) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 6) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 6) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 7) = Format(xlHoja1.Cells(I + 1, 4) + xlHoja1.Cells(I + 1, 5) - xlHoja1.Cells(I + 1, 6), "#,##0.00")
            
            xlHoja1.Cells(I + 1, 8) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Dep_H_Ejer_Ant, 0), "#,##0.00")
            xlHoja1.Cells(I + 1, 9) = IIf(prRs!Baja = "False", Format(prRs!Dep_H_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 10) = Format(prRs!Dep_H_Mes, "#,##0.00")
            xlHoja1.Cells(I + 1, 11) = Format(xlHoja1.Cells(I + 1, 9) + xlHoja1.Cells(I + 1, 10), "#,##0.00")
            If prRs!Baja = "False" Then
                xlHoja1.Cells(I + 1, 12) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 12) = Format(prRs!Dep_H_Mes_Ant, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 12) = Format(0, "#,##0.00")
                End If
            End If
            xlHoja1.Cells(I + 1, 13) = Format(xlHoja1.Cells(I + 1, 8) + xlHoja1.Cells(I + 1, 11) - xlHoja1.Cells(I + 1, 12), "#,##0.00")
            xlHoja1.Cells(I + 1, 14) = Format(xlHoja1.Cells(I + 1, 13) - xlHoja1.Cells(I + 1, 7), "#,##0.00")
        Else
            xlHoja1.Cells(I + 1, 1) = prRs!codigo & "-" & prRs!Serie
            xlHoja1.Cells(I + 1, 2) = Format(I - 2, "00000")
            xlHoja1.Cells(I + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(I + 1, 4) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            xlHoja1.Cells(I + 1, 5) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio.Text, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Then
                xlHoja1.Cells(I + 1, 6) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 6) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 6) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 7) = Format(xlHoja1.Cells(I + 1, 4) + xlHoja1.Cells(I + 1, 5) - xlHoja1.Cells(I + 1, 6), "#,##0.00")
            
            xlHoja1.Cells(I + 1, 8) = Format(IIf(Year(prRs!F_Compra) <> Me.mskAnio.Text, prRs!Dep_H_Ejer_Ant, 0), "#,##0.00")
            xlHoja1.Cells(I + 1, 9) = IIf(prRs!Baja = "False", Format(prRs!Dep_H_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 10) = Format(prRs!Dep_H_Mes, "#,##0.00")
            xlHoja1.Cells(I + 1, 11) = Format(xlHoja1.Cells(I + 1, 9) + xlHoja1.Cells(I + 1, 10), "#,##0.00")
            If prRs!Baja = "False" Then
                xlHoja1.Cells(I + 1, 12) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 12) = Format(prRs!Dep_H_Mes_Ant, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 12) = Format(0, "#,##0.00")
                End If
            End If
            xlHoja1.Cells(I + 1, 13) = Format(xlHoja1.Cells(I + 1, 8) + xlHoja1.Cells(I + 1, 11) - xlHoja1.Cells(I + 1, 12), "#,##0.00")
            xlHoja1.Cells(I + 1, 14) = Format(xlHoja1.Cells(I + 1, 13) - xlHoja1.Cells(I + 1, 7), "#,##0.00")
        End If
        prRs.MoveNext
    Wend
        
    xlHoja1.Cells.Select
    xlHoja1.Columns.AutoFit
    
    xlHoja1.Range("A1:N3").Select
    xlHoja1.Range("A1:N3").Font.Bold = True

    With xlHoja1.Range("A2:G2").Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With
    With xlHoja1.Range("H2:N2").Interior
        .ColorIndex = 35
        .Pattern = xlSolid
    End With
    
    xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Select
    xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlDiagonalUp).LineStyle = xlNone
    
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:N" & Trim(Str(I + 1))).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    
End Sub

Private Sub GeneraReporteSaldoAjustado(prRs As ADODB.Recordset)
    Dim I As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Currency
    
    I = 1
    
    With xlHoja1.Range("A1:S1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A1:T1").Merge
    xlHoja1.Range("A1:T1").FormulaR1C1 = " REPORTE DE SALDOS AJUSTADOS "
    
    With xlHoja1.Range("A2:L2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("A2:L2").Merge
    xlHoja1.Range("A2:L2").FormulaR1C1 = " COSTO DEL AJUSTADO DE ACTIVO FIJO "
    
    With xlHoja1.Range("M2:T2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = False
    End With
    xlHoja1.Range("M2:T2").Merge
    xlHoja1.Range("M2:T2").FormulaR1C1 = " DEPRECIACION AJUSTADA DE ACTIVOS FIJOS "
    
    prRs.MoveFirst
    While Not prRs.EOF
        I = I + 1
        
        If I = 2 Then
            xlHoja1.Cells(I + 1, 1) = "COD CTIVO"
            xlHoja1.Cells(I + 1, 2) = "COD. PATRIM."
            xlHoja1.Cells(I + 1, 3) = "F. ADQUIS"
            xlHoja1.Cells(I + 1, 4) = "SALDO ACT HIST"
            xlHoja1.Cells(I + 1, 5) = "SALDO AJUS AÑO ANT"
            xlHoja1.Cells(I + 1, 6) = "COM DEL AÑO"
            xlHoja1.Cells(I + 1, 7) = "RET DEL AÑO"
            xlHoja1.Cells(I + 1, 8) = "FAC DE AJUS"
            xlHoja1.Cells(I + 1, 9) = "REEXP VAL AJUS ANT"
            xlHoja1.Cells(I + 1, 10) = "COM AJUS"
            xlHoja1.Cells(I + 1, 11) = "RET AJUS"
            xlHoja1.Cells(I + 1, 12) = "VAL ACT AJUS"
            xlHoja1.Cells(I + 1, 13) = "DEP ACUM AÑO ANT"
            xlHoja1.Cells(I + 1, 14) = "REEXP DEP AJUS AÑO ANT"
            xlHoja1.Cells(I + 1, 15) = "DEP AJUS EJER MES ANT"
            xlHoja1.Cells(I + 1, 16) = "DEP AJUST DEL MES"
            xlHoja1.Cells(I + 1, 17) = "TOT DEP DEL EJER AJSU"
            xlHoja1.Cells(I + 1, 18) = "DEP AJUS ACUM RET"
            xlHoja1.Cells(I + 1, 19) = "DEP AJUS ACUM TOTAL"
            xlHoja1.Cells(I + 1, 20) = "VALOR EN LIBROS AJUS"
            
            I = I + 1

            xlHoja1.Cells(I + 1, 1) = prRs!codigo & "-" & prRs!Serie
            xlHoja1.Cells(I + 1, 2) = Format(I - 2, "00000")
            xlHoja1.Cells(I + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(I + 1, 4) = Format(prRs!Valor, "#,##0.00")
            xlHoja1.Cells(I + 1, 5) = IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, Format(prRs!Valor_Ajustado, "#,##0.00"))
            xlHoja1.Cells(I + 1, 6) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(I + 1, 7) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 7) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 7) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 8) = Format(prRs!F_Ajuste, "#,##0.000")
            xlHoja1.Cells(I + 1, 9) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(I + 1, 10) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0) * prRs!F_Ajuste, "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(I + 1, 11) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 11) = Format(prRs!Valor * prRs!F_Ajuste, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 11) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 12) = Format(xlHoja1.Cells(I + 1, 9) + xlHoja1.Cells(I + 1, 10) - xlHoja1.Cells(I + 1, 11), "#,##0.00")
            
            xlHoja1.Cells(I + 1, 13) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado), "#,##0.00")
            xlHoja1.Cells(I + 1, 14) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(I + 1, 15) = IIf(prRs!Baja = "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 16) = Format(prRs!Dep_A_Mes, "#,##0.00")
            xlHoja1.Cells(I + 1, 17) = Format(xlHoja1.Cells(I + 1, 15) + xlHoja1.Cells(I + 1, 16), "#,##0.00")
            xlHoja1.Cells(I + 1, 18) = IIf(prRs!Baja <> "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 19) = Format(xlHoja1.Cells(I + 1, 14) + xlHoja1.Cells(I + 1, 17) - xlHoja1.Cells(I + 1, 18), "#,##0.00")
            xlHoja1.Cells(I + 1, 20) = Format(xlHoja1.Cells(I + 1, 12) - xlHoja1.Cells(I + 1, 19), "#,##0.00")
        Else
            xlHoja1.Cells(I + 1, 1) = prRs!codigo & "-" & prRs!Serie
            xlHoja1.Cells(I + 1, 2) = Format(I - 2, "00000")
            xlHoja1.Cells(I + 1, 3) = "'" & Format(prRs!F_Compra, gsFormatoFechaView)
            xlHoja1.Cells(I + 1, 4) = Format(prRs!Valor, "#,##0.00")
            xlHoja1.Cells(I + 1, 5) = IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, Format(prRs!Valor_Ajustado, "#,##0.00"))
            xlHoja1.Cells(I + 1, 6) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0), "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(I + 1, 7) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 7) = Format(prRs!Valor, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 7) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 8) = Format(prRs!F_Ajuste, "#,##0.000")
            xlHoja1.Cells(I + 1, 9) = Format(IIf(Year(prRs!F_Compra) = Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(I + 1, 10) = Format(IIf(Year(prRs!F_Compra) = Year(gdFecSis), prRs!Valor, 0) * prRs!F_Ajuste, "#,##0.00")
            
            If prRs!Baja = "False" Or Not IsNumeric(prRs!F_Baja) Then
                xlHoja1.Cells(I + 1, 11) = Format(0, "#,##0.00")
            Else
                If Year(prRs!F_Baja) = Me.mskAnio Then
                    xlHoja1.Cells(I + 1, 11) = Format(prRs!Valor * prRs!F_Ajuste, "#,##0.00")
                Else
                    xlHoja1.Cells(I + 1, 11) = Format(0, "#,##0.00")
                End If
            End If
            
            xlHoja1.Cells(I + 1, 12) = Format(xlHoja1.Cells(I + 1, 9) + xlHoja1.Cells(I + 1, 10) - xlHoja1.Cells(I + 1, 11), "#,##0.00")
            
            xlHoja1.Cells(I + 1, 13) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado), "#,##0.00")
            xlHoja1.Cells(I + 1, 14) = Format(IIf(Year(prRs!F_Compra) >= Me.mskAnio, 0, prRs!Valor_Ajustado) * prRs!F_Ajuste, "#,##0.00")
            xlHoja1.Cells(I + 1, 15) = IIf(prRs!Baja = "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 16) = Format(prRs!Dep_A_Mes, "#,##0.00")
            xlHoja1.Cells(I + 1, 17) = Format(xlHoja1.Cells(I + 1, 15) + xlHoja1.Cells(I + 1, 16), "#,##0.00")
            xlHoja1.Cells(I + 1, 18) = IIf(prRs!Baja <> "False", Format(prRs!Dep_A_Mes_Ant, "#,##0.00"), Format(0, "#,##0.00"))
            xlHoja1.Cells(I + 1, 19) = Format(xlHoja1.Cells(I + 1, 14) + xlHoja1.Cells(I + 1, 17) - xlHoja1.Cells(I + 1, 18), "#,##0.00")
            xlHoja1.Cells(I + 1, 20) = Format(-xlHoja1.Cells(I + 1, 12) + xlHoja1.Cells(I + 1, 19), "#,##0.00")

        End If
        prRs.MoveNext
    Wend
        
    xlHoja1.Select
    xlHoja1.Columns.AutoFit
    
    xlHoja1.Range("A1:T3").Select
    xlHoja1.Range("A1:T3").Font.Bold = True

    With xlHoja1.Range("A2:L2").Interior
        .ColorIndex = 36
        .Pattern = xlSolid
    End With
    With xlHoja1.Range("M2:T2").Interior
        .ColorIndex = 35
        .Pattern = xlSolid
    End With
    
    xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Select
    xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlDiagonalDown).LineStyle = xlNone
    xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlDiagonalUp).LineStyle = xlNone
    
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
    With xlHoja1.Range("A2:T" & Trim(Str(I + 1))).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
        .ColorIndex = xlAutomatic
    End With
End Sub


Private Sub GetResumen()
'    Dim lnI As Long
'    Dim lnAnioAnt As Integer
'    Dim lnSumaVH As Currency
'    Dim lnSumaVA As Currency
'    Dim lnSumaVHM As Currency
'    Dim lnSumaVAM As Currency
'    Dim lnSumaVHMM As Currency
'    Dim lnSumaVAMM As Currency
'    Dim lnSumaVAjuste As Currency
'    Dim lnSumaVHMA As Currency
'    Dim lnSumaVAMA As Currency
'
'    flexRes.Clear
'    flexRes.Rows = 2
'    flexRes.FormaCabecera
'
'
'    lnSumaVH = 0
'    lnSumaVA = 0
'    lnSumaVHM = 0
'    lnSumaVAM = 0
'    lnSumaVHMM = 0
'    lnSumaVAMM = 0
'    lnSumaVAjuste = 0
'    lnSumaVHMA = 0
'    lnSumaVAMA = 0
'
'
'    For lnI = 1 To Me.Flex.Rows - 1
'        If lnAnioAnt <> 0 And lnAnioAnt <> Year(Flex.TextMatrix(lnI, 23)) And Year(Flex.TextMatrix(lnI, 23)) > 1998 Then
'            flexRes.AdicionaFila
'            flexRes.TextMatrix(flexRes.Rows - 1, 1) = lnAnioAnt
'            flexRes.TextMatrix(flexRes.Rows - 1, 2) = Format(lnSumaVH, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 3) = Format(lnSumaVA, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 4) = Format(lnSumaVHM, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 5) = Format(lnSumaVAM, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 6) = Format(lnSumaVHMM, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 7) = Format(lnSumaVAMM, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 8) = Format(lnSumaVAjuste, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 9) = Format(lnSumaVHMA, "#,##0.00")
'            flexRes.TextMatrix(flexRes.Rows - 1, 10) = Format(lnSumaVAMA, "#,##0.00")
'
'            lnSumaVH = 0
'            lnSumaVA = 0
'            lnSumaVHM = 0
'            lnSumaVAM = 0
'            lnSumaVHMM = 0
'            lnSumaVAMM = 0
'            lnSumaVAjuste = 0
'            lnSumaVHMA = 0
'            lnSumaVAMA = 0
'        End If
'
'        lnSumaVH = lnSumaVH + Flex.TextMatrix(lnI, 4)
'        If IsNumeric(Flex.TextMatrix(lnI, 6)) Then lnSumaVA = lnSumaVA + Flex.TextMatrix(lnI, 6)
'        lnSumaVHM = lnSumaVHM + Flex.TextMatrix(lnI, 9)
'        If IsNumeric(Flex.TextMatrix(lnI, 10)) Then lnSumaVAM = lnSumaVAM + Flex.TextMatrix(lnI, 10)
'        lnSumaVHMM = lnSumaVHMM + Flex.TextMatrix(lnI, 14)
'        If IsNumeric(Flex.TextMatrix(lnI, 15)) Then lnSumaVAMM = lnSumaVAMM + Flex.TextMatrix(lnI, 15)
'        If IsNumeric(Flex.TextMatrix(lnI, 16)) Then lnSumaVAjuste = lnSumaVAjuste + Flex.TextMatrix(lnI, 16)
'        lnSumaVHMA = lnSumaVHMA + Flex.TextMatrix(lnI, 26)
'        lnSumaVAMA = lnSumaVAMA + Flex.TextMatrix(lnI, 27)
'
'        lnAnioAnt = Year(Flex.TextMatrix(lnI, 23))
'    Next lnI
'
'    flexRes.AdicionaFila
'    flexRes.TextMatrix(flexRes.Rows - 1, 1) = lnAnioAnt
'    flexRes.TextMatrix(flexRes.Rows - 1, 2) = Format(lnSumaVH, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 3) = Format(lnSumaVA, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 4) = Format(lnSumaVHM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 5) = Format(lnSumaVAM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 6) = Format(lnSumaVHMM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 7) = Format(lnSumaVAMM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 8) = Format(lnSumaVAjuste, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 9) = Format(lnSumaVHMA, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 10) = Format(lnSumaVAMA, "#,##0.00")
'
'    lnSumaVH = 0
'    lnSumaVA = 0
'    lnSumaVHM = 0
'    lnSumaVAM = 0
'    lnSumaVHMM = 0
'    lnSumaVAMM = 0
'    lnSumaVAjuste = 0
'    lnSumaVHMA = 0
'    lnSumaVAMA = 0
'
'    For lnI = 1 To Me.flexRes.Rows - 1
'        lnSumaVH = lnSumaVH + flexRes.TextMatrix(lnI, 2)
'        lnSumaVA = lnSumaVA + flexRes.TextMatrix(lnI, 3)
'        lnSumaVHM = lnSumaVHM + flexRes.TextMatrix(lnI, 4)
'        lnSumaVAM = lnSumaVAM + flexRes.TextMatrix(lnI, 5)
'        lnSumaVHMM = lnSumaVHMM + flexRes.TextMatrix(lnI, 6)
'        lnSumaVAMM = lnSumaVAMM + flexRes.TextMatrix(lnI, 7)
'        lnSumaVAjuste = lnSumaVAjuste + flexRes.TextMatrix(lnI, 8)
'        lnSumaVHMA = lnSumaVHMA + flexRes.TextMatrix(lnI, 9)
'        lnSumaVAMA = lnSumaVAMA + flexRes.TextMatrix(lnI, 10)
'    Next lnI
'
'    flexRes.AdicionaFila
'    flexRes.TextMatrix(flexRes.Rows - 1, 1) = "TOTAL"
'    flexRes.TextMatrix(flexRes.Rows - 1, 2) = Format(lnSumaVH, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 3) = Format(lnSumaVA, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 4) = Format(lnSumaVHM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 5) = Format(lnSumaVAM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 6) = Format(lnSumaVHMM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 7) = Format(lnSumaVAMM, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 8) = Format(lnSumaVAjuste, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 9) = Format(lnSumaVHMA, "#,##0.00")
'    flexRes.TextMatrix(flexRes.Rows - 1, 10) = Format(lnSumaVAMA, "#,##0.00")
End Sub

Private Sub GeneraReporteJoyasAdjud(prRs As ADODB.Recordset)
    Dim I As Integer
    Dim K As Integer
    Dim j As Integer
    Dim nFila As Integer
    Dim nIni  As Integer
    Dim lNegativo As Boolean
    Dim sConec As String
    Dim lsSuma As String
    Dim sTipoGara As String
    Dim sTipoCred As String
    Dim lnAcum As Variant 'Currency
    Dim lnSer As Currency
    
    I = -1
    prRs.MoveFirst
    I = I + 1
    For j = 0 To prRs.Fields.Count - 1
        xlHoja1.Cells(I + 1, j + 1) = prRs.Fields(j).Name
    Next j
    
    While Not prRs.EOF
        I = I + 1
        For j = 0 To prRs.Fields.Count - 1
            xlHoja1.Cells(I + 1, j + 1) = prRs.Fields(j)
        Next j
        prRs.MoveNext
    Wend
    
'    i = i + 1
'    xlHoja1.Range("A1:A" & Trim(Str(i))).Font.Bold = True
'    xlHoja1.Range("B1:B" & Trim(Str(i))).Font.Bold = True
'    xlHoja1.Range("1:1").Font.Bold = True
'
'    xlHoja1.Range("D1:D" & Trim(Str(i))).NumberFormat = "#,##0.00"
'    xlHoja1.Range("F1:F" & Trim(Str(i))).NumberFormat = "#,##0.00"
'    xlHoja1.Range("I1:I" & Trim(Str(i))).NumberFormat = "#,##0.00"
'    xlHoja1.Range("J1:J" & Trim(Str(i))).NumberFormat = "#,##0.00"
'    xlHoja1.Range("N1:N" & Trim(Str(i))).NumberFormat = "#,##0.00"
'    xlHoja1.Range("O1:O" & Trim(Str(i))).NumberFormat = "#,##0.00"
'    xlHoja1.Range("P1:P" & Trim(Str(i))).NumberFormat = "#,##0.00"
'    xlHoja1.Range("Z1:Z" & Trim(Str(i))).NumberFormat = "#,##0.00"
'    xlHoja1.Range("AA1:AA" & Trim(Str(i))).NumberFormat = "#,##0.00"
'    xlHoja1.Range("AB1:AB" & Trim(Str(i))).NumberFormat = "#,##0.00"
'    xlHoja1.Range("AC1:AC" & Trim(Str(i))).NumberFormat = "#,##0.00"
'
'
'    xlHoja1.Range("S1:S" & Trim(Str(i))).NumberFormat = "dd/mm/yyyy"
'    xlHoja1.Range("W1:W" & Trim(Str(i))).NumberFormat = "dd/mm/yyyy"
'    xlHoja1.Range("Y1:Y" & Trim(Str(i))).NumberFormat = "dd/mm/yyyy"
'
'    xlHoja1.Range("D" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
'    xlHoja1.Range("F" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
'    xlHoja1.Range("I" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
'    xlHoja1.Range("J" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
'    xlHoja1.Range("N" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
'    xlHoja1.Range("O" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
'    xlHoja1.Range("P" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
'    xlHoja1.Range("Z" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
'    xlHoja1.Range("AA" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
'    xlHoja1.Range("AB" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
'    xlHoja1.Range("AC" & Trim(Str(i + 1))).FormulaR1C1 = "=SUM(R[-" & Trim(Str(i)) & "]C:R[-1]C)"
'
'
'    xlHoja1.Range("1:1").Font.Bold = True
'    xlHoja1.Columns.AutoFit
'
'    With xlHoja1.PageSetup
'        .LeftHeader = ""
'        .CenterHeader = "&""Arial,Negrita""&14Reporte de Activo Fijo Mes :  " & Trim(Left(Me.cmbMes, 15)) & " - " & Me.mskAnio.Text & " - " & Trim(Left(Me.cboTpo.Text, 15))
'        .RightHeader = "&P"
'        .LeftFooter = ""
'        .CenterFooter = ""
'        .RightFooter = ""
'        .PrintHeadings = False
'        .PrintGridlines = False
'        .PrintComments = xlPrintNoComments
'        .CenterHorizontally = False
'        .CenterVertically = False
'        .Orientation = xlLandscape
'        .Draft = False
'        .FirstPageNumber = xlAutomatic
'        .Order = xlDownThenOver
'        .BlackAndWhite = False
'        .Zoom = 70
'    End With
'
'    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlEdgeLeft)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlEdgeTop)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlEdgeBottom)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlEdgeRight)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlInsideVertical)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
'    With xlHoja1.Range("A1:AC" & Trim(Str(i + 1))).Borders(xlInsideHorizontal)
'        .LineStyle = xlContinuous
'        .Weight = xlThin
'        .ColorIndex = xlAutomatic
'    End With
End Sub

Private Sub Tab_Click(PreviousTab As Integer)
    If PreviousTab = 0 Then
        Call DepreTribuAf
    Else
        Call DepreContAF
    End If
End Sub
