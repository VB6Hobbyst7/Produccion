VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSegDesgravamenRetPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Caption"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6735
   Icon            =   "frmSegDesgravamenRetPago.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6765
   ScaleWidth      =   6735
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   4200
      TabIndex        =   19
      Top             =   6360
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5400
      TabIndex        =   20
      Top             =   6360
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   10821
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   520
      TabCaption(0)   =   "Origen"
      TabPicture(0)   =   "frmSegDesgravamenRetPago.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraOperacion"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FraCtaIF"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FraPeriodo"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "fraTrama"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "dlgArchivo"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      Begin MSComDlg.CommonDialog dlgArchivo 
         Left            =   5760
         Top             =   480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame fraTrama 
         Caption         =   "Trama Seguro"
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
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   6255
         Begin VB.CommandButton cmdCargar 
            Caption         =   "Cargar"
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
            Left            =   4920
            TabIndex        =   7
            Top             =   240
            Width           =   1215
         End
         Begin Sicmact.TxtBuscar TxtBuscaTrama 
            Height          =   345
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   609
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
            TipoBusqueda    =   6
            sTitulo         =   ""
            EnabledText     =   0   'False
         End
      End
      Begin VB.Frame FraPeriodo 
         Caption         =   " Periodo "
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
         Height          =   735
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   3855
         Begin VB.CommandButton cmdSeleccionar 
            Caption         =   "Seleccionar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   2520
            TabIndex        =   4
            Top             =   240
            Width           =   1170
         End
         Begin VB.TextBox txtAño 
            Alignment       =   2  'Center
            Height          =   315
            Left            =   1560
            MaxLength       =   4
            TabIndex        =   3
            Top             =   240
            Width           =   735
         End
         Begin VB.ComboBox cboMes 
            Height          =   315
            ItemData        =   "frmSegDesgravamenRetPago.frx":0326
            Left            =   120
            List            =   "frmSegDesgravamenRetPago.frx":0330
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame FraCtaIF 
         Caption         =   " Cuenta Institución Financiera "
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
         Height          =   1695
         Left            =   120
         TabIndex        =   8
         Top             =   2040
         Width           =   6255
         Begin Sicmact.TxtBuscar txtBuscaEntidad 
            Height          =   345
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1995
            _ExtentX        =   3519
            _ExtentY        =   609
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
            sTitulo         =   ""
            EnabledText     =   0   'False
         End
         Begin VB.Label lblDescCtaBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   120
            TabIndex        =   11
            Top             =   750
            Width           =   6015
         End
         Begin VB.Label lblDescBanco 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   2205
            TabIndex        =   10
            Top             =   375
            Width           =   3930
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo Pago :"
            Height          =   285
            Left            =   120
            TabIndex        =   12
            Top             =   1200
            Width           =   825
         End
         Begin VB.Label lblTipoPago 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1080
            TabIndex        =   13
            Top             =   1170
            Width           =   1410
         End
      End
      Begin VB.Frame fraOperacion 
         Caption         =   " Datos de la operación "
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
         Height          =   2175
         Left            =   120
         TabIndex        =   14
         Top             =   3840
         Width           =   6255
         Begin VB.TextBox txtMovDesc 
            Height          =   720
            Left            =   720
            MaxLength       =   500
            MultiLine       =   -1  'True
            TabIndex        =   18
            Top             =   840
            Width           =   5370
         End
         Begin MSMask.MaskEdBox txtFecha 
            Height          =   315
            Left            =   720
            TabIndex        =   16
            Top             =   345
            Width           =   1155
            _ExtentX        =   2037
            _ExtentY        =   556
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblMonto 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "0.00"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   1350
            TabIndex        =   22
            Top             =   1700
            Width           =   1695
         End
         Begin VB.Label lblMontoText 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Monto "
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   240
            TabIndex        =   21
            Top             =   1750
            Width           =   570
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H00E0E0E0&
            BackStyle       =   1  'Opaque
            Height          =   345
            Left            =   120
            Top             =   1680
            Width           =   2955
         End
         Begin VB.Label Label7 
            BackStyle       =   0  'Transparent
            Caption         =   "Fecha :"
            Height          =   285
            Left            =   120
            TabIndex        =   15
            Top             =   360
            Width           =   705
         End
         Begin VB.Label Label1 
            Caption         =   "Glosa :"
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   840
            Width           =   705
         End
      End
   End
End
Attribute VB_Name = "frmSegDesgravamenRetPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************
'Nombre : frmSegDesgravamenRetPago
'Descripcion:Formulario para el Registrode de Pago de Seguros
'Creacion: PASIERS1362014
'*****************************
Option Explicit
Dim fsopecod As String
Dim oDSeg As DSeguros
Dim oOpe As DOperacion
Dim oNContFunc As NContFunciones
Dim lsCtaContBanco As String
Dim nTipoDoc As TpoDoc
Dim fsMatrizSeguro As Currency
Dim rs As ADODB.Recordset
Dim FTSegAge() As TSegAgencia
Dim FTRetPagSegDet() As TRetPagoSegDet
Dim objPista As COMManejador.Pista 'ARLO20170217

Public Sub Inicio(ByVal psOpeCod As String)
       fsopecod = psOpeCod
       CargaComboConstante 1010, cboMes
       txtBuscaEntidad.psRaiz = "Cuentas de Bancos"
       Set oOpe = New DOperacion
       txtBuscaEntidad.rs = oOpe.GetOpeObj(fsopecod, "2")
       Me.Show 1
End Sub
Private Sub CboMes_Click()
    If Not cboMes.ListIndex = -1 Then
        txtAño.SetFocus
    End If
End Sub
Private Sub cmdAceptar_Click()
    Dim oDocPago As clsDocPago
    Dim oCtasIF As NCajaCtaIF
    Dim oDoc As DDocumento
    Dim oNSeg As NSeguros
    Dim oImp As NContImprimir
    
    Dim lsDocVoucher As String
    Dim lsPersCodIf As String
    Dim lsCtaEntidadOrig As String
    Dim lsPersNombre As String
    Dim lsSubCuentaIF As String
    Dim lsEntidadOrig As String
    Dim lsDocNro As String
    Dim lsFecha As String
    Dim lsDocNroTmp As String
    Dim lsDocVoucherTmp As String
    Dim lsPlanillaNro As String
    Dim lsMovNro As String
    Dim lsTpoIf As String
    Dim lsCtaBanco As String
    Dim lsImpre As String
    
    If Not ValidaDatos Then Exit Sub
    
    Set oNContFunc = New NContFunciones
    Set oDocPago = New clsDocPago
    Set oCtasIF = New NCajaCtaIF
    Set oDoc = New DDocumento
    
    lsTpoIf = Mid(txtBuscaEntidad.Text, 1, 2)
    lsCtaEntidadOrig = Trim(lblDescCtaBanco.Caption)
    lsPersCodIf = Mid(txtBuscaEntidad, 4, 13)
    lsCtaBanco = Mid(txtBuscaEntidad.Text, 18, Len(txtBuscaEntidad.Text))
    lsSubCuentaIF = oCtasIF.SubCuentaIF(Mid(txtBuscaEntidad.Text, 4, 13))
    lsEntidadOrig = oCtasIF.NombreIF(lsPersCodIf)
    
    If nTipoDoc = TpoDocCheque Then
        lsDocVoucher = oNContFunc.GeneraDocNro(TpoDocVoucherEgreso, Mid(fsopecod, 3, 1))
        oDocPago.InicioCheque lsDocNro, True, Mid(txtBuscaEntidad.Text, 4, 13), fsopecod, lsPersNombre, Me.Caption, Trim(txtMovDesc.Text), CCur(lblMonto.Caption), gdFecSis, gsNomCmacRUC, lsSubCuentaIF, lsEntidadOrig, lsCtaEntidadOrig, lsDocVoucher, True, gsCodAge ', , , lsTpoIf, Mid(txtBuscaEntidad.Text, , 4, 13), lsCtaBanco
        If oDocPago.vbOk Then
            lsFecha = oDocPago.vdFechaDoc
            lsDocNroTmp = oDocPago.vsNroDoc
            lsDocVoucherTmp = oDocPago.vsNroVoucher
        Else
            Exit Sub
        End If
    Else
            Do While True
                lsPlanillaNro = InputBox("Ingrese el Nro. de Planilla", "Planilla de Pagos", lsPlanillaNro)
                If lsPlanillaNro = "" Then Exit Sub
                If oDoc.GetValidaDocProv("", CLng(nTipoDoc), lsPlanillaNro) Then
                    MsgBox "Nro. de carta ya ha sido ingresada, verifique..!", vbInformation, "Aviso"
                Else
                    lsDocNroTmp = lsPlanillaNro
                    lsDocVoucherTmp = ""
                    gnMgIzq = 17
                    gnMgDer = 0
                    gnMgSup = 12
                    Exit Do
                End If
            Loop
    End If
    Set oDoc = Nothing
    
    If MsgBox("¿Esta seguro de realizar la operación?", vbYesNo + vbQuestion, "Aviso") = vbNo Then Exit Sub
    Set oNSeg = New NSeguros
    Call oNSeg.GrabarRetiroPagoSeguroDesgravIncen(gdFecSis, gsCodAge, gsCodUser, fsopecod, txtAño.Text + IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0", "") & Trim(Right(cboMes.Text, 2)), _
                                                                            nTipoDoc, lsDocNroTmp, txtFecha.Text, Trim(Replace(Replace(txtMovDesc.Text, Chr(10), ""), Chr(13), "")), FTSegAge, FTRetPagSegDet, CDbl(lblMonto.Caption), lsCtaContBanco, Mid(txtBuscaEntidad.Text, 4, 13), _
                                                                            lsTpoIf, lsCtaBanco, lsMovNro)
    
    Set oNSeg = Nothing
    Set oImp = New NContImprimir
    lsImpre = oImp.ImprimeAsientoContable(lsMovNro, gnLinPage, gnColPage, Replace(Replace(Me.Caption, "ME", ""), "MN", ""))
    EnviaPrevio lsImpre & oImpresora.gPrnSaltoPagina, Replace(Replace(Me.Caption, "ME", ""), "MN", ""), gnLinPage, False
    SetInicioControl
    Des_HabilitaControles 0
    cboMes.SetFocus
    Set oImp = Nothing
        
        'ARLO20170217
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " Se Grabo la Operación "
        Set objPista = Nothing
        '****
    Unload Me
End Sub
Private Function ValidaDatos() As Boolean
ValidaDatos = False
    If lsCtaContBanco = "" Then
        MsgBox "No seleccionó la cuenta de la Institución Financiera", vbInformation, "Aviso"
        txtBuscaEntidad.SetFocus
        Exit Function
    End If
    If nTipoDoc = 0 Then
        MsgBox "No se especificó el Tipo de Pago", vbInformation, "Aviso"
        txtBuscaEntidad.SetFocus
        Exit Function
    End If
    If Trim(Replace(Replace(txtMovDesc.Text, Chr(10), ""), Chr(13), "")) = "" Then
        MsgBox "Debe ingresar la glosa", vbInformation, "Aviso"
        txtMovDesc.SetFocus
        Exit Function
    End If
    If Not (CCur(lblMonto.Caption) > 0) Then
        MsgBox "No se ha generado ningún Monto. Verifique que el archivo de las tramas sea el correcto.", vbInformation, "Aviso"
        Exit Function
    End If
ValidaDatos = True
End Function
Private Sub cmdCancelar_Click()
    SetInicioControl
    Des_HabilitaControles 0
    cboMes.SetFocus
End Sub
Private Sub cmdCargar_Click()
    Dim psArchivoALeer As String
    Dim psAchivoAGrabar As String
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja As Excel.Worksheet
    Dim bExiste As Boolean
    Dim bEncontrado As Boolean
    Dim fs As New Scripting.FileSystemObject
    Dim oDSeguro As DSeguros
    Set oDSeguro = New DSeguros
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim pdFecha As Date

    If TxtBuscaTrama.Text = "" Then
        MsgBox "Asegurese de Cargar Correctamente la Trama.", vbInformation, "Aviso."
        TxtBuscaTrama.SetFocus
        Exit Sub
    End If
    On Error GoTo ErrCarga
    psArchivoALeer = TxtBuscaTrama.Text
    bExiste = fs.FileExists(psArchivoALeer)
    
    If bExiste = False Then
        MsgBox "Ha Ocurrido un Error con la Carga del Archivo. Verifique que el archivo existe en la ruta.", vbExclamation, "Aviso!!!"
        Exit Sub
    End If
    
    Set xlAplicacion = New Excel.Application
    Set xlLibro = xlAplicacion.Workbooks.Open(psArchivoALeer)
    bEncontrado = False
    For Each xlHoja In xlLibro.Worksheets
        If UCase(xlHoja.Name) = UCase("Trama" & Trim(Left(cboMes.Text, 50)) & IIf(Mid(fsopecod, 3, 1) = "1", txtAño.Text & "MN", txtAño.Text & "ME")) Then
            bEncontrado = True
            xlHoja.Activate
            Exit For
        End If
    Next
    If bEncontrado = False Then
        ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, False
        MsgBox "El archivo seleccionado no parece ser el correcto.", vbInformation, "Aviso!!!"
        Exit Sub
    End If
    Dim lnColCta, lnColAge As Integer
    Dim lnFilaMonto, lnColMonto, I, j As Integer
    Dim lnColFechaVig As Date
    Dim nMonto As Currency
    Dim bexisteAge As Boolean
    Dim bexisteCta As Boolean
    Dim FTRetPagSegDetAlt() As TRetPagoSegDet
    
    bexisteAge = False
    bexisteCta = False
    Select Case fsopecod
        Case OpeCGOtrosOpeRetPagSeguroDesgravamenMN, OpeCGOtrosOpeRetPagSeguroDesgravamenME
            
            lnFilaMonto = 2
            lnColMonto = 20
            lnColFechaVig = 19
            lnColCta = 2
            
            Do While xlHoja.Cells(lnFilaMonto, lnColMonto) <> ""
                If lnFilaMonto = 2 Then
                    ReDim Preserve FTRetPagSegDetAlt(1)
                    FTRetPagSegDetAlt(1).sCtaCod = xlHoja.Cells(lnFilaMonto, lnColCta)
                    FTRetPagSegDetAlt(1).sAgeCod = Format(Mid(CStr(xlHoja.Cells(lnFilaMonto, lnColCta)), 4, 2), "00")
                    FTRetPagSegDetAlt(1).nMonto = CCur(xlHoja.Cells(lnFilaMonto, lnColMonto))
                    FTRetPagSegDetAlt(1).dFechaVig = CCur(xlHoja.Cells(lnFilaMonto, lnColFechaVig))
                Else
                    ReDim Preserve FTRetPagSegDetAlt(UBound(FTRetPagSegDetAlt) + 1)
                    FTRetPagSegDetAlt(UBound(FTRetPagSegDetAlt)).sCtaCod = xlHoja.Cells(lnFilaMonto, lnColCta)
                    FTRetPagSegDetAlt(UBound(FTRetPagSegDetAlt)).sAgeCod = Format(Mid(CStr(xlHoja.Cells(lnFilaMonto, lnColCta)), 4, 2), "00")
                    FTRetPagSegDetAlt(UBound(FTRetPagSegDetAlt)).nMonto = CCur(xlHoja.Cells(lnFilaMonto, lnColMonto))
                    FTRetPagSegDetAlt(UBound(FTRetPagSegDetAlt)).dFechaVig = CCur(xlHoja.Cells(lnFilaMonto, lnColFechaVig))
                End If
                 lnFilaMonto = lnFilaMonto + 1
            Loop
                ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, False
            For j = 1 To UBound(FTRetPagSegDetAlt)
                If j = 1 Then
                    ReDim Preserve FTSegAge(1)
                    FTSegAge(1).sAgencia = FTRetPagSegDetAlt(j).sAgeCod
                    FTSegAge(1).nMonto = oDSeg.GetMontoSegDesgravamenxRetPagSeguro(FTRetPagSegDetAlt(j).sCtaCod, FTRetPagSegDetAlt(j).nMonto, FTRetPagSegDetAlt(j).dFechaVig)
                    ReDim Preserve FTRetPagSegDet(1)
                    FTRetPagSegDet(1).sCtaCod = FTRetPagSegDetAlt(j).sCtaCod
                    FTRetPagSegDet(1).sAgeCod = FTRetPagSegDetAlt(j).sAgeCod
                    FTRetPagSegDet(1).nMonto = oDSeg.GetMontoSegDesgravamenxRetPagSeguro(FTRetPagSegDetAlt(j).sCtaCod, FTRetPagSegDetAlt(j).nMonto, FTRetPagSegDetAlt(j).dFechaVig)
                Else
                    For I = 1 To UBound(FTSegAge)
                        If FTRetPagSegDetAlt(j).sAgeCod = FTSegAge(I).sAgencia Then
                            FTSegAge(I).nMonto = FTSegAge(I).nMonto + oDSeg.GetMontoSegDesgravamenxRetPagSeguro(FTRetPagSegDetAlt(j).sCtaCod, FTRetPagSegDetAlt(j).nMonto, FTRetPagSegDetAlt(j).dFechaVig)
                            bexisteAge = True
                        End If
                    Next
                    If Not bexisteAge Then
                        ReDim Preserve FTSegAge(UBound(FTSegAge) + 1)
                        FTSegAge(UBound(FTSegAge)).sAgencia = FTRetPagSegDetAlt(j).sAgeCod
                        FTSegAge(UBound(FTSegAge)).nMonto = oDSeg.GetMontoSegDesgravamenxRetPagSeguro(FTRetPagSegDetAlt(j).sCtaCod, FTRetPagSegDetAlt(j).nMonto, FTRetPagSegDetAlt(j).dFechaVig)
                    End If
                    
                    For I = 1 To UBound(FTRetPagSegDet)
                        If FTRetPagSegDetAlt(j).sCtaCod = FTRetPagSegDet(I).sCtaCod Then
                            FTRetPagSegDet(I).nMonto = FTRetPagSegDet(I).nMonto + oDSeg.GetMontoSegDesgravamenxRetPagSeguro(FTRetPagSegDetAlt(j).sCtaCod, FTRetPagSegDetAlt(j).nMonto, FTRetPagSegDetAlt(j).dFechaVig)
                            bexisteCta = True
                        End If
                    Next
                    If Not bexisteCta Then
                        ReDim Preserve FTRetPagSegDet(UBound(FTRetPagSegDet) + 1)
                        FTRetPagSegDet(UBound(FTRetPagSegDet)).sCtaCod = FTRetPagSegDetAlt(j).sCtaCod
                        FTRetPagSegDet(UBound(FTRetPagSegDet)).sAgeCod = FTRetPagSegDetAlt(j).sAgeCod
                        FTRetPagSegDet(UBound(FTRetPagSegDet)).nMonto = oDSeg.GetMontoSegDesgravamenxRetPagSeguro(FTRetPagSegDetAlt(j).sCtaCod, FTRetPagSegDetAlt(j).nMonto, FTRetPagSegDetAlt(j).dFechaVig)
                    End If
                End If
                nMonto = nMonto + oDSeg.GetMontoSegDesgravamenxRetPagSeguro(FTRetPagSegDetAlt(j).sCtaCod, FTRetPagSegDetAlt(j).nMonto, FTRetPagSegDetAlt(j).dFechaVig)
                bexisteAge = False
                bexisteCta = False
            Next j
            pdFecha = DateAdd("D", -1, DateAdd("M", 1, CDate("01/" & IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0", "") & Trim(Right(cboMes.Text, 2)) + "/" & txtAño.Text)))
            If UBound(FTRetPagSegDet) > 0 Then
                For I = 1 To UBound(FTRetPagSegDet)
                    Set rs = oDSeguro.GetCuotaCredxRetPagSeguro(FTRetPagSegDet(I).sCtaCod, pdFecha)
                    If Not rs.EOF Then
                        FTRetPagSegDet(I).nCuota = rs!nCuota
                        FTRetPagSegDet(I).nNroCalen = rs!nNroCalen
                    End If
                Next
            End If
            
        Case OpeCGOtrosOpeRetPagSeguroIncendioMN, OpeCGOtrosOpeRetPagSeguroIncendioME
            lnColAge = 4
            lnFilaMonto = 3
            lnColMonto = 9
            lnColCta = 2
             
            Do While xlHoja.Cells(lnFilaMonto, lnColMonto) <> ""
                If lnFilaMonto = 3 Then
                    ReDim Preserve FTRetPagSegDetAlt(1)
                    FTRetPagSegDetAlt(1).sCtaCod = xlHoja.Cells(lnFilaMonto, lnColCta)
                    FTRetPagSegDetAlt(1).sAgeCod = Format(Mid(CStr(xlHoja.Cells(lnFilaMonto, lnColCta)), 4, 2), "00")
                    FTRetPagSegDetAlt(1).nMonto = CCur(xlHoja.Cells(lnFilaMonto, lnColMonto))
                Else
                    ReDim Preserve FTRetPagSegDetAlt(UBound(FTRetPagSegDetAlt) + 1)
                    FTRetPagSegDetAlt(UBound(FTRetPagSegDetAlt)).sCtaCod = xlHoja.Cells(lnFilaMonto, lnColCta)
                    FTRetPagSegDetAlt(UBound(FTRetPagSegDetAlt)).sAgeCod = Format(Mid(CStr(xlHoja.Cells(lnFilaMonto, lnColCta)), 4, 2), "00")
                    FTRetPagSegDetAlt(UBound(FTRetPagSegDetAlt)).nMonto = CCur(xlHoja.Cells(lnFilaMonto, lnColMonto))
                End If
                 lnFilaMonto = lnFilaMonto + 1
            Loop
                ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, False
            For j = 1 To UBound(FTRetPagSegDetAlt)
                If j = 1 Then
                    ReDim Preserve FTSegAge(1)
                    FTSegAge(1).sAgencia = FTRetPagSegDetAlt(j).sAgeCod
                    FTSegAge(1).nMonto = FTRetPagSegDetAlt(j).nMonto
                    
                    ReDim Preserve FTRetPagSegDet(1)
                    FTRetPagSegDet(1).sCtaCod = FTRetPagSegDetAlt(j).sCtaCod
                    FTRetPagSegDet(1).sAgeCod = FTRetPagSegDetAlt(j).sAgeCod
                    FTRetPagSegDet(1).nMonto = FTRetPagSegDetAlt(j).nMonto
                Else
                    For I = 1 To UBound(FTSegAge)
                        If FTRetPagSegDetAlt(j).sAgeCod = FTSegAge(I).sAgencia Then
                            FTSegAge(I).nMonto = FTSegAge(I).nMonto + FTRetPagSegDetAlt(j).nMonto
                            bexisteAge = True
                        End If
                    Next
                    If Not bexisteAge Then
                        ReDim Preserve FTSegAge(UBound(FTSegAge) + 1)
                        FTSegAge(UBound(FTSegAge)).sAgencia = FTRetPagSegDetAlt(j).sAgeCod
                        FTSegAge(UBound(FTSegAge)).nMonto = FTRetPagSegDetAlt(j).nMonto
                    End If
                    
                    For I = 1 To UBound(FTRetPagSegDet)
                        If FTRetPagSegDetAlt(j).sCtaCod = FTRetPagSegDet(I).sCtaCod Then
                            FTRetPagSegDet(I).nMonto = FTRetPagSegDet(I).nMonto + FTRetPagSegDetAlt(j).nMonto
                            bexisteCta = True
                        End If
                    Next
                    If Not bexisteCta Then
                        ReDim Preserve FTRetPagSegDet(UBound(FTRetPagSegDet) + 1)
                        FTRetPagSegDet(UBound(FTRetPagSegDet)).sCtaCod = FTRetPagSegDetAlt(j).sCtaCod
                        FTRetPagSegDet(UBound(FTRetPagSegDet)).sAgeCod = FTRetPagSegDetAlt(j).sAgeCod
                        FTRetPagSegDet(UBound(FTRetPagSegDet)).nMonto = FTRetPagSegDetAlt(j).nMonto
                    End If
                End If
                nMonto = nMonto + FTRetPagSegDetAlt(j).nMonto
                bexisteAge = False
                bexisteCta = False
            Next j
            pdFecha = DateAdd("D", -1, DateAdd("M", 1, CDate("01/" & IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0", "") & Trim(Right(cboMes.Text, 2)) + "/" & txtAño.Text)))
            If UBound(FTRetPagSegDet) > 0 Then
                For I = 1 To UBound(FTRetPagSegDet)
                    Set rs = oDSeguro.GetCuotaCredxRetPagSeguro(FTRetPagSegDet(I).sCtaCod, pdFecha)
                    If Not rs.EOF Then
                        FTRetPagSegDet(I).nCuota = rs!nCuota
                        FTRetPagSegDet(I).nNroCalen = rs!nNroCalen
                    End If
                Next
            End If
    End Select
    lblMonto.Caption = Format(nMonto, "#,#0.00")
    Des_HabilitaControles 2
    txtBuscaEntidad.SetFocus
    Exit Sub
ErrCarga:
    ExcelEnd psArchivoALeer, xlAplicacion, xlLibro, xlHoja, True
    MsgBox TextErr("Ha Ocurrido un Error con la Carga del Archivo. Verifique."), vbInformation, "Aviso"
End Sub
Private Sub cmdSeleccionar_Click()
    If Not ValidaSeleccionar Then Exit Sub
        Des_HabilitaControles 1
        TxtBuscaTrama.SetFocus
End Sub
Private Sub Des_HabilitaControles(ByVal pnHabilita As Integer)
Dim pbHabil, pbNoHabil As Boolean
pbHabil = True
pbNoHabil = False
Select Case pnHabilita
        Case 0
                    FraPeriodo.Enabled = pbHabil
                    fraTrama.Enabled = pbNoHabil
                    FraCtaIF.Enabled = pbNoHabil
                    fraOperacion.Enabled = pbNoHabil
                    cmdAceptar.Enabled = pbNoHabil
        Case 1
                    FraPeriodo.Enabled = pbNoHabil
                    fraTrama.Enabled = pbHabil
                    FraCtaIF.Enabled = pbNoHabil
                    fraOperacion.Enabled = pbNoHabil
                    cmdAceptar.Enabled = pbNoHabil
        Case 2
                    FraPeriodo.Enabled = pbNoHabil
                    fraTrama.Enabled = pbNoHabil
                    FraCtaIF.Enabled = pbHabil
                    fraOperacion.Enabled = pbHabil
                    cmdAceptar.Enabled = pbHabil
End Select

End Sub
Private Sub Form_Load()
    Select Case fsopecod
        Case OpeCGOtrosOpeRetPagSeguroDesgravamenMN, OpeCGOtrosOpeRetPagSeguroDesgravamenME
            Me.Caption = "Retiro por Pago Seguro Desgravamen " & IIf(Mid(fsopecod, 3, 1) = "1", "MN", "ME")
        Case OpeCGOtrosOpeRetPagSeguroIncendioMN, OpeCGOtrosOpeRetPagSeguroIncendioME
            Me.Caption = "Retiro por Pago Seguro Contra Incendio " & IIf(Mid(fsopecod, 3, 1) = "1", "MN", "ME")
    End Select
    SetInicioControl
    Des_HabilitaControles 0
End Sub
Private Function ValidaSeleccionar() As Boolean
    ValidaSeleccionar = True
     If Trim(cboMes.Text) = "" Then
        MsgBox "Seleccione correctamente el mes", vbInformation, "Aviso"
        ValidaSeleccionar = False
        cboMes.SetFocus
        Exit Function
    End If
    If Trim(txtAño.Text) = "" Or (Val(txtAño.Text) < 1900 Or Val(txtAño.Text) > 9972) Then
        MsgBox "Ingrese Correctamente el año.", vbInformation, "Aviso"
        ValidaSeleccionar = False
        txtAño.SetFocus
        Exit Function
    End If
    If CInt(Trim(Right(cboMes.Text, 2))) >= CInt(Mid(gdFecSis, 4, 2)) Or Val(txtAño.Text) > Right(gdFecSis, 4) Then
        MsgBox "El periodo seleccionado debe ser anterior al mes actual", vbInformation, "Aviso"
        ValidaSeleccionar = False
        cboMes.SetFocus
        Exit Function
    End If
    Set oDSeg = New DSeguros
    If Not oDSeg.RecuperaRetPagoSeguro(txtAño.Text, IIf(Len(Trim(Right(cboMes.Text, 2))) = 1, "0", "") & Trim(Right(cboMes.Text, 2)), fsopecod, CInt(Mid(fsopecod, 3, 1))).EOF Then
        MsgBox "El " & IIf(Mid(fsopecod, 3, 6) = Mid(OpeCGOtrosOpeRetPagSeguroDesgravamenMN, 3, 6), "Retiro por Pago Seguro Desgravamen", "Retiro por Pago Seguro Contra Incendio") & " para este periodo ya fue registrado", vbInformation, "Aviso"
        ValidaSeleccionar = False
        cboMes.SetFocus
        Exit Function
    End If
End Function
Private Sub SetInicioControl()
    cboMes.ListIndex = -1
    ReDim FTSegAge(0)
    ReDim FTRetPagSegDet(0)
    Me.txtAño.Text = ""
    Me.TxtBuscaTrama.Text = ""
    cmdCargar.Enabled = False
    Me.txtBuscaEntidad.Text = ""
    Me.lblDescBanco.Caption = ""
    Me.lblDescCtaBanco.Caption = ""
    Me.lblTipoPago.Caption = ""
    Me.txtFecha.Text = Format(gdFecSis, "dd/MM/yyyy")
    Me.txtMovDesc.Text = ""
    Me.lblMontoText.Caption = "Monto " & gsSimbolo
    Me.lblMonto.Caption = ""
End Sub
Private Sub txtAño_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 And Len(Trim(Me.txtAño.Text)) <> 0 Then
        cmdSeleccionar.SetFocus
    End If
End Sub
Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    If KeyAscii = 8 Then SoloNumeros = KeyAscii
    If KeyAscii = 13 Then SoloNumeros = KeyAscii
End Function
Private Sub txtBuscaEntidad_EmiteDatos()
Dim oCtaIf As NCajaCtaIF
Dim oOpe As DOperacion
Set oOpe = New DOperacion
Set oCtaIf = New NCajaCtaIF

lblDescBanco = oCtaIf.NombreIF(Mid(txtBuscaEntidad, 4, 13))
lblDescCtaBanco = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscaEntidad, 18, 10)) + " " + txtBuscaEntidad.psDescripcion
lsCtaContBanco = oOpe.EmiteOpeCta(fsopecod, "H", , txtBuscaEntidad.Text, ObjEntidadesFinancieras)
    If lsCtaContBanco = "" Then
        MsgBox "Cuentas Contables no determinadas Correctamente" & Chr(13) & "consulte con el Dpto. de Contabilidad", vbInformation, "Aviso"
        txtBuscaEntidad.Text = ""
        lblDescBanco.Caption = ""
        lblDescCtaBanco.Caption = ""
        lblTipoPago.Caption = ""
        nTipoDoc = 0
        Exit Sub
    End If
    If Mid(txtBuscaEntidad.Text, 4, 13) = "1090100824640" Then
        nTipoDoc = TpoDocCarta
        lblTipoPago.Caption = "Transferencia"
    Else
        nTipoDoc = TpoDocCheque
        lblTipoPago.Caption = "Cheque"
    End If
txtFecha.SetFocus
End Sub
Private Sub TxtBuscaTrama_Click(psCodigo As String, psDescripcion As String)
    On Error GoTo ErrBuscaTrama
    TxtBuscaTrama.Text = Empty
    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
    dlgArchivo.ShowOpen
    If dlgArchivo.FileName <> Empty Then
        TxtBuscaTrama.Text = dlgArchivo.FileName
        cmdCargar.Enabled = False
    Else
        TxtBuscaTrama.Text = "NO SE ABRIO NINGÚN ARCHIVO"
        Exit Sub
    End If
    If TxtBuscaTrama.Text <> "" Then
        psCodigo = TxtBuscaTrama.Text
        psDescripcion = TxtBuscaTrama.Text
        cmdCargar.Enabled = True
        cmdCargar.SetFocus
    End If
    Exit Sub
ErrBuscaTrama:
    If Err.Number = 32755 Then
    ElseIf Err.Number = 1004 Then
        MsgBox "Archivo en uso. Ciérrelo y luego proceda a reemplazar.", vbExclamation, "Aviso"
    Else
        MsgBox "Error al momento de seleccionar el archivo", vbCritical, "Aviso"
    End If
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMovDesc.SetFocus
    End If
End Sub
Private Sub lblMonto_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

