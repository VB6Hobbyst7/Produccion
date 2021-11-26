VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCapRegVouDep 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Voucher de Depósito"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6705
   Icon            =   "frmCapRegVouDep.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame5 
      Caption         =   "Glosa"
      Height          =   1095
      Left            =   120
      TabIndex        =   21
      Top             =   4440
      Width           =   6495
      Begin VB.TextBox txtGlosa 
         Height          =   735
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   6255
      End
   End
   Begin VB.Frame Frame4 
      Caption         =   "Tipo"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1455
      Left            =   120
      TabIndex        =   16
      Top             =   2880
      Width           =   6495
      Begin VB.TextBox txtMontoVoucher 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   3600
         MaxLength       =   15
         TabIndex        =   27
         Top             =   960
         Width           =   1305
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5040
         TabIndex        =   10
         Top             =   960
         Width           =   1335
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   5040
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.CheckBox ckbConfirmar 
         Caption         =   "&Confirmado"
         Height          =   375
         Left            =   5040
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.ComboBox cboMotivo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCapRegVouDep.frx":030A
         Left            =   840
         List            =   "frmCapRegVouDep.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox txtNumeroVoucher 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   840
         TabIndex        =   4
         Top             =   630
         Width           =   4095
      End
      Begin MSMask.MaskEdBox txtFechaVoucher 
         Height          =   285
         Left            =   840
         TabIndex        =   26
         Top             =   960
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         BackColor       =   16777215
         ForeColor       =   -2147483630
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
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
      Begin VB.Label Label8 
         Caption         =   "Monto"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Fecha"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   960
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Número"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Motivo"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   5160
      TabIndex        =   15
      Top             =   480
      Width           =   1455
      Begin VB.OptionButton OptMoneda 
         Caption         =   "&Extranjera"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   1
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton OptMoneda 
         Caption         =   "&Nacional"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Destino"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   1680
      Width           =   6495
      Begin SICMACT.TxtBuscar txtCtaIF 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   1755
         _ExtentX        =   3096
         _ExtentY        =   503
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin VB.Label lblCtaIF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   6195
      End
      Begin VB.Label lblNombreIF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1920
         TabIndex        =   23
         Top             =   240
         Width           =   4395
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Origen"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   11
      Top             =   480
      Width           =   4935
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   285
         Left            =   960
         TabIndex        =   0
         Top             =   240
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   503
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblNomPer 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   960
         TabIndex        =   22
         Top             =   600
         Width           =   3795
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Persona"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   735
      End
   End
   Begin MSMask.MaskEdBox txtFechaMov 
      Height          =   330
      Left            =   840
      TabIndex        =   25
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   582
      _Version        =   393216
      Appearance      =   0
      BackColor       =   16777215
      ForeColor       =   -2147483630
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
   Begin VB.Label Label1 
      Caption         =   "Fecha Registro"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmCapRegVouDep"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapRegVouDep
'*** Descripción : Formulario para registrar el vouchert de Depósito.
'*** Creación : ELRO el 20120530 07:40:21 PM, según OYP-RFC024-2012
'********************************************************************
Option Explicit

Public fnMovNroPen As Long


Private Sub cargarEntidadesFinanciera()
    Dim oDOperacion As New clases.DOperacion
    Dim lsComparar As String
    
    txtCtaIF = ""
    txtCtaIF.psCodigoPersona = ""
    lblNombreIF = ""
    lblCtaIF = ""

    Me.txtCtaIF.psRaiz = "Cuentas de Instituciones Financieras"
    lsComparar = "_1_[12]" & IIf(OptMoneda(0).value, "1", "2") & "%"
    Me.txtCtaIF.rs = oDOperacion.listarCuentasEntidadesFinacieras(lsComparar, IIf(OptMoneda(0).value, "1", "2"))
    
    If OptMoneda(1).value = True Then
        txtCtaIF.BackColor = &HC0FFC0
        lblNombreIF.BackColor = &HC0FFC0
        lblCtaIF.BackColor = &HC0FFC0
        txtMontoVoucher.BackColor = &HC0FFC0
    Else
        txtCtaIF.BackColor = &H80000005
        lblNombreIF.BackColor = &H80000005
        lblCtaIF.BackColor = &H80000005
        txtMontoVoucher.BackColor = &H80000005
    End If
    
                  
    Set oDOperacion = Nothing
End Sub

Private Sub cargarMotivo()
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim rsMotivos As ADODB.Recordset
    Set rsMotivos = New ADODB.Recordset
    
    Set rsMotivos = oNCOMCaptaGenerales.obtenerMotivoCapVoucherDeposito
    
    cboMotivo.Clear
    Do While Not rsMotivos.EOF
        cboMotivo.AddItem rsMotivos!vMotivo
        rsMotivos.MoveNext
    Loop
    cboMotivo.ListIndex = -1
    
    Set rsMotivos = Nothing
    Set oNCOMCaptaGenerales = Nothing
    
End Sub

Private Sub LimpiarCampos()
    TxtBCodPers = ""
    TxtBCodPers.psCodigoPersona = ""
    lblNomPer = ""
    txtCtaIF = ""
    txtCtaIF.psDescripcion = ""
    lblNombreIF = ""
    lblCtaIF = ""
    txtNumeroVoucher = ""
    txtFechaVoucher = "__/__/____"
    txtMontoVoucher = "0.00"
    txtGlosa = ""
End Sub

Private Function validarCampos() As Boolean
    Dim lsMensaje As String
    validarCampos = False
    
    If Trim(txtGlosa) = "" Then
        MsgBox "Debe ingresar la Glosa.", vbInformation, "Aviso"
        txtGlosa.SetFocus
        Exit Function
    End If
    
    If Trim(TxtBCodPers) = "" Then
        MsgBox "Debe seleccionar la Persona del Voucher.", vbInformation, "Aviso"
        TxtBCodPers.SetFocus
        Exit Function
    End If
    
    If Trim(txtCtaIF) = "" Then
        MsgBox "Debe seleccionar la Entidad Financiera.", vbInformation, "Aviso"
        txtCtaIF.SetFocus
        Exit Function
    End If
    
    If cboMotivo.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Motivo.", vbInformation, "Aviso"
        cboMotivo.SetFocus
        Exit Function
    End If
    
    If Trim(txtNumeroVoucher) = "" Then
        MsgBox "Debe ingresar el Número del Voucher.", vbInformation, "Aviso"
        txtNumeroVoucher.SetFocus
        Exit Function
    End If
    
    If Trim(txtMontoVoucher) = "" Then
        MsgBox "Debe ingresar el Monto del Voucher.", vbInformation, "Aviso"
        txtNumeroVoucher.SetFocus
        Exit Function
    End If
    
    If Not IsNumeric(txtMontoVoucher) Then
        MsgBox "Debe ingresar el Monto del Voucher.", vbInformation, "Aviso"
        txtNumeroVoucher.SetFocus
        Exit Function
    End If
    
    If CCur(txtMontoVoucher) = 0# Then
        MsgBox "Debe ingresar el Monto del Voucher.", vbInformation, "Aviso"
        txtMontoVoucher.SetFocus
        Exit Function
    End If
    
    If txtFechaVoucher = "__/__/____" Then
        MsgBox "Debe ingresar la Fecha del Voucher.", vbInformation, "Aviso"
        txtFechaVoucher.SetFocus
        Exit Function
    End If
    
    lsMensaje = ValidaFecha(txtFechaVoucher.Text)

    If Trim(lsMensaje) <> "" Then
        MsgBox lsMensaje, vbInformation, "!Aviso¡"
        txtFechaVoucher.SetFocus
        Exit Function
    End If
    
    If CDate(txtFechaVoucher) > gdFecSis Then
        MsgBox "Fecha del Voucher es mayor que la Fecha de Sistema.", vbInformation, "Aviso"
        txtFechaVoucher.SetFocus
        Exit Function
    End If
             

    
    validarCampos = True
End Function



Private Sub cboMotivo_KeyPress(KeyAscii As Integer)
 If KeyAscii = 13 Then
    txtNumeroVoucher.SetFocus
 End If
End Sub

Private Sub CmdAceptar_Click()
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Dim oNContFunciones As clases.NContFunciones
    Set oNContFunciones = New NContFunciones
    Dim lsMovNro As String
    Dim lnConfirmacion As Long
    Dim lcBanco, lsCtafiltro As String
  
    If validarCampos = False Then Exit Sub
    
    'EJVG20130531 BCRP Otra Cta Contable
    'lcBanco = "11" & IIf(OptMoneda(0).value, "1", "2") & "301"
    lcBanco = "11" & IIf(OptMoneda(0).value, "1", "2") & IIf(Mid(txtCtaIF, 4, 13) = "1090100822183", "2", "3") & "01"
    txtNumeroVoucher.Text = Trim(txtNumeroVoucher.Text)
    'END EJVG *******
    
    lsCtafiltro = oNContFunciones.GetFiltroObjetos(1, lcBanco, txtCtaIF, False)
    
    If lsCtafiltro = "" Then
        MsgBox "Esta cuenta contable " & lcBanco & " no esta registrado en CtaIFFiltro, comunicarse con TI", vbInformation, "Aviso"
        Exit Sub
    End If
                    
    If oNContFunciones.verificarUltimoNivelCta(lcBanco & lsCtafiltro) = False Then
       MsgBox "La Cuenta Contable " & lcBanco + lsCtafiltro & " no es de Ultimo Nivel, comunicarse con Contabilidad", vbInformation, "Aviso"
       Exit Sub
    End If
    

    If CDate(txtFechaVoucher) < gdFecSis Then
        frmCapRegVouDepPen.iniciarListado CDate(txtFechaVoucher), Left(txtCtaIF, 2), Mid(txtCtaIF, 4, 13), Mid(txtCtaIF, 18, 10), CCur(txtMontoVoucher), IIf(OptMoneda(0).value, "1", "2"), frmCapRegVouDep
        If fnMovNroPen = 0 Then
            MsgBox "Debe relacionar con una Pendiente el Voucher.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    lnConfirmacion = oNCOMCaptaGenerales.guardarVoucherDeposito(lsMovNro, _
                                                                gdFecSis, _
                                                                TxtBCodPers.psCodigoPersona, _
                                                                Left(txtCtaIF, 2), _
                                                                Mid(txtCtaIF, 4, 13), _
                                                                Mid(txtCtaIF, 18, 10), _
                                                                Right(cboMotivo, 2), _
                                                                ckbConfirmar, _
                                                                txtNumeroVoucher, _
                                                                CCur(txtMontoVoucher), _
                                                                CDate(txtFechaVoucher), _
                                                                Right(gsCodAge, 2), _
                                                                IIf(OptMoneda(0).value, "1", "2"), _
                                                                txtGlosa, _
                                                                gsOpeCod, _
                                                                lcBanco, _
                                                                lsCtafiltro, _
                                                                fnMovNroPen)
    If lnConfirmacion > 0 Then
        MsgBox "Se realizó correctamente la operación", vbInformation, "Aviso"
        LimpiarCampos
        Unload Me
    Else
        MsgBox "No se realizó correctamente la operación", vbInformation, "Aviso"
        LimpiarCampos
        Unload Me
    End If
    
    fnMovNroPen = 0
    lnConfirmacion = 0
    lsMovNro = 0
    Set oNCOMContFunciones = Nothing
    Set oNCOMCaptaGenerales = Nothing
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    txtFechaMov = gdFecSis
    cargarEntidadesFinanciera
    cargarMotivo
    LimpiarCampos
End Sub

Private Sub OptMoneda_Click(Index As Integer)
    Call cargarEntidadesFinanciera
End Sub

Private Sub TxtBCodPers_EmiteDatos()
    If TxtBCodPers = "" Then Exit Sub
    If TxtBCodPers.psCodigoPersona = gsCodPersUser Then
        MsgBox "No se puede registrar un vouchert de si mismo.", vbInformation, "Aviso"
        TxtBCodPers = ""
        Exit Sub
    End If
    lblNomPer = TxtBCodPers.psDescripcion
    txtCtaIF.SetFocus
End Sub

Private Sub txtCtaIF_EmiteDatos()
    Dim lsNombreBanco As String
    Dim lsCuenta As String
    
    Dim oNCajaCtaIF As clases.NCajaCtaIF
    Set oNCajaCtaIF = New clases.NCajaCtaIF
    Dim oDOperacion As clases.DOperacion
    Set oDOperacion = New clases.DOperacion
    
    If Me.txtCtaIF <> "" Then
        lsNombreBanco = oNCajaCtaIF.NombreIF(Mid(txtCtaIF, 4, 13))
        lsCuenta = oDOperacion.recuperaTipoCuentaEntidadFinaciera(Mid(txtCtaIF, 18, 10)) & " " & txtCtaIF.psDescripcion
        lblNombreIF = lsNombreBanco
        lblCtaIF = lsCuenta
    End If
      
    cboMotivo.SetFocus
    Set oNCajaCtaIF = Nothing
    Set oDOperacion = Nothing
End Sub

Private Sub txtFechaVoucher_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim lsMensaje As String
        lsMensaje = ValidaFecha(txtFechaVoucher.Text)
    
        If Trim(lsMensaje) <> "" Then
            MsgBox lsMensaje, vbInformation, "!Aviso¡"
            txtFechaVoucher.SetFocus
            Exit Sub
        ElseIf Trim(lsMensaje) = "" Then
            txtMontoVoucher.SetFocus
        End If
    End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       ckbConfirmar.SetFocus
    End If
End Sub

Private Sub txtMontoVoucher_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtMontoVoucher) Then
        MsgBox "Ingrese un monto válido", vbInformation, "Mensaje"
        Exit Sub
    End If
    KeyAscii = NumerosDecimales(txtMontoVoucher, KeyAscii, 15)
    If KeyAscii <> 13 Then Exit Sub
    If KeyAscii = 13 Then
       txtGlosa.SetFocus
    End If
End Sub

Private Sub txtNumeroVoucher_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFechaVoucher.SetFocus
    End If
End Sub
