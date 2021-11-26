VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmGestionCartaFianza 
   Caption         =   "Carta Fianza - "
   ClientHeight    =   6480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7425
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmGestionCartaFianza.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6480
   ScaleWidth      =   7425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fmGestion 
      Caption         =   "Carta Fianza"
      Height          =   5775
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7215
      Begin VB.ComboBox cboEstado 
         Height          =   315
         ItemData        =   "frmGestionCartaFianza.frx":030A
         Left            =   2040
         List            =   "frmGestionCartaFianza.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   2450
         Width           =   1575
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmGestionCartaFianza.frx":0371
         Left            =   120
         List            =   "frmGestionCartaFianza.frx":037B
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   2450
         Width           =   1575
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   2475
         Left            =   120
         MaxLength       =   2000
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   3120
         Width           =   6915
      End
      Begin VB.TextBox txtNroContrato 
         Height          =   300
         Left            =   2040
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1845
         Width           =   1815
      End
      Begin VB.TextBox txtNroFC 
         Height          =   300
         Left            =   120
         MaxLength       =   50
         TabIndex        =   3
         Top             =   1850
         Width           =   1815
      End
      Begin MSComCtl2.DTPicker txtFVencimiento 
         Height          =   285
         Left            =   5400
         TabIndex        =   5
         Top             =   1850
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   225116161
         CurrentDate     =   41586
      End
      Begin MSComCtl2.DTPicker txtFEmision 
         Height          =   285
         Left            =   3960
         TabIndex        =   4
         Top             =   1850
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   503
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   225116161
         CurrentDate     =   41586
      End
      Begin Sicmact.TxtBuscar txtPerProv 
         Height          =   315
         Left            =   120
         TabIndex        =   0
         Top             =   600
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   556
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin Sicmact.TxtBuscar txtBuscaIF 
         Height          =   360
         Left            =   120
         TabIndex        =   1
         Top             =   1200
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   635
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label lblestado 
         Caption         =   "Estado"
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   2200
         Width           =   855
      End
      Begin VB.Label Label8 
         Caption         =   "Moneda"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   2200
         Width           =   1095
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° Contrato"
         Height          =   195
         Left            =   2040
         TabIndex        =   17
         Top             =   1605
         Width           =   870
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción del Servicio"
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   2880
         Width           =   1665
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "N° Carta Fianza"
         Height          =   195
         Left            =   120
         TabIndex        =   15
         Top             =   1605
         Width           =   1140
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Emisión"
         Height          =   195
         Left            =   3960
         TabIndex        =   14
         Top             =   1600
         Width           =   720
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "F. Vencimiento"
         Height          =   195
         Left            =   5400
         TabIndex        =   13
         Top             =   1600
         Width           =   1050
      End
      Begin VB.Label lblDescProveedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Top             =   600
         Width           =   4695
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   750
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Banco"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   435
      End
      Begin VB.Label lblDescBanco 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   2040
         TabIndex        =   9
         Top             =   1200
         Width           =   4695
      End
   End
   Begin VB.CommandButton cmdAccionar 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   6240
      TabIndex        =   8
      Top             =   6000
      Width           =   1095
   End
End
Attribute VB_Name = "frmGestionCartaFianza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmCartaFianza
'** Descripción : Registro de las Cartas Fianza creado segun RFC1902190004
'** Creación : TORE, 20190717 10:45:00 AM
'********************************************************************
Option Explicit
Dim oCartaFianza As DCartaFianza
Dim oMov As DMov
Dim oOpe As DOperacion
Dim oCtaIf As NCajaCtaIF
Dim oRS As ADODB.Recordset
Dim oCons As DConstante


Private Sub Form_Load()
    'Centrar el formulario
    With Screen
        Move (.Width - Width) / 2, (.Height - Height) / 2, Width, Height
    End With
    Call CargarInstituciones
End Sub

Public Sub Inicia(ByVal psNroCartaFianza As String, ByRef psCodProv As String, ByRef psDescProv As String)
    Dim rsConst As New ADODB.Recordset
    Set oCartaFianza = New DCartaFianza
    Set oRS = New ADODB.Recordset
    Set oCons = New DConstante
    
    Dim I As Integer
    Set oRS = oCartaFianza.ObtenerInfoCartaFianzaNroCF(psNroCartaFianza)
    If Not oRS.EOF And Not oRS.BOF Then
        Me.Caption = Me.Caption & " Actualización"
        Set rsConst = oCons.RecuperaConstantes(5005)
        Call CargaCombo(rsConst, cboEstado, 0, 1)
        For I = 1 To oRS.RecordCount
            txtPerProv.Text = oRS!cPersCodProv
            lblDescProveedor.Caption = oRS!cPersNombreProv
            txtBuscaIF.Text = oRS!cCodIfi
            lblDescBanco.Caption = oRS!cNombreIFI
            txtNroContrato.Text = oRS!cNroContrato
            txtNroFC.Text = oRS!cNroCartaFianza
            txtFEmision.value = oRS!dFEmision
            txtFVencimiento.value = oRS!dFVencimiento
            txtDescripcion.Text = oRS!cDescripcion
            cboEstado.ListIndex = IndiceListaCombo(cboEstado, Trim((oRS!nEstadoCF)))
            cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, Trim((oRS!nMoneda)))
            oRS.MoveNext
        Next
        txtPerProv.Enabled = False
        txtNroFC.Enabled = False
        lblDescProveedor.Enabled = False
        
        cmdAccionar.Caption = "Actualizar"
    Else
        Call LimpiarControles(1)
        txtPerProv.Text = psCodProv
        lblDescProveedor.Caption = psDescProv
        cboEstado.Visible = False
        lblestado.Visible = False
    End If
    Me.Show 1
End Sub

Public Sub VerCF(ByVal psNroCartaFianza As String)
    Dim rsConst As New ADODB.Recordset
    Set oCartaFianza = New DCartaFianza
    Set oRS = New ADODB.Recordset
    Set oCons = New DConstante

    Dim I As Integer
    Set oRS = oCartaFianza.ObtenerInfoCartaFianzaNroCF(psNroCartaFianza)
    If Not oRS.EOF And Not oRS.BOF Then
        Me.Caption = Me.Caption & " Detalle"
        Set rsConst = oCons.RecuperaConstantes(5005)
        Call CargaCombo(rsConst, cboEstado, 0, 1)
        For I = 1 To oRS.RecordCount
            txtPerProv.Text = oRS!cPersCodProv
            lblDescProveedor.Caption = oRS!cPersNombreProv
            txtBuscaIF.Text = oRS!cCodIfi
            lblDescBanco.Caption = oRS!cNombreIFI
            txtNroContrato.Text = oRS!cNroContrato
            txtNroFC.Text = oRS!cNroCartaFianza
            txtFEmision.value = oRS!dFEmision
            txtFVencimiento.value = oRS!dFVencimiento
            txtDescripcion.Text = oRS!cDescripcion
            cboEstado.ListIndex = IndiceListaCombo(cboEstado, Trim((oRS!nEstadoCF)))
            cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, Trim((oRS!nMoneda)))
            oRS.MoveNext
        Next
    End If
    cmdAccionar.Visible = False
    txtPerProv.Enabled = False
    txtBuscaIF.Enabled = False
    'cboMoneda.Enabled = False
    'cboEstado.Enabled = False
    'txtFEmision.Enabled = False
    'txtFVencimiento.Enabled = False
    Me.Show 1
End Sub


Private Sub LimpiarControles(ByVal pnAccionar As Integer)
    If pnAccionar = 1 Then
        Me.Caption = Me.Caption & " Registro"
        txtNroFC.Enabled = True
        cmdAccionar.Caption = "Registrar"
        txtPerProv.Enabled = False
        lblDescProveedor.Enabled = False
        txtFEmision.value = gdFecSis
        txtFVencimiento.value = gdFecSis
    Else
        txtBuscaIF.Text = ""
        lblDescBanco.Caption = ""
        txtNroContrato.Text = ""
        txtNroFC.Text = ""
        txtFEmision.value = gdFecSis
        txtFVencimiento.value = gdFecSis
        txtDescripcion.Text = ""
    End If

End Sub

Private Sub cmdAccionar_Click()
    If ValidaCampos() = False Then
        Set oMov = New DMov
        Set oCartaFianza = New DCartaFianza
        
        Dim lsCodProv As String, lsNroContrato As String, lsDescripcion As String, lsNroCartaFIanza As String
        Dim lsCodIFI As String, ldFEmision As String, ldFVencimiento As String, lsUltimaActualizacion As String
        Dim lnMoneda As Integer, lnEstado As Integer
        Dim lsMsjValidacion As String
        
        lsUltimaActualizacion = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        lsCodProv = txtPerProv.Text
        lsNroContrato = Trim(txtNroContrato.Text)
        lsDescripcion = Trim(txtDescripcion.Text)
        lsNroCartaFIanza = Trim$(txtNroFC.Text)
        lsCodIFI = Trim(txtBuscaIF.Text) 'Right(Trim(txtBuscaIF.Text), 13)
        lnMoneda = CInt(Trim(Right(cboMoneda.Text, 1)))
        ldFEmision = Format(txtFEmision.value, "yyyyMMdd")
        ldFVencimiento = Format(txtFVencimiento.value, "yyyyMMdd")
        lnEstado = CInt(IIf(Trim(Right(cboEstado.Text, 1)) = "", 0, Trim(Right(cboEstado.Text, 1))))
        
        lsMsjValidacion = oCartaFianza.ObtenerValidacionCartaFianza(lsNroCartaFIanza, _
                                                                    lsNroContrato, _
                                                                   IIf(cmdAccionar.Caption = "Registrar", 1, 2))
        If lsMsjValidacion = "" Then
            oCartaFianza.RegistraCartaFianza lsNroCartaFIanza, lsCodProv, lsNroContrato, _
                                            lsDescripcion, lsCodIFI, lnMoneda, ldFEmision, ldFVencimiento, _
                                            lsUltimaActualizacion, lnEstado
        Else
            MsgBox lsMsjValidacion, vbInformation, "Aviso"
            Exit Sub
        End If
        Unload Me
    End If

End Sub

Private Sub txtPerProv_EmiteDatos()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oProv As DLogProveedor
    Set oProv = New DLogProveedor
    Dim lsCodProv As String, lsDescProv As String
    
    Me.lblDescProveedor.Caption = txtPerProv.psDescripcion
    lsCodProv = txtPerProv.psCodigoPersona
    lsDescProv = txtPerProv.psDescripcion
    If txtPerProv.psDescripcion <> "" Then
        Set rs = oProv.GetProveedorAgeRetBuenCont(lsCodProv)
        If rs.EOF And rs.BOF Then
            MsgBox "La persona ingresada no esta registrada como proveedor o tiene el estado de Desactivado, debe regsitrarlo o activarlo.", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Sub txtBuscaIF_EmiteDatos()
    lblDescBanco.Caption = txtBuscaIF.psDescripcion
End Sub

Private Sub CargarInstituciones()
    Set oOpe = New DOperacion
   txtBuscaIF.psRaiz = "Instituciones Financieras"
   txtBuscaIF.rs = oOpe.GetOpeObj("421301", "1")
End Sub

Private Function ValidaCampos() As Boolean
    If Trim(txtPerProv.Text) = "" Then
        MsgBox "Debe cargar información del Proveedor.", vbInformation, "Aviso"
        ValidaCampos = True
        Exit Function
    End If
    If Trim(txtBuscaIF.Text) = "" Then
        MsgBox "Debe cargar información de la IFI.", vbInformation, "Aviso"
        ValidaCampos = True
        Exit Function
    End If
    If Trim(cboMoneda.Text) = "" Then
        MsgBox "Debe seleccionar el tipo de moneda.", vbInformation, "Aviso"
        ValidaCampos = True
        Exit Function
    End If
    If Trim(txtNroContrato.Text) = "" Then
        MsgBox "Ingrese el N° de contrato.", vbInformation, "Aviso"
        ValidaCampos = True
        Exit Function
    End If
     If Not ValidaCadena(Trim(txtNroContrato.Text), 2) Then
        MsgBox "Solo se permite letras mayúsculas y números para el N° de Contrato", vbInformation, "Aviso"
        txtNroContrato.SetFocus
        ValidaCampos = True
        Exit Function
    End If
    If Trim(txtNroFC.Text) = "" Then
        MsgBox "Ingrese el N° de Carta Fianza.", vbInformation, "Aviso"
        ValidaCampos = True
        Exit Function
    End If
     If Not ValidaCadena(Trim(txtNroFC.Text), 2) Then
        MsgBox "Solo se permite letras mayúsculas y números para el N° de la Carta Fianza.", vbInformation, "Aviso"
        txtNroFC.SetFocus
        ValidaCampos = True
        Exit Function
    End If
    If Trim(txtDescripcion.Text) = "" Then
        MsgBox "Ingrese la descripción de la Carta Fianza.", vbInformation, "Aviso"
        txtDescripcion.SetFocus
        ValidaCampos = True
        Exit Function
    End If
    If Not ValidaCadena(Trim(txtDescripcion.Text), 1) Then
        MsgBox "No se permiten carácteres especiales en la descripción de la Carta Fianza.", vbInformation, "Aviso"
        txtDescripcion.SetFocus
        ValidaCampos = True
        Exit Function
    End If
    If txtFEmision.value > txtFVencimiento.value Then
        MsgBox "La fecha de emisión no debe ser mayor a la fecha de vencimiento.", vbInformation, "Aviso"
        txtFEmision.value = gdFecSis
        txtFEmision.SetFocus
        ValidaCampos = True
        Exit Function
    End If
       If txtFVencimiento.value < txtFEmision.value Then
        MsgBox "La fecha de vencimiento no debe ser menor a la fecha de emisión.", vbInformation, "Aviso"
        txtFVencimiento.value = gdFecSis
        txtFVencimiento.SetFocus
        ValidaCampos = True
        Exit Function
    End If
    ValidaCampos = False
End Function


Private Function ValidaCadena(ByVal cadena As String, ByVal tpoCampo As Integer) As Boolean
Dim tamanioCadena As String
Dim cadenaResultado As String
Dim caracteresValidos As String
Dim caracteresActual As String
Dim I As Integer

tamanioCadena = Len(cadena)
If tamanioCadena > 0 Then
    If tpoCampo = 1 Then
        caracteresValidos = " 0123456789abcdefghijklmnñopqrstuwxvyzABCDEFGHIJKLMNÑOPQRSTUWVXYZ-óíÓÍÑ,:."
    ElseIf tpoCampo = 2 Then
        caracteresValidos = " 0123456789ABCDEFGHIJKLMNOPQRSTUWVXYZ-"
    End If
    
    
    For I = 1 To tamanioCadena
        caracteresActual = Mid(cadena, I, 1)
        
        If InStr(caracteresValidos, caracteresActual) Then
            ValidaCadena = True
        Else
            ValidaCadena = False
            Exit Function
        End If
    Next
End If
End Function



Public Sub CargaCombo(ByVal prsCombo As ADODB.Recordset, ByVal CtrlCombo As ComboBox, ByVal pnFiel1 As Integer, ByVal pnFiel2 As Integer)
    CtrlCombo.Clear
    While Not prsCombo.EOF
        CtrlCombo.AddItem prsCombo.Fields(pnFiel1) & Space(100) & prsCombo.Fields(pnFiel2)
        prsCombo.MoveNext
    Wend
End Sub

