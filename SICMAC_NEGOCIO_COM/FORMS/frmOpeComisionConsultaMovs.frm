VERSION 5.00
Begin VB.Form frmOpeComisionConsultaMovs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comisión por"
   ClientHeight    =   2745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6705
   Icon            =   "frmOpeComisionConsultaMovs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2745
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6495
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   360
         Left            =   5280
         TabIndex        =   4
         Top             =   2160
         Width           =   1050
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   360
         Left            =   4080
         TabIndex        =   3
         Top             =   2160
         Width           =   1050
      End
      Begin VB.ComboBox cboCuenta 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2160
         Width           =   2205
      End
      Begin VB.ComboBox cboTipoPago 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   240
         Width           =   1850
      End
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   330
         Left            =   1200
         TabIndex        =   1
         Top             =   1200
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
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
      Begin VB.Label lblMonedaCom 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1200
         TabIndex        =   13
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblCuenta 
         Caption         =   "Cuenta:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   12
         Top             =   2220
         Width           =   1095
      End
      Begin VB.Label lblLabelNombre 
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   1730
         Width           =   735
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   760
         Width           =   615
      End
      Begin VB.Label lblMontoCom 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1800
         TabIndex        =   9
         Top             =   720
         Width           =   975
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1200
         TabIndex        =   8
         Top             =   1680
         Width           =   5055
      End
      Begin VB.Label lblCliente 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   1260
         Width           =   735
      End
      Begin VB.Label lblTipoPago 
         Caption         =   "Tipo de Pago:"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   285
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmOpeComisionConsultaMovs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmOpeComisionConsultaMovs
'** Descripción : Formulario para pago de comisiones por Consulta de Movimientos o Consulta de Saldos
'**               según TI-ERS012-2015
'** Creación : JUEZ, 20151229 09:00:00 AM
'*****************************************************************************************************

Dim fnMoneda As Moneda
Dim fsCuenta As String
Dim frsCliente As ADODB.Recordset
Dim fnTipoPago As Integer
Dim fsPersCod As String
Dim fsPersNombre As String

Public Function Inicia(ByVal pnTipoOpe As Integer, ByVal pnMonto As Integer, ByVal pnMoneda As Moneda, ByVal prsCliente As ADODB.Recordset, _
                       ByRef pnTipoPago As Integer, ByVal psPersCod As String, ByRef psPersNombre As String) 'pnTipoOpe 1 Consulta Movimientos - 2 Consulta Saldos
    fsCuenta = ""
    fnTipoPago = 0
    Set frsCliente = prsCliente
    CargaTipoPago
    Me.lblMontoCom.Caption = Format(pnMonto, "#,##0.00")
    Me.lblMonedaCom.Caption = IIf(pnMoneda = gMonedaNacional, "S/", "$")
    fnMoneda = pnMoneda
    Me.Show 1
    Inicia = fsCuenta
    pnTipoPago = fnTipoPago
    psPersCod = fsPersCod
    psPersNombre = fsPersNombre
End Function

Private Sub cboTipoPago_Click()
    If Trim(Right(cboTipoPago.Text, 1)) = 1 Then
        lblCuenta.Visible = False
        cboCuenta.Visible = False
    ElseIf Trim(Right(cboTipoPago.Text, 1)) = 2 Then
        lblCuenta.Visible = True
        cboCuenta.Visible = True
    Else
        lblCuenta.Visible = False
        cboCuenta.Visible = False
    End If
    'TxtBCodPers.SetFocus
End Sub

Private Sub cboTipoPago_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtBCodPers.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
    If ValidaDatos Then
        fnTipoPago = Trim(Right((cboTipoPago.Text), 1))
        fsCuenta = Left(cboCuenta.Text, 18)
        fsPersCod = TxtBCodPers.Text
        fsPersNombre = lblPersNombre.Caption
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Function ValidaDatos() As Boolean
    ValidaDatos = False
    
    If Trim(cboTipoPago.Text) = "" Then
        MsgBox "El tipo de pago no fue seleccionado", vbInformation, "Aviso"
        cboTipoPago.SetFocus
        Exit Function
    End If
    If Trim(TxtBCodPers.Text) = "" Then
        MsgBox "El cliente no fue seleccionado", vbInformation, "Aviso"
        cboTipoPago.SetFocus
        Exit Function
    End If
    If Trim(Right((cboTipoPago.Text), 1)) = "2" And Trim(cboCuenta.Text) = "" Then
        MsgBox "La cuenta no fue seleccionado", vbInformation, "Aviso"
        cboTipoPago.SetFocus
        Exit Function
    End If
    
    ValidaDatos = True
End Function

Private Sub CargaTipoPago()
    Dim rsConst As New ADODB.Recordset
    Dim clsGen As New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(2037)
    Set clsGen = Nothing
    
    cboTipoPago.Clear
    While Not rsConst.EOF
        cboTipoPago.AddItem rsConst.Fields(0) & Space(100) & rsConst.Fields(1)
        rsConst.MoveNext
    Wend
    cboTipoPago.ListIndex = -1
    cboTipoPago_Click
End Sub

Private Sub TxtBCodPers_EmiteDatos()
    If Trim(TxtBCodPers.Text) <> "" Then
    Dim oCom As New COMDCredito.DCOMCredito
    Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim bValidaCliente As Boolean
    Dim bValidaCuenta As Boolean
    
        Set rs = oCom.RecuperaDatosComision(TxtBCodPers.Text, 2)
        Set oCom = Nothing
        
        bValidaCliente = False
        bValidaCuenta = False
        frsCliente.MoveFirst
        Do While Not frsCliente.EOF
            If rs!cPersCod = frsCliente!cPersCod Then
                bValidaCliente = True
                Exit Do
            End If
            frsCliente.MoveNext
        Loop
        If Not bValidaCliente Then
            MsgBox "Debe seleccionar un cliente de la cuenta", vbInformation, "Aviso"
            TxtBCodPers.Text = ""
            TxtBCodPers.SetFocus
            Exit Sub
        End If
        lblPersNombre.Caption = rs!cPersNombre
        TxtBCodPers.Enabled = False
        
        Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
            Set rsCuentas = oDCapGen.GetCuentasPersona(TxtBCodPers.Text, , True, True, fnMoneda, , , , True, "0,2")
        Set oDCapGen = Nothing
        
        Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
        If Not rsCuentas.EOF And Not rsCuentas.BOF Then
            Do While Not rsCuentas.EOF
                If Mid(rsCuentas("cCtaCod"), 9, 1) = fnMoneda And Mid(rsCuentas("cCtaCod"), 6, 3) <> gCapPlazoFijo Then
                    If oNCapMov.ValidaSaldoCuenta(rsCuentas("cCtaCod"), CDbl(Me.lblMontoCom.Caption)) Then
                        cboCuenta.AddItem rsCuentas("cCtaCod")
                        bValidaCuenta = True
                    End If
                End If
                rsCuentas.MoveNext
            Loop
        End If
        If Trim(Right((cboTipoPago.Text), 1)) = "2" Then
            If Not bValidaCuenta Then
                MsgBox "El cliente selecionado no tiene cuentas disponibles a debitar", vbInformation, "Aviso"
                TxtBCodPers.Enabled = True
                TxtBCodPers.Text = ""
                lblPersNombre.Caption = ""
                TxtBCodPers.SetFocus
            Else
                cboCuenta.SetFocus
            End If
        End If
        Set oNCapMov = Nothing
    Else
        lblPersNombre.Caption = ""
    End If
End Sub
