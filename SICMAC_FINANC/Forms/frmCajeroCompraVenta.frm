VERSION 5.00
Begin VB.Form frmCajeroCompraVenta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Compra Venta Moneda Extranjera"
   ClientHeight    =   4200
   ClientLeft      =   1875
   ClientTop       =   2205
   ClientWidth     =   6570
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajeroCompraVenta.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   5220
      TabIndex        =   4
      Top             =   3750
      Width           =   1155
   End
   Begin VB.CommandButton CmdGuardar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   4050
      TabIndex        =   3
      Top             =   3750
      Width           =   1185
   End
   Begin VB.Frame fraTipoCambio 
      Height          =   1065
      Left            =   60
      TabIndex        =   9
      Top             =   3060
      Width           =   3825
      Begin VB.TextBox TxtMontoPagar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1665
         TabIndex        =   10
         Top             =   600
         Width           =   1695
      End
      Begin Sicmact.EditMoney txtImporte 
         Height          =   375
         Left            =   1665
         TabIndex        =   2
         Top             =   195
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTipoCambio 
         Caption         =   "Monto a Cambiar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   255
         Width           =   1575
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto a Pagar:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         Top             =   660
         Width           =   1575
      End
      Begin VB.Label lblsimbolosoles2 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3435
         TabIndex        =   12
         Top             =   690
         Width           =   285
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "$."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   11
         Top             =   285
         Width           =   180
      End
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Datos de la Persona"
      Height          =   2745
      Left            =   60
      TabIndex        =   7
      Top             =   315
      Width           =   6315
      Begin Sicmact.FlexEdit fgDocs 
         Height          =   1230
         Left            =   1515
         TabIndex        =   1
         Top             =   1335
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2170
         Cols0           =   4
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Documento-N° Doc.-Tipo"
         EncabezadosAnchos=   "450-1500-1800-0"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X"
         ListaControles  =   "0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.TxtBuscar txtBuscaPers 
         Height          =   330
         Left            =   180
         TabIndex        =   0
         Top             =   240
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   582
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
         TipoBusqueda    =   3
         sTitulo         =   ""
         TipoBusPers     =   1
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "N° Documentos :"
         Height          =   210
         Left            =   195
         TabIndex        =   8
         Top             =   1335
         Width           =   1200
      End
      Begin VB.Label lblPersDireccion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         TabIndex        =   6
         Top             =   930
         Width           =   5670
      End
      Begin VB.Label lblPersNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   180
         TabIndex        =   5
         Top             =   585
         Width           =   5670
      End
   End
   Begin VB.Label lblTpoCambioDia 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "0.00"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   330
      Left            =   5250
      TabIndex        =   17
      Top             =   3210
      Width           =   1140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo Cambio:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   210
      Left            =   4050
      TabIndex        =   16
      Top             =   3270
      Width           =   1080
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      Caption         =   "COMPRA MONEDA EXTRANJERA"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   1410
      TabIndex        =   15
      Top             =   75
      Width           =   4140
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00E0E0E0&
      BackStyle       =   1  'Opaque
      FillColor       =   &H00E0E0E0&
      Height          =   360
      Left            =   3975
      Top             =   3195
      Width           =   2430
   End
End
Attribute VB_Name = "frmCajeroCompraVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oPers As UPersona
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim lsAgencia As String
Dim lbSalir As Boolean
Dim lsDocumento  As String

Private Sub cmdGuardar_Click()
Dim oCajero As nCajero
Dim lsMovNro As String
Dim oGen  As NContFunciones
Set oGen = New NContFunciones
Set oCajero = New nCajero
If ValidaInterfaz = False Then Exit Sub
If MsgBox("Desea grabar la Operación de Compra/Venta??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If oCajero.GrabaCompraVenta(gsFormatoFecha, lsMovNro, gsOpeCod, gsOpeDesc, CCur(txtImporte), CCur(lblTpoCambioDia), txtBuscaPers) = 0 Then
        
        Dim oImp As NContImprimir
        Dim lsTexto As String
        Dim lbReimp As Boolean
        Set oImp = New NContImprimir
        
        
        lbReimp = True
        Do While lbReimp
            oImp.ImprimeBoletaCompraVenta lblTitulo, "", lblPersNombre, lblPersDireccion, lsDocumento, _
                    CCur(lblTpoCambioDia), gsOpeCod, CCur(txtImporte), CCur(TxtMontoPagar), gsNomAge, lsMovNro, sLpt
        
            If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbReimp = False
            End If
        Loop
        Set oImp = Nothing
        txtBuscaPers = ""
        lblPersDireccion = ""
        lblPersNombre = ""
        fgDocs.Clear
        fgDocs.FormaCabecera
        fgDocs.Rows = 2
        txtImporte = 0
        TxtMontoPagar = "0.00"
        txtBuscaPers.SetFocus
    End If
    Set oGen = Nothing
    Set oCajero = Nothing
End If

End Sub
Function ValidaInterfaz() As Boolean
ValidaInterfaz = True
If txtBuscaPers = "" Then
    MsgBox "Persona no Ingresada", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtBuscaPers.SetFocus
    Exit Function
End If
If Val(txtImporte) = 0 Then
    MsgBox "Importe de Operación no Ingresado", vbInformation, "Aviso"
    ValidaInterfaz = False
    txtImporte.SetFocus
    Exit Function
End If
If Val(TxtMontoPagar) = 0 Then
    MsgBox "Monto a Pagar no válido para Operación", vbInformation, "Aviso"
    ValidaInterfaz = False
    Exit Function
End If



End Function
Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If lbSalir Then
    Unload Me
End If
End Sub

Private Sub Form_Load()
Dim oOpe As DOperacion
Set oOpe = New DOperacion
CentraForm Me
Me.Caption = gsOpeDesc
txtImporte.psSoles False
lbSalir = False
lblTpoCambioDia = Format(IIf(gsOpeCod = gOpeCajeroMECompra, gnTipCambioC, gnTipCambioV), "#,#0.00")
If Val(lblTpoCambioDia) = 0 Then
    MsgBox "Tipo de Cambio no ha sido Ingresado. Por favor Ingrese Tipo Cambio del Día", vbInformation, "Aviso"
    lbSalir = True
    Exit Sub
End If
Select Case gsOpeCod
    Case gOpeCajeroMECompra
        lblTitulo = "COMPRA MONEDA EXTRANJERA"
        Me.lblMonto = "Monto a Pagar"
    Case gOpeCajeroMEVenta
        lblTitulo = "VENTA MONEDA EXTRANJERA"
        Me.lblMonto = "Monto a Recibir"
End Select
TxtMontoPagar = "0.00"

'falta definir el objeto area agencia con que va a trabajar
lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , gsCodAge, ObjCMACAgenciaArea)
lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H", , gsCodAge, ObjCMACAgenciaArea)

End Sub
Private Sub txtBuscaPers_EmiteDatos()
lblPersNombre = txtBuscaPers.psDescripcion
lblPersDireccion = txtBuscaPers.sPersDireccion
fgDocs.Clear
fgDocs.FormaCabecera
fgDocs.Rows = 2
lsDocumento = ""
If txtBuscaPers <> "" Then
    lsDocumento = txtBuscaPers.sPersNroDoc
    Set fgDocs.Recordset = txtBuscaPers.rsDocPers
End If
fgDocs.RowHeight(-1) = 230
fgDocs.RowHeight(0) = 280
txtImporte.SetFocus
End Sub
Private Sub txtImporte_Change()
  TxtMontoPagar = Format(txtImporte * CCur(lblTpoCambioDia), "#,#0.00")
End Sub
Private Sub txtImporte_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CmdGuardar.SetFocus
End If
End Sub
