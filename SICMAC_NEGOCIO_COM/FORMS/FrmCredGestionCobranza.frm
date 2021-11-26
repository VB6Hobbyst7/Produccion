VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form FrmCredGestionCobranza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Gestion de Cobranzas"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8445
   Icon            =   "FrmCredGestionCobranza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8445
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraCliente 
      Height          =   1365
      Left            =   120
      TabIndex        =   14
      Top             =   570
      Width           =   8265
      Begin VB.Label Label13 
         Caption         =   "Total Mora:"
         Height          =   195
         Left            =   5400
         TabIndex        =   27
         Top             =   1005
         Width           =   885
      End
      Begin VB.Label LblMora 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6360
         TabIndex        =   26
         Top             =   960
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Caption         =   "Total Dias Atraso :"
         Height          =   195
         Left            =   2090
         TabIndex        =   25
         Top             =   960
         Width           =   1365
      End
      Begin VB.Label LblDiasAtraso 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3480
         TabIndex        =   24
         Top             =   960
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label8 
         Caption         =   "Total Interes:"
         Height          =   195
         Left            =   5280
         TabIndex        =   23
         Top             =   645
         Width           =   1005
      End
      Begin VB.Label LblInteres 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6360
         TabIndex        =   22
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente"
         Height          =   195
         Left            =   90
         TabIndex        =   21
         Top             =   270
         Width           =   645
      End
      Begin VB.Label Label3 
         Caption         =   "Prestamo"
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Capital :"
         Height          =   195
         Left            =   2400
         TabIndex        =   19
         Top             =   645
         Width           =   1035
      End
      Begin VB.Label lblSaldoCapital 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3480
         TabIndex        =   18
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblCodPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   17
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblNomPers 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   2400
         TabIndex        =   16
         Top             =   240
         Width           =   5475
         WordWrap        =   -1  'True
      End
      Begin VB.Label lblMontoPrestamo 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
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
         TabIndex        =   15
         Top             =   600
         Width           =   1335
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
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
      Left            =   7140
      TabIndex        =   12
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
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
      Left            =   2820
      TabIndex        =   11
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdGrabar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Grabar"
      Enabled         =   0   'False
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
      Left            =   1500
      TabIndex        =   10
      Top             =   5880
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Enabled         =   0   'False
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
      Left            =   300
      TabIndex        =   9
      Top             =   5880
      Width           =   1095
   End
   Begin VB.Frame fraComentario 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   4230
      Width           =   8265
      Begin VB.TextBox txtComentario 
         Height          =   630
         Left            =   120
         MaxLength       =   235
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   840
         Width           =   8010
      End
      Begin MSMask.MaskEdBox mskFechaVencimiento 
         Height          =   330
         Left            =   3960
         TabIndex        =   4
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   582
         _Version        =   393216
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
      Begin MSMask.MaskEdBox mskFechaAviso 
         Height          =   315
         Left            =   1080
         TabIndex        =   5
         Top             =   240
         Width           =   1260
         _ExtentX        =   2223
         _ExtentY        =   556
         _Version        =   393216
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
      Begin VB.Label Label5 
         Caption         =   "Fecha Aviso:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   260
         Width           =   1125
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Vencimiento:"
         Height          =   195
         Left            =   2520
         TabIndex        =   7
         Top             =   255
         Width           =   1485
      End
      Begin VB.Label Label7 
         Caption         =   "Comentario:"
         Height          =   195
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   1125
      End
   End
   Begin VB.Frame fraHistoria 
      Caption         =   "Actuaciones Procesales"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2220
      Left            =   120
      TabIndex        =   0
      Top             =   1950
      Width           =   8265
      Begin SICMACT.FlexEdit grdHistoria 
         Height          =   1905
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   7965
         _ExtentX        =   14049
         _ExtentY        =   3360
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Usu-Fecha-Comentario-Fecha Aviso-Fecha Vencimiento"
         EncabezadosAnchos=   "350-500-2000-5000-2000-2000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-L-L"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin SICMACT.ActXCodCta AXCodCta 
      Height          =   465
      Left            =   120
      TabIndex        =   13
      Top             =   120
      Width           =   3705
      _ExtentX        =   6535
      _ExtentY        =   820
      Texto           =   "Crédito"
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "FrmCredGestionCobranza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then Call BuscaDatos(AXCodCta.NroCuenta)
End Sub

Private Sub BuscaDatos(ByVal psNroCredito As String)
Dim lbOk As Boolean
Dim lrValida As ADODB.Recordset
Dim oDCred As COMDCredito.DCOMCredActBD
Dim rs As New ADODB.Recordset
Dim rsP As New ADODB.Recordset
Dim oNCred As COMNCredito.NCOMCredito
Dim lrAct As New ADODB.Recordset
Dim lsComentario As String, lsFecha As String, lsUsuario As String
Dim lnItem As Integer
Dim oCredCalen As COMDCredito.DCOMCalendario
Dim oDCredR As COMDCredito.DCOMCredito
Set oDCredR = New COMDCredito.DCOMCredito


Dim intColumna As Integer
Dim lngcolor As Long

Dim lsmensaje As String
'On Error GoTo ControlError

    'Valida Contrato
    Set oDCredR = New COMDCredito.DCOMCredito
    Set rsP = New ADODB.Recordset
    Set rsP = oDCredR.RecuperaRelacPers(AXCodCta.NroCuenta)
        If Not (rsP.EOF And rsP.BOF) Then
              lblCodPers = rsP("cPersCod")
              lblNomPers = PstaNombre(rsP("cPersNombre"), True)
              Set oDCred = New COMDCredito.DCOMCredActBD
                Set rs = New ADODB.Recordset
                Set rs = oDCred.RecuperaColocacCred(AXCodCta.NroCuenta)
                  If Not (rs.EOF And rs.BOF) Then
                      LblDiasAtraso = rs("nDiasAtrasoAcum")
                  End If
                Set rs = Nothing
                
                Set rs = New ADODB.Recordset
                Set rs = oDCred.RecuperaColocaciones(AXCodCta.NroCuenta)
                  If Not (rs.EOF And rs.BOF) Then
                      lblMontoPrestamo = Format(rs("nMontoCol"), "#0.00")
                  End If
                Set rs = Nothing
              Set oDCred = Nothing
              
              Set oCredCalen = New COMDCredito.DCOMCalendario
                Set rs = New ADODB.Recordset
                    Set rs = oCredCalen.RecuperaCalendarioPagosPendiente(AXCodCta.NroCuenta)
                    LblMora.Caption = Format(rs!nIntMor, "#0.00")
                    LblInteres.Caption = Format(rs!nIntComp, "#0.00")
                    lblSaldoCapital.Caption = Format(rs!nSaldoCap, "#0.00")
                Set rs = Nothing
              Set oCredCalen = Nothing
            
            cmdNuevo.Enabled = True
            cmdGrabar.Enabled = False
            cmdCancelar.Enabled = True
            AXCodCta.Enabled = False
            
        End If
   Set rs = Nothing
   Set oDCredR = Nothing
            
        'Actuaciones Procesales
        Set oDCred = New COMDCredito.DCOMCredActBD
        Set lrAct = New ADODB.Recordset
        Set lrAct = oDCred.ObtenerGestionCobranza(AXCodCta.NroCuenta)
        grdHistoria.Clear
        grdHistoria.FormaCabecera
        grdHistoria.Rows = 2
        txtComentario.Text = ""
        If lrAct.EOF And lrAct.BOF Then
            MsgBox "Credito NO tiene Actuaciones Procesales registradas", vbInformation, "Aviso"
        Else
            lnItem = 0
            Do While Not lrAct.EOF
                lsComentario = Trim(lrAct!cComenta)
                lsFecha = Mid(lrAct!cMovNro, 7, 2) & "/" & Mid(lrAct!cMovNro, 5, 2) & "/" & Mid(lrAct!cMovNro, 1, 4)
                lsFecha = lsFecha & " " & Mid(lrAct!cMovNro, 9, 2) & ":" & Mid(lrAct!cMovNro, 11, 2) & ":" & Mid(lrAct!cMovNro, 13, 2)
                lsUsuario = Right(lrAct!cMovNro, 4)
                lnItem = lnItem + 1
                grdHistoria.AdicionaFila
                grdHistoria.TextMatrix(lnItem, 1) = lsUsuario
                grdHistoria.TextMatrix(lnItem, 2) = lsFecha
                grdHistoria.TextMatrix(lnItem, 3) = lsComentario
                grdHistoria.TextMatrix(lnItem, 4) = Format(lrAct!dFechaAviso, "dd/mm/yyyy")
                grdHistoria.TextMatrix(lnItem, 5) = Format(lrAct!dFechaVencimiento, "dd/mm/yyyy")
                 
                'primero se debe establecer la Fila: .row = intFila y el Color:             lngColor = vbBlue
                
                If CDate(lrAct!dFechaAviso) = CDate(gdFecSis) Then
                    grdHistoria.row = lnItem
                    lngcolor = RGB(248, 247, 199)
                    For intColumna = 1 To grdHistoria.Cols - 1
                        grdHistoria.Col = intColumna
                        grdHistoria.CellBackColor = lngcolor
                    Next
                End If
                
                lrAct.MoveNext
                
            Loop
        End If
        
        Set lrAct = Nothing
 

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Dim sCuenta As String
Dim nProducto As Producto
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        sCuenta = frmValTarCodAnt.Inicia(nProducto)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
    AXCodCta.CMAC = gsCodCMAC
    AXCodCta.Age = gsCodAge
    cmdCancelar.Enabled = True
    cmdGrabar.Enabled = False
    cmdNuevo.Enabled = True
End Sub

Private Sub grdHistoria_RowColChange()
    If cmdNuevo.Enabled = True Then
        Me.txtComentario.Text = grdHistoria.TextMatrix(grdHistoria.row, 3)
        Me.mskFechaAviso.Text = grdHistoria.TextMatrix(grdHistoria.row, 4)
        Me.mskFechaVencimiento.Text = grdHistoria.TextMatrix(grdHistoria.row, 5)
    End If
End Sub

Private Sub cmdCancelar_Click()
    Limpiar ' True
    cmdNuevo.Enabled = True
    cmdGrabar.Enabled = False
    AXCodCta.Enabled = True
    fraComentario.Enabled = False
End Sub

Private Sub Limpiar()
lblCodPers = ""
lblNomPers = ""
lblMontoPrestamo = ""
lblSaldoCapital = ""
LblDiasAtraso = ""
LblInteres = ""
LblMora = ""
grdHistoria.Clear
grdHistoria.FormaCabecera
grdHistoria.Rows = 2
txtComentario.Text = ""
mskFechaAviso.Text = "__/__/____"
mskFechaVencimiento.Text = "__/__/____"

AXCodCta.CMAC = gsCodCMAC
AXCodCta.Age = gsCodAge
AXCodCta.Prod = ""
AXCodCta.Cuenta = ""
AXCodCta.SetFocusCuenta

End Sub

Private Function fValidaData() As Boolean
Dim lbOk As Boolean
Dim psMensaje As String
lbOk = True

If Len(Trim(txtComentario.Text)) = 0 Then
    MsgBox "No ingreso el comentario", vbInformation, "Aviso"
    lbOk = False
    Exit Function
End If

psMensaje = ValidaFecha(mskFechaAviso.Text)
If psMensaje <> "" Then
    MsgBox psMensaje, vbInformation, "Aviso"
    lbOk = False
    Exit Function
End If

psMensaje = ValidaFecha(mskFechaVencimiento.Text)
If psMensaje <> "" Then
    MsgBox psMensaje, vbInformation, "Aviso"
    lbOk = False
    Exit Function
End If

If CDate(mskFechaAviso.Text) > CDate(mskFechaVencimiento.Text) Then
    MsgBox "La Fecha de Aviso no puede se mayor a la fecha de Vencimiento", vbInformation, "Aviso"
    lbOk = False
    Exit Function
End If

fValidaData = lbOk
End Function

Private Sub mskFechaAviso_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFechaVencimiento.SetFocus
    End If
End Sub

Private Sub mskFechaVencimiento_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtComentario.SetFocus
    End If
End Sub

Private Sub TxtComentario_KeyPress(KeyAscii As Integer)
     KeyAscii = fgIntfMayusculas(KeyAscii)
     If KeyAscii = 13 Then
        cmdGrabar.SetFocus
     End If
End Sub

Private Sub cmdGrabar_Click()
Dim loContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim lsMovNro As String
Dim oDCred As COMDCredito.DCOMCredActBD
 
'On Error GoTo ControlError
' Valida Datos a Grabar
If fValidaData = False Then
    Exit Sub
End If

If MsgBox(" Grabar Registro de Actuacion Procesal ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        
    'Genera el Mov Nro
    Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    Set oDCred = New COMDCredito.DCOMCredActBD
        Call oDCred.dInsertGestionCobranza(AXCodCta.NroCuenta, lsMovNro, _
               Trim(txtComentario.Text), Me.mskFechaAviso.Text, Me.mskFechaVencimiento.Text)
    Set oDCred = Nothing
    
    
    BuscaDatos (AXCodCta.NroCuenta)
    cmdNuevo.Enabled = True
    cmdGrabar.Enabled = False
    mskFechaAviso.Text = "__/__/____"
    mskFechaVencimiento.Text = "__/__/____"
    Me.txtComentario.Text = ""
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub

Private Sub cmdNuevo_Click()
    fraComentario.Enabled = True
    txtComentario.Text = ""
    mskFechaAviso.Text = "__/__/____"
    mskFechaVencimiento.Text = "__/__/____"
    mskFechaAviso.SetFocus
    cmdNuevo.Enabled = False
    cmdGrabar.Enabled = True
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub



