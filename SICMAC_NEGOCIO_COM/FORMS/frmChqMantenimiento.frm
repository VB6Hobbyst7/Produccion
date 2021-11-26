VERSION 5.00
Begin VB.Form frmChqMantenimiento 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9915
   Icon            =   "frmChqMantenimiento.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   9915
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   960
      Left            =   6570
      TabIndex        =   11
      Top             =   135
      Width           =   3255
      Begin VB.TextBox txtGlosa 
         Height          =   645
         Left            =   135
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   12
         Top             =   225
         Width           =   3030
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   350
      Left            =   5535
      TabIndex        =   10
      Top             =   480
      Width           =   915
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   350
      Left            =   60
      TabIndex        =   8
      Top             =   4860
      Width           =   915
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   350
      Left            =   8940
      TabIndex        =   7
      Top             =   4860
      Width           =   915
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   350
      Left            =   7620
      TabIndex        =   6
      Top             =   4860
      Width           =   1275
   End
   Begin VB.Frame fraCheque 
      Caption         =   "Datos del Cheque"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3615
      Left            =   60
      TabIndex        =   5
      Top             =   1140
      Width           =   9795
      Begin SICMACT.FlexEdit grdCheque 
         Height          =   3300
         Left            =   60
         TabIndex        =   9
         Top             =   240
         Width           =   9675
         _ExtentX        =   17066
         _ExtentY        =   5821
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Cheque-Cuenta-Banco-Estado-Fecha Reg.-Fecha Valor-Mon-Monto-Flag"
         EncabezadosAnchos=   "350-1200-1200-2000-1000-1000-1000-350-1000-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-6-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-2-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-L-C-C-C-R-L"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-2-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraBuscar 
      Caption         =   "Buscar..."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1035
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5370
      Begin VB.TextBox txtNumCheque 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2100
         MaxLength       =   20
         TabIndex        =   4
         Top             =   420
         Width           =   2235
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   1710
         TabIndex        =   3
         Top             =   405
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.OptionButton optBuscar 
         Caption         =   "Por N° Cheque"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   660
         Width           =   1515
      End
      Begin VB.OptionButton optBuscar 
         Caption         =   "Por Cuenta"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Value           =   -1  'True
         Width           =   1515
      End
   End
End
Attribute VB_Name = "frmChqMantenimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nOperacion As CaptacOperacion

Public Function ValidaDatos() As Boolean
If Trim(txtGlosa) = "" Then
    MsgBox "Debe ingresar la glosa del movimiento", vbInformation, "Aviso"
    ValidaDatos = False
    txtGlosa.SetFocus
    Exit Function
End If
Dim nFila As Long
If nOperacion = gChqOpeModFechaValor Then
    Dim dNuevaFechaValor As Date
    Dim dFechaRegistro As Date
    
    nFila = grdCheque.Row
    dFechaRegistro = CDate(grdCheque.TextMatrix(nFila, 5))
    dNuevaFechaValor = CDate(grdCheque.TextMatrix(nFila, 6))
    If dNuevaFechaValor <= dFechaRegistro Then
        MsgBox "Fecha de Valorización no puede ser menor o igual que la fecha de valorización", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
ElseIf nOperacion = gChqOpeValorInmediata Then
    Dim nEstado As ChequeEstado
    
    nFila = grdCheque.Row
    nEstado = CLng(Right(Trim(grdCheque.TextMatrix(nFila, 4)), 2))
    If nEstado <> gChqEstEnValorizacion Then
        MsgBox "El cheque no se encuentra en estado de valorización", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
End If
ValidaDatos = True
End Function

Public Sub Inicia(ByVal nOpe As CaptacOperacion, ByVal sDescOperacion As String)
nOperacion = nOpe
Me.Caption = "Cheque - " & sDescOperacion
If nOperacion = gChqOpeConsultaEstado Then
    cmdGrabar.Visible = False
    grdCheque.lbEditarFlex = False
ElseIf nOperacion = gChqOpeModFechaValor Then
    cmdGrabar.Caption = "&Modificar"
    grdCheque.lbEditarFlex = True
ElseIf nOperacion = gChqOpeValorInmediata Then
    cmdGrabar.Caption = "&Valorizar"
    grdCheque.lbEditarFlex = False
ElseIf nOperacion = gChqOpeExtRegistro Then
    cmdGrabar.Caption = "&Extornar"
    grdCheque.lbEditarFlex = False
    txtCuenta.Visible = False
    optBuscar(0).Visible = False
ElseIf nOperacion = gChqOpeExtValorInmediata Then
    cmdGrabar.Caption = "&Extornar"
    grdCheque.lbEditarFlex = False
End If
Me.Show 1
End Sub

Private Sub cmdBuscar_Click()
Dim sDato As String
Dim oCap As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsChq As New ADODB.Recordset

If nOperacion = gChqOpeExtRegistro Then
    sDato = Trim(txtNumCheque.Text)
    If sDato = "" Then
        MsgBox "Número de Cheque No Válido", vbInformation, "Aviso"
        txtNumCheque.SetFocus
        Exit Sub
    End If
    Set oCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    'Set grdCheque.Recordset = oCap.GetChequesRegistrados(sDato)
    Set rsChq = oCap.GetChequesRegistrados(sDato)
ElseIf nOperacion = gChqOpeExtValorInmediata Then
    If optBuscar(0).value = True Then
        sDato = txtCuenta.NroCuenta
        If Len(sDato) <> 18 Then
            MsgBox "Número de Cuenta Incompleto", vbInformation, "Aviso"
            txtCuenta.SetFocus
            Exit Sub
        End If
        Set oCap = New COMNCaptaGenerales.NCOMCaptaGenerales
       
        Set rsChq = oCap.GetChequesValorizadosInmediato(gdFecSis, sDato)
    ElseIf optBuscar(1).value = True Then
        sDato = Trim(txtNumCheque.Text)
        If sDato = "" Then
            MsgBox "Número de Cheque No Válido", vbInformation, "Aviso"
            txtNumCheque.SetFocus
            Exit Sub
        End If
        Set oCap = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsChq = oCap.GetChequesValorizadosInmediato(gdFecSis, , sDato)
    End If
Else
    If optBuscar(0).value = True Then
        sDato = txtCuenta.NroCuenta
        If Len(sDato) <> 18 Then
            MsgBox "Número de Cuenta Incompleto", vbInformation, "Aviso"
            txtCuenta.SetFocus
            Exit Sub
        End If
        Set oCap = New COMNCaptaGenerales.NCOMCaptaGenerales
        If nOperacion = gChqOpeModFechaValor Then
                Set rsChq = oCap.GetChequesOperaciones(sDato, , 1)
        ElseIf nOperacion = gChqOpeValorInmediata Then
                Set rsChq = oCap.GetChequesOperaciones(sDato, , 1)
        ElseIf nOperacion = gChqOpeConsultaEstado Then
                Set rsChq = oCap.GetChequesOperaciones(sDato)
        End If
                
    ElseIf optBuscar(1).value = True Then
        sDato = Trim(txtNumCheque.Text)
        If sDato = "" Then
            MsgBox "Número de Cheque No Válido", vbInformation, "Aviso"
            txtNumCheque.SetFocus
            Exit Sub
        End If
        Set oCap = New COMNCaptaGenerales.NCOMCaptaGenerales
        If nOperacion = gChqOpeModFechaValor Then
                Set rsChq = oCap.GetChequesOperaciones(, sDato, 1)
        ElseIf nOperacion = gChqOpeValorInmediata Then
                Set rsChq = oCap.GetChequesOperaciones(, sDato, 1)
        ElseIf nOperacion = gChqOpeConsultaEstado Then
                Set rsChq = oCap.GetChequesOperaciones(, sDato)
        End If
    End If
End If

If rsChq.EOF And rsChq.BOF Then
    grdCheque.Clear
    grdCheque.Rows = 2
    grdCheque.FormaCabecera
    MsgBox "NO se encontraron cheques con el criterio de búsqueda", vbInformation, "Aviso"
    Me.cmdCancelar.SetFocus
Else
    Set grdCheque.Recordset = rsChq
    grdCheque.FormateaColumnas
    txtGlosa.SetFocus
End If
Set oCap = Nothing
End Sub

Private Sub cmdCancelar_Click()
grdCheque.Clear
grdCheque.Rows = 2
grdCheque.FormaCabecera
txtCuenta.Visible = True
txtNumCheque.Visible = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Age = ""
txtCuenta.EnabledAge = True
txtCuenta.EnabledCMAC = False
txtCuenta.Cuenta = ""
txtCuenta.Prod = ""
txtGlosa = ""
txtNumCheque = ""
Me.optBuscar(0).SetFocus
End Sub

Private Sub cmdGrabar_Click()
If Not ValidaDatos Then Exit Sub

If MsgBox("¿Desea Grabar la Información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim oServ As COMNCaptaServicios.NCOMCaptaServicios
    Dim oMov As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String
    Dim sGlosa As String
    Dim sPerscod As String, sNroDoc As String, sIFTpo As String, sIFCta As String
    Dim nFila As Long, nmovnro As Long
    
    
    
    Set oMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oMov = Nothing
    sGlosa = Trim(txtGlosa.Text)
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    If nOperacion = gChqOpeModFechaValor Then
        Dim rsChq As New ADODB.Recordset
        Set rsChq = grdCheque.GetRsNew()
        oServ.ActualizaChequeFechaValorizacion sMovNro, sGlosa, rsChq
        cmdCancelar_Click
    ElseIf nOperacion = gChqOpeValorInmediata Then
        nFila = grdCheque.Row
        sPerscod = Right(Trim(grdCheque.TextMatrix(nFila, 3)), 13)
        sNroDoc = Trim(grdCheque.TextMatrix(nFila, 1))
        sIFTpo = Left(Right(Trim(grdCheque.TextMatrix(nFila, 3)), 15), 2)
        sIFCta = Trim(grdCheque.TextMatrix(nFila, 2))
            
        oServ.ValorizaChequeInmediato sMovNro, sGlosa, sNroDoc, sPerscod, sIFTpo, sIFCta
        
        grdCheque.EliminaFila nFila
        txtCuenta.CMAC = gsCodCMAC
        txtCuenta.Age = gsCodAge
'        txtCuenta.EnabledAge = True
        txtCuenta.EnabledCMAC = False
'        txtCuenta.Cuenta = ""
'        txtCuenta.Prod = ""
'        txtGlosa = ""
'        txtNumCheque = ""
       
        
    ElseIf nOperacion = gChqOpeExtRegistro Then
        nFila = grdCheque.Row
        sPerscod = Right(Trim(grdCheque.TextMatrix(nFila, 3)), 13)
        sNroDoc = Trim(grdCheque.TextMatrix(nFila, 1))
        sIFTpo = Left(Right(Trim(grdCheque.TextMatrix(nFila, 3)), 15), 2)
        nmovnro = grdCheque.TextMatrix(nFila, 9)
        oServ.ExtornaChequeRegistro sMovNro, sGlosa, sNroDoc, sPerscod, sIFTpo, nmovnro
        grdCheque.EliminaFila nFila
        txtCuenta.CMAC = gsCodCMAC
        txtCuenta.Age = gsCodAge
        txtCuenta.EnabledAge = False
        txtCuenta.EnabledCMAC = False
        txtCuenta.Cuenta = ""
        txtCuenta.Prod = ""
        txtGlosa = ""
        txtNumCheque = ""
    ElseIf nOperacion = gChqOpeExtValorInmediata Then
        nFila = grdCheque.Row
        sPerscod = Right(Trim(grdCheque.TextMatrix(nFila, 3)), 13)
        sNroDoc = Trim(grdCheque.TextMatrix(nFila, 1))
        sIFTpo = Left(Right(Trim(grdCheque.TextMatrix(nFila, 3)), 15), 2)
        nmovnro = grdCheque.TextMatrix(nFila, 9)
        oServ.ExtornaValorizaChequeInmediato sMovNro, sGlosa, sNroDoc, sPerscod, sIFTpo, nmovnro
        grdCheque.EliminaFila nFila
        txtCuenta.CMAC = gsCodCMAC
        txtCuenta.Age = gsCodAge
        txtCuenta.EnabledAge = True
        txtCuenta.EnabledCMAC = False
        txtCuenta.Cuenta = ""
        txtCuenta.Prod = ""
        txtGlosa = ""
        txtNumCheque = ""
    End If
    Set oServ = Nothing
    
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gCapAhorros, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
txtCuenta.Visible = True
txtNumCheque.Visible = False
txtCuenta.CMAC = gsCodCMAC
txtCuenta.EnabledCMAC = False
If nOperacion = gChqOpeExtRegistro Then
    txtCuenta.EnabledAge = False
Else
    txtCuenta.EnabledAge = True
End If
End Sub

Private Sub grdCheque_OnCellChange(pnRow As Long, pnCol As Long)
grdCheque.TextMatrix(pnRow, 9) = "M"
End Sub

Private Sub optBuscar_Click(Index As Integer)
Select Case Index
    Case 0
        txtCuenta.Prod = ""
        txtCuenta.Cuenta = ""
        txtCuenta.Visible = True
        txtNumCheque.Visible = False
    Case 1
        txtCuenta.Visible = False
        txtNumCheque.Visible = True
        txtNumCheque = ""
End Select
End Sub

Private Sub optBuscar_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
    Select Case Index
        Case 0
            txtCuenta.Prod = ""
            txtCuenta.Cuenta = ""
            txtCuenta.Visible = True
            txtNumCheque.Visible = False
            txtCuenta.SetFocusAge
        Case 1
            txtCuenta.Visible = False
            txtNumCheque.Visible = True
            txtNumCheque.SetFocus
            txtNumCheque = ""
    End Select
    End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar.SetFocus
End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Select Case nOperacion
        Case gChqOpeModFechaValor
            grdCheque.Col = 6
            grdCheque.SetFocus
        Case gChqOpeValorInmediata, gChqOpeExtRegistro, gChqOpeExtValorInmediata
            cmdGrabar.SetFocus
    End Select
End If
End Sub

Private Sub txtNumCheque_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdBuscar.SetFocus
End If
End Sub

