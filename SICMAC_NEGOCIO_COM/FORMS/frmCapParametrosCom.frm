VERSION 5.00
Begin VB.Form frmCapParametrosCom 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   Icon            =   "frmCapParametrosCom.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   1455
      Left            =   7320
      TabIndex        =   17
      Top             =   120
      Width           =   1335
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton cmdLimpiar 
         Caption         =   "Limpiar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   840
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "Eliminar"
      Height          =   375
      Left            =   1320
      TabIndex        =   7
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Editar"
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   7560
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Height          =   2895
      Left            =   120
      TabIndex        =   11
      Top             =   1560
      Width           =   8535
      Begin SICMACT.FlexEdit grdParam 
         Height          =   2475
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   4366
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   1
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Código-Comisión-Monto-Moneda-Tipo-Estado-Tag-nMoneda-nTipo"
         EncabezadosAnchos=   "0-1000-3000-1200-800-1200-700-0-0-0"
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
         ColumnasAEditar =   "X-X-X-3-X-X-6-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-4-0-0-0"
         EncabezadosAlineacion=   "C-C-L-R-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-2-0-0-0-0-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "#"
         SelectionMode   =   1
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos de Comisión "
      Height          =   1455
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7095
      Begin VB.ComboBox cboTipo 
         Height          =   315
         ItemData        =   "frmCapParametrosCom.frx":030A
         Left            =   5280
         List            =   "frmCapParametrosCom.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   840
         Width           =   1575
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmCapParametrosCom.frx":0386
         Left            =   3240
         List            =   "frmCapParametrosCom.frx":0390
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   840
         Width           =   1215
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   300
         Left            =   960
         TabIndex        =   1
         Top             =   840
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtNombreCom 
         Height          =   300
         Left            =   2280
         TabIndex        =   0
         Top             =   360
         Width           =   4575
      End
      Begin VB.Label Label4 
         Caption         =   "Tipo:"
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
         Left            =   4800
         TabIndex        =   16
         Top             =   900
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Moneda:"
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
         Left            =   2520
         TabIndex        =   15
         Top             =   885
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   14
         Top             =   885
         Width           =   735
      End
      Begin VB.Label lblCodCom 
         Alignment       =   2  'Center
         BackColor       =   &H8000000E&
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
         Height          =   300
         Left            =   960
         TabIndex        =   9
         Top             =   360
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Comisión:"
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
         Left            =   240
         TabIndex        =   13
         Top             =   405
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCapParametrosCom"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCapParametrosCom
'** Descripción : Formulario para administrar los gastos y comisiones de Ahorros según TI-ERS097-2013
'** Creación : JUEZ, 20130828 09:00:00 AM
'*****************************************************************************************************

Option Explicit

Dim clsCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim lnOperacion As Integer
Dim rs As ADODB.Recordset
Dim fbEdita As Boolean
Dim fnRowEdita As Integer
Dim fbBuscarCom As Boolean
Dim fsCodCom As String
Dim fsParProd As String 'JUEZ 20150930

'Public Sub Inicia(ByVal pnOperacion As Integer)
Public Sub Inicia(ByVal pnOperacion As Integer, ByVal psParProd As String) 'JUEZ 20150930
    fsParProd = psParProd 'JUEZ 20150930
    Me.Caption = "Comisiones y Gastos Diversos de " & IIf(fsParProd = "A", "Ahorros", "Créditos") 'JUEZ 20150930
    CargaParametros
    CargaTipos
    lnOperacion = pnOperacion
    fnRowEdita = 0
    fbBuscarCom = False
    fsCodCom = ""
    Select Case pnOperacion
        Case 1
            cmdEditar.Visible = False
            cmdEliminar.Visible = False
            lblCodCom.Caption = ObtieneCodigoComision
            fbEdita = False
        Case 2
            Frame3.Visible = False
            cmdEditar.Visible = False
            cmdEliminar.Visible = False
            Frame1.Enabled = False
            Frame1.Left = 720
            Frame2.Top = 120
            cmdCerrar.Top = 3120
            Me.Height = 4125
            fbEdita = False
        Case 3
            cmdLimpiar.Caption = "Cancelar"
            Frame1.Enabled = False
            Frame3.Visible = False
            fbEdita = True
    End Select
    Me.Show 1
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboTipo.SetFocus
    End If
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGuardar.SetFocus
    End If
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub CargaParametros(Optional ByVal pbConsultaDifAho As Boolean = False)
    Dim rsPar As ADODB.Recordset
    Dim lnFila As Integer
    Set clsCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    'Set rsPar = clsCapDef.GetParametrosComision(, pbConsultaDifAho)
    Set rsPar = clsCapDef.GetParametrosComision(, pbConsultaDifAho, fsParProd) 'JUEZ 20150930
    
    Call LimpiaFlex(grdParam)
    Do While Not rsPar.EOF
        grdParam.AdicionaFila
        lnFila = grdParam.row
        grdParam.TextMatrix(lnFila, 1) = rsPar!cParCod
        grdParam.TextMatrix(lnFila, 2) = rsPar!cParDesc
        grdParam.TextMatrix(lnFila, 3) = Format(rsPar!nParMonto, "#,##0.00")
        grdParam.TextMatrix(lnFila, 4) = rsPar!cParMoneda
        grdParam.TextMatrix(lnFila, 5) = rsPar!cParTipo
        grdParam.TextMatrix(lnFila, 6) = IIf(rsPar!bEstado, 1, 0)
        grdParam.TextMatrix(lnFila, 7) = rsPar!nParMoneda
        grdParam.TextMatrix(lnFila, 8) = rsPar!nParTipo
        rsPar.MoveNext
    Loop
    grdParam.TopRow = 1
    Set rsPar = Nothing
    Set clsCapDef = Nothing
End Sub

Private Sub CargaTipos()
    Dim rsConst As New ADODB.Recordset
    Dim clsGen As New COMDConstSistema.DCOMGeneral
    Set rsConst = clsGen.GetConstante(2036)
    Set clsGen = Nothing
    
    cboTipo.Clear
    While Not rsConst.EOF
        cboTipo.AddItem rsConst.Fields(0) & Space(100) & rsConst.Fields(1)
        rsConst.MoveNext
    Wend
End Sub

Private Function ValidaDatos() As Boolean
    ValidaDatos = False
    
    If Trim(txtNombreCom.Text) = "" Then
        MsgBox "Falta ingresar la descripción de la comisión", vbInformation, "Aviso"
        txtNombreCom.SetFocus
        Exit Function
    End If
    If txtMonto.Text <= 0 Then
        MsgBox "Favor de ingresar correctamente el monto de la comisión", vbInformation, "Aviso"
        txtMonto.SetFocus
        Exit Function
    End If
    If cboMoneda.Text = "" Then
        MsgBox "Falta elegir la moneda de la comisión", vbInformation, "Aviso"
        cboMoneda.SetFocus
        Exit Function
    End If
    If cboTipo.Text = "" Then
        MsgBox "Falta elegir el tipo de la comisión", vbInformation, "Aviso"
        cboTipo.SetFocus
        Exit Function
    End If
    
    ValidaDatos = True
End Function

Private Sub CmdEditar_Click()
    If grdParam.TextMatrix(grdParam.row, 0) <> "" Then
        fbEdita = True
        fnRowEdita = grdParam.row
        lblCodCom.Caption = grdParam.TextMatrix(grdParam.row, 1)
        txtNombreCom.Text = grdParam.TextMatrix(grdParam.row, 2)
        txtMonto.Text = grdParam.TextMatrix(grdParam.row, 3)
        cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, grdParam.TextMatrix(grdParam.row, 7))
        cboTipo.ListIndex = IndiceListaCombo(cboMoneda, grdParam.TextMatrix(grdParam.row, 8))
        Frame1.Enabled = True
        Frame3.Visible = True
        cmdEditar.Enabled = False
        cmdEliminar.Enabled = False
    End If
End Sub

Private Sub cmdEliminar_Click()
    If grdParam.TextMatrix(grdParam.row, 0) <> "" Then
        If MsgBox("¿Está seguro de eliminar el parámetro " + grdParam.TextMatrix(grdParam.row, 1) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Set clsCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
            'Call clsCapDef.ActualizaParametrosComision(3, grdParam.TextMatrix(grdParam.row, 1), "", 0, 0, 0)
            Call clsCapDef.ActualizaParametrosComision(3, grdParam.TextMatrix(grdParam.row, 1), "", 0, 0, 0, fsParProd) 'JUEZ 20150930
            Set clsCapDef = Nothing
            grdParam.TextMatrix(grdParam.row, 6) = 0
        End If
    End If
End Sub

Private Sub cmdGuardar_Click()
    If ValidaDatos Then
        If MsgBox("¿Está seguro de " & IIf(fbEdita, "actualizar", "registrar") & " los datos?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Set clsCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        'Call clsCapDef.ActualizaParametrosComision(IIf(fbEdita, 2, 1), lblCodCom.Caption, Trim(txtNombreCom.Text), txtMonto.Text, Trim(Right(cboMoneda.Text, 2)), Trim(Right(cboTipo.Text, 2)))
        Call clsCapDef.ActualizaParametrosComision(IIf(fbEdita, 2, 1), lblCodCom.Caption, Trim(txtNombreCom.Text), txtMonto.Text, Trim(Right(cboMoneda.Text, 2)), Trim(Right(cboTipo.Text, 2)), fsParProd) 'JUEZ 20150930
        Set clsCapDef = Nothing
        MsgBox "Los datos fueron " & IIf(fbEdita, "actualizados", "registrados") & " con éxito", vbInformation, "Aviso"
        
        If Not fbEdita Then grdParam.AdicionaFila
        If Not fbEdita Then grdParam.TextMatrix(grdParam.Rows - 1, 1) = lblCodCom.Caption
        grdParam.TextMatrix(IIf(fbEdita, fnRowEdita, grdParam.Rows - 1), 2) = Trim(txtNombreCom.Text)
        grdParam.TextMatrix(IIf(fbEdita, fnRowEdita, grdParam.Rows - 1), 3) = Format(Me.txtMonto.Text, "#,##0.00")
        grdParam.TextMatrix(IIf(fbEdita, fnRowEdita, grdParam.Rows - 1), 4) = Trim(Left(cboMoneda.Text, 20))
        grdParam.TextMatrix(IIf(fbEdita, fnRowEdita, grdParam.Rows - 1), 5) = Trim(Left(cboTipo.Text, 20))
        grdParam.TextMatrix(IIf(fbEdita, fnRowEdita, grdParam.Rows - 1), 6) = 1
        grdParam.TextMatrix(IIf(fbEdita, fnRowEdita, grdParam.Rows - 1), 7) = Trim(Right(cboMoneda.Text, 2))
        grdParam.TextMatrix(IIf(fbEdita, fnRowEdita, grdParam.Rows - 1), 8) = Trim(Right(cboTipo.Text, 2))
        grdParam.TopRow = grdParam.Rows - 1
        cmdLimpiar_Click
    End If
End Sub

Private Sub cmdLimpiar_Click()
    txtNombreCom.Text = ""
    txtMonto.Text = 0
    cboMoneda.ListIndex = -1
    cboTipo.ListIndex = -1
    If lnOperacion = 1 Then
        lblCodCom.Caption = ObtieneCodigoComision
    End If
    If lnOperacion = 3 Then
        cmdEditar.Enabled = True
        cmdEliminar.Enabled = True
        Frame1.Enabled = False
        Frame3.Visible = False
        fbEdita = False
        fnRowEdita = 0
    End If
End Sub

Private Function ObtieneCodigoComision() As String
    Set clsCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    'ObtieneCodigoComision = clsCapDef.ObtenerCodigoParametroComision
    ObtieneCodigoComision = clsCapDef.ObtenerCodigoParametroComision(fsParProd) 'JUEZ 20150930
End Function

Private Sub grdParam_DblClick()
    If fbBuscarCom Then
        fsCodCom = grdParam.TextMatrix(grdParam.row, 1)
        Unload Me
    End If
End Sub

Private Sub grdParam_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If fbBuscarCom Then
            fsCodCom = grdParam.TextMatrix(grdParam.row, 1)
            Unload Me
        End If
    End If
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cboMoneda.SetFocus
    End If
End Sub

Private Sub txtNombreCom_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        txtMonto.SetFocus
    End If
End Sub

Private Sub txtNombreCom_LostFocus()
    txtNombreCom.Text = UCase(txtNombreCom.Text)
End Sub

'Public Function BuscarComision()
Public Function BuscarComision(ByVal psParProd As String) 'JUEZ 20150930
    fsParProd = psParProd 'JUEZ 20150930
    Me.Caption = "Concepto de Comisiones" 'JUEZ 20150930
    fbBuscarCom = True
    fsCodCom = ""
    Frame1.Visible = False
    Frame3.Visible = False
    cmdEditar.Visible = False
    cmdEliminar.Visible = False
    Frame1.Enabled = False
    Frame1.Left = 720
    Frame2.Top = 120
    cmdCerrar.Top = 3120
    grdParam.Width = 7900
    grdParam.EncabezadosAnchos = "0-1000-6500-0-0-0-0-0-0-0"
    grdParam.ListaControles = "0-0-0-0-0-0-0-0-0-0"
    Frame2.Width = 8175
    Me.Width = 8500
    Me.Height = 3600
    fbEdita = False
    CargaParametros True
    grdParam.TabIndex = 0
    grdParam.TopRow = 1
    Me.Show 1
    BuscarComision = fsCodCom
End Function
