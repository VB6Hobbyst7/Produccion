VERSION 5.00
Begin VB.Form frmCapTasaInt 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5745
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11880
   Icon            =   "frmCapTasaInt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5745
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraTasa 
      Caption         =   " Buscar "
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
      Height          =   855
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   11655
      Begin VB.ComboBox cboTarifario 
         Height          =   315
         Left            =   7440
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   300
         Width           =   2655
      End
      Begin VB.ComboBox cboTpoPrograma 
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
         Left            =   3840
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2535
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   10320
         TabIndex        =   3
         Top             =   280
         Width           =   1035
      End
      Begin VB.ComboBox cboTipoTasa 
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
         Left            =   7920
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   960
         Width           =   1935
      End
      Begin VB.ComboBox cboMoneda 
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
         ItemData        =   "frmCapTasaInt.frx":030A
         Left            =   840
         List            =   "frmCapTasaInt.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   300
         Width           =   1695
      End
      Begin VB.Label lblTarifario 
         Caption         =   "Tarifario:"
         Height          =   375
         Left            =   6720
         TabIndex        =   24
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sub Producto:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   2760
         TabIndex        =   12
         Top             =   360
         Width           =   1020
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Tasa:"
         Height          =   195
         Left            =   7200
         TabIndex        =   9
         Top             =   960
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdHistorico 
      Caption         =   "Histórico de Cambios"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      Picture         =   "frmCapTasaInt.frx":030E
      TabIndex        =   19
      Top             =   5280
      Width           =   1875
   End
   Begin VB.CommandButton cmdNuevaBus 
      Caption         =   "&Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      Picture         =   "frmCapTasaInt.frx":0650
      TabIndex        =   11
      Top             =   5280
      Width           =   1035
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   5280
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8760
      TabIndex        =   5
      Top             =   5280
      Width           =   1035
   End
   Begin VB.Frame fraTarifa 
      Caption         =   " Tasas "
      Enabled         =   0   'False
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
      Height          =   4095
      Left            =   120
      TabIndex        =   6
      Top             =   960
      Width           =   11655
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7440
         TabIndex        =   23
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8520
         TabIndex        =   22
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5280
         TabIndex        =   18
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   4200
         TabIndex        =   17
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3120
         TabIndex        =   16
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   8520
         TabIndex        =   15
         Top             =   3600
         Visible         =   0   'False
         Width           =   1035
      End
      Begin VB.CheckBox chkTodosAgencia 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   1215
      End
      Begin VB.ListBox LstAgencias 
         Height          =   3210
         ItemData        =   "frmCapTasaInt.frx":0992
         Left            =   120
         List            =   "frmCapTasaInt.frx":0999
         Style           =   1  'Checkbox
         TabIndex        =   13
         Top             =   555
         Width           =   2895
      End
      Begin SICMACT.FlexEdit grdTasas 
         Height          =   2895
         Left            =   3120
         TabIndex        =   10
         Top             =   600
         Width           =   8175
         _ExtentX        =   11456
         _ExtentY        =   5106
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Monto Ini-Monto Fin-Plazo Ini-Plazo Fin-Ord?-Tasa Int-nTasaCod-Cambio-Activa-bEdit"
         EncabezadosAnchos=   "300-1200-1200-900-900-600-1000-0-0-0-0"
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
         ColumnasAEditar =   "X-1-2-3-4-5-6-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-4-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-R-R-R-C-R-L-C-C-C"
         FormatosEdit    =   "0-2-2-3-3-0-2-0-0-1-1"
         CantEntero      =   12
         CantDecimales   =   4
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label4 
         Caption         =   "Agencias Seleccionadas: "
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
         Left            =   3120
         TabIndex        =   21
         Top             =   240
         Width           =   2055
      End
      Begin VB.Label lblAgenciasSelec 
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
         Left            =   5280
         TabIndex        =   20
         Top             =   240
         Width           =   3615
      End
   End
End
Attribute VB_Name = "frmCapTasaInt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'JUEZ 20140220 Rediseño del formulario

Option Explicit
Dim nProducto As COMDConstantes.Producto
Dim nMoneda As COMDConstantes.Moneda
Dim nTipoTasa As COMDConstantes.CaptacTipoTasa
Dim nTpoPrograma As Integer
Dim bConsulta As Boolean
Dim sTitProd As String
Dim nNroReg As String
Dim nColEdita As Integer 'JUEZ 20140220
Dim bCheckTodosAge As Boolean, bCheckLista As Boolean 'JUEZ 20140220
Dim rsTasasAux As ADODB.Recordset 'JUEZ 20140220
Dim bPFPersJur As Boolean 'JUEZ 20140220
'By capi 21012009
Dim objPista As COMManejador.Pista

Dim sTipo As Integer 'JIPR20190109 ERS077-2018
Dim nTipoTarifario As Integer 'JIPR20190109 ERS077-2018


Private Function ValidaTasas() As Boolean
Dim i As Long, J As Long
Dim nMontoIni As Double, nMontoFin As Double
Dim nPlazoIni As Long, nPlazoFin As Long
Dim nMontoIniAux As Double, nMontoFinAux As Double
Dim nPlazoIniAux As Long, nPlazoFinAux As Long
Dim bOrdPag As String, bOrdPagAux As Boolean
Dim nTpoPrograma As Integer, nTpoProgramaAux As Integer
Dim nTasa As Double
For i = 1 To grdTasas.rows - 1
    nMontoIni = CDbl(grdTasas.TextMatrix(i, 1))
    nMontoFin = CDbl(grdTasas.TextMatrix(i, 2))
    nPlazoIni = CDbl(grdTasas.TextMatrix(i, 3))
    nPlazoFin = CDbl(grdTasas.TextMatrix(i, 4))
    nTasa = CDbl(grdTasas.TextMatrix(i, 6))
    bOrdPag = IIf(grdTasas.TextMatrix(i, 5) = "Si", True, False)
    
    If nTasa = 0 Then
        MsgBox "Monto de Tasa no válido, debe ser mayor a cero.", vbInformation, "Aviso"
        grdTasas.row = i
        grdTasas.Col = 6
        grdTasas.SetFocus
        Exit Function
    End If
    
    If nProducto = gCapAhorros Or nProducto = gCapCTS Then
        If nMontoIni = 0 And nMontoFin = 0 Then
            ValidaTasas = False
            MsgBox "Tasa tiene rangos de montos no válidos, deben ser mayor a cero.", vbInformation, "Aviso"
            grdTasas.row = i
            grdTasas.SetFocus
            Exit Function
        End If
    Else
        If nPlazoIni = 0 And nPlazoFin = 0 Then
            ValidaTasas = False
            MsgBox "Tasa tiene rangos de plazos no válidos, deben ser mayor a cero.", vbInformation, "Aviso"
            grdTasas.row = i
            grdTasas.SetFocus
            Exit Function
        End If
    End If
    For J = 1 To nNroReg - 1
        If J <> i Then
            nMontoIniAux = CDbl(grdTasas.TextMatrix(J, 1))
            nMontoFinAux = CDbl(grdTasas.TextMatrix(J, 2))
            nPlazoIniAux = CDbl(grdTasas.TextMatrix(J, 3))
            nPlazoFinAux = CDbl(grdTasas.TextMatrix(J, 4))
            bOrdPagAux = IIf(grdTasas.TextMatrix(J, 5) = "Si", True, False)
            'By Capi 20012008
            If grdTasas.TextMatrix(J, 8) = "." Then
            
            If nProducto = gCapAhorros Or nProducto = gCapCTS Then
                If ((nMontoIni >= nMontoIniAux And nMontoIni < nMontoFinAux) Or (nMontoFin <= nMontoFinAux And nMontoFin > nMontoIniAux)) _
                    And bOrdPag = bOrdPagAux Then
                    ValidaTasas = False
                    MsgBox "Tasa tiene rangos de montos, plazos, o la opción de Orden Pago No Válida no válidos.", vbInformation, "Aviso"
                    grdTasas.row = J
                    grdTasas.SetFocus
                    Exit Function
                End If
            Else
                If ((nMontoIni >= nMontoIniAux And nMontoIni < nMontoFinAux) Or (nMontoFin <= nMontoFinAux And nMontoFin > nMontoIniAux)) And _
                    ((nPlazoIni >= nPlazoIniAux And nPlazoIni <= nPlazoFinAux) Or (nPlazoFin <= nPlazoFinAux And nPlazoFin >= nPlazoIniAux)) Then
                    ValidaTasas = False
                    MsgBox "Tasa tiene rangos de montos, plazos, o la opción de Orden Pago No Válida no válidos.", vbInformation, "Aviso"
                    grdTasas.row = J
                    grdTasas.SetFocus
                    Exit Function
                End If
            End If
           End If
        End If
    Next J
Next i
ValidaTasas = True
End Function

Private Sub IniciaCombos(ByRef combo As ComboBox, ByVal nConstante As ConstanteCabecera)
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsGen As ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsGen = clsGen.GetConstante(nConstante, , "", " ") 'BRGO 20111228
    Set clsGen = Nothing

    Do While Not rsGen.EOF
        combo.AddItem rsGen("cDescripcion") & Space(100) & rsGen("nConsValor")
        rsGen.MoveNext
    Loop

    combo.ListIndex = 0
    rsGen.Close
    Set rsGen = Nothing
End Sub

'JIPR20190109 ERS077-2018
Private Sub IniciaComboTarifario(ByRef combo As ComboBox, ByVal nConstante As ConstanteCabecera)
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim rsGen As ADODB.Recordset
    Set clsGen = New COMDConstSistema.DCOMGeneral
    Set rsGen = clsGen.GetTarifario(nConstante)
    Set clsGen = Nothing

    Do While Not rsGen.EOF
        combo.AddItem rsGen("cCodTarifario") & Space(100) & rsGen("nIdTarifario")
        rsGen.MoveNext
    Loop

    combo.ListIndex = 0
    rsGen.Close
    Set rsGen = Nothing
End Sub

Public Sub inicia(Optional nProd As Producto = gCapAhorros, Optional bCons As Boolean = False, Optional ByVal pbPFPersJur As Boolean = False)
    
    bConsulta = bCons
    IniciaCombos cboMoneda, gMoneda
    IniciaCombos cboTipoTasa, gCaptacTipoTasa
    IniciaComboTarifario cboTarifario, sTipo 'JIPR20190109 ERS077-2018
    CargaAgencias
    cboMoneda.ListIndex = 0
    cboTipoTasa.ListIndex = 0
    nProducto = nProd
    bPFPersJur = pbPFPersJur 'JUEZ 20140220
     
    Select Case nProd
        Case gCapAhorros
            sTitProd = "Ahorros"
            IniciaCombos cboTpoPrograma, gCaptacSubProdAhorros
            'by capi 21092009
             gsOpeCod = gAhoMantTasaInteres
            cboTarifario.Visible = True 'JIPR20190109 ERS077-2018
            lblTarifario.Visible = True 'JIPR20190109 ERS077-2018
        Case gCapPlazoFijo
            sTitProd = "Plazo Fijo" & IIf(bPFPersJur, " PJ", "")
            grdTasas.ColWidth(5) = 0
            IniciaCombos cboTpoPrograma, gCaptacSubProdPlazoFijo
            'by capi 21092009
             gsOpeCod = gPFMantTasaInteres
            cboTarifario.Visible = False 'JIPR20190109 ERS077-2018
            lblTarifario.Visible = False 'JIPR20190109 ERS077-2018
            
            grdTasas.EncabezadosAnchos = "300-1300-1300-1100-1100-0-1000-0-0-0-0"
            grdTasas.ListaControles = "0-0-0-0-0-0-0-0-0-0-0"
        Case gCapCTS
            sTitProd = "CTS"
            grdTasas.ColWidth(5) = 0
            IniciaCombos cboTpoPrograma, gCaptacSubProdCTS
            'by capi 21092009
             gsOpeCod = gCTSMantTasaInteres
            cboTarifario.Visible = True 'JIPR20190109 ERS077-2018
            lblTarifario.Visible = True 'JIPR20190109 ERS077-2018
            
            grdTasas.EncabezadosAnchos = "300-1300-1300-1100-1100-0-1000-0-0-0-0"
            grdTasas.ListaControles = "0-0-0-0-0-0-0-0-0-0-0"
    End Select
    If bConsulta Then
        'cmdNuevo.Visible = False
       ' cmdEliminar.Visible = False
        'cmdGrabar.Visible = False
        grdTasas.lbEditarFlex = False
        cmdHistorico.Visible = False
        Me.Caption = "Captaciones - Tasas Interés - " & sTitProd & "  - Consulta"
    Else
        'cmdNuevo.Visible = True
        'cmdEliminar.Visible = True
        cmdGrabar.Visible = True
        Me.Caption = "Captaciones - Tasas Interés - " & sTitProd & " - Mantenimiento"
        'cmdNuevo.Enabled = False
        'cmdGrabar.Enabled = False
        'cmdEliminar.Enabled = False
        grdTasas.lbEditarFlex = True
    End If
    cmdImprimir.Enabled = False
    'cmdNuevaBus.Enabled = False
    cmdNuevaBus_Click
    Me.Show 1
End Sub

Private Sub cboMoneda_Click()
If cboMoneda.ListIndex = 0 Then
    nMoneda = gMonedaNacional
    grdTasas.BackColor = &HC0FFFF
Else
    nMoneda = gMonedaExtranjera
    grdTasas.BackColor = &HC0FFC0
End If
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cboTipoTasa.SetFocus
End If
End Sub

Private Sub cboTipoTasa_Click()
    nTipoTasa = CLng(Trim(Right(cboTipoTasa, 4)))
End Sub

Private Sub cboTipoTasa_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAplicar.SetFocus
End If
End Sub

Private Sub cboTpoPrograma_Click()
nTpoPrograma = CInt(Trim(Right(cboTpoPrograma, 4)))
'sTitProd = sTitProd & " - " & Trim(Left(cboTpoPrograma.Text, 20))
End Sub

'JIPR20190109 ERS077-2018
Private Sub cboTarifario_Click()
    nTipoTarifario = CLng(Trim(Right(cboTarifario, 4)))
End Sub

Private Sub CmdAceptar_Click()
    Dim i As Integer
    For i = 1 To 6
        If i <> 5 Then
            If grdTasas.TextMatrix(grdTasas.row, i) = "" Then
                MsgBox "De ingresar todos los datos", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    Next i
    
    If Not ValidaTasasNew Then Exit Sub
    
    HabilitaControlesModifGrid (True)
End Sub

Private Sub cmdAgregar_Click()
    Dim i As Integer
    Dim bSelect As Boolean
    bSelect = False
    For i = 0 To LstAgencias.ListCount - 1
        If LstAgencias.Selected(i) Then
            bSelect = True
            Exit For
        End If
    Next i
    
    If Not bSelect Then
        MsgBox "Debe tener seleccionada por lo menos una agencia", vbInformation, "Aviso"
        Exit Sub
    End If
    
    HabilitaControlesModifGrid (False)
    
    Set rsTasasAux = grdTasas.GetRsNew
    grdTasas.AdicionaFila
    nColEdita = grdTasas.row
    SendKeys "{Enter}"
End Sub

Private Sub cmdAplicar_Click()
'JUEZ 20140220 ******************
'    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
'    Dim rsTasa As ADODB.Recordset
'    Dim i As Integer
'    Dim L As ListItem
'    Dim sListaAgencias As String

    
'   Set rsTasa = New ADODB.Recordset

'    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    'Set rsTasa = clsDef.GetTarifario(nProducto, nMoneda, nTipoTasa, gsCodAge, nTpoPrograma)
'    Set clsDef = Nothing
'    nNroReg = 0
'    If Not (rsTasa.EOF And rsTasa.BOF) Then
'        Set grdTasas.Recordset = rsTasa
'        nNroReg = grdTasas.Rows
'        For i = 0 To LstAgencias.ListCount - 1
'            LstAgencias.Selected(i) = True
'        Next i
        'If Not bConsulta Then
            'cmdNuevo.Enabled = True
            'cmdEliminar.Enabled = True
            'cmdGrabar.Enabled = True
            'grdTasas.lbEditarFlex = True
        'End If
    'Else
        'If Not bConsulta Then
            'cmdNuevo.Enabled = True
            'cmdEliminar.Enabled = False
            'cmdGrabar.Enabled = True
            'grdTasas.lbEditarFlex = True
        'End If
'    End If
'grdTasas.FormateaColumnas

'cmdAplicar.Enabled = False
'cboMoneda.Enabled = False
'cboTipoTasa.Enabled = False
'cboTpoPrograma.Enabled = False
'cmdNuevaBus.Enabled = True
'cmdImprimir.Enabled = True
If Not bConsulta Then
    HabilitaControles (True)
Else
    HabilitaControlesConsulta (True)
End If
'END JUEZ ***********************
End Sub
'By Capi 15012008 comentado porque ahora solo se activara o desactivara
'Private Sub cmdEliminar_Click()
'Dim nFila As Long
'nFila = grdTasas.Row
'If grdTasas.TextMatrix(nFila, 7) = "" Then
'    grdTasas.EliminaFila nFila
'Else
'    MsgBox "No es posible eliminar una tasa ya registrada en el Tarifario", vbInformation, "Aviso"
'End If
'End Sub

Private Sub cmdCancelar_Click()
    HabilitaControlesModifGrid (True)
    Set grdTasas.Recordset = rsTasasAux
    Set rsTasasAux = Nothing
End Sub

Private Sub CmdGrabar_Click()
If Trim(grdTasas.TextMatrix(1, 1)) = "" Then
    MsgBox "Ingrese las Tasas", vbInformation, "Aviso"
    Exit Sub
End If
'If Not ValidaTasas() Then Exit Sub

If MsgBox("¿Desea grabar la información actualizada?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
    Exit Sub
End If

Dim clsTasa As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim i As Integer
Dim nCodTasa As Long
Dim nMontoIni As Double, nMontoFin As Double
Dim nPlazoIni As Long, nPlazoFin As Long
Dim bOrdPag As Boolean
'By Capi 15012008
Dim bActiva As Boolean
Dim oMov As COMDMov.DCOMMov
Dim x As Integer

Dim VCMovNro As String
'
Dim nValorTasa As Double
Dim nUltFila As Long

nUltFila = grdTasas.rows - 1

Set clsTasa = New COMNCaptaGenerales.NCOMCaptaDefinicion

For x = 0 To LstAgencias.ListCount - 1
    If LstAgencias.Selected(x) Then
        If bPFPersJur Then
            clsTasa.ActualizaTasaPF 0, 0, 0, 0, 0, bOrdPag, 0, "", False, nProducto, nTpoPrograma, nMoneda, Left(LstAgencias.List(x), 2)
        Else
            clsTasa.ActualizaTasa 0, 0, 0, 0, 0, bOrdPag, 0, "", False, nProducto, nTpoPrograma, nMoneda, Left(LstAgencias.List(x), 2), nTipoTarifario 'JIPR20190109 ERS077-2018 , nTipoTarifario
        End If
    End If
Next x

For i = 1 To grdTasas.rows - 1
'    If CDbl(grdTasas.TextMatrix(nUltFila, 6)) = 0 Then
'        MsgBox "No es posible agregar un nuevo registro si no completa los datos anteriores", vbInformation, "Aviso"
'        grdTasas.Col = 6
'        grdTasas.row = nUltFila
'        grdTasas.SetFocus
'        Exit Sub
'    Else
'        If grdTasas.TextMatrix(i, 10) <> "" Or grdTasas.TextMatrix(i, 8) <> "." Then
            nMontoIni = grdTasas.TextMatrix(i, 1)
            nMontoFin = grdTasas.TextMatrix(i, 2)
            nPlazoIni = grdTasas.TextMatrix(i, 3)
            nPlazoFin = grdTasas.TextMatrix(i, 4)
            nValorTasa = CDbl(grdTasas.TextMatrix(i, 6))
            bOrdPag = IIf(grdTasas.TextMatrix(i, 5) = ".", True, False)
            'By Capi 15012008
            bActiva = IIf(grdTasas.TextMatrix(i, 8) = ".", True, False)
            Set oMov = New COMDMov.DCOMMov
            VCMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            '
'            nValorTasa = CDbl(grdTasas.TextMatrix(i, 6))
'            If grdTasas.TextMatrix(i, 10) = "M" Then
'                nCodTasa = grdTasas.TextMatrix(i, 7)
'                'By Capi 15012008
'                clsTasa.ActualizaTasa nCodTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, nValorTasa, "", False
'                'If grdTasas.TextMatrix(i, 8) = "." Then
'                'clsTasa.ActualizaTasa nCodTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, nValorTasa, vcMovNro, True
'                clsTasa.NuevaTasa nProducto, nMoneda, nTipoTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, gsCodAge, nValorTasa, nTpoPrograma, VCMovNro, True, nCodTasa
'                'End If
'                '
'            ElseIf grdTasas.TextMatrix(i, 10) = "A" Then
'                nCodTasa = grdTasas.TextMatrix(i, 7)
'                'By Capi 15012008
'                clsTasa.ActualizaTasa nCodTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, nValorTasa, "", False
'                'If grdTasas.TextMatrix(i, 8) = "." Then
'                'clsTasa.ActualizaTasa nCodTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, nValorTasa, vcMovNro, True
'                If grdTasas.TextMatrix(i, 9) = "." Then
'                    clsTasa.NuevaTasa nProducto, nMoneda, nTipoTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, gsCodAge, nValorTasa, nTpoPrograma, VCMovNro, True, nCodTasa
'                Else
'                    clsTasa.NuevaTasa nProducto, nMoneda, nTipoTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, gsCodAge, nValorTasa, nTpoPrograma, VCMovNro, False, nCodTasa
'                End If
'                'End If
'
'            ElseIf grdTasas.TextMatrix(i, 10) = "N" Then
'                'By Capi 15012008 se agrego 2 nuevos parametros
'                clsTasa.NuevaTasa nProducto, nMoneda, nTipoTasa, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, gsCodAge, nValorTasa, nTpoPrograma, VCMovNro, True, nCodTasa
'
'            End If
            For x = 0 To LstAgencias.ListCount - 1
                If LstAgencias.Selected(x) Then
                    If bPFPersJur Then
                        clsTasa.NuevaTasaPF nProducto, nMoneda, 100, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, Left(LstAgencias.List(x), 2), nValorTasa, nTpoPrograma, VCMovNro, True, 0
                    Else
                        clsTasa.NuevaTasa nProducto, nMoneda, 100, nMontoIni, nMontoFin, nPlazoIni, nPlazoFin, bOrdPag, Left(LstAgencias.List(x), 2), nValorTasa, nTpoPrograma, VCMovNro, True, 0, nTipoTarifario 'JIPR20190109 ERS077-2018, nTipoTarifario
                    End If
                End If
            Next x
            
            'By Capi 21012009
            'If grdTasas.TextMatrix(i, 10) = "M" Or grdTasas.TextMatrix(i, 10) = "A" Or grdTasas.TextMatrix(i, 10) = "N" Then
            '    objPista.InsertarPista gsOpeCod, VCMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar
            'End If

        'End If
    'End If
Next i

objPista.InsertarPista gsOpeCod, VCMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar
MsgBox "Las tasas fueron actualizadas paras las agencias seleccionadas", vbInformation, "Aviso"

Set clsTasa = Nothing
cmdNuevaBus_Click
End Sub

Private Sub cmdHistorico_Click()
    Dim rs As ADODB.Recordset
    Dim sAgencias As String
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    
    Dim i As Integer
    Dim bSelect As Boolean
    bSelect = False
    For i = 0 To LstAgencias.ListCount - 1
        If LstAgencias.Selected(i) Then
            bSelect = True
            Exit For
        End If
    Next i
    
    If Not bSelect Then
        MsgBox "Debe tener seleccionada por lo menos una agencia", vbInformation, "Aviso"
        Exit Sub
    End If
    
    For i = 0 To LstAgencias.ListCount - 1
        If LstAgencias.Selected(i) Then
            sAgencias = Left(LstAgencias.List(i), 2)
            Exit For
        End If
    Next i
    
    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    If bPFPersJur Then
        Set rs = clsDef.GetTarifarioHistoricoPF(nProducto, nMoneda, sAgencias, nTpoPrograma)
    Else
        Set rs = clsDef.GetTarifarioHistorico(nProducto, nMoneda, sAgencias, nTpoPrograma, nTipoTarifario) 'JIPR20190109 ERS077-2018 , nTipoTarifario
    End If
    Set clsDef = Nothing
    If Not (rs.EOF And rs.BOF) Then
        frmCapTasaIntHist.inicia nMoneda, Trim(Mid(Trim(cboTpoPrograma.Text), 1, Len(Trim(cboTpoPrograma.Text)) - 3)), rs
    Else
        MsgBox "No se encontraron datos", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdImprimir_Click()
    Dim P As previo.clsprevio
    Dim sCad As String
    Dim rs As New ADODB.Recordset
    Dim LsCapImp As COMNCaptaGenerales.NCOMCaptaImpresion
    Dim i As Integer
    With rs
    'Crear RecordSet
        .Fields.Append "nMontoI", adCurrency
        .Fields.Append "nMontoF", adCurrency
        .Fields.Append "nPlazoI", adCurrency
        .Fields.Append "nPlazoF", adCurrency
        .Fields.Append "sOrdPago", adVarChar, 100
        .Fields.Append "sTasaI", adVarChar, 100
        .Fields.Append "sActiva", adVarChar, 100
        .Fields.Append "LogCambio", adVarChar, 100
        .Open
    'Llenar Recordset
        For i = 1 To grdTasas.rows - 1
            .AddNew
            .Fields("nMontoI") = CDbl(Me.grdTasas.TextMatrix(i, 1))
            .Fields("nMontoF") = CDbl(Me.grdTasas.TextMatrix(i, 2))
            .Fields("nPlazoI") = CDbl(Me.grdTasas.TextMatrix(i, 3))
            .Fields("nPlazoF") = CDbl(Me.grdTasas.TextMatrix(i, 4))
            .Fields("sOrdPago") = IIf(Me.grdTasas.TextMatrix(i, 5) = ".", "SI", "NO")
            .Fields("sTasaI") = Me.grdTasas.TextMatrix(i, 6)
            .Fields("sActiva") = IIf(Me.grdTasas.TextMatrix(i, 9) = ".", "SI", "NO")
            .Fields("LogCambio") = Me.grdTasas.TextMatrix(i, 8)
        Next i
    End With

    Set LsCapImp = New COMNCaptaGenerales.NCOMCaptaImpresion
        sCad = LsCapImp.ImprimirTasaInt(rs, sTitProd, gMonedaNacional, gsNomAge, gdFecSis, gsNomCmac, nMoneda)
    Set LsCapImp = Nothing
    
    Set P = New previo.clsprevio
        P.Show sCad, "TASAS DE INTERES", False, , gImpresora
    Set P = Nothing
End Sub

Private Sub cmdModificar_Click()
    If grdTasas.TextMatrix(grdTasas.row, 0) = "" Then
        MsgBox "Debe seleccionar al menos un registro", vbInformation, "Aviso"
        Exit Sub
    End If
    HabilitaControlesModifGrid (False)
    
    Set rsTasasAux = grdTasas.GetRsNew
    nColEdita = grdTasas.row
End Sub

Private Sub cmdNuevaBus_Click()
'CmdAplicar.Enabled = True
'CboMoneda.Enabled = True
'cboTipoTasa.Enabled = True
'cboTpoPrograma.Enabled = True
'cmdNuevaBus.Enabled = False
'CboMoneda.SetFocus
If Not bConsulta Then
    HabilitaControles (False)
    'cmdNuevo.Enabled = False
    'cmdEliminar.Enabled = False
Else
    HabilitaControlesConsulta (False)
End If
'cmdImprimir.Enabled = False
grdTasas.Clear
grdTasas.rows = 2
grdTasas.FormaCabecera
End Sub

Private Sub cmdNuevo_Click()
Dim nUltFila As Long
nUltFila = grdTasas.rows - 1
If grdTasas.TextMatrix(nUltFila, 1) <> "" Then
    If CDbl(grdTasas.TextMatrix(nUltFila, 6)) = 0 Then
        MsgBox "No es posible agregar un nuevo registro si no completa los datos anteriores", vbInformation, "Aviso"
        grdTasas.Col = 6
        grdTasas.row = nUltFila
        grdTasas.SetFocus
        Exit Sub
    End If
End If
grdTasas.AdicionaFila
nUltFila = grdTasas.rows - 1
grdTasas.lbEditarFlex = True
grdTasas.TextMatrix(nUltFila, 1) = "0.00"
grdTasas.TextMatrix(nUltFila, 2) = "0.00"
grdTasas.TextMatrix(nUltFila, 3) = "0"
grdTasas.TextMatrix(nUltFila, 4) = "0"
grdTasas.TextMatrix(nUltFila, 6) = "0.00"
grdTasas.SetFocus
cmdGrabar.Enabled = True
'cmdEliminar.Enabled = True
End Sub

Private Sub cmdQuitar_Click()
    If grdTasas.TextMatrix(grdTasas.row, 0) = "" Then
        MsgBox "Debe seleccionar al menos un registro", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(grdTasas.row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        grdTasas.EliminaFila grdTasas.row
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    'If nProducto = gCapCTS Or nProducto = gCapPlazoFijo Then
    '   grdTasas.ColWidth(10) = 0
    'End If
    'By Capi 20012009
    Set objPista = New COMManejador.Pista
   
    'End By

End Sub

Private Sub grdTasas_Click()
If grdTasas.row <> nColEdita Then
    grdTasas.lbEditarFlex = False
Else
    grdTasas.lbEditarFlex = True
End If
End Sub

Private Sub grdTasas_DblClick()
If grdTasas.row <> nColEdita Then
    grdTasas.lbEditarFlex = False
Else
    grdTasas.lbEditarFlex = True
End If
End Sub

Private Sub grdTasas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If grdTasas.row <> nColEdita Then
        grdTasas.lbEditarFlex = False
    Else
        grdTasas.lbEditarFlex = True
    End If
End If
End Sub

Private Sub grdTasas_OnCellChange(pnRow As Long, pnCol As Long)
'If grdTasas.TextMatrix(pnRow, 10) = "" Then grdTasas.TextMatrix(pnRow, 10) = "M"
'If pnCol = 6 Then
'    If grdTasas.TextMatrix(pnRow, pnCol) <> "" Then
'        If CDbl(grdTasas.TextMatrix(pnRow, pnCol)) < 0 Then
'            grdTasas.TextMatrix(pnRow, pnCol) = "0.00"
'        End If
'    End If
'End If
'cmdGrabar.Enabled = True
If pnRow <> nColEdita Then
    grdTasas.lbEditarFlex = False
Else
    grdTasas.lbEditarFlex = True
End If

End Sub

Private Sub grdTasas_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
'If grdTasas.TextMatrix(pnRow, 10) = "" Then grdTasas.TextMatrix(pnRow, 10) = "A"
'cmdGrabar.Enabled = True
End Sub

Private Sub grdTasas_OnRowAdd(pnRow As Long)
'grdTasas.TextMatrix(pnRow, 10) = "N"
'cmdGrabar.Enabled = True
End Sub

'JUEZ 20140220 **********************************************
Private Sub HabilitaControles(ByVal pbHabilita As Boolean)
    fraTasa.Enabled = Not pbHabilita
    fraTarifa.Enabled = pbHabilita
    chkTodosAgencia.Enabled = pbHabilita
    LstAgencias.Enabled = pbHabilita
    cmdAgregar.Visible = pbHabilita
    cmdModificar.Visible = pbHabilita
    cmdQuitar.Visible = pbHabilita
    cmdGrabar.Visible = pbHabilita
    cmdAceptar.Visible = Not pbHabilita
    cmdCancelar.Visible = Not pbHabilita
    cmdHistorico.Enabled = pbHabilita
    chkTodosAgencia.value = 0
    If pbHabilita = False Then
        Dim i As Integer
        For i = 0 To LstAgencias.ListCount - 1
            LstAgencias.Selected(i) = pbHabilita
        Next i
    End If
End Sub

Private Sub CargaAgencias()
    Dim oAge As COMDConstantes.DCOMAgencias
    Dim rsAgencias As ADODB.Recordset
    Set oAge = New COMDConstantes.DCOMAgencias
        Set rsAgencias = oAge.ObtieneAgencias()
    Set oAge = Nothing
    If rsAgencias Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        LstAgencias.Clear
        With rsAgencias
            Do While Not rsAgencias.EOF
                LstAgencias.AddItem rsAgencias!nConsValor & " " & Trim(rsAgencias!cConsDescripcion)
                rsAgencias.MoveNext
            Loop
        End With
    End If
End Sub

Private Sub chkTodosAgencia_Click()
    Dim i As Integer
    If Not bCheckLista Then
        bCheckTodosAge = True
        For i = 0 To LstAgencias.ListCount - 1
            LstAgencias.Selected(i) = IIf(chkTodosAgencia.value = 1, True, False)
        Next i
        bCheckTodosAge = False
    End If
End Sub

Private Sub lstAgencias_Click()
    Dim clsDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Dim rsTasa As ADODB.Recordset
    Dim sAgencia As String
    Dim K As Integer
    Dim sAgeSelect As String, nCantAgeSelect As Integer
    
    bCheckLista = True
    Call LimpiaFlex(grdTasas)

    For K = 0 To LstAgencias.ListCount - 1
        If LstAgencias.Selected(K) Then
            sAgencia = Left(LstAgencias.List(K), 2)
            sAgeSelect = Mid(LstAgencias.List(K), 4, Len(LstAgencias.List(K)) - 1)
            Exit For
        End If
    Next K
    
    Set rsTasa = New ADODB.Recordset

    Set clsDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
    If bPFPersJur Then
        Set rsTasa = clsDef.GetTarifarioPF(nProducto, nMoneda, nTipoTasa, sAgencia, nTpoPrograma)
    Else
        Set rsTasa = clsDef.GetTarifario(nProducto, nMoneda, nTipoTasa, sAgencia, nTpoPrograma, nTipoTarifario) 'JIPR20190109 ERS077-2018 , nTipoTarifario
    End If
    Set clsDef = Nothing
    nNroReg = 0
    If Not (rsTasa.EOF And rsTasa.BOF) Then
        Set grdTasas.Recordset = rsTasa
        nNroReg = grdTasas.rows
        For K = 0 To LstAgencias.ListCount - 1
            If LstAgencias.Selected(K) Then
                nCantAgeSelect = nCantAgeSelect + 1
            Else
                If Not bCheckTodosAge Then chkTodosAgencia.value = 0
            End If
        Next K
        lblAgenciasSelec.Caption = sAgeSelect
        nCantAgeSelect = nCantAgeSelect - 1
        If nCantAgeSelect > 0 Then
            lblAgenciasSelec.Caption = lblAgenciasSelec & " y Otras (" & nCantAgeSelect & ")"
        End If
    End If
    lblAgenciasSelec.Caption = IIf(sAgencia = "", "Ninguna", lblAgenciasSelec.Caption)
    bCheckLista = False
End Sub

Private Sub HabilitaControlesModifGrid(ByVal pbHabilita As Boolean)
    chkTodosAgencia.Enabled = pbHabilita
    LstAgencias.Enabled = pbHabilita
    cmdAgregar.Visible = pbHabilita
    cmdModificar.Visible = pbHabilita
    cmdQuitar.Visible = pbHabilita
    cmdGrabar.Visible = pbHabilita
    cmdAceptar.Visible = Not pbHabilita
    cmdCancelar.Visible = Not pbHabilita
    grdTasas.lbEditarFlex = Not pbHabilita
End Sub
Private Sub HabilitaControlesConsulta(ByVal pbHabilita As Boolean)
    fraTasa.Enabled = Not pbHabilita
    fraTarifa.Enabled = pbHabilita
    chkTodosAgencia.Enabled = pbHabilita
    LstAgencias.Enabled = pbHabilita
    grdTasas.lbEditarFlex = Not pbHabilita
    If pbHabilita = False Then
        Dim i As Integer
        For i = 0 To LstAgencias.ListCount - 1
            LstAgencias.Selected(i) = pbHabilita
        Next i
    End If
End Sub
Private Function ValidaTasasNew() As Boolean
Dim i As Long, J As Long
Dim nMontoIni As Double, nMontoFin As Double
Dim nPlazoIni As Long, nPlazoFin As Long
Dim nMontoIniAux As Double, nMontoFinAux As Double
Dim nPlazoIniAux As Long, nPlazoFinAux As Long
Dim nOrdPag As Integer, nOrdPagAux As Integer

For i = IIf(grdTasas.rows - 1 > 1, 2, 1) To grdTasas.rows - 1
    nMontoIni = CDbl(grdTasas.TextMatrix(i, 1))
    nMontoFin = CDbl(grdTasas.TextMatrix(i, 2))
    If CDbl(grdTasas.TextMatrix(i, 3)) > 999999999 Or CDbl(grdTasas.TextMatrix(i, 4)) > 999999999 Or _
        CDbl(IIf(grdTasas.rows - 1 > 1, grdTasas.TextMatrix(i - 1, 1), 0)) > 999999999 Or _
        CDbl(IIf(grdTasas.rows - 1 > 1, grdTasas.TextMatrix(i - 1, 4), 0)) > 999999999 Then
        MsgBox "Los plazos no estàn correctamente ingresados", vbInformation, "Aviso"
        ValidaTasasNew = False
        Exit Function
    End If
    nPlazoIni = CDbl(grdTasas.TextMatrix(i, 3))
    nPlazoFin = CDbl(grdTasas.TextMatrix(i, 4))
    nOrdPag = IIf(grdTasas.TextMatrix(i, 5) = ".", 1, 0)
    nMontoIniAux = CDbl(IIf(grdTasas.rows - 1 > 1, grdTasas.TextMatrix(i - 1, 1), 0))
    nMontoFinAux = CDbl(IIf(grdTasas.rows - 1 > 1, grdTasas.TextMatrix(i - 1, 2), 0))
    nPlazoIniAux = CDbl(IIf(grdTasas.rows - 1 > 1, grdTasas.TextMatrix(i - 1, 3), 0))
    nPlazoFinAux = CDbl(IIf(grdTasas.rows - 1 > 1, grdTasas.TextMatrix(i - 1, 4), 0))
    nOrdPagAux = IIf(grdTasas.TextMatrix(i - 1, 5) = ".", 1, 0)
    
    If nMontoFin < nMontoIni Then
        MsgBox "El Monto Final no puede ser menor al Monto Inicial en la fila " & i, vbInformation, "Aviso"
        ValidaTasasNew = False
        Exit Function
    End If
    If nPlazoFin < nPlazoIni Then
        MsgBox "El Plazo Final no puede ser menor al Plazo Inicial en la fila " & i, vbInformation, "Aviso"
        ValidaTasasNew = False
        Exit Function
    End If
'    If grdTasas.Rows - 1 > 1 And nOrdPag = nOrdPagAux Then
'        If nMontoIni <= nMontoFinAux Then
'            MsgBox "El Monto Inicial de la fila " & i & " no puede ser menor o igual al Monto Final de la fila " & i - 1, vbInformation, "Aviso"
'            ValidaTasasNew = False
'            Exit Function
'        End If
'        If nPlazoIni <= nPlazoFinAux Then
'            MsgBox "El Plazo Inicial de la fila " & i & " no puede ser menor o igual al Plazo Final de la fila " & i - 1, vbInformation, "Aviso"
'            ValidaTasasNew = False
'            Exit Function
'        End If
'    End If
Next i

ValidaTasasNew = True
End Function
'END JUEZ ***************************************************
