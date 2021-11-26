VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmGiroTarifarioMant 
   Caption         =   "Form3"
   ClientHeight    =   7155
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8790
   ControlBox      =   0   'False
   LinkTopic       =   "Form3"
   ScaleHeight     =   7155
   ScaleWidth      =   8790
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabGiros 
      Height          =   6495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Tarifario de giros"
      TabPicture(0)   =   "frmGiroTarifarioMant.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FETarifario"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdQuitaTar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdEditaTar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdNuevoTar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "Servicios Adicionales"
      TabPicture(1)   =   "frmGiroTarifarioMant.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "FEServiciosAdicionales"
      Tab(1).Control(1)=   "cmdEditarSA"
      Tab(1).Control(2)=   "cmdGuardarSA"
      Tab(1).Control(3)=   "cmdCancelarSA"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdCancelarSA 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   -72960
         TabIndex        =   22
         Top             =   6000
         Width           =   855
      End
      Begin VB.CommandButton cmdGuardarSA 
         Caption         =   "Guardar"
         Height          =   375
         Left            =   -73920
         TabIndex        =   21
         Top             =   6000
         Width           =   855
      End
      Begin VB.CommandButton cmdEditarSA 
         Caption         =   "Editar"
         Height          =   375
         Left            =   -74880
         TabIndex        =   20
         Top             =   6000
         Width           =   855
      End
      Begin SICMACT.FlexEdit FEServiciosAdicionales 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   19
         Top             =   360
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   9763
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Comisión-MN-ME-Codigo"
         EncabezadosAnchos=   "400-6000-900-900-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdNuevoTar 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   17
         Top             =   6000
         Width           =   855
      End
      Begin VB.CommandButton cmdEditaTar 
         Caption         =   "Editar"
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   6000
         Width           =   855
      End
      Begin VB.CommandButton cmdQuitaTar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   2040
         TabIndex        =   15
         Top             =   6000
         Width           =   855
      End
      Begin VB.Frame Frame2 
         Caption         =   "Comisiones"
         Height          =   1695
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   8295
         Begin VB.TextBox txtHasta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            TabIndex        =   24
            Top             =   840
            Width           =   1215
         End
         Begin VB.TextBox txtDesde 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   960
            TabIndex        =   23
            Top             =   840
            Width           =   1215
         End
         Begin VB.ListBox lsAgencias 
            Height          =   1185
            Left            =   4320
            Style           =   1  'Checkbox
            TabIndex        =   18
            Top             =   360
            Width           =   2595
         End
         Begin VB.ComboBox cboMoneda 
            Height          =   315
            ItemData        =   "frmGiroTarifarioMant.frx":0038
            Left            =   960
            List            =   "frmGiroTarifarioMant.frx":0042
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   480
            Width           =   1215
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            ItemData        =   "frmGiroTarifarioMant.frx":0056
            Left            =   960
            List            =   "frmGiroTarifarioMant.frx":0060
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   1140
            Width           =   1215
         End
         Begin VB.TextBox txtValor 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   2760
            TabIndex        =   6
            Top             =   1155
            Width           =   1215
         End
         Begin VB.CheckBox chktodos 
            Caption         =   "Todos"
            Height          =   255
            Left            =   4320
            TabIndex        =   5
            Top             =   120
            Width           =   975
         End
         Begin VB.CommandButton cmdAceptarTar 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   7080
            TabIndex        =   4
            Top             =   360
            Width           =   1095
         End
         Begin VB.CommandButton cmdCancelarTar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   7080
            TabIndex        =   3
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label Label10 
            Caption         =   "Moneda:"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   480
            Width           =   735
         End
         Begin VB.Label Label9 
            Caption         =   "Desde:"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label8 
            Caption         =   "Tipo:"
            Height          =   255
            Left            =   240
            TabIndex        =   11
            Top             =   1200
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Hasta"
            Height          =   255
            Left            =   2280
            TabIndex        =   10
            Top             =   840
            Width           =   615
         End
         Begin VB.Label Label6 
            Caption         =   "Valor"
            Height          =   255
            Left            =   2280
            TabIndex        =   9
            Top             =   1200
            Width           =   735
         End
      End
      Begin SICMACT.FlexEdit FETarifario 
         Height          =   3735
         Left            =   120
         TabIndex        =   14
         Top             =   2160
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   6588
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Moneda-Desde-Hasta-Tipo-Valor-Agencia-TarCod-Codigo"
         EncabezadosAnchos=   "400-1200-1200-1200-1200-1200-1200-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-6-X-X"
         ListaControles  =   "0-0-0-0-0-0-1-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R-R-C-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   7800
      TabIndex        =   0
      Top             =   6720
      Width           =   855
   End
End
Attribute VB_Name = "frmGiroTarifarioMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmGiroTarifarioMant
'** Descripción : Formulario para dar mantenimeinto al tarifaio de giros
'** Creación : RECO, 20140410 - ERS008-2014
'**********************************************************************************************

Option Explicit
Dim lsCadAge As String
Dim lsGiroTarCod As String
Dim lnOpeTpo As Integer

Public Sub Inicio(ByVal nTpoOpe As Integer, ByVal sTitulo As String)
    Me.Caption = "Giros - Tarifario - " & sTitulo
    If nTpoOpe = 1 Then
    
    Else
        Call HabilitarControles(False)
        cmdNuevoTar.Visible = False
        cmdEditaTar.Visible = False
        cmdQuitaTar.Visible = False
        cmdEditarSA.Visible = False
        cmdGuardarSA.Visible = False
        cmdCancelarSA.Visible = False
    End If
    Me.Show 1
End Sub

Public Sub HabilitarControles(ByVal bHabilita As Boolean)
    cboMoneda.Enabled = bHabilita
    cboTipo.Enabled = bHabilita
    txtDesde.Enabled = bHabilita
    txtHasta.Enabled = bHabilita
    txtValor.Enabled = bHabilita
    chktodos.Enabled = bHabilita
    lsAgencias.Enabled = bHabilita
    cmdAceptarTar.Enabled = bHabilita
    cmdCancelarTar.Enabled = bHabilita
End Sub

Public Sub LimpiarFormulario()
    cboMoneda.ListIndex = 0
    cboTipo.ListIndex = 0
    txtDesde.Text = "0.00"
    txtHasta.Text = "0.00"
    txtValor.Text = "0.00"
    chktodos.value = 0
    cmdAceptarTar.Enabled = True
    cmdCancelarTar.Enabled = True
    cmdNuevoTar.Enabled = True
    cmdEditaTar.Enabled = True
    Me.cmdEditarSA.Visible = True
    cmdQuitaTar.Enabled = True
    cmdGuardarSA.Visible = False
    cmdCancelarSA.Visible = False
    Call LimpiarListaAge
    Call CargarTarifario
    Call CargarTarifarioServiciosAdicionales
    lnOpeTpo = 1
End Sub

Public Sub LimpiarListaAge()
    Dim nIndex As Integer
    For nIndex = 0 To lsAgencias.ListCount - 1
        lsAgencias.Selected(nIndex) = False
    Next
End Sub

Public Sub SelecListaAgeTodos()
    Dim nIndex As Integer
    For nIndex = 0 To lsAgencias.ListCount - 1
        lsAgencias.Selected(nIndex) = True
    Next
End Sub


Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDesde.SetFocus
    End If
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtValor.SetFocus
    End If
End Sub

Private Sub chktodos_Click()
    If chktodos.value = 0 Then
        Call LimpiarListaAge
    Else
        Call SelecListaAgeTodos
    End If
End Sub

Public Sub CargarListaAgencia()

    Dim loCargaAg As COMDColocPig.DCOMColPFunciones
    Dim lrAgenc As ADODB.Recordset
    
    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
    Set lrAgenc = loCargaAg.dObtieneAgencias(True)
        
    Set loCargaAg = Nothing
    
    If lrAgenc Is Nothing Then
        MsgBox " No se encuentran las Agencias ", vbInformation, " Aviso "
    Else
        Me.lsAgencias.Clear
        With lrAgenc
            Do While Not .EOF
                lsAgencias.AddItem !cAgeCod & " " & Trim(!cAgeDescripcion)
                If !cAgeCod = gsCodAge Then
                    lsAgencias.Selected(lsAgencias.ListCount - 1) = True
                End If
                .MoveNext
            Loop
        End With
    End If
End Sub

Private Sub cmdAceptarTar_Click()
    Dim oServ As New COMNCaptaServicios.NCOMCaptaServicios
    Dim nmoneda As Integer
    Dim nTipo As Integer
    Dim nIndex As Integer
    Dim sCodigo As String
    Dim nCanAge As Integer
    Dim nIndexAge As Integer
    
    nIndexAge = 1
    If Me.cboMoneda.ListIndex = 0 Then
        nmoneda = 1
    Else
        nmoneda = 2
    End If
    
    If Me.cboTipo.ListIndex = 0 Then
        nTipo = 1
    Else
        nTipo = 2
    End If
    
    nCanAge = RecuperaListaAgencias
    
    If nCanAge = 0 Then
        MsgBox "Debe seleccionar por lo menos una agencia", vbCritical, "Aviso"
        Exit Sub
    End If
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios

    If MsgBox("¿Está seguro que desea guardar?.", vbYesNo, "Aviso") = vbYes Then
        If lnOpeTpo = 2 Then
            sCodigo = lsGiroTarCod
            oServ.ActualizaTarifarioGiro sCodigo
            oServ.EliminaAgeTarifarioGiro sCodigo
        Else
            sCodigo = oServ.RecuperaUltimoCodigoTarifarioGiro
        End If
        Call oServ.RegistraTarifarioGiros(sCodigo, Format(gdFecSis, "yyyy/MM/dd"), nmoneda, val(Me.txtDesde.Text), val(Me.txtHasta.Text), nTipo, val(Me.txtValor.Text), gsCodUser, 1)
        For nIndex = 1 To nCanAge
            Dim sAgeCod  As String
            sAgeCod = Mid(lsCadAge, nIndexAge, 2)
            Call oServ.RegistraTarifarioGirosDet(sCodigo, sAgeCod, 1)
            nIndexAge = nIndexAge + 3
        Next
        Call oServ.ActualizaParamComisionValorMin 'NAGL 20181006 Según RFC1807260001
        Call CargarTarifario
    End If
    
    Me.FETarifario.Enabled = True
    Me.SSTabGiros.TabVisible(1) = True
    Me.cmdNuevoTar.Visible = True
    Me.cmdEditaTar.Visible = True
    Me.cmdQuitaTar.Visible = True
    Call HabilitarControles(False)
    Call LimpiarFormulario
End Sub

Private Sub cmdCancelarSA_Click()
    Call LimpiarFormulario
    Me.FETarifario.Enabled = True
    Me.SSTabGiros.TabVisible(1) = True
    Me.cmdNuevoTar.Visible = True
    Me.cmdEditaTar.Visible = True
    Me.cmdQuitaTar.Visible = True
    Call LimpiarFormulario
    Call HabilitarControles(False)
End Sub

Private Sub cmdCancelarTar_Click()
    Me.FETarifario.Enabled = True
    Me.SSTabGiros.TabVisible(1) = True
    Me.cmdNuevoTar.Visible = True
    Me.cmdEditaTar.Visible = True
    Me.cmdQuitaTar.Visible = True
    Call LimpiarFormulario
    Call HabilitarControles(False)
End Sub

Private Sub cmdEditarSA_Click()
    cmdGuardarSA.Visible = True
    cmdCancelarSA.Visible = True
    Me.cmdEditarSA.Visible = False
    FEServiciosAdicionales.SetFocus
    FEServiciosAdicionales.Col = 2
    FEServiciosAdicionales.row = FEServiciosAdicionales.row
    SendKeys "{Enter}"
End Sub

Private Sub cmdEditaTar_Click()
    Me.FETarifario.Enabled = False
    Me.SSTabGiros.TabVisible(1) = False
    Me.cmdNuevoTar.Visible = False
    Me.cmdEditaTar.Visible = False
    Me.cmdQuitaTar.Visible = False
    Call HabilitarControles(True)
    Me.cboMoneda.SetFocus
    lnOpeTpo = 2 'Editar Dato
    Call RecuperaDatosTarifarioEditar
End Sub

Private Sub cmdGuardarSA_Click()
    Dim oServ As New COMNCaptaServicios.NCOMCaptaServicios
    Dim nIndex As Integer
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    If Me.FEServiciosAdicionales.TextMatrix(3, 3) > 0 Then
        MsgBox "No se puede modificar el valor en ME de Comisión por Emisión de Detlle de Operaciones", vbCritical, "Aviso"
        Me.FEServiciosAdicionales.TextMatrix(3, 3) = 0
        Exit Sub
    End If
    For nIndex = 1 To Me.FEServiciosAdicionales.rows - 1
        oServ.ActualizaRecuperaDatosGirosServAdic Me.FEServiciosAdicionales.TextMatrix(nIndex, 4), Me.FEServiciosAdicionales.TextMatrix(nIndex, 2), Me.FEServiciosAdicionales.TextMatrix(nIndex, 3)
    Next
    Call LimpiarFormulario
End Sub

Private Sub cmdNuevoTar_Click()
    Me.FETarifario.Enabled = False
    Me.SSTabGiros.TabVisible(1) = False
    Me.cmdNuevoTar.Visible = False
    Me.cmdEditaTar.Visible = False
    Me.cmdQuitaTar.Visible = False
    Call HabilitarControles(True)
    Me.cboMoneda.SetFocus
End Sub

Private Sub cmdQuitaTar_Click()
    Dim oServ As New COMNCaptaServicios.NCOMCaptaServicios
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Call oServ.ActualizaTarifarioGiro(Me.FETarifario.TextMatrix(Me.FETarifario.row, 8))
    Call oServ.ActualizaParamComisionValorMin 'NAGL 20181006 Según RFC1807260001
    Call CargarTarifario
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub FETarifario_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    frmListaAgenciaGiros.Inicio (Me.FETarifario.TextMatrix(Me.FETarifario.row, 8))
    Me.FETarifario.TextMatrix(Me.FETarifario.row, 6) = "VER"
End Sub

Private Sub Form_Load()
    Call CargarListaAgencia
    Call LimpiarFormulario
    Call HabilitarControles(False)
    'Call CargarTarifario
    'Call CargarTarifarioServiciosAdicionales
End Sub

Private Sub txtDesde_GotFocus()
    If txtDesde.Text = "0.00" Then
        txtDesde.Text = ""
    End If
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Me.txtHasta.SetFocus
        Case 8, 48 To 57
        Case 46
            If InStr(txtDesde.Text, ".") > 0 Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtDesde_LostFocus()
    If Len(txtDesde.Text) > 0 Then
        txtDesde.Text = Format(txtDesde.Text, "0.00")
    Else
        txtDesde.Text = "0.00"
    End If
End Sub

Private Sub txtHasta_GotFocus()
    If txtHasta.Text = "0.00" Then
        txtHasta.Text = ""
    End If
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Me.cboTipo.SetFocus
        Case 8, 48 To 57
        
        Case 46
            If InStr(txtHasta.Text, ".") > 0 Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
    End Select
End Sub

Private Sub txtHasta_LostFocus()
    If Len(txtHasta.Text) > 0 Then
        txtHasta.Text = Format(txtHasta.Text, "0.00")
    Else
        txtHasta.Text = "0.00"
    End If
End Sub

Private Sub txtValor_GotFocus()
    If txtValor.Text = "0.00" Then
        txtValor.Text = ""
    End If
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            Me.lsAgencias.SetFocus
        Case 8, 48 To 57
        
        Case 46
            If InStr(txtValor.Text, ".") > 0 Then
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
    End Select
    If Me.cboTipo.ListIndex = 1 Then
        If val(txtValor.Text) > 100 Then
            MsgBox "No se puede ingresar un valor mayor a 100, debido a que el tipo de valor es porcentual.", vbCritical, "Aviso"
            Me.txtValor.Text = "0.00"
            Me.txtValor.SetFocus
        End If
    End If
End Sub

Private Sub txtValor_LostFocus()
    If Len(txtValor.Text) > 0 Then
        txtValor.Text = Format(txtValor.Text, "0.00")
    Else
        txtValor.Text = "0.00"
    End If
End Sub

Public Sub CargarTarifario()
    Dim oServ As New COMNCaptaServicios.NCOMCaptaServicios
    Dim lrDatos As ADODB.Recordset
    Dim nIndex As Integer
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set lrDatos = New ADODB.Recordset
    
    Set lrDatos = oServ.RecuperaDatosTarifarioGiros
    
    If Not (lrDatos.EOF And lrDatos.BOF) Then
        Me.FETarifario.Clear
        Me.FETarifario.rows = 2
        Me.FETarifario.FormaCabecera
        For nIndex = 1 To lrDatos.RecordCount
            Me.FETarifario.AdicionaFila
            Me.FETarifario.TextMatrix(nIndex, 0) = nIndex
            Me.FETarifario.TextMatrix(nIndex, 1) = IIf(lrDatos!nmoneda = 1, "SOLES", "DOLARES")
            Me.FETarifario.TextMatrix(nIndex, 2) = Format(lrDatos!nDesde, gcFormView)
            Me.FETarifario.TextMatrix(nIndex, 3) = Format(lrDatos!nHasta, gcFormView)
            Me.FETarifario.TextMatrix(nIndex, 4) = IIf(lrDatos!nTipo = 1, "Fijo", "Porcentual")
            Me.FETarifario.TextMatrix(nIndex, 5) = Format(lrDatos!nValor, gcFormView)
            Me.FETarifario.TextMatrix(nIndex, 6) = "VER"
            Me.FETarifario.TextMatrix(nIndex, 7) = lrDatos!cGiroTarCod
            Me.FETarifario.TextMatrix(nIndex, 8) = lrDatos!cGiroTarCod
            lrDatos.MoveNext
        Next
    End If
    Set lrDatos = Nothing
End Sub

Public Sub CargarTarifarioServiciosAdicionales()
    Dim oServ As New COMNCaptaServicios.NCOMCaptaServicios
    Dim lrDatos As ADODB.Recordset
    Dim nIndex As Integer
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set lrDatos = New ADODB.Recordset
    
    Set lrDatos = oServ.RecuperaDatosGirosServAdic
    
    If Not (lrDatos.EOF And lrDatos.BOF) Then
        Me.FEServiciosAdicionales.Clear
        Me.FEServiciosAdicionales.rows = 2
        Me.FEServiciosAdicionales.FormaCabecera
        For nIndex = 1 To lrDatos.RecordCount
            Me.FEServiciosAdicionales.AdicionaFila
            Me.FEServiciosAdicionales.TextMatrix(nIndex, 0) = nIndex
            Me.FEServiciosAdicionales.TextMatrix(nIndex, 1) = lrDatos!cComision
            Me.FEServiciosAdicionales.TextMatrix(nIndex, 2) = Format(lrDatos!nMontoMN, gcFormView)
            Me.FEServiciosAdicionales.TextMatrix(nIndex, 3) = Format(lrDatos!nMontoME, gcFormView)
            Me.FEServiciosAdicionales.TextMatrix(nIndex, 4) = lrDatos!cGiroTarServAdicCod
            lrDatos.MoveNext
        Next
    End If
    Set lrDatos = Nothing
End Sub
Public Function RecuperaListaAgencias() As Integer
    Dim nIndex As Integer
    lsCadAge = ""
    RecuperaListaAgencias = 0
    For nIndex = 0 To Me.lsAgencias.ListCount - 1
        If Me.lsAgencias.Selected(nIndex) Then
            lsCadAge = lsCadAge & Left(Me.lsAgencias.List(nIndex), 2) & ","
            RecuperaListaAgencias = RecuperaListaAgencias + 1
        End If
    Next
    If lsCadAge = "" Then
        Exit Function
    End If
    lsCadAge = Mid(lsCadAge, 1, Len(lsCadAge) - 1)
End Function
Public Sub RecuperaDatosTarifarioEditar()
    Dim oServ As New COMNCaptaServicios.NCOMCaptaServicios
    Dim lrDatos As ADODB.Recordset
    Dim lrDatosDet As ADODB.Recordset
    
    Set oServ = New COMNCaptaServicios.NCOMCaptaServicios
    Set lrDatos = New ADODB.Recordset
    Set lrDatosDet = New ADODB.Recordset
    
    Set lrDatos = oServ.RecuperaDatosTarifarioEditar(Me.FETarifario.TextMatrix(Me.FETarifario.row, 8))
    Set lrDatosDet = oServ.RecuperaDatosTarifarioDetEditar(Me.FETarifario.TextMatrix(Me.FETarifario.row, 8))
    
    If Not (lrDatos.EOF And lrDatos.BOF) Then
        cboMoneda.ListIndex = IIf(lrDatos!nmoneda = 1, 0, 1)
        txtDesde.Text = lrDatos!nDesde
        cboTipo.ListIndex = IIf(lrDatos!nTipo = 1, 0, 1)
        txtHasta.Text = lrDatos!nHasta
        txtValor.Text = lrDatos!nValor
        lsGiroTarCod = lrDatos!cGiroTarCod
        If Not (lrDatosDet.EOF And lrDatosDet.BOF) Then
            Dim nIndex As Integer
            Dim nIndList As Integer
            For nIndex = 0 To lrDatosDet.RecordCount - 1
                For nIndList = 0 To Me.lsAgencias.ListCount - 1
                    Dim sAgeCod As String
                    Me.lsAgencias.ListIndex = nIndList
                    sAgeCod = Mid(Me.lsAgencias.Text, 1, 2)
                    If sAgeCod = lrDatosDet!cAgeCod Then
                        lsAgencias.Selected(nIndList) = True
                    End If
                Next
                lrDatosDet.MoveNext
            Next
        End If
    End If
    Set lrDatos = Nothing
    Set lrDatosDet = Nothing
End Sub

