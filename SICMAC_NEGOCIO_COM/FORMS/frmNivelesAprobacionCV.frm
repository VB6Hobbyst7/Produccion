VERSION 5.00
Begin VB.Form frmNivelesAprobacionCV 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Niveles de Aprobación C/V ME - Mantenimiento"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12240
   Icon            =   "frmNivelesAprobacionCV.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
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
      Left            =   11040
      TabIndex        =   20
      Top             =   6120
      Width           =   975
   End
   Begin VB.CommandButton cmdQuitar 
      Caption         =   "Quitar"
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
      Left            =   1920
      TabIndex        =   19
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "Mostrar Nivel"
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
      Left            =   3600
      TabIndex        =   18
      Top             =   6120
      Width           =   1575
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
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
      Left            =   240
      TabIndex        =   17
      Top             =   6120
      Width           =   1575
   End
   Begin VB.Frame Frame5 
      Height          =   3015
      Left            =   120
      TabIndex        =   21
      Top             =   3600
      Width           =   12015
      Begin SICMACT.FlexEdit feNivApr 
         Height          =   2295
         Left            =   120
         TabIndex        =   16
         Top             =   120
         Width           =   11775
         _ExtentX        =   20770
         _ExtentY        =   4048
         Cols0           =   11
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nivel-CodCargos-Cargos-Tipo-Firmas-Ag?-Desde $-Hasta $-TCC Más-TCV Menos"
         EncabezadosAnchos=   "300-1200-0-3000-0-800-800-1600-1600-1200-1200"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L-L-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Registro"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      Begin VB.Frame Frame4 
         Caption         =   "TCV Menos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   9120
         TabIndex        =   27
         Top             =   1200
         Width           =   2055
         Begin VB.TextBox txtTCVMas 
            Height          =   300
            Left            =   240
            TabIndex        =   11
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "TCC Más"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   6840
         TabIndex        =   26
         Top             =   1200
         Width           =   2175
         Begin VB.TextBox txtTCCMas 
            Height          =   300
            Left            =   480
            TabIndex        =   10
            Top             =   240
            Width           =   1215
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Rango $"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   6840
         TabIndex        =   9
         Top             =   240
         Width           =   4335
         Begin VB.TextBox txtRangoDe 
            Height          =   300
            Left            =   480
            TabIndex        =   7
            Top             =   240
            Width           =   1695
         End
         Begin VB.TextBox txtRangoA 
            Height          =   300
            Left            =   2520
            TabIndex        =   8
            Top             =   240
            Width           =   1695
         End
         Begin VB.CheckBox chkPatrimonio 
            Caption         =   "10% Patrimonio"
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
            Left            =   2400
            TabIndex        =   23
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label3 
            Caption         =   "De:"
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
            TabIndex        =   25
            Top             =   360
            Width           =   255
         End
         Begin VB.Label Label4 
            Caption         =   "A:"
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
            Left            =   2280
            TabIndex        =   24
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
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
         Left            =   8160
         TabIndex        =   15
         Top             =   2400
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
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
         Left            =   6840
         TabIndex        =   14
         Top             =   2400
         Width           =   1215
      End
      Begin VB.Frame Frame6 
         Caption         =   "Cargos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   6135
         Begin VB.ComboBox cboNivel 
            Height          =   315
            Left            =   960
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   720
            Width           =   4935
         End
         Begin VB.CheckBox chkAgencia 
            Caption         =   "Agencia"
            Height          =   375
            Left            =   2880
            TabIndex        =   13
            Top             =   2400
            Width           =   975
         End
         Begin VB.CommandButton cmdEliminarPorCargo 
            Caption         =   "-"
            Height          =   375
            Left            =   5520
            TabIndex        =   4
            Top             =   240
            Width           =   375
         End
         Begin VB.CommandButton cmdAgregarPorCargo 
            Caption         =   "+"
            Height          =   375
            Left            =   5040
            TabIndex        =   3
            Top             =   240
            Width           =   375
         End
         Begin VB.ComboBox cboTipoPorCargo 
            Height          =   315
            Left            =   2280
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   2535
         End
         Begin SICMACT.TxtBuscar txtBuscarCargo 
            Height          =   375
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   661
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
         Begin SICMACT.FlexEdit fePorCargo 
            Height          =   1095
            Left            =   120
            TabIndex        =   6
            Top             =   1200
            Width           =   5865
            _ExtentX        =   10345
            _ExtentY        =   1931
            Cols0           =   4
            HighLight       =   1
            EncabezadosNombres=   "-CargoCod-Cargo-Tipo"
            EncabezadosAnchos=   "300-0-4400-880"
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
            EncabezadosAlineacion=   "C-L-L-C"
            FormatosEdit    =   "0-1-1-0"
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            ColWidth0       =   300
            RowHeight0      =   300
         End
         Begin SICMACT.EditMoney txtNumFirmasPorCargo 
            Height          =   300
            Left            =   1440
            TabIndex        =   12
            Top             =   2400
            Width           =   495
            _ExtentX        =   873
            _ExtentY        =   529
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
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
         Begin VB.Label Label1 
            Caption         =   "Nivel"
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
            TabIndex        =   29
            Top             =   760
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "Nº de Vistos Necesarios"
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
            Left            =   120
            TabIndex        =   28
            Top             =   2400
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmNivelesAprobacionCV"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnPatrimonioEfectivo As Currency
Dim fbNuevo As Boolean
Dim fbActualiza As Boolean
Dim fnTipoReg As Integer
Dim fnRowNivApr As Integer
Public Sub InicioRegistroNiveles()
    Me.Caption = "Niveles de Aprobación C/V ME - Mantenimiento"
    fbNuevo = True
    fbActualiza = False
    fnTipoReg = 0
    CargaDatosNiveles
    feNivApr.TopRow = 1
    feNivApr.row = 1
    cmdCancelar.Cancel = False
    cmdCerrar.Cancel = False
    CargarCboTipoPorCargo
    Me.Show 1
End Sub
Private Sub CargaDatosNiveles()
    Dim oConst As COMDConstantes.DCOMConstantes
    Set oConst = New COMDConstantes.DCOMConstantes
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Set oCont = New COMNContabilidad.NCOMContFunciones

    lnPatrimonioEfectivo = 0
    Dim pdFecha As Date
    Dim psMes As String
    Dim psAnio As String
    pdFecha = DateAdd("d", -Day(gdFecSis), gdFecSis)
    pdFecha = DateAdd("d", 1, pdFecha)

    If Day(gdFecSis) >= 15 Then
        pdFecha = DateAdd("m", -1, pdFecha)
    Else
        pdFecha = DateAdd("m", -2, pdFecha)
    End If
    psMes = Month(pdFecha)
    psMes = Right("0" & Trim(psMes), 2)
    psAnio = Year(pdFecha)

    lnPatrimonioEfectivo = oCont.PatrimonioEfecAjustInfl(psAnio, psMes)
    Set oCont = Nothing


    txtBuscarCargo.lbUltimaInstancia = False
    txtBuscarCargo.psRaiz = "CARGOS DISPONIBLES PARA LOS NIVELES DE APROBACION"
    txtBuscarCargo.rs = oConst.ObtenerCargosArea

    Dim oNiv As COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer
    Set oNiv = New COMDCredito.DCOMNivelAprobacion

     Set rs = oNiv.RecuperaNivAprCV()
    Set oNiv = Nothing
    Call LimpiaFlex(feNivApr)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feNivApr.AdicionaFila
            lnFila = feNivApr.row
            feNivApr.TextMatrix(lnFila, 1) = rs!cNivelCod
            feNivApr.TextMatrix(lnFila, 2) = rs!cRHCargos
            feNivApr.TextMatrix(lnFila, 3) = rs!cRHCargoDescripcion
            feNivApr.TextMatrix(lnFila, 4) = rs!cTipoCargoDesc
            feNivApr.TextMatrix(lnFila, 5) = rs!nNroFirmas
            feNivApr.TextMatrix(lnFila, 6) = IIf(rs!bValidaAgencia = True, "Si", "No")
            feNivApr.TextMatrix(lnFila, 7) = rs!nMontoDesde
            feNivApr.TextMatrix(lnFila, 8) = rs!nMontoHasta
            feNivApr.TextMatrix(lnFila, 9) = rs!nTCCmas
            feNivApr.TextMatrix(lnFila, 10) = rs!nTCVmas
            rs.MoveNext
        Loop
        cmdEditar.Enabled = True
        cmdQuitar.Enabled = True
    Else
        cmdEditar.Enabled = False
        cmdQuitar.Enabled = False
        feNivApr.Enabled = False
    End If
    rs.Close
    Set rs = Nothing
    Call llenarCboNivel
End Sub

Private Sub chkPatrimonio_Click()
If chkPatrimonio.value = 1 Then
    txtRangoA.Text = Round(lnPatrimonioEfectivo * 0.1, 2)
Else
    txtRangoA = 0#
End If
End Sub

Private Sub cmdAgregarPorCargo_Click()
    Dim oConst As COMDConstantes.DCOMConstantes
    Dim rs As ADODB.Recordset
    Dim i As Integer

    If Trim(txtBuscarCargo) = "" Then
        MsgBox "Falta ingresar el cargo", vbInformation, "Aviso"
        txtBuscarCargo.SetFocus
        Exit Sub
    End If
    If Trim(cboTipoPorCargo.Text) = "" Then
        MsgBox "Falta seleccionar el tipo", vbInformation, "Aviso"
        cboTipoPorCargo.SetFocus
        Exit Sub
    End If
    If Trim(cboNivel.Text) = "" Then
        MsgBox "Falta selecionar el nivel", vbInformation, "Aviso"
        cboNivel.SetFocus
        Exit Sub
    End If

    For i = 1 To fePorCargo.Rows - 1
        If fePorCargo.TextMatrix(i, 0) <> "" Then
            If Trim(fePorCargo.TextMatrix(i, 1)) = Right(Trim(txtBuscarCargo), 6) Then
                MsgBox "El cargo ya fue ingresado", vbInformation, "Aviso"
                txtBuscarCargo.SetFocus
                Exit Sub
            End If
        End If
    Next i

    Set oConst = New COMDConstantes.DCOMConstantes
    Set rs = oConst.ObtenerCargosArea(Right(Trim(txtBuscarCargo), 6))

    If Not rs.EOF Then
        fePorCargo.AdicionaFila
        fePorCargo.TextMatrix(fePorCargo.row, 1) = Right(Trim(txtBuscarCargo), 6)
        fePorCargo.TextMatrix(fePorCargo.row, 2) = rs!Descripcion
        fePorCargo.TextMatrix(fePorCargo.row, 3) = IIf(Right(Trim(cboTipoPorCargo.Text), 1) = "N", "Nesesario", "Opcional")
        Call VerificaNumFirmas(False, txtNumFirmasPorCargo, fePorCargo)
        txtBuscarCargo = ""
        Me.cboTipoPorCargo.ListIndex = -1
    Else
        MsgBox "El codigo ingresado no existe", vbInformation, "Aviso"
        txtBuscarCargo.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
Call LimpiarControles
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub ActivarEditar()
    Dim oNiv As COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Call LimpiarControles
    Set oNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oNiv.RecuperaNivAprCV(feNivApr.TextMatrix(feNivApr.row, 1))
    If Not rs.EOF Then
         cboNivel.ListIndex = IndiceListaCombo(cboNivel, Trim(rs!cNivelCod))
         txtNumFirmasPorCargo.Text = rs!nNroFirmas
         chkAgencia.value = IIf(rs!bValidaAgencia = 1, 1, 0)
         txtRangoDe.Text = rs!nMontoDesde
         txtRangoA.Text = rs!nMontoHasta

         txtTCCMas.Text = rs!nTCCmas
         txtTCVMas.Text = rs!nTCVmas
         Do While Not rs.EOF
            fePorCargo.AdicionaFila
            fePorCargo.TextMatrix(fePorCargo.row, 1) = rs!cRHCargos
            fePorCargo.TextMatrix(fePorCargo.row, 2) = rs!cRHCargoDescripcion
            fePorCargo.TextMatrix(fePorCargo.row, 3) = IIf(Trim(rs!cTipoCargo) = "N", "Nesesario", "Opcional")
            rs.MoveNext
        Loop
    End If
    txtBuscarCargo = ""
    Me.cboTipoPorCargo.ListIndex = -1
    cmdGrabar.Caption = "Editar"
End Sub
Private Sub LimpiarControles()
    Call LimpiaFlex(fePorCargo)
    txtBuscarCargo.Text = ""
    cboTipoPorCargo.ListIndex = -1
    cboNivel.ListIndex = -1
    txtNumFirmasPorCargo.Text = ""
    chkAgencia.value = 0
    txtRangoDe.Text = ""
    txtRangoA.Text = ""
    chkPatrimonio.value = 0
    txtTCCMas.Text = ""
    txtTCVMas.Text = ""
    cmdGrabar.Caption = "Guardar"
End Sub

Private Sub cmdEditar_Click()
    Call ActivarEditar
End Sub

Private Sub cmdEliminarPorCargo_Click()
    If fePorCargo.TextMatrix(fePorCargo.row, 0) <> "" Then
        If MsgBox("¿Eliminar los datos de la fila " + CStr(fePorCargo.row) + " de la lista de Cargos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            fePorCargo.EliminaFila fePorCargo.row
            Call VerificaNumFirmas(False, txtNumFirmasPorCargo, fePorCargo)
        End If
    End If
End Sub
Private Sub CargarCboTipoPorCargo()
        cboTipoPorCargo.AddItem "Necesario " & Space(100) & "N"
        cboTipoPorCargo.AddItem "Opcional  " & Space(100) & "O"
End Sub

Private Function ValidaDatosNivApr() As Boolean
    ValidaDatosNivApr = False
    If Right(Trim(cboNivel.Text), 6) = "" Then
        MsgBox "Debe el nivel de aprobación", vbCritical
        ValidaDatosNivApr = False
        Exit Function
    End If

    Call VerificaNumFirmas(False, txtNumFirmasPorCargo, fePorCargo)
    If ValidaDatosGrid(fePorCargo, "Debe ingresar al menos un cargo", "Faltan datos en la lista de Cargos", 4) = False Then
        ValidaDatosNivApr = False
        Exit Function
    End If
    If Trim(txtRangoDe.Text) = "" Then
        MsgBox "Debe ingresar valor de rango de inicio", vbCritical
        ValidaDatosNivApr = False
        Exit Function
    End If
    If Trim(txtRangoA.Text) = "" Then
        MsgBox "Debe ingresar valor de rango Final", vbCritical
        ValidaDatosNivApr = False
        Exit Function
    End If
    If Trim(txtTCCMas.Text) = "" Then
        MsgBox "Debe ingresar valor de rango de TC de Compra (TCC)", vbCritical
        ValidaDatosNivApr = False
        Exit Function
    End If
    If Trim(txtTCVMas.Text) = "" Then
        MsgBox "Debe ingresar valor de rango de TC de Venta (TCV)", vbCritical
        ValidaDatosNivApr = False
        Exit Function
    End If

    ValidaDatosNivApr = True

End Function
Private Sub cmdGrabar_Click()

    If ValidaDatosNivApr Then
        Dim oNivApr As COMNCredito.NCOMNivelAprobacion
        Dim MatValores() As String, i As Integer
        fnRowNivApr = feNivApr.row
        ReDim MatValores(fePorCargo.Rows - 1, 2)
        For i = 1 To fePorCargo.Rows - 1
            MatValores(i - 1, 0) = fePorCargo.TextMatrix(i, 1)
            MatValores(i - 1, 1) = Left(Trim(fePorCargo.TextMatrix(i, 3)), 1)
        Next

        Set oNivApr = New COMNCredito.NCOMNivelAprobacion
        If MsgBox("¿Está seguro de actualizar los datos?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
        For i = 1 To fePorCargo.Rows - 1
            Call oNivApr.dInsertaNivAprCV(Right(Trim(cboNivel.Text), 6), fePorCargo.TextMatrix(i, 1), Left(Trim(fePorCargo.TextMatrix(i, 3)), 1), txtNumFirmasPorCargo.Text, IIf(chkAgencia.value = 1, 1, 0), txtRangoDe.Text, txtRangoA.Text, txtTCCMas.Text, txtTCVMas.Text)
        Next i
        MsgBox "Los datos se " & IIf(fbNuevo = True, "registraron", "actualizaron") & " correctamente", vbInformation, "Aviso"
        Call LimpiarControles
        CargaDatosNiveles
        feNivApr.TopRow = fnRowNivApr
        feNivApr.row = fnRowNivApr
        cmdGrabar.Caption = "Guardar"
    End If
End Sub

Private Sub cmdNuevo_Click()
    Call LimpiarControles
End Sub

Private Sub cmdQuitar_Click()
Dim oNiveles As COMDCredito.DCOMNivelAprobacion
Set oNiveles = New COMDCredito.DCOMNivelAprobacion
Call oNiveles.EliminarNivelesAprobacionCompraVenta(feNivApr.TextMatrix(feNivApr.row, 1), feNivApr.TextMatrix(feNivApr.row, 2))
MsgBox "datos se eliminaron correctamente ", vbCritical
CargaDatosNiveles
End Sub

Private Sub feNivApr_DblClick()
    If feNivApr.TextMatrix(feNivApr.row, feNivApr.Col) <> "" Then
        MuestraDatosGridNivApr
    End If
End Sub
Private Sub MuestraDatosGridNivApr()
    Dim oLista As COMDCredito.DCOMNivelAprobacion
    Dim rsDatos As ADODB.Recordset
    Dim MatTitulos() As String
    ReDim MatTitulos(1, 2)
    MatTitulos(0, 0) = feNivApr.TextMatrix(feNivApr.row, 4)
    MatTitulos(0, 1) = "Tipo"
    If feNivApr.Col = 1 Then
        Set oLista = New COMDCredito.DCOMNivelAprobacion
            Set rsDatos = oLista.RecuperaNivAprValoresCV(feNivApr.TextMatrix(feNivApr.row, 1))
        Set oLista = Nothing
        frmCredListaDatos.Inicio feNivApr.TextMatrix(feNivApr.row, 4), rsDatos, , 2, MatTitulos
    End If
End Sub
Private Sub txtNumFirmasPorCargo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 46 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtNumFirmasPorCargo_LostFocus()
    Call VerificaNumFirmas(True, txtNumFirmasPorCargo, fePorCargo)
End Sub
Private Sub VerificaNumFirmas(ByVal pbVerificaReg As Boolean, ByVal pNroFirmas As EditMoney, ByVal pFlex As FlexEdit)
    If pbVerificaReg = False Then
        Dim fnNes As Integer, fnOpc As Integer, i As Integer
        fnNes = 0
        fnOpc = 0
        For i = 1 To pFlex.Rows - 1
            If pFlex.TextMatrix(i, 0) <> "" Then
                If Left(Trim(pFlex.TextMatrix(i, 3)), 1) = "N" Then
                    fnNes = fnNes + 1
                Else
                    fnOpc = fnOpc + 1
                End If
            End If
        Next i
        If fnNes = 0 And fnOpc = 0 Then
            pNroFirmas.Text = 0
        ElseIf fnNes = 0 And fnOpc <> 0 Then
            If CInt(pNroFirmas.Text) > 0 Then
                pNroFirmas.Text = pNroFirmas.Text
            Else
                pNroFirmas.Text = 1
            End If
        ElseIf fnNes <> 0 Then
            If CInt(pNroFirmas.Text) < fnNes Then
                pNroFirmas.Text = fnNes
            End If
        End If
    Else
        If pNroFirmas.value > pFlex.Rows - 1 Then
            MsgBox "El nro de Firmas no puede ser mayor a la cantidad de registros de la lista", vbInformation, "Aviso"
        End If
    End If
    pNroFirmas.Text = CInt(pNroFirmas)
End Sub
Public Sub llenarCboNivel()
Dim oNiveles As COMDCredito.DCOMNivelAprobacion
Set oNiveles = New COMDCredito.DCOMNivelAprobacion
Dim oRs As ADODB.Recordset
Set oRs = New ADODB.Recordset
Set oRs = oNiveles.ObtenerNivelesAprobacionCompraVentaTabla("")
    If Not (oRs.BOF Or oRs.EOF) Then
        Do While Not oRs.EOF
        cboNivel.AddItem oRs!cNivelDes & Space(200) & oRs!cNivelCod
        oRs.MoveNext
        Loop
    End If
End Sub

