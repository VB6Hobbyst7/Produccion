VERSION 5.00
Begin VB.Form frmCredPolizas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Polizas"
   ClientHeight    =   5340
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9120
   Icon            =   "frmCredPolizas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5340
   ScaleWidth      =   9120
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkSumAsegurada 
      Caption         =   "Suma asegurada en Soles (S/.)"
      Height          =   255
      Left            =   120
      TabIndex        =   42
      Top             =   2640
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      Caption         =   "Tipo de Pago de la Poliza"
      Height          =   1215
      Left            =   120
      TabIndex        =   39
      Top             =   3480
      Width           =   8895
      Begin VB.CommandButton cmdRelaciones 
         Caption         =   "&Relaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7440
         TabIndex        =   43
         Top             =   720
         Width           =   1275
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Pago Adelantado"
         Height          =   255
         Left            =   120
         TabIndex        =   41
         Top             =   360
         Value           =   -1  'True
         Width           =   1575
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Pago en Cuotas"
         Height          =   375
         Left            =   120
         TabIndex        =   40
         Top             =   720
         Width           =   1695
      End
   End
   Begin VB.TextBox txtTc 
      Alignment       =   1  'Right Justify
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
      Height          =   345
      Left            =   5760
      TabIndex        =   36
      Text            =   "0.00"
      Top             =   1575
      Width           =   1395
   End
   Begin VB.TextBox txtNumAnio 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   3600
      TabIndex        =   34
      Text            =   "0"
      Top             =   3000
      Width           =   795
   End
   Begin VB.ComboBox cboMoneda 
      Height          =   315
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   32
      Top             =   2160
      Width           =   1410
   End
   Begin VB.CommandButton cmdCalculaPrima 
      Caption         =   "Calcular Prima"
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
      Left            =   7320
      TabIndex        =   31
      ToolTipText     =   "Calcula la Prima Neta usando una Garantía relacionada"
      Top             =   1580
      Width           =   1575
   End
   Begin VB.TextBox txtNumCertif 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   930
      TabIndex        =   29
      Top             =   2220
      Width           =   1395
   End
   Begin VB.TextBox txtPrimaNeta 
      Alignment       =   1  'Right Justify
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
      Height          =   345
      Left            =   1080
      TabIndex        =   27
      Text            =   "0.00"
      Top             =   1580
      Width           =   1395
   End
   Begin VB.TextBox txtSumaA 
      Alignment       =   1  'Right Justify
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4320
      TabIndex        =   26
      Text            =   "0.00"
      Top             =   2190
      Width           =   1395
   End
   Begin VB.CheckBox chkExterna 
      Caption         =   "Poliza Externa"
      Height          =   240
      Left            =   120
      TabIndex        =   24
      Top             =   3075
      Width           =   1440
   End
   Begin VB.CommandButton cmdexaminar 
      Caption         =   "E&xaminar"
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
      Left            =   7740
      TabIndex        =   22
      Top             =   75
      Width           =   1230
   End
   Begin VB.ComboBox cboTipo 
      Height          =   315
      Left            =   5880
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   3000
      Width           =   3090
   End
   Begin VB.TextBox txtMonto 
      Alignment       =   1  'Right Justify
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
      Height          =   345
      Left            =   3450
      TabIndex        =   11
      Text            =   "0.00"
      Top             =   1580
      Width           =   1395
   End
   Begin VB.CommandButton cmdBuscaCont 
      Caption         =   "..."
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
      Left            =   8640
      TabIndex        =   3
      Top             =   525
      Width           =   390
   End
   Begin VB.CommandButton cmdBuscaAseg 
      Caption         =   "..."
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
      Left            =   8640
      TabIndex        =   2
      Top             =   975
      Width           =   390
   End
   Begin VB.TextBox txtNumPoliza 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1065
      TabIndex        =   0
      Top             =   75
      Width           =   1440
   End
   Begin VB.Frame fracontrol 
      Height          =   585
      Left            =   75
      TabIndex        =   14
      Top             =   4680
      Width           =   8940
      Begin VB.CommandButton Command1 
         Caption         =   "Autorización>>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   5160
         TabIndex        =   38
         Top             =   165
         Width           =   1635
      End
      Begin VB.CommandButton cmdGarantias 
         Caption         =   "Ga&rantias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   3480
         TabIndex        =   23
         Top             =   165
         Width           =   1035
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   75
         TabIndex        =   21
         Top             =   165
         Width           =   930
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1005
         TabIndex        =   20
         Top             =   165
         Width           =   900
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   60
         TabIndex        =   19
         Top             =   165
         Width           =   930
      End
      Begin VB.CommandButton cmdsalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   7845
         TabIndex        =   18
         Top             =   165
         Width           =   975
      End
      Begin VB.CommandButton cmdImprimir 
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
         Left            =   2925
         Picture         =   "frmCredPolizas.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   17
         ToolTipText     =   "Imprimir Solicitud"
         Top             =   150
         Visible         =   0   'False
         Width           =   435
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eli&minar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1920
         TabIndex        =   16
         Top             =   165
         Width           =   915
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
         Height          =   345
         Left            =   6870
         TabIndex        =   15
         Top             =   165
         Width           =   975
      End
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "Total. T/C:"
      Height          =   195
      Left            =   4920
      TabIndex        =   37
      Top             =   1650
      Width           =   780
   End
   Begin VB.Label lblNumAnio 
      AutoSize        =   -1  'True
      Caption         =   "Nº Años :"
      Height          =   195
      Left            =   2880
      TabIndex        =   35
      Top             =   3030
      Width           =   675
   End
   Begin VB.Label lblMoneda 
      AutoSize        =   -1  'True
      Caption         =   "Moneda Póliza:"
      Height          =   195
      Left            =   6360
      TabIndex        =   33
      Top             =   2220
      Width           =   1095
   End
   Begin VB.Shape Shape1 
      DrawMode        =   16  'Merge Pen
      Height          =   615
      Left            =   120
      Top             =   1440
      Width           =   8895
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "Nº Certif. :"
      Height          =   195
      Left            =   120
      TabIndex        =   30
      Top             =   2220
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Prima Neta:"
      Height          =   195
      Left            =   195
      TabIndex        =   28
      Top             =   1650
      Width           =   825
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Suma Asegurada US$:"
      Height          =   195
      Left            =   2655
      TabIndex        =   25
      Top             =   2220
      Width           =   1620
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "Tipo:"
      Height          =   195
      Left            =   5370
      TabIndex        =   12
      Top             =   3030
      Width           =   360
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Prima Total:"
      Height          =   195
      Left            =   2595
      TabIndex        =   10
      Top             =   1650
      Width           =   840
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Contratante:"
      Height          =   195
      Left            =   75
      TabIndex        =   9
      Top             =   555
      Width           =   870
   End
   Begin VB.Label LblContPersNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   2430
      TabIndex        =   8
      Top             =   540
      Width           =   6165
   End
   Begin VB.Label LblContPersCod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   1065
      TabIndex        =   7
      Top             =   540
      Width           =   1350
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Aseguradora:"
      Height          =   195
      Left            =   75
      TabIndex        =   6
      Top             =   1005
      Width           =   945
   End
   Begin VB.Label LblAsegPersNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   2430
      TabIndex        =   5
      Top             =   960
      Width           =   6165
   End
   Begin VB.Label LblAsegPersCod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   1065
      TabIndex        =   4
      Top             =   990
      Width           =   1350
   End
   Begin VB.Label Label1 
      Caption         =   "Numero:"
      Height          =   240
      Left            =   75
      TabIndex        =   1
      Top             =   150
      Width           =   615
   End
End
Attribute VB_Name = "frmCredPolizas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nTipoOperacion As Integer '0 Nuevo...1 Modificar
Dim lnTotal As Double
Private fbHayPolExterna  As Boolean 'WIOR 20130829
Sub Limpiar_Controles(bTipoLimpieza As Boolean)  'WIOR 20120312
    txtMonto.Text = "0.00"
    txtPrimaNeta.Text = "0.00"
    txtNumAnio.Text = "0"
    lnTotal = 0
    txtNumPoliza.Text = ""
    txtNumCertif.Text = "" 'peac 20071227
   If bTipoLimpieza = True Then 'WIOR 20120312
        LblAsegPersCod.Caption = ""
        LblAsegPersNombre.Caption = ""
        LblContPersCod.Caption = ""
        LblContPersNombre.Caption = ""
    Else
        cmdBuscaAseg.Enabled = False 'WIOR 20120312
        cmdBuscaCont.Enabled = False 'WIOR 20120312
    End If
    cboTipo.ListIndex = -1
    cboMoneda.ListIndex = -1
    chkExterna.value = 0
    txtSumaA.Text = "0.00"
    txtTc.Text = "0.00"
End Sub

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    SendKeys "{Tab}", True
End If
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then chkExterna.SetFocus
End Sub

'JUEZ 20130417 ********************************************************
Private Sub chkExterna_Click()
If Not fbHayPolExterna Then  'WIOR 20130829
    If chkExterna.value = 1 Then
        Dim oCred As New COMDCredito.DCOMCredito
        If oCred.ExisteComisionVigente(LblContPersCod.Caption, gComisionEvalPolEndosada) = False Then
            MsgBox "Para poder registrar póliza externa primero el cliente debe realizar el pago por Evaluación de Póliza Endosada", vbInformation, "Aviso"
            chkExterna.value = 0
        End If
    End If
End If
End Sub
'END JUEZ *************************************************************

Private Sub chkExterna_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtSumaA.SetFocus
    fEnfoque txtSumaA
End If
End Sub

Private Sub chkSumAsegurada_Click()
    If chkSumAsegurada.value = 1 Then
        Label2.Caption = "Suma Asegurada S/."
    Else
        Label2.Caption = "Suma Asegurada US$"
    End If
End Sub

Private Sub cmdBuscaAseg_Click()
Dim oPers As COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        LblAsegPersCod.Caption = oPers.sPersCod
        LblAsegPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
    'txtMonto.SetFocus
    'fEnfoque txtMonto
End Sub

Private Sub cmdBuscaCont_Click()
Dim oPers As COMDPersona.UCOMPersona
    Set oPers = frmBuscaPersona.Inicio
    If Not oPers Is Nothing Then
        LblContPersCod.Caption = oPers.sPersCod
        LblContPersNombre.Caption = oPers.sPersNombre
    End If
    Set oPers = Nothing
    cmdBuscaAseg.SetFocus
End Sub

Private Sub cmdCalculaPrima_Click()
Dim nMonedaSumaAsegurada As Integer

If MsgBox("¿Está seguro que desea Calcular la Prima Neta y la Prima Total?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    If CDbl(Me.txtSumaA.Text) <= 0 Then
        MsgBox "La suma asegurada debe tener un monto mayor a cero para poder hacer los cálculos", vbInformation, "Atención"
        Exit Sub
    End If
    
    Dim oPol As COMDCredito.DCOMPoliza
    Dim rs As ADODB.Recordset
    Set oPol = New COMDCredito.DCOMPoliza
       
    Set rs = oPol.CargaDatosPrimaNetaPoliza(Me.txtNumPoliza, CDbl(Me.txtSumaA.Text), gdFecSis)
    If Not rs.EOF Then
        txtPrimaNeta.Text = Format(rs!PrimaNeta, "#0.00")
        txtMonto.Text = Format(rs!Total, "#0.00")
        If rs!MonePoliza = 1 Then
            txtTc.Text = Format(rs!Total * IIf(chkSumAsegurada.value = 1, 1, rs!TC), "#0.00")
        Else
            txtTc.Text = Format(rs!Total * 1, "#0.00")
        End If
    End If
    Set oPol = Nothing
End Sub

Private Sub cmdCancelar_Click()
    'Call Limpiar_Controles
    Call Limpiar_Controles(True) 'WIOR 20120315
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)
    cmdCalculaPrima.Enabled = False
    fbHayPolExterna = False 'WIOR 20130829
End Sub

Private Sub CmdEditar_Click()
nTipoOperacion = 1
Call Habilita_Grabar(True)
Call Habilita_Datos(True)
CmdGrabar.Enabled = True
cmdEditar.Enabled = False
cmdEliminar.Enabled = False
cmdGarantias.Enabled = False
cmdcancelar.Enabled = True
cmdCalculaPrima.Enabled = True
lnTotal = 0
End Sub

Private Sub cmdEliminar_Click()
Dim oPol As COMDCredito.DCOMPoliza

If MsgBox("Esta seguro que desea eliminar la Poliza?" & vbCrLf & "Se eliminarán las garantias asociadas", vbQuestion + vbYesNo, "Mensaje") = vbYes Then
    Set oPol = New COMDCredito.DCOMPoliza
    Call oPol.EliminaPoliza(txtNumPoliza.Text)
    Set oPol = Nothing
    cmdEliminar.Enabled = False
    cmdEditar.Enabled = False
End If
End Sub

Private Sub cmdExaminar_Click()
Dim oPol As COMDCredito.DCOMPoliza
Dim rs As ADODB.Recordset
Set oPol = New COMDCredito.DCOMPoliza
'*** PEAC 20080201
Dim nEstPol As Integer
fbHayPolExterna = False 'WIOR 20130829
Call frmCredPolizaListado.Inicio(nBusqueda)
'peac 20071128 se modifico este componente
Set rs = oPol.CargaDatosPoliza(frmCredPolizaListado.sNumPoliza)
If Not rs.EOF Then
    txtNumPoliza.Text = frmCredPolizaListado.sNumPoliza
    txtMonto.Text = Format(rs!nMontoPrimaTotal, "#0.00")
    '*** PEAC 20080626
    txtTc.Text = Format(rs!nTotal, "#0.00")
    
    LblContPersCod.Caption = rs!cPersCodContr
    LblContPersNombre.Caption = rs!Contratante
    LblAsegPersCod.Caption = rs!cPersCodAseg
    LblAsegPersNombre.Caption = rs!Aseguradora
    cboTipo.ListIndex = IndiceListaCombo(cboTipo, Trim(str(rs!nTipoPoliza)), 2)
    cboMoneda.ListIndex = IndiceListaCombo(cboMoneda, Trim(str(rs!nmoneda)), 2)
    'WIOR 20130829 ************************************
    If CBool(rs!bPolizaExterna) Then
        fbHayPolExterna = True
    End If
    'WIOR FIN ******************************************
    chkExterna.value = rs!bPolizaExterna
    txtSumaA.Text = rs!nSumaAsegurada
    nEstPol = rs!nEstado
    'peac 20071128
    txtNumCertif.Text = IIf(IsNull(rs!cCodPolizaAseg), " ", rs!cCodPolizaAseg) 'cNumCertif
    txtPrimaNeta.Text = Format(rs!nPrimaNeta, "#0.00")
    txtNumAnio.Text = Format(rs!nNumAnio, "#0")
    Me.cmdRelaciones.Enabled = True 'WIOR 20120315
    If rs!iTipoPago = 1 Then
        Option1.value = True
        Option2.value = False
        'LblNroCredito.Visible = False
        'txtNroCredito.Visible = False
        'cmdValidar.Visible = False
    Else
        Option1.value = False
        Option2.value = True
        'LblNroCredito.Visible = True
        'txtNroCredito.Visible = True
        'txtNroCredito.Text = rs!vCtaCod
        'cmdValidar.Visible = True
    End If
End If
Set oPol = Nothing
cmdGarantias.Enabled = True

'*** PEAC 20080412
'If nEstPol = 3 Then
'    cmdEditar.Enabled = False
'    cmdGarantias.Enabled = False
'    cmdEliminar.Enabled = False
'    cmdGrabar.Enabled = False
'Else
'    cmdEditar.Enabled = True
'    cmdGarantias.Enabled = True
'    cmdEliminar.Enabled = True
'    cmdGrabar.Enabled = True
'End If
End Sub

Private Sub cmdGarantias_Click()
'MAVM 20090922
frmCredPolizaGarantia.Inicio 0
'Call frmCredPolizaGarantia.Garantias_Poliza(LblContPersCod.Caption, LblContPersNombre.Caption, txtNumPoliza.Text)
End Sub

Private Sub cmdGrabar_Click()
Dim oPol As COMDCredito.DCOMPoliza
Dim nTC As Double
If Valida_Datos = False Then Exit Sub
fbHayPolExterna = False 'WIOR 20130829
If MsgBox("¿Esta seguro de registrar la poliza ?", vbQuestion + vbYesNo) = vbNo Then Exit Sub

Set oPol = New COMDCredito.DCOMPoliza
If nTipoOperacion = 0 Then
    'peac 20071128 se agregó "Trim(txtNumCertif.Text), CDbl(txtPrimaNeta.Text)"
     'MAVM se agrego el campo iTipoPago 1: Pago Desembolso, 2: Pago en cuotas
     Call oPol.RegistroPoliza(txtNumPoliza.Text, LblContPersCod.Caption, LblAsegPersCod.Caption, CDbl(txtMonto.Text), CInt(Trim(Right(cboTipo.Text, 20))), gdFecSis, chkExterna.value, CDbl(txtSumaA.Text), Trim(txtNumCertif.Text), CDbl(txtPrimaNeta.Text), CInt(Trim(Right(cboMoneda.Text, 20))), CInt(txtNumAnio.Text), IIf(Option1.value = True, "1", "2"))

Else
    'peac 20071128 se agregó "Trim(txtNumCertif.Text), CDbl(txtPrimaNeta.Text)"
    'MAVM 20090920
    'MADM 20111006 nTc
    nTC = -1
    If txtTc.Text <> "" Then
        nTC = CDbl(txtTc.Text)
    End If
    Call oPol.ModificaPoliza(txtNumPoliza.Text, LblContPersCod.Caption, LblAsegPersCod.Caption, CDbl(txtMonto.Text), CInt(Trim(Right(cboTipo.Text, 20))), , gdFecSis, , , , , , chkExterna.value, CDbl(txtSumaA.Text), Trim(txtNumCertif.Text), CDbl(txtPrimaNeta.Text), CInt(Trim(Right(cboMoneda.Text, 20))), CInt(txtNumAnio.Text), nTC, IIf(Option1.value = True, 1, 2))
End If
'JUEZ 20130417 ******************************************************
If chkExterna.value = 1 Then
    Dim oCredAct As COMDCredito.DCOMCredActBD
    Set oCredAct = New COMDCredito.DCOMCredActBD
    
    'WIOR 20150618 ***
    If Not fbHayPolExterna Then
        Call oCredAct.dUpdateComision(LblContPersCod.Caption, gComisionEvalPolEndosada)
    End If
    'WIOR FIN ********
    
    fbHayPolExterna = True 'WIOR 20130829
End If
'END JUEZ ***********************************************************
Set oPol = Nothing
    CmdGrabar.Enabled = False
    cmdEditar.Enabled = True
    cmdEliminar.Enabled = True
    cmdGarantias.Enabled = True
    cmdcancelar.Enabled = False
    cmdCalculaPrima.Enabled = False
    Call Habilita_Grabar(False)
    Call Habilita_Datos(False)

End Sub

Private Sub cmdNuevo_Click()
    Dim oPol As COMDCredito.DCOMPoliza
    Set oPol = New COMDCredito.DCOMPoliza
    nTipoOperacion = 0
    'Call Limpiar_Controles
    Call Limpiar_Controles(True) 'WIOR 20120315
    txtNumPoliza.Text = oPol.RecuperaNumeroPoliza
    Set oPol = Nothing
    Call Habilita_Grabar(True)
    Call Habilita_Datos(True)
    CmdGrabar.Enabled = True
    cmdEditar.Enabled = False
    cmdEliminar.Enabled = False
    cmdGarantias.Enabled = False
    cmdcancelar.Enabled = True
    fbHayPolExterna = False 'WIOR 20130829
End Sub
'WIOR 20120315-INICIO***********
Private Sub cmdRelaciones_Click()
If Me.LblContPersCod.Caption <> "" Then
    frmCredGarantPol.Inicio (Trim(Me.LblContPersCod.Caption))
Else
    MsgBox "Debe Ingresar el contratante para ver sus relaciones.CREDITO-GARANTIA-POLIZA", vbInformation, "AVISO"
End If

End Sub
'FIN***********************
Private Sub cmdsalir_Click()
    fbHayPolExterna = False 'WIOR 20130829
    Unload Me
End Sub


Private Sub Command1_Click()
    frmCredPolizasAprobacion.CargarDatos (txtNumPoliza.Text)
    frmCredPolizasAprobacion.Show 1
End Sub

Private Sub Form_Load()
Dim oCons As COMDConstantes.DCOMConstantes
Dim rs As ADODB.Recordset
Dim RM As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes
Set rs = oCons.RecuperaConstantes(9066)
Set RM = oCons.RecuperaConstantes(1011)
Set oCons = Nothing

Call CentraForm(Me)
Call Llenar_Combo_con_Recordset(rs, cboTipo)
Call Llenar_Combo_con_Recordset(RM, cboMoneda)
Call Habilita_Grabar(False)
Call Habilita_Datos(False)
End Sub

Sub Habilita_Grabar(ByVal pbHabilita As Boolean)
    CmdGrabar.Visible = pbHabilita
    cmdNuevo.Visible = Not pbHabilita
End Sub

Private Sub Option1_Click()
'        LblNroCredito.Visible = False
'        txtNroCredito.Visible = False
'        cmdValidar.Visible = False
End Sub

Private Sub Option2_Click()
'    LblNroCredito.Visible = True
'    txtNroCredito.Visible = True
'    cmdValidar.Visible = True
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosDecimales(txtMonto, KeyAscii)
     If KeyAscii = 13 Then
        cboTipo.SetFocus
     End If
End Sub

Private Sub txtMonto_LostFocus()
If txtMonto.Text = "" Then
    txtMonto.Text = "0.00"
Else
    txtMonto.Text = Format(txtMonto.Text, "#0.00")
End If
End Sub

Function Valida_Datos() As Boolean
Dim sNumCretif As String 'WIOR 20121203
Valida_Datos = True
'If CDbl(txtMonto.Text) = 0 Then
'    MsgBox "Debe indicar el valor de la prima", vbInformation, "Mensaje"
'    txtMonto.SetFocus
'    Valida_Datos = False
'    Exit Function
'End If
If LblContPersCod.Caption = "" Then
    MsgBox "Debe indicar el contratante", vbInformation, "Mensaje"
    Valida_Datos = False
    cmdBuscaCont.SetFocus
    Exit Function
End If
If LblAsegPersCod.Caption = "" Then
    MsgBox "Debe indicar la aseguradora", vbInformation, "Mensaje"
    Valida_Datos = False
    cmdBuscaAseg.SetFocus
    Exit Function
End If
If cboTipo.ListIndex = -1 Then
    MsgBox "Debe indicar el tipo de poliza", vbInformation, "Mensaje"
    Valida_Datos = False
    cboTipo.SetFocus
    Exit Function
End If

If cboMoneda.ListIndex = -1 Then
    MsgBox "Debe indicar la moneda de la poliza", vbInformation, "Mensaje"
    Valida_Datos = False
    cboMoneda.SetFocus
    Exit Function
End If


If CDbl(txtSumaA.Text) = 0 Then
    MsgBox "Debe indicar la suma asegurada", vbInformation, "Mensaje"
    Valida_Datos = False
    txtSumaA.SetFocus
End If
'WIOR 20130410 **************************
If Trim(txtNumCertif.Text) = "" Then
    MsgBox "Debe Ingresar un numero de Certificado.", vbInformation, "Mensaje"
    Valida_Datos = False
    txtNumCertif.SetFocus
End If
'WIOR FIN *******************************
'WIOR 20121203 **************************
sNumCretif = Replace(Replace(txtNumCertif.Text, "-", ""), ".", "")
If IsNumeric(sNumCretif) Then
    If CDbl(sNumCretif) = 0 Then
        MsgBox "Debe Ingresar un numero de Certificado Valido.", vbCritical, "Mensaje"
        Valida_Datos = False
        txtNumCertif.SetFocus
    End If
End If
'WIOR FIN *******************************
End Function

Private Sub Habilita_Datos(ByVal pbHabilita As Boolean)
    txtNumCertif.Enabled = pbHabilita 'peac 20071227
    txtNumAnio.Enabled = pbHabilita
    txtSumaA.Enabled = pbHabilita 'peac 20071227
    cmdBuscaAseg.Enabled = pbHabilita
    cmdBuscaCont.Enabled = pbHabilita
    cboTipo.Enabled = pbHabilita
    cboMoneda.Enabled = pbHabilita
    chkExterna.Enabled = pbHabilita
    cmdRelaciones.Enabled = pbHabilita  'WIOR 20120315
End Sub

Private Sub txtNumCertif_GotFocus()
    fEnfoque txtNumCertif
End Sub

'peac 20071128 convierte a mayuscula miestras escribe
Private Sub txtNumCertif_Change()
'Dim i As Integer
'    txtNumCertif.Text = UCase(txtNumCertif.Text)
'    i = Len(txtNumCertif.Text)
'    txtNumCertif.SelStart = i
End Sub


Private Sub txtSumaA_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosDecimales(txtSumaA, KeyAscii)
'     If KeyAscii = 13 Then cmdGrabar.SetFocus
End Sub

Private Sub txtSumaA_LostFocus()
If txtSumaA.Text = "" Then
    txtSumaA.Text = "0.00"
Else
    txtSumaA.Text = Format(txtSumaA.Text, "#0.00")
End If
End Sub


