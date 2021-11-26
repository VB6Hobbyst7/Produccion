VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredPolizaConsulta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Poliza"
   ClientHeight    =   5955
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7815
   Icon            =   "frmCredPolizaConsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   7815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6755
      TabIndex        =   10
      Top             =   5540
      Width           =   1000
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5710
      TabIndex        =   9
      Top             =   5540
      Width           =   1000
   End
   Begin VB.TextBox txtSumaA 
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
      Left            =   3975
      TabIndex        =   8
      Text            =   "0.00"
      Top             =   2925
      Width           =   1395
   End
   Begin VB.CheckBox chkExterna 
      Caption         =   "Poliza Externa"
      Enabled         =   0   'False
      Height          =   240
      Left            =   1050
      TabIndex        =   7
      Top             =   2970
      Width           =   1440
   End
   Begin VB.ComboBox cboEstado 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3975
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   600
      Width           =   3390
   End
   Begin VB.TextBox txtNumPoliza 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1065
      TabIndex        =   0
      Top             =   75
      Width           =   1440
   End
   Begin VB.TextBox txtCodigo 
      Enabled         =   0   'False
      Height          =   345
      Left            =   1050
      TabIndex        =   1
      Top             =   600
      Width           =   1470
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
      Left            =   1050
      TabIndex        =   3
      Text            =   "0.00"
      Top             =   2010
      Width           =   1395
   End
   Begin VB.ComboBox cboTipo 
      Enabled         =   0   'False
      Height          =   315
      Left            =   3975
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2025
      Width           =   3315
   End
   Begin VB.Frame fraGarantias 
      Caption         =   "Garantias"
      Height          =   2115
      Left            =   150
      TabIndex        =   11
      Top             =   3375
      Width           =   7590
      Begin SICMACT.FlexEdit FeGarantias 
         Height          =   1665
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Width           =   7290
         _ExtentX        =   12859
         _ExtentY        =   2937
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Garantía-Titular-Fec. Tasación"
         EncabezadosAnchos=   "0-1600-4200-1400"
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
         ColumnasAEditar =   "X-X-X-3"
         ListaControles  =   "0-0-0-2"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C"
         FormatosEdit    =   "0-0-0-0"
         lbUltimaInstancia=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin MSMask.MaskEdBox mskVigencia 
      Height          =   345
      Left            =   1050
      TabIndex        =   5
      Top             =   2475
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox mskVencimien 
      Height          =   345
      Left            =   3975
      TabIndex        =   6
      Top             =   2475
      Width           =   1365
      _ExtentX        =   2408
      _ExtentY        =   609
      _Version        =   393216
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label8 
      Caption         =   "Suma Asegurada:"
      Height          =   240
      Left            =   2700
      TabIndex        =   26
      Top             =   2970
      Width           =   1515
   End
   Begin VB.Label Label7 
      Caption         =   "Estado:"
      Enabled         =   0   'False
      Height          =   240
      Left            =   2850
      TabIndex        =   25
      Top             =   675
      Width           =   690
   End
   Begin VB.Label Label6 
      Caption         =   "Vencimiento:"
      Enabled         =   0   'False
      Height          =   240
      Left            =   2700
      TabIndex        =   24
      Top             =   2550
      Width           =   915
   End
   Begin VB.Label Label5 
      Caption         =   "Vigencia:"
      Enabled         =   0   'False
      Height          =   240
      Left            =   150
      TabIndex        =   23
      Top             =   2550
      Width           =   765
   End
   Begin VB.Label Label1 
      Caption         =   "Código:"
      Enabled         =   0   'False
      Height          =   240
      Left            =   150
      TabIndex        =   22
      Top             =   150
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Num Poliza:"
      Enabled         =   0   'False
      Height          =   255
      Left            =   150
      TabIndex        =   21
      Top             =   675
      Width           =   885
   End
   Begin VB.Label LblAsegPersCod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   1065
      TabIndex        =   20
      Top             =   1515
      Width           =   1350
   End
   Begin VB.Label LblAsegPersNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   2430
      TabIndex        =   19
      Top             =   1515
      Width           =   4950
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Aseguradora:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   75
      TabIndex        =   18
      Top             =   1530
      Width           =   945
   End
   Begin VB.Label LblContPersCod 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   1065
      TabIndex        =   17
      Top             =   1065
      Width           =   1350
   End
   Begin VB.Label LblContPersNombre 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H8000000D&
      Height          =   345
      Left            =   2430
      TabIndex        =   16
      Top             =   1065
      Width           =   4950
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Contratante:"
      Enabled         =   0   'False
      Height          =   195
      Left            =   75
      TabIndex        =   15
      Top             =   1080
      Width           =   870
   End
   Begin VB.Label Label3 
      Caption         =   "Prima:"
      Enabled         =   0   'False
      Height          =   240
      Left            =   150
      TabIndex        =   14
      Top             =   2055
      Width           =   915
   End
   Begin VB.Label Label4 
      Caption         =   "Tipo:"
      Enabled         =   0   'False
      Height          =   315
      Left            =   2700
      TabIndex        =   13
      Top             =   2025
      Width           =   615
   End
End
Attribute VB_Name = "frmCredPolizaConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fdVigenciaIni As Date
Dim fdVencimientoIni As Date

Public Sub Inicio(ByVal pcNumPoliza As String)
    Dim oPol As COMDCredito.DCOMPoliza
    Dim rs As ADODB.Recordset
    
    Dim bExisteDatos As Boolean  'RECO20150520 ERS010-2015
    
    Set oPol = New COMDCredito.DCOMPoliza
    
    Set rs = oPol.CargaDatosPoliza(pcNumPoliza)
    bExisteDatos = False 'RECO20150520 ERS010-2015
    If Not rs.EOF Then
        txtNumPoliza.Text = pcNumPoliza
        txtMonto.Text = Format(rs!nMontoPrimaTotal, "#0.00")
        LblContPersCod.Caption = rs!cPersCodContr
        LblContPersNombre.Caption = rs!Contratante
        LblAsegPersCod.Caption = rs!cPersCodAseg
        LblAsegPersNombre.Caption = rs!Aseguradora
        cboTipo.ListIndex = IndiceListaCombo(cboTipo, Trim(str(rs!nTipoPoliza)), 2)
        cboEstado.ListIndex = IndiceListaCombo(cboTipo, Trim(str(rs!nEstado)), 2)
        mskVigencia.Text = IIf(IsNull(rs!dVigencia), "__/__/____", rs!dVigencia)
        mskVencimien.Text = IIf(IsNull(rs!dVencimiento), "__/__/____", rs!dVencimiento)
        txtCodigo.Text = IIf(IsNull(rs!cCodPolizaAseg), "", rs!cCodPolizaAseg)
        chkExterna.value = rs!bPolizaExterna
        'EJVG20150623 ***
        cmdGrabar.Visible = rs!bPolizaExterna
        mskVigencia.Enabled = rs!bPolizaExterna
        mskVencimien.Enabled = rs!bPolizaExterna
        fdVigenciaIni = CDate(IIf(IsNull(rs!dVigencia), "01/01/1900", rs!dVigencia))
        fdVencimientoIni = CDate(IIf(IsNull(rs!dVencimiento), "01/01/1900", rs!dVencimiento))
        'END EJVG *******
        txtSumaA.Text = Format(rs!nSumaAsegurada, "#0.00")
        
        Set rs = oPol.Garantias_x_Poliza(txtNumPoliza.Text)
        With FeGarantias
            .Clear
            .FormaCabecera
            .Rows = 2
            .rsFlex = rs
        End With
        bExisteDatos = True 'RECO20150520 ERS010-2015
    End If
    Set oPol = Nothing
    If bExisteDatos = True Then 'RECO20150520 ERS010-2015
        Me.Show 1
    Else
        MsgBox "No existen Datos.", vbInformation, "Alerta"
    End If
End Sub

Sub Limpiar_Controles()
    txtMonto.Text = ""
    txtNumPoliza.Text = ""
    LblAsegPersCod.Caption = ""
    LblAsegPersNombre.Caption = ""
    LblContPersCod.Caption = ""
    LblContPersNombre.Caption = ""
    cboEstado.ListIndex = -1
    cboTipo.ListIndex = -1
    txtCodigo.Text = ""
    mskVigencia.Text = "__/__/____"
    mskVencimien.Text = "__/__/____"
    With FeGarantias
        .Clear
        .FormaCabecera
        .Rows = 2
    End With
End Sub

Private Sub cmdGrabar_Click()
    Dim oGarantia As COMDCredito.DCOMGarantia
    Dim lsFecha As String
    On Error GoTo ErrGrabar
    
    cmdGrabar.Enabled = False
    If Not validarGrabar Then
        cmdGrabar.Enabled = True
        Exit Sub
    End If
    
    If MsgBox("¿Esta seguro de actualiza los datos de la Poliza?", vbQuestion + vbYesNo) = vbNo Then
        cmdGrabar.Enabled = True
        Exit Sub
    End If
    
    Screen.MousePointer = 11
    Set oGarantia = New COMDCredito.DCOMGarantia
    oGarantia.ActualizarPeriodoPoliza txtNumPoliza.Text, CDate(mskVigencia.Text), CDate(mskVencimien.Text)
    fdVigenciaIni = CDate(mskVigencia.Text)
    fdVencimientoIni = CDate(mskVencimien.Text)
    Screen.MousePointer = 0
    
    MsgBox "Se ha actualizado satisfactoriamente el periodo de la Poliza", vbInformation, "Aviso"
    Set oGarantia = Nothing
    
    Exit Sub
ErrGrabar:
    Screen.MousePointer = 0
    cmdGrabar.Enabled = True
    MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Function validarGrabar() As Boolean
    Dim lsFecha As String
    
    lsFecha = ValidaFecha(mskVigencia.Text)
    If Len(lsFecha) > 0 Then
        MsgBox lsFecha, vbInformation, "Aviso"
        EnfocaControl mskVigencia
        Exit Function
    End If
    lsFecha = ValidaFecha(mskVencimien.Text)
    If Len(lsFecha) > 0 Then
        MsgBox lsFecha, vbInformation, "Aviso"
        EnfocaControl mskVencimien
        Exit Function
    End If
    If CDate(mskVigencia.Text) >= CDate(mskVigencia.Text) Then
        MsgBox "La fecha de Vigencia no puede ser mayor o igual a la fecha de vencimiento", vbInformation, "Aviso"
        EnfocaControl mskVigencia
        Exit Function
    End If
    If CDate(mskVigencia.Text) < fdVigenciaIni Then
        MsgBox "La fecha de Vigencia que se está editando no puede ser menor al " & Format(fdVigenciaIni, gsFormatoFechaView), vbInformation, "Aviso"
        EnfocaControl mskVigencia
        Exit Function
    End If
    If CDate(mskVencimien.Text) < fdVencimientoIni Then
        MsgBox "La fecha de Vencimiento que se está editando no puede ser menor al " & Format(fdVencimientoIni, gsFormatoFechaView), vbInformation, "Aviso"
        EnfocaControl mskVencimien
        Exit Function
    End If
    
    validarGrabar = True
End Function

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oCons As COMDConstantes.DCOMConstantes
Dim rs As ADODB.Recordset

Call CentraForm(Me)
Set oCons = New COMDConstantes.DCOMConstantes
Set rs = oCons.RecuperaConstantes(9066)
Call Llenar_Combo_con_Recordset(rs, cboTipo)
Set rs = oCons.RecuperaConstantes(9068)
Call Llenar_Combo_con_Recordset(rs, cboEstado)
Set oCons = Nothing

End Sub
