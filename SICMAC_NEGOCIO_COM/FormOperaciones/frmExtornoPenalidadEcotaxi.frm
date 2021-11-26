VERSION 5.00
Begin VB.Form frmExtornoPenalidadEcotaxi 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno Penalidad EcoTaxi"
   ClientHeight    =   4530
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11745
   Icon            =   "frmExtornoPenalidadEcotaxi.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4530
   ScaleWidth      =   11745
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnCancelar 
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
      Height          =   350
      Left            =   80
      TabIndex        =   16
      Top             =   4155
      Width           =   1095
   End
   Begin VB.Frame fraGlosa 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   975
      Left            =   9000
      TabIndex        =   14
      Top             =   0
      Width           =   2655
      Begin VB.TextBox txtGlosa 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   240
         Width           =   2415
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   975
      Left            =   1800
      TabIndex        =   7
      Top             =   0
      Width           =   7095
      Begin VB.CommandButton btnBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   6120
         TabIndex        =   13
         Top             =   360
         Width           =   900
      End
      Begin VB.Frame fraUsuario 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   60
         TabIndex        =   9
         Top             =   240
         Width           =   6135
         Begin SICMACT.TxtBuscar txtUser 
            Height          =   330
            Left            =   765
            TabIndex        =   10
            Top             =   120
            Width           =   1080
            _ExtentX        =   1905
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
            sTitulo         =   ""
         End
         Begin VB.Label lblNomUser 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   330
            Left            =   1845
            TabIndex        =   12
            Top             =   120
            Width           =   4140
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Usuario:"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   180
            Width           =   585
         End
      End
      Begin SICMACT.ActXCodCta ActXCtaCod 
         Height          =   405
         Left            =   120
         TabIndex        =   8
         Top             =   320
         Visible         =   0   'False
         Width           =   3705
         _ExtentX        =   6535
         _ExtentY        =   714
         Texto           =   "Nro Crédito :"
         EnabledCta      =   -1  'True
      End
   End
   Begin VB.OptionButton optNroCuenta 
      Caption         =   "Nro. Cuenta"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   550
      Width           =   1215
   End
   Begin VB.OptionButton optUsuario 
      Caption         =   "Usuario"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   240
      Value           =   -1  'True
      Width           =   975
   End
   Begin VB.CommandButton btnExtornar 
      Caption         =   "&Extornar"
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
      Height          =   350
      Left            =   9420
      TabIndex        =   3
      Top             =   4155
      Width           =   1095
   End
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
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
      Height          =   350
      Left            =   10560
      TabIndex        =   2
      Top             =   4155
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Penalidades a Extornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3120
      Left            =   50
      TabIndex        =   1
      Top             =   970
      Width           =   11655
      Begin SICMACT.FlexEdit fePenalidad 
         Height          =   2730
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   11460
         _ExtentX        =   20214
         _ExtentY        =   4815
         Cols0           =   15
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   $"frmExtornoPenalidadEcotaxi.frx":030A
         EncabezadosAnchos=   "350-2000-0-2000-2000-2500-0-2000-2000-2500-2000-2000-2000-2000-2000"
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-C-C-L-C-L-C-L-R-R-R-R-R"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-2-2-2-2-2"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
   End
   Begin VB.Frame fraOpcBuscar 
      Caption         =   "Buscar por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   960
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   1695
   End
End
Attribute VB_Name = "frmExtornoPenalidadEcotaxi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    CentraForm Me

    Dim oGen As COMDConstSistema.DCOMGeneral
    Set oGen = New COMDConstSistema.DCOMGeneral
    txtUser.psRaiz = "USUARIOS "
    txtUser.rs = oGen.GetUserAreaAgenciaResumenIngEgre("026", Right(gsCodAge, 2))
    Set oGen = Nothing
    
    Me.ActXCtaCod.Age = gsCodAge
    Me.ActXCtaCod.CMAC = gsCodCMAC
    Me.ActXCtaCod.Prod = "517"
End Sub
Private Sub optUsuario_Click()
    Me.fraUsuario.Visible = True
    Me.ActXCtaCod.Visible = False
    Me.txtUser.Text = ""
    Me.lblNomUser.Caption = ""
    Me.txtUser.SetFocus
End Sub
Private Sub optNroCuenta_Click()
    Me.fraUsuario.Visible = False
    Me.ActXCtaCod.Visible = True
    Me.ActXCtaCod.Cuenta = ""
    Me.ActXCtaCod.SetFocus
End Sub
Private Sub txtUser_EmiteDatos()
    Me.lblNomUser.Caption = Me.txtUser.psDescripcion
    If lblNomUser <> "" Then
        Me.btnBuscar.SetFocus
    End If
End Sub
Private Sub txtUser_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.btnBuscar.SetFocus
    End If
End Sub
Private Sub ActXCtaCod_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.btnBuscar.SetFocus
    End If
End Sub
Private Sub btnBuscar_Click()
    Dim oNCred As COMNCredito.NCOMCredito
    Dim rsPenalidad As ADODB.Recordset
    Dim lnTpoBusqueda As Integer
    Dim lsBusqueda As String
    
    FormatearGrillaPenalidad
    
    If optUsuario.value = True And txtUser.Text = "" Then
        MsgBox "Ud. debe seleccionar el usuario", vbInformation, "Aviso"
        Me.txtUser.SetFocus
        Exit Sub
    End If
    If optNroCuenta.value = True And Len(Me.ActXCtaCod.NroCuenta) <> 18 Then
        MsgBox "Ud. debe especificar el Nro de Cuenta", vbInformation, "Aviso"
        Me.ActXCtaCod.SetFocus
        Exit Sub
    End If
    lnTpoBusqueda = IIf(optUsuario.value = True, 1, 0)
    lsBusqueda = IIf(optUsuario.value = True, Me.txtUser.Text, Me.ActXCtaCod.NroCuenta)
    
    Set oNCred = New COMNCredito.NCOMCredito
    Set rsPenalidad = New ADODB.Recordset
    
    Set rsPenalidad = oNCred.ListaPenalidadEcoTaxiAExtornar(lnTpoBusqueda, lsBusqueda, gdFecSis)
    
    If RSVacio(rsPenalidad) Then
        MsgBox "No se encontraron Penalidades a Extornar", vbInformation, "Aviso"
        Exit Sub
    End If

     Do While Not rsPenalidad.EOF
        fePenalidad.AdicionaFila
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 1) = rsPenalidad!cMovNro
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 2) = rsPenalidad!nMovNro
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 3) = rsPenalidad!cCtaCodCreTit
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 4) = rsPenalidad!cCtaCodAhoTit
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 5) = rsPenalidad!cPersNombreTit
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 6) = rsPenalidad!nMotivoPenalidad
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 7) = rsPenalidad!cMotivoPenalidad
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 8) = rsPenalidad!cCtaCodAhoConcesionario
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 9) = rsPenalidad!cPersNombreConce
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 10) = Format(rsPenalidad!nMontoRetConce, "##,##0.00")
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 11) = Format(rsPenalidad!nITFCargoConce, "##,##0.00")
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 12) = Format(rsPenalidad!nMontoDepTit, "##,##0.00")
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 13) = Format(rsPenalidad!nITFCargoTit, "##,##0.00")
        Me.fePenalidad.TextMatrix(Me.fePenalidad.Row, 14) = Format(rsPenalidad!nMontoFavorInstitucion, "##,##0.00")
        rsPenalidad.MoveNext
     Loop

     Me.btnExtornar.Enabled = True
End Sub
Private Sub btnCancelar_Click()
    Me.txtGlosa.Text = ""
    Me.ActXCtaCod.Cuenta = ""
    Me.optUsuario.value = True
    Me.optNroCuenta.value = False
    optUsuario_Click
    FormatearGrillaPenalidad
    Me.btnExtornar.Enabled = False
End Sub
Private Sub btnExtornar_Click()
    Dim oBase As COMDCredito.DCOMCredActBD
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lnMotivoPenalidad As Integer
    Dim lsCtaCodCreTit As String, lsCtaCodAhoTit As String
    Dim lnMontoDepTit As Double, lnITFCargoTit As Double
    Dim lsCtaCodAhoConce As String
    Dim lnMontoRetConce As Double, lnITFCargoConce As Double
    Dim nRow As Integer
    Dim bTransac As Boolean
    
    If Not validaDatos Then Exit Sub
    
    If MsgBox("¿Esta seguro de extornar la Penalidad EcoTaxi seleccionada?", vbYesNo + vbInformation, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    Set oBase = New COMDCredito.DCOMCredActBD
    nRow = Me.fePenalidad.Row
    
    lsMovNro = Trim(Me.fePenalidad.TextMatrix(nRow, 1))
    lnMovNro = CLng(Trim(Me.fePenalidad.TextMatrix(nRow, 2)))
    lnMotivoPenalidad = CInt(Trim(Me.fePenalidad.TextMatrix(nRow, 6)))

    lsCtaCodCreTit = Trim(Me.fePenalidad.TextMatrix(nRow, 3))
    lsCtaCodAhoTit = Trim(Me.fePenalidad.TextMatrix(nRow, 4))
    lnMontoDepTit = CDbl(Trim(Me.fePenalidad.TextMatrix(nRow, 12)))
    lnITFCargoTit = CDbl(Trim(Me.fePenalidad.TextMatrix(nRow, 13)))

    lsCtaCodAhoConce = Trim(Me.fePenalidad.TextMatrix(nRow, 8))
    lnMontoRetConce = CDbl(Trim(Me.fePenalidad.TextMatrix(nRow, 10)))
    lnITFCargoConce = CDbl(Trim(Me.fePenalidad.TextMatrix(nRow, 11)))
    
On Error GoTo Error

    bTransac = False
    oBase.dBeginTrans
    bTransac = True

    lsMovNro = oBase.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    'Extorno Cargo Concesionario
    Call oBase.CapExtornoCargoAho(0, lnMovNro, gOtrOpeExtPenalidadEcoTaxi, lsCtaCodAhoConce, lsMovNro, "", lnMontoRetConce)
    If lnITFCargoConce > 0 Then
        Sleep (1000)
        lsMovNro = oBase.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Call oBase.CapExtornoCargoAho(0, lnMovNro, gOtrOpeExtPenalidadEcoTaxi, lsCtaCodAhoConce, lsMovNro, "", lnITFCargoConce)
    End If
    If lnMotivoPenalidad = 1 Then
        'Extorno Deposito a Titular
        Sleep (1000)
        lsMovNro = oBase.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Call oBase.CapExtornoAbonoAho(lsMovNro, lnMovNro, lnMovNro, gOtrOpeExtPenalidadEcoTaxi, lsCtaCodAhoTit, lsMovNro, "", lnMontoDepTit)
        If lnITFCargoTit > 0 Then
            Sleep (1000)
            lsMovNro = oBase.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
            Call oBase.CapExtornoCargoAho(0, lnMovNro, gOtrOpeExtPenalidadEcoTaxi, lsCtaCodAhoTit, lsMovNro, "", lnITFCargoConce)
        End If
    End If
    
    oBase.dCommitTrans
    bTransac = False
    
    MsgBox "Se ha Extornado con éxito la Penalidad EcoTaxi seleccionada", vbInformation, "Aviso"
    If MsgBox("¿Desea extornar otra Penalidad EcoTaxi?", vbYesNo + vbInformation, "Aviso") = vbYes Then
        Me.fePenalidad.EliminaFila nRow
        Me.txtGlosa.Text = ""
    Else
        Unload Me
    End If
    
    Set oBase = Nothing
    Exit Sub
Error:
    If bTransac Then
        oBase.dRollbackTrans
        Set oBase = Nothing
    End If
    err.Raise err.Number, "Error En Extorno", err.Description
End Sub
Private Sub btnSalir_Click()
    Unload Me
End Sub
Private Function validaDatos() As Boolean
    If Me.fePenalidad.Row <= 1 And Len(Me.fePenalidad.TextMatrix(1, 1)) = 0 Then
        validaDatos = False
        MsgBox "Ud. debe seleccionar la Penalidad a Extornar", vbInformation, "Aviso"
        Exit Function
    End If
    If Len(Trim(Me.txtGlosa.Text)) = 0 Then
        validaDatos = False
        MsgBox "Ud. debe ingresar la Glosa con respecto al Extorno", vbInformation, "Aviso"
        Exit Function
    End If
    validaDatos = True
End Function
Private Sub FormatearGrillaPenalidad()
    Me.fePenalidad.Clear
    Me.fePenalidad.FormaCabecera
    Me.fePenalidad.Rows = 2
End Sub
