VERSION 5.00
Begin VB.Form frmCredNewNivAprHist 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Historial de Aprobaciones"
   ClientHeight    =   3750
   ClientLeft      =   14895
   ClientTop       =   -795
   ClientWidth     =   11790
   Icon            =   "frmCredNewNivAprHist.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   11790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Cancel          =   -1  'True
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
      Height          =   360
      Left            =   10440
      TabIndex        =   12
      Top             =   3240
      Width           =   1170
   End
   Begin VB.Frame Frame1 
      Caption         =   " Valores sugeridos por el Analista "
      Height          =   975
      Left            =   8400
      TabIndex        =   2
      Top             =   75
      Width           =   3255
      Begin VB.Label Label6 
         Caption         =   "Tasa:"
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
         Left            =   120
         TabIndex        =   10
         Top             =   645
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "%"
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
         Left            =   1320
         TabIndex        =   9
         Top             =   645
         Width           =   255
      End
      Begin VB.Label lblCuotas 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   2520
         TabIndex        =   8
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label9 
         Caption         =   "Nº Cuotas:"
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
         Left            =   1680
         TabIndex        =   7
         Top             =   645
         Width           =   855
      End
      Begin VB.Label lblTasa 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   600
         TabIndex        =   6
         Top             =   600
         Width           =   645
      End
      Begin VB.Label Label28 
         Caption         =   "Monto Solicitado:"
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
         Left            =   120
         TabIndex        =   5
         Top             =   345
         Width           =   1310
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1440
         TabIndex        =   4
         Top             =   285
         Width           =   390
      End
      Begin VB.Label lblMontoSug 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000004&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1860
         TabIndex        =   3
         Top             =   285
         Width           =   1305
      End
   End
   Begin VB.CommandButton CmdBuscar 
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
      Left            =   3960
      TabIndex        =   0
      Top             =   210
      Width           =   900
   End
   Begin SICMACT.ActXCodCta ActxCta 
      Height          =   435
      Left            =   240
      TabIndex        =   1
      Top             =   180
      Width           =   3660
      _ExtentX        =   6456
      _ExtentY        =   767
      Texto           =   "Credito :"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin SICMACT.FlexEdit feHistorial 
      Height          =   1935
      Left            =   120
      TabIndex        =   11
      Top             =   1200
      Width           =   11535
      _ExtentX        =   20346
      _ExtentY        =   3413
      Cols0           =   12
      HighLight       =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      EncabezadosNombres=   "-cNivAprCod-Nivel-Firmas Solic-Firmas Aprob.-Usuario-Estado-Result-Monto-Tasa-Nº Cuotas-Comentario"
      EncabezadosAnchos=   "0-0-2100-1200-1300-900-700-700-1300-900-950-1100"
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
      ColumnasAEditar =   "X-X-X-X-X-X-6-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-4-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-L-C-C-C-C-C-R-R-C-R"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0"
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmCredNewNivAprHist"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oDNiv As COMDCredito.DCOMNivelAprobacion
Dim rs As ADODB.Recordset

Public Sub InicioCredito(ByVal psCtaCod As String)
    ActxCta.NroCuenta = psCtaCod
    ActxCta.Enabled = False
    CmdBuscar.Enabled = False
    CargaDatos (ActxCta.NroCuenta)
    Me.Show 1
End Sub

Public Sub Inicio()
    'If gnAgenciaCredEval = 0 Then
    '    MsgBox "Agencia No configurada para este proceso.", vbInformation, "Aviso"
    '    Exit Sub
    'End If
    ActxCta.CMAC = "109"
    ActxCta.Age = gsCodAge
    Me.Show 1
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CargaDatos (ActxCta.NroCuenta)
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
    Call LimpiaFlex(feHistorial)
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Call FrmVerCredito.Inicio(oPers.sPersCod, , True)
        ActxCta.SetFocusCuenta
    End If
    Set oPers = Nothing
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    Dim lnFila As Integer
    Set oDNiv = New COMDCredito.DCOMNivelAprobacion
    CargaDatos = False
    
    'RECO20150217 ERS173-2014****************************
    Dim oCredito As New DCOMCredito
'    Dim oCliPre As New COMNCredito.NCOMCredito 'COMENTADO POR ARLO 20170722
    Dim oDR As New Recordset
    Dim bClientPref As Boolean
    
    On Error GoTo ErrCargar
    Screen.MousePointer = 11
    
    Set oDR = oCredito.RecuperaRelacPers(psCtaCod)
    
    If Not (oDR.EOF And oDR.BOF) Then
        'bClientPref = oCliPre.ValidarClientePreferencial(oDR!cPersCod) 'COMENTADO POR ARLO 20170722
        bClientPref = False 'ARLO 20170722
    Else
        Exit Function
    End If
    Set rs = oDNiv.RecuperaHistorialCredAprobados(psCtaCod, IIf(bClientPref = True, 2, 1))
    'RECO FIN************************
    'Set rs = oDNiv.RecuperaHistorialCredAprobados(psCtaCod)
    Set oDNiv = Nothing
    Call LimpiaFlex(feHistorial)
    If Not rs.EOF Then
        If Not IsNull(rs!cNivAprDesc) Then
            Do While Not rs.EOF
                feHistorial.AdicionaFila
                lnFila = feHistorial.row
                feHistorial.TextMatrix(lnFila, 1) = rs!cNivAprCod
                feHistorial.TextMatrix(lnFila, 2) = rs!cNivAprDesc
                feHistorial.TextMatrix(lnFila, 3) = rs!nFirmasSolic
                feHistorial.TextMatrix(lnFila, 4) = rs!nFirmasAprob
                feHistorial.TextMatrix(lnFila, 5) = rs!cUserApr
                feHistorial.TextMatrix(lnFila, 6) = IIf(rs!nEstado = 0, "", "1")
                feHistorial.TextMatrix(lnFila, 7) = IIf(rs!nResultado = 0, "", "A")
                feHistorial.TextMatrix(lnFila, 8) = IIf(rs!nMonto = 0, "", Format(rs!nMonto, "#,##0.00"))
                feHistorial.TextMatrix(lnFila, 9) = IIf(rs!nTasa = 0, "", Format(rs!nTasa, "#,##0.0000"))
                feHistorial.TextMatrix(lnFila, 10) = IIf(rs!nCuotas = 0, "", rs!nCuotas)
                feHistorial.TextMatrix(lnFila, 11) = "..."
                rs.MoveNext
            Loop
        End If
        rs.MoveFirst
        lblMoneda.Caption = IIf(Mid(psCtaCod, 9, 1) = "1", "S/.", "$")
        lblMontoSug.Caption = Format(rs!nMontoSug, "#,##0.00")
        lblTasa.Caption = Format(rs!nTasaSug, "#,##0.0000")
        lblCuotas.Caption = rs!nCuotaSug
    End If
    Screen.MousePointer = 0
    CargaDatos = True
    Exit Function
ErrCargar:
    Screen.MousePointer = 0
    CargaDatos = False
    MsgBox Err.Description, vbCritical, "Aviso"
End Function

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub feHistorial_Click()
    If feHistorial.TextMatrix(feHistorial.row, 3) <> "" Then
        If feHistorial.Col = 11 Then
            Set oDNiv = New COMDCredito.DCOMNivelAprobacion
                Set rs = oDNiv.RecuperaDatosCredResultado(ActxCta.NroCuenta, feHistorial.TextMatrix(feHistorial.row, 1), feHistorial.TextMatrix(feHistorial.row, 5))
            Set oDNiv = Nothing
            If Not rs.EOF Then
                frmCredListaDatos.InicioTextBox "Comentarios", rs!cComent
            End If
        End If
    End If
End Sub
