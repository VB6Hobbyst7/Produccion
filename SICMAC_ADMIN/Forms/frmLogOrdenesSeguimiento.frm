VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogOrdenesSeguimiento 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Seguimiento de Ordenes"
   ClientHeight    =   6915
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   13425
   Icon            =   "frmLogOrdenesSeguimiento.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6915
   ScaleWidth      =   13425
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   375
      Left            =   12240
      TabIndex        =   17
      Top             =   6480
      Width           =   1095
   End
   Begin VB.CommandButton cmdDetalle 
      Caption         =   "Detalle"
      Height          =   375
      Left            =   11080
      TabIndex        =   16
      Top             =   6480
      Width           =   1095
   End
   Begin VB.Frame fraFiltro 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Filtro de Búsqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1335
      Left            =   50
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      Begin VB.ComboBox cboTpoDoc 
         Height          =   315
         ItemData        =   "frmLogOrdenesSeguimiento.frx":030A
         Left            =   1200
         List            =   "frmLogOrdenesSeguimiento.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   300
         Width           =   2055
      End
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmLogOrdenesSeguimiento.frx":0412
         Left            =   4200
         List            =   "frmLogOrdenesSeguimiento.frx":041F
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   300
         Width           =   2055
      End
      Begin VB.CheckBox chkSoloOrden 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Sólo Considerar (Nro. Orden o Glosa):"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   3615
      End
      Begin VB.TextBox txtNroOrden 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   375
         Left            =   3960
         TabIndex        =   9
         Top             =   720
         Width           =   3015
      End
      Begin VB.CommandButton cmdBuscar 
         Appearance      =   0  'Flat
         Caption         =   "Buscar"
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
         Left            =   7080
         MaskColor       =   &H00400040&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   720
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mskIni 
         Height          =   300
         Left            =   7200
         TabIndex        =   10
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskFin 
         Height          =   300
         Left            =   9120
         TabIndex        =   7
         Top             =   300
         Width           =   1200
         _ExtentX        =   2117
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Tpo Doc.:"
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
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Moneda:"
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
         Left            =   3360
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Desde:"
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
         Left            =   6480
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Hasta:"
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
         Left            =   8520
         TabIndex        =   6
         Top             =   360
         Width           =   975
      End
   End
   Begin Sicmact.FlexEdit dgOrden 
      Height          =   4695
      Left            =   45
      TabIndex        =   15
      Top             =   1680
      Width           =   13335
      _ExtentX        =   23521
      _ExtentY        =   8281
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Nro. Orden-Tpo-Proveedor-Fecha-Usuario-Glosa-Moneda-Monto-nMovNro"
      EncabezadosAnchos=   "300-1200-400-2600-800-700-5000-1000-1200-0"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-C-L-C-C-L-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblResultBusq 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Resultados de Búsqueda"
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
      Left            =   5280
      TabIndex        =   13
      Top             =   1340
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.Label lblTotalIzq 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total: 0 Registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1380
      Width           =   1935
   End
   Begin VB.Label lblTotalDer 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Total: 0 Registros"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H008080FF&
      Height          =   255
      Left            =   11520
      TabIndex        =   14
      Top             =   1380
      Width           =   1935
   End
End
Attribute VB_Name = "frmLogOrdenesSeguimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'ARLO 20170125******************
Dim objPista As COMManejador.Pista
'*******************************
Private Sub chkSoloOrden_Click()
    txtNroOrden.Text = ""
    EstadoControles IIf(chkSoloOrden.value, 1, 0)
    If CBool(chkSoloOrden.value) Then
        txtNroOrden.BackColor = RGB(252, 250, 207)
    Else
        txtNroOrden.BackColor = vbWhite
    End If
End Sub
Private Function ValidaInterfaz() As Boolean
    ValidaInterfaz = False
    If Not (CBool(chkSoloOrden.value)) Then
        If cboTpoDoc.ListIndex = -1 Then
            MsgBox "Asegure se haber seleccionado el tipo de Documento.", vbInformation, "Mensaje"
            cboTpoDoc.SetFocus
            Exit Function
        End If
        If cboMoneda.ListIndex = -1 Then
            MsgBox "Asegure se haber seleccionado el tipo de Moneda.", vbInformation, "Mensaje"
            cboMoneda.SetFocus
            Exit Function
        End If
        If Not (IsDate(mskIni.Text)) Then
            MsgBox "Asegure se haber ingresado la fecha desde correcta.", vbInformation, "Mensaje"
            mskIni.SetFocus
            Exit Function
        End If
        If Not (IsDate(mskFin.Text)) Then
            MsgBox "Asegure se haber ingresado la Fecha Hasta correcta.", vbInformation, "Mensaje"
            mskFin.SetFocus
            Exit Function
        End If
        If CDate(mskIni) > CDate(mskFin) Then
            MsgBox "La Fecha Desde no puede ser mayor a la Fecha Hasta.", vbInformation, "Mensaje"
            mskIni.SetFocus
            Exit Function
        End If
    Else
        If Len(Trim(txtNroOrden.Text)) = 0 Then
             MsgBox "No se ha ingresado ninguna consideración para la búsqueda. Verifique.", vbInformation, "Mensaje"
            txtNroOrden.SetFocus
            Exit Function
        End If
    End If
    ValidaInterfaz = True
End Function
Private Sub cmdBuscar_Click()
Dim rs As ADODB.Recordset
Dim psCtaCont As String
Dim psDocNro As String
Dim row As Integer
Dim nDocTpo As Integer
Dim i As Integer
Dim oProv As DLogGeneral
Set oProv = New DLogGeneral

    If Not ValidaInterfaz Then Exit Sub
        lblResultBusq.Visible = True
        lblResultBusq.Caption = "Buscando..."
        lblTotalIzq = "Total: 0 Registros"
        lblTotalDer = "Total: 0 Registros"
        LimpiaFlex dgOrden
        DoEvents
    If Not CBool(chkSoloOrden.value) Then
        If Right(cboTpoDoc.Text, 1) = 0 Then
            psCtaCont = "25M60"
        ElseIf Right(cboTpoDoc.Text, 1) = 1 Then
            psCtaCont = "25M601"
        Else
            psCtaCont = "25M60202"
        End If
        If Right(cboMoneda, 1) = 0 Then
            psCtaCont = Replace(psCtaCont, "M", "_")
        ElseIf Right(cboMoneda, 1) = 1 Then
            psCtaCont = Replace(psCtaCont, "M", "1")
        Else
            psCtaCont = Replace(psCtaCont, "M", "2")
        End If
        Set rs = oProv.ObtieneOrdenesxSeguimiento(psCtaCont, CDate(mskIni.Text), CDate(mskFin.Text))
    Else
        psDocNro = Trim(txtNroOrden.Text)
        Set rs = oProv.ObtieneOrdenesxSeguimientoxNroOrden(psDocNro, CDate(mskIni.Text), CDate(mskFin.Text))
    End If
    Do While Not rs.EOF
        dgOrden.AdicionaFila
        lblResultBusq.Caption = "Mostrando Resultados..."
        row = dgOrden.row
        dgOrden.TextMatrix(row, 1) = rs!cDocNro
        dgOrden.TextMatrix(row, 2) = rs!Tipo
        dgOrden.TextMatrix(row, 3) = rs!cPersNombre
        dgOrden.TextMatrix(row, 4) = rs!dDocFecha
        dgOrden.TextMatrix(row, 5) = rs!cuser
        dgOrden.TextMatrix(row, 6) = rs!cMovDesc
        dgOrden.TextMatrix(row, 7) = rs!Moneda
        dgOrden.TextMatrix(row, 8) = Format(rs!nMonto, "#,#0.00")
        dgOrden.TextMatrix(row, 9) = rs!nMovNro
        lblTotalIzq = "Total: " + CStr(row) + " Registros"
        lblTotalDer = "Total: " + CStr(row) + " Registros"
        For i = 0 To 10000
        Next i
        rs.MoveNext
    Loop
    'ARLO 20170125
    gsopecod = LogPistaseguientoOrdenes
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Busqueda de Ordenes"
    Set objPista = Nothing
    '***********
    lblResultBusq.Caption = "Resultados de Búsqueda"
End Sub

Private Sub cmdDetalle_Click()
    If Not Len(dgOrden.TextMatrix(dgOrden.row, 1)) = 0 Then
        frmLogOrdenDet.Inicio dgOrden.TextMatrix(dgOrden.row, 9), dgOrden.TextMatrix(dgOrden.row, 1), dgOrden.TextMatrix(dgOrden.row, 2), dgOrden.TextMatrix(dgOrden.row, 4), dgOrden.TextMatrix(dgOrden.row, 7), dgOrden.TextMatrix(dgOrden.row, 8), dgOrden.TextMatrix(dgOrden.row, 3), dgOrden.TextMatrix(dgOrden.row, 6)
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub dgOrden_DblClick()
    If Not Len(dgOrden.TextMatrix(dgOrden.row, 1)) = 0 And Not dgOrden.row = 0 Then
        frmLogOrdenDet.Inicio dgOrden.TextMatrix(dgOrden.row, 9), dgOrden.TextMatrix(dgOrden.row, 1), dgOrden.TextMatrix(dgOrden.row, 2), dgOrden.TextMatrix(dgOrden.row, 4), dgOrden.TextMatrix(dgOrden.row, 7), dgOrden.TextMatrix(dgOrden.row, 8), dgOrden.TextMatrix(dgOrden.row, 3), dgOrden.TextMatrix(dgOrden.row, 6)
    End If
End Sub
Private Sub Form_Load()
    EstadoControles 0
    mskIni.Text = DateAdd("M", -1, gdFecSis)
    mskFin.Text = gdFecSis
End Sub
Private Sub EstadoControles(ByVal pnEstado As Integer)
    Select Case pnEstado
        Case 0:
                        chkSoloOrden.value = 0
                        txtNroOrden.Enabled = False
                        lblResultBusq.Visible = False
                        cboTpoDoc.Enabled = True
                        cboMoneda.Enabled = True
                        mskIni.Enabled = True
                        mskFin.Enabled = True
       Case 1
                        cboTpoDoc.Enabled = False
                        cboMoneda.Enabled = False
                        'mskIni.Enabled = False
                        'mskFin.Enabled = False
                        txtNroOrden.Enabled = True
    End Select
End Sub
Private Sub mskFin_GotFocus()
    mskFin.SelStart = 0
    mskFin.SelLength = 50
End Sub
Private Sub mskFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdBuscar.SetFocus
    End If
End Sub
Private Sub mskIni_GotFocus()
    mskIni.SelStart = 0
    mskIni.SelLength = 50
End Sub
Private Sub mskIni_KeyPress(KeyAscii As Integer)
     If KeyAscii = 13 Then
        Me.mskFin.SetFocus
    End If
End Sub
Private Sub txtNroOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdBuscar.SetFocus
    End If
End Sub
