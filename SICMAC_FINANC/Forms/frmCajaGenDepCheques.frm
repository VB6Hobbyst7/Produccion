VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajaGenDepCheques 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6300
   ClientLeft      =   690
   ClientTop       =   1455
   ClientWidth     =   10710
   Icon            =   "frmCajaGenDepCheques.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6300
   ScaleWidth      =   10710
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDepositar 
      Caption         =   "&Depositar"
      Height          =   390
      Left            =   7815
      TabIndex        =   20
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdDevolver 
      Caption         =   "De&volver"
      Height          =   390
      Left            =   6480
      TabIndex        =   22
      Top             =   5760
      Width           =   1335
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   9165
      TabIndex        =   21
      Top             =   5760
      Width           =   1335
   End
   Begin VB.Frame Frame1 
      Caption         =   "Depositar en :"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1155
      Left            =   105
      TabIndex        =   13
      Top             =   4500
      Width           =   10350
      Begin VB.TextBox txtMovDesc 
         Height          =   720
         Left            =   5835
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   18
         Top             =   330
         Width           =   4440
      End
      Begin Sicmact.TxtBuscar txtBuscarBanco 
         Height          =   345
         Left            =   810
         TabIndex        =   14
         Top             =   255
         Width           =   2160
         _ExtentX        =   3810
         _ExtentY        =   609
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
         ForeColor       =   -2147483635
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
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
         Height          =   195
         Left            =   5850
         TabIndex        =   19
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lblCuentaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   810
         TabIndex        =   17
         Top             =   645
         Width           =   4950
      End
      Begin VB.Label lblBancoDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   345
         Left            =   2985
         TabIndex        =   16
         Top             =   270
         Width           =   2760
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Banco :"
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
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Top             =   300
         Width           =   675
      End
   End
   Begin VB.Frame fraLista 
      Caption         =   "Lista de Cheques"
      Height          =   3720
      Left            =   105
      TabIndex        =   7
      Top             =   765
      Width           =   10350
      Begin Sicmact.FlexEdit fgCheques 
         Height          =   2850
         Left            =   135
         TabIndex        =   8
         Top             =   270
         Width           =   10065
         _ExtentX        =   17754
         _ExtentY        =   5027
         Cols0           =   16
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmCajaGenDepCheques.frx":030A
         EncabezadosAnchos=   "350-350-2300-1500-1500-1000-1200-1200-1200-2500-0-0-0-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-L-C-R-L-L-L-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-2-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lbltotalCartera 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   345
         Left            =   1680
         TabIndex        =   12
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "En Cartera :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   435
         TabIndex        =   11
         Top             =   3300
         Width           =   1035
      End
      Begin VB.Label lbltotalDepositar 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "0.00"
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
         Height          =   345
         Left            =   8370
         TabIndex        =   10
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "A Depositar :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   195
         Left            =   7125
         TabIndex        =   9
         Top             =   3300
         Width           =   1125
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   390
         Left            =   6885
         Top             =   3210
         Width           =   3315
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   390
         Left            =   195
         Top             =   3210
         Width           =   3315
      End
   End
   Begin VB.CheckBox chktodas 
      Caption         =   "Todas las Areas"
      Height          =   195
      Left            =   195
      TabIndex        =   6
      Top             =   75
      Width           =   1710
   End
   Begin MSMask.MaskEdBox txtfecha 
      Height          =   360
      Left            =   9195
      TabIndex        =   2
      Top             =   255
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   635
      _Version        =   393216
      ForeColor       =   8388608
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Frame FraAreas 
      Height          =   645
      Left            =   105
      TabIndex        =   0
      Top             =   90
      Width           =   7320
      Begin Sicmact.TxtBuscar txtBuscarArea 
         Height          =   330
         Left            =   750
         TabIndex        =   3
         Top             =   210
         Width           =   1110
         _ExtentX        =   1958
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
         ForeColor       =   8388608
      End
      Begin VB.Label lblAreaDesc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   1875
         TabIndex        =   5
         Top             =   210
         Width           =   5295
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Areas :"
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
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   270
         Width           =   615
      End
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Fecha :"
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
      Height          =   195
      Left            =   8460
      TabIndex        =   1
      Top             =   315
      Width           =   660
   End
End
Attribute VB_Name = "frmCajaGenDepCheques"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oDocRec As NDocRec
Dim oOpe As DOperacion
Dim lsCtaDebe As String
Dim lsCtaHaber As String
Dim objPista As COMManejador.Pista 'ARLO20170217

Private Sub chktodas_Click()
txtBuscarArea = ""
lblAreaDesc = ""
If chktodas.value = 1 Then
    CargaChequesNoDepositados "", ""
    FraAreas.Enabled = False
Else
    Limpiar
    FraAreas.Enabled = True
End If
End Sub
Sub CargaChequesNoDepositados(ByVal psAreaCod As String, ByVal psAgeCod As String)
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set rs = oDocRec.GetChequesNoDepositados(Mid(gsOpeCod, 3, 1), psAreaCod, psAgeCod)
Limpiar
If Not rs.EOF And Not rs.BOF Then
    Set fgCheques.Recordset = rs
    lbltotalCartera = Format(fgCheques.SumaRow(6), "#,#0.00")
End If
rs.Close
Set rs = Nothing

End Sub
Sub Limpiar()
fgCheques.Clear
fgCheques.FormaCabecera
fgCheques.Rows = 2
lbltotalCartera = "0.00"
lbltotalDepositar = "0.00"
End Sub

Private Sub cmdDepositar_Click()
Dim lsMovNro As String
Dim lnMovNro As Long
Dim oCont As NContFunciones
Dim oCaja As nCajaGeneral
On Error GoTo ErrDeposito

Set oCaja = New nCajaGeneral
Set oCont = New NContFunciones
If fgCheques.TextMatrix(1, 0) = "" Then Exit Sub
If Trim(Len(txtMovDesc)) = 0 Then
    MsgBox "Descripción de Operación no válida", vbInformation, "aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If
If Val(lbltotalDepositar) = 0 Then
    MsgBox "Seleccione algun cheque para realizar el depósito", vbInformation, "Aviso"
    Exit Sub
End If
If Len(Trim(txtBuscarBanco)) = 0 Then
    MsgBox "Seleccione cuenta en la cual realizará el depósito respectivo", vbInformation, "Aviso"
    Exit Sub
End If

'*** PEAC 20110709
If Not oCont.PermiteModificarAsiento(Format(Me.txtfecha, gsFormatoMovFecha), False) Then
   MsgBox "No se permite realizar operaciones con fecha de un Mes Contable Cerrado.", vbInformation, "Atención"
   Exit Sub
End If
'Set oCont = Nothing
'*** FIN PEAC

If MsgBox("Desea Realizar el depósito respectivo de cheques seleccionados??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    Me.cmdDepositar.Enabled = False
    lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", , txtBuscarBanco, ObjEntidadesFinancieras)
    lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
    lsMovNro = oCont.GeneraMovNro(txtfecha, gsCodAge, gsCodUser)
    oCaja.GrabaDepositoCheques lsMovNro, gsOpeCod, txtMovDesc, lsCtaDebe, txtBuscarBanco, lsCtaHaber, CCur(lbltotalDepositar), fgCheques.GetRsNew
    ImprimeAsientoContable lsMovNro
    CargaChequesNoDepositados Mid(txtBuscarArea, 1, 3), Mid(txtBuscarArea, 4, 2)
    txtBuscarBanco = ""
    lblBancoDesc = ""
    lblCuentaDesc = ""
    txtMovDesc = ""
    Me.cmdDepositar.Enabled = True
End If

Exit Sub
ErrDeposito:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
    cmdDepositar.Enabled = True
End Sub

Private Sub cmdDevolver_Click()
Dim lsMovNro As String
Dim lnMovNro As Long
Dim oCont As NContFunciones
Dim oCaja As nCajaGeneral
On Error GoTo cmdDevolverErr

Set oCaja = New nCajaGeneral
Set oCont = New NContFunciones
If fgCheques.TextMatrix(1, 0) = "" Then Exit Sub
If Trim(Len(txtMovDesc)) = 0 Then
    MsgBox "Descripción de Operación no válida", vbInformation, "aviso"
    txtMovDesc.SetFocus
    Exit Sub
End If
If Val(lbltotalDepositar) = 0 Then
    MsgBox "Seleccione algun cheque para realizar el depósito", vbInformation, "Aviso"
    Exit Sub
End If

'*** PEAC 20110709
If Not oCont.PermiteModificarAsiento(Format(Me.txtfecha, gsFormatoMovFecha), False) Then
   MsgBox "No se permite realizar operaciones con fecha de un Mes Contable Cerrado.", vbInformation, "Atención"
   Exit Sub
End If
Set oCont = Nothing
'*** FIN PEAC

If MsgBox("Desea Realizar el Devolución de Cheques seleccionados??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsCtaDebe = oOpe.EmiteOpeCta(gsOpeCod, "D", 1)
    lsCtaHaber = oOpe.EmiteOpeCta(gsOpeCod, "H")
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oCaja.GrabaDevolucionCheques lsMovNro, gsOpeCod, txtMovDesc, lsCtaDebe, txtBuscarBanco, lsCtaHaber, CCur(lbltotalDepositar), fgCheques.GetRsNew
    ImprimeAsientoContable lsMovNro
    CargaChequesNoDepositados Mid(txtBuscarArea, 1, 3), Mid(txtBuscarArea, 4, 2)
    txtBuscarBanco = ""
    lblBancoDesc = ""
    lblCuentaDesc = ""
    txtMovDesc = ""
    'ARLO20170217
    Set objPista = New COMManejador.Pista
    objPista.InsertarPista gsOpeCod, gsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", "Devolución de Cheques"
    '****
End If
Exit Sub
cmdDevolverErr:
    MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgCheques_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim lnTotal As Currency
Dim i As Integer
If fgCheques.TextMatrix(1, 0) <> "" Then
    For i = 1 To fgCheques.Rows - 1
        If fgCheques.TextMatrix(i, 1) <> "" Then
            lnTotal = lnTotal + fgCheques.TextMatrix(i, 6)
        End If
    Next
End If
lbltotalDepositar = Format(lnTotal, "#,#0.00")

End Sub

Private Sub Form_Load()
Set oOpe = New DOperacion
Me.Caption = gsOpeDesc
txtfecha = gdFecSis
CentraForm Me
Set oDocRec = New NDocRec
txtBuscarArea.rs = oOpe.GetOpeObj(gsOpeCod, "1")
txtBuscarBanco.psRaiz = "Cuentas de Bancos"
txtBuscarBanco.rs = oOpe.GetOpeObj(gsOpeCod, "2")

End Sub

Private Sub txtBuscarArea_EmiteDatos()
lblAreaDesc = txtBuscarArea.psDescripcion
If txtBuscarArea.OK Then
    CargaChequesNoDepositados Mid(txtBuscarArea, 1, 3), Mid(txtBuscarArea, 4, 2)
    If lblAreaDesc <> "" Then
        txtfecha.SetFocus
    End If
End If
End Sub

Private Sub txtBuscarBanco_EmiteDatos()
Dim oCtaIf As NCajaCtaIF
Set oCtaIf = New NCajaCtaIF
lblBancoDesc = oCtaIf.NombreIF(Mid(txtBuscarBanco, 4, 13))
lblCuentaDesc = oCtaIf.EmiteTipoCuentaIF(Mid(txtBuscarBanco, 18, 10)) + " " + txtBuscarBanco.psDescripcion
If txtBuscarBanco <> "" Then
    txtMovDesc.SetFocus
End If
End Sub

Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    cmdDepositar.SetFocus
End If
End Sub
