VERSION 5.00
Begin VB.Form frmCajeroRegFaltSob 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4680
   ClientLeft      =   2970
   ClientTop       =   2430
   ClientWidth     =   5385
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajeroRegFaltSob.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   5385
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   4095
      TabIndex        =   1
      Top             =   4215
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   4140
      Left            =   90
      TabIndex        =   0
      Top             =   0
      Width           =   5205
      Begin VB.TextBox txtMonto 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
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
         Height          =   330
         Left            =   3495
         TabIndex        =   8
         Text            =   "0.00"
         Top             =   3720
         Width           =   1500
      End
      Begin VB.TextBox txtMovdesc 
         Height          =   660
         Left            =   180
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   2910
         Width           =   4845
      End
      Begin Sicmact.TxtBuscar txtBuscarCajero 
         Height          =   330
         Left            =   855
         TabIndex        =   5
         Top             =   195
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
         ForeColor       =   8388608
      End
      Begin Sicmact.FlexEdit fgFalt 
         Height          =   1815
         Left            =   180
         TabIndex        =   6
         Top             =   960
         Width           =   4830
         _ExtentX        =   8520
         _ExtentY        =   3201
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Fecha Reg-Total-Cubierto-Pendiente-nMovNroReg"
         EncabezadosAnchos=   "350-1000-1000-1000-1000-0"
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
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R-R-R-L"
         FormatosEdit    =   "0-0-2-2-2-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Monto :"
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
         Height          =   210
         Left            =   2580
         TabIndex        =   9
         Top             =   3780
         Width           =   615
      End
      Begin VB.Label lblNomCajero 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   300
         Left            =   855
         TabIndex        =   4
         Top             =   555
         Width           =   4080
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cajero :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   180
         TabIndex        =   3
         Top             =   240
         Width           =   630
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         BackStyle       =   1  'Opaque
         Height          =   360
         Left            =   2355
         Top             =   3705
         Width           =   2655
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   2880
      TabIndex        =   2
      Top             =   4215
      Width           =   1230
   End
End
Attribute VB_Name = "frmCajeroRegFaltSob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCajero As nCajero

Private Sub cmdAceptar_Click()
Dim oCont As NContFunciones
Set oCont = New NContFunciones
Dim lnMovNroReg As Long
Dim lsMovNro As String
Dim lnImporte As Currency
If Valida = False Then Exit Sub

lnMovNroReg = fgFalt.TextMatrix(fgFalt.Row, 5)
lnImporte = CCur(txtMonto)
If MsgBox("Desea Realizar la Operación seleccionada??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    lsMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    oCajero.GrabaRegularizaFaltante lsMovNro, gsOpeCod, txtMovdesc, txtMonto, txtMonto, lnMovNroReg
    
    Dim oImp As NContImprimir
    Dim lsTexto As String
    Dim lbReimp As Boolean
    Set oImp = New NContImprimir
    
    lbReimp = True
    Do While lbReimp
        oImp.ImprimeBoletaGeneral gsOpeDescPadre, gsOpeDescHijo, gsOpeCod, txtMonto, gsNomAge, lsMovNro, _
                 sLpt, "[" & txtBuscarCajero & "] " & lblNomCajero, "", "", Trim(txtMovdesc)
    
        If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            lbReimp = False
        End If
    Loop
    Set oImp = Nothing
    CargaFaltantes
    txtMovdesc = ""
    txtMonto = "0.00"
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgFalt_RowColChange()
If fgFalt.TextMatrix(1, 0) <> "" Then
    txtMonto = Format(fgFalt.TextMatrix(fgFalt.Row, 4), "#,#0.00")
End If
End Sub

Private Sub Form_Load()
Dim oGen As DGeneral
Set oGen = New DGeneral
Set oCajero = New nCajero
Me.Caption = gsOpeDesc
CentraForm Me
txtBuscarCajero.psRaiz = "CAJEROS"
txtBuscarCajero.rs = oGen.GetUserAreaAgencia(gsCodArea, gsCodAge)
End Sub
Private Sub txtBuscarCajero_EmiteDatos()
lblNomCajero = Me.txtBuscarCajero.psDescripcion
If lblNomCajero = "" Then
    txtBuscarCajero.SetFocus
Else
    CargaFaltantes
End If
End Sub
Sub CargaFaltantes()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

fgFalt.Clear
fgFalt.FormaCabecera
fgFalt.Rows = 2
txtMovdesc = ""
txtMonto = "0.00"
Set rs = oCajero.GetFaltanteCajero(gsOpeCod, gsCodUser, gsCodAge)
If Not rs.EOF And Not rs.BOF Then
    Set fgFalt.Recordset = rs
    fgFalt.SetFocus
Else
    MsgBox "No se encontraron datos para el Usuario Seleccionado", vbInformation, "Aviso"
End If
rs.Close: Set rs = Nothing
End Sub
Private Sub txtMonto_GotFocus()
fEnfoque txtMonto
End Sub
Private Sub txtMonto_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMonto, KeyAscii)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub
Private Sub txtMonto_LostFocus()
If Len(Trim(Me.txtMonto)) = 0 Then txtMonto = 0
txtMonto = Format(txtMonto, "#,#0.00")
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    txtMonto.SetFocus
End If
End Sub
Function Valida() As Boolean
Valida = True
If txtBuscarCajero = "" Then
   MsgBox "Usuario no ha sido seleccionado", vbInformation, "Aviso"
   txtBuscarCajero.SetFocus
   Valida = False
   Exit Function
End If
If fgFalt.TextMatrix(1, 0) = "" Then
    MsgBox "Usuario no Posee faltantes pendiente", vbInformation, "Aviso"
    txtBuscarCajero.SetFocus
    Valida = False
    Exit Function
End If
If Len(Trim(txtMovdesc)) = "" Then
    MsgBox "Descripcióin de Operación no ingresada", vbInformation, "Aviso"
    txtMovdesc.SetFocus
    Valida = False
    Exit Function
End If
If Val(txtMonto) = 0 Then
    MsgBox "Monto de Operación no ingresado", vbInformation, "Aviso"
    txtMonto.SetFocus
    Valida = False
    Exit Function
End If
If CCur(txtMonto) > CCur(fgFalt.TextMatrix(fgFalt.Row, 4)) Then
    MsgBox "Monto de Operación no puede ser mayor a pendiente Selecciondo", vbInformation, "Aviso"
    txtMonto = Format(fgFalt.TextMatrix(fgFalt.Row, 4), "#,#0.00")
    txtMonto.SetFocus
    Valida = False
    Exit Function
End If
End Function
