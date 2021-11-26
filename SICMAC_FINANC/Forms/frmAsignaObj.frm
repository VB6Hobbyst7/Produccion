VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmAsignaObj 
   ClientHeight    =   5535
   ClientLeft      =   1200
   ClientTop       =   1515
   ClientWidth     =   8655
   Icon            =   "frmAsignaObj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   8655
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNewObjeto 
      Appearance      =   0  'Flat
      Caption         =   "&+"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   150
      TabIndex        =   10
      ToolTipText     =   "Nuevo Objeto"
      Top             =   1275
      Width           =   285
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   7200
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4860
      Width           =   1110
   End
   Begin VB.Frame fraTitulo 
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
      Height          =   825
      Left            =   120
      TabIndex        =   11
      Top             =   75
      Width           =   8415
      Begin VB.TextBox txtDesc 
         BackColor       =   &H00F0FFFF&
         Enabled         =   0   'False
         Height          =   375
         Left            =   1800
         MaxLength       =   60
         TabIndex        =   13
         Top             =   300
         Width           =   6375
      End
      Begin VB.TextBox txtCod 
         BackColor       =   &H00F0FFFF&
         BeginProperty DataFormat 
            Type            =   0
            Format          =   "0"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
         Enabled         =   0   'False
         Height          =   375
         Left            =   180
         MaxLength       =   20
         TabIndex        =   12
         Top             =   300
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   5940
      TabIndex        =   7
      ToolTipText     =   "Asignar Objeto a Cuenta"
      Top             =   4860
      Width           =   1155
   End
   Begin MSDataGridLib.DataGrid grdObjs 
      Height          =   2610
      Left            =   120
      TabIndex        =   9
      Top             =   1260
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   4604
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   2
      RowHeight       =   15
      RowDividerStyle =   6
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cObjetoCod"
         Caption         =   "Código"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cObjetoDesc"
         Caption         =   "Descripción"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         BeginProperty Column00 
            DividerStyle    =   6
            ColumnWidth     =   1709.858
         EndProperty
         BeginProperty Column01 
            ColumnWidth     =   6134.74
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame2 
      Caption         =   "  Objeto   "
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
      Height          =   1425
      Left            =   120
      TabIndex        =   15
      Top             =   3975
      Width           =   8415
      Begin VB.TextBox txtTipo 
         Alignment       =   2  'Center
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
         Height          =   315
         Left            =   1545
         MaxLength       =   1
         TabIndex        =   4
         Top             =   930
         Width           =   300
      End
      Begin VB.TextBox txtImpre 
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
         Height          =   315
         Left            =   3825
         TabIndex        =   6
         Top             =   930
         Width           =   1755
      End
      Begin VB.TextBox txtObjetoDesc 
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
         ForeColor       =   &H8000000D&
         Height          =   315
         Left            =   2340
         TabIndex        =   1
         Top             =   300
         Width           =   5865
      End
      Begin VB.TextBox txtFiltro 
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
         Height          =   315
         Left            =   2025
         TabIndex        =   5
         Top             =   930
         Width           =   1635
      End
      Begin VB.TextBox txtSec 
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
         Height          =   315
         Left            =   270
         MaxLength       =   1
         TabIndex        =   2
         Top             =   930
         Width           =   495
      End
      Begin VB.TextBox txtNiv 
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
         Height          =   315
         Left            =   930
         MaxLength       =   2
         TabIndex        =   3
         Top             =   930
         Width           =   435
      End
      Begin Sicmact.TxtBuscar txtObjetoCod 
         Height          =   330
         Left            =   270
         TabIndex        =   0
         Top             =   300
         Width           =   2055
         _extentx        =   3995
         _extenty        =   582
         appearance      =   1
         appearance      =   1
         font            =   "frmAsignaObj.frx":030A
         appearance      =   1
         stitulo         =   ""
      End
      Begin VB.Label lbl05 
         Caption         =   "Tipo"
         Height          =   195
         Left            =   1530
         TabIndex        =   20
         Top             =   735
         Width           =   375
      End
      Begin VB.Label lbl04 
         Caption         =   "Filtro de Impresión"
         Height          =   195
         Left            =   3825
         TabIndex        =   19
         Top             =   735
         Width           =   1395
      End
      Begin VB.Label lbl03 
         Caption         =   "Filtro de Selección"
         Height          =   195
         Left            =   2025
         TabIndex        =   18
         Top             =   735
         Width           =   1395
      End
      Begin VB.Label lbl02 
         Caption         =   "Nivel"
         Height          =   195
         Left            =   930
         TabIndex        =   17
         Top             =   735
         Width           =   375
      End
      Begin VB.Label lbl01 
         Caption         =   "Orden"
         Height          =   195
         Left            =   285
         TabIndex        =   16
         Top             =   735
         Width           =   495
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Relación de Objetos"
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
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   975
      Width           =   1995
   End
End
Attribute VB_Name = "frmAsignaObj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim rsObj   As New ADODB.Recordset
Dim sCtaCod As String, sCtaDesc As String, sTipo As String
Dim nOrdenObj As Integer
Dim rsCtaObj As ADODB.Recordset
Dim grdActivo As Boolean, lPressMouse As Boolean

Dim clsObjeto  As DObjeto
Dim nTopSec As Integer

'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(psCod As String, psDesc As String, psTipo As String, Optional pntopsec As Integer = 0)
nTopSec = pntopsec
sCtaCod = psCod
sCtaDesc = psDesc
sTipo = psTipo
Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
If AsignaObjeto Then
   Unload Me
End If
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub cmdNewObjeto_Click()
frmMntObjetos.Show 1
grdObjs.SetFocus
Set rsObj = clsObjeto.CargaObjeto()
Set grdObjs.DataSource = rsObj
txtObjetoCod.rs = rsObj
End Sub

Private Sub Form_Load()
nOrdenObj = 1  'Inicialmente los objetos se ordenan por descripcion
txtCod.Text = sCtaCod
txtDesc.Text = sCtaDesc
Select Case sTipo
   Case "0"
        fraTitulo.Caption = " Cuenta Contable "
        lbl05.Caption = "Opc"
        
   Case "2"
        fraTitulo.Caption = " Operación "
        lbl04.Visible = False
        txtImpre.Visible = False
        
   Case "3"
        fraTitulo.Caption = " Proveedor "
        lbl01.Visible = False
        lbl02.Visible = False
        lbl03.Visible = False
        lbl04.Visible = False
        lbl05.Visible = False
        txtSec.Visible = False
        txtNiv.Visible = False
        txtFiltro.Visible = False
        txtImpre.Visible = False
        txtTipo.Visible = False
        txtSec.Text = "0"
        txtNiv.Text = "0"
        txtTipo.Text = "0"
End Select

frmAsignaObj.Caption = fraTitulo.Caption & ": Asignación de Objetos"

Set clsObjeto = New DObjeto
Set rsObj = clsObjeto.CargaObjeto()
Set grdObjs.DataSource = rsObj

txtObjetoCod.rs = rsObj
txtObjetoCod.EditFlex = False
txtObjetoCod.lbUltimaInstancia = False
txtObjetoCod.TipoBusqueda = BuscaDatoEnGrid

txtNiv = "1"
If sTipo = "0" Then
   txtSec = nTopSec
End If

End Sub

Private Sub grdObjs_DblClick()
If grdActivo Then
   AsignaObjeto
End If
End Sub

Private Sub grdObjs_GotFocus()
grdObjs.MarqueeStyle = dbgHighlightRow
grdActivo = True
End Sub

Private Sub grdObjs_KeyPress(KeyAscii As Integer)
If grdActivo Then
   If KeyAscii = 13 Then
      cmdAceptar_Click
      KeyAscii = 0
   End If
End If
End Sub
Private Sub grdObjs_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 40 Or KeyCode = 38 Or KeyCode = 34 Or KeyCode = 33 Then
   If grdActivo Then
      txtObjetoCod.Text = rsObj!cObjetoCod
   End If
   txtObjetoDesc.Text = rsObj!cObjetoDesc
End If
End Sub
Private Sub grdObjs_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
lPressMouse = True
End Sub

Private Sub grdObjs_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
lPressMouse = True
End Sub

Private Sub grdObjs_LostFocus()
grdActivo = False
grdObjs.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub grdObjs_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
If grdActivo Or lPressMouse Then
    If Not rsObj.EOF Then
    txtObjetoCod.Text = rsObj!cObjetoCod
    lPressMouse = False
    txtObjetoDesc.Text = rsObj!cObjetoDesc
    End If
End If
End Sub

Private Sub txtObjetoCod_Change()
Dim Criterio As String
If Len(Trim(txtObjetoCod.Text)) > 0 Then
   Criterio = "cObjetoCod LIKE '" & txtObjetoCod.Text & "*'"
   BuscaDato Criterio, rsObj, 1, False
   txtObjetoDesc.Text = rsObj!cObjetoDesc
End If
End Sub

Private Sub txtObjetoCod_EmiteDatos()
txtObjetoDesc = txtObjetoCod.psDescripcion
If txtObjetoDesc <> "" Then
   If sTipo = "3" Then
        If cmdAceptar.Visible Then
            cmdAceptar.SetFocus
        End If
   Else
        If txtSec.Visible Then
            txtSec.SetFocus
        End If
   End If
End If
End Sub

Private Sub txtObjetoCod_GotFocus()
txtObjetoCod.SelStart = Len(txtObjetoCod.Text)
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If Len(txtSec) = 0 Then
   MsgBox " Secuencia no puede estar vacio ...! ", vbCritical, "Error de Asignación"
   txtSec.SetFocus
   Exit Function
End If
If Len(txtNiv) = 0 Then
   MsgBox " Nivel no puede ser 0 ...! ", vbCritical, "Error de Asignación"
   txtNiv.SetFocus
   Exit Function
End If

If sTipo = "0" Then
  If Val(txtSec) > nTopSec Then
     MsgBox "El Orden no puede ser mayor que " & nTopSec & " ...!", vbInformation, "Orden no válido"
     txtSec.SetFocus
     Exit Function
  End If
End If
ValidaDatos = True
End Function

Private Function AsignaObjeto() As Boolean
On Error GoTo AsignaErr
AsignaObjeto = False
If Not ValidaDatos() Then
   Exit Function
End If
If MsgBox("  ¿ Seguro de Asignar Objeto  ?  ", vbOKCancel, "Confirmación de Asignación") = vbOk Then
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   
   Select Case sTipo
      Case "0"
         Dim clsCtaCont As New DCtaCont
         clsCtaCont.InsertaCtaObj sCtaCod, txtSec, rsObj!cObjetoCod, txtNiv, txtFiltro, txtImpre, gsMovNro, rsObj!nObjetoNiv
         clsCtaCont.InsertaCtaObjFiltro sCtaCod, txtSec, rsObj!cObjetoCod, "", gsMovNro
         Set clsCtaCont = Nothing
         Dim oCon As New DConecta
         Dim sSql As String
      Case "2"
         Dim clsOperacion As New DOperacion
         clsOperacion.InsertaOpeObj sCtaCod, txtSec, rsObj!cObjetoCod, "", Val(txtNiv.Text), txtFiltro, gsMovNro
         Set clsOperacion = Nothing
   End Select
            'ARLO20170208
            Dim lsObjeCod As String
            Set objPista = New COMManejador.Pista
            'gsOpeCod = LogPistaMantClasifOperacion
            lsObjeCod = rsObj!cObjetoCod
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, gIngresarSistema, "Se Asigno el |Obejto : " & lsObjeCod & " | " & txtObjetoDesc.Text & " a la Operacion : " & txtDesc.Text
            Set objPista = Nothing
            '*******
   AsignaObjeto = True
End If
Exit Function
AsignaErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Function


Private Sub txtSec_GotFocus()
fEnfoque txtSec
End Sub

Private Sub txtSec_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtNiv.SetFocus
End If
End Sub

Private Sub txtFiltro_GotFocus()
fEnfoque txtFiltro
End Sub

Private Sub txtFiltro_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 13 And KeyAscii <> 8 Then
   If InStr("_0123456789^[]%", Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End If
If KeyAscii = 13 Then
   If txtImpre.Visible Then
      txtImpre.SetFocus
   Else
      cmdAceptar.SetFocus
   End If
End If
End Sub
Private Sub txtImpre_GotFocus()
fEnfoque txtImpre
End Sub

Private Sub txtImpre_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii <> 13 And KeyAscii <> 8 Then
   If InStr("_X", Chr(KeyAscii)) = 0 Then
      KeyAscii = 0
   End If
End If
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub

Private Sub txtNiv_GotFocus()
txtNiv.SelStart = Len(txtNiv.Text)
End Sub

Private Sub txtNiv_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   If txtTipo.Visible Then
      txtTipo.SetFocus
   Else
      If txtFiltro.Visible Then
         txtFiltro.SetFocus
      Else
         cmdAceptar.SetFocus
      End If
   End If
End If
End Sub

Private Sub txtTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   If InStr("01", txtTipo.Text) <> 0 Then
      If txtFiltro.Visible Then
         txtFiltro.SetFocus
      Else
         cmdAceptar.SetFocus
      End If
   Else
      MsgBox "Valor no válido ...", vbCritical, "Error"
   End If
End If
End Sub
