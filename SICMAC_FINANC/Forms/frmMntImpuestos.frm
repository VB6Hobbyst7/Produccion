VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntImpuestos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impuestos: Mantenimiento"
   ClientHeight    =   4005
   ClientLeft      =   2070
   ClientTop       =   2085
   ClientWidth     =   9840
   Icon            =   "frmMntImpuestos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   9840
   StartUpPosition =   2  'CenterScreen
   Begin MSDataGridLib.DataGrid grdImp 
      Height          =   2535
      Left            =   150
      TabIndex        =   0
      Top             =   225
      Width           =   9660
      _ExtentX        =   17039
      _ExtentY        =   4471
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   17
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
      ColumnCount     =   6
      BeginProperty Column00 
         DataField       =   "cCtaContCod"
         Caption         =   "Cuenta Contable"
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
         DataField       =   "cCtaContDesc"
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
      BeginProperty Column02 
         DataField       =   "cImpAbrev"
         Caption         =   "Abreviatura"
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
      BeginProperty Column03 
         DataField       =   "nImpTasa"
         Caption         =   "Tasa (%)"
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
      BeginProperty Column04 
         DataField       =   "cImpDestino"
         Caption         =   "Destino"
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
      BeginProperty Column05 
         DataField       =   "nCalculo"
         Caption         =   "Calculo Neto"
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
         BeginProperty Column00 
            DividerStyle    =   6
            ColumnAllowSizing=   0   'False
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   6
            ColumnWidth     =   3974.74
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            DividerStyle    =   6
            ColumnWidth     =   1049.953
         EndProperty
         BeginProperty Column03 
            Alignment       =   2
            DividerStyle    =   6
            ColumnWidth     =   840.189
         EndProperty
         BeginProperty Column04 
            Alignment       =   2
            DividerStyle    =   6
            ColumnWidth     =   720
         EndProperty
         BeginProperty Column05 
            ColumnWidth     =   1214.929
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   3960
      TabIndex        =   10
      Top             =   3510
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   7260
      TabIndex        =   12
      Top             =   3510
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   8490
      TabIndex        =   11
      Top             =   3510
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   390
      Left            =   2730
      TabIndex        =   9
      Top             =   3510
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   390
      Left            =   1500
      TabIndex        =   8
      Top             =   3510
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   270
      TabIndex        =   7
      Top             =   3510
      Width           =   1215
   End
   Begin Sicmact.TxtBuscar txtCta 
      Height          =   330
      Left            =   270
      TabIndex        =   1
      Top             =   2955
      Visible         =   0   'False
      Width           =   1725
      _ExtentX        =   3043
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
   Begin VB.TextBox txtTasa 
      Height          =   315
      Left            =   7080
      TabIndex        =   4
      Top             =   2955
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.TextBox txtAbrev 
      Height          =   315
      Left            =   5970
      TabIndex        =   3
      Top             =   2955
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtDesc 
      Enabled         =   0   'False
      Height          =   315
      Left            =   1980
      TabIndex        =   2
      Top             =   2955
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.TextBox txtDestino 
      Height          =   315
      Left            =   7890
      MaxLength       =   1
      TabIndex        =   5
      Top             =   2955
      Visible         =   0   'False
      Width           =   750
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   8490
      TabIndex        =   13
      Top             =   3510
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.TextBox txtCalculo 
      Height          =   315
      Left            =   8655
      MaxLength       =   1
      TabIndex        =   14
      Top             =   2955
      Visible         =   0   'False
      Width           =   1080
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   150
      TabIndex        =   6
      Top             =   3420
      Width           =   9645
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   180
      Top             =   2880
      Width           =   9615
   End
End
Attribute VB_Name = "frmMntImpuestos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lConsulta As Boolean
Dim lNuevo As Boolean
Dim sSql As String
Dim rsImp As ADODB.Recordset
Dim clsImp As DImpuesto
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(plConsulta As Boolean)
lConsulta = plConsulta
Me.Show 1
End Sub

Private Sub cmdImprimir_Click()
Dim sTexto As String
Dim rs1 As ADODB.Recordset
Set rs1 = New ADODB.Recordset
Dim Cta As String
Dim Desc As String
Dim Abrev As String
Set clsImp = New DImpuesto

sTexto = ""
'Set rs1 = grdImp.DataSource
Set rs1 = clsImp.CargaImpuesto(, adLockOptimistic)
sTexto = sTexto & "Cuenta Contable" & Space(10) & "Descripcion" & Space(30) & "Abrev." & Space(10) & "Tasa (%)" & oImpresora.gPrnSaltoLinea
sTexto = sTexto & String(90, "=") & oImpresora.gPrnSaltoLinea

  
   Do While Not rs1.EOF
      Cta = rs1(0)
      Desc = rs1(1)
      Abrev = rs1(2)
      
      sTexto = sTexto & rs1(0) & Space(20 - Len(Cta)) & rs1(1) & Space(45 - Len(Desc)) & rs1(2) & Space(18 - Len(Abrev)) & rs1(3) & oImpresora.gPrnSaltoLinea
      rs1.MoveNext
   Loop
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantImpuesto
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", Me.Caption & " | Se Imprimio los Impuestos "
            Set objPista = Nothing
            '*******
RSClose rs1
EnviaPrevio sTexto, "Lista de Impuestos", gnLinPage, False
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim oCta As New DCtaCont
Set clsImp = New DImpuesto
CargaImpuestos
If lConsulta Then
   cmdNuevo.Visible = False
   cmdModificar.Visible = False
   cmdEliminar.Visible = False
   cmdImprimir.Left = cmdNuevo.Left
Else
   txtCta.rs = oCta.CargaCtaCont(" cCtaContCod LIKE '2_[12]%' ")
   Set oCta = Nothing
   txtCta.EditFlex = False
   txtCta.TipoBusqueda = BuscaGrid
   txtCta.sTitulo = "Cuenta Contable"
End If

CentraForm Me
End Sub
Private Sub CargaImpuestos()
Set rsImp = clsImp.CargaImpuesto(, adLockOptimistic)
Set grdImp.DataSource = rsImp
End Sub

Private Sub cmdNuevo_Click()
txtCta.Text = ""
txtDesc.Text = ""
txtAbrev.Text = ""
txtTasa.Text = ""
txtDestino.Text = ""
ActivaBotones False
txtCta.Enabled = True
txtCta.SetFocus
lNuevo = True
End Sub

Private Sub cmdModificar_Click()
If rsImp.EOF Then
   MsgBox "No existen datos para Eliminar", vbInformation, "¡Aviso!"
   Exit Sub
End If
txtCta.Text = rsImp!cCtaContCod
txtDesc.Text = rsImp!cCtaContDesc
txtAbrev.Text = rsImp!cImpAbrev
txtTasa.Text = rsImp!nImpTasa
txtDestino.Text = rsImp!cImpDestino
txtCalculo.Text = rsImp!nCalculo
ActivaBotones False
txtCta.Enabled = False
txtAbrev.SetFocus
lNuevo = False
End Sub

Private Sub cmdEliminar_Click()
Dim sELIM As String, cCtaImp
Dim lrs   As ADODB.Recordset
If rsImp.EOF Then
   MsgBox "No existen datos para Eliminar", vbInformation, "¡Aviso!"
   Exit Sub
End If
If MsgBox(" ¿ Está seguro de Eliminar el Impuesto ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
   Dim oDoc As DDocumento
   Set oDoc = New DDocumento
   Set lrs = oDoc.CargaDocImpuesto(-1, rsImp!cCtaContCod)
   If lrs.EOF Then
      clsImp.EliminaImpuesto rsImp!cCtaContCod
      rsImp.Delete adAffectCurrent
            
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantImpuesto
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " | Cuenta Contable : " & rsImp!cCtaContCod
            Set objPista = Nothing
            '*******
   Else
      MsgBox "Impuesto esta Asignado a Documentos", vbInformation, "¡Aviso!"
   End If
   Set oDoc = Nothing
   RSClose lrs
End If
grdImp.SetFocus
End Sub

Private Sub cmdCancelar_Click()
ActivaBotones True
grdImp.SetFocus
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If Len(txtCta.Text) = 0 Then
   MsgBox "  Cuenta de impuesto no válida ...  ", vbInformation, "Error"
   txtCta.SetFocus
   Exit Function
End If

If Len(txtAbrev.Text) = 0 Then
   MsgBox "  Abreviatura de impuesto no válida ...  ", vbInformation, "Error"
   txtAbrev.SetFocus
   Exit Function
End If

If Len(txtTasa.Text) = 0 Then
   MsgBox "  Tasa de impuesto no válida ...  ", vbInformation, "Error"
   txtTasa.SetFocus
   Exit Function
End If

If Len(txtDestino.Text) = 0 Then
   MsgBox "  Condición de Destino Tasa de impuesto no válida ...  ", vbInformation, "Error"
   txtTasa.SetFocus
   Exit Function
End If

Dim prs As ADODB.Recordset
Set prs = clsImp.CargaImpuesto()
If Not prs.EOF Then
   prs.Find "cImpAbrev = '" & Me.txtAbrev & "'"
   If Not prs.EOF Then
        If lNuevo Then
           MsgBox "La abreviatura ya está relacionada con otro Impuesto...      ", vbInformation, "¡Aviso!"
           RSClose prs
           Exit Function
        Else
           If prs!cCtaContCod <> txtCta Then
              MsgBox "La abreviatura ya está relacionada con otro Impuesto...      ", vbInformation, "¡Aviso!"
              RSClose prs
              Exit Function
           End If
        End If
   End If
End If


ValidaDatos = True

End Function

Private Sub cmdAceptar_Click()
On Error GoTo AceptarErr
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro de grabar datos del impuesto ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   Select Case lNuevo
      Case True
            clsImp.InsertaImpuesto txtCta.Text, txtAbrev.Text, Format(txtTasa.Text, gsFormatoNumeroView), txtDestino, Me.txtCalculo.Text, gsMovNro
                  
      Case False
            clsImp.ActualizaImpuesto txtCta.Text, txtAbrev.Text, Format(txtTasa.Text, gsFormatoNumeroView), txtDestino, Me.txtCalculo.Text, gsMovNro
   End Select
   CargaImpuestos
   rsImp.Find "cCtaContCod = '" & txtCta.Text & "'", , , 1
End If
            'ARLO20170208
            Dim lsAccion As String
            Set objPista = New COMManejador.Pista
            If lNuevo Then
            lsAccion = "1"
            Else:
            lsAccion = "2"
            End If
            gsOpeCod = LogPistaMantImpuesto
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, Me.Caption & " | Cuenta Contable : " & txtCta.Text & " | Descripcion : " & txtDesc.Text
            Set objPista = Nothing
            '*******
ActivaBotones True
grdImp.SetFocus
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsImp
Set clsImp = Nothing
End Sub

Private Sub grdImp_GotFocus()
grdImp.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdImp_HeadClick(ByVal ColIndex As Integer)
If Not rsImp Is Nothing Then
   If Not rsImp.EOF Then
      rsImp.Sort = grdImp.Columns(ColIndex).DataField
   End If
End If
End Sub

Private Sub grdImp_LostFocus()
grdImp.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub txtCta_EmiteDatos()
txtDesc = txtCta.psDescripcion
If txtDesc <> "" And txtAbrev.Visible Then
   txtAbrev.SetFocus
End If
End Sub

Private Sub txtCta_GotFocus()
txtCta.BackColor = "&H00F5FFDD"
End Sub
Private Sub txtCta_LostFocus()
txtCta.BackColor = "&H80000005"
End Sub
Private Sub txtCta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtabrev_GotFocus()
txtAbrev.BackColor = "&H00F5FFDD"
End Sub
Private Sub txtabrev_LostFocus()
txtAbrev.BackColor = "&H80000005"
End Sub
Private Sub TxtAbrev_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   txtTasa.SetFocus
End If
End Sub
Private Sub txttasa_GotFocus()
txtTasa.BackColor = "&H00F5FFDD"
End Sub
Private Sub txttasa_LostFocus()
txtTasa.BackColor = "&H80000005"
End Sub
Private Sub txtTasa_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtTasa, KeyAscii, 6, 2)
If KeyAscii = 13 Then
   txtDestino.SetFocus
End If
End Sub
Private Sub TxtDestino_GotFocus()
txtDestino.BackColor = "&H00F5FFDD"
End Sub
Private Sub txtDestino_LostFocus()
txtDestino.BackColor = "&H80000005"
End Sub
Private Sub txtDestino_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
Else
   If KeyAscii <> 8 Then
      KeyAscii = IIf(InStr("01", Chr(KeyAscii)) > 0, KeyAscii, 0)
   End If
End If
End Sub

Sub ActivaBotones(lActiva As Boolean)
If lActiva Then
   grdImp.Height = 3120
Else
   grdImp.Height = 2640
End If
cmdNuevo.Visible = lActiva
cmdModificar.Visible = lActiva
cmdEliminar.Visible = lActiva
cmdImprimir.Visible = lActiva
cmdSalir.Visible = lActiva
cmdAceptar.Visible = Not lActiva
cmdCancelar.Visible = Not lActiva
txtCta.Visible = Not lActiva
txtDesc.Visible = Not lActiva
txtAbrev.Visible = Not lActiva
txtTasa.Visible = Not lActiva
txtDestino.Visible = Not lActiva
txtCalculo.Visible = Not lActiva
End Sub

