VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntCtaContPend 
   Caption         =   "Cuentas de Pendientes: Mantenimiento"
   ClientHeight    =   4140
   ClientLeft      =   2085
   ClientTop       =   2100
   ClientWidth     =   8715
   Icon            =   "frmMntCtaContPend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8715
   Begin MSDataGridLib.DataGrid grdImp 
      Height          =   3090
      Left            =   180
      TabIndex        =   0
      Top             =   225
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5450
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
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
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   4
         ScrollBars      =   2
         BeginProperty Column00 
            DividerStyle    =   6
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   2220.094
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   6
            ColumnWidth     =   5625.071
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraControl 
      Height          =   585
      Left            =   300
      TabIndex        =   5
      Top             =   3360
      Width           =   8085
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   390
         Left            =   120
         TabIndex        =   1
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   390
         Left            =   1380
         TabIndex        =   2
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   390
         Left            =   2640
         TabIndex        =   3
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         CausesValidation=   0   'False
         Height          =   390
         Left            =   6720
         TabIndex        =   4
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         CausesValidation=   0   'False
         Height          =   390
         Left            =   6720
         TabIndex        =   9
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   5460
         TabIndex        =   8
         Top             =   150
         Visible         =   0   'False
         Width           =   1215
      End
   End
   Begin VB.TextBox txtDesc 
      Height          =   315
      Left            =   2670
      TabIndex        =   7
      Top             =   2925
      Width           =   5715
   End
   Begin Sicmact.TxtBuscar txtCta 
      Height          =   330
      Left            =   420
      TabIndex        =   6
      Top             =   2925
      Width           =   2235
      _ExtentX        =   3942
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
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   180
      Top             =   2850
      Width           =   8415
   End
End
Attribute VB_Name = "frmMntCtaContPend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nPos As Variant
Dim sImp As String
Dim sSql As String
Dim nOPE As Integer
Dim rs As New ADODB.Recordset
Dim YaTrans As Boolean
Dim vCuenta As String
Dim oCon As DConecta

'ARLO20170208****
Dim objPista As COMManejador.Pista
Dim lsPalabra, lsAccion As String
'************'

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Set oCon = New DConecta
oCon.AbreConexion
CargaDatos
nOPE = 0

Dim clsCta As DCtaCont
Set clsCta = New DCtaCont
txtCta.rs = clsCta.CargaCtaCont(" cCtaContCod LIKE '[12]_[^0]%' ", "CtaCont")
txtCta.lbUltimaInstancia = False
txtCta.EditFlex = False
txtCta.TipoBusqueda = BuscaGrid

End Sub

Private Sub CargaDatos()
sImp = "SELECT cctacontCod, cCtaContDesc " & _
       "FROM   CtaContPend "
Set rs = oCon.CargaRecordSet(sImp, adLockOptimistic)
Set grdImp.DataSource = rs
End Sub

Private Sub cmdNuevo_Click()
txtCta.Text = ""
txtDesc.Text = ""
Botones2
txtCta.SetFocus
nOPE = 1
End Sub

Private Sub cmdModificar_Click()
vCuenta = rs!cCtaContCod
txtCta.Text = rs!cCtaContCod
txtDesc.Text = rs!cCtaContDesc
nPos = rs.Bookmark
Botones2
txtCta.SetFocus
nOPE = 2
End Sub

Private Sub cmdEliminar_Click()
Dim sELIM As String, cCtaImp
On Error GoTo ErrElimina
cCtaImp = rs!cCtaContCod
If MsgBox(" ¿ Está seguro de eliminar el Cuenta de Pendiente ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
   sELIM = "DELETE from  CtaContPend where cCtaContCod = '" & cCtaImp & "' "
   If YaTrans Then
      oCon.RollbackTrans
      YaTrans = False
   End If
   YaTrans = True
   oCon.BeginTrans
   oCon.Ejecutar sELIM
   oCon.CommitTrans
               'ARLO20170208
            Set objPista = New COMManejador.Pista
            txtDesc.Text = rs!cCtaContDesc
            gsOpeCod = LogPistaManCtaPend
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se Elimino la Cuenta Pendiente Cuenta :" & cCtaImp & " Descripcion : " & txtDesc.Text
            Set objPista = Nothing
            '*******
   rs.Delete adAffectCurrent
   YaTrans = False
End If
grdImp.SetFocus
Exit Sub
ErrElimina:
    MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdCancelar_Click()
Botones1
grdImp.SetFocus
End Sub

Private Sub cmdAceptar_Click()
Dim rd As New ADODB.Recordset
Dim sINS As String
Dim sDUP As String
On Error GoTo AceptarErr

If Len(txtCta.Text) = 0 Then
   MsgBox "  Cuenta de impuesto no válida ...  ", vbCritical, "Error"
   txtCta.SetFocus
   Exit Sub
End If

If Len(txtDesc.Text) = 0 Then
   MsgBox "  Ingresar Descripción de la Pendiente ...  ", vbCritical, "Error"
   txtDesc.SetFocus
   Exit Sub
End If

If txtCta.Text <> vCuenta Then
   sDUP = "Select * FROM CtaContPend where cCtaContCod = '" & txtCta.Text & "'"
   Set rd = oCon.CargaRecordSet(sDUP)
   If Not rd.EOF Then
      MsgBox "  El Cuenta de Pendiente ya está registrado ... ", vbInformation, "Aviso"
      txtCta.SetFocus
      Exit Sub
   End If
End If

If MsgBox(" ¿ Seguro de grabar datos del Cuenta ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   Select Case nOPE
      Case 1
           sINS = "INSERT into CtaContPend (cCtaContCod, cCtaContDesc) " & _
                  "values ('" & txtCta.Text & "' , '" & txtDesc.Text & "' )"
      Case 2
           sINS = "UPDATE CtaContPend " & _
                  "SET cCtaContCod = '" & txtCta.Text & "' , " & _
                  "    cCtaContDesc = '" & txtDesc.Text & "' " & _
                  " where cCtaContcod = '" & vCuenta & "' "
   End Select
                 
   YaTrans = True
   oCon.BeginTrans
   oCon.Ejecutar sINS
   oCon.CommitTrans
   YaTrans = False
   CargaDatos
   Select Case nOPE
      Case 1
         rs.Find "cctacontcod = '" & txtCta.Text & "'", , , 1
      Case 2
         rs.Bookmark = nPos
   End Select
End If
cmdCancelar_Click
grdImp.SetFocus
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            If (nOPE = 1) Then
            lsPalabra = "Agrego"
            lsAccion = "1"
            Else: lsPalabra = "Modifico"
            objPista = "2"
            End If
            gsOpeCod = LogPistaManCtaPend
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, "Se " & lsPalabra & " Cuenta Pendiente Cuenta :" & txtCta.Text & " Descripcion : " & txtDesc.Text
            Set objPista = Nothing
            '*******
Exit Sub
AceptarErr:
   If YaTrans Then
      oCon.RollbackTrans
      YaTrans = False
   End If
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdBuscar_Click()
Dim clsBusca As New DescObjeto.ClassDescObjeto
'ManejaBoton False
Botones1
clsBusca.BuscarDato rs, 0, "Cuenta Contable"
'nOrdenCta = clsBusca.gnOrdenBusca
If clsBusca.lbOk Then
   txtCta = clsBusca.gsSelecCod
   txtDesc = clsBusca.gsSelecDesc
   txtDesc.SetFocus
End If
Set clsBusca = Nothing
Botones2
End Sub

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing
End Sub

Private Sub grdImp_GotFocus()
grdImp.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdImp_LostFocus()
grdImp.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub txtCta_EmiteDatos()
txtDesc = txtCta.psDescripcion
If txtDesc <> "" Then
   cmdAceptar.SetFocus
End If
End Sub

Private Sub txtCta_GotFocus()
txtCta.BackColor = "&H00F5FFDD"
End Sub
Private Sub txtCta_LostFocus()
txtCta.BackColor = "&H80000005"
End Sub

Sub Botones1()
grdImp.Height = 3090
cmdNuevo.Visible = True
cmdModificar.Visible = True
cmdEliminar.Visible = True
cmdSalir.Visible = True
cmdAceptar.Visible = False
cmdCancelar.Visible = False
End Sub

Sub Botones2()
grdImp.Height = 2640
cmdNuevo.Visible = False
cmdModificar.Visible = False
cmdEliminar.Visible = False
cmdSalir.Visible = False
cmdAceptar.Visible = True
cmdCancelar.Visible = True
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub
