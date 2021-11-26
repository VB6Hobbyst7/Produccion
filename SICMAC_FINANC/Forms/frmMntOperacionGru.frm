VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntOperacionGru 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Operaciones: Grupos: "
   ClientHeight    =   4140
   ClientLeft      =   2070
   ClientTop       =   2085
   ClientWidth     =   8715
   Icon            =   "frmMntOperacionGru.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4140
   ScaleWidth      =   8715
   Begin MSDataGridLib.DataGrid grdOpe 
      Height          =   3210
      Left            =   120
      TabIndex        =   0
      Top             =   210
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   5662
      _Version        =   393216
      AllowUpdate     =   0   'False
      BorderStyle     =   0
      HeadLines       =   2
      RowHeight       =   19
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
         DataField       =   "cOpeGruCod"
         Caption         =   "Grupo"
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
         DataField       =   "cOpeGruDesc"
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
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            DividerStyle    =   6
            ColumnAllowSizing=   0   'False
            ColumnWidth     =   1800
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   6
            ColumnWidth     =   6045.166
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   390
      Left            =   1440
      TabIndex        =   4
      Top             =   3570
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   390
      Left            =   210
      TabIndex        =   3
      Top             =   3570
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   390
      Left            =   2670
      TabIndex        =   5
      Top             =   3570
      Width           =   1215
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   390
      Left            =   6000
      TabIndex        =   7
      Top             =   3570
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   390
      Left            =   3900
      TabIndex        =   6
      Top             =   3570
      Width           =   1215
   End
   Begin VB.TextBox txtOpeDesc 
      Height          =   315
      Left            =   2250
      MaxLength       =   50
      TabIndex        =   2
      Top             =   3030
      Visible         =   0   'False
      Width           =   6075
   End
   Begin VB.TextBox txtOpeCod 
      Height          =   315
      Left            =   420
      MaxLength       =   2
      TabIndex        =   1
      Top             =   3030
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   7230
      TabIndex        =   9
      Top             =   3570
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      CausesValidation=   0   'False
      Height          =   390
      Left            =   7230
      TabIndex        =   8
      Top             =   3570
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   540
      Left            =   120
      TabIndex        =   10
      Top             =   3480
      Width           =   8415
   End
   Begin VB.Shape Shape1 
      Height          =   465
      Left            =   120
      Top             =   2955
      Width           =   8415
   End
End
Attribute VB_Name = "frmMntOperacionGru"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lConsulta As Boolean
Dim lNuevo As Boolean

Dim rsGru As ADODB.Recordset
Dim clsOpe As DOperacion
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

Set clsOpe = New DOperacion

sTexto = ""
Set rs1 = clsOpe.CargaOpeGru(, adLockOptimistic)

sTexto = sTexto & "GRUPO" & Space(10) & "DESCRIPCION" & Space(30) & oImpresora.gPrnSaltoLinea
sTexto = sTexto & String(50, "=") & oImpresora.gPrnSaltoLinea

  
   Do While Not rs1.EOF
      Cta = rs1(0)
      Desc = rs1(1)
     
      
      sTexto = sTexto & rs1(0) & Space(15 - Len(Cta)) & rs1(1) & oImpresora.gPrnSaltoLinea
      rs1.MoveNext
   Loop
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantClasifOperacion
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio los Gupos de Operaciones "
            Set objPista = Nothing
            '*******
RSClose rs1
EnviaPrevio sTexto, "Lista de Impuestos", gnLinPage, False

End Sub

Private Sub cmdModificar_Click()
ActivaBotones False
lNuevo = False
txtOpeCod = rsGru!cOpeGruCod
txtOpeDesc = rsGru!cOpeGruDesc
txtOpeCod.Enabled = False
txtOpeDesc.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
If lConsulta Then
   Me.Caption = Me.Caption & "Consulta"
   cmdNuevo.Visible = False
   cmdModificar.Visible = False
   cmdEliminar.Visible = False
   cmdImprimir.Left = cmdNuevo.Left
Else
   Me.Caption = Me.Caption & "Mantenimiento"
End If
Set clsOpe = New DOperacion
CargaGrupos
End Sub
Private Sub CargaGrupos()
Set rsGru = clsOpe.CargaOpeGru(, adLockOptimistic)
Set grdOpe.DataSource = rsGru
End Sub

Private Sub cmdNuevo_Click()
ActivaBotones False
lNuevo = True
txtOpeCod = ""
txtOpeDesc = ""
txtOpeCod.Enabled = True
txtOpeCod.SetFocus
End Sub

Private Sub cmdEliminar_Click()
On Error GoTo EliminaErr
If MsgBox(" ¿ Está seguro de eliminar Operación ? ", vbQuestion + vbYesNo, "Mensaje de confirmación") = vbYes Then
   clsOpe.EliminaOpeGru rsGru!cOpeGruCod
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantClasifOperacion
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & "| CopeCod : " & rsGru!cOpeGruCod
            Set objPista = Nothing
            '*******
   rsGru.Delete adAffectCurrent
End If
grdOpe.SetFocus
Exit Sub
EliminaErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdCancelar_Click()
ActivaBotones True
grdOpe.SetFocus
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If Len(txtOpeCod.Text) = 0 Then
   MsgBox " Operación no válida ...  ", vbInformation, "¡Aviso!"
   txtOpeCod.SetFocus
   Exit Function
End If

If Len(txtOpeDesc.Text) = 0 Then
   MsgBox "  Descripción no puede ser vacio ...  ", vbInformation, "Error"
   txtOpeDesc.SetFocus
   Exit Function
End If

Dim prs As ADODB.Recordset
Set prs = clsOpe.CargaOpeGru(, adLockOptimistic)
If Not prs.EOF Then
   prs.Find "cOpeGruDesc = '" & Me.txtOpeDesc & "'"
   If Not prs.EOF Then
        If lNuevo Then
           MsgBox "Existe Operación creada con la misma descripción...      ", vbInformation, "¡Aviso!"
           RSClose prs
           Exit Function
        Else
           If prs!cOpeGruCod <> Me.txtOpeCod Then
              MsgBox "Existe Operación creada con la misma descripción...      ", vbInformation, "¡Aviso!"
              RSClose prs
              Exit Function
           End If
        End If
   End If
End If

ValidaDatos = True




End Function

Private Sub cmdAceptar_Click()
On Error GoTo ErrAceptar
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro de grabar datos ? ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   Select Case lNuevo
      Case True
            clsOpe.InsertaOpeGru txtOpeCod.Text, txtOpeDesc, gsMovNro
                  
      Case False
            clsOpe.ActualizaOpeGru txtOpeCod.Text, txtOpeDesc, gsMovNro
   End Select
   CargaGrupos
   rsGru.Find "cOpeGruCod = '" & txtOpeCod.Text & "'", , , 1
End If
            'ARLO20170208
            Dim lsAccion As String
            Set objPista = New COMManejador.Pista
            If lNuevo Then
            lsAccion = "1"
            Else:
            lsAccion = "2"
            End If
            gsOpeCod = LogPistaMantClasifOperacion
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, Me.Caption & " | CodOpe : " & txtOpeCod.Text & " | Descripcion : " & txtOpeDesc.Text
            Set objPista = Nothing
            '*******
ActivaBotones True
grdOpe.SetFocus
Exit Sub
ErrAceptar:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub Form_Unload(Cancel As Integer)
RSClose rsGru
Set clsOpe = Nothing
End Sub

Private Sub grdOpe_GotFocus()
grdOpe.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdOpe_HeadClick(ByVal ColIndex As Integer)
If Not rsGru Is Nothing Then
   If Not rsGru.EOF Then
      rsGru.Sort = grdOpe.Columns(ColIndex).DataField
   End If
End If
End Sub

Private Sub grdOpe_LostFocus()
grdOpe.MarqueeStyle = dbgNoMarquee
End Sub


Private Sub txtOpeCod_GotFocus()
txtOpeCod.BackColor = "&H00F5FFDD"
End Sub
Private Sub txtOpeCod_LostFocus()
txtOpeCod.BackColor = "&H80000005"
End Sub
Private Sub txtOpeCod_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   txtOpeDesc.SetFocus
End If
End Sub

Private Sub txtOpeDesc_GotFocus()
txtOpeDesc.BackColor = "&H00F5FFDD"
End Sub

Private Sub txtOpeDesc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub

Private Sub txtOpeDesc_LostFocus()
txtOpeDesc.BackColor = "&H80000005"
End Sub

Sub ActivaBotones(lActiva As Boolean)
If lActiva Then
   grdOpe.Height = 3210
Else
   grdOpe.Height = 2670
End If
cmdNuevo.Visible = lActiva
cmdModificar.Visible = lActiva
cmdEliminar.Visible = lActiva
cmdImprimir.Visible = lActiva
cmdSalir.Visible = lActiva
cmdAceptar.Visible = Not lActiva
cmdCancelar.Visible = Not lActiva
txtOpeCod.Visible = Not lActiva
txtOpeDesc.Visible = Not lActiva
End Sub

