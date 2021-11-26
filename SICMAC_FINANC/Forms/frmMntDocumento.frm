VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntDocumento 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Documentos: Mantenimiento"
   ClientHeight    =   5460
   ClientLeft      =   3015
   ClientTop       =   1995
   ClientWidth     =   7020
   Icon            =   "frmMntDocumento.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   7020
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5340
      Left            =   60
      TabIndex        =   11
      Top             =   0
      Width           =   6855
      Begin MSDataGridLib.DataGrid grdDoc 
         Height          =   4275
         Left            =   240
         TabIndex        =   0
         Top             =   300
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   7541
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   2
         RowHeight       =   17
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
         ColumnCount     =   3
         BeginProperty Column00 
            DataField       =   "nDocTpo"
            Caption         =   "Codigo"
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
            DataField       =   "cDocDesc"
            Caption         =   "Documento"
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
            DataField       =   "cDocAbrev"
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
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   2
            BeginProperty Column00 
               Alignment       =   2
               DividerStyle    =   6
               ColumnWidth     =   675.213
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   4034.835
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnWidth     =   1094.74
            EndProperty
         EndProperty
      End
      Begin VB.TextBox tDocAbrev 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   5280
         MaxLength       =   3
         TabIndex        =   8
         Top             =   4050
         Width           =   1095
      End
      Begin VB.TextBox tDocDesc 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   1260
         MaxLength       =   40
         TabIndex        =   7
         Top             =   4050
         Width           =   4035
      End
      Begin VB.TextBox tDocTpo 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   340
         Left            =   540
         MaxLength       =   3
         TabIndex        =   6
         Top             =   4050
         Width           =   735
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   400
         Left            =   300
         TabIndex        =   1
         Top             =   4725
         Width           =   1100
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   400
         Left            =   1440
         TabIndex        =   2
         Top             =   4725
         Width           =   1100
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   400
         Left            =   2580
         TabIndex        =   3
         Top             =   4725
         Width           =   1100
      End
      Begin VB.CommandButton cmdImpuestos 
         Caption         =   "&Impuestos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   400
         Left            =   3720
         TabIndex        =   4
         Top             =   4725
         Width           =   1100
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   400
         Left            =   4320
         TabIndex        =   9
         Top             =   4725
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         CausesValidation=   0   'False
         Height          =   400
         Left            =   5460
         TabIndex        =   5
         Top             =   4725
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         CausesValidation=   0   'False
         Height          =   400
         Left            =   5460
         TabIndex        =   10
         Top             =   4725
         Visible         =   0   'False
         Width           =   1100
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   240
         Top             =   3975
         Width           =   6375
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   240
         TabIndex        =   12
         Top             =   4650
         Width           =   6375
      End
   End
End
Attribute VB_Name = "frmMntDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OpGraba As String

Dim lConsulta As Boolean
Dim clsDoc    As DDocumento
Dim rsDoc     As ADODB.Recordset
Dim psSql     As String
'ARLO2010208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(plConsulta As Boolean)
lConsulta = plConsulta
Me.Show 1
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub CargaDocumento()
Set rsDoc = clsDoc.CargaDocumento(, , adLockOptimistic)
Set grdDoc.DataSource = rsDoc
End Sub
Private Sub Form_Load()
Set clsDoc = New DDocumento
ActivaControlMnt False
ActivaBotones True
CargaDocumento
If lConsulta Then
   cmdNuevo.Visible = False
   cmdModificar.Visible = False
   cmdEliminar.Visible = False
   cmdImpuestos.Left = cmdNuevo.Left
End If
CentraForm Me
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
tDocTpo = Format(tDocTpo, "00")
tDocDesc = UCase(tDocDesc)
tDocAbrev = UCase(tDocAbrev)

If Len(tDocTpo) = 0 Or Len(tDocDesc) = 0 Or Len(tDocAbrev) = 0 Then
   MsgBox "No se pueden grabar campos en blanco...      ", vbInformation, "¡Aviso!"
   Exit Function
End If

Dim prs As ADODB.Recordset
Set prs = clsDoc.CargaDocumento(, tDocAbrev)
If Not prs.EOF Then
   If OpGraba = "1" Then
      MsgBox "La abreviatura ya está relacionada con otro documento...      ", vbInformation, "¡Aviso!"
      RSClose prs
      Exit Function
   Else
      If prs!nDocTpo <> tDocTpo Then
         MsgBox "La abreviatura ya está relacionada con otro documento...      ", vbInformation, "¡Aviso!"
         RSClose prs
         Exit Function
      End If
   End If
End If
RSClose prs
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
Dim rs As New ADODB.Recordset
On Error GoTo AceptarErr
If Not ValidaDatos Then
   Exit Sub
End If
If MsgBox(" ¿ Está seguro de grabar los datos ?      ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
   Select Case OpGraba
     Case "1"
          clsDoc.InsertaDocumento tDocTpo, tDocDesc, tDocAbrev, gsMovNro
     Case "2"
          clsDoc.ActualizaDocumento tDocTpo, tDocDesc, tDocAbrev, gsMovNro
   End Select
   CargaDocumento
   rsDoc.Find "nDocTpo = '" & tDocTpo & "'"
End If
ActivaControlMnt False
ActivaBotones True
            'ARLO20170208
            Dim lsAccion As String
            Set objPista = New COMManejador.Pista
            If (OpGraba = 1) Then
            lsAccion = "1"
            Else:
            lsAccion = "2"
            End If
            gsOpeCod = LogPistaMantDocumento
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, Me.Caption & " | Codigo : " & tDocTpo & " | Descripcion : " & tDocDesc & " | Abreviatura : " & tDocAbrev
            Set objPista = Nothing
            '*******
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdCancelar_Click()
ActivaControlMnt False
ActivaBotones True
End Sub

Private Sub cmdNuevo_Click()
tDocTpo = ""
tDocDesc = ""
tDocAbrev = ""
OpGraba = "1"
ActivaControlMnt True
ActivaBotones False
tDocTpo.SetFocus

End Sub

Private Sub cmdEliminar_Click()
On Error GoTo EliminarErr
If Not rsDoc.EOF Then
   If MsgBox(" ¿ Está seguro de eliminar el documento ?      ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
      clsDoc.EliminaDocumento rsDoc!nDocTpo
      rsDoc.Delete adAffectCurrent
      grdDoc.SetFocus
   End If
Else
   MsgBox "No hay documentos registrados...      ", vbInformation, "¡Aviso!"
End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantDocumento
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " | Codigo : " & tDocTpo & " | Descripcion : " & tDocDesc & " | Abreviatura : " & tDocAbrev
            Set objPista = Nothing
            '*******
Exit Sub
EliminarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub

Private Sub cmdImpuestos_Click()
Dim nPos As Variant
If Not rsDoc.EOF Then
   nPos = rsDoc.Bookmark
   gnDocTpo = rsDoc("nDocTpo")
   gsDocDesc = rsDoc("cDocDesc")
   frmMntDocumentoImp.Inicio lConsulta
Else
   MsgBox "No hay documentos registrados...      ", vbCritical, "Error"
End If
End Sub

Private Sub cmdModificar_Click()
OpGraba = "2"
If Not rsDoc.EOF Then
   tDocTpo = rsDoc("nDocTpo")
   tDocDesc = rsDoc("cDocDesc")
   tDocAbrev = rsDoc("cDocAbrev")
   ActivaControlMnt True
   ActivaBotones False
   tDocTpo.Enabled = False
   tDocDesc.SetFocus
Else
   MsgBox "No hay documentos registrados...      ", vbCritical, "Error"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set clsDoc = Nothing
RSClose rsDoc
End Sub

Private Sub grdDoc_HeadClick(ByVal ColIndex As Integer)
If Not rsDoc Is Nothing Then
   If Not rsDoc.EOF Then
      rsDoc.Sort = grdDoc.Columns(ColIndex).DataField
   End If
End If
End Sub

Private Sub tDocTpo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   tDocTpo = Format(tDocTpo, "00")
   tDocDesc.SetFocus
End If
End Sub

Private Sub tDocDesc_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii, True)
If KeyAscii = 13 Then
   tDocDesc = UCase(tDocDesc)
   tDocAbrev.SetFocus
End If
End Sub

Private Sub tDocAbrev_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub

Private Sub grdDoc_GotFocus()
grdDoc.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdDoc_LostFocus()
grdDoc.MarqueeStyle = dbgNoMarquee
End Sub

Sub ActivaBotones(plActiva As Boolean)
If plActiva Then
   grdDoc.Height = 4215
Else
   grdDoc.Height = 3690
End If
grdDoc.Enabled = plActiva
cmdNuevo.Visible = plActiva
cmdEliminar.Visible = plActiva
cmdModificar.Visible = plActiva
cmdImpuestos.Visible = plActiva
cmdSalir.Visible = plActiva
cmdAceptar.Visible = Not plActiva
cmdCancelar.Visible = Not plActiva
End Sub

Sub ActivaControlMnt(plActiva As Boolean)
tDocTpo.Enabled = plActiva
tDocDesc.Enabled = plActiva
tDocAbrev.Enabled = plActiva
End Sub

