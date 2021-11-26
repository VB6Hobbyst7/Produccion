VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntDocumentoImp 
   Caption         =   "Documentos: Relación de Impuestos"
   ClientHeight    =   4185
   ClientLeft      =   2460
   ClientTop       =   2715
   ClientWidth     =   7365
   Icon            =   "frmMntDocumentoImp.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Impuestos "
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
      Height          =   3165
      Left            =   120
      TabIndex        =   10
      Top             =   900
      Width           =   7095
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   400
         Left            =   4380
         TabIndex        =   8
         Top             =   2400
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.Frame fraCar 
         Caption         =   "Carácter "
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
         Height          =   1290
         Left            =   2040
         TabIndex        =   12
         Top             =   1575
         Visible         =   0   'False
         Width           =   1755
         Begin VB.OptionButton OpOb 
            Caption         =   "Obligatorio"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   6
            Top             =   450
            Width           =   1275
         End
         Begin VB.OptionButton OpOp 
            Caption         =   "Opcional"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   825
            Width           =   1095
         End
      End
      Begin VB.Frame fraUbi 
         Caption         =   "Ubicación "
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
         Height          =   1290
         Left            =   180
         TabIndex        =   11
         Top             =   1575
         Visible         =   0   'False
         Width           =   1635
         Begin VB.OptionButton OpH 
            Caption         =   "Haber"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   5
            Top             =   825
            Width           =   855
         End
         Begin VB.OptionButton OpD 
            Caption         =   "Debe"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   300
            TabIndex        =   4
            Top             =   450
            Width           =   915
         End
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   400
         Left            =   5700
         TabIndex        =   9
         Top             =   2400
         Visible         =   0   'False
         Width           =   1200
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   400
         Left            =   180
         TabIndex        =   1
         Top             =   2625
         Width           =   1200
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Quitar"
         Height          =   400
         Left            =   1440
         TabIndex        =   2
         Top             =   2625
         Width           =   1200
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   400
         Left            =   5700
         TabIndex        =   3
         Top             =   2625
         Width           =   1200
      End
      Begin VB.Frame fraCta 
         Caption         =   "Cuenta Contable "
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
         Height          =   840
         Left            =   180
         TabIndex        =   13
         Top             =   375
         Visible         =   0   'False
         Width           =   6735
         Begin Sicmact.TxtBuscar tCtaCod 
            Height          =   345
            Left            =   150
            TabIndex        =   17
            Top             =   375
            Width           =   2415
            _ExtentX        =   4260
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
         End
         Begin VB.TextBox tCtaDesc 
            Enabled         =   0   'False
            Height          =   315
            Left            =   2580
            MaxLength       =   30
            TabIndex        =   14
            Top             =   375
            Width           =   3975
         End
      End
      Begin MSDataGridLib.DataGrid grdImp 
         Height          =   2190
         Left            =   180
         TabIndex        =   0
         Top             =   300
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   3863
         _Version        =   393216
         AllowUpdate     =   0   'False
         BorderStyle     =   0
         HeadLines       =   2
         RowHeight       =   19
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
         ColumnCount     =   4
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
            DataField       =   "Clase"
            Caption         =   "Clase"
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
            DataField       =   "Caracter"
            Caption         =   "Caracter"
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
               ColumnAllowSizing=   0   'False
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   2910.047
            EndProperty
            BeginProperty Column02 
               Alignment       =   2
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   854.929
            EndProperty
            BeginProperty Column03 
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1110.047
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   690
      Left            =   120
      TabIndex        =   15
      Top             =   75
      Width           =   7095
      Begin VB.Label lblTitulo 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00C0C000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "XXXX"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         TabIndex        =   16
         Top             =   210
         Width           =   6735
      End
   End
End
Attribute VB_Name = "frmMntDocumentoImp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sClase As String, sOpcion As String
Dim rsCta As New ADODB.Recordset
Dim clsImp As DImpuesto
Dim clsDoc As DDocumento
Dim lConsulta As Boolean
'ARLO20170208****
Dim objPista As COMManejador.Pista
'************

Public Sub Inicio(plConsulta As Boolean)
lConsulta = plConsulta
Me.Show 1
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
lblTitulo.Caption = gsDocDesc
If lConsulta Then
   cmdAgregar.Visible = False
   cmdQuitar.Visible = False
End If

Set clsDoc = New DDocumento
Set clsImp = New DImpuesto
MuestraImpuestos
tCtaCod.rs = clsImp.CargaCtaImpuesto()
tCtaCod.EditFlex = False
tCtaCod.psRaiz = "Impuestos"
CentraForm Me
End Sub
Sub MuestraImpuestos()
Set rsCta = clsDoc.CargaDocImpuesto(gnDocTpo)
Set grdImp.DataSource = rsCta
End Sub

Private Sub cmdCancelar_Click()
ActivaControles True
grdImp.SetFocus
End Sub

Private Sub cmdAgregar_Click()
tCtaCod = ""
tCtaDesc = ""
OpD.value = True
OpOb.value = True
ActivaControles False
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo AceptarErr
sClase = "H"
If OpD.value Then
   sClase = "D"
End If

sOpcion = "2"
If OpOb.value Then
   sOpcion = "1"
End If

gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
If MsgBox(" ¿ Está seguro de agregar el Impuesto ?      ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   clsDoc.AsignaDocImpuesto gnDocTpo, tCtaCod, sClase, sOpcion, gsMovNro
   MuestraImpuestos
End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantDocumento
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "1", Me.Caption & " | Codigo : " & gnDocTpo & " | Cuenta Contable : " & tCtaCod
            Set objPista = Nothing
            '*******
ActivaControles True
Exit Sub
AceptarErr:
   MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
End Sub
Private Sub cmdQuitar_Click()
If rsCta.EOF And rsCta.BOF Then
   MsgBox "No hay impuestos disponibles...       ", vbInformation, "¡Aviso!"
   Exit Sub
Else
   tCtaCod = rsCta!cCtaContCod
End If
If MsgBox(" ¿ Está seguro de quitar el impuesto ?      ", vbQuestion + vbYesNo, "¡Confirmación!") = vbYes Then
   clsDoc.DesasignaDocImpuesto gnDocTpo, tCtaCod
   MuestraImpuestos
End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            gsOpeCod = LogPistaMantDocumento
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", Me.Caption & " | Codigo : " & gnDocTpo & " | Cuenta Contable : " & tCtaCod
            Set objPista = Nothing
            '*******
End Sub

Private Sub grdImp_GotFocus()
grdImp.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdImp_LostFocus()
grdImp.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub OpD_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   OpOb.SetFocus
End If
End Sub
Private Sub OpH_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   OpOb.SetFocus
End If
End Sub

Private Sub OpOb_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub

Private Sub OpOp_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub

Private Sub tCtaCod_EmiteDatos()
tCtaDesc = tCtaCod.psDescripcion
   If tCtaCod.Text <> "" And OpD.Visible Then
      OpD.SetFocus
   End If
End Sub

Sub ActivaControles(plActiva As Boolean)
cmdAgregar.Visible = plActiva
cmdQuitar.Visible = plActiva
grdImp.Visible = plActiva
fraCta.Visible = Not plActiva
fraUbi.Visible = Not plActiva
fraCar.Visible = Not plActiva
cmdAceptar.Visible = Not plActiva
cmdCancelar.Visible = Not plActiva
cmdSalir.Visible = plActiva
End Sub


