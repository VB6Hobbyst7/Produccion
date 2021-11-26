VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form frmMntBienesNavidad 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Regalo por Fiestas para Personal"
   ClientHeight    =   5475
   ClientLeft      =   2805
   ClientTop       =   2025
   ClientWidth     =   10425
   Icon            =   "frmMntBienesNavidad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   10425
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   5340
      Left            =   60
      TabIndex        =   18
      Top             =   0
      Width           =   10275
      Begin VB.CommandButton cmdListado 
         Caption         =   "&Listado"
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
         Left            =   5100
         TabIndex        =   14
         Top             =   4725
         Width           =   1425
      End
      Begin VB.CommandButton cmdComprobante 
         Caption         =   "&Comprobantes"
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
         TabIndex        =   13
         Top             =   4725
         Width           =   1365
      End
      Begin MSDataGridLib.DataGrid grd 
         Height          =   3750
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   9915
         _ExtentX        =   17489
         _ExtentY        =   6615
         _Version        =   393216
         AllowUpdate     =   0   'False
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
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "cPersNombre"
            Caption         =   "Persona"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   "nCanasta"
            Caption         =   "Canasta"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column02 
            DataField       =   "nPavo"
            Caption         =   "Pavo"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column03 
            DataField       =   "njuguete"
            Caption         =   "Juguete"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column04 
            DataField       =   "nigv"
            Caption         =   "IGV"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column05 
            DataField       =   "nTotal"
            Caption         =   "Total"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "#,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            MarqueeStyle    =   3
            ScrollBars      =   2
            BeginProperty Column00 
               DividerStyle    =   6
               ColumnAllowSizing=   -1  'True
               ColumnWidth     =   3390.236
            EndProperty
            BeginProperty Column01 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column02 
               Alignment       =   1
               ColumnAllowSizing=   0   'False
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column03 
               Alignment       =   1
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column04 
               Alignment       =   1
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column05 
               Alignment       =   1
               ColumnWidth     =   1200.189
            EndProperty
         EndProperty
      End
      Begin VB.TextBox txtJuguete 
         Alignment       =   1  'Right Justify
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
         Left            =   6390
         MaxLength       =   40
         TabIndex        =   7
         Top             =   4140
         Width           =   1200
      End
      Begin VB.TextBox txtPersNombre 
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
         Left            =   360
         MaxLength       =   60
         TabIndex        =   4
         Top             =   4140
         Width           =   3615
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   400
         Left            =   300
         TabIndex        =   10
         Top             =   4725
         Width           =   1125
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   400
         Left            =   1440
         TabIndex        =   11
         Top             =   4725
         Width           =   1125
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Height          =   400
         Left            =   2580
         TabIndex        =   12
         Top             =   4725
         Width           =   1125
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   400
         Left            =   7800
         TabIndex        =   16
         Top             =   4725
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         CausesValidation=   0   'False
         Height          =   400
         Left            =   8940
         TabIndex        =   15
         Top             =   4725
         Width           =   1125
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         CausesValidation=   0   'False
         Height          =   400
         Left            =   8940
         TabIndex        =   17
         Top             =   4725
         Visible         =   0   'False
         Width           =   1125
      End
      Begin VB.TextBox txtCanasta 
         Alignment       =   1  'Right Justify
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
         Left            =   3990
         MaxLength       =   15
         TabIndex        =   5
         Top             =   4140
         Width           =   1200
      End
      Begin VB.TextBox txtPavo 
         Alignment       =   1  'Right Justify
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
         Left            =   5190
         MaxLength       =   15
         TabIndex        =   6
         Top             =   4140
         Width           =   1200
      End
      Begin VB.Frame Frame2 
         Height          =   645
         Left            =   240
         TabIndex        =   20
         Top             =   120
         Width           =   6915
         Begin VB.ComboBox cboTipo 
            Height          =   315
            ItemData        =   "frmMntBienesNavidad.frx":030A
            Left            =   600
            List            =   "frmMntBienesNavidad.frx":0314
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   210
            Width           =   2745
         End
         Begin VB.TextBox txtAnio 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   4050
            MaxLength       =   4
            TabIndex        =   1
            Top             =   210
            Width           =   1035
         End
         Begin VB.CommandButton cmdProcesar 
            Caption         =   "&Procesar"
            Height          =   345
            Left            =   5370
            TabIndex        =   2
            Top             =   210
            Width           =   1335
         End
         Begin VB.Label Label3 
            Caption         =   "Tipo"
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
            TabIndex        =   23
            Top             =   270
            Width           =   495
         End
         Begin VB.Label Label2 
            Caption         =   "Año"
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
            Left            =   3570
            TabIndex        =   21
            Top             =   270
            Width           =   495
         End
      End
      Begin VB.TextBox txtIgv 
         Alignment       =   1  'Right Justify
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
         Left            =   7590
         MaxLength       =   40
         TabIndex        =   8
         Top             =   4140
         Width           =   1200
      End
      Begin VB.TextBox txtTotal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
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
         Left            =   8790
         MaxLength       =   40
         TabIndex        =   9
         Top             =   4140
         Width           =   1200
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00E0E0E0&
         Height          =   495
         Left            =   240
         Top             =   4065
         Width           =   9915
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   540
         Left            =   240
         TabIndex        =   19
         Top             =   4650
         Width           =   9885
      End
   End
   Begin RichTextLib.RichTextBox rtxtAsiento 
      Height          =   315
      Left            =   4590
      TabIndex        =   22
      Top             =   300
      Visible         =   0   'False
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      _Version        =   393217
      Enabled         =   0   'False
      ScrollBars      =   2
      TextRTF         =   $"frmMntBienesNavidad.frx":037A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmMntBienesNavidad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OpGraba As String

Dim lConsulta As Boolean
Dim clsBien   As DBienesNavidad
Dim rs        As ADODB.Recordset

Public Sub Inicio(plConsulta As Boolean)
lConsulta = plConsulta
Me.Show 1
End Sub

Private Sub cboTipo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtAnio.SetFocus
End If
End Sub

Private Sub cmdComprobante_Click()
Dim rsImp As ADODB.Recordset
Dim sImpre As String
Dim sImpr1 As String
gdFecha = CDate("31/12/" & txtAnio)
Set rsImp = clsBien.CargaBeneficiarios(Right(cboTipo, 1), gdFecha)
If rsImp.EOF Then
   MsgBox "No existen datos para Emitir Comprobantes...!", vbInformation, "¡Aviso!"
   RSClose rsImp
   Exit Sub
End If
sImpr1 = ""
sImpre = PrnSet("MI", 3)
Do While Not rsImp.EOF
   gnImporte = rsImp!nCanasta + rsImp!nPavo + rsImp!nJuguete + rsImp!nIGV
   Linea sImpre, BON & Justifica(gsNomCmac, 50) & Space(7) & "RUC : " & gsRUC
   Linea sImpre, Justifica("AG. SEDE INSTITUCIONAL", 50) & Space(7) & "TRANSFERENCIA GRATUITA"
   Linea sImpre, "AV. ESPAÑA  No.2611"
   Linea sImpre, String(76, "-")
   Linea sImpre, Space(57) & Format(gdFecha & " " & GetHoraServer, "dd/mm/yyyy hh:mm:ss"), 3
   Linea sImpre, Centra("COMPROBANTES DE RETIRO DE BIENES", 72) & BOFF
   Linea sImpre, String(76, "-"), 2
   Linea sImpre, "BENEFICIARIO    : " & rsImp!cPersNombre
   Linea sImpre, "                  ENTREGA POR POLITICA INSTITUCIONAL (AGUINALDO)"
   Linea sImpre, "                  ", 3
   Linea sImpre, "     CONCEPTOS                                    IMPORTE    "
   Linea sImpre, "    ------------------------------------------------------------", 2
   Linea sImpre, "     CANASTA NAVIDEÑA (SEGUN RELACION)          " & PrnVal(rsImp!nCanasta, 14, 2), 2
   Linea sImpre, "     PAVO                                       " & PrnVal(rsImp!nPavo, 14, 2), 2
   Linea sImpre, "     JUGUETES                                   " & PrnVal(rsImp!nJuguete, 14, 2), 2
   Linea sImpre, "    ------------------------------------------------------------"
   Linea sImpre, "                               VALOR VENTA      " & PrnVal(rsImp!nCanasta + rsImp!nPavo + rsImp!nJuguete, 14, 2)
   Linea sImpre, "                               I.G.V.           " & PrnVal(rsImp!nIGV, 14, 2) & BON
   Linea sImpre, "                               TOTAL         " & Justifica(gsSimbolo, 3) & PrnVal(gnImporte, 14, 2) & BOFF
   Linea sImpre, "                            =====================================", 3
   Linea sImpre, "SON: " & ConvNumLet(gnImporte), 2
   Linea sImpre, "COPIA SIN DERECHO A CREDITO FISCAL DEL I.G.V.", 2
   Linea sImpre, "RECIBI CONFORME" & BON, 4
   Linea sImpre, Centra(String(Len(rsImp!cPersNombre) + 4, "_"), 72)
   Linea sImpre, Centra("  " & rsImp!cPersNombre & "  ", 72) & BOFF
   Linea sImpre, Chr(12), 0
   If Len(sImpre) > 4000 Then
      sImpr1 = sImpr1 & sImpre
      sImpre = ""
   End If
   rsImp.MoveNext
Loop
sImpre = sImpr1 & sImpre
RSClose rsImp
EnviaPrevio ImpreCarEsp(sImpre), "CANASTAS PARA PERSONAL: COMPROBANTES", gnLinPage, False
End Sub

Private Sub cmdListado_Click()
Dim rsImp As ADODB.Recordset
Dim sImpre As String
Dim nPag   As Integer, nLin As Integer
Dim nTotCanasta As Currency
Dim nTotPavo    As Currency
Dim nTotJuguete As Currency
Dim nTotIgv     As Currency
Dim nTotTotal   As Currency

gdFecha = CDate("31/12/" & txtAnio)
Set rsImp = clsBien.CargaBeneficiarios(Right(cboTipo, 1), gdFecha)
If rsImp.EOF Then
   MsgBox "No existen datos para Emitir Comprobantes...!", vbInformation, "¡Aviso!"
   RSClose rsImp
   Exit Sub
End If
sImpre = PrnSet("MI", 3)
nPag = 0
nLin = gnLinPage
Do While Not rsImp.EOF
   If nLin > gnLinPage - 6 Then
      If nPag > 0 Then sImpre = sImpre & Chr(12)
      nPag = nPag + 1
      Linea sImpre, BON & Justifica(gsNomCmac, 50) & Space(8) & "RUC : " & gsRUC
      Linea sImpre, Justifica("AG. SEDE INSTITUCIONAL", 50) & "  Pag. " & Format(nPag, "000") & "    " & Format(gdFecSis, "dd/mm/yyyy")
      Linea sImpre, Justifica("AV. ESPAÑA  No.2611", 50)
      Linea sImpre, Centra("LISTADO DE COMPROBANTES DE RETIRO DE BIENES", 72) & BOFF & CON
      Linea sImpre, String(125, "-")
      Linea sImpre, Justifica(" PERSONA", 50) & " " & Justifica("  CANASTA", 14) & " " & Justifica("    PAVO", 14) & " " & Justifica("   JUGUETE", 14) & " " & Justifica("    I.G.V.", 14) & " " & Justifica("   TOTAL", 14)
      Linea sImpre, String(125, "-") & COFF
      nLin = 7
   End If
   nTotCanasta = nTotCanasta + rsImp!nCanasta
   nTotPavo = nTotPavo + rsImp!nPavo
   nTotJuguete = nTotJuguete + rsImp!nJuguete
   nTotIgv = nTotIgv + rsImp!nIGV
   nTotTotal = nTotTotal + rsImp!nTotal
   
   gnImporte = rsImp!nCanasta + rsImp!nPavo + rsImp!nJuguete + rsImp!nIGV
   Linea sImpre, CON & Justifica(rsImp!cPersNombre, 45) & " " & PrnVal(rsImp!nCanasta, 14, 2) & " " & PrnVal(rsImp!nPavo, 14, 2) & " " & PrnVal(rsImp!nJuguete, 14, 2) & " " & PrnVal(rsImp!nIGV, 14, 2) & " " & PrnVal(gnImporte, 14, 2) & COFF
   nLin = nLin + 1
   rsImp.MoveNext
Loop
RSClose rsImp
Linea sImpre, String(125, "-")
Linea sImpre, BON & CON & Justifica(" ", 45) & " " & PrnVal(nTotCanasta, 14, 2) & " " & PrnVal(nTotPavo, 14, 2) & " " & PrnVal(nTotJuguete, 14, 2) & " " & PrnVal(nTotIgv, 14, 2) & " " & PrnVal(nTotTotal, 14, 2) & COFF & BOFF

sImpre = ImpreCarEsp(sImpre)
EnviaPrevio sImpre, "CANASTA PARA PERSONAL: LISTADO", gnLinPage
End Sub

Private Sub cmdProcesar_Click()
gdFecha = CDate("31/12/" & txtAnio)
CargaDatos
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub CargaDatos()
Set rs = clsBien.CargaBeneficiarios(Right(cboTipo, 1), gdFecha, , adLockOptimistic)
Set grd.DataSource = rs
End Sub

Private Sub Form_Load()
CentraForm Me
Set clsBien = New DBienesNavidad
AbreConexion
ActivaControlMnt False
ActivaBotones True
gsSimbolo = gcMN
If lConsulta Then
   cmdNuevo.Visible = False
   cmdModificar.Visible = False
   cmdEliminar.Visible = False
End If
txtAnio = Year(gdFecSis)
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
txtPersNombre = Trim(txtPersNombre)
If Len(txtPersNombre) = 0 Then
   MsgBox "Falta indicar Empleado beneficiado...      ", vbInformation, "¡Aviso!"
   Exit Function
End If
If nVal(txtTotal) = 0 Then
   MsgBox "No se asignó ningún concepto a Empleado... !", vbInformation, "¡Aviso!"
End If
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
On Error GoTo AceptarErr
If Not ValidaDatos Then
   Exit Sub
End If
gdFecha = CDate("31/12/" & txtAnio)

If MsgBox(" ¿ Está seguro de grabar los datos ?      ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
   Select Case OpGraba
     Case "1"
          clsBien.InsertaBeneficiario Right(cboTipo, 1), gdFecha, txtPersNombre, nVal(txtCanasta), nVal(txtPavo), nVal(txtJuguete), nVal(txtIgv)
     Case "2"
          clsBien.ActualizaBeneficiario Right(cboTipo, 1), gdFecha, txtPersNombre, nVal(txtCanasta), nVal(txtPavo), nVal(txtJuguete), nVal(txtIgv)
   End Select
   CargaDatos
   rs.Find "cPersNombre = '" & txtPersNombre & "'"
End If
ActivaControlMnt False
ActivaBotones True
grd.SetFocus
Exit Sub
AceptarErr:
   MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub cmdCancelar_Click()
ActivaControlMnt False
ActivaBotones True
grd.SetFocus
End Sub

Private Sub cmdNuevo_Click()
OpGraba = "1"
ActivaControlMnt True
ActivaBotones False
txtPersNombre.SetFocus
End Sub

Private Sub cmdEliminar_Click()
If Not rs.EOF Then
   If MsgBox(" ¿ Está seguro de eliminar el documento ?      ", vbQuestion + vbYesNo, "Confirmación") = vbYes Then
      clsBien.EliminaBeneficiario Right(cboTipo, 1), gdFecha, rs!cPersNombre
      rs.Delete adAffectCurrent
      grd.SetFocus
   End If
Else
   MsgBox "No hay datos registrados...! ", vbInformation, "¡Aviso!"
End If
End Sub

Private Sub cmdModificar_Click()
OpGraba = "2"
If Not rs.EOF Then
   txtPersNombre = Format(rs!cPersNombre, gsFormatoNumeroView)
   txtCanasta = Format(rs!nCanasta, gsFormatoNumeroView)
   txtPavo = Format(rs!nPavo, gsFormatoNumeroView)
   txtJuguete = Format(rs!nJuguete, gsFormatoNumeroView)
   txtIgv = Format(rs!nIGV, gsFormatoNumeroView)
   txtTotal = Format(rs!nTotal, gsFormatoNumeroView)
   ActivaControlMnt True
   ActivaBotones False
   txtPersNombre.SetFocus
Else
   MsgBox "No hay datos registrados...!", vbInformation, "¡Aviso!"
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set clsBien = Nothing
CierraConexion
RSClose rs
End Sub

Private Sub grd_HeadClick(ByVal ColIndex As Integer)
If Not rs Is Nothing Then
   If Not rs.EOF Then
      rs.Sort = grd.Columns(ColIndex).DataField
   End If
End If
End Sub

Private Sub txtIgv_GotFocus()
fEnfoque txtIgv
End Sub

Private Sub txtIgv_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtIgv, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtIgv = Format(txtIgv, gsFormatoNumeroView)
   cmdAceptar.SetFocus
End If
End Sub

Private Sub txtPersNombre_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   txtPersNombre = UCase(txtPersNombre)
   txtCanasta.SetFocus
End If
End Sub

Private Sub txtPavo_GotFocus()
fEnfoque txtPavo
End Sub

Private Sub txtPavo_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPavo, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtPavo = Format(txtPavo, gsFormatoNumeroView)
   txtIgv = Format(Round((nVal(txtCanasta) + nVal(txtPavo) + nVal(txtJuguete)) * gnIGVValor, 2), gsFormatoNumeroView)
   txtTotal = nVal(txtCanasta) + nVal(txtPavo) + nVal(txtJuguete) + nVal(txtIgv)
   txtJuguete.SetFocus
End If
End Sub

Private Sub txtCanasta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtCanasta, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtCanasta = Format(txtCanasta, gsFormatoNumeroView)
   txtIgv = Format(Round((nVal(txtCanasta) + nVal(txtPavo) + nVal(txtJuguete)) * gnIGVValor, 2), gsFormatoNumeroView)
   txtTotal = nVal(txtCanasta) + nVal(txtPavo) + nVal(txtJuguete) + nVal(txtIgv)
   txtPavo.SetFocus
End If
End Sub

Private Sub txtJuguete_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtJuguete, KeyAscii, 14, 2)
If KeyAscii = 13 Then
   txtJuguete = Format(txtJuguete, gsFormatoNumeroView)
   txtIgv = Format(Round((nVal(txtCanasta) + nVal(txtPavo) + nVal(txtJuguete)) * gnIGVValor, 2), gsFormatoNumeroView)
   txtTotal = nVal(txtCanasta) + nVal(txtPavo) + nVal(txtJuguete) + nVal(txtIgv)
   txtIgv.SetFocus
End If
End Sub

Private Sub grd_GotFocus()
grd.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grd_LostFocus()
grd.MarqueeStyle = dbgNoMarquee
End Sub

Sub ActivaBotones(plActiva As Boolean)
If plActiva Then
   grd.Height = 3750
Else
   grd.Height = 3150
End If
grd.Enabled = plActiva
cmdNuevo.Visible = plActiva
cmdEliminar.Visible = plActiva
cmdModificar.Visible = plActiva
cmdSalir.Visible = plActiva
cmdAceptar.Visible = Not plActiva
cmdCancelar.Visible = Not plActiva
End Sub

Sub ActivaControlMnt(plActiva As Boolean)
   txtPersNombre.Enabled = plActiva
   txtCanasta.Enabled = plActiva
   txtPavo.Enabled = plActiva
   txtJuguete.Enabled = plActiva
   txtIgv.Enabled = plActiva
End Sub

