VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMktActividad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Registro de Actividades"
   ClientHeight    =   6675
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8355
   Icon            =   "frmMktActividad.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6675
   ScaleWidth      =   8355
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnNuevo 
      Caption         =   "&Nuevo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   70
      TabIndex        =   8
      Top             =   6270
      Width           =   1000
   End
   Begin VB.CommandButton btnEditar 
      Caption         =   "&Editar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   1200
      TabIndex        =   9
      Top             =   6270
      Width           =   1000
   End
   Begin VB.CommandButton btnCerrar 
      Cancel          =   -1  'True
      Caption         =   "&Cerrar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7290
      TabIndex        =   10
      Top             =   6270
      Width           =   1000
   End
   Begin TabDlg.SSTab TabActividad 
      Height          =   6180
      Left            =   45
      TabIndex        =   11
      Top             =   45
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   10901
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   6
      TabHeight       =   617
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Actividades"
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feActividad"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraRegistro"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame fraRegistro 
         Caption         =   "Registro"
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
         Height          =   2855
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   8055
         Begin VB.CheckBox chkPreProg 
            Caption         =   " Pre Programado"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   5520
            TabIndex        =   2
            Top             =   720
            Width           =   1605
         End
         Begin VB.TextBox txtDescripcion 
            Height          =   840
            Left            =   1920
            MaxLength       =   255
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   5
            Top             =   1440
            Width           =   5730
         End
         Begin VB.TextBox txtActividad 
            Height          =   300
            Left            =   1920
            MaxLength       =   199
            TabIndex        =   1
            Top             =   720
            Width           =   3255
         End
         Begin VB.CommandButton btnCancelar 
            Caption         =   "&Cancelar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   6650
            TabIndex        =   7
            Top             =   2400
            Width           =   1000
         End
         Begin VB.CommandButton btnGuardar 
            Caption         =   "&Guardar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   5550
            TabIndex        =   6
            Top             =   2400
            Width           =   1000
         End
         Begin VB.ComboBox cboTpoActividad 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   340
            Width           =   3225
         End
         Begin MSMask.MaskEdBox txtFechaIni 
            Height          =   300
            Left            =   1920
            TabIndex        =   3
            Top             =   1080
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txtFechaFin 
            Height          =   300
            Left            =   5160
            TabIndex        =   4
            Top             =   1080
            Width           =   1170
            _ExtentX        =   2064
            _ExtentY        =   529
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label8 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   480
            TabIndex        =   24
            Top             =   1440
            Width           =   975
         End
         Begin VB.Label Label7 
            Caption         =   "Fecha Fin:"
            Height          =   255
            Left            =   4320
            TabIndex        =   23
            Top             =   1080
            Width           =   855
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha Inicio:"
            Height          =   255
            Left            =   480
            TabIndex        =   22
            Top             =   1080
            Width           =   975
         End
         Begin VB.Label Label1 
            Caption         =   "Tipo de Actividad:"
            Height          =   255
            Left            =   480
            TabIndex        =   14
            Top             =   360
            Width           =   1335
         End
         Begin VB.Label Label2 
            Caption         =   "Nombre Actividad:"
            Height          =   255
            Left            =   480
            TabIndex        =   13
            Top             =   720
            Width           =   1335
         End
      End
      Begin MSComctlLib.ListView lstAhorros 
         Height          =   2790
         Left            =   -74910
         TabIndex        =   15
         Top             =   495
         Width           =   9390
         _ExtentX        =   16563
         _ExtentY        =   4921
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         Enabled         =   0   'False
         NumItems        =   12
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro."
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   1
            Text            =   "Fecha"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Producto"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Agencia"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Nro. Cuenta"
            Object.Width           =   3881
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Nro. Cta Antigua"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Estado"
            Object.Width           =   2469
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Participación"
            Object.Width           =   2470
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   8
            Text            =   "SaldoCont"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   9
            Text            =   "SaldoDisp"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "Motivo de Bloque"
            Object.Width           =   7231
         EndProperty
         BeginProperty ColumnHeader(12) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   11
            Text            =   "Moneda"
            Object.Width           =   2540
         EndProperty
      End
      Begin Sicmact.FlexEdit feActividad 
         Height          =   2730
         Left            =   120
         TabIndex        =   16
         Top             =   3280
         Width           =   8025
         _ExtentX        =   14155
         _ExtentY        =   4815
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Tipo Actividad-Nombre Actividad-Prog-Fecha Inicio-Fecha Fin-Descripción-ActividadId-TipoActividadId"
         EncabezadosAnchos=   "350-2000-3000-500-1200-1200-5000-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-4-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-C-C-C-L-C-C"
         FormatosEdit    =   "0-1-1-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "SOLES"
         Height          =   195
         Left            =   -71445
         TabIndex        =   21
         Top             =   3465
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "TOTAL AHORROS"
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
         Left            =   -73185
         TabIndex        =   20
         Top             =   3465
         Width           =   1590
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "DOLARES"
         Height          =   195
         Left            =   -68475
         TabIndex        =   19
         Top             =   3465
         Width           =   765
      End
      Begin VB.Label lblSolesAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -70815
         TabIndex        =   18
         Top             =   3375
         Width           =   2145
      End
      Begin VB.Label lblDolaresAho 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
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
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   -67680
         TabIndex        =   17
         Top             =   3375
         Width           =   2145
      End
   End
End
Attribute VB_Name = "frmMktActividad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fbNuevo As Boolean
Dim fnCodActividad As Long

Private Sub Form_Load()
    CentraForm Me
    ListarTipoActividadesActivas
    MuestraActividades
    fnCodActividad = 0
    fbNuevo = True
End Sub
Private Sub btnNuevo_Click()
    limpiar
    Me.cboTpoActividad.SetFocus
End Sub
Private Sub btnCancelar_Click()
    limpiar
End Sub
Private Sub btnCerrar_Click()
    Unload Me
End Sub
Private Sub btnEditar_Click()
    If feActividad.TextMatrix(feActividad.Row, 0) = "" Then
        MsgBox "Ud. debe seleccionar la Actividad a Editar", vbInformation, "Aviso"
        feActividad.SetFocus
        Exit Sub
    End If
    
    fnCodActividad = feActividad.TextMatrix(feActividad.Row, 7)
    Me.cboTpoActividad.ListIndex = IndiceListaCombo(Me.cboTpoActividad, CLng(Trim(feActividad.TextMatrix(feActividad.Row, 8))))
    Me.txtActividad.Text = feActividad.TextMatrix(feActividad.Row, 2)
    Me.chkPreProg.value = IIf(Trim(feActividad.TextMatrix(feActividad.Row, 3)) = ".", 1, 0)
    Me.txtFechaIni.Text = feActividad.TextMatrix(feActividad.Row, 4)
    Me.txtFechaFin.Text = feActividad.TextMatrix(feActividad.Row, 5)
    Me.txtDescripcion.Text = feActividad.TextMatrix(feActividad.Row, 6)
    fbNuevo = False
End Sub
Private Sub btnGuardar_Click()
    Dim oGasto As DGastosMarketing
    Dim lsNombreAct As String, lsDescripcion As String
    Dim lbPreProgramado As Boolean
    Dim ldFecIni As Date, ldFecFin As Date
    Dim lnTpoAct As Long
    
    On Error GoTo ErrorGuardarActividad
    
    If validaGrabar = False Then Exit Sub
    Set oGasto = New DGastosMarketing
    
    lsNombreAct = UCase(Trim(Me.txtActividad.Text))
    lbPreProgramado = IIf(Me.chkPreProg.value = "1", True, False)
    ldFecIni = CDate(Me.txtFechaIni.Text)
    ldFecFin = CDate(Me.txtFechaFin.Text)
    lsDescripcion = Trim(Me.txtDescripcion.Text)
    lnTpoAct = CLng(Trim(Right(Me.cboTpoActividad.Text, 5)))
    
    If MsgBox("Esta seguro de guardar los datos de la Actividad?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    If fbNuevo Then
        Call oGasto.InsertaActividad(lsNombreAct, lbPreProgramado, ldFecIni, ldFecFin, lsDescripcion, lnTpoAct)
    Else
        If fnCodActividad = 0 Then
            MsgBox "Ud. debe seleccionar la Actividad a editar", vbInformation, "Aviso"
            Me.feActividad.SetFocus
            Exit Sub
        End If
        Call oGasto.ActualizaActividad(fnCodActividad, lsNombreAct, lbPreProgramado, ldFecIni, ldFecFin, lsDescripcion, lnTpoAct)
    End If
    
    MsgBox "Se ha grabado con éxito la Actividad", vbInformation, "Aviso"
    limpiar
    MuestraActividades
    Exit Sub
ErrorGuardarActividad:
    Err.Raise Err.Number, "Error Guardar", Err.Description
End Sub
Private Function validaGrabar() As Boolean
    Dim i As Integer
    Dim sCad As String
    validaGrabar = True
    If Me.cboTpoActividad.ListIndex = -1 Then
        MsgBox "Falta seleccionar el Tipo de Actividad", vbInformation, "Aviso"
        Me.cboTpoActividad.SetFocus
        validaGrabar = False
        Exit Function
    End If
    If Len(Trim(Me.txtActividad.Text)) = 0 Then
        MsgBox "Falta ingresar el Nombre de la Actividad", vbInformation, "Aviso"
        Me.txtActividad.SetFocus
        validaGrabar = False
        Exit Function
    End If
    If Len(Trim(Me.txtFechaIni.Text)) <> 10 Then
        MsgBox "Falta ingresar la Fecha de Inicio de la Actividad", vbInformation, "Aviso"
        Me.txtFechaIni.SetFocus
        validaGrabar = False
        Exit Function
    End If
    sCad = ValidaFecha(txtFechaIni.Text)
    If Not Trim(sCad) = "" Then
        MsgBox sCad, vbInformation, "Aviso"
        txtFechaIni.SetFocus
        validaGrabar = False
        Exit Function
    End If
    If Len(Trim(Me.txtFechaFin.Text)) <> 10 Then
        MsgBox "Falta ingresar la Fecha de Fin de la Actividad", vbInformation, "Aviso"
        Me.txtFechaFin.SetFocus
        validaGrabar = False
        Exit Function
    End If
    sCad = ValidaFecha(txtFechaFin.Text)
    If Not Trim(sCad) = "" Then
        MsgBox sCad, vbInformation, "Aviso"
        txtFechaFin.SetFocus
        validaGrabar = False
        Exit Function
    End If
    If CDate(Me.txtFechaFin.Text) < CDate(Me.txtFechaIni.Text) Then
        MsgBox "La Fecha de Fin de la Actividad NO puede ser menor que la de Inicio", vbInformation, "Aviso"
        Me.txtFechaFin.SetFocus
        validaGrabar = False
        Exit Function
    End If
    If Len(Trim(Me.txtDescripcion.Text)) = 0 Then
        MsgBox "Falta ingresar la Descripción de la Actividad", vbInformation, "Aviso"
        Me.txtDescripcion.SetFocus
        validaGrabar = False
        Exit Function
    End If
    If Not (feActividad.Rows - 1 = 1 And feActividad.TextMatrix(1, 0) = "") Then
        For i = 1 To feActividad.Rows - 1
            'Que no sean iguales el Nombre, Tipo Actividad, Fecha Inicio y Fecha Fin de la Actividad
            If CLng(Trim(feActividad.TextMatrix(i, 8))) = CLng(Trim(Right(Me.cboTpoActividad.Text, 5))) And Trim(feActividad.TextMatrix(i, 4)) = Trim(Me.txtFechaIni.Text) And Trim(feActividad.TextMatrix(i, 5)) = Trim(Me.txtFechaFin.Text) Then
                If Trim(feActividad.TextMatrix(i, 2)) = Trim(Me.txtActividad.Text) Then
                    If fbNuevo Then
                        MsgBox "La Actividad que se está creando ya existe como Tipo de Actividad y coincide con fechas de Inicio y Fin, verifique", vbInformation, "Aviso"
                        feActividad.SetFocus
                        feActividad.Row = i
                        feActividad.Col = 1
                        validaGrabar = False
                        Exit Function
                    Else
                        If CLng(Trim(feActividad.TextMatrix(i, 7))) <> fnCodActividad Then
                            MsgBox "La Actividad que se está editando ya existe como Tipo de Actividad y coincide con fechas de Inicio y Fin, verifique", vbInformation, "Aviso"
                            feActividad.SetFocus
                            feActividad.Row = i
                            feActividad.Col = 2
                            validaGrabar = False
                            Exit Function
                        End If
                    End If
                End If
            End If
        Next
    End If
End Function
Private Sub limpiar()
    Me.cboTpoActividad.ListIndex = -1
    Me.txtActividad.Text = ""
    Me.chkPreProg.value = 0
    Me.txtFechaIni.Text = "__/__/____"
    Me.txtFechaFin.Text = "__/__/____"
    Me.txtDescripcion.Text = ""
    fnCodActividad = 0
    fbNuevo = True
End Sub
Private Sub cboTpoActividad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtActividad.SetFocus
    End If
End Sub
Private Sub txtActividad_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        Me.chkPreProg.SetFocus
    End If
End Sub
Private Sub chkPreProg_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        Me.txtFechaIni.SetFocus
    End If
End Sub
Private Sub txtFechaIni_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        Me.txtFechaFin.SetFocus
    End If
End Sub
'Private Sub txtFechaIni_LostFocus()
'    Dim sCad As String
'    sCad = ValidaFecha(txtFechaIni.Text)
'    If Not Trim(sCad) = "" Then
'        MsgBox sCad, vbInformation, "Aviso"
'        txtFechaIni.SetFocus
'        Exit Sub
'    End If
'End Sub
Private Sub txtFechaFin_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        Me.txtDescripcion.SetFocus
    End If
End Sub
'Private Sub txtFechaFin_LostFocus()
'    Dim sCad As String
'    sCad = ValidaFecha(txtFechaFin.Text)
'    If Not Trim(sCad) = "" Then
'        MsgBox sCad, vbInformation, "Aviso"
'        txtFechaFin.SetFocus
'        Exit Sub
'    End If
'End Sub
Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras2(KeyAscii, True)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Me.btnGuardar.SetFocus
    End If
End Sub
Private Sub ListarTipoActividadesActivas()
    Dim oGasto As DGastosMarketing
    Dim rs As ADODB.Recordset
    Set oGasto = New DGastosMarketing
    Set rs = New ADODB.Recordset
    
    Set rs = oGasto.RecuperaTipoActividadxEstado(True)
    Do While Not rs.EOF
        Me.cboTpoActividad.AddItem Trim(rs!cNombre) & Space(200) & Trim(rs!nId)
        rs.MoveNext
    Loop
End Sub
Private Sub MuestraActividades()
    Dim oGasto As DGastosMarketing
    Dim rsActividad As ADODB.Recordset
    Set oGasto = New DGastosMarketing
    Set rsActividad = New ADODB.Recordset
    
    Call FormateaFlex(feActividad)
    Set rsActividad = oGasto.RecuperaActividad
    If Not RSVacio(rsActividad) Then
        Do While Not rsActividad.EOF
            feActividad.AdicionaFila
            feActividad.TextMatrix(feActividad.Row, 1) = rsActividad!cTipoActNombre
            feActividad.TextMatrix(feActividad.Row, 2) = rsActividad!cNombre
            feActividad.TextMatrix(feActividad.Row, 3) = IIf(rsActividad!bPreProgramado = True, "1", "")
            feActividad.TextMatrix(feActividad.Row, 4) = Format(rsActividad!dFechaIni, "dd/mm/yyyy")
            feActividad.TextMatrix(feActividad.Row, 5) = Format(rsActividad!dFechaFin, "dd/mm/yyyy")
            feActividad.TextMatrix(feActividad.Row, 6) = Format(rsActividad!cDescripcion, "dd/mm/yyyy")
            feActividad.TextMatrix(feActividad.Row, 7) = rsActividad!nId
            feActividad.TextMatrix(feActividad.Row, 8) = rsActividad!nTipoActividadId
            rsActividad.MoveNext
        Loop
    End If
    Set oGasto = Nothing
    Set rsActividad = Nothing
End Sub
