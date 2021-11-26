VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColRecComision 
   Appearance      =   0  'Flat
   Caption         =   "Recuperaciones - Comisión de Cobranza"
   ClientHeight    =   5220
   ClientLeft      =   1920
   ClientTop       =   2895
   ClientWidth     =   7650
   Icon            =   "frmColRecComision.frx":0000
   LinkTopic       =   "Form1"
   Moveable        =   0   'False
   ScaleHeight     =   5220
   ScaleWidth      =   7650
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraDatosComision 
      Height          =   825
      Left            =   120
      TabIndex        =   12
      Top             =   3840
      Width           =   7485
      Begin VB.ComboBox cboCategoria 
         Height          =   315
         ItemData        =   "frmColRecComision.frx":030A
         Left            =   4800
         List            =   "frmColRecComision.frx":0317
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   360
         Width           =   1950
      End
      Begin VB.TextBox txtCodigoComision 
         Height          =   285
         Left            =   5940
         TabIndex        =   14
         Top             =   375
         Visible         =   0   'False
         Width           =   705
      End
      Begin VB.ComboBox cboTipoComision 
         Height          =   315
         ItemData        =   "frmColRecComision.frx":0345
         Left            =   2520
         List            =   "frmColRecComision.frx":034F
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   360
         Width           =   1455
      End
      Begin SICMACT.EditMoney AxMoneyIni 
         Height          =   315
         Left            =   120
         TabIndex        =   22
         Top             =   360
         Width           =   1180
         _ExtentX        =   2090
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney AxMoneyFin 
         Height          =   315
         Left            =   1320
         TabIndex        =   23
         Top             =   360
         Width           =   1180
         _ExtentX        =   2090
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney AxMoneyValor 
         Height          =   315
         Left            =   3960
         TabIndex        =   24
         Top             =   360
         Width           =   825
         _ExtentX        =   1455
         _ExtentY        =   556
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BackColor       =   -2147483633
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Categoria"
         Height          =   195
         Index           =   5
         Left            =   4860
         TabIndex        =   21
         Top             =   165
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "codigo"
         Height          =   195
         Index           =   4
         Left            =   5940
         TabIndex        =   20
         Top             =   225
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Valor"
         Height          =   195
         Index           =   3
         Left            =   4080
         TabIndex        =   19
         Top             =   180
         Width           =   645
      End
      Begin VB.Label Label1 
         Caption         =   "Rango Final"
         Height          =   195
         Index           =   2
         Left            =   1440
         TabIndex        =   18
         Top             =   180
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Rango Inicial"
         Height          =   195
         Index           =   1
         Left            =   270
         TabIndex        =   17
         Top             =   180
         Width           =   1005
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Comision"
         Height          =   195
         Index           =   0
         Left            =   2520
         TabIndex        =   16
         Top             =   120
         Width           =   1005
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
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
      Height          =   375
      Left            =   4155
      TabIndex        =   9
      Top             =   4740
      Width           =   1005
   End
   Begin VB.CommandButton cmdNuevo 
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
      Height          =   375
      Left            =   1995
      TabIndex        =   8
      Top             =   4740
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancelar 
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
      Height          =   375
      Left            =   5235
      TabIndex        =   7
      Top             =   4740
      Width           =   1005
   End
   Begin VB.CommandButton cmdEditar 
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
      Height          =   375
      Left            =   3075
      TabIndex        =   6
      Top             =   4740
      Width           =   1005
   End
   Begin VB.Frame Frame2 
      Caption         =   "Comisiones"
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
      Height          =   2985
      Left            =   90
      TabIndex        =   5
      Top             =   810
      Width           =   7485
      Begin MSComctlLib.ListView lstComision 
         Height          =   2565
         Left            =   180
         TabIndex        =   11
         Top             =   270
         Width           =   7140
         _ExtentX        =   12594
         _ExtentY        =   4524
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   265
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Rango Inicial"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Rango Final"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tipo Comision"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "Valor"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "CodComision"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Categoria"
            Object.Width           =   1058
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Estudio Juridico"
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
      Height          =   645
      Left            =   90
      TabIndex        =   3
      Top             =   90
      Width           =   7485
      Begin SICMACT.TxtBuscar AxBuscarAbogado 
         Height          =   285
         Left            =   180
         TabIndex        =   10
         Top             =   270
         Width           =   1815
         _ExtentX        =   3201
         _ExtentY        =   503
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
         TipoBusqueda    =   3
         TipoBusPers     =   1
      End
      Begin VB.TextBox txtNomPers 
         Height          =   285
         Left            =   1980
         Locked          =   -1  'True
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   270
         Width           =   5355
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6315
      TabIndex        =   0
      Top             =   4740
      Width           =   1005
   End
   Begin RichTextLib.RichTextBox rtfPre 
      Height          =   195
      Left            =   90
      TabIndex        =   2
      Top             =   3960
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   344
      _Version        =   393217
      Enabled         =   -1  'True
      TextRTF         =   $"frmColRecComision.frx":0367
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Label lblTipo 
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
      ForeColor       =   &H8000000D&
      Height          =   285
      Left            =   8190
      TabIndex        =   1
      Top             =   990
      Width           =   285
   End
End
Attribute VB_Name = "frmColRecComision"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REGISTRO DE COMISIONES DE ESTUDIOS JURIDICOS
'Archivo:  frmColRecComision.frm
'LAYG   :  01/08/2001.
'Resumen:  Nos permite registrar las Comisiones de los Abogados
Option Explicit

Dim fsAccion As String

Dim fnNroComision As Integer
Dim fnRangoIni As Long, fnRangoFin As Long
Dim fnValor As Double
Dim fnTipComis As Integer, fnTipCateg As Integer

Private Sub Limpiar(ByVal pnNivel As Integer)
    If pnNivel = 2 Then
        Me.AxBuscarAbogado.Text = ""
        Me.txtNomPers.Text = ""
    End If
    Me.lstComision.ListItems.Clear
    Me.AxMoneyIni.value = 0
    Me.AxMoneyFin.value = 0
    'Me.cboTipoComision.Clear
    Me.cboTipoComision.ListIndex = -1
    Me.AxMoneyValor = 0
    'Me.cboCategoria.Clear
    Me.cboCategoria.ListIndex = -1
End Sub
    
Private Sub HabilitaControles(ByVal pbCmdNuevo As Boolean, ByVal pbCmdGrabar As Boolean, _
        ByVal pbCmdEditar As Boolean, ByVal pbCmdCancelar As Boolean, ByVal pbCmdSalir As Boolean, _
        ByVal pbAxBuscarAbogado As Boolean, ByVal pbfraDatosComision As Boolean)
    cmdNuevo.Enabled = pbCmdNuevo
    cmdGrabar.Enabled = pbCmdGrabar
    cmdEditar.Enabled = pbCmdEditar
    cmdCancelar.Enabled = pbCmdCancelar
    cmdSalir.Enabled = pbCmdSalir
    AxBuscarAbogado.Enabled = pbAxBuscarAbogado
    fraDatosComision.Enabled = pbfraDatosComision
End Sub

Private Sub AxBuscarAbogado_EmiteDatos()
    txtNomPers.Text = AxBuscarAbogado.psDescripcion
    If Len(AxBuscarAbogado.Text) > 0 Then
        Call HabilitaControles(True, False, True, False, True, True, False)
        CargaListaComision (AxBuscarAbogado)
        AxBuscarAbogado.Enabled = False
        cmdNuevo.SetFocus
    End If
End Sub

Function CargaListaComision(ByVal psCodAbogado As String) As Boolean
Dim loRegComision As COMNColocRec.NComColRecComision 'NColRecComision
Dim lrComis As New ADODB.Recordset
Dim litmX As ListItem
Dim lnItem As Integer

Dim lsmensaje As String

Set loRegComision = New COMNColocRec.NComColRecComision
    Set lrComis = loRegComision.nObtieneListaComisionAbogado(psCodAbogado, lsmensaje)
    If Trim(lsmensaje) <> "" Then
        MsgBox lsmensaje, vbInformation, "Aviso"
        Exit Function
    End If
        If lrComis.BOF And lrComis.EOF Then
            MsgBox "Estudio Juridico no tiene comisiones asignadas ", vbInformation, "Aviso"
            lstComision.ListItems.Clear
            Exit Function
        Else
            Do While Not lrComis.EOF
                lnItem = lnItem + 1
                Set litmX = lstComision.ListItems.Add(, , lnItem)
                    litmX.SubItems(1) = Format(lrComis!nRangIni, "#0.00")
                    litmX.SubItems(2) = Format(lrComis!nRangFin, "#0.00")
                    litmX.SubItems(3) = IIf(lrComis!nTipComis = 1, "Moneda", "Porcentaje")
                    litmX.SubItems(4) = Format(lrComis!nValor, "#0.00")
                    litmX.SubItems(5) = Format(lrComis!nComisionCod)
                    litmX.SubItems(6) = Format(lrComis!nCategoria)
                lrComis.MoveNext
            Loop
        End If
    Set lrComis = Nothing
Set loRegComision = Nothing

End Function

Private Sub AxMoneyFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cboTipoComision.SetFocus
End If
End Sub

Private Sub AxMoneyIni_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.AxMoneyFin.SetFocus
End If
End Sub

Private Sub AxMoneyValor_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Me.cboCategoria.SetFocus
End If
End Sub

Private Sub cboCategoria_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub

Private Sub cboTipoComision_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   AxMoneyValor.Enabled = True
   AxMoneyValor.SetFocus
End If
End Sub

Private Sub cmdcancelar_Click()
If fsAccion = "EDITAR" Then
    Call Limpiar(1)
    Call HabilitaControles(True, False, True, False, True, True, False)
    CargaListaComision (AxBuscarAbogado)
Else
    Call Limpiar(2)
    Call HabilitaControles(True, False, True, False, True, True, False)
End If
fsAccion = ""
End Sub

Private Sub cmdEditar_Click()
If lstComision.ListItems.Count > 0 Then
    fsAccion = "EDITAR"
    Call HabilitaControles(False, True, False, True, True, False, True)
End If
End Sub

Private Function ValidaDataGrabar() As Boolean
ValidaDataGrabar = True
'Valida ingreso de informacion
    If Me.AxMoneyIni.value >= Me.AxMoneyFin.value Then
       MsgBox "El Rango Inicial debe ser Menor al Rango Final", vbInformation, "Aviso"
       AxMoneyIni.SetFocus
       ValidaDataGrabar = False
       Exit Function
    End If
    If Len(AxMoneyValor.Text) = 0 Or AxMoneyValor.Text = "." Then
        MsgBox "Por favor Ingrese un Valor", vbInformation, "Aviso"
        Me.AxMoneyValor.SetFocus
        ValidaDataGrabar = False
        Exit Function
    End If
    'Categoria
    If Me.cboCategoria.Text = "SIN G.REAL" Then
        fnTipCateg = 1
    ElseIf Me.cboCategoria.Text = "CON G.REAL" Then
        fnTipCateg = 2
    ElseIf Me.cboCategoria.Text = "T. EXTRAJUDICIAL" Then
        fnTipCateg = 3
    Else
       MsgBox "Escoja un Tipo de Comisión", vbInformation, "Aviso"
       Me.cboTipoComision.SetFocus
       ValidaDataGrabar = False
       Exit Function
    End If
    
    'Tipo de Comision
    If Me.cboTipoComision.Text = "MONEDA" Then
        fnTipComis = 1
    Else
        If Me.cboTipoComision.Text = "PORCENTAJE" Then
         fnTipComis = 2
        Else
           MsgBox "Escoja un Tipo de Comisión", vbInformation, "Aviso"
           Me.cboTipoComision.SetFocus
           ValidaDataGrabar = False
           Exit Function
        End If
    End If
    
fnRangoIni = Me.AxMoneyIni.value
fnRangoFin = Me.AxMoneyFin.value
fnValor = Me.AxMoneyValor.value
If fsAccion = "EDITAR" Then
    fnNroComision = Trim(Me.txtCodigoComision.Text)
End If
End Function
Private Sub cmdGrabar_Click()
'On Error GoTo ControlError
Dim loGrabar As COMNColocRec.NComColRecComision

If ValidaDataGrabar = False Then
    Exit Sub
End If

If MsgBox(" Grabar Comision de Abogado ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
        
        Select Case fsAccion
            Case "NUEVO"
                Set loGrabar = New COMNColocRec.NComColRecComision
                    Call loGrabar.nGrabarNuevaComisionAbogado(Me.AxBuscarAbogado.Text, fnTipComis, fnValor, fnRangoIni, fnRangoFin, fnTipCateg, False)
                Set loGrabar = Nothing
            Case "EDITAR"
                Set loGrabar = New COMNColocRec.NComColRecComision
                    Call loGrabar.nGrabarModificaComisionAbogado(fnNroComision, fnTipComis, fnValor, fnRangoIni, fnRangoFin, fnTipCateg, False)
                Set loGrabar = Nothing
        End Select

        Limpiar (1)
        Call HabilitaControles(True, False, True, False, True, True, False)
        
        CargaListaComision (AxBuscarAbogado)
        cmdNuevo.SetFocus
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub

End Sub

Private Sub cmdNuevo_Click()
Dim loGen As DGeneral
    If Len(AxBuscarAbogado.Text) > 0 Then
        Call HabilitaControles(False, True, False, True, True, False, True)
        fsAccion = "NUEVO"
        Me.AxMoneyIni.value = 0
        Me.AxMoneyFin.value = 0
        'cboTipoComision.Clear
        Me.cboTipoComision.ListIndex = 0
        Me.AxMoneyValor.value = 0
        Me.AxMoneyIni.SetFocus
    End If
End Sub

Private Sub CmdSalir_Click()
     Unload Me
End Sub

Private Sub Form_Load()
    Call Limpiar(2)
    Call HabilitaControles(False, False, False, False, True, True, False)
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub lstComision_Click()
    If lstComision.ListItems.Count > 0 Then
        Me.AxMoneyIni.value = Format(Trim(lstComision.SelectedItem.SubItems(1)), "####0.00")
        Me.AxMoneyFin.value = Format(Trim(lstComision.SelectedItem.SubItems(2)), "####0.00")
        If lstComision.SelectedItem.SubItems(3) = "Moneda" Then
            Me.cboTipoComision.ListIndex = 0
        ElseIf lstComision.SelectedItem.SubItems(3) = "Porcentaje" Then
            Me.cboTipoComision.ListIndex = 1
        End If
        Me.AxMoneyValor.value = Format(Trim(lstComision.SelectedItem.SubItems(4)), "####0.00")
        Me.txtCodigoComision.Text = Trim(lstComision.SelectedItem.SubItems(5))
        'Me.cboCategoria.ListIndex = lstComision.SelectedItem.SubItems(6) - 1
        If lstComision.SelectedItem.SubItems(6) = "1" Then
            Me.cboCategoria.ListIndex = 0
        ElseIf lstComision.SelectedItem.SubItems(6) = "2" Then
            Me.cboCategoria.ListIndex = 1
        ElseIf lstComision.SelectedItem.SubItems(6) = "3" Then
            Me.cboCategoria.ListIndex = 2
        End If
        
        Call HabilitaControles(True, False, True, False, True, True, False)
        fsAccion = ""
    End If
    
End Sub


