VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRCDParametro 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe RCD - Mantenimiento de Parametros"
   ClientHeight    =   4665
   ClientLeft      =   1440
   ClientTop       =   3150
   ClientWidth     =   6570
   Icon            =   "frmRCDParametro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   6570
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraLista 
      Caption         =   "Parámetros"
      Height          =   2790
      Left            =   75
      TabIndex        =   20
      Top             =   90
      Width           =   6360
      Begin MSComctlLib.ListView lstParametros 
         Height          =   2535
         Left            =   120
         TabIndex        =   0
         Top             =   180
         Width           =   6180
         _ExtentX        =   10901
         _ExtentY        =   4471
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Mes"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Fecha Emision"
            Object.Width           =   2293
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Monto Minimo"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   3
            Text            =   "Tipo Cambio"
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "U.I.T."
            Object.Width           =   1940
         EndProperty
      End
   End
   Begin VB.Frame FraDatos 
      Height          =   1005
      Left            =   60
      TabIndex        =   14
      Top             =   2880
      Width           =   6390
      Begin VB.TextBox txtUIT 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   4860
         TabIndex        =   5
         Top             =   480
         Width           =   1200
      End
      Begin VB.TextBox txtCambio 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3615
         TabIndex        =   4
         Top             =   480
         Width           =   1200
      End
      Begin VB.TextBox txtMontoMin 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   2385
         TabIndex        =   3
         Top             =   480
         Width           =   1200
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   330
         Left            =   1170
         TabIndex        =   2
         Top             =   480
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblMes 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   330
         Left            =   105
         TabIndex        =   1
         Top             =   480
         Width           =   1035
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6120
         Y1              =   420
         Y2              =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Valor UIT"
         Height          =   195
         Index           =   3
         Left            =   5130
         TabIndex        =   19
         Top             =   150
         Width           =   675
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Cambio"
         Height          =   195
         Index           =   2
         Left            =   3840
         TabIndex        =   18
         Top             =   135
         Width           =   885
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto Minimo"
         Height          =   195
         Index           =   1
         Left            =   2535
         TabIndex        =   17
         Top             =   150
         Width           =   990
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Emisión"
         Height          =   195
         Left            =   1200
         TabIndex        =   16
         Top             =   165
         Width           =   1035
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Index           =   0
         Left            =   255
         TabIndex        =   15
         Top             =   150
         Width           =   300
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   60
      TabIndex        =   13
      Top             =   3900
      Width           =   6405
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Imprimir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   3840
         TabIndex        =   10
         Top             =   195
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.CommandButton cmdsalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   5160
         TabIndex        =   11
         Top             =   180
         Width           =   1095
      End
      Begin VB.CommandButton cmdeliminar 
         Caption         =   "&Eliminar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2700
         TabIndex        =   9
         Top             =   180
         Width           =   975
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
         Height          =   405
         Left            =   120
         TabIndex        =   6
         Top             =   195
         Width           =   1155
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
         Height          =   405
         Left            =   1455
         TabIndex        =   7
         Top             =   195
         Width           =   1095
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
         Height          =   405
         Left            =   1455
         TabIndex        =   8
         Top             =   195
         Width           =   1095
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   240
         TabIndex        =   12
         Top             =   195
         Width           =   1035
      End
   End
End
Attribute VB_Name = "frmRCDParametro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lbNuevo As Boolean
Dim lsFechaRCD As String
Dim fsServConsol As String

Private Sub cboMoneda_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.cmdGrabar.Visible And Me.cmdGrabar.Enabled Then
        Me.cmdGrabar.SetFocus
    End If
End If
End Sub

Private Sub cmdCancelar_Click()
    lbNuevo = False
    InHabilitar
    lstParametros.SetFocus
End Sub

Private Sub CmdEditar_Click()
If Len(Trim(lblMes)) <> 0 Then
    lbNuevo = False
    Habilitar
    Me.txtFecha.Enabled = False
    Me.txtMontoMin.SetFocus
Else
    MsgBox "Dato no válido", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdeliminar_Click()
'Dim lsSQL  As String
'Dim loBase As COMConecta.DCOMConecta
Dim oRCD As COMDCredito.DCOMRCD
If Len(Trim(lblMes)) <> 0 Then
    If MsgBox("Desea Eliminar el Registro", vbYesNo + vbInformation, "Aviso") = vbYes Then
        lbNuevo = False
        Set oRCD = New COMDCredito.DCOMRCD
        Call oRCD.EliminaParametro(fsServConsol, Trim(Me.lblMes.Caption))
        Set oRCD = Nothing
        'lsSQL = "Delete " & fsServConsol & "RCDParametro WHERE CMES='" & Trim(Me.lblMes) & "' "
        'Set loBase = New COMConecta.DCOMConecta
        '    loBase.AbreConexion
        '    loBase.Ejecutar (lsSQL)
        'Set loBase = Nothing
        LlenaLista
    End If
Else
    MsgBox "Dato no válido", vbInformation, "Aviso"
End If

End Sub

Private Sub cmdGrabar_Click()
'Dim lsSQL As String
'Dim loBase As COMConecta.DCOMConecta
Dim oRCD As COMDCredito.DCOMRCD

If Valida = True Then
    'lsFechaRCD = Format(txtFecha.Text, "yyyymmdd")
    
    lblMes = lsFechaRCD
    If MsgBox("Desea Grabar la Información?", vbYesNo + vbInformation, "Aviso") = vbYes Then
        'Set loBase = New COMConecta.DCOMConecta
        '    loBase.AbreConexion
            Set oRCD = New COMDCredito.DCOMRCD
            If lbNuevo = True Then
        '        lsSQL = " INSERT INTO " & fsServConsol & "RCDParametro (cMes,dFecha,nMontoMin,nCambioFijo,nUIT,cCodUsu,dFecMod) " _
                    & " VALUES('" & lsFechaRCD & "','" & Format(Me.txtFecha, "mm/dd/yyyy") & "'," _
                    & txtMontoMin & "," & TxtCambio & "," & txtUIT & ",'" & gsCodUser & "','" _
                    & FechaHora(gdFecSis) & "')"
                Call oRCD.InsertaParametro(fsServConsol, lsFechaRCD, CDate(txtFecha.Text), CDbl(txtMontoMin.Text), _
                                            CDbl(txtCambio.Text), CDbl(txtUIT.Text), gsCodUser, gdFecSis)
            Else
            '    lsSQL = "UPDATE " & fsServConsol & "RCDParametro SET " _
                    & "cMes='" & lsFechaRCD & "'," & "dFecha='" & Format(Me.txtFecha, "mm/dd/yyyy") & "'," _
                    & "nMontoMin=" & Me.txtMontoMin & "," & "nCambioFijo=" & Me.txtCambio & "," _
                    & "nUIT =" & Me.txtUIT & "," _
                    & "cCodUsu='" & gsCodUser & "'," & "dFecMod='" & FechaHora(gdFecSis) & "' " _
                    & "WHERE cMes='" & Trim(lblMes) & "' "
                Call oRCD.ModificaParametro(fsServConsol, lsFechaRCD, CDate(txtFecha.Text), CDbl(txtMontoMin.Text), _
                                             CDbl(txtCambio.Text), CDbl(txtUIT.Text), gsCodUser, gdFecSis)
            End If
            'loBase.Ejecutar (lsSQL)
        'Set loBase = Nothing
        lbNuevo = False
        LlenaLista
        InHabilitar
        lstParametros.SetFocus
    End If
End If

End Sub
Private Sub Refresco()
If lstParametros.ListItems.Count > 0 Then
    lblMes = Me.lstParametros.SelectedItem
    txtFecha = Me.lstParametros.SelectedItem.SubItems(1)
    txtMontoMin = Me.lstParametros.SelectedItem.SubItems(2)
    txtCambio = Me.lstParametros.SelectedItem.SubItems(3)
    txtUIT = Me.lstParametros.SelectedItem.SubItems(4)
End If
End Sub


Private Sub cmdNuevo_Click()
lbNuevo = True
Habilitar
Limpiar
txtFecha.SetFocus
End Sub
Private Sub Limpiar()
'cboMoneda = DatoTablaCodigo("87", "42")
Me.txtFecha = Format(gdFecSis, "dd/mm/yyyy")
Me.txtCambio = "0.00"
Me.txtMontoMin = "0.00"
Me.txtUIT = "0.00"
Me.lblMes = ""
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Form_Load()
Dim loConstSist As COMDConstSistema.NCOMConstSistema
    Set loConstSist = New COMDConstSistema.NCOMConstSistema
        fsServConsol = loConstSist.LeeConstSistema(gConstSistServCentralRiesgos)
    Set loConstSist = Nothing
    
    InHabilitar
    'Call LlenaCombo("87", CboMoneda)
    LlenaLista
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
Private Sub lstParametros_Click()
    Refresco
End Sub
Private Sub lstParametros_DblClick()
    cmdEditar.value = True
End Sub

Private Sub lstParametros_GotFocus()
Refresco
End Sub

Private Sub lstParametros_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdEditar.value = True
End If
End Sub
Private Sub lstParametros_KeyUp(KeyCode As Integer, Shift As Integer)
If KeyCode = 38 Or KeyCode = 40 Then
    Refresco
End If
End Sub

Private Sub txtCambio_GotFocus()
fEnfoque txtCambio
End Sub

Private Sub txtCambio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(Me.txtCambio, KeyAscii, , 3)
If KeyAscii = 13 Then
    Me.txtUIT.SetFocus
End If
End Sub
Private Sub txtCambio_LostFocus()
Me.txtCambio = Format(Val(Me.txtCambio), "#0.000")
End Sub

Private Sub txtFecha_GotFocus()
fEnfoque txtFecha
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtMontoMin.SetFocus
End If
End Sub

Private Sub txtMontoMin_GotFocus()
fEnfoque txtMontoMin
End Sub

Private Sub txtMontoMin_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoMin, KeyAscii)
If KeyAscii = 13 Then
    txtCambio.SetFocus
End If
End Sub

Private Sub txtMontoMin_LostFocus()
Me.txtMontoMin = Format(Val(Me.txtMontoMin), "#0.00")
End Sub

Private Sub txtUIT_GotFocus()
    fEnfoque txtUIT
End Sub

Private Sub txtUIT_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(Me.txtUIT, KeyAscii)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If

End Sub
Private Sub Habilitar()
Me.FraDatos.Enabled = True
Me.fraLista.Enabled = False
Me.txtFecha.Enabled = True
Me.cmdNuevo.Visible = False
Me.cmdEditar.Visible = False
Me.cmdGrabar.Visible = True
Me.cmdCancelar.Visible = True
Me.cmdeliminar.Enabled = False
Me.cmdImprimir.Enabled = False

End Sub
Private Sub InHabilitar()
Me.FraDatos.Enabled = False
Me.fraLista.Enabled = True

Me.cmdNuevo.Visible = True
Me.cmdEditar.Visible = True
Me.cmdGrabar.Visible = False
Me.cmdCancelar.Visible = False
Me.cmdeliminar.Enabled = True
Me.cmdImprimir.Enabled = True
End Sub
Private Function Valida() As Boolean
Dim oRCD As COMDCredito.DCOMRCD

lsFechaRCD = Format(txtFecha.Text, "yyyymmdd")
If lbNuevo = True Then
    Set oRCD = New COMDCredito.DCOMRCD
    If oRCD.VerificaDuplicado(fsServConsol, lsFechaRCD) = True Then
        MsgBox "Fecha ya ha sido Ingresada", vbInformation, "Aviso"
        Valida = False
        Set oRCD = Nothing
        Exit Function
    Else
        Valida = True
    End If
    Set oRCD = Nothing
End If
If ValFecha(Me.txtFecha) = False Then
    Valida = False
    txtFecha.SetFocus
    Exit Function
End If
If Len(Trim(txtMontoMin)) = 0 Then
    Valida = False
    MsgBox "Monto Mínimo no Válido", vbInformation, "Aviso"
    txtMontoMin.SetFocus
    Exit Function
Else
    Valida = True
End If
If Len(Trim(txtCambio)) = 0 Then
    Valida = False
    MsgBox "Tipo de Cambio no Válido", vbInformation, "Aviso"
    txtCambio.SetFocus
    Exit Function
Else
    Valida = True
End If
If Len(Trim(Me.txtUIT)) = 0 Then
    Valida = False
    MsgBox "Valor UIT no Válido", vbInformation, "Aviso"
    txtUIT.SetFocus
    Exit Function
Else
    Valida = True
End If
End Function
Private Sub txtUIT_LostFocus()
Me.txtUIT = Format(Val(Me.txtUIT), "#0.00")
End Sub
Private Sub LlenaLista()
'Dim lsSQL As String
Dim lrs As New ADODB.Recordset
'Dim loBase As COMConecta.DCOMConecta
Dim Item As ListItem
Dim oRCD As COMDCredito.DCOMRCD

Me.lstParametros.ListItems.Clear
Set oRCD = New COMDCredito.DCOMRCD
Set lrs = oRCD.ObtenerParametros(fsServConsol)
Set oRCD = Nothing

'lsSQL = "Select * from " & fsServConsol & "RCDParametro Order By dFecha Desc"

'Set loBase = New COMConecta.DCOMConecta
'    loBase.AbreConexion
'    Set lrs = loBase.CargaRecordSet(lsSQL)
'Set loBase = Nothing
    
    If Not (lrs.BOF And lrs.EOF) Then
        Do While Not lrs.EOF
            Set Item = Me.lstParametros.ListItems.Add(, , lrs!cMes)
            Item.SubItems(1) = Format(lrs!dFecha, "dd/mm/yyyy")
            Item.SubItems(2) = Format(lrs!nMontoMin, "#0.00")
            Item.SubItems(3) = Format(lrs!nCambioFijo, "#0.000")
            Item.SubItems(4) = Format(lrs!nUIT, "#0.00")
            lrs.MoveNext
        Loop
    End If
    lrs.Close
    Set lrs = Nothing

End Sub
'Private Function VerificaDuplicado(lsMes As String) As Boolean
'Dim lsSQL  As String
'Dim rs As ADODB.Recordset
'Dim loBase As COMConecta.DCOMConecta
'lsSQL = "Select cMes From " & fsServConsol & "RCDParametro where cMes='" & Trim(lsMes) & "'"
'
'Set loBase = New COMConecta.DCOMConecta
'    loBase.AbreConexion
'    Set rs = loBase.CargaRecordSet(lsSQL)
'Set loBase = Nothing
'    If rs.BOF And rs.EOF Then
'        VerificaDuplicado = False
'    Else
'        VerificaDuplicado = True
'    End If
'    rs.Close
'    Set rs = Nothing
'End Function
