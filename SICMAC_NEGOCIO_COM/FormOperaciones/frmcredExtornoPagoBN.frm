VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmcredExtornoPagoBN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extorno Pago BN - Convenio / Corresponsalia"
   ClientHeight    =   3765
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8655
   Icon            =   "frmcredExtornoPagoBN.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3765
   ScaleWidth      =   8655
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmMotExtorno 
      Caption         =   "Motivos del Extorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   2700
      Left            =   3120
      TabIndex        =   11
      Top             =   120
      Visible         =   0   'False
      Width           =   2845
      Begin VB.CommandButton cmdExtContinuar 
         Caption         =   "&Continuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   840
         TabIndex        =   14
         Top             =   2040
         Width           =   1095
      End
      Begin VB.TextBox txtDetExtorno 
         BackColor       =   &H00C0FFFF&
         Height          =   750
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMotivos 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmcredExtornoPagoBN.frx":030A
         Left            =   240
         List            =   "frmcredExtornoPagoBN.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalles del Extorno"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblExtCmb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture2 
      Height          =   540
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   8505
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   8565
      Begin VB.ComboBox cbotipo 
         Height          =   315
         Left            =   720
         TabIndex        =   7
         Text            =   "cbotipo"
         Top             =   120
         Width           =   1575
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
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
         Left            =   7200
         TabIndex        =   6
         Top             =   80
         Width           =   1245
      End
      Begin MSMask.MaskEdBox txtFec 
         Height          =   300
         Left            =   3960
         TabIndex        =   9
         Top             =   120
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Proceso :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   2640
         TabIndex        =   10
         Top             =   120
         Width           =   1230
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Tipo :"
         BeginProperty Font 
            Name            =   "MS Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000007&
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   120
         Width           =   435
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operaciones de Extorno"
      Height          =   2295
      Left            =   0
      TabIndex        =   3
      Top             =   600
      Width           =   8625
      Begin MSComctlLib.ListView LstOpExt 
         Height          =   1995
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   3519
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
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   11
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "N� Cuenta"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Operacion"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Hora"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Movimiento"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Monto"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Usuario"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "CodOpe"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "nPrePago"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   8
            Text            =   "nMovNroR"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   9
            Text            =   "nMovNroC"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   10
            Text            =   "cMovC"
            Object.Width           =   0
         EndProperty
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   660
      Left            =   0
      ScaleHeight     =   600
      ScaleWidth      =   8505
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   3000
      Width           =   8565
      Begin VB.CommandButton cmdSalir 
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
         Height          =   375
         Left            =   7155
         TabIndex        =   2
         Top             =   105
         Width           =   1245
      End
      Begin VB.CommandButton cmdExtorno 
         Caption         =   "&Extorno"
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
         Left            =   5820
         TabIndex        =   1
         Top             =   105
         Width           =   1245
      End
   End
End
Attribute VB_Name = "frmcredExtornoPagoBN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim id_Log As Integer
Dim sCadImpre As String


Sub llenar_cbo()
cbotipo.Clear
cbotipo.AddItem "Cobros", 0
cbotipo.AddItem "Pagos", 1
cbotipo.AddItem "BCP", 2
cbotipo.ListIndex = 0
End Sub

Private Sub CargaPagosEnLote()
Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim L As ListItem
Dim i As Integer

    Set oDCred = New COMDCredito.DCOMCreditos
    Set R = oDCred.RecuperaPagosEnLoteBN(Me.cbotipo.ListIndex, gdFecSis, CDate(Me.txtFec.Text))
    Set oDCred = Nothing
    LstOpExt.ListItems.Clear
    For i = 0 To R.RecordCount - 1
        Set L = LstOpExt.ListItems.Add(, , R!cCodCtafull)
        L.SubItems(1) = Trim(R!cMovDesc)
        L.SubItems(2) = R!cHora
        L.SubItems(3) = R!nMovNro
        L.SubItems(4) = Format(R!nMonto, "#0.00")
        L.SubItems(5) = R!cUsuario
        L.SubItems(6) = R!cOpecod
        L.SubItems(7) = R!nPrepago
        L.SubItems(8) = R!nMovNroR
        L.SubItems(9) = R!nMovNroC
        L.SubItems(10) = R!cMovC
        id_Log = R!id
        R.MoveNext
    Next i
    R.Close
    Set R = Nothing
    If LstOpExt.ListItems.count > 0 Then
        cmdExtorno.Enabled = True
    Else
        cmdExtorno.Enabled = False
        MsgBox "No se Encontraron pagos en Lote", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdBuscar_Click()
  If Me.txtFec.Text = "" Then
    MsgBox "Fecha no v�lida", vbInformation, "Aviso"
    Exit Sub
 ElseIf Not IsDate(txtFec) Then
    MsgBox "Fecha no v�lida", vbInformation, "Aviso"
    Exit Sub
 End If
        
  If Me.cbotipo.ListIndex <> -1 Then
    Call CargaPagosEnLote
  Else
    MsgBox "seleccione un tipo de Archivo", vbInformation, "Aviso"
  End If
End Sub

Private Sub cmdExtContinuar_Click()

Dim oNCred As COMNCredito.NCOMCredito
Dim i As Integer
Dim a As Integer
Dim err As Integer
Dim MatDatos(11) As String
Dim sMensaje As String
Dim oPrevio As previo.clsprevio
Dim sImpreBoleta_1 As String
Dim sImpreBoleta_2() As String
Dim sImpreBoletaAho_1() As String
Dim sImpreBoletaAho_2() As String

'****cti3
Dim DatosExtorna(1) As String
'***************CTI3  (ferimoro)  01102018
If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
    MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
    Exit Sub
End If

    If MsgBox("Se va A Extornar Todo el Proceso de Pago Autom�tico, Desea Continuar ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        '***CTI3 (ferimoro) *****************
        frmMotExtorno.Visible = False
        Me.cmbMotivos.ListIndex = -1
        Me.txtDetExtorno.Text = ""
        '************************************
        Exit Sub
    End If
    sCadImpre = ""
    
'MatDatos(0) = LstOpExt.SelectedItem.Text
    Set oNCred = New COMNCredito.NCOMCredito
    
    For i = 0 To LstOpExt.ListItems.count - 1
        If i = 0 And id_Log > 0 Then
          Call oNCred.ExtornarCredPagoCorrespBN(id_Log, Me.cbotipo.ListIndex)
        End If
        
        MatDatos(0) = LstOpExt.ListItems(i + 1).Text
        For a = 1 To 10
            MatDatos(a) = LstOpExt.ListItems(i + 1).SubItems(a)
        Next a
        
        '**** cti3
        frmMotExtorno.Visible = False
        DatosExtorna(0) = cmbMotivos.Text
        DatosExtorna(1) = txtDetExtorno.Text
               
        Call oNCred.ExtornarOperacionCredito(MatDatos, gdFecSis, gsCodUser, gsCodAge, gsNomAge, sLpt, gsInstCmac, gsCodCMAC, _
                                        gsUser, sMensaje, sImpreBoleta_1, sImpreBoleta_2, sImpreBoletaAho_1, sImpreBoletaAho_2, gbImpTMU, , , DatosExtorna)
        sCadImpre = sCadImpre & sImpreBoleta_1
        err = oNCred.ExtornarPagoAutomaticoBN(gdFecSis, gsCodUser, gsCodAge, CCur(LstOpExt.ListItems(i + 1).SubItems(4)), CLng(LstOpExt.ListItems(i + 1).SubItems(8)), CLng(LstOpExt.ListItems(i + 1).SubItems(9)), LstOpExt.ListItems(i + 1).SubItems(10), LstOpExt.ListItems(i + 1).SubItems(1), DatosExtorna)
    Next i
    
    id_Log = 0
    Set oNCred = Nothing
    
    If sMensaje <> "" Then
        MsgBox sMensaje, vbInformation, "Mensaje"
    Else
        Set oPrevio = New previo.clsprevio
        oPrevio.Show sCadImpre, ""
         Set oPrevio = Nothing
        MsgBox "Extorno Finalizado", vbInformation, "Aviso"
        'Call cmdBuscar_Click
    End If
    LstOpExt.ListItems.Clear
    cmdExtorno.Enabled = False

End Sub

Private Sub cmdExtorno_Click()
'******CTI3 (ferimoro) 27092018
 frmMotExtorno.Visible = True
 cmdExtorno.Enabled = False
 CmdBuscar.Enabled = False
 cmbMotivos.SetFocus
'******************************
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
'******CTI3 (ferimoro) 18102018
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes

Set R = oCons.ObtenerConstanteExtornoMotivo

Set oCons = Nothing
Call Llenar_Combo_MotivoExtorno(R, cmbMotivos)

End Sub
Private Sub Form_Load()
    CentraForm Me
    Call CargaControles
    sCadImpre = ""
    llenar_cbo
    id_Log = 0
    Me.txtFec.Text = "__/__/____"
    Me.txtFec.Enabled = True
End Sub

Private Sub txtDetExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub
