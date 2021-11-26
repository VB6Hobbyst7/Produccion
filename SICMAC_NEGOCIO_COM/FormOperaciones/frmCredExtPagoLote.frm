VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form frmCredExtPagoLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno de Pago en Lote"
   ClientHeight    =   3345
   ClientLeft      =   1875
   ClientTop       =   2460
   ClientWidth     =   8760
   Icon            =   "frmCredExtPagoLote.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3345
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
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
      Left            =   3000
      TabIndex        =   6
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
         Left            =   860
         TabIndex        =   9
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
         TabIndex        =   8
         Top             =   1200
         Width           =   2415
      End
      Begin VB.ComboBox cmbMotivos 
         BackColor       =   &H00C0FFFF&
         Height          =   315
         ItemData        =   "frmCredExtPagoLote.frx":030A
         Left            =   240
         List            =   "frmCredExtPagoLote.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   600
         Width           =   2415
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalles del Extorno"
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   960
         Width           =   1575
      End
      Begin VB.Label lblExtCmb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture1 
      Height          =   660
      Left            =   60
      ScaleHeight     =   600
      ScaleWidth      =   8505
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   2580
      Width           =   8565
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
         Left            =   90
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   105
         Width           =   1245
      End
      Begin VB.CommandButton cmdSalir 
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
         Height          =   375
         Left            =   7155
         TabIndex        =   3
         Top             =   105
         Width           =   1245
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Operaciones de Extorno"
      Height          =   2415
      Left            =   60
      TabIndex        =   0
      Top             =   105
      Width           =   8625
      Begin MSComctlLib.ListView LstOpExt 
         Height          =   1995
         Left            =   195
         TabIndex        =   1
         Top             =   255
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
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nª Cuenta"
            Object.Width           =   0
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
            Object.Width           =   4410
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
      End
   End
End
Attribute VB_Name = "frmCredExtPagoLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargaPagosEnLote()
Dim oDCred As COMDCredito.DCOMCreditos
Dim R As ADODB.Recordset
Dim L As ListItem
Dim i As Integer

    Set oDCred = New COMDCredito.DCOMCreditos
    Set R = oDCred.RecuperaPagosEnLote(gdFecSis)
    Set oDCred = Nothing
    LstOpExt.ListItems.Clear
    For i = 0 To R.RecordCount - 1
        Set L = LstOpExt.ListItems.Add(, , "")
        L.SubItems(1) = R!cMovDesc
        L.SubItems(2) = R!cHora
        L.SubItems(3) = R!nMovNro
        L.SubItems(4) = Format(R!nMonto, "#0.00")
        L.SubItems(5) = R!cUsuario
        L.SubItems(6) = R!cOpecod
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
    Call CargaPagosEnLote
End Sub


Private Sub cmdExtContinuar_Click()
Dim oNCred As COMNCredito.NCOMCredito
Dim sImpreBoleta As String 'JIPR 22/06/2018
Dim oPrevio As previo.clsprevio 'JIPR 22/06/2018
'****cti3
Dim DatosExtorna(1) As String

If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
    MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
    Exit Sub
End If

    If MsgBox("Se va A Extornar el Pago en Lote, Desea Continuar ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
    
        '***CTI3 (ferimoro) *****************
        frmMotExtorno.Visible = False
        Me.cmbMotivos.ListIndex = -1
        Me.txtDetExtorno.Text = ""
        '************************************
    
        Exit Sub
    End If
    Set oNCred = New COMNCredito.NCOMCredito
    
    '**** cti3
    frmMotExtorno.Visible = False
    DatosExtorna(0) = cmbMotivos.Text
    DatosExtorna(1) = txtDetExtorno.Text
    
    'JIPR sImpreBoleta 22/06/2018
    Call oNCred.ExtornarPagoEnLote(CLng(LstOpExt.SelectedItem.SubItems(3)), gdFecSis, gsCodUser, gsCodAge, sImpreBoleta, DatosExtorna)

   'JIPR 22/06/2018
   If sImpreBoleta <> "" Then
      Set oPrevio = New previo.clsprevio
      oPrevio.Show sImpreBoleta, ""
      LstOpExt.ListItems.Clear
      
        '***CTI3 (ferimoro) *****************
        frmMotExtorno.Visible = False
        Me.cmbMotivos.ListIndex = -1
        Me.txtDetExtorno.Text = ""
        Me.CmdBuscar.Enabled = True
        '************************************
      
    Else

        Set oNCred = Nothing
        MsgBox "Extorno de Pago en Lote Finalizado", vbInformation, "Aviso"
        LstOpExt.ListItems.Clear
        cmdExtorno.Enabled = False
        
        '***CTI3 (ferimoro) *****************
        frmMotExtorno.Visible = False
        Me.cmbMotivos.ListIndex = -1
        Me.txtDetExtorno.Text = ""
        Me.CmdBuscar.Enabled = True
        '************************************
        
    End If  'JIPR 22/06/2018
End Sub

Private Sub cmdExtorno_Click()
'******CTI3 (ferimoro) 27092018
 frmMotExtorno.Visible = True
 cmdExtorno.Enabled = False
 CmdBuscar.Enabled = False
 cmbMotivos.SetFocus
'******************************
End Sub

'Private Sub cmdExtorno_Click()
'Dim oNCred As COMNCredito.NCOMCredito
'Dim sImpreBoleta As String 'JIPR 22/06/2018
'Dim oPrevio As previo.clsprevio 'JIPR 22/06/2018
'
'
'    If MsgBox("Se va A Extornar el Pago en Lote, Desea Continuar ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
'        Exit Sub
'    End If
'    Set oNCred = New COMNCredito.NCOMCredito
'    'JIPR sImpreBoleta 22/06/2018
'    Call oNCred.ExtornarPagoEnLote(CLng(LstOpExt.SelectedItem.SubItems(3)), gdFecSis, gsCodUser, gsCodAge, sImpreBoleta)
'
'   'JIPR 22/06/2018
'   If sImpreBoleta <> "" Then
'      Set oPrevio = New previo.clsprevio
'      oPrevio.Show sImpreBoleta, ""
'      LstOpExt.ListItems.Clear
'    Else
'
'    Set oNCred = Nothing
'    MsgBox "Extorno de Pago en Lote Finalizado", vbInformation, "Aviso"
'    LstOpExt.ListItems.Clear
'    cmdExtorno.Enabled = False
'    End If  'JIPR 22/06/2018
'End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    Call CargaControles
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

Private Sub txtDetExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub
