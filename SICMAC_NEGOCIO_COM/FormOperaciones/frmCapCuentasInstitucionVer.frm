VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCapCuentasInstitucionVer 
   Caption         =   "Cuentas de la Institución"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6405
   Icon            =   "frmCapCuentasInstitucionVer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6405
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Cuentas"
      Height          =   2895
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   6375
      Begin MSComctlLib.ListView Lst 
         Height          =   2580
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   6135
         _ExtentX        =   10821
         _ExtentY        =   4551
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cuenta"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Relacion"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Estado"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   3000
      Width           =   1215
   End
   Begin VB.CommandButton cmbAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3000
      Width           =   1215
   End
End
Attribute VB_Name = "frmCapCuentasInstitucionVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'***Nombre      : frmCapCuentasInstitucionVer ----SUBIDO DESDE LA 60
'***Descripción : Formulario para mostrar las cuentas al momento de seleccionar institución
'***Creación    : MARG el 20171201, según TI-ERS 065-2017
'************************************************************************************************
Option Explicit
Dim CtlCtaCod As ActXCodCta

Public Sub Inicio(ByVal psPersCod As String, Optional ByRef pCtlCtaCod As ActXCodCta)
    Dim rs As ADODB.Recordset
    Dim clsMant As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim item As ListItem
    
    Set clsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rs = clsMant.ObtenerCuentasIntitucion(psPersCod)
    Set clsMant = Nothing
    Lst.ListItems.Clear
    Do Until rs.EOF
        Set item = Lst.ListItems.Add(, , rs!cCtaCod)
        item.SubItems(1) = rs!Relacion
        item.SubItems(2) = rs!Estado
        rs.MoveNext
    Loop
    Set rs = Nothing
'    Set CtlCtaCod = pCtlCtaCod
    Me.Show vbModal
End Sub

Private Sub cmbAceptar_Click()
    Call Lst_DblClick
End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Activate()
    Lst.SetFocus
End Sub

Private Sub Lst_DblClick()
   If Not Lst.SelectedItem Is Nothing Then
'            CtlCtaCod.NroCuenta = Lst.ListItems(Lst.SelectedItem.Index).Text
'            CtlCtaCod.SetFocusCuenta
            frmCapDepositoLote.txtCuentaInstitucion = Lst.ListItems(Lst.SelectedItem.Index).Text
            Unload Me
  End If
   
End Sub

Private Sub Lst_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call Lst_DblClick
    End If
End Sub

