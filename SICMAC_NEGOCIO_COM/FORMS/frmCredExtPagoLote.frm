VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
        L.SubItems(3) = R!nmovnro
        L.SubItems(4) = Format(R!nMonto, "#0.00")
        L.SubItems(5) = R!cUsuario
        L.SubItems(6) = R!cOpecod
        R.MoveNext
    Next i
    R.Close
    Set R = Nothing
    If LstOpExt.ListItems.Count > 0 Then
        cmdExtorno.Enabled = True
    Else
        cmdExtorno.Enabled = False
        MsgBox "No se Encontraron pagos en Lote", vbInformation, "Aviso"
    End If
End Sub

Private Sub CmdBuscar_Click()
    Call CargaPagosEnLote
End Sub

Private Sub cmdExtorno_Click()
Dim oNCred As COMNCredito.NCOMCredito
        
    If MsgBox("Se va A Extornar el Pago en Lote, Desea Continuar ?", vbInformation + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then
        Exit Sub
    End If
    Set oNCred = New COMNCredito.NCOMCredito
    Call oNCred.ExtornarPagoEnLote(CLng(LstOpExt.SelectedItem.SubItems(3)), gdFecSis, gsCodUser, gsCodAge)
    Set oNCred = Nothing
    MsgBox "Extorno de Pago en Lote Finalizado", vbInformation, "Aviso"
    LstOpExt.ListItems.Clear
    cmdExtorno.Enabled = False
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
End Sub
