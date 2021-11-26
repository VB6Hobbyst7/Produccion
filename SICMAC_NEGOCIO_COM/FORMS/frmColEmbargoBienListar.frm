VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColEmbargoBienListar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Listado de Bienes"
   ClientHeight    =   3450
   ClientLeft      =   6645
   ClientTop       =   5190
   ClientWidth     =   6075
   Icon            =   "frmColEmbargoBienListar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3450
   ScaleWidth      =   6075
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   0
      TabIndex        =   2
      Top             =   2640
      Width           =   6015
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   1560
         TabIndex        =   7
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   4440
         TabIndex        =   6
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   1560
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   3000
         TabIndex        =   3
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6015
      Begin MSComctlLib.ListView lsvBienes 
         Height          =   1695
         Left            =   120
         TabIndex        =   10
         Top             =   840
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   2990
         View            =   3
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
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bien"
            Object.Width           =   8819
         EndProperty
      End
      Begin VB.TextBox txtBuscarBien 
         Height          =   285
         Left            =   120
         TabIndex        =   9
         Text            =   "[escriba aqui para buscar en la lista]"
         Top             =   480
         Width           =   5655
      End
      Begin VB.TextBox txtBien 
         Height          =   285
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Visible         =   0   'False
         Width           =   5655
      End
      Begin VB.Label lblSubTpoBien 
         Caption         =   "SUbTipoBien"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5655
      End
   End
End
Attribute VB_Name = "frmColEmbargoBienListar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim lsBien As String
'Dim lnSubTpoBien As Integer
'Dim lnCodBien As Integer
'Public Function Inicio(ByVal psSubTpoBien As String, ByVal psSubTpoBienDesc As String) As String
'    lnSubTpoBien = psSubTpoBien
'    Me.lblSubTpoBien = "Listado: " + psSubTpoBienDesc
'    Me.Show 1
'    Inicio = lsBien
'    lsBien = "[Ingresar Bien]" + Space(100) + "0"
'End Function
'
'Private Sub cmdAceptar_Click()
'    If lsvBienes.ListItems.Count > 0 Then
'        lsBien = lsvBienes.ListItems.Item(lsvBienes.SelectedItem.Index)
'    End If
'    Unload Me
'End Sub
'
'Private Sub cmdCancelar_Click()
'    cmdNuevo.Visible = True
'    cmdAceptar.Visible = True
'
'    cmdGrabar.Visible = False
'    cmdCancelar.Visible = False
'
'    lsvBienes.Visible = True
'    Me.txtBuscarBien.Visible = True
'    Me.txtBien.Visible = False
'    Me.txtBien.Text = ""
'End Sub
'
'Private Sub cmdGrabar_Click()
'    If Me.txtBien.Text = "" Then
'        MsgBox "Ingrese el Bien en la Casilla"
'        Me.txtBien.SetFocus
'        Exit Sub
'    End If
'
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'
'
'    If MsgBox("Seguro de Registrar los Datos", vbYesNo, "Aviso") = vbYes Then
'        Dim clsMov As COMNContabilidad.NCOMContFunciones
'        Dim nCorrelativo As Integer
'        Dim sMovNro As String
'        Set clsMov = New COMNContabilidad.NCOMContFunciones
'
'
'
'        sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'        If lnCodBien = 0 Then
'            Set oColRec = New COMNColocRec.NCOMColRecCredito
'            nCorrelativo = oColRec.ObtenerBienCorrelativo(lnSubTpoBien)
'            oColRec.guardarBienEmbargo lnSubTpoBien, nCorrelativo, Me.txtBien.Text, sMovNro
'            lsBien = Me.txtBien.Text + Space(100) + CStr(nCorrelativo)
'        Else
'            oColRec.modificarBienEmbargo lnSubTpoBien, lnCodBien
'            oColRec.guardarBienEmbargo lnSubTpoBien, lnCodBien, Me.txtBien.Text, sMovNro
'            lsBien = Me.txtBien.Text + Space(100) + CStr(lnCodBien)
'        End If
'
'        MsgBox "Se ha Guardado el Bien", vbInformation, "Aviso"
'        Unload Me
'    End If
'End Sub
'
'Private Sub cmdModificar_Click()
'    Me.cmdGrabar.Visible = True
'    Me.cmdCancelar.Visible = True
'
'    Me.cmdModificar.Visible = False
'    Me.cmdNuevo.Visible = False
'    Me.cmdAceptar.Visible = False
'    Me.txtBuscarBien.Visible = False
'    Me.lsvBienes.Visible = False
'
'    Me.txtBien.Visible = True
'    Me.txtBien.Text = Trim(Left(lsvBienes.ListItems.Item(lsvBienes.SelectedItem.Index), 50))
'    Me.txtBien.SetFocus
'
'    lnCodBien = Right(lsvBienes.ListItems.Item(lsvBienes.SelectedItem.Index), 4)
'End Sub
'
'Private Sub cmdNuevo_Click()
'
'    cmdNuevo.Visible = False
'    cmdAceptar.Visible = False
'    cmdModificar.Visible = False
'    cmdGrabar.Visible = True
'    cmdCancelar.Visible = True
'
'    lsvBienes.Visible = False
'    Me.txtBuscarBien.Visible = False
'    Me.txtBien.Visible = True
'    Me.txtBien.Text = ""
'    Me.txtBien.SetFocus
'
'End Sub
'
'Private Sub cmdSalir_Click()
'    Unload Me
'End Sub
'
'Private Sub Form_Load()
'    Dim rsBienes As Recordset
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'    Dim Lst As ListItem
'    lnCodBien = 0
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set rsBienes = oColRec.ObtenerBienesEmbargo(lnSubTpoBien)
'
'    If Not rsBienes.EOF And Not rsBienes.BOF Then
'          lsvBienes.ListItems.Clear
'
'          Do While Not rsBienes.EOF
'             If Not rsBienes.BOF Then
'                    Set Lst = lsvBienes.ListItems.Add(, , rsBienes(2) + Space(100) + CStr(rsBienes(1)))
'
'              End If
'                   rsBienes.MoveNext
'           Loop
'    End If
'    rsBienes.Close
'    Set rsBienes = Nothing
'
'
'End Sub
'
'Private Sub lsvBienes_Click()
'    If lsvBienes.ListItems.Count > 0 Then
'        Me.cmdModificar.Visible = True
'    End If
'End Sub
'
'Private Sub lsvBienes_DblClick()
'    If lsvBienes.ListItems.Count > 0 Then
'        lsBien = lsvBienes.ListItems.Item(lsvBienes.SelectedItem.Index)
'        Unload Me
'    End If
'
'End Sub
'
'Private Sub lsvBienes_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.cmdAceptar.SetFocus
'    End If
'End Sub
'
'Private Sub txtBien_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        Me.cmdGrabar.SetFocus
'    End If
'End Sub
'
'Private Sub txtBuscarBien_Change()
'    Dim O_Item As ListItem
'
'    If lsvBienes.ListItems.Count > 0 Then
'        Set O_Item = lsvBienes.FindItem(Me.txtBuscarBien.Text, 0, 1, lvwPartial)
'
'        If O_Item Is Nothing Then
'           MsgBox "No se ha encontrado el elemento buscado"
'        Else
'           O_Item.EnsureVisible
'           O_Item.Selected = True
'
'        End If
'    End If
'End Sub
'
'Private Sub txtBuscarBien_Click()
'    txtBuscarBien.Text = ""
'End Sub
'
'Private Sub txtBuscarBien_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        lsvBienes.SetFocus
'    End If
'End Sub
