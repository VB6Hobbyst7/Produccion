VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmMntBienDeprecia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depreciación de Activos: "
   ClientHeight    =   4905
   ClientLeft      =   1230
   ClientTop       =   1815
   ClientWidth     =   7275
   Icon            =   "frmMntBienDeprecia.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4905
   ScaleWidth      =   7275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraRubroCbo 
      Height          =   705
      Left            =   150
      TabIndex        =   28
      Top             =   60
      Visible         =   0   'False
      Width           =   6945
      Begin VB.ComboBox cboRubro 
         Height          =   315
         Left            =   975
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   240
         Width           =   5730
      End
      Begin VB.Label Label2 
         Caption         =   "Rubros :"
         Height          =   210
         Left            =   165
         TabIndex        =   29
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame fraRubroDet 
      Height          =   1965
      Left            =   150
      TabIndex        =   31
      Top             =   795
      Visible         =   0   'False
      Width           =   6945
      Begin MSComctlLib.ListView LstDetRub 
         Height          =   1665
         Left            =   90
         TabIndex        =   4
         Top             =   195
         Width           =   6750
         _ExtentX        =   11906
         _ExtentY        =   2937
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
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "CodRub"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Codigo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Item"
            Object.Width           =   970
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descripcion"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Fecha Adq."
            Object.Width           =   1764
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Ubicacion"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Vida Util"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Valor Historico"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.Frame fraRubro 
      Height          =   2700
      Left            =   150
      TabIndex        =   30
      Top             =   60
      Visible         =   0   'False
      Width           =   6945
      Begin MSComctlLib.ListView LstRubros 
         Height          =   2370
         Left            =   120
         TabIndex        =   0
         Top             =   225
         Width           =   6720
         _ExtentX        =   11853
         _ExtentY        =   4180
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
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7937
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Archivo"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame fraRubroDetDat 
      Enabled         =   0   'False
      Height          =   1380
      Left            =   150
      TabIndex        =   21
      Top             =   2730
      Visible         =   0   'False
      Width           =   6945
      Begin VB.TextBox TxtValHis 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   5595
         TabIndex        =   10
         Top             =   960
         Width           =   1215
      End
      Begin VB.TextBox TxtVidaUtil 
         Alignment       =   1  'Right Justify
         Height          =   315
         Left            =   3180
         TabIndex        =   9
         Top             =   960
         Width           =   930
      End
      Begin VB.TextBox txtUbic 
         Height          =   300
         Left            =   3930
         TabIndex        =   6
         Top             =   195
         Width           =   2880
      End
      Begin VB.TextBox TxtDesc 
         Height          =   300
         Left            =   1185
         TabIndex        =   7
         Top             =   570
         Width           =   5625
      End
      Begin VB.TextBox TxtCod 
         Height          =   300
         Left            =   1185
         TabIndex        =   5
         Top             =   195
         Width           =   1575
      End
      Begin MSMask.MaskEdBox TxtFecAdq 
         Height          =   300
         Left            =   1185
         TabIndex        =   8
         Top             =   945
         Width           =   1155
         _ExtentX        =   2037
         _ExtentY        =   529
         _Version        =   393216
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label LblValHis 
         AutoSize        =   -1  'True
         Caption         =   "Valor Historico :"
         Height          =   195
         Left            =   4410
         TabIndex        =   27
         Top             =   1020
         Width           =   1110
      End
      Begin VB.Label LblVidaUtil 
         AutoSize        =   -1  'True
         Caption         =   "VidaUtil :"
         Height          =   195
         Left            =   2490
         TabIndex        =   26
         Top             =   1020
         Width           =   630
      End
      Begin VB.Label LblFecAdq 
         AutoSize        =   -1  'True
         Caption         =   "Fecha Adq :"
         Height          =   195
         Left            =   195
         TabIndex        =   25
         Top             =   1005
         Width           =   870
      End
      Begin VB.Label LblUbic 
         AutoSize        =   -1  'True
         Caption         =   "Ubicacion :"
         Height          =   195
         Left            =   2955
         TabIndex        =   24
         Top             =   210
         Width           =   810
      End
      Begin VB.Label LblDetdes 
         AutoSize        =   -1  'True
         Caption         =   "Descripcion :"
         Height          =   195
         Left            =   180
         TabIndex        =   23
         Top             =   615
         Width           =   930
      End
      Begin VB.Label lblDetCod 
         AutoSize        =   -1  'True
         Caption         =   "Codigo :"
         Height          =   195
         Left            =   210
         TabIndex        =   22
         Top             =   225
         Width           =   585
      End
   End
   Begin VB.Frame fraRubroDat 
      Enabled         =   0   'False
      Height          =   1230
      Left            =   150
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   6945
      Begin VB.TextBox TxtNomArch 
         Height          =   330
         Left            =   1275
         TabIndex        =   2
         Top             =   705
         Width           =   5385
      End
      Begin VB.TextBox TxtRubDes 
         Height          =   330
         Left            =   1275
         TabIndex        =   1
         Top             =   270
         Width           =   5385
      End
      Begin VB.Label LblNomArch 
         Caption         =   "&Archivo :"
         Height          =   285
         Left            =   225
         TabIndex        =   20
         Top             =   765
         Width           =   885
      End
      Begin VB.Label LblDes 
         Caption         =   "Descripcion :"
         Height          =   285
         Left            =   225
         TabIndex        =   19
         Top             =   315
         Width           =   1020
      End
   End
   Begin VB.Frame Frame1 
      Height          =   675
      Left            =   150
      TabIndex        =   15
      Top             =   4080
      Width           =   6945
      Begin VB.CommandButton CmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   5280
         TabIndex        =   14
         Top             =   210
         Width           =   1440
      End
      Begin VB.CommandButton CmdElim 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   3585
         TabIndex        =   13
         Top             =   210
         Width           =   1440
      End
      Begin VB.CommandButton CmdEditar 
         Caption         =   "E&ditar"
         Height          =   375
         Left            =   1875
         TabIndex        =   12
         Top             =   210
         Width           =   1440
      End
      Begin VB.CommandButton CmdAdd 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   165
         TabIndex        =   11
         Top             =   210
         Width           =   1440
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   165
         TabIndex        =   16
         Top             =   210
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         CausesValidation=   0   'False
         Height          =   375
         Left            =   1875
         TabIndex        =   17
         Top             =   210
         Visible         =   0   'False
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmMntBienDeprecia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nAccion As Integer
Dim nTab    As Integer
Dim lbRubro As Boolean
Dim oDep    As New DAjusteDeprecia

Public Sub Inicio(pbRubro As Boolean)
lbRubro = pbRubro
Me.Show 1
End Sub

Private Sub CargaComboRubros()
Dim sSql As String
Dim R As New ADODB.Recordset
    cboRubro.Clear
    Set R = oDep.CargaAjusteDeprecia(, adLockOptimistic)
    RSLlenaCombo R, cboRubro
    RSClose R
End Sub

Private Sub CargaListaRubros()
Dim sSql As String
Dim R As New ADODB.Recordset
Dim l As ListItem
    LstRubros.ListItems.Clear
    Set R = oDep.CargaAjusteDeprecia(, adLockOptimistic)
    Do While Not R.EOF
        Set l = LstRubros.ListItems.Add(, , Trim(Str(R!nCodigo)))
        l.SubItems(1) = Trim(R!cDescrip)
        l.SubItems(2) = Trim(R!cNomArch)
        R.MoveNext
    Loop
    RSClose R
End Sub
Private Sub CargaListaDetRubros(ByVal nCod As Long)
Dim sSql As String
Dim R As New ADODB.Recordset
Dim l As ListItem
   LstDetRub.ListItems.Clear
   If cboRubro.ListIndex >= 0 Then
      Set R = oDep.CargaAjusteDepreciaDet(nCod)
      Do While Not R.EOF
          Set l = LstDetRub.ListItems.Add(, , Trim(Str(R!nCodigo)))
          l.SubItems(1) = Trim(R!cBSCod)
          l.SubItems(2) = Trim(Str(R!nItem))
          l.SubItems(3) = Trim(R!cDescrip)
          l.SubItems(4) = Format(R!dFecAdq, gsFormatoFechaView)
          l.SubItems(5) = Trim(R!cUbicacion)
          l.SubItems(6) = Trim(R!nVidaUtil)
          l.SubItems(7) = Trim(R!nValHis)
          R.MoveNext
      Loop
      RSClose R
   End If
End Sub

Private Sub cboRubro_Click()
    Call CargaListaDetRubros(Val(Right(cboRubro.Text, 6)))
End Sub

Private Sub cboRubro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        LstDetRub.SetFocus
    End If
End Sub

Private Sub cmdAceptar_Click()
Dim sSql As String
Dim nPos As Integer
Dim nMax As Long
Dim R As New ADODB.Recordset
Dim nComb As Integer
Dim nCodigo As Long
Dim l As ListItem
   If lbRubro Then
      nPos = LstRubros.SelectedItem.Index
      If nAccion = 1 Then
         oDep.InsertaAjusteDeprecia nCodigo, TxtRubDes, TxtNomArch, "", ""
         CargaListaRubros
         DoEvents
         LstRubros.ListItems(LstRubros.ListItems.Count).Selected = True
      Else
         oDep.ActualizaAjusteDeprecia LstRubros.SelectedItem.Text, TxtRubDes, TxtNomArch, "", ""
         CargaListaRubros
         LstRubros.ListItems(nPos).Selected = True
      End If
   Else
      nPos = LstDetRub.SelectedItem.Index
      If nAccion = 1 Then
         nMax = oDep.GetItemDeprecia(Trim(Right(cboRubro.Text, 2)))
         oDep.InsertaAjusteDepreciaDet Trim(Right(cboRubro.Text, 2)), nMax, Trim(TxtCod.Text), Trim(txtUbic.Text), txtDesc.Text, Format(TxtFecAdq.Text, gsFormatoFecha), TxtVidaUtil, nVal(TxtValHis.Text)
         CargaListaDetRubros Val(Right(cboRubro.Text, 6))
         DoEvents
         LstDetRub.ListItems(LstDetRub.ListItems.Count).Selected = True
      Else
         oDep.ActualizaAjusteDepreciaDet Trim(Right(cboRubro.Text, 2)), LstDetRub.SelectedItem.SubItems(2), Trim(TxtCod.Text), Trim(txtUbic), txtDesc, Format(TxtFecAdq, gsFormatoFecha), nVal(TxtVidaUtil), nVal(TxtValHis.Text)
         CargaListaDetRubros Val(Right(cboRubro.Text, 6))
         LstDetRub.ListItems(nPos).Selected = True
      End If
   End If
   ActivaBotones True
   If lbRubro Then
      LstRubros.SetFocus
   Else
      LstDetRub.SetFocus
   End If
End Sub

Private Sub CmdAdd_Click()
    nAccion = 1
   ActivaBotones False
   If lbRubro Then
       TxtRubDes.Text = ""
       TxtNomArch.Text = ""
      TxtRubDes.SetFocus
   Else
       TxtCod.Text = ""
       txtUbic.Text = ""
       txtDesc.Text = ""
       TxtFecAdq.Text = "__/__/____"
       TxtVidaUtil.Text = ""
       TxtValHis.Text = ""
       fEnfoque TxtCod
       TxtCod.SetFocus
   End If
    
End Sub

Private Sub cmdCancelar_Click()
   If lbRubro Then
      fraRubro.Enabled = True
      fraRubroDat.Enabled = False
      LstRubros.SetFocus
   Else
      fraRubroDet.Enabled = True
      fraRubroDetDat.Enabled = False
      LstDetRub.SetFocus
   End If
   ActivaBotones True
End Sub

Private Sub cmdEditar_Click()
   nAccion = 2
   ActivaBotones False
   
   If lbRubro Then
      TxtRubDes.SetFocus
   Else
      TxtCod.SetFocus
   End If
End Sub

Private Sub CmdElim_Click()
    If MsgBox("Esta Seguro que Desea Eliminar el Registro ?", vbInformation + vbYesNo, "Aviso") = vbYes Then
       If lbRubro Then
         oDep.EliminaAjusteDeprecia LstRubros.SelectedItem.Text
         LstRubros.ListItems.Remove LstRubros.SelectedItem.Index
         LstRubros.SetFocus
       Else
         oDep.EliminaAjusteDepreciaDet Trim(Right(cboRubro.Text, 2)), LstDetRub.SelectedItem.SubItems(2)
         LstDetRub.ListItems.Remove LstDetRub.SelectedItem.Index
         LstDetRub.SetFocus
       End If
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
   CentraForm Me
   Set oDep = New DAjusteDeprecia
   If lbRubro Then
      fraRubro.Visible = True
      fraRubroDat.Visible = True
      CargaListaRubros
   Else
      fraRubroDet.Visible = True
      fraRubroCbo.Visible = True
      fraRubroDetDat.Visible = True
      CargaComboRubros
      If cboRubro.ListCount > 0 Then
          cboRubro.ListIndex = 0
      End If
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oDep = Nothing
End Sub


Private Sub LstDetRub_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LstDetRub.SortKey = ColumnHeader.SubItemIndex
    LstDetRub.SortOrder = lvwAscending
    LstDetRub.Sorted = True
End Sub

Private Sub LstDetRub_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If LstDetRub.ListItems.Count > 0 Then
        TxtCod.Text = LstDetRub.SelectedItem.SubItems(1)
        txtUbic.Text = LstDetRub.SelectedItem.SubItems(5)
        txtDesc.Text = LstDetRub.SelectedItem.SubItems(3)
        TxtFecAdq.Text = LstDetRub.SelectedItem.SubItems(4)
        TxtVidaUtil.Text = LstDetRub.SelectedItem.SubItems(6)
        TxtValHis.Text = LstDetRub.SelectedItem.SubItems(7)
    End If
End Sub

Private Sub LstDetRub_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CmdAdd.SetFocus
    End If
End Sub

Private Sub LstRubros_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LstRubros.SortKey = ColumnHeader.SubItemIndex
    LstRubros.SortOrder = lvwAscending
    LstRubros.Sorted = True
End Sub

Private Sub LstRubros_DblClick()
    If LstRubros.ListItems.Count > 0 Then
        Call cmdEditar_Click
    End If
End Sub

Private Sub LstRubros_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If LstRubros.ListItems.Count > 0 Then
        TxtRubDes.Text = LstRubros.SelectedItem.SubItems(1)
        TxtNomArch.Text = LstRubros.SelectedItem.SubItems(2)
    End If
End Sub

Private Sub LstRubros_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If CmdAdd.Enabled Then
            CmdAdd.SetFocus
        End If
    End If
End Sub

Private Sub txtcod_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        fEnfoque txtUbic
        txtUbic.SetFocus
    End If
End Sub

Private Sub txtdesc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        fEnfoque TxtFecAdq
        TxtFecAdq.SetFocus
    End If
End Sub


Private Sub TxtFecAdq_KeyPress(KeyAscii As Integer)
Dim b As Boolean
    If KeyAscii = 13 Then
        Call TxtFecAdq_Validate(b)
        If Not b Then
            fEnfoque TxtVidaUtil
            TxtVidaUtil.SetFocus
        End If
    End If
End Sub

Private Sub TxtFecAdq_Validate(Cancel As Boolean)
Dim Cad As String
    Cad = ValidaFecha(TxtFecAdq.Text)
    If Len(Trim(Cad)) > 0 Then
        If cmdcancelar.Enabled Then
            MsgBox Cad, vbInformation, "Aviso"
        End If
        Cancel = True
    End If
End Sub

Private Sub TxtNomArch_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub TxtRubDes_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        fEnfoque TxtNomArch
        TxtNomArch.SetFocus
    End If
End Sub

Private Sub txtUbic_KeyPress(KeyAscii As Integer)
    KeyAscii = Letras(KeyAscii, True)
    If KeyAscii = 13 Then
        fEnfoque txtDesc
        txtDesc.SetFocus
    End If
End Sub

Private Sub TxtValHis_KeyPress(KeyAscii As Integer)
   KeyAscii = NumerosDecimales(TxtValHis, KeyAscii, 16, 2)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub

Private Sub TxtValHis_LostFocus()
    If Len(Trim(TxtValHis.Text)) = 0 Then
        TxtValHis.Text = "0.00"
    End If
End Sub

Private Sub TxtVidaUtil_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosEnteros(KeyAscii)
    If KeyAscii = 13 Then
        fEnfoque TxtValHis
        TxtValHis.SetFocus
    End If
End Sub

Private Sub TxtVidaUtil_LostFocus()
    If Len(Trim(TxtVidaUtil.Text)) = 0 Then
        TxtVidaUtil.Text = "0"
    End If
End Sub

Private Sub ActivaBotones(pbActiva As Boolean)
CmdAdd.Visible = pbActiva
cmdEditar.Visible = pbActiva
CmdElim.Visible = pbActiva
cmdAceptar.Visible = Not pbActiva
cmdcancelar.Visible = Not pbActiva
If lbRubro Then
   fraRubroDat.Enabled = Not pbActiva
   fraRubro.Enabled = pbActiva
Else
   fraRubroDetDat.Enabled = Not pbActiva
   fraRubroDet.Enabled = pbActiva
End If
End Sub
