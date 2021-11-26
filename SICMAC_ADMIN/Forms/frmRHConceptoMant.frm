VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmRHConceptoMant 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6405
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9075
   ForeColor       =   &H00800000&
   Icon            =   "frmRHConceptoMant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6405
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraConceptos 
      Caption         =   "Conceptos Remunerativos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   5955
      Left            =   30
      TabIndex        =   7
      Top             =   0
      Width           =   9030
      Begin VB.Frame fraDatos 
         Caption         =   "Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2700
         Left            =   75
         TabIndex        =   26
         Top             =   210
         Width           =   8895
         Begin VB.TextBox txtForEdit 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Courier"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1140
            Left            =   75
            MultiLine       =   -1  'True
            TabIndex        =   44
            Top             =   1440
            Width           =   8730
         End
         Begin VB.TextBox txtCtaCnt 
            Appearance      =   0  'Flat
            Height          =   280
            Left            =   7170
            MaxLength       =   50
            TabIndex        =   42
            Top             =   572
            Width           =   1665
         End
         Begin VB.CheckBox ChkMesTrab 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "% Mes Trabajado"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   5790
            TabIndex        =   34
            Top             =   930
            Width           =   1560
         End
         Begin VB.TextBox txtNomImpre 
            Appearance      =   0  'Flat
            Height          =   280
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   33
            Top             =   885
            Width           =   1620
         End
         Begin VB.CheckBox CheckImp5 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            Caption         =   "A.5ta.Cat"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   7485
            TabIndex        =   32
            Top             =   915
            Width           =   1350
         End
         Begin VB.TextBox txtNemonico 
            Appearance      =   0  'Flat
            Height          =   280
            Left            =   4080
            MaxLength       =   50
            TabIndex        =   31
            Top             =   572
            Width           =   1620
         End
         Begin VB.ComboBox cmbGrupo 
            Height          =   315
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   555
            Width           =   2595
         End
         Begin VB.TextBox txtNomCon 
            Appearance      =   0  'Flat
            Height          =   280
            Left            =   750
            MaxLength       =   50
            TabIndex        =   29
            Top             =   255
            Width           =   4950
         End
         Begin VB.ComboBox cmbConcep 
            Height          =   315
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   28
            Top             =   870
            Width           =   2580
         End
         Begin VB.TextBox txtOrden 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   7170
            MaxLength       =   50
            TabIndex        =   27
            Text            =   "0"
            Top             =   253
            Width           =   630
         End
         Begin VB.Label lblFormula 
            Caption         =   "Formula"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   45
            Top             =   1230
            Width           =   1245
         End
         Begin VB.Label lblCtaCnt 
            Caption         =   "Cta Cnt :"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   5895
            TabIndex        =   43
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblNomImpre 
            Caption         =   "Impresión"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3375
            TabIndex        =   41
            Top             =   900
            Width           =   870
         End
         Begin VB.Label lblOrden 
            Caption         =   "Orden"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   5940
            TabIndex        =   40
            Top             =   268
            Width           =   615
         End
         Begin VB.Label lblNemoConep 
            Caption         =   "Nemo:"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   3420
            TabIndex        =   39
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lblGrupo 
            Caption         =   "Grupo"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   90
            TabIndex        =   38
            Top             =   585
            Width           =   615
         End
         Begin VB.Label lblNomConcep 
            Caption         =   "Nombre"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   90
            TabIndex        =   37
            Top             =   290
            Width           =   735
         End
         Begin VB.Label lblTipConcep 
            Caption         =   "Tipo:"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   90
            TabIndex        =   36
            Top             =   900
            Width           =   615
         End
         Begin VB.Label lblCodigoL 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   195
            Left            =   8010
            TabIndex        =   35
            Top             =   300
            Width           =   810
         End
      End
      Begin VB.Frame fraConcepto 
         Caption         =   "Concepto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   2895
         Left            =   75
         TabIndex        =   24
         Top             =   2970
         Width           =   5895
         Begin MSComctlLib.ListView lvwCon 
            Height          =   2610
            Left            =   60
            TabIndex        =   25
            Top             =   225
            Width           =   5760
            _ExtentX        =   10160
            _ExtentY        =   4604
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
      End
      Begin VB.Frame fraCamposTabla 
         Caption         =   "Tabla/Campos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   975
         Left            =   6045
         TabIndex        =   19
         Top             =   4890
         Width           =   2910
         Begin VB.ComboBox cmbCampos 
            Height          =   315
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   600
            Width           =   2070
         End
         Begin VB.ComboBox cmbTablas 
            Height          =   315
            Left            =   750
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   225
            Width           =   2070
         End
         Begin VB.Label lblCampos 
            Caption         =   "Campos"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   75
            TabIndex        =   23
            Top             =   660
            Width           =   675
         End
         Begin VB.Label lblTablas 
            Caption         =   "Tablas"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   60
            TabIndex        =   22
            Top             =   285
            Width           =   615
         End
      End
      Begin VB.Frame fraOperadores 
         Caption         =   "Operadores"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   1935
         Left            =   6045
         TabIndex        =   8
         Top             =   2940
         Width           =   2910
         Begin VB.ComboBox cmbOpeFec 
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   1545
            Width           =   2070
         End
         Begin VB.ComboBox cmbOpeCad 
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   1215
            Width           =   2070
         End
         Begin VB.ComboBox cmbOpeLog 
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   870
            Width           =   2070
         End
         Begin VB.ComboBox cmbOpeAri 
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   540
            Width           =   2070
         End
         Begin VB.ComboBox cmbOpeInterna 
            Height          =   315
            Left            =   735
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   225
            Width           =   2070
         End
         Begin VB.Label lblOpeInterna 
            Caption         =   "Internos"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   45
            TabIndex        =   18
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label lblOpeFec 
            Caption         =   "Fecha"
            ForeColor       =   &H00800000&
            Height          =   225
            Left            =   90
            TabIndex        =   17
            Top             =   1590
            Width           =   1335
         End
         Begin VB.Label lblOpeCad 
            Caption         =   "Cadena"
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   75
            TabIndex        =   16
            Top             =   1275
            Width           =   1215
         End
         Begin VB.Label lblOpeLog 
            Caption         =   "Logicos"
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   60
            TabIndex        =   15
            Top             =   915
            Width           =   1215
         End
         Begin VB.Label lblOpeAri 
            Caption         =   "Arimet."
            ForeColor       =   &H00800000&
            Height          =   210
            Left            =   60
            TabIndex        =   14
            Top             =   585
            Width           =   750
         End
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   350
      Left            =   8070
      TabIndex        =   4
      Top             =   6015
      Width           =   975
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   350
      Left            =   60
      TabIndex        =   3
      Top             =   6015
      Width           =   975
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   350
      Left            =   1080
      TabIndex        =   2
      Top             =   6015
      Width           =   975
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   350
      Left            =   2085
      TabIndex        =   1
      Top             =   6015
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   350
      Left            =   3105
      TabIndex        =   0
      Top             =   6015
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   350
      Left            =   60
      TabIndex        =   5
      Top             =   6015
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   350
      Left            =   1080
      TabIndex        =   6
      Top             =   6015
      Visible         =   0   'False
      Width           =   975
   End
End
Attribute VB_Name = "frmRHConceptoMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbEditado As Boolean
Dim lsCodCon As String
Dim lnTipo As TipoOpe

Dim lsCaption As String

Public Sub Ini(pnTipo As TipoOpe, psCaption As String)
    lnTipo = pnTipo
    lsCaption = psCaption
    Me.Show 1
End Sub

Private Sub CheckImp5_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbOpeInterna.SetFocus
    End If
End Sub

Private Sub ChkMesTrab_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.CheckImp5.SetFocus
    End If
End Sub

Private Sub cmbCampos_Click()
    Dim lnStar As Integer
    Dim lsTablaCampo As String
    
    If cmbCampos.ListIndex = -1 Or Not txtForEdit.Enabled Then
        cmbCampos.ListIndex = -1
        Exit Sub
    End If
    
    lnStar = txtForEdit.SelStart
    lsTablaCampo = Trim(Mid(cmbTablas.Text, 1, 30)) & "." & Trim(Mid(cmbCampos.Text, 1, 30))
    txtForEdit = Mid(txtForEdit, 1, txtForEdit.SelStart) & lsTablaCampo & Mid(txtForEdit, txtForEdit.SelStart + 1)
    
    txtForEdit.SelStart = lnStar + Len(lsTablaCampo)
    If txtForEdit.Enabled Then txtForEdit.SetFocus
    cmbCampos.ListIndex = -1
End Sub

Private Sub cmbCampos_KeyPress(KeyAscii As Integer)
    Dim lnStar As Integer
    Dim lsTablaCampo As String
End Sub

Private Sub cmbConcep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtNomImpre.SetFocus
End Sub

Private Sub cmbGrupo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.txtNemonico.SetFocus
End Sub

Private Sub cmbOpeAri_Click()
    Dim lnStar As Integer
    If cmbOpeAri.ListIndex = -1 Or Not txtForEdit.Enabled Then
        cmbOpeAri.ListIndex = -1
        Exit Sub
    End If
    lnStar = txtForEdit.SelStart
    txtForEdit = Mid(txtForEdit, 1, txtForEdit.SelStart) & Trim(Right(Trim(Mid(Me.cmbOpeAri, 1, 30)), 2)) & Mid(txtForEdit, txtForEdit.SelStart + 1)
    
    cmbOpeAri.ListIndex = -1
    txtForEdit.SelStart = lnStar + 1
    If txtForEdit.Enabled Then txtForEdit.SetFocus
End Sub

Private Sub cmbOpeAri_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmbOpeLog.SetFocus
End Sub

Private Sub cmbOpeCad_Click()
    Dim lnStar As Integer
    If cmbOpeCad.ListIndex = -1 Or Not txtForEdit.Enabled Then
        cmbOpeCad.ListIndex = -1
        Exit Sub
    End If
    
    lnStar = txtForEdit.SelStart + 1
    txtForEdit = Mid(txtForEdit, 1, txtForEdit.SelStart) & Trim(Mid(Me.cmbOpeCad, 1, 30)) & Mid(txtForEdit, txtForEdit.SelStart + 1)
    cmbOpeCad.ListIndex = -1
    txtForEdit.SelStart = lnStar
    txtForEdit.SelLength = 5
    
    If txtForEdit.Enabled Then txtForEdit.SetFocus
End Sub

Private Sub cmbOpeCad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmbOpeFec.SetFocus
End Sub

Private Sub cmbOpeFec_Click()
    Dim lnStar As Integer
    Dim lnLen As Integer
    If cmbOpeFec.ListIndex = -1 Or Not txtForEdit.Enabled Then
        cmbOpeFec.ListIndex = -1
        Exit Sub
    End If
    
    lnStar = txtForEdit.SelStart
    txtForEdit = Mid(txtForEdit, 1, txtForEdit.SelStart) & Trim(Right(Trim(Left(Me.cmbOpeFec.Text, 100)), 10)) & Mid(txtForEdit, txtForEdit.SelStart + 1)
    
    lnLen = Len(Trim(Right(Trim(Left(Me.cmbOpeFec.Text, 100)), 10)))
    
    cmbOpeFec.ListIndex = -1
    txtForEdit.SelStart = lnStar
    txtForEdit.SelLength = lnLen
    
    If txtForEdit.Enabled Then txtForEdit.SetFocus
End Sub

Private Sub cmbOpeFec_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmbTablas.SetFocus
End Sub

Private Sub cmbOpeInterna_Click()
    Dim lnStar As Integer
    
    If cmbOpeInterna.ListIndex = -1 Or Not txtForEdit.Enabled Then
        cmbOpeInterna.ListIndex = -1
        Exit Sub
    End If
    
    lnStar = txtForEdit.SelStart
    txtForEdit = Mid(txtForEdit, 1, txtForEdit.SelStart) & Trim(Mid(cmbOpeInterna, 1, 30)) & Mid(txtForEdit, txtForEdit.SelStart + 1)
    If Trim(Mid(cmbOpeInterna, 1, 2)) = "SI" Then
        txtForEdit.SelStart = lnStar + Len(Trim(Mid(cmbOpeInterna, 1, 30))) - 3
    Else
        txtForEdit.SelStart = lnStar + InStr(Trim(Mid(cmbOpeInterna, 1, 30)), "(")
        If InStr(Trim(Mid(cmbOpeInterna, 1, 30)), ")") - InStr(Trim(Mid(cmbOpeInterna, 1, 30)), "(") - 1 > 0 Then txtForEdit.SelLength = InStr(Trim(Mid(cmbOpeInterna, 1, 30)), ")") - InStr(Trim(Mid(cmbOpeInterna, 1, 30)), "(") - 1
    End If
    cmbOpeInterna.ListIndex = -1
    
    txtForEdit.SetFocus
End Sub

Private Sub cmbOpeInterna_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmbOpeAri.SetFocus
    End If
End Sub

Private Sub cmbOpeLog_Click()
    Dim lnStar As Integer
    Dim lsValor As String
    If cmbOpeLog.ListIndex = -1 Or Not txtForEdit.Enabled Then
        cmbOpeLog.ListIndex = -1
        Exit Sub
    End If
    lnStar = txtForEdit.SelStart
    
    lsValor = Trim(Right(Trim(Mid(cmbOpeLog, 1, 30)), 6))
    txtForEdit = Mid(txtForEdit, 1, txtForEdit.SelStart) & lsValor & Mid(txtForEdit, txtForEdit.SelStart + 1)
    
    If lsValor = "SI(,,)" Then
        lnStar = lnStar + 3
    ElseIf lsValor = "()" Then
        lnStar = lnStar + 1
    Else
        lnStar = lnStar + Len(lsValor)
    End If
    cmbOpeLog.ListIndex = -1
    txtForEdit.SelStart = lnStar
    If txtForEdit.Enabled Then txtForEdit.SetFocus
End Sub

Private Sub cmbOpeLog_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmbOpeCad.SetFocus
End Sub

Private Sub cmbTablas_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Me.cmbCampos.SetFocus
End Sub

Private Sub CmdCancelar_Click()
    ClearScreen
    Activa False
End Sub

Private Sub cmdEliminar_Click()
    Dim lsCodigo As String
    Dim oCon As NRHConcepto
    
    If Me.txtNemonico.Text = "" Then Exit Sub
    
    Set oCon = New NRHConcepto
    If lvwCon.SelectedItem Is Nothing Then Exit Sub
    
    If MsgBox("Desea Eliminar el Concepto ?, el Borrado será logico", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    
    
    oCon.EliminaConcepto lvwCon.SelectedItem
    lvwCon.ListItems.Remove lvwCon.SelectedItem.Index
End Sub

Private Sub cmdGrabar_Click()
    Dim llAux As ListItem
    Dim lsCod As String
    Dim lsNum As String
    Dim lsNumPorMes As String
    Dim oCon As NRHConcepto
    Set oCon = New NRHConcepto
    Dim lsIdNemonico As String
    
    If Not Valida Then Exit Sub
    
    If MsgBox("Desea Guradar los cambios ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    lsNum = IIf(CheckImp5.value = 1, "1", "0")
    lsNumPorMes = IIf(Me.ChkMesTrab.value = 1, "1", "0")
    
    lsIdNemonico = Mid(Me.cmbGrupo.Text, InStr(1, Me.cmbGrupo.Text, "[") + 1, InStr(1, Me.cmbGrupo.Text, "]") - InStr(1, Me.cmbGrupo.Text, "[") - 1)
    
    Me.txtForEdit.Text = Replace(Me.txtForEdit.Text, Chr$(13), "")
    Me.txtForEdit.Text = Replace(Me.txtForEdit.Text, Chr$(10), "")
    Me.txtForEdit.Text = Replace(Me.txtForEdit.Text, Chr$(12), "")
     
    If lbEditado Then
        oCon.ModificaConcepto Me.lblCodigoL.Caption, Me.txtNomCon.Text, Trim(Right(cmbConcep, 5)), Me.txtForEdit.Text, Me.txtOrden.Text, Me.txtNomImpre.Text, lsNum, lsNumPorMes, Me.txtCtaCnt.Text, lsIdNemonico & Me.txtNemonico.Text, GetMovNro(gsCodUser, gsCodAge)
    Else
        oCon.AgregaConcepto Me.lblCodigoL.Caption, Me.txtNomCon.Text, Trim(Right(cmbConcep, 5)), Me.txtForEdit.Text, Me.txtOrden.Text, Me.txtNomImpre.Text, lsNum, lsNumPorMes, Me.txtCtaCnt.Text, lsIdNemonico & Me.txtNemonico.Text, GetMovNro(gsCodUser, gsCodAge), Right(Me.cmbGrupo.Text, 1)
    End If
    
    If lbEditado Then lvwCon.ListItems.Remove lvwCon.SelectedItem.Index
    
    Activa False
    lbEditado = False
    GetData
End Sub

Private Sub cmdImprimir_Click()
    Dim oPrevio As Previo.clsPrevio
    Dim oCon As NRHConcepto
    Dim lsCadena As String
    Set oCon = New NRHConcepto
    Set oPrevio = New Previo.clsPrevio

    
    lsCadena = oCon.GetReporte(gsNomAge, gsEmpresa, gdFecSis, gsCodUser)
        
    If lsCadena <> "" Then oPrevio.Show lsCadena, Caption, True, 66
    
    Set oPrevio = Nothing
    Set oCon = Nothing
End Sub

Private Sub cmdModificar_Click()
    If Me.txtNemonico.Text = "" Then Exit Sub
    If Me.lvwCon.ListItems.Count = 0 Then Exit Sub
    lbEditado = True
    Activa True
End Sub

Private Sub Activa(pbValor As Boolean)
    Me.fraDatos.Enabled = pbValor
    cmdCancelar.Visible = pbValor
    cmdGrabar.Visible = pbValor
    cmdNuevo.Visible = Not pbValor
    cmdModificar.Visible = Not pbValor
    cmdEliminar.Enabled = Not pbValor
    cmdImprimir.Enabled = Not pbValor
    cmdSalir.Enabled = Not pbValor
    'lvwCon.Enabled = Not pbValor
End Sub

Private Sub cmdNuevo_Click()
    ClearScreen
    lbEditado = False
    Activa True
    txtNomCon.SetFocus
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

'Borrado Fisico de Conceptos
Private Sub cmdSalir_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
    Dim lsCodigo As String
    Dim sqlD As String
    
'    If Shift = 7 And Button = 2 Then
'        lsCodigo = lvwCon.SelectedItem.ListSubItems(2)
'
'        If MsgBox("Desea Eliminar el Concepto ?, el Borrado será logico", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
'
'        sqlD = "Delete Concepto Where cConcepCod = '" & lsCodigo & "'"
'
'        dbCmact.BeginTrans
'        dbCmact.Execute sqlD
'        dbCmact.CommitTrans
'
'        lvwCon.ListItems.Remove lvwCon.SelectedItem.Index
'    End If
    
End Sub

Private Sub Form_Load()
    Dim sqlR As String
    Dim rsR As New ADODB.Recordset
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    
    Caption = lsCaption
    
    Set rsR = oCon.GetConstante(6028)
    CargaCombo rsR, Me.cmbConcep
    
    rsR.Close
    Set rsR = oCon.GetConstante(6029)
    CargaCombo rsR, Me.cmbOpeAri
    
    rsR.Close
    Set rsR = oCon.GetConstante(6030)
    CargaCombo rsR, Me.cmbOpeCad
    
    rsR.Close
    Set rsR = oCon.GetConstante(6031)
    CargaCombo rsR, Me.cmbOpeLog
    
    rsR.Close
    Set rsR = oCon.GetConstante(6032)
    CargaCombo rsR, Me.cmbOpeFec

    rsR.Close
    Set rsR = oCon.GetConstante(6010)
    CargaCombo rsR, Me.cmbGrupo
    
    rsR.Close
    Set rsR = oCon.GetConstante(6034)
    CargaCombo rsR, Me.cmbOpeInterna
    
    GetData
    
    rsR.Close
    Set rsR = Nothing
    Activa False
    
    If lnTipo = gTipoOpeConsulta Then
        Me.cmdModificar.Enabled = False
        Me.cmdNuevo.Enabled = False
        Me.cmdEliminar.Enabled = False
        Me.cmdImprimir.Enabled = False
    ElseIf lnTipo = gTipoOpeReporte Then
        Me.cmdModificar.Enabled = False
        Me.cmdNuevo.Enabled = False
        Me.cmdEliminar.Enabled = False
        Me.fraConcepto.Enabled = False
        Me.fraOperadores.Enabled = False
    End If
End Sub

Private Sub GetData()
    Dim llAux As ListItem
    Dim rsD As ADODB.Recordset
    Dim oCon As DRHConcepto
    Set oCon = New DRHConcepto
    Set rsD = New ADODB.Recordset

    Set rsD = oCon.GetConceptos
    
    'rsD.Fields.Count
    lvwCon.ColumnHeaders.Clear
    lvwCon.ListItems.Clear
    lvwCon.HideColumnHeaders = False
    lvwCon.ColumnHeaders.Add , , "Codigo", 1000
    lvwCon.ColumnHeaders.Add , , "Nombre", 4000
    lvwCon.ColumnHeaders.Add , , "Tipo", 1
    lvwCon.ColumnHeaders.Add , , "Nemonico", 2200
    lvwCon.ColumnHeaders.Add , , "Formula", 5000
    lvwCon.ColumnHeaders.Add , , "Orden", 1
    lvwCon.ColumnHeaders.Add , , "5ta.Cat", 1
    lvwCon.ColumnHeaders.Add , , "Nom.Impre", 1
    lvwCon.ColumnHeaders.Add , , "Porcentaje Mes", 1
    lvwCon.ColumnHeaders.Add , , "CtaCnt", 2000, 1
    lvwCon.View = lvwReport
    
    If Not RSVacio(rsD) Then
        While Not rsD.EOF
            Set llAux = lvwCon.ListItems.Add(, , Trim(rsD!codigo))
            llAux.Bold = True
            llAux.SubItems(1) = Trim(rsD!Descrip)
            llAux.SubItems(2) = Trim(rsD!Tipo)
            llAux.SubItems(3) = Trim(rsD!Nemo)
            llAux.SubItems(4) = Trim(rsD!Formula)
            llAux.SubItems(5) = Trim(rsD!Orden)
            llAux.SubItems(6) = IIf(rsD!Imp5, "1", "0")
            llAux.SubItems(7) = IIf(IsNull(rsD!Impre), "", Trim(rsD!Impre))
            llAux.SubItems(8) = IIf(rsD!MesTrab, "1", "0")
            llAux.SubItems(9) = IIf(IsNull(rsD!CtaCont), "", rsD!CtaCont)
            rsD.MoveNext
        Wend
    End If
    
    rsD.Close
    Set rsD = oCon.GetTablasAlias("1")
    CargaCombo rsD, Me.cmbTablas
    
    rsD.Close
    Set rsD = Nothing
End Sub

Private Sub cmbTablas_Click()
    Dim oTab As DRHConcepto
    Set oTab = New DRHConcepto
    Dim rsR As ADODB.Recordset
    Set rsR = New ADODB.Recordset
    
    If cmbTablas.ListIndex = -1 Then Exit Sub
    
    Set rsR = oTab.GetTablasAlias(Trim(Right(Me.cmbTablas, 3)))
    CargaCombo rsR, cmbCampos, , 2, 1
    
    rsR.Close
    Set rsR = Nothing
End Sub

Private Sub lvwCon_DblClick()
    Dim lsValor As String
    Dim lnStar As Integer
    On Error GoTo ERROR
    lnStar = txtForEdit.SelStart
    lsValor = lvwCon.SelectedItem.SubItems(3)
    txtForEdit = Mid(txtForEdit, 1, txtForEdit.SelStart) & lsValor & Mid(txtForEdit, txtForEdit.SelStart + 1)
    
    txtForEdit.SelStart = lnStar + Len(lsValor)
    If txtForEdit.Enabled Then txtForEdit.SetFocus
ERROR:
End Sub

Private Sub lvwCon_ItemClick(ByVal Item As MSComctlLib.ListItem)
    If Not Me.fraDatos.Enabled Then
        lsCodCon = lvwCon.SelectedItem
        Me.lblCodigoL.Caption = lvwCon.SelectedItem
        txtNomCon.Text = lvwCon.SelectedItem.ListSubItems(1)
        txtNemonico.Text = Mid(lvwCon.SelectedItem.ListSubItems(3), 3)
        txtForEdit.Text = lvwCon.SelectedItem.ListSubItems(4)
        txtOrden.Text = lvwCon.SelectedItem.ListSubItems(5)
        UbicaCombo cmbConcep, lvwCon.SelectedItem.ListSubItems(2)
        UbicaCombo Me.cmbGrupo, Left(lsCodCon, 1)
        CheckImp5.value = CCur(lvwCon.SelectedItem.ListSubItems(6))
        ChkMesTrab.value = CCur(lvwCon.SelectedItem.ListSubItems(8))
        txtNomImpre.Text = lvwCon.SelectedItem.ListSubItems(7)
        Me.txtCtaCnt.Text = lvwCon.SelectedItem.ListSubItems(9)
    End If
End Sub

Private Sub lvwCon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then lvwCon_DblClick
End Sub

Private Sub txtCtaCnt_GotFocus()
    txtCtaCnt.SelStart = 0
    txtCtaCnt.SelLength = 50
End Sub

Private Sub txtCtaCnt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmbConcep.SetFocus
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub

Private Sub txtForEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdGrabar.Visible Then
            cmdGrabar.SetFocus
        Else
            Me.lvwCon.SetFocus
        End If
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtNemonico_GotFocus()
    txtNemonico.SelStart = 0
    txtNemonico.SelLength = 30
End Sub

Private Sub txtNemonico_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtCtaCnt.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtNomCon_GotFocus()
    txtNomCon.SelStart = 0
    txtNomCon.SelLength = 100
End Sub

'master@fi.upm.es
Private Sub txtNomCon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtOrden.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub ClearScreen()
    Me.txtForEdit = ""
    Me.cmbConcep.ListIndex = -1
    Me.cmbGrupo.ListIndex = -1
    Me.txtNomCon = ""
    Me.txtNemonico = ""
    Me.txtNomImpre = ""
    Me.CheckImp5.value = 0
End Sub

Private Function Valida() As Boolean
    If Me.cmbConcep.ListIndex = -1 Then
        MsgBox "Debe Ingresar el Tipo de Concepto.", vbInformation, "Aviso"
        cmbConcep.SetFocus
        Valida = False
    ElseIf cmbGrupo.ListIndex = -1 Then
        MsgBox "Debe Ingresar el Grupo.", vbInformation, "Aviso"
        cmbGrupo.SetFocus
        Valida = False
    ElseIf Trim(txtNomCon) = "" Then
        MsgBox "Debe Ingresar el Tipo de Concepto.", vbInformation, "Aviso"
        txtNomCon.SetFocus
        Valida = False
    ElseIf Trim(Me.txtNomImpre) = "" Then
        MsgBox "Debe Ingresar el nombre de impresion del Concepto.", vbInformation, "Aviso"
        txtNomImpre.SetFocus
        Valida = False
    ElseIf Not IsNumeric(Trim(Me.txtOrden)) Then
        MsgBox "Debe Ingresar el Numero de Orden del Concepto.", vbInformation, "Aviso"
        txtOrden.SetFocus
        Valida = False
    Else
        Valida = True
    End If

End Function

Private Sub txtNomImpre_GotFocus()
    txtNomImpre.SelStart = 0
    txtNomImpre.SelLength = 20
End Sub

Private Sub txtNomImpre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.ChkMesTrab.SetFocus
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtOrden_Change()
    If txtOrden = "" Then txtOrden = "0"
End Sub

Private Sub txtOrden_GotFocus()
    txtOrden.SelStart = 0
    txtOrden.SelLength = 50
End Sub

Private Sub txtOrden_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmbGrupo.SetFocus
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub


