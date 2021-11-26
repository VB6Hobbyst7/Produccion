VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmModifyActivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MODIFCAR ACTIVO FIJO"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8940
   Icon            =   "frmModifyActivo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   8940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtYear 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2520
      MaxLength       =   4
      TabIndex        =   10
      Top             =   5040
      Width           =   1170
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   7560
      TabIndex        =   9
      Top             =   2760
      Width           =   960
   End
   Begin Sicmact.TxtBuscar TxtSerie 
      Height          =   330
      Left            =   1800
      TabIndex        =   1
      Top             =   570
      Width           =   1965
      _ExtentX        =   2778
      _ExtentY        =   661
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
      TipoBusqueda    =   2
      TipoBusPers     =   1
   End
   Begin VB.Frame FraAct 
      Caption         =   "Activo Fijo"
      Height          =   3255
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   8655
      Begin VB.TextBox txtMontoIni 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         MaxLength       =   30
         TabIndex        =   20
         Text            =   "0.00"
         Top             =   1320
         Width           =   1125
      End
      Begin VB.TextBox txtMarca 
         Height          =   285
         Left            =   1650
         TabIndex        =   14
         Top             =   2040
         Width           =   3855
      End
      Begin VB.TextBox txtSeriePlaca 
         Height          =   285
         Left            =   1650
         TabIndex        =   13
         Top             =   2760
         Width           =   3855
      End
      Begin VB.TextBox txtModelo 
         Height          =   285
         Left            =   1650
         TabIndex        =   12
         Top             =   2400
         Width           =   3855
      End
      Begin VB.TextBox txtSerieNew 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3720
         TabIndex        =   8
         Top             =   600
         Width           =   2535
      End
      Begin VB.TextBox txtNomAct 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1680
         MaxLength       =   100
         TabIndex        =   4
         Top             =   240
         Width           =   5850
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   345
         Left            =   6360
         TabIndex        =   3
         Top             =   2760
         Width           =   960
      End
      Begin VB.TextBox txtDescripcion 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1680
         MaxLength       =   45
         TabIndex        =   2
         Top             =   960
         Width           =   4770
      End
      Begin Sicmact.TxtBuscar txtAF 
         Height          =   315
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1530
         _ExtentX        =   2699
         _ExtentY        =   556
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
      End
      Begin MSMask.MaskEdBox mskFecIng 
         Height          =   285
         Left            =   1665
         TabIndex        =   15
         Top             =   1680
         Width           =   1305
         _ExtentX        =   2302
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label lblMontoIni 
         AutoSize        =   -1  'True
         Caption         =   "Monto Ini. Anual :"
         Height          =   195
         Left            =   120
         TabIndex        =   21
         Top             =   1350
         Width           =   1245
      End
      Begin VB.Label lnlFecIng 
         AutoSize        =   -1  'True
         Caption         =   "F. Ingreso :"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   1725
         Width           =   795
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Marca :"
         Height          =   195
         Left            =   90
         TabIndex        =   18
         Top             =   2040
         Width           =   540
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Modelo :"
         Height          =   195
         Left            =   90
         TabIndex        =   17
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Serie/Placa :"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   2760
         Width           =   930
      End
      Begin VB.Label lblDescri 
         Caption         =   "Descripción :"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   960
         Width           =   1095
      End
      Begin VB.Label lblSerie 
         Caption         =   "Serie :"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   620
         Width           =   495
      End
   End
   Begin VB.Label lblYear 
      Caption         =   "Año : "
      Height          =   255
      Left            =   2040
      TabIndex        =   11
      Top             =   5055
      Width           =   375
   End
End
Attribute VB_Name = "frmModifyActivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lcBSCod As String
Dim lcSerie As String
Dim ldFecha As Date
Dim lsSerieMod As String
'Dim lnMIni As Currency
'Dim lnMHist As Currency
'Dim lnMAjus As Currency
'Dim lnVida_Util As Integer
'Dim lnPerDep As Integer
'Dim lsAreaAge As String
'Dim lsPersCod As String
'Dim lsComentario As String

Dim lbIngreso As Boolean
Dim lbModifica As Boolean
Dim lbElimina As Boolean

Private Sub cmdGrabar_Click()
    Dim oAF As DMov
    Set oAF = New DMov
    Dim La As String
    Dim lnMovNro As Long
    Dim lsMovNro As String
    Dim lnMovNroAjus As Long
    Dim lsMovNroAjus As String
    Dim lsCtaCont As String
    Dim oOpe As DOperacion
    Set oOpe = New DOperacion
    
    Dim lsCtaOpeBS As String
   
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
   Dim psDescrip As String
    'If Not Valida Then Exit Sub
    
    '*** PEAC 20120328
    If Trim(Me.txtSerie.Text) <> Trim(Me.txtSerieNew.Text) Then
        If oALmacen.BuscaNuevaSerie(Me.txtSerieNew.Text, psDescrip) Then
            MsgBox "El nuevo codigo de serie que trata de ingresar ya existe con el producto" + Chr(10) + psDescrip + ".", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    '*** FIN PEAC
    
    If MsgBox("Desea guardar los cambios ?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
        oAF.BeginTrans
           
            oAF.ActualizaActivo Me.txtAF.Text, Me.txtSerie.Text, Me.txtSerieNew.Text, Me.txtDescripcion.Text, Val(Me.txtYear.Text), Me.txtMontoIni.Text, Format(CDate(Me.mskFecIng.Text), "yyyymmdd"), Me.txtMarca.Text, Me.txtModelo.Text, Me.txtSeriePlaca.Text
         
        oAF.CommitTrans
      
      MsgBox "Los datos se grabaron correctamente.", vbInformation, "Aviso"
      
      LimpiaCampos
      
    Set oAF = Nothing
    'Unload Me

End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Dim oGen As DGeneral
    Set oGen = New DGeneral


    Me.txtAF.rs = oALmacen.GetAFBienes
    'Me.txtAF.rs = oALmacen.GetBienesActivosFijos()
        
    Set oALmacen = Nothing
    'Me.mskFecIng = "01/01/2002"
    If lbModifica Or lbElimina Then
        Me.txtAF.Text = lcBSCod
        txtAF_EmiteDatos
    End If
    
End Sub

Private Sub mskFecIng_GotFocus()
    Me.mskFecIng.SelStart = 0
    Me.mskFecIng.SelLength = 50
End Sub

Private Sub mskFecIng_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMarca.SetFocus
    End If
End Sub

Private Sub txtAF_EmiteDatos()

    Dim rsAlma As ADODB.Recordset
    Set rsAlma = New ADODB.Recordset

    'If txtAF.psDescripcion <> "" Then
    '    Me.txtNomAct.Text = Mid(txtAF.psDescripcion, InStr(txtAF.psDescripcion, " ") + 1, InStr(txtAF.psDescripcion, "[") - 7)
    '    Me.txtSerie.Text = Mid(txtAF.psDescripcion, InStr(txtAF.psDescripcion, "[") + 1, InStr(txtAF.psDescripcion, "]") - InStr(txtAF.psDescripcion, "[") - 1)
    ' Else
    '    Me.txtNomAct.Text = ""
    '    Me.txtSerie.Text = ""
    'End If
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    If txtAF.Text <> "" Then
        Me.txtNomAct.Text = txtAF.psDescripcion
        
        '*** PEAC 20120327
        LimpiaCampos

        Set rsAlma = oALmacen.GetAFBSSerieMod(txtAF.Text, Val(Me.txtYear.Text))
        If (rsAlma.EOF And rsAlma.BOF) Then
            MsgBox "Este Bien NO tiene Series Asignadas.", vbInformation, "Aviso"
        End If
        Me.txtSerie.rs = rsAlma
        '*** FIN PEAC
        
    End If
    
    Set oALmacen = Nothing
    'Me.mskFecIng = txtAF.psdActivacion
    'Me.txtYear.Text = Format(txtAF.psnAnio, "#,##0")
End Sub

Private Sub txtDescripcion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Me.txtMontoIni.Enabled = False Then
            Me.txtMarca.SetFocus
        Else
            Me.txtMontoIni.SetFocus
        End If
        
    End If
End Sub

Private Sub txtMarca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtModelo.SetFocus
    End If
End Sub

Private Sub txtModelo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtSeriePlaca.SetFocus
    End If
End Sub

Private Sub txtMontoIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFecIng.SetFocus
    Else
        KeyAscii = NumerosDecimales(Me.txtMontoIni, KeyAscii, 15)
    End If
End Sub

Private Sub txtMontoIni_LostFocus()
    txtMontoIni.SelStart = 0
    txtMontoIni.SelLength = 50
End Sub

Private Sub txtSerie_EmiteDatos()

'*** PEAC 20120327
'    Me.txtDescripcion.Text = TxtSerie.psDescripcion
'    Me.txtSerieNew.Text = Me.TxtSerie.Text
    
Dim oALmacen As DLogAlmacen
Set oALmacen = New DLogAlmacen
    
If Not (txtSerie.rs Is Nothing) Then
    If Not (txtSerie.rs.EOF And txtSerie.rs.BOF) Then
        Me.txtDescripcion.Text = txtSerie.rs(1)
        Me.txtSerieNew.Text = Me.txtSerie.rs(0)
        Me.txtMarca.Text = Me.txtSerie.rs(4)
        Me.txtModelo.Text = Me.txtSerie.rs(5)
        Me.txtSeriePlaca.Text = Me.txtSerie.rs(6)
        Me.txtMontoIni.Text = Me.txtSerie.rs(3)
        Me.mskFecIng.Text = Me.txtSerie.rs(2)
        
        If oALmacen.BuscaSiSerieFueDepre(Me.txtSerieNew.Text) Then
            MsgBox "Esta Serie ya fue depreciada, solo podrá modificar algunos campos.", vbInformation, "Aviso"
            Me.txtSerieNew.Enabled = False
            Me.txtMontoIni.Enabled = False
            Me.mskFecIng.Enabled = False
            Me.txtDescripcion.SetFocus
        Else
            Me.txtSerieNew.Enabled = True
            Me.txtMontoIni.Enabled = True
            Me.mskFecIng.Enabled = True
            Me.txtSerieNew.SetFocus
        End If
    End If
End If


'*** FIN PEAC
    
End Sub

Private Sub txtSerie_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtSerieNew.SetFocus
    End If
End Sub

Private Sub txtSerieNew_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtDescripcion.SetFocus
    End If
End Sub

Private Sub txtSeriePlaca_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
'    If KeyAscii <> Asc("9") Then
'        'KeyAscii = 8 es el retroceso o BackSpace
'        If KeyAscii <> 8 Then
'            KeyAscii = 0
'        End If
'    End If
End Sub
 Private Sub LimpiaCampos()
        txtSerie.Text = ""
        Me.txtSerieNew.Text = ""
        Me.txtDescripcion.Text = ""
        Me.txtMontoIni.Text = 0#
        Me.mskFecIng.Text = "__/__/____"
        Me.txtMarca.Text = ""
        Me.txtModelo.Text = ""
        Me.txtSeriePlaca.Text = ""
 End Sub
