VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmBNDBaja 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Baja de Bien No Depreciable"
   ClientHeight    =   2835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   Icon            =   "frmBNDBaja.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2835
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   7890
      TabIndex        =   15
      Top             =   2430
      Width           =   1125
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   9090
      TabIndex        =   14
      Top             =   2430
      Width           =   1125
   End
   Begin VB.Frame fraOpe 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Baja de Activo Fijo"
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
      Height          =   2325
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   10200
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   1005
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   1785
         Width           =   9105
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   1005
         TabIndex        =   1
         Top             =   1365
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Sicmact.TxtBuscar txtAgeO 
         Height          =   345
         Left            =   990
         TabIndex        =   3
         Top             =   855
         Width           =   1590
         _extentx        =   2566
         _extenty        =   609
         appearance      =   0
         appearance      =   0
         font            =   "frmBNDBaja.frx":030A
         enabled         =   0   'False
         appearance      =   0
         enabledtext     =   0   'False
      End
      Begin Sicmact.TxtBuscar txtSerie 
         Height          =   315
         Left            =   6750
         TabIndex        =   4
         Top             =   495
         Width           =   3360
         _extentx        =   5927
         _extenty        =   556
         appearance      =   0
         appearance      =   0
         font            =   "frmBNDBaja.frx":0336
         appearance      =   0
         tipobusqueda    =   2
         lbultimainstancia=   0   'False
      End
      Begin Sicmact.TxtBuscar txtBS 
         Height          =   345
         Left            =   990
         TabIndex        =   5
         Top             =   480
         Width           =   1590
         _extentx        =   2593
         _extenty        =   609
         appearance      =   0
         appearance      =   0
         font            =   "frmBNDBaja.frx":0362
         appearance      =   0
      End
      Begin VB.Label lblSerie 
         Caption         =   "Serie :"
         Height          =   225
         Left            =   6735
         TabIndex        =   13
         Top             =   270
         Width           =   810
      End
      Begin VB.Label lblOrigen 
         Caption         =   "Origen"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   210
         Left            =   135
         TabIndex        =   12
         Top             =   225
         Width           =   2205
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   10080
         Y1              =   1290
         Y2              =   1275
      End
      Begin VB.Line Line1 
         X1              =   45
         X2              =   10095
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Label lblAgeOG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2610
         TabIndex        =   11
         Top             =   870
         Width           =   7500
      End
      Begin VB.Label lblAgeO 
         Caption         =   "Agencia O"
         Height          =   180
         Left            =   105
         TabIndex        =   10
         Top             =   930
         Width           =   840
      End
      Begin VB.Label lblBien 
         Caption         =   "Bien :"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   510
         Width           =   840
      End
      Begin VB.Label lblBienG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2610
         TabIndex        =   8
         Top             =   495
         Width           =   4155
      End
      Begin VB.Label lblComentario 
         Caption         =   "Coment."
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   780
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha"
         Height          =   225
         Left            =   150
         TabIndex        =   6
         Top             =   1410
         Width           =   810
      End
   End
End
Attribute VB_Name = "frmBNDBaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnMovNroIni As Long
Dim lnAnio As Long

Private Sub cmdGrabar_Click()
    Dim oMov As DMov
    Set oMov = New DMov
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsMovNro As String
    Dim lnMovNro As Long
        
    If Me.txtBS.Text = "" Then
        MsgBox "Debe ingresar un codigo de Bien.", vbInformation, "Aviso"
        Me.txtBS.SetFocus
        Exit Sub
    ElseIf Me.txtSerie.Text = "" Then
        MsgBox "Debe ingresar una serie valida.", vbInformation, "Aviso"
        Me.txtSerie.SetFocus
        Exit Sub
    ElseIf Not IsDate(mskFecha.Text) Then
        MsgBox "Debe ingresar una fecha valida.", vbInformation, "Aviso"
        Me.mskFecha.SetFocus
        Exit Sub
    ElseIf txtComentario.Text = "" Then
        MsgBox "Debe ingresar un comentario valido.", vbInformation, "Aviso"
        Me.txtComentario.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Desea grabar la Transferencia del Activo Fijo", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oMov.BeginTrans
        lsMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        oMov.InsertaMov lsMovNro, gnBajaAF, Me.txtComentario.Text
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        oMov.InsertaMovBSBND Me.txtBS.Text, Me.txtSerie.Text, lnMovNro
        oALmacen.AFActualizaBaja lnAnio, Me.txtBS.Text, Me.txtSerie.Text, CDate(Me.mskFecha.Text)
    oMov.CommitTrans
    
    MsgBox "EL Activo Fijo " & Me.txtBS.Text & "-" & Me.txtSerie.Text & " ha sido dado de baja ", vbInformation, "Aviso"
    
    Me.txtBS.Text = ""
    Me.txtSerie.Text = ""
    Me.lblBienG.Caption = ""
    Me.lblSerie.Caption = ""
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    
    Me.mskFecha.Text = gdFecSis
    
    Me.txtAgeO.rs = oArea.GetAgenciasAreas
    Me.txtBS.rs = oALmacen.GetBNDBienes
End Sub

Private Sub mskFecha_GotFocus()
    mskFecha.SelStart = 0
    mskFecha.SelLength = 50
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtComentario.SetFocus
    End If
End Sub

Private Sub txtAgeO_EmiteDatos()
    Me.lblAgeOG.Caption = txtAgeO.psDescripcion
End Sub

Private Sub txtBS_EmiteDatos()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    If txtBS.Text <> "" Then
        Me.lblBienG.Caption = txtBS.psDescripcion
        Me.txtSerie.rs = oALmacen.GetBNDBSSerie(txtBS.Text)
    End If
    
    Set oALmacen = Nothing
End Sub

Private Sub txtComentario_GotFocus()
    txtComentario.SelStart = 0
    txtComentario.SelLength = 300
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGrabar.SetFocus
    End If
End Sub

Private Sub txtSerie_EmiteDatos()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If txtBS.Text <> "" And txtSerie.Text <> "" Then
        Set rs = oALmacen.GetBNDBSDetalle(txtBS.Text, txtSerie.Text)
        Me.txtAgeO.Text = rs.Fields(0) & rs.Fields(1)
        txtAgeO_EmiteDatos
    End If
    
    Set oALmacen = Nothing
    Set rs = Nothing
End Sub



