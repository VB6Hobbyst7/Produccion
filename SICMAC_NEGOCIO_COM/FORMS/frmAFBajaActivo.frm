VERSION 5.00
Begin VB.Form frmAFBajaActivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Baja de Activo Fijo"
   ClientHeight    =   2805
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9030
   Icon            =   "frmAFBajaActivo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2805
   ScaleWidth      =   9030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
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
      ForeColor       =   &H00C00000&
      Height          =   2325
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   8940
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   1005
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   1350
         Width           =   7845
      End
      Begin Sicmact.TxtBuscar txtAgeO 
         Height          =   345
         Left            =   990
         TabIndex        =   4
         Top             =   855
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   609
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
         Enabled         =   0   'False
         Enabled         =   0   'False
         Appearance      =   0
         EnabledText     =   0   'False
      End
      Begin Sicmact.TxtBuscar txtSerie 
         Height          =   315
         Left            =   6750
         TabIndex        =   5
         Top             =   495
         Width           =   2115
         _ExtentX        =   3731
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
         TipoBusqueda    =   2
      End
      Begin Sicmact.TxtBuscar txtBS 
         Height          =   345
         Left            =   990
         TabIndex        =   6
         Top             =   480
         Width           =   1470
         _ExtentX        =   2593
         _ExtentY        =   609
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
      Begin VB.Label lblComentario 
         Caption         =   "Coment."
         Height          =   210
         Left            =   120
         TabIndex        =   13
         Top             =   1380
         Width           =   780
      End
      Begin VB.Label lblBienG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2610
         TabIndex        =   12
         Top             =   495
         Width           =   4140
      End
      Begin VB.Label lblBien 
         Caption         =   "Bien :"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   510
         Width           =   840
      End
      Begin VB.Label lblAgeO 
         Caption         =   "Agencia O"
         Height          =   180
         Left            =   105
         TabIndex        =   10
         Top             =   930
         Width           =   840
      End
      Begin VB.Label lblAgeOG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2610
         TabIndex        =   9
         Top             =   870
         Width           =   6225
      End
      Begin VB.Line Line1 
         X1              =   45
         X2              =   8910
         Y1              =   1260
         Y2              =   1260
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   15
         X2              =   8940
         Y1              =   1290
         Y2              =   1290
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
         ForeColor       =   &H000000C0&
         Height          =   210
         Left            =   135
         TabIndex        =   8
         Top             =   225
         Width           =   2205
      End
      Begin VB.Label lblSerie 
         Caption         =   "Serie :"
         Height          =   225
         Left            =   6735
         TabIndex        =   7
         Top             =   270
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   7875
      TabIndex        =   1
      Top             =   2400
      Width           =   1125
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   6675
      TabIndex        =   0
      Top             =   2400
      Width           =   1125
   End
End
Attribute VB_Name = "frmAFBajaActivo"
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
    End If
    
    If MsgBox("Desa grabar la Transferencia del Activo Fijo", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oMov.BeginTrans
        lsMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        oMov.InsertaMov lsMovNro, gnBajaAF, Me.txtComentario.Text
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        oMov.InsertaMovBSAF lnAnio, lnMovNroIni, 1, Me.txtBS.Text, Me.txtSerie.Text, lnMovNro
        oALmacen.AFActualizaBaja lnAnio, Me.txtBS.Text, Me.txtSerie.Text
    oMov.CommitTrans
    
    MsgBox "EL Activo Fijo " & Me.txtBS.Text & "-" & Me.txtSerie.Text & " ha sido dado de baja ", vbInformation, "Aviso"
    
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
    
    Me.txtAgeO.rs = oArea.GetAgenciasAreas
    Me.txtBS.rs = oALmacen.GetAFBienes
End Sub

Private Sub txtAgeO_EmiteDatos()
    Me.lblAgeOG.Caption = txtAgeO.psDescripcion
End Sub

Private Sub txtBS_EmiteDatos()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    If txtBS.Text <> "" Then
        Me.lblBienG.Caption = txtBS.psDescripcion
        Me.txtSerie.rs = oALmacen.GetAFBSSerie(txtBS.Text)
    End If
    
    Set oALmacen = Nothing
End Sub

Private Sub txtSerie_EmiteDatos()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If txtBS.Text <> "" And txtSerie.Text <> "" Then
        Set rs = oALmacen.GetAFBSDetalle(txtBS.Text, txtSerie.Text)
        lnMovNroIni = rs.Fields(2)
        lnAnio = rs.Fields(3)
        Me.txtAgeO.Text = rs.Fields(0) & rs.Fields(1)
        txtAgeO_EmiteDatos
    End If
    
    Set oALmacen = Nothing
    Set rs = Nothing
End Sub


