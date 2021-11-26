VERSION 5.00
Begin VB.Form frmAFAsiganPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignacion de Activo Fijo"
   ClientHeight    =   3435
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9000
   Icon            =   "frmAFAsiganPersona.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraOpe 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Asignacion de Activo Fijo"
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
      Height          =   2970
      Left            =   30
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
         Top             =   1980
         Width           =   7845
      End
      Begin Sicmact.TxtBuscar txtPersonaO 
         Height          =   345
         Left            =   990
         TabIndex        =   4
         Top             =   855
         Width           =   1590
         _extentx        =   2805
         _extenty        =   609
         appearance      =   0
         appearance      =   0
         font            =   "frmAFAsiganPersona.frx":030A
         enabled         =   0   'False
         appearance      =   0
         tipobusqueda    =   7
         tipobuspers     =   1
         enabledtext     =   0   'False
      End
      Begin Sicmact.TxtBuscar txtSerie 
         Height          =   315
         Left            =   6450
         TabIndex        =   5
         Top             =   495
         Width           =   2400
         _extentx        =   4233
         _extenty        =   556
         appearance      =   0
         appearance      =   0
         font            =   "frmAFAsiganPersona.frx":0336
         appearance      =   0
         tipobusqueda    =   2
      End
      Begin Sicmact.TxtBuscar txtBS 
         Height          =   345
         Left            =   990
         TabIndex        =   6
         Top             =   480
         Width           =   1605
         _extentx        =   2831
         _extenty        =   609
         appearance      =   0
         appearance      =   0
         font            =   "frmAFAsiganPersona.frx":0362
         appearance      =   0
      End
      Begin Sicmact.TxtBuscar txtPersonaD 
         Height          =   345
         Left            =   1005
         TabIndex        =   17
         Top             =   1568
         Width           =   1590
         _extentx        =   2805
         _extenty        =   609
         appearance      =   0
         appearance      =   0
         font            =   "frmAFAsiganPersona.frx":038E
         appearance      =   0
         tipobusqueda    =   7
         tipobuspers     =   1
         enabledtext     =   0   'False
      End
      Begin VB.Label lblBienG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2610
         TabIndex        =   16
         Top             =   495
         Width           =   3870
      End
      Begin VB.Label lblBien 
         Caption         =   "Bien :"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   510
         Width           =   840
      End
      Begin VB.Label lblPersonaO 
         Caption         =   "Persona O"
         Height          =   180
         Left            =   105
         TabIndex        =   14
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
         TabIndex        =   13
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
      Begin VB.Label lblPersona 
         Caption         =   "Persona D :"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   1650
         Width           =   840
      End
      Begin VB.Label lblAgeDG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2610
         TabIndex        =   11
         Top             =   1590
         Width           =   6225
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
         TabIndex        =   10
         Top             =   225
         Width           =   2205
      End
      Begin VB.Label lblDestino 
         Caption         =   "Destino"
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
         TabIndex        =   9
         Top             =   1350
         Width           =   2205
      End
      Begin VB.Label lblComentario 
         Caption         =   "Coment."
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   2010
         Width           =   780
      End
      Begin VB.Label lblSerie 
         Caption         =   "Serie :"
         Height          =   225
         Left            =   6450
         TabIndex        =   7
         Top             =   255
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   7845
      TabIndex        =   1
      Top             =   3030
      Width           =   1125
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   6645
      TabIndex        =   0
      Top             =   3030
      Width           =   1125
   End
End
Attribute VB_Name = "frmAFAsiganPersona"
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
    ElseIf Me.txtPersonaD.Text = "" Then
        MsgBox "Debe ingresar la persona de destino.", vbInformation, "Aviso"
        Me.txtPersonaD.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Desea grabar la Asignación del Activo Fijo.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oMov.BeginTrans
        lsMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        oMov.InsertaMov lsMovNro, gnAsignaAF, Me.txtComentario.Text, gMovEstContabNoContable
        lnMovNro = oMov.GetnMovNro(lsMovNro)
        oMov.InsertaMovBSAF lnAnio, lnMovNroIni, 1, Me.txtBS.Text, Me.txtSerie.Text, lnMovNro
        oMov.InsertaMovGasto lnMovNro, Me.txtPersonaO.Text, "", txtPersonaD.Text
        oALmacen.AFActualizaPersona Me.txtPersonaD.Text, lnAnio, Me.txtBS.Text, Me.txtSerie.Text
    oMov.CommitTrans
    
    MsgBox "EL Activo Fijo " & Me.txtBS.Text & "-" & Me.txtSerie.Text & " ha sido transferido a : " & Me.txtPersonaO.Text & " " & Me.lblAgeDG.Caption
    
    Unload Me
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    Me.txtBS.rs = oALmacen.GetAFBienes
End Sub

Private Sub txtPersonaD_EmiteDatos()
    Me.lblAgeDG.Caption = txtPersonaD.psDescripcion
    
    If txtPersonaD.psDescripcion <> "" Then
        Me.txtComentario.SetFocus
    End If
End Sub

Private Sub txtPersonaO_EmiteDatos()
    Me.lblAgeOG.Caption = txtPersonaO.psDescripcion
End Sub

Private Sub txtBS_EmiteDatos()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    
    If txtBS.Text <> "" Then
        Me.lblBienG.Caption = txtBS.psDescripcion
        Me.txtSerie.rs = oALmacen.GetAFBSSerie(txtBS.Text, Year(gdFecSis))
    End If
    
    Set oALmacen = Nothing
End Sub

Private Sub txtSerie_EmiteDatos()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oPer As UPersona
    Set oPer = New UPersona
    
    If txtBS.Text <> "" And txtSerie.Text <> "" Then
        Set rs = oALmacen.GetAFBSDetallePersona(txtBS.Text, txtSerie.Text)
        '*** PEAC 20110223
        If rs.EOF And rs.BOF Then
            MsgBox "No se encontró registros para esta consulta.", vbOKOnly + vbInformation, "Atención"
            txtSerie.Text = ""
            Exit Sub
        End If
        
        lnMovNroIni = rs.Fields(2)
        lnAnio = rs.Fields(3)
        Me.txtPersonaO.Text = rs.Fields(0) & ""
        If Me.txtPersonaO.Text <> "" Then
            oPer.ObtieneClientexCodigo Me.txtPersonaO.Text
            Me.lblAgeOG.Caption = oPer.sPersNombre
        End If
        Me.txtPersonaD.SetFocus
    End If
    
    Set oALmacen = Nothing
    Set rs = Nothing
End Sub

