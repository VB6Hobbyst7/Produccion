VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContingDesestimar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contingencias: Desestimar"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6105
   Icon            =   "frmContingDesestimar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   345
      Left            =   1440
      TabIndex        =   14
      Top             =   4320
      Width           =   1050
   End
   Begin VB.CommandButton cmdDesestimar 
      Caption         =   "Desestimar"
      Height          =   345
      Left            =   120
      TabIndex        =   13
      Top             =   4320
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTabConting 
      Height          =   4020
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   7091
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Desestimar Contingencia"
      TabPicture(0)   =   "frmContingDesestimar.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).ControlCount=   2
      Begin VB.Frame Frame2 
         Caption         =   " Desestimación "
         Height          =   1575
         Left            =   120
         TabIndex        =   15
         Top             =   2280
         Width           =   5535
         Begin VB.TextBox txtGlosaDesest 
            Height          =   735
            Left            =   1080
            MultiLine       =   -1  'True
            TabIndex        =   19
            Top             =   670
            Width           =   4280
         End
         Begin VB.Label Label16 
            Caption         =   "Fecha :"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   350
            Width           =   735
         End
         Begin VB.Label lblFecDesest 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1080
            TabIndex        =   17
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label9 
            Caption         =   "Glosa : "
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame Frame1 
         Caption         =   " Datos Generales "
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   5535
         Begin VB.Label txtDesc 
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   615
            Left            =   1080
            TabIndex        =   11
            Top             =   990
            Width           =   4275
         End
         Begin VB.Label Label7 
            Caption         =   "Descripción : "
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   975
            Width           =   975
         End
         Begin VB.Label lblUsuario 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   4080
            TabIndex        =   10
            Top             =   645
            Width           =   1275
         End
         Begin VB.Label Label5 
            Caption         =   "Usuario :"
            Height          =   255
            Left            =   3120
            TabIndex        =   9
            Top             =   690
            Width           =   735
         End
         Begin VB.Label lblFecRegistro 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   4080
            TabIndex        =   8
            Top             =   315
            Width           =   1275
         End
         Begin VB.Label Label3 
            Caption         =   "F. Registro :"
            Height          =   255
            Left            =   3120
            TabIndex        =   7
            Top             =   345
            Width           =   855
         End
         Begin VB.Label lblOrigen 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1080
            TabIndex        =   6
            Top             =   315
            Width           =   1875
         End
         Begin VB.Label Label2 
            Caption         =   "Monto : "
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   690
            Width           =   735
         End
         Begin VB.Label Label1 
            Caption         =   "Origen :"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   350
            Width           =   735
         End
         Begin VB.Label lblProvision 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1650
            TabIndex        =   3
            Top             =   645
            Width           =   1305
         End
         Begin VB.Label lblMoneda 
            Alignment       =   2  'Center
            BackColor       =   &H80000004&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1080
            TabIndex        =   2
            Top             =   645
            Width           =   495
         End
      End
   End
End
Attribute VB_Name = "frmContingDesestimar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmContingDesestimar
'** Descripción : Desestimacion de Contingencias creado segun RFC056-2012 (Nuevo)
'** Creación : JUEZ, 20120709 09:00:00 AM
'********************************************************************

Option Explicit
Dim rs As ADODB.Recordset
Dim oConting As DContingencia
Dim oGen As DGeneral
Dim sNumRegistro As String
Dim psOpeCod As String
Dim nTipoConting As Integer

Public Sub Inicio(ByVal psNumRegistro As String)
    sNumRegistro = psNumRegistro
    psOpeCod = gDesestimarConting
    Set oConting = New DContingencia
    Set rs = oConting.BuscaContingenciaParaDesestimar(sNumRegistro)
    lblOrigen.Caption = rs!cOrigen
    lblFecRegistro.Caption = Format(rs!dFechaReg, "dd/mm/yyyy")
    lblMoneda.Caption = rs!cmoneda
    lblProvision.Caption = Format(rs!nMontoAprox, "#,##0.00")
    lblUsuario.Caption = rs!cUsuReg
    txtDesc.Caption = rs!cContigDesc
    
    lblFecDesest.Caption = gdFecSis
    
    Me.Show 1
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub cmdDesestimar_Click()
    If Len(Trim(txtGlosaDesest.Text)) = 0 Then
        MsgBox "Falta ingresar la glosa", vbInformation, "Aviso"
        txtGlosaDesest.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Está seguro de Desestimar la Contingencia? ", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Set oConting = New DContingencia
    Dim sMensaje As String
    Dim oMov As DMov
    Set oMov = New DMov
    gsMovNro = oMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    gsGlosa = "Desestimar Contingencia"
    
    Call oConting.DesestimaContingencia(sNumRegistro, Trim(txtGlosaDesest.Text), gsMovNro)
    
    oMov.BeginTrans
    oMov.InsertaMov gsMovNro, psOpeCod, Trim(txtGlosaDesest.Text)
    oMov.CommitTrans
    Set oMov = Nothing
    
    MsgBox "Se ha desestimado con éxito la Contingencia", vbInformation, "Aviso"
    
    Unload Me
End Sub

