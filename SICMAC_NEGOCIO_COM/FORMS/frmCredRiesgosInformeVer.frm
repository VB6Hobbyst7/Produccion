VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredRiesgosInformeVer 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe de Riesgos"
   ClientHeight    =   5640
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7755
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   6480
      TabIndex        =   4
      Top             =   5160
      Width           =   1215
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4935
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   8705
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Datos"
      TabPicture(0)   =   "frmCredRiesgosInformeVer.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "frm"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblNivel"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lblGlosa"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lblSalida"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblIngreso"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label4"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label5"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "ActxCta"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).ControlCount=   9
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   420
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   741
         Texto           =   "Credito"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Salida Exp.:"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   1680
         Width           =   840
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Ingreso Exp.:"
         Height          =   195
         Left            =   480
         TabIndex        =   9
         Top             =   1200
         Width           =   930
      End
      Begin VB.Label lblIngreso 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   320
         Left            =   1440
         TabIndex        =   8
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label lblSalida 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   320
         Left            =   1440
         TabIndex        =   7
         Top             =   1680
         Width           =   2175
      End
      Begin VB.Label lblGlosa 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   2055
         Left            =   1440
         TabIndex        =   6
         Top             =   2640
         Width           =   5655
      End
      Begin VB.Label lblNivel 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   1440
         TabIndex        =   5
         Top             =   2160
         Width           =   3615
      End
      Begin VB.Label frm 
         AutoSize        =   -1  'True
         Caption         =   "Nivel de Riesgo:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   2160
         Width           =   1170
      End
      Begin VB.Label Label1 
         Caption         =   "Conclusión General:"
         Height          =   435
         Left            =   480
         TabIndex        =   1
         Top             =   2640
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmCredRiesgosInformeVer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub Inicio(ByVal psCtaCod As String)
Call LlenarCampos(psCtaCod)
Me.Show 1
End Sub


Private Sub cmdCerrar_Click()
Unload Me
End Sub

Public Sub LlenarCampos(ByVal pnNroCuenta As String)
Dim oCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Set oCredito = New COMDCredito.DCOMCredito
Set rsCredito = oCredito.ObtenerInformeRiesgo(Trim(pnNroCuenta), 1)

Me.ActxCta.NroCuenta = pnNroCuenta
ActxCta.Enabled = False

If Not (rsCredito.EOF And rsCredito.BOF) Then
    Me.lblingreso.Caption = rsCredito!Ingreso
    Me.lblSalida.Caption = rsCredito!Salida
    Me.lblNivel.Caption = rsCredito!Nivel
    Me.lblGlosa.Caption = rsCredito!Glosa
End If

End Sub

Private Sub Form_Load()
Call CentraForm(Me)
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
