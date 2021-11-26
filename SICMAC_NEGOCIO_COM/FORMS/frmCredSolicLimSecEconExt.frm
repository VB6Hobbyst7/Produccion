VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredSolicLimSecEconExt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno de Autorización/Rechazo"
   ClientHeight    =   4455
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7215
   Icon            =   "frmCredSolicLimSecEconExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   7215
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   360
      Left            =   5880
      TabIndex        =   22
      Top             =   3960
      Width           =   1170
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Extornar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   120
      TabIndex        =   21
      Top             =   3960
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   1440
      TabIndex        =   20
      Top             =   3960
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      _ExtentX        =   12303
      _ExtentY        =   6588
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Autorización/Rechazo de Solicitud"
      TabPicture(0)   =   "frmCredSolicLimSecEconExt.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "ActXCodCta"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdBuscar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      Begin VB.Frame Frame1 
         Caption         =   " Datos del Crédito "
         Height          =   2655
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   6735
         Begin VB.Label lblNuevoPorcCart 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5160
            TabIndex        =   19
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label12 
            Caption         =   "Nuevo % Cartera:"
            Height          =   255
            Left            =   3720
            TabIndex        =   18
            Top             =   1500
            Width           =   1335
         End
         Begin VB.Label lblCartera 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   5160
            TabIndex        =   17
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label10 
            Caption         =   "% Cartera:"
            Height          =   255
            Left            =   3720
            TabIndex        =   16
            Top             =   1140
            Width           =   1335
         End
         Begin VB.Label lblResultado 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1560
            TabIndex        =   15
            Top             =   2160
            Width           =   1935
         End
         Begin VB.Label Label5 
            Caption         =   "Resultado:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2220
            Width           =   1335
         End
         Begin VB.Label lblGlosa 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1560
            TabIndex        =   13
            Top             =   1800
            Width           =   5055
         End
         Begin VB.Label Label3 
            Caption         =   "Glosa:"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1860
            Width           =   1335
         End
         Begin VB.Label lblMontoCredMN 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1560
            TabIndex        =   11
            Top             =   1440
            Width           =   1455
         End
         Begin VB.Label Label1 
            Caption         =   "Monto Crédito MN:"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   1500
            Width           =   1335
         End
         Begin VB.Label lblCliente 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1560
            TabIndex        =   9
            Top             =   360
            Width           =   5055
         End
         Begin VB.Label lblSectorDesc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1560
            TabIndex        =   8
            Top             =   720
            Width           =   3255
         End
         Begin VB.Label lblMontoSectorMN 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   1560
            TabIndex        =   7
            Top             =   1080
            Width           =   1455
         End
         Begin VB.Label Label6 
            Caption         =   "Cliente: "
            Height          =   255
            Left            =   120
            TabIndex        =   6
            Top             =   405
            Width           =   615
         End
         Begin VB.Label Label7 
            Caption         =   "Sector:"
            Height          =   255
            Left            =   120
            TabIndex        =   5
            Top             =   765
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Monto Sector MN:"
            Height          =   255
            Left            =   120
            TabIndex        =   4
            Top             =   1140
            Width           =   1335
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   375
         Left            =   3840
         TabIndex        =   2
         Top             =   480
         Width           =   375
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   873
         Texto           =   "Credito"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmCredSolicLimSecEconExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredSolicLimSecEconExt
'** Descripción : Formulario para extornar la autorización/rechazo solicitudes de límites de
'**               créditos por Sector económico creado segun TI-ERS029-2013
'** Creación : JUEZ, 20140609 09:00:00 AM
'**********************************************************************************************

Option Explicit

Dim oDPersGen As COMDPersona.DCOMPersGeneral
Dim rs As ADODB.Recordset

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
If Len(ActXCodCta.NroCuenta) = 18 Then
    CargarDatos
Else
    MsgBox "Favor de digitar él crédito correctamente", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdAceptar_Click()
    If MsgBox("Se va a extornar la autorización/rechazo, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Set oDPersGen = New COMDPersona.DCOMPersGeneral
        Call oDPersGen.ActualizarSolicitudAutorizacionRiesgos(ActXCodCta.NroCuenta, "", 0)
    Set oDPersGen = Nothing
    cmdCancelar_Click
End Sub

Private Sub cmdBuscar_Click()
'Dim oCredito As COMDCredito.DCOMCredito
'Dim R As ADODB.Recordset
'Dim oPers As COMDPersona.UCOMPersona
'    cmdCancelar_Click
'    Set oPers = frmBuscaPersona.Inicio()
'    If Not oPers Is Nothing Then
'        Call FrmVerCredito.Inicio(oPers.sPersCod, , , True, ActXCodCta)
'        ActXCodCta.SetFocusCuenta
'    End If
'    Set oPers = Nothing
    ActXCodCta.NroCuenta = frmCredPersEstado.Inicio(Array(gColocEstSolic), "Creditos para Extornar Autorizacion/Rechazos")
    ActXCodCta.SetFocusCuenta
End Sub

Private Sub CargarDatos()
Set oDPersGen = New COMDPersona.DCOMPersGeneral
Set rs = oDPersGen.RecuperaAutorizacionRechazoRiesgos(ActXCodCta.NroCuenta)
Set oDPersGen = Nothing
    If Not (rs.EOF And rs.BOF) Then
        lblCliente.Caption = rs!cPersNombre
        lblSectorDesc.Caption = rs!cSectorDesc
        lblMontoSectorMN.Caption = Format(rs!nMontoSectorMN, "#,##0.00")
        lblMontoCredMN.Caption = Format(rs!nMontoCredMN, "#,##0.00")
        lblGlosa.Caption = rs!cGlosa
        lblResultado.Caption = rs!cResult
        lblCartera.Caption = Format(rs!PorcCartera, "#,##0.00")
        lblNuevoPorcCart.Caption = Format(rs!NuevoPorcCart, "#,##0.00")
        ActXCodCta.Enabled = False
        cmdbuscar.Enabled = False
        cmdAceptar.Enabled = True
    Else
        MsgBox "No se encontraron datos", vbInformation, "Aviso"
        ActXCodCta.NroCuenta = ""
        ActXCodCta.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    ActXCodCta.Enabled = True
    cmdbuscar.Enabled = True
    cmdAceptar.Enabled = False
    ActXCodCta.NroCuenta = ""
    lblCliente.Caption = ""
    lblSectorDesc.Caption = ""
    lblMontoSectorMN.Caption = ""
    lblMontoCredMN.Caption = ""
    lblGlosa.Caption = ""
    lblResultado.Caption = ""
    lblCartera.Caption = ""
    lblNuevoPorcCart.Caption = ""
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
