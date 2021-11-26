VERSION 5.00
Begin VB.Form frmLogProvCtas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "frmLogProvCtas.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   3210
      TabIndex        =   9
      Top             =   2955
      Width           =   1110
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   360
      Left            =   4380
      TabIndex        =   8
      Top             =   2955
      Width           =   1110
   End
   Begin VB.Frame fraCta 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Cuentas"
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
      Height          =   1605
      Left            =   30
      TabIndex        =   1
      Top             =   1290
      Width           =   5460
      Begin VB.CheckBox chkME 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4980
         TabIndex        =   12
         Top             =   1065
         Width           =   270
      End
      Begin VB.CheckBox chkMN 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   4980
         TabIndex        =   11
         Top             =   465
         Width           =   270
      End
      Begin Sicmact.TxtBuscar txtME 
         Height          =   360
         Left            =   2355
         TabIndex        =   5
         Top             =   990
         Width           =   2505
         _ExtentX        =   3413
         _ExtentY        =   635
         Appearance      =   0
         BackColor       =   8454016
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
      Begin Sicmact.TxtBuscar txtMN 
         Height          =   360
         Left            =   2355
         TabIndex        =   2
         Top             =   390
         Width           =   2505
         _ExtentX        =   3413
         _ExtentY        =   635
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
         sTitulo         =   ""
      End
      Begin VB.Label Label2 
         Caption         =   "Moneda Extranjera"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   210
         Left            =   240
         TabIndex        =   4
         Top             =   1065
         Width           =   2295
      End
      Begin VB.Label lblMN 
         Caption         =   "Moneda Nacional"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   240
         TabIndex        =   3
         Top             =   465
         Width           =   2295
      End
   End
   Begin Sicmact.TxtBuscar txtAge 
      Height          =   360
      Left            =   60
      TabIndex        =   0
      Top             =   885
      Width           =   1050
      _ExtentX        =   3413
      _ExtentY        =   635
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
      sTitulo         =   ""
   End
   Begin VB.Label lblNomPers 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   60
      TabIndex        =   10
      Top             =   390
      Width           =   5415
   End
   Begin VB.Label lblPersCod 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   300
      Left            =   60
      TabIndex        =   7
      Top             =   60
      Width           =   5415
   End
   Begin VB.Label lblAgencia 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   1125
      TabIndex        =   6
      Top             =   915
      Width           =   4365
   End
End
Attribute VB_Name = "frmLogProvCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Ini(psPersCod As String, psPersNombre As String, psCtaCodMN As String, psCtaCodME As String)
    Me.lblPersCod.Caption = psPersCod
    Me.lblNomPers.Caption = psPersNombre
    
    Me.Show 1
End Sub

Private Sub cmdAceptar_Click()
    Dim oProv As DLogProveedor
    Set oProv = New DLogProveedor
    
    If MsgBox("Desea Grabar las cuentas ? , solo se grabaran las cuentas marcadas con check, solo se puede grabar una cuenta por moneda.", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oProv.SetProvCtas Me.lblPersCod.Caption, IIf(Me.chkMN.value = 1, Me.txtMN.Text, ""), IIf(Me.chkME.value = 1, Me.txtME.Text, ""), GetMovNro(gsCodUser, gsCodAge)
    
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oCon As DConstantes
    Set oCon = New DConstantes
    

    Me.txtAge.rs = oCon.GetAgencias(, , True)
    
End Sub

Private Sub txtAge_EmiteDatos()
    Dim oProv As DLogProveedor
    Set oProv = New DLogProveedor
    
    Me.lblAgencia.Caption = Me.txtAge.psDescripcion
    
    If txtAge <> "" Then
        Me.txtMN.rs = oProv.GetProvCtas(Me.lblPersCod.Caption, Me.txtAge.Text, gMonedaNacional, gbBitCentral)
        Me.txtME.rs = oProv.GetProvCtas(Me.lblPersCod.Caption, Me.txtAge.Text, gMonedaExtranjera, gbBitCentral)
    End If
    
End Sub

Private Sub txtME_GotFocus()
    If txtAge.Text = "" Then
        MsgBox "Debe elegir la agencia donde buscar las cuentas.", vbInformation, "Aviso"
        Me.txtAge.SetFocus
    End If
End Sub

Private Sub txtMN_GotFocus()
    If txtAge.Text = "" Then
        MsgBox "Debe elegir la agencia donde buscar las cuentas.", vbInformation, "Aviso"
        Me.txtAge.SetFocus
    End If
End Sub

