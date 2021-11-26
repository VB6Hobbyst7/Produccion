VERSION 5.00
Begin VB.Form frmColPReimpresionComprobanteAdj 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reimpresión Comprobante de Pago"
   ClientHeight    =   3510
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6375
   Icon            =   "frmColPReimpresionComprobanteAdj.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3510
   ScaleWidth      =   6375
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   310
      Left            =   5040
      TabIndex        =   10
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   310
      Left            =   3840
      TabIndex        =   12
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Timer tmrAlerta 
      Left            =   3000
      Top             =   240
   End
   Begin VB.CommandButton cmdImprmir 
      Caption         =   "Imprimir"
      Height          =   310
      Left            =   2640
      TabIndex        =   11
      Top             =   3120
      Width           =   1215
   End
   Begin VB.Frame frTipoComp 
      Caption         =   " Comprobantes de Pago "
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   6
      Top             =   1920
      Width           =   6135
      Begin VB.CheckBox chkUsuario 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   480
         Width           =   2775
      End
      Begin VB.CheckBox chkContabilidad 
         Caption         =   "Contabilidad"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   720
         Width           =   2775
      End
      Begin VB.CheckBox chkEmisor 
         Caption         =   "Emisor"
         Height          =   255
         Left            =   360
         TabIndex        =   7
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos del Cliente "
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   6135
      Begin VB.Label lblDireccion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   960
         TabIndex        =   5
         Top             =   600
         Width           =   5055
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   300
         Left            =   960
         TabIndex        =   4
         Top             =   315
         Width           =   5055
      End
      Begin VB.Label Label2 
         Caption         =   "Domicilio:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   680
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
   End
   Begin SICMACT.ActXCodCta_New ActXCodCta 
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1508
      Texto           =   " Crédito "
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
      CMAC            =   "109"
   End
End
Attribute VB_Name = "frmColPReimpresionComprobanteAdj"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmColPReimpresionComprobanteAdj
'** Descripción : Formulario que reimprime comprobante de pago de creditos adjudicados.
'** Creación    : RECO, 20160229 - ERS040-2015
'**********************************************************************************************

Option Explicit

Dim rsDatos  As New ADODB.Recordset
Dim bChkEmisor As Boolean
Dim bChkUsuario As Boolean
Dim bChkContabilidad As Boolean

'Dim tmrAlerta As New Timer
Dim nEfecCont As Integer
Dim nEfec As Integer
Dim nEfecControl As Integer

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Dim oColP As New COMNColoCPig.NCOMColPContrato
        Set rsDatos = oColP.PignoReimpresionCompAdj(ActXCodCta.NroCuenta)
        
        If Not (rsDatos.EOF And rsDatos.BOF) Then
            lblNombre.Caption = rsDatos!cPersNombre
            lblDireccion.Caption = rsDatos!cPersDireccDomicilio
        Else
            tmrAlerta.Enabled = True
            nEfecControl = 1
            MsgBox "No se encontró información de la cuenta.", vbInformation, "Alerta"
            Call LimpiarFomulario
        End If
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiarFomulario
End Sub

Private Sub cmdImprmir_Click()
    Dim sMsj As String
    If ActXCodCta.NroCuenta <> "" Then
        sMsj = ValidaDatos
        tmrAlerta.Interval = 500
        If sMsj = "" Then
            Screen.MousePointer = 11
            Call ImprimeComprobanteAdj(True, ActXCodCta.NroCuenta, , , , , CantidadComprobante, chkEmisor.value, chkUsuario.value, chkContabilidad.value)
            Screen.MousePointer = 0
        Else
            tmrAlerta.Enabled = True
            nEfecControl = 2
            MsgBox sMsj, vbInformation, "Alerta"
        End If
    End If
End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Function CantidadComprobante() As Integer
    Dim nIndice As Integer
    CantidadComprobante = 0
    If chkEmisor.value = 1 Then
        CantidadComprobante = 1 + CantidadComprobante
    End If
    If chkUsuario.value = 1 Then
        CantidadComprobante = 1 + CantidadComprobante
    End If
    If chkContabilidad.value = 1 Then
        CantidadComprobante = 1 + CantidadComprobante
    End If
End Function
Private Function ValidaDatos() As String
    ValidaDatos = ""
    If chkEmisor.value = 0 And chkUsuario.value = 0 And chkContabilidad.value = 0 Then
        ValidaDatos = "Debe seleccionar un tipo de comprobante."
        Exit Function
    End If
End Function
Private Sub LimpiarFomulario()
    ActXCodCta.NroCuenta = ""
    lblNombre.Caption = ""
    lblDireccion.Caption = ""
    chkEmisor.value = 0
    chkUsuario.value = 0
    chkContabilidad.value = 0
    ActXCodCta.CMAC = "109"
    ActXCodCta.Age = gsCodAge
    ActXCodCta.Prod = "705"
End Sub

Private Sub Form_Load()
    Call LimpiarFomulario
End Sub

Private Sub tmrAlerta_Timer()
    If nEfec = 0 Then
        If nEfecControl = 2 Then
            frTipoComp.ForeColor = &H800000
        Else
            ActXCodCta.ForeColor = &H800000
        End If
        nEfec = 1
    Else
        If nEfecControl = 2 Then
            frTipoComp.ForeColor = &HFF0000
        Else
            ActXCodCta.ForeColor = &HFF0000
        End If
        nEfec = 0
    End If
    nEfecCont = nEfecCont + 1
    If nEfecCont = 4 Then
        tmrAlerta.Enabled = False
        nEfecCont = 0
    End If
End Sub
