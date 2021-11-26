VERSION 5.00
Begin VB.Form frmCredRetiroConvenio 
   Caption         =   "Retirar crédito de convenio"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8340
   Icon            =   "frmCredRetiroConvenio.frx":0000
   LinkTopic       =   "Form3"
   ScaleHeight     =   3075
   ScaleWidth      =   8340
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRetirar 
      Caption         =   "Retirar"
      Height          =   315
      Left            =   4680
      TabIndex        =   16
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   315
      Left            =   5880
      TabIndex        =   15
      Top             =   2640
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   315
      Left            =   7080
      TabIndex        =   14
      Top             =   2640
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Crétido"
      Height          =   1935
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   8055
      Begin VB.Label lblEstado 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   5880
         TabIndex        =   13
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label11 
         Caption         =   "Estado Actual:"
         Height          =   255
         Left            =   4800
         TabIndex        =   12
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblInstitucion 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         Top             =   1440
         Width           =   6615
      End
      Begin VB.Label lblTpoCred 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   10
         Top             =   720
         Width           =   2895
      End
      Begin VB.Label lblTpoProd 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   9
         Top             =   1080
         Width           =   2895
      End
      Begin VB.Label lblDOI 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   6360
         TabIndex        =   8
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label6 
         Caption         =   "DOI:"
         Height          =   255
         Left            =   5880
         TabIndex        =   7
         Top             =   360
         Width           =   495
      End
      Begin VB.Label lblCliente 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1320
         TabIndex        =   6
         Top             =   360
         Width           =   4215
      End
      Begin VB.Label Label4 
         Caption         =   "Institución:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   1440
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo Producto:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Crédito:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   615
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   661
      Texto           =   "Credito"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
End
Attribute VB_Name = "frmCredRetiroConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public fsCodInstitucion As String
Public fsCodModular As String
Public fsCargo As String
Public fsCARBEN As String
Public fsT_Plani As String

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    Dim loCred As New COMDCredito.DCOMCreditos
    Dim loDR  As New ADODB.Recordset
     If KeyAscii = 13 Then
        Set loCred = New COMDCredito.DCOMCreditos
        ActXCodCta.Enabled = False
        Set loDR = loCred.ObtieneDatosCredPersona(ActXCodCta.NroCuenta)
        
        If Not (loDR.EOF And loDR.BOF) Then
                lblCliente.Caption = loDR!cPersNombre
                lblDOI.Caption = loDR!cPersIDnro
                lblTpoCred.Caption = loDR!cTpoCred
                lblTpoProd.Caption = loDR!CTpoProd
                lblEstado.Caption = loDR!cEstado
                lblInstitucion.Caption = loDR!cPersNombConv
                fsCodInstitucion = loDR!cPersCodConv
        Else
            MsgBox "No se encontro datos de crédito, verifique el número", vbExclamation, "Aviso"
            Call Limpiar
        End If
     End If
End Sub

Public Sub Limpiar()
    fsCodInstitucion = ""
    fsCodModular = ""
    fsCargo = ""
    fsCARBEN = ""
    fsT_Plani = ""
    ActXCodCta.NroCuenta = ""
    lblCliente.Caption = ""
    lblDOI.Caption = ""
    lblTpoCred.Caption = ""
    lblTpoProd.Caption = ""
    lblEstado.Caption = ""
    lblInstitucion = ""
    ActXCodCta.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
    Call Limpiar
End Sub

Private Sub cmdRetirar_Click()
    Dim loCred As New COMDCredito.DCOMCreditos
    Set loCred = New COMDCredito.DCOMCreditos
    
    Dim loCredHis As COMDCredito.DCOMCredito
    Set loCredHis = New COMDCredito.DCOMCredito
    
    Dim oDR As New ADODB.Recordset
    Set oDR = New ADODB.Recordset
    
    Set oDR = loCred.ObtieneCtaConvenioPostDesembolso(ActXCodCta.NroCuenta)
    
    If Not (oDR.EOF And oDR.BOF) Then
        loCred.EliminaConvenioCredPosDesembolso (ActXCodCta.NroCuenta)
        'JAME *****'
        loCredHis.RegistroHistorialConvenio ActXCodCta.NroCuenta, gsCodUser, Format(gdFecSis & " " & GetHoraServer, "yyyy/MM/dd hh:mm:ss"), fsCodInstitucion, " ", 3
        'JAME FIN'
        MsgBox "La eliminación culmino con exito", vbInformation, "Aviso"
    Else
        MsgBox "La cuenta no cumple con las caracteristica necesario para realizar esta operación", vbCritical, "Aviso"
    End If
    Call Limpiar
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub
