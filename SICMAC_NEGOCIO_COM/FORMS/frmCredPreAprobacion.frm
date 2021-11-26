VERSION 5.00
Begin VB.Form frmCredPreAprobacion 
   Caption         =   "Aprobacion Por Niveles de Credito"
   ClientHeight    =   4395
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6765
   Icon            =   "frmCredPreAprobacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   6765
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   8
      Top             =   3480
      Width           =   6495
      Begin VB.CommandButton cmdRechazar 
         Caption         =   "&Rechazar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   10
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aprobar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1920
         TabIndex        =   9
         Top             =   240
         Width           =   1215
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   6495
      Begin VB.TextBox TxtMontoApr 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   285
         Left            =   1080
         TabIndex        =   11
         Top             =   1200
         Width           =   1170
      End
      Begin VB.CommandButton cmdNvaConsulta 
         Caption         =   "Nueva Consulta"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4800
         TabIndex        =   7
         Top             =   240
         Width           =   1575
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   240
         TabIndex        =   4
         Top             =   240
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label3 
         Caption         =   "Monto:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label lblTitular 
         Height          =   255
         Left            =   1080
         TabIndex        =   6
         Top             =   840
         Width           =   5055
      End
      Begin VB.Label Label1 
         Caption         =   "Titular:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   840
         Width           =   615
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   6495
      Begin VB.TextBox txtComentario 
         Height          =   1095
         Left            =   1200
         TabIndex        =   1
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label2 
         Caption         =   "Comentario:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCredPreAprobacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CargaDatos (ActxCta.NroCuenta)
    End If
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
    Dim objNivApr As COMDCredito.DCOMNivelAprobacion
    Set objNivApr = New COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    Set rs = objNivApr.CargarDatosPreApr(psCtaCod, gsCodCargo)
    
    If rs.RecordCount <> "0" Then
        If rs!iTipoAprob = "0" Then
            lblTitular = PstaNombre(Trim(rs!cPersNombre))
            txtComentario.Text = rs!vComentario
            cmdAceptar.Enabled = True
            cmdRechazar.Enabled = True
            TxtMontoApr.Text = Format(rs!nMonto, "#0.00")
            TxtMontoApr.Enabled = True
        Else
            MsgBox "El Credito Ya fue Aprobado", vbInformation, "Aviso"
            cmdAceptar.Enabled = False
            cmdRechazar.Enabled = False
            LimpiaControles Me
            lblTitular.Caption = ""
            ActxCta.CMAC = gsCodCMAC
            ActxCta.Age = gsCodAge
            TxtMontoApr.Enabled = False
        End If
    Else
        MsgBox "No se Encontro el Credito", vbCritical, "Aviso"
        LimpiaControles Me
        lblTitular.Caption = ""
        ActxCta.CMAC = gsCodCMAC
        ActxCta.Age = gsCodAge
        cmdAceptar.Enabled = False
        cmdRechazar.Enabled = False
        TxtMontoApr.Enabled = False
    End If

    Set objNivApr = Nothing
    Set rs = Nothing
End Function

Private Sub cmdAceptar_Click()
    Dim objNivApr As COMDCredito.DCOMNivelAprobacion
    Set objNivApr = New COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim i As Integer
    
    If MsgBox("Se Va A Grabar los Datos, Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        objNivApr.ModificarEstadoPreApr ActxCta.NroCuenta, gsCodCargo, txtComentario.Text, 1, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        Set rs = objNivApr.CargarEncargaturaNivApr(gsCodUser)
        If rs.RecordCount <> "0" Then
            For i = 0 To rs.RecordCount - 1
                objNivApr.ModificarEstadoPreApr ActxCta.NroCuenta, Mid(Trim(rs!cCargoCod), 4, 6), txtComentario.Text, 1, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Next i
        End If
        
        MsgBox "Los Datos Se Registraron Correctamente", vbInformation, "AVISO"
        LimpiaControles Me
        lblTitular.Caption = ""
        ActxCta.CMAC = gsCodCMAC
        ActxCta.Age = gsCodAge
        
        Set objNivApr = Nothing
        Set rs = Nothing
        
    End If
    
End Sub

Private Sub cmdNvaConsulta_Click()
    LimpiaControles Me
    lblTitular.Caption = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    TxtMontoApr.Enabled = False
End Sub

Private Sub cmdRechazar_Click()
    Dim objNivApr As COMDCredito.DCOMNivelAprobacion
    Set objNivApr = New COMDCredito.DCOMNivelAprobacion
        If MsgBox("Se Va A Grabar los Datos, Desea Continuar ?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            objNivApr.ModificarEstadoPreApr ActxCta.NroCuenta, gsCodCargo, txtComentario.Text, 2, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            MsgBox "Los Datos Se Registraron Correctamente", vbInformation, "AVISO"
            LimpiaControles Me
            lblTitular.Caption = ""
            ActxCta.CMAC = gsCodCMAC
            ActxCta.Age = gsCodAge
        End If
End Sub

Private Sub Form_Load()
    CentraForm Me
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
End Sub

Private Sub TxtMontoApr_Change()
    If Mid(ActxCta.NroCuenta, 9, 1) = "1" Then
        TxtMontoApr.ForeColor = vbBlue
    Else
        TxtMontoApr.ForeColor = &H289556
    End If
End Sub
