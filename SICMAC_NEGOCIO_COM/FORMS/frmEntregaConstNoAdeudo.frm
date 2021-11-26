VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmEntregaConstNoAdeudo 
   Caption         =   "Entrega de Constancia de No Adeudo"
   ClientHeight    =   1905
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7380
   Icon            =   "frmEntregaConstNoAdeudo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1905
   ScaleWidth      =   7380
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   3000
         TabIndex        =   1
         Top             =   840
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtFEntrega 
         Height          =   300
         Left            =   3360
         TabIndex        =   2
         Top             =   240
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "F. de Entrega:"
         Height          =   255
         Left            =   1680
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmEntregaConstNoAdeudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim sPersCod As String
Dim iConstNoAdeudo As Integer

Public Sub Inicio(ByVal iConstanciaNoAdeudo As Integer, ByVal sPersCodigo As String)
    sPersCod = sPersCodigo
    iConstNoAdeudo = iConstanciaNoAdeudo
End Sub

Private Sub cmdAceptar_Click()
    Dim objCOMNCredito As COMNCredito.NCOMCredito
    Set objCOMNCredito = New COMNCredito.NCOMCredito
    
    If txtFEntrega.Text <> "__/__/____" Then
        objCOMNCredito.ActualizarEstadoConstNoAdeudo iConstNoAdeudo, sPersCod, txtFEntrega.Text
        Unload Me
        frmListarConstNoAdeudo.BuscarDatos
    Else
        MsgBox "Debe Ingresar la Fecha de Entrega", vbCritical, "Aviso"
    End If
    
End Sub

Public Sub CargarDatosModificar(ByVal ConstNoAdeudoId As Integer)
'    Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
'    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
'    Dim rs As ADODB.Recordset
'    Set rs = objCOMNAuditoria.ObtenerRevisionXId(gRevisionId)
'    Dim i, J, K As Integer
'    txtCodigo.Text = rs("cPersCod")
'    txtCodCta.Text = rs("vCodCta")
'    mskPeriodo1Del = rs("cFCierre")
'    txtTCambio.Text = rs("mTCambio")
'    txtFRegistro.Text = rs("cFRegistro")
'    txtCAnalista.Text = rs("vCAnalista")
'    txtDeudor.Text = rs("cPersNombre")
'    txtDireccion.Text = rs("cPersDireccDomicilio")
'    txtGiro.Text = rs("vGiro")
'    txtFSDCmac.Text = rs("cFSDCMAC")
'    txtMontoCMAC.Text = Format(rs("mMontoCMAC"), "#,##0.00")
'    cboTMonedaCMAC.Text = rs("cTMonedaCMAC")
'    txtFSDSF.Text = rs("cFSDSF")
'    txtMontoSF.Text = Format(rs("mMontoSF"), "#,##0.00")
'    cboTMonedaSF.Text = rs("cTMonedaSF")
'    txtPNorm.Text = rs("cPNorm")
'    txtPCPP.Text = rs("cPCPP")
'    txtPDefic.Text = rs("cPDefic")
'    txtPDud.Text = rs("cPDud")
'    txtPPerd.Text = rs("cPPerd")
'        If rs("vCalificacionCMAC") <> "" Then
'            For i = 0 To 4
'                If i = CInt(rs("vCalificacionCMAC")) Then
'                    txtN(i).Text = "x"
'                End If
'            Next i
'        End If
'
'        If rs("vCalificacionSF") <> "" Then
'            For J = 0 To 4
'                If J = CInt(rs("vCalificacionSF")) Then
'                    txtSF(J).Text = "x"
'                End If
'            Next J
'        End If
'
'        If rs("vCalificacionOCI") <> "" Then
'            For K = 0 To 4
'                If K = CInt(rs("vCalificacionOCI")) Then
'                    txtOCI(K).Text = "x"
'                End If
'            Next K
'        End If
'    txtSituacion.Text = rs("vSituacion")
'    txtDesarrollo.Text = rs("vDesarrollo")
'    txtGarantia.Text = rs("vGarantia")
'    txtInformacion.Text = rs("vInformacion")
'    txtEvaluacion.Text = rs("vEvaluacion")
'    txtComentario.Text = rs("vComentario")
'    txtConclusion.Text = rs("vConclusion")
End Sub

Private Sub txtFEntrega_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       cmdAceptar.SetFocus
    End If
End Sub

Private Sub txtFEntrega_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFEntrega.Text)
        If Not Trim(sCad) = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            Exit Sub
        End If
        If CDate(txtFEntrega.Text) > gdFecSis Then
            MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
            txtFEntrega.SetFocus
            Exit Sub
        End If
End Sub
