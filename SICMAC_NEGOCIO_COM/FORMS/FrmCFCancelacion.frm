VERSION 5.00
Begin VB.Form FrmCFCancelacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Carta Fianza -Cancelacion de CF Emitida"
   ClientHeight    =   5640
   ClientLeft      =   2265
   ClientTop       =   1845
   ClientWidth     =   7665
   Icon            =   "FrmCFCancelacion.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5640
   ScaleWidth      =   7665
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   180
      TabIndex        =   28
      Top             =   4860
      Width           =   7335
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   31
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5940
         TabIndex        =   30
         Top             =   195
         Width           =   1155
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1680
         TabIndex        =   29
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Afianzado"
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
      Height          =   645
      Left            =   120
      TabIndex        =   19
      Top             =   660
      Width           =   7425
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Afianzado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   705
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   21
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2340
         TabIndex        =   20
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   4920
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Acreedor"
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
      Height          =   675
      Left            =   120
      TabIndex        =   15
      Top             =   1320
      Width           =   7410
      Begin VB.Label lblNomAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2340
         TabIndex        =   18
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   4950
      End
      Begin VB.Label lblCodAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1080
         TabIndex        =   17
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1245
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Acreedor "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   16
         Top             =   240
         Width           =   690
      End
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Cancelacion"
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
      Height          =   1425
      Left            =   120
      TabIndex        =   12
      Top             =   3420
      Width           =   7380
      Begin VB.TextBox TxtComenta 
         Height          =   570
         Left            =   1095
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   690
         Width           =   5730
      End
      Begin VB.ComboBox CboMotivoR 
         Height          =   315
         Left            =   1095
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   285
         Width           =   5760
      End
      Begin VB.Label Label10 
         Caption         =   "Comentario"
         Height          =   255
         Left            =   180
         TabIndex        =   14
         Top             =   720
         Width           =   900
      End
      Begin VB.Label Label4 
         Caption         =   "Motivo"
         Height          =   255
         Left            =   180
         TabIndex        =   13
         Top             =   300
         Width           =   615
      End
   End
   Begin VB.CommandButton CmdExaminar 
      Caption         =   "E&xaminar..."
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
      Left            =   6240
      TabIndex        =   1
      Top             =   180
      Width           =   1230
   End
   Begin VB.Frame FraCredito 
      Caption         =   "Carta Fianza"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   7395
      Begin VB.Label Label9 
         Caption         =   "Analista"
         Height          =   255
         Left            =   180
         TabIndex        =   27
         Top             =   900
         Width           =   720
      End
      Begin VB.Label lblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   26
         Top             =   840
         Width           =   3420
      End
      Begin VB.Label lblFecVencCF 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5700
         TabIndex        =   24
         Top             =   840
         Width           =   1590
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento"
         Height          =   195
         Left            =   4740
         TabIndex        =   23
         Top             =   900
         Width           =   870
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   180
         TabIndex        =   11
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Monto"
         Height          =   255
         Left            =   4740
         TabIndex        =   10
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   4740
         TabIndex        =   9
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label8 
         Caption         =   "Modalidad"
         Height          =   255
         Left            =   180
         TabIndex        =   8
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lblTipoCF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   7
         Top             =   240
         Width           =   3420
      End
      Begin VB.Label lblMontoSol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5700
         TabIndex        =   6
         Top             =   540
         Width           =   1590
      End
      Begin VB.Label lblModalidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   5
         Top             =   540
         Width           =   3420
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5700
         TabIndex        =   4
         Top             =   240
         Width           =   1590
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   390
      Left            =   180
      TabIndex        =   32
      Top             =   120
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   688
      Texto           =   "Cta Fianza"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Label lblEstado 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   4080
      TabIndex        =   25
      Top             =   180
      Width           =   1605
   End
End
Attribute VB_Name = "FrmCFCancelacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*  APLICACION : Carta Fianza
'*  ARCHIVO : frmCFDevolucion
'*  CREACION: 06/09/2007     - CAPI
'*************************************************************************
'*  RESUMEN: CANCELACION CARTA FIANZA EMITIDA
'***************************************************************************

Option Explicit
Dim vCodCta As String

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(ActXCodCta.NroCuenta)) > 0 Then
            Call CargaDatosR(ActXCodCta.NroCuenta)
        Else
            Call LimpiarControles
        End If
    End If
End Sub

Private Sub CboMotivoR_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TxtComenta.SetFocus
    End If
End Sub

Private Sub cmdCancelar_Click()
    LimpiarControles
    ActXCodCta.SetFocus
End Sub

Private Sub cmdexaminar_Click()
Dim lsCta As String
    lsCta = frmCFPersEstado.Inicio(Array(gColocEstVigNorm), "Cancelacion de Carta Fianza Emitida", Array(gColCFComercial, gColCFPYME))
    If Len(Trim(lsCta)) > 0 Then
        ActXCodCta.NroCuenta = lsCta
        Call CargaDatosR(lsCta)
    Else
        Call LimpiarControles
    End If
End Sub

'PROCEDIMIENTO QUE CARGA LOS DATOS QUE SE REQUIEREN PARA EL FORMULARIO
Sub CargaDatosR(ByVal psCodCta As String)

Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim R As New ADODB.Recordset
Dim loConstante As COMDConstantes.DCOMConstantes
Dim lbTienePermiso As Boolean


'On Error GoTo ErrorCargaDat
    
'ActXCodCta.Enabled = False 'FRHU20131120

    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set R = oCF.RecuperaCartaFianzaDevolucion(psCodCta)
    Set oCF = Nothing
    'FRHU20131120
    If R Is Nothing Then
        MsgBox "Esta Carta Fianza ya se encuentra en Estado Cancelado"
        Exit Sub
    End If
    'END FRHU20131120
    If Not R.BOF And Not R.EOF Then
        lblCodigo.Caption = R!cPersCod
        lblNombre.Caption = PstaNombre(R!cPersNombre)
    
        lblCodAcreedor.Caption = R!cPersAcreedor
        lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
    
        
        If Mid(Trim(psCodCta), 6, 1) = "1" Then
            lblTipoCF = "COMERCIALES "
        ElseIf Mid(Trim(psCodCta), 6, 1) = "2" Then
            lblTipoCF = "MICROEMPRESA "
        End If
        
        If Mid(Trim(psCodCta), 9, 1) = "1" Then
            lblmoneda = "Soles"
        ElseIf Mid(Trim(psCodCta), 9, 1) = "2" Then
            lblmoneda = "Dolares"
        End If
        lblanalista.Caption = PstaNombre(IIf(IsNull(R!cAnalista), "", R!cAnalista))
        lblMontoSol.Caption = IIf(IsNull(R!nMontoSol), "", Format(R!nMontoSol, "#0.00"))
        lblFecVencCF.Caption = IIf(IsNull(R!dVencSol), "", Format(R!dVencSol, "dd/mm/yyyy"))
        'lblFinalidad.Caption = IIf(IsNull(R!cFinalidad), "", R!cFinalidad)

        Set loConstante = New COMDConstantes.DCOMConstantes
            'lblModalidad.Caption = loConstante.DameDescripcionConstante(gColCFModalidad, R!nModalidad)'comento JOEP20181222 CP
            'JOEP20181222 CP
            If R!nModalidad = 13 Then
                lblModalidad.Caption = R!OtrsModalidades
            Else
                lblModalidad.Caption = loConstante.DameDescripcionConstante(gColCFModalidad, R!nModalidad)
            End If
            'JOEP20181222 CP
        Set loConstante = Nothing
        
        LblEstado.Caption = fgEstadoColCFDesc(R!nPrdEstado)
        'lblModalidad = DatoTablaCodigo("D1", IIf(IsNull(Trim(reg!cModalidad)), "", reg!cModalidad))
        
        FraDatos.Enabled = True
        cmdGrabar.Enabled = True
        CboMotivoR.SetFocus
    'FRHU20131120
    Else
        MsgBox "Esta Carta Fianza ya se encuentra en Estado Cancelado"
        Exit Sub
    'END FRHU20131120
    End If
Exit Sub

ErrorCargaDat:
    MsgBox "Error Nº [" & str(Err.Number) & "] " & Err.Description, vbCritical, "Error del Sistema"
    Exit Sub
End Sub



Private Sub CmdGrabar_Click()
Dim loNCartaFianza As COMNCartaFianza.NCOMCartaFianza 'NCartaFianza

Dim loContFunct As COMNContabilidad.NCOMContFunciones 'NContFunciones
Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnDevuelta As String
Dim lsComenta As String
Dim lnMonto As Double

vCodCta = ActXCodCta.NroCuenta
lnDevuelta = CInt(Trim(Right(CboMotivoR.Text, 2)))
'lnDevuelta = Trim(CboMotivoR.Text)

lsComenta = Replace(Trim(TxtComenta), "'", " ", , , vbTextCompare)
lnMonto = CDbl(Me.lblMontoSol.Caption)

If MsgBox("Desea Grabar Cancelacion de Carta Fianza Emitida", vbInformation + vbYesNo, "Sugerencia de Analista") = vbYes Then
           
    Dim OCon As New COMConecta.DCOMConecta
    Dim sql As String
    Dim cFecCan As String
    
    cFecCan = Mid(gdFecSis, 7, 4) + Mid(gdFecSis, 4, 2) + Mid(gdFecSis, 1, 2)
    OCon.AbreConexion
    sql = " exec  Col_CFCancelacion  '" & ActXCodCta.NroCuenta & "' , '" & cFecCan & "' ,'" & gsCodUser & "' , '" & gsCodAge & "'," & lnDevuelta
    OCon.Ejecutar (sql)
    OCon.CierraConexion
    Set OCon = Nothing
               
    cmdGrabar.Enabled = False
    LimpiarControles
End If

End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    LimpiarControles
    Call CargaComboConstante(gColocMotivRechazo, CboMotivoR)
    CboMotivoR.ListIndex = 0
End Sub

Sub LimpiarControles()
   ActXCodCta.Enabled = True
   ActXCodCta.NroCuenta = fgIniciaAxCuentaCF
   lblCodigo.Caption = ""
   lblNombre.Caption = ""
   lblCodAcreedor.Caption = ""
   lblNomAcreedor.Caption = ""
   lblTipoCF.Caption = ""
   lblmoneda.Caption = ""
   lblMontoSol.Caption = ""
   lblModalidad.Caption = ""
   lblanalista.Caption = ""
   lblFecVencCF.Caption = ""
   LblEstado.Caption = ""
   TxtComenta = ""
   CboMotivoR.ListIndex = -1
   FraDatos.Enabled = False
   cmdGrabar.Enabled = False
End Sub

Private Sub TxtComenta_KeyPress(KeyAscii As Integer)
     KeyAscii = fgIntfMayusculas(KeyAscii)
     If KeyAscii = 13 Then
        cmdGrabar.SetFocus
     End If
End Sub
