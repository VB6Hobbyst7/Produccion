VERSION 5.00
Begin VB.Form frmCFEditarModEm 
   Caption         =   "Cartas Fianza -Edición de Modalidad"
   ClientHeight    =   7725
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7575
   Icon            =   "frmCFEditarModEm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   7575
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraActualizarModalidad 
      Caption         =   "Actualizar Modalidad"
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
      Height          =   2595
      Left            =   120
      TabIndex        =   31
      Top             =   4320
      Width           =   7410
      Begin VB.Frame frOtrsMod 
         Height          =   615
         Left            =   120
         TabIndex        =   38
         Top             =   720
         Width           =   7215
         Begin VB.TextBox txtOtrsModa 
            Height          =   300
            Left            =   1560
            TabIndex        =   40
            Top             =   220
            Width           =   5415
         End
         Begin VB.Label Label1 
            Caption         =   "Otras Modalidades"
            Height          =   300
            Left            =   120
            TabIndex        =   39
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.Frame frGlosa 
         Height          =   1215
         Left            =   120
         TabIndex        =   35
         Top             =   1320
         Width           =   7215
         Begin VB.TextBox txtGlosa 
            Height          =   855
            Left            =   1560
            MaxLength       =   700
            MultiLine       =   -1  'True
            TabIndex        =   36
            Top             =   240
            Width           =   5430
         End
         Begin VB.Label lblGlosa 
            AutoSize        =   -1  'True
            Caption         =   "Glosa:"
            ForeColor       =   &H80000008&
            Height          =   195
            Left            =   120
            TabIndex        =   37
            Top             =   240
            Width           =   450
         End
      End
      Begin VB.ComboBox cboModalidad 
         Height          =   315
         Left            =   1680
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   360
         Width           =   3120
      End
      Begin VB.Label lblNuevaMod 
         AutoSize        =   -1  'True
         Caption         =   "Nueva Modalidad:"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   240
         TabIndex        =   32
         Top             =   360
         Width           =   1305
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Aval"
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
      TabIndex        =   27
      Top             =   1880
      Width           =   7410
      Begin VB.Label lblNomAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   30
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   5010
      End
      Begin VB.Label lblCodAvalado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   29
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Aval"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   28
         Top             =   240
         Width           =   315
      End
   End
   Begin VB.Frame FraCredito 
      Caption         =   "Carta Fianza"
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
      Height          =   1695
      Left            =   120
      TabIndex        =   11
      Top             =   2520
      Width           =   7395
      Begin VB.Label lblMontoApr 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         TabIndex        =   34
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   5760
         TabIndex        =   22
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lblModalidad 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   525
         Left            =   1020
         TabIndex        =   21
         Top             =   540
         Width           =   3420
      End
      Begin VB.Label lblTipoCF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   20
         Top             =   240
         Width           =   3420
      End
      Begin VB.Label Label8 
         Caption         =   "Modalidad"
         Height          =   255
         Left            =   180
         TabIndex        =   19
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   4680
         TabIndex        =   18
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Monto"
         Height          =   255
         Left            =   4680
         TabIndex        =   17
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   180
         TabIndex        =   16
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Vencimiento"
         Height          =   195
         Left            =   4680
         TabIndex        =   15
         Top             =   960
         Width           =   870
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
         Left            =   5760
         TabIndex        =   14
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label lblAnalista 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1020
         TabIndex        =   13
         Top             =   1200
         Width           =   3420
      End
      Begin VB.Label Label9 
         Caption         =   "Analista"
         Height          =   255
         Left            =   180
         TabIndex        =   12
         Top             =   1260
         Width           =   720
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
      TabIndex        =   7
      Top             =   1200
      Width           =   7410
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Acreedor "
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   690
      End
      Begin VB.Label lblCodAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   9
         Tag             =   "txtcodigo"
         Top             =   240
         Width           =   1305
      End
      Begin VB.Label lblNomAcreedor 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   8
         Tag             =   "txtnombre"
         Top             =   240
         Width           =   5010
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
      TabIndex        =   3
      Top             =   540
      Width           =   7425
      Begin VB.Label lblNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   2280
         TabIndex        =   6
         Tag             =   "txtnombre"
         Top             =   210
         Width           =   5025
      End
      Begin VB.Label lblCodigo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   960
         TabIndex        =   5
         Tag             =   "txtcodigo"
         Top             =   210
         Width           =   1305
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Afianzado"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Frame Frame4 
      Height          =   675
      Left            =   120
      TabIndex        =   0
      Top             =   6960
      Width           =   7380
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   1680
         TabIndex        =   24
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5940
         TabIndex        =   2
         Top             =   195
         Width           =   1155
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "&Modificar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   480
         TabIndex        =   1
         Top             =   180
         Width           =   1155
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   390
      Left            =   120
      TabIndex        =   23
      Top             =   0
      Width           =   3645
      _ExtentX        =   6429
      _ExtentY        =   688
      Texto           =   "Cta Fianza"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.Label lblPoliza 
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
      Left            =   6000
      TabIndex        =   26
      Top             =   240
      Width           =   1335
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Nº Folio :"
      Height          =   195
      Index           =   1
      Left            =   5280
      TabIndex        =   25
      Top             =   240
      Width           =   645
   End
End
Attribute VB_Name = "frmCFEditarModEm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'*  APLICACION : Carta Fianza
'*  ARCHIVO : frmCFEditarModEm
'*  CREACION: 11/03/2013     - WIOR
'*************************************************************************
'*  RESUMEN: EDITAR MODALIDAD DE CARTA FIANZA EMITIDA
'***************************************************************************
Option Explicit
Dim objPista As COMManejador.Pista

'JOEP20181226 CP
Private Sub cboModalidad_Click()
If Trim(Right(cboModalidad.Text, 9)) = "13" Then
    frmCFEditarModEm.Height = 8250
    fraActualizarModalidad.Height = 2595
    Frame4.top = 6960
    frOtrsMod.Visible = True
    frOtrsMod.BorderStyle = 0
    frOtrsMod.top = 720
    frGlosa.top = 1320
    frGlosa.BorderStyle = 0
Else
    frOtrsMod.Visible = False
    txtOtrsModa.Text = ""
    frmCFEditarModEm.Height = 7545
    Frame4.top = 6240
    frGlosa.top = 650
    frGlosa.BorderStyle = 0
    fraActualizarModalidad.top = 4320
    fraActualizarModalidad.Height = 1890
End If
End Sub
'JOEP20181226 CP

Private Sub cmdCancelar_Click()
    LimpiarControles
    ActXCodCta.SetFocus
'JEOP20181226 CP
    frOtrsMod.Visible = False
    txtOtrsModa.Text = ""
    frmCFEditarModEm.Height = 7545
    Frame4.top = 6240
    frGlosa.top = 650
    frGlosa.BorderStyle = 0
    fraActualizarModalidad.top = 4320
    fraActualizarModalidad.Height = 1890
'JEOP20181226 CP
End Sub

Private Sub CargaDatos(ByVal psCodCta As String)
Dim oCF As COMDCartaFianza.DCOMCartaFianza
Dim loConstante As COMDConstantes.DCOMConstantes
Dim R As New ADODB.Recordset

    ActXCodCta.Enabled = False
    'Call CargaControles 'Comento JOEP20181226 CP
    Call CP_CargaComboxEdi(49000) 'JOEP20181226 CP
    
    Set oCF = New COMDCartaFianza.DCOMCartaFianza
    Set R = oCF.RecuperaCartaFianzaDetalle(psCodCta)
    Set oCF = Nothing

    If Not R.BOF And Not R.EOF Then
        If IIf(IsNull(R!nRenovacion), "0", R!nRenovacion) = 0 Then
            lblCodigo.Caption = R!cPersCod
            lblNombre.Caption = PstaNombre(R!cPersNombre)
        
            lblCodAcreedor.Caption = R!cPersAcreedor
            lblNomAcreedor.Caption = PstaNombre(R!cPersNomAcre)
            'JOEP20181226 CP
            If R!nModalidad = 13 Then
                lblModalidad.Caption = UCase(R!OtrsModalidades)
            Else
                lblModalidad.Caption = UCase(R!Modalidad)
            End If
            'JOEP20181226 CP
            lblCodAvalado.Caption = IIf(IsNull(R!cPersAvalado), "", R!cPersAvalado)
            If R!cAvalNombre <> "" Then
                lblNomAvalado.Caption = IIf(IsNull(PstaNombre(R!cAvalNombre)), "", PstaNombre(R!cAvalNombre))
            End If
            
            lblTipoCF = IIf(IsNull(R!cConsDescripcion), "", R!cConsDescripcion)
            
            If Mid(Trim(psCodCta), 9, 1) = "1" Then
                lblMoneda = "Soles"
            ElseIf Mid(Trim(psCodCta), 9, 1) = "2" Then
                lblMoneda = "Dolares"
            End If
            
            lblAnalista.Caption = PstaNombre(IIf(IsNull(R!cAnalista), "", R!cAnalista))
            lblMontoApr.Caption = IIf(IsNull(R!nMontoApr), "", Format(R!nMontoApr, "#0.00"))
            lblFecVencCF.Caption = IIf(IsNull(R!dVencimiento), "", Format(R!dVencimiento, "dd/mm/yyyy"))
            fraActualizarModalidad.Enabled = True
            cmdModificar.Enabled = True
            lblPoliza.Caption = IIf(IsNull(R!nPoliza), "0", R!nPoliza)
            cboModalidad.RemoveItem (IndiceListaCombo(cboModalidad, R!nModalidad))
            cboModalidad.SetFocus
        Else
            MsgBox "No se puede modificar la Modalidad de la Carta Fianza, Ya fue Renovada", vbInformation, "AVISO"
            Exit Sub
        End If
    Else
        MsgBox "No existe Carta Fianza", vbInformation, "AVISO"
        Exit Sub
    End If
Exit Sub

ErrorCargaDat:
    MsgBox "Error Nº [" & str(Err.Number) & "] " & Err.Description, vbCritical, "Error del Sistema"
    Exit Sub
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CargaDatos (ActXCodCta.NroCuenta)
    End If
End Sub

Private Sub cmdModificar_Click()
  If ValidaDatos Then
    If MsgBox("Desea Grabar los Datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Dim oCF As COMDCredito.DCOMCredActBD
        Dim lsMovNro As String
        
        Set oCF = New COMDCredito.DCOMCredActBD
        
        lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

        
        'Call oCF.dUpdateColocCartaFianza(Trim(ActXCodCta.NroCuenta), , Trim(Right(cboModalidad.Text, 5)), , , , , , , "01/01/1950") 'Comento JOEP20181226 CP
        Call oCF.dUpdateColocCartaFianza(Trim(ActXCodCta.NroCuenta), , Trim(Right(cboModalidad.Text, 5)), , , , , , , "01/01/1950", , Trim(UCase(txtOtrsModa.Text))) 'JOEP20181226 CP Trim(UCase(txtOtrsModa.Text))
        
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, "Se cambió la modalidad :" & Trim(lblModalidad.Caption) & " a " & Trim(Left(cboModalidad.Text, Len(cboModalidad.Text) - 5)) & ", Por: " & Trim(txtGlosa.Text), Trim(ActXCodCta.NroCuenta), gCodigoCuenta
        Set objPista = Nothing
        Set oCF = Nothing
        MsgBox "Datos Guardados Satisfactoriamente", vbInformation, "Aviso"
        LimpiarControles
    End If
End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    LimpiarControles
    gsOpeCod = gCredModificacionModCF
    
'JOEP20181226 CP
    frOtrsMod.Visible = False
    txtOtrsModa.Text = ""
    frmCFEditarModEm.Height = 7545
    Frame4.top = 6240
    frGlosa.top = 650
    frGlosa.BorderStyle = 0
    fraActualizarModalidad.top = 4320
    fraActualizarModalidad.Height = 1890
'JOEP20181226 CP

End Sub

'Comento JOEP20181226 CP
'Private Sub CargaControles()
'    'Carga Modalidad de Carta Fianza
'    Call CargaComboConstante(gColCFModalidad, cboModalidad)
'    Call CambiaTamañoCombo(cboModalidad, 300)
'End Sub
'Comento JOEP20181226 CP

Sub LimpiarControles()
   ActXCodCta.Enabled = True
   ActXCodCta.NroCuenta = fgIniciaAxCuentaCF
   lblCodigo.Caption = ""
   lblNombre.Caption = ""
   lblCodAcreedor.Caption = ""
   lblNomAcreedor.Caption = ""
   lblCodAvalado.Caption = ""
   lblNomAvalado.Caption = ""
   lblTipoCF.Caption = ""
   lblMoneda.Caption = ""
   lblMontoApr.Caption = ""
   lblModalidad.Caption = ""
   lblAnalista.Caption = ""
   lblFecVencCF.Caption = ""
   lblPoliza.Caption = ""
   fraActualizarModalidad.Enabled = False
   cmdModificar.Enabled = False
   cboModalidad.Clear
   txtGlosa.Text = ""
   txtOtrsModa.Text = "" 'JOEP20181226 CP
End Sub

Private Function ValidaDatos() As Boolean

If Trim(cboModalidad.Text) = "" Then
    MsgBox "Seleccione la Modalidad.", vbInformation, "Aviso"
    cboModalidad.SetFocus
    ValidaDatos = False
    Exit Function
End If

If Trim(txtGlosa.Text) = "" Then
    MsgBox "Ingrese la Glosa.", vbInformation, "Aviso"
    txtGlosa.SetFocus
    ValidaDatos = False
    Exit Function
End If

'JOEP20181226 CP
If frOtrsMod.Visible = True And Trim(txtOtrsModa.Text) = "" Then
    MsgBox "Ingrese la descripción de la Modalidad.", vbInformation, "Aviso"
    txtOtrsModa.SetFocus
    ValidaDatos = False
    Exit Function
End If
'JOEP20181226 CP

ValidaDatos = True
End Function


Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
 KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        cmdModificar.SetFocus
    End If
End Sub

'JOEP20181218 CP
Private Sub CP_CargaComboxEdi(ByVal nParCod As Long)
Dim objCatalogoLlenaCombox As COMDCredito.DCOMCredito
Dim rsCatalogoCombox As ADODB.Recordset
Set objCatalogoLlenaCombox = New COMDCredito.DCOMCredito
Set rsCatalogoCombox = objCatalogoLlenaCombox.getCatalogoCombo("514", nParCod)

If Not (rsCatalogoCombox.BOF And rsCatalogoCombox.EOF) Then
    If nParCod = 49000 Then
        cboModalidad.Clear
        Call Llenar_Combo_con_Recordset(rsCatalogoCombox, cboModalidad)
        Call CambiaTamañoCombo(cboModalidad, 300)
    End If
End If
Set objCatalogoLlenaCombox = Nothing
RSClose rsCatalogoCombox
End Sub
'JOEP20181218 CP
