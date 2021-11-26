VERSION 5.00
Begin VB.Form frmCredRechazo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Rechazo de Credito"
   ClientHeight    =   6150
   ClientLeft      =   4260
   ClientTop       =   2025
   ClientWidth     =   7395
   Icon            =   "frmCredRechazo.frx":0000
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   7395
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame3 
      Caption         =   "Credito"
      Height          =   1245
      Left            =   60
      TabIndex        =   29
      Top             =   0
      Width           =   7320
      Begin VB.Frame FraListaCred 
         Caption         =   "&Lista Creditos"
         Height          =   960
         Left            =   4860
         TabIndex        =   32
         Top             =   180
         Width           =   2385
         Begin VB.ListBox LstCred 
            Height          =   645
            Left            =   60
            TabIndex        =   33
            Top             =   225
            Width           =   2265
         End
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "&Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3870
         TabIndex        =   31
         Top             =   525
         Width           =   900
      End
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   435
         Left            =   195
         TabIndex        =   30
         Top             =   495
         Width           =   3660
         _ExtentX        =   6456
         _ExtentY        =   767
         Texto           =   "Credito :"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin VB.Frame FraCredito 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   1365
      Left            =   75
      TabIndex        =   17
      Top             =   2325
      Width           =   7275
      Begin VB.Label LblMoneda 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4890
         TabIndex        =   28
         Top             =   960
         Width           =   1590
      End
      Begin VB.Label LblDestinoC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   27
         Top             =   600
         Width           =   1950
      End
      Begin VB.Label LblMontoS 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   26
         Top             =   930
         Width           =   1260
      End
      Begin VB.Label LblTipoC 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1920
         TabIndex        =   25
         Top             =   285
         Width           =   3780
      End
      Begin VB.Label Label8 
         Caption         =   "Destino:"
         Height          =   255
         Left            =   615
         TabIndex        =   24
         Top             =   630
         Width           =   735
      End
      Begin VB.Label Label7 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   4155
         TabIndex        =   23
         Top             =   945
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Monto Solicitado:"
         Height          =   255
         Left            =   585
         TabIndex        =   22
         Top             =   945
         Width           =   1335
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   570
         TabIndex        =   21
         Top             =   270
         Width           =   615
      End
      Begin VB.Label Label11 
         Caption         =   "Datos del Crédito"
         Height          =   255
         Left            =   225
         TabIndex        =   20
         Top             =   -30
         Width           =   1260
      End
      Begin VB.Label Label14 
         Caption         =   "Estado:"
         Height          =   270
         Left            =   4170
         TabIndex        =   19
         Top             =   630
         Width           =   585
      End
      Begin VB.Label LblEstCredito 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4875
         TabIndex        =   18
         Top             =   630
         Width           =   1605
      End
   End
   Begin VB.Frame FraCliente 
      Enabled         =   0   'False
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   75
      TabIndex        =   9
      Top             =   1290
      Width           =   7275
      Begin VB.Label LblDIdent 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   4905
         TabIndex        =   16
         Top             =   255
         Width           =   1575
      End
      Begin VB.Label LblCodCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   15
         Top             =   240
         Width           =   1470
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   600
         TabIndex        =   13
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Doc. Identidad:"
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   3735
         TabIndex        =   12
         Top             =   255
         Width           =   1095
      End
      Begin VB.Label LblNomCliente 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   1440
         TabIndex        =   11
         Top             =   600
         Width           =   5040
      End
      Begin VB.Label Label12 
         Caption         =   "Datos del Cliente"
         Height          =   195
         Left            =   225
         TabIndex        =   10
         Top             =   -30
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Enabled         =   0   'False
      Height          =   1860
      Left            =   60
      TabIndex        =   3
      Top             =   3735
      Width           =   7275
      Begin VB.TextBox TxtComenta 
         Height          =   570
         Left            =   1500
         MaxLength       =   255
         MultiLine       =   -1  'True
         TabIndex        =   35
         Top             =   1140
         Width           =   5070
      End
      Begin VB.ComboBox CmbMotivo 
         Height          =   315
         Left            =   1500
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   34
         Top             =   705
         Width           =   5040
      End
      Begin VB.Label Label4 
         Caption         =   "Motivo:"
         Height          =   255
         Left            =   600
         TabIndex        =   8
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "Analista:"
         Height          =   255
         Left            =   585
         TabIndex        =   7
         Top             =   390
         Width           =   720
      End
      Begin VB.Label Label13 
         Caption         =   "Rechazo del Crédito"
         Height          =   225
         Left            =   210
         TabIndex        =   6
         Top             =   -30
         Width           =   1440
      End
      Begin VB.Label LblAnalista 
         BackColor       =   &H8000000E&
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   1515
         TabIndex        =   5
         Top             =   330
         Width           =   5040
      End
      Begin VB.Label Label10 
         Caption         =   "Comentario:"
         Height          =   255
         Left            =   570
         TabIndex        =   4
         Top             =   1305
         Width           =   900
      End
   End
   Begin VB.CommandButton CmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   1290
      TabIndex        =   2
      Top             =   5655
      Width           =   1275
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   5865
      TabIndex        =   1
      Top             =   5640
      Width           =   1455
   End
   Begin VB.CommandButton CmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   45
      TabIndex        =   0
      Top             =   5655
      Width           =   1230
   End
End
Attribute VB_Name = "frmCredRechazo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'1:Rechazo
'2:Retiro
Private pnRechazoRetiro As Integer
Private pbRefinanc As Boolean
Private fbAutorizacion As Boolean
Dim objPista As COMManejador.Pista


'20060329
'Modificamos este método para que puede diferenciar el orden de
'rechazo:
'1. Al rechazar la solicitud de crédito
'2. Al Rechazar el credito sugerido
'para elllo agregamos un índice como argumento del procedimiento
Public Sub Rechazar(Optional nTipoRechazo As Long = 1, Optional ByVal psCtaCod As String = "") 'RECO20160627 ERS002-2016 Se agregó "psCtaCod"
'    pnRechazoRetiro = 1
'    Me.Caption = "Rechazar Credito"
    
    Select Case nTipoRechazo
        Case 3
            '20060329
            'en este caso se rechaza solicitud
            Me.Caption = "Rechazar Solicitud de Crédito"
            'pnRechazoRetiro = 3
        Case 4
            '20060329
            'en este caso se rechaza la sugerencia de credito
            Me.Caption = "Rechazar Crédito Sugerido"
            'pnRechazoRetiro = 4
    End Select
    
    '20060330
    'asignamos el tipo de rechazo
    pnRechazoRetiro = nTipoRechazo
    'RECO20160627 ERS002-2016 ******************************
    If psCtaCod <> "" Then
        ActxCta.NroCuenta = psCtaCod
        Call CargaDatos(psCtaCod)
        Call HabilitaIngreso(True)
        CmdNuevo.Enabled = False
        fbAutorizacion = True
    End If
    'RECO FIN **********************************************
    Me.Show 1
    
End Sub

Public Sub Retirar()
    pnRechazoRetiro = 2
    Me.Caption = "Anular Credito"
    Me.Show 1
End Sub

'JUEZ 20121112 **************************************************************
Public Sub RechazarPorNiveles(ByVal psCtaCod As String, Optional nTipoRechazo As Long = 1)
    pnRechazoRetiro = nTipoRechazo
    ActxCta.NroCuenta = psCtaCod
    
    pnRechazoRetiro = nTipoRechazo
    ActxCta.Enabled = False
    LstCred.Enabled = False
    CmdBuscar.Enabled = False
    CmdNuevo.Visible = False
    ActxCta_KeyPress (13)
    Me.Show 1
End Sub
'END JUEZ *******************************************************************

Public Sub HabilitaIngreso(ByVal pbHabilita As Boolean)
    Frame3.Enabled = Not pbHabilita
    CmdGrabar.Enabled = pbHabilita
    Frame1.Enabled = pbHabilita
End Sub

Public Sub LimpiarPantalla()
    LimpiaControles Me
    ActxCta.NroCuenta = ""
    ActxCta.Age = gsCodAge
    ActxCta.CMAC = gsCodCMAC
    HabilitaIngreso False
    LstCred.Clear
    fbAutorizacion = False 'RECO20160627 ERS002-2016
End Sub

Public Function CargaDatos(ByVal psCtaCod As String) As Boolean
'Dim oDCred As COMDCredito.DCOMCredito
Dim oNegCred As COMNCredito.NCOMCredito
Dim MatRechazo As Variant
Dim bRefinanciado As Boolean

Dim R As ADODB.Recordset

    On Error GoTo ErrorCargaDatos
    '20060330
    'código obsoleto de la caja
'    Set oDCred = New COMDCredito.DCOMCredito
    
    'If pnRechazoRetiro = 1 Then 'Rechazo
    '    Set R = oDCred.RecuperaDatosComunes(psCtaCod, , Array(gColocEstSolic, gColocEstSug))
    'Else
    '    Set R = oDCred.RecuperaDatosComunes(psCtaCod, , Array(gColocEstAprob, gColocEstRefMor, gColocEstRefNorm, gColocEstRefVenc))
    'End If
'    If pnRechazoRetiro = 1 Then 'Rechazo
'        MatRechazo = Array(gColocEstSolic, gColocEstSug)
'    Else
'        MatRechazo = Array(gColocEstAprob, gColocEstRefMor, gColocEstRefNorm, gColocEstRefVenc)
'    End If
    
    '20060330
    'nuevo codigo
    'Se modificaron las lineas arriba
    'se asigna un selector para filtrar los datos a recuperar
    Select Case pnRechazoRetiro
        Case 1
            'rechazo de credito
            MatRechazo = Array(gColocEstSolic, gColocEstSug)
        Case 2
            'retiro de crédito
            MatRechazo = Array(gColocEstAprob, gColocEstRefMor, gColocEstRefNorm, gColocEstRefVenc)
        Case 3
            'rechazo de solicutd de credito
            MatRechazo = Array(gColocEstSolic)
        Case 4
            'rechazo de credtito sugerido
            MatRechazo = Array(gColocEstSug)
    End Select
    
    Set oNegCred = New COMNCredito.NCOMCredito
    Call oNegCred.CargarDatosRechazo(psCtaCod, MatRechazo, R, bRefinanciado)
    Set oNegCred = Nothing
    
    If Not R.BOF And Not R.EOF Then
        If CDate(Format(IIf(IsNull(R!dVigencia), "01/01/1900", R!dVigencia), "dd/mm/yyyy")) = gdFecSis Or IsNull(R!dVigencia) Or Not IsNull(R!dVigencia) Then
            CmdGrabar.Enabled = True
            CargaDatos = True
            LblCodCliente.Caption = R!cPersCod
            LblDIdent.Caption = IIf(IsNull(R!DNI), "", R!DNI)
            LblNomCliente.Caption = PstaNombre(R!cTitular)
            LblTipoC.Caption = R!cTipoCredDescrip
            LblDestinoC.Caption = R!cDestinoDescripcion
            LblEstCredito.Caption = R!cEstado
            LblMontoS.Caption = Format(R!nMontoSol, "#0.00")
            LblMoneda.Caption = R!cMoneda
            LblAnalista.Caption = PstaNombre(R!cAnalista)
            'Set oNegCred = New COMNCredito.NCOMCredito
            pbRefinanc = bRefinanciado 'oNegCred.EsRefinanciado(psCtaCod)
            'Set oNegCred = Nothing
        Else
            CargaDatos = False
            CmdGrabar.Enabled = False
        End If
    Else
        CargaDatos = False
        CmdGrabar.Enabled = False
    End If
    'R.Close
    'Set R = Nothing
    Exit Function

ErrorCargaDatos:
        MsgBox Err.Description, vbCritical, "Aviso"

End Function

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not CargaDatos(ActxCta.NroCuenta) Then
            MsgBox "No se pudo Encontrar el Credito", vbInformation, "Aviso"
            HabilitaIngreso False
        Else
            HabilitaIngreso True
            'CmbMotivo.SetFocus
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona 'UPersona
'20060330
'agregamos una variable que contendar los estados a filtrar
Dim cEstadoCredito, cMensaje As String

    On Error GoTo ErrorCmdBuscar_Click
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        
        '20060330
        'código obsoleto de la caja
'        If pnRechazoRetiro = 1 Then 'Rechazo
'            Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstSolic, gColocEstSug))
'        Else
'            Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstAprob, 0))
'        End If
        
        '20060330
        'Modificamos las líneas anteriores
        'agregamos un selector para las opciones creadas
        Select Case pnRechazoRetiro
            Case 1
                
                'rechazo de credito solicitado y sugerido
                cEstadoCredito = Array(gColocEstSolic, gColocEstSug)
                
            Case 2
            
                'retiro de credito
                cEstadoCredito = Array(gColocEstAprob, 0)
                
            Case 3
            
                'rechazo de credito solicitado
                cEstadoCredito = Array(gColocEstSolic)
                
            Case 4
            
                'rechazo de credito sugerido
                cEstadoCredito = Array(gColocEstSug)
                
        End Select
        
        'llenamos el recordset con los resultados de la consulta
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , cEstadoCredito)
        
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
    End If
    If LstCred.ListCount = 0 Then
'        If pnRechazoRetiro = 1 Then
'            MsgBox "El Cliente No Tiene Creditos Solicitados o Sugeridos", vbInformation, "Aviso"
'        Else
'            MsgBox "El Cliente No Tiene Creditos Aprobados", vbInformation, "Aviso"
'        End If

        '20060330
        'Cambiamos las lineas anteriores por un selector
        Select Case pnRechazoRetiro
            Case 1
                'no hay creditos solicitados o sugeridos
                cMensaje = "El Cliente No Tiene Creditos Solicitados o Sugeridos"
                
            Case 2
                'no hay creditos aprobados
                cMensaje = "El Cliente No Tiene Creditos Aprobados"
                
            Case 3
                'no hay creditos solicitados
                cMensaje = "El Cliente No Tiene Creditos Solicitados"
                
            Case 4
                'no hay creditos sugeridos
                cMensaje = "El Cliente No Tiene Creditos Sugeridos"
                
        End Select
        
        'mostramos el comentario
        MsgBox cMensaje, vbInformation, "Aviso"
        
    End If
        
    Exit Sub

ErrorCmdBuscar_Click:
    MsgBox Err.Description, vbCritical, "Aviso"

End Sub

Private Sub cmdGrabar_Click()
Dim oNegCredito As COMNCredito.NCOMCredito
Dim oCredNiv As New COMDCredito.DCOMNivelAprobacion 'RECO20160627 ERS002-2016
Dim sMovNro As String 'RECO20160627 ERS002-2016
Dim sError, cMensaje As String
    If Not ValidarDatos Then Exit Sub 'FRHU 20140514 Observacion Otros
    '20060330
    'código obsoleto de la caja
'    If pnRechazoRetiro = 1 Then
'        If MsgBox("Se va a Rechazar el Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
'            Exit Sub
'        End If

        '20060330
        'Agregamos un selector
        Select Case pnRechazoRetiro
            Case 1
                'mensaje de rechazo de credito
                cMensaje = "Se va a Rechazar el Credito, Desea Continuar ?"
                gsOpeCod = gCredRechazarCreditos
            Case 2
                'mensaje de retiro de credito
                cMensaje = "Se va a Retirar el Credito, Desea Continuar ?"
                gsOpeCod = gCredRetirarCreditos
            Case 3
                'mensaje de rechazo de credito solicitado
                cMensaje = "Se va rechazar la Solicitud de Crédito. ¿Desea continuar?"
                gsOpeCod = gCredRechazarSolicitud
            Case 4
                'mensaje de rechazo de credito sugerido
                cMensaje = "Se va rechazar el Crédito Sugerido. ¿Desea continuar?"
                gsOpeCod = gCredRechazarSugerencia
        End Select
    
        'realizamos la pregunta
        If MsgBox(cMensaje, vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Exit Sub
        End If
'    Else
'        If MsgBox("Se va a Retirar el Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
'            Exit Sub
'        End If
'    End If
    Set oNegCredito = New COMNCredito.NCOMCredito
    If pnRechazoRetiro = 1 Then
        sError = oNegCredito.RechazoRetiroCredito(ActxCta.NroCuenta, TxtComenta.Text, CInt(Trim(Right(CmbMotivo.Text, 20))), _
            gdFecSis, gsCodAge, gsCodUser, CDbl(LblMontoS.Caption), pnRechazoRetiro, False)
    
    ''*** PEAC 20090126
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActxCta.NroCuenta, gCodigoCuenta
        
    Else
        sError = oNegCredito.RechazoRetiroCredito(ActxCta.NroCuenta, TxtComenta.Text, CInt(Trim(Right(CmbMotivo.Text, 20))), _
            gdFecSis, gsCodAge, gsCodUser, CDbl(LblMontoS.Caption), pnRechazoRetiro, pbRefinanc)
            
    ''*** PEAC 20090126
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar, , ActxCta.NroCuenta, gCodigoCuenta
    'RECO20160627 ERS002-2016******************************************************************************
    'If fbAutorizacion Then
    sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Call oCredNiv.ActualizaEstadoAutoExoCab(ActxCta.NroCuenta, EstadoAutoExonera.gEstadoRechazado, sMovNro)
    'End If
    'RECO FIN *********************************************************************************************
    End If
    Set oNegCredito = Nothing
    If fbAutorizacion Then
        Call cmdsalir_Click
    Else
        Call cmdNuevo_Click
    End If
    If sError = "" Then
        HabilitaIngreso False
        'LimpiaPantalla
    Else
        MsgBox sError, vbInformation, "Aviso"
        HabilitaIngreso True
    End If
End Sub

Function VerificaTipoCredito() As String
'Devuelve AGRICOLA si  para creditos agricolas
'Devuelve Comercial si es para pequeñas empresas
'Devuelve consumo si es de consumo
 Dim sTipoCredito As String
 Dim sSubTipoCredito As String
 Dim sCuenta As String

 sTipoCredito = Mid(ActxCta.NroCuenta, 6, 1) & "00"
 sSubTipoCredito = Mid(ActxCta.NroCuenta, 6, 1) & Mid(ActxCta.NroCuenta, 7, 2)
 If (sTipoCredito = 100 And sSubTipoCredito = "102") Or _
    (sTipoCredito = "200" And sSubTipoCredito = "202") Then
    VerificaTipoCredito = "AGRICOLA"
 ElseIf (sTipoCredito = "300" And sSubTipoCredito = "301") Or _
         (sTipoCredito = "300" And sSubTipoCredito = "302") Or _
         (sTipoCredito = "300" And sSubTipoCredito = "303") Or _
         (sTipoCredito = "300" And sSubTipoCredito = "304") Or _
         (sTipoCredito = "300" And sSubTipoCredito = "305") Or _
         (sTipoCredito = "300" And sSubTipoCredito = "320") Then
          VerificaTipoCredito = "CONSUMO"
 End If
End Function
Private Sub cmdNuevo_Click()
    Call LimpiarPantalla
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Dim oCons As COMDConstantes.DCOMConstantes
Dim rs As ADODB.Recordset
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    pbRefinanc = False
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
Set oCons = New COMDConstantes.DCOMConstantes
Set rs = oCons.RecuperaConstantes(gColocMotivRechazo)
Set oCons = Nothing
Call Llenar_Combo_con_Recordset(rs, CmbMotivo)
    fbAutorizacion = False 'RECO20160627 ERS002-2016
'    Call CargaComboConstante(gColocMotivRechazo, CmbMotivo)
    CentraSdi Me
        
Set objPista = New COMManejador.Pista
gsOpeCod = gCredRechazarCreditos
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub LstCred_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
            ActxCta.NroCuenta = LstCred.Text
        End If
    End If
End Sub
'FRHU 20140514 Observacion
Private Function ValidarDatos() As Boolean
    ValidarDatos = True
    If CmbMotivo.ListIndex = -1 Then
        MsgBox "Debe seleccionar el Motivo"
        ValidarDatos = False
    End If
End Function
'FIN FRHU 20140514
