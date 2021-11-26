VERSION 5.00
Begin VB.Form frmCredReprogSolicitud 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Solicitud de Reprogramación"
   ClientHeight    =   5490
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9045
   Icon            =   "frmCredReprogSolicitud.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5490
   ScaleWidth      =   9045
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   7680
      TabIndex        =   27
      Top             =   5040
      Width           =   1170
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "Grabar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6360
      TabIndex        =   2
      Top             =   5040
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   3
      Top             =   5040
      Width           =   1170
   End
   Begin VB.Frame FraSolicitud 
      Caption         =   " Solicitud "
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   22
      Top             =   2400
      Width           =   8775
      Begin VB.CheckBox chkAutorizacionPF 
         Caption         =   "Solicitar Reprogramación por fuerza mayor"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   31
         Top             =   2040
         Width           =   4095
      End
      Begin VB.TextBox txtDiasReprog 
         Height          =   300
         Left            =   720
         MaxLength       =   4
         TabIndex        =   29
         Top             =   315
         Width           =   975
      End
      Begin VB.CheckBox chkAutorizacion 
         Caption         =   "Solicitar Reprogramación excepcional"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   25
         Top             =   1800
         Width           =   4095
      End
      Begin VB.TextBox txtMotivo 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   720
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   720
         Width           =   7935
      End
      Begin VB.Label lblFechaSol 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3600
         TabIndex        =   30
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label Label6 
         Caption         =   "Dias:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   28
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "Motivo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   24
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Nueva Fecha Venc.:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   2040
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Datos "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   5
      Top             =   600
      Width           =   8775
      Begin VB.Label Label1 
         Caption         =   "Cliente :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   21
         Top             =   370
         Width           =   735
      End
      Begin VB.Label lblNomCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   20
         Top             =   330
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "Cuota Venc:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1235
         Width           =   1095
      End
      Begin VB.Label lblCuotaVenc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   18
         Top             =   1200
         Width           =   735
      End
      Begin VB.Label Label5 
         Caption         =   "Monto Desemb :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   810
         Width           =   1215
      End
      Begin VB.Label lblMontoDesemb 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   16
         Top             =   765
         Width           =   1335
      End
      Begin VB.Label Label9 
         Caption         =   "Fecha Cuota Venc :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3090
         TabIndex        =   15
         Top             =   1230
         Width           =   1455
      End
      Begin VB.Label lblFecCuotaVenc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4680
         TabIndex        =   14
         Top             =   1200
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha Venc. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3480
         TabIndex        =   13
         Top             =   810
         Width           =   1095
      End
      Begin VB.Label lblFecVenc 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4680
         TabIndex        =   12
         Top             =   770
         Width           =   1335
      End
      Begin VB.Label Label15 
         Caption         =   "D.O.I. :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   11
         Top             =   370
         Width           =   615
      End
      Begin VB.Label lblDOI 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7440
         TabIndex        =   10
         Top             =   330
         Width           =   1215
      End
      Begin VB.Label Label17 
         Caption         =   "Moneda :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   9
         Top             =   810
         Width           =   735
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7440
         TabIndex        =   8
         Top             =   765
         Width           =   1215
      End
      Begin VB.Label Label19 
         Caption         =   "Atraso:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6600
         TabIndex        =   7
         Top             =   1230
         Width           =   615
      End
      Begin VB.Label lblAtraso 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7440
         TabIndex        =   6
         Top             =   1200
         Width           =   1215
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "Buscar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3960
      TabIndex        =   4
      Top             =   120
      Width           =   930
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      Texto           =   "Credito"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.CommandButton cmdAutorizar 
      Caption         =   "Autorizar"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   6240
      TabIndex        =   26
      Top             =   5040
      Width           =   1290
   End
End
Attribute VB_Name = "frmCredReprogSolicitud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmCredReprogSolicitud
'** Descripción : Formulario para solicitar reprogramación de créditos según TI-ERS010-2016
'** Creación : JUEZ, 20160216 09:00:00 AM
'*****************************************************************************************************

Option Explicit

Private Enum TipoAcceso
    TipoSolicitado = 0
    TipoAutorizado = 1
End Enum

Dim oDCred As COMDCredito.DCOMCredito
Dim fnTipo As TipoAcceso
Dim fnCuotasReprog As Integer
Dim fnPersoneria As Integer
Dim sMovNro As String
Dim fnSaldoReprog As Double

Dim nTipoOp As Integer 'Agrego JOEP20171214 ACTA220-2017
Dim nTpPermisoSolic As Integer 'Agrego JOEP20171214 ACTA220-2017

Dim nValidadOkey As Integer 'Agrego JOEP20171214 ACTA220-2017

Public Sub Inicio(ByVal pnTipo As Integer)
Dim PermisoAproReprog As COMNCredito.NCOMCredito 'Agrego JOEP20171214 ACTA220-2017
Set PermisoAproReprog = New COMNCredito.NCOMCredito '(3:Cargo si cumple Condicion, RF,Asesor,tasador),(4:Cargo si no cumple Condicion, Analista,Jefe Agencia,Coordinador de Credito)Agrego JOEP20171214 ACTA220-2017
nValidadOkey = 0 'Agrego JOEP20171214 ACTA220-2017

    fnTipo = pnTipo
    ActXCodCta.CMAC = gsCodCMAC
    ActXCodCta.Age = gsCodAge
    fnSaldoReprog = 0
    Select Case fnTipo
        Case TipoSolicitado
            chkAutorizacion.Visible = True
            FraSolicitud.Height = 2350
            CmdGrabar.Visible = True
            cmdAutorizar.Visible = False
            Me.Caption = "Solicitud de Reprogramación"
            
            'Inicion Agrego JOEP20171214 ACTA220-2017
            chkAutorizacion.Enabled = False
            chkAutorizacionPF.Enabled = False
            
            nTpPermisoSolic = PermisoAproReprog.ObtieneTipoPermisoReprog(gsCodCargo) ' Obtener el tipo de Permiso, Segun Cargo
                        
            If nTpPermisoSolic = 1 Or nTpPermisoSolic = 3 Or nTpPermisoSolic = 4 Then 'Agrego JOEP2018 nTpPermisoSolic = 1
            Else
                MsgBox "No tiene permiso para Solicitar", vbInformation, "Aviso"
                Exit Sub
            End If
            'Fin Agrego JOEP20171214 ACTA220-2017
            
        Case TipoAutorizado
            chkAutorizacion.Visible = False
            FraSolicitud.Height = 1815
            CmdGrabar.Visible = False
            cmdAutorizar.Visible = True
            Me.Caption = "Autorización de Reprogramación"
    End Select
            
    ValidarFechaActual
    Me.Show 1
End Sub

Private Sub ValidarFechaActual()
Dim lsFechaValidador As String

    lsFechaValidador = validarFechaSistema
    If lsFechaValidador <> "" Then
        If gdFecSis <> CDate(lsFechaValidador) Then
            MsgBox "La Fecha de tu sesión en el Negocio no coincide con la fecha del Sistema", vbCritical, "Aviso"
            Unload Me
            End
        End If
    End If
End Sub

Private Sub Limpiar()
    ActXCodCta.CMAC = gsCodCMAC
    ActXCodCta.Age = gsCodAge
    ActXCodCta.Prod = ""
    ActXCodCta.Cuenta = ""
    LblNomCliente.Caption = ""
    lblDOI.Caption = ""
    lblMontoDesemb.Caption = ""
    lblFecVenc.Caption = ""
    LblMoneda.Caption = ""
    lblCuotaVenc.Caption = ""
    lblFecCuotaVenc.Caption = ""
    lblAtraso.Caption = ""
    txtDiasReprog.Text = ""
    lblFechaSol.Caption = "__/__/____"
    txtMotivo.Text = ""
    chkAutorizacion.value = 0
    chkAutorizacionPF.value = 0 'Agrego JOEP20171214 ACTA220-2017
    HabilitaControles True
    If ActXCodCta.Enabled Then ActXCodCta.SetFocusProd
End Sub

Private Sub HabilitaControles(ByVal pbHabilitaBus As Boolean)
    ActXCodCta.Enabled = pbHabilitaBus
    CmdBuscar.Enabled = pbHabilitaBus
    FraSolicitud.Enabled = Not pbHabilitaBus
    If fnTipo = TipoSolicitado Then
        CmdGrabar.Enabled = Not pbHabilitaBus
    ElseIf fnTipo = TipoAutorizado Then
        cmdAutorizar.Enabled = Not pbHabilitaBus
    End If
End Sub

'Inicio Agrego JOEP20171214 ACTA220-2017
Private Sub chkAutorizacion_Click()
    If chkAutorizacion.value = 1 And chkAutorizacionPF.value = 1 Then
        MsgBox "Solo puede seleccionar una opción", vbInformation, "Aviso"
        chkAutorizacion.value = 0
        Exit Sub
    End If
    
    If chkAutorizacion.value = 1 Then
        nTipoOp = 1
    ElseIf chkAutorizacion.value = 0 And chkAutorizacionPF.value = 0 Then
        nTipoOp = 0
    End If
End Sub

Private Sub chkAutorizacionPF_Click()
    If chkAutorizacion.value = 1 And chkAutorizacionPF.value = 1 Then
        MsgBox "Solo puede seleccionar una opción", vbInformation, "Aviso"
        chkAutorizacionPF.value = 0
        Exit Sub
    End If
    
    If chkAutorizacionPF.value = 1 Then
        nTipoOp = 2
    ElseIf chkAutorizacion.value = 0 And chkAutorizacionPF.value = 0 Then
        nTipoOp = 0
    End If
End Sub
'Fin JOEP20171214 ACTA220-2017

Private Sub cmdAutorizar_Click()
Dim oNCred As COMNCredito.NCOMCredito

        If MsgBox("Se va a autorizar la reprogramación del crédito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                
        Set oNCred = New COMNCredito.NCOMCredito
            Call oNCred.RegistraReprogramacionEstado(ActXCodCta.NroCuenta, gdFecSis, gEstReprogAutorizado, fnSaldoReprog, sMovNro, CInt(lblAtraso.Caption), _
                                                     fnCuotasReprog, CDate(lblFecCuotaVenc.Caption), CDate(lblFechaSol.Caption))
        Set oNCred = Nothing
        
        MsgBox "Se ha autorizado la reprogramación del crédito", vbInformation, "Aviso"
        Limpiar
End Sub

Private Sub cmdBuscar_Click()
Dim R As ADODB.Recordset
Dim oDPers As COMDPersona.UCOMPersona
    Limpiar
    Set oDPers = frmBuscaPersona.Inicio()
    If Not oDPers Is Nothing Then
        Call FrmVerCredito.Inicio(oDPers.sPersCod, , , True, ActXCodCta)
        ActXCodCta.SetFocusCuenta
    End If
    Set oDPers = Nothing
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(ActXCodCta.NroCuenta) = 18 Then
            If CargaDatos Then
                HabilitaControles False
                If fnTipo = TipoSolicitado Then
                    If txtDiasReprog.Enabled Then txtDiasReprog.SetFocus
                ElseIf fnTipo = TipoAutorizado Then
                    If cmdAutorizar.Enabled Then cmdAutorizar.SetFocus
                End If
            Else
                HabilitaControles True
            End If
        Else
            MsgBox "Ingrese correctamente el crédito", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Function CargaDatos()
Dim rsCred As ADODB.Recordset, rs As ADODB.Recordset, rsSol As ADODB.Recordset
Dim lsEstadosValida As String
    CargaDatos = False
    
    Set oDCred = New COMDCredito.DCOMCredito
        Set rsCred = oDCred.RecuperaDatosCreditoVigente(ActXCodCta.NroCuenta, gdFecSis, 0)
    Set oDCred = Nothing
    
    If Not rsCred.EOF And Not rsCred.BOF Then
        If fnTipo = TipoSolicitado Then
            
            If rsCred!EsFaeReactiva = 1 Then  'ANGC20200824
                MsgBox "No está permitido reprogramar créditos con campaña REACTÍVA PERÚ o FAE MYPE" & vbCrLf & "Por favor coordinar con Finanzas y Productos Crediticios", vbInformation, "Aviso"
                Exit Function
            End If
            
            If rsCred!cTpoProdCod = "515" Or rsCred!cTpoProdCod = "516" Then
                MsgBox "El proceso de reprogramación no está disponible para créditos Leasing", vbInformation, "Aviso"
                Exit Function
            End If
            
            'lsEstadosValida = gEstReprogSolicitado & "," & gEstReprogPropuesto & "," & gEstReprogAutorizado & "," & gEstReprogAprobado 'comento JOEP20171214 ACTA220-2017
            lsEstadosValida = gEstReprogSolicitado & "," & gEstReprogPropuesto & "," & gEstReprogAutorizado & "," & gEstReprogAprobado & "," & gEstReprogSolicitadoReprogramado 'Agrego JOEP20171214 ACTA220-2017
            
            Set oDCred = New COMDCredito.DCOMCredito
                Set rs = oDCred.RecuperaColocacReprogramado(ActXCodCta.NroCuenta, lsEstadosValida)
            Set oDCred = Nothing
            
            If Not rs.EOF And Not rs.BOF Then
                'INICIO JOEP20171214 ACTA220-2017
                If rs!nPrdEstado = gEstReprogSolicitadoReprogramado Then
                    MsgBox "El crédito ya se encuentra en proceso de reprogramación. No es necesario seguir con los otros procesos, por favor ir a la Opción de Reprogramación.", vbInformation, "Aviso"
                Else
                'FIN JOEP20171214 ACTA220-2017
                    MsgBox "El crédito ya se encuentra en proceso de reprogramación. Su estado actual es " & rs!cPrdEstado, vbInformation, "Aviso"
                End If
                
                Limpiar
                Exit Function
            End If
        ElseIf fnTipo = TipoAutorizado Then
            Set oDCred = New COMDCredito.DCOMCredito
                Set rs = oDCred.RecuperaColocacReprogramado(ActXCodCta.NroCuenta, gEstReprogPropuesto)
            Set oDCred = Nothing
            
            If Not rs.EOF And Not rs.BOF Then
                If rs!nPrdEstado <> 208 Then 'Agrego JOEP20171214 ACTA220-2017
                    If rs!bSolicAutorizacion Then
                        Set oDCred = New COMDCredito.DCOMCredito
                            Set rsSol = oDCred.RecuperaColocacReprogramadoEstado(ActXCodCta.NroCuenta, gEstReprogSolicitado)
                        Set oDCred = Nothing
                        lblFechaSol.Caption = Format(rsSol!dFecNuevaCuotaVenc, "dd/MM/yyyy")
                        txtDiasReprog.Text = DateDiff("d", CDate(rs!dFecCuotaVenc), CDate(rsSol!dFecNuevaCuotaVenc))
                        txtMotivo.Text = rsSol!cMotivo
                        txtDiasReprog.Locked = True
                        txtMotivo.Locked = True
                    Else
                        MsgBox "El crédito no está disponible para ser autorizado", vbInformation, "Aviso"
                        Limpiar
                        Exit Function
                    End If
                    
                'Inicio Agrego JOEP20171214 ACTA220-2017
                Else
                    MsgBox "El crédito no está disponible para ser autorizado", vbInformation, "Aviso"
                    Limpiar
                    Exit Function
                End If
                'FIn Agrego JOEP20171214 ACTA220-2017
                
            Else
                MsgBox "El crédito no está disponible para ser autorizado", vbInformation, "Aviso"
                Limpiar
                Exit Function
            End If
        End If
        
        LblNomCliente.Caption = rsCred!cPersNombre
        lblDOI.Caption = Trim(rsCred!nDoi)
        lblMontoDesemb.Caption = Format(rsCred!nMontoCol, "#,##0.00")
        lblFecVenc.Caption = Format(rsCred!dFecVenc, "dd/MM/yyyy")
        LblMoneda.Caption = rsCred!cMoneda
        lblCuotaVenc.Caption = rsCred!nCuotaVenc
        lblFecCuotaVenc.Caption = Format(rsCred!dFecCuotaVenc, "dd/MM/yyyy")
        lblAtraso.Caption = rsCred!nDiasAtraso
        fnCuotasReprog = CInt(rsCred!nCuotasApr) - (CInt(rsCred!nCuotaVenc) - 1)
        fnPersoneria = CInt(rsCred!nPersPersoneria)
        fnSaldoReprog = rsCred!nSaldo
        CargaDatos = True
    Else
        MsgBox "No existen datos del crédito o no se encuentra vigente", vbInformation, "Aviso"
    End If
End Function

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub txtDiasReprog_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        txtMotivo.SetFocus
    End If
'Inicio Agrego JOEP20171214 ACTA220-2017
    If txtDiasReprog <> "" Then
        'If txtDiasReprog > 29 Then Covid joep20200725
        If txtDiasReprog >= 1 Then
            chkAutorizacion.Enabled = True
            chkAutorizacionPF.Enabled = True
        Else
            chkAutorizacion.Enabled = False
            chkAutorizacionPF.Enabled = False
            chkAutorizacion.value = 0
            chkAutorizacionPF.value = 0
        End If
    End If
'FIn Agrego JOEP20171214 ACTA220-2017

End Sub

Private Sub txtDiasReprog_KeyUp(KeyCode As Integer, Shift As Integer)
    If Trim(txtDiasReprog) <> "" Then
        lblFechaSol.Caption = DateAdd("d", CLng(txtDiasReprog), CDate(lblFecCuotaVenc.Caption))
    Else
        lblFechaSol.Caption = CDate(lblFecCuotaVenc.Caption)
    End If
'Inicio Agrego JOEP20171214 ACTA220-2017
    If txtDiasReprog <> "" Then
        'If txtDiasReprog > 29 Then
        If txtDiasReprog >= 1 Then 'Covid 20200725
            chkAutorizacion.Enabled = True
            chkAutorizacionPF.Enabled = True
        Else
            chkAutorizacion.Enabled = False
            chkAutorizacionPF.Enabled = False
            chkAutorizacion.value = 0
            chkAutorizacionPF.value = 0
        End If
    End If
'FIn Agrego JOEP20171214 ACTA220-2017
End Sub

Private Sub txtDiasReprog_LostFocus()
    If Trim(txtDiasReprog) <> "" Then
        lblFechaSol.Caption = DateAdd("d", CLng(txtDiasReprog), CDate(lblFecCuotaVenc.Caption))
    Else
        lblFechaSol.Caption = CDate(lblFecCuotaVenc.Caption)
    End If
    
'Inicio Agrego JOEP20171214 ACTA220-2017
    If txtDiasReprog <> "" Then
        'If txtDiasReprog > 29 Then
        If txtDiasReprog >= 1 Then 'Covid 20200725
            chkAutorizacion.Enabled = True
            chkAutorizacionPF.Enabled = True
        Else
            chkAutorizacion.Enabled = False
            chkAutorizacionPF.Enabled = False
            chkAutorizacion.value = 0
            chkAutorizacionPF.value = 0
        End If
    End If
'FIn Agrego JOEP20171214 ACTA220-2017
End Sub

Private Sub cmdGrabar_Click()
Dim oNCred As COMNCredito.NCOMCredito
Dim oDCredAct As COMDCredito.DCOMCredActBD
Dim rsExonera As ADODB.Recordset

nValidadOkey = 0

    If ValidaDatos Then
        Dim i As Integer
        Dim bExisteExonera As Boolean
        bExisteExonera = False
        'If chkAutorizacion.value = 1 Then 'comento JOEP20171214 ACTA220-2017
        If chkAutorizacion.value = 1 Or chkAutorizacionPF.value = 1 Then 'Agrego JOEP20171214 ACTA220-2017
            'Set rsExonera = frmCredReprogSolicExonera.ObtieneExoneracionesSolicitud 'comento JOEP20171214 ACTA220-2017
            Set rsExonera = frmCredReprogSolicExonera.ObtieneExoneracionesSolicitud(nTipoOp) 'Agrego JOEP20171214 ACTA220-2017
            
            If Not rsExonera.EOF And Not rsExonera.BOF Then
                rsExonera.MoveFirst
                For i = 0 To rsExonera.RecordCount - 1
                    If rsExonera.Fields(2) = 1 Then
                        bExisteExonera = True
                        rsExonera.MoveFirst
                        Exit For
                    End If
                    rsExonera.MoveNext
                Next i
            End If
            If Not bExisteExonera Then
                MsgBox "Debe seleccionar por lo menos una exoneración, de lo contrario desactivar el check de Solicitud de Autorización", vbInformation, "Aviso"
                chkAutorizacion.SetFocus
                Exit Sub
            End If
        End If
                       
        'If Not ValidaPromedioDiasAtraso Then MsgBox "Los dias de atraso promedio del crédito no cumple con el reglamento." & Chr(13) & "Se debe actualizar los informes de visita por parte de los Coordinadores de Créditos o Jefe de Agencia", vbInformation, "Aviso" 'Comento JOEP20171214 ACTA220-2017
        
        'Inicio Agrego JOEP20171214 ACTA220-2017
    If (nTpPermisoSolic <> 1) Then 'Mejora JOEP2018
        If nTpPermisoSolic <> 4 Then
            If Not ValidaCondicion Then
                Exit Sub
            End If
        Else
            If Not ValidaCondicion Then
                nValidadOkey = 1
            Else
                MsgBox "La solicitud de crédito debe ser atendido por: RF ó Asesor al Cliente ó Tasador.", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    Else 'Mejora JOEP2018
         If Not ValidaCondicion Then
            nValidadOkey = 1
        Else
            If nValidadOkey = 1 Then
            Else
                MsgBox "La solicitud de crédito debe ser atendido por: RF ó Asesor al Cliente ó Tasador.", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
    End If
        'Fin Agrego JOEP20171214 ACTA220-2017
        
        If MsgBox("Se va a solicitar la reprogramación del crédito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        
        sMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        If Not ValidaCondicion Then 'Agrego JOEP20171214 ACTA220-2017
            Set oNCred = New COMNCredito.NCOMCredito
                Call oNCred.RegistraReprogramacionEstado(ActXCodCta.NroCuenta, gdFecSis, gEstReprogSolicitado, fnSaldoReprog, sMovNro, CInt(lblAtraso.Caption), fnCuotasReprog, _
                                                         CDate(lblFecCuotaVenc.Caption), CDate(lblFechaSol.Caption), txtMotivo.Text, IIf(chkAutorizacion.value = 1, chkAutorizacion.value, IIf(chkAutorizacionPF.value = 1, chkAutorizacionPF.value, 0)), , rsExonera)
            Set oNCred = Nothing
        Else
            'Inicio Agrego JOEP20171214 ACTA220-2017
            Set oNCred = New COMNCredito.NCOMCredito
                Call oNCred.RegistraReprogramacionEstado(ActXCodCta.NroCuenta, gdFecSis, gEstReprogSolicitadoReprogramado, fnSaldoReprog, sMovNro, CInt(lblAtraso.Caption), fnCuotasReprog, _
                                                         CDate(lblFecCuotaVenc.Caption), CDate(lblFechaSol.Caption), txtMotivo.Text, IIf(chkAutorizacion.value = 1, chkAutorizacion.value, IIf(chkAutorizacionPF.value = 1, chkAutorizacionPF.value, 0)), , rsExonera)
            Set oNCred = Nothing
            'Fin Agrego JOEP20171214 ACTA220-2017
        End If
        
        If GenerarSolicitudReprogramacion Then
            MsgBox "La solicitud de reprogramación fue registrada correctamente", vbInformation, "Aviso"
            Limpiar
        Else
            Set oDCredAct = New COMDCredito.DCOMCredActBD
                Call oDCredAct.dEliminaColocacReprogramado(ActXCodCta.NroCuenta, gEstReprogSolicitado, sMovNro)
            Set oDCredAct = Nothing
        End If
    End If
End Sub

'Inicio Agrego JOEP20171214 ACTA220-2017
Private Function ValidaCondicion() As Boolean
Dim rsVCon As ADODB.Recordset
Set oDCred = New COMDCredito.DCOMCredito

Set rsVCon = oDCred.ObtieneCondiciones(ActXCodCta.NroCuenta)

ValidaCondicion = True

If nValidadOkey <> 1 Then
      'Debe estar al dia
        If CInt(lblAtraso.Caption) > 0 Then
            MsgBox "El crédito no está al día tiene " & lblAtraso.Caption & " días de atraso, por favor comunicarse con su Analista de Crédito, para su evaluación.", vbInformation, "Aviso"
            ValidaCondicion = False
            Exit Function
        End If
        
        'Max. 29 dias
        If DateDiff("d", CDate(lblFecCuotaVenc.Caption), CDate(lblFechaSol.Caption)) > 29 Then
            MsgBox "Los dias a reprogramar exceden los 29 días, por favor comunicarse con su Analista de Crédito, para su evaluación.", vbInformation, "Aviso"
            ValidaCondicion = False
            Exit Function
        End If
        
        If Not (rsVCon.EOF And rsVCon.BOF) Then
            
            '8 dias max. en su ultimo pago
            If rsVCon!nDiasAtrasoUltPago > 8 Then
                MsgBox "Tiene " & rsVCon!nDiasAtrasoUltPago & " días de atraso en su último Pago, por favor comunicarse con su Analista de Crédito, para su evaluación.", vbInformation, "Aviso"
                ValidaCondicion = False
                Exit Function
            End If
            
            'Ult. RCC 100
            If rsVCon!nRCC <> 100 Then
                MsgBox "No se encuentra 100% Normal en el Ultimo RCC, por favor comunicarse con su Analista de Crédito, para su evaluación.", vbInformation, "Aviso"
                ValidaCondicion = False
                Exit Function
            End If
            
            'Inicio JOEP2018
                If (rsVCon!nTotalCuota = 1 And nTpPermisoSolic = 4) Then 'Si es un Credito a una sola cuota y le solicita una Analista.
                    ValidaCondicion = False
                    Exit Function
                ElseIf (rsVCon!nTotalCuota = 1 And nTpPermisoSolic = 1) Then 'Si es un Credito a una sola cuota y le solicita una JA,Coord.
                    nValidadOkey = 1
                ElseIf (rsVCon!nTotalCuota = 1 And nTpPermisoSolic = 3) Then
                    MsgBox "La solicitud de crédito debe ser atendido por: Analista de Crédito.", vbInformation, "Aviso"
                    ValidaCondicion = False
                    Exit Function
                End If
            'Fin JOEP2018
        End If
    
        '8 dias max. Promedio en los ult. 6 meses.
        If Not ValidaPromedioDiasAtraso Then
            MsgBox "Excede los 8 dias de atraso max. promedio en los ultimos 6 meses, por favor comunicarse con su Analista de Crédito, para su evaluación.", vbInformation, "Aviso"
            ValidaCondicion = False
            Exit Function
        End If
Else
    ValidaCondicion = False
End If

   Set oDCred = Nothing

End Function
'Fin Agrego JOEP20171214 ACTA220-2017

Private Function ValidaDatos() As Boolean
'JOEP20200703 Cambio covid-19
Dim rsCovid As ADODB.Recordset
Dim objCovid As COMDCredito.DCOMCredito
Set objCovid = New COMDCredito.DCOMCredito

Dim nDiasAtraso As Integer
nDiasAtraso = IIf(txtDiasReprog.Text = "", -1, txtDiasReprog.Text)

Set rsCovid = objCovid.ReprogramacionValidaDatosCovid(ActXCodCta.NroCuenta, nDiasAtraso, Format(lblFechaSol.Caption, "yyyymmdd"), Format(lblFecCuotaVenc.Caption, "yyyymmdd"), txtMotivo.Text, lblAtraso.Caption)

ValidaDatos = False

If Not (rsCovid.BOF And rsCovid.EOF) Then
        If rsCovid!MsgBox <> "" Then
            MsgBox rsCovid!MsgBox, vbInformation, "Aviso"
                Select Case rsCovid!SetFocus
                    Case 1
                        txtDiasReprog.SetFocus
                    Case 2
                        txtMotivo.SetFocus
                End Select
            Exit Function
        End If
End If
'JOEP20200703 Cambio covid-19

'comento JOEP20200703 Cambio covid-19
'ValidaDatos = False
    
'    If txtDiasReprog.Text = "" Or txtDiasReprog.Text = "0" Then
'        MsgBox "Debe ingresar correctamente los dias", vbInformation, "Aviso"
'        txtDiasReprog.SetFocus
'        Exit Function
'    End If
'    If Not IsDate(lblFechaSol.Caption) Then
'        MsgBox "Ingrese una fecha válida", vbInformation, "Aviso"
'        txtDiasReprog.SetFocus
'        Exit Function
'    End If
'
'    If CDate(lblFechaSol.Caption) <= CDate(gdFecSis) Then
'        MsgBox "La nueva fecha no debe ser menor o igual a la fecha actual", vbInformation, "Aviso"
'        txtDiasReprog.SetFocus
'        Exit Function
'    End If
'
'    If CDate(lblFechaSol.Caption) <= CDate(lblFecCuotaVenc.Caption) Then
'        MsgBox "La nueva fecha no debe ser menor o igual a la fecha de vencimiento de la cuota", vbInformation, "Aviso"
'        txtDiasReprog.SetFocus
'        Exit Function
'    End If
'
'    If Trim(Replace(Replace(txtMotivo.Text, Chr(10), ""), Chr(13), "")) = "" Then
'        MsgBox "Ingrese correctamente el motivo", vbInformation, "Aviso"
'        txtMotivo.SetFocus
'        Exit Function
'    End If
'comento JOEP20200703 Cambio covid-19
  
    'If chkAutorizacion.value = 0 Then 'comento JOEP20171214 ACTA220-2017
    'If chkAutorizacion.value = 0 And chkAutorizacionPF.value = 0 Then 'Se comento para que valide >29 y <29 JOEP20200413
        If VerificaReprogramacionAnterior Then
            MsgBox "El crédito ya fue reprogramado anteriormente, no podrá registrarse una nueva solicitud en un plazo no menor de 12 meses de la reprogramación", vbInformation, "Aviso"
            Exit Function
        End If
        
'Comento JOEP20171214 ACTA220-2017
'        If CInt(lblAtraso.Caption) > 0 Then
'            MsgBox "El crédito debe estar al dia", vbInformation, "Aviso"
'            Exit Function
'        End If
'
'        If DateDiff("d", CDate(lblFecCuotaVenc.Caption), CDate(lblFechaSol.Caption)) > 29 Then
'            MsgBox "Los dias a reprogramar no pueden ser mas de 29 días", vbInformation, "Aviso"
'            Exit Function
'        End If
'Comento JOEP20171214 ACTA220-2017

    'End If
    
    ValidaDatos = True
End Function

Private Function ValidaPromedioDiasAtraso() As Boolean
Dim oDParam As COMDCredito.DCOMParametro
Dim nDiasAtrasoProm As Integer

    ValidaPromedioDiasAtraso = False
    
    Set oDParam = New COMDCredito.DCOMParametro
    
    Set oDCred = New COMDCredito.DCOMCredito
        nDiasAtrasoProm = oDCred.ObtienePromedioDiasAtraso(ActXCodCta.NroCuenta, gdFecSis, ObtieneParametro(3502))
    Set oDCred = Nothing
    
    If nDiasAtrasoProm <= ObtieneParametro(3501) Then ValidaPromedioDiasAtraso = True
End Function

Private Function ObtieneParametro(ByVal pnCodigo As Long) As Double
Dim oDParam As COMDCredito.DCOMParametro
Dim rs As ADODB.Recordset

    Set oDParam = New COMDCredito.DCOMParametro
        ObtieneParametro = oDParam.RecuperaValorParametro(pnCodigo)
    Set oDParam = Nothing
End Function

Private Function VerificaReprogramacionAnterior() As Boolean
Dim rs As ADODB.Recordset
    VerificaReprogramacionAnterior = True
    
    Set oDCred = New COMDCredito.DCOMCredito
        Set rs = oDCred.RecuperaColocacReprogramadoEstado(ActXCodCta.NroCuenta, gEstReprogReprogramado)
    Set oDCred = Nothing
    
    If Not rs.BOF And Not rs.EOF Then
        If DateDiff("m", CDate(rs!dPrdEstado), CDate(gdFecSis)) <= 12 Then Exit Function
    End If
    
    VerificaReprogramacionAnterior = False
End Function

Private Function GenerarSolicitudReprogramacion() As Boolean
Dim oDAge As COMDConstantes.DCOMAgencias
Dim sAgeUbiGeoDesc As String
Dim rsRelac As ADODB.Recordset
Dim cPersNombre As String
Dim cDOI As String
Dim cMotivo1 As String, cMotivo2 As String, cMotivo3 As String, cMotivo4 As String

Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Dim sArchivo As String
    
    GenerarSolicitudReprogramacion = False
    
    On Error GoTo ErrorGenerarSolicitud
    
    Set oDAge = New COMDConstantes.DCOMAgencias
        sAgeUbiGeoDesc = RTrim(oDAge.RecuperaAgencias(gsCodAge)!cUbiGeoDescripcion)
    Set oDAge = Nothing
    
    If fnPersoneria <> 1 Then
        Set oDCred = New COMDCredito.DCOMCredito
            Set rsRelac = oDCred.RecuperaRelacPers(ActXCodCta.NroCuenta)
        Set oDCred = Nothing
        
        Do While Not rsRelac.EOF
            If rsRelac!nPersPersoneria = 1 Then
                cPersNombre = rsRelac!cPersNombre
                cDOI = rsRelac!DNI
                rsRelac.MoveLast
            End If
            rsRelac.MoveNext
        Loop
    Else
        cPersNombre = LblNomCliente.Caption
        cDOI = lblDOI.Caption
    End If
    
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = False
    
    Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\SolicitudReprogramacionCred.doc")
    Set oRange = oWord.ActiveDocument.Content
    
    sArchivo = App.Path & "\Spooler\SolicitudReprog_" & ActXCodCta.NroCuenta & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyyMMdd") & "_" & Format(Time, "hhmmss") & ".doc"
    oDoc.SaveAs (sArchivo)
    
    With oWord.Selection.Find
        .Text = "<<cLugar>>"
        .Replacement.Text = sAgeUbiGeoDesc
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<cFecha>>"
        .Replacement.Text = ArmaFecha(gdFecSis)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<cAgeDescripcion>>"
        .Replacement.Text = gsNomAge
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<cCtaCod>>"
        .Replacement.Text = ActXCodCta.NroCuenta
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With

    With oRange.Find
        .ClearFormatting
        .Text = "<<cMotivo>>"
        .Forward = True
        .Wrap = wdFindContinue
        Do While .Execute
            oRange.Text = txtMotivo.Text
        Loop
    End With
    
    With oWord.Selection.Find
        .Text = "<<cNombreCliente>>"
        .Replacement.Text = PstaNombre(cPersNombre)
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    With oWord.Selection.Find
        .Text = "<<cDOI>>"
        .Replacement.Text = cDOI
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .Execute Replace:=wdReplaceAll
    End With
    
    oDoc.Close
    Set oDoc = Nothing
    
    Set oWord = CreateObject("Word.Application")
    oWord.Visible = True
    
    Set oDoc = oWord.Documents.Open(sArchivo)
    Set oDoc = Nothing
    Set oWord = Nothing
    GenerarSolicitudReprogramacion = True
    
    Exit Function
ErrorGenerarSolicitud:
    MsgBox "Hubo un error en la generación de la solicitud: " & Err.Description, vbInformation, "Aviso"
End Function

Private Sub cmdCancelar_Click()
    Limpiar
End Sub

Private Sub txtMotivo_KeyPress(KeyAscii As Integer)
    If Len(txtMotivo.Text) = 0 Then
        KeyAscii = fgIntfMayusculas(KeyAscii)
    End If
End Sub

Private Sub txtMotivo_LostFocus()
    If Len(txtMotivo.Text) > 0 Then
        txtMotivo.Text = UCase(Left(txtMotivo.Text, 1)) & Mid(txtMotivo.Text, 2, Len(txtMotivo.Text))
    End If
End Sub

