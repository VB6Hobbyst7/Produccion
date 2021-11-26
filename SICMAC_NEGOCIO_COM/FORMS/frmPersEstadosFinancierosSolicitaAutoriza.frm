VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmPersEstadosFinancierosSolicitaAutoriza 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registrar / Aprobar/Anular Solicitud  de Edición de EEFF"
   ClientHeight    =   5835
   ClientLeft      =   105
   ClientTop       =   435
   ClientWidth     =   8280
   Icon            =   "frmPersEstadosFinancierosSolicitaAutoriza.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   8280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTabEdicionEF 
      Height          =   5295
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   8175
      _ExtentX        =   14420
      _ExtentY        =   9340
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Solicita Edición de EEFF"
      TabPicture(0)   =   "frmPersEstadosFinancierosSolicitaAutoriza.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label8"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label3"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label2"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSalirSol"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdLimpiarSol"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdRegSol"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "TextMotivoEdicion"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "TextUltActualizacion"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "TextFechaReg"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "TextAgencia"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "TextUsuarioSolicitante"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtNombreCliente"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).ControlCount=   15
      TabCaption(1)   =   "Aprobar / Anular Edición de EEFF"
      TabPicture(1)   =   "frmPersEstadosFinancierosSolicitaAutoriza.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label14"
      Tab(1).Control(1)=   "Label13"
      Tab(1).Control(2)=   "Label12"
      Tab(1).Control(3)=   "Label11"
      Tab(1).Control(4)=   "Label10"
      Tab(1).Control(5)=   "Label9"
      Tab(1).Control(6)=   "cmdSalirApr"
      Tab(1).Control(7)=   "cmdAnularApr"
      Tab(1).Control(8)=   "cmdRegApr"
      Tab(1).Control(9)=   "TextClienteApr"
      Tab(1).Control(10)=   "TextUsuarioSolicitanteApr"
      Tab(1).Control(11)=   "TextAgenciaApr"
      Tab(1).Control(12)=   "TextFechaRegApr"
      Tab(1).Control(13)=   "TextUltActualizacionApr"
      Tab(1).Control(14)=   "TextMotivoAprob"
      Tab(1).ControlCount=   15
      Begin VB.TextBox txtNombreCliente 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   18
         Top             =   600
         Width           =   4815
      End
      Begin VB.TextBox TextUsuarioSolicitante 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   17
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox TextAgencia 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   16
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox TextFechaReg 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   15
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox TextUltActualizacion 
         Enabled         =   0   'False
         Height          =   300
         Left            =   2160
         TabIndex        =   14
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox TextMotivoEdicion 
         Height          =   1455
         Left            =   2160
         MultiLine       =   -1  'True
         TabIndex        =   13
         Text            =   "frmPersEstadosFinancierosSolicitaAutoriza.frx":0342
         Top             =   3000
         Width           =   5775
      End
      Begin VB.CommandButton cmdRegSol 
         Caption         =   "Registrar"
         Height          =   495
         Left            =   2160
         TabIndex        =   12
         Top             =   4560
         Width           =   1935
      End
      Begin VB.CommandButton cmdLimpiarSol 
         Caption         =   "Limpiar"
         Height          =   495
         Left            =   4200
         TabIndex        =   11
         Top             =   4560
         Width           =   1935
      End
      Begin VB.CommandButton cmdSalirSol 
         Caption         =   "Salir"
         Height          =   495
         Left            =   6360
         TabIndex        =   10
         Top             =   4560
         Width           =   1575
      End
      Begin VB.TextBox TextMotivoAprob 
         Height          =   1455
         Left            =   -72840
         MultiLine       =   -1  'True
         TabIndex        =   9
         Text            =   "frmPersEstadosFinancierosSolicitaAutoriza.frx":0355
         Top             =   3000
         Width           =   5775
      End
      Begin VB.TextBox TextUltActualizacionApr 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -72840
         TabIndex        =   8
         Top             =   2520
         Width           =   2895
      End
      Begin VB.TextBox TextFechaRegApr 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -72840
         TabIndex        =   7
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox TextAgenciaApr 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -72840
         TabIndex        =   6
         Top             =   1560
         Width           =   4815
      End
      Begin VB.TextBox TextUsuarioSolicitanteApr 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -72840
         TabIndex        =   5
         Top             =   1080
         Width           =   4815
      End
      Begin VB.TextBox TextClienteApr 
         Enabled         =   0   'False
         Height          =   300
         Left            =   -72840
         TabIndex        =   4
         Top             =   600
         Width           =   4815
      End
      Begin VB.CommandButton cmdRegApr 
         Caption         =   "Aprobar"
         Height          =   495
         Left            =   -72840
         TabIndex        =   3
         Top             =   4560
         Width           =   1935
      End
      Begin VB.CommandButton cmdAnularApr 
         Caption         =   "Anular"
         Height          =   495
         Left            =   -70800
         TabIndex        =   2
         Top             =   4560
         Width           =   1935
      End
      Begin VB.CommandButton cmdSalirApr 
         Caption         =   "Salir"
         Height          =   495
         Left            =   -68640
         TabIndex        =   1
         Top             =   4560
         Width           =   1575
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   1320
         TabIndex        =   30
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Solicitante :"
         Height          =   195
         Left            =   480
         TabIndex        =   29
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label5 
         Caption         =   "Agencia :"
         Height          =   255
         Left            =   1200
         TabIndex        =   28
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Registro :"
         Height          =   255
         Left            =   720
         TabIndex        =   27
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Ultima Actualización :"
         Height          =   255
         Left            =   360
         TabIndex        =   26
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label8 
         Caption         =   "Motivo de la Edición :"
         Height          =   255
         Left            =   360
         TabIndex        =   25
         Top             =   3000
         Width           =   1815
      End
      Begin VB.Label Label9 
         Caption         =   "Aprobación/Anulación :"
         Height          =   495
         Left            =   -74760
         TabIndex        =   24
         Top             =   3120
         Width           =   1815
      End
      Begin VB.Label Label10 
         Caption         =   "Ultima Actualización :"
         Height          =   255
         Left            =   -74640
         TabIndex        =   23
         Top             =   2520
         Width           =   1575
      End
      Begin VB.Label Label11 
         Caption         =   "Fecha Registro :"
         Height          =   255
         Left            =   -74280
         TabIndex        =   22
         Top             =   2040
         Width           =   1215
      End
      Begin VB.Label Label12 
         Caption         =   "Agencia :"
         Height          =   255
         Left            =   -73800
         TabIndex        =   21
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Usuario Solicitante :"
         Height          =   195
         Left            =   -74520
         TabIndex        =   20
         Top             =   1080
         Width           =   1410
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Cliente :"
         Height          =   195
         Left            =   -73680
         TabIndex        =   19
         Top             =   600
         Width           =   570
      End
   End
   Begin MSMask.MaskEdBox txtFecEFmod 
      Height          =   300
      Left            =   2280
      TabIndex        =   31
      Top             =   0
      Width           =   1200
      _ExtentX        =   2117
      _ExtentY        =   529
      _Version        =   393216
      BackColor       =   16777215
      Enabled         =   0   'False
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Fecha de EEFF a Modificar"
      Height          =   255
      Left            =   120
      TabIndex        =   33
      Top             =   0
      Width           =   2055
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Usuario Solicitante :"
      Height          =   195
      Left            =   240
      TabIndex        =   32
      Top             =   1680
      Width           =   1410
   End
End
Attribute VB_Name = "frmPersEstadosFinancierosSolicitaAutoriza"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************
'** Nombre      : frmPersEstadosFinancierosSolicitaAutoriza                                                       *
'** Descripción : Formulario para solicitar y aprobar la edicion de los estados financieros     *
'** Referencia  : ERS051-2017                                                                   *
'** Creación    : EAAS, 20170915 09:00:00 AM                                                    *
'************************************************************************************************
Option Explicit
Dim nCodEF As Integer
Dim sUser As String
Public Sub Inicio(ByVal pnCodEF As String, ByVal psPersCod As String, ByVal psFechaEF As String, Optional pnOp As Integer = -1)
  
   
    Dim nNumEditEF As Integer
    Dim bAutoriza As Boolean
    Dim oRS As ADODB.Recordset
    Dim oRs2 As ADODB.Recordset
    Dim oDFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDFormatosEval = New COMDCredito.DCOMFormatosEval
    Set oRS = oDFormatosEval.RecuperaEEFF(pnCodEF, pnOp)
    nCodEF = pnCodEF
    If pnOp = 0 Then
        sUser = gsCodUser
        Set oRs2 = oDFormatosEval.RecuperaUsuarioSolicitanteEdicionEEFF(sUser)
        txtFecEFmod.Text = Format(psFechaEF, "dd/mm/yyyy")
        nCodEF = pnCodEF
        txtNombreCliente = oRS!cPersNombre
        TextUsuarioSolicitante = sUser & "-" & oRs2!cPersNombre
        TextAgencia = gsCodAge & "-" & oRs2!cAgeDescripcion
        TextFechaReg = gdFecSis
        SSTabEdicionEF.TabVisible(1) = False
        
    Else
        sUser = oRS!cPersUsuarioSolicitante
        Set oRs2 = oDFormatosEval.RecuperaUsuarioSolicitanteEdicionEEFF(sUser)
        txtFecEFmod = Format(psFechaEF, "dd/mm/yyyy")
        TextClienteApr = oRS!cPersNombre
        TextUsuarioSolicitanteApr = sUser & "-" & oRs2!cPersNombre
        TextAgenciaApr = gsCodAge & "-" & oRs2!cAgeDescripcion
        TextFechaRegApr = gdFecSis
        SSTabEdicionEF.TabVisible(0) = True
        Set oRs2 = oDFormatosEval.RecuperaUsuarioSolicitanteEdicionEEFF(sUser)
        txtFecEFmod.Text = Format(psFechaEF, "dd/mm/yyyy")
        nCodEF = pnCodEF
        txtNombreCliente = oRS!cPersNombre
        TextUsuarioSolicitante = sUser & "-" & oRs2!cPersNombre
        TextAgencia = gsCodAge & "-" & oRs2!cAgeDescripcion
        TextFechaReg = gdFecSis
        TextMotivoEdicion = oRS!cGlosaSolicitud
        TextMotivoEdicion.Enabled = False
        cmdRegSol.Enabled = False
        cmdLimpiarSol.Enabled = False
        cmdSalirSol.Enabled = False
        
    End If
    
    Set oDFormatosEval = Nothing
    If oRS!dFecha = "01/01/1900" Then
    nNumEditEF = oRS.RecordCount - 1
    TextUltActualizacion = psFechaEF
    TextUltActualizacionApr = psFechaEF
    Else
    TextUltActualizacion = oRS!dFecha
    TextUltActualizacionApr = oRS!dFecha
    End If
    Me.Show 1
End Sub

Private Sub cmdAnularApr_Click()
Dim sMotivo As String
    sMotivo = TextMotivoAprob.Text
    Dim lcMovNroEF As String
    lcMovNroEF = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Dim oRS As ADODB.Recordset
    Dim oDFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDFormatosEval = New COMDCredito.DCOMFormatosEval
    If MsgBox("Esta anulando la solicitud, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Set oRS = oDFormatosEval.GrabarAnulacionSolicitudEdicionEF(sMotivo, nCodEF, sUser, gsCodUser, lcMovNroEF)
        MsgBox "La solicitud de edición del EEFF ha sido anulada. ", vbInformation, "Aviso"
        Unload Me
    
End Sub

Private Sub cmdLimpiarSol_Click()
    TextMotivoEdicion.Text = ""
End Sub

Private Sub cmdRegApr_Click()
    Dim sMotivo As String
    sMotivo = TextMotivoAprob.Text
    Dim lcMovNroEF As String
    Dim oRS As ADODB.Recordset
    Dim oDFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDFormatosEval = New COMDCredito.DCOMFormatosEval
    If MsgBox("Esta aprobando la solicitud, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        lcMovNroEF = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Set oRS = oDFormatosEval.GrabarAprobacionSolicitudEdicionEF(sMotivo, nCodEF, sUser, gsCodUser, lcMovNroEF)
        MsgBox "La solicitud de edición del EEFF ha sido aprobada. ", vbInformation, "Aviso"
        Unload Me
    
End Sub

Private Sub cmdRegSol_Click()
    Dim sMotivo As String
    Dim lcMovNroEF As String
    sMotivo = TextMotivoEdicion.Text
    Dim oRS As ADODB.Recordset
    Dim oDFormatosEval As COMDCredito.DCOMFormatosEval
    Set oDFormatosEval = New COMDCredito.DCOMFormatosEval
    If MsgBox("Esta solicitando la edición del EEFF, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        lcMovNroEF = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Set oRS = oDFormatosEval.GrabarSolicitudEdicionEF(sMotivo, nCodEF, gsCodUser, lcMovNroEF)
        MsgBox "Su solicitud de edición del EEFF ha sido registrada, comuníquese con el Dpt. de Riesgos. ", vbInformation, "Aviso"
        Unload Me
    
End Sub

Private Sub cmdSalirApr_Click()
    Unload Me
End Sub

Private Sub cmdSalirSol_Click()
    Unload Me
End Sub

Private Sub TextMotivoEdicion_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras3(KeyAscii, True)
End Sub

