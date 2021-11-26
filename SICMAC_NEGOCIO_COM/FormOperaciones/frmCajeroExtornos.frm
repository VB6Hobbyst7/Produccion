VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCajeroExtornos 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   4800
   ClientLeft      =   1935
   ClientTop       =   2400
   ClientWidth     =   9630
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "Arial"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCajeroExtornos.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   9630
   ShowInTaskbar   =   0   'False
   Begin VB.Frame frmMotExtorno 
      Caption         =   "Motivos del Extorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3180
      Left            =   3480
      TabIndex        =   18
      Top             =   840
      Visible         =   0   'False
      Width           =   2845
      Begin VB.ComboBox cmbMotivos 
         BackColor       =   &H00C0FFC0&
         Height          =   330
         ItemData        =   "frmCajeroExtornos.frx":030A
         Left            =   240
         List            =   "frmCajeroExtornos.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   600
         Width           =   2415
      End
      Begin VB.TextBox txtDetExtorno 
         BackColor       =   &H00C0FFC0&
         Height          =   1230
         Left            =   240
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   20
         Top             =   1200
         Width           =   2415
      End
      Begin VB.CommandButton cmdExtContinuar 
         Caption         =   "&Continuar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   860
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label lblExtCmb 
         BackStyle       =   0  'Transparent
         Caption         =   "Motivo"
         Height          =   255
         Left            =   240
         TabIndex        =   23
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Detalles del Extorno"
         Height          =   255
         Left            =   240
         TabIndex        =   22
         Top             =   960
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Height          =   630
      Left            =   90
      TabIndex        =   9
      Top             =   315
      Width           =   9480
      Begin SICMACT.TxtBuscar txtBuscarUser 
         Height          =   330
         Left            =   5460
         TabIndex        =   15
         Top             =   195
         Width           =   1110
         _ExtentX        =   1958
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         sTitulo         =   ""
      End
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "&Procesar"
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
         Left            =   8010
         TabIndex        =   2
         Top             =   150
         Width           =   1335
      End
      Begin MSMask.MaskEdBox txtDesde 
         Height          =   315
         Left            =   795
         TabIndex        =   0
         Top             =   195
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8388608
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txthasta 
         Height          =   315
         Left            =   2655
         TabIndex        =   1
         Top             =   195
         Width           =   1050
         _ExtentX        =   1852
         _ExtentY        =   556
         _Version        =   393216
         ForeColor       =   8388608
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Usuario :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   4605
         TabIndex        =   14
         Top             =   225
         Width           =   720
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Hasta :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   2025
         TabIndex        =   11
         Top             =   240
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Desde :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   210
         Left            =   150
         TabIndex        =   10
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame FraLista 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   75
      TabIndex        =   8
      Top             =   960
      Width           =   9480
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7935
         TabIndex        =   7
         Top             =   3090
         Width           =   1350
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "&Confirmar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6660
         TabIndex        =   12
         Top             =   3090
         Width           =   1275
      End
      Begin VB.TextBox txtMovDesc 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   120
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   3120
         Width           =   6285
      End
      Begin VB.CommandButton cmdExtornar 
         Caption         =   "&Extornar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6660
         TabIndex        =   6
         Top             =   3090
         Width           =   1275
      End
      Begin VB.TextBox txtConcepto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   150
         Locked          =   -1  'True
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   4
         Top             =   2220
         Width           =   9240
      End
      Begin SICMACT.FlexEdit fgListaCG 
         Height          =   1770
         Left            =   150
         TabIndex        =   3
         Top             =   240
         Width           =   9225
         _ExtentX        =   16272
         _ExtentY        =   3122
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Fecha-Operación-Usu-Persona--Importe-nMovNro-nMoneda-cMovDesc"
         EncabezadosAnchos=   "350-1500-1800-600-3000-400-1200-0-0-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-L-C-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-2-0-0-0"
         TextArray0      =   "N°"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbOrdenaCol     =   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblGlosa 
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   165
         TabIndex        =   17
         Top             =   2895
         Width           =   855
      End
      Begin VB.Label lblDescripcion 
         Caption         =   "Descripción :"
         Height          =   240
         Left            =   180
         TabIndex        =   16
         Top             =   2010
         Width           =   1050
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000003&
         X1              =   9525
         X2              =   30
         Y1              =   2850
         Y2              =   2850
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         X1              =   10275
         X2              =   0
         Y1              =   2790
         Y2              =   2790
      End
   End
   Begin SICMACT.Usuario User 
      Left            =   0
      Top             =   0
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.Label lblTitulo 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      Caption         =   "EXTORNOS CAJERO "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   270
      Left            =   3023
      TabIndex        =   13
      Top             =   60
      Width           =   3585
   End
End
Attribute VB_Name = "frmCajeroExtornos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim lsAreaCod As String
Dim lsAgeCod As String
Dim lsCtaDebe As String
Dim lsCtaHaber As String
'MIOL 20120914, SEGUN RQ12270 ***********************
Dim loVistoElectronico As frmVistoElectronico
Dim lbVistoVal As Boolean
'END MIOL *******************************************

Dim nroBtn As Integer 'CTI3 (ferimoro) 19102018

Public Sub Inicia(ByVal sCaption As String, Optional ByVal psCodOpeExt As String = "")
Me.Caption = sCaption
Me.Show 1
End Sub

Private Sub GridFormatoMoneda()
Dim i As Long
Dim nMoneda As COMDConstantes.Moneda
For i = 1 To fgListaCG.Rows - 1
    nMoneda = CLng(fgListaCG.TextMatrix(i, 8))
    If nMoneda = COMDConstantes.gMonedaExtranjera Then
        fgListaCG.row = i
        fgListaCG.BackColorRow &HC0FFC0
    End If
Next i
End Sub
'****CTI3 (ferimoro)     09102018
Sub LimpiarOpc2()
    frmMotExtorno.Visible = False
    Me.cmbMotivos.ListIndex = -1
    Me.txtDetExtorno.Text = ""
    Frame1.Enabled = True
    FraLista.Enabled = True
    cmdExtornar.Enabled = True
End Sub
Private Sub cmdConfirmar_Click()
Dim OCon As COMNContabilidad.NCOMContFunciones
Dim lsMovNro As String, sSimbolo As String
Dim nMovNroHab As Long
Dim lnImporte As Double
Dim ldFechaMov As Date
Dim nMoneda As COMDConstantes.Moneda

Dim lsCadImp As String

Set OCon = New COMNContabilidad.NCOMContFunciones
If Len(Trim(txtMovdesc)) = 0 Then
    MsgBox "Debe ingresar una glosa válida", vbInformation, "Aviso"
    txtMovdesc.SetFocus
    Exit Sub
End If
If fgListaCG.TextMatrix(fgListaCG.row, 1) <> "" Then
    nMovNroHab = CLng(fgListaCG.TextMatrix(fgListaCG.row, 7))
    lnImporte = CCur(fgListaCG.TextMatrix(fgListaCG.row, 6))
    ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
    nMoneda = CLng(fgListaCG.TextMatrix(fgListaCG.row, 8))
Else
    MsgBox "Antes de Confirmar se debe Procesar", vbInformation, "Aviso"
    Exit Sub
End If

If MsgBox("Desea Confirmar la Habilitación respectiva??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    'MIOL 20120914, SEGUN RQ12270 ************************************
    If gsOpeCod = "901017H" Then 'MIOL 20130308, SEGUN ERS008-2013 SE CAMBIO "901017" POR "901017H"
'         Set loVistoElectronico = New frmVistoElectronico
'         lbVistoVal = loVistoElectronico.Inicio(3, gsOpeCod)
'         If lbVistoVal = False Then
'            Unload Me
'            Exit Sub
'         End If
          cmdSalir.Visible = True 'MIOL 20130308, SEGUN ERS008-2013
    End If
    
    'END MIOL ********************************************************
    lsMovNro = OCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oCajero = New COMNCajaGeneral.NCOMCajero
    If oCajero.GrabaConfHabilitaAgencia(lsMovNro, gsOpeCod, Trim(txtMovdesc.Text), nMovNroHab) = 0 Then
        Dim oContImp As COMNContabilidad.NCOMContImprimir
        Dim lbOk As Boolean
        Set oContImp = New COMNContabilidad.NCOMContImprimir
        lbOk = True
        
        lsCadImp = oContImp.ImprimeBoletahabilitacion(lblTitulo.Caption, "CONFIRMACION HAB. EN EFECTIVO", _
                     fgListaCG.TextMatrix(fgListaCG.row, 3), fgListaCG.TextMatrix(fgListaCG.row, 4), gsCodUser, txtBuscarUser.psDescripcion, nMoneda, gsOpeCod, _
                     lnImporte, gsNomAge, lsMovNro, sLpt, gsCodCMAC, gbImpTMU)
        Do While lbOk
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsCadImp & Chr$(12)
                Print #nFicSal, ""
            Close #nFicSal
            If MsgBox("Desea Reimprimir Boleta de Operación??", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbOk = False
            End If
        Loop
        Set oContImp = Nothing
        fgListaCG.EliminaFila fgListaCG.row
        If Not (fgListaCG.TextMatrix(1, 1) = "" And fgListaCG.Rows = 2) Then
            If MsgBox("Desea realizar otra Confirmación ??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                FraLista.Enabled = True
                fgListaCG.SetFocus
                txtMovdesc = ""
                txtConcepto = ""
            Else
                Unload Me
            End If
        Else
            Unload Me
        End If
    End If
    
End If

Set oCajero = Nothing
Set OCon = Nothing
End Sub
' VAPA 20161024
Private Sub CargarUserMatrix()
Dim oGen As COMDConstSistema.DCOMGeneral
Dim rs As ADODB.Recordset
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Set rs = New ADODB.Recordset
Set oGen = New COMDConstSistema.DCOMGeneral
If gsOpeCod = "902031" Then
   txtBuscarUser.rs = oGen.GetUserAgenciaCierre(gdFecSis)
Else
    txtBuscarUser.rs = oGen.GetUserAreaAgencia(User.cAreaCodAct, gsCodAge)
End If
End Sub
'END VAPA
Private Sub cmdExtContinuar_Click()
Dim OCon As COMNContabilidad.NCOMContFunciones
Dim oCaja As COMNCajaGeneral.NCOMCajaGeneral
Dim lnMovNro As Long
Dim lsMovNroExt As String
Dim lnImporte As Double
Dim ldFechaMov As Date
Dim lsCtaCod As String
Dim oImp As COMNContabilidad.NCOMContImprimir
Dim lsTexto As String
Dim lbReimp As Boolean
Dim nMoneda As Moneda

Set OCon = New COMNContabilidad.NCOMContFunciones
Set oCaja = New COMNCajaGeneral.NCOMCajaGeneral

Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Dim lsUserDest As String
Dim lnNombreDest As String

Dim lsCadImp As String
Dim lsOpeCodExtornoCap As String 'CTI6 ERS0112020
Dim lsCtaAhoExt As String 'CTI6 ERS0112020
Dim lnMontoAhorroExt As Double 'CTI6 ERS0112020
Dim lsOperacionDescExt As String 'CTI6 ERS0112020
Dim lnITFAhoExt As Double 'CTI6 ERS0112020
Dim lnMovNroAExt As Long 'CTI6 ERS0112020
Dim lsClienteExt As String 'CTI6 ERS0112020

'*** PEAC 20081002
Dim lbResultadoVisto As Boolean
Dim sPersVistoCod  As String
Dim sPersVistoCom As String
Dim loVistoElectronico As frmVistoElectronico
Dim oCajero As COMNCajaGeneral.NCOMCajero 'madm 20110202
Set oCajero = New COMNCajaGeneral.NCOMCajero 'madm 20110202
Set loVistoElectronico = New frmVistoElectronico

'***CTI3 (FERIMORO)   02102018
Dim DatosExtorna(1) As String

If fgListaCG.TextMatrix(1, 0) = "" Then Exit Sub

'***************CTI3  (ferimoro)  01102018
If cmbMotivos.ListIndex = -1 Or Len(txtDetExtorno.Text) <= 0 Then
    MsgBox "Debe ingresar el motivo y/o detalle del Extorno", vbInformation, "Aviso"
    Exit Sub
End If
'If Len(Trim(txtMovDesc)) = 0 Then
'    MsgBox "Debe ingresar una glosa válida.", vbInformation, "Aviso"
'    txtMovDesc.SetFocus
'    Exit Sub
'End If

        '***CTI3 (ferimoro)    02102018
        frmMotExtorno.Visible = False
        DatosExtorna(0) = cmbMotivos.Text
        DatosExtorna(1) = txtDetExtorno.Text

        lnMovNro = fgListaCG.TextMatrix(fgListaCG.row, 7)
        lnImporte = CDbl(fgListaCG.TextMatrix(fgListaCG.row, 6))
        ldFechaMov = CDate(fgListaCG.TextMatrix(fgListaCG.row, 1))
        lsCtaCod = fgListaCG.TextMatrix(fgListaCG.row, 4)
        nMoneda = CLng(fgListaCG.TextMatrix(fgListaCG.row, 8))
        
    '*** PEAC 20081001 - visto electronico ******************************************************
    '*** en estos extornos de operaciones pedirá visto electrónico
    
    Select Case gsOpeCod
         'madm 20110202
         Case 901029
             If oCajero.YaRealizoDevBilletaje(gsCodUser, gdFecSis, gsCodAge) Then
                MsgBox "Ud. ha realizado la operación de registro de efectivo, esta operación no esta disponible después del registro de efectivo", vbInformation, "Aviso"
                Unload Me
                Exit Sub
            End If
         'end madm
         Case "901025", "901028", "909002", "909003", gCompraMECargoCtaAhorro, gVentaMECargoCtaAhorro 'CTI6 ERS0112020

            ' *** RIRO SEGUN TI-ERS108-2013 ***
                Dim nMovNroOperacion As Long
                nMovNroOperacion = 0
                If fgListaCG.row >= 1 And Len(Trim(fgListaCG.TextMatrix(fgListaCG.row, 7))) > 0 Then
                    nMovNroOperacion = lnMovNro
                End If
            ' *** FIN RIRO ***

             lbResultadoVisto = loVistoElectronico.Inicio(3, gsOpeCod, , , nMovNroOperacion)
             If Not lbResultadoVisto Then
                 Unload Me
                 Exit Sub
             End If
    End Select

    '*** FIN PEAC ************************************************************
        
If MsgBox("Desea Realizar el Extorno respectivo??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    If gdFecSis <> Format(ldFechaMov, "dd/mm/yyyy") Then
        If MsgBox("Se va a Realizar el Extorno de Movimientos de dias anteriores" & vbCrLf & " Desea Proseguir??", vbYesNo + vbQuestion) = vbNo Then Exit Sub
    End If
    lsMovNroExt = OCon.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If gsOpeCod = COMDConstantes.gServExtCobFideicomiso Then   'Extorno Cobranza de Fideicomiso
        Dim oGraba As COMNCaptaServicios.NCOMCaptaFideicomiso
        Set oGraba = New COMNCaptaServicios.NCOMCaptaFideicomiso
        If oGraba.ExtornoPagoFideicomiso(lnMovNro, lsMovNroExt, lsCtaCod, lnImporte) Then
            Set OCon = Nothing
                Set oImp = New COMNContabilidad.NCOMContImprimir
                lsCadImp = oImp.ImprimeBoletaExtornos(lblTitulo.Caption, txtMovdesc.Text, gsOpeCod, lnImporte, fgListaCG.TextMatrix(fgListaCG.row, 3), _
                            fgListaCG.TextMatrix(fgListaCG.row, 4), gsNomAge, lsMovNroExt, gsInstCmac, nMoneda, sLpt, 0, gsCodCMAC)
          
                Do
                    nFicSal = FreeFile
                    Open sLpt For Output As nFicSal
                        Print #nFicSal, lsCadImp & Chr$(12)
                        Print #nFicSal, ""
                    Close #nFicSal
                    
                Loop While MsgBox("Desea Reimprimir Boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbYes
                
                Set oImp = Nothing
                
            If MsgBox("Desea realizar otro extorno de Movimiento??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                fgListaCG.EliminaFila fgListaCG.row
                fgListaCG.Clear
                FraLista.Enabled = True
                fgListaCG.SetFocus
                txtMovdesc = ""
                txtConcepto = ""
            Else
                Unload Me
            End If
        
        End If
        
        Set oGraba = Nothing
        
    End If
    'cti3
    Dim loContFunct As COMNContabilidad.NCOMContFunciones 'CTI6 ERS0112020
    Set loContFunct = New COMNContabilidad.NCOMContFunciones 'CTI6 ERS0112020
    Dim lsMovNroCapExtorno As String 'CTI6 ERS0112020
    lsMovNroCapExtorno = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser) 'CTI6 ERS0112020
    Set loContFunct = Nothing 'CTI6 ERS0112020
    ' If oCaja.GrabaExtornoMov(gdFecSis, ldFechaMov, lsMovNroExt, lnMovNro, gsOpeCod, Trim(txtMovdesc.Text), lnImporte) = 0 Then
    If oCaja.GrabaExtornoMov(gdFecSis, ldFechaMov, lsMovNroExt, lnMovNro, gsOpeCod, Trim(txtMovdesc.Text), lnImporte, DatosExtorna) = 0 Then
        Set OCon = Nothing
        Set oImp = New COMNContabilidad.NCOMContImprimir
        Set rs = oCaja.GetCajeroDestino(lnMovNro)
        
        If rs Is Nothing Then
            lsUserDest = ""
            lnNombreDest = ""
        Else
            If rs.EOF And rs.BOF Then
                lsUserDest = ""
                lnNombreDest = ""
            Else
                lsUserDest = rs.Fields(0)
                lnNombreDest = rs.Fields(1)
                lsUserDest = "Para : " & Left(lsUserDest & "-" & lnNombreDest, 35)
            End If
        End If
        
        'APRI20180201 INC20180105004
         Dim oCapAut As COMDConstSistema.DCOMTCEspPermiso
         Set oCapAut = New COMDConstSistema.DCOMTCEspPermiso
         If oCapAut.ObtieneOpeTipoCambioEspecialCliente(lnMovNro) Then
            Call oCapAut.OpeExtTCEspecialCliente(lnMovNro)
         End If
        'END APRI
        
        
        lbReimp = True
        If gsOpeCod = COMDConstantes.gServExtCobEdelnor Then
           lsCadImp = oImp.ImprimeBoletaExtornos(lblTitulo.Caption, fgListaCG.TextMatrix(fgListaCG.row, 9), gsOpeCod, lnImporte, Left(fgListaCG.TextMatrix(fgListaCG.row, 3) & "-" & fgListaCG.TextMatrix(fgListaCG.row, 4), 26), _
                lsUserDest, gsNomAge, lsMovNroExt, gsInstCmac, nMoneda, sLpt, lnMovNro, gsCodCMAC)
        Else
           lsCadImp = oImp.ImprimeBoletaExtornos(lblTitulo.Caption, Trim(txtMovdesc.Text), gsOpeCod, lnImporte, Left(fgListaCG.TextMatrix(fgListaCG.row, 3) & "-" & fgListaCG.TextMatrix(fgListaCG.row, 4), 26), _
                lsUserDest, gsNomAge, lsMovNroExt, gsInstCmac, nMoneda, sLpt, lnMovNro, gsCodCMAC)
        End If
        Do
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsCadImp & Chr$(12)
                Print #nFicSal, ""
            Close #nFicSal
        Loop While MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbYes
        
        'CTI6 ERS0112020
        If lsOpeCodExtornoCap <> "" Then
            If (lsOpeCodExtornoCap = gAhoCargoCompra Or lsOpeCodExtornoCap = gAhoCargoVenta) And Len(lsCadImp) > 0 Then
                Dim ClsMov As COMDMov.DCOMMov, sCodUserBusExt As String, sMovNroBusExt As String
                Set ClsMov = New COMDMov.DCOMMov
                sMovNroBusExt = "": sCodUserBusExt = ""
                sMovNroBusExt = lsMovNroExt
                sCodUserBusExt = Right(sMovNroBusExt, 4)
                        
            
                Dim lsFechaHoraGrabExt As String
                lsFechaHoraGrabExt = fgFechaHoraGrab(lsMovNroCapExtorno)
                Set oImp = New COMNContabilidad.NCOMContImprimir
                lsClienteExt = fgListaCG.TextMatrix(fgListaCG.row, 4)
                lsCadImp = oImp.nPrintReciboExtorCargoCta(gsNomAge, lsFechaHoraGrabExt, "", lsCtaAhoExt, lnITFAhoExt, _
                lsClienteExt, lsOperacionDescExt, lnMontoAhorroExt, 0, lnMovNroAExt, gsCodUser, "", "", sCodUserBusExt, gImpresora, gbImpTMU)
                 Do
                    nFicSal = FreeFile
                    Open sLpt For Output As nFicSal
                        Print #nFicSal, lsCadImp & Chr$(12)
                        Print #nFicSal, ""
                    Close #nFicSal
                Loop While MsgBox("Reimprimir Recibo de Extorno del Abono a Cuenta") = vbYes
            End If
        End If
        'END
        
        Set oImp = Nothing
        If MsgBox("Desea realizar otro extorno de Movimiento??", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            fgListaCG.EliminaFila fgListaCG.row
            fgListaCG.Clear
            FraLista.Enabled = True
            fgListaCG.SetFocus
            txtMovdesc = ""
            txtConcepto = ""
        Else
            Unload Me
        End If
    End If
    Call CargarUserMatrix 'VAPA 20163110
    '*** PEAC 20081002
        loVistoElectronico.RegistraVistoElectronico (lnMovNro)
    '*** FIN PEAC
           
End If
Set oCajero = Nothing
Set oCaja = Nothing
End Sub

Private Sub cmdProcesar_Click()
Dim rs As ADODB.Recordset
Dim sUser As String
Dim dDesde As Date, dHasta As Date
Set rs = New ADODB.Recordset
Dim lsOperacion As String
Dim lsAgenciaCierre As String
Dim oCierre As COMDConstSistema.DCOMGeneral 'vapa20161024
Set oCierre = New COMDConstSistema.DCOMGeneral 'vapa
dDesde = CDate(txtDesde.Text)
dHasta = CDate(txtHasta.Text)

If ValFecha(txtDesde) = False Then Exit Sub
If ValFecha(txtHasta) = False Then Exit Sub

If CDate(dDesde) > CDate(dHasta) Then
    MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
    Exit Sub
End If
'MIOL 20130308, SEGUN ERS008-2013 *****
If gsOpeCod = "901017H" Then
    gsOpeCod = "901017"
End If
'END MIOL *****************************
lsOperacion = gsOpeCod
sUser = txtBuscarUser.Text
'vapa20161024
If sUser = "" And gsOpeCod = "902031" Then
MsgBox "Debe Elegir un Usuario para realizar el extorno", vbInformation, "Aviso"
Exit Sub
End If
lsAgenciaCierre = oCierre.GetAgenciaxUser(sUser)
'End Vapa
Set oCajero = New COMNCajaGeneral.NCOMCajero
Select Case gsOpeCod
    Case COMDConstSistema.gOpeBoveAgeExtHabCajero
        lsOperacion = COMDConstSistema.gOpeBoveAgeHabCajero
        Set rs = oCajero.GetMovHabBovAgeCajero(lsOperacion, lsAreaCod, lsAgeCod, CDate(dDesde), CDate(dHasta), sUser)
    Case COMDConstSistema.gOpeHabCajConfHabBovAge
        lsOperacion = COMDConstSistema.gOpeBoveAgeHabCajero & "','" & COMDConstSistema.gOpeHabCajTransfEfectCajeros
        Set rs = oCajero.GetMovHabBovAgeCajero(lsOperacion, lsAreaCod, lsAgeCod, CDate(dDesde), CDate(dHasta), sUser)
    Case COMDConstSistema.gOpeBoveAgeConfDevCaj
        lsOperacion = COMDConstSistema.gOpeHabCajDevABove
        Set rs = oCajero.GetMovHabCajDevABove(lsOperacion, lsAreaCod, lsAgeCod, CDate(dDesde), CDate(dHasta))
    Case COMDConstSistema.gOpeHabCajExtDevABove
        lsOperacion = COMDConstSistema.gOpeHabCajDevABove
        Set rs = oCajero.GetDevCajero(lsOperacion, CDate(dDesde), CDate(dHasta), sUser)
    Case COMDConstSistema.gOpeHabCajExtDevBilletaje
        lsOperacion = COMDConstSistema.gOpeHabCajDevBilletaje
        Set rs = oCajero.GetDevCajero(lsOperacion, CDate(dDesde), CDate(dHasta), sUser)
    Case COMDConstSistema.gOpeHabCajExtConfHabBovAge
        lsOperacion = COMDConstSistema.gOpeHabCajConfHabBovAge
        Set rs = oCajero.GetConfHabBovAgencia(lsOperacion, CDate(dDesde), CDate(dHasta), lsAreaCod, lsAgeCod, sUser)
    Case COMDConstSistema.gOpeHabCajExtTransfEfectCajeros
        lsOperacion = COMDConstSistema.gOpeHabCajTransfEfectCajeros
        Set rs = oCajero.GetDevCajero(lsOperacion, CDate(dDesde), CDate(dHasta), sUser)
    Case COMDConstSistema.gOpeCajeroMEExtCompra
        lsOperacion = COMDConstSistema.gOpeCajeroMECompra
        Set rs = oCajero.GetCompraVenta(lsOperacion, CDate(dDesde), CDate(dHasta), sUser)
    'MADM 20100201
    Case COMDConstSistema.gOpeBoveAgeExtPreCuadre
        lsOperacion = COMDConstSistema.gOpeHabCajDevBilletaje
        Set rs = oCajero.GetDevCajeroPreCuadre(lsOperacion, CDate(dDesde), CDate(dHasta), sUser)
    'END MADM
    'MADM 20110926
    Case COMDConstSistema.gOpeBoveAgeExtConfDevCaj
        lsOperacion = COMDConstSistema.gOpeBoveAgeExtConfDevCaj
        Set rs = oCajero.GetDevCajeroConfDevolucionBove(lsOperacion, CDate(dDesde), CDate(dHasta), sUser)
    'END MADM
    Case COMDConstSistema.gOpeCajeroMEExtVenta
        lsOperacion = COMDConstSistema.gOpeCajeroMEVenta
        Set rs = oCajero.GetCompraVenta(lsOperacion, CDate(dDesde), CDate(dHasta), sUser)
    Case COMDConstSistema.gOpeCajeroMEExtCompraEsp
        lsOperacion = COMDConstSistema.gOpeCajeroMECompraEsp
        Set rs = oCajero.GetCompraVenta(lsOperacion, CDate(dDesde), CDate(dHasta), sUser)
    Case COMDConstSistema.gOpeCajeroMEExtVentaEsp
        lsOperacion = COMDConstSistema.gOpeCajeroMEVentaEsp
        Set rs = oCajero.GetCompraVenta(lsOperacion, CDate(dDesde), CDate(dHasta), sUser)
    Case gServExtCobHidrandina, gServExtCobSedalib, gServExtCobEdelnor, gServExtCobSATTInfraccion, _
            gServExtCobSATTReciboDerecho, gServExtCobSATTReciboDerechoOficEsp
        Select Case gsOpeCod
            Case COMDConstantes.gServExtCobHidrandina
                lsOperacion = COMDConstantes.gServCobHidrandina
            Case COMDConstantes.gServExtCobSedalib
                lsOperacion = COMDConstantes.gServCobSedalib
            Case COMDConstantes.gServExtCobEdelnor
                lsOperacion = COMDConstantes.gServCobEdelnor
            Case COMDConstantes.gServExtCobSATTInfraccion
                lsOperacion = COMDConstantes.gServCobSATTInfraccion
            Case COMDConstantes.gServExtCobSATTReciboDerecho
                lsOperacion = COMDConstantes.gServCobSATTReciboDerecho
            Case COMDConstantes.gServExtCobSATTReciboDerechoOficEsp
                lsOperacion = COMDConstantes.gServCobSATTReciboDerechoOficEsp
        End Select
        Set rs = oCajero.GetServicios(lsOperacion, CDate(dDesde), CDate(dHasta), Trim(sUser), gsCodAge)
    Case COMDConstSistema.gOpeHabCajExtIngEfectRegulaFalt
        lsOperacion = COMDConstSistema.gOpeHabCajIngEfectRegulaFalt & "','" & COMDConstSistema.gOpeHabCajDevClienteRegulaSob & "','" & COMDConstSistema.gOpeHabCajIngRegulaSob
        Set rs = oCajero.GetIngEfectFalt(lsOperacion, CDate(dDesde), CDate(dHasta), Trim(sUser), gsCodAge)
    'EXTORNOS DE SERVICIOS DE LA CAJA METROPOLITANA
    Case COMDConstantes.gServExtCobFideicomiso
        lsOperacion = COMDConstantes.gServCobFideicomiso
        Set rs = oCajero.GetMovFideicomiso(lsOperacion, CDate(dDesde), CDate(dHasta), Trim(sUser), gsCodAge)
    Case COMDConstantes.gServExtCobFoncodes
        lsOperacion = COMDConstantes.gServCobFoncodes
        Set rs = oCajero.GetMovFoncodes(lsOperacion, CDate(dDesde), CDate(dHasta), Trim(sUser), gsCodAge)
    Case COMDConstantes.gServExtCobPlanBici
        lsOperacion = COMDConstantes.gServCobPlanBici
        Set rs = oCajero.GetMovPlanBici(lsOperacion, CDate(dDesde), CDate(dHasta), Trim(sUser), gsCodAge)
    Case COMDConstSistema.gOpeBoveAgeExtRegEfect 'DAOR 20080204
        lsOperacion = COMDConstSistema.gOpeBoveAgeRegEfect
        Set rs = oCajero.ObtenerRegistrosDeEfectivo(lsOperacion, CDate(dDesde), CDate(dHasta), txtBuscarUser.Text, lsAgeCod)
    Case COMDConstSistema.gOpeBoveAgeExtRegSobFalt
        lsOperacion = COMDConstSistema.gOpeBoveAgeRegSobrante & "','" & gOpeBoveAgeRegFaltante
        Set rs = oCajero.GetMovRegSobFalt(lsOperacion, lsAgeCod, CDate(dDesde), CDate(dHasta))
    Case COMDConstSistema.gOpeHabCajExtRegSobFalt
        lsOperacion = COMDConstSistema.gOpeHabCajRegSobrante & "','" & COMDConstSistema.gOpeHabCajRegFaltante
        Set rs = oCajero.GetMovRegSobFalt(lsOperacion, lsAgeCod, CDate(dDesde), CDate(dHasta), sUser)
    'extorno de cierre de agencia
    Case COMDConstSistema.gOpeCajaExtCierreAgenica
        lsOperacion = COMDConstSistema.gOpeCajaCierreAgencia
       ' Set rs = oCajero.GetCierreAgencias(gdFecSis, lsOperacion, "", gsCodAge)    'comentado por VAPA 20161024
        Set rs = oCajero.GetCierreAgencias(gdFecSis, lsOperacion, "", lsAgenciaCierre) ' VAPA 20161024
End Select
fgListaCG.Clear
fgListaCG.FormaCabecera
fgListaCG.Rows = 2
If Not rs.EOF And Not rs.BOF Then
    Set fgListaCG.Recordset = rs
    fgListaCG.FormatoPersNom 4
    If fgListaCG.Enabled And fgListaCG.Visible Then fgListaCG.SetFocus
    GridFormatoMoneda
    If gsOpeCod = COMDConstantes.gServExtCobFideicomiso Then
        fgListaCG.ColWidth(5) = 0
    End If
Else
    MsgBox "Datos no encontrados", vbInformation, "Aviso"
End If
rs.Close
Set rs = Nothing
Set oCajero = Nothing
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgListaCG_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
  If fgListaCG.TextMatrix(1, 0) <> "" Then
  
    KeyAscii = 0
    If cmdExtornar.Visible Then cmdExtornar.Enabled = True: cmdExtornar.SetFocus
    If cmdConfirmar.Visible Then cmdConfirmar.Enabled = True: cmdConfirmar.SetFocus

  End If
    'txtMovDesc.SetFocus
End If
End Sub

Private Sub fgListaCG_RowColChange()
    txtConcepto = fgListaCG.TextMatrix(fgListaCG.row, 9)
End Sub

Private Sub Form_Load()

Dim oOpe As New COMDConstSistema.DCOMOperacion
Dim oGen As COMDConstSistema.DCOMGeneral

Dim rs As ADODB.Recordset
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Set rs = New ADODB.Recordset
Set oOpe = New COMDConstSistema.DCOMOperacion
Set oGen = New COMDConstSistema.DCOMGeneral

txtDesde = gdFecSis
txtHasta = gdFecSis
User.Inicio gsCodUser
Me.txtBuscarUser.psRaiz = "USUARIOS"


''txtBuscarUser.rs = oGen.GetUserAreaAgencia(User.cAreaCodAct, gsCodAge)
Call CargaControles
Call CargarUserMatrix ' VAPA 20161024


lblTitulo = Trim(Replace(gsOpeDesc, "-", "", 1, , vbTextCompare))
cmdExtornar.Visible = False
cmdConfirmar.Visible = False
Select Case gsOpeCod
    'MADM 20110926 - 901035
    Case 901035
        'CTI3
        txtMovdesc.Enabled = False
        txtMovdesc.BackColor = &H80000004
        '**************************************
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
        lsAreaCod = User.AreaCod
        lsAgeCod = gsCodAge
'        txtBuscarUser = gsCodUser
'        txtBuscarUser.Enabled = True
    'END MADM
    Case gOpeBoveAgeExtHabCajero
        'CTI3
        txtMovdesc.Enabled = False
        txtMovdesc.BackColor = &H80000004
        '******************************
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
        lsAreaCod = User.AreaCod
        lsAgeCod = gsCodAge
    Case gOpeHabCajConfHabBovAge
        cmdConfirmar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
        txtBuscarUser = gsCodUser
        txtBuscarUser.Enabled = False
        lsAreaCod = User.AreaCod
        lsAgeCod = gsCodAge
    'MIOL 20130308, SEGUN ERS008-2013 ***************
        Case "901017H"
        cmdConfirmar.Visible = True
        cmdSalir.Visible = False
        txtDesde.Enabled = False
        txtHasta.Enabled = False
        txtBuscarUser = gsCodUser
        txtBuscarUser.Enabled = False
        lsAreaCod = User.AreaCod
        lsAgeCod = gsCodAge
    'END MIOL ***************************************
    Case gOpeBoveAgeExtRegEfect 'DAOR 20080204
        'CTI3
        txtMovdesc.Enabled = False
        txtMovdesc.BackColor = &H80000004
        '***********************************
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
        txtBuscarUser = gsUsuarioBOVEDA
        txtBuscarUser.Enabled = False
        lsAreaCod = User.AreaCod
        lsAgeCod = gsCodAge
    Case gOpeBoveAgeExtRegSobFalt
        'CTI3
        txtMovdesc.Enabled = False
        txtMovdesc.BackColor = &H80000004
        '***********************************
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
        txtBuscarUser = gsUsuarioBOVEDA
        txtBuscarUser.Enabled = False
        lsAreaCod = User.AreaCod
        lsAgeCod = gsCodAge
    Case gOpeBoveAgeConfDevCaj
        cmdConfirmar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
        txtBuscarUser = gsUsuarioBOVEDA
        txtBuscarUser.Enabled = False
        lsAreaCod = User.AreaCod
        lsAgeCod = gsCodAge
    Case gOpeHabCajExtConfHabBovAge, gOpeHabCajExtRegSobFalt
        'CTI3
        txtMovdesc.Enabled = False
        txtMovdesc.BackColor = &H80000004
        '***********************************
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
        lsAreaCod = User.AreaCod
        lsAgeCod = gsCodAge
    Case gOpeHabCajExtDevABove, gOpeHabCajExtDevBilletaje, 901029
        'CTI3
        txtMovdesc.Enabled = False
        txtMovdesc.BackColor = &H80000004
        '***********************************
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
    Case gOpeHabCajExtTransfEfectCajeros
        'CTI3
        txtMovdesc.Enabled = False
        txtMovdesc.BackColor = &H80000004
        '***********************************
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
    Case gOpeCajeroMEExtCompra, gOpeCajeroMEExtVenta, gOpeCajeroMEExtCompraEsp, gOpeCajeroMEExtVentaEsp
        'CTI3
        txtMovdesc.Enabled = False
        txtMovdesc.BackColor = &H80000004
        '***********************************
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
    Case gServExtCobHidrandina, gServExtCobSedalib, gServExtCobEdelnor, gServExtCobSATTInfraccion, _
            gServExtCobSATTReciboDerecho, gServExtCobSATTReciboDerechoOficEsp
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
        
   'FIDEICOMISO, FONCODES, PLAN BICI
    Case gServExtCobFideicomiso, gServExtCobFoncodes, gServExtCobPlanBici
        fgListaCG.EncabezadosNombres = "N° -Fecha - Operación - Usuario - Documento/Referencia -  - Importe - cMovNro - cOpeCod"
        'CTI3
        txtMovdesc.Enabled = False
        txtMovdesc.BackColor = &H80000004
        '***********************************
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
        fgListaCG.ColWidth(5) = 0
        
    Case gOpeHabCajExtIngEfectRegulaFalt
        fgListaCG.EncabezadosNombres = "N° -Fecha - Operación - Usuario - Concepto - Importe - cMovNro - cOpeCod"
        'Me.lblTitulo = gsOpeDescHijo
        'CTI3
        txtMovdesc.Enabled = False
        txtMovdesc.BackColor = &H80000004
        '***********************************
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
    Case gOpeCajaExtCierreAgenica
        'CTI3
        txtMovdesc.Enabled = False
        txtMovdesc.BackColor = &H80000004
        '***********************************
        cmdExtornar.Visible = True
        txtDesde.Enabled = False
        txtHasta.Enabled = False
        txtBuscarUser.Enabled = True
End Select
Set oOpe = Nothing
Set oGen = Nothing
Me.Caption = gsOpeCod & " - " & lblTitulo
End Sub

Private Sub txtConcepto_KeyDown(KeyCode As Integer, Shift As Integer)
    'KeyAscii = fgIntfMayusculas(KeyAscii)
End Sub

Private Sub txtDesde_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    If ValFecha(txtDesde) = False Then Exit Sub
    If CDate(txtDesde) > CDate(txtHasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
    txtHasta.SetFocus
End If
End Sub
Private Sub txtDetExtorno_KeyPress(KeyAscii As Integer)
If KeyAscii = 39 Then KeyAscii = 0: Exit Sub
'If KeyAscii = 13 Then SendKeys "{TAB}": Exit Sub
End Sub

Private Sub txtHasta_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    If ValFecha(txtHasta) = False Then Exit Sub
    If CDate(txtDesde) > CDate(txtHasta) Then
        MsgBox "Fecha Inicial no puede ser mayor que la Final", vbInformation, "Aviso"
        Exit Sub
    End If
    cmdProcesar.SetFocus
End If
End Sub
Private Sub txtMovDesc_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    KeyAscii = 0
    If cmdExtornar.Visible Then cmdExtornar.SetFocus
    If cmdConfirmar.Visible Then cmdConfirmar.SetFocus
End If
End Sub
Private Sub cmdExtornar_Click()
nroBtn = 2

If fgListaCG.TextMatrix(fgListaCG.row, 1) = "" Then
 MsgBox "Antes de Confirmar se debe Procesar", vbInformation, "Aviso"
 Exit Sub
End If

'******CTI3 (ferimoro) 27092018
 frmMotExtorno.Visible = True
 Frame1.Enabled = False
 FraLista.Enabled = False
 cmdExtornar.Enabled = False
 cmbMotivos.SetFocus
 cmbMotivos.ListIndex = -1
 txtDetExtorno.Text = ""
'******************************

End Sub
'******CTI3 (ferimoro) 18102018
Private Sub CargaControles()
Dim oCons As COMDConstantes.DCOMConstantes
Dim R As ADODB.Recordset

Set oCons = New COMDConstantes.DCOMConstantes

Set R = oCons.ObtenerConstanteExtornoMotivo

Set oCons = Nothing
Call Llenar_Combo_MotivoExtorno(R, cmbMotivos)

End Sub
