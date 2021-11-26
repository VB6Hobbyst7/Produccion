VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmSegSepelioCobroPrima 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cobro Prima - Seguro Sepelio"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6135
   Icon            =   "frmSegSepelioCobroPrima.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame4 
      Height          =   615
      Left            =   240
      TabIndex        =   9
      Top             =   2400
      Width           =   5655
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   300
         Left            =   4440
         TabIndex        =   15
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   120
         TabIndex        =   13
         Top             =   210
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   300
         Left            =   1210
         TabIndex        =   12
         Top             =   210
         Width           =   1095
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   " Pago Efectivo "
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   3840
      TabIndex        =   8
      Top             =   1800
      Width           =   2055
      Begin SICMACT.EditMoney lblMonto 
         Height          =   255
         Left            =   600
         TabIndex        =   14
         Top             =   240
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.Label lblMoneda 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Certificado "
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   240
      TabIndex        =   7
      Top             =   1800
      Width           =   2535
      Begin VB.Label lblNumCerti 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Asegurado "
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   600
      Width           =   5655
      Begin VB.TextBox TxtBuscarPersona 
         Enabled         =   0   'False
         Height          =   285
         Left            =   960
         TabIndex        =   17
         Top             =   230
         Width           =   2055
      End
      Begin VB.CommandButton cmbBuscar 
         Caption         =   "Buscar"
         Height          =   300
         Left            =   3070
         TabIndex        =   16
         Top             =   200
         Width           =   735
      End
      Begin VB.Label lblDOI 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   795
         Width           =   1575
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   525
         Width           =   4575
      End
      Begin VB.Label Label3 
         Caption         =   "DOI:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   820
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   525
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5318
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Seguro Vida"
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmSegSepelioCobroPrima"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lsFecAfiliacion As String
Dim lnEstado As Integer
Dim lnTipo As Integer 'APRI20180209 ERS028-2017
Dim lbPrimaAnual As Boolean 'APRI20181023 ERS071-2018

Public Function Inicia(ByVal psPersCod As String, Optional ByVal nTipo As Integer = 0) As Boolean
'COMENTADO APRI20171024
'    Inicia = False
'    If VerificaAfiliacionDesactivada(psPersCod) Then
'        Call CargarDatos(psPersCod)
'        Inicia = True
'        TxtBuscarPersona.Enabled = False
'        cmdCancelar.Enabled = False
'        lblMonto.Enabled = False
'        If MsgBox("El cliente tiene un seguro de sepelio en estado Inactivo. ¿Desea volver a Activarlo?", vbYesNo + vbQuestion, "Alerta") = vbYes Then
'            Me.Show 1
'        End If
'    End If
    'APRI20171024 ERS028-2017
        lnTipo = nTipo
        Call CargarDatos(psPersCod)
        TxtBuscarPersona.Enabled = False
        If lnTipo = 1 Then
            cmdCancelar.Enabled = False
        Else
            cmdCancelar.Enabled = True
        End If
        lblMonto.Enabled = False
        
        cmdAceptar.Enabled = True
        cmbBuscar.Enabled = False
        
        Me.Show 1
    'END APRI

End Function

Private Sub cmbBuscar_Click()
    Dim oSegSep As New COMNCaptaGenerales.NCOMSeguros
    Dim ClsPersona As New COMDPersona.DCOMPersonas
    Dim oPersona As COMDPersona.UCOMPersona
    Dim rsDatos As New ADODB.Recordset
    Dim R As New ADODB.Recordset
    Dim lsPersCod As String
    lblMonto.Enabled = False
    Set ClsPersona = New COMDPersona.DCOMPersonas
    
    'lsPersCod = frmBuscaPersona.Inicio 'TxtBuscarPersona.psCodigoPersona
    Set oPersona = frmBuscaPersona.Inicio 'TxtBuscarPersona.psCodigoPersona
    
    If Not oPersona Is Nothing Then
        lsPersCod = oPersona.sPersCod
        TxtBuscarPersona.Text = lsPersCod
    End If
    If lsPersCod <> "" Then
        Set ClsPersona = New COMDPersona.DCOMPersonas
        Set R = ClsPersona.BuscaCliente(lsPersCod, BusquedaCodigo)
        Set rsDatos = oSegSep.ObtieneDatosPersSegSepelio(lsPersCod)
        If Not (rsDatos.EOF And rsDatos.BOF) Then
        
            If oSegSep.SepelioVerificaPagoMes(Format(gdFecSis, "yyyyMMdd"), lsPersCod) = True Then
            'If rsDatos!nSegEstado = 1 Then 'RECO20160413
                'MsgBox "Cliente ya registra pago del mes. Proxima fecha de pago :" & DateAdd("d", 30, rsDatos!dFecAfiliacion), vbInformation, "ALERTA SICMAC MAYNAS" 'RECO20160413
                MsgBox "Cliente ya registra pago del mes", vbInformation, "Alerta SICMACM"
                Call LimpiaFormulario
                Exit Sub
            End If
            lblDOI.Caption = R!cPersIDnroDNI
            lblNombre.Caption = R!cPersNombre
            lblMonto.Text = Format(rsDatos!nMontoPrima, gsFormatoNumeroView)
            lblMoneda.Caption = "S/."
            lblNumCerti.Caption = rsDatos!cNumCertificado
            lsFecAfiliacion = rsDatos!dFecAfiliacion
            lnEstado = rsDatos!nSegEstado
            lbPrimaAnual = rsDatos!bPrimaAnual 'APRI20181023 ERS071-2018
            cmdAceptar.Enabled = True
            cmbBuscar.Enabled = False
        Else
            MsgBox "La persona no cuenta con un  seguro de sepelio activo.", vbInformation, "ALERTA SICMAC MAYNAS"
            Call LimpiaFormulario
        End If
    Else
        'MsgBox "No se encontraron datos de la persona.", vbInformation, "ALERTA SICMAC MAYNAS"
        Call LimpiaFormulario
    End If
    Set oSegSep = Nothing
End Sub

Private Sub CmdAceptar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII
    Call Grabar
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaFormulario
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub LimpiaFormulario()
    TxtBuscarPersona.Text = ""
    lblDOI.Caption = ""
    lblNombre.Caption = ""
    lblMonto.Text = ""
    lblMoneda.Caption = ""
    lblNumCerti.Caption = ""
    lsFecAfiliacion = ""
    lnEstado = 0
    cmdAceptar.Enabled = False
    cmbBuscar.Enabled = True
End Sub

Private Sub Grabar()
    Dim clsCapMov As New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCont As New COMNContabilidad.NCOMContFunciones
    Dim ClsMov As New COMDMov.DCOMMov
    Dim oBol As New COMNCaptaGenerales.NCOMCaptaImpresion
    Dim oBITF As New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCap As New COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim oSeg As New COMNCaptaGenerales.NCOMSeguros
    Dim lnMovNro As Long
    Dim lnMovNroITF As Long
    Dim lbBan As Boolean
    Dim CodOpe As String
    Dim lnMonto As Currency
    Dim lsMov As String
    Dim lsBoleta As String
    Dim lsDocumento As String
    Dim lsBoletaITF As String
    On Error GoTo Error
    
    lnMonto = lblMonto.Text
    lsMov = FechaHora(gdFecSis)
    lsDocumento = lblNumCerti.Caption
   
    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If MsgBox("Desea Grabar la Información", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        
        Dim sPersLavDinero As String
        Dim nMontoLavDinero As Double
        Dim lnMoneda As String
        Dim nTC As Double
        Dim sFecAfiliacion As String
        
        'sFecAfiliacion = DateAdd("D", 30, lsFecAfiliacion)
        'sFecAfiliacion = DateAdd("M", 1, lsFecAfiliacion) 'APRI20171027 ERS028-2017
        sFecAfiliacion = IIf(lbPrimaAnual, DateAdd("M", 12, lsFecAfiliacion), DateAdd("M", 1, lsFecAfiliacion)) 'APRI20181023 ERS071-2018
        'sFecAfiliacion = IIf(lnEstado = 1, sFecAfiliacion, sFecAfiliacion)
        lnMovNro = clsCapMov.OtrasOperaciones(lsMov, gServCobSepelioPrima, lnMonto, lsDocumento, "Cobro Prima Seguro Sepelio", gMonedaNacional, TxtBuscarPersona.Text, , , , , , , gnMovNro)
        Call clsCap.AgregaSegSepelioAfiliacionhis(lsDocumento, "", Format(sFecAfiliacion, "yyyyMMdd"), lsMov, lnMovNro, TxtBuscarPersona.Text, gsCodAge, 501) 'APRI20171027 ERS028-2017 CAMBIÒ ESTADO 1 POR 501
        Call oSeg.ActualizaEstadoSeguroSepelio(lsDocumento, sFecAfiliacion, lnMovNro, 501) 'APRI20171027 ERS028-2017 CAMBIÒ ESTADO 1 POR 501
        If gnMovNro = 0 Then
            MsgBox "La operación no se realizó, favor intentar nuevamente", vbInformation, "Aviso"
            Exit Sub
        End If
        
        Set clsCont = Nothing
        Set clsCapMov = Nothing
        Set ClsMov = Nothing
               
        
        MsgBox "Debe realizar el cobro de S./" & lblMonto.Text, vbInformation, "Aviso"
        lsBoleta = oSeg.ImprimeBoletaAfilicacionSeguroSepelio(lnMovNro, lsMov, gsNomAge, gbImpTMU)
        
        Do
           If Trim(lsBoleta) <> "" Then
                lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    Print #nFicSal, ""
                Close #nFicSal
          End If
          
          If Trim(lsBoletaITF) <> "" Then
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoletaITF
                Print #nFicSal, ""
            Close #nFicSal
          End If
            
        Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
        Set oBol = Nothing
    End If
    Call LimpiaFormulario
    'APRI20180206 ERS028-2017
    If lnTipo = 1 Then
        Unload Me
    End If
    'END APRI
    Exit Sub
Error:
      MsgBox Str(err.Number) & err.Description
End Sub
Private Function VerificaAfiliacionDesactivada(ByVal psPersCod As String) As Boolean
    Dim oSeg As New COMNCaptaGenerales.NCOMSeguros
    Dim rs As New ADODB.Recordset
    
    Set rs = oSeg.VerificaAfiliacionDesactivada(psPersCod)
    VerificaAfiliacionDesactivada = False
    If Not (rs.BOF And rs.EOF) Then
        VerificaAfiliacionDesactivada = True
    End If
End Function
Private Sub CargarDatos(ByVal psPersCod As String)
    Dim oSegSep As New COMNCaptaGenerales.NCOMSeguros
    Dim ClsPersona As New COMDPersona.DCOMPersonas
    Dim rsDatos As New ADODB.Recordset
    Dim R As New ADODB.Recordset
    Dim lsPersCod As String
    
    lsPersCod = psPersCod
    If lsPersCod <> "" Then
        Set ClsPersona = New COMDPersona.DCOMPersonas
        Set R = ClsPersona.BuscaCliente(lsPersCod, BusquedaCodigo)
        Set rsDatos = oSegSep.ObtieneDatosPersSegSepelio(lsPersCod)
        If Not (rsDatos.EOF And rsDatos.BOF) Then
            lblDOI.Caption = R!cPersIDnroDNI
            lblNombre.Caption = R!cPersNombre
            lblMonto.Text = Format(rsDatos!nMontoPrima, gsFormatoNumeroView)
            lblMoneda.Caption = "S/."
            lblNumCerti.Caption = rsDatos!cNumCertificado
            TxtBuscarPersona.Text = lsPersCod
            lsFecAfiliacion = rsDatos!dFecAfiliacion 'APRI20171027 ERS028-2017
            lbPrimaAnual = rsDatos!bPrimaAnual 'APRI20181023 ERS071-2018
        Else
            MsgBox "La persona no cuenta con un  seguro de sepelio activo.", vbInformation, "ALERTA SICMAC MAYNAS"
            Call LimpiaFormulario
        End If
    Else
        MsgBox "No se encontraron datos de la persona.", vbInformation, "ALERTA SICMAC MAYNAS"
        Call LimpiaFormulario
    End If
End Sub

Private Sub TxtBuscarPersona_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
End Sub
