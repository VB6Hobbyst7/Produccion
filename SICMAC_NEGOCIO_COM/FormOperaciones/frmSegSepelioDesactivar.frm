VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSegSepelioDesactivar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Anular Seguro Sepelio"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "frmSegSepelioDesactivar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   6090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   240
      TabIndex        =   13
      Top             =   2640
      Width           =   5655
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   300
         Left            =   4440
         TabIndex        =   16
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdDesactivar 
         Caption         =   "Desactivar"
         Enabled         =   0   'False
         Height          =   300
         Left            =   2250
         TabIndex        =   15
         Top             =   240
         Width           =   1095
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   300
         Left            =   3340
         TabIndex        =   14
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   " Datos de Afiliación "
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   7
      Top             =   1680
      Width           =   5655
      Begin SICMACT.ActXCodCta_New ActXCodCta 
         Height          =   735
         Left            =   2760
         TabIndex        =   10
         Top             =   120
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1296
         Texto           =   "Cuenta"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblMonto 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   1440
         TabIndex        =   17
         Top             =   525
         Width           =   1215
      End
      Begin VB.Label lblMoneda 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   960
         TabIndex        =   12
         Top             =   525
         Width           =   375
      End
      Begin VB.Label lblCertificado 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   270
         Left            =   960
         TabIndex        =   11
         Top             =   240
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "DOI:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   550
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Certificado:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   " Asegurado "
      ForeColor       =   &H00FF0000&
      Height          =   1215
      Left            =   240
      TabIndex        =   1
      Top             =   480
      Width           =   5655
      Begin VB.CommandButton cmbBuscar 
         Caption         =   "Buscar"
         Height          =   280
         Left            =   3060
         TabIndex        =   19
         Top             =   210
         Width           =   720
      End
      Begin VB.TextBox TxtBuscarPersona 
         Height          =   285
         Left            =   960
         TabIndex        =   18
         Top             =   230
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Código:"
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Cliente:"
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   525
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "DOI:"
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   820
         Width           =   615
      End
      Begin VB.Label lblNombre 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   3
         Top             =   525
         Width           =   4575
      End
      Begin VB.Label lblDOI 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   960
         TabIndex        =   2
         Top             =   795
         Width           =   1575
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   5953
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Desactivar"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSegSepelioDesactivar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************
'*  CREADOR:       RECO,Renzo Cordova                                               *
'*  FECHA:         08/02/2016                                                       *
'*  DESCRIPCION:   Formulario creado para desactivar la afiliacion als eguro sepelio*
'************************************************************************************

Option Explicit

Dim MinIng(3) As Integer
Dim MaxIng(3) As Integer
Dim MaxPer(3) As Integer
'Dim DerArrep(3) As Integer
Dim DerArrep As Integer 'APRI20171024 ERS028-2017
Dim sFechaAfilia As Date

Private Sub cmbBuscar_Click()
    Dim oSegSep As New COMNCaptaGenerales.NCOMSeguros
    Dim ClsPersona As New COMDPersona.DCOMPersonas
    Dim oPersona As COMDPersona.UCOMPersona
    Dim rsDatos As New ADODB.Recordset
    Dim R As New ADODB.Recordset
    Dim lsPersCod As String
    
    lblMonto.Enabled = False
    
    Set ClsPersona = New COMDPersona.DCOMPersonas
    
    
    Set oPersona = frmBuscaPersona.Inicio
    
    If Not oPersona Is Nothing Then
        lsPersCod = oPersona.sPersCod
        TxtBuscarPersona.Text = lsPersCod
    End If
    
    'lsPersCod = TxtBuscarPersona.psCodigoPersona
    If lsPersCod <> "" Then
        Set ClsPersona = New COMDPersona.DCOMPersonas
        Set R = ClsPersona.BuscaCliente(lsPersCod, BusquedaCodigo)
        Set rsDatos = oSegSep.ObtieneDatosPersSegSepelio(lsPersCod)
        If Not (rsDatos.EOF And rsDatos.BOF) Then
            lblDOI.Caption = R!cPersIDnroDNI
            lblNombre.Caption = R!cPersNombre
            lblMonto.Caption = Format(rsDatos!nMontoPrima, gsFormatoNumeroView)
            lblMoneda.Caption = "S/."
            lblCertificado.Caption = rsDatos!cNumCertificado
            ActXCodCta.NroCuenta = rsDatos!cCtaCodDebito
            sFechaAfilia = rsDatos!dFecAfiliacionInicio 'APRI20170725 CAMBIO A dFecAfiliacionInicio
            TxtBuscarPersona.Enabled = False
            cmdDesactivar.Enabled = True
        Else
            MsgBox "El cliente no registra seguro activo.", vbInformation, "ALERTA"
            Call LimpiarFormulario
        End If
    Else
        Call LimpiarFormulario
    End If
End Sub

Private Sub cmdCancelar_Click()
    Call LimpiarFormulario
End Sub

Private Sub cmdDesactivar_Click()
    Dim oSegSep As New COMNCaptaGenerales.NCOMSeguros
    Dim loVistoElectronico As New frmVistoElectronico
    Dim nDiastrans As Integer
    Dim nEstado As Integer
     Dim sGlosa As String 'APRI20171024 ERS028-2017
    If loVistoElectronico.Inicio(17, "401596") Then
        nDiastrans = DateDiff("D", sFechaAfilia, gdFecSis)
        'If DerArrep(2) >= nDiastrans Then
        If DerArrep >= nDiastrans Then 'APRI20171024 ERS028-2017
            nEstado = 3
            sGlosa = "DEVOLUCION POR ANULACIÓN SEGURO SEPELIO ANTES DEL PERIODO DE ARREPENTIMIENTO"
        Else
            nEstado = 5
            sGlosa = "ANULACIÓN SEGURO SEPELIO POSTERIOR DEL PERIODO DE ARREPENTIMIENTO"
        End If
        Call GrabarDevolucion(nEstado, sGlosa)
        Call LimpiarFormulario
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Call CargarCriterios
    ActXCodCta.Enabled = False
End Sub

Private Sub LimpiarFormulario()
    TxtBuscarPersona.Text = ""
    lblDOI.Caption = ""
    lblCertificado.Caption = ""
    lblNombre.Caption = ""
    lblMoneda.Caption = ""
    lblMonto.Caption = ""
    ActXCodCta.NroCuenta = ""
    TxtBuscarPersona.Enabled = True
    cmdDesactivar.Enabled = False
End Sub

Private Sub CargarCriterios()
    Dim oSeg As New COMNCaptaGenerales.NCOMSeguros
    Dim rs As New ADODB.Recordset
    Dim nIndex As Integer
    
    Set rs = oSeg.SepelioObtieneCriterios
    
    For nIndex = 0 To rs.RecordCount - 1
        If rs!nCriterioID = 1 Then
            MinIng(0) = rs!nAnio: MinIng(1) = rs!nMes: MinIng(2) = rs!nDia
        ElseIf rs!nCriterioID = 2 Then
            MaxIng(0) = rs!nAnio: MaxIng(1) = rs!nMes: MaxIng(2) = rs!nDia
        ElseIf rs!nCriterioID = 3 Then
            MaxPer(0) = rs!nAnio: MaxPer(1) = rs!nMes: MaxPer(2) = rs!nDia
'COMENTADO APRI20171024 ERS028-2017
'        ElseIf rs!nCriterioID = 4 Then
'            DerArrep(0) = rs!nAnio: DerArrep(1) = rs!nMes: DerArrep(2) = rs!nDia
        End If
        rs.MoveNext
    Next
    Set oSeg = Nothing
    DerArrep = oSeg.SegSepelioParametros(7) 'APRI20171024 ERS028-2017
    
End Sub

Private Sub GrabarDevolucion(ByVal nEstado As Integer, ByVal psGlosa As String)
    'APRI20171024 ADD PARAMETRO psGlosa
    Dim clsCapMov As New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCont As New COMNContabilidad.NCOMContFunciones
    Dim clsMov As New COMDMov.DCOMMov
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
    
    On Error GoTo Error
    
    lnMonto = lblMonto.Caption
    lsMov = FechaHora(gdFecSis)
    lsDocumento = lblCertificado.Caption
   
    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If MsgBox("Está seguro de la desafiliación", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        
        Dim sPersLavDinero As String
        Dim nMontoLavDinero As Double
        Dim lnMoneda As String
        Dim nTC As Double
        If nEstado = 3 Then
            'lnMovNro = clsCapMov.OtrasOperaciones(lsMov, "401596", lnMonto, lsDocumento, "Seg. Sepelio- Devolución por Anulacion", gMonedaNacional, TxtBuscarPersona.Text, , , , , , , gnMovNro)
            lnMovNro = clsCapMov.OtrasOperaciones(lsMov, "401596", lnMonto, lsDocumento, psGlosa, gMonedaNacional, TxtBuscarPersona.Text, , , , , , , gnMovNro) 'APRI20171024 ERS028-2017 ADD  psGlosa
            If gnMovNro = 0 Then
                MsgBox "La operación no se realizó, favor intentar nuevamente", vbInformation, "Aviso"
            Exit Sub
            End If
        Else
             'APRI20171024 ERS028-2017
            clsCap.AgregaMov lsMov, "300153", psGlosa, gMovEstContabNoContable, gMovFlagVigente
            lnMovNro = clsCap.GetnMovNro(lsMov)
            'END APRI
        End If
        Call clsCap.AgregaSegSepelioAfiliacionhis(lsDocumento, "", gdFecSis, lsMov, lnMovNro, TxtBuscarPersona.Text, gsCodAge, 502) 'APRI20171024 ERS028-2017 CHANGE nEstado To 502
        Call oSeg.ActualizaEstadoSeguroSepelio(lsDocumento, gdFecSis, lnMovNro, 502) 'APRI20171024 ERS028-2017 CHANGE nEstado To 502
        
        
        Set clsCont = Nothing
        Set clsCapMov = Nothing
        Set clsMov = Nothing
               
        If nEstado = 3 Then
            MsgBox "Devolver al cliente S./" & lblMonto.Caption, vbInformation, "Aviso"
            lsBoleta = oSeg.ImprimeBoletaAfilicacionSeguroSepelio(lnMovNro, lsMov, gsNomAge, gbImpTMU)
        Else
            'lsBoleta = oSeg.ImprimeBoletaDesafilicacionSegSepelio(lnMovNro, lsMov, gsNomAge, gbImpTMU)
        End If
        If lsBoleta <> "" Then
            Do
                If Trim(lsBoleta) <> "" Then
                    lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea
                    nFicSal = FreeFile
                    Open sLpt For Output As nFicSal
                        Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                        Print #nFicSal, ""
                    Close #nFicSal
                End If
            Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
        End If
    End If
    Exit Sub
Error:
      MsgBox str(Err.Number) & Err.Description
End Sub
