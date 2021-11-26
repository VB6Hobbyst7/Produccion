VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmSegSepelioAfiliacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Afiliación a Seguros"
   ClientHeight    =   7080
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8790
   Icon            =   "frmSegSepelioAfiliacion.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7080
   ScaleWidth      =   8790
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   310
      Left            =   7440
      TabIndex        =   27
      Top             =   6570
      Width           =   1095
   End
   Begin VB.Frame frSegSepelio3 
      Caption         =   "Beneficiarios"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   2295
      Left            =   240
      TabIndex        =   12
      Top             =   4200
      Width           =   8295
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   310
         Left            =   960
         TabIndex        =   17
         Top             =   1880
         Width           =   855
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   310
         Left            =   120
         TabIndex        =   16
         Top             =   1880
         Width           =   855
      End
      Begin SICMACT.FlexEdit feBeneficiarios 
         Height          =   1335
         Left            =   120
         TabIndex        =   15
         Top             =   480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   2355
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nombre-Parentesco-Participación"
         EncabezadosAnchos=   "300-5200-1400-1000"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3"
         ListaControles  =   "0-0-3-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-R"
         FormatosEdit    =   "0-0-0-2"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.OptionButton optBeneficiarioNo 
         Caption         =   "No"
         Height          =   195
         Left            =   360
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optBeneficiarioSi 
         Caption         =   "Si"
         Height          =   195
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin VB.Frame frSegSepelio2 
      Caption         =   "Asegurado"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   1095
      Left            =   240
      TabIndex        =   10
      Top             =   3120
      Width           =   8295
      Begin SICMACT.FlexEdit feAsegurado 
         Height          =   735
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1296
         Cols0           =   4
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Código-DNI-Nombre"
         EncabezadosAnchos=   "300-1400-1700-4500"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-X-X"
         ListaControles  =   "0-1-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-L"
         FormatosEdit    =   "0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame frSegSepelio1 
      Caption         =   "¿Desea asegurarse o asegurar a algún familiar?"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   19
      Top             =   2160
      Width           =   8295
      Begin VB.CommandButton cmdBuscarCtaSep 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   5920
         TabIndex        =   28
         Top             =   420
         Width           =   200
      End
      Begin VB.OptionButton optAseguraSeplNo 
         Caption         =   "No"
         Height          =   195
         Left            =   240
         TabIndex        =   26
         Top             =   480
         Value           =   -1  'True
         Width           =   615
      End
      Begin VB.OptionButton optAseguraSeplSi 
         Caption         =   "Si"
         Height          =   195
         Left            =   1080
         TabIndex        =   25
         Top             =   480
         Width           =   615
      End
      Begin VB.Frame Frame6 
         Caption         =   "Prima:"
         Height          =   670
         Left            =   6240
         TabIndex        =   22
         Top             =   180
         Width           =   1935
         Begin VB.CheckBox ChkPrimaAnualSS 
            Caption         =   "Anual"
            Height          =   255
            Left            =   1080
            TabIndex        =   31
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtPrimaSepl 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "Certificado:"
         Height          =   670
         Left            =   1830
         TabIndex        =   20
         Top             =   180
         Width           =   1575
         Begin VB.TextBox txtCertificadoSepelio 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   115
            TabIndex        =   21
            Top             =   240
            Width           =   1335
         End
      End
      Begin SICMACT.ActXCodCta_New ActXCodCtaSepl 
         Height          =   735
         Left            =   3480
         TabIndex        =   24
         Top             =   180
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1296
         Texto           =   "Cuenta a Debitar:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Guardar"
      Height          =   310
      Left            =   6360
      TabIndex        =   18
      Top             =   6570
      Width           =   1095
   End
   Begin VB.Frame frSegTarjeta 
      Caption         =   "¿Desea asegurar su tarjeta?"
      Enabled         =   0   'False
      ForeColor       =   &H00FF0000&
      Height          =   975
      Left            =   240
      TabIndex        =   2
      Top             =   480
      Width           =   8295
      Begin VB.CommandButton cmdBuscaCtaTarj 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   290
         Left            =   5960
         TabIndex        =   29
         Top             =   420
         Width           =   200
      End
      Begin VB.Frame Frame3 
         Caption         =   "Certificado:"
         Height          =   670
         Left            =   1830
         TabIndex        =   8
         Top             =   180
         Width           =   1575
         Begin VB.TextBox txtCertificadoTarjeta 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   115
            TabIndex        =   9
            Top             =   240
            Width           =   1335
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Prima:"
         Height          =   670
         Left            =   6240
         TabIndex        =   6
         Top             =   180
         Width           =   1935
         Begin VB.CheckBox ChkPrimaAnualST 
            Caption         =   "Anual"
            Height          =   255
            Left            =   1080
            TabIndex        =   30
            Top             =   240
            Width           =   735
         End
         Begin VB.TextBox txtPrimaTarj 
            Enabled         =   0   'False
            Height          =   285
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
      End
      Begin SICMACT.ActXCodCta_New ActXCodCtaTarj 
         Height          =   735
         Left            =   3480
         TabIndex        =   5
         Top             =   180
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1296
         Texto           =   "Cuenta a Debitar:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.OptionButton optAseguraTarjSi 
         Caption         =   "Si"
         Height          =   195
         Left            =   1080
         TabIndex        =   4
         Top             =   480
         Width           =   615
      End
      Begin VB.OptionButton optAseguraTarjNo 
         Caption         =   "No"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   480
         Value           =   -1  'True
         Width           =   615
      End
   End
   Begin ComctlLib.TabStrip SegSepelioParamNumCertificadoHis 
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   2778
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Seguro de Tarjetas"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5205
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   9181
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   1
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Seguro Sepelio"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmSegSepelioAfiliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim oNSegTar As COMNCaptaGenerales.NCOMSeguros
Dim oDSegTar As COMDCaptaGenerales.DCOMSeguros
Dim rs As ADODB.Recordset
Dim sNumTarj As String
Dim fsPersCod As String
Dim fsPersNombre As String
Dim fnParamNroSolicMes As Integer
'Dim fnMontoPrima As Double
'Dim fnMontoPrimaAnual As Double 'APRI20181023 ERS071-2018
'APRI20190405 MEJORA
Dim fnMontoPrimaSS As Double
Dim fnMontoPrimaAnualSS As Double
Dim fnMontoPrimaST As Double
Dim fnMontoPrimaAnualST As Double
'END APRI
Dim sPVV As String
Dim sPVVOrig As String
Dim nEstadoSegSepelio As Integer
Dim bOpeManual As Boolean
Dim MinIng(3) As Integer
Dim MaxIng(3) As Integer
Dim MaxPer(3) As Integer
Dim DerArrep(3) As Integer
Dim sFechaAfilia As Date

Public Sub Inicio(ByVal psCtaCod As String, Optional ByVal psNumtarjeta As String = "", Optional ByVal pnTipoBusqueda As TipoBusquedaSepelio = 1)
    Dim oDPersona As New COMDPersona.DCOMPersonas
    Dim oCons As New COMDConstSistema.DCOMGeneral
    Dim oSeg As New COMNCaptaGenerales.NCOMSeguros
    Dim rsDatos As New ADODB.Recordset
    Dim rsDatosCobroPrima As New ADODB.Recordset
    Dim nValorCandado As Integer
    
    nValorCandado = oCons.LeeConstSistema(521)
    
    If nValorCandado = 0 Then
        If pnTipoBusqueda <> gSegTpoBusPersCod Then
            Set rsDatos = oDPersona.ObtieneTitularCuenta(psCtaCod)
            If Not (rsDatos.EOF And rsDatos.BOF) Then
                fsPersCod = rsDatos!cPersCod
                fsPersNombre = rsDatos!cPersNombre
            End If
            If fsPersCod = "" Then Exit Sub
        Else
            fsPersCod = psCtaCod
        End If
        bOpeManual = False
        Call IniciaAfiliacion(fsPersCod, psNumtarjeta, 1, psCtaCod) 'APRI20171007 ADD psCtaCod ERS028-2017
    End If
End Sub

Public Sub IniciaAfilManual()
    Dim oDSeg As New COMDCaptaGenerales.DCOMSeguros
    Dim oCons As New COMDConstSistema.DCOMGeneral
    Dim rs As New ADODB.Recordset
    Dim nValorCandado As Integer
    
    nValorCandado = oCons.LeeConstSistema(521)
    If nValorCandado = 0 Then
        Set rs = oDSeg.RecuperaSegSepelioParametro(2)
        fnMontoPrimaSS = rs!nValor
        'APRI20181023 ERS071-2018
        Set rs = oDSeg.RecuperaSegSepelioParametro(10)
        fnMontoPrimaAnualSS = rs!nValor
        'END APRI
        txtPrimaSepl.Text = Format(fnMontoPrimaSS, "#0.00") & " "
        txtCertificadoSepelio.Text = ""
        bOpeManual = True
        frSegSepelio1.Enabled = True
        frSegSepelio2.Enabled = True
        frSegSepelio3.Enabled = True
        optAseguraSeplSi.value = True
        optBeneficiarioSi.value = True
        Call optAseguraSeplSi_Click
        Me.Show 1
    End If
End Sub

Private Sub IniciaAfiliacion(ByVal psPersCod As String, ByVal psNumtarjeta As String, pnTipoAfiliacion As Integer, Optional ByVal psCtaCod As String = "")
'APRI20171007 ADD psCtaCod ERS028-2017
    Dim bActivarSegTarj As Boolean
    Dim bActivarSegSepel As Boolean
    sNumTarj = psNumtarjeta
    
    Dim oDSeg As New COMDCaptaGenerales.DCOMSeguros
    Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
    
    Set rs = oDSeg.RecuperaSegTarjetaParametro(100)
        fnParamNroSolicMes = rs!nParamValor
     
    If pnTipoAfiliacion = 1 And psNumtarjeta <> "" Then
        Set rs = oDSeg.RecuperaNroSolicitudAfiliacionMes(sNumTarj, CInt(Right(gdFecSis, 4)), CInt(Mid(gdFecSis, 4, 2)))
        If Not rs.EOF Then
            If rs!nNroSolicMes < fnParamNroSolicMes Then
                Call oNSegTar.InsertaNroSolicitudAfiliacionMes(sNumTarj, gdFecSis)
                bActivarSegTarj = True
            ElseIf rs!nNroSolicMes = fnParamNroSolicMes Then
                Call oNSegTar.InsertaNroSolicitudAfiliacionMes(sNumTarj, gdFecSis)
                bActivarSegTarj = False
            Else
                bActivarSegTarj = False
            End If
        Else
            Call oNSegTar.InsertaNroSolicitudAfiliacionMes(sNumTarj, gdFecSis)
            bActivarSegTarj = True
        End If
        Set rs = oDSeg.RecuperaSegTarjetaParametro(101)
        fnMontoPrimaST = rs!nParamValor
        'APRI20181023 ERS071-2018
        Set rs = oDSeg.RecuperaSegTarjetaParametro(106)
        fnMontoPrimaAnualST = rs!nParamValor
        'END APRI
        txtPrimaTarj.Text = Format(fnMontoPrimaST, "#0.00") & " "
        txtCertificadoTarjeta.Text = oDSeg.ObtenerSegTarjetaNumCertificado
    End If
    
    If VerificaAfiliacionActiva(Format(gdFecSis, "yyyyMMdd"), psPersCod) Then
        bActivarSegSepel = False
    Else
        '*************APRI20171007 ERS028-2017*****************
        If psCtaCod <> "" Then
            Dim ClsPersona As COMDPersona.DCOMPersonas
            Dim R As New ADODB.Recordset
            Dim nEdad As Integer
            Dim sMsj As String

            If Mid(psCtaCod, 6, 3) = "232" Or Mid(psCtaCod, 6, 3) = "234" Then
                ActXCodCtaSepl.CMAC = "109"
                ActXCodCtaSepl.Age = Mid(psCtaCod, 4, 2)
                ActXCodCtaSepl.Prod = Mid(psCtaCod, 6, 3)
                ActXCodCtaSepl.Cuenta = Mid(psCtaCod, 9, 10)
            End If

            Set ClsPersona = New COMDPersona.DCOMPersonas
            Set R = ClsPersona.BuscaCliente(psPersCod, BusquedaCodigo)
            nEdad = EdadPersona(R!dPersNacCreac, gdFecSis)
            
            sMsj = ValidaCriterios(nEdad)
            If sMsj = "" Then
                feAsegurado.TextMatrix(1, 1) = R!cPersCod
                feAsegurado.TextMatrix(1, 2) = R!cPersIDnroDNI
                feAsegurado.TextMatrix(1, 3) = R!cPersNombre
            End If
            
        End If
        '*************END APRI20171007***************
    
        Set rs = oDSeg.RecuperaSegSepelioParametro(2)
        fnMontoPrimaSS = rs!nValor
        'APRI20181023 ERS071-2018
        Set rs = oDSeg.RecuperaSegSepelioParametro(10)
        fnMontoPrimaAnualSS = rs!nValor
        'END APRI
        txtPrimaSepl.Text = Format(fnMontoPrimaSS, "#0.00") & " "
        txtCertificadoSepelio.Text = ""
        bActivarSegSepel = True
    End If
    
        frSegTarjeta.Enabled = True
        If bActivarSegSepel Then
            optAseguraSeplSi.value = True
            optBeneficiarioSi.value = True
            frSegSepelio1.Enabled = True
            frSegSepelio2.Enabled = True
            frSegSepelio3.Enabled = True
        End If
        If psNumtarjeta = "" Then
            frSegTarjeta.Enabled = False
        End If
        
        Set oNSegTar = Nothing
        Set oDSeg = Nothing
        If pnTipoAfiliacion = 1 Then
        
        If Not bActivarSegTarj And Not bActivarSegSepel Then
            Exit Sub
        End If
        'If Not frmSegSepelioCobroPrima.Inicia(psPersCod) Then 'COMENTADO APRI20171024 ERS028-2017
            Dim clsSegN As New COMNCaptaGenerales.NCOMSeguros
            Dim rsValidaExiste As New ADODB.Recordset
            Set rsValidaExiste = clsSegN.DevulveClienteSeguroActivo(fsPersCod)
            If Not (rsValidaExiste.EOF And rsValidaExiste.BOF) Then
            Else
                Call optAseguraTarjNo_Click
                Me.Show 1
            End If
        'End If
    End If
End Sub
'APRI20181023 ERS071-2018
Private Sub ChkPrimaAnualSS_Click()
    If ChkPrimaAnualSS.value Then
        txtPrimaSepl.Text = Format(fnMontoPrimaAnualSS, "#0.00") & " "
    Else
        txtPrimaSepl.Text = Format(fnMontoPrimaSS, "#0.00") & " "
    End If
End Sub

Private Sub ChkPrimaAnualST_Click()
    If ChkPrimaAnualST.value Then
        txtPrimaTarj.Text = Format(fnMontoPrimaAnualST, "#0.00") & " "
    Else
        txtPrimaTarj.Text = Format(fnMontoPrimaST, "#0.00") & " "
    End If
End Sub
'END APRI
Private Sub CmdAceptar_Click()
    Dim msj As String
    Dim bSepl As Boolean
    Dim bTarj As Boolean
    Dim rsValidaExiste As ADODB.Recordset
    Dim clsSegN As New COMNCaptaGenerales.NCOMSeguros
    cmdAceptar.Enabled = False 'APRI20170614 SEGUN INCIDENTE INC1706130007
    
    
    If optAseguraTarjSi.value = True Then
        'APRI20190408 MEJORA
        If ChkPrimaAnualST.value And CDbl(txtPrimaTarj.Text) = 0 Then
            MsgBox "Producto no aplica a prima anual.", vbInformation, "Alerta"
            cmdAceptar.Enabled = True
            Exit Sub
        End If
        'END APRI

        bTarj = RegistrarAfiliacionTarjeta
        If Not bTarj Then 'APRI20171004 MEJORA
            cmdAceptar.Enabled = True
        End If
    End If
    
    If optAseguraSeplSi.value = True Then
        fsPersCod = feAsegurado.TextMatrix(1, 1) 'APRI20170614  SEGUN SATI TIC1706140002

        If Not fsPersCod = "" Then
            Set rsValidaExiste = clsSegN.DevulveClienteSeguroActivo(fsPersCod)
            If Not (rsValidaExiste.EOF And rsValidaExiste.BOF) Then
                MsgBox "El cliente ya tiene un seguro sepelio afiliado", vbInformation, "Alerta"
                Call LimpiarFormulario
                cmdAceptar.Enabled = True 'APRI20170614 SEGUN INCIDENTE INC1706130007
                Exit Sub
            End If
            'APRI20180206 ERS028-2017
            If clsSegN.SepelioVerificaProducto(fsPersCod) And ActXCodCtaSepl.NroCuenta = "" Then
                MsgBox "El cliente tiene cuenta(s) vigente(s), no es posible afiliar el seguro en modo Pago Efectivo!", vbInformation, "Alerta"
                cmdAceptar.Enabled = True
                Exit Sub
            End If
            'END APRI
        End If
        If txtCertificadoSepelio.Text = "" And optAseguraSeplSi.value = True Then
            MsgBox "Debe llenar la casilla de [Num. Certificado].", vbInformation, "Alerta"
            cmdAceptar.Enabled = True 'APRI20170614 SEGUN INCIDENTE INC1706130007
            Exit Sub
        End If
    
        If feAsegurado.TextMatrix(1, 1) = "" Or feAsegurado.TextMatrix(1, 2) = "" Or feAsegurado.TextMatrix(1, 3) = "" Then
            MsgBox "Los datos del Asegurado no deben estar vacíos.", vbInformation, "Alerta"
            cmdAceptar.Enabled = True 'APRI20170614 SEGUN INCIDENTE INC1706130007
            Exit Sub
        End If
        
        msj = ValidaDatosBeneficiario
        
        If Mid(ActXCodCtaSepl.NroCuenta, 9, 1) = 2 Then
            MsgBox "No se puede vincular cuentas en dolares", vbInformation, "Alerta SICMAC"
            Call cmdBuscarCtaSep_Click
            cmdAceptar.Enabled = True 'APRI20170614 SEGUN INCIDENTE INC1706130007
            Exit Sub
        End If
        
        If msj = "" Then
            bSepl = RegistrarAfiliacionSepelio(bOpeManual)
            If Not bSepl Then
            cmdAceptar.Enabled = True 'APRI20170614 SEGUN INCIDENTE INC1706130007
            Exit Sub
            End If
            
        Else
            MsgBox msj, vbInformation, "Alerta"
            cmdAceptar.Enabled = True 'APRI20170614 SEGUN INCIDENTE INC1706130007
            Exit Sub
        End If
    End If
    If bTarj = True Or bSepl = True Then
        Unload Me
    End If

End Sub

Private Sub cmdBuscaCtaTarj_Click()
    Dim oDCapGen As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim loCuentas As New COMDPersona.UCOMProdPersona
    Dim rsCuentas As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
    Set rs = oDSegTar.RecuperaPersonaTarjeta(sNumTarj)
    If rs.EOF Then
        MsgBox "No se pueden obtener los datos de la persona", vbInformation, "Aviso"
        fsPersCod = ""
        Exit Sub
    End If
    fsPersCod = rs!cPersCod
    fsPersNombre = rs!cPersNombre
    Set oDSegTar = Nothing
    
    If Trim(fsPersCod) <> "" Then
        Set rsCuentas = oDCapGen.GetCuentasPersona(fsPersCod, , True, True, , , , , True, "0,2")
        Set oDCapGen = Nothing
    End If
        
    If rsCuentas.RecordCount > 0 Then
        Set loCuentas = frmProdPersona.Inicio(fsPersNombre, rsCuentas)
        If loCuentas.sCtaCod <> "" Then
            ActXCodCtaTarj.CMAC = "109"
            ActXCodCtaTarj.Age = Mid(loCuentas.sCtaCod, 4, 2)
            ActXCodCtaTarj.Prod = Mid(loCuentas.sCtaCod, 6, 3)
            ActXCodCtaTarj.Cuenta = Mid(loCuentas.sCtaCod, 9, 10)
            'ActXCodCtaTarj.SetFocus 'APRI20181023 MEJORA
        End If
        Set loCuentas = Nothing
    Else
        MsgBox "El cliente no tiene cuentas de ahorro activas", vbInformation, "Aviso"
        fsPersCod = ""
    End If
End Sub

Private Sub cmdBuscarCtaSep_Click()
    Dim oDCapGen As New COMDCaptaGenerales.DCOMCaptaGenerales
    Dim loCuentas As New COMDPersona.UCOMProdPersona
    Dim rsCuentas As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    
    Dim clsMant As New COMNCaptaGenerales.NCOMCaptaGenerales
    
    If fsPersCod = "" And bOpeManual = True Then
        If Not feAsegurado.TextMatrix(1, 1) = "" Then
            fsPersCod = feAsegurado.TextMatrix(1, 1)
            Set rs = clsMant.GetCuentasPersona(fsPersCod)
        Else
            MsgBox "Debe ingresar el asegurado", vbInformation, "Aviso"
            fsPersCod = ""
            Exit Sub
        End If
    Else
        Set rs = clsMant.GetCuentasPersona(fsPersCod)
    End If
    
    If rs.EOF Then
        MsgBox "No se pueden obtener los datos de la persona", vbInformation, "Aviso"
        fsPersCod = ""
        Exit Sub
    End If
    
    Set oDSegTar = Nothing
    
    If Trim(fsPersCod) <> "" Then
        Set rsCuentas = oDCapGen.GetCuentasPersona(fsPersCod, , True, True, , , , , True, "0,2")
        Set oDCapGen = Nothing
    End If
        
    If rsCuentas.RecordCount > 0 Then
        Set loCuentas = frmProdPersona.Inicio(fsPersNombre, rsCuentas)
        If loCuentas.sCtaCod <> "" Then
            ActXCodCtaSepl.CMAC = "109"
            ActXCodCtaSepl.Age = Mid(loCuentas.sCtaCod, 4, 2)
            ActXCodCtaSepl.Prod = Mid(loCuentas.sCtaCod, 6, 3)
            ActXCodCtaSepl.Cuenta = Mid(loCuentas.sCtaCod, 9, 10)
            'ActXCodCtaSepl.SetFocus
        End If
        Set loCuentas = Nothing
    Else
        MsgBox "El cliente no tiene cuentas de ahorro activas", vbInformation, "Aviso"
        fsPersCod = ""
    End If
End Sub

Private Sub cmdsalir_Click()
    Call LimpiarFormulario
    Unload Me
End Sub

Private Function ValidaDatos(ByVal oCuenta As ActXCodCta_New) As Boolean
    ValidaDatos = False
    If Trim(fsPersCod) = "" Then
        MsgBox "No hay datos de la persona", vbInformation, "Aviso"
        Exit Function
    End If
    ValidaDatos = True
End Function

Function SoloNumeros(ByVal KeyAscii As Integer) As Integer
    If InStr("0123456789", Chr(KeyAscii)) = 0 Then
        SoloNumeros = 0
    Else
        SoloNumeros = KeyAscii
    End If
    If KeyAscii = 8 Then SoloNumeros = KeyAscii
    If KeyAscii = 13 Then SoloNumeros = KeyAscii
End Function

Private Sub cmdAgregar_Click()
    feBeneficiarios.AdicionaFila
End Sub

Private Sub cmdQuitar_Click()
    feBeneficiarios.EliminaFila (feBeneficiarios.row)
End Sub

Private Sub CargaComboParentesco()
    Dim oConst As New COMDConstantes.DCOMConstantes
    
    Set rs = oConst.ObtenerVarRecuperaciones(1006)
    feBeneficiarios.CargaCombo oConst.ObtenerVarRecuperaciones(1006)
    Set oConst = Nothing
End Sub

Private Sub feAsegurado_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim ClsPersona As COMDPersona.DCOMPersonas
    Dim oCred As COMDCredito.DCOMCredito
    Dim nValor As Integer
    Dim R As New ADODB.Recordset
    Dim nEdad As Integer
    Dim sMsj As String
    If psDataCod <> "" Then
        Set ClsPersona = New COMDPersona.DCOMPersonas
        Set R = ClsPersona.BuscaCliente(psDataCod, BusquedaCodigo)
        nEdad = EdadPersona(R!dPersNacCreac, gdFecSis)
        
        sMsj = ValidaCriterios(nEdad)
        If sMsj <> "" Then
            MsgBox sMsj, vbInformation, "Alerta SICMAC"
            feAsegurado.TextMatrix(pnRow, 1) = ""
            feAsegurado.TextMatrix(pnRow, 2) = ""
            feAsegurado.TextMatrix(pnRow, 3) = ""
        Else
            feAsegurado.TextMatrix(pnRow, 1) = R!cPersCod
            feAsegurado.TextMatrix(pnRow, 2) = R!cPersIDnroDNI
            feAsegurado.TextMatrix(pnRow, 3) = R!cPersNombre
            
             'APRI20180201 ERS028-2017
            fsPersCod = R!cPersCod
            fsPersNombre = R!cPersNombre
            'END APRI
            
        End If
    Else
        'APRI20180201 ERS028-2017
        fsPersCod = ""
        fsPersNombre = ""
        'END APRI
    End If
End Sub

Private Sub feBeneficiarios_OnCellChange(pnRow As Long, pnCol As Long)
    Dim nIndice As Integer
    Dim nSumaTotPart As Double
    Select Case pnCol
        Case 2
            Dim nConyugeDup As Integer
            Dim j As Integer
            
            For j = 1 To feBeneficiarios.Rows - 1
                If Val(Right(feBeneficiarios.TextMatrix(j, 2), 2)) = 0 Then
                    nConyugeDup = nConyugeDup + 1
                    If nConyugeDup > 1 Then
                        MsgBox "No se puede agregar otro [Cónyuge]", vbInformation, "Alerta"
                        feBeneficiarios.TextMatrix(j, 2) = ""
                    End If
                End If
            Next
        Case 3
            If InStr(feBeneficiarios.TextMatrix(pnRow, 3), "%") Then
            Else
                If Val(feBeneficiarios.TextMatrix(pnRow, 3)) > 100 Then
                    MsgBox "El valor de participación no puede ser mayor de 100%", vbInformation, "Alerta"
                    feBeneficiarios.TextMatrix(pnRow, 3) = 100
                End If
                If Val(feBeneficiarios.TextMatrix(pnRow, 3)) < 1 Then
                    MsgBox "El valor de participación no puede ser menor de 1%", vbInformation, "Alerta"
                    feBeneficiarios.TextMatrix(pnRow, 3) = 100
                End If
                feBeneficiarios.TextMatrix(pnRow, 3) = feBeneficiarios.TextMatrix(pnRow, 3) & "%"
            End If
            For nIndice = 1 To feBeneficiarios.Rows - 1
            '*********'APRI20170623 MEJORA
                If feBeneficiarios.TextMatrix(nIndice, 3) = "" Then
                    MsgBox "El valor de participación no puede estar vacio. Favor verifique.", vbInformation, "Alerta"
                    Exit For
                Else
                    nSumaTotPart = nSumaTotPart + Mid(feBeneficiarios.TextMatrix(nIndice, 3), 1, Len(feBeneficiarios.TextMatrix(nIndice, 3)) - 1)
                End If
            '**********END APRI20170623
            'nSumaTotPart = nSumaTotPart + Mid(feBeneficiarios.TextMatrix(nIndice, 3), 1, Len(feBeneficiarios.TextMatrix(nIndice, 3)) - 1)
            Next
            If nSumaTotPart > 100 Then
                MsgBox "El total de participación no puede superar el 100%", vbInformation, "Alerta"
                feBeneficiarios.TextMatrix(pnRow, 3) = 0 & "%"
            End If
    End Select
    feBeneficiarios.TextMatrix(pnRow, pnCol) = UCase(feBeneficiarios.TextMatrix(pnRow, pnCol))
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub

Private Sub Form_Load()
    Call CargaComboParentesco
    Call CargarCriterios
    feAsegurado.AdicionaFila
End Sub

Private Sub optAseguraSeplNo_Click()
    Call HabilitaControles(False, txtCertificadoSepelio, ActXCodCtaSepl, txtPrimaSepl, cmdBuscarCtaSep)
    frSegSepelio3.Enabled = False
    If optAseguraTarjNo.value = True And optAseguraSeplNo.value = True Then
        cmdAceptar.Enabled = False
    End If
    ChkPrimaAnualSS.Enabled = False 'APRI20181023 ERS071-2018
End Sub

Private Sub optAseguraSeplSi_Click()
    Call HabilitaControles(True, txtCertificadoSepelio, ActXCodCtaSepl, txtPrimaSepl, cmdBuscarCtaSep)
    frSegSepelio3.Enabled = True
    cmdAceptar.Enabled = True
    ChkPrimaAnualSS.Enabled = True 'APRI20181023 ERS071-2018
End Sub

Private Sub optAseguraTarjNo_Click()
    Call HabilitaControles(False, txtCertificadoTarjeta, ActXCodCtaTarj, txtPrimaTarj, cmdBuscaCtaTarj)
    If optAseguraTarjNo.value = True And optAseguraSeplNo.value = True Then
        cmdAceptar.Enabled = False
    End If
    ChkPrimaAnualST.Enabled = False 'APRI20181023 ERS071-2018
End Sub

Private Sub optAseguraTarjSi_Click()
    Call HabilitaControles(True, txtCertificadoTarjeta, ActXCodCtaTarj, txtPrimaTarj, cmdBuscaCtaTarj)
    cmdAceptar.Enabled = True
    ChkPrimaAnualST.Enabled = True 'APRI20181023 ERS071-2018
End Sub

Private Sub HabilitaControles(ByVal pbValor As Boolean, ByVal poTxtCertificado As TextBox, ByVal poTxtCuenta As ActXCodCta_New, ByVal poTxtPrima As TextBox, ByVal o As CommandButton)
    poTxtCertificado.Enabled = pbValor
    'poTxtCuenta.Enabled = pbValor
    poTxtCuenta.Enabled = Not pbValor 'APRI20180201 ERS028-2017
    o.Enabled = pbValor
End Sub

Private Function RegistrarAfiliacionTarjeta() As Boolean
    Dim oNCapMov As New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oNCapGen As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim oNContFunc As New COMNContabilidad.NCOMContFunciones
    Dim sCodOpe As CaptacOperacion
    Dim rsCuenta As New ADODB.Recordset
    Dim nMonto As Double
    Dim sCuenta As String
    Dim sProd As String
    Dim nTpoPrograma As String
    Dim sMovNro As String
    Dim lsmensaje As String
    Dim lsBoleta As String
    Dim nSaldo As Double
    
    RegistrarAfiliacionTarjeta = False
    If ActXCodCtaTarj.NroCuenta = "" Then
        MsgBox "Debe ingresar un número de cuenta para el seguro de tarjeta", vbInformation, "Alerta"
        Exit Function
    End If
    If ValidaDatos(ActXCodCtaTarj) Then
        nMonto = CDbl(txtPrimaTarj.Text)
        sCuenta = ActXCodCtaTarj.NroCuenta
        sProd = ActXCodCtaTarj.Prod
        
        Set rsCuenta = oNCapGen.GetDatosCuenta(sCuenta)
        Set oNCapGen = Nothing
        nTpoPrograma = rsCuenta!nTpoPrograma
    
'***********************COMENTADO APRI20171026 ERS028-2017*******************************
'        If sProd = gCapAhorros Then
'            If nTpoPrograma <> 0 And nTpoPrograma <> 5 And nTpoPrograma <> 6 And nTpoPrograma <> 8 Then
'                MsgBox "El subproducto de la cuenta no es permitido para esta operación", vbInformation, "Aviso"
'                Exit Function
'            End If
'        Else
'            If nTpoPrograma <> 0 And nTpoPrograma <> 1 Then
'                MsgBox "El subproducto de la cuenta no es permitido para esta operación", vbInformation, "Aviso"
'                Exit Function
'            End If
'        End If
        
        'APRI20171004 ERS028-2017
        Set oNSegTar = New COMNCaptaGenerales.NCOMSeguros
        If sProd = "233" Then
            MsgBox "Este tipo de cuenta no esta permitido para esta operación", vbInformation, "Aviso"
            Exit Function
        Else
            If Not oNSegTar.SepelioVerificaTpoPrograma(2, sProd, nTpoPrograma) Then
                MsgBox "El subproducto de la cuenta no es permitido para esta operación", vbInformation, "Aviso"
                Exit Function
            End If
        End If
        Set oNSegTar = Nothing
        'END APRI
        
        If Mid(sCuenta, 9, 1) = gMonedaExtranjera Then
            Dim ObjTc As New COMDConstSistema.NCOMTipoCambio
            Dim nTC As Double
            nTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoMes)
            Set ObjTc = Nothing
            nMonto = Round(nMonto / nTC, 2)
        End If
        
        If sProd = gCapAhorros Then
            sCodOpe = gAhoCargoAfilSegTarjeta
        Else
            sCodOpe = gCTSCargoAfilSegTarjeta
        End If
        
        'APRI20171004 MEJORA
            If sNumTarj <> "" Then
             Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
                If oDSegTar.VerificaSegTarjetaAfiliacion(sNumTarj) Then
                   MsgBox "La tarjeta ya está afiliada al Seguro", vbInformation, "Aviso"
                   Exit Function
                End If
             Set oDSegTar = Nothing
            End If
        'END APRI
        
        Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
        If oDSegTar.ValidaExisteRegistroNroCertificado(Trim(txtCertificadoTarjeta.Text)) Then
            MsgBox "El Número de Certificado ya fue registrado anteriormente", vbInformation, "Aviso"
            Exit Function
        End If
        
        Set oDSegTar = Nothing
        
        If oNCapMov.ValidaSaldoCuenta(sCuenta, nMonto, gAhoCargoAfilSegTarjeta) Then
            
            If MsgBox("Se va a realizar el Cargo a la cuenta por Afiliación de Tarjeta, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Function
            
            sMovNro = oNContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            Set oNContFunc = Nothing
        
            
            nSaldo = oNCapMov.CapCargoCuentaSegTarjeta(sCuenta, nMonto, sCodOpe, sMovNro, "Cuenta = " & sCuenta & ", Tarjeta = " & sNumTarj, gsNomAge, sLpt, gsCodCMAC, gsCodAge, lsmensaje, gbImpTMU, Trim(txtCertificadoTarjeta.Text), sNumTarj, gdFecSis, fsPersCod, lsBoleta, ChkPrimaAnualST.value)
            'APRI20181023 ERS071-2018 ADD ChkPrimaAnualST.Value
            
            If Trim(lsmensaje) <> "" Then
                MsgBox lsmensaje, vbInformation
                Exit Function
            End If
                    
            Dim nFicSal As Integer
            Do
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                Print #nFicSal, lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                Close #nFicSal
            Loop Until MsgBox("¿Desea reimprimir el voucher?", vbQuestion + vbYesNo, "Aviso") = vbNo
            RegistrarAfiliacionTarjeta = True
        Else
            MsgBox "La Cuenta NO posee saldo suficiente", vbInformation, "Aviso"
        End If
        Set oNCapMov = Nothing
    End If
End Function

Private Function RegistrarAfiliacionSepelio(ByVal pbOpeManual As Boolean)
    Dim oNCapMov As New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oNCapGen As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim oNContFunc As New COMNContabilidad.NCOMContFunciones
    Dim oNSegSep As New COMNCaptaGenerales.NCOMSeguros
    Dim sCodOpe As CaptacOperacion
    Dim rsCuenta As New ADODB.Recordset
    Dim nMonto As Double
    Dim sCuenta As String
    Dim sProd As String
    Dim nTpoPrograma As String
    Dim sMovNro As String
    Dim nMovNro As Long
    Dim lsmensaje As String
    Dim lsBoleta As String
    Dim nSaldo As Double
    Dim index As Integer
    
    RegistrarAfiliacionSepelio = False
    If Not bOpeManual Or ActXCodCtaSepl.NroCuenta <> "" Then
        If ActXCodCtaSepl.NroCuenta <> "" Then
            nMonto = CDbl(txtPrimaSepl.Text)
            sCuenta = ActXCodCtaSepl.NroCuenta
            sProd = ActXCodCtaSepl.Prod
            
            If Not (sProd = gCapAhorros Or sProd = gCapPlazoFijo Or sProd = gCapCTS) Then
                MsgBox "Cuenta invalida", vbInformation, "Alerta"
                Exit Function
            End If

            Set rsCuenta = oNCapGen.GetDatosCuenta(sCuenta)
            Set oNCapGen = Nothing

            If (rsCuenta.BOF And rsCuenta.EOF) Then
                MsgBox "La cuenta ingresada no es valida", vbInformation, "Alerta"
                Exit Function
            End If
            
        'End If
        nTpoPrograma = rsCuenta!nTpoPrograma

'***********************COMENTADO POR APRI20171004***************************************************
'        If sProd = gCapAhorros Then
'            If nTpoPrograma <> 0 And nTpoPrograma <> 5 And nTpoPrograma <> 6 And nTpoPrograma <> 8 Then
'                MsgBox "El subproducto de la cuenta no es permitido para esta operación", vbInformation, "Aviso"
'                RegistrarAfiliacionSepelio = False
'                Exit Function
'            End If
'        Else
'            If nTpoPrograma <> 0 And nTpoPrograma <> 1 Then
'                MsgBox "El subproducto de la cuenta no es permitido para esta operación", vbInformation, "Aviso"
'                RegistrarAfiliacionSepelio = False
'                Exit Function
'            End If
'        End If
        
        'APRI20171004 ERS028-2017
        Set oNSegSep = New COMNCaptaGenerales.NCOMSeguros
        If sProd = "233" Then
            MsgBox "Este tipo de cuenta no esta permitido para esta operación", vbInformation, "Aviso"
            RegistrarAfiliacionSepelio = False
            Exit Function
        Else
            If Not oNSegSep.SepelioVerificaTpoPrograma(1, sProd, nTpoPrograma) Then
                MsgBox "El subproducto de la cuenta no es permitido para esta operación", vbInformation, "Aviso"
                RegistrarAfiliacionSepelio = False
                Exit Function
            End If
        End If
        Set oNSegSep = Nothing
        'END APRI
            
        If Mid(sCuenta, 9, 1) = gMonedaExtranjera Then
            Dim ObjTc As New COMDConstSistema.NCOMTipoCambio
            Dim nTC As Double
            nTC = ObjTc.EmiteTipoCambio(gdFecSis, TCFijoMes)
            Set ObjTc = Nothing
            nMonto = Round(nMonto / nTC, 2)
        End If
            
        If sProd = gCapAhorros Then
            sCodOpe = gAhoCargoAfilSegSepelio
        Else
            sCodOpe = gCTSCargoAfilSegSepelio
        End If
            
        Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
        If oDSegTar.ValidaExisteRegistroNroCertificado(Trim(txtCertificadoSepelio.Text), gSegCertificadoSepelio) Then
            MsgBox "El Número de Certificado ya fue registrado anteriormente", vbInformation, "Aviso"
            RegistrarAfiliacionSepelio = False
            Exit Function
        End If
            
        Set oDSegTar = Nothing
            
        If oNCapMov.ValidaSaldoCuenta(sCuenta, nMonto, gAhoCargoAfilSegSepelio) Then
                
            If MsgBox("Se va a realizar el Cargo a la cuenta por Afiliación de Seguro de Sepelio, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Function
                
                sMovNro = oNContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
                Set oNContFunc = Nothing
            
                
                nSaldo = oNCapMov.CapCargoCuentaSegSepelio(sCuenta, nMonto, sCodOpe, sMovNro, "Afiliacion Seguro Sepelio Cuenta = " & sCuenta, gsNomAge, sLpt, gsCodCMAC, gsCodAge, lsmensaje, gbImpTMU, Trim(txtCertificadoSepelio.Text), sNumTarj, gdFecSis, feAsegurado.TextMatrix(1, 1), lsBoleta, gsCodAge, nMovNro, , ChkPrimaAnualSS.value)
                'APRI20181023 ERS071-2018 ADD ChkPrimaAnualSS.value
                
                If Trim(lsmensaje) <> "" Then
                    MsgBox lsmensaje, vbInformation
                    Exit Function
                Else
                    If optBeneficiarioSi.value = True Then
                        For index = 1 To feBeneficiarios.Rows - 1
                            Call oNSegSep.RegistraAseguradoSepelio(Trim(txtCertificadoSepelio.Text), feBeneficiarios.TextMatrix(index, 1), Right(feBeneficiarios.TextMatrix(index, 2), 3), Mid(feBeneficiarios.TextMatrix(index, 3), 1, Len(feBeneficiarios.TextMatrix(index, 3)) - 1), 1)
                        Next
                    End If
                End If
                        
                Dim nFicSal As Integer
                lsBoleta = oNSegSep.ImprimeBoletaAfilicacionSeguroSepelio(nMovNro, sMovNro, gsNomAge, gbImpTMU)
                Do
                    nFicSal = FreeFile
                    Open sLpt For Output As nFicSal
                    Print #nFicSal, lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                    Close #nFicSal
                Loop Until MsgBox("¿Desea reimprimir el voucher?", vbQuestion + vbYesNo, "Aviso") = vbNo
                RegistrarAfiliacionSepelio = True
            Else
                MsgBox "La Cuenta NO posee saldo suficiente", vbInformation, "Aviso"
                Call cmdBuscarCtaSep_Click
                RegistrarAfiliacionSepelio = False
            End If
            Set oNCapMov = Nothing
        Else
            If RealizarPagoSinDebitoCta Then
            For index = 1 To feBeneficiarios.Rows - 1
                Dim nPart As Integer
                
                If Len(feBeneficiarios.TextMatrix(index, 3)) = 4 Then
                    nPart = Left(feBeneficiarios.TextMatrix(index, 3), 3)
                ElseIf Len(feBeneficiarios.TextMatrix(index, 3)) = 3 Then
                    nPart = Left(feBeneficiarios.TextMatrix(index, 3), 2)
                Else
                    nPart = Left(feBeneficiarios.TextMatrix(index, 3), 1)
                End If
                If optBeneficiarioSi.value = True Then
                    Call oNSegSep.RegistraAseguradoSepelio(Trim(txtCertificadoSepelio.Text), feBeneficiarios.TextMatrix(index, 1), Right(feBeneficiarios.TextMatrix(index, 2), 3), nPart, 1)
                End If
            Next
            RegistrarAfiliacionSepelio = True
            End If
        End If
    Else
        '***********APRI 20170612 SEGUN SATI TIC1706060013*********************
        Set oDSegTar = New COMDCaptaGenerales.DCOMSeguros
        If oDSegTar.ValidaExisteRegistroNroCertificado(Trim(txtCertificadoSepelio.Text), gSegCertificadoSepelio) Then
            MsgBox "El Número de Certificado ya fue registrado anteriormente", vbInformation, "Aviso"
            RegistrarAfiliacionSepelio = False
            Exit Function
        End If
            
        Set oDSegTar = Nothing
        '**********END APRI20170612********************
        
        If RealizarPagoSinDebitoCta Then
            If optBeneficiarioSi.value = True Then
                For index = 1 To feBeneficiarios.Rows - 1
                    Call oNSegSep.RegistraAseguradoSepelio(Trim(txtCertificadoSepelio.Text), feBeneficiarios.TextMatrix(index, 1), Right(feBeneficiarios.TextMatrix(index, 2), 3), Mid(feBeneficiarios.TextMatrix(index, 3), 1, Len(feBeneficiarios.TextMatrix(index, 3)) - 1), 1)
                Next
            End If
            RegistrarAfiliacionSepelio = True
        End If
    End If
End Function
Private Function VerificaAfiliacionActiva(ByVal psFecha As String, ByVal psPersCod As String) As Boolean
    Dim oNCOMSeg As New COMNCaptaGenerales.NCOMSeguros
    Dim rsDatos As New ADODB.Recordset
    
    VerificaAfiliacionActiva = False
    Set rsDatos = oNCOMSeg.ObtieneClienteAfiliacionsepelio(psFecha, psPersCod)
    If Not (rsDatos.EOF And rsDatos.BOF) Then
        nEstadoSegSepelio = rsDatos!nEstado
        If rsDatos!nEstado = "501" Then 'APRI2180201 ERS028-2017
        'If rsDatos!nEstado = gSegEstadoAfiliado Then
            VerificaAfiliacionActiva = True
        End If
    End If
End Function
Private Sub LimpiarFormulario()
    nEstadoSegSepelio = 0
    ActXCodCtaTarj.NroCuenta = ""
    ActXCodCtaSepl.NroCuenta = ""
    fsPersCod = ""
    feAsegurado.Clear
    FormateaFlex feAsegurado
    feBeneficiarios.Clear
    FormateaFlex feBeneficiarios
    txtCertificadoSepelio.Text = ""
    txtCertificadoTarjeta.Text = ""
    feAsegurado.AdicionaFila
End Sub

Private Function RealizarPagoSinDebitoCta() As Boolean
    Dim clsCapMov As New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCont As New COMNContabilidad.NCOMContFunciones
    Dim clsCapM As New COMDCaptaGenerales.DCOMCaptaMovimiento
    Dim ClsMov As New COMDMov.DCOMMov
    Dim clsSegN As New COMNCaptaGenerales.NCOMSeguros
    Dim oNSegSep As New COMNCaptaGenerales.NCOMSeguros
    Dim rsValidaExiste As ADODB.Recordset
    Dim lnMovNro As Long
    Dim lsOpeCod As String
    Dim lnMonto As Currency
    Dim Moneda As String
    Dim lsMov As String
    
    On Error GoTo Error
        fsPersCod = feAsegurado.TextMatrix(1, 1)
        Set rsValidaExiste = clsSegN.DevulveClienteSeguroActivo(fsPersCod)
        If Not (rsValidaExiste.EOF And rsValidaExiste.BOF) Then
            MsgBox "El cliente ya tiene un seguro afiliado", vbInformation, "Alerta"
            Call LimpiarFormulario
            Exit Function
        End If
        lnMonto = txtPrimaSepl.Text
        lsMov = FechaHora(gdFecSis)
        Dim lbBan As Boolean
        
        lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
        If MsgBox("¿Desea Grabar la Información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            Dim lnMoneda As String
            Dim nTC As Double
            
            lnMovNro = clsCapMov.OtrasOperaciones(lsMov, gServActSepelioManual, lnMonto, "", "Afiliación Efectivo Seguro Sepelio", gMonedaNacional, fsPersCod, , , , , , , gnMovNro)
            clsCapM.AgregaSegSepelioAfiliacion txtCertificadoSepelio.Text, "", gdFecSis, lsMov, lnMovNro, feAsegurado.TextMatrix(1, 1), gsCodAge, ChkPrimaAnualSS.value
            'APRI20181023 ERS071-2018 ADD ChkPrimaAnualSS.value
            clsCapM.AgregaSegSepelioAfiliacionhis txtCertificadoSepelio.Text, "", gdFecSis, lsMov, lnMovNro, feAsegurado.TextMatrix(1, 1), gsCodAge, 501 'APRI20171024 ERS028-2017 501
            
            
            Dim index As Integer
            If optBeneficiarioSi.value = True Then
                For index = 1 To feBeneficiarios.Rows - 1
                    Call oNSegSep.RegistraAseguradoSepelio(Trim(txtCertificadoSepelio.Text), feBeneficiarios.TextMatrix(index, 1), Right(feBeneficiarios.TextMatrix(index, 2), 3), Mid(feBeneficiarios.TextMatrix(index, 3), 1, Len(feBeneficiarios.TextMatrix(index, 3)) - 1), 1)
                Next
            End If
            If gnMovNro = 0 Then
                MsgBox "La operación no se realizó, favor intentar nuevamente", vbInformation, "Aviso"
                Exit Function
            End If
        
        Set clsCont = Nothing
        Set clsCapMov = Nothing
        Set ClsMov = Nothing
        
        Dim oBol As New COMNCaptaGenerales.NCOMCaptaImpresion
        
        Dim lsBoleta As String
        lsBoleta = oNSegSep.ImprimeBoletaAfilicacionSeguroSepelio(lnMovNro, lsMov, gsNomAge, gbImpTMU)
        Set oBol = Nothing
        
        Dim oBITF As New COMNCaptaGenerales.NCOMCaptaMovimiento
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
        Set oBol = Nothing
        Unload Me
    End If
  
    Exit Function
Error:
      MsgBox Str(err.Number) & err.Description
End Function

Private Function ValidaDatosBeneficiario() As String
    Dim nIndice As Integer
    Dim nConyugeDup As Integer
    If optBeneficiarioSi.value = True Then
        For nIndice = 1 To feBeneficiarios.Rows - 1
            'APRI20171016 MEJORA PENDIENTE YIHU-06082017
            If feBeneficiarios.TextMatrix(nIndice, 1) = "" Or feBeneficiarios.TextMatrix(nIndice, 2) = "" Or feBeneficiarios.TextMatrix(nIndice, 3) = "" Then
                ValidaDatosBeneficiario = "No se aceptan datos vacíos en la sección de beneficiario"
            End If
            If feBeneficiarios.TextMatrix(nIndice, 3) <> "" And Len(feBeneficiarios.TextMatrix(nIndice, 3)) = 2 Then
                If Left(feBeneficiarios.TextMatrix(nIndice, 3), 1) = 0 Then
                    ValidaDatosBeneficiario = "No se acepta el valor 0 en el campo [Participación] para los beneficiarios."
                End If
            End If
'            If feBeneficiarios.TextMatrix(nIndice, 1) = "" Then ValidaDatosBeneficiario = "No se aceptan datos vacíos en la sección de beneficiario"
'            If feBeneficiarios.TextMatrix(nIndice, 2) = "" Then ValidaDatosBeneficiario = "No se aceptan datos vacíos en la sección de beneficiario"
        Next
    End If
End Function

Private Sub optBeneficiarioNo_Click()
    Call HabilitaControlGrid(feBeneficiarios, False, 2)
End Sub

Private Sub optBeneficiarioSi_Click()
    Call HabilitaControlGrid(feBeneficiarios, True, 2)
End Sub

Private Sub HabilitaControlGrid(ByVal poFlex As FlexEdit, ByVal pbValor As Boolean, ByVal pnOpe As Integer)
    poFlex.Enabled = pbValor
    If pnOpe = 2 Then
        cmdAgregar.Enabled = pbValor
        cmdQuitar.Enabled = pbValor
    End If
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
'COMENTADO APRI20171005 ERS028-2017
'        ElseIf rs!nCriterioID = 4 Then
'            DerArrep(0) = rs!nAnio: DerArrep(1) = rs!nMes: DerArrep(2) = rs!nDia
        End If
        rs.MoveNext
    Next
End Sub

Private Function ValidaCriterios(ByVal pnEdadPers As Integer) As String
    ValidaCriterios = ""
    If pnEdadPers < MinIng(0) Then
        ValidaCriterios = "La persona no puede ser menor de " & MinIng(0) & " años"
    End If
    If pnEdadPers > MaxIng(0) Then
        ValidaCriterios = "La persona no puede ser mayor de " & MaxIng(0) & " años"
    End If
End Function
