VERSION 5.00
Begin VB.Form frmOpeComisionCred 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6255
   Icon            =   "frmOpeComisionCred.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
      Enabled         =   0   'False
      Height          =   360
      Left            =   3600
      TabIndex        =   4
      Top             =   3000
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   4920
      TabIndex        =   5
      Top             =   3000
      Width           =   1170
   End
   Begin VB.Frame Frame2 
      Caption         =   " Tipo de Cobro "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   2040
      Width           =   6015
      Begin VB.ComboBox cboTipoCobro 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label lblLabelMonto 
         Caption         =   "Comisión S/:"
         Height          =   255
         Left            =   3840
         TabIndex        =   12
         Top             =   420
         Width           =   975
      End
      Begin VB.Label lblComision 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   4920
         TabIndex        =   3
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "..."
      Height          =   330
      Left            =   3840
      TabIndex        =   1
      Top             =   150
      Width           =   375
   End
   Begin VB.Frame Frame1 
      Caption         =   " Cliente "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   6
      Top             =   600
      Width           =   6015
      Begin VB.Label lblCliente 
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Width           =   4935
      End
      Begin VB.Label lblDOI 
         Alignment       =   2  'Center
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   840
         TabIndex        =   9
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label8 
         Caption         =   "Nombre:"
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
         TabIndex        =   8
         Top             =   400
         Width           =   615
      End
      Begin VB.Label Label9 
         Caption         =   "D.O.I.:"
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
         TabIndex        =   7
         Top             =   890
         Width           =   615
      End
   End
   Begin SICMACT.ActXCodCta ActXCodCta 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   873
      Texto           =   "Crédito"
      EnabledCMAC     =   -1  'True
      EnabledCta      =   -1  'True
      EnabledProd     =   -1  'True
      EnabledAge      =   -1  'True
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   360
      Left            =   5040
      TabIndex        =   13
      Top             =   2520
      Width           =   1050
   End
End
Attribute VB_Name = "frmOpeComisionCred"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmOpeComisionCred
'** Descripción : Formulario para pago de comisiones diversas de Créditos según TI-ERS012-2015
'** Creación : JUEZ, 20151229 09:00:00 AM
'*****************************************************************************************************
Option Explicit
Dim fsOpeCod As CaptacOperacion
Dim fsPersCod As String
Dim fnCodParam As Integer
Dim fnValorCom As Double

Public Sub Inicia(ByVal psOpeCod As CaptacOperacion, ByVal psOpeDesc As String, ByVal pnCodParam As Integer)
    fsOpeCod = psOpeCod
    fnCodParam = pnCodParam
    Me.Caption = "Comisión " & psOpeDesc
    ActXCodCta.CMAC = gsCodCMAC
    ActXCodCta.Age = gsCodAge
    CargaTipoCobro
    CargaMontoComision
    Me.Show 1
End Sub

Private Sub CargaMontoComision()
Dim oPar As COMDColocPig.DCOMColPCalculos

    Set oPar = New COMDColocPig.DCOMColPCalculos
        fnValorCom = oPar.dObtieneColocParametro(fnCodParam)
    Set oPar = Nothing

    lblComision.Caption = Format(fnValorCom, "#,##0.00")
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    If Len(ActXCodCta.NroCuenta) = 18 Then
        Dim oCred As COMDCredito.DCOMCredito
        Set oCred = New COMDCredito.DCOMCredito
        If fsOpeCod = gComiCredDupCronograma Then
                CargarDatos
                ActXCodCta.Enabled = False
                cmdBuscar.Enabled = False
                cmdGenerar.Enabled = True
                cboTipoCobro.SetFocus
        Else
            CargarDatos
        End If
    End If
End Sub

Private Sub cboTipoCobro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub

Private Sub cboTipoCobro_LostFocus()
    If Trim(Right(cboTipoCobro.Text, 2)) = "1" Then
        lblComision.Caption = Format(fnValorCom, "#,##0.00")
    ElseIf Trim(Right(cboTipoCobro.Text, 2)) = "2" Then
        lblComision.Caption = "0.00"
    End If
End Sub

Private Sub cmdGenerar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII
    
    If MsgBox("Desea Grabar la Información?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    Dim oCred As COMDCredito.DCOMCredActBD
    Set oCred = New COMDCredito.DCOMCredActBD
    Dim oCredMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oCredMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCont As COMNContabilidad.NCOMContFunciones
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    Dim lsMov As String
    Dim lsBoleta As String
    Dim lsGlosa As String
    Dim lsTitVoucher As String
    
    If fsOpeCod = gComiCredDupCronograma Then
        lsGlosa = "Duplicado de Cronograma: " & CStr(ActXCodCta.NroCuenta)
        lsTitVoucher = "DUPLICADO DE CRONOGRAMA"
    End If
    
    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    gnMovNro = 0
    If Trim(Right(cboTipoCobro.Text, 2)) = "1" Then
        Call oCredMov.OtrasOperaciones(lsMov, fsOpeCod, CDbl(lblComision.Caption), ActXCodCta.NroCuenta, lsGlosa & " cliente: " & lblCliente.Caption, gMonedaNacional, fsPersCod, , , , , , , gnMovNro)
        If gnMovNro <> 0 Then
            Call oCred.dInsertComision(gnMovNro, ActXCodCta.NroCuenta, CDbl(lblComision.Caption), 0, 0)
            Set oCred = Nothing
            Dim oBol As COMNCredito.NCOMCredDoc
            Set oBol = New COMNCredito.NCOMCredDoc
                lsBoleta = oBol.ImprimeBoletaComision(lsTitVoucher, Left("Total pago comision", 36), "", Str(CDbl(lblComision.Caption)), lblCliente.Caption, lblDOI.Caption, "________" & Mid(ActXCodCta.NroCuenta, 9, 1), False, "", "", , gdFecSis, gsNomAge, gsCodUser, sLpt, , gbImpTMU)
            Set oBol = Nothing
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
            
            If fsOpeCod = gComiCredDupCronograma Then
                ImprimeCronogramaPagos
            End If
            
            cmdCancelar_Click
            'INICIO JHCU ENCUESTA 16-10-2019
            Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
            'FIN

        Else
            MsgBox "Hubo un error en el registro", vbInformation, "Aviso"
        End If
    ElseIf Trim(Right(cboTipoCobro.Text, 2)) = "2" Then
        Dim loVistoElectronico As frmVistoElectronico
        Dim lbVistoVal As Boolean
        Set loVistoElectronico = New frmVistoElectronico
        lbVistoVal = loVistoElectronico.inicio(9, "")
        If lbVistoVal Then
            If fsOpeCod = gComiCredDupCronograma Then
                ImprimeCronogramaPagos
            End If
            cmdCancelar_Click
        Else
            Exit Sub
        End If
    End If
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona
    Limpiar
    Set oPers = frmBuscaPersona.inicio()
    If Not oPers Is Nothing Then
        Call FrmVerCredito.inicio(oPers.sPersCod, , , True, ActXCodCta)
        ActXCodCta.SetFocusCuenta
    End If
    Set oPers = Nothing
End Sub

Private Sub CargarDatos()
    Dim oCred As COMDCredito.DCOMCredito
    Dim R As ADODB.Recordset
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.RecuperaDatosComision(ActXCodCta.NroCuenta, 1)
    If Not (R.EOF And R.BOF) Then
        lblCliente.Caption = R!cPersNombre
        lblDOI.Caption = R!cPersIDnro
        fsPersCod = R!cPersCod
    
        Set R = Nothing
        cmdGenerar.Enabled = True
        cmdGenerar.SetFocus
    Else
        MsgBox "No existen Datos.", vbInformation, "Aviso"
        Limpiar
    End If
End Sub

Private Sub Limpiar()
    ActXCodCta.Prod = ""
    ActXCodCta.Cuenta = ""
    lblCliente.Caption = ""
    lblDOI.Caption = ""
    cboTipoCobro.ListIndex = -1
    CargaMontoComision
    cmdGenerar.Enabled = False
    ActXCodCta.Enabled = True
    cmdBuscar.Enabled = True
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    ActXCodCta.SetFocusProd
End Sub

Private Sub CargaTipoCobro()
Dim oDConst As COMDConstantes.DCOMConstantes
Dim rs As ADODB.Recordset

    Set oDConst = New COMDConstantes.DCOMConstantes
        Set rs = oDConst.RecuperaConstantes(3041)
    Set oDConst = Nothing
    
    cboTipoCobro.Clear
    Do While Not rs.EOF
        cboTipoCobro.AddItem Trim(rs!cConsDescripcion) & Space(200) & Trim(rs!nConsValor)
        rs.MoveNext
    Loop
    rs.Close
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub ImprimeCronogramaPagos()
Dim oCredDoc As COMNCredito.NCOMCredDoc
Dim oCred As COMDCredito.DCOMCredito
Dim Prev As previo.clsprevio
Dim rsDatos As ADODB.Recordset
Dim sCadImp As String
    Set oCred = New COMDCredito.DCOMCredito
        Set rsDatos = oCred.RecuperaDatosComunes(ActXCodCta.NroCuenta)
    Set oCred = Nothing
    
    Set oCredDoc = New COMNCredito.NCOMCredDoc
        sCadImp = oCredDoc.ImprimePlandePagosDuplicado(ActXCodCta.NroCuenta, gsNomAge, Format(gdFecSis, "dd/mm/yyyy"), gsCodUser, rsDatos!nMontoCol, IIf(IsNull(rsDatos!bMiVivienda), 0, rsDatos!bMiVivienda), , gsNomCmac, IIf(IsNull(rsDatos!bCuotaCom), 0, rsDatos!bCuotaCom), IIf(IsNull(rsDatos!nCalendDinamico), 0, rsDatos!nCalendDinamico), IIf(IsNull(rsDatos!bMiVivienda), 0, rsDatos!bMiVivienda), True)
    Set oCredDoc = Nothing
    
    Set Prev = New clsprevio
        Prev.Show oImpresora.gPrnTpoLetraSansSerif1PDef & sCadImp & oImpresora.gPrnTamLetra10CPIDef, ""
    Set Prev = Nothing
End Sub
