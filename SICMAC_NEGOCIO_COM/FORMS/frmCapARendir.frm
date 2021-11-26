VERSION 5.00
Begin VB.Form frmCapARendir 
   ClientHeight    =   3525
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6285
   Icon            =   "frmCapARendir.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3525
   ScaleWidth      =   6285
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraMoneda 
      Caption         =   "Moneda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   915
      Left            =   3960
      TabIndex        =   13
      Top             =   1920
      Width           =   2175
      Begin VB.OptionButton optMoneda 
         Caption         =   "Moneda &Extranjera"
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
         Index           =   1
         Left            =   180
         TabIndex        =   15
         Top             =   540
         Width           =   1695
      End
      Begin VB.OptionButton optMoneda 
         Caption         =   "Moneda &Nacional"
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
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   240
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   1
      Top             =   3000
      Width           =   1095
   End
   Begin VB.CommandButton cmdDesembolsar 
      Caption         =   "&Desembolsar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3840
      TabIndex        =   0
      Top             =   3000
      Width           =   1095
   End
   Begin VB.Frame Frame2 
      Caption         =   "Datos de Víaticos"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1455
      Left            =   120
      TabIndex        =   12
      Top             =   1920
      Width           =   3735
      Begin VB.TextBox txtMontoDesembolsar 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   18
         Top             =   960
         Width           =   1335
      End
      Begin VB.TextBox txtFechaSolicitud 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   600
         Width           =   1335
      End
      Begin VB.TextBox txtNumeroSolicitud 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   2280
         Locked          =   -1  'True
         TabIndex        =   16
         Top             =   240
         Width           =   1335
      End
      Begin VB.Label lblMontoDesembolsar 
         Caption         =   "Monto a Desembolsar S/."
         Height          =   375
         Left            =   120
         TabIndex        =   21
         Top             =   960
         Width           =   1815
      End
      Begin VB.Label Label6 
         Caption         =   "Fecha Apr. Solicitud"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label5 
         Caption         =   "N° de Solicitud"
         Height          =   375
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Datos del Colaborador"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1815
      Left            =   120
      TabIndex        =   7
      Top             =   0
      Width           =   6015
      Begin VB.TextBox txtMovNro 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4320
         TabIndex        =   22
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtCodigoAgencia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   5
         Top             =   1320
         Width           =   710
      End
      Begin VB.TextBox txtCodigoArea 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   960
         Width           =   710
      End
      Begin VB.TextBox txtNombreAgencia 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   1320
         Width           =   4335
      End
      Begin VB.TextBox txtNombreArea 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   1560
         Locked          =   -1  'True
         TabIndex        =   4
         Top             =   960
         Width           =   4335
      End
      Begin VB.TextBox txtNombreColaborador 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   840
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   600
         Width           =   5055
      End
      Begin SICMACT.TxtBuscar txtCodigoColaborador 
         Height          =   350
         Left            =   840
         TabIndex        =   23
         Top             =   240
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label Label2 
         Caption         =   "Área:"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Label4 
         Caption         =   "Código:"
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Agencia:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Nombre:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCapARendir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'************************************************************************************************
'***Nombre      : frmCapARendir
'***Descripción : Formulario para Desembolsar las Solicitudes de los Váticos y Otros Gastos A Rendir.
'***Creación    : ELRO el 20120423, según OYP-RFC005-2012 y OYP-RFC016-2012
'************************************************************************************************


Public Sub iniciarDesembolso(ByVal pnOpeCod As CaptacOperacion, ByVal psTitulo As String)
 If pnOpeCod = gOtrOpeDesParGas Then
    Me.Caption = psTitulo
 ElseIf pnOpeCod = gOtrOpeDesParVia Then
    Me.Caption = psTitulo
 End If
Show 1
End Sub


Private Sub limpiarCampos()
    txtNombreColaborador = ""
    txtCodigoArea = ""
    txtNombreArea = ""
    txtCodigoAgencia = ""
    txtNombreAgencia = ""
    txtMovNro = ""
    txtNumeroSolicitud = ""
    txtFechaSolicitud = "__/__/____"
    txtMontoDesembolsar = "0.00"
    
End Sub

Private Sub imprimirBoleta(ByVal psBoleta As String, ByVal psDJ As String, Optional ByVal sMensaje As String = "Boleta Operación y DJ")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, psBoleta & psDJ
    Print #nFicSal, psBoleta & psDJ
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub


Private Function validarDatosColaborador() As Boolean

validarDatosColaborador = False

If txtCodigoColaborador = gsCodPersUser Then
    MsgBox "No puedes realizar esta operación con prorpio usuario ", vbInformation, "Aviso"
    cmdDesembolsar.Enabled = False
    txtCodigoColaborador = ""
    txtCodigoColaborador.SetFocus
    Exit Function
End If

If txtCodigoColaborador = "" Then
    MsgBox "Falta ingresar el código de colaborador", vbInformation, "Aviso"
    cmdDesembolsar.Enabled = False
    txtCodigoColaborador = ""
    txtCodigoColaborador.SetFocus
    Exit Function
End If

validarDatosColaborador = True
End Function

Private Sub cargarDatosColaboradorViaticosParaDesembolsar()
Dim oNCOMCajaGeneral As COMNCajaGeneral.NCOMCajaGeneral
Set oNCOMCajaGeneral = New COMNCajaGeneral.NCOMCajaGeneral
Dim rsViaticosParaDesembolsar As ADODB.Recordset
Set rsViaticosParaDesembolsar = New ADODB.Recordset

If validarDatosColaborador = False Then
    Exit Sub
End If

Call limpiarCampos

Set rsViaticosParaDesembolsar = oNCOMCajaGeneral.obtenerAprobacionARendirViaticosParaDesembolsar(IIf(optMoneda.iTem(0), "1", "2"), txtCodigoColaborador.psCodigoPersona)

If Not rsViaticosParaDesembolsar.BOF And Not rsViaticosParaDesembolsar.EOF Then

    txtNombreColaborador = rsViaticosParaDesembolsar!cPersNombre
    txtCodigoArea = rsViaticosParaDesembolsar!cAreaCod
    txtNombreArea = rsViaticosParaDesembolsar!cAreaDescripcion
    txtCodigoAgencia = rsViaticosParaDesembolsar!cAgeCod
    txtNombreAgencia = rsViaticosParaDesembolsar!cAgeDescripcion
    txtNumeroSolicitud = rsViaticosParaDesembolsar!cDocNro
    txtFechaSolicitud = rsViaticosParaDesembolsar!cFecApro
    txtMontoDesembolsar = Format(rsViaticosParaDesembolsar!nImporte, "#,##0.00")
    txtMovNro = rsViaticosParaDesembolsar!nMovNro
    
    If cmdDesembolsar.Enabled = False Then
        cmdDesembolsar.Enabled = True
    End If
    
    If optMoneda.iTem(0).value Then
        lblMontoDesembolsar = "Monto a Desembolsar S/."
        txtMontoDesembolsar.BackColor = &H80000005
    Else
        lblMontoDesembolsar = "Monto a Desembolsar $."
        txtMontoDesembolsar.BackColor = &HC0FFC0
    End If
         
Else
    MsgBox "No tiene Solicitud de Viaticos pendientes a desembolsar " & IIf(optMoneda.iTem(0).value, "en Moneda Nacional", "en Moneda Extranjera"), vbInformation, "Aviso"
    cmdDesembolsar.Enabled = False
    txtCodigoColaborador = ""
    txtCodigoColaborador.SetFocus
    
End If

Set oNCOMCajaGeneral = Nothing
Set rsViaticosParaDesembolsar = Nothing
End Sub

Private Sub cargarDatosColaboradorARendirCuentasParaDesembolsar()
Dim oNCOMCajaGeneral As COMNCajaGeneral.NCOMCajaGeneral
Set oNCOMCajaGeneral = New COMNCajaGeneral.NCOMCajaGeneral
Dim rsARendirCuentasParaDesembolsar As ADODB.Recordset
Set rsARendirCuentasParaDesembolsar = New ADODB.Recordset

If validarDatosColaborador = False Then
    Exit Sub
End If

Call limpiarCampos

Set rsARendirCuentasParaDesembolsar = oNCOMCajaGeneral.obtenerAprobacionARendirCuentasParaDesembolsar(IIf(optMoneda.iTem(0), "1", "2"), txtCodigoColaborador.psCodigoPersona)

If Not rsARendirCuentasParaDesembolsar.BOF And Not rsARendirCuentasParaDesembolsar.EOF Then

    txtNombreColaborador = rsARendirCuentasParaDesembolsar!cPersNombre
    txtCodigoArea = rsARendirCuentasParaDesembolsar!cAreaCod
    txtNombreArea = rsARendirCuentasParaDesembolsar!cAreaDescripcion
    txtCodigoAgencia = rsARendirCuentasParaDesembolsar!cAgeCod
    txtNombreAgencia = rsARendirCuentasParaDesembolsar!cAgeDescripcion
    txtNumeroSolicitud = rsARendirCuentasParaDesembolsar!cDocNro
    txtFechaSolicitud = rsARendirCuentasParaDesembolsar!cFecApro
    txtMontoDesembolsar = Format(rsARendirCuentasParaDesembolsar!nImporte, "#,##0.00")
    txtMovNro = rsARendirCuentasParaDesembolsar!nMovNro
    
    If cmdDesembolsar.Enabled = False Then
        cmdDesembolsar.Enabled = True
    End If
    
    If optMoneda.iTem(0).value Then
        lblMontoDesembolsar = "Monto a Desembolsar S/."
        txtMontoDesembolsar.BackColor = &H80000005
    Else
        lblMontoDesembolsar = "Monto a Desembolsar US$."
        txtMontoDesembolsar.BackColor = &HC0FFC0
    End If
         
Else
    MsgBox "No tiene Solicitud de Viaticos pendientes a desembolsar " & IIf(optMoneda.iTem(0).value, "en Moneda Nacional", "en Moneda Extranjera"), vbInformation, "Aviso"
    cmdDesembolsar.Enabled = False
    txtCodigoColaborador = ""
    txtCodigoColaborador.SetFocus
    
End If

Set oNCOMCajaGeneral = Nothing
Set rsARendirCuentasParaDesembolsar = Nothing
End Sub

Private Sub cmdDesembolsar_Click()
    Dim oNCOMCaptaMovimiento As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oNCOMCaptaMovimiento = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Dim lsMovNro As String, lsBoleta As String, lsDJ As String
    
    If MsgBox("¿Desea Grabar la Operación?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Set oNCOMCaptaMovimiento = Nothing
        Set oNCOMContFunciones = Nothing
        Exit Sub
    Else
        oNCOMCaptaMovimiento.IniciaImpresora gImpresora
        lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, _
                                                   gsCodAge, _
                                                   gsCodUser)
        
        
        If gsOpeCod = gOtrOpeDesParGas Then
            Call oNCOMCaptaMovimiento.registrarDesembolsoParaARendirCuentas(CCur(txtMontoDesembolsar), _
                                                                            CLng(txtMovNro), _
                                                                            gsOpeCod, _
                                                                            lsMovNro, _
                                                                            CDate(txtFechaSolicitud), _
                                                                            "Desembolso por A Rendir Cuentas", _
                                                                            txtNombreColaborador, _
                                                                            gsNomAge, _
                                                                            IIf(optMoneda(0).value = True, Moneda.gMonedaNacional, Moneda.gMonedaExtranjera), _
                                                                            lsBoleta, _
                                                                            lsDJ, _
                                                                            sLpt, _
                                                                            gbImpTMU)
        ElseIf gsOpeCod = gOtrOpeDesParVia Then
            Call oNCOMCaptaMovimiento.registrarDesembolsoParaViatico(CCur(txtMontoDesembolsar), _
                                                                     CLng(txtMovNro), _
                                                                     gsOpeCod, _
                                                                     lsMovNro, _
                                                                     CDate(txtFechaSolicitud), _
                                                                     "Desembolso por Viáticos", _
                                                                     txtNombreColaborador, _
                                                                     gsNomAge, _
                                                                     IIf(optMoneda(0).value = True, Moneda.gMonedaNacional, Moneda.gMonedaExtranjera), _
                                                                     lsBoleta, _
                                                                     lsDJ, _
                                                                     sLpt, _
                                                                     gbImpTMU)
   
        End If
        
        If Trim(lsBoleta) <> "" Then
            Call imprimirBoleta(lsBoleta, lsDJ)
        Else
            MsgBox "No se realizo la operación", vbInformation, "Aviso"
        End If
    End If
    

           
    Set oNCOMCaptaMovimiento = Nothing
    Set oNCOMContFunciones = Nothing
    lsMovNro = ""
    lsBoleta = ""
    cmdSalir_Click
End Sub

Private Sub cmdSalir_Click()
    Call limpiarCampos
    Unload Me
End Sub

Private Sub Form_Load()
    If optMoneda.iTem(0).value Then
        lblMontoDesembolsar = "Monto a Desembolsar S/."
        txtMontoDesembolsar.BackColor = &H80000005
    Else
        lblMontoDesembolsar = "Monto a Desembolsar US$."
        txtMontoDesembolsar.BackColor = &HC0FFC0
    End If
End Sub

Private Sub optMoneda_Click(Index As Integer)
    If gsOpeCod = gOtrOpeDesParGas Then
        Call cargarDatosColaboradorARendirCuentasParaDesembolsar
    ElseIf gsOpeCod = gOtrOpeDesParVia Then
        Call cargarDatosColaboradorViaticosParaDesembolsar
    End If
End Sub

Private Sub txtCodigoColaborador_EmiteDatos()
    If gsOpeCod = gOtrOpeDesParGas Then
        Call cargarDatosColaboradorARendirCuentasParaDesembolsar
    ElseIf gsOpeCod = gOtrOpeDesParVia Then
        Call cargarDatosColaboradorViaticosParaDesembolsar
    End If
End Sub

