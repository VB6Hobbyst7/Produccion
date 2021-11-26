VERSION 5.00
Begin VB.Form frmDepositoCajeroCorresponsal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Depósito Recaudo Cajero Corresponsal"
   ClientHeight    =   4395
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5475
   Icon            =   "frmDepositoCajeroCorresponsal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   5475
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   315
      Left            =   2625
      TabIndex        =   19
      Top             =   3975
      Width           =   855
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   315
      Left            =   3540
      TabIndex        =   20
      Top             =   3975
      Width           =   855
   End
   Begin VB.TextBox txtMontoPagarCC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3780
      MaxLength       =   18
      TabIndex        =   18
      Top             =   3555
      Width           =   1545
   End
   Begin VB.TextBox txtMontoCobradoCC 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3780
      Locked          =   -1  'True
      MaxLength       =   18
      TabIndex        =   17
      Top             =   3195
      Width           =   1545
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Datos del Cajero Corresponsal"
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
      Height          =   1995
      Left            =   150
      TabIndex        =   6
      Top             =   1095
      Width           =   5175
      Begin VB.TextBox txtUbicacionCC 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   14
         Top             =   1440
         Width           =   3015
      End
      Begin VB.TextBox txtDireccionCC 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   13
         Top             =   1080
         Width           =   3015
      End
      Begin VB.TextBox txtDoiCC 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   720
         Width           =   3015
      End
      Begin VB.TextBox txtOperadorCC 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1890
         Locked          =   -1  'True
         TabIndex        =   11
         Top             =   360
         Width           =   3015
      End
      Begin VB.Label Label6 
         Appearance      =   0  'Flat
         Caption         =   "Ubicación:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   1470
         Width           =   1170
      End
      Begin VB.Label Label5 
         Appearance      =   0  'Flat
         Caption         =   "Direc. del Establec."
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   240
         TabIndex        =   9
         Top             =   1110
         Width           =   1590
      End
      Begin VB.Label Label3 
         Appearance      =   0  'Flat
         Caption         =   "DOI"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   780
         Width           =   690
      End
      Begin VB.Label Label2 
         Appearance      =   0  'Flat
         Caption         =   "Operador:"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   420
         Width           =   690
      End
   End
   Begin VB.Frame pnlBusquedaConvenio 
      Appearance      =   0  'Flat
      Caption         =   "Búsqueda de Cajero Corresponsal"
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
      Height          =   885
      Left            =   150
      TabIndex        =   0
      Top             =   90
      Width           =   5175
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         Left            =   780
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   330
         Width           =   1590
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3150
         MaxLength       =   18
         TabIndex        =   4
         Top             =   330
         Width           =   1335
      End
      Begin VB.CommandButton btnBuscarCC 
         Appearance      =   0  'Flat
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   314
         Left            =   4500
         TabIndex        =   5
         Top             =   330
         Width           =   400
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         Caption         =   "Moneda"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   690
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         Caption         =   "Código: "
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   360
         Width           =   690
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   315
      Left            =   4470
      TabIndex        =   21
      Top             =   3975
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "S/."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3510
      TabIndex        =   23
      Top             =   3555
      Width           =   240
   End
   Begin VB.Label Label9 
      Caption         =   "S/."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   3510
      TabIndex        =   22
      Top             =   3240
      Width           =   240
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      Caption         =   "Monto a Pagar"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2220
      TabIndex        =   16
      Top             =   3590
      Width           =   1170
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      Caption         =   "Monto Cobrado"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   2220
      TabIndex        =   15
      Top             =   3245
      Width           =   1170
   End
End
Attribute VB_Name = "frmDepositoCajeroCorresponsal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'-------------------------------------------------------------------------------------------------
' NOMBRE         : "frmDepositoCajeroCorresponsal"
' DESCRIPCION    : Formulario creado para efectuar pagos de recaudacion de cajeros corresponsales
' CREACION       : RIRO, 20160109 10:00 AM
'-------------------------------------------------------------------------------------------------

Option Explicit

'declarando variables que se mostraran en el detalle
Private nCodCC As Long
Private sCodCC As String
Private nMontoRecaudado As Currency
Private sPersNombre As String
Private sPersCod As String
Private sDoi As String
Private sDireccion As String
Private sUbigeo As String
Private sNombreCC As String
Private sNombreComercialCC As String
Private sOpeCod As CaptacOperacion
'fin de declaracion de variables

Public Sub Inicio(ByVal psOpeCod As CaptacOperacion)
  sOpeCod = psOpeCod
  Me.Show 1
End Sub

Private Function Validacion() As String

    Dim sValidacion As String
    sValidacion = ""
    
    If Not IsNumeric(txtMontoPagarCC.Text) Then
        sValidacion = "El valor a pagar debe ser numérico" & vbNewLine
    End If
    If (nMontoRecaudado <> CCur(txtMontoPagarCC.Text)) Then
        sValidacion = sValidacion & "El monto a pagar debe ser el igual al monto recaudado" & vbNewLine
    End If
    If (CCur(txtMontoPagarCC.Text) <= 0) Then
        sValidacion = "El monto a pagar debe ser mayor que cero ""0.00""" & vbNewLine
    End If
    If (nCodCC <= 0 Or Trim(sCodCC) = "") Then
        sValidacion = sValidacion & "Antes de efectuar el pago, debe seleccionar un Cajero Corresponsal" & vbNewLine
    End If
    If Len(Trim(sValidacion)) = 0 Then
        If ObtenerMontoPago <> CCur(txtMontoPagarCC.Text) Then
            sValidacion = sValidacion & "Monto de pago es diferente al monto a pagar, se le recomienda volver a cargar los datos del cajero corresponsal" & vbNewLine
        End If
    End If
    
    Validacion = sValidacion
End Function

Private Sub btnBuscarCC_Click()
txtCodigo_KeyPress (13)
End Sub

Private Sub cmdCancelar_Click()
    If MsgBox("Desea cancelar el actual proceso de pago?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        LimpiarCC
        txtCodigo.SetFocus
    End If
End Sub

Private Sub ImprimeVoucher(ByVal pnMovNro As Long, _
                           ByVal psTitulo As String, _
                           ByVal psOperador As String, _
                           ByVal psCodOperador As String, _
                           ByVal psMontoV As Currency, _
                           ByVal pnMoneda As Integer)

'   Set clsMov = Nothing
    Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
    Dim lsBoleta As String
    
    Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
    lsBoleta = oBol.ImprimeBoleta(psTitulo, "Pago Recaudo", "", str(Format(CCur(txtMontoPagarCC.Text), "#0.00")), psOperador, "________" & Trim(Right(cboMoneda, 3)), psCodOperador, 0, "0", "Cod. Cajero", 0, 0, False, False, , , , False, , "Nro Ope. : " & str(pnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False)
    Set oBol = Nothing
    
    
    Do
       If Trim(lsBoleta) <> "" Then
            'lsBoleta = lsBoleta & oImpresora.gPrnSaltoLinea
            nFicSal = FreeFile
            Open sLpt For Output As nFicSal
                Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                Print #nFicSal, ""
            Close #nFicSal
      End If
    Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
    
End Sub

Private Sub cmdGrabar_Click()
    
    Dim lnMonto As Currency
    Dim lsMov As String
    Dim lsOpeCod As CaptacOperacion
    Dim lnMoneda As Integer
    Dim lnMovNro As Long
    Dim bResp As Boolean
    Dim sMensaje As String
    Dim sValidacion As String
    Dim clsCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim clsCont As COMNContabilidad.NCOMContFunciones
    Dim sBoleta As String

    ' variables de voucher
    Dim sTitulo As String
    Dim sOperador As String
    Dim sCodOperador As String
    Dim nMontoV As Currency
    Dim nMonedaV As Integer
    ' end variables de voucher

    On Error GoTo Error
    
    sValidacion = Validacion
    
    If Len(Trim(sValidacion)) > 0 Then
        MsgBox "Se presentaron las siguientes observaciones:" & vbNewLine & sValidacion, vbInformation, "Aviso"
        txtMontoPagarCC.SelStart = 0
        txtMontoPagarCC.SelLength = Len(txtMontoPagarCC.Text)
        txtMontoPagarCC.SetFocus
        Exit Sub
    End If
    Set clsCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set clsCont = New COMNContabilidad.NCOMContFunciones
    
    lnMonto = txtMontoPagarCC.Text
    lsMov = clsCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If MsgBox("Desea Grabar la Información", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        lnMoneda = CInt(Right(Me.cboMoneda.Text, 3))
        lsOpeCod = sOpeCod
        bResp = clsCapMov.PagoRecaudoCajeroCorresponsal(lsMov, lsOpeCod, nCodCC, sPersCod, lnMonto, lnMoneda, sCodCC, sMensaje, lnMovNro)
        If Not bResp Then
            If Trim(sMensaje) = "" Then
                sMensaje = "Se presentó un error error durante el proceso de pago en la Apliacación, favor de consultar con T.I."
            End If
            MsgBox sMensaje, vbInformation, "Aviso"
            Exit Sub
        Else
            MsgBox "El proceso de pago fue exitoso", vbInformation, "Aviso"
            clsCapMov.ObtieneDatosVoucherCajeroCorresponsal lnMovNro, sTitulo, sOperador, sCodOperador, nMontoV, nMonedaV
                        
            sBoleta = clsCapMov.ImprimeBoleta(sTitulo, "Cajero Corresponsal", gOtrOpePagoRecaudoCajeroCorresponsal, str(Format(nMontoV, "#0.00")), _
                                              sOperador, "________" & nMonedaV, sCodOperador, 0, "0", "Cod. Cajero", 0, 0, False, False, _
                                              , , , , , "Nro Ope. : " & str(lnMovNro), , CDate(Mid(lsMov, 7, 2) & "/" & Mid(lsMov, 5, 2) & "/" & Mid(lsMov, 1, 4)), _
                                              gsNomAge, Right(lsMov, 4), , , , , , , , , , , , , , , False, , 0, , , , gbImpTMU, , , sOperador)
  
            Do
               If Trim(sBoleta) <> "" Then
                    nFicSal = FreeFile
                    Open sLpt For Output As nFicSal
                        Print #nFicSal, oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & sBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
                        Print #nFicSal, ""
                    Close #nFicSal
              End If
            Loop While MsgBox("Desea Re Imprimir ?", vbQuestion + vbYesNo, "Aviso") = vbYes
                        
            LimpiarCC
            txtCodigo.SetFocus
        End If
        Set clsCont = Nothing
        Set clsCapMov = Nothing
    End If
    Exit Sub
Error:
      MsgBox "Se presentó un error: " & str(Err.Number) & Err.Description & vbNewLine & "Favor de verificar si se concretó el pago." & vbNewLine & "Se limpiará el formulario", vbInformation, "Aviso"
      LimpiarCC
      txtCodigo.SetFocus
End Sub

Private Sub cmdSalir_Click()
    If MsgBox("Desea salir del formulario de pago?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim clsCon As COMDConstantes.DCOMConstantes
    Set clsCon = New COMDConstantes.DCOMConstantes
    CargaCombo Me.cboMoneda, clsCon.RecuperaConstantes(gMoneda)
    Me.cboMoneda.ListIndex = 0
    Me.cboMoneda.Enabled = False
    txtMontoPagarCC.Text = 0
    txtCodigo.MaxLength = 13
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        ObtenerDatosCC
    Else
    KeyAscii = Letras(KeyAscii, True)
    End If
End Sub

Private Sub ObtenerDatosCC()
    
    Dim rsCC As ADODB.Recordset
    Dim clsMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim bResult As Boolean
    
    bResult = False
    sCodCC = Replace(Trim(txtCodigo.Text), "'", "")
    
    If Len(Trim(sCodCC)) = 0 Then
        MsgBox "Debe Ingresar un código de operador de cajero corresponsal", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set clsMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set rsCC = clsMov.ObtieneDatosCajeroCorresponsal(sCodCC)
    
    If Not rsCC Is Nothing Then
        If Not rsCC.EOF And Not rsCC.BOF Then
            If rsCC.RecordCount > 0 Then
                bResult = True
            End If
        End If
    End If
    If Not bResult Then
        MsgBox "El valor ingresado no corresponde a un cajero corresponsal", vbInformation, "Aviso"
        Set clsMov = Nothing
        Set rsCC = Nothing
        txtCodigo.Text = ""
        Exit Sub
    Else
        nCodCC = rsCC!nIdCC
        sCodCC = rsCC!cCodigoCC
        nMontoRecaudado = rsCC!nMontoRecaudo
        sPersNombre = rsCC!cPersNombre
        sPersCod = rsCC!cPersCod
        sDoi = rsCC!cDOI
        sDireccion = rsCC!cPersDireccCC
        sUbigeo = rsCC!cUbiGeoDescripcion
        sNombreCC = rsCC!cNombreCC
        sNombreComercialCC = rsCC!cNombreComercialCC
        
        txtOperadorCC.Text = sPersNombre
        txtDoiCC.Text = sDoi
        txtDireccionCC.Text = sDireccion
        txtUbicacionCC.Text = sUbigeo
        txtMontoCobradoCC.Text = nMontoRecaudado
        txtCodigo.Enabled = False
        btnBuscarCC.Enabled = False
        txtMontoPagarCC.SelStart = 0
        txtMontoPagarCC.SelLength = Len(txtMontoCobradoCC.Text)
        txtMontoPagarCC.SetFocus
        
    End If
End Sub

Private Function ObtenerMontoPago() As Currency
    Dim rsCC As ADODB.Recordset
    Dim clsMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim bResult As Boolean
    Dim nResp As Currency
    
    bResult = False
        
    Set clsMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Set rsCC = clsMov.ObtieneDatosCajeroCorresponsal(Replace(Trim(txtCodigo.Text), "'", ""))
    
    If Not rsCC Is Nothing Then
        If Not rsCC.EOF And Not rsCC.BOF Then
            If rsCC.RecordCount > 0 Then
                bResult = True
            End If
        End If
    End If
    If bResult Then
        nResp = rsCC!nMontoRecaudo
    End If
    Set clsMov = Nothing
    Set rsCC = Nothing
    ObtenerMontoPago = nResp
End Function


Private Sub LimpiarCC()
    nCodCC = 0
    sCodCC = ""
    nMontoRecaudado = 0
    sPersNombre = ""
    sPersCod = ""
    sDoi = ""
    sDireccion = ""
    sUbigeo = ""
    sNombreCC = ""
    sNombreComercialCC = ""
    
    txtCodigo.Text = ""
    txtOperadorCC.Text = sPersNombre
    txtDoiCC.Text = sDoi
    txtDireccionCC.Text = sDireccion
    txtUbicacionCC.Text = sUbigeo
    txtMontoCobradoCC.Text = nMontoRecaudado
    txtMontoPagarCC.Text = 0
    cboMoneda.ListIndex = 0
    txtCodigo.Enabled = True
    btnBuscarCC.Enabled = True
End Sub

Private Sub txtMontoPagarCC_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGrabar.SetFocus
    Else
        KeyAscii = NumerosDecimales(txtMontoPagarCC, KeyAscii, 8, 2, False)
    End If
    
End Sub

Private Sub txtMontoPagarCC_LostFocus()

    If IsNumeric(txtMontoPagarCC.Text) Then
        txtMontoPagarCC.Text = txtMontoPagarCC.Text * 1
    Else
        MsgBox "El valor ingresado en el Monto a Pagar debe ser numérico", vbInformation, "Aviso"
        txtMontoPagarCC.Text = 0
        txtMontoPagarCC.SelStart = 0
        txtMontoPagarCC.SelLength = Len(txtMontoPagarCC.Text)
        txtMontoPagarCC.SetFocus
    End If

End Sub




