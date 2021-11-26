VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmColPVentaLotePrepara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Crédito Pignoraticio - Preparación de Venta en Lote de Prendas Adjudicadas"
   ClientHeight    =   3390
   ClientLeft      =   975
   ClientTop       =   2385
   ClientWidth     =   6885
   Icon            =   "frmColPVentaLotePrepara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin MSComDlg.CommonDialog dlgGrabar 
      Left            =   4680
      Top             =   2640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "Subasta"
      Height          =   3105
      Left            =   120
      TabIndex        =   11
      Top             =   105
      Width           =   6675
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5430
         TabIndex        =   23
         Top             =   2130
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   2880
         TabIndex        =   8
         Top             =   1485
         Width           =   975
      End
      Begin VB.TextBox txtEstado 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   315
         Left            =   5295
         TabIndex        =   3
         Top             =   300
         Width           =   1200
      End
      Begin VB.TextBox txtNumVenta 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         Height          =   285
         Left            =   750
         TabIndex        =   0
         Top             =   315
         Width           =   645
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   5400
         TabIndex        =   10
         Top             =   1485
         Width           =   975
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   4200
         TabIndex        =   9
         Top             =   1485
         Width           =   975
      End
      Begin VB.Frame Frame4 
         Caption         =   "Precios del Oro "
         Height          =   600
         Left            =   135
         TabIndex        =   12
         Top             =   735
         Width           =   6375
         Begin VB.TextBox txtPreOro21 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   5430
            MaxLength       =   6
            TabIndex        =   7
            Top             =   210
            Width           =   750
         End
         Begin VB.TextBox txtPreOro18 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   3795
            MaxLength       =   6
            TabIndex        =   6
            Top             =   210
            Width           =   750
         End
         Begin VB.TextBox txtPreOro16 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   2250
            MaxLength       =   6
            TabIndex        =   5
            Top             =   225
            Width           =   750
         End
         Begin VB.TextBox txtPreOro14 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            Height          =   285
            Left            =   615
            MaxLength       =   6
            TabIndex        =   4
            Top             =   225
            Width           =   750
         End
         Begin VB.Label Label5 
            Caption         =   "14 K :"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   16
            Top             =   255
            Width           =   540
         End
         Begin VB.Label Label5 
            Caption         =   "16 K :"
            Height          =   225
            Index           =   2
            Left            =   1755
            TabIndex        =   15
            Top             =   255
            Width           =   525
         End
         Begin VB.Label Label5 
            Caption         =   "18 K :"
            Height          =   225
            Index           =   3
            Left            =   3315
            TabIndex        =   14
            Top             =   240
            Width           =   525
         End
         Begin VB.Label Label5 
            Caption         =   "21 K :"
            Height          =   225
            Index           =   4
            Left            =   4950
            TabIndex        =   13
            Top             =   240
            Width           =   540
         End
      End
      Begin MSMask.MaskEdBox txtFecVenta 
         Height          =   315
         Left            =   2115
         TabIndex        =   1
         Top             =   300
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtHorVenta 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "H:mm:ss"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   3082
            SubFormatType   =   4
         EndProperty
         Height          =   315
         Left            =   3825
         TabIndex        =   2
         Top             =   285
         Width           =   690
         _ExtentX        =   1217
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   5
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin RichTextLib.RichTextBox rtfImp 
         Height          =   360
         Left            =   6150
         TabIndex        =   21
         Top             =   2280
         Visible         =   0   'False
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   635
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmColPVentaLotePrepara.frx":030A
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtfCartas 
         Height          =   360
         Left            =   5325
         TabIndex        =   22
         Top             =   2280
         Visible         =   0   'False
         Width           =   450
         _ExtentX        =   794
         _ExtentY        =   635
         _Version        =   393217
         Enabled         =   -1  'True
         TextRTF         =   $"frmColPVentaLotePrepara.frx":038A
      End
      Begin MSComctlLib.ProgressBar prgList 
         Height          =   330
         Left            =   150
         TabIndex        =   24
         Top             =   2160
         Visible         =   0   'False
         Width           =   5190
         _ExtentX        =   9155
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Label Label3 
         Caption         =   "Estado :"
         Height          =   255
         Index           =   1
         Left            =   4665
         TabIndex        =   20
         Top             =   345
         Width           =   645
      End
      Begin VB.Label Label4 
         Caption         =   "Número :"
         Height          =   255
         Left            =   105
         TabIndex        =   19
         Top             =   330
         Width           =   660
      End
      Begin VB.Label Label3 
         Caption         =   "Hora :"
         Height          =   255
         Index           =   0
         Left            =   3375
         TabIndex        =   18
         Top             =   330
         Width           =   525
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   1545
         TabIndex        =   17
         Top             =   345
         Width           =   630
      End
   End
End
Attribute VB_Name = "frmColPVentaLotePrepara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'Archivo:  frmColPVentaLotePrepara.frm
'DAOR   :  30/01/2008.
'Resumen:  Nos permite ingresar o actualizar los precios del oro con
'          que van a ser procesados los listado y planillas para venta lote
'***************************************************************************
Option Explicit

Dim fnTasaImpuesto As Double
Dim fnTasaPreparacionRemate As Double
Dim fnTasaIGV As Double
Dim pVerCodAnt As Boolean

Dim fsVentaCadaAgencia As String
Dim fnJoyasDet As Integer

Dim vNumAvisos As Integer
Dim MuestraImpresion As Boolean
Dim vRTFImp As String, vBuffer As String
Dim vCont As Double
Dim vNomAge As String

Private Sub cmdAgencia_Click()
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub
Private Sub HabilitaControles(ByVal pbEditar As Boolean, ByVal pbGrabar As Boolean, ByVal pbSalir As Boolean, _
    ByVal pbCancelar As Boolean, ByVal pbFecVenta As Boolean, ByVal pbHorVenta As Boolean, _
    ByVal pbPreOro14 As Boolean, ByVal pbPreOro16 As Boolean, ByVal pbPreOro18 As Boolean, ByVal pbPreOro21 As Boolean, _
    ByVal pbImpAvisSuba As Boolean, ByVal pbImpPlanVenta As Boolean, ByVal pbImpListVenta As Boolean)

    cmdEditar.Enabled = pbEditar
    cmdGrabar.Enabled = pbGrabar
    cmdSalir.Enabled = pbSalir
    cmdCancelar.Enabled = pbCancelar
    txtFecVenta.Enabled = pbFecVenta
    txtHorVenta.Enabled = pbHorVenta
    txtPreOro14.Enabled = pbPreOro14
    txtPreOro16.Enabled = pbPreOro16
    txtPreOro18.Enabled = pbPreOro18
    txtPreOro21.Enabled = pbPreOro21
    
   
End Sub
'Permite no reconocer los datos ingresados
Private Sub cmdCancelar_Click()

On Error GoTo ControlError

Call HabilitaControles(True, False, True, False, False, False, False, False, False, False, True, False, False)
Limpiar
VeriDatRem
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Permite editar los campos editables de una Venta
Private Sub CmdEditar_Click()
    Call HabilitaControles(False, True, False, True, True, True, _
        True, True, True, True, False, False, False)
    txtFecVenta.SetFocus
End Sub

'permite grabar los cambios ingresados
Private Sub cmdGrabar_Click()

On Error GoTo ControlError
Dim loGrabSub As COMNColoCPig.NCOMColPRecGar

Call HabilitaControles(True, False, True, False, False, False, False, False, False, False, True, True, True)


Set loGrabSub = New COMNColoCPig.NCOMColPRecGar
    Call loGrabSub.nRecGarGrabaDatosPreparaCredPignoraticio("V", Me.txtNumVenta, gColPRecGarEstNoIniciado, Format(txtFecVenta.Text & " " & txtHorVenta.Text, "mm/dd/yyyy hh:mm"), fsVentaCadaAgencia, Val(Me.txtPreOro14), Val(Me.txtPreOro16), Val(Me.txtPreOro18), Val(Me.txtPreOro21), , , , False)
Set loGrabSub = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "

End Sub


Private Sub cmdSalir_Click()
Unload frmSelectAgencias
Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CargaParametros
    Limpiar
End Sub

Private Sub Limpiar()
Dim loDatos As COMNColoCPig.NCOMColPRecGar
Dim lrDatosVen As ADODB.Recordset
Dim lsMensaje As String
Dim lsUltVenta As String

    txtPreOro14 = Format(0, "#0.00")
    txtPreOro16 = Format(0, "#0.00")
    txtPreOro18 = Format(0, "#0.00")
    txtPreOro21 = Format(0, "#0.00")

    Set lrDatosVen = New ADODB.Recordset
    Set loDatos = New COMNColoCPig.NCOMColPRecGar
        lsUltVenta = loDatos.nObtieneNroUltimoProceso("V", fsVentaCadaAgencia, lsMensaje)
        If Trim(lsMensaje) <> "" Then
            MsgBox lsMensaje, vbInformation, "Aviso"
            Exit Sub
        End If
        Set lrDatosVen = loDatos.nObtieneDatosProcesoRGCredPig("V", lsUltVenta, fsVentaCadaAgencia, lsMensaje)
        If Trim(lsMensaje) <> "" Then
            MsgBox lsMensaje, vbInformation, "Aviso"
            Exit Sub
        End If
    Set loDatos = Nothing
    If lrDatosVen Is Nothing Then Exit Sub

    Me.txtNumVenta = lrDatosVen!cNroProceso
    Me.txtFecVenta = Format(lrDatosVen!dProceso, "dd/mm/yyyy")
    Me.txtHorVenta = Format(lrDatosVen!dProceso, "hh:mm")
    Me.txtPreOro14 = Format(lrDatosVen!nPrecioK14, "#0.00")
    Me.txtPreOro16 = Format(lrDatosVen!nPrecioK16, "#0.00")
    Me.txtPreOro18 = Format(lrDatosVen!nPrecioK18, "#0.00")
    Me.txtPreOro21 = Format(lrDatosVen!nPrecioK21, "#0.00")
    
    If lrDatosVen!nRGEstado = gColPRecGarEstNoIniciado Then
        txtEstado = "NO INICIADO"
        If Val(txtPreOro14) > 0 And Val(txtPreOro16) > 0 And Val(txtPreOro18) > 0 _
            And Val(txtPreOro21) > 0 Then
        End If
    ElseIf lrDatosVen!nRGEstado = gColPRecGarEstIniciado Then
        txtEstado = "INICIADO"
        cmdEditar.Enabled = False
    Else
        MsgBox " No existe el remate generado", vbCritical, " Error de Sistema "
        cmdEditar.Enabled = False
        txtNumVenta = ""
        txtFecVenta = Format("01/01/2000", "dd/mm/yyyy")
        txtHorVenta = Format("00:00", "hh:mm")
    End If
    Set lrDatosVen = Nothing
    
End Sub

'Valida el campo txtFecVenta
Private Sub txtFecVenta_GotFocus()
fEnfoque txtFecVenta
End Sub
Private Sub txtFecVenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtHorVenta.SetFocus
End If
End Sub
Private Sub txtFecVenta_LostFocus()
If Not ValFecha(txtFecVenta) Then
    txtFecVenta.SetFocus
ElseIf DateDiff("d", txtFecVenta, gdFecSis) > 0 Then
    MsgBox " Fecha no debe ser anterior a la fecha actual", vbInformation, " Aviso "
    txtFecVenta.SetFocus
End If
'VeriDatRem
End Sub

Private Sub txtFecVenta_Validate(Cancel As Boolean)
If Not ValFecha(txtFecVenta) Then
    Cancel = True
ElseIf DateDiff("d", txtFecVenta, gdFecSis) > 0 Then
    MsgBox " Fecha no debe ser anterior a la fecha actual", vbInformation, " Aviso "
    Cancel = True
End If

End Sub

'Valida el campo txtHorVenta
Private Sub txtHorVenta_GotFocus()
fEnfoque txtHorVenta
End Sub
Private Sub txtHorVenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtPreOro14.SetFocus
End If
End Sub
Private Sub txtHorVenta_LostFocus()
If Not ValidaHora(txtHorVenta) Then
    txtHorVenta.SetFocus
End If
'VeriDatRem
End Sub

'Valida el campo txtpreoro14
Private Sub txtPreOro14_GotFocus()
fEnfoque txtPreOro14
End Sub
Private Sub txtPreOro14_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro14, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro14 = Format(txtPreOro14, "#0.00")
    txtPreOro16.SetFocus
End If
End Sub
Private Sub txtPreOro14_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro16
Private Sub txtPreOro16_GotFocus()
fEnfoque txtPreOro16
End Sub
Private Sub txtPreOro16_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro16, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro16 = Format(txtPreOro16, "#0.00")
    txtPreOro18.SetFocus
End If
End Sub
Private Sub txtPreOro16_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro18
Private Sub txtPreOro18_GotFocus()
fEnfoque txtPreOro18
End Sub
Private Sub txtPreOro18_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro18, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro18 = Format(txtPreOro18, "#0.00")
    txtPreOro21.SetFocus
End If
End Sub
Private Sub txtPreOro18_LostFocus()
    VeriPreOro
End Sub

'Valida el campo txtpreoro21
Private Sub txtPreOro21_GotFocus()
fEnfoque txtPreOro21
End Sub
Private Sub txtPreOro21_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreOro21, KeyAscii)
If KeyAscii = 13 Then
    txtPreOro21 = Format(txtPreOro21, "#0.00")
    cmdGrabar.SetFocus
End If
End Sub
Private Sub txtPreOro21_LostFocus()
    VeriPreOro
End Sub

' Permite activar la opción de procesar solo cuando están ingresados los campos
' fecha y hora de Venta
Public Sub VeriDatRem()
    If Len(txtNumVenta) > 0 And Len(txtFecVenta) = 10 And Len(txtHorVenta) = 5 Then
    End If
End Sub

' Permite activar la opción de grabar solo cuando están ingresados los campos
' del precio del oro
Public Sub VeriPreOro()
    If Val(txtPreOro14) > 0 And Val(txtPreOro16) > 0 And Val(txtPreOro18) > 0 _
        And Val(txtPreOro21) > 0 Then
        cmdGrabar.Enabled = True
    End If
End Sub


Private Sub CargaParametros()

Dim loParam As COMDColocPig.DCOMColPCalculos
Dim loConstSis As COMDConstSistema.NCOMConstSistema
Dim lnProcesoCadaAgencia As Integer

Set loParam = New COMDColocPig.DCOMColPCalculos
    fnTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnTasaIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
Set loParam = Nothing


Set loConstSis = New COMDConstSistema.NCOMConstSistema
    fnJoyasDet = loConstSis.LeeConstSistema(109) ' Joyas en Detalle
    lnProcesoCadaAgencia = loConstSis.LeeConstSistema(121)  ' gConstSistPigRemateCadaAg
    If lnProcesoCadaAgencia = 1 Then  ' En cada agencia
        fsVentaCadaAgencia = gsCodCMAC & gsCodAge
    Else
        fsVentaCadaAgencia = gsCodCMAC & "00"
    End If
Set loConstSis = Nothing
End Sub

Private Function FormatoContratro(pContrato As String) As String
  FormatoContratro = Mid(pContrato, 1, 2) & "-" & Mid(pContrato, 3, 4) & "-" & Mid(pContrato, 7, 5) & "-" & Mid(pContrato, 12, 1)
End Function

