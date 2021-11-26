VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmColPSubastaRegVenta 
   Caption         =   "Crédito Pignoraticio : Registrar Venta en Subasta"
   ClientHeight    =   6705
   ClientLeft      =   930
   ClientTop       =   2010
   ClientWidth     =   8070
   Icon            =   "frmColPSubastaRegVenta.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   8070
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraContenedor 
      Height          =   6225
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   30
      Width           =   7815
      Begin VB.Frame fraContenedor 
         Caption         =   "Adjudicatario"
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
         Height          =   1035
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   4095
         Width           =   7455
         Begin VB.TextBox txtNomAdj 
            Enabled         =   0   'False
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
            Left            =   825
            TabIndex        =   17
            Tag             =   "txtnombre"
            Top             =   570
            Width           =   3780
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "&Buscar..."
            Enabled         =   0   'False
            Height          =   300
            Left            =   2520
            TabIndex        =   0
            Top             =   225
            Width           =   930
         End
         Begin VB.TextBox txtCodAdj 
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   825
            TabIndex        =   16
            Tag             =   "txtcodigo"
            Top             =   225
            Width           =   1455
         End
         Begin VB.TextBox txtTriAdj 
            Enabled         =   0   'False
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
            Left            =   6030
            TabIndex        =   15
            Tag             =   "txtTributario"
            Top             =   570
            Width           =   1080
         End
         Begin VB.TextBox txtNatAdj 
            Enabled         =   0   'False
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
            Left            =   6045
            TabIndex        =   14
            Tag             =   "txtDocumento"
            Top             =   240
            Width           =   1080
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Nombre :"
            Height          =   225
            Index           =   7
            Left            =   135
            TabIndex        =   21
            Top             =   615
            Width           =   735
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Código :"
            Height          =   225
            Index           =   8
            Left            =   120
            TabIndex        =   20
            Top             =   285
            Width           =   615
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Doc.Natural : "
            Height          =   255
            Index           =   2
            Left            =   4875
            TabIndex        =   19
            Top             =   270
            Width           =   1110
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Doc.Tributario : "
            Height          =   255
            Index           =   3
            Left            =   4890
            TabIndex        =   18
            Top             =   615
            Width           =   1110
         End
      End
      Begin VB.Frame fraContenedor 
         Height          =   600
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   5160
         Width           =   7455
         Begin VB.TextBox txtPreVentaBruta 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFFF&
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   6165
            MaxLength       =   9
            TabIndex        =   1
            Text            =   "0"
            Top             =   180
            Width           =   1035
         End
         Begin VB.TextBox txtPreBaseVenta 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3465
            TabIndex        =   9
            Text            =   "0"
            Top             =   150
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.TextBox txtDeuda 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   825
            TabIndex        =   8
            Text            =   "0"
            Top             =   150
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Deuda :"
            Height          =   225
            Index           =   13
            Left            =   75
            TabIndex        =   12
            Top             =   195
            Visible         =   0   'False
            Width           =   675
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Precio Venta :"
            Height          =   225
            Index           =   14
            Left            =   5070
            TabIndex        =   11
            Top             =   225
            Width           =   1050
         End
         Begin VB.Label lblEtiqueta 
            Caption         =   "Precio Base :"
            Height          =   225
            Index           =   6
            Left            =   2325
            TabIndex        =   10
            Top             =   195
            Visible         =   0   'False
            Width           =   1020
         End
      End
      Begin MSMask.MaskEdBox txtNroDocumento 
         Height          =   330
         Left            =   5445
         TabIndex        =   2
         Top             =   5790
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   582
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   9
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "###-#####"
         PromptChar      =   "_"
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   120
         TabIndex        =   24
         Top             =   240
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.ActXColPDesCon AXDesCon 
         Height          =   3495
         Left            =   120
         TabIndex        =   25
         Top             =   585
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   6165
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Nro.Documento :"
         Height          =   225
         Index           =   10
         Left            =   4095
         TabIndex        =   22
         Top             =   5835
         Width           =   1305
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6750
      TabIndex        =   5
      Top             =   6300
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3870
      TabIndex        =   3
      Top             =   6300
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5355
      TabIndex        =   4
      Top             =   6300
      Width           =   975
   End
   Begin VB.Label lblMensaje 
      Caption         =   "<F12> Número Contrato Antiguo"
      Height          =   285
      Left            =   120
      TabIndex        =   23
      Top             =   6240
      Width           =   2655
   End
End
Attribute VB_Name = "frmColPSubastaRegVenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REGISTRO DE VENTA DE SUBASTA.
'Archivo:  frmColPSubastaRegVenta.frm
'LAYG   :  18/07/2001.
'Resumen:  Nos permite registrar una venta de un contrato que se está Subastando
Option Explicit

Dim fnVarNroSubasta As String
Dim fsVarNroSubasta As String
Dim pDifeDiasRema As Integer
Dim fnVarTasaImpuesto As Double
Dim fnVarTasaPreparacionRemate As Double
Dim fnVarTasaIGV As Double
Dim pAgeRemSub As String * 2

Dim fnVarPrecioAdjudica As Currency

'Inicializa el formulario
Public Sub Inicio(ByVal psNroProceso As String)
    fsVarNroSubasta = psNroProceso
    CargaParametros
    Limpiar
    Me.Show 1
End Sub

'Inicializa las variables
Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    AXDesCon.Limpiar
    txtCodAdj = ""
    txtNomAdj = ""
    txtNatAdj = ""
    txtTriAdj = ""
    txtDeuda.Text = Format(0, "#0.00")
    txtPreBaseVenta.Text = Format(0, "#0.00")
    txtPreVentaBruta.Text = Format(0, "#0.00")
End Sub

'Busca el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)

Dim lbOk As Boolean, lnTotalVta As Double
Dim lrValida As ADODB.Recordset
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lsMensaje As String
'On Error GoTo ControlError

    'Valida Contrato
    Set lrValida = New ADODB.Recordset
    Set loValContrato = New COMNColoCPig.NCOMColPValida
    Set lrValida = loValContrato.nValidaRegVentaSubastaCredPignoraticio(psNroContrato, "A", fsVarNroSubasta, lsMensaje)
    If Trim(lsMensaje) <> "" Then
        MsgBox lsMensaje, vbInformation, "Aviso"
        Exit Sub
    End If
            
    If lrValida Is Nothing Then ' Hubo un Error
        Limpiar
        Set lrValida = Nothing
        Exit Sub
    End If
    
    If lrValida!bExcepVta = 0 Then
        If DateDiff("d", lrValida!dPrdEstado, gdFecSis) <= 60 Then
            MsgBox "Contrato con menos de 60 dias de adjudicado para ser vendido.", vbOKOnly + vbInformation, "AVISO"
            Limpiar
            Set lrValida = Nothing
            Exit Sub
        End If
    Else
            MsgBox "Contrato marcado de no venta por excepcion.", vbOKOnly + vbInformation, "AVISO"
            Limpiar
            Set lrValida = Nothing
            Exit Sub
    End If
    
    lnTotalVta = loValContrato.CalcPrecioVta(gdFecSis, lrValida!cCtaCod)
       
    
    Set loValContrato = Nothing
    
    
    'Muestra Datos
    lbOk = fgMuestraCredPig_AXDesCon(psNroContrato, Me.AXDesCon, False)

   ' txtDeuda = Format(lrValida!nDeuda, "#0.00")
    txtPreVentaBruta.Text = Format(lnTotalVta, "#,##0.00")
    'txtPreBaseVenta = Format(lrValida!nRemSubBaseVta, "#0.00")
    'txtPreVentaBruta = Format(lrValida!nRemSubBaseVta, "#0.00")
    
    fnVarPrecioAdjudica = Format(lrValida!nValRegistroAdj, "#0.00")
    
    
    Set lrValida = Nothing
        
    cmdGrabar.Enabled = True
        
    AXCodCta.Enabled = False
    cmdBuscar.Enabled = True
    cmdBuscar.SetFocus

Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox "Error: " & Err.Number & " " & Err.Description & vbCr & _
        "Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Call BuscaContrato(AXCodCta.NroCuenta)
End Sub

'Busca el Adjudicatario por nombre y/o documento
Private Sub cmdBuscar_Click()
Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String
Dim liFil As Integer
Dim ls As String
On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
Set loPers = frmBuscaPersona.Inicio

If Not loPers Is Nothing Then
    lsPersCod = loPers.sPersCod
    Me.txtCodAdj = loPers.sPersCod
    txtNomAdj = PstaNombre(loPers.sPersNombre, False)
    
    txtPreVentaBruta.Enabled = True
    txtPreVentaBruta.SetFocus

End If

Set loPers = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Cancela el proceso actual e inicializa uno nuevo
Private Sub cmdCancelar_Click()
    Limpiar
    cmdGrabar.Enabled = False
    cmdBuscar.Enabled = False
    txtPreVentaBruta.Enabled = False
    txtNroDocumento.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocus
End Sub

'Graba los cambios en la base de datos
Private Sub cmdGrabar_Click()

'On Error GoTo ControlError
Dim loContFunct As COMNContabilidad.NCOMContFunciones
Dim loGrabarVta As COMNColoCPig.NCOMColPContrato

Dim lsMovNro As String
Dim lsFechaHoraGrab As String

Dim lnVtaNeta As Currency
Dim lnImpuestoVtaSub As Currency
Dim lnCostVtaAdj As Currency

If Len(txtCodAdj) <> 13 Then
    MsgBox " Falta ingresar el Adjudicatario ", vbInformation, " Aviso "
    cmdBuscar.Enabled = True
    cmdBuscar.SetFocus
    Exit Sub
ElseIf CCur(txtPreVentaBruta) <= 0 Then
    MsgBox " Falta ingresar el precio de venta bruta ", vbInformation, " Aviso "
    txtPreVentaBruta.SetFocus
    Exit Sub
ElseIf Len(Trim(txtNroDocumento)) <> 8 Then
    MsgBox " Falta ingresar el número de documento ", vbInformation, " Aviso "
    'txtNroDocumento.SetFocus
    Exit Sub
End If

'asigna valores a variables
lnVtaNeta = Round(Val(txtPreVentaBruta.Text) / (1 + fnVarTasaIGV), 2)
lnImpuestoVtaSub = Val(txtPreVentaBruta.Text) - lnVtaNeta

If MsgBox(" Grabar Venta de Joyas en Subasta ? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
        cmdGrabar.Enabled = False
        cmdBuscar.Enabled = False
        txtPreVentaBruta.Enabled = False
        txtNroDocumento.Enabled = False
        
        'Genera el Mov Nro
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
            lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        
        lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
        Set loGrabarVta = New COMNColoCPig.NCOMColPContrato
            'Grabar Venta de Remate
            Call loGrabarVta.nSubastaVentaCredPignoraticio(AXCodCta.NroCuenta, fsVarNroSubasta, lsFechaHoraGrab, _
                 lsMovNro, CCur(Val(Me.txtPreVentaBruta.Text)), lnVtaNeta, lnImpuestoVtaSub, fnVarPrecioAdjudica, _
                 Val(Me.AXDesCon.Oro14), Val(Me.AXDesCon.Oro16), Val(Me.AXDesCon.Oro18), Val(Me.AXDesCon.Oro21), False)
        Set loGrabarVta = Nothing

        'Impresión
        If MsgBox(" Desea realizar impresión de Recibo ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
            Dim lscadimp As String
            Dim oImpColP As COMNColoCPig.NCOMColPImpre
            Set oImpColP = New COMNColoCPig.NCOMColPImpre
                lscadimp = oImpColP.ImpRecVtaSub(AXCodCta.NroCuenta, txtNomAdj.Text, CCur(Val(txtPreVentaBruta.Text)), gsNomAge, gdFecSis, gsCodUser, gImpresora)
            Set oImpColP = Nothing
            
            Dim loPrevio As previo.clsprevio
            Set loPrevio = New previo.clsprevio
                loPrevio.PrintSpool sLpt, lscadimp, False
            Set loPrevio = Nothing
            
            Do While True
            
                If MsgBox("Desea reimprimir ?", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
                    
                    Set loPrevio = New previo.clsprevio
                        loPrevio.PrintSpool sLpt, lscadimp, False
                    Set loPrevio = Nothing

                Else
                    Exit Do
                End If
            Loop
        End If
        txtNroDocumento = txtNroDocumento + 1
        txtNroDocumento = String(8 - Len(txtNroDocumento), "0") & Trim(str(txtNroDocumento))
        Limpiar
        AXCodCta.Enabled = True
        AXCodCta.SetFocus
        
Else
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
End If
Exit Sub


ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

'Finaliza el formulario
Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And AXCodCta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gColConsuPrendario, False)
        If sCuenta <> "" Then
            AXCodCta.NroCuenta = sCuenta
            AXCodCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub txtNomAdj_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloLetras(KeyAscii)
End Sub

'Valida el campo txtnrodocumento
Private Sub txtNroDocumento_KeyPress(KeyAscii As Integer)
Dim loValid As COMNColoCPig.NCOMColPValida
Dim lbExiste As Boolean
If KeyAscii = 13 And Len(Trim(txtNroDocumento)) = 8 Then
    Set loValid = New COMNColoCPig.NCOMColPValida
        lbExiste = loValid.nDocumentoEmitido(3, txtNroDocumento.Text, "'" & geColPVtaSubasta & "'")
    Set loValid = Nothing
    If lbExiste = True Then
        MsgBox "Número de Boleta duplicada" & vbCr & "Ingrese un número diferente", vbInformation, " Aviso "
    Else
        cmdGrabar.Enabled = True
        cmdGrabar.SetFocus
    End If
End If
End Sub

'Valida el campo txtpreventabruta
Private Sub txtPreVentaBruta_GotFocus()
    fEnfoque txtPreVentaBruta
End Sub
Private Sub txtPreVentaBruta_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtPreVentaBruta, KeyAscii)
If KeyAscii = 13 Then
    If Val(txtPreVentaBruta.Text) <= 0 Then
        MsgBox " Precio Venta debe ser mayor a 0 ", vbInformation, " Aviso "
        txtPreVentaBruta.SetFocus
    Else
        txtNroDocumento.Enabled = True
        txtNroDocumento.SetFocus
    End If
End If
End Sub
Private Sub txtPreVentaBruta_LostFocus()
txtPreVentaBruta = Format(Val(txtPreVentaBruta), "#0.00")
VeriVenBru
End Sub

'Procedimiento de verificación de la venta bruta
Private Sub VeriVenBru()
    If Val(txtPreVentaBruta.Text) < Val(txtPreBaseVenta.Text) Then
        MsgBox " Precio Venta debe ser mayor a Precio Base ", vbInformation, " Aviso "
        txtPreVentaBruta.SetFocus
    End If
End Sub

Private Sub CargaParametros()
Dim loParam As COMDColocPig.DCOMColPCalculos
Set loParam = New COMDColocPig.DCOMColPCalculos
    fnVarTasaPreparacionRemate = loParam.dObtieneColocParametro(gConsColPTasaPreparaRemate)
    fnVarTasaImpuesto = loParam.dObtieneColocParametro(gConsColPTasaImpuesto)
    fnVarTasaIGV = loParam.dObtieneColocParametro(gConsColPTasaIGV)
    'pAgeRemSub = Right(ReadVarSis("CPR", "cAgeRemSub"), 2)
    'pDifeDiasRema = Val(ReadVarSis("CPR", "nDifeDiasRema"))
Set loParam = Nothing
End Sub

