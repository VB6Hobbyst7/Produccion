VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{2E37A9B8-9906-4590-9F3A-E67DB58122F5}#7.0#0"; "OcxLabelX.ocx"
Begin VB.Form frmColPMovs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Picnoraticios: Consulta de Kardex por Contrato"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7470
   Icon            =   "frmColPMovs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   7470
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese Nro de Contrato"
      Height          =   5040
      Left            =   180
      TabIndex        =   9
      Top             =   180
      Width           =   7125
      Begin VB.CommandButton cmdBuscar 
         Height          =   345
         Left            =   3900
         Picture         =   "frmColPMovs.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Buscar ..."
         Top             =   390
         Width           =   420
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos del Cliente"
         Height          =   2190
         Left            =   105
         TabIndex        =   14
         Top             =   990
         Width           =   6885
         Begin OcxLabelX.LabelX lblCodigo 
            Height          =   450
            Left            =   855
            TabIndex        =   15
            Top             =   330
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   794
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lblNombre 
            Height          =   420
            Left            =   180
            TabIndex        =   16
            Top             =   1020
            Width           =   6660
            _ExtentX        =   11748
            _ExtentY        =   741
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lblDireccion 
            Height          =   450
            Left            =   180
            TabIndex        =   17
            Top             =   1680
            Width           =   5100
            _ExtentX        =   8996
            _ExtentY        =   794
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lblTelefono 
            Height          =   450
            Left            =   5295
            TabIndex        =   18
            Top             =   1635
            Width           =   1530
            _ExtentX        =   2699
            _ExtentY        =   794
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin OcxLabelX.LabelX lblDocumento 
            Height          =   450
            Left            =   4830
            TabIndex        =   23
            Top             =   240
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   794
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Left            =   3990
            TabIndex        =   24
            Top             =   270
            Width           =   825
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   225
            TabIndex        =   22
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Left            =   225
            TabIndex        =   21
            Top             =   780
            Width           =   555
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Direccion"
            Height          =   195
            Left            =   225
            TabIndex        =   20
            Top             =   1425
            Width           =   675
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono"
            Height          =   195
            Left            =   5265
            TabIndex        =   19
            Top             =   1455
            Width           =   630
         End
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   390
         Left            =   4785
         TabIndex        =   8
         Top             =   4440
         Width           =   1305
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   3090
         TabIndex        =   7
         Top             =   4440
         Width           =   1305
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Imprimir"
         Height          =   390
         Left            =   1440
         TabIndex        =   6
         Top             =   4440
         Width           =   1305
      End
      Begin VB.Frame Frame2 
         Caption         =   "Impresión"
         Height          =   780
         Left            =   4200
         TabIndex        =   13
         Top             =   3330
         Width           =   2535
         Begin VB.OptionButton optOpcionImpresion 
            Caption         =   "Impresora"
            Height          =   225
            Index           =   1
            Left            =   1350
            TabIndex        =   5
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optOpcionImpresion 
            Caption         =   "Pantalla"
            Height          =   225
            Index           =   0
            Left            =   225
            TabIndex        =   4
            Top             =   285
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Activar fechas de búsqueda"
         Height          =   225
         Left            =   240
         TabIndex        =   1
         Top             =   3315
         Width           =   2415
      End
      Begin VB.Frame fraPeriodo1 
         Enabled         =   0   'False
         Height          =   795
         Left            =   120
         TabIndex        =   10
         Top             =   3300
         Width           =   3705
         Begin MSMask.MaskEdBox mskPeriodo1Al 
            Height          =   330
            Left            =   2265
            TabIndex        =   3
            Top             =   345
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   582
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
         Begin MSMask.MaskEdBox mskPeriodo1Del 
            Height          =   315
            Left            =   540
            TabIndex        =   2
            Top             =   345
            Width           =   1260
            _ExtentX        =   2223
            _ExtentY        =   556
            _Version        =   393216
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
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
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   435
            Width           =   435
         End
         Begin VB.Label Label2 
            Caption         =   "Al :"
            Height          =   240
            Left            =   1935
            TabIndex        =   11
            Top             =   390
            Width           =   450
         End
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   555
         Left            =   135
         TabIndex        =   0
         Top             =   390
         Width           =   4440
         _ExtentX        =   7832
         _ExtentY        =   979
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
End
Attribute VB_Name = "frmColPMovs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then BuscaContrato (AXCodCta.NroCuenta)
End Sub

Private Sub Check1_Click()
fraPeriodo1.Enabled = IIf(Check1.value = 1, True, False)
mskPeriodo1Al.Text = "__/__/____"
mskPeriodo1Del.Text = "__/__/____"
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check1.value = 1 Then
        Check1.SetFocus
    Else
        cmdMostrar.SetFocus
    End If
End If
End Sub

Private Sub cmdBuscar_Click()

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsEstados As String
Dim loPersContrato As COMDColocPig.DCOMColPContrato
Dim lrContratos As ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim lrCredPigPersonas As ADODB.Recordset

On Error GoTo ControlError

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
Set loPers = Nothing

' Selecciona Estados
lsEstados = gColPEstAdjud & "," & gColPEstAnNoD & "," & gColPEstCance & "," & gColPEstChafa & _
"," & gColPEstDesem & "," & gColPEstDifer & "," & gColPEstPRema & "," & gColPEstRegis & _
"," & gColPEstRemat & "," & gColPEstRenov & "," & gColPEstSubas & "," & gColPEstVenci

If Trim(lsPersCod) <> "" Then
    Set loPersContrato = New COMDColocPig.DCOMColPContrato
        Set lrContratos = loPersContrato.dObtieneCredPigDePersona(lsPersCod, lsEstados, Mid(gsCodAge, 4, 2))
    Set loPersContrato = Nothing
End If

Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(lsPersNombre, lrContratos)
    If loCuentas.sCtaCod <> "" Then
        AXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        AXCodCta.SetFocusCuenta
    Else
        Limpiar
    End If
Set loCuentas = Nothing

BuscaContrato AXCodCta.NroCuenta
    
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub


'Cancela el proceso actual e inicializa uno nuevo
Private Sub cmdCancelar_Click()
    Limpiar
    'cmdMostrar.Enabled = False
    AXCodCta.Enabled = True
    AXCodCta.SetFocusCuenta
End Sub

Private Sub cmdMostrar_Click()
If Len(Trim(lblCodigo.Caption)) = 0 Then
    MsgBox "Seleccione un credito", vbExclamation, "Aviso"
    cmdBuscar.SetFocus
    Exit Sub
End If
If Check1.value = 1 Then
    If IsDate(mskPeriodo1Del.Text) = False Then
        MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
        mskPeriodo1Del.SetFocus
        Exit Sub
    End If
    If IsDate(mskPeriodo1Al.Text) = False Then
        MsgBox "Ingrese una fecha correcta", vbExclamation, "Aviso"
        mskPeriodo1Al.SetFocus
        Exit Sub
    End If
End If

EjecutaReporte

End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    Limpiar
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

Private Sub Limpiar()
    Me.AXCodCta.NroCuenta = fgIniciaAxCuentaPignoraticio
    lblCodigo.Caption = ""
    lblNombre.Caption = ""
    lblDireccion.Caption = ""
    lblTelefono.Caption = ""
    lblDocumento.Caption = ""
    Check1.value = 0
End Sub
'Permite buscar el contrato ingresado
Private Sub BuscaContrato(ByVal psNroContrato As String)
Dim loValContrato As COMNColoCPig.NCOMColPValida
Dim lrValida As ADODB.Recordset
Dim lbOk As Boolean
Dim loMuestraContrato As COMDColocPig.DCOMColPContrato
Dim lrCredPigPersonas As ADODB.Recordset

On Error GoTo ControlError
    
'    'Muestra Datos
    
    Set loMuestraContrato = New COMDColocPig.DCOMColPContrato

    Set lrCredPigPersonas = loMuestraContrato.dObtieneDatosCreditoPignoraticioPersonas(psNroContrato)
    
    With lrCredPigPersonas
        If .BOF Then
            Limpiar
        Else
            lblCodigo.Caption = !cPersCod
            lblNombre.Caption = !cPersNombre & " " & !cpersapellido
            lblDireccion.Caption = !cPersDireccDomicilio
            lblTelefono.Caption = !cPersTelefono
            If Len(!NroDNI) > 0 Then
                lblDocumento.Caption = "DNI " & !NroDNI
            ElseIf Len(!NroRuc) > 0 Then
                lblDocumento.Caption = "RUC " & !NroRuc
            Else
                lblDocumento.Caption = ""
            End If
            AXCodCta.Enabled = False
        End If
        Check1.SetFocus
    End With
    
    Set lrCredPigPersonas = Nothing
    Set lrValida = Nothing
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub EjecutaReporte()
Dim loRep As COMNColoCPig.NCOMColPRepo
Dim lsCadImp As String
Dim loPrevio As previo.clsprevio
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel
Dim X As Integer

Dim lsmensaje As String

Dim psDescOperacion As String
psDescOperacion = "Movimientos de Contrato"
Set loRep = New COMNColoCPig.NCOMColPRepo
    loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
    ' Reporte de Movimientos de Contrato
     lsCadImp = loRep.nRepo128080_ListadoMovsxContrato(Trim(lblCodigo.Caption), Trim(lblNombre.Caption), Trim(lblDireccion.Caption), Trim(lblTelefono.Caption), Trim(lblDocumento.Caption), AXCodCta.NroCuenta, Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, Check1.value, lsmensaje, gImpresora)
        
    If Trim(lsmensaje) <> "" Then
         MsgBox lsmensaje, vbInformation, "Aviso"
         Exit Sub
    End If
    
    If optOpcionImpresion(0).value = True Then
        lsDestino = "P"
    ElseIf optOpcionImpresion(1).value = True Then
        lsDestino = "A"
    End If
    
Set loRep = Nothing
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New previo.clsprevio
            If lsDestino = "P" Then
                loPrevio.Show lsCadImp, psDescOperacion, True
            ElseIf lsDestino = "A" Then
                frmImpresora.Show 1
                loPrevio.PrintSpool sLpt, lsCadImp, True
            End If
        Set loPrevio = Nothing
    Else
        MsgBox "No Existen Datos para el reporte", vbInformation, "Aviso"
    End If
End Sub
 

Private Sub mskPeriodo1Al_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdMostrar.SetFocus
End If
End Sub

Private Sub mskPeriodo1Del_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    mskPeriodo1Al.SetFocus
End If
End Sub
