VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{2E37A9B8-9906-4590-9F3A-E67DB58122F5}#7.0#0"; "OcxLabelX.ocx"
Begin VB.Form frmColPContratosxCliente 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Consulta de Contratos por Cliente"
   ClientHeight    =   7215
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7695
   Icon            =   "frmColPContratosxCliente.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7215
   ScaleWidth      =   7695
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Ingrese Nro de Contrato"
      Height          =   6870
      Left            =   225
      TabIndex        =   23
      Top             =   210
      Width           =   7275
      Begin VB.CommandButton cmdAgencia 
         Caption         =   "A&gencias..."
         Height          =   345
         Left            =   4680
         TabIndex        =   40
         Top             =   2820
         Width           =   1020
      End
      Begin VB.CheckBox chkEstados 
         Caption         =   "Activar estados de busqueda"
         Height          =   225
         Left            =   405
         TabIndex        =   4
         Top             =   3540
         Width           =   2415
      End
      Begin VB.Frame Frame4 
         Enabled         =   0   'False
         Height          =   2025
         Left            =   210
         TabIndex        =   39
         Top             =   3525
         Width           =   6915
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Anulado no Desembolsado"
            Height          =   420
            Index           =   12
            Left            =   5400
            TabIndex        =   17
            Top             =   390
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Chafaloneado"
            Height          =   240
            Index           =   11
            Left            =   3650
            TabIndex        =   16
            Top             =   1380
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Anulado"
            Height          =   240
            Index           =   10
            Left            =   3650
            TabIndex        =   15
            Top             =   1050
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Subastado"
            Height          =   240
            Index           =   9
            Left            =   3650
            TabIndex        =   14
            Top             =   720
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Adjudicado"
            Height          =   240
            Index           =   8
            Left            =   3650
            TabIndex        =   13
            Top             =   390
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Renovado"
            Height          =   240
            Index           =   7
            Left            =   1900
            TabIndex        =   12
            Top             =   1380
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Para Remate"
            Height          =   240
            Index           =   6
            Left            =   1900
            TabIndex        =   11
            Top             =   1050
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Rematado"
            Height          =   240
            Index           =   5
            Left            =   1900
            TabIndex        =   10
            Top             =   720
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Vencido"
            Height          =   240
            Index           =   4
            Left            =   1900
            TabIndex        =   9
            Top             =   390
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Cancelado"
            Height          =   240
            Index           =   3
            Left            =   150
            TabIndex        =   8
            Top             =   1380
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Diferido"
            Height          =   240
            Index           =   2
            Left            =   150
            TabIndex        =   7
            Top             =   1050
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Desembolsado"
            Height          =   240
            Index           =   1
            Left            =   150
            TabIndex        =   6
            Top             =   720
            Width           =   1470
         End
         Begin VB.CheckBox chkBuscar 
            Caption         =   "Registrado"
            Height          =   240
            Index           =   0
            Left            =   150
            TabIndex        =   5
            Top             =   390
            Width           =   1470
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Activar fechas de búsqueda"
         Height          =   225
         Left            =   375
         TabIndex        =   1
         Top             =   2550
         Width           =   2415
      End
      Begin VB.Frame fraPeriodo1 
         Enabled         =   0   'False
         Height          =   795
         Left            =   180
         TabIndex        =   36
         Top             =   2550
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
         Begin VB.Label Label2 
            Caption         =   "Al :"
            Height          =   240
            Left            =   1935
            TabIndex        =   38
            Top             =   390
            Width           =   450
         End
         Begin VB.Label Label1 
            Caption         =   "Del :"
            Height          =   255
            Left            =   120
            TabIndex        =   37
            Top             =   435
            Width           =   435
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Impresión"
         Height          =   780
         Left            =   210
         TabIndex        =   35
         Top             =   5730
         Width           =   2535
         Begin VB.OptionButton optOpcionImpresion 
            Caption         =   "Pantalla"
            Height          =   225
            Index           =   0
            Left            =   225
            TabIndex        =   18
            Top             =   285
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optOpcionImpresion 
            Caption         =   "Impresora"
            Height          =   225
            Index           =   1
            Left            =   1350
            TabIndex        =   19
            Top             =   285
            Width           =   1020
         End
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "&Imprimir"
         Height          =   390
         Left            =   2850
         TabIndex        =   20
         Top             =   6030
         Width           =   1305
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   4290
         TabIndex        =   21
         Top             =   6030
         Width           =   1305
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         Height          =   390
         Left            =   5670
         TabIndex        =   22
         Top             =   6015
         Width           =   1305
      End
      Begin VB.Frame Frame3 
         Caption         =   "Datos del Cliente"
         Height          =   2190
         Left            =   195
         TabIndex        =   24
         Top             =   285
         Width           =   6885
         Begin VB.CommandButton cmdBuscar 
            Height          =   360
            Left            =   2700
            Picture         =   "frmColPContratosxCliente.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   0
            ToolTipText     =   "Buscar ..."
            Top             =   360
            Width           =   390
         End
         Begin OcxLabelX.LabelX lblCodigo 
            Height          =   450
            Left            =   855
            TabIndex        =   25
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
            Left            =   165
            TabIndex        =   26
            Top             =   1005
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
            Left            =   165
            TabIndex        =   27
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
            TabIndex        =   28
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
            TabIndex        =   29
            Top             =   240
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   794
            FondoBlanco     =   0   'False
            Resalte         =   0
            Bold            =   0   'False
            Alignment       =   0
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Teléfono"
            Height          =   195
            Left            =   5265
            TabIndex        =   34
            Top             =   1455
            Width           =   630
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Direccion"
            Height          =   195
            Left            =   225
            TabIndex        =   33
            Top             =   1425
            Width           =   675
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Nombre"
            Height          =   195
            Left            =   225
            TabIndex        =   32
            Top             =   780
            Width           =   555
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Código"
            Height          =   195
            Left            =   225
            TabIndex        =   31
            Top             =   360
            Width           =   495
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Documento"
            Height          =   195
            Left            =   3990
            TabIndex        =   30
            Top             =   270
            Width           =   825
         End
      End
   End
End
Attribute VB_Name = "frmColPContratosxCliente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Check1_Click()
fraPeriodo1.Enabled = IIf(Check1.value = 1, True, False)
mskPeriodo1Al.Text = "__/__/____"
mskPeriodo1Del.Text = "__/__/____"
End Sub

Private Sub Check1_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Check1.value = 1 Then
        mskPeriodo1Del.SetFocus
    Else
        chkEstados.SetFocus
    End If
End If
End Sub


Private Sub chkBuscar_KeyPress(Index As Integer, KeyAscii As Integer)
If KeyAscii = 13 Then
    If Index < 12 Then
        chkBuscar(Index + 1).SetFocus
    ElseIf Index = 12 Then
        cmdMostrar.SetFocus
    End If
End If
End Sub

Private Sub chkEstados_Click()
Dim i As Integer
Frame4.Enabled = IIf(chkEstados.value = 1, True, False)
For i = 0 To 12
    chkBuscar(i).value = 0
Next

End Sub

Private Sub chkEstados_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If chkEstados.value = 1 Then
        chkBuscar(0).SetFocus
    Else
        cmdMostrar.SetFocus
    End If
End If
End Sub

Private Sub cmdAgencia_Click()
    'Selec Age. a realizar Remate
    frmSelectAgencias.Inicio Me
    frmSelectAgencias.Show 1
End Sub

Private Sub CmdBuscar_Click()

Dim loPers As COMDPersona.UCOMPersona
Dim lsPersCod As String, lsPersNombre As String
Dim lsPersDireccDomicilio As String
Dim lsPersTelefono As String
Dim lsPersIdNroDNI As String
Dim lsPersIdNroRUC As String

Dim lsEstados As String

On Error GoTo ControlError

Limpiar

Set loPers = New COMDPersona.UCOMPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
    lsPersCod = loPers.sPersCod
    lsPersNombre = loPers.sPersNombre
    lsPersDireccDomicilio = loPers.sPersDireccDomicilio
    lsPersTelefono = loPers.sPersTelefono
    lsPersIdNroDNI = loPers.sPersIdnroDNI
    lsPersIdNroRUC = loPers.sPersIdnroRUC
Set loPers = Nothing

lblCodigo.Caption = lsPersCod
lblNombre.Caption = lsPersNombre
lblDireccion.Caption = lsPersDireccDomicilio
lblTelefono.Caption = lsPersTelefono
If Len(lsPersIdNroDNI) > 0 Then
    lblDocumento.Caption = "DNI " & lsPersIdNroDNI
Else
    If Len(lsPersIdNroRUC) > 0 Then
        lblDocumento.Caption = "RUC " & lsPersIdNroRUC
    Else
        lblDocumento.Caption = ""
    End If
End If
 
Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub Limpiar()
    lblCodigo.Caption = ""
    lblNombre.Caption = ""
    lblDireccion.Caption = ""
    lblTelefono.Caption = ""
    lblDocumento.Caption = ""
    Check1.value = 0
    chkEstados.value = 0
End Sub

Private Sub cmdMostrar_Click()

Dim lsEstados As String

If Len(Trim(lblCodigo.Caption)) = 0 Then
    MsgBox "Seleccione un cliente", vbExclamation, "Aviso"
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

If chkEstados.value = 1 Then
    If chkBuscar(0).value = 1 Then
        lsEstados = gColPEstRegis
    End If
    If chkBuscar(1).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstDesem
        Else
            lsEstados = lsEstados & ", " & gColPEstDesem
        End If
    End If
    If chkBuscar(2).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstDifer
        Else
            lsEstados = lsEstados & ", " & gColPEstDifer
        End If
    End If
    If chkBuscar(3).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstCance
        Else
            lsEstados = lsEstados & ", " & gColPEstCance
        End If
    End If
    If chkBuscar(4).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstVenci
        Else
            lsEstados = lsEstados & ", " & gColPEstVenci
        End If
    End If
    If chkBuscar(5).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstRemat
        Else
            lsEstados = lsEstados & ", " & gColPEstRemat
        End If
    End If
    If chkBuscar(6).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstPRema
        Else
            lsEstados = lsEstados & ", " & gColPEstPRema
        End If
    End If
    If chkBuscar(7).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstRenov
        Else
            lsEstados = lsEstados & ", " & gColPEstRenov
        End If
    End If
    If chkBuscar(8).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstAdjud
        Else
            lsEstados = lsEstados & ", " & gColPEstAdjud
        End If
    End If
    If chkBuscar(9).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstSubas
        Else
            lsEstados = lsEstados & ", " & gColPEstSubas
        End If
    End If
    If chkBuscar(10).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstAnula
        Else
            lsEstados = lsEstados & ", " & gColPEstAnula
        End If
    End If
    If chkBuscar(11).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstChafa
        Else
            lsEstados = lsEstados & ", " & gColPEstChafa
        End If
    End If
    If chkBuscar(12).value = 1 Then
        If Len(Trim(lsEstados)) = 0 Then
            lsEstados = gColPEstAnNoD
        Else
            lsEstados = lsEstados & ", " & gColPEstAnNoD
        End If
    End If
    If Len(Trim(lsEstados)) = 0 Then
        MsgBox "Ud. debe seleccionar al menos una opcion de estados", vbExclamation, "Aviso"
        chkBuscar(0).SetFocus
        Exit Sub
    End If
Else
    lsEstados = gColPEstAdjud & "," & gColPEstAnNoD & "," & gColPEstCance & "," & gColPEstChafa & _
                "," & gColPEstDesem & "," & gColPEstDifer & "," & gColPEstPRema & "," & gColPEstRegis & _
                "," & gColPEstRemat & "," & gColPEstRenov & "," & gColPEstSubas & "," & gColPEstVenci & ", " & gColPEstAnula
End If

EjecutaReporte lsEstados

End Sub


Private Sub EjecutaReporte(ByVal lsEstados As String)
Dim loRep As COMNColoCPig.NCOMColPRepo 'NColPRepo
Dim lsCadImp As String
Dim loPrevio As previo.clsPrevio
Dim lsDestino As String  ' P= Previo // I = Impresora // A = Archivo // E = Excel
Dim x As Integer
Dim lnAge As Integer
Dim lsmensaje As String

Dim psDescOperacion As String
psDescOperacion = "Lista de Contratos por Cliente"
    
    ' Reporte de  Contratos x Cliente
    lsCadImp = ""
    For lnAge = 1 To frmSelectAgencias.List1.ListCount
        If frmSelectAgencias.List1.Selected(lnAge - 1) = True Then
            Set loRep = New COMNColoCPig.NCOMColPRepo
            loRep.Inicio gsNomCmac, gsNomAge, gsCodUser, gdFecSis
            lsCadImp = lsCadImp & loRep.nRepo128035_ListadoContratosxCliente(Mid(frmSelectAgencias.List1.List(lnAge - 1), 1, 2), Trim(lblCodigo.Caption), Trim(lblNombre.Caption), Trim(lblDireccion.Caption), Trim(lblTelefono.Caption), Trim(lblDocumento.Caption), lsEstados, Me.mskPeriodo1Del.Text, Me.mskPeriodo1Al.Text, Check1.value, lsmensaje, gImpresora)
            If Trim(lsmensaje) <> "" Then
                 MsgBox lsmensaje, vbInformation, "Aviso"
                 Exit Sub
            End If
            Set loRep = Nothing
        End If
    Next lnAge
    
    If optOpcionImpresion(0).value = True Then
        lsDestino = "P"
    ElseIf optOpcionImpresion(1).value = True Then
        lsDestino = "A"
    End If
    
    If Len(Trim(lsCadImp)) > 0 Then
        Set loPrevio = New previo.clsPrevio
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


Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub mskPeriodo1Al_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    chkEstados.SetFocus
End If
End Sub

Private Sub mskPeriodo1Del_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    mskPeriodo1Al.SetFocus
End If
End Sub

