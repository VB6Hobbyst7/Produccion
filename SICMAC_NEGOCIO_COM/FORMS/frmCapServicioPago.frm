VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapServicioPago 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   4320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11160
   Icon            =   "frmCapServicioPago.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4320
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab stbGeneral 
      Height          =   4095
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   7223
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Gestión de Convenios de Pago"
      TabPicture(0)   =   "frmCapServicioPago.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "FRAConvenio"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "FRAComision"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "FRAPenalidad"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FRACuenta"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "FRAGeneral"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdBaja"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdSalir"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdGuardar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdEditar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "FRAComisionTransf"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.Frame FRAComisionTransf 
         Caption         =   "Com. Transf"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   9270
         TabIndex        =   25
         Top             =   480
         Width           =   1275
         Begin VB.TextBox txtComisionTransf 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   170
            TabIndex        =   26
            Text            =   "0.00"
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   3600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   375
         Left            =   9480
         TabIndex        =   7
         Top             =   3600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   8040
         TabIndex        =   8
         Top             =   3600
         Width           =   1335
      End
      Begin VB.CommandButton cmdBaja 
         Caption         =   "&Dar de baja"
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   3600
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Frame FRAGeneral 
         Caption         =   "Datos Generales de Convenio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1095
         Left            =   120
         TabIndex        =   16
         Top             =   2400
         Width           =   10695
         Begin VB.TextBox txtDescripcion 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1995
            MaxLength       =   100
            TabIndex        =   6
            Top             =   650
            Width           =   8535
         End
         Begin VB.TextBox txtNombre 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1995
            MaxLength       =   50
            TabIndex        =   5
            Top             =   240
            Width           =   8535
         End
         Begin VB.Label Label2 
            Caption         =   "Descripción:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   900
            TabIndex        =   21
            Top             =   690
            Width           =   1095
         End
         Begin VB.Label lblNombre 
            Caption         =   "Nombre del convenio:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   285
            Width           =   1935
         End
      End
      Begin VB.Frame FRACuenta 
         Caption         =   "Cuenta de Pago Convenio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1215
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   7695
         Begin VB.TextBox txtMoneda 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   1150
            Locked          =   -1  'True
            TabIndex        =   18
            Top             =   720
            Width           =   1215
         End
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "..."
            Height          =   360
            Left            =   3720
            TabIndex        =   1
            ToolTipText     =   "Busca cliente por nombre, documento o codigo"
            Top             =   240
            Width           =   375
         End
         Begin SICMACT.ActXCodCta txtCuenta 
            Height          =   375
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   3630
            _ExtentX        =   6403
            _ExtentY        =   661
            Texto           =   "Cuenta N°:"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
         Begin VB.Label lblMoneda 
            Caption         =   "Moneda:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   360
            TabIndex        =   19
            Top             =   765
            Width           =   735
         End
      End
      Begin VB.Frame FRAPenalidad 
         Caption         =   "Penalidad"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   1215
         Left            =   7920
         TabIndex        =   14
         Top             =   1200
         Width           =   2895
         Begin VB.TextBox txtLimite 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   1
            EndProperty
            Height          =   330
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   4
            Top             =   600
            Width           =   1215
         End
         Begin VB.TextBox txtPenalidad 
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """S/."" #,##0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
            Height          =   330
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   3
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Lim. Max. Ope.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   720
            Width           =   1455
         End
         Begin VB.Label lblMontoPenalidad 
            Caption         =   "Monto S/.:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame FRAComision 
         Caption         =   "Comision"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   7920
         TabIndex        =   13
         Top             =   480
         Width           =   1185
         Begin VB.TextBox txtComision 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BeginProperty DataFormat 
               Type            =   1
               Format          =   """S/."" #,##0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   2
            EndProperty
            Height          =   330
            Left            =   120
            MaxLength       =   10
            TabIndex        =   2
            Text            =   "0.00"
            Top             =   240
            Width           =   945
         End
      End
      Begin VB.Frame FRAConvenio 
         Caption         =   "Intitución Convenio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   735
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   7695
         Begin VB.TextBox txtConvenio 
            Appearance      =   0  'Flat
            Height          =   330
            Left            =   2040
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   240
            Width           =   5535
         End
         Begin SICMACT.TxtBuscar txtCodPers 
            Height          =   350
            Left            =   120
            TabIndex        =   0
            Top             =   240
            Width           =   1860
            _extentx        =   3281
            _extenty        =   609
            appearance      =   1
            appearance      =   1
            font            =   "frmCapServicioPago.frx":0326
            appearance      =   1
            tipobusqueda    =   3
            stitulo         =   ""
         End
      End
   End
End
Attribute VB_Name = "frmCapServicioPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmCapServicioPago
'*** Descripción : Formulario para registrar un convenio.
'*** Creación : ELRO el 20130701 04:14:23 PM, según RFC1306270002
'********************************************************************
Option Explicit

Dim fnAccion As Integer '1:Nuevo, 2:Editar
Dim fnIdSerPag As Long

Public Function inicarMantenimiento()
cmdEditar.Visible = True
cmdBaja.Visible = True
Me.Caption = "Mantenimiento de Convenio de Pago"
FRACuenta.Enabled = False
FRAGeneral.Enabled = False
FRAComision.Enabled = False
FRAComisionTransf.Enabled = False 'RIRO20150430 ERS146-2014
FRAPenalidad.Enabled = False
fnAccion = 2
Show 1
End Function

Public Function iniciarRegistro()
cmdGuardar.Visible = True
Me.Caption = "Registro de Convenio de Pago"
fnAccion = 1
Show 1
End Function

Private Sub iniciar()
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = False
txtCuenta.EnabledCta = False
End Sub

Private Sub LimpiarCampos()
txtCodPers = ""
txtConvenio = ""
txtCuenta.CMAC = ""
txtCuenta.Age = ""
txtCuenta.Prod = ""
txtCuenta.Cuenta = ""
txtMoneda = ""
txtNombre = ""
txtDescripcion = ""
txtComision = "0.00" 'RIRO20150430 ERS146-2014
txtComisionTransf.Text = "0.00" 'RIRO20150430 ERS146-2014
txtPenalidad = ""
txtLimite = ""
End Sub

Private Sub limpiarCamposEditar()
txtCodPers.Text = ""
txtConvenio = ""
txtCuenta.texto = ""
txtMoneda = ""
txtNombre = ""
txtDescripcion = ""
txtComision.Text = "0.00" 'RIRO20150430 ERS146-2014
txtComisionTransf.Text = "0.00" 'RIRO20150430 ERS146-2014
txtPenalidad = ""
txtLimite = ""
End Sub

Private Function validarDatos() As Boolean
validarDatos = False
If Trim(txtComision) = "" Then
    MsgBox "Debe ingresar la comisión.", vbInformation, "Aviso"
    Exit Function
End If

If Trim(txtPenalidad) = "" Then
    MsgBox "Debe ingresar la penalidad.", vbInformation, "Aviso"
    Exit Function
End If

If Trim(txtLimite) = "" Then
    MsgBox "Debe ingresar el límite máximo de operaciones diarias.", vbInformation, "Aviso"
    Exit Function
End If

If Trim(txtNombre) = "" Then
    MsgBox "Debe ingresar el nombre del convenio.", vbInformation, "Aviso"
    Exit Function
End If

If Trim(txtDescripcion) = "" Then
    MsgBox "Debe ingresar la descripción del convenio.", vbInformation, "Aviso"
    Exit Function
End If

If txtCuenta.CMAC = "" Or txtCuenta.Age = "" Or txtCuenta.Prod = "" Or txtCuenta.Cuenta = "" Then
    MsgBox "Debe ingresar la cuenta del convenio.", vbInformation, "Aviso"
    Exit Function
End If

If CDec(txtComision) < 0# Then
    MsgBox "La comisión debe ser mayor a cero.", vbInformation, "Aviso"
    Exit Function
End If

If CDec(txtPenalidad) < 0# Then
    MsgBox "La penalidad debe ser mayor a cero.", vbInformation, "Aviso"
    Exit Function
End If

'If CInt(txtLimite) < 0# Then 'RIRO20150814 Comentado
If CLng(txtLimite) < 0# Then 'RIRO20150814
    MsgBox "El límite máximo de las operaciones de diarias debe ser mayor a cero.", vbInformation, "Aviso"
    Exit Function
End If

validarDatos = True
End Function


Private Sub cmdBaja_Click()
If fnIdSerPag = 0 Then
    MsgBox "Debe seleccionar un convenio.", vbInformation, "Aviso"
    Exit Sub
End If

If MsgBox("¿Esta seguro de dar de baja el convenio?", vbYesNo, "Aviso") = vbYes Then
    Dim oNCOMCaptaGenerales As COMNCaptaGenerales.NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim oNCOMContFunciones As NCOMContFunciones
    Set oNCOMContFunciones = New NCOMContFunciones
    Dim lsMovNro As String
    Dim lnConfirmacion As Long
    
    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    lnConfirmacion = oNCOMCaptaGenerales.darBajaServicioPago(fnIdSerPag, lsMovNro)
    If lnConfirmacion > 0 Then
        MsgBox "Se dio de baja el convenio satisfactoriamente.", vbInformation, "Aviso"
        LimpiarCampos
        fnIdSerPag = 0
    Else
        MsgBox "No se dio de baja del convenio satisfactoriamente.", vbCritical, "Aviso"
    End If

    Set oNCOMCaptaGenerales = Nothing
    Set oNCOMContFunciones = Nothing
End If
End Sub

Private Sub cmdBuscar_Click()
If txtCodPers = "" Then Exit Sub
Dim oNCOMCaptaGenerales As COMNCaptaGenerales.NCOMCaptaGenerales
Set oNCOMCaptaGenerales = New COMNCaptaGenerales.NCOMCaptaGenerales
Dim rsCuentas As ADODB.Recordset
Set rsCuentas = New ADODB.Recordset
Dim rsCuenta As ADODB.Recordset
Set rsCuenta = New ADODB.Recordset
Dim oUCapCuenta As UCapCuenta
Set oUCapCuenta = New UCapCuenta

Set rsCuentas = oNCOMCaptaGenerales.obtenerCuentasPersonaServicioPago(txtCodPers.psCodigoPersona)

If Not (rsCuentas.EOF And rsCuentas.EOF) Then
    Do While Not rsCuentas.EOF
        frmCapMantenimientoCtas.lstCuentas.AddItem rsCuentas("cCtaCod") & Space(2) & rsCuentas("cRelacion") & Space(2) & Trim(rsCuentas("cEstado"))
        rsCuentas.MoveNext
    Loop
        
    Set oUCapCuenta = frmCapMantenimientoCtas.inicia
        
    If Not oUCapCuenta Is Nothing Then
        If oUCapCuenta.sCtaCod <> "" Then
            txtCuenta.CMAC = Left(oUCapCuenta.sCtaCod, 3)
            txtCuenta.Age = Mid(oUCapCuenta.sCtaCod, 4, 2)
            txtCuenta.Prod = Mid(oUCapCuenta.sCtaCod, 6, 3)
            txtCuenta.Cuenta = Right(oUCapCuenta.sCtaCod, 10)
            If Mid(oUCapCuenta.sCtaCod, 9, 1) = "1" Then
                txtMoneda = "SOLES"
            Else
                txtMoneda = "DÓLARES"
            End If
            txtComision.SetFocus
        Else
            MsgBox "La persona no tiene una cuenta disponible para registrar un Servicio de Pago.", vbInformation, "Aviso"
        End If
    End If
End If

Set oUCapCuenta = Nothing
Set rsCuentas = Nothing
Set oUCapCuenta = Nothing
Set oNCOMCaptaGenerales = Nothing

End Sub

Private Sub cmdEditar_Click()
If Trim(txtCodPers.Text) = "" Then Exit Sub
FRAConvenio.Enabled = False
FRACuenta.Enabled = True
FRAGeneral.Enabled = True
FRAComision.Enabled = True
FRAComisionTransf.Enabled = True 'RIRO20150430 ERS146-2014
FRAPenalidad.Enabled = True
cmdEditar.Visible = False
cmdBaja.Visible = False
cmdGuardar.Visible = True
End Sub

Private Sub cmdGuardar_Click()

If validarDatos = False Then Exit Sub

Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
Dim lsMovNro As String
Dim lnVerificarNombre, lnConfirmar As Long
Dim lsCuenta As String

Dim oNCOMContFunciones As NCOMContFunciones
Set oNCOMContFunciones = New NCOMContFunciones

lsCuenta = txtCuenta.CMAC & txtCuenta.Age & txtCuenta.Prod & txtCuenta.Cuenta
lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
lnVerificarNombre = oNCOMCaptaGenerales.confirmarNombreConvenio(txtCodPers.psCodigoPersona, UCase(txtNombre))

If fnAccion = 1 Then
    If lnVerificarNombre = -1 Then
        'RIRO20150430 ERS146-2014, ADD txtComisionTransf
        lnConfirmar = oNCOMCaptaGenerales.guardarServicioPago(txtCodPers.psCodigoPersona, _
                                                                        UCase(txtNombre), _
                                                                        UCase(txtDescripcion), lsCuenta, _
                                                                        txtComision, txtPenalidad, txtLimite, _
                                                                        Format(gdFecSis, "yyyyMMdd"), lsMovNro, , _
                                                                        txtComisionTransf.Text)

        If lnConfirmar > 0 Then
            MsgBox "Se guardo satisfactoriamente el convenio.", vbInformation, "Aviso"
            LimpiarCampos
        Else
            MsgBox "No se guardo el convenio.", vbCritical, "Aviso"
        End If
    Else
        MsgBox "El nombre del convenio ya existe.", vbCritical, "Aviso"
        txtNombre.SetFocus
        Exit Sub
    End If
    
Else
    If lnVerificarNombre = -1 Or lnVerificarNombre = fnIdSerPag Then
        'RIRO20150430 ERS146-2014, ADD txtComisionTransf
        lnConfirmar = oNCOMCaptaGenerales.guardarServicioPago(txtCodPers.psCodigoPersona, UCase(txtNombre), _
                                                              UCase(txtDescripcion), lsCuenta, _
                                                              txtComision, txtPenalidad, txtLimite, _
                                                              Format(gdFecSis, "yyyyMMdd"), lsMovNro, fnIdSerPag, _
                                                              txtComisionTransf.Text)

        If lnConfirmar > 0 Then
            MsgBox "Se modifico satisfactoriamente el convenio.", vbInformation, "Aviso"
            LimpiarCampos
            FRAConvenio.Enabled = True
            FRACuenta.Enabled = False
            FRAGeneral.Enabled = False
            FRAComision.Enabled = False
            FRAComisionTransf.Enabled = False 'RIRO20150430 ERS146-2014
            FRAPenalidad.Enabled = False
        Else
            MsgBox "No se modifico el convenio.", vbCritical, "Aviso"
        End If

    Else
        MsgBox "El nombre del convenio ya existe.", vbCritical, "Aviso"
        txtNombre.SetFocus
        Exit Sub
    End If
End If

Set oNCOMCaptaGenerales = Nothing
Set oNCOMContFunciones = Nothing
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
iniciar
LimpiarCampos
fnIdSerPag = 0
End Sub

Private Sub TxtCodPers_EmiteDatos()
Dim oCOMNPersona As COMNPersona.NCOMPersona
Set oCOMNPersona = New COMNPersona.NCOMPersona

Dim nExisteRol As Boolean

nExisteRol = oCOMNPersona.verficarRolPersona(11, txtCodPers.psCodigoPersona)
If nExisteRol = False Then
    txtCodPers = ""
    MsgBox "Persona no tiene el rol correspondiente.", vbInformation, "Aviso"
    Exit Sub
End If

If txtCodPers.psCodigoPersona = "" Then Exit Sub

If txtCodPers.psCodigoPersona = gsCodPersUser Then
    MsgBox "No se puede realizar la operación de si mismo.", vbInformation, "Aviso"
    txtCodPers = ""
    Exit Sub
End If
txtConvenio = txtCodPers.psDescripcion
If fnAccion = 2 Then
    Dim oNCOMCaptaGenerales As NCOMCaptaGenerales
    Set oNCOMCaptaGenerales = New NCOMCaptaGenerales
    Dim rsConvenios As ADODB.Recordset
    Set rsConvenios = New ADODB.Recordset
    Dim rsConvenio As ADODB.Recordset
    Set rsConvenio = New ADODB.Recordset
    
    Set rsConvenios = oNCOMCaptaGenerales.obtenerConveniosPersona(txtCodPers.psCodigoPersona)
     
    If Not (rsConvenios.EOF And rsConvenios.EOF) Then
        Do While Not rsConvenios.EOF
            frmCapMantenimientoConvenio.lstConvenios.AddItem rsConvenios!cCodSerPago & Space(2) & rsConvenios!cFecReg & Space(15) & rsConvenios!Id_SerPag
            rsConvenios.MoveNext
        Loop
        fnIdSerPag = frmCapMantenimientoConvenio.iniciarFormulario
    Else
        fnIdSerPag = 0
    End If
    
    If fnIdSerPag > 0 Then
        Set rsConvenio = oNCOMCaptaGenerales.obtenerConvenio(fnIdSerPag)
        If Not (rsConvenio.BOF And rsConvenio.EOF) Then
            txtCuenta.CMAC = Left(rsConvenio!cCtaCod, 3)
            txtCuenta.Age = Mid(rsConvenio!cCtaCod, 4, 2)
            txtCuenta.Prod = Mid(rsConvenio!cCtaCod, 6, 3)
            txtCuenta.Cuenta = Right(rsConvenio!cCtaCod, 10)
            If Mid(rsConvenio!cCtaCod, 9, 1) = "1" Then
                txtMoneda = "SOLES"
            Else
                txtMoneda = "DÓLARES"
            End If
            txtComision = Format$(rsConvenio!nComision, "##,##0.00")
            txtComisionTransf = Format$(rsConvenio!nComisionTrans, "##,##0.00")
            txtPenalidad = Format$(rsConvenio!nPenalidad, "##,##0.00") 'RIRO20150430 ERS146-2014
            txtLimite = rsConvenio!nLimMaxOpeDia
            txtNombre = rsConvenio!cNomSerPag
            txtDescripcion = rsConvenio!cDesSerPag
        Else
            limpiarCamposEditar
        End If
    Else
        limpiarCamposEditar
    End If
End If
Set oCOMNPersona = Nothing
End Sub

Private Sub txtComision_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtComision) Then
        txtComision = "0.00"
    End If
    KeyAscii = NumerosDecimales(txtComision, KeyAscii, , 2)
    If KeyAscii <> 13 Then Exit Sub
    If KeyAscii = 13 Then
       'txtPenalidad.SetFocus
       txtComisionTransf.SetFocus
    End If
End Sub
'RIRO20150430 ERS146-2014 *****************************
Private Sub txtComision_GotFocus()
    txtComision.SelStart = 0
    txtComision.SelLength = Len(txtComision.Text)
    txtComision.SetFocus
End Sub
Private Sub txtComision_LostFocus()
    If Not IsNumeric(txtComision) Then
        MsgBox "El número ingresado debe ser numérico", vbInformation, "Aviso"
        txtComision = "0.00"
        txtComision.SetFocus
    Else
        txtComision = Format(txtComision.Text, "#0.00")
    End If
End Sub
Private Sub txtComisionTransf_Change()
    If Not IsNumeric(txtComisionTransf.Text) Then
        txtComisionTransf.Text = "0.00"
    End If
End Sub
Private Sub txtComisionTransf_GotFocus()
    txtComisionTransf.SelStart = 0
    txtComisionTransf.SelLength = Len(txtComisionTransf.Text)
    txtComisionTransf.SetFocus
End Sub
Private Sub txtComisionTransf_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtComisionTransf, KeyAscii, , 2)
    If KeyAscii <> 13 Then Exit Sub
    If KeyAscii = 13 Then
       txtPenalidad.SetFocus
    End If
End Sub
Private Sub txtComisionTransf_LostFocus()
    If Not IsNumeric(txtComisionTransf) Then
        MsgBox "El número ingresado debe ser numérico", vbInformation, "Aviso"
        txtComisionTransf.Text = "0.00"
        txtComisionTransf.SetFocus
    Else
        txtComisionTransf.Text = Format(txtComisionTransf.Text, "#0.00")
    End If
End Sub

'END ERS146-2014 *************************************
Private Sub txtLimite_KeyPress(KeyAscii As Integer)
    If Not IsNumeric(txtLimite) Then
        txtLimite = ""
    End If
    KeyAscii = NumerosDecimales(txtLimite, KeyAscii, , 0)
    
    If KeyAscii <> 13 Then Exit Sub
    
    If KeyAscii = 13 Then
       txtNombre.SetFocus
    End If
End Sub
Private Sub txtPenalidad_KeyPress(KeyAscii As Integer)
If Not IsNumeric(txtPenalidad) Then
    txtPenalidad = ""
End If

KeyAscii = NumerosDecimales(txtPenalidad, KeyAscii)

If KeyAscii <> 13 Then Exit Sub

If KeyAscii = 13 Then
   txtLimite.SetFocus
End If
End Sub
