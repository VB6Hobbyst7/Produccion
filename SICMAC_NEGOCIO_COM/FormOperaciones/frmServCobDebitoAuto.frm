VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmServCobDebitoAuto 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Débitos Automáticos"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9495
   Icon            =   "frmServCobDebitoAuto.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6750
   ScaleWidth      =   9495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   6495
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   11456
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Débitos Automáticos"
      TabPicture(0)   =   "frmServCobDebitoAuto.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "cmdElimDeb"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtCuenta"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdVerReglas"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "FraCliente"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "SSTabServCred"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "FraComision"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdAceptar"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCerrar"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCancelar"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdBuscar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   325
         Left            =   3840
         TabIndex        =   17
         Top             =   520
         Width           =   375
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   360
         Left            =   6600
         TabIndex        =   9
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton cmdCerrar 
         Cancel          =   -1  'True
         Caption         =   "Cerrar"
         Height          =   360
         Left            =   7920
         TabIndex        =   10
         Top             =   6000
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   360
         Left            =   5280
         TabIndex        =   8
         Top             =   6000
         Width           =   1215
      End
      Begin VB.Frame FraComision 
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   5880
         Width           =   1695
         Begin VB.Label lblComision 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   960
            TabIndex        =   16
            Top             =   150
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Comisión:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   180
            Width           =   735
         End
      End
      Begin TabDlg.SSTab SSTabServCred 
         Height          =   2580
         Left            =   120
         TabIndex        =   5
         Top             =   3240
         Width           =   9015
         _ExtentX        =   15901
         _ExtentY        =   4551
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   8
         TabHeight       =   520
         TabCaption(0)   =   "Servicios"
         TabPicture(0)   =   "frmServCobDebitoAuto.frx":0326
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "FEServicios"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "cmdServAgregar"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).Control(2)=   "cmdServEliminar"
         Tab(0).Control(2).Enabled=   0   'False
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "Créditos"
         TabPicture(1)   =   "frmServCobDebitoAuto.frx":0342
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "FECreditos"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "cmdCredAgregar"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "cmdCredEliminar"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin VB.CommandButton cmdCredEliminar 
            Caption         =   "Eliminar"
            Height          =   345
            Left            =   -73680
            TabIndex        =   7
            Top             =   2080
            Width           =   1095
         End
         Begin VB.CommandButton cmdCredAgregar 
            Caption         =   "Agregar"
            Height          =   345
            Left            =   -74880
            TabIndex        =   6
            Top             =   2080
            Width           =   1095
         End
         Begin VB.CommandButton cmdServEliminar 
            Caption         =   "Eliminar"
            Height          =   345
            Left            =   1320
            TabIndex        =   3
            Top             =   2080
            Width           =   1095
         End
         Begin VB.CommandButton cmdServAgregar 
            Caption         =   "Agregar"
            Height          =   345
            Left            =   120
            TabIndex        =   2
            Top             =   2080
            Width           =   1095
         End
         Begin SICMACT.FlexEdit FECreditos 
            Height          =   1515
            Left            =   -74880
            TabIndex        =   13
            Top             =   480
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   2672
            Cols0           =   6
            HighLight       =   1
            VisiblePopMenu  =   -1  'True
            EncabezadosNombres=   "#-Crédito-Saldo Cap.-Titular-Monto Max-FecVencCuota"
            EncabezadosAnchos=   "250-2100-1400-3450-1200-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            EncabezadosAlineacion=   "-0-0--0-0"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            SelectionMode   =   1
            ColWidth0       =   255
            RowHeight0      =   300
         End
         Begin SICMACT.FlexEdit FEServicios 
            Height          =   1575
            Left            =   120
            TabIndex        =   19
            Top             =   480
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   2672
            Cols0           =   8
            HighLight       =   1
            VisiblePopMenu  =   -1  'True
            EncabezadosNombres=   "#-Empresa-Convenio-Cod. Usuario-Día Pago 1-Día Pago 2-Monto Max-cId"
            EncabezadosAnchos=   "250-2100-1750-1150-970-970-1200-0"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "L-L-L-C-C-C-R-C"
            FormatosEdit    =   "0-0-2-0-2-0-2-0"
            TextArray0      =   "#"
            SelectionMode   =   1
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            ColWidth0       =   255
            RowHeight0      =   300
            TipoBusPersona  =   1
         End
      End
      Begin VB.Frame FraCliente 
         Height          =   2055
         Left            =   120
         TabIndex        =   11
         Top             =   1080
         Width           =   9015
         Begin SICMACT.FlexEdit FECliente 
            Height          =   1755
            Left            =   120
            TabIndex        =   12
            Top             =   240
            Width           =   8775
            _ExtentX        =   15478
            _ExtentY        =   3096
            Cols0           =   9
            HighLight       =   1
            AllowUserResizing=   3
            VisiblePopMenu  =   -1  'True
            EncabezadosNombres=   "#-Codigo-Nombre-Relacion-Direccion-ID-Firma Oblig-Grupo-Presente"
            EncabezadosAnchos=   "250-1500-3200-1500-0-0-0-1000-1000"
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-8"
            ListaControles  =   "0-0-0-0-0-0-0-0-4"
            EncabezadosAlineacion=   "C-L-L-L-C-C-C-L-C"
            FormatosEdit    =   "0-0-0-0-0-0-0-0-0"
            TextArray0      =   "#"
            lbEditarFlex    =   -1  'True
            lbFlexDuplicados=   0   'False
            lbUltimaInstancia=   -1  'True
            TipoBusqueda    =   3
            ColWidth0       =   255
            RowHeight0      =   300
            TipoBusPersona  =   1
         End
      End
      Begin VB.CommandButton cmdVerReglas 
         Caption         =   "Ver Reglas"
         Height          =   325
         Left            =   7800
         TabIndex        =   1
         Top             =   520
         Width           =   1215
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   120
         TabIndex        =   0
         Top             =   480
         Width           =   3630
         _ExtentX        =   6403
         _ExtentY        =   661
         Texto           =   "Cuenta N°:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.CommandButton cmdElimDeb 
         Caption         =   "Eliminar Débitos"
         Height          =   360
         Left            =   120
         TabIndex        =   18
         Top             =   6000
         Visible         =   0   'False
         Width           =   1815
      End
   End
End
Attribute VB_Name = "frmServCobDebitoAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'** Nombre : frmSegTarjetaAfiliacion
'** Descripción : Formulario para afiliar seleccionar una cuenta para asociar al débito
'**               automático para pagos de servicios de recaudo o créditos creado segun
'**               TI-ERS144-2014
'** Creación : JUEZ, 20150130 09:00:00 AM
'****************************************************************************************

Option Explicit

Dim oNCapGen As COMNCaptaGenerales.NCOMCaptaGenerales
Dim bProcesoNuevo As Boolean
Dim strReglas As String
Dim MatServCred As Variant
Dim i As Integer
Dim nCodID As Long

Public Sub inicia(ByVal psOpeCod As CaptacOperacion)
Dim oDGen As COMDConstSistema.DCOMGeneral
    gsOpeCod = psOpeCod
    txtCuenta.CMAC = "109"
    txtCuenta.Age = gsCodAge
    nCodID = 0
    HabilitaControles True, False
    If psOpeCod = gServCobDebitoAuto Then
        Set oDGen = New COMDConstSistema.DCOMGeneral
        lblComision.Caption = Format(oDGen.GetParametro(2000, 2157), "#,##0.00") & " "
        Caption = "Registro de Débitos Automáticos"
    Else
        FraCliente.Visible = False
        cmdVerReglas.Visible = False
        lblComision.Visible = False
        cmdElimDeb.Visible = True
        SSTabServCred.Top = 1080
        cmdElimDeb.Top = 3840
        cmdAceptar.Top = 3840
        cmdCancelar.Top = 3840
        cmdCerrar.Top = 3840
        SSTab1.Height = 4335
        Me.Height = 5055
        Me.Caption = "Edición de Débitos Automáticos"
    End If
Set MatServCred = Nothing
Me.Show 1
End Sub

Public Sub IniciaConsulta(ByVal psCtaCod As String)
gsOpeCod = gServCobDebitoAutoEdit
FraCliente.Visible = False
cmdVerReglas.Visible = False
lblComision.Visible = False
cmdElimDeb.Visible = False
SSTabServCred.Top = 1080
cmdElimDeb.Top = 3840
cmdAceptar.Top = 3840
cmdCancelar.Top = 3840
cmdCerrar.Top = 3840
SSTab1.Height = 4335
Me.Height = 5055
Me.Caption = "Consulta de Débitos Automáticos"
txtCuenta.NroCuenta = psCtaCod
ObtieneDatosCuenta psCtaCod
If FEServicios.TextMatrix(1, 0) = "" And FECreditos.TextMatrix(1, 0) = "" Then
    Unload Me
    Exit Sub
End If
cmdServAgregar.Visible = False
cmdServEliminar.Visible = False
cmdCredAgregar.Visible = False
cmdCredEliminar.Visible = False
cmdAceptar.Visible = False
cmdCancelar.Visible = False
Me.Show 1
End Sub

Private Sub CmdAceptar_Click()
    'ANDE 20180228 Comprobar si tiene acceso la opción como RFIII
    Dim bPermitirEjecucionOperacion As Boolean
    Dim oCaja As New COMNCajaGeneral.NCOMCajaGeneral
    bPermitirEjecucionOperacion = oCaja.PermitirEjecucionOperacion(gsCodUser, gsOpeCod, "0")
    If Not bPermitirEjecucionOperacion Then
        End
    End If
    'fin Comprobacion si es RFIII

Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oNContFunc As COMNContabilidad.NCOMContFunciones
Dim rsServ As ADODB.Recordset
Dim rsCred As ADODB.Recordset
Dim lsMensaje As String
Dim sMovNro As String
Dim VistoElectronico As frmVistoElectronico
Dim lbResultadoVisto As Boolean
Dim lbRegistro As Boolean
Dim lsMsjErr As String
Dim lsBoleta As String

If FEServicios.TextMatrix(1, 0) = "" And FECreditos.TextMatrix(1, 0) = "" Then
    MsgBox "Es necesario que se agregue al menos un servicio o un crédito", vbInformation, "Aviso"
    Exit Sub
End If

Set oNCapGen = New COMNCaptaGenerales.NCOMCaptaGenerales

Set rsServ = FEServicios.GetRsNew
If Not rsServ Is Nothing Then
    lsMensaje = oNCapGen.ValidaServCobDebitoAutoDet(rsServ, txtCuenta.NroCuenta, gServConvenio)
    If lsMensaje <> "" Then
        MsgBox lsMensaje, vbInformation, "Aviso"
        Exit Sub
    End If
End If

Set rsCred = FECreditos.GetRsNew
If Not rsCred Is Nothing Then
    lsMensaje = oNCapGen.ValidaServCobDebitoAutoDet(rsCred, txtCuenta.NroCuenta, gServCredito)
    If lsMensaje <> "" Then
        MsgBox lsMensaje, vbInformation, "Aviso"
        Exit Sub
    End If
End If

Set oNCapGen = Nothing

If gsOpeCod = gServCobDebitoAuto Then
    If bProcesoNuevo = True Then
        If ValidarReglasPersonas = False Then
            MsgBox "Las personas seleccionadas no tienen suficientes poderes para realizar el registro", vbInformation
            Exit Sub
        End If
    End If
End If

If MsgBox("Se va a " & IIf(gsOpeCod = gServCobDebitoAuto, "registrar", "actualizar") & " el Servicio de Débito Automático, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub

Set VistoElectronico = New frmVistoElectronico
lbResultadoVisto = False
lbResultadoVisto = VistoElectronico.Inicio(3, gsOpeCod)
If Not lbResultadoVisto Then
    Exit Sub
End If

Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    If gsOpeCod = gServCobDebitoAuto Then
        lbRegistro = oNCapMov.RegistrarServCobDebitoAuto(gdFecSis, gsCodAge, gsCodUser, txtCuenta.NroCuenta, CDbl(lblComision.Caption), gsOpeCod, rsServ, rsCred, lsMsjErr, gsNomAge, gbImpTMU, lsBoleta)
    Else
        lbRegistro = oNCapMov.ActualizarServCobDebitoAuto(gdFecSis, gsCodAge, gsCodUser, txtCuenta.NroCuenta, gsOpeCod, rsServ, rsCred, lsMsjErr, gsNomAge, gbImpTMU, lsBoleta)
    End If
Set oNCapMov = Nothing
If Not lbRegistro Then
    MsgBox "Ocurrió un inconveniente en " & IIf(gsOpeCod = gServCobDebitoAuto, "el registro", "la actualización") & ". Intentar nuevamente. " & lsMsjErr, vbInformation, "Aviso"
    Exit Sub
End If

Dim nFicSal As Integer
    Do
        nFicSal = FreeFile
        Open sLpt For Output As nFicSal
        Print #nFicSal, lsBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
        Close #nFicSal
    Loop Until MsgBox("¿Desea reimprimir el voucher?", vbQuestion + vbYesNo, "Aviso") = vbNo

MsgBox "El Servicio de Débito Automático fue " & IIf(gsOpeCod = gServCobDebitoAuto, "registrado", "actualizado") & " correctamente", vbInformation, "Aviso"
'INICIO JHCU ENCUESTA 16-10-2019
Encuestas gsCodUser, gsCodAge, "ERS0292019", gsOpeCod
'FIN
cmdCancelar_Click
End Sub

Private Sub cmdBuscar_Click()
Dim clsPers As COMDPersona.UCOMPersona

Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio

If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As New ADODB.Recordset
    Dim clsCap As COMDCaptaGenerales.DCOMCaptaGenerales
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuenta
    
    sPers = clsPers.sPersCod
    Set clsCap = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsPers = clsCap.GetCuentasPersona(sPers, , True, True)

    Set clsCap = Nothing
    If Not (rsPers.EOF And rsPers.EOF) Then
        Do While Not rsPers.EOF
            sCta = rsPers("cCtaCod")
            sRelac = rsPers("cRelacion")
            sEstado = Trim(rsPers("cEstado"))
            frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
            rsPers.MoveNext
        Loop
        Set clsCuenta = New UCapCuenta
        Set clsCuenta = frmCapMantenimientoCtas.inicia
        If Not clsCuenta Is Nothing Then
            If clsCuenta.sCtaCod <> "" Then
                txtCuenta.NroCuenta = clsCuenta.sCtaCod
                If txtCuenta.Prod = gCapPlazoFijo Then
                    MsgBox "Solo está permitido las cuentas de Ahorro y CTS", vbInformation, "Aviso"
                    txtCuenta.NroCuenta = ""
                    txtCuenta.CMAC = "109"
                    txtCuenta.Age = gsCodAge
                End If
                txtCuenta.SetFocusCuenta
            End If
        End If
        Set clsCuenta = Nothing
    Else
        MsgBox "Persona no posee ninguna cuenta de captaciones o no tiene cuentas en la agencia.", vbInformation, "Aviso"
    End If
    rsPers.Close
    Set rsPers = Nothing
End If
Set clsPers = Nothing
txtCuenta.SetFocusCuenta
End Sub

Private Sub cmdCancelar_Click()
Dim oDGen As COMDConstSistema.DCOMGeneral
Set oDGen = New COMDConstSistema.DCOMGeneral
    txtCuenta.NroCuenta = ""
    txtCuenta.CMAC = "109"
    txtCuenta.Age = gsCodAge
    HabilitaControles True, False
    LimpiaFlex FECliente
    LimpiaFlex FEServicios
    LimpiaFlex FECreditos
    lblComision.Caption = Format(oDGen.GetParametro(2000, 2157), "#,##0.00") & " "
    nCodID = 0
    Set MatServCred = Nothing
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdCredAgregar_Click()
    ReDim MatServCred(0, 0)
    frmServCobDebitoAutoServCred.inicia gServCredito, MatServCred
    If MatServCred(0, 0) <> "" Then
        If Mid(CStr(MatServCred(0, 0)), 9, 1) <> Mid(txtCuenta.NroCuenta, 9, 1) Then
            MsgBox "El crédito ingresado tiene una moneda diferente a la de la cuenta a debitar.", vbInformation, "Aviso"
            Exit Sub
        End If
        For i = 1 To FECreditos.Rows - 1
            If CStr(FECreditos.TextMatrix(i, 1)) = CStr(MatServCred(0, 0)) Then
                MsgBox "El crédito ingresado ya se encuentra en la lista", vbInformation, "Aviso"
                Exit Sub
            End If
        Next i
        FECreditos.AdicionaFila
        FECreditos.TextMatrix(FECreditos.row, 0) = FECreditos.row
        FECreditos.TextMatrix(FECreditos.row, 1) = MatServCred(0, 0)
        FECreditos.TextMatrix(FECreditos.row, 2) = Format(MatServCred(0, 1), "#,##0.00")
        FECreditos.TextMatrix(FECreditos.row, 3) = MatServCred(0, 2)
        FECreditos.TextMatrix(FECreditos.row, 4) = MatServCred(0, 3)
        FECreditos.TextMatrix(FECreditos.row, 5) = MatServCred(0, 4)
    End If
    Set MatServCred = Nothing
End Sub

Private Sub cmdCredEliminar_Click()
    If FECreditos.TextMatrix(FECreditos.row, 0) <> "" Then
        If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(FECreditos.row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            FECreditos.EliminaFila FECreditos.row
        End If
    End If
End Sub

Private Sub cmdElimDeb_Click()
If MsgBox("¿Está seguro de eliminar y desafiliar todos los servicios y créditos para esta cuenta?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim oNCapGen As COMNCaptaGenerales.NCOMCaptaMovimiento
    
    Set oNCapGen = New COMNCaptaGenerales.NCOMCaptaMovimiento
    oNCapGen.EliminarServCobDebitoAuto gdFecSis, gsCodAge, gsCodUser, txtCuenta.NroCuenta, gsOpeCod, nCodID
    Set oNCapGen = Nothing
    
    MsgBox "Se ha eliminado el Servicio de Débito Automático para la cuenta " & txtCuenta.NroCuenta, vbInformation, "Aviso"
    Call cmdCerrar_Click
End If
End Sub

Private Sub cmdServAgregar_Click()
    ReDim MatServCred(0, 0)
    frmServCobDebitoAutoServCred.inicia 1, MatServCred
    
    If MatServCred(0, 0) <> "" Then
        If CStr(MatServCred(0, 7)) <> Mid(txtCuenta.NroCuenta, 9, 1) Then
            MsgBox "La cuenta del convenio ingresado tiene una moneda diferente a la de la cuenta a debitar.", vbInformation, "Aviso"
            Exit Sub
        End If
        For i = 1 To FEServicios.Rows - 1
            If CStr(FEServicios.TextMatrix(i, 2)) = CStr(MatServCred(0, 1)) And CStr(FEServicios.TextMatrix(i, 3)) = CStr(MatServCred(0, 2)) Then
                MsgBox "El convenio y el cliente ingresado ya se encuentran en la lista", vbInformation, "Aviso"
                Exit Sub
            End If
        Next i
        FEServicios.AdicionaFila
        FEServicios.TextMatrix(FEServicios.row, 0) = FEServicios.row
        FEServicios.TextMatrix(FEServicios.row, 1) = MatServCred(0, 0)
        FEServicios.TextMatrix(FEServicios.row, 2) = MatServCred(0, 1)
        FEServicios.TextMatrix(FEServicios.row, 3) = MatServCred(0, 2)
        FEServicios.TextMatrix(FEServicios.row, 4) = MatServCred(0, 3)
        FEServicios.TextMatrix(FEServicios.row, 5) = MatServCred(0, 4)
        FEServicios.TextMatrix(FEServicios.row, 6) = MatServCred(0, 5)
        FEServicios.TextMatrix(FEServicios.row, 7) = MatServCred(0, 6)
    End If
    Set MatServCred = Nothing
End Sub

Private Sub cmdServEliminar_Click()
    If FEServicios.TextMatrix(FEServicios.row, 0) <> "" Then
        If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(FEServicios.row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
            FEServicios.EliminaFila FECreditos.row
        End If
    End If
End Sub

Private Sub cmdVerReglas_Click()
    If strReglas <> "" Then
        Call frmCapVerReglas.inicia(strReglas)
    Else
        MsgBox "Cuenta no tiene reglas definidas", vbInformation, "Aviso"
    End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    KeyAscii = SoloNumeros(KeyAscii)
    If KeyAscii = 13 Then
        Dim sCta As String
        sCta = txtCuenta.NroCuenta
        If Len(sCta) = 18 Then
            ObtieneDatosCuenta sCta
        Else
            MsgBox "Ingresar correctamente la cuenta", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim rsCta As ADODB.Recordset, rsRel As New ADODB.Recordset, rsDeb As ADODB.Recordset
Dim nRow As Long
Dim sMsg As String, sPersona As String

nCodID = 0
Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
    sMsg = clsCap.ValidaCuentaOperacion(sCuenta)
Set clsCap = Nothing
If sMsg = "" Then
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsCta = New ADODB.Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        Set rsDeb = clsMant.ValidaServCobDebitoAuto(sCuenta)
        If Not (rsDeb.EOF And rsDeb.BOF) Then
            nCodID = rsDeb("nCodID")
            If gsOpeCod = gServCobDebitoAuto Then
                MsgBox "La cuenta " & sCuenta & " ya se encuentra registrada en el Servicio de Débito Automático", vbInformation, "Aviso"
                Exit Sub
            Else
                CargarServCredDebitoAuto rsDeb("nMovNro")
                HabilitaControles False, True
                Exit Sub
            End If
        Else
            If gsOpeCod = gServCobDebitoAutoEdit Then
                MsgBox "La cuenta " & sCuenta & " no está registrada en el Servicio de Débito Automático", vbInformation, "Aviso"
                Exit Sub
            End If
        End If
        strReglas = IIf(IsNull(rsCta!cReglas), "", rsCta!cReglas)
        FECliente.lbEditarFlex = True
    
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        
        sPersona = ""
        
        Dim dlsMant As COMDCaptaGenerales.DCOMCaptaGenerales
        Set dlsMant = New COMDCaptaGenerales.DCOMCaptaGenerales
                
        Do While Not rsRel.EOF
            If rsRel("cPersCod") = gsCodPersUser Then
                MsgBox "Ud. No puede hacer operaciones con sus propias cuentas.", vbInformation, "Aviso"
                txtCuenta.NroCuenta = ""
                txtCuenta.CMAC = "109"
                txtCuenta.Age = gsCodAge
                txtCuenta.SetFocusCuenta
                Exit Sub
            End If
            If sPersona <> rsRel("cPersCod") Then
                FECliente.AdicionaFila
                nRow = FECliente.Rows - 1
                FECliente.TextMatrix(nRow, 1) = rsRel("cPersCod")
                FECliente.TextMatrix(nRow, 2) = UCase(PstaNombre(rsRel("Nombre")))
                FECliente.TextMatrix(nRow, 3) = UCase(rsRel("Relacion")) & Space(50) & Trim(rsRel("nPrdPersRelac"))
                FECliente.TextMatrix(nRow, 4) = rsRel("Direccion") & ""
                FECliente.TextMatrix(nRow, 5) = rsRel("ID N°")
                 
                If rsRel("cGrupo") <> "" Then
                    bProcesoNuevo = True
                    FECliente.TextMatrix(nRow, 7) = rsRel("cGrupo")
                Else
                    bProcesoNuevo = False
                    FECliente.TextMatrix(nRow, 6) = IIf(IsNull(rsRel("cobligatorio")) Or rsRel("cobligatorio") = "N", "NO", IIf(rsRel("cobligatorio") = "S", "SI", "OPCIONAL"))
                End If
                sPersona = rsRel("cPersCod")
            End If
            rsRel.MoveNext
        Loop
        HabilitaControles False, True
        cmdAceptar.SetFocus
    End If
Else
    MsgBox sMsg, vbInformation, "Operacion"
    'txtCuenta.SetFocus 'ande 20171013 corrección de error de llamada de procedimiento inválido
End If
Set clsMant = Nothing
End Sub
Private Function HabilitaControles(ByVal pbHabCarga As Boolean, ByVal pbHabDatos As Boolean)
    txtCuenta.Enabled = pbHabCarga
    cmdBuscar.Enabled = pbHabCarga
    SSTabServCred.Enabled = pbHabDatos
    cmdAceptar.Enabled = pbHabDatos
    If gsOpeCod = gServCobDebitoAuto Then
        cmdVerReglas.Enabled = pbHabDatos
        FraCliente.Enabled = pbHabDatos
    Else
        cmdElimDeb.Enabled = pbHabDatos
    End If
End Function
Private Function ValidarReglasPersonas() As Boolean
 Dim sReglas() As String
    Dim sGrupos() As String
    Dim sTemporal As String
    Dim v1, v2 As Variant
    Dim bAprobado As Boolean
    Dim intRegla, i, J As Integer
    
    If Trim(strReglas) = "" Then
        ValidarReglasPersonas = False
        Exit Function
    End If
    sReglas = Split(strReglas, "-")
    For i = 1 To FECliente.Rows - 1
        If FECliente.TextMatrix(i, 8) = "." Then
            If J = 0 Then
               sTemporal = sTemporal & FECliente.TextMatrix(i, 7)
            Else
               sTemporal = sTemporal & "," & FECliente.TextMatrix(i, 7)
            End If
            J = J + 1
        End If
    Next
    If Trim(sTemporal) = "" Then
        ValidarReglasPersonas = False
        Exit Function
    End If
    sGrupos = Split(sTemporal, ",")
    For Each v1 In sReglas
        bAprobado = True
        For Each v2 In sGrupos
            If InStr(CStr(v1), CStr(v2)) = 0 Then
                bAprobado = False
                Exit For
            End If
        Next
        If bAprobado Then
            If UBound(sGrupos) = UBound(Split(CStr(v1), "+")) Then
                Exit For
            Else
                bAprobado = False
            End If
        End If
    Next
    ValidarReglasPersonas = bAprobado
End Function
Private Sub CargarServCredDebitoAuto(ByVal pnMovNro As Long)
Dim oDCapGen As COMDCaptaGenerales.DCOMCaptaGenerales
Dim rsCobDet As ADODB.Recordset
    Set oDCapGen = New COMDCaptaGenerales.DCOMCaptaGenerales
    Set rsCobDet = oDCapGen.ObtenerServCobDebitoAutoDet(pnMovNro, gServConvenio)
    Do While Not rsCobDet.EOF
        FEServicios.AdicionaFila
        i = FEServicios.row
        FEServicios.TextMatrix(i, 1) = rsCobDet!cPersNombre
        FEServicios.TextMatrix(i, 2) = rsCobDet!cCodConvenio
        FEServicios.TextMatrix(i, 3) = rsCobDet!cCodCliente
        FEServicios.TextMatrix(i, 4) = rsCobDet!nDia1
        FEServicios.TextMatrix(i, 5) = rsCobDet!nDia2
        FEServicios.TextMatrix(i, 6) = Format(rsCobDet!nMontoMax, "#,##0.00")
        FEServicios.TextMatrix(i, 7) = rsCobDet!cId
        rsCobDet.MoveNext
    Loop
    
    Set rsCobDet = oDCapGen.ObtenerServCobDebitoAutoDet(pnMovNro, gServCredito)
    Do While Not rsCobDet.EOF
        FECreditos.AdicionaFila
        i = FECreditos.row
        FECreditos.TextMatrix(i, 1) = rsCobDet!cCtaCod
        FECreditos.TextMatrix(i, 2) = Format(rsCobDet!nSaldo, "#,##0.00")
        FECreditos.TextMatrix(i, 3) = rsCobDet!cPersNombre
        FECreditos.TextMatrix(i, 4) = Format(rsCobDet!nMontoMax, "#,##0.00")
        FECreditos.TextMatrix(i, 5) = rsCobDet!dFecVenc
        rsCobDet.MoveNext
    Loop
End Sub
