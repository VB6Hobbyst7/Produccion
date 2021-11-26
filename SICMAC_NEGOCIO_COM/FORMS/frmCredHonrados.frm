VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredHonrados 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Créditos Honrados"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   Icon            =   "frmCredHonrados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   10920
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   5655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   9975
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Registro de Créditos Honrados"
      TabPicture(0)   =   "frmCredHonrados.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraDatosCred"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "fraCredHorandos"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdQuitar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdSalir"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   9240
         TabIndex        =   13
         Top             =   5160
         Width           =   1335
      End
      Begin VB.CommandButton cmdQuitar 
         Appearance      =   0  'Flat
         Caption         =   "Quitar"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   5160
         Width           =   1335
      End
      Begin VB.Frame fraCredHorandos 
         Caption         =   "Créditos Honrados"
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
         Height          =   2655
         Left            =   120
         TabIndex        =   11
         Top             =   2400
         Width           =   10455
         Begin SICMACT.FlexEdit feCredHonrados 
            Height          =   2295
            Left            =   120
            TabIndex        =   15
            Top             =   240
            Width           =   10215
            _ExtentX        =   18018
            _ExtentY        =   4048
            Cols0           =   11
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Crédito-Titular-DOI-Moneda-Monto Honrado-Monto Devuelto-Saldo Pendiente-FecHoramiento-MovHonra-Codigo"
            EncabezadosAnchos=   "500-1800-3200-1200-1200-1200-1300-1300-0-0-0"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-L-L-L-R-R-R-L-L-C"
            FormatosEdit    =   "0-0-0-0-0-2-2-2-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin VB.Frame fraDatosCred 
         Caption         =   "Datos del Crédito"
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
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   480
         Width           =   10455
         Begin VB.TextBox txtMontoHonrado 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   350
            Left            =   5160
            TabIndex        =   16
            Top             =   1320
            Width           =   1575
         End
         Begin VB.ComboBox cmbMoneda 
            Height          =   315
            Left            =   4440
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   1320
            Width           =   600
         End
         Begin VB.CommandButton cmdCancelar 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   8640
            TabIndex        =   10
            Top             =   1320
            Width           =   1215
         End
         Begin VB.CommandButton cmdAceptar 
            Caption         =   "Aceptar"
            Height          =   375
            Left            =   7200
            TabIndex        =   9
            Top             =   1320
            Width           =   1335
         End
         Begin VB.CommandButton cmdBuscaCuenta 
            Caption         =   "..."
            Height          =   375
            Left            =   3840
            TabIndex        =   3
            Top             =   360
            Width           =   375
         End
         Begin SICMACT.ActXCodCta ActXCodCta 
            Height          =   375
            Left            =   120
            TabIndex        =   2
            Top             =   360
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   661
            Texto           =   "Crédito:"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
         Begin VB.Label lblDOI 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1200
            TabIndex        =   8
            Top             =   1320
            Width           =   1800
         End
         Begin VB.Label Label3 
            Caption         =   "Monto Honrado: "
            Height          =   255
            Left            =   3120
            TabIndex        =   7
            Top             =   1400
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "DOI: "
            Height          =   255
            Left            =   720
            TabIndex        =   6
            Top             =   1400
            Width           =   375
         End
         Begin VB.Label Label1 
            Caption         =   "Titular: "
            Height          =   255
            Left            =   600
            TabIndex        =   5
            Top             =   900
            Width           =   495
         End
         Begin VB.Label lblTitular 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   1200
            TabIndex        =   4
            Top             =   840
            Width           =   8640
         End
      End
   End
End
Attribute VB_Name = "frmCredHonrados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina:           frmCredHonrados
'***     Descripcion:      Registro de Créditos Honrados
'***     Creado por:       WIOR
'***     Maquina:          TIF-1-19
'***     Fecha-Tiempo:     10/10/2013 08:30:00 AM
'***     Ultima Modificacion: Creacion de la Opcion
'*****************************************************************************************
Option Explicit
Private fsPersCod As String
Private fnMontoHonrado As Double
Private fdFecCancelacion As Date
Private fnMoneda As Integer
Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    CargaDatos (Trim(ActXCodCta.NroCuenta))
End If
End Sub

Private Sub cmbMoneda_Change()
If Trim(cmbMoneda.Text) = "$." Then
    fnMoneda = 2
Else
    fnMoneda = 1
End If
End Sub

Private Sub cmdAceptar_Click()
On Error GoTo ErrorProceso
Dim oCredito As COMNCredito.NCOMCredito

If Not ValidaDatos Then Exit Sub

If MsgBox("Estas seguro de Registrar el Crédito?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    Set oCredito = New COMNCredito.NCOMCredito
    Call oCredito.OperacionCredHonrado(1, Trim(ActXCodCta.NroCuenta), CDbl(txtMontoHonrado.Text), 0, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), fdFecCancelacion)
    Set oCredito = Nothing
    MsgBox "Crédito Registrado Satisfactoriamente.", vbInformation, "Aviso"
    LimpiarDatos
    MostrarCreditosHonrados
End If

Exit Sub
ErrorProceso:
    Set oCredito = Nothing
    MsgBox "Error: " & err.Description, vbCritical, "Error en Proceso"
End Sub

Private Sub cmdBuscaCuenta_Click()
Dim oPersona As COMDPersona.UCOMPersona
Dim loPersCreditos As COMDCredito.DCOMCredito
Dim lrCreditos As New ADODB.Recordset
Dim loCuentas As COMDPersona.UCOMProdPersona

LimpiarDatos

On Error GoTo ControlError

Set oPersona = Nothing
Set oPersona = New COMDPersona.UCOMPersona
Set oPersona = frmBuscaPersona.Inicio
If oPersona Is Nothing Then Exit Sub

fsPersCod = oPersona.sPersCod
If Trim(oPersona.sPersCod) <> "" Then
    Set loPersCreditos = New COMDCredito.DCOMCredito
    Set lrCreditos = loPersCreditos.CreditosCanceladosSinHonramiento(oPersona.sPersCod)
    Set loPersCreditos = Nothing
End If

If Not (lrCreditos.EOF And lrCreditos.BOF) Then
Set loCuentas = New COMDPersona.UCOMProdPersona
    Set loCuentas = frmProdPersona.Inicio(oPersona.sPersNombre, lrCreditos)
    If loCuentas.sCtaCod <> "" Then
        ActXCodCta.NroCuenta = Mid(loCuentas.sCtaCod, 1, 18)
        ActXCodCta.Enabled = False
        Call ActXCodCta_KeyPress(13)
    End If
Else
    MsgBox "Persona No cuenta con créditos Cancelados", vbInformation, "Aviso"
End If
Set loCuentas = Nothing
Exit Sub
ControlError:
MsgBox "Error: " & err.Description, vbCritical, "Error"
End Sub

Private Sub cmdCancelar_Click()
LimpiarDatos
End Sub

Private Sub cmdQuitar_Click()
On Error GoTo ErrorProceso

Dim oCredito As New COMNCredito.NCOMCredito
Dim lnCodigo As Long
Dim lsCodigo As String
lsCodigo = ""
If MsgBox("Estás seguro de quitar el Crédito Honrado?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    'INICIO ORCR-20140913*********
    'lsCodigo = Trim(feCredHonrados.TextMatrix(feCredHonrados.row, 9))
    lsCodigo = Trim(feCredHonrados.TextMatrix(feCredHonrados.row, 10))
    'FIN ORCR-20140913************
    lnCodigo = oCredito.ObtenerCodigoCredHonrado(lsCodigo, gdFecSis)
    
    If lnCodigo = 0 Then
        MsgBox "No se puede quitar el Crédito, por superar el día de su registro o ya cuenta con pagos.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    Call oCredito.OperacionCredHonrado(3, , , , , , lnCodigo)
    MsgBox "Créditos Quitado Satisfactoriamente.", vbInformation, "Aviso"
    MostrarCreditosHonrados
End If

Exit Sub
ErrorProceso:
    Set oCredito = Nothing
    MsgBox "Error: " & err.Description, vbCritical, "Error en Proceso"
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub CargaMoneda()
cmbMoneda.Clear
cmbMoneda.AddItem "S/."
cmbMoneda.AddItem "$."
End Sub



Private Sub Form_Load()
LimpiarDatos
CargaMoneda
MostrarCreditosHonrados
End Sub

Private Sub LimpiarDatos()
ActXCodCta.NroCuenta = ""
ActXCodCta.Enabled = True
ActXCodCta.CMAC = gsCodCMAC
ActXCodCta.Age = gsCodAge
ActXCodCta.EnabledCMAC = False
lblTitular.Caption = ""
lblDOI.Caption = ""
cmbMoneda.Enabled = True
fnMontoHonrado = 0
fnMoneda = 0
txtMontoHonrado.Text = ""
fsPersCod = ""
fdFecCancelacion = "01/01/1900"
End Sub

Private Sub CargaDatos(ByVal psCtaCod As String)
Dim oCredito As New COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim nEstado As Integer

nEstado = oCredito.RecuperaEstadoCredito(psCtaCod)

If nEstado = gColocEstCancelado Or nEstado = gColocEstRecCanJud Or nEstado = gColocEstRecCanCast Then 'WIOR 20150824 gColocEstRecCanCast
    
    Set rsCredito = oCredito.DatosCreditoCanceladoAHonrar(psCtaCod)
    
    If Not (rsCredito.EOF And rsCredito.BOF) Then
        lblTitular.Caption = Trim(rsCredito!Cliente)
        lblDOI.Caption = Trim(rsCredito!NroDoc)
        txtMontoHonrado.Text = Format(rsCredito!MontoTotal, "###," & String(15, "#") & "#0.00")
        fnMoneda = CInt(Mid(psCtaCod, 9, 1))
        fdFecCancelacion = CDate(rsCredito!FechaCancelacion)
        If fnMoneda = 2 Then
            txtMontoHonrado.ForeColor = &H289556
            cmbMoneda.Text = "$."
        Else
            cmbMoneda.Text = "S/."
            txtMontoHonrado.ForeColor = vbBlue
        End If
        cmbMoneda.Enabled = False
    Else
        fdFecCancelacion = "01/01/1900"
        lblTitular.Caption = ""
        lblDOI.Caption = ""
        cmbMoneda.Enabled = True
    End If
Else
    MsgBox "Credito no Esta Cancelado.", vbInformation, "Aviso"
    LimpiarDatos
    Exit Sub
End If

End Sub

Private Sub txtMontoHonrado_Change()
If fnMoneda = 2 Then
    txtMontoHonrado.ForeColor = &H289556
Else
    txtMontoHonrado.ForeColor = vbBlue
End If
End Sub

Private Sub txtMontoHonrado_GotFocus()
fEnfoque txtMontoHonrado
End Sub

Private Sub txtMontoHonrado_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoHonrado, KeyAscii)
End Sub

Private Sub txtMontoHonrado_LostFocus()
If Len(Trim(txtMontoHonrado.Text)) = 0 Then
     txtMontoHonrado.Text = "0.00"
End If
txtMontoHonrado.Text = Format(txtMontoHonrado.Text, "###," & String(15, "#") & "#0.00")
End Sub

Private Sub MostrarCreditosHonrados()
Dim oCredito As New COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset
Dim i As Long

Set rsCredito = oCredito.ObtenerCreditosHonrados("%")
LimpiaFlex feCredHonrados
If Not (rsCredito.EOF And rsCredito.BOF) Then
    For i = 1 To rsCredito.RecordCount
        feCredHonrados.AdicionaFila
        feCredHonrados.TextMatrix(i, 1) = Trim(rsCredito!cCtaCod)
        feCredHonrados.TextMatrix(i, 2) = Trim(rsCredito!cPersNombre)
        feCredHonrados.TextMatrix(i, 3) = Trim(rsCredito!cPersIDnro)
        feCredHonrados.TextMatrix(i, 4) = Trim(rsCredito!Moneda)
        feCredHonrados.TextMatrix(i, 5) = Format(rsCredito!nMontoHonrado, "###," & String(15, "#") & "#0.00")
        feCredHonrados.TextMatrix(i, 6) = Format(rsCredito!nMontoDevuelto, "###," & String(15, "#") & "#0.00")
        'INICIO ORCR-20140913*********
        'feCredHonrados.TextMatrix(i, 7) = CDate(rsCredito!dFecHonramiento)
        'feCredHonrados.TextMatrix(i, 8) = Trim(rsCredito!cMovHonra)
        'feCredHonrados.TextMatrix(i, 9) = Trim(rsCredito!nCodCredHonrado)
        feCredHonrados.TextMatrix(i, 7) = Format(rsCredito!nMontoHonrado - rsCredito!nMontoDevuelto, "###," & String(15, "#") & "#0.00")
        feCredHonrados.TextMatrix(i, 8) = CDate(rsCredito!dFecHonramiento)
        feCredHonrados.TextMatrix(i, 9) = Trim(rsCredito!cMovHonra)
        feCredHonrados.TextMatrix(i, 10) = Trim(rsCredito!nCodCredHonrado)
        'FIN ORCR-20140913************
        rsCredito.MoveNext
    Next i
    feCredHonrados.row = 1
    feCredHonrados.TopRow = 1
End If
End Sub


Private Function ValidaDatos() As Boolean
ValidaDatos = True

If Len(ActXCodCta.NroCuenta) < 18 Then
    MsgBox "Ingrese la Cuenta", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

If Trim(lblTitular.Caption) = "" Then
    MsgBox "Cargue el Crédito antes de Grabar.", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

If txtMontoHonrado.Text = "0.00" Or txtMontoHonrado.Text = "" Then
    MsgBox "Ingrese el Monto Honrado", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If
End Function
