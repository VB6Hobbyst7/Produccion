VERSION 5.00
Begin VB.Form frmEnvioEstadoCtaAfiliacion 
   Caption         =   "Afiliación Envio de Estado de Cuenta"
   ClientHeight    =   6990
   ClientLeft      =   7260
   ClientTop       =   3720
   ClientWidth     =   7230
   Icon            =   "frmEnvioEstadoCtaAfiliacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   7230
   Begin VB.Frame Frame1 
      Caption         =   "Solicitud de Afiliación"
      Height          =   6735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6975
      Begin VB.CommandButton cmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   240
         TabIndex        =   24
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5640
         TabIndex        =   23
         Top             =   6240
         Width           =   975
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   375
         Left            =   4440
         TabIndex        =   22
         Top             =   6240
         Width           =   975
      End
      Begin VB.Frame Frame2 
         Caption         =   " Datos de Cliente"
         ForeColor       =   &H00FF0000&
         Height          =   1695
         Left            =   240
         TabIndex        =   3
         Top             =   360
         Width           =   6375
         Begin VB.TextBox txtDireccion 
            Height          =   285
            Left            =   1080
            TabIndex        =   14
            Top             =   1200
            Width           =   5055
         End
         Begin VB.TextBox txtCorreo 
            Height          =   285
            Left            =   3600
            TabIndex        =   12
            Top             =   840
            Width           =   2535
         End
         Begin VB.TextBox TxtBuscarPersona 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1080
            TabIndex        =   5
            Top             =   230
            Width           =   2055
         End
         Begin VB.CommandButton cmbBuscar 
            Caption         =   "Buscar"
            Height          =   300
            Left            =   3240
            TabIndex        =   4
            Top             =   200
            Width           =   735
         End
         Begin VB.Label Label5 
            Caption         =   "Dirección:"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   1200
            Width           =   975
         End
         Begin VB.Label Label4 
            Caption         =   "Correo:"
            Height          =   255
            Left            =   2880
            TabIndex        =   11
            Top             =   840
            Width           =   615
         End
         Begin VB.Label lblDOI 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1080
            TabIndex        =   10
            Top             =   795
            Width           =   1575
         End
         Begin VB.Label lblNombre 
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1080
            TabIndex        =   9
            Top             =   525
            Width           =   5055
         End
         Begin VB.Label Label3 
            Caption         =   "DOI:"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   820
            Width           =   615
         End
         Begin VB.Label Label2 
            Caption         =   "Cliente:"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   525
            Width           =   615
         End
         Begin VB.Label Label1 
            Caption         =   "Código:"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Envio"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   1335
         Left            =   240
         TabIndex        =   2
         Top             =   4800
         Width           =   6375
         Begin VB.OptionButton OptFisico 
            Caption         =   "Recibir la informacion al domicilio por medio fisico."
            Height          =   255
            Left            =   480
            TabIndex        =   21
            Top             =   960
            Width           =   3855
         End
         Begin VB.OptionButton OptElectronico 
            Caption         =   "Recibir la informacion al correo por medio electronico."
            Height          =   255
            Left            =   480
            TabIndex        =   20
            Top             =   720
            Value           =   -1  'True
            Width           =   4335
         End
         Begin VB.Label Label6 
            Caption         =   $"frmEnvioEstadoCtaAfiliacion.frx":030A
            Height          =   495
            Left            =   240
            TabIndex        =   19
            Top             =   240
            Width           =   5895
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Productos"
         Enabled         =   0   'False
         ForeColor       =   &H00FF0000&
         Height          =   2535
         Left            =   240
         TabIndex        =   1
         Top             =   2160
         Width           =   6375
         Begin VB.Frame Frame5 
            Caption         =   "Créditos"
            Height          =   1695
            Left            =   240
            TabIndex        =   17
            Top             =   720
            Width           =   6015
            Begin SICMACT.FlexEdit FEProductos 
               Height          =   1335
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   5775
               _ExtentX        =   10186
               _ExtentY        =   2249
               Cols0           =   5
               HighLight       =   1
               AllowUserResizing=   3
               RowSizingMode   =   1
               EncabezadosNombres=   "--Cuenta-SubProducto-Medio Envío"
               EncabezadosAnchos=   "350-400-2000-2500-1200"
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
                  Name            =   "Tahoma"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ColumnasAEditar =   "X-1-X-X-X"
               TextStyleFixed  =   4
               ListaControles  =   "0-4-0-0-0"
               BackColor       =   16777215
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               BackColorControl=   -2147483643
               EncabezadosAlineacion=   "C-C-C-C-C"
               FormatosEdit    =   "0-0-0-0-0"
               SelectionMode   =   1
               lbUltimaInstancia=   -1  'True
               ColWidth0       =   345
               RowHeight0      =   300
               ForeColorFixed  =   -2147483635
               CellBackColor   =   16777215
            End
         End
         Begin VB.OptionButton OptAhorros 
            Caption         =   "Ahorros"
            Height          =   195
            Left            =   1440
            TabIndex        =   16
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton OptCreditos 
            Caption         =   "Créditos"
            Height          =   195
            Left            =   240
            TabIndex        =   15
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
      End
   End
End
Attribute VB_Name = "frmEnvioEstadoCtaAfiliacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmEnvioEstadoCtaAfiliacion
'** Descripción : Formulario para afiliar el envio de estado de cuenta TI-ERS036-2017
'** Creación : APRI, 20180520 09:00:00 AM
'**********************************************************************************************
Dim frsCredVig As ADODB.Recordset
Dim frsCuentasAh As ADODB.Recordset
Private ContMatCta As Integer
Private MatCta(100, 3) As String
Private fnTipo As Integer
Public Function Inicia(ByVal nTipo As String) As Boolean
    fnTipo = nTipo
    If fnTipo = 2 Then
        Me.Caption = "Desafiliación Envio de Estado de Cuenta"
        Frame1.Caption = "Solicitud de Desafiliación"
    End If
    Me.Show 1
End Function
Private Sub cmbBuscar_Click()

    Dim ClsPersona As New COMDPersona.DCOMPersonas
    Dim oPersona As COMDPersona.UCOMPersona
    Dim R As New ADODB.Recordset
    Dim lsPersCod As String
    Dim oProductos As COMNCaptaGenerales.NCOMCaptaGenerales
    Set ClsPersona = New COMDPersona.DCOMPersonas
    
    Set oPersona = frmBuscaPersona.Inicio
    
    If Not oPersona Is Nothing Then
        lsPersCod = oPersona.sPersCod
        TxtBuscarPersona.Text = lsPersCod
    End If
    If lsPersCod <> "" Then
        Set ClsPersona = New COMDPersona.DCOMPersonas
        Set R = ClsPersona.BuscaCliente(lsPersCod, BusquedaCodigo)
        
            lblDOI.Caption = IIf(R!nPersPersoneria = 1, R!cPersIDnroDNI, R!cPersIDnroRUC)
            lblNombre.Caption = R!cPersNombre
            txtCorreo.Text = R!cPersEmail
            txtDireccion.Text = R!cPersDireccDomicilio
            
            Frame3.Enabled = True
            If fnTipo = 1 Then
                Frame4.Enabled = True
                txtCorreo.Enabled = True
                txtDireccion.Enabled = True
                txtCorreo.SetFocus
            End If

            cmdGrabar.Enabled = True
            cmbBuscar.Enabled = False
        
        Set oProductos = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set frsCredVig = oProductos.CargarProductosVigentes(lsPersCod, 1, fnTipo)
        Set frsCuentasAh = oProductos.CargarProductosVigentes(lsPersCod, 2, fnTipo)
        Set oProductos = Nothing
        
        OptCreditos_Click
      
    Else
        Call LimpiaFormulario
    End If

End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaFormulario
End Sub

Private Sub LimpiaFormulario()
    TxtBuscarPersona.Text = ""
    lblDOI.Caption = ""
    lblNombre.Caption = ""
    txtCorreo.Text = ""
    txtDireccion.Text = ""
    If fnTipo = 1 Then
        txtCorreo.Enabled = True
        txtDireccion.Enabled = True
        FEProductos.EncabezadosAnchos = "350-400-2000-2500-0"
    ElseIf fnTipo = 2 Then
        txtCorreo.Enabled = False
        txtDireccion.Enabled = False
        FEProductos.EncabezadosAnchos = "350-400-2000-2500-1200"
    End If
    cmdGrabar.Enabled = False
    cmbBuscar.Enabled = True
    OptCreditos.value = True
    OptElectronico.value = True
    LimpiaFlex FEProductos
    Set frsCredVig = Nothing
    Set frsCuentasAh = Nothing
    Frame3.Enabled = False
    Frame4.Enabled = False
    For i = 0 To 99
        MatCta(i, 0) = ""
        MatCta(i, 1) = ""
        MatCta(i, 2) = ""
    Next i
   
    ContMatCta = 0
    
End Sub

Private Sub cmdGrabar_Click()
Dim cNroItem As String 'APRI20190615 SATI RFC1902040001
    If validar Then
        If fnTipo = 1 Then
            If MsgBox("Se procedera a realizar la afiliación, ¿Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Else
            If MsgBox("Se procedera a realizar la desafiliación, ¿Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        End If
            Dim oCaptaGeneral As New COMNCaptaGenerales.NCOMCaptaGenerales
            Dim clsprevio As previo.clsprevio
            Dim lsCadImp As String
            oCaptaGeneral.IniciaImpresora gImpresora
            Set clsprevio = New previo.clsprevio
            Set oCaptaGeneral = New COMNCaptaGenerales.NCOMCaptaGenerales
            cNroItem = oCaptaGeneral.AfiliarDesafiliarEnvioEstadoCta(fnTipo, MatCta, ContMatCta, TxtBuscarPersona.Text, IIf(OptElectronico.value, txtCorreo.Text, txtDireccion.Text), IIf(OptElectronico.value, 1, 2), GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), IIf(gsCodCargo <> "007014", 1, 2))
            
            If gsCodCargo <> "007014" Then 'APRI20190615 SATI RFC1902040001
                lsCadImp = oCaptaGeneral.GeneraSolicitudAfiliacionDesafiliacionEnvioEstadoCta(fnTipo, lblNombre.Caption, lblDOI.Caption, gdFecSis, IIf(OptCreditos.value, "Creditos", "Ahorros"), MatCta, ContMatCta, IIf(OptElectronico.value, "Electronico", "Fisico"), Trim(txtCorreo.Text), Trim(txtDireccion.Text), gsCodAge, gsNomCmac, gsNomAge, gsCodUser, cNroItem)
            
                clsprevio.Show lsCadImp, "Solicitud de Apertura", True, , gImpresora
            Else
                If fnTipo = 1 Then
                    MsgBox "El Codigo de afiliación es: " & cNroItem, vbInformation, "Aviso"
                Else
                    MsgBox "El Codigo de desafiliación es: " & cNroItem, vbInformation, "Aviso"
                End If
            End If
            'clsprevio.PrintSpool sLpt, oImpresora.gPrnTpoLetraSansSerif1PDef & oImpresora.gPrnTamLetra10CPIDef & sImpreDocs, False, gnLinPage
            Set oCaptaGeneral = Nothing
            Set clsprevio = Nothing

           Call cmdCancelar_Click
      
            
    End If
End Sub
Private Function validar() As Boolean
    validar = True
    If OptElectronico.value And txtCorreo.Text = "" And fnTipo = 1 Then
        MsgBox "Debe ingresar el correo del cliente.", vbInformation, "Aviso"
        txtCorreo.SetFocus
        validar = False
    ElseIf OptFisico.value And txtDireccion.Text = "" And fnTipo = 1 Then
        MsgBox "Debe ingresar la dirección del cliente.", vbInformation, "Aviso"
        txtCorreo.SetFocus
        validar = False
    ElseIf ContMatCta <= 0 Then
        MsgBox "Debe seleccionar al menos una cuenta.", vbInformation, "Aviso"
        validar = False
    End If
End Function
Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_Load()
    LimpiaFormulario
End Sub

Private Sub OptAhorros_Click()
LimpiaFlex FEProductos
ContMatCta = 0
For i = 0 To 99
        MatCta(i, 0) = ""
        MatCta(i, 1) = ""
        MatCta(i, 2) = ""
Next i
CargarCuentas (2)
End Sub

Private Sub OptCreditos_Click()
LimpiaFlex FEProductos
ContMatCta = 0
For i = 0 To 99
    MatCta(i, 0) = ""
    MatCta(i, 1) = ""
    MatCta(i, 2) = ""
Next i
CargarCuentas (1)
End Sub
Private Sub CargarCuentas(ByVal pnTipo As Integer)
 
 If pnTipo = 1 Then
        Frame5.Caption = "Créditos"
        
        'If Not frsCredVig Is Nothing Then Exit Sub
        
        If frsCredVig.RecordCount <= 0 Then
            FEProductos.Enabled = False
            If fnTipo = 1 Then
                MsgBox "El Cliente no posee cuentas de créditos para afiliar", vbInformation + vbAceptar, "Aviso"
            Else
                MsgBox "El Cliente no posee cuentas de créditos para desafiliar", vbInformation + vbAceptar, "Aviso"
            End If
            Exit Sub
        Else
            FEProductos.Enabled = True
            If frsCredVig.RecordCount > 0 Then frsCredVig.MoveFirst
            Do While Not frsCredVig.EOF
             
                FEProductos.AdicionaFila , , True
                FEProductos.TextMatrix(frsCredVig.Bookmark, 0) = frsCredVig.Bookmark
                FEProductos.TextMatrix(frsCredVig.Bookmark, 2) = frsCredVig!cCtaCod
                FEProductos.TextMatrix(frsCredVig.Bookmark, 3) = frsCredVig!cProducto
                FEProductos.TextMatrix(frsCredVig.Bookmark, 4) = frsCredVig!cMedioEnvio
                
                frsCredVig.MoveNext
            Loop
            
            FEProductos.Rows = frsCredVig.RecordCount + 1
            'FEProductos.row = 1
        End If

Else
        Frame5.Caption = "Ahorros"
         'If Not frsCuentasAh Is Nothing Then Exit Sub
        If frsCuentasAh.RecordCount <= 0 Then
            FEProductos.Enabled = False
            If fnTipo = 1 Then
                MsgBox "El Cliente no posee cuentas de ahorros para afiliar", vbInformation + vbAceptar, "Aviso"
            Else
                MsgBox "El Cliente no posee cuentas de ahorros para desafiliar", vbInformation + vbAceptar, "Aviso"
            End If
            Exit Sub
        Else
            FEProductos.Enabled = True
            If frsCuentasAh.RecordCount > 0 Then frsCuentasAh.MoveFirst
            Do While Not frsCuentasAh.EOF
    
                FEProductos.AdicionaFila , , True
                FEProductos.TextMatrix(frsCuentasAh.Bookmark, 0) = frsCuentasAh.Bookmark
                FEProductos.TextMatrix(frsCuentasAh.Bookmark, 2) = frsCuentasAh!cCtaCod
                FEProductos.TextMatrix(frsCuentasAh.Bookmark, 3) = frsCuentasAh!SubProducto
                FEProductos.TextMatrix(frsCuentasAh.Bookmark, 4) = frsCuentasAh!cMedioEnvio
                
                frsCuentasAh.MoveNext
            Loop
            
            FEProductos.Rows = frsCuentasAh.RecordCount + 1
            'FEProductos.row = 1
        End If
        
End If
    FEProductos.row = 1
    FEProductos.lbEditarFlex = True
End Sub
Private Sub FEProductos_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If Trim(FEProductos.TextMatrix(pnRow, pnCol)) <> "." Then 'Sin Check
        Call EliminarCta(FEProductos.TextMatrix(pnRow, 2))
    Else 'Con Check
        Call AdicionaCta(FEProductos.TextMatrix(pnRow, 2), FEProductos.TextMatrix(pnRow, 3), FEProductos.TextMatrix(pnRow, 4))
    End If
End Sub
Private Sub EliminarCta(ByVal psCtaCod As String)
Dim i As Integer
Dim nPos As Integer

    
    nPos = -1
    For i = 0 To ContMatCta - 1
        If MatCta(i, 0) = psCtaCod Then
            nPos = i
            Exit For
        End If
    Next i
    If nPos <> -1 Then
        For i = nPos To ContMatCta - 2
            MatCta(i, 0) = MatCta(i + 1, 0)
            MatCta(i, 1) = MatCta(i + 1, 1)
            MatCta(i, 2) = MatCta(i + 1, 2)
        Next i
        MatCta(ContMatCta - 1, 0) = ""
        MatCta(ContMatCta - 1, 1) = ""
        MatCta(ContMatCta - 1, 2) = ""
        ContMatCta = ContMatCta - 1
    End If
        
End Sub
Private Sub AdicionaCta(ByVal psCtaCod As String, ByVal psTpoCred As String, ByVal psMedioEnvio As String)
    ContMatCta = ContMatCta + 1
    MatCta(ContMatCta - 1, 0) = psCtaCod
    MatCta(ContMatCta - 1, 1) = psTpoCred
    MatCta(ContMatCta - 1, 2) = psMedioEnvio
End Sub
