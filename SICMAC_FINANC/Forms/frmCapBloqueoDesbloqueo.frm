VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCapBloqueoDesbloqueo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8715
   Icon            =   "frmCapBloqueoDesbloqueo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7185
   ScaleWidth      =   8715
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   1155
      TabIndex        =   8
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6615
      TabIndex        =   5
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7560
      TabIndex        =   6
      Top             =   6720
      Width           =   855
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   105
      TabIndex        =   7
      Top             =   6720
      Width           =   855
   End
   Begin VB.Frame fraBloqueo 
      Caption         =   "Bloqueo"
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
      Height          =   3690
      Left            =   105
      TabIndex        =   10
      Top             =   2940
      Width           =   8520
      Begin TabDlg.SSTab tabBloqueo 
         Height          =   3375
         Left            =   105
         TabIndex        =   4
         Top             =   210
         Width           =   8310
         _ExtentX        =   14658
         _ExtentY        =   5953
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TabCaption(0)   =   "Bloqueo Retiro"
         TabPicture(0)   =   "frmCapBloqueoDesbloqueo.frx":030A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "grdRetiro"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Bloqueo Total"
         TabPicture(1)   =   "frmCapBloqueoDesbloqueo.frx":0326
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "grdTotal"
         Tab(1).ControlCount=   1
         Begin Sicmact2000.FlexEdit grdRetiro 
            Height          =   2850
            Left            =   105
            TabIndex        =   3
            Top             =   420
            Width           =   8100
            _ExtentX        =   14288
            _ExtentY        =   5027
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Est-Motivo-Comentario-Fecha-Usu-cConsValor"
            EncabezadosAnchos=   "250-400-3000-6000-850-500-0"
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-3-X-X-X"
            ListaControles  =   "0-4-0-0-0-0-0"
            EncabezadosAlineacion=   "C-C-L-L-C-C-C"
            FormatosEdit    =   "0-0-1-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
         End
         Begin Sicmact2000.FlexEdit grdTotal 
            Height          =   2850
            Left            =   -74895
            TabIndex        =   17
            Top             =   420
            Width           =   8100
            _ExtentX        =   14288
            _ExtentY        =   5027
            Cols0           =   7
            HighLight       =   1
            AllowUserResizing=   3
            EncabezadosNombres=   "#-Est-Motivo-Comentario-Fecha-Usu-cConsValor"
            EncabezadosAnchos=   "250-400-3000-6000-850-500-0"
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-1-X-3-X-X-X"
            ListaControles  =   "0-4-0-0-0-0-0"
            EncabezadosAlineacion=   "C-C-L-L-C-C-C"
            FormatosEdit    =   "0-0-1-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
         End
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta"
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
      Height          =   2745
      Left            =   105
      TabIndex        =   9
      Top             =   105
      Width           =   8520
      Begin VB.Frame fraDatosCuenta 
         Height          =   2010
         Left            =   105
         TabIndex        =   11
         Top             =   630
         Width           =   8310
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCliente 
            Height          =   1275
            Left            =   105
            TabIndex        =   2
            Top             =   630
            Width           =   8100
            _ExtentX        =   14288
            _ExtentY        =   2249
            _Version        =   393216
            _NumberOfBands  =   1
            _Band(0).Cols   =   2
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Estado:"
            Height          =   195
            Left            =   4935
            TabIndex        =   18
            Top             =   285
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Apertura :"
            Height          =   195
            Left            =   105
            TabIndex        =   16
            Top             =   278
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Cuenta :"
            Height          =   195
            Left            =   2100
            TabIndex        =   15
            Top             =   285
            Width           =   960
         End
         Begin VB.Label lblEstado 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   330
            Left            =   5565
            TabIndex        =   14
            Top             =   210
            Width           =   2640
         End
         Begin VB.Label lblApertura 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   840
            TabIndex        =   13
            Top             =   210
            Width           =   1170
         End
         Begin VB.Label lblTipoCuenta 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   330
            Left            =   3150
            TabIndex        =   12
            Top             =   210
            Width           =   1695
         End
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         Height          =   375
         Left            =   3780
         TabIndex        =   1
         Top             =   250
         Width           =   435
      End
      Begin Sicmact2000.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   105
         TabIndex        =   0
         Top             =   250
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   767
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblCuenta 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   435
         Left            =   4410
         TabIndex        =   19
         Top             =   210
         Width           =   3690
      End
   End
End
Attribute VB_Name = "frmCapBloqueoDesbloqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As Producto
Public bConsulta As Boolean

Private Sub SetupGridCliente()
Dim I As Integer
grdCliente.Cols = 12
For I = 1 To grdCliente.Rows - 1
    grdCliente.MergeCol(I) = True
Next I
grdCliente.MergeCells = flexMergeFree
grdCliente.ColWidth(0) = 100
grdCliente.ColWidth(1) = 3500
grdCliente.ColWidth(2) = 400
grdCliente.ColWidth(3) = 3500
grdCliente.ColWidth(4) = 1500
grdCliente.ColWidth(5) = 1000
grdCliente.ColWidth(6) = 600
grdCliente.ColWidth(7) = 1500
grdCliente.ColWidth(8) = 0
grdCliente.ColWidth(9) = 0
grdCliente.ColWidth(10) = 0
grdCliente.ColWidth(11) = 0

grdCliente.TextMatrix(0, 1) = "Nombre"
grdCliente.TextMatrix(0, 2) = "RE"
grdCliente.TextMatrix(0, 3) = "Direccion"
grdCliente.TextMatrix(0, 4) = "Zona"
grdCliente.TextMatrix(0, 5) = "Fono"
grdCliente.TextMatrix(0, 6) = "ID"
grdCliente.TextMatrix(0, 7) = "ID N°"
End Sub

Public Sub Inicia(ByVal nProd As Producto, Optional bCons As Boolean = False)
nProducto = nProd
bConsulta = bCons
Select Case nProd
    Case gCapAhorros
        txtCuenta.Prod = Trim(Str(gCapAhorros))
        Me.Caption = "Captaciones - Bloqueo/Desbloqueo - Ahorros"
    Case gCapPlazoFijo
        txtCuenta.Prod = Trim(Str(gCapPlazoFijo))
        Me.Caption = "Captaciones - Bloqueo/Desbloqueo - Plazo Fijo"
    Case gCapCTS
        txtCuenta.Prod = Trim(Str(gCapCTS))
        Me.Caption = "Captaciones - Mantenimiento - CTS"
End Select

If bConsulta Then
    cmdGrabar.Visible = False
    grdRetiro.lbEditarFlex = False
    grdTotal.lbEditarFlex = False
Else
    cmdGrabar.Visible = True
    grdRetiro.lbEditarFlex = True
    grdTotal.lbEditarFlex = True
End If
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Age = Mid(gsCodAge, 4, 2)
txtCuenta.EnabledCMAC = False
txtCuenta.EnabledAge = False
txtCuenta.EnabledProd = False
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
fraBloqueo.Enabled = False
fraDatosCuenta.Enabled = False
SetupGridCliente
Me.Show 1
End Sub

Private Sub ObtieneDatosCuenta(ByVal sCuenta As String)
Dim clsMant As NCapMantenimiento
Dim rsCta As Recordset, rsRel As Recordset
Dim nEstado As CaptacEstado
Dim sSQL As String, sMoneda As String

Set clsMant = New NCapMantenimiento
Set rsCta = New Recordset
Set rsCta = clsMant.GetDatosCuenta(sCuenta)

If Not (rsCta.EOF And rsCta.BOF) Then
    nEstado = rsCta("cPrdEstado")
    If (nEstado <> gCapEstAnuladaAct) Or (nEstado <> gCapEstAnuladaInac) Or (nEstado <> gCapEstCanceladaAct) Or (nEstado <> gCapEstCanceladaInac) Then
        lblApertura = Format$(rsCta("dApertura"), "dd mmm yyyy")
        lblEstado = UCase(rsCta("cEstado"))
        lblTipoCuenta = UCase(rsCta("cTipoCuenta"))
        If Mid(rsCta("cCtaCod"), 9, 1) = "1" Then
            sMoneda = "NACIONAL"
            lblCuenta.ForeColor = &HC00000
        Else
            sMoneda = "EXTRANJERA"
            lblCuenta.ForeColor = &H8000&
        End If
        Select Case nProducto
            Case gCapAhorros
                lblCuenta = "AHORROS " & IIf(rsCta("bOrdPag"), "CON ORDEN DE PAGO", "SIN ORDEN DE PAGO") & " - MONEDA " & sMoneda
            Case gCapPlazoFijo
                lblCuenta = "PLAZO FIJO - MONEDA " & sMoneda
            Case gCapCTS
                lblCuenta = "CTS - MONEDA " & sMoneda
        End Select
        rsCta.Close
        Set rsCta = clsMant.GetProductoPersona(sCuenta)
        If Not (rsCta.EOF And rsCta.BOF) Then
            Set grdCliente.Recordset = rsCta
            cmdBuscar.Enabled = False
            txtCuenta.Enabled = False
        Else
            MsgBox "Cuenta no posee relacion con Persona", vbExclamation, "Aviso"
            txtCuenta.SetFocusCuenta
        End If
        rsCta.Close
        Dim sCta As String
        sCta = txtCuenta.NroCuenta
        'Obtiene los datos del bloqueo
        Set rsCta = clsMant.GetCapBloqueos(sCta, gCapTpoBlqRetiro, gCaptacMotBloqueoRet)
        If Not (rsCta.EOF And rsCta.BOF) Then
            Set grdRetiro.Recordset = rsCta
        End If
        Set rsCta = clsMant.GetCapBloqueos(sCta, gCapTpoBlqTotal, gCaptacMotBloqueoTot)
        If Not (rsCta.EOF And rsCta.BOF) Then
            Set grdTotal.Recordset = rsCta
        End If
        Set rsCta = Nothing
        SetupGridCliente
        cmdCancelar.Enabled = True
        cmdImprimir.Enabled = True
        fraBloqueo.Enabled = True
        fraDatosCuenta.Enabled = True
        cmdGrabar.Enabled = True
        grdRetiro.SetFocus
    Else
        MsgBox "La cuenta se encuentra Cancelada o Anulada.", vbInformation, "Aviso"
        txtCuenta.SetFocusCuenta
        Exit Sub
    End If
Else
    MsgBox "Cuenta no existe", vbInformation, "Aviso"
    txtCuenta.SetFocusCuenta
End If
End Sub

Private Sub cmdBuscar_Click()
Dim clsPers As UPersona
Set clsPers = New UPersona
Set clsPers = frmBuscaPersona.Inicio
If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As Recordset
    Dim clsCap As NCapMantenimiento
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuentas
    
    sPers = clsPers.sPersCod
    Set clsCap = New NCapMantenimiento
    
    Set rsPers = clsCap.GetCuentasPersona(sPers, nProducto)
    Set clsCap = Nothing
    If Not (rsPers.EOF And rsPers.EOF) Then
        Do While Not rsPers.EOF
            sCta = rsPers("cCtaCod")
            sRelac = rsPers("cRelacion")
            sEstado = Trim(rsPers("cEstado"))
            frmCapMantenimientoCtas.lstCuentas.AddItem sCta & Space(2) & sRelac & Space(2) & sEstado
            rsPers.MoveNext
        Loop
        Set clsCuenta = frmCapMantenimientoCtas.Inicia
        If clsCuenta.sCtaCod <> "" Then
            txtCuenta.Cuenta = Mid(clsCuenta.sCtaCod, 9, 10)
        End If
        Set clsCuenta = Nothing
    Else
        MsgBox "Persona no posee ninguna cuenta de captaciones.", vbInformation, "Aviso"
    End If
    rsPers.Close
    Set rsPers = Nothing
End If
txtCuenta.SetFocusCuenta
End Sub

Private Sub cmdCancelar_Click()
txtCuenta.Enabled = True
cmdBuscar.Enabled = True
cmdImprimir.Enabled = False
grdCliente.Clear
grdCliente.Rows = 2
SetupGridCliente
grdRetiro.Clear
grdTotal.Clear
grdRetiro.FormaCabecera
grdRetiro.Rows = 2
grdTotal.FormaCabecera
grdTotal.Rows = 2
lblApertura = ""
lblCuenta = ""
lblTipoCuenta = ""
lblEstado = ""
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
txtCuenta.Cuenta = ""
txtCuenta.SetFocusCuenta
End Sub

Private Sub cmdGrabar_Click()
Dim rsRet As Recordset, rsTot As Recordset
Dim clsMant As NCapMantenimiento
Dim clsMov As NContFunciones
Dim sMovNro As String
If MsgBox("Está seguro de grabar??", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Set rsRet = grdRetiro.GetRsNew
    Set rsTot = grdTotal.GetRsNew
    
    Set clsMov = New NContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Set clsMov = Nothing
    
    Set clsMant = New NCapMantenimiento
    clsMant.ActualizaBloqueos txtCuenta.NroCuenta, rsRet, rsTot, sMovNro
    Set clsMant = Nothing
    cmdCancelar_Click
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub grdRetiro_OnCellChange(pnRow As Long, pnCol As Long)
grdRetiro.TextMatrix(pnRow, 1) = 1
grdRetiro.TextMatrix(pnRow, 4) = Format$(gdFecSis, "dd-mm-yyyy")
End Sub

Private Sub grdRetiro_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If grdRetiro.TextMatrix(pnRow, 1) = "" Then
    grdRetiro.TextMatrix(pnRow, 4) = ""
Else
    grdRetiro.TextMatrix(pnRow, 4) = Format$(gdFecSis, "dd-mm-yyyy")
End If
End Sub

Private Sub grdTotal_OnCellChange(pnRow As Long, pnCol As Long)
grdTotal.TextMatrix(pnRow, 1) = 1
grdTotal.TextMatrix(pnRow, 4) = Format$(gdFecSis, "dd-mm-yyyy")
End Sub

Private Sub grdTotal_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
If grdTotal.TextMatrix(pnRow, 1) = "" Then
    grdTotal.TextMatrix(pnRow, 4) = ""
Else
    grdTotal.TextMatrix(pnRow, 4) = Format$(gdFecSis, "dd-mm-yyyy")
End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCta
End If
End Sub
