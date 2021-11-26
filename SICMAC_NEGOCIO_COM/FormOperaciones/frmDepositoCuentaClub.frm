VERSION 5.00
Begin VB.Form frmDepositoCuentaClub 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7845
   Icon            =   "frmDepositoCuentaClub.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   7845
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   3885
      TabIndex        =   8
      Top             =   6825
      Width           =   1170
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5250
      TabIndex        =   9
      Top             =   6825
      Width           =   1170
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   330
      Left            =   6615
      TabIndex        =   10
      Top             =   6825
      Width           =   1065
   End
   Begin VB.Frame fraDetalle 
      Caption         =   "Detalle:"
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
      Height          =   2640
      Left            =   105
      TabIndex        =   12
      Top             =   3990
      Width           =   7590
      Begin SICMACT.EditMoney txtMonto 
         Height          =   330
         Left            =   6195
         TabIndex        =   5
         Top             =   1260
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
      End
      Begin VB.TextBox txtGlosa 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1890
         TabIndex        =   4
         Top             =   735
         Width           =   5475
      End
      Begin VB.TextBox txtReferencia 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1890
         TabIndex        =   3
         Top             =   315
         Width           =   1590
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6195
         TabIndex        =   7
         Top             =   2100
         Width           =   1170
      End
      Begin VB.Label lblItf 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   6195
         TabIndex        =   6
         Top             =   1680
         Width           =   1170
      End
      Begin VB.Label Label6 
         Caption         =   "Total:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   5145
         TabIndex        =   18
         Top             =   2145
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "ITF."
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
         Left            =   5145
         TabIndex        =   17
         Top             =   1755
         Width           =   855
      End
      Begin VB.Label lblMonto 
         Caption         =   "Monto S/."
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
         Left            =   5145
         TabIndex        =   16
         Top             =   1335
         Width           =   1065
      End
      Begin VB.Label Label3 
         Caption         =   "Nota: Esta operación no reducirá el monto de la Caja del RF, abonará dicho monto a la cuenta del club de trabajadores selecionado"
         Height          =   855
         Left            =   210
         TabIndex        =   15
         Top             =   1365
         Width           =   4635
      End
      Begin VB.Label Label2 
         Caption         =   "Glosa:"
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
         Left            =   210
         TabIndex        =   14
         Top             =   780
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Doc. Referencia:"
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
         Left            =   210
         TabIndex        =   13
         Top             =   315
         Width           =   1590
      End
   End
   Begin VB.Frame fraCuenta 
      Caption         =   "Cuenta del Club"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   3795
      Left            =   105
      TabIndex        =   1
      Top             =   105
      Width           =   7590
      Begin VB.Frame fraClientes 
         Caption         =   "Clientes"
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
         Height          =   2745
         Left            =   105
         TabIndex        =   11
         Top             =   840
         Width           =   7365
         Begin SICMACT.FlexEdit grdClientes 
            Height          =   2430
            Left            =   105
            TabIndex        =   2
            Top             =   210
            Width           =   7155
            _ExtentX        =   12621
            _ExtentY        =   4286
            Cols0           =   6
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "#-Codigo-Nombre-Relacion-Personeria-Obligatorio"
            EncabezadosAnchos=   "500-1500-3500-1200-0-0"
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
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnasAEditar =   "X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-L-L-L-L-L"
            FormatosEdit    =   "0-0-0-0-0-0"
            TextArray0      =   "#"
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   495
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   435
         Left            =   105
         TabIndex        =   0
         Top             =   315
         Width           =   3690
         _ExtentX        =   6509
         _ExtentY        =   767
         Texto           =   "Cuenta:"
         EnabledCta      =   -1  'True
         EnabledAge      =   -1  'True
         Prod            =   "232"
         CMAC            =   "109"
      End
   End
End
Attribute VB_Name = "frmDepositoCuentaClub"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************************************************
'* NOMBRE         : "frmDepositoCuentaClub"
'* DESCRIPCION    : Formulario creado realizar depósitos directos a la cuenta de los trabajadores, segun ERS: TI-ERS145-2013
'* CREACION       : RIRO, 20131017 10:00 AM
'************************************************************************************************************************************************

Option Explicit

Private sCuenta As String
Private lbITFCtaExonerada As Boolean
Private sOpeCod As String
Public nProducto As COMDConstantes.Producto

Private Sub cmdCancelar_Click()
    grdClientes.Clear
    grdClientes.Rows = 2
    grdClientes.FormaCabecera
    
    txtReferencia.Text = Empty
    txtGlosa.Text = Empty
    txtMonto.value = 0
    
    txtReferencia.Enabled = False
    txtGlosa.Enabled = False
    txtMonto.Enabled = False
    
    cmdGrabar.Enabled = False
    cmdCancelar.Enabled = False
    
    lblItf.Caption = "0.00"
    lblTotal.Caption = "0.00"
    
    txtCuenta.EnabledAge = True
    txtCuenta.EnabledProd = False
    txtCuenta.EnabledCta = True
    
    txtCuenta.NroCuenta = ""
    txtCuenta.Prod = "232"
    txtCuenta.CMAC = "109"
    txtCuenta.SetFocus
    
End Sub

Private Sub CmdGrabar_Click()
    
    Dim sNroDoc As String
    Dim nMonto As Double
    Dim sCuenta As String
    Dim lsmensaje As String
    Dim nComixDep As Double
    
    Dim loLavDinero As frmMovLavDinero
    Dim objPersona As COMDPersona.DCOMPersonas
    Set objPersona = New COMDPersona.DCOMPersonas
    
    Set loLavDinero = New frmMovLavDinero
    
    Dim loMov As COMDMov.DCOMMov
    Set loMov = New COMDMov.DCOMMov
    Dim lnLogEcotaxi As Integer
    Dim oNCOMContImprimir As COMNContabilidad.NCOMContImprimir
    Set oNCOMContImprimir = New COMNContabilidad.NCOMContImprimir

On Error GoTo ErrGraba

    nMonto = CDbl(IIf(IsNumeric(lblTotal.Caption), lblTotal.Caption, 0#))
    If nMonto = 0 Then
        MsgBox "Monto debe ser mayor a cero", vbInformation, "Aviso"
        If txtMonto.Enabled Then txtMonto.SetFocus
        Exit Sub
    End If
    
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaMovimiento
    
    If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
       
        Dim sMovNro As String, sPersLavDinero As String
        Dim clsMov As COMNContabilidad.NCOMContFunciones
        Dim nSaldo As Double, nPorcDisp As Double
        Dim nMontoLavDinero As Double, nTC As Double
        Dim clsLav As COMNCaptaGenerales.NCOMCaptaDefinicion, clsExo As COMNCaptaServicios.NCOMCaptaServicios
        Dim previo As New previo.clsprevio
        Dim lsBoletaImp As String
        Dim lsBoletaImpITF As String
        Dim nFicSal As Integer
        Dim lsBoletaCVME As String
        
        sCuenta = txtCuenta.NroCuenta
        Set clsLav = New COMNCaptaGenerales.NCOMCaptaDefinicion
        Set clsExo = New COMNCaptaServicios.NCOMCaptaServicios
        
        If Not clsExo.EsCuentaExoneradaLavadoDinero(sCuenta) Then
            Set clsExo = Nothing
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            If nMonto >= Round(nMontoLavDinero * nTC, 2) Then
                sPersLavDinero = loLavDinero.inicia(, , , , False, True, nMonto, sCuenta, , True, , , , , , CInt(Mid(sCuenta, 9, 1)), , gnTipoREU, gnMontoAcumulado, gsOrigen)
                If loLavDinero.OrdPersLavDinero = "" Then
                    Exit Sub
                End If
            End If
        Else
            Set clsExo = Nothing
        End If
        Dim oCont As COMNContabilidad.NCOMContFunciones
        Set oCont = New COMNContabilidad.NCOMContFunciones
        Set clsCap = New COMNCaptaGenerales.NCOMCaptaMovimiento
        sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        Dim sDocReferencia As String
        sDocReferencia = Trim(txtReferencia.Text)
        
        nSaldo = clsCap.CapAbonoCuentaAho(sCuenta, nMonto, gAhoDepDirectoClub, sMovNro, Trim(txtGlosa.Text), , , , , , , _
                                        , , gsNomAge, sLpt, , , , _
                                        , , gsCodCMAC, , , IIf(IsNumeric(lblItf.Caption), CDbl(lblItf.Caption), CDbl("0")), _
                                        , gITFCobroCargo, , , , , lsBoletaImp, , , , , , , , , , , , , , , , , , , , , 0, , , sDocReferencia)
    
    If Trim(lsBoletaImp) <> "" Then ImprimeBoleta lsBoletaImp
        
        Set clsLav = Nothing
        Set clsCap = Nothing
                
        gVarPublicas.LimpiaVarLavDinero
        
        MsgBox "La Operacion se registró correctamente", vbInformation, "Aviso"
        cmdCancelar_Click

    End If
      
    Set loLavDinero = Nothing
    Set oNCOMContImprimir = Nothing
    Exit Sub
ErrGraba:
        MsgBox Err.Description, vbExclamation, "Error"
        Exit Sub
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    Me.Caption = "Depósitos a Cuenta del Club de Trabajadores"
    txtCuenta.EnabledAge = True
    sOpeCod = "200264"
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        
        If Left(txtCuenta.Cuenta, 1) <> "1" Then
            MsgBox "Solo se admite depósito a cuentas en soles", vbInformation, "Aviso"
            Exit Sub
        End If
        
        Dim objValidar As COMNCaptaGenerales.NCOMCaptaMovimiento
        Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
        Dim rsCta As ADODB.Recordset
        
        grdClientes.Clear
        grdClientes.Rows = 2
        grdClientes.FormaCabecera
        
        Set objValidar = New COMNCaptaGenerales.NCOMCaptaMovimiento
        Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
        Set rsCta = New ADODB.Recordset
        
        sCuenta = txtCuenta.NroCuenta
        Set rsCta = clsMant.GetDatosCuenta(sCuenta)
        If Not (rsCta.EOF And rsCta.BOF) Then
            If Not objValidar.ValidaEstadoCuenta(sCuenta, False) Then
                MsgBox "La cuenta NO Tiene un estado valido para la operacion", vbInformation + vbDefaultButton1, "Aviso"
                Set objValidar = Nothing
                Set clsMant = Nothing
                Set rsCta = Nothing
                Exit Sub
            End If
            
            
            Dim rsClientes As ADODB.Recordset
            Set rsClientes = New ADODB.Recordset
            Set rsClientes = clsMant.GetProductoPersona(sCuenta)
            grdClientes.AdicionaFila
            If Not (rsClientes.EOF And rsClientes.BOF) Then
                Dim i As Integer
                For i = 1 To rsClientes.RecordCount
                    grdClientes.TextMatrix(i, 1) = rsClientes("cperscod")
                    grdClientes.TextMatrix(i, 2) = rsClientes("nombre")
                    grdClientes.TextMatrix(i, 3) = rsClientes("crelacion")
                    rsClientes.MoveNext
                    If i < rsClientes.RecordCount Then
                        grdClientes.AdicionaFila
                    End If
                Next
                txtCuenta.EnabledAge = False
                txtCuenta.EnabledProd = False
                txtCuenta.EnabledCta = False
                txtReferencia.Enabled = True
                txtGlosa.Enabled = True
                txtMonto.Enabled = True
                
                cmdGrabar.Enabled = True
                cmdCancelar.Enabled = True
                
                txtReferencia.SetFocus
                Exit Sub
            Else
                MsgBox "Cuenta no posee relacion con Persona", vbExclamation, "Aviso"
                txtCuenta.SetFocusCuenta
                Exit Sub
            End If
        Else
            MsgBox "No existe cuenta", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
End Sub

Private Sub txtItf_Click()

End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtMonto.Enabled Then
            txtMonto.SetFocus
        End If
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub txtMonto_Change()

    Dim nITF As Double
    Dim nRedondeoITF As Double
    
    nITF = fgITFCalculaImpuesto(CDbl(txtMonto.value))
    nRedondeoITF = fgDiferenciaRedondeoITF((nITF))
    
    If nRedondeoITF > 0 Then
        nITF = nITF - nRedondeoITF
    End If
    
    If txtMonto.Text = "." Then
        txtMonto.value = 0
    End If
    
    lblItf.Caption = Format(nITF, "#,##0.00")
    lblTotal.Caption = Format(txtMonto.value, "#,##0.00")
    
End Sub

Private Sub txtMonto_GotFocus()
txtMonto.SelStart = 0
txtMonto.SelLength = 20
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdGrabar.Enabled Then
            cmdGrabar.SetFocus
        End If
    End If
End Sub

Private Sub txtReferencia_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If txtGlosa.Enabled Then
            txtGlosa.SetFocus
        End If
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    If nProducto = gCapCTS Then
        sBoleta = sBoleta & oImpresora.gPrnSaltoLinea
    End If
    Print #nFicSal, sBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    Print #nFicSal, ""
    Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub
