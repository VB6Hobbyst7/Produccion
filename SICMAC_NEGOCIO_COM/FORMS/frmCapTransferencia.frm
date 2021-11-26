VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmCapTransferencia 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "frmCapTransferencia.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   9600
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CdlgFile 
      Left            =   1425
      Top             =   7095
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   105
      TabIndex        =   11
      Top             =   7035
      Width           =   915
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8550
      TabIndex        =   10
      Top             =   7035
      Width           =   915
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7530
      TabIndex        =   9
      Top             =   7035
      Width           =   915
   End
   Begin VB.Frame fraCuentaAbono 
      Caption         =   "Cuenta Abono"
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
      Height          =   3855
      Left            =   105
      TabIndex        =   20
      Top             =   3075
      Width           =   9375
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   1140
         TabIndex        =   7
         Top             =   3360
         Width           =   915
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
         Height          =   375
         Left            =   180
         TabIndex        =   6
         Top             =   3360
         Width           =   915
      End
      Begin SICMACT.FlexEdit grdCuentaAbono 
         Height          =   2655
         Left            =   180
         TabIndex        =   5
         Top             =   300
         Width           =   9075
         _ExtentX        =   16007
         _ExtentY        =   4683
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Cuenta-Titular-Monto S/.-Monto $"
         EncabezadosAnchos=   "250-1800-3800-1400-1400"
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
         ColumnasAEditar =   "X-X-X-3-4"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-R-R"
         FormatosEdit    =   "0-0-0-2-2"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   285
         ForeColorFixed  =   -2147483630
      End
      Begin SICMACT.ActXCodCta txtCuentaAbo 
         Height          =   375
         Left            =   2220
         TabIndex        =   8
         Top             =   3360
         Visible         =   0   'False
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "TOTAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   3060
         TabIndex        =   24
         Top             =   2940
         Width           =   3075
      End
      Begin VB.Label lblTotalME 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   7500
         TabIndex        =   23
         Top             =   2940
         Width           =   1395
      End
      Begin VB.Label lblTotalMN 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   6120
         TabIndex        =   22
         Top             =   2940
         Width           =   1400
      End
   End
   Begin VB.Frame fraCuentaCargo 
      Caption         =   "Cuenta Cargo"
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
      Height          =   2970
      Left            =   105
      TabIndex        =   12
      Top             =   60
      Width           =   9375
      Begin VB.TextBox txtIdAut 
         Height          =   330
         Left            =   1560
         TabIndex        =   26
         Top             =   750
         Width           =   1380
      End
      Begin VB.CommandButton cmdObtDatos 
         Caption         =   "&Obtener Datos"
         Height          =   375
         Left            =   3825
         TabIndex        =   1
         Top             =   300
         Width           =   1230
      End
      Begin VB.Frame fraGlosa 
         Caption         =   "Glosa"
         Height          =   795
         Left            =   5580
         TabIndex        =   25
         Top             =   2010
         Width           =   3615
         Begin VB.TextBox txtGlosa 
            Height          =   435
            Left            =   120
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   240
            Width           =   3375
         End
      End
      Begin VB.Frame fraTipoCambio 
         Caption         =   "Tipo Cambio"
         Height          =   915
         Left            =   7680
         TabIndex        =   15
         Top             =   1110
         Width           =   1575
         Begin VB.Label lblTCV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   660
            TabIndex        =   19
            Top             =   510
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "TCV:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   570
            Width           =   360
         End
         Begin VB.Label lblTCC 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Left            =   660
            TabIndex        =   17
            Top             =   210
            Width           =   795
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "TCC:"
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   180
            TabIndex        =   16
            Top             =   270
            Width           =   360
         End
      End
      Begin VB.Frame fraMontoCargo 
         Caption         =   "Monto Total Cargo"
         Height          =   915
         Left            =   5580
         TabIndex        =   13
         Top             =   1110
         Width           =   1995
         Begin SICMACT.EditMoney txtMontoCargo 
            Height          =   375
            Left            =   120
            TabIndex        =   3
            Top             =   300
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   661
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   255
            Text            =   "0"
         End
         Begin VB.Label lblMon 
            Caption         =   "S/."
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
            Height          =   255
            Left            =   1620
            TabIndex        =   21
            Top             =   360
            Width           =   315
         End
      End
      Begin SICMACT.ActXCodCta txtCuenta 
         Height          =   375
         Left            =   180
         TabIndex        =   0
         Top             =   285
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   "Cuenta N°"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin SICMACT.FlexEdit grdCliente 
         Height          =   1635
         Left            =   180
         TabIndex        =   2
         Top             =   1170
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   2884
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Nombre-RE-cperscod-CCodRelacion"
         EncabezadosAnchos=   "250-4000-600-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Id Autorización"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   225
         TabIndex        =   27
         Top             =   810
         Width           =   1290
      End
      Begin VB.Label lblMensaje 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   4065
         TabIndex        =   14
         Top             =   240
         Width           =   5235
      End
   End
End
Attribute VB_Name = "frmCapTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As Producto
Dim nMoneda As Moneda
Dim nOperacion As CaptacOperacion
Private Const nLongPrimerRegistro = 66
Private Const nLongSegundoRegistro = 33
Dim sError As String, sCodPers As String



'***************Variabres Agregadas********************
Dim Gtitular As String
Dim GAutNivel As String
Dim GAutMontoFinSol As Double
Dim GAutMontoFinDol As Double
Dim GMontoAprobado As Double, GNroID As Long, GPersCod As String

'********************************************************



Private Sub ObtieneDatosCuentasAbonar(ByVal sArchivo As String)
Dim sCad As String
Dim bPrimeraLinea As Boolean
Dim nMontoTotal As Double, nSumaTotal As Double
Dim nNumReg As Long, nItem As Long
Dim dFechaAbono As Date, dFechaProceso As Date
Dim sCuentaAbono As String, sCuentaCargo As String
Dim nMontoAbono As Double
On Error GoTo ErrFileOpen
Open sArchivo For Input As #1
bPrimeraLinea = True
nItem = 0
nSumaTotal = 0
sError = ""
Do While Not EOF(1)
    Line Input #1, sCad
    If sCad <> "" Then
        If bPrimeraLinea Then
            If Len(sCad) = nLongPrimerRegistro Then
                sCodPers = Left(sCad, 13)
                sCad = Mid(sCad, 14, Len(sCad) - 13)
                sCuentaCargo = Left(sCad, 18)
                sCad = Mid(sCad, 19, Len(sCad) - 18)
                If ObtieneDatosCuenta(sCuentaCargo, True) Then
                    nNumReg = CLng(Trim(Left(sCad, 8)))
                    sCad = Mid(sCad, 9, Len(sCad) - 8)
                    nMontoTotal = CDbl(Trim(Mid(sCad, 1, 9)) & "." & Trim(Mid(sCad, 10, 2)))
                    sCad = Mid(sCad, 12, Len(sCad) - 11)
                    dFechaAbono = CDate(Mid(sCad, 7, 2) & "/" & Mid(sCad, 5, 2) & "/" & Mid(sCad, 1, 4))
                    sCad = Mid(sCad, 9, Len(sCad) - 8)
                    If DateDiff("d", gdFecSis, dFechaAbono) >= 0 Then
                        dFechaProceso = CDate(Mid(sCad, 7, 2) & "/" & Mid(sCad, 5, 2) & "/" & Mid(sCad, 1, 4))
                        If DateDiff("d", gdFecSis, dFechaProceso) < 0 Then
                            sError = sError & "Fecha de Proceso es mayor que la fecha actual" & gPrnSaltoLinea
                        End If
                    Else
                        sError = sError & "Fecha de Abono es menor que la fecha actual" & gPrnSaltoLinea
                    End If
                    txtCuenta.Age = sCuentaCargo
                End If
                bPrimeraLinea = False
            Else
                sError = sError & "Longitud del primer registro no coincide con formato establecido" & gPrnSaltoLinea
            End If
        Else
            sCad = Mid(sCad, 5, Len(sCad) - 4)
            sCuentaAbono = Left(sCad, 18)
            sCad = Mid(sCad, 19, Len(sCad) - 18)
            nMontoAbono = CDbl(Trim(Mid(sCad, 1, 9)) & "." & Trim(Mid(sCad, 10, 2)))
            If Not CuentaExisteEnLista(sCuentaAbono) Then
                ObtieneDatosCuentaAbono sCuentaAbono, True, nMontoAbono
            Else
                sError = sError & "Cuenta N° " & sCuentaAbono & "Duplicada en la relación" & gPrnSaltoLinea
            End If
            nSumaTotal = nSumaTotal + nMontoAbono
            nItem = nItem + 1
        End If
    End If
Loop
Close #1
If nItem <> nNumReg Then
    sError = sError & "Número de Cuentas NO coincide con el total de registros enviados. " & nNumReg & " - " & nItem & gPrnSaltoLinea
End If
If Round(nMontoTotal, 2) - Round(nSumaTotal, 2) <> 0 Then
    sError = sError & "Monto Total NO coincide con la SUMA TOTAL de MONTOS A ABONAR. " & nMontoTotal & " - " & nSumaTotal & gPrnSaltoLinea
End If
If sError <> "" Then
    Dim oPrevio As previo.clsPrevio
    Set oPrevio = New previo.clsPrevio
    oPrevio.Show sError, "Errores Cargo Abono en Lote", True, , gImpresora
    Set oPrevio = Nothing
    cmdCancelar_Click
    Exit Sub
End If
txtMontoCargo.value = nMontoTotal
CalculaTotales
cmdGrabar.Enabled = True
cmdCancelar.Enabled = True
fraCuentaAbono.Enabled = True
fraMontoCargo.Enabled = True
txtCuenta.Enabled = False
cmdObtDatos.Enabled = False
grdCuentaAbono.lbEditarFlex = False
Exit Sub
ErrFileOpen:
    Close #1
    cmdCancelar_Click
    MsgBox Err.Description, vbExclamation, "Error"
End Sub

Private Function CuentaExisteEnLista(ByVal sCuenta As String) As Boolean
Dim bExito As Boolean
Dim I As Long
Dim sCuentaLista As String
bExito = False
For I = 1 To grdCuentaAbono.Rows - 1
    sCuentaLista = grdCuentaAbono.TextMatrix(I, 1)
    If sCuenta = sCuentaLista Then
        bExito = True
        Exit For
    End If
Next I
CuentaExisteEnLista = bExito
End Function

Private Sub CalculaTotales()
Dim I As Long, nFila As Long, nCol As Long
Dim nAcumMN As Double, nAcumME As Double, nMonto As Double
Dim bValida As Boolean
nFila = grdCuentaAbono.Row
nCol = grdCuentaAbono.Col
nAcumMN = 0
nAcumME = 0
For I = 1 To grdCuentaAbono.Rows - 1
    If grdCuentaAbono.TextMatrix(I, 3) <> "" Then
        nAcumMN = nAcumMN + CDbl(grdCuentaAbono.TextMatrix(I, 3))
    End If
    If grdCuentaAbono.TextMatrix(I, 4) <> "" Then
        nAcumME = nAcumME + CDbl(grdCuentaAbono.TextMatrix(I, 4))
    End If
Next I
nMonto = txtMontoCargo.value
bValida = True
If nMoneda = gMonedaNacional Then
    If nMonto < nAcumMN Then
        MsgBox "SUMA TOTAL supera al monto establecido para cargar.", vbInformation, "Aviso"
        bValida = False
    ElseIf nMonto = nAcumMN Then
        cmdAgregar.Enabled = False
    Else
        cmdAgregar.Enabled = True
    End If
Else
    If nMonto < nAcumME Then
        MsgBox "SUMA TOTAL supera al monto establecido para cargar.", vbInformation, "Aviso"
        bValida = False
    ElseIf nMonto = nAcumME Then
        cmdAgregar.Enabled = False
    Else
        cmdAgregar.Enabled = True
    End If
End If
grdCuentaAbono.Row = nFila
grdCuentaAbono.Col = nCol
If bValida Then
    lblTotalMN = Format$(nAcumMN, "#,##0.00")
    lblTotalME = Format$(nAcumME, "#,##0.00")
Else
    grdCuentaAbono.TextMatrix(nFila, 3) = "0.00"
    grdCuentaAbono.TextMatrix(nFila, 4) = "0.00"
End If
End Sub

Private Function ObtieneDatosCuenta(ByVal sCuenta As String, Optional bArchivo As Boolean = False) As Boolean
Dim clsMant As NCapMantenimiento
Dim clsCap As NCapMovimientos
Dim rsCta As Recordset, rsRel As Recordset
Dim nEstado As CaptacEstado
Dim nRow As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Set clsCap = New NCapMovimientos
sMsg = clsCap.ValidaCuentaOperacion(sCuenta)
Set clsCap = Nothing
If sMsg = "" Then
    Set clsMant = New NCapMantenimiento
    Set rsCta = New Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        nMoneda = CLng(Mid(sCuenta, 9, 1))
        If nMoneda = gMonedaNacional Then
            sMoneda = "MONEDA NACIONAL"
            txtMontoCargo.BackColor = &HC0FFFF
            lblMon.Caption = "S/."
        Else
            sMoneda = "MONEDA EXTRANJERA"
            txtMontoCargo.BackColor = &HC0FFC0
            lblMon.Caption = "$"
        End If
        
        If rsCta("bOrdPag") Then
            lblMensaje = "AHORROS CON ORDEN DE PAGO" & Chr$(13) & sMoneda
        Else
            lblMensaje = "AHORROS SIN ORDEN DE PAGO" & Chr$(13) & sMoneda
        End If
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        sPersona = ""
        Do While Not rsRel.EOF
            If sPersona <> rsRel("cPersCod") Then
                grdCliente.AdicionaFila
                nRow = grdCliente.Rows - 1
                grdCliente.TextMatrix(nRow, 1) = UCase(PstaNombre(rsRel("Nombre")))
                grdCliente.TextMatrix(nRow, 2) = Left(UCase(rsRel("Relacion")), 2)
                grdCliente.TextMatrix(nRow, 3) = rsRel!cPersCod
                grdCliente.TextMatrix(nRow, 4) = Trim(rsRel("nPrdPersRelac"))
                sPersona = rsRel("cPersCod")
            End If
            rsRel.MoveNext
        Loop
        rsRel.Close
        Set rsRel = Nothing
        txtCuenta.Enabled = False
        txtMontoCargo.Enabled = True
        txtMontoCargo.SetFocus
        cmdCancelar.Enabled = True
        txtCuenta.Age = Mid(sCuenta, 4, 2)
        txtCuenta.Cuenta = Mid(sCuenta, 9, 10)
        fraCuentaAbono.Enabled = True
        cmdAgregar.Enabled = True
        ObtieneDatosCuenta = True
    End If
Else
    If bArchivo Then
        sError = sError & sMsg & gPrnSaltoLinea
    Else
        MsgBox sMsg, vbInformation, "Operacion"
        txtCuenta.SetFocus
    End If
    ObtieneDatosCuenta = False
End If
End Function

Private Function ObtieneDatosCuentaAbono(ByVal sCuenta As String, Optional bArchivo As Boolean = False, _
        Optional nMonto As Double = 0) As Boolean
Dim clsMant As NCapMantenimiento
Dim clsCap As NCapMovimientos
Dim rsCta As Recordset, rsRel As Recordset
Dim nEstado As CaptacEstado
Dim nFila As Long
Dim sMsg As String, sMoneda As String, sPersona As String
Dim nMonedaAbono As Moneda
Set clsCap = New NCapMovimientos
sMsg = clsCap.ValidaCuentaOperacion(sCuenta, True)
Set clsCap = Nothing
If sMsg = "" Then
    Set clsMant = New NCapMantenimiento
    Set rsCta = New Recordset
    Set rsCta = clsMant.GetDatosCuenta(sCuenta)
    If Not (rsCta.EOF And rsCta.BOF) Then
        grdCuentaAbono.AdicionaFila
        nFila = grdCuentaAbono.Rows - 1
        grdCuentaAbono.TextMatrix(nFila, 1) = sCuenta
        nMonedaAbono = CLng(Mid(sCuenta, 9, 1))
        
        Set rsRel = clsMant.GetPersonaCuenta(sCuenta)
        sPersona = ""
        Do While Not rsRel.EOF
            If sPersona <> rsRel("cPersCod") And rsRel("nPrdPersRelac") = gCapRelPersTitular Then
                grdCuentaAbono.TextMatrix(nFila, 2) = UCase(PstaNombre(rsRel("Nombre")))
                Exit Do
            End If
            rsRel.MoveNext
        Loop
        rsRel.Close
        Set rsRel = Nothing
        
        If nMonedaAbono = gMonedaNacional Then
            grdCuentaAbono.BackColorRow vbWhite
            grdCuentaAbono.BackColorControl = vbWhite
            grdCuentaAbono.TextMatrix(nFila, 3) = nMonto
        Else
            grdCuentaAbono.BackColorRow &HC0FFC0
            grdCuentaAbono.BackColorControl = &HC0FFC0
            grdCuentaAbono.TextMatrix(nFila, 4) = nMonto
        End If
        
        If Not bArchivo Then
            grdCuentaAbono.lbEditarFlex = True
            grdCuentaAbono.SetFocus
            cmdEliminar.Enabled = True
            cmdGrabar.Enabled = True
        End If
        ObtieneDatosCuentaAbono = True
    End If
Else
    If bArchivo Then
        sError = sError & sMsg & gPrnSaltoLinea
    Else
        MsgBox sMsg, vbInformation, "Operacion"
        cmdAgregar.SetFocus
    End If
    ObtieneDatosCuentaAbono = False
End If
If Not bArchivo Then
    txtCuentaAbo.Visible = False
End If
End Function

Private Sub LimpiaControles()
grdCliente.Clear
grdCliente.Rows = 2
grdCliente.FormaCabecera
grdCuentaAbono.Clear
grdCuentaAbono.Rows = 2
grdCuentaAbono.FormaCabecera
txtMontoCargo.value = 0
cmdGrabar.Enabled = False
txtCuenta.Age = ""
txtCuenta.Cuenta = ""
txtCuentaAbo.Age = ""
txtCuentaAbo.Cuenta = ""
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
fraCuentaAbono.Enabled = False
txtGlosa = ""
fraGlosa.Enabled = False
txtCuenta.Enabled = True
txtCuenta.SetFocus
lblMensaje = ""
lblTotalMN = ""
lblTotalME = ""
End Sub

Private Sub cmdAgregar_Click()
txtCuentaAbo.Age = ""
txtCuentaAbo.Cuenta = ""
txtCuentaAbo.Visible = True
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
txtMontoCargo.Enabled = False
txtCuentaAbo.SetFocus
End Sub

Private Sub cmdCancelar_Click()
LimpiaControles
End Sub

Private Sub CmdEliminar_Click()
If MsgBox("¿Desea Eliminar la cuenta de la Relación?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    grdCuentaAbono.EliminaFila grdCuentaAbono.Row
    If Trim(grdCuentaAbono.TextMatrix(1, 1)) = "" Then
        cmdEliminar.Enabled = False
        cmdGrabar.Enabled = False
    End If
    CalculaTotales
End If
End Sub
'****Agregado MPBR
Private Function ObtTitular() As String
Dim I As Integer
For I = 1 To grdCliente.Rows - 1
If Right(grdCliente.TextMatrix(I, 4), 2) = "10" Then
      ObtTitular = Trim(grdCliente.TextMatrix(I, 3))
      Exit For
  End If
Next I
End Function
Private Sub cmdGrabar_Click()
Dim nMontoCargo As Double
Dim sCuenta As String, sGlosa As String
nMontoCargo = txtMontoCargo.value
sCuenta = txtCuenta.NroCuenta

If lblTotalMN = "" Or lblTotalME = "" Then
    MsgBox "Debe ingresar cuenta(s) para el abono", vbInformation, "Aviso"
    cmdAgregar.SetFocus
    Exit Sub
End If

If nMontoCargo = 0 Then
    MsgBox "Monto de Cargo debe ser mayor a cero", vbInformation, "Aviso"
    txtMontoCargo.SetFocus
    Exit Sub
End If
If nMoneda = gMonedaNacional Then
    If nMontoCargo <> CDbl(lblTotalMN) Then
        MsgBox "Suma total no coincide como monto de cargo", vbInformation, "Aviso"
        cmdAgregar.SetFocus
        Exit Sub
    End If
Else
    If nMontoCargo <> CDbl(lblTotalME) Then
        MsgBox "Suma total no coincide como monto de cargo", vbInformation, "Aviso"
        cmdAgregar.SetFocus
        Exit Sub
    End If
End If

If MsgBox("¿Está seguro de grabar la información?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    Dim clsCap As NCapMovimientos
    Dim sMovNro As String
    Dim clsMov As NContFunciones
    Dim clsMant As NCapMantenimiento
    Dim rsCtaAbo As Recordset
    Set clsMov = New NContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set clsMov = Nothing
    
     On Error GoTo ErrGraba
    
    Dim clsLav As nCapDefinicion, clsExo As NCapServicios, sPersLavDinero As String
    Dim nMontoLavDinero As Double, nTC As Double

    'Realiza la Validación para el Lavado de Dinero
    sCuenta = txtCuenta.NroCuenta
    Set clsLav = New nCapDefinicion
    'If clsLav.EsOperacionEfectivo(Trim(nOperacion)) Then
        Set clsExo = New NCapServicios
        If Not clsExo.EsCuentaExoneradaLavadoDinero(sCuenta) Then
            Set clsExo = Nothing
            sPersLavDinero = ""
            nMontoLavDinero = clsLav.GetCapParametro(gMonOpeLavDineroME)
            Set clsLav = Nothing
            If nMoneda = gMonedaNacional Then
                Dim clsTC As nTipoCambio
                Set clsTC = New nTipoCambio
                nTC = clsTC.EmiteTipoCambio(gdFecSis, TCFijoDia)
                Set clsTC = Nothing
            Else
                nTC = 1
            End If
            If nMontoCargo >= Round(nMontoLavDinero * nTC, 2) Then
                sPersLavDinero = IniciaLavDinero()
                If sPersLavDinero = "" Then Exit Sub
            End If
        Else
            Set clsExo = Nothing
        End If
    'Else
    '    Set clsLav = Nothing
   ' End If
    
       
    Set clsCap = New NCapMovimientos
    Set rsCtaAbo = grdCuentaAbono.GetRsNew()
    sGlosa = Trim(txtGlosa)
    'ALPA 20081009********************************
    'If clsCap.CapTransferenciaAho(sCuenta, nMontoCargo, sMovNro, rsCtaAbo, sGlosa, gsNomAge, sLpt, , CDbl(Me.lblTCC.Caption), CDbl(Me.lblTCV.Caption)) Then
    If clsCap.CapTransferenciaAho(sCuenta, nMontoCargo, sMovNro, rsCtaAbo, sGlosa, gsNomAge, sLpt, , CDbl(Me.lblTCC.Caption), CDbl(Me.lblTCV.Caption), , , , , , gnMovNro) Then
        cmdCancelar_Click
    End If
    'ALPA 20081009********************************
    If gnMovNro > 0 Then
        Call frmMovLavDinero.InsertarLavDinero(, , , gnMovNro, , , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen)
    End If
    '*********************************************
End If
Exit Sub
ErrGraba:
    MsgBox Err.Description, vbExclamation, "Error"
    Exit Sub
End Sub


Private Function IniciaLavDinero() As String
Dim I As Long
Dim nRelacion As CaptacRelacPersona
Dim sPersCod As String, sNombre As String, sDocId As String, sDireccion As String
Dim nPersoneria As PersPersoneria, sOperacion As String, sTipoCuenta As String
Dim nMonto As Double
Dim sCuenta As String
sOperacion = CStr(nOperacion)

For I = 1 To grdCuentaAbono.Rows - 1
    nRelacion = CLng(Trim(Right(grdCliente.TextMatrix(I, 3), 4)))
'    If npersoneria = gPersonaNat Then
'        If nRelacion = gCapRelPersApoderado Or nRelacion = gCapRelPersTitular Then
'            sPersCod = grdCliente.TextMatrix(i, 3)
'            sNombre = grdCliente.TextMatrix(i, 1)
'            sDireccion = ""
'            sDocId = ""
'            Exit For
'        End If
'    Else
'        If nRelacion = gCapRelPersTitular Then
            sPersCod = grdCliente.TextMatrix(I, 3)
            sNombre = grdCliente.TextMatrix(I, 1)
            sDireccion = ""
            sDocId = ""
            Exit For
'        End If
'    End If
Next I
nMonto = txtMontoCargo.value
sCuenta = txtCuenta.NroCuenta
'If sPersCodCMAC <> "" Then
'    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, False, True, nMonto, sCuenta, sOperacion, , sTipoCuenta)
'Else
    'ALPA 20081009******************************************************************************************************************
    'IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, False, nMonto, sCuenta, soperacion, , sTipoCuenta)
    IniciaLavDinero = frmMovLavDinero.Inicia(sPersCod, sNombre, sDireccion, sDocId, True, False, nMonto, sCuenta, sOperacion, , sTipoCuenta, , , , , , , gnTipoREU, gnMontoAcumulado, gsOrigen)
    '*******************************************************************************************************************************
'End If
End Function



Private Sub cmdObtDatos_Click()
Dim sArchivo As String
On Local Error Resume Next
CdlgFile.CancelError = True
'Especificar las extensiones a usar
CdlgFile.DefaultExt = "*.txt"
CdlgFile.Filter = "Textos (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
CdlgFile.ShowOpen
If Err Then
    sArchivo = "" 'Cancelada la operación de abrir
Else
    sArchivo = CdlgFile.Filename
    ObtieneDatosCuentasAbonar sArchivo
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(gCapAhorros, False)
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Me.Caption = "Captaciones - Ahorros - Transferencia de Cuentas"
txtCuenta.CMAC = gsCodCMAC
txtCuenta.Prod = Trim(gCapAhorros)
txtCuentaAbo.CMAC = gsCodCMAC
txtCuentaAbo.Prod = Trim(gCapAhorros)
txtCuenta.EnabledProd = False
txtCuentaAbo.EnabledProd = False
txtCuenta.EnabledCMAC = False
txtCuentaAbo.EnabledCMAC = False
txtCuentaAbo.Visible = False
Dim clsGen As nTipoCambio
Dim rsTC As Recordset
Set clsGen = New nTipoCambio
lblTCC = Format$(clsGen.EmiteTipoCambio(gdFecSis, TCCompra), "#0.0000")
lblTCV = Format$(clsGen.EmiteTipoCambio(gdFecSis, TCVenta), "#0.0000")
Set clsGen = Nothing
fraCuentaAbono.Enabled = False
fraGlosa.Enabled = False
cmdGrabar.Enabled = False
cmdCancelar.Enabled = False
End Sub

Private Sub grdCuentaAbono_OnCellChange(pnRow As Long, pnCol As Long)
Dim nMonCta As Moneda
Dim nMonto As Double
nMonCta = CLng(Mid(grdCuentaAbono.TextMatrix(pnRow, 1), 9, 1))
nMonto = CDbl(grdCuentaAbono.TextMatrix(pnRow, pnCol))

If pnCol = 3 Or pnCol = 4 Then
    If nMoneda = gMonedaNacional Then
        If nMonCta = nMoneda Then
            grdCuentaAbono.TextMatrix(pnRow, 4) = "0.00"
        Else
            If pnCol = 4 Then
                grdCuentaAbono.TextMatrix(pnRow, 3) = Format$(nMonto * CDbl(lblTCV), "#0.00")
            Else
                grdCuentaAbono.TextMatrix(pnRow, 4) = Format$(nMonto / CDbl(lblTCV), "#0.00")
            End If
        End If
    ElseIf nMoneda = gMonedaExtranjera Then
        If nMonCta = nMoneda Then
            grdCuentaAbono.TextMatrix(pnRow, 3) = "0.00"
        Else
            If pnCol = 4 Then
                grdCuentaAbono.TextMatrix(pnRow, 3) = Format$(nMonto * CDbl(lblTCC), "#0.00")
            Else
                grdCuentaAbono.TextMatrix(pnRow, 4) = Format$(nMonto / CDbl(lblTCC), "#0.00")
            End If
        End If
    End If
End If
CalculaTotales
End Sub

Private Sub grdCuentaAbono_RowColChange()
If grdCuentaAbono.TextMatrix(grdCuentaAbono.Row, 1) <> "" Then
    If CLng(Mid(grdCuentaAbono.TextMatrix(grdCuentaAbono.Row, 1), 9, 1)) = gMonedaNacional Then
        grdCuentaAbono.BackColorControl = vbWhite
    Else
        grdCuentaAbono.BackColorControl = &HC0FFC0
    End If
End If
End Sub


Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String
    sCta = txtCuenta.NroCuenta
    ObtieneDatosCuenta sCta
End If
End Sub

Private Sub txtCuentaAbo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Dim sCta As String, sCtaCargo As String
    sCta = txtCuentaAbo.NroCuenta
    sCtaCargo = txtCuenta.NroCuenta
    If sCta = sCtaCargo Then
        MsgBox "La Cuenta de Abono no puede ser la misma cuenta de Cargo.", vbInformation, "Aviso"
        txtCuentaAbo.SetFocusCuenta
        Exit Sub
    End If
    If Not CuentaExisteEnLista(sCta) Then
        ObtieneDatosCuentaAbono sCta
        cmdGrabar.Enabled = True
        cmdCancelar.Enabled = True
        txtMontoCargo.Enabled = True
    Else
        MsgBox "Cuenta ya se encuentra en la lista.", vbInformation, "Aviso"
        txtCuentaAbo.SetFocusCuenta
    End If
End If
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    cmdAgregar.SetFocus
End If
End Sub

Private Sub txtIdAut_KeyPress(KeyAscii As Integer)

 Dim rs As ADODB.Recordset
   Set rs = New ADODB.Recordset
   Gtitular = ObtTitular
   If Gtitular = "" Then
    MsgBox "Esta cuenta no tiene titular", vbOKOnly + vbInformation, "Atención"
    Exit Sub
   End If
   nOperacion = gAhoTransferencia
   If KeyAscii = 13 And Trim(txtIdAut.Text) <> "" And Len(txtCuenta.NroCuenta) = 18 Then
      'Set rs = gAutorizacion.SAA(Left(CStr(nOperacion), 4) & "00", Vusuario, txtCuenta.NroCuenta, GTitular, CInt(nMoneda), CLng(txtIdAut.Text))
     If rs.State = 1 Then
       If rs.RecordCount > 0 Then
        txtMontoCargo.Text = rs!nMontoAprobado
      Else
          MsgBox "No Existe este Id de Autorización para esta cuenta." & vbCrLf & "Consulte las Operaciones Pendientes.", vbOKOnly + vbInformation, "Atención"
          txtIdAut.Text = ""
       End If
       
     End If
   End If

 If (KeyAscii < Asc("0") Or KeyAscii > Asc("9")) And Not (KeyAscii = 13 Or KeyAscii = 8) Then
      KeyAscii = 0
   End If
End Sub

Private Sub txtMontoCargo_GotFocus()
txtMontoCargo.MarcaTexto
End Sub

Private Sub txtMontoCargo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If txtMontoCargo.value > 0 Then
        fraCuentaAbono.Enabled = True
        fraGlosa.Enabled = True
        cmdEliminar.Enabled = False
        txtGlosa.SetFocus
    Else
        MsgBox "Monto debe ser mayor a cero", vbInformation, "Aviso"
        txtMontoCargo.SetFocus
    End If
End If
End Sub



Private Function Cargousu(ByVal NomUser As String)
 Dim SQLAUX As String, RsAUX As ADODB.Recordset, oConecta As DConecta
 SQLAUX = "   SELECT RC.CRHCARGOCOD FROM RRHH RH "
 SQLAUX = SQLAUX & "  INNER JOIN RHCARGOS RC ON RC.CPERSCOD=RH.CPERSCOD "
 SQLAUX = SQLAUX & " WHERE RH.CUSER='" & Vusuario & "'"
 Set oConecta = New DConecta
 Set RsAUX = New ADODB.Recordset
    oConecta.AbreConexion
    RsAUX.CursorLocation = adUseClient
    Set RsAUX = oConecta.CargaRecordSet(SQLAUX)
    Cargousu = RsAUX.Fields(0).value
End Function

