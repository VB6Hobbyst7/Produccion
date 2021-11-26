VERSION 5.00
Begin VB.Form frmCredRefinanc 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Refinanciar Creditos"
   ClientHeight    =   6330
   ClientLeft      =   2325
   ClientTop       =   1755
   ClientWidth     =   7440
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   7440
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   6270
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   7245
      Begin VB.Frame Frame2 
         Height          =   2730
         Left            =   120
         TabIndex        =   5
         Top             =   3405
         Width           =   6990
         Begin VB.CheckBox ChkCapGastos 
            Caption         =   "Capitalizar Gastos"
            Height          =   240
            Left            =   4110
            TabIndex        =   18
            ToolTipText     =   "Active en Check si desea Capitalizar los Gastos"
            Top             =   285
            Visible         =   0   'False
            Width           =   2295
         End
         Begin VB.CommandButton cmdcancelar 
            Caption         =   "&Cancelar"
            Enabled         =   0   'False
            Height          =   345
            Left            =   5640
            TabIndex        =   8
            Top             =   2145
            Width           =   1140
         End
         Begin VB.CommandButton cmdaceptar 
            Caption         =   "&Aceptar"
            Enabled         =   0   'False
            Height          =   345
            Left            =   4440
            TabIndex        =   7
            Top             =   2145
            Width           =   1140
         End
         Begin SICMACT.ActXCodCta ActxCta 
            Height          =   420
            Left            =   165
            TabIndex        =   6
            Top             =   210
            Width           =   3735
            _ExtentX        =   6588
            _ExtentY        =   741
            Texto           =   "Credito :"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
         Begin VB.Label lblTipoCredito 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1320
            TabIndex        =   22
            Top             =   720
            Width           =   5430
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Credito:"
            Height          =   195
            Left            =   180
            TabIndex        =   21
            Top             =   720
            Width           =   900
         End
         Begin VB.Label lblTipoProducto 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1320
            TabIndex        =   20
            Top             =   1080
            Width           =   5430
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Producto :"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   1080
            Width           =   1095
         End
         Begin VB.Label LblIntComp 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   4200
            TabIndex        =   14
            Top             =   1770
            Width           =   1260
         End
         Begin VB.Label Label1 
            Caption         =   "Monto a Capitalizar :"
            Height          =   225
            Left            =   2715
            TabIndex        =   13
            Top             =   1785
            Width           =   1515
         End
         Begin VB.Label Label3 
            Caption         =   "Saldo Capital :"
            Height          =   225
            Left            =   180
            TabIndex        =   12
            Top             =   1785
            Width           =   1050
         End
         Begin VB.Label lblsalcap 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1320
            TabIndex        =   11
            Top             =   1755
            Width           =   1260
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Titular :"
            Height          =   195
            Left            =   180
            TabIndex        =   10
            Top             =   1425
            Width           =   525
         End
         Begin VB.Label lbltitular 
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   270
            Left            =   1320
            TabIndex        =   9
            Top             =   1425
            Width           =   5430
         End
      End
      Begin VB.Frame FraDatos 
         Height          =   3285
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   7005
         Begin VB.CommandButton CmdSalir 
            Caption         =   "&Salir"
            Height          =   375
            Left            =   2865
            TabIndex        =   17
            Top             =   2760
            Width           =   1320
         End
         Begin VB.CommandButton CmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   375
            Left            =   1500
            TabIndex        =   4
            Top             =   2760
            Width           =   1320
         End
         Begin VB.CommandButton CmdNuevo 
            Caption         =   "&Adicionar"
            Height          =   375
            Left            =   135
            TabIndex        =   3
            Top             =   2760
            Width           =   1320
         End
         Begin SICMACT.FlexEdit FECredRef 
            Height          =   2385
            Left            =   120
            TabIndex        =   2
            Top             =   270
            Width           =   6795
            _ExtentX        =   11986
            _ExtentY        =   4207
            Cols0           =   12
            HighLight       =   1
            AllowUserResizing=   3
            RowSizingMode   =   1
            EncabezadosNombres=   "-Credito-Monto-Capital-Inter Comp.-Inter. Morat.-Inter. Gracia-Inter. Susp.-Inter. Reprog-Gastos-MontoSol-Titular"
            EncabezadosAnchos=   "400-1200-1200-1200-1200-1200-1200-1200-1200-1200-0-2500"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-C-R-R-R-R-R-R-R-R-C-L"
            FormatosEdit    =   "0-0-2-2-2-2-2-2-2-2-0-0"
            SelectionMode   =   1
            lbUltimaInstancia=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   8421376
         End
         Begin VB.Label LblMontoRef 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            Caption         =   "0.00"
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
            Height          =   285
            Left            =   5850
            TabIndex        =   16
            Top             =   2805
            Width           =   1080
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "A Refinanciar :"
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
            Left            =   4470
            TabIndex        =   15
            Top             =   2835
            Width           =   1290
         End
      End
   End
End
Attribute VB_Name = "frmCredRefinanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public nProducto As Producto
Private MatCredRef() As String
Private nPrestamo As Double
Private dFecVig As Date
Private nTasa As Double
Private MatCalend As Variant
Private nMoneda As Integer
Private bCapitlaInt As Boolean
Public nDestino As Integer 'JAME20140328


Public Function Inicio(ByVal pnMoneda As Moneda, pMatCalend As Variant, ByVal pbCapitlaInt As Boolean, ByVal pbSustiDeudor As Boolean, Optional ByRef pnDestino As Integer = 0) As Variant
Dim i As Integer
Dim nMonto As Double
Dim oCredito As COMDCredito.DCOMCredito
Dim R As New ADODB.Recordset
    bCapitlaInt = pbCapitlaInt
    nMoneda = pnMoneda
    If IsArray(pMatCalend) Then
        nMonto = 0
        For i = 0 To UBound(pMatCalend) - 1
            FECredRef.AdicionaFila
            FECredRef.TextMatrix(i + 1, 1) = pMatCalend(i, 0)
            FECredRef.TextMatrix(i + 1, 2) = pMatCalend(i, 1)
            FECredRef.TextMatrix(i + 1, 3) = pMatCalend(i, 2)
            FECredRef.TextMatrix(i + 1, 4) = pMatCalend(i, 3)
            FECredRef.TextMatrix(i + 1, 5) = pMatCalend(i, 4)
            FECredRef.TextMatrix(i + 1, 6) = pMatCalend(i, 5)
            FECredRef.TextMatrix(i + 1, 7) = pMatCalend(i, 6)
            FECredRef.TextMatrix(i + 1, 8) = pMatCalend(i, 7)
            FECredRef.TextMatrix(i + 1, 9) = pMatCalend(i, 8)
            FECredRef.TextMatrix(i + 1, 10) = pMatCalend(i, 9)
            
            Set oCredito = New DCOMCredito
            Set R = oCredito.RecuperaRelacPers(pMatCalend(i, 0))
            Set oCredito = Nothing
            
            If Not (R.EOF And R.BOF) Then
                Do Until R.EOF
                    If R!nOrden = 1 Then
                       FECredRef.TextMatrix(i + 1, 11) = PstaNombre(R!cPersNombre)
                       Exit Do
                    End If
                    R.MoveNext
                Loop
            End If
            
            nMonto = nMonto + CDbl(FECredRef.TextMatrix(FECredRef.row, 2))
        Next i
        LblMontoRef.Caption = Format(nMonto, "#0.00")
    End If
    
    If pbSustiDeudor Then
       Label2.Caption = "A Sustituir :"
       frmCredRefinanc.Caption = "Sustituir Créditos"
    Else
       Label2.Caption = "A Refinanciar :"
       frmCredRefinanc.Caption = "Refinanciar Créditos"
    End If
    
    Me.Show 1
    pnDestino = nDestino
    Inicio = MatCredRef
End Function

Private Sub HabiltaIngreso(ByVal pbHabilita As Boolean)
    ActxCta.Enabled = pbHabilita
    Label5.Enabled = pbHabilita
    lbltitular.Enabled = pbHabilita
    Label3.Enabled = pbHabilita
    lblsalcap.Enabled = pbHabilita
    'cmdaceptar.Enabled = pbHabilita
    'cmdcancelar.Enabled = pbHabilita
    FraDatos.Enabled = Not pbHabilita
End Sub
Private Sub LimpiaPantalla()
    LimpiaControles Me
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
End Sub

Private Function CargaDatos(ByVal psCtaCod As String) As Boolean
Dim oDCred As COMDCredito.DCOMCredito
Dim oNCred As COMNCredito.NCOMCredito
Dim R As ADODB.Recordset
Dim lnInteresDiferido As Currency 'ALPA20150908
    On Error GoTo ErrorCargaDatos
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.RecuperaDatosComunes(psCtaCod, False)
    Set oDCred = Nothing
    
    'ARLO20180321 ERS070 - 2017 - ANEXO 02
    If (FECredRef.TextMatrix(FECredRef.rows - 1, 1)) = "" Then
        dFecVig = "12:00:00 AM"
    End If
    'ARLO20180321 ERS070 - 2017 - ANEXO 02
    
    If Not R.BOF And Not R.EOF Then
        CargaDatos = True
        lbltitular.Caption = Trim(R!cTitular)
        lblsalcap.Caption = Format(R!nSaldo, "#0.00")
        nPrestamo = CDbl(Format(R!nMontoCol, "#0.00"))
        'Jame 200142803 *************************************
        If dFecVig <> "12:00:00 AM" Then
            If CDate(Format(R!dVigencia, "dd/mm/yyyy")) < dFecVig Then
                nDestino = R!nColocDestino
            End If
        Else
            nDestino = R!nColocDestino
        End If
        'Fin Jame **********************************
        dFecVig = CDate(Format(R!dVigencia, "dd/mm/yyyy"))
        nTasa = CDbl(Format(R!nTasaInteres, "#0.00"))
        'ALPA 20100606***************
        lblTipoCredito.Caption = R!cTpoCredDes
        lblTipoProducto.Caption = R!cTpoProdDes
        '****************************
        R.Close
        Set R = Nothing
        
        Set oNCred = New COMNCredito.NCOMCredito
        MatCalend = oNCred.RecuperaMatrizCalendarioPendiente(psCtaCod)
        
        'ALPA 20150908**************************************
        Dim oCredito As COMDCredito.DCOMCredito
        Set oCredito = New COMDCredito.DCOMCredito
        Dim R2 As ADODB.Recordset
        Set R2 = New ADODB.Recordset
        Set R2 = oCredito.RecuperaColocacCred(psCtaCod)
        If R2.RecordCount > 0 Then
            lnInteresDiferido = 0
            R2.Close
        End If
        '***************************************************
        
        
        'LblIntComp.Caption = oNCred.InteresGastosAFecha(psCtaCod, gdFecSis, nPrestamo, nFecVig, nTasa)
        'If ChkCapGastos.value = 1 Then
        '    LblIntComp.Caption = oNCred.MatrizInteresGastosAFecha(psCtaCod, MatCalend, gdFecSis)
        'Else
            'MAVM 11 Set 2009 se agrego el Monto del Interes de Gracia
            LblIntComp.Caption = Format(oNCred.MatrizInteresGastosAFecha(psCtaCod, MatCalend, gdFecSis) + oNCred.MatrizInteresGraAFecha(psCtaCod, MatCalend, gdFecSis), "#0.00")
            'LblIntComp.Caption = Format(CDbl(LblIntComp.Caption) - oNCred.MatrizGastosFecha(psCtaCod, MatCalend) + oNCred.TotalGastosAFecha(psCtaCod, gdFecSis), "#0.00") 'JUEZ 20140922
            LblIntComp.Caption = Format(CDbl(LblIntComp.Caption) - oNCred.MatrizGastosFecha(psCtaCod, MatCalend) + oNCred.TotalGastosAFecha(psCtaCod, gdFecSis), "#0.00") + lnInteresDiferido 'ALPA 20150908
            
        'End If
        ActxCta.Enabled = False
        ChkCapGastos.Enabled = False
        cmdaceptar.Enabled = True
        cmdcancelar.Enabled = True
        Set oNCred = Nothing
    Else
        CargaDatos = False
        R.Close
        ActxCta.Enabled = True
        'cmdaceptar.Enabled = False
        'cmdcancelar.Enabled = False
        Set R = Nothing
    End If
    Exit Function

ErrorCargaDatos:
    MsgBox Err.Description, vbCritical, "Aviso"
    
End Function


Private Sub ActxCta_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(nProducto, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Not CargaDatos(ActxCta.NroCuenta) Then
            MsgBox "No se Pudo encontrar el Credito o No esta Vigente", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
End Sub

Private Sub CmdAceptar_Click()
Dim oNegCred As COMNCredito.NCOMCredito
Dim i As Integer
Dim nTipoCambioFijo As Double
Dim oGeneral As COMDConstSistema.DCOMGeneral
 Dim lnInteresDiferido As Currency
    If Len(Trim(lbltitular.Caption)) <= 0 Then
        MsgBox "Digite un Credito", vbInformation, "Aviso"
        Exit Sub
    End If
    
    'ALPA 20150908**************************************
        Dim oCredito As COMDCredito.DCOMCredito
        Set oCredito = New COMDCredito.DCOMCredito
        Dim R2 As ADODB.Recordset
        Set R2 = New ADODB.Recordset
        Set R2 = oCredito.RecuperaColocacCred(ActxCta.NroCuenta)
    If R2.RecordCount > 0 Then
            lnInteresDiferido = 0
            R2.Close
    End If
    '***************************************************
    
    Set oNegCred = New COMNCredito.NCOMCredito
        FECredRef.AdicionaFila
        FECredRef.TextMatrix(FECredRef.rows - 1, 1) = ActxCta.NroCuenta
        If bCapitlaInt Then
            FECredRef.TextMatrix(FECredRef.rows - 1, 2) = CDbl(Format(CDbl(LblIntComp.Caption) + CDbl(lblsalcap.Caption), "#0.00"))
        Else
            FECredRef.TextMatrix(FECredRef.rows - 1, 2) = CDbl(Format(CDbl(lblsalcap.Caption) + CDbl(lnInteresDiferido), "#0.00"))
        End If
        FECredRef.TextMatrix(FECredRef.rows - 1, 3) = lblsalcap.Caption
        FECredRef.TextMatrix(FECredRef.rows - 1, 4) = Format(oNegCred.MatrizInteresCompAFecha(ActxCta.NroCuenta, MatCalend, gdFecSis, nPrestamo, nTasa) + oNegCred.MatrizInteresCompVencidoFecha(ActxCta.NroCuenta, MatCalend), "#0.00") + lnInteresDiferido
        FECredRef.TextMatrix(FECredRef.rows - 1, 5) = Format(oNegCred.MatrizInteresMorFecha(ActxCta.NroCuenta, MatCalend), "#0.00")
        FECredRef.TextMatrix(FECredRef.rows - 1, 6) = Format(oNegCred.MatrizInteresGraciaFecha(ActxCta.NroCuenta, MatCalend, gdFecSis, nPrestamo), "#0.00")
        FECredRef.TextMatrix(FECredRef.rows - 1, 7) = Format(oNegCred.MatrizInteresSuspensoFecha(ActxCta.NroCuenta, MatCalend), "#0.00")
        FECredRef.TextMatrix(FECredRef.rows - 1, 8) = Format(oNegCred.MatrizInteresReprogramadoFecha(ActxCta.NroCuenta, MatCalend), "#0.00")
        'FECredRef.TextMatrix(FECredRef.Rows - 1, 9) = Format(oNegCred.MatrizGastosFecha(ActxCta.NroCuenta, MatCalend), "#0.00")
        FECredRef.TextMatrix(FECredRef.rows - 1, 9) = Format(oNegCred.TotalGastosAFecha(ActxCta.NroCuenta, gdFecSis), "#0.00") 'JUEZ 20140922
        FECredRef.TextMatrix(FECredRef.rows - 1, 11) = lbltitular
    Set oNegCred = Nothing
    
    Set oGeneral = New COMDConstSistema.DCOMGeneral
    nTipoCambioFijo = oGeneral.EmiteTipoCambio(gdFecSis, TCFijoMes)
    Set oGeneral = Nothing
    If nMoneda <> CInt(Mid(ActxCta.NroCuenta, 9, 1)) Then
        If nMoneda = gMonedaNacional Then 'De Dolares a Soles
            FECredRef.TextMatrix(FECredRef.rows - 1, 10) = Format(CDbl(FECredRef.TextMatrix(FECredRef.rows - 1, 2)) * nTipoCambioFijo, "#0.00")
        Else 'De Soles a Dolares
            FECredRef.TextMatrix(FECredRef.rows - 1, 10) = Format(CDbl(FECredRef.TextMatrix(FECredRef.rows - 1, 2)) / nTipoCambioFijo, "#0.00")
        End If
    Else
        FECredRef.TextMatrix(FECredRef.rows - 1, 10) = FECredRef.TextMatrix(FECredRef.rows - 1, 2)
    End If
    
    HabiltaIngreso False
    Call LimpiaPantalla
     LblMontoRef.Caption = "0.00"
    For i = 1 To FECredRef.rows - 1
        LblMontoRef.Caption = CDbl(LblMontoRef.Caption) + CDbl(FECredRef.TextMatrix(i, 10))
    Next i
    cmdaceptar.Enabled = False
    cmdcancelar.Enabled = False
End Sub

Private Sub cmdCancelar_Click()
    HabiltaIngreso False
    Call LimpiaPantalla
    cmdaceptar.Enabled = False
    cmdcancelar.Enabled = False
End Sub

Private Sub CmdEliminar_Click()
    FECredRef.EliminaFila FECredRef.row
End Sub

Private Sub cmdNuevo_Click()
    HabiltaIngreso True
    Call LimpiaPantalla
    ActxCta.SetFocusProd
    cmdcancelar.Enabled = True
    ChkCapGastos.Enabled = True
    ChkCapGastos.value = 0
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
       
    If KeyCode = vbKeyF12 And ActxCta.Enabled = True Then 'F12
        Dim bRetSinTarjeta As Boolean
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.inicia(nProducto, bRetSinTarjeta)
        If sCuenta <> "" Then
            ActxCta.NroCuenta = sCuenta
            ActxCta.SetFocusCuenta
        End If
    End If
    
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    CentraSdi Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim i As Integer

    If Trim(FECredRef.TextMatrix(1, 1)) <> "" Then
        ReDim MatCredRef(FECredRef.rows - 1, 10)
        For i = 1 To FECredRef.rows - 1
            MatCredRef(i - 1, 0) = FECredRef.TextMatrix(i, 1)
            MatCredRef(i - 1, 1) = FECredRef.TextMatrix(i, 2)
            MatCredRef(i - 1, 2) = FECredRef.TextMatrix(i, 3)
            MatCredRef(i - 1, 3) = FECredRef.TextMatrix(i, 4)
            MatCredRef(i - 1, 4) = FECredRef.TextMatrix(i, 5)
            MatCredRef(i - 1, 5) = FECredRef.TextMatrix(i, 6)
            MatCredRef(i - 1, 6) = FECredRef.TextMatrix(i, 7)
            MatCredRef(i - 1, 7) = FECredRef.TextMatrix(i, 8)
            MatCredRef(i - 1, 8) = FECredRef.TextMatrix(i, 9)
            MatCredRef(i - 1, 9) = FECredRef.TextMatrix(i, 10)
            MatCredRef(i - 1, 10) = val(Replace(FECredRef.TextMatrix(i, 4), ",", "")) + val(Replace(FECredRef.TextMatrix(i, 5), ",", "")) + val(Replace(FECredRef.TextMatrix(i, 6), ",", "")) + val(Replace(FECredRef.TextMatrix(i, 9), ",", "")) 'JOEP20171222 226-2017
        Next i
    Else
        ReDim MatCredRef(0, 0)
    End If
End Sub
