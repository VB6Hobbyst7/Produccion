VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCredTransfRecupeGarant 
   Caption         =   "Garantia a Recuperaciones"
   ClientHeight    =   5850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   8535
   Begin VB.Frame Frame5 
      Height          =   825
      Left            =   120
      TabIndex        =   11
      Top             =   4800
      Width           =   8250
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   1800
         TabIndex        =   13
         Top             =   255
         Width           =   1500
      End
      Begin VB.CommandButton CmdTransferir 
         Caption         =   "&Transferir"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   180
         TabIndex        =   12
         Top             =   255
         Width           =   1500
      End
      Begin MSComctlLib.ProgressBar PBBarra 
         Height          =   285
         Left            =   3600
         TabIndex        =   14
         Top             =   360
         Visible         =   0   'False
         Width           =   4560
         _ExtentX        =   8043
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
   End
   Begin VB.Frame FraCta 
      Height          =   735
      Left            =   3960
      TabIndex        =   8
      Top             =   360
      Width           =   3975
      Begin SICMACT.ActXCodCta ActxCta 
         Height          =   375
         Left            =   360
         TabIndex        =   9
         Top             =   120
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   661
         Texto           =   "Crédito"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3615
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   8175
      Begin VB.CommandButton CmdPasarARec 
         Caption         =   "A Recuperaciones"
         Height          =   465
         Left            =   6360
         TabIndex        =   2
         Top             =   3000
         Width           =   1650
      End
      Begin VB.Frame Frame3 
         Height          =   630
         Left            =   240
         TabIndex        =   15
         Top             =   2880
         Width           =   2805
         Begin VB.OptionButton OptSelecc 
            Caption         =   "&Ninguno"
            Height          =   240
            Index           =   1
            Left            =   1680
            TabIndex        =   3
            Top             =   240
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton OptSelecc 
            Caption         =   "&Todos"
            Height          =   240
            Index           =   0
            Left            =   135
            TabIndex        =   16
            Top             =   240
            Width           =   960
         End
      End
      Begin SICMACT.FlexEdit FECreditos 
         Height          =   2655
         Left            =   240
         TabIndex        =   10
         Top             =   240
         Width           =   7815
         _ExtentX        =   13785
         _ExtentY        =   4683
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "-Tran-CodPer-Nombre-cCtaCod-Nro Garantia-DesTipo-DescGaran-Direccion-Estado"
         EncabezadosAnchos=   "400-400-1600-4200-2000-1200-2500-4200-3200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame OptBusqueda 
      Caption         =   "Busqueda"
      Height          =   1095
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8055
      Begin VB.Frame FraBusqNom 
         Height          =   735
         Left            =   3720
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   3975
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "&Buscar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   2160
            TabIndex        =   18
            Top             =   120
            Width           =   1380
         End
      End
      Begin VB.Frame frmTipoBusqueda 
         Height          =   735
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3495
         Begin VB.OptionButton OptionBusqueda 
            Caption         =   "General"
            Height          =   255
            Index           =   2
            Left            =   2520
            TabIndex        =   7
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton OptionBusqueda 
            Caption         =   "Por Nombre"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   6
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OptionBusqueda 
            Caption         =   "Por Cuenta"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   5
            Top             =   240
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmCredTransfRecupeGarant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBuscar_Click()
Dim oPers As COMDPersona.UCOMPersona 'UPersona
    If OptionBusqueda(1).value Then
        Set oPers = frmBuscaPersona.Inicio
        If Not oPers Is Nothing Then
            Call CargaDatos(oPers.sPersCod, 2)
        End If
    Else
        Call CargaDatos("", 1)
    End If
    'cmdImprimirActa.Enabled = False
End Sub
Private Sub CmdPasarARec_Click()
Dim nCol As Integer
    nCol = FECreditos.Col
    If Trim(FECreditos.TextMatrix(FECreditos.Row, 1)) <> "" Then
        FECreditos.Col = 1
        If FECreditos.CellBackColor = vbGreen Then
            FECreditos.CellBackColor = vbWhite
            FECreditos.TextMatrix(FECreditos.Row, 1) = "."
            CmdPasarARec.Caption = "A Recuperaciones"
        Else
            FECreditos.CellBackColor = vbGreen
            FECreditos.TextMatrix(FECreditos.Row, 1) = "R"
            CmdPasarARec.Caption = "A Recuperaciones"
        End If
    End If
    FECreditos.Col = nCol
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub CmdTransferir_Click()
Dim i As Integer
Dim nCol As Integer
Dim oNCred As COMNCredito.NCOMGarantia
Dim nMaxBarra As Integer
Dim nContCred As Integer
Dim lsmensaje As String

Dim rs As ADODB.Recordset

    If MsgBox("Garantias seran Transferidas a Recuperaciones, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    If Trim(FECreditos.TextMatrix(1, 1)) = "" Then
        MsgBox "No Se Encontraron Registros para Transferir", vbInformation, "Aviso"
        Exit Sub
    End If
    Screen.MousePointer = 11
    PBBarra.Visible = True
    nCol = FECreditos.Col
    FECreditos.Col = 1
    nContCred = 0
    For i = 1 To FECreditos.Rows - 1
        FECreditos.Row = i
        If FECreditos.CellBackColor = vbGreen Then
            nContCred = nContCred + 1
        End If
    Next i
    nMaxBarra = nContCred
    nContCred = 0
    
    'Mandar los Datos a un Recordset para Grabar en Lote
    Set rs = New ADODB.Recordset

    With rs
        'Crear RecordSet
        .Fields.Append "cGarantia", adVarChar, 8
        .Fields.Append "dEstadoAdju", adDate
        .Fields.Append "nEstadoAdju", adInteger
        .Open
        'Llenar Recordset
    
    For i = 1 To FECreditos.Rows - 1
        FECreditos.Row = i
        If FECreditos.CellBackColor = vbGreen Then
            nContCred = nContCred + 1
            .AddNew
            .Fields("cGarantia") = FECreditos.TextMatrix(i, 5)
            .Fields("dEstadoAdju") = gdFecSis
            .Fields("nEstadoAdju") = 7
            FECreditos.CellBackColor = vbRed
            PBBarra.value = (nContCred / nMaxBarra) * 100
        End If
    Next i
    End With
    
    Set oNCred = New COMNCredito.NCOMGarantia
    If Not (rs.EOF And rs.BOF) Then
        Call oNCred.TransferirGarantiasARecuperacionesLote(rs, gsCodUser, gsCodAge, lsmensaje)
        If Trim(lsmensaje) <> "" Then
            MsgBox lsmensaje, vbInformation, "Aviso"
            Exit Sub
        End If
         MsgBox "Datos se guardaron correctamente", vbInformation, "Aviso"
    Else
        MsgBox "Seleccione un Crédito", vbInformation, "Aviso"
        Screen.MousePointer = 0
        Exit Sub
    End If
                
    Set oNCred = Nothing
    FECreditos.Col = nCol
    PBBarra.Visible = False
    CmdPasarARec.Enabled = False
    CmdTransferir.Enabled = False
    'cmdImprimirActa.Enabled = True
            
    Screen.MousePointer = 0
    
    Call Impresion
    
End Sub
Private Sub Impresion()
Dim oPrev As previo.clsprevio
Dim sCad As String
Dim i As Integer

Dim loImpre As COMNColocRec.NCOMColRecImpre
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset

   With rs
        'Crear RecordSet
''        .Fields.Append "cTran", adVarChar, 2
''        .Fields.Append "cCtaCod", adVarChar, 20
''        .Fields.Append "cCond", adVarChar, 2
''        .Fields.Append "cCliente", adVarChar, 100
''        .Fields.Append "nMonto", adCurrency
''        .Fields.Append "nSaldo", adCurrency
''        .Fields.Append "cEstado", adVarChar, 2
        .Fields.Append "cTran", adVarChar, 2
        .Fields.Append "cPersCod", adVarChar, 13
        .Fields.Append "cCliente", adVarChar, 100
        .Fields.Append "cCtaCod", adVarChar, 20
        .Fields.Append "cNumGarant", adVarChar, 8
        .Fields.Append "cDesTipo", adVarChar, 100
        .Fields.Append "cDireccion", adVarChar, 100
        .Open
        'Llenar Recordset
        For i = 1 To FECreditos.Rows - 1
            FECreditos.Row = i
            FECreditos.Col = 1
            If FECreditos.CellBackColor = vbRed Then
                .AddNew
''                .Fields("cTran") = IIf(IsNull(FECreditos.TextMatrix(i, 1)), "", FECreditos.TextMatrix(i, 1))
''                .Fields("cCtaCod") = IIf(IsNull(FECreditos.TextMatrix(i, 2)), "", FECreditos.TextMatrix(i, 2))
''                .Fields("cCond") = IIf(IsNull(FECreditos.TextMatrix(i, 3)), "", FECreditos.TextMatrix(i, 3))
''                .Fields("cCliente") = IIf(IsNull(FECreditos.TextMatrix(i, 4)), "", FECreditos.TextMatrix(i, 4))
''                .Fields("nMonto") = IIf(IsNull(FECreditos.TextMatrix(i, 5)), 0, FECreditos.TextMatrix(i, 5))
''                .Fields("nSaldo") = IIf(IsNull(FECreditos.TextMatrix(i, 6)), 0, FECreditos.TextMatrix(i, 6))
''                .Fields("cEstado") = IIf(IsNull(FECreditos.TextMatrix(i, 7)), 0, FECreditos.TextMatrix(i, 7))
                .Fields("cTran") = IIf(IsNull(FECreditos.TextMatrix(i, 1)), "", FECreditos.TextMatrix(i, 1))
                .Fields("cPersCod") = IIf(IsNull(FECreditos.TextMatrix(i, 2)), "", FECreditos.TextMatrix(i, 2))
                .Fields("cCliente") = IIf(IsNull(FECreditos.TextMatrix(i, 3)), "", FECreditos.TextMatrix(i, 3))
                .Fields("cCtaCod") = IIf(IsNull(FECreditos.TextMatrix(i, 4)), "", FECreditos.TextMatrix(i, 4))
                .Fields("cNumGarant") = IIf(IsNull(FECreditos.TextMatrix(i, 5)), "", FECreditos.TextMatrix(i, 5))
                .Fields("cDesTipo") = IIf(IsNull(FECreditos.TextMatrix(i, 6)), "", FECreditos.TextMatrix(i, 6))
                .Fields("cDireccion") = IIf(IsNull(FECreditos.TextMatrix(i, 8)), 0, FECreditos.TextMatrix(i, 8))
            End If
        Next i
    End With
    
    Set loImpre = New COMNColocRec.NCOMColRecImpre
        sCad = loImpre.ImpresionTransferenciaAdjudicados(rs, gsNomCmac, gdFecSis, gsNomAge, gsCodUser)
    Set loImpre = Nothing
    
    rs.Close
    
    Set oPrev = New previo.clsprevio
    oPrev.Show sCad, "Transferencia A Recuperaciones"
    Set oPrev = Nothing
    
    
End Sub
Private Sub FECreditos_OnChangeCombo()
Dim nCol As Integer
    If Trim(FECreditos.TextMatrix(1, 1)) = "" Then
        Exit Sub
    End If
    nCol = FECreditos.Col
    FECreditos.Col = 1
    If FECreditos.CellBackColor = vbGreen Then
        CmdPasarARec.Caption = "No Pasar A Recupe."
    Else
        CmdPasarARec.Caption = "A Recuperaciones"
    End If
    
    
    
    FECreditos.Col = nCol
End Sub





Private Sub OptionBusqueda_Click(Index As Integer)
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
    If Index = 0 Then
        FraCta.Visible = True
        FraBusqNom.Visible = False
    Else
        FraCta.Visible = False
        FraBusqNom.Visible = True
    End If
End Sub
Private Sub CargaDatos(ByVal psBusqueda As String, ByVal pnTipoBusq As Integer)
Dim oCredito As COMDCredito.DCOMCredito
Dim oNCredito As COMNCredito.NCOMCredito
Dim R As ADODB.Recordset
Dim MatCalend As Variant

    On Error GoTo ErrorCargaDatos
    Set oCredito = New COMDCredito.DCOMCredito
    Select Case pnTipoBusq
        Case 1 'Total
            Set R = oCredito.RecuperaCreditosParaAdjudicados
        Case 2 'Por Nombre
            Set R = oCredito.RecuperaCreditosParaAdjudicados(psBusqueda)
        Case 3 'Por Cuenta
            Set R = oCredito.RecuperaCreditosParaAdjudicados(, psBusqueda)
    End Select
    LimpiaFlex FECreditos
    Set oCredito = Nothing
    If R.BOF And R.EOF Then
        MsgBox "No se Encontraron Registros", vbInformation, "Aviso"
        R.Close
        Set R = Nothing
        Exit Sub
    End If
    Do While Not R.EOF
        FECreditos.AdicionaFila
        FECreditos.TextMatrix(R.Bookmark, 1) = "."
        FECreditos.TextMatrix(R.Bookmark, 2) = R!cPersCod
        FECreditos.TextMatrix(R.Bookmark, 3) = PstaNombre(R!cPers)
        FECreditos.TextMatrix(R.Bookmark, 4) = R!cCtaCod
        FECreditos.TextMatrix(R.Bookmark, 5) = R!cNumGarant
        FECreditos.TextMatrix(R.Bookmark, 6) = R!cDesTipoGarantia
        FECreditos.TextMatrix(R.Bookmark, 7) = R!cDescripcion
        FECreditos.TextMatrix(R.Bookmark, 8) = R!cDireccion
        FECreditos.TextMatrix(R.Bookmark, 9) = R!cDesEstado
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    Exit Sub

ErrorCargaDatos:
    MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub OptSelecc_Click(Index As Integer)
Dim i As Integer
Dim nCol As Integer
    nCol = FECreditos.Col
    If FECreditos.TextMatrix(1, 1) <> "" Then
        If Index = 0 Then
            For i = 1 To FECreditos.Rows - 1
                FECreditos.Col = 1
                FECreditos.Row = i
                If FECreditos.CellBackColor <> vbRed Then
                    FECreditos.CellBackColor = vbGreen
                    CmdPasarARec.Caption = "No Pasar A Recupe."
                End If
            Next i
        Else
            For i = 1 To FECreditos.Rows - 1
                FECreditos.Col = 1
                FECreditos.Row = i
                If FECreditos.CellBackColor <> vbRed Then
                    FECreditos.CellBackColor = vbWhite
                    CmdPasarARec.Caption = "A Recuperaciones"
                End If
            Next i
            
        End If
    End If
    FECreditos.Col = nCol
End Sub
