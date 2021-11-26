VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmCredAsigGastosLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Asignacion de Gastos en Lote"
   ClientHeight    =   6870
   ClientLeft      =   1320
   ClientTop       =   1365
   ClientWidth     =   8895
   Icon            =   "frmCredAsigGastosLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6870
   ScaleWidth      =   8895
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Height          =   3690
      Left            =   75
      TabIndex        =   5
      Top             =   3090
      Width           =   8745
      Begin VB.Frame Frame3 
         Height          =   570
         Left            =   90
         TabIndex        =   9
         Top             =   3045
         Width           =   2535
         Begin VB.OptionButton OptGasto 
            Caption         =   "Ninguno"
            Height          =   195
            Index           =   1
            Left            =   1275
            TabIndex        =   11
            Top             =   255
            Value           =   -1  'True
            Width           =   990
         End
         Begin VB.OptionButton OptGasto 
            Caption         =   "Todos"
            Height          =   195
            Index           =   0
            Left            =   105
            TabIndex        =   10
            Top             =   255
            Width           =   840
         End
      End
      Begin VB.CommandButton CmdAsignarGasto 
         Caption         =   "Asignar &Gasto"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   2790
         TabIndex        =   8
         Top             =   3180
         Width           =   1440
      End
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
         Height          =   390
         Left            =   7215
         TabIndex        =   7
         Top             =   3180
         Width           =   1350
      End
      Begin SICMACT.FlexEdit FECredito 
         Height          =   2865
         Left            =   105
         TabIndex        =   6
         Top             =   180
         Width           =   8520
         _extentx        =   15028
         _extenty        =   5054
         cols0           =   5
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "--Credito-Cliente-Saldo Cap"
         encabezadosanchos=   "400-400-2000-4000-1200"
         font            =   "frmCredAsigGastosLote.frx":030A
         font            =   "frmCredAsigGastosLote.frx":0336
         font            =   "frmCredAsigGastosLote.frx":0362
         font            =   "frmCredAsigGastosLote.frx":038E
         font            =   "frmCredAsigGastosLote.frx":03BA
         fontfixed       =   "frmCredAsigGastosLote.frx":03E6
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         columnasaeditar =   "X-1-X-X-X"
         listacontroles  =   "0-4-0-0-0"
         encabezadosalineacion=   "C-C-C-C-C"
         formatosedit    =   "0-0-0-0-0"
         selectionmode   =   1
         lbeditarflex    =   -1
         lbflexduplicados=   0
         lbbuscaduplicadotext=   -1
         colwidth0       =   405
         rowheight0      =   300
         forecolorfixed  =   -2147483635
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gastos"
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
      Height          =   3060
      Left            =   75
      TabIndex        =   0
      Top             =   30
      Width           =   8730
      Begin VB.CommandButton CmdAplicar 
         Caption         =   "&Aplicar"
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
         Left            =   7125
         TabIndex        =   4
         Top             =   2520
         Width           =   1455
      End
      Begin VB.ComboBox CmbTipoCred 
         Height          =   315
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2535
         Width           =   2160
      End
      Begin SICMACT.FlexEdit FEGastos 
         Height          =   2010
         Left            =   165
         TabIndex        =   1
         Top             =   255
         Width           =   8430
         _extentx        =   14870
         _extenty        =   3545
         cols0           =   11
         highlight       =   1
         allowuserresizing=   1
         encabezadosnombres=   "-Codigo-Gasto-Aplicado-Ran. Ini-Ran Fin-Tipo Valor-Valor-Moneda--"
         encabezadosanchos=   "400-1200-3000-1200-1200-1200-1200-1200-1200-0-0"
         font            =   "frmCredAsigGastosLote.frx":0414
         font            =   "frmCredAsigGastosLote.frx":0440
         font            =   "frmCredAsigGastosLote.frx":046C
         font            =   "frmCredAsigGastosLote.frx":0498
         font            =   "frmCredAsigGastosLote.frx":04C4
         fontfixed       =   "frmCredAsigGastosLote.frx":04F0
         columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X"
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-L-C-R-R-C-C-C-C-C"
         formatosedit    =   "0-0-0-0-2-2-0-0-0-0-0"
         selectionmode   =   1
         colwidth0       =   405
         rowheight0      =   300
         forecolorfixed  =   -2147483635
      End
      Begin MSComCtl2.Animation AnmBuscar 
         Height          =   675
         Left            =   5100
         TabIndex        =   12
         Top             =   2340
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   1191
         _Version        =   393216
         FullWidth       =   53
         FullHeight      =   45
      End
      Begin VB.Label lblMensaje 
         BackColor       =   &H80000005&
         Caption         =   "  Buscando ..."
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   705
         Left            =   3780
         TabIndex        =   13
         Top             =   2325
         Visible         =   0   'False
         Width           =   2130
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Producto :"
         Height          =   195
         Left            =   150
         TabIndex        =   2
         Top             =   2565
         Width           =   1320
      End
   End
End
Attribute VB_Name = "frmCredAsigGastosLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private R As ADODB.Recordset
Private nFila As Integer
Dim objPista As COMManejador.Pista


Private Sub CargaControles()
Dim oDCred As COMDCredito.DCOMCredito
Dim RTemp As ADODB.Recordset

    On Error GoTo ERRORCargaControles
    Set oDCred = New COMDCredito.DCOMCredito
    Set RTemp = oDCred.RecuperaTiposCredito
    CmbTipoCred.Clear
    Do While Not RTemp.EOF
        CmbTipoCred.AddItem RTemp!cConsDescripcion & Space(250) & RTemp!nConsValor
        RTemp.MoveNext
    Loop
    RTemp.Close
    Set RTemp = Nothing
    Set oDCred = Nothing
    Exit Sub

ERRORCargaControles:
    MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub CargaGastos()
Dim oDGasto As COMDCredito.DCOMGasto
Dim R As ADODB.Recordset

    On Error GoTo ErrorCargaGastos
    LimpiaFlex FEGastos
    Set oDGasto = New COMDCredito.DCOMGasto
    Set R = oDGasto.RecuperaGastosAplicablesCuotas(, "'MA'")
    Set oDGasto = Nothing
    Do While Not R.EOF
        FEGastos.AdicionaFila
        FEGastos.TextMatrix(R.Bookmark, 1) = Trim(str(R!nPrdConceptoCod))
        FEGastos.TextMatrix(R.Bookmark, 2) = Trim(R!cdescripcion)
        FEGastos.TextMatrix(R.Bookmark, 3) = "CUOTA"
        FEGastos.TextMatrix(R.Bookmark, 4) = Format(R!nInicial, "#0.00")
        FEGastos.TextMatrix(R.Bookmark, 5) = Format(R!nFinal, "#0.00")
        FEGastos.TextMatrix(R.Bookmark, 6) = IIf(R!nTpoValor = 1, "VALOR", "PORCENTAJE")
        FEGastos.TextMatrix(R.Bookmark, 7) = Format(R!nValor, "#0.00")
        FEGastos.TextMatrix(R.Bookmark, 8) = IIf(R!nMoneda = gMonedaNacional, "SOLES", "DOLARES")
        FEGastos.TextMatrix(R.Bookmark, 9) = Trim(str(R!nMoneda))
        FEGastos.TextMatrix(R.Bookmark, 10) = Trim(str(R!nTpoValor))
        R.MoveNext
    Loop
    R.Close
    Set R = Nothing
    
    Exit Sub

ErrorCargaGastos:
        MsgBox Err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdAplicar_Click()
Dim oDCred As COMDCredito.DCOMCredito
Dim i As Integer
    
    If Trim(FEGastos.TextMatrix(0, 1)) = "" Then
        MsgBox "No existen Gastos Aplicables a las Cuotas", vbInformation, "Aviso"
        Exit Sub
    End If
    
    LimpiaFlex FECredito
    Set R = Nothing
    Set oDCred = New COMDCredito.DCOMCredito
    
    lblMensaje.Visible = True
    Call AbrirControlAnimation(AnmBuscar, 0)
    'AnmBuscar.Open App.path & "\Videos\FINDCOMP.AVI"
    'AnmBuscar.Play
    'AnmBuscar.Visible = True
    'DoEvents
    Set R = oDCred.RecuperaCreditosParaAsignarGasto(CDbl(FEGastos.TextMatrix(FEGastos.Row, 4)), _
        CDbl(FEGastos.TextMatrix(FEGastos.Row, 5)), gColocCalendAplCuota, _
        Mid(Trim(Right(CmbTipoCred.Text, 10)), 1, 1), Trim(FEGastos.TextMatrix(FEGastos.Row, 9)), _
        Trim(FEGastos.TextMatrix(FEGastos.Row, 1)))
    Set oDCred = Nothing
    lblMensaje.Visible = False
    Call CerrarControlAnimation(AnmBuscar)
    'AnmBuscar.Stop
    'AnmBuscar.Visible = False
    
    Set FECredito.Recordset = R
    For i = 1 To FECredito.Rows - 1
        FECredito.TextMatrix(i, 0) = i
    Next i
    nFila = FEGastos.Row
    If FECredito.Rows <= 1 Then
        FECredito.Enabled = False
    Else
        FECredito.Enabled = True
    End If
    CmdAsignarGasto.Enabled = True
End Sub

Private Sub CmdAsignarGasto_Click()
Dim i As Integer
Dim oNCred As COMNCredito.NCOMCredito
Dim MatCuentas() As String

    If FECredito.Rows <= 1 Then
        MsgBox "No Existen Creditos para Asignar Gastos", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("Se va a Asignar Gastos a los Creditos, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    If Trim(FECredito.TextMatrix(1, 2)) = "" Then
        MsgBox "No Existen Creditos para Asignar Gastos", vbInformation, "Aviso"
        Exit Sub
    End If
    
    ReDim MatCuentas(0)
    For i = 1 To FECredito.Rows - 1
        If FECredito.TextMatrix(i, 1) = "." Then
            ReDim Preserve MatCuentas(UBound(MatCuentas) + 1)
            MatCuentas(UBound(MatCuentas) - 1) = FECredito.TextMatrix(i, 2)
        End If
    Next i
    
    Set oNCred = New COMNCredito.NCOMCredito
    Call oNCred.AsignarGastoLoteCreditoTotal(CDbl(FEGastos.TextMatrix(nFila, 4)), CDbl(FEGastos.TextMatrix(nFila, 5)), MatCuentas, _
                                            Trim(FEGastos.TextMatrix(nFila, 1)), CDbl(FEGastos.TextMatrix(nFila, 7)), CInt(FEGastos.TextMatrix(nFila, 10)))
    
    ''*** PEAC 20090126
    objPista.InsertarPista gsOpeCod, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gInsertar
    
    MsgBox "Se realizó la Asignación de Gastos en Lote", vbInformation, "Mensaje"
    Set oNCred = Nothing
    CmdAsignarGasto.Enabled = False
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub FEGastos_OnCellChange(pnRow As Long, pnCol As Long)
    Call cmdAplicar_Click
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
    CentraSdi Me
    Call CargaControles
    Call CargaGastos
    
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCredRegistrarGastosLote
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set objPista = Nothing
End Sub

Private Sub OptGasto_Click(Index As Integer)
Dim i As Integer
    If FECredito.Rows <= 1 Then
        Exit Sub
    End If
    If FECredito.TextMatrix(1, 2) = "" Then
        Exit Sub
    End If
    If Index = 0 Then
        For i = 1 To FECredito.Rows - 1
            FECredito.TextMatrix(i, 1) = "1"
        Next i
    Else
        For i = 1 To FECredito.Rows - 1
            FECredito.TextMatrix(i, 1) = "0"
        Next i
    End If

End Sub


