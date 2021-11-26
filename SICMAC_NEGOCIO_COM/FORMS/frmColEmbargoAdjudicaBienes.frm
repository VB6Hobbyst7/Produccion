VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmColEmbargoAdjudicaBienes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Valorizar Bienes a Adjudicar"
   ClientHeight    =   5505
   ClientLeft      =   1845
   ClientTop       =   4470
   ClientWidth     =   12090
   Icon            =   "frmColEmbargoAdjudicaBienes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   12090
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame2 
      Caption         =   "Bienes a Adjudicar"
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
      Height          =   3375
      Left            =   120
      TabIndex        =   13
      Top             =   1320
      Width           =   11895
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgBienAdjudica 
         Height          =   3015
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   11655
         _ExtentX        =   20558
         _ExtentY        =   5318
         _Version        =   393216
         Cols            =   14
         BackColorFixed  =   -2147483644
         ForeColorSel    =   10944511
         BackColorBkg    =   -2147483644
         _NumberOfBands  =   1
         _Band(0).Cols   =   14
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   120
      TabIndex        =   7
      Top             =   4680
      Width           =   11895
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   375
         Left            =   10440
         TabIndex        =   8
         Top             =   240
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Distribucion de Saldos"
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
      TabIndex        =   0
      Top             =   120
      Width           =   11895
      Begin VB.Label lblInteresMora 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   6960
         TabIndex        =   12
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblGastos 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9720
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblInteresComp 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   3840
         TabIndex        =   10
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lblCapital 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   840
         TabIndex        =   9
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label lbl2 
         Caption         =   "Gastos"
         Height          =   255
         Left            =   9000
         TabIndex        =   6
         Top             =   630
         Width           =   735
      End
      Begin VB.Label lbl3 
         Caption         =   "Interes Comp."
         Height          =   255
         Left            =   2640
         TabIndex        =   5
         Top             =   630
         Width           =   1095
      End
      Begin VB.Label lbl1 
         Caption         =   "Capital"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   615
         Width           =   735
      End
      Begin VB.Label lb4 
         Caption         =   "Interes Mora."
         Height          =   255
         Left            =   5880
         TabIndex        =   3
         Top             =   630
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Cuenta:"
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   615
      End
      Begin VB.Label lblCuenta 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Top             =   240
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmColEmbargoAdjudicaBienes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'Dim lbResult As Boolean
'Dim lsCtaCod As String
'Dim lnCapital As Currency
'Dim lnIntComp As Currency
'Dim lnIntMora As Currency
'Dim lnGastos As Currency
'Dim ldFechaAdj As Date
'Dim lfgFlex As FlexEdit
'Dim lsMovNro As String
'Dim lsMoneda As String
'Dim lnNroFilas As Integer
'Public Function Inicio(ByVal psCtaCod As String, ByVal pfgEmbargados As FlexEdit, ByVal pdFechaAdj As Date, ByVal psMovNro As String, Optional pnNroFilas As Integer = 1) As Boolean
'     lsCtaCod = psCtaCod
'     ldFechaAdj = pdFechaAdj
'     lsMovNro = psMovNro
'     lnNroFilas = pnNroFilas
'     Set lfgFlex = pfgEmbargados
'     If Not cargarDistribucionSaldo Then
'        lbResult = False
'     Else
'        Me.Show 1
'     End If
'
'     Inicio = lbResult
'End Function
'
'Private Sub cargarFlex(ByVal pnFlex As FlexEdit)
'    Dim i As Integer
'    Dim nItem As Integer
'
'
'    If Mid(Me.lblCuenta, 9, 1) = "1" Then
'        lsMoneda = "NACIONAL" + Space(50) + "1"
'    ElseIf Mid(Me.lblCuenta, 9, 1) = "2" Then
'        lsMoneda = "EXTRANJERA" + Space(50) + "2"
'    End If
'    nItem = 0
'    Me.fgBienAdjudica.Rows = lnNroFilas + 1 'pnFlex.Rows
'
'    For i = 1 To pnFlex.Rows - 1
'        If pnFlex.TextMatrix(i, 1) = "." Then
'            If Right(pnFlex.TextMatrix(i, 2), 1) = "1" Then
'                nItem = nItem + 1
'                'Me.fgBienAdjudica.AdicionaFila
'                Me.fgBienAdjudica.TextMatrix(nItem, 0) = nItem
'                Me.fgBienAdjudica.TextMatrix(nItem, 1) = i
'                Me.fgBienAdjudica.TextMatrix(nItem, 2) = Trim(Right(pnFlex.TextMatrix(i, 4), 5))
'                Me.fgBienAdjudica.TextMatrix(nItem, 3) = Trim(Right(pnFlex.TextMatrix(i, 5), 5))
'                Me.fgBienAdjudica.TextMatrix(nItem, 4) = pnFlex.TextMatrix(i, 6)
'                Me.fgBienAdjudica.TextMatrix(nItem, 5) = pnFlex.TextMatrix(i, 8)
'                Me.fgBienAdjudica.TextMatrix(nItem, 6) = "0.00"
'                Me.fgBienAdjudica.TextMatrix(nItem, 7) = "0.00"
'                Me.fgBienAdjudica.TextMatrix(nItem, 8) = "0.00"
'                Me.fgBienAdjudica.TextMatrix(nItem, 9) = "0.00"
'                Me.fgBienAdjudica.TextMatrix(nItem, 10) = lsMoneda 'pnFlex.TextMatrix(i, 20) 'Moneda
'                Me.fgBienAdjudica.TextMatrix(nItem, 11) = pnFlex.TextMatrix(i, 17) 'Partida
'                Me.fgBienAdjudica.TextMatrix(nItem, 12) = IIf(pnFlex.TextMatrix(i, 18) = "", "0.00", pnFlex.TextMatrix(i, 18)) 'Tasacion
'                Me.fgBienAdjudica.TextMatrix(nItem, 13) = pnFlex.TextMatrix(i, 19) 'Fec. Tasacion
'              End If
'        End If
'    Next i
'End Sub
'Private Function valida() As Boolean
'    valida = True
'    Dim i As Integer
'    With Me.fgBienAdjudica
'        For i = 1 To .Rows - 1
'            If .TextMatrix(i, 6) = "" Then
'              MsgBox "Debe Ingresar el valor de Adjudicacion", vbInformation, "AVISO"
'              .Col = 6
'              .SetFocus
'              valida = False
'              Exit Function
'            End If
'            If .TextMatrix(i, 7) = "" Then
'              MsgBox "Debe Ingresar el Capital ", vbInformation, "AVISO"
'              .Col = 7
'              .SetFocus
'              valida = False
'              Exit Function
'            End If
'            If .TextMatrix(i, 8) = "" Then
'              MsgBox "Debe Ingresar el Interes", vbInformation, "AVISO"
'              .Col = 8
'              .SetFocus
'              valida = False
'              Exit Function
'            End If
'            If .TextMatrix(i, 9) = "" Then
'              MsgBox "Debe Ingresar el Gasto", vbInformation, "AVISO"
'              .Col = 9
'              .SetFocus
'              valida = False
'              Exit Function
'            End If
'            If .TextMatrix(i, 10) = "" Then
'              MsgBox "Debe Seleccionar la Moneda a valorar", vbInformation, "AVISO"
'              .Col = 10
'              .SetFocus
'              valida = False
'              Exit Function
'            End If
'        Next i
'    End With
'End Function
'
'Private Sub cmdAceptar_Click()
''    If Not valida Then
''        Exit Sub
''    End If
'    If MsgBox("Se va ha proceder a Adjudicar los bienes", vbYesNo, "Ajudicar Bienes") = vbYes Then
'        Dim oColRec As COMNColocRec.NCOMColRecCredito
'        Dim oCnt As COMNContabilidad.NCOMContFunciones
'        Dim i As Integer
'        Dim nMontoTotal As Currency
'        Dim sTipoBien As String
'        Dim dFecTasacion As Date
'        Set oCnt = New COMNContabilidad.NCOMContFunciones
'        Set oColRec = New COMNColocRec.NCOMColRecCredito
'        nMontoTotal = 0#
'        With Me.fgBienAdjudica
'            For i = 1 To Me.fgBienAdjudica.Rows - 1
'
'                If Left(Trim(Right(.TextMatrix(i, 3), 5)), 2) <> "94" Then
'                       sTipoBien = getCodTipoBien(Mid(Trim(Right(.TextMatrix(i, 3), 5)), 1, 2))
'                Else
'                        sTipoBien = getCodTipoBien(Trim(Right(.TextMatrix(i, 3), 5)))
'                End If
'                If .TextMatrix(i, 12) = "0.00" Or .TextMatrix(i, 12) = "0" Or .TextMatrix(i, 12) = "" Then
'                    dFecTasacion = ldFechaAdj
'                Else
'                    dFecTasacion = CDate(.TextMatrix(i, 13))
'                 End If
'
'                oCnt.InsertarBienAdjudicado Me.lblCuenta, "", sTipoBien, _
'                            .TextMatrix(i, 5), "1", CDbl(.TextMatrix(i, 6)), CDbl(.TextMatrix(i, 12)), CDbl(.TextMatrix(i, 12)), ldFechaAdj, _
'                           dFecTasacion, CInt(Right(.TextMatrix(i, 10), 1)), CDbl(.TextMatrix(i, 7)), CDbl(.TextMatrix(i, 8)), CDbl(.TextMatrix(i, 9)), 1, lsMovNro
'
'                nMontoTotal = nMontoTotal + CCur(.TextMatrix(.row, 6))
'            Next i
'            oColRec.guardarPagoCredAdjudicacion "136303", Me.lblCuenta, nMontoTotal, gdFecSis, Right(lsMoneda, 1), lsMovNro
'
'        End With
'        lbResult = True
'        Unload Me
'    End If
'End Sub
'
'Private Sub fgBienAdjudica_EnterCell()
'    With fgBienAdjudica
'        If (.Col > 6 And .Col < 10) Then
'            .CellBackColor = &H80000018
'            .Tag = ""
'        End If
'    End With
'End Sub
'
'Private Sub fgBienAdjudica_KeyDown(KeyCode As Integer, Shift As Integer)
'    With Me.fgBienAdjudica
'     Select Case KeyCode
'        Case 46
'            .Tag = fgBienAdjudica
'            If (.Col > 6 And .Col < 10) Then
'                 .TextMatrix(.row, .Col) = ""
'                 fgBienAdjudica = ""
'            End If
'
'
'      End Select
'    End With
'End Sub
'
'Private Sub fgBienAdjudica_KeyPress(KeyAscii As Integer)
'    With Me.fgBienAdjudica
'        If .Col > 6 And .Col < 10 Then
'            If KeyAscii = 13 Then
'                If .TextMatrix(.row, .Col) = "" Then
'                    .TextMatrix(.row, .Col) = "0.00"
'                End If
'                If Left(.TextMatrix(.row, .Col), 1) = "." Then
'                    .TextMatrix(.row, .Col) = Right(.TextMatrix(.row, .Col), Len(.TextMatrix(.row, .Col)) - 1)
'                End If
'
'                If Right(.TextMatrix(.row, .Col), 1) = "." Then
'                    .TextMatrix(.row, .Col) = Left(.TextMatrix(.row, .Col), Len(.TextMatrix(.row, .Col)) - 1)
'                End If
'
'                .TextMatrix(.row, .Col) = Format(.TextMatrix(.row, .Col), "##,##0.00")
'                .TextMatrix(.row, 6) = Format(CCur(IIf(.TextMatrix(.row, 7) = "", "0.00", .TextMatrix(.row, 7))) + CCur(IIf(.TextMatrix(.row, 8) = "", "0.00", .TextMatrix(.row, 8))) + CCur(IIf(.TextMatrix(.row, 9) = "", "0.00", .TextMatrix(.row, 9))), "##,##0.00")
'
'                If .Col = 9 And .row = .Rows - 1 Then
'                    .CellBackColor = &H8000000E
'                    Me.cmdAceptar.SetFocus
'                ElseIf .Col = 9 And .row < .Rows - 1 Then
'                   .CellBackColor = &H8000000E
'                    .row = .row + 1
'                    .Col = 7
'                    .CellBackColor = &H80000018
'                Else
'                   .CellBackColor = &H8000000E
'                   .Col = .Col + 1
'                    .CellBackColor = &H80000018
'                End If
'
'            ElseIf KeyAscii = 8 Then
'                If Len(.TextMatrix(.row, .Col)) > 0 Then
'
'                            .TextMatrix(.row, .Col) = Mid(.TextMatrix(.row, .Col), 1, Len(.TextMatrix(.row, .Col)) - 1)
'                            'fgBienAdjudica = Format(.TextMatrix(.Row, .Col), "##,##0")
'                            fgBienAdjudica = .TextMatrix(.row, .Col)
'
'
'                End If
'            Else
'                    If (InStr("0123456789.", Chr(KeyAscii)) = 0) Then 'compara si es un numero
'                            KeyAscii = 0
'                            .TextMatrix(.row, .Col) = ""
'                    Else 'entra si es un numero
'                            .TextMatrix(.row, .Col) = .TextMatrix(.row, .Col) + Chr(KeyAscii)
'                            'fgBienAdjudica = Format(.TextMatrix(.Row, .Col), "##,##0.00")
'                            fgBienAdjudica = .TextMatrix(.row, .Col)
'
'                    End If
'
'
'
'
'            End If
'        End If
'    End With
'End Sub
'
'Private Sub fgBienAdjudica_LeaveCell()
'    With Me.fgBienAdjudica
'        If .Col > 6 And .Col < 10 Then
'            .CellBackColor = &H8000000E
'            If .TextMatrix(.row, .Col) = "" Then
'                .TextMatrix(.row, .Col) = "0.00"
'            End If
'        End If
'    End With
'End Sub
'
'Private Sub Form_Load()
'    lbResult = False
'    Me.lblCuenta = lsCtaCod
'    Me.lblCapital = Format(lnCapital, "##,##0.00")
'    Me.lblInteresComp = Format(lnIntComp, "##,##0.00")
'    Me.lblInteresMora = Format(lnIntMora, "##,##0.00")
'    Me.lblGastos = Format(lnGastos, "##,##0.00")
'    ConfigFlexCabecera
'    cargarFlex lfgFlex
'
'
'End Sub
'
'Private Function cargarDistribucionSaldo() As Boolean
'    Dim oColRec As COMNColocRec.NCOMColRecCredito
'    Dim rsDistribucion As Recordset
'    Set oColRec = New COMNColocRec.NCOMColRecCredito
'    Set rsDistribucion = New Recordset
'    cargarDistribucionSaldo = True
'    Set rsDistribucion = oColRec.obtenerDistribucionSaldos(lsCtaCod)
'    If Not (rsDistribucion.EOF And rsDistribucion.BOF) Then
'        Me.lblCapital = rsDistribucion!Capital
'        Me.lblInteresComp = rsDistribucion!InteresComp
'        Me.lblInteresMora = rsDistribucion!InteresMora
'        Me.lblGastos = rsDistribucion!Gastos
'        lnCapital = rsDistribucion!Capital
'        lnIntComp = rsDistribucion!InteresComp
'        lnIntMora = rsDistribucion!InteresMora
'        lnGastos = rsDistribucion!Gastos
'
'    Else
'        MsgBox "Aun no se ha realizado la Distribucion de Saldos de la cuenta:" + Me.lblCuenta, vbExclamation, "AVISO"
'        cargarDistribucionSaldo = False
'        'Unload Me
'    End If
'End Function
'Private Function getCodTipoBien(ByVal psCod As String) As String
'
'    Select Case psCod
'        Case "91" 'Mobiliario y Equipo
'            getCodTipoBien = "B03" 'Maquinaria y Equipos de uso generico
'
'        Case "92" 'Maquinarias y Motores
'            getCodTipoBien = "B04" 'Maquinaria y Equipos Especializado
'
'        Case "93" 'Unidades de Transporte
'            getCodTipoBien = "B07" 'Vehiculos
'
'        Case "9401" 'Inmubeles-->Terreno
'            getCodTipoBien = "A03" 'Terreno Urbano
'
'        Case "9402" 'Inmubeles-->Construccion
'            getCodTipoBien = "A01" 'Vivienda Urbano
'
'    End Select
'End Function
'Private Sub ConfigFlexCabecera()
'
'
'    With Me.fgBienAdjudica
'        '#-NroBien-TipoBien-SubTipo-Bien-Descripcion-Valor Adju.-Capital-Interes-Gastos-Moneda-Partida-Valor Tasa-Fec_Tasac
'        .TextMatrix(0, 0) = "#"
'        .TextMatrix(0, 1) = "NroBien"
'        .TextMatrix(0, 2) = "TipoBien"
'        .TextMatrix(0, 3) = "SubTipo"
'        .TextMatrix(0, 4) = "Bien"
'        .TextMatrix(0, 5) = "Descripcion"
'        .TextMatrix(0, 6) = "Valor Adju."
'        .TextMatrix(0, 7) = "Capital"
'        .TextMatrix(0, 8) = "Interes"
'        .TextMatrix(0, 9) = "Gastos"
'        .TextMatrix(0, 10) = "Moneda"
'        .TextMatrix(0, 11) = "Partida"
'        .TextMatrix(0, 12) = "Valor Tasa"
'        .TextMatrix(0, 13) = "Fec_Tasac"
'
'
'         'Ancho de las Columnas
'         '400-0-0-0-2000-3000-1200-1200-1200-1000-1200-0-0-0
'        .ColWidth(0) = 400
'        .ColWidth(1) = 0 'nMovNro
'        .ColWidth(2) = 0 'Fecha
'        .ColWidth(3) = 0 'Nro_Credito
'        .ColWidth(4) = 2000 'Cliente
'        .ColWidth(5) = 3000 'Nro_Expediente
'        .ColWidth(6) = 1200 'Nro_Resolucion
'        .ColWidth(7) = 1200 'Moneda
'        .ColWidth(8) = 1200 'Capital
'        .ColWidth(9) = 1000 'Interes
'        .ColWidth(10) = 1200 'Gastos
'        .ColWidth(11) = 0 'Bienes_Embargados
'        .ColWidth(12) = 0 'Bienes_Embargados
'        .ColWidth(13) = 0 'Bienes_Embargados
'
'
'        'Alineacion del contenido en las Celdas
'        .ColAlignment(0) = flexAlignLeftCenter
'        .ColAlignment(1) = flexAlignLeftCenter
'        .ColAlignment(2) = flexAlignLeftCenter
'        .ColAlignment(3) = flexAlignCenterCenter
'        .ColAlignment(4) = flexAlignLeftCenter
'        .ColAlignment(5) = flexAlignLeftCenter
'        .ColAlignment(6) = flexAlignRightCenter
'         .ColAlignment(7) = flexAlignRightCenter
'        .ColAlignment(8) = flexAlignRightCenter
'        .ColAlignment(9) = flexAlignRightCenter
'        .ColAlignment(10) = flexAlignLeftCenter
'       .ColAlignment(11) = flexAlignCenterCenter
'       .ColAlignment(12) = flexAlignCenterCenter
'       .ColAlignment(13) = flexAlignCenterCenter
'
'
'
'    End With
'End Sub
'
