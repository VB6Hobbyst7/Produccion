VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmRHPrestamosAdmOtros 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creditos en Otras Entidades"
   ClientHeight    =   6720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9195
   Icon            =   "frmRHPrestamosAdmOtros.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6720
   ScaleWidth      =   9195
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   6150
      TabIndex        =   15
      Top             =   6285
      Width           =   975
   End
   Begin VB.Frame fraPlanilla 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Planilla"
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   60
      TabIndex        =   4
      Top             =   30
      Width           =   7065
      Begin VB.CommandButton cmdBuscaPla 
         Height          =   375
         Left            =   3645
         Picture         =   "frmRHPrestamosAdmOtros.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   218
         Width           =   375
      End
      Begin Sicmact.TxtBuscar txtPlanillas 
         Height          =   300
         Left            =   90
         TabIndex        =   5
         Top             =   270
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   529
         Appearance      =   0
         BackColor       =   12648447
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Appearance      =   0
         sTitulo         =   ""
      End
      Begin MSComCtl2.DTPicker txtFecPla 
         Height          =   315
         Left            =   2085
         TabIndex        =   7
         Top             =   248
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   64552961
         CurrentDate     =   36963
      End
      Begin VB.Label lblFecha 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   1515
         TabIndex        =   9
         Top             =   308
         Width           =   495
      End
      Begin VB.Label lblPlanillaG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   90
         TabIndex        =   8
         Top             =   630
         Width           =   6780
      End
   End
   Begin VB.Frame fraCuentas 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Cuentas"
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
      Height          =   5100
      Left            =   45
      TabIndex        =   3
      Top             =   1125
      Width           =   9105
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   345
         Left            =   1050
         TabIndex        =   19
         Top             =   4605
         Width           =   870
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   345
         Left            =   120
         TabIndex        =   18
         Top             =   4605
         Width           =   870
      End
      Begin Sicmact.FlexEdit Flex 
         Height          =   4260
         Left            =   75
         TabIndex        =   17
         Top             =   240
         Width           =   8985
         _ExtentX        =   15849
         _ExtentY        =   7514
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Persona-Nombre-Moneda-Institut-Monto (S/.)-Monto (US$)-TipoCambio"
         EncabezadosAnchos=   "300-1300-2600-1000-1200-1000-1000-01"
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
         ColumnasAEditar =   "X-1-X-3-4-5-6-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-1-0-3-3-0-0-0"
         BackColor       =   -2147483639
         EncabezadosAlineacion=   "C-L-L-L-L-R-R-C"
         FormatosEdit    =   "0-0-0-0-0-2-2-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   7
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   300
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
         CellBackColor   =   -2147483639
      End
      Begin VB.Label lblTotDol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFC0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   7500
         TabIndex        =   22
         Top             =   4605
         Width           =   1005
      End
      Begin VB.Label lblTotSol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0.00"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6510
         TabIndex        =   21
         Top             =   4605
         Width           =   1005
      End
      Begin VB.Label Label4 
         Appearance      =   0  'Flat
         BackColor       =   &H00E0E0E0&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Total"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Left            =   5250
         TabIndex        =   20
         Top             =   4605
         Width           =   3255
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8175
      TabIndex        =   2
      Top             =   6285
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   45
      TabIndex        =   1
      Top             =   6285
      Width           =   975
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   7170
      TabIndex        =   0
      Top             =   6285
      Width           =   975
   End
   Begin RichTextLib.RichTextBox R 
      Height          =   180
      Left            =   1245
      TabIndex        =   16
      Top             =   6420
      Visible         =   0   'False
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   318
      _Version        =   393217
      Appearance      =   0
      TextRTF         =   $"frmRHPrestamosAdmOtros.frx":040C
   End
   Begin VB.Frame fraTC 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Tipo Cambio"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1095
      Left            =   7110
      TabIndex        =   10
      Top             =   30
      Width           =   2040
      Begin VB.Label lblTCF 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   720
         TabIndex        =   14
         Top             =   600
         Width           =   1080
      End
      Begin VB.Label lblTCFL 
         AutoSize        =   -1  'True
         Caption         =   "TCF:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   630
         Width           =   420
      End
      Begin VB.Label lblTCC 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   315
         Left            =   720
         TabIndex        =   12
         Top             =   240
         Width           =   1080
      End
      Begin VB.Label lblTCCL 
         AutoSize        =   -1  'True
         Caption         =   "TCC:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   270
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmRHPrestamosADMOtros"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdBuscaPla_Click()
    If Me.txtPlanillas.Text = "" Then
        MsgBox "Debe elegir una planilla.", vbInformation, "Aviso"
        Me.txtPlanillas.SetFocus
        Exit Sub
    End If

    BuscaDatosPlanilla
    fraPlanilla.Enabled = False
    Suma
End Sub

Private Sub cmdCancelar_Click()
    fraPlanilla.Enabled = True
    Flex.Clear
    Flex.Rows = 2
    Flex.FormaCabecera
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea Eliminar el registro " & Me.Flex.TextMatrix(Me.Flex.Row, 1) & " ? ", vbQuestion + vbYesNo + vbDefaultButton2, "Aviso") = vbNo Then Exit Sub
    Flex.EliminaFila Me.Flex.Row
    Flex.SetFocus
    Me.cmdGrabar.Enabled = True
End Sub

Private Sub cmdGrabar_Click()
    Dim L As ListItem
    Dim sPla As String, sFec As String, sFecAct As String
    Dim sEmp As String, sCta As String, sPers As String
    Dim nMonto As Double
    Dim sTCC As String, sCuota As String, sFecVenc As String
    Dim sqlEmpCon As String
    Dim rsEmpCon As ADODB.Recordset
    Set rsEmpCon = New ADODB.Recordset
    Dim oCred As DRHPrestamosAdm
    Set oCred = New DRHPrestamosAdm
    Dim oMov As DMov
    Set oMov = New DMov
    Dim lsMovNro As String
    Dim lnI As Integer
    
    For lnI = 1 To Me.Flex.Rows - 1
        If Me.Flex.TextMatrix(lnI, 1) = "" Then
            MsgBox "Debe ingresar un empleado valido.", vbInformation, "Aviso"
            Flex.Col = 1
            Flex.Row = lnI
            Me.Flex.SetFocus
            Exit Sub
        ElseIf Me.Flex.TextMatrix(lnI, 3) = "" Then
            MsgBox "Debe ingresar una moneda valido.", vbInformation, "Aviso"
            Flex.Col = 3
            Flex.Row = lnI
            Me.Flex.SetFocus
            Exit Sub
        ElseIf Me.Flex.TextMatrix(lnI, 4) = "" Then
            MsgBox "Debe ingresar una institucion valida.", vbInformation, "Aviso"
            Flex.Col = 4
            Flex.Row = lnI
            Me.Flex.SetFocus
            Exit Sub
        ElseIf Right(Me.Flex.TextMatrix(lnI, 3), 3) = Moneda.gMonedaNacional And Not IsNumeric(Me.Flex.TextMatrix(lnI, 5)) Then
            MsgBox "Debe ingresar un monto en soles valido.", vbInformation, "Aviso"
            Flex.Col = 5
            Flex.Row = lnI
            Me.Flex.SetFocus
            Exit Sub
        ElseIf Right(Me.Flex.TextMatrix(lnI, 3), 3) = Moneda.gMonedaExtranjera And Not IsNumeric(Me.Flex.TextMatrix(lnI, 6)) Then
            MsgBox "Debe ingresar un monto en dolares valido.", vbInformation, "Aviso"
            Flex.Col = 6
            Flex.Row = lnI
            Me.Flex.SetFocus
            Exit Sub
        End If
    Next lnI
    
    lsMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        
    Me.MousePointer = 11
    sPla = Me.txtPlanillas.Text
    sFec = Format$(CDate(txtFecPla.value), gsFormatoMovFecha)
    sFecAct = FechaHora(gdFecSis)
    sTCC = lblTCC
     
    If MsgBox("Desea grabar la informacion??", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oCred.BeginTrans
        'Elimina todo lo existente en la tabla de conceptos por empleado y en la tabla de otros prestamos
        oCred.EliminaCreditosAdmOtros Me.txtPlanillas.Text, Format(CDate(Me.txtFecPla), gsFormatoMovFecha)
        
        'Recorre la lista y graba todo lo seleccionado
        For lnI = 1 To Me.Flex.Rows - 1
            sCta = Right(Me.Flex.TextMatrix(lnI, 3), 3)
            sPers = Me.Flex.TextMatrix(lnI, 1)
            If sCta = Moneda.gMonedaNacional Then
                nMonto = CDbl(Me.Flex.TextMatrix(lnI, 5))
                sTCC = "NULL"
            Else
                nMonto = CDbl(Me.Flex.TextMatrix(lnI, 6))
                sTCC = lblTCC
            End If
            
            oCred.InsertaCreditosAdmOtros sPers, sPla, sFec, sCta, CCur(nMonto), sTCC, lsMovNro, Trim(Right(Me.Flex.TextMatrix(lnI, 4), 15)), Format(Me.Flex.TextMatrix(lnI, 5), "#.00")
        Next lnI
    oCred.CommitTrans
    
    MsgBox "Grabación completa", vbInformation, "Aviso"
    Me.MousePointer = 0
    cmdCancelar_Click
End Sub

Private Sub cmdImprimir_Click()
    Dim i As Integer
    Dim lsCadena As String
    Dim lnTotal As Currency
    Dim lnTotalD As Currency
    Dim lnPagina As Long
    Dim lnItem As Long
    
    Dim lsNombre As String * 45
    Dim lsCredito As String * 15
    Dim lsMontoS As String * 15
    Dim lsMontoD As String * 15
    
    Dim oPrevio As clsPrevio
    Set oPrevio = New clsPrevio
5
    lsCredito = ""
    lsCadena = ""
    lnTotal = 0
    lnTotalD = 0
    lsCadena = lsCadena & CabeceraPagina("CREDITOS ADMINISTRATIVOS", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis)
    lsCadena = lsCadena & Encabezado("Nombre;10; ;26;Credito;20;Monto S/;25;Monto $;17; ;21;", lnItem)
    For i = 1 To Me.Flex.Rows - 1
        If Right(Flex.TextMatrix(i, 3), 2) = Moneda.gMonedaNacional Then
        
            lsNombre = Flex.TextMatrix(i, 2)
            RSet lsMontoS = Flex.TextMatrix(i, 5)
            RSet lsMontoD = Flex.TextMatrix(i, 6)
            
            lnTotal = lnTotal + CCur(Flex.TextMatrix(i, 5))
            lnTotalD = lnTotalD + CCur(Flex.TextMatrix(i, 6))
            
            lsCadena = lsCadena & lsNombre & "  " & lsCredito & "  " & lsMontoS & "  " & lsMontoD & oImpresora.gPrnSaltoLinea
            
            lnItem = lnItem + 1
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("CREDITOS ADMINISTRATIVOS", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis)
                lsCadena = lsCadena & Encabezado("Nombre;10; ;26;Credito;20;Monto S/;25;Monto $;17; ;21;", lnItem)
            End If
        End If
    Next i
    lsNombre = ""
    lsCredito = ""
    RSet lsMontoS = Format(lnTotal, "#,##0.00")
    RSet lsMontoD = Format(lnTotalD, "#,##0.00")
    lsCadena = lsCadena & String(116, "=") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsNombre & "  " & lsCredito & "  " & lsMontoS & "  " & lsMontoD & oImpresora.gPrnSaltoLinea
    
    lnTotal = 0
    lnTotalD = 0
    lsCadena = lsCadena & CabeceraPagina("CREDITOS ADMINISTRATIVOS", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis, "2")
    lsCadena = lsCadena & Encabezado("Nombre;10; ;26;Credito;20;Monto S/;25;Monto $;17; ;21;", lnItem)
    For i = 1 To Me.Flex.Rows - 1
        If Right(Flex.TextMatrix(i, 3), 2) = Moneda.gMonedaExtranjera Then
            
            lsNombre = Flex.TextMatrix(i, 2)
            RSet lsMontoS = Flex.TextMatrix(i, 5)
            RSet lsMontoD = Flex.TextMatrix(i, 6)
            
            lnTotal = lnTotal + CCur(Flex.TextMatrix(i, 5))
            lnTotalD = lnTotalD + CCur(Flex.TextMatrix(i, 6))
            
            lsCadena = lsCadena & lsNombre & "  " & lsCredito & "  " & lsMontoS & "  " & lsMontoD & oImpresora.gPrnSaltoLinea
            
            lnItem = lnItem + 1
            If lnItem > 54 Then
                lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
                lsCadena = lsCadena & CabeceraPagina("CREDITOS ADMINISTRATIVOS", lnPagina, lnItem, gsNomAge, gsEmpresa, gdFecSis)
                lsCadena = lsCadena & Encabezado("Nombre;10; ;26;Credito;20;Monto S/;25;Monto $;17; ;21;", lnItem)
            End If
        End If
    Next i
    lsNombre = ""
    lsCredito = ""
    RSet lsMontoS = Format(lnTotal, "#,##0.00")
    RSet lsMontoD = Format(lnTotalD, "#,##0.00")
    lsCadena = lsCadena & String(116, "=") & oImpresora.gPrnSaltoLinea
    lsCadena = lsCadena & lsNombre & "  " & lsCredito & "  " & lsMontoS & "  " & lsMontoD & oImpresora.gPrnSaltoLinea
    
    oPrevio.Show lsCadena, Caption, True
End Sub

Private Sub cmdNuevo_Click()
    Flex.AdicionaFila
    Flex.SetFocus
    Me.cmdGrabar.Enabled = True
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Flex_OnChangeCombo()
    If Flex.Col = 3 Then
        Flex.TextMatrix(Flex.Row, 6) = Format(0, "#0.00")
        Flex.TextMatrix(Flex.Row, 5) = Format(0, "#0.00")
    End If
End Sub

Private Sub Flex_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    
    If Me.Flex.Col = 5 Then
        If Right(Flex.TextMatrix(Flex.Row, 3), 2) = Moneda.gMonedaNacional Then
            Flex.TextMatrix(Flex.Row, 6) = Format(0, "#0.00")
        End If
    ElseIf Me.Flex.Col = 6 Then
        If Right(Flex.TextMatrix(Flex.Row, 3), 2) = Moneda.gMonedaExtranjera Then
            Flex.TextMatrix(Flex.Row, 5) = Format(Flex.TextMatrix(Flex.Row, 6) * Me.lblTCC.Caption, "#0.00")
        End If
    End If
End Sub

Private Sub Flex_RowColChange()
    Dim rsE As ADODB.Recordset
    Dim oCons As DConstantes
    Dim oCred As DRHPrestamosAdm
    
    If Me.Flex.Col = 3 Then
        Set oCons = New DConstantes
        Set rsE = New ADODB.Recordset
        Set rsE = oCons.GetConstante(gMoneda)
        Me.Flex.CargaCombo rsE
    ElseIf Me.Flex.Col = 4 Then
        Set rsE = New ADODB.Recordset
        Set oCred = New DRHPrestamosAdm
        Set rsE = oCred.GetRHCreditosAdmInst
        Me.Flex.CargaCombo rsE
    ElseIf Me.Flex.Col = 5 Then
        If Right(Flex.TextMatrix(Flex.Row, 3), 2) = Moneda.gMonedaExtranjera Then
            Flex.ColumnasAEditar = "X-1-X-3-4-X-6-X"
        Else
            Flex.ColumnasAEditar = "X-1-X-3-4-5-X-X"
        End If
    ElseIf Me.Flex.Col = 6 Then
        If Right(Flex.TextMatrix(Flex.Row, 3), 2) = Moneda.gMonedaExtranjera Then
            Flex.ColumnasAEditar = "X-1-X-3-4-X-6-X"
        Else
            Flex.ColumnasAEditar = "X-1-X-3-4-5-X-X"
        End If
    End If
    Suma
End Sub

Private Sub Form_Load()
    Dim oPla As DActualizaDatosConPlanilla
    Set oPla = New DActualizaDatosConPlanilla
    
    GetTipCambio gdFecSis, Not gbBitCentral
    lblTCC = Format$(gnTipCambioV, "#0.0000")
    lblTCF = Format$(gnTipCambio, "#0.0000")

    Me.txtPlanillas.rs = oPla.GetPlanillas(, True)

    cmdGrabar.Enabled = False
    txtFecPla.value = gdFecSis
End Sub

Private Sub BuscaDatosPlanilla()
    Dim rsEmp As ADODB.Recordset
    Dim sPla As String, sFec As String
    Dim L As ListItem
    Dim oCred As DRHPrestamosAdm
    Set oCred = New DRHPrestamosAdm
    
    sPla = Me.txtPlanillas.Text
    sFec = Format$(CDate(txtFecPla.value), "yyyymmdd")
    Set rsEmp = New ADODB.Recordset
    rsEmp.CursorLocation = adUseClient
    
    Set rsEmp = oCred.GetCreditosAdmOtros(sPla, sFec)
    
    Set rsEmp.ActiveConnection = Nothing
    If Not (rsEmp.EOF And rsEmp.BOF) Then
        Me.Flex.rsFlex = rsEmp
    Else
        MsgBox "No se han registrado prestamos para esta planilla", vbInformation, "Aviso"
    End If
End Sub

Private Sub txtPlanillas_EmiteDatos()
    Me.lblPlanillaG.Caption = Me.txtPlanillas.psDescripcion
End Sub

Private Sub Suma()
    Dim lnI As Integer
    Dim lnSumSol As Currency
    Dim lnSumDol As Currency
    
    lnSumSol = 0
    lnSumDol = 0
    
    For lnI = 1 To Me.Flex.Rows - 1
        If IsNumeric(Me.Flex.TextMatrix(lnI, 5)) Then lnSumSol = lnSumSol + CCur(Me.Flex.TextMatrix(lnI, 5))
        If IsNumeric(Me.Flex.TextMatrix(lnI, 6)) Then lnSumDol = lnSumDol + CCur(Me.Flex.TextMatrix(lnI, 6))
    Next lnI
    
    Me.lblTotSol.Caption = Format(lnSumSol, "#,##0.00")
    Me.lblTotDol.Caption = Format(lnSumDol, "#,##0.00")
End Sub
