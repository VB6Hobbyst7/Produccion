VERSION 5.00
Begin VB.Form frmOpeComisionContMoneda 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Comisión por Conteo de Monedas"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmOpeComisionContMoneda.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   360
      Left            =   4200
      TabIndex        =   3
      Top             =   5680
      Width           =   1170
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   360
      Left            =   3000
      TabIndex        =   2
      Top             =   5680
      Width           =   1170
   End
   Begin VB.Frame Frame3 
      Caption         =   " Comisión "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   9
      Top             =   4680
      Width           =   5295
      Begin VB.Label lblComision 
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3960
         TabIndex        =   13
         Top             =   345
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Comisión :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   12
         Top             =   375
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Total :"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   375
         Width           =   495
      End
      Begin VB.Label lblTotal 
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
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   720
         TabIndex        =   10
         Top             =   345
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   1200
      Width           =   5295
      Begin SICMACT.FlexEdit fgMonedas 
         Height          =   3090
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   5100
         _ExtentX        =   8996
         _ExtentY        =   5450
         Cols0           =   6
         FixedCols       =   2
         HighLight       =   2
         AllowUserResizing=   3
         EncabezadosNombres=   "N°-Descripción-Cantidad-Monto-cEfectivoCod-nEfectivoValor"
         EncabezadosAnchos=   "350-2420-800-1200-0-0"
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
            Size            =   6.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-2-3-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-R-R-C-C"
         FormatosEdit    =   "0-0-3-4-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         lbEditarFlex    =   -1  'True
         Enabled         =   0   'False
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         CellBackColor   =   -2147483633
      End
   End
   Begin VB.Frame frmOpeComisionOtros 
      Caption         =   " Cliente "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   5295
      Begin SICMACT.TxtBuscar TxtBCodPers 
         Height          =   330
         Left            =   840
         TabIndex        =   0
         Top             =   300
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label lblNumDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   4080
         TabIndex        =   14
         Top             =   300
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Cliente: "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   780
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   340
         Width           =   615
      End
      Begin VB.Label lblCliente 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   840
         TabIndex        =   4
         Top             =   735
         Width           =   4335
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   360
      Left            =   4440
      TabIndex        =   15
      Top             =   5160
      Width           =   930
   End
End
Attribute VB_Name = "frmOpeComisionContMoneda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmOpeComisionContMoneda
'** Descripción : Formulario para pago de comisiones por Consulta de Movimientos o Consulta de Saldos
'**               según TI-ERS012-2015
'** Creación : JUEZ, 20151229 09:00:00 AM
'*****************************************************************************************************
Option Explicit
Dim fnPorcComision As Double
Dim fnMontoExonerado As Double

Private Sub cmdAceptar_Click()
Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
Dim oBol As COMNCaptaGenerales.NCOMCaptaImpresion
Dim rsMonedas As ADODB.Recordset
Dim lnMovNro As Long
Dim lsMovNro As String
Dim lsBoleta As String

    If CDbl(lblComision.Caption) > 0 Then
        
        Set rsMonedas = fgMonedas.GetRsNew
        lsMovNro = GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
        lnMovNro = oNCapMov.OtrasOperaciones(lsMovNro, gComiAhoContMoneda, CDbl(lblComision.Caption), lblNumDoc.Caption, "", "1", TxtBCodPers.Text, , , , , , , gnMovNro)
        Set oNCapMov = Nothing
        
        If gnMovNro > 0 Then
            Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
                oNCapMov.GrabarOpeComisionConteoMoneda gnMovNro, rsMonedas
            Set oNCapMov = Nothing
            
            Set oBol = New COMNCaptaGenerales.NCOMCaptaImpresion
                lsBoleta = oBol.ImprimeBoleta("OTRAS COMISIONES", "Comisión Conteo de Monedas", gComiAhoContMoneda, CStr(lblComision.Caption), lblCliente.Caption, "________" & gMonedaNacional, lblNumDoc.Caption, 0, "0", "Nro Documento", 0, 0, False, False, , , , False, , "Nro Ope. : " & Str(lnMovNro), , gdFecSis, gsNomAge, gsCodUser, sLpt, , False, 0)
            Set oBol = Nothing
            
            If Trim(lsBoleta) <> "" Then ImprimeBoleta lsBoleta
            
            gnMovNro = 0
            cmdCancelar_Click
        Else
            MsgBox "Hubo un problema en el registro", vbInformation, "Aviso"
        End If
    Else
        MsgBox "La comisión a cobrar debe ser mayor a cero", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim oNCapDef As COMNCaptaGenerales.NCOMCaptaDefinicion
    Set oNCapDef = New COMNCaptaGenerales.NCOMCaptaDefinicion
        fnPorcComision = oNCapDef.GetCapParametro(2159)
        fnMontoExonerado = oNCapDef.GetCapParametro(2160)
    Set oNCapDef = Nothing
    Limpiar
End Sub

Private Sub TxtBCodPers_EmiteDatos()
    Dim R As ADODB.Recordset
    Dim oCred As COMDCredito.DCOMCredito
    Set oCred = New COMDCredito.DCOMCredito
    Set R = oCred.RecuperaDatosComision(TxtBCodPers.Text, 2)
    Set oCred = Nothing
    If Not R.EOF And Not R.BOF Then
        lblCliente.Caption = R!cPersNombre
        lblNumDoc.Caption = R!cPersIDnro
        CargaBilletajes
        TxtBCodPers.Enabled = False
        fgMonedas.Enabled = True
        fgMonedas.Col = 2
        fgMonedas.SetFocus
    Else
        Limpiar
    End If
    Set R = Nothing
End Sub

Private Sub CargaBilletajes()
Dim sql As String
Dim rs As ADODB.Recordset
Dim oContFunct As COMNContabilidad.NCOMContFunciones
Dim oEfec As COMDCajaGeneral.DCOMEfectivo
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim lnFila As Long
Dim i As Integer

Set oContFunct = New COMNContabilidad.NCOMContFunciones
Set oEfec = New COMDCajaGeneral.DCOMEfectivo

Set rs = New ADODB.Recordset
Set oCajero = New COMNCajaGeneral.NCOMCajero
Set rs = oEfec.EmiteBilletajes(gMonedaNacional, "M")

fgMonedas.FontFixed.Bold = True
fgMonedas.Clear
fgMonedas.FormaCabecera
fgMonedas.Rows = 2
Do While Not rs.EOF
    fgMonedas.AdicionaFila
    lnFila = fgMonedas.row
    fgMonedas.TextMatrix(lnFila, 1) = rs!Descripcion
    fgMonedas.TextMatrix(lnFila, 2) = Format(rs!Cantidad, "#,##0")
    fgMonedas.TextMatrix(lnFila, 3) = Format(rs!Monto, "#,##0.00")
    fgMonedas.TextMatrix(lnFila, 4) = rs!cEfectivoCod
    fgMonedas.TextMatrix(lnFila, 5) = rs!nEfectivoValor
    rs.MoveNext
Loop

rs.Close
Set rs = Nothing
Set oContFunct = Nothing
Set oEfec = Nothing
fgMonedas.Col = 2
fgMonedas.TopRow = 1
End Sub

Private Sub Limpiar()
    TxtBCodPers.Text = ""
    lblCliente.Caption = ""
    lblNumDoc.Caption = ""
    TxtBCodPers.Enabled = True
    CargaBilletajes
    fgMonedas.Enabled = False
    lblTotal.Caption = "0.00"
    lblComision.Caption = "0.00"
End Sub

Private Sub cmdCancelar_Click()
    Limpiar
    TxtBCodPers.SetFocus
End Sub

Private Sub fgMonedas_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
Dim lnValor As Currency
Dim lnComision As Currency

Select Case pnCol
    Case 3
        lnValor = CCur(IIf(fgMonedas.TextMatrix(pnRow, pnCol) = "", "0", fgMonedas.TextMatrix(pnRow, pnCol)))
        If Residuo(lnValor, CCur(fgMonedas.TextMatrix(pnRow, 5))) Then
            fgMonedas.TextMatrix(pnRow, 2) = Format(Round(lnValor / fgMonedas.TextMatrix(pnRow, 5), 0), "#,##0")
        Else
            Cancel = False
            Exit Sub
        End If
    Case 2
        lnValor = CCur(IIf(fgMonedas.TextMatrix(pnRow, pnCol) = "", "0", fgMonedas.TextMatrix(pnRow, pnCol)))
        fgMonedas.TextMatrix(pnRow, 3) = Format(lnValor * CCur(IIf(fgMonedas.TextMatrix(pnRow, 5) = "", "0", fgMonedas.TextMatrix(pnRow, 5))), "#,##0.00")
End Select

lblTotal.Caption = Format(fgMonedas.SumaRow(3), "#,##0.00")
lnComision = (CCur(lblTotal.Caption) - fnMontoExonerado) * (fnPorcComision / 100)
lnComision = IIf(lnComision < 0, 0, lnComision)
lblComision.Caption = Format(lnComision, "#,##0.00")
End Sub

Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub
