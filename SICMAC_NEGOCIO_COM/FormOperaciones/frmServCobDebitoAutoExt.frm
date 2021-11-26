VERSION 5.00
Begin VB.Form frmServCobDebitoAutoExt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Extorno Registro Débito Automático"
   ClientHeight    =   4575
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7560
   Icon            =   "frmServCobDebitoAutoExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4575
   ScaleWidth      =   7560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   6360
      TabIndex        =   11
      Top             =   4080
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Forma de Busqueda"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   5055
      Begin VB.OptionButton optBusqueda 
         Caption         =   "&Número de Movimiento"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.OptionButton optBusqueda 
         Caption         =   "&Usuario"
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
         Index           =   1
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox txtMovimiento 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   2160
         TabIndex        =   0
         Top             =   360
         Width           =   1575
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   3840
         TabIndex        =   7
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Height          =   2775
      Left            =   120
      TabIndex        =   5
      Top             =   1200
      Width           =   7335
      Begin SICMACT.FlexEdit FEMovimientos 
         Height          =   2445
         Left            =   120
         TabIndex        =   10
         Top             =   240
         Width           =   7095
         _ExtentX        =   12515
         _ExtentY        =   4313
         Cols0           =   6
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Movimiento-Usuario-Cuenta-Moneda-Nro Serv"
         EncabezadosAnchos=   "250-1400-1150-1900-1200-800"
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
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-C-C-L-L-R"
         FormatosEdit    =   "0-0-0-0-1-3"
         TextArray0      =   "#"
         SelectionMode   =   1
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   255
         RowHeight0      =   300
         TipoBusPersona  =   1
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Glosa"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   5280
      TabIndex        =   3
      Top             =   120
      Width           =   2175
      Begin VB.TextBox txtGlosa 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdExtonar 
      Caption         =   "&Extornar"
      Height          =   375
      Left            =   5160
      TabIndex        =   2
      Top             =   4080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   975
   End
End
Attribute VB_Name = "frmServCobDebitoAutoExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************************************************
'** Nombre : frmServCobDebitoAutoExt
'** Descripción : Formulario extornar afiliacion de una cuenta al débito automático para
'**               pagos de servicios de recaudo o créditos creado segun TI-ERS144-2014
'** Creación : JUEZ, 20150130 09:00:00 AM
'****************************************************************************************

Option Explicit

Dim oNCapGen As COMNCaptaGenerales.NCOMCaptaGenerales
Dim RS As ADODB.Recordset

Private Sub cmdBuscar_Click()
Dim lsMensaje As String

lsMensaje = IIf(optBusqueda.iTem(0).value, "Falta ingresar el Nro Movimiento", "Falta ingresar el Usuario")

If Trim(txtMovimiento) = "" Then
    MsgBox lsMensaje, vbInformation, "Aviso"
    txtMovimiento.SetFocus
    Exit Sub
End If

If optBusqueda.iTem(0).value = True And Not IsNumeric(txtMovimiento) Then
    MsgBox "Ingrese Nro de Operacion", vbInformation, "Aviso"
    txtMovimiento.SetFocus
    Exit Sub
End If

If optBusqueda.iTem(0).value = False And IsNumeric(txtMovimiento) Then
    MsgBox "Ingrese un Usuario", vbInformation, "Aviso"
    txtMovimiento.SetFocus
    Exit Sub
End If

CargaMovimientos
End Sub

Private Sub cmdCancelar_Click()
txtMovimiento = ""
txtGlosa.Text = ""
Call LimpiaFlex(FEMovimientos)
If optBusqueda.iTem(0).value = True Then
    txtMovimiento.MaxLength = 9
Else
    txtMovimiento.MaxLength = 4
End If
End Sub

Private Sub cmdExtonar_Click()
Dim oVistoElectronico As frmVistoElectronico
Set oVistoElectronico = New frmVistoElectronico
Dim lbRegistro As Boolean
Dim lbResultadoVisto As Boolean

If FEMovimientos.TextMatrix(FEMovimientos.Rows - 1, 0) = "" Then
    MsgBox "No existe movimiento para extornar", vbInformation, "Aviso"
    Exit Sub
End If
If Len(Trim(txtGlosa)) = 0 Then
    MsgBox "Por favor ingrese la glosa del extorno", vbInformation, "Aviso"
   
    txtGlosa.SetFocus
    Exit Sub
End If

Set oVistoElectronico = New frmVistoElectronico
lbResultadoVisto = False
lbResultadoVisto = oVistoElectronico.Inicio(3, gsOpeCod)
If Not lbResultadoVisto Then
    Exit Sub
End If

If MsgBox("¿Esta seguro que desea eliminar el Nro de Operación: " & FEMovimientos.TextMatrix(FEMovimientos.row, 1) & "?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
    Dim oNCapMov As COMNCaptaGenerales.NCOMCaptaMovimiento
    Set oNCapMov = New COMNCaptaGenerales.NCOMCaptaMovimiento
    Dim oNCOMContFunciones As COMNContabilidad.NCOMContFunciones
    Set oNCOMContFunciones = New COMNContabilidad.NCOMContFunciones
    Dim lsBoleta As String, lsMovNro As String
    
    lsMovNro = oNCOMContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    lbRegistro = oNCapMov.ExtornarServCobDebitoAuto(CLng(FEMovimientos.TextMatrix(FEMovimientos.row, 1)), lsMovNro, gsOpeCod, _
                                                    FEMovimientos.TextMatrix(FEMovimientos.row, 3), UCase(txtGlosa), gsCodAge, gsNomAge, gsCodCMAC, lsBoleta, sLpt, gbImpTMU)
    
    If lbRegistro Then
        oVistoElectronico.RegistraVistoElectronico (CLng(FEMovimientos.TextMatrix(FEMovimientos.row, 1)))
         
        If Trim(lsBoleta) <> "" Then ImprimeBoleta (lsBoleta)
        
        FEMovimientos.EliminaFila FEMovimientos.row
        txtGlosa = ""
    End If
    Set oNCapMov = Nothing
    Set oVistoElectronico = Nothing
End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
gsOpeCod = gExtornoServCobRegDebitoAuto
End Sub

Private Sub optBusqueda_Click(Index As Integer)
txtMovimiento = ""
Call LimpiaFlex(FEMovimientos)
If optBusqueda.iTem(0).value = True Then
    txtMovimiento.MaxLength = 9
Else
    txtMovimiento.MaxLength = 4
End If
txtMovimiento.SetFocus
End Sub

Private Sub CargaMovimientos()
Set oNCapGen = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set RS = oNCapGen.ObtenerServCobDebitoAuto(IIf(optBusqueda.iTem(0).value = False, 0, txtMovimiento), IIf(optBusqueda.iTem(0).value = False, Trim(UCase(txtMovimiento)), ""), Format(gdFecSis, "yyyyMMdd"), gsCodAge)
Set oNCapGen = Nothing

If Not RS.BOF Or Not RS.EOF Then
    Do While Not RS.EOF
        FEMovimientos.AdicionaFila
        FEMovimientos.TextMatrix(FEMovimientos.row, 0) = FEMovimientos.row
        FEMovimientos.TextMatrix(FEMovimientos.row, 1) = RS!nMovNro
        FEMovimientos.TextMatrix(FEMovimientos.row, 2) = RS!cUser
        FEMovimientos.TextMatrix(FEMovimientos.row, 3) = RS!cCtaCod
        FEMovimientos.TextMatrix(FEMovimientos.row, 4) = RS!cmoneda
        FEMovimientos.TextMatrix(FEMovimientos.row, 5) = RS!nNumServ
        RS.MoveNext
    Loop
Else
    MsgBox "No se encontraron datos", vbInformation, "Aviso"
    LimpiaFlex FEMovimientos
End If
End Sub

Private Sub ImprimeBoleta(ByVal sBoleta As String, Optional ByVal sMensaje As String = "Boleta Operación")
Dim nFicSal As Integer
Do
    nFicSal = FreeFile
    Open sLpt For Output As nFicSal
    Print #nFicSal, sBoleta & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    'Print #nFicSal, ""
    Close #nFicSal
Loop Until MsgBox("¿Desea Re-Imprimir " & sMensaje & " ?", vbQuestion + vbYesNo, "Aviso") = vbNo
End Sub

Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        cmdExtonar.SetFocus
    End If
End Sub

Private Sub txtMovimiento_KeyPress(KeyAscii As Integer)
    If optBusqueda.iTem(1).value Then
        KeyAscii = SoloLetras2(KeyAscii, True)
    ElseIf optBusqueda.iTem(0).value Then
        KeyAscii = SoloNumeros(KeyAscii)
    End If
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub
