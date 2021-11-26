VERSION 5.00
Begin VB.Form frmCapNoCobroInactivas 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CUENTAS EXONERADAS DEL DESCUENTO DE INACTIVIDAD"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11565
   Icon            =   "frmCapNoCobroInactivas.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   1335
      Left            =   60
      TabIndex        =   4
      Top             =   0
      Width           =   11295
      Begin VB.Frame Frame3 
         Height          =   1095
         Left            =   120
         TabIndex        =   11
         Top             =   120
         Width           =   2535
         Begin VB.OptionButton OptOpciones 
            Caption         =   "Por Cuenta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   180
            TabIndex        =   14
            Top             =   760
            Width           =   1695
         End
         Begin VB.OptionButton OptOpciones 
            Caption         =   "Por Agencia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   180
            TabIndex        =   13
            Top             =   400
            Width           =   1695
         End
         Begin VB.OptionButton OptOpciones 
            Caption         =   "Todas las Cuentas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   180
            TabIndex        =   12
            Top             =   120
            Width           =   1935
         End
      End
      Begin VB.CommandButton CmdProcesar 
         Caption         =   "&Procesar"
         Height          =   375
         Left            =   9240
         TabIndex        =   10
         Top             =   480
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.Frame fraCuenta 
         Height          =   675
         Left            =   3120
         TabIndex        =   7
         Top             =   320
         Visible         =   0   'False
         Width           =   4410
         Begin VB.CommandButton cmdBuscar 
            Caption         =   "..."
            Height          =   375
            Left            =   3705
            TabIndex        =   9
            Top             =   225
            Width           =   375
         End
         Begin SICMACT.ActXCodCta txtCuenta 
            Height          =   435
            Left            =   120
            TabIndex        =   8
            Top             =   225
            Width           =   3615
            _ExtentX        =   6376
            _ExtentY        =   767
            Texto           =   "Cuenta N°"
            EnabledCta      =   -1  'True
            EnabledAge      =   -1  'True
            Prod            =   "232"
            CMAC            =   "109"
         End
      End
      Begin VB.Frame FraAgencia 
         Height          =   675
         Left            =   3120
         TabIndex        =   5
         Top             =   320
         Visible         =   0   'False
         Width           =   3720
         Begin VB.ComboBox cboAgencia 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   255
            Width           =   3495
         End
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7920
      TabIndex        =   2
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Height          =   5805
      Left            =   60
      TabIndex        =   0
      Top             =   1320
      Width           =   11295
      Begin SICMACT.FlexEdit grdCargaDatos 
         Height          =   5430
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   9578
         Cols0           =   7
         HighLight       =   2
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nro. Cuenta-Cliente-Exonerar-SaldoC-SaldoD-Exon_Desembolso"
         EncabezadosAnchos=   "1000-1800-5150-850-0-0-2000"
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
         ColumnasAEditar =   "X-X-X-3-4-5-X"
         TextStyleFixed  =   1
         ListaControles  =   "0-0-0-4-4-4-0"
         BackColor       =   -2147483639
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-C-R-R-L"
         FormatosEdit    =   "0-0-0-0-2-2-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   0
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   1005
         RowHeight0      =   300
         ForeColor       =   -2147483630
         ForeColorFixed  =   -2147483630
         CellForeColor   =   -2147483630
         CellBackColor   =   -2147483639
      End
   End
End
Attribute VB_Name = "frmCapNoCobroInactivas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by

Private Sub CargaAgencias()
Dim clsAge As COMDConstantes.DCOMAgencias, rstemp As ADODB.Recordset, i As Integer
   Set clsAge = New COMDConstantes.DCOMAgencias
   Set rstemp = clsAge.RecuperaAgencias()
   i = 0
        cboAgencia.AddItem "<Seleccione>" & Space(100 - Len("<Seleccione>")) & "00"
   While Not rstemp.EOF
        cboAgencia.AddItem Trim(rstemp!cAgeDescripcion) & Space(100 - Len(Trim(rstemp!cAgeDescripcion))) & rstemp!cAgeCod
        rstemp.MoveNext
   Wend
   If cboAgencia.ListCount > 0 Then cboAgencia.ListIndex = 0
   Set rstemp = Nothing
   Set clsAge = Nothing
End Sub
Private Sub CargaGrid(ByVal rsDatos As Recordset)
  Dim i As Integer
    i = 1
  Set Me.grdCargaDatos.Recordset = rsDatos
 If rsDatos.State = 1 Then rsDatos.Close
  
End Sub

Private Sub CboAgencia_Click()
Dim clmant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rstemp As New ADODB.Recordset

If cboAgencia.ListIndex > 0 Then
    Set clmant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rstemp = clmant.GetCuentasAhorro(True, Right(cboAgencia.Text, 2))
    If Not rstemp.EOF Then
        Screen.MousePointer = vbHourglass
        Call CargaGrid(rstemp)
    End If
    Screen.MousePointer = vbDefault
    
End If
Set clmant = Nothing
End Sub


Private Sub cmdBuscar_Click()
Dim clsPers As COMDPersona.UCOMPersona

Set clsPers = New COMDPersona.UCOMPersona
Set clsPers = frmBuscaPersona.Inicio

If Not clsPers Is Nothing Then
    Dim sPers As String
    Dim rsPers As ADODB.Recordset
    Dim clsCap As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim sCta As String
    Dim sRelac As String * 15
    Dim sEstado As String
    Dim clsCuenta As UCapCuenta
    sPers = clsPers.sPerscod
    Set clsCap = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rsPers = clsCap.GetCuentasPersona(sPers, Producto.gCapAhorros, , , , , gsCodAge)
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
            txtCuenta.Age = Mid(clsCuenta.sCtaCod, 4, 2)
            txtCuenta.SetFocusCuenta
            SendKeys "{Enter}"
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

Private Sub cmdImprimir_Click()
'    Dim oPrevio As previo.clsprevio, lsCadena As String
'    Dim oRep As COMNCaptaGenerales.NCOMCaptaReportes
'
'    Set oRep = New COMNCaptaGenerales.NCOMCaptaReportes
'    Set oPrevio = New previo.clsprevio
'
'    lsCadena = oRep.ReporteNoCobroInactivas(IIf(chkAgencia.value = vbChecked, Right(cboAgencia, 2), ""))
'
'    oPrevio.Show lsCadena, Caption, True, 66
'    Set oRep = Nothing
    ImprimirListaCtaInactivas
End Sub

Private Sub CmdProcesar_Click()
     Dim clmant As COMNCaptaGenerales.NCOMCaptaGenerales
     Dim rstemp As New ADODB.Recordset
    
     Set clmant = New COMNCaptaGenerales.NCOMCaptaGenerales
     Set rstemp = clmant.GetCuentasAhorro(True, "")
     If Not rstemp.EOF Then
         Call CargaGrid(rstemp)
     End If
      
     Set clmant = Nothing
     Set rstemp = Nothing
End Sub

Private Sub cmdsalir_Click()
 Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If KeyCode = vbKeyF12 And txtCuenta.Enabled = True Then 'F12
        Dim sCuenta As String
        sCuenta = frmValTarCodAnt.Inicia(Producto.gCapAhorros, False)
        If Val(Mid(sCuenta, 6, 3)) <> Producto.gCapAhorros Then
            MsgBox "Esta operación no le corresponde a este producto.", vbOKOnly + vbInformation, App.Title
            Exit Sub
        End If
        If sCuenta <> "" Then
            txtCuenta.NroCuenta = sCuenta
            txtCuenta.SetFocusCuenta
        End If
    End If
End Sub

Private Sub Form_Load()
  Call CargaAgencias
  OptOpciones(0).value = False
  OptOpciones(1).value = False
  OptOpciones(2).value = False
  'By Capi 20012009
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCapExoneraDsctoInactivas
'End By

End Sub

Private Sub grdCargaDatos_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
 Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales, bdescinact As Boolean, sMovNro As String
 Dim clsMov As COMNContabilidad.NCOMContFunciones
 Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
   
  
If grdCargaDatos.TextMatrix(pnRow, pnCol) = "." Then
     bdescinact = True
Else
     bdescinact = False
End If

   Set clsMov = New COMNContabilidad.NCOMContFunciones
   sMovNro = clsMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
   Set clsMov = Nothing

   Call clsMant.ActualizaCtaDescInact(grdCargaDatos.TextMatrix(pnRow, 1), bdescinact, sMovNro)
    'By Capi 21012009
     objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, , grdCargaDatos.TextMatrix(pnRow, 1), gCodigoCuenta
    'End by
            

   Set clsMant = Nothing
End Sub

Private Sub OptOpciones_Click(Index As Integer)
   grdCargaDatos.Clear
   grdCargaDatos.Rows = 2
   grdCargaDatos.FormaCabecera
   
   If OptOpciones(0).value = True Then
      CmdProcesar.Visible = True
      FraAgencia.Visible = False
      fraCuenta.Visible = False
   ElseIf OptOpciones(1).value = True Then
      CmdProcesar.Visible = False
      FraAgencia.Visible = True
      fraCuenta.Visible = False
   ElseIf OptOpciones(2).value = True Then
      CmdProcesar.Visible = False
      FraAgencia.Visible = False
      fraCuenta.Visible = True
   End If
End Sub

Private Sub txtCuenta_KeyPress(KeyAscii As Integer)
Dim clmant As COMNCaptaGenerales.NCOMCaptaGenerales
Dim rstemp As ADODB.Recordset

If KeyAscii = 13 Then
    Set clmant = New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rstemp = clmant.GetCuentasAhorro(True, "", txtCuenta.NroCuenta)
    If Not rstemp.EOF Then
        Call CargaGrid(rstemp)
    End If
End If
Set clmant = Nothing
End Sub

Public Sub ImprimirListaCtaInactivas()
    Dim oPrevio As previo.clsprevio, lsCadena As String
    Dim oRep As COMNCaptaGenerales.NCOMCaptaReportes
    Dim rs As New ADODB.Recordset
    Dim i As Integer
 
    With rs
    'Crear RecordSet
    .Fields.Append "cCtaCod", adVarChar, 20
    .Fields.Append "cPersNom", adVarChar, 100
    .Fields.Append "nSaldoCont", adCurrency
    .Fields.Append "nSaldoDisp", adCurrency
    .Open
    'Llenar Recordset
     For i = 1 To grdCargaDatos.Rows - 1
     
       If grdCargaDatos.TextMatrix(i, 3) = "." Then
           .AddNew
           .Fields("cCtaCod") = grdCargaDatos.TextMatrix(i, 1)
           .Fields("cPersNom") = grdCargaDatos.TextMatrix(i, 2)
           .Fields("nSaldoCont") = grdCargaDatos.TextMatrix(i, 4)
           .Fields("nSaldoDisp") = grdCargaDatos.TextMatrix(i, 5)
       End If
     Next
    End With
    
    Set oRep = New COMNCaptaGenerales.NCOMCaptaReportes
    Set oPrevio = New previo.clsprevio
    
    lsCadena = oRep.ReporteNoCobroInactivasN(rs)
    
    oPrevio.Show lsCadena, Caption, True, 66, gImpresora
    Set oRep = Nothing
End Sub
