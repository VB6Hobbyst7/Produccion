VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPigVentaJoyas 
   Caption         =   "Venta de Joyas "
   ClientHeight    =   7500
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10155
   FillColor       =   &H00000040&
   ForeColor       =   &H00000040&
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   10155
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   6915
      Left            =   105
      TabIndex        =   3
      Top             =   60
      Width           =   9930
      Begin VB.Frame frTipoVenta 
         Caption         =   "Tipo Venta"
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
         Height          =   885
         Left            =   90
         TabIndex        =   24
         Top             =   135
         Width           =   2340
         Begin VB.OptionButton optTipoVenta 
            Caption         =   "Al Titular"
            Height          =   195
            Index           =   2
            Left            =   135
            TabIndex        =   27
            Top             =   615
            Width           =   1470
         End
         Begin VB.OptionButton optTipoVenta 
            Caption         =   "Por Responsabilidad"
            Height          =   195
            Index           =   1
            Left            =   135
            TabIndex        =   26
            Top             =   420
            Width           =   1890
         End
         Begin VB.OptionButton optTipoVenta 
            Caption         =   "A Terceros"
            Height          =   195
            Index           =   0
            Left            =   135
            TabIndex        =   25
            Top             =   225
            Value           =   -1  'True
            Width           =   1470
         End
      End
      Begin VB.Frame frDocumento 
         Caption         =   "Documento"
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
         Height          =   885
         Left            =   2490
         TabIndex        =   4
         Top             =   135
         Width           =   4260
         Begin VB.OptionButton OptDocumento 
            Caption         =   "Boleta"
            Height          =   195
            Index           =   0
            Left            =   150
            TabIndex        =   8
            Top             =   285
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton OptDocumento 
            Caption         =   "Factura"
            Height          =   210
            Index           =   1
            Left            =   150
            TabIndex        =   7
            Top             =   555
            Width           =   855
         End
         Begin VB.CheckBox chkCambiar 
            Alignment       =   1  'Right Justify
            Height          =   195
            Left            =   3915
            TabIndex        =   5
            ToolTipText     =   "Cambiar Número..."
            Top             =   450
            Width           =   195
         End
         Begin MSMask.MaskEdBox mskNroDocumento 
            Height          =   300
            Left            =   2580
            TabIndex        =   6
            Top             =   345
            Width           =   1230
            _ExtentX        =   2170
            _ExtentY        =   529
            _Version        =   393216
            BackColor       =   16777215
            ForeColor       =   16711680
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSerDocumento 
            Height          =   300
            Left            =   1995
            TabIndex        =   33
            Top             =   345
            Width           =   555
            _ExtentX        =   979
            _ExtentY        =   529
            _Version        =   393216
            ForeColor       =   16711680
            Enabled         =   0   'False
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            PromptChar      =   "_"
         End
         Begin VB.Label Label11 
            Caption         =   "Número :"
            Height          =   195
            Left            =   1260
            TabIndex        =   9
            Top             =   405
            Width           =   645
         End
      End
      Begin SICMACT.FlexEdit feJoyas 
         Height          =   3825
         Left            =   165
         TabIndex        =   10
         Top             =   2055
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   6747
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   "Sec-Contrato-Pza-Descripcion-PVenta-IGV-Item-nRemate"
         EncabezadosAnchos=   "400-1700-350-5845-1200-0-0-0"
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
         ColumnasAEditar =   "X-1-2-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-R-L-R-R-R-R"
         FormatosEdit    =   "0-0-3-0-2-2-3-3"
         CantEntero      =   10
         TextArray0      =   "Sec"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Frame frCliente 
         Caption         =   "Datos del Cliente"
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
         Height          =   870
         Left            =   90
         TabIndex        =   28
         Top             =   1005
         Width           =   9750
         Begin VB.CommandButton cmdBuscar 
            Height          =   555
            Left            =   8895
            Picture         =   "frmPigVentaJoyas.frx":0000
            Style           =   1  'Graphical
            TabIndex        =   32
            Top             =   195
            Width           =   720
         End
         Begin MSComctlLib.ListView lstCliente 
            Height          =   630
            Left            =   60
            TabIndex        =   29
            Top             =   165
            Width           =   8745
            _ExtentX        =   15425
            _ExtentY        =   1111
            View            =   3
            LabelEdit       =   1
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            AllowReorder    =   -1  'True
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   1
            NumItems        =   5
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Código"
               Object.Width           =   2293
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Nombre o Razón Social"
               Object.Width           =   4851
            EndProperty
            BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   2
               Text            =   "Documento"
               Object.Width           =   1764
            EndProperty
            BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   3
               Text            =   "Dirección"
               Object.Width           =   4586
            EndProperty
            BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Alignment       =   2
               SubItemIndex    =   4
               Text            =   "Telefono"
               Object.Width           =   1764
            EndProperty
         End
      End
      Begin VB.Frame frJoyas 
         Caption         =   "Joyas"
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
         Height          =   4980
         Left            =   90
         TabIndex        =   11
         Top             =   1845
         Width           =   9750
         Begin VB.CommandButton cmdEliminar 
            Caption         =   "&Eliminar"
            Height          =   315
            Left            =   1245
            TabIndex        =   13
            Top             =   4320
            Width           =   930
         End
         Begin VB.CommandButton cmdAgregar 
            Caption         =   "&Agregar"
            Height          =   315
            Left            =   165
            TabIndex        =   12
            Top             =   4320
            Width           =   930
         End
         Begin VB.Label lblTotalS 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00404080&
            Height          =   315
            Left            =   8400
            TabIndex        =   23
            Top             =   4545
            Width           =   1200
         End
         Begin VB.Label Label9 
            Caption         =   "Total S/.   :"
            Height          =   195
            Left            =   7455
            TabIndex        =   22
            Top             =   4605
            Width           =   840
         End
         Begin VB.Label lblTotalD 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H80000008&
            Height          =   315
            Left            =   6060
            TabIndex        =   21
            Top             =   4545
            Width           =   1200
         End
         Begin VB.Label Label7 
            Caption         =   "Total US$ :"
            Height          =   195
            Left            =   5100
            TabIndex        =   20
            Top             =   4605
            Width           =   840
         End
         Begin VB.Label Label6 
            Caption         =   "I.G.V.        :"
            Height          =   195
            Left            =   7455
            TabIndex        =   19
            Top             =   4200
            Width           =   840
         End
         Begin VB.Label Label5 
            Caption         =   "SubTotal   :"
            Height          =   195
            Left            =   5100
            TabIndex        =   18
            Top             =   4200
            Width           =   840
         End
         Begin VB.Label Label4 
            Caption         =   "% Dcto.     :"
            Height          =   195
            Left            =   2760
            TabIndex        =   17
            Top             =   4200
            Width           =   840
         End
         Begin VB.Label lblSubTotal 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   6060
            TabIndex        =   16
            Top             =   4140
            Width           =   1200
         End
         Begin VB.Label lblDescuento 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   3705
            TabIndex        =   15
            Top             =   4140
            Width           =   1200
         End
         Begin VB.Label lblIGV 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   1  'Fixed Single
            ForeColor       =   &H00C00000&
            Height          =   315
            Left            =   8400
            TabIndex        =   14
            Top             =   4140
            Width           =   1200
         End
      End
      Begin VB.Label lblTipoCambio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         BorderStyle     =   1  'Fixed Single
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
         Height          =   315
         Left            =   8490
         TabIndex        =   31
         Top             =   570
         Width           =   1335
      End
      Begin VB.Label Label12 
         Caption         =   "Tipo de Cambio    :"
         Height          =   255
         Left            =   6945
         TabIndex        =   30
         Top             =   615
         Width           =   1410
      End
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   6210
      TabIndex        =   2
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   7530
      TabIndex        =   1
      Top             =   7080
      Width           =   1095
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   8850
      TabIndex        =   0
      Top             =   7080
      Width           =   1095
   End
End
Attribute VB_Name = "frmPigVentaJoyas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'****************************************************
'* frmPigVentaJoyas - Venta de Joyas en Tienda
'* EAFA - 01/10/2002
'****************************************************
Dim fnVarSubTotal As Currency
Dim fnVarIGV As Currency
Dim fnVarTotalS As Currency
Dim fnVarTotalD As Currency
Dim fnSw As Integer
Dim fnVarTipDoc As Integer
Dim fsVarSerDoc As String
Dim fsVarNumDoc As String
Dim lsPersCod As String
Dim lsPersNombre As String
Dim fnVarTipoVenta As Integer
Dim fnVarTipoProceso As Integer
Dim fmJoyas(30, 6) As String
Dim fsDescrip As String
Dim nFilas As Integer

Public Sub Inicio()
GetTipCambio (gdFecSis)
lblTipoCambio.Caption = Format(gnTipCambioC, "##0.000 ")
Limpiar
Me.Show 1
End Sub

Private Sub chkCambiar_Click()
If chkCambiar.value Then
   mskNroDocumento.Enabled = True
   mskNroDocumento.SetFocus
Else
   mskNroDocumento.Enabled = False
End If
End Sub

Private Sub CmdAgregar_Click()
    If feJoyas.Rows <= 20 Then
        feJoyas.AdicionaFila
        
        If feJoyas.Rows >= 2 Then
            cmdEliminar.Enabled = True
        End If
    Else
        cmdAgregar.Enabled = False
        MsgBox "Sólo puede ingresar como máximo veinte piezas", vbInformation, "Aviso"
    End If
    If fnSw = 0 Then
       frTipoVenta.Enabled = False
       frCliente.Enabled = True
       frDocumento.Enabled = True
       cmdGrabar.Enabled = True
       feJoyas.Enabled = True
       OptDocumento(0).value = True
       OptDocumento_Click (0)
       fnSw = 1
    End If
    feJoyas.SetFocus
End Sub

Private Sub cmdBuscar_Click()
Dim loPers As UPersona
Dim lsEstados As String
Dim lstTmpCliente As ListItem
Dim loPersContrato As dColPContrato
Dim lrContratos As ADODB.Recordset
Dim loCuentas As UProdPersona

On Error GoTo ControlError

Set loPers = New UPersona
    Set loPers = frmBuscaPersona.Inicio
    If loPers Is Nothing Then Exit Sub
     lsPersCod = loPers.sPersCod
     lsPersNombre = loPers.sPersNombre
    
     lstCliente.ListItems.Clear
    Set lstTmpCliente = lstCliente.ListItems.Add(, , Trim(lsPersCod))
          lstTmpCliente.SubItems(1) = Trim(lsPersNombre)
          If loPers.sPersPersoneria = gPersonaNat Then
             lstTmpCliente.SubItems(2) = Trim(IIf(IsNull(loPers.sPersIdnroDNI), "", loPers.sPersIdnroDNI))
          Else
             lstTmpCliente.SubItems(2) = Trim(IIf(IsNull(loPers.sPersIdnroRUC), "", loPers.sPersIdnroRUC))
          End If
          lstTmpCliente.SubItems(3) = Trim(IIf(IsNull(loPers.sPersDireccDomicilio), "", loPers.sPersDireccDomicilio))
          lstTmpCliente.SubItems(4) = Trim(IIf(IsNull(loPers.sPersTelefono), "", loPers.sPersTelefono))
          
    Set loPers = Nothing

    Exit Sub

ControlError:   ' Rutina de control de errores.
    MsgBox " Error: " & Err.Number & " " & Err.Description & vbCr & _
        " Avise al Area de Sistemas ", vbInformation, " Aviso "
End Sub

Private Sub cmdCancelar_Click()
Limpiar
End Sub

Private Sub cmdeliminar_Click()
    
    feJoyas.EliminaFila feJoyas.Row
    If feJoyas.Rows <= 20 Then
        cmdAgregar.Enabled = True
    End If
    SumaColumnas
  
End Sub

Private Sub cmdGrabar_Click()
Dim loContFunct As NContFunciones
Dim lrGraba As NPigContrato
Dim rsJoyas As Recordset
Dim lsMovNro As String
Dim lsFechaHoraGrab As String

If lsPersCod = "" Then
   MsgBox "Debe indicar el Nombre del Comprador...", vbInformation, "Aviso"
   Exit Sub
End If
If Val(lblTotalS.Caption) = 0 Then
   MsgBox "Debe ingresar al Menos una Joya a vender...", vbInformation, "Aviso"
   Exit Sub
End If
If mskNroDocumento.Text = "" Then
    MsgBox "Debe Espeicificar el Documento a Emitir...", vbInformation, "Aviso"
   Exit Sub
End If

If MsgBox("Desea Grabar la " & fsDescrip & "? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False
    'Genera el Mov Nro
    Set loContFunct = New NContFunciones
          lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
        
    lsFechaHoraGrab = fgFechaHoraGrab(lsMovNro)
    
    Set rsJoyas = feJoyas.GetRsNew
           nFilas = CargaMatrix(rsJoyas)
    Set rsJoyas = Nothing
    
    Set lrGraba = New NPigContrato
        Call lrGraba.nRegistraVentaJoyas(lsMovNro, fnVarTipDoc, fsVarSerDoc & fsVarNumDoc, lsPersCod, _
                                                                    Val(Format$(lblSubTotal.Caption, "#0.00")), Val(Format$(lblIGV.Caption, "#0.00")), _
                                                                    Val(Format$(lblTotalS.Caption, "#0.00")), gMonedaNacional, fnVarTipoVenta, fmJoyas, nFilas)
        '****** Impresion del Comprobante de Venta
        
        
    Set lrGraba = Nothing
    Limpiar
End If

End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub feJoyas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If feJoyas.TextMatrix(feJoyas.Row, 1) = "" Then
       cmdeliminar_Click
    End If
End If
End Sub

Private Sub FEJoyas_OnCellChange(pnRow As Long, pnCol As Long)
Dim oPigCalculos As NPigCalculos
Dim oPigFunciones As DPigFunciones
Dim oPigJoyas As dPigContrato
Dim lrDatos As ADODB.Recordset

Dim lnPOro As Currency

    If feJoyas.Col = 2 Then
       If feJoyas.TextMatrix(feJoyas.Row, 2) <> "" Then                                                          'Número de Pieza a Vender
          feJoyas.TextMatrix(feJoyas.Row, 6) = feJoyas.TextMatrix(feJoyas.Row, 0)            '*** Asigna el valor del Item
          
          Set lrDatos = New ADODB.Recordset
          Set oPigJoyas = New dPigContrato
                 Set lrDatos = oPigJoyas.dObtieneDetalleJoyaVenta(Right(feJoyas.TextMatrix(feJoyas.Row, 1), 18), Val(feJoyas.TextMatrix(feJoyas.Row, 2)))
          Set oPigJoyas = Nothing
        
          If lrDatos.EOF And lrDatos.BOF Then                                                     'No se encuentra
              MsgBox "Pieza no Existe o No se Encuentra en la Tienda", vbInformation, "Aviso"
              Exit Sub
          End If
          If lrDatos!nSituacionPieza <> gPigSituacionDisponible Then              'Vendida
              MsgBox "Pieza ya ha sido vendida...", vbInformation, "Aviso"
              Exit Sub
          End If
          If lrDatos!nUbicaPieza <> Val(gsCodAge) Then                                     'No se Encuentra en la Tienda
              MsgBox "Pieza no se Encuentra en la Tienda...", vbInformation, "Aviso"
              Exit Sub
          End If
          
          Set oPigFunciones = New DPigFunciones
          
          feJoyas.TextMatrix(feJoyas.Row, 3) = Trim(lrDatos!Tipo) & " DE " & Trim(lrDatos!Material) & " DE " & Format(lrDatos!pNeto, "#,##0.00") & " grs. " & Trim(lrDatos!Observacion) & " " & Trim(lrDatos!ObservAdic)
          feJoyas.TextMatrix(feJoyas.Row, 7) = lrDatos!nRemate
          If optTipoVenta(0).value = True Then                                                                                                      ' Venta A Terceros
              lnPOro = oPigFunciones.GetPrecioMaterial(gPigTipoValorVenta, Val(lrDatos!Material), gMonedaExtranjera)
              feJoyas.TextMatrix(feJoyas.Row, 4) = Round((lnPOro * Val(lrDatos!pNeto)) + Val(lrDatos!TasAdicion) * Val(lblTipoCambio.Caption), 2)
              feJoyas.TextMatrix(feJoyas.Row, 5) = Round(feJoyas.TextMatrix(feJoyas.Row, 4) * 0.18, 4)    ' IGV
          ElseIf optTipoVenta(1).value = True Or optTipoVenta(2).value = True Then                                   ' Venta Por Responsabilidad   Or   Venta Al Titular
              feJoyas.TextMatrix(feJoyas.Row, 4) = Round(Val(lrDatos!nValorDeuda), 2)
              feJoyas.TextMatrix(feJoyas.Row, 5) = Round(feJoyas.TextMatrix(feJoyas.Row, 4) * 0.18, 4)    ' IGV
          Else
              MsgBox "No se ha Definido el Proceso para este Tipo de Venta...", vbInformation, "Aviso"
              Exit Sub
          End If
           
           Set oPigFunciones = Nothing
           Set lrDatos = Nothing
          
           SumaColumnas
       End If
    End If
    
End Sub

Private Sub FEJoyas_RowColChange()
Dim oPigFuncion As DPigFunciones

    Set oPigFuncion = New DPigFunciones
    
    Select Case feJoyas.Col
    Case 1
        Set rsTipoJoya = oPigFuncion.GetConstante(gColocPigTipoJoya)
        feJoyas.CargaCombo rsTipoJoya
        Set rsTipoJoya = Nothing
    Case 2
        Set rsSTipoJoya = oPigFuncion.GetConstante(gColocPigSubTipoJoya)
        feJoyas.CargaCombo rsSTipoJoya
        Set rsSTipoJoya = Nothing
    Case 3
        Set rsMaterial = oPigFuncion.GetConstante(gColocPigMaterial)
        feJoyas.CargaCombo rsMaterial
        Set rsMaterial = Nothing
    Case 4
        Set rsEstadoJoya = oPigFuncion.GetConstante(gColocPigEstConservaJoya)
        feJoyas.CargaCombo rsEstadoJoya
        Set rsEstadoJoya = Nothing
    
    End Select
    
    Set oPigFuncion = Nothing

End Sub

Private Sub SumaColumnas()
Dim i As Integer
Dim oPigCalculos As NPigCalculos
    
    fnVarSubTotal = 0: fnVarIGV = 0: fnVarTotalS = 0: fnVarTotalD = 0
   
    fnVarSubTotal = feJoyas.SumaRow(4)
    fnVarIGV = Round(feJoyas.SumaRow(5), 2)
    fnVarTotalS = fnVarSubTotal + fnVarIGV
    fnVarTotalD = Round(fnVarTotalS / Val(lblTipoCambio.Caption), 2)
    
    lblSubTotal.Caption = Format(fnVarSubTotal, "###,##0.00 ")
    lblIGV.Caption = Format(fnVarIGV, "###,##0.00 ")
    lblTotalS.Caption = Format(fnVarTotalS, "###,##0.00 ")
    lblTotalD.Caption = Format(fnVarTotalD, "###,##0.00 ")
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub OptDocumento_Click(Index As Integer)
Dim lrDocumento As DPigFunciones

Set lrDocumento = New DPigFunciones
Select Case Index
          Case 0                       '************* BOLETA
                   fnVarTipDoc = Val(gPigTipoBoleta)
                   fsVarSerDoc = lrDocumento.GetSerDocumento(Val(gsCodAge), fnVarTipDoc)
                   fsVarNumDoc = lrDocumento.GetNumDocumento(fnVarTipDoc, fsVarSerDoc)
                   fsDescrip = "Boleta"
          Case 1                       '************* FACTURA
                   fnVarTipDoc = Val(gPigTipoFactura)
                   fsVarSerDoc = lrDocumento.GetSerDocumento(Val(gsCodAge), fnVarTipDoc)
                   fsVarNumDoc = lrDocumento.GetNumDocumento(fnVarTipDoc, fsVarSerDoc)
                   fsDescrip = "Factura"
End Select
mskSerDocumento.Text = fsVarSerDoc
mskNroDocumento.Text = fsVarNumDoc

Set lrDocumento = Nothing
End Sub

Private Sub optTipoVenta_Click(Index As Integer)
    Select Case Index
           Case 0
                    fnVarTipoVenta = gPigTipoVentaATerceros                     '***** Venta a Terceros
           Case 1
                    fnVarTipoVenta = gPigTipoVentaPorResponsabilidad            '***** Venta Por Responsabilidad
           Case 2
                    fnVarTipoVenta = gPigTipoVentaAlTitular                     '***** Venta al Titular
    End Select
End Sub

Private Function CargaMatrix(ByVal prPiezas As Recordset) As Integer
Dim Fila As Integer
Fila = 0
Do While Not prPiezas.EOF
    fmJoyas(Fila, 0) = prPiezas!Item
    fmJoyas(Fila, 1) = prPiezas!Contrato
    fmJoyas(Fila, 2) = prPiezas!Pza
    fmJoyas(Fila, 3) = prPiezas!pVenta
    fmJoyas(Fila, 4) = prPiezas!IGV
    fmJoyas(Fila, 5) = prPiezas!nRemate
    Fila = Fila + 1
    prPiezas.MoveNext
Loop
CargaMatrix = Fila - 1
End Function

Private Sub Limpiar()
frTipoVenta.Enabled = True
frCliente.Enabled = False
frDocumento.Enabled = False
feJoyas.Clear
feJoyas.Rows = 2
feJoyas.FormaCabecera
feJoyas.Enabled = False
cmdEliminar.Enabled = False
cmdGrabar.Enabled = False
mskSerDocumento.Text = ""
mskNroDocumento.Text = ""
optTipoVenta(0).value = True
optTipoVenta_Click (0)
fnSw = 0
End Sub
