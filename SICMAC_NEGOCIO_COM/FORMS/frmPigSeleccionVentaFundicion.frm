VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "msdatlst.ocx"
Begin VB.Form frmPigSeleccionVentaFundicion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Selección de Joyas Para Venta/Fundición"
   ClientHeight    =   6930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12300
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   12300
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   9465
      TabIndex        =   37
      Top             =   6495
      Width           =   1200
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   10905
      TabIndex        =   36
      Top             =   6495
      Width           =   1200
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   345
      Left            =   8025
      TabIndex        =   35
      Top             =   6495
      Width           =   1200
   End
   Begin VB.Frame Frame1 
      Height          =   6345
      Left            =   105
      TabIndex        =   0
      Top             =   45
      Width           =   12060
      Begin VB.CommandButton cmdEjecutar 
         Caption         =   "&Ejecutar"
         Height          =   360
         Left            =   10980
         TabIndex        =   26
         Top             =   840
         Width           =   855
      End
      Begin VB.Frame frmSeleccion 
         Caption         =   "Tipo de Selección"
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
         Height          =   1260
         Left            =   120
         TabIndex        =   23
         Top             =   180
         Width           =   1800
         Begin VB.OptionButton optSeleccion 
            Caption         =   "Fundición"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   1
            Left            =   180
            TabIndex        =   25
            Top             =   855
            Width           =   1455
         End
         Begin VB.OptionButton optSeleccion 
            Caption         =   "Venta"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Index           =   0
            Left            =   150
            TabIndex        =   24
            Top             =   300
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame frmCriterios 
         Caption         =   "Busquedad Por"
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
         Height          =   1320
         Left            =   2010
         TabIndex        =   8
         Top             =   150
         Width           =   8775
         Begin MSDataListLib.DataCombo cboRemDesde 
            Height          =   315
            Left            =   1305
            TabIndex        =   14
            Top             =   285
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.OptionButton optCriterios 
            Caption         =   "Contrato"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   3
            Left            =   3765
            TabIndex        =   13
            Top             =   825
            Width           =   1155
         End
         Begin VB.TextBox txtCtaCod 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   285
            Left            =   5055
            TabIndex        =   12
            Top             =   780
            Width           =   1995
         End
         Begin VB.OptionButton optCriterios 
            Caption         =   "Agencia"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   2
            Left            =   3720
            TabIndex        =   11
            Top             =   360
            Width           =   1155
         End
         Begin VB.OptionButton optCriterios 
            Caption         =   "Material"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   1
            Left            =   135
            TabIndex        =   10
            Top             =   780
            Width           =   1155
         End
         Begin VB.OptionButton optCriterios 
            Caption         =   "Remate"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   0
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Value           =   -1  'True
            Width           =   1155
         End
         Begin MSDataListLib.DataCombo cboRemHasta 
            Height          =   315
            Left            =   2550
            TabIndex        =   15
            Top             =   270
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboMaterial 
            Height          =   315
            Left            =   1305
            TabIndex        =   16
            Top             =   705
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboAgeDesde 
            Height          =   315
            Left            =   5055
            TabIndex        =   17
            Top             =   345
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin MSDataListLib.DataCombo cboAgeHasta 
            Height          =   315
            Left            =   7005
            TabIndex        =   18
            Top             =   315
            Width           =   1440
            _ExtentX        =   2540
            _ExtentY        =   556
            _Version        =   393216
            Text            =   ""
         End
         Begin VB.Label lblText1 
            Alignment       =   2  'Center
            Caption         =   "AL"
            Height          =   210
            Left            =   2190
            TabIndex        =   22
            Top             =   330
            Width           =   405
         End
         Begin VB.Label lblText2 
            Alignment       =   2  'Center
            Caption         =   "AL"
            Height          =   210
            Left            =   6540
            TabIndex        =   21
            Top             =   405
            Width           =   405
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "Joyas"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   4740
         Left            =   120
         TabIndex        =   1
         Top             =   1515
         Width           =   11835
         Begin SICMACT.FlexEdit feSeleccionadas 
            Height          =   3630
            Left            =   5730
            TabIndex        =   19
            Top             =   570
            Width           =   6000
            _ExtentX        =   10583
            _ExtentY        =   6403
            Cols0           =   13
            ScrollBars      =   1
            HighLight       =   2
            EncabezadosNombres=   "Item-Contrato-Pza-Tipo-Material-PNeto-nRemate-nValDeuda-nValDeudaComp-nValAdjudicado-OtroProceso-Estado-VAdicional"
            EncabezadosAnchos=   "0-1700-350-900-750-922-0-0-0-0-0-0-1150"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-12"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "R-R-C-L-L-R-R-R-C-R-L-L-R"
            FormatosEdit    =   "3-3-0-0-0-2-3-2-2-2-0-1-2"
            TextArray0      =   "Item"
            lbEditarFlex    =   -1  'True
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin VB.CommandButton cmdSelUno 
            Caption         =   ">"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5040
            TabIndex        =   5
            Top             =   870
            Width           =   540
         End
         Begin VB.CommandButton cmdRegresaUno 
            Caption         =   "<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5040
            TabIndex        =   4
            Top             =   1710
            Width           =   540
         End
         Begin VB.CommandButton cmdSelTodos 
            Caption         =   ">>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5040
            TabIndex        =   3
            Top             =   2565
            Width           =   540
         End
         Begin VB.CommandButton cmdRegresaTodos 
            Caption         =   "<<"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   405
            Left            =   5040
            TabIndex        =   2
            Top             =   3405
            Width           =   540
         End
         Begin SICMACT.FlexEdit feDisponibles 
            Height          =   3630
            Left            =   120
            TabIndex        =   20
            Top             =   570
            Width           =   4845
            _ExtentX        =   8546
            _ExtentY        =   6403
            Cols0           =   12
            ScrollBars      =   1
            HighLight       =   2
            EncabezadosNombres=   "Item-Contrato-Pza-Tipo-Material-PNeto-nRemate-nValDeuda-nValDeudaComp-nValAdjudicado-OtroProceso-Estado"
            EncabezadosAnchos=   "0-1700-350-900-750-972-730-1100-1350-1350-0-1500"
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
            ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X"
            ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0"
            EncabezadosAlineacion=   "R-R-C-L-L-R-R-R-R-R-L-L"
            FormatosEdit    =   "3-3-0-0-0-2-3-2-2-2-0-1"
            AvanceCeldas    =   1
            TextArray0      =   "Item"
            SelectionMode   =   1
            lbPuntero       =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            RowHeight0      =   300
         End
         Begin VB.Label lblTotPNetoSel 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   9885
            TabIndex        =   34
            Top             =   4290
            Width           =   1185
         End
         Begin VB.Label Label9 
            Caption         =   "Tot. Peso Neto :"
            Height          =   195
            Left            =   8640
            TabIndex        =   33
            Top             =   4365
            Width           =   1200
         End
         Begin VB.Label lblPzaSeleccionadas 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   7485
            TabIndex        =   32
            Top             =   4290
            Width           =   1020
         End
         Begin VB.Label Label7 
            Caption         =   "Seleccionadas :"
            Height          =   195
            Left            =   6255
            TabIndex        =   31
            Top             =   4365
            Width           =   1155
         End
         Begin VB.Label lblTotPNetoDisp 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   3630
            TabIndex        =   30
            Top             =   4290
            Width           =   1185
         End
         Begin VB.Label Label5 
            Caption         =   "Tot. Peso Neto :"
            Height          =   210
            Left            =   2370
            TabIndex        =   29
            Top             =   4350
            Width           =   1200
         End
         Begin VB.Label lblPzaDisponibles 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   1  'Fixed Single
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   345
            Left            =   1215
            TabIndex        =   28
            Top             =   4290
            Width           =   1020
         End
         Begin VB.Label Label3 
            Caption         =   "Disponibles :"
            Height          =   210
            Left            =   240
            TabIndex        =   27
            Top             =   4350
            Width           =   915
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Seleccionadas"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   5730
            TabIndex        =   7
            Top             =   240
            Width           =   6000
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Disponibles"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000D&
            Height          =   360
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   4845
         End
      End
   End
End
Attribute VB_Name = "frmPigSeleccionVentaFundicion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************************************************
'* frmPigSeleccionVentaFundicion : Seleccion de Joyas Para Venta ó Fundición
'* EAFA - 22/10/2002
'**************************************************************
Dim fnVarCriterio As Integer
Dim psCondicionAdic As String
Dim fmJoyas() As String
Dim fnTipoProceso As Integer

Private Sub cmdCancelar_Click()
limpiaInicia
End Sub

Private Sub cmdEjecutar_Click()
Dim lrDatos As ADODB.Recordset
Dim lrPigContrato As DPigContrato

If fnVarCriterio = 0 Then
    If Not IsNull(cboRemDesde.Text) And Not IsNull(cboRemHasta.Text) Then
        psCondicion = " Where CPJT.nTipoTasacion = " & gPigTipoTasacNor
        psCondicionAdic = psCondicionAdic & " And CPEJ.nRemate Between " & Val(cboRemDesde.BoundText) & " And " & Val(cboRemHasta.BoundText)
    Else
        MsgBox "Debe Especificar correctamente los Números de Remate", vbInformation, "Aviso"
        Exit Sub
    End If
ElseIf fnVarCriterio = 1 Then
    If Not IsNull(cboMaterial.Text) Then
        psCondicion = " Where CPJT.nTipoTasacion = " & gPigTipoTasacNor & " And CPJT.nMaterial =" & Val(cboMaterial.BoundText)
    Else
        MsgBox "Debe Especificar el Tipo de Material", vbInformation, "Aviso"
        Exit Sub
    End If
ElseIf fnVarCriterio = 2 Then
    If Not IsNull(cboAgeDesde.Text) And Not IsNull(cboAgeHasta.Text) Then
        psCondicion = " Where CPJT.nTipoTasacion = " & gPigTipoTasacNor & " And Convert(int, Substring(CPJT.cCtaCod,4,2)) Between Convert(int," & cboAgeDesde.BoundText & ") And Convert(int," & cboAgeHasta.BoundText & ") "
    Else
        MsgBox "Debe Especificar Correctamente las Agencias para la Transferencia", vbInformation, "Aviso"
        Exit Sub
    End If
ElseIf fnVarCriterio = 3 Then
    If Not IsNull(txtCtaCod.Text) Then
        psCondicion = " Where CPJT.nTipoTasacion = " & gPigTipoTasacNor & " And CPJT.cCtaCod = '" & txtCtaCod.Text & "'"
    Else
        MsgBox "Debe Especificar el Número de Contrato", vbInformation, "Aviso"
        Exit Sub
    End If
End If

Set lrDatos = New ADODB.Recordset
Set lrPigContrato = New DPigContrato
    Set lrDatos = lrPigContrato.dObtieneJoyasVentaFundicion(psCondicion, psCondicionAdic, fnTipoProceso)
Set lrPigContrato = Nothing

If lrDatos Is Nothing Or (lrDatos.BOF And lrDatos.EOF) Then
    MsgBox "No se Encuentro Información", vbInformation, "Aviso"
    Exit Sub
Else
    feDisponibles.Clear
    feDisponibles.Rows = 2
    feDisponibles.FormaCabecera
    Do While Not lrDatos.EOF
          feDisponibles.AdicionaFila
          feDisponibles.TextMatrix(feDisponibles.Row, 1) = lrDatos!cCtaCod
          feDisponibles.TextMatrix(feDisponibles.Row, 2) = lrDatos!iTem
          feDisponibles.TextMatrix(feDisponibles.Row, 3) = lrDatos!Tipo
          feDisponibles.TextMatrix(feDisponibles.Row, 4) = lrDatos!Material
          feDisponibles.TextMatrix(feDisponibles.Row, 5) = Format(lrDatos!pNeto, "##,##0.00")
          feDisponibles.TextMatrix(feDisponibles.Row, 6) = Format(lrDatos!nRemate, "###0")
          feDisponibles.TextMatrix(feDisponibles.Row, 7) = Format(lrDatos!nValorDeuda, "##,##0.00")
          feDisponibles.TextMatrix(feDisponibles.Row, 8) = Format(lrDatos!nValorDeudaComp, "##,##0.00")
          feDisponibles.TextMatrix(feDisponibles.Row, 9) = Format(lrDatos!nValorProceso, "##,##0.00")                  '*** Valor Adjudicado
          feDisponibles.TextMatrix(feDisponibles.Row, 10) = IIf(IsNull(lrDatos!OtroProceso), "", lrDatos!OtroProceso)
          feDisponibles.TextMatrix(feDisponibles.Row, 11) = Trim(lrDatos!DescEstado)
          lrDatos.MoveNext
    Loop
    ActivaDesactivaControles
    SumaColumnas
    frmSeleccion.Enabled = False
    frmCriterios.Enabled = False
    cmdEjecutar.Enabled = False
End If
End Sub

Private Sub cmdGrabar_Click()
Dim loContFunct As NContFunciones
Dim lrPigContrato As NPigContrato
Dim lsMovNro As String
Dim prJoyas As Recordset
Dim filas As Long

If MsgBox("Desea Grabar la Transferencia? ", vbYesNo + vbQuestion + vbDefaultButton1, " Aviso ") = vbYes Then
    cmdGrabar.Enabled = False

    'Genera el Mov Nro
    Set loContFunct = New NContFunciones
         lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set loContFunct = Nothing
    
    'Traslada todo el Flex a un RecordSet
    Set prJoyas = feSeleccionadas.GetRsNew
    filas = CargaMatrix(prJoyas)

    Set lrPigContrato = New NPigContrato
        Call lrPigContrato.nRegistraSelVentaFundicion(fmJoyas, filas, fnTipoProceso, lsMovNro)
    Set lrPigContrato = Nothing
    limpiaInicia
End If
End Sub

Private Sub cmdRegresaTodos_Click()
Do While feSeleccionadas.SumaRow(5) > 0
    feDisponibles.AdicionaFila
    For C = 1 To feSeleccionadas.Cols - 2
        feDisponibles.TextMatrix(feDisponibles.Row, C) = feSeleccionadas.TextMatrix(feSeleccionadas.Row, C)
    Next
    feDisponibles.TextMatrix(feDisponibles.Row, 5) = Format(feDisponibles.TextMatrix(feDisponibles.Row, 5), "##,##0.00")
    feSeleccionadas.EliminaFila (feSeleccionadas.Row)
Loop
ActivaDesactivaControles
SumaColumnas
End Sub

Private Sub cmdRegresaUno_Click()
feDisponibles.AdicionaFila
For C = 1 To feDisponibles.Cols - 1
    feDisponibles.TextMatrix(feDisponibles.Row, C) = feSeleccionadas.TextMatrix(feSeleccionadas.Row, C)
Next
feDisponibles.TextMatrix(feDisponibles.Row, 5) = Format(feDisponibles.TextMatrix(feDisponibles.Row, 5), "##,##0.00")
feSeleccionadas.EliminaFila (feSeleccionadas.Row)
ActivaDesactivaControles
SumaColumnas
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdSelTodos_Click()
Do While feDisponibles.SumaRow(5) > 0
    feSeleccionadas.AdicionaFila
    For C = 1 To feDisponibles.Cols - 1
        feSeleccionadas.TextMatrix(feSeleccionadas.Row, C) = feDisponibles.TextMatrix(feDisponibles.Row, C)
    Next
    feSeleccionadas.TextMatrix(feSeleccionadas.Row, 5) = Format(feSeleccionadas.TextMatrix(feSeleccionadas.Row, 5), "##,##0.00")
    feDisponibles.EliminaFila (feDisponibles.Row)
Loop
ActivaDesactivaControles
SumaColumnas
End Sub

Private Sub cmdSelUno_Click()
feSeleccionadas.AdicionaFila
For C = 1 To feDisponibles.Cols - 1
    feSeleccionadas.TextMatrix(feSeleccionadas.Row, C) = feDisponibles.TextMatrix(feDisponibles.Row, C)
Next
feDisponibles.EliminaFila (feDisponibles.Row)
ActivaDesactivaControles
SumaColumnas
End Sub

Private Sub Form_Activate()
limpiaInicia
ActivaDesactivaControles
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub optCriterios_Click(Index As Integer)
cboRemDesde.Text = ""
cboRemHasta.Text = ""
cboMaterial.Text = ""
cboAgeDesde.Text = ""
cboAgeHasta.Text = ""
lblText2.Caption = ""
txtCtaCod.Text = ""
Select Case Index
            Case 0
                    cboRemDesde.Visible = True
                    cboRemHasta.Visible = True
                    lblText1.Visible = True
                    cboMaterial.Visible = False
                    cboAgeDesde.Visible = False
                    cboAgeHasta.Visible = False
                    lblText2.Visible = False
                    txtCtaCod.Visible = False
                    fnVarCriterio = 0
            Case 1
                    cboRemDesde.Visible = False
                    cboRemHasta.Visible = False
                    lblText1.Visible = False
                    cboMaterial.Visible = True
                    cboAgeDesde.Visible = False
                    cboAgeHasta.Visible = False
                    lblText2.Visible = False
                    txtCtaCod.Visible = False
                    fnVarCriterio = 1
            Case 2
                    cboRemDesde.Visible = False
                    cboRemHasta.Visible = False
                    lblText1.Visible = False
                    cboMaterial.Visible = False
                    cboAgeDesde.Visible = True
                    cboAgeHasta.Visible = True
                    lblText2.Visible = True
                    txtCtaCod.Visible = False
                    fnVarCriterio = 2
            Case 3
                    cboRemDesde.Visible = False
                    cboRemHasta.Visible = False
                    lblText1.Visible = False
                    cboMaterial.Visible = False
                    cboAgeDesde.Visible = False
                    cboAgeHasta.Visible = False
                    lblText2.Visible = False
                    txtCtaCod.Visible = True
                    fnVarCriterio = 3
                    txtCtaCod.SetFocus
End Select
End Sub

Private Sub CargaCombos()
CargaRemates
CargaMaterial
CargaAgencias
End Sub

Private Sub CargaMaterial()
Dim lrDatos As ADODB.Recordset
Dim lrPigContrato As DPigContrato

Set lrDatos = New ADODB.Recordset
Set lrPigContrato = New DPigContrato
      Set lrDatos = lrPigContrato.dEvaluacion(gColocPigMaterial)
Set lrPigContrato = Nothing

Set cboMaterial.RowSource = lrDatos
       cboMaterial.ListField = "cConsDescripcion"
       cboMaterial.BoundColumn = "nConsValor"

Set lrDatos = Nothing
End Sub

Private Sub CargaRemates()
Dim lrDatos As ADODB.Recordset
Dim lrPigFunciones As DPigFunciones

Set lrDatos = New ADODB.Recordset
Set lrPigFunciones = New DPigFunciones
      Set lrDatos = lrPigFunciones.GetListaRemates()
Set lrPigFunciones = Nothing

Set cboRemDesde.RowSource = lrDatos
       cboRemDesde.ListField = "nRemate"
       cboRemDesde.BoundColumn = "nRemate"

Set cboRemHasta.RowSource = lrDatos
       cboRemHasta.ListField = "nRemate"
       cboRemHasta.BoundColumn = "nRemate"

Set lrDatos = Nothing
End Sub

Private Sub CargaAgencias()
Dim lrDatos As ADODB.Recordset
Dim lrPigFunciones As DPigFunciones

Set lrDatos = New ADODB.Recordset
Set lrPigFunciones = New DPigFunciones
      Set lrDatos = lrPigFunciones.GetListaAgencias()
Set lrPigFunciones = Nothing

Set cboAgeDesde.RowSource = lrDatos
       cboAgeDesde.ListField = "cAgeDescripcion"
       cboAgeDesde.BoundColumn = "cAgeCod"

Set cboAgeHasta.RowSource = lrDatos
       cboAgeHasta.ListField = "cAgeDescripcion"
       cboAgeHasta.BoundColumn = "cAgeCod"

Set lrDatos = Nothing
End Sub

Private Sub optSeleccion_Click(Index As Integer)
Select Case Index
            Case 0            '******** Para Venta
                    psCondicionAdic = " And CPEJ.nEstadoJoya in ( " & gPigTipoAdjudicacionCaja & "," & gPigTipoFundicion & ")"
                    fnTipoProceso = gPigTipoVentas
            Case 1            '******** Para Fundición
                    psCondicionAdic = " And CPEJ.nEstadoJoya in ( " & gPigTipoAdjudicacionCaja & "," & gPigTipoVentas & ")"
                    fnTipoProceso = gPigTipoFundicion
End Select
End Sub

Private Sub ActivaDesactivaControles()
Dim X As Integer
If feDisponibles.SumaRow(5) = 0 Then
    cmdSelUno.Enabled = False
    cmdSelTodos.Enabled = False
    feSeleccionadas.SetFocus
Else
    cmdSelUno.Enabled = True
    cmdSelTodos.Enabled = True
    feDisponibles.SetFocus
End If

If feSeleccionadas.SumaRow(5) = 0 Then
    cmdRegresaUno.Enabled = False
    cmdRegresaTodos.Enabled = False
    feDisponibles.SetFocus
    cmdGrabar.Enabled = False
Else
    cmdRegresaUno.Enabled = True
    cmdRegresaTodos.Enabled = True
    feSeleccionadas.SetFocus
    cmdGrabar.Enabled = True
End If
End Sub

Private Sub SumaColumnas()
Dim prJoyas1 As Recordset
Dim i As Integer

lblTotPNetoDisp.Caption = Format(feDisponibles.SumaRow(5), "###,##0.00 ")
If Val(lblTotPNetoDisp.Caption) > 0 Then
    Set prJoyas = feDisponibles.GetRsNew
    lblPzaDisponibles.Caption = Format(CuentaItems(prJoyas), "##,##0 ")
    Set prJoyas = Nothing
Else
    lblPzaDisponibles.Caption = Format("0", "##,##0 ")
End If

lblTotPNetoSel.Caption = Format(feSeleccionadas.SumaRow(5), "###,##0.00 ")
If Val(lblTotPNetoSel.Caption) > 0 Then
    Set prJoyas = feSeleccionadas.GetRsNew
    lblPzaSeleccionadas.Caption = Format(CuentaItems(prJoyas), "##,##0 ")
    Set prJoyas = Nothing
Else
    lblPzaSeleccionadas.Caption = Format("0", "##,##0 ")
End If
End Sub

Private Function CuentaItems(ByVal prJoyas As Recordset) As Long
Dim pnItems As Integer
pnItems = 0
Do While Not prJoyas.EOF
      pnItems = pnItems + 1
      prJoyas.MoveNext
Loop
CuentaItems = pnItems
End Function

Private Function CargaMatrix(ByVal prPiezasSel As Recordset) As Integer
Dim Fila As Integer
ReDim fmJoyas(feSeleccionadas.Rows - 1, 8)

Fila = 0
Do While Not prPiezasSel.EOF
    fmJoyas(Fila, 0) = prPiezasSel!Contrato
    fmJoyas(Fila, 1) = prPiezasSel!Pza
    fmJoyas(Fila, 2) = prPiezasSel!pNeto
    fmJoyas(Fila, 3) = prPiezasSel!nRemate
    fmJoyas(Fila, 4) = prPiezasSel!nValDeuda
    fmJoyas(Fila, 5) = prPiezasSel!nValDeudaComp
    fmJoyas(Fila, 6) = prPiezasSel!nValAdjudicado
    fmJoyas(Fila, 7) = prPiezasSel!VAdicional
    fmJoyas(Fila, 8) = prPiezasSel!OtroProceso
    Fila = Fila + 1
    prPiezasSel.MoveNext
Loop
CargaMatrix = Fila - 1
End Function

Private Sub limpiaInicia()
cboMaterial.Visible = False
cboAgeDesde.Visible = False
cboAgeHasta.Visible = False
lblText2.Visible = False
txtCtaCod.Visible = False
lblTotPNetoDisp.Caption = Format("0.00", "###,##0.00 ")
lblPzaDisponibles.Caption = Format("0", "##,##0 ")
lblTotPNetoSel.Caption = Format("0.00", "###,##0.00 ")
lblPzaSeleccionadas.Caption = Format("0", "##,##0 ")
feDisponibles.Clear
feDisponibles.Rows = 2
feDisponibles.FormaCabecera
feSeleccionadas.Clear
feSeleccionadas.Rows = 2
feSeleccionadas.FormaCabecera
CargaCombos
optCriterios(0).value = True
optCriterios_Click (0)
optSeleccion(0).value = True
optSeleccion_Click (0)
ActivaDesactivaControles
frmSeleccion.Enabled = True
frmCriterios.Enabled = True
cmdEjecutar.Enabled = True
End Sub

Private Sub txtCtaCod_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdEjecutar_Click
End If
End Sub
