VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogMantBienesAdjud 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento: Tasación de Bienes Adjudicados"
   ClientHeight    =   8085
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13350
   Icon            =   "frmLogMantBienesAdjud.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8085
   ScaleWidth      =   13350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame2 
      Caption         =   "Historial de Monto Venta"
      ForeColor       =   &H00000080&
      Height          =   2535
      Left            =   9840
      TabIndex        =   38
      Top             =   3600
      Width           =   3255
      Begin Sicmact.FlexEdit grdVenta 
         Height          =   2175
         Left            =   120
         TabIndex        =   39
         Top             =   240
         Width           =   3015
         _ExtentX        =   5318
         _ExtentY        =   3836
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Fecha-Valor"
         EncabezadosAnchos=   "400-1200-1200"
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
         ColumnasAEditar =   "X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-R"
         FormatosEdit    =   "0-5-2"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         SelectionMode   =   1
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame fraDetalle 
      ForeColor       =   &H8000000B&
      Height          =   4455
      Left            =   120
      TabIndex        =   2
      Top             =   3480
      Width           =   13095
      Begin VB.Frame fraActualizaVenta 
         Caption         =   "Datos Monto Venta"
         ForeColor       =   &H00FF0000&
         Height          =   735
         Left            =   2040
         TabIndex        =   31
         Top             =   3600
         Width           =   7575
         Begin MSMask.MaskEdBox txtFecVenta 
            Height          =   255
            Left            =   720
            TabIndex        =   36
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   450
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.CommandButton cmdCancelaVenta 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6120
            TabIndex        =   34
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdGrabarVenta 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   4800
            TabIndex        =   33
            Top             =   240
            Width           =   1215
         End
         Begin Sicmact.EditMoney txtValVenta 
            Height          =   300
            Left            =   3120
            TabIndex        =   37
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
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
            Text            =   "0.00"
            Enabled         =   -1  'True
         End
         Begin VB.Label Label3 
            Caption         =   "Fecha:"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lblEtqValorVenta 
            Caption         =   "Valor:"
            Height          =   255
            Left            =   2040
            TabIndex        =   32
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdActVenta 
         Caption         =   "Actualizar Venta"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Frame fraNuevaTasacion 
         Caption         =   "Datos Tasación"
         ForeColor       =   &H000000C0&
         Height          =   735
         Left            =   2040
         TabIndex        =   21
         Top             =   2760
         Width           =   7575
         Begin VB.CommandButton cmdCancela 
            Caption         =   "Cancelar"
            Height          =   375
            Left            =   6120
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
         Begin VB.CommandButton cmdGrabar 
            Caption         =   "Grabar"
            Height          =   375
            Left            =   4800
            TabIndex        =   26
            Top             =   240
            Width           =   1215
         End
         Begin Sicmact.EditMoney txtValTasacion 
            Height          =   300
            Left            =   3120
            TabIndex        =   25
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   529
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
            Text            =   "0.00"
            Enabled         =   -1  'True
         End
         Begin MSMask.MaskEdBox txtFecTasacion 
            Height          =   300
            Left            =   720
            TabIndex        =   22
            Top             =   240
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label13 
            Caption         =   "Fecha:"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lblEtqValor 
            Caption         =   "Valor:"
            Height          =   255
            Left            =   2040
            TabIndex        =   23
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.CommandButton cmdNuevaTasacion 
         Caption         =   "Actualizar Tasación"
         Height          =   375
         Left            =   120
         TabIndex        =   20
         Top             =   2880
         Width           =   1575
      End
      Begin VB.Frame Frame4 
         Caption         =   "Detalle de Bien Adjudicado"
         ForeColor       =   &H00000080&
         Height          =   2535
         Left            =   120
         TabIndex        =   7
         Top             =   120
         Width           =   6015
         Begin VB.Label lblValVenta 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4200
            TabIndex        =   29
            Top             =   2160
            Width           =   1695
         End
         Begin VB.Label lblEtqValVenta 
            Caption         =   "Valor Venta:"
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   2760
            TabIndex        =   28
            Top             =   2160
            Width           =   1335
         End
         Begin VB.Label lblValAdjud 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4200
            TabIndex        =   19
            Top             =   1800
            Width           =   1695
         End
         Begin VB.Label lblEtqValAdjud 
            Caption         =   "Valor Adjudic:"
            Height          =   255
            Left            =   2760
            TabIndex        =   18
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label lblTipoBien 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   17
            Top             =   1440
            Width           =   4815
         End
         Begin VB.Label Label9 
            Caption         =   "Tipo Bien"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1440
            Width           =   855
         End
         Begin VB.Label lblFecAdjud 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   15
            Top             =   1800
            Width           =   1455
         End
         Begin VB.Label Label7 
            Caption         =   "Fec.Adjudic:"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   1800
            Width           =   975
         End
         Begin VB.Label lblNumBien 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   1080
            TabIndex        =   13
            Top             =   240
            Width           =   1455
         End
         Begin VB.Label lblPElectronica 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4200
            TabIndex        =   12
            Top             =   240
            Width           =   1695
         End
         Begin VB.Label Label4 
            Caption         =   "P.Electronica:"
            Height          =   255
            Left            =   2760
            TabIndex        =   11
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblDescripcion 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   735
            Left            =   1080
            TabIndex        =   10
            Top             =   600
            Width           =   4815
         End
         Begin VB.Label Label2 
            Caption         =   "Descripción:"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   855
         End
         Begin VB.Label Label1 
            Caption         =   "N°"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Historial de Tasación"
         ForeColor       =   &H00000080&
         Height          =   2535
         Left            =   6240
         TabIndex        =   3
         Top             =   120
         Width           =   3375
         Begin Sicmact.FlexEdit grdTasacion 
            Height          =   2175
            Left            =   120
            TabIndex        =   4
            Top             =   240
            Width           =   3135
            _ExtentX        =   5530
            _ExtentY        =   3836
            Cols0           =   3
            HighLight       =   1
            AllowUserResizing=   1
            RowSizingMode   =   1
            EncabezadosNombres=   "N°-Fecha-Valor"
            EncabezadosAnchos=   "400-1200-1200"
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
            ColumnasAEditar =   "X-X-X"
            TextStyleFixed  =   4
            ListaControles  =   "0-0-0"
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            BackColorControl=   -2147483643
            EncabezadosAlineacion=   "C-R-R"
            FormatosEdit    =   "0-5-2"
            AvanceCeldas    =   1
            TextArray0      =   "N°"
            SelectionMode   =   1
            lbEditarFlex    =   -1  'True
            lbUltimaInstancia=   -1  'True
            lbFormatoCol    =   -1  'True
            lbBuscaDuplicadoText=   -1  'True
            ColWidth0       =   405
            RowHeight0      =   300
            ForeColorFixed  =   -2147483630
         End
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Bienes Adjudicados"
      ForeColor       =   &H00000080&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13095
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   11040
         TabIndex        =   6
         Top             =   720
         Width           =   1455
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "Aplicar"
         Height          =   375
         Left            =   11040
         TabIndex        =   5
         Top             =   240
         Width           =   1455
      End
      Begin Sicmact.FlexEdit grdBienesAdjudicados 
         Height          =   3015
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   10695
         _ExtentX        =   18865
         _ExtentY        =   5318
         Cols0           =   8
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   "N°-Agencia-Num.Adju-Descripción-Fec.Adju-Valor Adju.-Capital-Int. y Otros"
         EncabezadosAnchos=   "400-0-800-4000-1200-1200-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-R-R-R-R"
         FormatosEdit    =   "0-0-0-5-5-2-2-2"
         AvanceCeldas    =   1
         TextArray0      =   "N°"
         SelectionMode   =   1
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmLogMantBienesAdjud"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdAplicar_Click()
    Dim NumAdj As Integer
    If grdBienesAdjudicados.row >= 1 And grdBienesAdjudicados.TextMatrix(grdBienesAdjudicados.row, 2) <> "" Then
        NumAdj = grdBienesAdjudicados.TextMatrix(grdBienesAdjudicados.row, 2)
        If NumAdj >= 0 Then
            Me.fraDetalle.Enabled = True
            Call ObtenerDatosBienAdjud(NumAdj)
        Else
            MsgBox "Debe seleccionar un bien adjudicado para registrar venta", vbInformation + vbOKOnly, "SICMACM"
            Exit Sub
        End If
    Else
        MsgBox "No existen datos en el listado", vbInformation + vbOKOnly, "SICMACM"
    End If
    Me.fraNuevaTasacion.Visible = False 'NAGL 20180618
    Me.fraActualizaVenta.Visible = False 'NAGL 20180618
End Sub
Private Sub CargaDatos()
    Dim oCnt As DLogBieSer
    Dim rs As ADODB.Recordset
    Set oCnt = New DLogBieSer
    
    Set rs = oCnt.ObtenerListaBienesAdjudicado()
    If rs.EOF Then MsgBox "No se encontraron datos.", vbInformation, "Mensaje"
    grdBienesAdjudicados.Clear
    grdBienesAdjudicados.FormaCabecera
    grdBienesAdjudicados.Rows = 2
    grdBienesAdjudicados.rsFlex = rs
    Set oCnt = Nothing
End Sub
Private Sub ObtenerDatosBienAdjud(ByVal nNumAdjudicacion As Integer)
    Dim oCnt As DLogBieSer
    Set oCnt = New DLogBieSer
    Dim prs As ADODB.Recordset
    Set prs = oCnt.ObtenerDatosBienAdjudicado(nNumAdjudicacion)

    If Not prs.EOF Then
        Me.lblNumBien.Caption = Format(prs!nBienAdjudCod, "000000")
        lblPElectronica.Caption = prs!cNumPartElectronica
        lblTipoBien.Caption = prs!cTipoBien
        lblDescripcion.Caption = prs!cDescripcion
        lblValAdjud.Caption = Format(prs!nValAdjudicacion, "#,###,##0.00")
        lblFecAdjud.Caption = Format(prs!dFecAdjudicacion, "dd/MM/yyyy")
        lblValVenta.Caption = Format(prs!nMontoPreVenta, "#,###,##0.00") 'JUCS TI-ERS002-2017
        If prs!nMoneda = 1 Then
            Me.lblEtqValor.Caption = "Valor: (S/)"
            Me.lblEtqValAdjud.Caption = "Valor Adjud:(S/)"
            Me.lblEtqValVenta.Caption = "Valor Venta: (S/)"  'JUCS TI-ERS002-2017
            Me.lblEtqValorVenta.Caption = "Valor: (S/)" 'JUCS TI-ERS002-2017
        Else
            Me.lblEtqValor.Caption = "Valor: (US$)"
            Me.lblEtqValAdjud.Caption = "Valor Adjud:(US$)"
            Me.lblEtqValVenta.Caption = "Valor Venta: (US$)"  'JUCS TI-ERS002-2017
            Me.lblEtqValorVenta.Caption = "Valor: (US$)" 'JUCS TI-ERS002-2017
        End If
    End If
    
    Set prs = oCnt.ListarTasasBienAdjudicado(nNumAdjudicacion)
    grdTasacion.Clear
    grdTasacion.FormaCabecera
    If Not prs.EOF Then
        grdTasacion.Rows = 2
        grdTasacion.rsFlex = prs
    End If
    
    'JUCS TI-ERS002-2017
   Set prs = oCnt.ListarMontoVentaBienAdjudicado(nNumAdjudicacion)
    grdVenta.Clear
    grdVenta.FormaCabecera
    'If Not prs.EOF Then
        grdVenta.Rows = 2
        grdVenta.rsFlex = prs
    'End If Comentado by NAGL 20180618
    'END JUCS TI-ERS002-2017
    Set oCnt = Nothing
    Set prs = Nothing
End Sub

Private Sub cmdCancela_Click()
    fraNuevaTasacion.Visible = False
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub
 'JUCS TI-ERS002-2017
Private Sub cmdCancelaVenta_Click()
    fraActualizaVenta.Visible = False
End Sub
Private Sub cmdGrabarVenta_Click()
    Dim oFunciones As NContFunciones
    Dim oCnt As DLogBieSer
    Set oCnt = New DLogBieSer
    Dim prs As ADODB.Recordset
    Dim sMovNro As String
    If Not IsDate(txtFecVenta.Text) Then
        MsgBox "La fecha ingresada no es correcta", vbExclamation, "Aviso"
        Me.txtFecVenta.SetFocus
        Exit Sub
    End If
    If Me.txtValVenta.Text <= 0 Then
        MsgBox "El valor de la venta no puede ser menor que cero", vbExclamation, "Aviso"
        Me.txtValVenta.SetFocus
        Exit Sub
    End If
    If valFecha(txtFecVenta) Then 'NAGL Cambió de grdTasacion.TextMatrix(1, 1) a valFecha(txtFecVenta)
        If grdVenta.TextMatrix(1, 1) <> "" Then
            If DateDiff("d", CDate(grdVenta.TextMatrix(1, 1)), CDate(Me.txtFecVenta.Text)) < 0 Then 'NAGL Cambió de grdTasacion a grdVenta
                MsgBox "La fecha ingresada no puede ser menor a la fecha del último monto de la Venta", vbExclamation, "Aviso"
                Me.txtFecVenta.SetFocus
                Exit Sub
            End If
        End If
    Else
        Exit Sub 'NAGL 20180618
    End If
    If MsgBox("¿Está seguro de haber ingresado correctamente los datos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set oFunciones = New NContFunciones
        sMovNro = oFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set oFunciones = Nothing
        
        Call oCnt.RegistrarVentaBienAdjudicado(lblNumBien.Caption, txtFecVenta, txtValVenta, sMovNro)
        Set oCnt = Nothing
        Set prs = Nothing
        Set oCnt = New DLogBieSer
        Set prs = oCnt.ListarTasasBienAdjudicado(nNumAdjudicacion)
        If Not prs.EOF Then
            grdTasacion.Clear
            grdTasacion.FormaCabecera
            grdTasacion.Rows = 2
            grdTasacion.rsFlex = prs
        End If
        Call LimpiarControles
        Me.fraDetalle.Enabled = False
        Me.fraNuevaTasacion.Visible = False
        Me.fraActualizaVenta.Visible = False 'NAGL 20180618
        MsgBox "El registro se realizó de forma exitosa"
    End If
    Set oCnt = Nothing
    Set prs = Nothing
End Sub
'END JUCS-ERS002-2017
Private Sub cmdGrabar_Click()
    Dim oFunciones As NContFunciones
    Dim oCnt As DLogBieSer
    Set oCnt = New DLogBieSer
    Dim prs As ADODB.Recordset
    Dim sMovNro As String
    If Not IsDate(txtFecTasacion.Text) Then
        MsgBox "La fecha ingresada no es correcta", vbExclamation, "Aviso"
        Me.txtFecTasacion.SetFocus
        Exit Sub
    End If
    If Me.txtValTasacion.Text <= 0 Then
        MsgBox "El valor de la tasación no puede ser menor que cero", vbExclamation, "Aviso"
        Me.txtValTasacion.SetFocus
        Exit Sub
    End If
    If IsDate(grdTasacion.TextMatrix(1, 1)) = True Then
        If DateDiff("d", CDate(grdTasacion.TextMatrix(1, 1)), CDate(Me.txtFecTasacion.Text)) < 0 Then
            MsgBox "La fecha ingresada no puede ser menor a la fecha de la última tasación", vbExclamation, "Aviso"
            Me.txtFecTasacion.SetFocus
            Exit Sub
        End If
    End If
    If MsgBox("¿Está seguro de haber ingresado correctamente los datos?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set oFunciones = New NContFunciones
        sMovNro = oFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set oFunciones = Nothing
        
        Call oCnt.RegistrarTasaBienAdjudicado(lblNumBien.Caption, txtFecTasacion, txtValTasacion, sMovNro)
        Set oCnt = Nothing
        Set prs = Nothing
        Set oCnt = New DLogBieSer
        Set prs = oCnt.ListarTasasBienAdjudicado(nNumAdjudicacion)
        If Not prs.EOF Then
            grdTasacion.Clear
            grdTasacion.FormaCabecera
            grdTasacion.Rows = 2
            grdTasacion.rsFlex = prs
        End If
        Call LimpiarControles
        Me.fraDetalle.Enabled = False
        Me.fraNuevaTasacion.Visible = False
        Me.fraActualizaVenta.Visible = False 'NAGL 20180618
        MsgBox "El registro se realizó de forma exitosa"
    End If
    Set oCnt = Nothing
    Set prs = Nothing
End Sub
Sub LimpiarControles()
    Me.lblDescripcion.Caption = ""
    Me.lblEtqValAdjud.Caption = "Valor Adjud:"
    Me.lblEtqValor.Caption = "Valor:"
    Me.lblEtqValVenta.Caption = "Valor Venta:" 'JUCS TI-ERS002-2017
    Me.lblEtqValorVenta.Caption = "Valor:"  'JUCS TI-ERS002-2017
    Me.lblFecAdjud.Caption = ""
    Me.lblNumBien.Caption = ""
    Me.lblPElectronica.Caption = ""
    Me.lblTipoBien.Caption = ""
    Me.lblValAdjud.Caption = ""
    Me.lblValVenta.Caption = "" 'JUCS TI-ERS002-2017
    Me.txtValVenta.Text = "0.00" 'JUCS TI-ERS002-2017
    Me.txtFecTasacion.Text = "__/__/____"
    Me.txtFecVenta.Text = "__/__/____"  'JUCS TI-ERS002-2017
    Me.txtValTasacion.Text = "0.00"
    grdVenta.Clear  'JUCS TI-ERS002-2017
    grdVenta.FormaCabecera  'JUCS TI-ERS002-2017
    grdVenta.Rows = 2  'JUCS TI-ERS002-2017
    grdTasacion.Clear
    grdTasacion.FormaCabecera
    grdTasacion.Rows = 2
    grdVenta.Clear  'JUCS TI-ERS002-2017
    grdVenta.FormaCabecera  'JUCS TI-ERS002-2017
    
End Sub
Private Sub cmdNuevaTasacion_Click()
    fraNuevaTasacion.Visible = True
    Me.txtFecTasacion.SetFocus
    txtFecTasacion.Text = "__/__/____" 'NAGL 20180618
    txtValTasacion.Text = "0.00" 'NAGL 20180618
End Sub
'JUCS TI-ERS002-2018
Private Sub txtFecVenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtValVenta.SetFocus 'NAGL Agregó txtValVenta
    End If
End Sub
'JUCS TI-ERS002-2018

'JUCS TI-ERS002-2017
Private Sub txtValVenta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGrabarVenta.SetFocus 'NAGL Agregó cmdGrabarVenta
    End If
End Sub
'FIN JUCS TI-ERS002-2017

'JUCS TI-ERS002-2017
Private Sub cmdActVenta_Click()
fraActualizaVenta.Visible = True
Me.txtFecVenta.SetFocus 'NAGL 20180618
txtFecVenta.Text = "__/__/____" 'NAGL 20180618
txtValVenta.Text = "0.00" 'NAGL 20180618
End Sub
'END JUCS TI-ERS002-2017
Private Sub Form_Load()
    Me.fraDetalle.Enabled = False
    fraNuevaTasacion.Visible = False
    fraActualizaVenta.Visible = False 'JUCS TI-ERS002-2017
    Call CargaDatos
End Sub
Private Sub txtFecTasacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtValTasacion.SetFocus
    End If
End Sub
Private Sub txtValTasacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGrabar.SetFocus
    End If
End Sub

