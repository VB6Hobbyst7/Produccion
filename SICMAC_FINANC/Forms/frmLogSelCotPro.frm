VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "Tabctl32.ocx"
Begin VB.Form frmLogSelCotPro 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6690
   ClientLeft      =   885
   ClientTop       =   1740
   ClientWidth     =   10080
   Icon            =   "frmLogSelCotPro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   10080
   ShowInTaskbar   =   0   'False
   Begin TabDlg.SSTab sstPropuesta 
      Height          =   3795
      Left            =   195
      TabIndex        =   11
      Top             =   900
      Width           =   9735
      _ExtentX        =   17171
      _ExtentY        =   6694
      _Version        =   393216
      Tabs            =   2
      TabHeight       =   520
      ForeColor       =   8388608
      TabCaption(0)   =   "Técnica"
      TabPicture(0)   =   "frmLogSelCotPro.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).ControlCount=   0
      TabCaption(1)   =   "Económica"
      TabPicture(1)   =   "frmLogSelCotPro.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lblEtiqueta(2)"
      Tab(1).Control(1)=   "fgeBS"
      Tab(1).Control(2)=   "fgeTot"
      Tab(1).ControlCount=   3
      Begin Sicmact.FlexEdit fgeTot 
         Height          =   915
         Left            =   -74925
         TabIndex        =   12
         Top             =   2805
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   1614
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "------Total---Total Prov"
         EncabezadosAnchos=   "400-0-2000-650-900-900-900-900-900-900"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         BackColor       =   -2147483624
         EncabezadosAlineacion=   "C-L-L-L-R-R-R-R-R-R"
         FormatosEdit    =   "0-0-0-0-2-2-2-2-2-2"
         AvanceCeldas    =   1
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin Sicmact.FlexEdit fgeBS 
         Height          =   2220
         Left            =   -74925
         TabIndex        =   13
         Top             =   585
         Width           =   9585
         _ExtentX        =   16907
         _ExtentY        =   3916
         Cols0           =   10
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-Precio-Total-Cant.Prov-Prec.Prov-Total Prov"
         EncabezadosAnchos=   "400-0-2000-650-900-900-900-900-900-900"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Microsoft Sans Serif"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-7-8-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0"
         EncabezadosAlineacion=   "C-L-L-L-R-R-R-R-R-R"
         FormatosEdit    =   "0-0-0-0-2-2-2-2-2-2"
         AvanceCeldas    =   1
         TextArray0      =   "Item"
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         Appearance      =   0
         ColWidth0       =   405
         RowHeight0      =   285
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Detalle"
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
         Height          =   195
         Index           =   2
         Left            =   -74820
         TabIndex        =   14
         Top             =   375
         Width           =   750
      End
   End
   Begin VB.CommandButton cmdCot 
      Caption         =   "&Grabar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   7770
      TabIndex        =   3
      Top             =   5505
      Width           =   1290
   End
   Begin VB.CommandButton cmdCot 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   0
      Left            =   7770
      TabIndex        =   2
      Top             =   4935
      Width           =   1290
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   390
      Left            =   7770
      TabIndex        =   4
      Top             =   6090
      Width           =   1290
   End
   Begin Sicmact.Usuario Usuario 
      Left            =   0
      Top             =   6105
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin Sicmact.FlexEdit fgeCot 
      Height          =   1605
      Left            =   180
      TabIndex        =   1
      Top             =   4950
      Width           =   6915
      _ExtentX        =   12197
      _ExtentY        =   2831
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-Cotización-Codigo-Proveedor-Ok"
      EncabezadosAnchos=   "400-2200-0-3500-350"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-4"
      ListaControles  =   "0-0-0-0-4"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-C"
      FormatosEdit    =   "0-0-0-0-0"
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   285
   End
   Begin Sicmact.TxtBuscar txtSelNro 
      Height          =   285
      Left            =   330
      TabIndex        =   0
      Top             =   555
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   503
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
      TipoBusqueda    =   2
      sTitulo         =   ""
   End
   Begin VB.Label lblProvee 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3345
      TabIndex        =   10
      Top             =   555
      Width           =   4245
   End
   Begin VB.Label lblProvEti 
      Caption         =   "Proveedor"
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
      Height          =   240
      Left            =   3495
      TabIndex        =   9
      Top             =   360
      Width           =   960
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Proceso de selección"
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
      Height          =   210
      Index           =   1
      Left            =   420
      TabIndex        =   8
      Top             =   360
      Width           =   1950
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Area :"
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
      Height          =   210
      Index           =   0
      Left            =   495
      TabIndex        =   7
      Top             =   75
      Width           =   555
   End
   Begin VB.Label lblAreaDes 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   1215
      TabIndex        =   6
      Top             =   45
      Width           =   3705
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Cotizaciones"
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
      Height          =   210
      Index           =   5
      Left            =   285
      TabIndex        =   5
      Top             =   4755
      Width           =   1185
   End
End
Attribute VB_Name = "frmLogSelCotPro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim psFrmTpo As String
Dim psTpoPro As String

Public Sub Inicio(ByVal psFormTpo As String, ByVal psTipoProp As String)
'Presentacion de propuestas [1] o evaluacion [2]
psFrmTpo = psFormTpo
'Propuesta tecnica [1] o economica [2]
psTpoPro = psTipoProp
Me.Show 1
End Sub

Private Sub cmdCot_Click(Index As Integer)
    Dim clsDGnral As DLogGeneral
    Dim clsDMov  As DLogMov
    Dim sSelNro As String, sSelCotNro As String, sSelTraNro As String, sBSCod As String
    Dim sActualiza As String, sProvee As String
    Dim nCont As Integer, nSum As Integer, nResult As Integer
    Dim nCantidad As Currency, nPrecio As Currency
    
    Select Case Index
        Case 0:
            'CANCELAR
            If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                Call Limpiar
                Call CargaTxtSelNro
            End If
        Case 1:
            'GRABACION
            sSelNro = txtSelNro.Text
            For nCont = 1 To fgeCot.Rows - 1
                If fgeCot.TextMatrix(nCont, 4) = "." Then
                    sSelCotNro = fgeCot.TextMatrix(nCont, 1)
                    sProvee = fgeCot.TextMatrix(nCont, 3)
                    Exit For
                End If
            Next
            If sSelNro = "" Or sSelCotNro = "" Then
                MsgBox "Falta seleccionar una cotización ", vbInformation, " Aviso"
                Exit Sub
            End If
            If psFrmTpo = "1" Then
                'PROPUESTA
                If psTpoPro = "1" Then
                    'TECNICA
                    
                ElseIf psTpoPro = "2" Then
                    'ECONOMICA

                    If CCur(IIf(fgeTot.TextMatrix(1, 9) = "", 0, fgeTot.TextMatrix(1, 9))) = 0 Then
                        MsgBox "Falta determiar cantidades y precios", vbInformation, " Aviso"
                        Exit Sub
                    End If
                    'Verifica si cantidades son iguales
                    For nCont = 1 To fgeBS.Rows - 1
                        If CCur(IIf(fgeBS.TextMatrix(nCont, 4) = "", 0, fgeBS.TextMatrix(nCont, 4))) <> CCur(IIf(fgeBS.TextMatrix(nCont, 7) = "", 0, fgeBS.TextMatrix(nCont, 7))) Then
                            nSum = nSum + 1
                        End If
                    Next
                    If nSum > 0 Then
                        If MsgBox("Cantidades ingresadas son diferentes a las solicitadas " & vbCr & "¿ Deseas continuar con la grabación ? ", vbQuestion + vbYesNo, " Aviso ") = vbNo Then
                            Exit Sub
                        End If
                    End If
                    'GRABAR
                    If MsgBox("¿ Estás seguro de Grabar la información ingresada ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                        Set clsDGnral = New DLogGeneral
                        sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        Set clsDGnral = Nothing
                        
                        sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                        Set clsDMov = New DLogMov
                        
                        clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoCotizacion))
                        clsDMov.InsertaMovRef sSelTraNro, sSelNro
                        
                        For nCont = 1 To fgeBS.Rows - 1
                            sBSCod = fgeBS.TextMatrix(nCont, 1)
                            nCantidad = CCur(IIf(fgeBS.TextMatrix(nCont, 7) = "", 0, fgeBS.TextMatrix(nCont, 7)))
                            nPrecio = CCur(IIf(fgeBS.TextMatrix(nCont, 8) = "", 0, fgeBS.TextMatrix(nCont, 8)))
                            clsDMov.ActualizaSelCotDetalle sSelCotNro, sBSCod, nCantidad, nPrecio, sActualiza
                        Next
                        
                        'Ejecuta todos los querys en una transacción
                        'nResult = clsDMov.EjecutaBatch
                        Set clsDMov = Nothing
                        
                        If nResult = 0 Then
                            fgeBS.lbEditarFlex = False
                            cmdCot(0).Enabled = False
                            cmdCot(1).Enabled = False
                            Call CargaTxtSelNro
                        Else
                            MsgBox "Error al grabar la información", vbInformation, " Aviso "
                        End If
                    End If
                End If
            ElseIf psFrmTpo = "2" Then
                'EVALUACION TECNICA Y ECONOMICA
                If MsgBox("¿ Estás seguro de Adjudicar a " & sProvee & vbCr & " el proceso de Selección ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
                    Set clsDGnral = New DLogGeneral
                    sSelTraNro = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDGnral = Nothing
                    
                    sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
                    Set clsDMov = New DLogMov
                    
                    'Inserta MOV - MOVREF
                    clsDMov.InsertaMov sSelTraNro, Trim(Str(gLogOpeSelTramite)), "", Trim(Str(gLogSelEstadoProcAdju))
                    clsDMov.InsertaMovRef sSelTraNro, sSelNro
                    
                    'Actualiza LogSeleccion
                    clsDMov.ActualizaSeleccionAdju sSelNro, sSelCotNro, _
                         sActualiza
                        
                    'Ejecuta todos los querys en una transacción
                    'nResult = clsDMov.EjecutaBatch
                    Set clsDMov = Nothing
                    
                    If nResult = 0 Then
                        fgeCot.lbEditarFlex = False
                        cmdCot(0).Enabled = False
                        cmdCot(1).Enabled = False
                        Call CargaTxtSelNro
                    Else
                        MsgBox "Error al grabar la información", vbInformation, " Aviso "
                    End If
                End If
            End If
        Case Else
            MsgBox "Tipo de comando no reconocido", vbInformation, " Aviso"
    End Select
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub


Private Sub fgeBS_OnCellChange(pnRow As Long, pnCol As Long)
    If psFrmTpo = "1" And psTpoPro = "2" Then
        If pnCol = 7 Or pnCol = 8 Then
            fgeBS.TextMatrix(pnRow, 9) = Format(CCur(IIf(fgeBS.TextMatrix(pnRow, 7) = "", 0, fgeBS.TextMatrix(pnRow, 7))) * CCur(IIf(fgeBS.TextMatrix(pnRow, 8) = "", 0, fgeBS.TextMatrix(pnRow, 8))), "#,##0.00")
            fgeTot.TextMatrix(1, 9) = Format(fgeBS.SumaRow(9), "#,##0.00")
        End If
    End If
End Sub

Private Sub fgeCot_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If psFrmTpo = "2" Then
        If fgeCot.TextMatrix(pnRow, 4) = "." Then
            cmdCot(0).Enabled = True
            cmdCot(1).Enabled = True
        Else
            cmdCot(0).Enabled = False
            cmdCot(1).Enabled = False
        End If
    End If
End Sub

Private Sub fgeCot_OnRowChange(pnRow As Long, pnCol As Long)
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Dim sSelCotNro As String
    Dim nCont  As Integer
    
    If psFrmTpo = "1" Then
        'PROPUESTAS
        'Verifica que siempre este por lo menos UNO
        If fgeCot.TextMatrix(1, 1) = "" Then
            Exit Sub
        End If
        lblProvee.Caption = fgeCot.TextMatrix(fgeCot.Row, 3)
        
        If psTpoPro = "1" Then
            'TECNICA
        
        ElseIf psTpoPro = "2" Then
            'ECONOMICA
            fgeBS.lbEditarFlex = True
            sSelCotNro = fgeCot.TextMatrix(fgeCot.Row, 1)
            cmdCot(0).Enabled = True
            cmdCot(1).Enabled = True
            'Muestra datos
            Set clsDAdq = New DLogAdquisi
            Set rs = New ADODB.Recordset
            Set rs = clsDAdq.CargaSelCotDetalle(SelCotDetUnRegistro, sSelCotNro)
            If rs.RecordCount > 0 Then
                Set fgeBS.Recordset = rs
            End If
            Set rs = Nothing
            Set clsDAdq = Nothing
            'Color en las COLUMNAS
            For nCont = 1 To fgeBS.Rows - 1
                fgeBS.Row = nCont
                fgeBS.Col = 7
                fgeBS.CellForeColor = vbBlue '&HC0FFC0 '(verde)
                fgeBS.Col = 8
                fgeBS.CellForeColor = vbBlue '&HC0FFC0 '(verde)
            Next
            fgeTot.TextMatrix(1, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
            fgeTot.TextMatrix(1, 9) = Format(fgeBS.SumaRow(9), "#,##0.00")
        End If
    End If
End Sub

Private Sub Form_Load()
    Call CentraForm(Me)
    'Carga información de la relación usuario-area
    Usuario.Inicio gsCodUser
    If Len(Usuario.AreaCod) = 0 Then
        MsgBox "Usuario no determinado", vbInformation, "Aviso"
        Exit Sub
    End If
    lblAreaDes.Caption = Usuario.AreaNom
    
    If psFrmTpo = "1" Then
        If psTpoPro = "1" Then
            Me.Caption = "Propuesta Técnica"
            sstPropuesta.TabVisible(1) = False
        ElseIf psTpoPro = "2" Then
            Me.Caption = "Propuesta Económica"
            sstPropuesta.TabVisible(0) = False
        End If
        fgeCot.EncabezadosAnchos = "400-2200-0-3500-0"
    ElseIf psFrmTpo = "2" Then
        Me.Caption = "Evaluación de Propuestas"
        lblProvEti.Visible = False
        lblProvee.Visible = False
    Else
        MsgBox "Tipo de formulario no reconocido", vbInformation, " Aviso"
    End If
    
    Call CargaTxtSelNro
End Sub

Private Sub CargaTxtSelNro()
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    
    Set rs = clsDAdq.CargaSeleccion(SelTodosEstado, , gLogSelEstadoCotizacion)
    If rs.RecordCount > 0 Then
        txtSelNro.rs = rs
    Else
        txtSelNro.Enabled = False
    End If
    Set rs = Nothing
    Set clsDAdq = Nothing
End Sub

Private Sub txtSelNro_EmiteDatos()
    Dim clsDAdq As DLogAdquisi
    Dim rs As ADODB.Recordset
    Dim sSelNro As String, sAdqNro As String
    Dim sBSCod As String, sLogSelCotNroAnt As String
    Dim nCont As Integer, pColIni As Integer
    
    
    If txtSelNro.Ok = False Then Exit Sub
    
    Set clsDAdq = New DLogAdquisi
    Set rs = New ADODB.Recordset
    Call Limpiar
    sSelNro = txtSelNro.Text
    
    Set rs = clsDAdq.CargaSelCotiza(sSelNro)
    If rs.RecordCount > 0 Then
        Set fgeCot.Recordset = rs
        If psFrmTpo = "1" Then
            'PROPUESTA
            If psTpoPro = "1" Then
                'TECNICA
            ElseIf psTpoPro = "2" Then
                'ECONOMICA
                Call fgeCot_OnRowChange(fgeCot.Row, fgeCot.Col)
            End If
        ElseIf psFrmTpo = "2" Then
            'EVALUACION TECNICA Y ECONOMICA
            Set rs = clsDAdq.CargaSelCotDetalle(SelCotDetUnRegEvalua, "", sSelNro)
            If rs.RecordCount > 0 Then
                fgeCot.lbEditarFlex = True

                'Base izquierda del FLEX
                Set fgeBS.Recordset = rs
                fgeTot.TextMatrix(1, 6) = Format(fgeBS.SumaRow(6), "#,##0.00")
                'Carga los detalles
                Set rs = clsDAdq.CargaSelCotDetalle(SelCotDetUnRegEvaluaTodos, "", sSelNro)
                pColIni = 7
                fgeBS.Cols = (fgeBS.Cols - 3) + ((fgeCot.Rows - 1) * 3)
                fgeTot.Cols = fgeBS.Cols
                fgeTot.TextMatrix(1, 2) = "T O T A L E S"
                For nCont = 1 To fgeCot.Rows - 1
                    fgeBS.TextMatrix(0, pColIni) = "P" & fgeCot.TextMatrix(nCont, 0) & " - Cantidad"
                    fgeBS.TextMatrix(0, pColIni + 1) = "P" & fgeCot.TextMatrix(nCont, 0) & " - Precio"
                    fgeBS.TextMatrix(0, pColIni + 2) = "P" & fgeCot.TextMatrix(nCont, 0) & " - Total"
                    
                    fgeTot.TextMatrix(0, pColIni + 2) = "P" & fgeCot.TextMatrix(nCont, 0) & " - Total"
                    pColIni = pColIni + 3
                Next
                
                For nCont = 1 To fgeBS.Rows - 1
                    sBSCod = fgeBS.TextMatrix(nCont, 1)
                    pColIni = 7
                    rs.MoveFirst
                    sLogSelCotNroAnt = rs!cLogSelCotNro
                    Do While Not rs.EOF
                        If sLogSelCotNroAnt <> rs!cLogSelCotNro Then
                            pColIni = pColIni + 3
                            sLogSelCotNroAnt = rs!cLogSelCotNro
                        End If
                        If sBSCod = rs!cBSCod Then
                            fgeBS.Row = nCont
                            fgeBS.Col = pColIni
                            fgeBS.CellForeColor = vbBlue
                            fgeBS.Col = pColIni + 1
                            fgeBS.CellForeColor = vbBlue
                            fgeBS.TextMatrix(nCont, pColIni) = Format(rs!nLogSelCotDetCantidad, "#,##0.00")
                            fgeBS.TextMatrix(nCont, pColIni + 1) = Format(rs!nLogSelCotDetPrecio, "#,##0.00")
                            fgeBS.TextMatrix(nCont, pColIni + 2) = Format(rs!Total, "#,##0.00")
                        End If
                        rs.MoveNext
                    Loop
                Next
            End If
            pColIni = 9
            For nCont = 1 To (fgeCot.Rows - 1)
                fgeTot.TextMatrix(1, pColIni) = Format(fgeBS.SumaRow(pColIni), "#,##0.00")
                pColIni = pColIni + 3
            Next
        End If
    End If
End Sub

Private Sub Limpiar()
    lblProvee.Caption = ""
    fgeCot.Clear
    fgeCot.FormaCabecera
    fgeCot.Rows = 2
    fgeBS.Clear
    fgeBS.FormaCabecera
    fgeBS.Rows = 2
    fgeTot.Clear
    fgeTot.FormaCabecera
    fgeTot.Rows = 2
End Sub
