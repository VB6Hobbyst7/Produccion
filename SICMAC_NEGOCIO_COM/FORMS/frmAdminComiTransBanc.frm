VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmAdminComiTransBanc 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5835
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10785
   FillStyle       =   0  'Solid
   Icon            =   "frmAdminComiTransBanc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5835
   ScaleWidth      =   10785
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab tbComisiones 
      Height          =   5715
      Left            =   30
      TabIndex        =   29
      Top             =   30
      Width           =   10695
      _ExtentX        =   18865
      _ExtentY        =   10081
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "Nuevas Comisiones"
      TabPicture(0)   =   "frmAdminComiTransBanc.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label6"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label7"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label8"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label9"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label10"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label11"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label12"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label13"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtDesde"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtHasta"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtBuscarBanco"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtBanco"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cbTipo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cbMoneda"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "cbTipoCalculo"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "cbMonedaComision"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdAceptar"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "cmdCancelar"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdEditar"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdQuitar"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "cmdDefinir"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdSalir"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cbPlaza"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "txtComision"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtMontoMinimo"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "grdComision"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).ControlCount=   31
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdComision 
         Height          =   2895
         Left            =   240
         TabIndex        =   22
         Top             =   2280
         Width           =   10215
         _ExtentX        =   18018
         _ExtentY        =   5106
         _Version        =   393216
         Rows            =   3
         Cols            =   12
         FixedRows       =   2
         AllowBigSelection=   0   'False
         FocusRect       =   2
         GridLines       =   2
         GridLinesUnpopulated=   3
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   12
         _Band(0).GridLinesBand=   2
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   3
      End
      Begin SICMACT.EditMoney txtMontoMinimo 
         Height          =   315
         Left            =   4320
         TabIndex        =   26
         Top             =   5295
         Width           =   1095
         _extentx        =   1931
         _extenty        =   556
         font            =   "frmAdminComiTransBanc.frx":0326
         text            =   "0"
         enabled         =   -1
      End
      Begin SICMACT.EditMoney txtComision 
         Height          =   315
         Left            =   7680
         TabIndex        =   19
         Top             =   1800
         Width           =   855
         _extentx        =   1508
         _extenty        =   556
         font            =   "frmAdminComiTransBanc.frx":0352
         text            =   "0"
         enabled         =   -1
      End
      Begin VB.ComboBox cbPlaza 
         Height          =   315
         Left            =   8760
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   840
         Width           =   1695
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "Salir"
         Height          =   315
         Left            =   9360
         TabIndex        =   28
         Top             =   5280
         Width           =   1095
      End
      Begin VB.CommandButton cmdDefinir 
         Caption         =   "Definir"
         Height          =   315
         Left            =   5520
         TabIndex        =   27
         Top             =   5280
         Width           =   975
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   315
         Left            =   1320
         TabIndex        =   24
         Top             =   5280
         Width           =   975
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   315
         Left            =   240
         TabIndex        =   23
         Top             =   5280
         Width           =   975
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   315
         Left            =   9240
         TabIndex        =   21
         Top             =   1800
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Height          =   315
         Left            =   9240
         TabIndex        =   20
         Top             =   1440
         Width           =   1215
      End
      Begin VB.ComboBox cbMonedaComision 
         Height          =   315
         Left            =   6000
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox cbTipoCalculo 
         Height          =   315
         Left            =   4080
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox cbMoneda 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1800
         Width           =   1455
      End
      Begin VB.ComboBox cbTipo 
         Height          =   315
         Left            =   6450
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.TextBox txtBanco 
         Height          =   315
         Left            =   2040
         Locked          =   -1  'True
         TabIndex        =   2
         Top             =   840
         Width           =   3855
      End
      Begin SICMACT.TxtBuscar txtBuscarBanco 
         Height          =   315
         Left            =   240
         TabIndex        =   1
         Top             =   840
         Width           =   1695
         _extentx        =   2990
         _extenty        =   556
         appearance      =   1
         appearance      =   1
         font            =   "frmAdminComiTransBanc.frx":037E
         appearance      =   1
         stitulo         =   ""
      End
      Begin SICMACT.EditMoney txtHasta 
         Height          =   315
         Left            =   3000
         TabIndex        =   13
         Top             =   1800
         Width           =   735
         _extentx        =   1296
         _extenty        =   556
         font            =   "frmAdminComiTransBanc.frx":03AA
         text            =   "0"
         enabled         =   -1
      End
      Begin SICMACT.EditMoney txtDesde 
         Height          =   315
         Left            =   2040
         TabIndex        =   12
         Top             =   1800
         Width           =   735
         _extentx        =   1296
         _extenty        =   556
         font            =   "frmAdminComiTransBanc.frx":03D6
         text            =   "0"
         enabled         =   -1
      End
      Begin VB.Label Label13 
         Caption         =   "Monto Mínimo S/."
         Height          =   255
         Left            =   2880
         TabIndex        =   25
         Top             =   5355
         Width           =   1455
      End
      Begin VB.Label Label12 
         Caption         =   "Comisión:"
         Height          =   255
         Left            =   7680
         TabIndex        =   18
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   6000
         TabIndex        =   31
         Top             =   1560
         Width           =   735
      End
      Begin VB.Label Label10 
         Caption         =   "Tipo Cálculo:"
         Height          =   255
         Left            =   4080
         TabIndex        =   15
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label9 
         Caption         =   "Condiciones de la comisión"
         Height          =   255
         Left            =   4080
         TabIndex        =   14
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "_"
         Height          =   255
         Left            =   2775
         TabIndex        =   30
         Top             =   1755
         Width           =   255
      End
      Begin VB.Label Label7 
         Caption         =   "Rango:"
         Height          =   255
         Left            =   2040
         TabIndex        =   11
         Top             =   1560
         Width           =   615
      End
      Begin VB.Label Label6 
         Caption         =   "Moneda:"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "Montos de la Operación"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   1320
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Plaza:"
         Height          =   255
         Left            =   8260
         TabIndex        =   6
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label3 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   6000
         TabIndex        =   3
         Top             =   900
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Características de la Operación"
         Height          =   255
         Left            =   6000
         TabIndex        =   4
         Top             =   480
         Width           =   2655
      End
      Begin VB.Label Label1 
         Caption         =   "Banco"
         Height          =   255
         Left            =   240
         TabIndex        =   0
         Top             =   480
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmAdminComiTransBanc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Option Explicit
'
'Private Operacion As Integer   '1 : Registrar     2: Editar
'Public nPermiso As Integer     '1 : Mantenimiento 2: Consulta
'Private nIdComision As Integer
'
'Private Sub cbMoneda_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If txtDesde.Enabled Then txtDesde.SetFocus
'    End If
'End Sub
'
'Private Sub cbMonedaComision_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If txtComision.Enabled Then txtComision.SetFocus
'    End If
'End Sub
'
'Private Sub cbPlaza_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If cbMoneda.Enabled Then cbMoneda.SetFocus
'    End If
'End Sub
'
'Private Sub cbTipo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If cbPlaza.Enabled Then cbPlaza.SetFocus
'    End If
'End Sub
'
'Private Sub cbTipoCalculo_Click()
'txtComision.value = 0
'End Sub
'
'Private Sub cbTipoCalculo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If cbMonedaComision.Enabled Then cbMonedaComision.SetFocus
'    End If
'End Sub
'
'Private Sub CmdAceptar_Click()
'
'    Dim nIdComision As String
'    Dim cidBanco As String
'    Dim nTipoOpe As Integer
'    Dim nPlaza As Integer
'    Dim nMonedaOpe As Integer
'    Dim nTipoCalc As Integer
'    Dim nMonedaComi As Integer
'    Dim nEstado As Integer
'    Dim nDesde As Double
'    Dim nHasta As Double
'    Dim nMontoComi As Double
'    Dim nMontoMinimo As Double
'    Dim sMensaje As String
'
'    On Error GoTo error
'
'    If Validacion Then
'
'        If Operacion = 1 Then
'            sMensaje = "¿Está seguro de Agregar una NUEVA comisión?"
'        ElseIf Operacion = 2 Then
'            sMensaje = "¿Está seguro de Actualizar los datos de la comisión seleccionada?"
'        End If
'
'        If MsgBox(sMensaje, vbQuestion + vbYesNo, "Aviso") = vbYes Then
'
'            cidBanco = Mid(txtBuscarBanco.Text, 4, 13)
'            nTipoOpe = val(Trim(Right(cbTipo.Text, 5)))
'            nPlaza = val(Trim(Right(cbPlaza.Text, 5)))
'            nMonedaOpe = val(Trim(Right(cbMoneda.Text, 5)))
'            nDesde = txtDesde.value
'            nHasta = txtHasta.value
'            nTipoCalc = val(Trim(Right(cbTipoCalculo.Text, 5)))
'            nMonedaComi = val(Trim(Right(cbMonedaComision.Text, 5)))
'            nMontoComi = txtComision.value
'            nMontoMinimo = txtMontoMinimo.value
'            nIdComision = val(grdComision.TextMatrix(grdComision.row, 11))
'            nEstado = 1
'
'            Dim oDeficion As New COMNCaptaGenerales.NCOMCaptaDefinicion
'
'            If Operacion = 1 Then
'                Call oDeficion.RegistraComisionTransferencia(cidBanco, nTipoOpe, nPlaza, _
'                nMonedaOpe, nDesde, nHasta, nTipoCalc, nMonedaComi, nMontoComi, nMontoMinimo, nEstado)
'
'            ElseIf Operacion = 2 Then
'                Call oDeficion.ActualizaComisionTransferencia(nIdComision, cidBanco, nTipoOpe, nPlaza, _
'                nMonedaOpe, nDesde, nHasta, nTipoCalc, nMonedaComi, nMontoComi, nMontoMinimo, nEstado)
'
'            Else
'
'            End If
'            Limpiar
'            ListarComisiones
'            MsgBox "Los datos de la comisión se guradaron correctamente", vbInformation, "Aviso"
'        End If
'
'
'
'
'    Else
'        MsgBox "Verificar que los datos ingresados para la comisión, sean los correctos y no coincidan con otra comision existente", vbInformation, "Aviso"
'
'    End If
'
'    Exit Sub
'
'error:
'    MsgBox Err.Description, vbCritical, "Aviso"
'
'End Sub
'
'Private Sub ListarComisiones()
'    GetComisionTransferencia
'End Sub
'
'Private Sub GetComisionTransferencia()
'
'    Dim rs As ADODB.Recordset
'    Dim oDeficion As New COMNCaptaGenerales.NCOMCaptaDefinicion
'    Dim oConstante As New NConstSistemas
'
'    'Dim oCont As New COMNContabilidad.NCOMContFunciones
'
'    'sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
'
'
'
'    Dim i As Integer
'    Dim nFila As Integer
'    Dim nComisionMin As Double
'    Set rs = oDeficion.GetComisionTransferencia
'    'nComisionMin = oConstante.ActualizaConstSistemas(464, oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), txtMontoMinimo.value)
'    nComisionMin = CDbl(oConstante.LeeConstSistema(464))
'    txtMontoMinimo.value = Format(nComisionMin, "#0.00")
'
'    grdComision.Rows = 3
'
'    If Not rs Is Nothing Then
'        If Not rs.BOF And Not rs.EOF Then
'            nFila = rs.RecordCount
'            For i = 0 To nFila - 1
'                grdComision.TextMatrix(i + 2, 0) = i + 1
'                grdComision.TextMatrix(i + 2, 1) = rs!BANCO
'                grdComision.TextMatrix(i + 2, 2) = rs!Tipo
'                grdComision.TextMatrix(i + 2, 3) = rs!PLAZA
'                grdComision.TextMatrix(i + 2, 4) = rs!MONEDAOPE
'                grdComision.TextMatrix(i + 2, 5) = Format(rs!nDesde, "#0.00")
'                grdComision.TextMatrix(i + 2, 6) = Format(rs!nHasta, "#0.00")
'                grdComision.TextMatrix(i + 2, 7) = rs!TIPOCALC
'                grdComision.TextMatrix(i + 2, 8) = Format(rs!MONEDA_COMI, "#0.00")
'                grdComision.TextMatrix(i + 2, 9) = Format(rs!MONTO_COMI, "#0.00")
'                grdComision.TextMatrix(i + 2, 10) = rs!cidBanco
'                grdComision.TextMatrix(i + 2, 11) = rs!nIdComision
'                If i < nFila - 1 Then
'                    AdicionaRow grdComision
'                    rs.MoveNext
'                End If
'            Next
'        End If
'    End If
'
'End Sub
'
'Private Function Validacion() As Boolean
'
'    Dim cCodBanco, sPlaza, sMoneda As String
'
'    Dim i, nValm, nTipo As Integer
'    Dim nRangoMin, nRangoMax As Double
'    Dim nMin, nMax As Double
'    Dim bResultado As Boolean
'
'    bResultado = True
'
'    cCodBanco = Mid(txtBuscarBanco.Text, 4, 13)
'    sPlaza = Trim(Right(cbPlaza.Text, 5))
'    sMoneda = Trim(Right(cbMoneda.Text, 5))
'    nRangoMin = CDbl(txtDesde.value)
'    nRangoMax = CDbl(txtHasta.value)
'    nTipo = CInt(Trim(Right(cbTipo.Text, 5)))
'
'    ' Verificando si Grid tiene registros
'    If grdComision.Rows > 2 And Len(Trim(grdComision.TextMatrix(2, 1))) = 0 Then
'        If nRangoMin >= nRangoMax Then
'            bResultado = False
'        Else
'            bResultado = True
'        End If
'        Validacion = bResultado
'        Exit Function
'    End If
'
'    ' Validando Registros
'    For i = 2 To grdComision.Rows - 1
'        If Not nIdComision = Trim(grdComision.TextMatrix(i, 11)) Then
'            nValm = 0
'            ' Validando Banco
'            If Trim(Right(grdComision.TextMatrix(i, 1), 15)) = cCodBanco Then
'                nValm = nValm + 1
'            End If
'            ' Validando Plaza
'            If Trim(Right(grdComision.TextMatrix(i, 3), 15)) = sPlaza Then
'                nValm = nValm + 1
'            End If
'            ' Validando Moneda
'            If Trim(Right(grdComision.TextMatrix(i, 4), 5)) = sMoneda Then
'                nValm = nValm + 1
'            End If
'            'Validando Tipo: Emision / Recepcion
'            If Trim(Right(grdComision.TextMatrix(i, 2), 5)) = nTipo Then
'                nValm = nValm + 1
'            End If
'            ' Validando Rango
'            nMin = CDbl(IIf(Len(Trim(grdComision.TextMatrix(i, 5))) > 0, Trim(grdComision.TextMatrix(i, 5)), 0))
'            nMax = CDbl(IIf(Len(Trim(grdComision.TextMatrix(i, 6))) > 0, Trim(grdComision.TextMatrix(i, 6)), 0))
'            If nRangoMin < nMin Then
'                If nRangoMax >= nMin Then
'                    nValm = nValm + 1
'                End If
'            ElseIf nRangoMin <= nMax Then
'                nValm = nValm + 1
'            End If
'            If nValm = 5 Then
'                bResultado = False
'            End If
'        End If
'    Next
'    If nRangoMin >= nRangoMax Then
'        bResultado = False
'    End If
'    If txtComision.value <= 0 Then
'        bResultado = False
'    End If
'    If Len(Trim(cCodBanco)) = 0 Then
'        bResultado = False
'    End If
'    Validacion = bResultado
'
'End Function
'
'Private Sub cmdCancelar_Click()
'    Limpiar
'End Sub
'
'Private Sub cmdDefinir_Click()
'
'    On Error GoTo error
'
'    Dim oDefinicion As New COMNCaptaGenerales.NCOMCaptaDefinicion
'    Dim nComisionMinima As Double
'
'    If MsgBox("¿Se actualizará el monto mínimo de comisión por transferencia" & vbNewLine & "Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'
'        nComisionMinima = txtMontoMinimo.value
'        oDefinicion.ActualizaMinComisionTransferencia (nComisionMinima)
'
'        MsgBox "Valor actualizado.", vbInformation, "Aviso"
'        'Limpiar
'        If txtBuscarBanco.Enabled Then txtBuscarBanco.SetFocus
'    End If
'
'    Exit Sub
'
'error:
'    MsgBox Err.Description, vbCritical, "Aviso"
'
'End Sub
'
'Private Sub CmdEditar_Click()
'
'    Dim nRow As Integer, nIndice As Integer
'    nRow = grdComision.row
'
'    If grdComision.Rows = 3 And Len(Trim(grdComision.TextMatrix(grdComision.row, 1))) = 0 Then
'       MsgBox "No hay comisiones disponibles", vbInformation, "Aviso"
'       Exit Sub
'
'    ElseIf Len(Trim(grdComision.TextMatrix(grdComision.row, 1))) = 0 Then
'        MsgBox "Debe seleccionar una comision válida", vbInformation, "Aviso"
'        Exit Sub
'
'    End If
'
'    For nIndice = 1 To grdComision.Cols - 1
'        grdComision.col = nIndice
'        grdComision.CellBackColor = &H8000000D
'        grdComision.CellForeColor = &H80000009
'    Next
'
'    'Llenando datos en controles
'    grdComision.Enabled = False
'    txtBuscarBanco.Text = "01." & grdComision.TextMatrix(nRow, 10)
'    txtBuscarBanco_EmiteDatos
'
'    'Combo Tipo
'    For nIndice = 0 To cbTipo.ListCount - 1
'        cbTipo.ListIndex = nIndice
'        If Trim(Right(cbTipo.Text, 5)) = Trim(Right(grdComision.TextMatrix(nRow, 2), 5)) Then
'            nIndice = cbTipo.ListCount
'        End If
'    Next
'
'    'Combo Plaza
'    For nIndice = 0 To cbPlaza.ListCount - 1
'        cbPlaza.ListIndex = nIndice
'        If Trim(Right(cbPlaza.Text, 5)) = Trim(Right(grdComision.TextMatrix(nRow, 3), 5)) Then
'            nIndice = cbPlaza.ListCount
'        End If
'    Next
'
'    'Combo Moneda
'    For nIndice = 0 To cbMoneda.ListCount - 1
'        cbMoneda.ListIndex = nIndice
'        If Trim(Right(cbMoneda.Text, 5)) = Trim(Right(grdComision.TextMatrix(nRow, 4), 5)) Then
'            nIndice = cbMoneda.ListCount
'        End If
'    Next
'
'    'Montos
'    txtDesde.value = grdComision.TextMatrix(nRow, 5) ' desde
'    txtHasta.value = grdComision.TextMatrix(nRow, 6) ' hasta
'
'    'Combo Tipo Calculo
'    For nIndice = 0 To cbTipoCalculo.ListCount - 1
'        cbTipoCalculo.ListIndex = nIndice
'        If Trim(Right(cbTipoCalculo.Text, 5)) = Trim(Right(grdComision.TextMatrix(nRow, 7), 5)) Then
'            nIndice = cbTipoCalculo.ListCount
'        End If
'    Next
'
'    'Combo Moneda Comision
'    For nIndice = 0 To cbMonedaComision.ListCount - 1
'        cbMonedaComision.ListIndex = nIndice
'        If Trim(Right(cbMonedaComision.Text, 5)) = Trim(Right(grdComision.TextMatrix(nRow, 8), 5)) Then
'            nIndice = cbMonedaComision.ListCount
'        End If
'    Next
'
'    'Comision
'    txtComision.value = grdComision.TextMatrix(nRow, 9)
'    nIdComision = IIf(IsNumeric(grdComision.TextMatrix(nRow, 11)), CInt(grdComision.TextMatrix(nRow, 11)), 0)
'    Operacion = 2
'    cmdQuitar.Enabled = False
'    cmdDefinir.Enabled = False
'    If txtComision.Enabled Then txtComision.SetFocus
'
'End Sub
'
'Private Sub cmdQuitar_Click()
'
'    Dim nRow, nIndice As Integer
'    nRow = grdComision.row
'
'    If grdComision.Rows = 3 And Len(Trim(grdComision.TextMatrix(grdComision.row, 1))) = 0 Then
'       MsgBox "No hay comisiones disponibles", vbInformation, "Aviso"
'       Exit Sub
'
'    ElseIf Len(Trim(grdComision.TextMatrix(grdComision.row, 1))) = 0 Then
'        MsgBox "Debe seleccionar una comision válida", vbInformation, "Aviso"
'        Exit Sub
'
'    End If
'
'    For nIndice = 1 To grdComision.Cols - 1
'        grdComision.col = nIndice
'        grdComision.CellBackColor = &H8000000D
'        grdComision.CellForeColor = &H80000009
'    Next
'
'    If MsgBox("¿Esta seguro de quitar la comision seleccionada?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'        Dim oDeficion As New COMNCaptaGenerales.NCOMCaptaDefinicion
'        nRow = grdComision.row
'        Call oDeficion.ActualizaComisionTransferencia(Trim(grdComision.TextMatrix(nRow, 11)), _
'                                                      Trim(Right(grdComision.TextMatrix(nRow, 1), 15)), _
'                                                      Trim(Right(grdComision.TextMatrix(nRow, 2), 5)), _
'                                                      Trim(Right(grdComision.TextMatrix(nRow, 3), 5)), _
'                                                      Trim(Right(grdComision.TextMatrix(nRow, 4), 5)), _
'                                                      Trim(grdComision.TextMatrix(nRow, 5)), _
'                                                      Trim(grdComision.TextMatrix(nRow, 6)), _
'                                                      Trim(Right(grdComision.TextMatrix(nRow, 7), 5)), _
'                                                      Trim(Right(grdComision.TextMatrix(nRow, 8), 5)), _
'                                                      Trim(grdComision.TextMatrix(nRow, 9)), _
'                                                      txtMontoMinimo.value, _
'                                                      0)
'        Limpiar
'        ListarComisiones
'        MsgBox "Comisión eliminada.", vbInformation, "Aviso"
'        txtBuscarBanco.SetFocus
'    Else
'        For nIndice = 1 To grdComision.Cols - 1
'            grdComision.col = nIndice
'            grdComision.CellBackColor = &H80000009
'            grdComision.CellForeColor = &H80000007
'        Next
'
'    End If
'
'End Sub
'
'Private Sub cmdsalir_Click()
'    If MsgBox("¿Desea salir del formulario de administracion de comisiones?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'        Unload Me
'    End If
'End Sub
'
'Private Sub Form_Load()
'
'Me.Caption = "Administración Comisiones por Transferencia Bancaria"
'
'grdComision.TextMatrix(0, 0) = "#"
'grdComision.TextMatrix(1, 0) = "#"
'grdComision.TextMatrix(0, 1) = "Banco"
'grdComision.TextMatrix(1, 1) = "Banco"
'grdComision.TextMatrix(0, 2) = "Características Op."
'grdComision.TextMatrix(0, 3) = "Características Op."
'grdComision.TextMatrix(1, 2) = "Tipo"
'grdComision.TextMatrix(1, 3) = "Plaza"
'grdComision.TextMatrix(0, 4) = "Montos de la Operacion"
'grdComision.TextMatrix(0, 5) = "Montos de la Operacion"
'grdComision.TextMatrix(0, 6) = "Montos de la Operacion"
'grdComision.TextMatrix(1, 4) = "Moneda"
'grdComision.TextMatrix(1, 5) = "Desde"
'grdComision.TextMatrix(1, 6) = "Hasta"
'grdComision.TextMatrix(0, 7) = "Condiciones de la Comisión"
'grdComision.TextMatrix(0, 8) = "Condiciones de la Comisión"
'grdComision.TextMatrix(0, 9) = "Condiciones de la Comisión"
'grdComision.TextMatrix(1, 7) = "Tipo Calc."
'grdComision.TextMatrix(1, 8) = "Moneda"
'grdComision.TextMatrix(1, 9) = "Monto"
'
'grdComision.MergeCells = 2
'grdComision.MergeCol(0) = True
'grdComision.MergeCol(1) = True
'grdComision.MergeCol(2) = True
'grdComision.MergeCol(3) = True
'grdComision.MergeCol(4) = True
'grdComision.MergeCol(5) = True
'grdComision.MergeCol(6) = True
'grdComision.MergeCol(7) = True
'grdComision.MergeCol(8) = True
'grdComision.MergeCol(9) = True
'
'grdComision.MergeRow(0) = True
'grdComision.MergeRow(1) = True
'
'grdComision.SelectionMode = flexSelectionFree
'
''Obteniendo lista de entidades financieras
'Dim clsBanco As COMNCajaGeneral.NCOMCajaCtaIF
'Set clsBanco = New COMNCajaGeneral.NCOMCajaCtaIF
'txtBuscarBanco.psRaiz = "BANCOS"
'txtBuscarBanco.rs = clsBanco.CargaCtasIF(gMonedaNacional, "0[1]%", MuestraInstituciones)
'grdComision.ColWidth(0) = 300
'grdComision.ColWidth(1) = 3000
'grdComision.ColWidth(2) = 1200
'grdComision.ColWidth(3) = 1200
'grdComision.ColWidth(10) = 0
'grdComision.ColWidth(11) = 0
'grdComision.ColAlignmentFixed(-1) = flexAlignCenterCenter
'grdComision.ColAlignment(0) = flexAlignCenterCenter
'
'Dim oConstante As COMDConstSistema.DCOMGeneral
'Dim rsConstante As ADODB.Recordset
'Set oConstante = New COMDConstSistema.DCOMGeneral
'
'' cargando cbTipo
'Set rsConstante = oConstante.GetConstante("10032", , "'10[^0]'")
'CargaCombo cbTipo, rsConstante
'cbTipo.ListIndex = 0
'Set rsConstante = Nothing
'
''cargando plaza
'Set rsConstante = oConstante.GetConstante("10032", , "'20[^0]'")
'CargaCombo cbPlaza, rsConstante
'cbPlaza.ListIndex = 0
'Set rsConstante = Nothing
'
''cargando moneda operacion
'Set rsConstante = oConstante.GetConstante("10032", , "'30[^0]'")
'CargaCombo cbMoneda, rsConstante
'cbMoneda.ListIndex = 0
'Set rsConstante = Nothing
'
''cargando tipo cálculo
'Set rsConstante = oConstante.GetConstante("10032", , "'40[^0]'")
'CargaCombo cbTipoCalculo, rsConstante
'cbTipoCalculo.ListIndex = 0
'Set rsConstante = Nothing
'
''cargando moneda comision
'Set rsConstante = oConstante.GetConstante("10032", , "'50[^0]'")
'CargaCombo cbMonedaComision, rsConstante
'cbMonedaComision.ListIndex = 0
'Set rsConstante = Nothing
'Operacion = 1
'ListarComisiones
'
'If nPermiso = 2 Then
'    txtBuscarBanco.Enabled = False
'    txtBanco.Enabled = False
'    cbTipo.Enabled = False
'    cbPlaza.Enabled = False
'    cbMoneda.Enabled = False
'    txtDesde.Enabled = False
'    txtHasta.Enabled = False
'    cbTipoCalculo.Enabled = False
'    cbMonedaComision.Enabled = False
'    txtComision.Enabled = False
'    cmdAceptar.Enabled = False
'    cmdCancelar.Enabled = False
'    cmdEditar.Enabled = False
'    cmdQuitar.Enabled = False
'    txtMontoMinimo.Enabled = False
'    cmdDefinir.Enabled = False
'    grdComision.Enabled = True
'End If
'
'End Sub
'
'Private Sub txtBanco_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'    If cbTipo.Enabled Then cbTipo.SetFocus
'    End If
'End Sub
'
'Private Sub txtBuscarBanco_EmiteDatos()
'    Dim oNCajaCtaIF As New clases.NCajaCtaIF
'    Dim oDOperacion As New clases.DOperacion
'
'    txtBanco.Text = ""
'    If txtBuscarBanco.Text <> "" Then
'        txtBanco.Text = oNCajaCtaIF.NombreIF(Mid(txtBuscarBanco.Text, 4, 13))
'    End If
'    Set oNCajaCtaIF = Nothing
'    Set oDOperacion = Nothing
'    cbTipo.SetFocus
'End Sub
'
'Private Sub txtComision_Change()
'
'Dim nTipoCalculo As Integer
'Dim nComision As Double
'
'nTipoCalculo = CInt(val(Trim(Right(cbTipoCalculo.Text, 5))))
'nComision = val(txtComision.value)
'
'If nTipoCalculo = 402 Then
'    If nComision > 100 Then
'        txtComision.value = 100
'        MsgBox "El porcentaje máximo es 100%", vbInformation, "Aviso"
'    End If
'End If
'
'End Sub
'
'Private Sub txtComision_GotFocus()
'txtComision.SelStart = 0
'txtComision.SelLength = 20
'End Sub
'
'Private Sub txtComision_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If cmdAceptar.Enabled Then cmdAceptar.SetFocus
'    End If
'End Sub
'
'Private Sub txtDesde_GotFocus()
'    txtDesde.SelStart = 0
'    txtDesde.SelLength = 20
'End Sub
'
'Private Sub txtDesde_KeyPress(KeyAscii As Integer)
'
'    If KeyAscii = 13 Then
'        If txtHasta.Enabled Then txtHasta.SetFocus
'    End If
'
'End Sub
'
'Private Sub txtHasta_GotFocus()
'txtHasta.SelStart = 0
'txtHasta.SelLength = 20
'End Sub
'
'Private Sub txthasta_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If cbTipoCalculo.Enabled Then cbTipoCalculo.SetFocus
'    End If
'End Sub
'
'Private Sub txtMontoMinimo_GotFocus()
'txtMontoMinimo.SelStart = 0
'txtMontoMinimo.SelLength = 20
'End Sub
'
'Private Sub Limpiar()
'
'Dim nIndice As Integer
'
'For nIndice = 1 To grdComision.Cols - 1
'    grdComision.col = nIndice
'    grdComision.CellBackColor = &HFFFFFF
'    grdComision.CellForeColor = &H80000007
'Next
'
'nIdComision = 0
'txtBanco.Text = ""
'txtBuscarBanco.Text = ""
'cbTipo.ListIndex = 0
'cbMoneda.ListIndex = 0
'cbPlaza.ListIndex = 0
'txtDesde.value = 0
'txtHasta.value = 0
'cbTipoCalculo.ListIndex = 0
'cbMonedaComision.ListIndex = 0
'txtComision.value = 0
'grdComision.Enabled = True
'grdComision.Rows = 3
'grdComision.TextMatrix(2, 1) = ""
'grdComision.TextMatrix(2, 2) = ""
'grdComision.TextMatrix(2, 3) = ""
'grdComision.TextMatrix(2, 4) = ""
'grdComision.TextMatrix(2, 5) = ""
'grdComision.TextMatrix(2, 6) = ""
'grdComision.TextMatrix(2, 7) = ""
'grdComision.TextMatrix(2, 8) = ""
'grdComision.TextMatrix(2, 9) = ""
'grdComision.TextMatrix(2, 10) = ""
'grdComision.TextMatrix(2, 11) = ""
'ListarComisiones
'If txtBuscarBanco.Enabled Then txtBuscarBanco.SetFocus
'Operacion = 1
'cmdAceptar.Enabled = True
'cmdCancelar.Enabled = True
'cmdEditar.Enabled = True
'cmdQuitar.Enabled = True
'cmdDefinir.Enabled = True
'
'End Sub
'
'Private Sub txtMontoMinimo_KeyPress(KeyAscii As Integer)
'    If KeyAscii = 13 Then
'        If cmdDefinir.Enabled Then cmdDefinir.SetFocus
'    End If
'End Sub
