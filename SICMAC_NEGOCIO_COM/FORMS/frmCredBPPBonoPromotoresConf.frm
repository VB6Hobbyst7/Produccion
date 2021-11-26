VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredBPPBonoPromotoresConf 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Parámetros Bono Promotores"
   ClientHeight    =   4830
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6390
   Icon            =   "frmCredBPPBonoPromotoresConf.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4830
   ScaleWidth      =   6390
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   5160
      TabIndex        =   6
      Top             =   4320
      Width           =   1095
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   4320
      Width           =   1095
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   2415
      Left            =   120
      TabIndex        =   1
      Top             =   1800
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   4260
      _Version        =   393216
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Minorista y No Minorista"
      TabPicture(0)   =   "frmCredBPPBonoPromotoresConf.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feMinoYNoMino"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "cmdAgregar"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "cmdQuitar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdModificar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).ControlCount=   4
      TabCaption(1)   =   "Convenio"
      TabPicture(1)   =   "frmCredBPPBonoPromotoresConf.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "feConvenio"
      Tab(1).Control(1)=   "cmdAgregarCon"
      Tab(1).Control(2)=   "cmdQuitarCon"
      Tab(1).Control(3)=   "cmdModificarCon"
      Tab(1).ControlCount=   4
      Begin VB.CommandButton cmdModificarCon 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   -70080
         TabIndex        =   15
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitarCon 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   -70080
         TabIndex        =   14
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregarCon 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   -70080
         TabIndex        =   13
         Top             =   480
         Width           =   1095
      End
      Begin VB.CommandButton cmdModificar 
         Caption         =   "Modificar"
         Height          =   375
         Left            =   4920
         TabIndex        =   12
         Top             =   1440
         Width           =   1095
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "Quitar"
         Height          =   375
         Left            =   4920
         TabIndex        =   11
         Top             =   960
         Width           =   1095
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         Height          =   375
         Left            =   4920
         TabIndex        =   10
         Top             =   480
         Width           =   1095
      End
      Begin SICMACT.FlexEdit feMinoYNoMino 
         Height          =   1815
         Left            =   120
         TabIndex        =   16
         Top             =   480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3201
         Cols0           =   5
         HighLight       =   1
         EncabezadosNombres=   "#-Desde-Hasta-Porcentaje-Aux"
         EncabezadosAnchos=   "0-1500-1500-1200-0"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-R-R-C"
         FormatosEdit    =   "0-2-2-2-0"
         CantEntero      =   12
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
      Begin SICMACT.FlexEdit feConvenio 
         Height          =   1815
         Left            =   -74880
         TabIndex        =   17
         Top             =   480
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   3201
         Cols0           =   5
         HighLight       =   1
         EncabezadosNombres=   "#-Desde-Hasta-Porcentaje-Aux"
         EncabezadosAnchos=   "0-1500-1500-1200-0"
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
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-R-R-C"
         FormatosEdit    =   "0-2-2-2-0"
         CantEntero      =   15
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
      End
   End
   Begin VB.Frame fraConfig 
      Caption         =   "Configuración General"
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
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6135
      Begin SICMACT.EditMoney txtMinCartera 
         Height          =   300
         Left            =   2880
         TabIndex        =   7
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtMontoMaxBon 
         Height          =   300
         Left            =   2880
         TabIndex        =   8
         Top             =   720
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtTEM 
         Height          =   300
         Left            =   2880
         TabIndex        =   9
         Top             =   1080
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto máximo de bonificación S/.:"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   720
         Width           =   2460
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "T.E.M. Mínima (no convenio)(%):"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   1080
         Width           =   2325
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Monto mínimo de colocaciones S/.:"
         Height          =   195
         Left            =   240
         TabIndex        =   2
         Top             =   360
         Width           =   2535
      End
   End
End
Attribute VB_Name = "frmCredBPPBonoPromotoresConf"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''*****************************************************************************************************
''** Nombre      : frmCredBPPBonoPromotoresConf
''** Descripción : Formulario para la configuración del bono promotores
''** Creación    : WIOR, 20140620 10:00:00 AM
''*****************************************************************************************************
'Option Explicit
'Private fnNoMoverFilaMinoYNoMino As Integer
'Private fnNoMoverFilaConvenio As Integer
'Private fbPorGrabarMinoYNoMino As Boolean
'Private fbPorGrabarConvenio As Boolean
'
'Private Sub cmdAgregar_Click()
'If ValidaDatosFlex(feMinoYNoMino, IIf(fnNoMoverFilaMinoYNoMino = 0, feMinoYNoMino.Rows - 2, feMinoYNoMino.Rows - 1), "Minorista y No Minorista") Then
'    feMinoYNoMino.AdicionaFila
'    fnNoMoverFilaMinoYNoMino = feMinoYNoMino.Rows - 1
'    feMinoYNoMino.lbEditarFlex = True
'    fbPorGrabarMinoYNoMino = True
'    feMinoYNoMino.ColumnasAEditar = "X-1-2-3-X"
'    feMinoYNoMino.SetFocus
'    SendKeys "{Enter}"
'End If
'End Sub
'
'Private Sub cmdAgregarCon_Click()
'If ValidaDatosFlex(feConvenio, IIf(fnNoMoverFilaConvenio = 0, feConvenio.Rows - 2, feConvenio.Rows - 1), "Convenio") Then
'    feConvenio.AdicionaFila
'    fnNoMoverFilaConvenio = feConvenio.Rows - 1
'    feConvenio.lbEditarFlex = True
'    fbPorGrabarConvenio = True
'    feConvenio.ColumnasAEditar = "X-1-2-3-X"
'    feConvenio.SetFocus
'    SendKeys "{Enter}"
'End If
'End Sub
'
'Private Sub cmdCerrar_Click()
'Unload Me
'End Sub
'
'Private Sub cmdGuardar_Click()
'If ValidaDatos Then
'    If MsgBox("Estas seguro de guardar los Datos?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'        feMinoYNoMino.ColumnasAEditar = "X-1-2-3-X"
'        feConvenio.ColumnasAEditar = "X-1-2-3-X"
'        fbPorGrabarConvenio = False
'        fbPorGrabarMinoYNoMino = False
'        feMinoYNoMino.lbEditarFlex = False
'        feConvenio.lbEditarFlex = False
'        fnNoMoverFilaMinoYNoMino = feMinoYNoMino.Rows - 1
'        fnNoMoverFilaConvenio = feConvenio.Rows - 1
'        Dim oBPP As COMNCredito.NCOMBPPR
'        Dim i As Integer
'        Set oBPP = New COMNCredito.NCOMBPPR
'
'        Call oBPP.OpeConfigPromotoresGen(2)
'        Call oBPP.OpeConfigPromotoresGen(1, CDbl(txtMinCartera.Text), CDbl(txtMontoMaxBon.Text), CDbl(txtTEM.Text))
'
'        Call oBPP.OpeConfigPromotoresGenDet(2, 1)
'        For i = 1 To feMinoYNoMino.Rows - 1
'            Call oBPP.OpeConfigPromotoresGenDet(1, 1, CDbl(feMinoYNoMino.TextMatrix(i, 1)), CDbl(feMinoYNoMino.TextMatrix(i, 2)), CDbl(feMinoYNoMino.TextMatrix(i, 3)))
'        Next i
'
'        Call oBPP.OpeConfigPromotoresGenDet(2, 2)
'        For i = 1 To feConvenio.Rows - 1
'            Call oBPP.OpeConfigPromotoresGenDet(1, 2, CDbl(feConvenio.TextMatrix(i, 1)), CDbl(feConvenio.TextMatrix(i, 2)), CDbl(feConvenio.TextMatrix(i, 3)))
'        Next i
'
'        MsgBox "Se guardo correctamente los datos", vbInformation, "Aviso"
'    End If
'End If
'End Sub
'
'Private Sub cmdModificar_Click()
'If Not fbPorGrabarMinoYNoMino Then
'    If feMinoYNoMino.row = feMinoYNoMino.Rows - 1 Then
'        feMinoYNoMino.ColumnasAEditar = "X-1-X-3-X"
'    Else
'        feMinoYNoMino.ColumnasAEditar = "X-X-X-3-X"
'    End If
'    fnNoMoverFilaMinoYNoMino = feMinoYNoMino.row
'    feMinoYNoMino.lbEditarFlex = True
'    fbPorGrabarMinoYNoMino = True
'    feMinoYNoMino.SetFocus
'    SendKeys "{Enter}"
'Else
'    MsgBox "Favor de guardar primero antes de intentar modificar algun registro", vbInformation, "Aviso"
'End If
'End Sub
'
'Private Sub cmdModificarCon_Click()
'If Not fbPorGrabarConvenio Then
'    If feConvenio.row = feConvenio.Rows - 1 Then
'        feConvenio.ColumnasAEditar = "X-1-X-3-X"
'    Else
'        feConvenio.ColumnasAEditar = "X-X-X-3-X"
'    End If
'    fnNoMoverFilaConvenio = feConvenio.row
'    feConvenio.lbEditarFlex = True
'    fbPorGrabarConvenio = True
'    feConvenio.SetFocus
'    SendKeys "{Enter}"
'Else
'    MsgBox "Favor de guardar primero antes de intentar modificar algun registro", vbInformation, "Aviso"
'End If
'End Sub
'
'Private Sub cmdQuitar_Click()
'If MsgBox("Estas seguro de quitar el último registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'    feMinoYNoMino.EliminaFila feMinoYNoMino.Rows - 1
'
'    If Trim(feMinoYNoMino.TextMatrix(1, 0)) = "" And (feMinoYNoMino.Rows - 1) = 1 Then
'        fnNoMoverFilaMinoYNoMino = 0
'    Else
'        fnNoMoverFilaMinoYNoMino = feMinoYNoMino.Rows - 1
'    End If
'End If
'End Sub
'
'Private Sub cmdQuitarCon_Click()
'If MsgBox("Estas seguro de quitar el último registro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
'    feConvenio.EliminaFila feConvenio.Rows - 1
'
'    If Trim(feConvenio.TextMatrix(1, 0)) = "" And (feConvenio.Rows - 1) = 1 Then
'        fnNoMoverFilaConvenio = 0
'    Else
'        fnNoMoverFilaConvenio = feConvenio.Rows - 1
'    End If
'End If
'End Sub
'
'Private Sub feConvenio_RowColChange()
'If fbPorGrabarConvenio Then
'    feConvenio.row = IIf(fnNoMoverFilaConvenio = 0, 1, fnNoMoverFilaConvenio)
'End If
'End Sub
'
'Private Sub feMinoYNoMino_RowColChange()
'If fbPorGrabarMinoYNoMino Then
'    feMinoYNoMino.row = IIf(fnNoMoverFilaMinoYNoMino = 0, 1, fnNoMoverFilaMinoYNoMino)
'End If
'End Sub
'
'Private Function ValidaDatosFlex(ByVal pFE As FlexEdit, ByVal pFila As Integer, ByVal psPestana As String) As Boolean
'On Error GoTo ValidaError
'ValidaDatosFlex = True
'Dim i As Integer
'
'If pFila > 0 Then
'
'For i = 1 To 3
'    If Trim(pFE.TextMatrix(pFila, i)) = "" Then
'        ValidaDatosFlex = False
'        MsgBox "Ingrese el valor de la fila " & pFila & " en la columna " & Choose(i, "Desde", "Hasta", "Porcentaje") & " en la pestaña " & psPestana & ".", vbInformation, "Aviso"
'        Exit Function
'    End If
'
'
'    If CDbl(pFE.TextMatrix(pFila, i)) < 0 Then
'        ValidaDatosFlex = False
'        MsgBox "Ingrese un valor mayor o igual a 0 de la fila " & pFila & " en la columna " & Choose(i, "Desde", "Hasta", "Porcentaje") & " en la pestaña " & psPestana & ".", vbInformation, "Aviso"
'        Exit Function
'    End If
'
'    If i = 1 And pFila = 1 Then
'        If CDbl(pFE.TextMatrix(pFila, i)) <> 0 Then
'            ValidaDatosFlex = False
'            MsgBox "El valor de la Columna Desde en la fila  " & pFila & " debe ser 0 en la pestaña " & psPestana & ".", vbInformation, "Aviso"
'            Exit Function
'        End If
'    End If
'
'    If i = 2 Then
'        If CDbl(pFE.TextMatrix(pFila, i - 1)) > CDbl(pFE.TextMatrix(pFila, i)) Then
'            ValidaDatosFlex = False
'            MsgBox "El valor de la Columna Hasta debe ser mayor a la Columna Desde en la fila  " & pFila & " en la pestaña " & psPestana & ".", vbInformation, "Aviso"
'            Exit Function
'        End If
'    End If
'
'    If i = 3 Then
'        If CDbl(pFE.TextMatrix(pFila, i)) > 100 Then
'            ValidaDatosFlex = False
'            MsgBox "El valor del Porcentaje no debe ser mayor a 100 de la fila " & pFila & " en la pestaña " & psPestana & ".", vbInformation, "Aviso"
'            Exit Function
'        End If
'    End If
'Next i
'
'    If pFila > 1 Then
'        If CCur(CDbl(pFE.TextMatrix(pFila, 1)) - CDbl(pFE.TextMatrix(pFila - 1, 2))) <> 0.01 Then
'            ValidaDatosFlex = False
'            MsgBox "Ingrese correctamente el valor de la Columna Desde en la fila " & pFila & " con relacion a la columna Hasta de la fila " & (pFila - 1) & " en la pestaña " & psPestana & ".", vbInformation, "Aviso"
'            Exit Function
'        End If
'
'        If CDbl(pFE.TextMatrix(pFila, 3)) > CDbl(pFE.TextMatrix(pFila - 1, 3)) Then
'            ValidaDatosFlex = False
'            MsgBox "El valor de la Columna Porcentaje de la fila " & pFila & " no puede ser mayor que de la fila " & (pFila - 1) & " en la pestaña " & psPestana & ".", vbInformation, "Aviso"
'            Exit Function
'        End If
'
'    End If
'End If
'
'Exit Function
'ValidaError:
'ValidaDatosFlex = False
'MsgBox err.Description, vbCritical, "Error"
'End Function
'
'Private Function ValidaDatos() As Boolean
'ValidaDatos = True
'Dim i As Integer
'
'If Trim(txtMinCartera.Text) = "" Then
'    ValidaDatos = False
'    MsgBox "Ingrese El Monto Mínimo de Colocaciones.", vbInformation, "Aviso"
'    txtMinCartera.SetFocus
'    Exit Function
'End If
'
'If CDbl(txtMinCartera.Text) <= 0 Then
'    ValidaDatos = False
'    MsgBox "El Monto Mínimo de Colocaciones tiene que ser mayor a 0.", vbInformation, "Aviso"
'    txtMinCartera.SetFocus
'    Exit Function
'End If
'
'If Trim(txtMontoMaxBon.Text) = "" Then
'    ValidaDatos = False
'    MsgBox "Ingrese El Monto Máximo de Bonifiación.", vbInformation, "Aviso"
'    txtMontoMaxBon.SetFocus
'    Exit Function
'End If
'
'If CDbl(txtMontoMaxBon.Text) < 0 Then
'    ValidaDatos = False
'    MsgBox "El Monto Máximo de Bonifiación tiene que ser mayor a 0.", vbInformation, "Aviso"
'    txtMontoMaxBon.SetFocus
'    Exit Function
'End If
'
'
'If Trim(txtTEM.Text) = "" Then
'    ValidaDatos = False
'    MsgBox "Ingrese la T.E.M. Mínima", vbInformation, "Aviso"
'    txtTEM.SetFocus
'    Exit Function
'End If
'
'If CDbl(txtTEM.Text) < 0 Or CDbl(txtTEM.Text) > 100 Then
'    ValidaDatos = False
'    MsgBox "El valor de la T.E.M. debe estar entre 0 y 100.", vbInformation, "Aviso"
'    txtTEM.SetFocus
'    Exit Function
'End If
'
'For i = 1 To feMinoYNoMino.Rows - 1
'    ValidaDatos = ValidaDatosFlex(feMinoYNoMino, i, "Minorista y No Minorista")
'    If Not ValidaDatos Then
'        Exit Function
'    End If
'Next i
'
'For i = 1 To feConvenio.Rows - 1
'    ValidaDatos = ValidaDatosFlex(feConvenio, i, "Convenio")
'    If Not ValidaDatos Then
'        Exit Function
'    End If
'Next i
'
'End Function
'
'Private Sub CargaControles()
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'Set oBPP = New COMNCredito.NCOMBPPR
'Set rsBPP = oBPP.ObtenerConfigPromotoresGen()
'
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    txtMinCartera.Text = Format(rsBPP!nMontoMinCol, "###," & String(15, "#") & "#0.00")
'    txtMontoMaxBon.Text = Format(rsBPP!nMontoMaxBon, "###," & String(15, "#") & "#0.00")
'    txtTEM.Text = Format(rsBPP!nTem, "###," & String(15, "#") & "#0.00")
'End If
'
'Call LimpiaFlex(feMinoYNoMino)
'Call CargaGrilla(1, feMinoYNoMino)
'Call LimpiaFlex(feConvenio)
'Call CargaGrilla(2, feConvenio)
'
'End Sub
'Private Sub CargaGrilla(ByVal pnTipo As Integer, ByRef pFE As FlexEdit)
'Dim oBPP As COMNCredito.NCOMBPPR
'Dim rsBPP As ADODB.Recordset
'Dim i As Integer
'
'Set oBPP = New COMNCredito.NCOMBPPR
'Set rsBPP = oBPP.ObtenerConfigPromotoresGenDet(pnTipo)
'
'
'If Not (rsBPP.EOF And rsBPP.BOF) Then
'    For i = 1 To rsBPP.RecordCount
'        pFE.AdicionaFila
'        pFE.TextMatrix(i, 1) = Format(rsBPP!nDesde, "###," & String(15, "#") & "#0.00")
'        pFE.TextMatrix(i, 2) = Format(rsBPP!nHasta, "###," & String(15, "#") & "#0.00")
'        pFE.TextMatrix(i, 3) = Format(rsBPP!nPorcentaje, "###," & String(15, "#") & "#0.00")
'        rsBPP.MoveNext
'    Next i
'End If
'
'End Sub
'
'Private Sub Form_Load()
'CargaControles
'End Sub
