VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmArqueoExpedientesAho 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Arqueos de expedientes de ahorro"
   ClientHeight    =   8025
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10500
   Icon            =   "frmArqueoExpedientesAho.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8025
   ScaleWidth      =   10500
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   9240
      TabIndex        =   16
      Top             =   7560
      Width           =   1095
   End
   Begin VB.Frame fraConformidad 
      Caption         =   "Conformidad de expedientes de ahorro"
      Height          =   3375
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Width           =   10095
      Begin MSComCtl2.DTPicker txtFecFin 
         Height          =   300
         Left            =   2640
         TabIndex        =   21
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   169607169
         CurrentDate     =   42935
      End
      Begin MSComCtl2.DTPicker txtFecIni 
         Height          =   300
         Left            =   720
         TabIndex        =   20
         Top             =   360
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   169607169
         CurrentDate     =   42935
      End
      Begin VB.CheckBox ckbSelTodos 
         Caption         =   "Seleccionar todo"
         Height          =   255
         Left            =   8160
         TabIndex        =   17
         Top             =   360
         Width           =   1695
      End
      Begin SICMACT.FlexEdit flxConformidad 
         Height          =   2415
         Left            =   240
         TabIndex        =   14
         Top             =   960
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   4260
         Cols0           =   14
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmArqueoExpedientesAho.frx":030A
         EncabezadosAnchos=   "0-2200-3000-1200-1200-2400-2000-800-600-1200-1200-2000-2200-2000"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-9-10-11-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0-4-3-3-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "Item"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbOrdenaCol     =   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.CommandButton cmdConsultar 
         Caption         =   "Consultar"
         Height          =   375
         Left            =   4200
         TabIndex        =   13
         Top             =   300
         Width           =   1215
      End
      Begin VB.Label lblIni 
         AutoSize        =   -1  'True
         Caption         =   "Del :"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   420
         Width           =   330
      End
      Begin VB.Label lblFin 
         AutoSize        =   -1  'True
         Caption         =   "Al :"
         Height          =   195
         Left            =   2400
         TabIndex        =   18
         Top             =   420
         Width           =   225
      End
   End
   Begin VB.Frame fraInvolucrados 
      Caption         =   "Personal Involucrado"
      Height          =   2655
      Left            =   240
      TabIndex        =   6
      Top             =   1200
      Width           =   10095
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         Height          =   375
         Left            =   1440
         TabIndex        =   9
         Top             =   2040
         Width           =   1215
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "Nuevo"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   2040
         Width           =   1095
      End
      Begin SICMACT.FlexEdit flxPersonalInvolucrado 
         Height          =   1575
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   2778
         Cols0           =   6
         HighLight       =   1
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Código-Usuario-Nombre-Cargo-Relación con arqueo"
         EncabezadosAnchos=   "800-1600-1200-2600-2800-1600"
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
         ColumnasAEditar =   "X-1-X-X-X-5"
         ListaControles  =   "0-1-0-0-0-3"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "L-C-C-C-C-R"
         FormatosEdit    =   "0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         ColWidth0       =   795
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblAgencia 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   7560
         TabIndex        =   11
         Top             =   2160
         Width           =   2055
      End
      Begin VB.Label Label3 
         Caption         =   "Agencia"
         Height          =   255
         Left            =   6600
         TabIndex        =   10
         Top             =   2160
         Width           =   855
      End
   End
   Begin VB.Frame fraDatosArqueo 
      Caption         =   "Datos del Arqueo"
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   10095
      Begin VB.TextBox txtOtroTipoArqueo 
         Height          =   285
         Left            =   7320
         TabIndex        =   5
         Top             =   360
         Width           =   2415
      End
      Begin VB.ComboBox cboTipoArqueo 
         Height          =   315
         Left            =   5520
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo de Arqueo"
         Height          =   255
         Left            =   4080
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label lblFechaHora 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1680
         TabIndex        =   2
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha/Hora Inicio"
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
   End
   Begin SICMACT.Usuario oUsuario 
      Left            =   240
      Top             =   360
      _ExtentX        =   820
      _ExtentY        =   820
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   7920
      TabIndex        =   15
      Top             =   7560
      Width           =   1095
   End
End
Attribute VB_Name = "frmArqueoExpedientesAho"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'ANDE 20170706 ERS021-2017
Option Explicit
Dim lbConforme As Boolean
Dim lnRowConformidad, lnItemSelComentario, lnItemSelTipoArqueo As Integer
Dim lbComentarioDeficiente As Boolean
Dim laPersonalInvolucrado() As Variant
Dim lcPersonalInvolucrado As String
Dim lrDatos As ADODB.Recordset 'datos cuentas aperturadas

Private Sub cboOrdenar_Change()
End Sub


Private Sub cboTipoArqueo_Change()
       Call TipoDeArqueoSel
End Sub

Private Sub cboTipoArqueo_Click()
    Call TipoDeArqueoSel
End Sub

Public Sub TipoDeArqueoSel()
    Dim cTipoArqueo As String
    cTipoArqueo = Right(cboTipoArqueo.Text, 1)
    If EsNumero(cTipoArqueo) Then
        lnItemSelTipoArqueo = CInt(cTipoArqueo)
        
        If lnItemSelTipoArqueo = 3 Then
            txtOtroTipoArqueo.Enabled = True
            txtOtroTipoArqueo.BackColor = vbWhite
        Else
            txtOtroTipoArqueo.Enabled = False
            txtOtroTipoArqueo.BackColor = vbGrayed
        End If
    Else
        lnItemSelTipoArqueo = -1
    End If
End Sub

Private Sub ckbSelTodos_Click()

    Dim nRows, i As Integer
    nRows = flxConformidad.Rows
    If ckbSelTodos.value = 1 Then
        For i = 1 To nRows - 1
            flxConformidad.TextMatrix(i, 9) = "."
            flxConformidad.SeleccionaChekTecla
        Next i
    Else
        For i = 1 To nRows - 1
            flxConformidad.TextMatrix(i, 9) = ""
        Next i
    End If
    
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub cmdConsultar_Click()
    
    
    On Error GoTo ErrorDeConsulta
    Dim cDel, cAl As String
    cDel = txtFecIni.value
    cAl = txtFecFin.value
    
    If ValidarFechas(cDel, cAl) Then
        ckbSelTodos.value = 0
        'limpiar flxConformidad
        Call LimpiarFlxConformidad
        Dim oCapta As New COMNCaptaGenerales.NCOMCaptaGenerales
            Dim i As Integer
            
            Set lrDatos = oCapta.ObtenerCuentasAperturadas(cDel, cAl, gsCodAge)
            
            If Not (lrDatos.BOF And lrDatos.EOF) Then
                flxConformidad.Enabled = True
                lrDatos.MoveFirst
                i = 1
                Do While Not lrDatos.EOF
                    flxConformidad.AdicionaFila
                    flxConformidad.TextMatrix(i, 1) = lrDatos!Cuenta
                    flxConformidad.TextMatrix(i, 2) = lrDatos!Nombre
                    flxConformidad.TextMatrix(i, 3) = lrDatos!FechaApe
                    flxConformidad.TextMatrix(i, 4) = lrDatos!Hora
                    flxConformidad.TextMatrix(i, 5) = lrDatos!tipo
                    flxConformidad.TextMatrix(i, 6) = Format(lrDatos!MontoApe, "#,##0.00")
                    flxConformidad.TextMatrix(i, 7) = lrDatos!cUsu
                    flxConformidad.TextMatrix(i, 8) = lrDatos!Firma
                    flxConformidad.TextMatrix(i, 12) = lrDatos!OrdenPago
                    flxConformidad.TextMatrix(i, 13) = lrDatos!Producto
                    i = i + 1
                    lrDatos.MoveNext
                Loop
                cmdProcesar.Enabled = True
            Else
                MsgBox "No se encontraron datos.", vbInformation, "Aviso"
            End If
            
            Dim J As Integer
            
            For J = 1 To i - 1
                flxConformidad.TextMatrix(J, 9) = "."
            Next J
    End If
    Exit Sub
ErrorDeConsulta:
    flxConformidad.Enabled = False
    MsgBox "No se pudo hacer la consulta.", vbError + vbOKOnly, "Error"
End Sub

Private Sub cmdEliminar_Click()
    flxPersonalInvolucrado.EliminaFila (flxPersonalInvolucrado.Rows - 1)
End Sub

Private Sub cmdNuevo_Click()
    flxPersonalInvolucrado.AdicionaFila
End Sub

Private Sub cmdProcesar_Click()
    If Validar Then
        Dim oWord As Word.Application
        Dim oDoc As Document
        
        Dim rCiudad, R, rCuentasApeMesArqueo As ADODB.Recordset
        Dim oCapta As New COMNCaptaGenerales.NCOMCaptaGenerales
        Dim i, nTotalRowsFlxConformidad, nCodigoArqueo, nTotalAperturasSistMN, nTotalAperturasSistME, nTotalCuentasConformeMN, nTotalCuentasConformeME, _
        nDiferenciaMN, nDifirenciaME, nRowsPersonalInvolucrado, nFila, nColunma, nTipoArqueo, nPerfil As Integer
        Dim cArchivo, cNombreArchivo, cCiudad, cDel, cAl, cCuenta, cCliente, cFechApertura, cMotivo, cMoneda, cPersCod, cNombreInvolucrado, cCargo, cPerfil, _
        cDescOtro, cFecArqueo, cCodigoArqueo, cComentario As String
        Dim bRegistroFaltantes, bNuevaFila As Boolean
        Dim aCuentasConformes() As Variant
              
        On Error GoTo ErrorProcesarArchivo
                
                
        'obteniendo codigo de arqueo
        cFecArqueo = Trim(lblFechaHora.Caption)
        Set R = oCapta.ObtenerCodigoArqueo(cFecArqueo)
        If Not (R.BOF And R.EOF) Then
            cCodigoArqueo = R!codigo
        End If
        'guardando deficientes y/op faltantes
        nTotalRowsFlxConformidad = flxConformidad.Rows
        
        For i = 1 To nTotalRowsFlxConformidad - 1
            cComentario = Trim(flxConformidad.TextMatrix(i, 10))
            cMotivo = Trim(flxConformidad.TextMatrix(i, 11))
            If cComentario <> "" Then
                    cCuenta = Trim(flxConformidad.TextMatrix(i, 1))
                    Call oCapta.RegistrarDF(cCodigoArqueo, cCuenta, gdFecSis, cComentario, cMotivo)
            End If
        Next i
        
        'iniciando documento word
        Set oWord = CreateObject("Word.Application")
        oWord.Visible = False
        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\PlantillaArqueo.doc")
        
        cDel = txtFecIni.value
        cAl = txtFecFin.value
        cNombreArchivo = "ArqueoExpAho_" & Format(gdFecSis, "YYYYMMDD") & Format(Time, "hhmmss") & ".doc"
        cArchivo = App.Path & "\FormatoCarta\" & cNombreArchivo
        Set rCiudad = oCapta.ObtenerCiudad(gsCodAge)
        If Not (rCiudad.BOF And rCiudad.EOF) Then
            Dim cAgeCod2 As String
            cAgeCod2 = rCiudad!cAgeCod
            If cAgeCod2 = gsCodAge Then
                cCiudad = rCiudad!Ciudad
            Else
                MsgBox "Error al procesar arqueo.", vbError + vbOKOnly, "Error"
                Exit Sub
            End If
        End If
        
        oDoc.SaveAs (cArchivo)
        
        'datos generales
        With oWord.Selection.Find
            .Text = "<<ciudad>>"
            .Replacement.Text = cCiudad
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        With oWord.Selection.Find
            .Text = "<<HH:MM:SS>>"
            .Replacement.Text = CStr(Time)
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<día>>"
            .Replacement.Text = Day(gdFecSis)
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<mes>>"
            .Replacement.Text = MonthName(Month(gdFecSis))
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<año>>"
            .Replacement.Text = Year(gdFecSis)
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<De>>"
            .Replacement.Text = cDel
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<Al>>"
            .Replacement.Text = cAl
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<participantes>>"
            .Replacement.Text = Replace(lcPersonalInvolucrado, "/", " ")
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        'listando aperturas según sistema
        Set R = ObtenerTotalCuentasAperturas(cDel, cAl, gsCodAge)
        Dim bHaySoles, bHayDolares As Boolean
        Dim nCantidadDatosObtenidos As Integer
        
        bHaySoles = False
        bHayDolares = False
        nCantidadDatosObtenidos = 1 'contador de datos
        
        nTotalAperturasSistMN = 0
        nTotalAperturasSistME = 0
        
        If Not (R.BOF And R.EOF) Then
            R.MoveFirst
            Do While Not R.EOF
                
                If R!Moneda = "ME" Then
                    bHayDolares = True
                    With oWord.Selection.Find
                        .Text = "<<aperdolaho1>>"
                        .Replacement.Text = IIf(R!Ahorro = "", 0, R!Ahorro)
                        .Forward = False
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    With oWord.Selection.Find
                        .Text = "<<aperdolpfi1>>"
                        .Replacement.Text = IIf(R!Plazofijo = "", 0, R!Plazofijo)
                        .Forward = False
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    With oWord.Selection.Find
                        .Text = "<<aperdolcts1>>"
                        .Replacement.Text = IIf(R!CTS = "", 0, R!CTS)
                        .Forward = False
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    
                    nTotalAperturasSistME = R!Total ' guardando total ME
                    
                    With oWord.Selection.Find
                        .Text = "<<aperdoltotal1>>"
                        .Replacement.Text = nTotalAperturasSistME
                        .Forward = False
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                    End With
                Else
                    bHaySoles = True
                     With oWord.Selection.Find
                        .Text = "<<apersolaho1>>"
                        .Replacement.Text = IIf(R!Ahorro = "", 0, R!Ahorro)
                        .Forward = False
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    With oWord.Selection.Find
                        .Text = "<<apersolpfi1>>"
                        .Replacement.Text = IIf(R!Plazofijo = "", 0, R!Plazofijo)
                        .Forward = False
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    With oWord.Selection.Find
                        .Text = "<<apersolcts1>>"
                        .Replacement.Text = IIf(R!CTS = "", 0, R!CTS)
                        .Forward = False
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                    End With
                    
                    nTotalAperturasSistMN = R!Total 'guardando total ME
                    
                    With oWord.Selection.Find
                        .Text = "<<apersoltotal1>>"
                        .Replacement.Text = nTotalAperturasSistMN
                        .Forward = False
                        .Wrap = wdFindContinue
                        .Format = False
                        .Execute Replace:=wdReplaceAll
                    End With
                End If
                nCantidadDatosObtenidos = nCantidadDatosObtenidos + 1
                R.MoveNext
            Loop
        Else
            'no hay datos por tando 0
            With oWord.Selection.Find
                .Text = "<<aperdolaho1>>"
                .Replacement.Text = "0"
                .Forward = False
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<aperdolpfi1>>"
                .Replacement.Text = "0"
                .Forward = False
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<aperdolcts1>>"
                .Replacement.Text = "0"
                .Forward = False
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            With oWord.Selection.Find
                .Text = "<<aperdoltotal1>>"
                .Replacement.Text = "0"
                .Forward = False
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
        
             With oWord.Selection.Find
                .Text = "<<apersolaho1>>"
                .Replacement.Text = "0"
                .Forward = False
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<apersolpfi1>>"
                .Replacement.Text = "0"
                .Forward = False
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            With oWord.Selection.Find
                .Text = "<<apersolcts1>>"
                .Replacement.Text = "0"
                .Forward = False
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
            
            With oWord.Selection.Find
                .Text = "<<apersoltotal1>>"
                .Replacement.Text = "0"
                .Forward = False
                .Wrap = wdFindContinue
                .Format = False
                .Execute Replace:=wdReplaceAll
            End With
        End If
        
        If nCantidadDatosObtenidos = 2 Then
            If bHayDolares Then
                 With oWord.Selection.Find
                    .Text = "<<apersolaho1>>"
                    .Replacement.Text = "0"
                    .Forward = False
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<apersolpfi1>>"
                    .Replacement.Text = "0"
                    .Forward = False
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<apersolcts1>>"
                    .Replacement.Text = "0"
                    .Forward = False
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<apersoltotal1>>"
                    .Replacement.Text = "0"
                    .Forward = False
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
            End If
            
            If bHaySoles Then
                 With oWord.Selection.Find
                    .Text = "<<aperdolaho1>>"
                    .Replacement.Text = "0"
                    .Forward = False
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<aperdolpfi1>>"
                    .Replacement.Text = "0"
                    .Forward = False
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                With oWord.Selection.Find
                    .Text = "<<aperdolcts1>>"
                    .Replacement.Text = "0"
                    .Forward = False
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
                
                With oWord.Selection.Find
                    .Text = "<<aperdoltotal1>>"
                    .Replacement.Text = "0"
                    .Forward = False
                    .Wrap = wdFindContinue
                    .Format = False
                    .Execute Replace:=wdReplaceAll
                End With
            End If
        End If
        
        Set R = Nothing
        
        'listando cuentas conformes
        nTotalCuentasConformeMN = 0
        nTotalCuentasConformeME = 0
        
        aCuentasConformes = ObtenerCuentasConformes
        With oWord.Selection.Find
            .Text = "<<apersolaho2>>"
            .Replacement.Text = aCuentasConformes(0)
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<apersolpfi2>>"
            .Replacement.Text = aCuentasConformes(1)
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<apersolcts2>>"
            .Replacement.Text = aCuentasConformes(2)
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<aperdolaho2>>"
            .Replacement.Text = aCuentasConformes(3)
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<aperdolpfi2>>"
            .Replacement.Text = aCuentasConformes(4)
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<aperdolcts2>>"
            .Replacement.Text = aCuentasConformes(5)
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<apersoltotal2>>"
            .Replacement.Text = aCuentasConformes(6)
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<aperdoltotal2>>"
            .Replacement.Text = aCuentasConformes(7)
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        
        nTotalCuentasConformeMN = aCuentasConformes(6)
        nTotalCuentasConformeME = aCuentasConformes(7)
        
        'resumen arqueo
        nDiferenciaMN = nTotalAperturasSistMN - nTotalCuentasConformeMN
        nDifirenciaME = Abs(nTotalAperturasSistME - nTotalCuentasConformeME)
        
        With oWord.Selection.Find
            .Text = "<<apersistsol>>"
            .Replacement.Text = nTotalAperturasSistMN
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<aperfissol>>"
            .Replacement.Text = nTotalCuentasConformeMN
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
          With oWord.Selection.Find
            .Text = "<<aperdifsol>>"
            .Replacement.Text = nDiferenciaMN
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<apersistdol>>"
            .Replacement.Text = nTotalAperturasSistME
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<aperfisdol>>"
            .Replacement.Text = nTotalCuentasConformeME
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "<<aperdifdol>>"
            .Replacement.Text = nDifirenciaME
            .Forward = False
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
            
        'listando faltantes o no habidos
        nTotalRowsFlxConformidad = flxConformidad.Rows - 1
        cCuenta = ""
        cCliente = ""
        cFechApertura = ""
        
        For i = 1 To nTotalRowsFlxConformidad
            cMotivo = flxConformidad.TextMatrix(i, 11)
            If flxConformidad.TextMatrix(i, 9) = "" And cMotivo = "" Then
                cCuenta = flxConformidad.TextMatrix(i, 1)
                cCliente = flxConformidad.TextMatrix(i, 2)
                cFechApertura = flxConformidad.TextMatrix(i, 3)
                cMoneda = IIf(Mid(flxConformidad.TextMatrix(i, 1), 9, 1) = "1", "MN", "ME")
                ' ordenando según formato
                With oDoc.Application.ActiveDocument
                    .Tables(4).Rows.Add
                    .Tables(4).Cell(i + 1, 1).Range.InsertAfter cCliente
                    .Tables(4).Cell(i + 1, 2).Range.InsertAfter cCuenta
                    .Tables(4).Cell(i + 1, 3).Range.InsertAfter cFechApertura
                    .Tables(4).Cell(i + 1, 4).Range.InsertAfter cMoneda
                End With
            End If
        Next i
            
        'listando cuentas deficientes
        For i = 1 To nTotalRowsFlxConformidad
            cMotivo = flxConformidad.TextMatrix(i, 11)
            If flxConformidad.TextMatrix(i, 9) = "" And cMotivo <> "" Then
                cCuenta = flxConformidad.TextMatrix(i, 1)
                cCliente = flxConformidad.TextMatrix(i, 2)
                cFechApertura = flxConformidad.TextMatrix(i, 3)
                cMoneda = IIf(Mid(flxConformidad.TextMatrix(i, 1), 9, 1) = "1", "MN", "ME")
                ' ordenando según formato
                With oDoc.Application.ActiveDocument
                    .Tables(5).Rows.Add
                    .Tables(5).Cell(i + 1, 1).Range.InsertAfter cCliente
                    .Tables(5).Cell(i + 1, 2).Range.InsertAfter cCuenta
                    .Tables(5).Cell(i + 1, 3).Range.InsertAfter cFechApertura
                    .Tables(5).Cell(i + 1, 4).Range.InsertAfter cMoneda
                End With
            End If
        Next i
        
        'listando cuentas con observaciones
        For i = 1 To nTotalRowsFlxConformidad
            cMotivo = flxConformidad.TextMatrix(i, 11)
            If flxConformidad.TextMatrix(i, 9) = "" And cMotivo <> "" Then
                cCuenta = flxConformidad.TextMatrix(i, 1)
                cCliente = flxConformidad.TextMatrix(i, 2)
                cFechApertura = flxConformidad.TextMatrix(i, 3)
                cMoneda = IIf(Mid(flxConformidad.TextMatrix(i, 1), 9, 1) = "1", "MN", "ME")
                
                ' ordenando según formato
                With oDoc.Application.ActiveDocument
                    .Tables(6).Rows.Add
                    .Tables(6).Cell(i + 1, 1).Range.InsertAfter cCliente
                    .Tables(6).Cell(i + 1, 2).Range.InsertAfter cCuenta
                    .Tables(6).Cell(i + 1, 3).Range.InsertAfter cFechApertura
                    .Tables(6).Cell(i + 1, 4).Range.InsertAfter cMotivo
                    .Tables(6).Cell(i + 1, 5).Range.InsertAfter cMoneda
                End With
            End If
        Next i
        
        'listando cuentas aperturadas mismo mes arqueo
        Set rCuentasApeMesArqueo = oCapta.ObtenerCuentasAperturadas(cDel, cAl, gsCodAge)
        i = 1
        If Not (rCuentasApeMesArqueo.BOF And rCuentasApeMesArqueo.EOF) Then
            rCuentasApeMesArqueo.MoveFirst
            Do While Not rCuentasApeMesArqueo.EOF
                cCliente = rCuentasApeMesArqueo!Nombre
                cCuenta = rCuentasApeMesArqueo!Cuenta
                cFechApertura = rCuentasApeMesArqueo!FechaApe
                cMoneda = rCuentasApeMesArqueo!Moneda
                
                'ordenando según formato
                With oDoc.Application.ActiveDocument
                    .Tables(7).Rows.Add
                    .Tables(7).Cell(i + 1, 1).Range.InsertAfter cCliente
                    .Tables(7).Cell(i + 1, 2).Range.InsertAfter cCuenta
                    .Tables(7).Cell(i + 1, 3).Range.InsertAfter cFechApertura
                    .Tables(7).Cell(i + 1, 4).Range.InsertAfter cMoneda
                End With
                i = i + 1
                rCuentasApeMesArqueo.MoveNext
            Loop
        Else
            MsgBox "No hay cuentas en el mes de arqueo.", vbInformation + vbOKOnly, "Aviso"
        End If
        
        Set rCuentasApeMesArqueo = Nothing
        
        'agregando firmas del personal involucrado
                
        nRowsPersonalInvolucrado = flxPersonalInvolucrado.Rows - 1
        bNuevaFila = False
        nFila = 1
        nColunma = 1
        nPerfil = 0
        nTipoArqueo = CInt(Right(cboTipoArqueo.Text, 1))
        cDescOtro = txtOtroTipoArqueo.Text
        
        For i = 1 To nRowsPersonalInvolucrado
            cPersCod = flxPersonalInvolucrado.TextMatrix(i, 1)
            cNombreInvolucrado = flxPersonalInvolucrado.TextMatrix(i, 3)
            cCargo = flxPersonalInvolucrado.TextMatrix(i, 4)
            cPerfil = flxPersonalInvolucrado.TextMatrix(i, 5)
            
            If cPerfil = "Líder" Then
                nPerfil = 1
            ElseIf cPerfil = "Veedor" Then
                nPerfil = 2
            Else
                nPerfil = 3
            End If
                        
            With oDoc.Application.ActiveDocument
                If bNuevaFila Then
                    .Tables(8).Rows.Add
                End If
                .Tables(8).Cell(nFila, nColunma).Range.InsertAfter vbCrLf & vbCrLf & "------------------------------------------------" & vbCrLf & cNombreInvolucrado & vbCrLf & cCargo
                
            End With
            If nColunma < 2 Then
                nColunma = nColunma + 1
            Else
                nColunma = 1
                bNuevaFila = True
                nFila = nFila + 1
            End If
            'guardando pista arqueo
            Call oCapta.GuardarPistaArqueoExpAho(cCodigoArqueo, cPersCod, nTipoArqueo, cDescOtro, nPerfil, gsCodAge, cFecArqueo)
        
        Next i
       
        oDoc.Close
        Set oDoc = Nothing
        Set oWord = Nothing
        
        Set oWord = CreateObject("Word.Application")
        oWord.Visible = True
        Set oDoc = oWord.Documents.Open(cArchivo)
        Set oDoc = Nothing
        Set oWord = Nothing
        lcPersonalInvolucrado = ""
        
        If MsgBox("Arqueo finalizado con éxito. ¿Desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
            Unload Me
        End If
        
        Exit Sub
ErrorProcesarArchivo:
        Set oDoc = Nothing
        Set oWord = Nothing
        lcPersonalInvolucrado = ""
        MsgBox "No se pudo hacer el arqueo.", vbError + vbOKOnly, "Error"
    End If
End Sub



Private Sub flxConformidad_DblClick()
    Call CombrobarCheckConformidad
End Sub

Private Sub flxConformidad_KeyPress(KeyAscii As Integer)
    If KeyAscii = 3 Or KeyAscii = 22 Then
        KeyAscii = 0
    End If
End Sub

Private Sub flxConformidad_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    If flxConformidad.TextMatrix(pnRow, pnCol) = "." Then
        flxConformidad.TextMatrix(pnRow, pnCol + 1) = ""
        flxConformidad.TextMatrix(pnRow, pnCol + 2) = ""
    End If
End Sub

Private Sub flxConformidad_OnChangeCombo()
    Dim cSel, cLetra As String
    Dim nLongCadena, nRow, nCol As Integer
    nRow = flxConformidad.row
    nCol = flxConformidad.Col
    cSel = flxConformidad.TextMatrix(nRow, nCol)
    If cSel <> "" Then
        cLetra = Right(cSel, 1)
        If EsNumero(cLetra) Then
            lnItemSelComentario = CInt(cLetra)
            nLongCadena = Len(cSel) - 1
            cSel = Trim(Left(cSel, nLongCadena))
            flxConformidad.TextMatrix(nRow, nCol) = cSel
        End If
    End If
End Sub

Private Sub flxPersonalInvolucrado_OnChangeCombo()
    Dim nRow, nCol, nLongCadena, nRows, i, J As Integer
    Dim cSel, cSel2 As String
    Dim bLideresDuplicados As Boolean
    nRow = flxPersonalInvolucrado.row
    nCol = 5
    cSel = flxPersonalInvolucrado.TextMatrix(nRow, nCol)
    nRows = flxPersonalInvolucrado.Rows - 1
    ReDim Preserve laPersonalInvolucrado(nRows)
    
    For i = 1 To nRows
        cSel = flxPersonalInvolucrado.TextMatrix(i, nCol)
        If EsNumero(Right(cSel, 1)) Then
            nLongCadena = Len(cSel) - 1
            cSel = Trim(Left(cSel, nLongCadena))
        End If
        laPersonalInvolucrado(i - 1) = cSel
    Next i
    
    For i = 1 To nRows
        cSel = flxPersonalInvolucrado.TextMatrix(i, nCol)
        If cSel <> "" Then
            If cSel <> "Líder" Then
                nLongCadena = Len(cSel) - 1
                cSel = Trim(Left(cSel, nLongCadena))
            End If
            
            For J = 1 To nRows
                If i <> J Then
                    cSel2 = flxPersonalInvolucrado.TextMatrix(J, nCol)
                    If cSel2 <> "" Then
                        If cSel2 <> "Líder" Then
                            nLongCadena = Len(cSel2) - 1
                            cSel2 = Trim(Left(cSel2, nLongCadena))
                        End If
                        
                        If cSel = cSel2 Then
                            bLideresDuplicados = True
                            Exit For
                        End If
                    End If
                End If
            Next J
        End If
    Next i
    
    If bLideresDuplicados Then
        MsgBox "No puede haber relación duplicada en el arqueo. Quite a uno.", vbInformation, "Aviso"
        flxPersonalInvolucrado.TextMatrix(nRow, nCol) = ""
        Exit Sub
    Else
        cSel = flxPersonalInvolucrado.TextMatrix(nRow, nCol)
        If cSel <> "" Then
            If EsNumero(Right(cSel, 1)) Then
                nLongCadena = Len(cSel) - 1
                cSel = Trim(Left(cSel, nLongCadena))
            End If
            flxPersonalInvolucrado.TextMatrix(nRow, nCol) = cSel
        End If
    End If
End Sub

Private Sub flxPersonalInvolucrado_OnEnterTextBuscar(psDataCod As String, pnRow As Long, pnCol As Long, pbEsDuplicado As Boolean)
    Dim bResultado As Boolean
    Dim oAcceso As New comdpersona.UCOMAcceso
    Dim RDatos As ADODB.Recordset
    Dim oPersona As New comdpersona.DCOMPersonas
    Dim i, iFila, nCantidadFilas As Integer
    
    bResultado = False
    i = flxPersonalInvolucrado.row
    iFila = 1
    nCantidadFilas = flxPersonalInvolucrado.Rows
    
    Set RDatos = oPersona.BuscaCliente(flxPersonalInvolucrado.TextMatrix(i, 1), BusquedaEmpleadoCodigo)
    If Not (RDatos.BOF And RDatos.EOF) Then
        oUsuario.Inicio RDatos!cUser
        
        For iFila = 1 To nCantidadFilas - 1
            If RDatos!cUser = flxPersonalInvolucrado.TextMatrix(iFila, 2) Then
                MsgBox "El usuario ya fue agregado. Solo se permite agregar una ves a los usuarios.", vbInformation, "Aviso"
                flxPersonalInvolucrado.EliminaFila i
                Exit Sub
            End If
        Next iFila
        
        RDatos.MoveFirst
        Do While Not RDatos.EOF
            flxPersonalInvolucrado.TextMatrix(i, 1) = RDatos!cPersCod
            flxPersonalInvolucrado.TextMatrix(i, 2) = RDatos!cUser
            flxPersonalInvolucrado.TextMatrix(i, 3) = RDatos!cPersNombre
            flxPersonalInvolucrado.TextMatrix(i, 4) = oUsuario.PersCargo
            RDatos.MoveNext
        Loop
    Else
        MsgBox "Usuario no válido.", vbInformation + vbOKOnly, "Aviso"
        
        flxPersonalInvolucrado.TextMatrix(nCantidadFilas - 1, 1) = ""
        flxPersonalInvolucrado.TextMatrix(nCantidadFilas - 1, 2) = ""
                
    End If
    
End Sub

Private Sub Form_Load()
    Dim rTipoArqueo As ADODB.Recordset
    Dim rMes As ADODB.Recordset
    
    'configurando controles
    txtOtroTipoArqueo.Enabled = False
    txtOtroTipoArqueo.BackColor = vbGrayed
    flxConformidad.Enabled = False
    cmdProcesar.Enabled = False
    lbConforme = False
    
    'fin bloqueo
    
    ReDim Preserve laPersonalInvolucrado(0) ' tiene memoria reserva para un item, significa no hay nada
    lcPersonalInvolucrado = ""
    lblAgencia.Caption = gsNomAge
    lblFechaHora.Caption = CStr(gdFecSis) & " " & CStr(Time)
    
    Set rTipoArqueo = LeerConstante("1109")
    Set rMes = LeerConstante("1010")
    
    If Not (rTipoArqueo.BOF And rTipoArqueo.EOF) Then
        Do While Not rTipoArqueo.EOF
            cboTipoArqueo.AddItem (rTipoArqueo!cConstante)
            rTipoArqueo.MoveNext
        Loop
        cboTipoArqueo.Text = cboTipoArqueo.List(0)
    Else
        MsgBox "No se encontró datos"
    End If
    
    Call ObtenerUsuariosParticipantes
    'agregando valor a los date rangos
    txtFecIni = gdFecSis - 1
    txtFecFin = gdFecSis
End Sub


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
    If KeyCode = 113 And Shift = 0 Then
        KeyCode = 10
    End If
End Sub



Public Function LeerConstante(ByVal cConsValor As String) As ADODB.Recordset
    Dim R As New ADODB.Recordset
    Dim oConstante As New COMDConstantes.DCOMConstantes
    
    On Error GoTo ErrorCargaConstante
    Set R = oConstante.ObtenerConstante(cConsValor)
    Set LeerConstante = R
    Exit Function
ErrorCargaConstante:
    err.Raise err.Number, "Error al leer constante", err.Description
End Function

Public Sub ObtenerUsuariosParticipantes()
    Dim oCap As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim cUserLogeado As String
    Dim cUserVisto As String
    Dim cUserParticipante As String
    Dim RDatos As ADODB.Recordset
    Dim i As Integer
    Dim rRelConArqueo As ADODB.Recordset
    Set rRelConArqueo = LeerConstante("1110")
    flxPersonalInvolucrado.CargaCombo rRelConArqueo
        
    cUserVisto = gcUsuarioVistoArqExpAho
    cUserLogeado = gsCodUser
    For i = 1 To 2
        If i = 1 Then
            cUserParticipante = cUserVisto
        Else
            cUserParticipante = cUserLogeado
        End If
        Set RDatos = oCap.ObtenerDatosParticipantes(cUserParticipante)
        
        If Not (RDatos.BOF And RDatos.EOF) Then
            RDatos.MoveFirst
            flxPersonalInvolucrado.AdicionaFila
            flxPersonalInvolucrado.TextMatrix(i, 1) = RDatos!codigo
            flxPersonalInvolucrado.TextMatrix(i, 2) = RDatos!Usuario
            flxPersonalInvolucrado.TextMatrix(i, 3) = RDatos!Nombre
            flxPersonalInvolucrado.TextMatrix(i, 4) = RDatos!Cargo
            
        End If
    Next i
    
End Sub

Public Sub CargarCombosGridflxConformidad(ByVal pnCombo As Integer)
    Dim rDatosMotivo, rDatosComentario, rVacio As ADODB.Recordset
    
    Set rDatosMotivo = LeerConstante("1111")
    Set rDatosComentario = LeerConstante("1112")
    If pnCombo = 1 Then
        flxConformidad.CargaCombo rDatosMotivo
    ElseIf pnCombo = 2 Then
        flxConformidad.CargaCombo rDatosComentario
    Else
        flxConformidad.LimpiarCombo
    End If
End Sub

Public Sub CombrobarCheckConformidad()
    Dim nRow, nCol As Integer
    nRow = flxConformidad.row
    nCol = flxConformidad.Col
    
    If nCol = 10 Or nCol = 11 Then
                
        If flxConformidad.TextMatrix(nRow, 9) = "." Then
            lbConforme = True
        Else
            lbConforme = False
        End If
        
        If nCol = 10 Then
            If lbConforme = False Then
                CargarCombosGridflxConformidad 1
            Else
                MsgBox "No puede agregar un comentario porque dio conformidad al registro.", vbInformation, "Aviso"
            End If
            Exit Sub
        End If
        
        If nCol = 11 Then
            If lbConforme = False Then
                If flxConformidad.TextMatrix(nRow, 10) = "" Then
                    MsgBox "Debe seleccionar el comentario.", vbInformation, "Aviso"
                    CargarCombosGridflxConformidad 3
                ElseIf Trim(flxConformidad.TextMatrix(nRow, 10)) = "Faltante" Then
                    MsgBox "Comentario no válido para agregar motivos.", vbInformation, "Aviso"
                    CargarCombosGridflxConformidad 3
                ElseIf lnItemSelComentario = 2 Or Trim(flxConformidad.TextMatrix(nRow, 10)) = "Presenta deficiencia" Then
                    CargarCombosGridflxConformidad 2
                Else
                    MsgBox "No hay motivos para el comentario elegido.", vbInformation, "Aviso"
                    CargarCombosGridflxConformidad 3
                End If
            Else
                MsgBox "No puede agregar un motivo porque dio conformidad al registro.", vbInformation, "Aviso"
                CargarCombosGridflxConformidad 3
            End If
            Exit Sub
        End If
        
    End If
End Sub

Public Function EsNumero(ByVal pcValor As String) As Boolean
    Dim cABC As String
    Dim nLongABC, i As Integer
    Dim bEsNumero As Boolean
    cABC = "abcdefghijklmnñopqrstuvwxyz"
    bEsNumero = False
    
    If InStr(1, cABC, LCase(pcValor)) = 0 Then
        bEsNumero = True
    End If
    EsNumero = bEsNumero
End Function

Public Function Validar() As Boolean
    'validar campos de arqueo
    Validar = True
    If lblFechaHora.Caption = "" Then
        MsgBox "Fecha no válida.", vbInformation + vbOKOnly, "Aviso"
        Validar = False
        Exit Function
    End If
    
    If lnItemSelTipoArqueo = -1 Then
        MsgBox "Tipo de arqueo no válido.", vbInformation + vbOKOnly, "Aviso"
        Validar = False
        cboTipoArqueo.SetFocus
        Exit Function
    Else
        If lnItemSelTipoArqueo = 3 And txtOtroTipoArqueo = "" Then
            MsgBox "Debe especificar el tipo de arqueo.", vbInformation + vbOKOnly, "Aviso"
            Validar = False
            txtOtroTipoArqueo.SetFocus
            Exit Function
        End If
    End If
    'validar personal involucrado
    If lblAgencia.Caption = "" Then
        MsgBox "Agencia no válida.", vbInformation, "Aviso"
        Validar = False
        Exit Function
    End If
    
    Dim nInvolucrados, i As Integer
    Dim bHayLider As Boolean
    bHayLider = False
        
    If UBound(laPersonalInvolucrado) > 0 Then ' 0 significa no hay cambios
        nInvolucrados = UBound(laPersonalInvolucrado) - 1
        For i = 0 To nInvolucrados
            If laPersonalInvolucrado(i) = "" Then
                MsgBox "Un usuario no tiene relación con el arqueo, debe asignarle una.", vbInformation + vbOKOnly, "Aviso"
                Validar = False
                flxPersonalInvolucrado.SetFocus
                Exit Function
            Else
                If laPersonalInvolucrado(i) = "Líder" Then
                    bHayLider = True
                End If
                lcPersonalInvolucrado = lcPersonalInvolucrado & flxPersonalInvolucrado.TextMatrix(i + 1, 3) & ";"
            End If
        Next i
    Else
        nInvolucrados = flxPersonalInvolucrado.Rows - 1
        For i = 1 To nInvolucrados
            If flxPersonalInvolucrado.TextMatrix(i, 5) = "" Then
                MsgBox "Un usuario no tiene relación con el arqueo, debe asignarle una.", vbInformation + vbOKOnly, "Aviso"
                Validar = False
                flxPersonalInvolucrado.SetFocus
                Exit Function
            Else
                If flxPersonalInvolucrado.TextMatrix(i, 5) = "Líder" Then
                    bHayLider = True
                End If
                lcPersonalInvolucrado = lcPersonalInvolucrado & flxPersonalInvolucrado.TextMatrix(i, 3) & ";"
            End If
            
        Next i
    End If
    
    If bHayLider Then
        lcPersonalInvolucrado = Mid(lcPersonalInvolucrado, 1, Len(lcPersonalInvolucrado) - 1) + "."
    Else
        MsgBox "Es necesario que un usuario sea líder.", vbInformation + vbOKOnly, "Aviso"
        lcPersonalInvolucrado = ""
        Validar = False
        Exit Function
    End If
   
    'validar conformidad
    
    Dim nCuentasAperturadas As Integer
    
    nCuentasAperturadas = flxConformidad.Rows
    
    For i = 1 To nCuentasAperturadas - 1
        If flxConformidad.TextMatrix(i, 9) = "." Then
            If flxConformidad.TextMatrix(i, 10) <> "" Then
                MsgBox "No puede poner comentario a una cuenta si ya dio conformidad a la misma.", vbInformation, "Aviso"
                flxConformidad.SetFocus
                Validar = False
                Exit Function
            End If
            If flxConformidad.TextMatrix(i, 11) <> "" Then
                MsgBox "No puede poner motivo a una cuenta si ya dio conformidad a la misma.", vbInformation, "Aviso"
                Validar = False
                flxConformidad.SetFocus
                Exit Function
            End If
        Else
            If flxConformidad.TextMatrix(i, 10) = "" Then
                MsgBox "Debe poner comentario a una cuenta si no dio conformidad a la misma.", vbInformation, "Aviso"
                Validar = False
                Exit Function
            ElseIf Trim(flxConformidad.TextMatrix(i, 10)) = "Presenta deficiencia" Then
                If flxConformidad.TextMatrix(i, 11) = "" Then
                    MsgBox "Debe poner motivo a una cuenta si no dio conformidad a la misma, y si eligió como comentario " _
                    & Chr(34) & "Presenta deficiencia" & Chr(34) & ".", vbInformation, "Aviso"
                    
                    Validar = False
                    flxConformidad.SetFocus
                    Exit Function
                End If
            Else
                If flxConformidad.TextMatrix(i, 11) <> "" Then
                    MsgBox "No puede poner motivo a una cuenta si elegió como comentario " & Chr(34) & "Faltante" & Chr(34) & " .", vbInformation, "Aviso"
                    Validar = False
                    flxConformidad.SetFocus
                    Exit Function
                End If
            End If
        End If
    Next i
End Function


Public Function ObtenerTotalCuentasAperturas(ByVal pcDel As String, ByVal pcAl As String, ByVal pcAgeCod As String) As ADODB.Recordset
    Dim oCapta As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim R As ADODB.Recordset
    
    Set R = oCapta.ObtenerTotalCuentasAperturadas(pcDel, pcAl, pcAgeCod)
    Set ObtenerTotalCuentasAperturas = R
    Set R = Nothing
    
End Function

Public Function ObtenerCuentasConformes() As Variant
    Dim i, nTotalAhorroMN, nTotalPFMN, nTotalCTSMN, nTotalAhorroME, nTotalPFME, nTotalCTSME As Integer
    Dim aCuentasConformes() As Variant
    
    lrDatos.MoveFirst
    i = 1
    nTotalAhorroMN = 0
    nTotalPFMN = 0
    nTotalCTSMN = 0
    nTotalAhorroME = 0
    nTotalPFME = 0
    nTotalCTSME = 0
        
    ReDim Preserve aCuentasConformes(8)
    
    Do While Not lrDatos.EOF
        If Trim(lrDatos!Producto) = "Ahorro" Then
            If Trim(lrDatos!Moneda) = "MN" And flxConformidad.TextMatrix(i, 9) = "." Then
                nTotalAhorroMN = nTotalAhorroMN + 1
            End If
            If Trim(lrDatos!Moneda) = "ME" And flxConformidad.TextMatrix(i, 9) = "." Then
                nTotalAhorroME = nTotalAhorroME + 1
            End If
        End If
        If Trim(lrDatos!Producto) = "Plazo fijo" Then
            If Trim(lrDatos!Moneda) = "MN" And flxConformidad.TextMatrix(i, 9) = "." Then
                nTotalPFMN = nTotalPFMN + 1
            End If
            If lrDatos!Moneda = "ME" And flxConformidad.TextMatrix(i, 9) = "." Then
                nTotalPFME = nTotalPFME + 1
            End If
        End If
        If lrDatos!Producto = "CTS" Then
            If lrDatos!Moneda = "MN" And flxConformidad.TextMatrix(i, 9) = "." Then
                nTotalCTSMN = nTotalCTSMN + 1
            End If
            If lrDatos!Moneda = "ME" And flxConformidad.TextMatrix(i, 9) = "." Then
                nTotalCTSME = nTotalCTSME + 1
            End If
        End If
        i = i + 1
        lrDatos.MoveNext
    Loop
    
    aCuentasConformes(0) = nTotalAhorroMN
    aCuentasConformes(1) = nTotalPFMN
    aCuentasConformes(2) = nTotalCTSMN
    aCuentasConformes(3) = nTotalAhorroME
    aCuentasConformes(4) = nTotalPFME
    aCuentasConformes(5) = nTotalCTSME
    aCuentasConformes(6) = nTotalAhorroMN + nTotalPFMN + nTotalCTSMN
    aCuentasConformes(7) = nTotalAhorroME + nTotalPFME + nTotalCTSME
        
    ObtenerCuentasConformes = aCuentasConformes
    
End Function

Public Function LimpiarFlxConformidad()
    Dim i, nRows As Integer
    nRows = flxConformidad.Rows - 1
    
    For i = 1 To nRows
        flxConformidad.EliminaFila (nRows - 1)
        nRows = flxConformidad.Rows
    Next i
    
End Function


Private Sub TxtFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdConsultar.SetFocus
    End If
End Sub

Private Sub TxtFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecFin.SetFocus
    End If
End Sub

Private Sub txtOtroTipoArqueo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtFecIni.SetFocus
    End If
End Sub

Private Function ValidarFechas(ByVal dateIni As Date, ByVal dateFin As Date) As Boolean
    If (dateIni = dateFin) Then
        MsgBox "Fechas válidas, pero no pueden ser iguales.", vbInformation + vbOKOnly, "Aviso"
        ValidarFechas = False
        Exit Function
    ElseIf (Year(dateFin) < Year(dateIni)) Then
        MsgBox "Fecha inválida. Fecha de fin no debe ser menor a la fecha de inicio.", vbInformation + vbOKOnly, "Aviso"
        ValidarFechas = False
        Exit Function
    ElseIf (Month(dateFin) < Month(dateIni)) And (Year(dateFin) <= Year(dateIni)) Then
        MsgBox "Fecha inválida. Fecha de fin no debe ser menor a la fecha de inicio.", vbInformation + vbOKOnly, "Aviso"
        ValidarFechas = False
        Exit Function
    ElseIf (Day(dateFin) < Day(dateIni)) And (Month(dateFin) <= Month(dateIni)) And (Year(dateFin) <= Year(dateIni)) Then
        MsgBox "Fecha inválida. Fecha de fin no debe ser menor a la fecha de inicio.", vbInformation + vbOKOnly, "Aviso"
        ValidarFechas = False
        Exit Function
    Else
        ValidarFechas = True
    End If
End Function
'END ANDE ERS021-2017
