VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAFBajaActivo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Baja de Activo Fijo"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12240
   Icon            =   "frmAFBajaActivo.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   12240
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExportar 
      Caption         =   "&Exportar"
      Height          =   360
      Left            =   8520
      TabIndex        =   20
      Top             =   6600
      Width           =   1125
   End
   Begin VB.CheckBox chkEstadistico 
      Caption         =   "Solo estadístico / No genera asiento contable"
      Height          =   435
      Left            =   360
      TabIndex        =   17
      Top             =   6600
      Width           =   3255
   End
   Begin VB.Frame fraOpe 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      Caption         =   "Baja de Activo Fijo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   6525
      Left            =   45
      TabIndex        =   2
      Top             =   0
      Width           =   12120
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   360
         Left            =   10890
         TabIndex        =   25
         Top             =   920
         Width           =   1125
      End
      Begin VB.CheckBox chkTodos 
         Caption         =   "Todos"
         Height          =   255
         Left            =   10680
         TabIndex        =   16
         Top             =   240
         Visible         =   0   'False
         Width           =   1335
      End
      Begin MSMask.MaskEdBox mskFecha 
         Height          =   315
         Left            =   1005
         TabIndex        =   9
         Top             =   5565
         Width           =   1410
         _ExtentX        =   2487
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.TextBox txtComentario 
         Appearance      =   0  'Flat
         Height          =   450
         Left            =   1005
         MaxLength       =   300
         MultiLine       =   -1  'True
         TabIndex        =   3
         Top             =   5985
         Width           =   9105
      End
      Begin Sicmact.TxtBuscar txtBS 
         Height          =   345
         Left            =   1440
         TabIndex        =   4
         Top             =   600
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         Appearance      =   0
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
      End
      Begin Sicmact.FlexEdit FeAdj 
         Height          =   4095
         Left            =   120
         TabIndex        =   10
         Top             =   1360
         Width           =   11895
         _ExtentX        =   20981
         _ExtentY        =   7223
         Cols0           =   16
         HighLight       =   1
         AllowUserResizing=   1
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmAFBajaActivo.frx":030A
         EncabezadosAnchos=   "400-400-1800-3500-1200-1200-1200-1200-800-800-0-0-1200-0-0-0"
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
         ColumnasAEditar =   "X-1-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-4-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-R-R-R-R-C-C-C-C-C-R-C-C"
         FormatosEdit    =   "0-0-0-0-5-2-2-2-0-0-0-0-0-3-0-0"
         AvanceCeldas    =   1
         TextArray0      =   "Nº"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin Sicmact.TxtBuscar txtAreaAgeCod 
         Height          =   345
         Left            =   1440
         TabIndex        =   18
         Top             =   240
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         Appearance      =   0
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
      End
      Begin Sicmact.TxtBuscar txtSerieCod 
         Height          =   345
         Left            =   1440
         TabIndex        =   22
         Top             =   960
         Width           =   2055
         _ExtentX        =   3625
         _ExtentY        =   609
         Appearance      =   0
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
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Serie :"
         Height          =   195
         Left            =   240
         TabIndex        =   24
         Top             =   960
         Width           =   450
      End
      Begin VB.Label lblSerieNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3600
         TabIndex        =   23
         Top             =   960
         Width           =   4860
      End
      Begin VB.Label lblAreaAgeNombre 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3600
         TabIndex        =   21
         Top             =   240
         Width           =   4860
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Área/Agencia :"
         Height          =   195
         Left            =   240
         TabIndex        =   19
         Top             =   270
         Width           =   1080
      End
      Begin VB.Label lblFecha 
         Caption         =   "Fecha"
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   5610
         Width           =   810
      End
      Begin VB.Label lblComentario 
         Caption         =   "Coment."
         Height          =   210
         Left            =   120
         TabIndex        =   7
         Top             =   6000
         Width           =   780
      End
      Begin VB.Label lblBienG 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   3600
         TabIndex        =   6
         Top             =   600
         Width           =   4860
      End
      Begin VB.Label lblBien 
         AutoSize        =   -1  'True
         Caption         =   "Bien :"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   10905
      TabIndex        =   1
      Top             =   6600
      Width           =   1125
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   9705
      TabIndex        =   0
      Top             =   6600
      Width           =   1125
   End
   Begin Sicmact.TxtBuscar txtSerie 
      Height          =   315
      Left            =   1200
      TabIndex        =   11
      Top             =   7800
      Width           =   3360
      _ExtentX        =   5927
      _ExtentY        =   556
      Appearance      =   0
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
      TipoBusqueda    =   2
      lbUltimaInstancia=   0   'False
   End
   Begin Sicmact.TxtBuscar txtAgeO 
      Height          =   345
      Left            =   5565
      TabIndex        =   13
      Top             =   7800
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   609
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
      Enabled         =   0   'False
      Appearance      =   0
      EnabledText     =   0   'False
   End
   Begin VB.Label lblAgeOG 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Enabled         =   0   'False
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7185
      TabIndex        =   15
      Top             =   7815
      Width           =   1260
   End
   Begin VB.Label lblAgeO 
      Caption         =   "Agencia O"
      Height          =   180
      Left            =   4680
      TabIndex        =   14
      Top             =   7875
      Width           =   840
   End
   Begin VB.Label lblSerie 
      Caption         =   "Serie :"
      Height          =   225
      Left            =   120
      TabIndex        =   12
      Top             =   7800
      Width           =   810
   End
End
Attribute VB_Name = "frmAFBajaActivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lnMovNroIni As Long
Dim lnAnio As Long
Dim I As Integer
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

'EJVG20130627 ***
Private Sub cmdBuscar_Click()
    Dim oBien As New DBien
    Dim rsSeries As New ADODB.Recordset
    Dim lsAreaAgeCod As String
    
    If txtAreaAgeCod.Text <> "" Then
        lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
    End If
    
    FeAdj.Clear
    FeAdj.FormaCabecera
    FeAdj.Rows = 2
    
    Set rsSeries = oBien.GetAFBienesPaBaja(lsAreaAgeCod, txtBS.Text, txtSerieCod.Text)
    Set oBien = Nothing
    If rsSeries.EOF Then
        MsgBox "Este Activo no tiene Series creadas.", vbInformation, "Mensaje"
        Exit Sub
    End If
    FeAdj.rsFlex = rsSeries
    MsgBox "Se cargaron los datos satisfactoriamente.", vbOKOnly + vbInformation, "Aviso"
End Sub
Private Sub cmdExportar_Click()
    Dim xlsAplicacion As New Excel.Application
    Dim xlsLibro As Excel.Workbook
    Dim xlsHoja As Excel.Worksheet
    Dim lnFila As Long, lnColumna As Long, lnColumnaMax As Long
    Dim I As Long, J As Long
    Dim lsArchivo As String
    Dim bOK As Boolean
    
On Error GoTo ErrExportar
    
    If FlexVacio(FeAdj) Then
        MsgBox "No hay información para exportar a formato Excel", vbInformation, "Aviso"
        Exit Sub
    Else 'Se haya seleccionado registros
        For I = 0 To FeAdj.Rows - 1
            If FeAdj.TextMatrix(I, 1) = "." Then 'OK
                bOK = True
                Exit For
            End If
        Next
        If Not bOK Then
            MsgBox "No hay información para exportar a formato Excel", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    Screen.MousePointer = 11
    
    lsArchivo = "\spooler\RptBajaActivos" & UCase(gsCodUser) & "_" & Format(gdFecSis, "YYYYMMDD") & "_" & Format$(Time(), "HHMMSS") & ".xlsx"
    Set xlsLibro = xlsAplicacion.Workbooks.Add

    Set xlsHoja = xlsLibro.Worksheets.Add
    xlsHoja.Name = "Reporte Baja de Activos"
    xlsHoja.Cells.Font.Name = "Arial"
    xlsHoja.Cells.Font.Size = 9
    
    lnFila = 2
    
    For I = 0 To FeAdj.Rows - 1
        lnColumna = 2
        If I = 0 Or (I > 0 And FeAdj.TextMatrix(I, 1) = ".") Then 'OK
            For J = 0 To FeAdj.Cols - 1
                If J > 1 And FeAdj.ColWidth(J) > 0 Then
                    xlsHoja.Cells(lnFila, lnColumna) = FeAdj.TextMatrix(I, J)
                    lnColumna = lnColumna + 1
                    lnColumnaMax = lnColumna
                End If
            Next
            lnFila = lnFila + 1
        End If
    Next

    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).Interior.Color = RGB(191, 191, 191)
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).Font.Bold = True
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(2, lnColumnaMax - 1)).HorizontalAlignment = xlCenter
    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(lnFila - 1, lnColumnaMax - 1)).Borders.Weight = xlThin

    xlsHoja.Range(xlsHoja.Cells(2, 2), xlsHoja.Cells(lnFila - 1, lnColumnaMax - 1)).EntireColumn.AutoFit
    
    MsgBox "Se ha exportado satisfactoriamente la información", vbInformation, "Aviso"
    
    xlsHoja.SaveAs App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    
    Screen.MousePointer = 0
    
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlsHoja = Nothing
        
        'ARLO 20160126 ***
        gsopecod = LogPistaReportesActivoFijo
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Exportaron los Activos Fijos "
        Set objPista = Nothing
        '**************
        
    Exit Sub
ErrExportar:
    MsgBox Err.Description, vbCritical, "Aviso"
End Sub
'END EJVG *******
Private Sub cmdGrabar_Click()
    Dim oMov As DMov
    Set oMov = New DMov
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim lsMovNro As String
    Dim lnMovNro As Long
    Dim lsCtaCont As String '*** PEAC 20120413
    Dim lsCtaDH As String
    Dim lsCtaOtro As String
    Dim lsCtaDebeOtro As String
    Dim lsCtaHaberOtro As String
    Dim I As Integer
    Dim lsMovNroR As String
    Dim oPrevio As clsPrevio
    Dim oAsiento As NContImprimir
    Set oPrevio = New clsPrevio
    Set oAsiento = New NContImprimir
    Dim lnItem As Integer
    Dim lnMovNroR As Long
    Dim lcCuentaSeleccionados As Integer
    Dim lsDetalle As String
    Dim nLin As Integer
    Dim lnValorI, lnValorDepre, lnValorPorDepre As Double
    Dim lsTexto As String
    Dim J As Integer
    Dim lnCuenta As Integer
    'EJVG20130719 ***
    Dim lnMontoDeterioro As Currency
    Dim lnValorResidual As Integer
    Dim lsCtaDebeOtro2 As String
    'END EJVG *******
    Dim RsAF As ADODB.Recordset
    Set RsAF = New ADODB.Recordset
    Dim lsValFecha As String
    
    lsValFecha = ValidaFecha(mskFecha)
    If Len(lsValFecha) > 0 Then
        MsgBox lsValFecha, vbInformation, "Aviso"
        mskFecha.SetFocus
        Exit Sub
    End If
    If Len(Trim(txtComentario.Text)) = 0 Then
        MsgBox "Debe ingresar un comentario valido.", vbInformation, "Aviso"
        Me.txtComentario.SetFocus
        Exit Sub
    End If
    '-----------------------------------------------
    lcCuentaSeleccionados = 0
    For I = 1 To Me.FeAdj.Rows - 1
        If FeAdj.TextMatrix(I, 1) = "." Then
            lcCuentaSeleccionados = lcCuentaSeleccionados + 1
        End If
    Next
    If lcCuentaSeleccionados = 0 Then
        MsgBox "Seleccione al menos un bien para continuar.", vbInformation, "Aviso"
        Exit Sub
    End If
    '-----------------------------------------------
    lnCuenta = 0
    For I = 1 To Me.FeAdj.Rows - 1
        If FeAdj.TextMatrix(I, 1) = "." Then
            lnCuenta = lnCuenta + 1
        End If
    Next
    If lnCuenta > 99 Then
        MsgBox "Solo puede dar de baja hasta 99 items, procese y vuelva a seleccionar los items.", vbInformation, "Aviso"
        Exit Sub
    End If
    '-----------------------------------------------

    If MsgBox("¿Desea dar de Baja los Activos Fijos seleccionados " & IIf(Me.chkEstadistico.value = 1, "SIN generar Asiento Contable.", "y generar sus asientos contables") & "?", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oMov.BeginTrans
    
'        lsMovNroR = oMov.GeneraMovNro(CDate(Me.mskFecha.Text), Right(gsCodAge, 2), gsCodUser)
'        oMov.InsertaMov lsMovNroR, gnBajaAF, "Baja de Activos Fijos - Resumen ", 10
'        lnMovNroR = oMov.GetnMovNro(lsMovNroR)

'        oMov.GeneraAsientoRes lnMovNroR, lnMovNro
'        oMov.InsertaMovRef lnMovNroR, lnMovNro
        
'        lsMovNro = oMov.GeneraMovNro(CDate(Me.mskFecha.Text), Right(gsCodAge, 2), gsCodUser)
'        oMov.InsertaMov lsMovNro, gnBajaAF, Me.txtComentario.Text, 13
'        lnMovNro = oMov.GetnMovNro(lsMovNro)
        
'        oMov.InsertaMovBSAF lnAnio, lnMovNroIni, 1, Me.txtBS.Text, Me.txtSerie.Text, lnMovNro
'        oALmacen.AFActualizaBaja lnAnio, Me.txtBS.Text, Me.txtSerie.Text, CDate(Me.mskFecha.Text)
        
        '*** PEAC 20120413
        'lnItem = 0
        lsCtaDH = "H"
        For I = 1 To Me.FeAdj.Rows - 1
            If FeAdj.TextMatrix(I, 1) = "." Then
                                                
'                For j = 1 To 10000 * 3
'                Next
                                                
                lsMovNro = oMov.GeneraMovNro(CDate(Me.mskFecha.Text), Right(gsCodAge, 2), gsCodUser)
                oMov.InsertaMov lsMovNro, gnBajaAF, Me.txtComentario.Text, 10
                lnMovNro = oMov.GetnMovNro(lsMovNro)
                
                lnItem = 0
                lnItem = lnItem + 1
                
                oMov.InsertaMovBSAF FeAdj.TextMatrix(I, 11), FeAdj.TextMatrix(I, 10), lnItem, FeAdj.TextMatrix(I, 12), FeAdj.TextMatrix(I, 2), lnMovNro
                oALmacen.AFActualizaBaja FeAdj.TextMatrix(I, 11), FeAdj.TextMatrix(I, 12), FeAdj.TextMatrix(I, 2), CDate(Me.mskFecha.Text)
                'EJVG20130719 ***
                lnMontoDeterioro = CCur(FeAdj.TextMatrix(I, 14))
                lnValorResidual = CInt(FeAdj.TextMatrix(I, 15))
                'END EJVG *******
                '*** PEAC 20120510
                If Me.chkEstadistico.value = 0 Then
                
                    Set RsAF = oMov.BuscaCtaPlantillaAF(gnBajaAF, Left(FeAdj.TextMatrix(I, 2), 6), lsCtaDH)
    
                    lsCtaCont = IIf(lsCtaDH = "H", RsAF!cCtaContCodH, RsAF!cCtaContCodD)
                    lsCtaCont = Replace(lsCtaCont, "AG", Right(FeAdj.TextMatrix(I, 9), 2))
    
                    lsCtaOtro = IIf(lsCtaDH = "D", RsAF!cCtaContCodH, RsAF!cCtaContCodD)
                    lsCtaOtro = Replace(lsCtaOtro, "AG", Right(FeAdj.TextMatrix(I, 9), 2))
    
                    lsCtaDebeOtro = IIf(IsNull(RsAF!cCtaContCodOtroD), "", RsAF!cCtaContCodOtroD)
                    lsCtaDebeOtro = Replace(lsCtaDebeOtro, "AG", Right(FeAdj.TextMatrix(I, 9), 2))
    
    '                lsCtaHaberOtro = IIf(IsNull(RsAF!cCtaContCodOtroH), "", RsAF!cCtaContCodOtroH)
    '                lsCtaHaberOtro = Replace(lsCtaHaberOtro, "AG", Right(RsAF!cAgeCod, 2))
    
                    oMov.InsertaMovCta lnMovNro, lnItem, lsCtaCont, Round(FeAdj.TextMatrix(I, 5) * -1, 2)
                    lnItem = lnItem + 1
                    oMov.InsertaMovCta lnMovNro, lnItem, lsCtaOtro, Round(FeAdj.TextMatrix(I, 6), 2)
                    lnItem = lnItem + 1
                    'EJVG20130719 ***
                    'Teóricamente el Monto de Deterioro es igual a Round(FeAdj.TextMatrix(i, 7), 2) - lnValorResidual
                    If lnMontoDeterioro > 0 Then
                        lsCtaDebeOtro2 = IIf(IsNull(RsAF!cCtaContCodOtro2D), "", RsAF!cCtaContCodOtro2D)
                        lsCtaDebeOtro2 = Replace(lsCtaDebeOtro2, "AG", Right(FeAdj.TextMatrix(I, 9), 2))
                        oMov.InsertaMovCta lnMovNro, lnItem, lsCtaDebeOtro2, Round(FeAdj.TextMatrix(I, 7), 2) - lnValorResidual
                        lnItem = lnItem + 1
                        oMov.InsertaMovCta lnMovNro, lnItem, lsCtaDebeOtro, lnValorResidual
                    Else
                        oMov.InsertaMovCta lnMovNro, lnItem, lsCtaDebeOtro, Round(FeAdj.TextMatrix(I, 7), 2)
                    End If
                    'END EJVG *******
'                    oMov.GeneraAsientoRes lnMovNroR, lnMovNro
'                    oMov.InsertaMovRef lnMovNroR, lnMovNro

                     FeAdj.TextMatrix(I, 13) = lnMovNro
                     
                    'lsTexto = oAsiento.ImprimeAsientoContable(lsMovNro, 60, 80, Caption)
                Else
                    FeAdj.TextMatrix(I, 13) = 0
                End If
                '*** FIN PEAC

            End If
        Next
        
        '*** PEAC 20120510
'        If Me.chkEstadistico.value = 0 Then
'            lsMovNroR = oMov.GeneraMovNro(CDate(Me.mskFecha.Text), Right(gsCodAge, 2), gsCodUser)
'            'lsMovNro = oMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
'            oMov.InsertaMov lsMovNroR, gnBajaAF, "Baja de Activos Fijos - Resumen ", 10
'            lnMovNroR = oMov.GetnMovNro(lsMovNroR)
'            oMov.GeneraAsientoRes lnMovNroR, lnMovNro
'            oMov.InsertaMovRef lnMovNroR, lnMovNro
'        End If
        '*** FIN PEAC
        
    oMov.CommitTrans
    
    For I = 1 To Me.FeAdj.Rows - 1
        If CDbl(Val(FeAdj.TextMatrix(I, 13))) > 0 Then
            'lsTexto = lsTexto + oAsiento.ImprimeAsientoContable(lsMovNro, 60, 80, Caption, "Serie:" & FeAdj.TextMatrix(I, 2))
            lsTexto = lsTexto + oAsiento.ImprimeAsientoContable(oMov.GetcMovNro(CDbl(FeAdj.TextMatrix(I, 13))), 60, 80, Caption, "Serie:" & FeAdj.TextMatrix(I, 2))
        End If
    Next
    If Len(lsTexto) > 0 Then
        lsTexto = lsTexto + oImpresora.gPrnSaltoPagina
    End If
    ''oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNroR, 60, 80, Caption + "-Ref:" + lsMovNro), Caption, True
    
    lsDetalle = oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsDetalle = lsDetalle & "  DETALLE DE ACTIVOS FIJOS DADOS DE BAJA" & oImpresora.gPrnSaltoLinea
    lsDetalle = lsDetalle & "  --------------------------------------" & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea
    lsDetalle = lsDetalle & "   SERIE               DESCRIPCION                   FECHA           VALOR INI.    VALOR DEPRE.      POR DEPRE.    COD.BIEN     MOV.CONT." & oImpresora.gPrnSaltoLinea
    lsDetalle = lsDetalle & "  ---------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
'    nLin = 5
    lnValorI = 0
    lnValorDepre = 0
    lnValorPorDepre = 0

    For I = 1 To Me.FeAdj.Rows - 1
        If FeAdj.TextMatrix(I, 1) = "." Then
        lsDetalle = lsDetalle & Space(3) & Justifica(FeAdj.TextMatrix(I, 2), 20) & Justifica(FeAdj.TextMatrix(I, 3), 30) & Justifica(FeAdj.TextMatrix(I, 4), 10) & Right(Space(16) & Format(FeAdj.TextMatrix(I, 5), "##,###,##0.00"), 16) & Right(Space(16) & Format(FeAdj.TextMatrix(I, 6), "##,###,##0.00"), 16) & Right(Space(16) & Format(FeAdj.TextMatrix(I, 7), "##,###,##0.00"), 16) & Space(2) & Justifica(FeAdj.TextMatrix(I, 12), 20) & IIf(FeAdj.TextMatrix(I, 13) = 0, "", oMov.GetcMovNro(FeAdj.TextMatrix(I, 13))) & oImpresora.gPrnSaltoLinea
'        nLin = nLin + 1
        lnValorI = lnValorI + CDbl(FeAdj.TextMatrix(I, 5))
        lnValorDepre = lnValorDepre + CDbl(FeAdj.TextMatrix(I, 6))
        lnValorPorDepre = lnValorPorDepre + CDbl(FeAdj.TextMatrix(I, 7))
        End If
    Next
    lsDetalle = lsDetalle & "  ---------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea
    lsDetalle = lsDetalle & Space(63) & Right(Space(16) & Format(lnValorI, "##,###,##0.00"), 16) & Right(Space(16) & Format(lnValorDepre, "##,###,##0.00"), 16) & Right(Space(16) & Format(lnValorPorDepre, "##,###,##0.00"), 16) & oImpresora.gPrnSaltoLinea
    lsDetalle = lsDetalle & "  ---------------------------------------------------------------------------------------------------------------------------------------" & oImpresora.gPrnSaltoLinea ' oImpresora.gPrnSaltoPagina

        'ARLO 20160126 ***
        gsopecod = LogPistaBajaActivo
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsopecod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "3", "Se dio de Baja los Activos Fijos seleccionados "
        Set objPista = Nothing
        '**************

    'oPrevio.Show oAsiento.ImprimeAsientoContable(lsMovNroR, 60, 80, Caption) + lsDetalle, Caption, True
    oPrevio.Show lsTexto + lsDetalle, Caption, True
        
        
'    MsgBox "EL Activo Fijo " & Me.txtBS.Text & "-" & Me.txtSerie.Text & " ha sido dado de baja ", vbInformation, "Aviso"
    
    txtAreaAgeCod.Text = ""
    lblAreaAgeNombre.Caption = ""
    txtBS.Text = ""
    lblBienG.Caption = ""
    txtSerieCod.Text = ""
    lblSerieNombre.Caption = ""
    
    FeAdj.Clear
    FeAdj.FormaCabecera
    FeAdj.Rows = 2

    mskFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
    Me.txtComentario.Text = ""
    Me.chkEstadistico.value = 0
    
    CargarControles
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Command1_Click()
            For I = 1 To Me.FeAdj.Rows - 1
                FeAdj.TextMatrix(I, 2) = " "
            Next
End Sub

Private Sub Command2_Click()
            For I = 1 To Me.FeAdj.Rows - 1
                FeAdj.TextMatrix(I, 2) = "."
            Next

End Sub

'EJVG20130627 ***
Private Sub FeAdj_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(FeAdj.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
'END EJVG *******
Private Sub Form_Load()
    CentraForm Me
    CargarControles
    mskFecha.Text = Format(gdFecSis, "dd/mm/yyyy")
End Sub
Private Sub CargarControles()
    Dim oArea As New DActualizaDatosArea
    Dim oBien As New DBien
    
    txtAreaAgeCod.rs = oArea.GetAgenciasAreas()
    txtBS.rs = oBien.RecuperaCategoriasBienPaObjeto(True, "")
    txtSerieCod.rs = oBien.RecuperaSeriesPaObjeto("", "")
    Set oBien = Nothing
End Sub

Private Sub mskFecha_GotFocus()
    mskFecha.SelStart = 0
    mskFecha.SelLength = 50
End Sub

Private Sub mskFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtComentario.SetFocus
    End If
End Sub
'EJVG20130626 ***
Private Sub txtAreaAgeCod_EmiteDatos()
    Dim oBien As New DBien
    Dim lsAreaAgeCod As String
    
    Screen.MousePointer = 11
    lblAreaAgeNombre.Caption = ""
    txtBS.Text = ""
    lblBienG.Caption = ""
    If txtAreaAgeCod.Text <> "" Then
        lblAreaAgeNombre.Caption = txtAreaAgeCod.psDescripcion
        lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
        txtBS.rs = oBien.RecuperaCategoriasBienPaObjeto(False, lsAreaAgeCod)
    Else
        txtBS.rs = oBien.RecuperaCategoriasBienPaObjeto(True, "")
    End If
    txtBS_EmiteDatos
    Screen.MousePointer = 0
    Set oBien = Nothing
End Sub
'END EJVG *******
'Private Sub txtAgeO_EmiteDatos()
'    Me.lblAgeOG.Caption = txtAgeO.psDescripcion
'End Sub
Private Sub txtAreaAgeCod_LostFocus()
    If txtAreaAgeCod.Text = "" Then
        lblAreaAgeNombre.Caption = ""
    End If
End Sub
'Private Sub txtBS_EmiteDatos()
'    Dim oALmacen As DLogAlmacen
'    Set oALmacen = New DLogAlmacen
'
'    Dim rsSeries As ADODB.Recordset
'    Set rsSeries = New ADODB.Recordset
'    Dim lsAreaAgeCod As String
'    If txtBS.Text <> "" Then
'        If txtAreaAgeCod.Text <> "" Then
'            lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
'        End If
'        '*** PEAC 20120417
'
'        Me.lblBienG.Caption = txtBS.psDescripcion
''        Me.txtSerie.rs = oALmacen.GetAFBSSerie(txtBS.Text, Year(Me.mskFecha.Text))
'
'            'Set rsSeries = oALmacen.GetAFBSSerie(txtBS.Text, Year(Me.mskFecha.Text), 1, Me.chkTodos.value)
'            Set rsSeries = oALmacen.GetAFBSSerie(txtBS.Text, Year(Me.mskFecha.Text), 1, Me.chkTodos.value, lsAreaAgeCod) 'EJVG20130626
'
'            If rsSeries.EOF Then
'                MsgBox "Este Activo no tiene Series creadas.", vbInformation, "Mensaje"
'                FeAdj.Clear
'                FeAdj.FormaCabecera
'                FeAdj.Rows = 2
'                Exit Sub
'            End If
'
'            FeAdj.Clear
'            FeAdj.FormaCabecera
'            FeAdj.Rows = 2
'            FeAdj.rsFlex = rsSeries
'
'            MsgBox "Se cargaron los datos satisfactoriamente.", vbOKOnly + vbInformation, "Aviso"
'
'        '*** FIN PEAC
'
'    End If
'
'    Set oALmacen = Nothing
'End Sub

Private Sub txtBS_EmiteDatos()
    Dim oBien As New DBien
    Dim lsAreaAgeCod As String

    Screen.MousePointer = 11
    lblBienG.Caption = ""
    If txtBS.Text <> "" Then
       lblBienG.Caption = txtBS.psDescripcion
    End If
    If txtAreaAgeCod.Text <> "" Then
        lsAreaAgeCod = Left(txtAreaAgeCod.Text, 3) & IIf(Mid(txtAreaAgeCod.Text, 4, 2) = "", "01", Mid(txtAreaAgeCod.Text, 4, 2))
    End If
    txtSerieCod.Text = ""
    lblSerieNombre.Caption = ""
    txtSerieCod.rs = oBien.RecuperaSeriesPaObjeto(lsAreaAgeCod, txtBS.Text)
    txtSerieCod_EmiteDatos
    Screen.MousePointer = 0
    Set oBien = Nothing
End Sub
Private Sub txtBS_LostFocus()
    If txtBS.Text = "" Then
        lblBienG.Caption = ""
    End If
End Sub
Private Sub txtSerieCod_EmiteDatos()
    lblSerieNombre.Caption = ""
    FeAdj.Clear
    FeAdj.FormaCabecera
    FeAdj.Rows = 2
    If txtSerieCod.Text <> "" Then
        lblSerieNombre.Caption = txtSerieCod.psDescripcion
    End If
End Sub
Private Sub txtSerieCod_LostFocus()
    If txtSerieCod.Text = "" Then
        lblSerieNombre.Caption = ""
    End If
End Sub
Private Sub txtComentario_GotFocus()
    txtComentario.SelStart = 0
    txtComentario.SelLength = 300
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdGrabar.SetFocus
    End If
End Sub

'Private Sub txtSerie_EmiteDatos()
'    Dim oALmacen As DLogAlmacen
'    Set oALmacen = New DLogAlmacen
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'
'    If txtBS.Text <> "" And txtSerie.Text <> "" Then
'        Set rs = oALmacen.GetAFBSDetalle(txtBS.Text, txtSerie.Text)
'        '*** PEAC 20110223
'        If rs.EOF And rs.BOF Then
'            MsgBox "No existe datos para esta consulta.", vbOKOnly + vbInformation, "Atención"
'            txtSerie.Text = ""
'            Exit Sub
'        End If
'
'        lnMovNroIni = rs.Fields(2)
'        lnAnio = rs.Fields(3)
'        Me.txtAgeO.Text = rs.Fields(0) & rs.Fields(1)
'
'        '*** PEAC 20120417
'        Me.lblValorIni.Caption = CStr(rs.Fields(4))
'        Me.lblValorDepre.Caption = CStr(rs.Fields(5))
'        Me.lblPorDepre.Caption = CStr(rs.Fields(4) - rs.Fields(5))
'        '*** FIN PEAC
'
'        txtAgeO_EmiteDatos
'    End If
'
'    Set oALmacen = Nothing
'    Set rs = Nothing
'End Sub
