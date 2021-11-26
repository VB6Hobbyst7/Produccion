VERSION 5.00
Begin VB.Form frmCapTasaEspAprRech 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9090
   Icon            =   "frmCapTasaEspAprRech.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   9090
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fraDatosSolicitud 
      Caption         =   "Datos Solicitud"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2520
      Left            =   90
      TabIndex        =   12
      Top             =   2880
      Width           =   4380
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Nro Solicitud:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   450
         Width           =   1170
      End
      Begin VB.Label lblNumSol 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1350
         TabIndex        =   21
         Top             =   315
         Width           =   1095
      End
      Begin VB.Label lblComentario 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   735
         Left            =   90
         TabIndex        =   20
         Top             =   1620
         Width           =   4155
      End
      Begin VB.Label lblFecha 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1350
         TabIndex        =   19
         Top             =   1005
         Width           =   2580
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Fecha:"
         Height          =   195
         Left            =   105
         TabIndex        =   18
         Top             =   1110
         Width           =   495
      End
      Begin VB.Label lblUsu 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   330
         Left            =   1350
         TabIndex        =   17
         Top             =   660
         Width           =   1095
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Solictado por:"
         Height          =   195
         Left            =   90
         TabIndex        =   16
         Top             =   795
         Width           =   975
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Comentario:"
         Height          =   195
         Left            =   90
         TabIndex        =   15
         Top             =   1395
         Width           =   840
      End
   End
   Begin VB.CommandButton cmdActualizar 
      Caption         =   "Ac&tualizar"
      Height          =   375
      Left            =   135
      TabIndex        =   11
      Top             =   5505
      Width           =   1100
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   7875
      TabIndex        =   10
      Top             =   5505
      Width           =   1100
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "&Rechazar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4575
      TabIndex        =   7
      Top             =   5505
      Width           =   1100
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "&Aprobar"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3270
      TabIndex        =   6
      Top             =   5505
      Width           =   1100
   End
   Begin VB.Frame fraDatos 
      Caption         =   "Datos Aprobación/Rechazo"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2505
      Left            =   4500
      TabIndex        =   2
      Top             =   2880
      Width           =   4470
      Begin SICMACT.EditMoney txtTasa 
         Height          =   330
         Left            =   1485
         TabIndex        =   24
         Top             =   650
         Width           =   1350
         _ExtentX        =   2381
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Text            =   "0"
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkPermanente 
         Alignment       =   1  'Right Justify
         Caption         =   "Es Permanente"
         Height          =   210
         Left            =   135
         TabIndex        =   23
         Top             =   360
         Visible         =   0   'False
         Width           =   1530
      End
      Begin VB.TextBox txtComentario 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   90
         MaxLength       =   100
         MultiLine       =   -1  'True
         TabIndex        =   8
         ToolTipText     =   "(Escriba aquí su comentario)"
         Top             =   1605
         Width           =   4155
      End
      Begin SICMACT.EditMoney txtMonto 
         Height          =   330
         Left            =   1485
         TabIndex        =   3
         Top             =   1035
         Width           =   1860
         _ExtentX        =   3281
         _ExtentY        =   582
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
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
      Begin VB.Label lblMon 
         AutoSize        =   -1  'True
         Caption         =   "S/."
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3420
         TabIndex        =   14
         Top             =   1110
         Width           =   300
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "% TEA"
         Height          =   195
         Left            =   2910
         TabIndex        =   13
         Top             =   750
         Width           =   480
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Comentario:"
         Height          =   195
         Left            =   135
         TabIndex        =   9
         Top             =   1380
         Width           =   840
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tasa Aprobada:"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   750
         Width           =   1140
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Monto Apertura:"
         Height          =   195
         Left            =   135
         TabIndex        =   4
         Top             =   1110
         Width           =   1140
      End
   End
   Begin VB.Frame fra 
      Caption         =   "Solicitudes de Aprobación"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2670
      Left            =   90
      TabIndex        =   0
      Top             =   135
      Width           =   8895
      Begin SICMACT.FlexEdit grdSolicitud 
         Height          =   2265
         Left            =   90
         TabIndex        =   1
         Top             =   315
         Width           =   8670
         _ExtentX        =   15293
         _ExtentY        =   3995
         Cols0           =   16
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmCapTasaEspAprRech.frx":030A
         EncabezadosAnchos=   "350-3700-1700-900-900-400-1200-550-0-0-0-0-0-0-0-0"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-C-C-C-R-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "0-1-0-0-0-2-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapTasaEspAprRech"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by

'Modificado por RIRO el 201304111052
Private Sub ActualizaFramesDatos()
Dim nFila As Long
Dim sMovNro As String
nFila = grdSolicitud.row
If nFila > 0 Then
    
    sMovNro = grdSolicitud.TextMatrix(nFila, 11)
    lblUsu = Right(sMovNro, 4)
    lblFecha = Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
    lblComentario = grdSolicitud.TextMatrix(nFila, 8)
    txtTasa.Text = Format(grdSolicitud.TextMatrix(nFila, 4), "#0.0000")
    txtMonto.Text = Format(grdSolicitud.TextMatrix(nFila, 6), "#0.00")
    lblMon = grdSolicitud.TextMatrix(nFila, 5)
    txtTasa.BackColor = grdSolicitud.CellBackColor
    txtMonto.BackColor = grdSolicitud.CellBackColor
    cmdAprobar.Enabled = True
    cmdRechazar.Enabled = True
    fraDatos.Enabled = True
    lblNumSol.Caption = grdSolicitud.TextMatrix(nFila, 13)
    
'    sMovNro = grdSolicitud.TextMatrix(nFila, 10)
'    lblUsu = Right(sMovNro, 4)
'    lblFecha = Mid(sMovNro, 7, 2) & "/" & Mid(sMovNro, 5, 2) & "/" & Left(sMovNro, 4) & " " & Mid(sMovNro, 9, 2) & ":" & Mid(sMovNro, 11, 2) & ":" & Mid(sMovNro, 13, 2)
'    lblComentario = grdSolicitud.TextMatrix(nFila, 6)
'    txtTasa.Text = Format(grdSolicitud.TextMatrix(nFila, 3), "#0.0000")
'    txtMonto.Text = Format(grdSolicitud.TextMatrix(nFila, 5), "#0.00")
'    lblMon = grdSolicitud.TextMatrix(nFila, 4)
'    txtTasa.BackColor = grdSolicitud.CellBackColor
'    txtMonto.BackColor = grdSolicitud.CellBackColor
'    cmdAprobar.Enabled = True
'    cmdRechazar.Enabled = True
'    fraDatos.Enabled = True
'    lblNumSol.Caption = grdSolicitud.TextMatrix(nFila, 12)

Else
    lblUsu = ""
    lblFecha = ""
    lblComentario = ""
    txtComentario.Text = ""
    lblNumSol.Caption = ""
    
    txtTasa.Text = "0.0000"
    txtMonto.Text = "0.00"
    lblMon = ""
    txtTasa.BackColor = &HFFFFFF
    txtMonto.BackColor = &HFFFFFF
    cmdAprobar.Enabled = False
    cmdRechazar.Enabled = False
    fraDatos.Enabled = False
    
End If
End Sub

'Modificado por RIRO, se efectuo cambio por modificacion en el diseño del grid
Private Sub DarFormatoGrid()
Dim i As Long, J As Long
Dim nFila As Long, nCol As Long
For i = 1 To grdSolicitud.Rows - 1
    'grdSolicitud.TextMatrix(i, 3) = Format$(ConvierteTNAaTEA(CDbl(grdSolicitud.TextMatrix(i, 3))), "#,##0.00")
    'grdSolicitud.TextMatrix(i, 5) = Format$(CDbl(grdSolicitud.TextMatrix(i, 5)), "#,##0.00")
    
    If grdSolicitud.TextMatrix(i, 3) <> "" Then
    
        grdSolicitud.TextMatrix(i, 3) = Format$(CDbl(grdSolicitud.TextMatrix(i, 3)), "#,##0.00")
    
    End If
    
    If grdSolicitud.TextMatrix(i, 4) <> "" Then
    
        grdSolicitud.TextMatrix(i, 4) = Format$(CDbl(grdSolicitud.TextMatrix(i, 4)), "#,##0.00")
    
    End If
    
    If grdSolicitud.TextMatrix(i, 6) <> "" Then
    
        grdSolicitud.TextMatrix(i, 6) = Format$(CDbl(grdSolicitud.TextMatrix(i, 6)), "#,##0.00")
    
    End If
    
    If grdSolicitud.TextMatrix(i, 8) = "2" Then
        nFila = grdSolicitud.row
        nCol = grdSolicitud.col
        grdSolicitud.row = i
        For J = 1 To grdSolicitud.Cols - 1
            grdSolicitud.col = J
            grdSolicitud.CellBackColor = &HC0FFC0
        Next J
        grdSolicitud.row = nFila
        grdSolicitud.col = nCol
    End If
Next i
'Agregado Por RIRO
ValidaTasas
End Sub



Private Sub GetSolicitudesAprobacion()
Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
Dim rsTasa As New ADODB.Recordset

Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
'Set rsTasa = oServ.GetDatosCapTasaEspecial(0)
Set rsTasa = oserv.GetDatosTasaEspecialSol()
Set oserv = Nothing
If rsTasa.EOF And rsTasa.BOF Then
    MsgBox "No existen solicitudes de Aprobación/Rechazo de Tasas Especiales", vbInformation, "Aviso"
End If
Set grdSolicitud.Recordset = rsTasa

Set oserv = Nothing
End Sub

Private Sub cmdActualizar_Click()
GetSolicitudesAprobacion
DarFormatoGrid
ActualizaFramesDatos
End Sub

'Agregafo por RIRO el 201304111115
'Se usuario no esta autorizado, devuelve el valor 999, caso contrario devuelve la tea adicional
Private Function validarPermisosAprobacion(Optional Aprobar As Boolean = True) As String

    Dim oCaptaServicios As COMDCaptaServicios.DCOMCaptaServicios
    Dim rs As ADODB.Recordset
    Dim nTeaAdicional, nTemporal, nPlazo, nPlazoSolicitado, nTeaTarifada, nTeaSolicitada As Double
    Dim nSubPr, nContar, nProducto, nSubProducto, i As Integer
    Dim sMensaje As String
    Dim sGrupos() As String
    
    i = 0
    sMensaje = ""
    nContar = 0
    nTeaAdicional = 0
    nTemporal = 0
    nSubPr = 0
    nPlazo = 100000
    nTeaSolicitada = IIf(Aprobar, CDbl(txtTasa.Text), CDbl(grdSolicitud.TextMatrix(grdSolicitud.row, 4)))
    nPlazoSolicitado = CDbl(grdSolicitud.TextMatrix(grdSolicitud.row, 7))
    nTeaTarifada = CDbl(grdSolicitud.TextMatrix(grdSolicitud.row, 3))
    nProducto = CInt(grdSolicitud.TextMatrix(grdSolicitud.row, 12))
    nSubProducto = CInt(Mid(grdSolicitud.TextMatrix(grdSolicitud.row, 14), 4, 1))
    
    sGrupos = Split(gsGruposUser, ",")
    Set oCaptaServicios = New COMDCaptaServicios.DCOMCaptaServicios
        
    For i = LBound(sGrupos) To UBound(sGrupos)
            
        Set rs = oCaptaServicios.ObtenerPermisoTea(sGrupos(i), gsCodAge)
        
        If Not rs.EOF Then
            
                nTemporal = CDbl(Format$(ConvierteTNAaTEA(rs("nTeaadicional")), "#,##0.00"))
                                    
            If nTemporal > nTeaAdicional Then
                 nTeaAdicional = nTemporal
            End If
                        
                nTemporal = CDbl(rs("nDias"))
                        
            If nTemporal < nPlazo Then
                 nPlazo = nTemporal
            End If
                                       
                nTemporal = CDbl(rs("nSubProductos"))
                        
            If nTemporal = 1 Then
                 nSubPr = 1
            End If
                        
            nContar = nContar + 1
          
        End If
            
    Next
    
    If nContar > 0 Then
    
        If nProducto = 233 Then
        
            If nPlazoSolicitado < nPlazo Then
                  sMensaje = "* El plazo solicitado es menor al plazo autorizado por el nivel de aprobacion" & vbNewLine
            End If
            
        Else
        
            If nSubPr <> 1 Then
                
                sMensaje = "* No tiene permiso para aprobar solicitudes de tasa especial para los productos: Ahorro y CTS" & vbNewLine
        
            End If
        
        End If
        
        If CDbl(nTeaAdicional) < (nTeaSolicitada - nTeaTarifada) Then
            sMensaje = sMensaje & "* La tasa adicional solicitada es superior a la autorizada" & vbNewLine
        End If
        
        If Trim(sMensaje) <> "" Then
        
            sMensaje = "Se presentaron las siguientes observaciones: " & vbNewLine & vbNewLine & sMensaje
        
        End If
            
    Else
        sMensaje = "No tiene permiso para aprobar o rechazar la solicitud de tasa especial"
    
    End If
        
    Set oCaptaServicios = Nothing
    Set rs = Nothing
    
    validarPermisosAprobacion = sMensaje
         
End Function

'Modificado por RIRO el 201304111121
Private Sub cmdAprobar_Click()

Dim nTasa As Double, nMonto As Double
Dim nProd As COMDConstantes.Producto, nMon As COMDConstantes.Moneda
Dim sComent, sPersona, sMensaje As String ' Agregado por RIRO 20130411
Dim nPlazo As Double, bPermanente As Boolean
Dim nFila As Long
Dim nTeaAdicional, nTasaTarif, nTasaSolicitada As Double '20130411RIRO

If Trim(grdSolicitud.TextMatrix(grdSolicitud.row, 2)) = "" Then
    MsgBox "Esta solicitud no puede ser aprobada porque se efectuaron cambios " & _
    " internos en el sistema, se agradece generar nuevamente la solicitud", vbExclamation, "Aviso"
    Exit Sub
End If

sMensaje = validarPermisosAprobacion '20130411RIRO

If Trim(sMensaje) <> "" Then

    MsgBox sMensaje, vbInformation, "Aviso"
    Exit Sub

End If

nFila = grdSolicitud.row
nTasa = CDbl(txtTasa.Text) '20130411RIRO
nTasaTarif = CDbl(grdSolicitud.TextMatrix(grdSolicitud.row, 3)) '20130411RIRO
nTasaSolicitada = CDbl(grdSolicitud.TextMatrix(grdSolicitud.row, 4)) '20130411RIRO
nPlazo = CDbl(grdSolicitud.TextMatrix(nFila, 7)) ' Modificado por RIRO 20130411

If nTasa = 0 Or nTasa >= 100 Then
    MsgBox "Tasa no Válida", vbInformation, "Error"
    txtTasa.SetFocus
    Exit Sub
End If

nMonto = txtMonto.value

If nMonto = 0 Then
    MsgBox "Monto de Apertura no Válida", vbInformation, "Error"
    txtMonto.SetFocus
    Exit Sub
End If

sComent = Trim(txtComentario.Text)

If sComent = "" Then
    MsgBox "Comentario no Válido", vbInformation, "Error"
    txtComentario.SetFocus
    Exit Sub
End If

If MsgBox("¿Desea Aprobar la Solicitud de Tasa Especial?", vbQuestion + vbYesNo) = vbYes Then
    
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String
    Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
    Dim nNumSolicitud As Long
    
    Dim sSProducto As String ' 20130411RIRO
    
    Set oCont = New COMNContabilidad.NCOMContFunciones
    sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Set oCont = Nothing
    
    sPersona = grdSolicitud.TextMatrix(nFila, 9) ' Modificado por RIRO 20130411
    nMon = CInt(grdSolicitud.TextMatrix(nFila, 10)) ' Modificado por RIRO 20130411
    nNumSolicitud = CLng(grdSolicitud.TextMatrix(nFila, 13)) ' Modificado por RIRO 20130411
    nProd = CDbl(grdSolicitud.TextMatrix(nFila, 12)) ' Modificado por RIRO 20130411
    
    If nProd = gCapPlazoFijo Then
        nPlazo = CDbl(grdSolicitud.TextMatrix(nFila, 7)) ' Modificado por RIRO 20130411
    Else
        nPlazo = 0
    End If
    
    sSProducto = grdSolicitud.TextMatrix(nFila, 14) ' 20130411RIRO
    
    nTasa = Format$(ConvierteTEAaTNA(nTasa), "#0.0000")
    nTasaTarif = Format$(ConvierteTEAaTNA(nTasaTarif), "#0.000000") ' 20130411RIRO
    nTasaSolicitada = Format$(ConvierteTEAaTNA(nTasaSolicitada), "#0.000000") ' 20130411RIRO
    
    bPermanente = chkPermanente.value
    
    Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
    oserv.AgregaCapTasaEspecial nNumSolicitud, sPersona, nProd, nMon, 1, sMovNro, nTasa, sComent, nMonto, , nPlazo, , bPermanente, sSProducto, nTasaTarif, nTasaSolicitada
    'By Capi 21012009
     objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Aprobacion", str(nNumSolicitud), gNumeroSolicitud
    'End by

    Set oserv = Nothing
    cmdActualizar_Click
    
    txtComentario.Text = "" ' 20130411RIRO
            
    If txtComentario.Enabled Then
            
        txtComentario.SetFocus
                
    End If
    
End If

End Sub

Private Sub cmdRechazar_Click()
    
    Dim nTasa As Double, nMonto As Double
    Dim nProd As COMDConstantes.Producto, nMon As COMDConstantes.Moneda
    Dim sComent, sPersona, sMensaje As String
    Dim nPlazo As Integer
    Dim nFila As Long
    Dim nTeaAdicional, nTasaTarif, nTasaSolicitada As Double ' 20130411RIRO
    
    sMensaje = validarPermisosAprobacion(False)
 
    If Trim(sMensaje) <> "" Then

        MsgBox sMensaje, vbInformation, "Aviso"
        Exit Sub

    End If
 
    nTasa = CDbl(grdSolicitud.TextMatrix(grdSolicitud.row, 4))  ' 20130411RIRO
    nTasaTarif = CDbl(grdSolicitud.TextMatrix(grdSolicitud.row, 3)) ' 20130411RIRO
    nTasaSolicitada = CDbl(grdSolicitud.TextMatrix(grdSolicitud.row, 4)) '20130411RIRO
    
    If nTasa = 0 Or nTasa >= 100 Then
        MsgBox "Tasa no Válida", vbInformation, "Error"
        txtTasa.SetFocus
        Exit Sub
    End If
    nMonto = txtMonto.value
    If nMonto = 0 Then
        MsgBox "Monto de Apertura no Válida", vbInformation, "Error"
        txtMonto.SetFocus
        Exit Sub
    End If
    sComent = Trim(txtComentario.Text)
    If sComent = "" Then
        MsgBox "Comentario no Válido", vbInformation, "Error"
        txtComentario.SetFocus
        Exit Sub
    End If
    nFila = grdSolicitud.row
    
    If MsgBox("¿Desea RECHAZAR la Solicitud de Tasa Especial?", vbQuestion + vbYesNo) = vbYes Then
        Dim oCont As COMNContabilidad.NCOMContFunciones
        Dim sMovNro As String
        Dim oserv As COMDCaptaServicios.DCOMCaptaServicios
        Dim nNumSolicitud As Long
        Dim sSProducto As String ' 20130411RIRO
        
        Set oCont = New COMNContabilidad.NCOMContFunciones
        sMovNro = oCont.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set oCont = Nothing
        
        sPersona = grdSolicitud.TextMatrix(nFila, 9)
        nMon = CInt(grdSolicitud.TextMatrix(nFila, 10))
        nNumSolicitud = CLng(grdSolicitud.TextMatrix(nFila, 13))
        nProd = CInt(grdSolicitud.TextMatrix(nFila, 12))
        If nProd = gCapPlazoFijo Then
            nPlazo = CInt(grdSolicitud.TextMatrix(nFila, 7))
        Else
            nPlazo = 0
        End If
        
        sSProducto = grdSolicitud.TextMatrix(nFila, 14) '201304RIRO
        
        nTasa = Format$(ConvierteTEAaTNA(nTasa), "#0.000000")
        nTasaTarif = Format$(ConvierteTEAaTNA(nTasaTarif), "#0.000000") ' 20130411RIRO
        nTasaSolicitada = Format$(ConvierteTEAaTNA(nTasaSolicitada), "#0.000000") 'MODIFICADO POR "RIRO" EL 22/11/2012
        
        Set oserv = New COMDCaptaServicios.DCOMCaptaServicios
        oserv.AgregaCapTasaEspecial nNumSolicitud, sPersona, nProd, nMon, 4, sMovNro, nTasa, sComent, nMonto, , nPlazo, , , sSProducto, nTasaTarif, nTasaSolicitada
        
        'By Capi 21012009
         objPista.InsertarPista gsOpeCod, sMovNro, gsCodPersUser, GetMaquinaUsuario, gInsertar, "Rechazo", str(nNumSolicitud), gNumeroSolicitud
        'End by
        Set oserv = Nothing
        cmdActualizar_Click
        
        ' 20130411RIRO ************
        txtComentario.Text = ""
            
        If txtComentario.Enabled Then
                    
           txtComentario.SetFocus
                
        End If
        ' END RIRO ****************
        
    End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Tasa Especial - Aprobación/Rechazo"
Me.Icon = LoadPicture(App.path & gsRutaIcono)
GetSolicitudesAprobacion
DarFormatoGrid
ActualizaFramesDatos
'By Capi 20012009
Set objPista = New COMManejador.Pista
gsOpeCod = gCapAprobRechTasasPreferen
'End By
End Sub

Private Sub grdSolicitud_Click()
    If grdSolicitud.Rows < 2 Then
        Exit Sub
    End If
End Sub

Private Sub grdSolicitud_DblClick()
    If grdSolicitud.Rows < 2 Then
        Exit Sub
    End If
End Sub

Private Sub grdSolicitud_GotFocus()
ActualizaFramesDatos
End Sub

Private Sub grdSolicitud_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtTasa.SetFocus
End If
End Sub

Private Sub grdSolicitud_RowColChange()
ActualizaFramesDatos
End Sub

Private Sub txtMonto_Change()
 
 If txtMonto.Text = "." Then
        MsgBox "Debe ingresar valores numéricos", vbInformation, "Aviso"
        txtMonto.Text = Format(grdSolicitud.TextMatrix(grdSolicitud.row, 6), "#0.00")
 End If
 
End Sub

Private Sub txtTasa_Change()
        
    If Not IsNumeric(txtTasa.Text) Then
        MsgBox "Debe ingresar valores numéricos", vbInformation, "Aviso"
        txtTasa.Text = Format(grdSolicitud.TextMatrix(grdSolicitud.row, 4), "#0.0000")
        txtTasa.MarcaTexto
    Else
        If txtTasa.Text > 100 Then
            txtTasa.Text = "100.00"
        End If
            
    End If
    
End Sub

Private Sub txtTasa_GotFocus()
    txtTasa.MarcaTexto
End Sub

Private Sub txtTasa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        txtMonto.SetFocus
    End If
End Sub

Private Sub txtMonto_GotFocus()
txtMonto.MarcaTexto
End Sub

Private Sub txtMonto_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    txtComentario.SetFocus
End If
End Sub

Private Sub txtComentario_GotFocus()
txtComentario.SelStart = 0
txtComentario.SelLength = Len(txtComentario.Text)
End Sub

Private Sub txtComentario_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    cmdAprobar.SetFocus
End If
End Sub

Private Sub ValidaTasas()

    Dim dConec As COMConecta.DCOMConecta
    Dim nSolicitud  As Integer
    Dim RSOtro As ADODB.Recordset
    Dim i As Double

    Set dConec = New COMConecta.DCOMConecta

    dConec.AbreConexion

    For i = 1 To grdSolicitud.Rows - 1
        If Trim(grdSolicitud.TextMatrix(i, 4)) = "" Or _
            Trim(grdSolicitud.TextMatrix(i, 3)) = "" Then
                nSolicitud = grdSolicitud.TextMatrix(i, 13)
                Set RSOtro = dConec.Ejecutar("select top 1 c.nNumSolicitud, c.cMovNro, c.cPersCod, c.nProducto, c.nMoneda, c.cCtaCod,c.nMonto, c.nPlazo, c.ntasa, c.nestado, c.ccomentario, c.NEXTORNO, c.bPermanente, c.ctipo,c.cSubProducto, c.nTasaTarif,c.nTasaSolicitada, p.nPersPersoneria from CapTasaEspecial c inner join Persona p on c.cPersCod = p.cPersCod where nNumSolicitud = " & nSolicitud & " order by cMovNro desc")
                If Not RSOtro.EOF Then
                    grdSolicitud.TextMatrix(i, 3) = "0.00"
                    grdSolicitud.TextMatrix(i, 4) = Format(ConvierteTNAaTEA(RSOtro!nTasa), "#,##0.00")
                    grdSolicitud.TextMatrix(i, 14) = "20300"
                End If
        End If
    Next
 
    dConec.CierraConexion
    Set dConec = Nothing
    Set RSOtro = Nothing
   
End Sub

