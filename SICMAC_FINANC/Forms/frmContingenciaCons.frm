VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmContingenciaCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Contingencias: Consulta"
   ClientHeight    =   4545
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15360
   Icon            =   "frmContingenciaCons.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4545
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtBusqueda 
      Height          =   375
      Left            =   1200
      TabIndex        =   7
      Top             =   4080
      Width           =   3735
   End
   Begin VB.CommandButton cmdVerDetMonto 
      Caption         =   "Ver Detalle Monto"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   11760
      TabIndex        =   6
      Top             =   4080
      Width           =   2250
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "Imprimir"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   9480
      TabIndex        =   3
      Top             =   4080
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdVerIT 
      Caption         =   "Ver IT"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10680
      TabIndex        =   2
      Top             =   4080
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   14160
      TabIndex        =   1
      Top             =   4080
      Width           =   1170
   End
   Begin TabDlg.SSTab SSTabConting 
      Height          =   3780
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   15195
      _ExtentX        =   26802
      _ExtentY        =   6668
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Listado de Registro de Contingencias"
      TabPicture(0)   =   "frmContingenciaCons.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "feContingentes"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin Sicmact.FlexEdit feContingentes 
         Height          =   3315
         Left            =   0
         TabIndex        =   5
         Top             =   480
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   5847
         Cols0           =   14
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   $"frmContingenciaCons.frx":0326
         EncabezadosAnchos=   "385-1350-1200-1800-900-1200-1820-2200-1000-1200-900-1200-1200-1200"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-9-X-11-12-X"
         TextStyleFixed  =   3
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C-R-C-C-C-R-C-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-2-0-0-2-2-2-0-2-0"
         CantEntero      =   15
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbFormatoCol    =   -1  'True
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   390
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "Extornar"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   10680
      TabIndex        =   4
      Top             =   4080
      Visible         =   0   'False
      Width           =   1050
   End
   Begin VB.Label lblBuscar 
      Caption         =   "Buscar :"
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
      Left            =   240
      TabIndex        =   8
      Top             =   4155
      Width           =   855
   End
End
Attribute VB_Name = "frmContingenciaCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmContingenciaCons
'** Descripción : Consulta de Contingencias por Areas creado segun RFC056-2012
'** Creación : JUEZ, 20120618 09:00:00 AM
'********************************************************************

Option Explicit
Dim rs As ADODB.Recordset
Dim oConting As DContingencia
Dim oGen As DGeneral
Dim sNumRegistro As String

Private Sub cmdCerrar_Click()
    Unload Me
End Sub
'Comentado por TORE28032018
'Private Sub cmdExtornar_Click()
'    sNumRegistro = DBGrdConting.Columns(0)
'    If sNumRegistro <> "" Then
'        Set oConting = New DContingencia
'        Set rs = oConting.BuscaContigenciasxArea(gsCodArea, sNumRegistro)
'        If rs!nEstado <> 1 Then
'            MsgBox "No puede extornar la contingencia seleccionada, presenta informe técnico", vbExclamation, "Aviso!"
'            Exit Sub
'        End If
'
'        If MsgBox("Está seguro de extornar la contingencia ", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
'
'        Call oConting.ExtornaContingencia(sNumRegistro)
'        MsgBox "Se ha extornado la Contingencia", vbInformation, "Aviso"
'        Call Extorno
'    End If
'End Sub

Private Sub cmdImprimir_Click()
    Dim lsCadImp As String
    'sNumRegistro = DBGrdConting.Columns(0)
    'If sNumRegistro <> "" Then
    Screen.MousePointer = 11
    lsCadImp = ImprimirRegistroContingencia
    If Len(Trim(lsCadImp)) > 0 Then
        EnviaPrevio lsCadImp, "Registro de Contingencias", gnLinPage, False
    Else
        MsgBox "No Existen Datos para el reporte", vbInformation, "Aviso"
    End If
    Screen.MousePointer = 0
    'End If
End Sub

Public Function Consulta(Optional PrimerIngreso As Integer = 0)
    Screen.MousePointer = 11
    Set oConting = New DContingencia
    Set rs = oConting.BuscaContigenciasxArea(gsCodArea, , "1,3")
    'Set DBGrdConting.DataSource = rs
    'DBGrdConting.Refresh
    
    'CROB20170731
    feContingentes.Clear
    feContingentes.FormaCabecera
    feContingentes.Rows = 2
    If Not rs.EOF And Not rs.BOF Then
        Set feContingentes.Recordset = rs
        feContingentes.Row = rs.RecordCount - 1
    End If 'CROB20170731
    
    Screen.MousePointer = 0
    If rs.RecordCount = 0 Then
      MsgBox "No se Encontraron Datos Registrados", vbInformation, "Aviso"
      cmdImprimir.Visible = False
      cmdVerIT.Visible = False
    Else
      'cmdImprimir.Visible = True
      'cmdVerIT.Visible = True
    End If
    Me.Caption = "Contingencias: Consulta"
    If PrimerIngreso = 1 Then
        Me.Show 1
    End If
End Function
'Cometado por TORE28032018
'Public Function Extorno(Optional PrimerIngreso As Integer = 0)
'    Screen.MousePointer = 11
'    Set oConting = New DContingencia
'    Set rs = oConting.BuscaContigenciasxArea(gsCodArea, , "1")
'    Set DBGrdConting.DataSource = rs
'    DBGrdConting.Refresh
'    Screen.MousePointer = 0
'    If rs.RecordCount = 0 Then
'      MsgBox "No se Encontraron Datos Registrados", vbInformation, "Aviso"
'      cmdExtornar.Visible = False
'    Else
'      cmdExtornar.Visible = True
'    End If
'    Me.Caption = "Contingencias: Extornar"
'    If PrimerIngreso = 1 Then
'        Me.Show 1
'    End If
'End Function

Public Function ImprimirRegistroContingencia() As String
    Dim lsCad As String
    Dim cAreaNombre As String
    Dim oAreas As DActualizaDatosArea
    Set oConting = New DContingencia
    Set rs = oConting.BuscaContigenciasxArea(gsCodArea, , "1,2")
    Set oAreas = New DActualizaDatosArea
    cAreaNombre = oAreas.GetNombreAreas(gsCodArea)
            
    Dim psTitulo As String
    Dim pnAnchoLinea As Integer
    Dim RDatosUser As ADODB.Recordset
    Dim lnNegritaON As String
    Dim lnNegritaOFF As String
    
    psTitulo = "CONTINGENCIAS REGISTRADAS"
    pnAnchoLinea = 125

    CON = PrnSet("C+")
    COFF = PrnSet("C-")
    BON = PrnSet("B+")
    BOFF = PrnSet("B-")
    
    lsCad = CON & BON & Chr$(10) & Chr$(10) & Centra(" " & psTitulo & " ", pnAnchoLinea) & Chr$(10)
    lsCad = lsCad & Centra(" " & gdFecSis & " ", pnAnchoLinea) & Chr$(10) & Chr$(10)
    lsCad = lsCad & "Area: " & cAreaNombre
    lsCad = lsCad & space(pnAnchoLinea - 45)
    lsCad = lsCad & "Usuario: " & FillText(gsCodUser, 5, " ") & BOFF & Chr$(10)
    lsCad = lsCad & String(pnAnchoLinea, "-") & Chr$(10)
    lsCad = lsCad & "Fecha Reg.  Tipo Contingencia  Moneda        Monto  Origen                              Tipo Evento de Perdida" & Chr$(10)
    lsCad = lsCad & String(pnAnchoLinea, "-") & Chr$(10)
    Do While Not rs.EOF
        lsCad = lsCad & Left(rs!dFechaReg & String(10, " "), 10) & space(2)
        lsCad = lsCad & Left(rs!cTpoConting & String(18, " "), 18) & space(3)
        lsCad = lsCad & Left(rs!cMoneda & String(3, " "), 3)
        lsCad = lsCad & Right(String(14, " ") & Format(rs!nMonto, "#,###,##0.00"), 14) & space(2)
        lsCad = lsCad & Left(rs!cOrigen & String(30, " "), 34) & space(2)
        lsCad = lsCad & Left(rs!cTpoEvPerdida & String(22, " "), 37)
        lsCad = lsCad & Chr$(10)
        rs.MoveNext
    Loop
    lsCad = lsCad & String(pnAnchoLinea, "-") & COFF & Chr$(10)
    
    ImprimirRegistroContingencia = lsCad
End Function
'Cometado por TORE28032018
'Private Sub cmdVerIT_Click()
'    sNumRegistro = DBGrdConting.Columns(0)
'    If sNumRegistro <> "" Then
'        Set oConting = New DContingencia
'        Set rs = oConting.BuscaContigenciaSeleccionada(sNumRegistro)
'        If rs!nEstado = 1 Then
'            MsgBox "La contigencia seleccionada no tiene ningun informe técnico", vbInformation, "Aviso"
'            Exit Sub
'        End If
'        frmContingInformeTecCons.Consulta sNumRegistro, 1
'    End If
'End Sub

'CROB20170731 -> TORE(Actualizacion)
Private Sub feContingentes_OnCellChange(pnRow As Long, pnCol As Long)
    Dim sNumRegistro As String
    Dim nProvision As Double
    Dim cCtaContable As String
    Dim nMRecuperado As Double
    Dim nResul As Integer
    Set oConting = New DContingencia
    
    If feContingentes.col = 9 Then
        sNumRegistro = feContingentes.TextMatrix(feContingentes.Row, 1)
        nProvision = feContingentes.TextMatrix(feContingentes.Row, 9)
        If nProvision < 0 Then
            
            Set oConting = New DContingencia
            Set rs = oConting.BuscaContigenciasxArea(gsCodArea, , "1,3")
            
            feContingentes.Clear
            feContingentes.FormaCabecera
            feContingentes.Rows = 2
            If Not rs.EOF And Not rs.BOF Then
                Set feContingentes.Recordset = rs
                feContingentes.Row = rs.RecordCount - 1
            End If
            txtBusqueda.Text = ""
        
            MsgBox "No se puede registrar montos negativos", vbInformation, "Aviso"
        Else
            Set rs = oConting.ActualizarProvisionContingencia(sNumRegistro, nProvision, gsCodUser)
        nResul = rs!nResultado
        
        Set oConting = New DContingencia
        Set rs = oConting.BuscaContigenciasxArea(gsCodArea, , "1,3")
        
        feContingentes.Clear
        feContingentes.FormaCabecera
        feContingentes.Rows = 2
        If Not rs.EOF And Not rs.BOF Then
            Set feContingentes.Recordset = rs
            feContingentes.Row = rs.RecordCount - 1
        End If
        txtBusqueda.Text = ""
        
        If nResul = 1 Then
            MsgBox "Monto de la provisión actualizado", vbInformation, "Aviso"
        Else
            MsgBox "No se puede actualizar la provisión debido a que el evento se encuentra finalizado", vbInformation, "Aviso"
        End If
        End If
    End If
    
    'TORE07032018
    If feContingentes.col = 11 Then
        sNumRegistro = feContingentes.TextMatrix(feContingentes.Row, 1)
        cCtaContable = feContingentes.TextMatrix(feContingentes.Row, 11)
        Set rs = oConting.ActualizarCtaContableContingencia(cCtaContable, sNumRegistro)
        nResul = rs!nResultado
        
        Set oConting = New DContingencia
        Set rs = oConting.BuscaContigenciasxArea(gsCodArea, , "1,3")
        
        feContingentes.Clear
        feContingentes.FormaCabecera
        feContingentes.Rows = 2
        If Not rs.EOF And Not rs.BOF Then
            Set feContingentes.Recordset = rs
            feContingentes.Row = rs.RecordCount - 1
        End If
        txtBusqueda.Text = ""
        
        If nResul = 1 Then
            MsgBox "La cuenta contable fue actualizado", vbInformation, "Aviso"
        Else
            MsgBox "La cuenta contable no es la permitida", vbInformation, "Aviso"
        End If
    End If
    
    If feContingentes.col = 12 Then
        sNumRegistro = feContingentes.TextMatrix(feContingentes.Row, 1)
        nMRecuperado = feContingentes.TextMatrix(feContingentes.Row, 12)
        
        If nMRecuperado < 0 Then
            
            Set oConting = New DContingencia
            Set rs = oConting.BuscaContigenciasxArea(gsCodArea, , "1,3")
            
            feContingentes.Clear
            feContingentes.FormaCabecera
            feContingentes.Rows = 2
            If Not rs.EOF And Not rs.BOF Then
                Set feContingentes.Recordset = rs
                feContingentes.Row = rs.RecordCount - 1
            End If
            txtBusqueda.Text = ""
            
            MsgBox "No se puede registrar montos negativos", vbInformation, "Aviso"
        Else
            Set rs = oConting.ActualizarMontoRecuperadoContingencia(sNumRegistro, nMRecuperado, gsCodUser)
            nResul = rs!nResultado
            
            Set oConting = New DContingencia
            Set rs = oConting.BuscaContigenciasxArea(gsCodArea, , "1,3")
            
            feContingentes.Clear
            feContingentes.FormaCabecera
            feContingentes.Rows = 2
            If Not rs.EOF And Not rs.BOF Then
                Set feContingentes.Recordset = rs
                feContingentes.Row = rs.RecordCount - 1
            End If
            txtBusqueda.Text = ""
            
            If nResul = 1 Then
                MsgBox "Monto recuperado actualizado", vbInformation, "Aviso"
            Else
                MsgBox "Error en la actualización del monto recuperado", vbInformation, "Aviso"
            End If
        End If
        
        
        
    End If
    
    'END TORE
    
    Set oConting = Nothing
End Sub

Private Sub cmdVerDetMonto_Click()
    sNumRegistro = feContingentes.TextMatrix(feContingentes.Row, 1)
    frmContingenciaDetMontos.Consultar (sNumRegistro)
End Sub


Private Sub txtBusqueda_Change()
    Dim i As Integer
        
    Dim rsFiltro As ADODB.Recordset
    Set rsFiltro = rs.Clone
    
    If txtBusqueda.Text = "'" Then txtBusqueda.Text = ""
    
    If Trim(txtBusqueda.Text) <> "" Then
        rsFiltro.Filter = " cNumRegistro LIKE '*" + Trim(txtBusqueda.Text) + "*'"
    End If
    
    feContingentes.Clear
    feContingentes.FormaCabecera
    Call LimpiaFlex(feContingentes)
    For i = 1 To rsFiltro.RecordCount
        feContingentes.AdicionaFila

            feContingentes.TextMatrix(i, 1) = rsFiltro!cNumRegistro
            feContingentes.TextMatrix(i, 2) = rsFiltro!dFechaReg
            feContingentes.TextMatrix(i, 3) = rsFiltro!cTpoConting
            feContingentes.TextMatrix(i, 4) = rsFiltro!cMoneda
            feContingentes.TextMatrix(i, 5) = FormatNumber(rsFiltro!nMonto, 2)
            feContingentes.TextMatrix(i, 6) = rsFiltro!cOrigen
            feContingentes.TextMatrix(i, 7) = rsFiltro!cTpoEvPerdida
            feContingentes.TextMatrix(i, 8) = rsFiltro!cReportado
            feContingentes.TextMatrix(i, 9) = FormatNumber(rsFiltro!nProvision, 2)
            feContingentes.TextMatrix(i, 11) = FormatNumber(rsFiltro!cCtaContCod, 2)
            feContingentes.TextMatrix(i, 12) = FormatNumber(rsFiltro!nMontoRecup, 2)
            feContingentes.TextMatrix(i, 13) = rsFiltro!cEstado
        rsFiltro.MoveNext
    Next i
End Sub
'CROB20170731

