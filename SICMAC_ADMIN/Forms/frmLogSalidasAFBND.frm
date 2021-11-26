VERSION 5.00
Begin VB.Form frmLogSalidasAFBND 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Salidas de Actibo Fijo"
   ClientHeight    =   6705
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   Icon            =   "frmLogSalidasAFBND.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6705
   ScaleWidth      =   9690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdValida 
      Caption         =   "&Valida"
      Height          =   360
      Left            =   4995
      TabIndex        =   7
      Top             =   6300
      Width           =   1005
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   8640
      TabIndex        =   6
      Top             =   6300
      Width           =   1005
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   300
      Left            =   8640
      TabIndex        =   5
      Top             =   5910
      Width           =   945
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   360
      Left            =   90
      TabIndex        =   4
      Top             =   6300
      Width           =   1005
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   300
      Left            =   7650
      TabIndex        =   3
      Top             =   5910
      Width           =   945
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   360
      Left            =   90
      TabIndex        =   2
      Top             =   6300
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   360
      Left            =   1200
      TabIndex        =   1
      Top             =   6300
      Width           =   1005
   End
   Begin Sicmact.FlexEdit Flex 
      Height          =   5775
      Left            =   45
      TabIndex        =   0
      Top             =   90
      Width           =   9630
      _ExtentX        =   16986
      _ExtentY        =   10186
      Cols0           =   27
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   $"frmLogSalidasAFBND.frx":08CA
      EncabezadosAnchos=   "500-800-1000-1200-3000-2000-1200-1200-1200-1200-1200-1000-2000-0-0-1200-1200-3000-1000-1000-1000-1200-1200-1200-3000-1200-1200"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-2-3-X-5-6-7-8-9-10-11-X-13-X-15-16-17-18-19-20-21-22-23-X-25-26"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-1-0-0-0-0-0-0-2-1-0-1-0-0-0-0-0-0-0-0-0-1-0-0-0"
      BackColor       =   16777215
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-R-L-L-L-R-R-R-R-L-C-L-L-L-R-R-L-R-R-R-L-L-L-L-L-L"
      FormatosEdit    =   "0-3-3-0-0-0-2-2-2-3-0-0-0-0-0-3-3-0-3-3-3-0-0-0-0-0-0"
      TextArray0      =   "#"
      lbUltimaInstancia=   -1  'True
      lbPuntero       =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   495
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
      CellBackColor   =   16777215
   End
End
Attribute VB_Name = "frmLogSalidasAFBND"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim ldIni As Date
Dim ldFin As Date

Public Sub Ini(pdIni As Date, pdFin As Date)
    ldIni = pdIni
    ldFin = pdFin
    Me.Show 1
End Sub

Private Sub cmdEditar_Click()
    Activa True
End Sub

Private Sub cmdEliminar_Click()
    If MsgBox("Desea Eliminar el Registro. " & Me.Flex.row & "", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    Me.Flex.EliminaFila Flex.row
End Sub

Private Sub cmdGrabar_Click()
    Dim oDep As DLogDeprecia
    Set oDep = New DLogDeprecia
    Dim lnI As Integer
    
    For lnI = 1 To Me.Flex.Rows - 1
        If Not IsNumeric(Me.Flex.TextMatrix(lnI, 1)) Then
            MsgBox "Debe ingresar el año de activación.", vbInformation, "Aviso"
            Flex.Col = 1
            Flex.row = lnI
            Flex.SetFocus
            Exit Sub
        ElseIf Not IsNumeric(Me.Flex.TextMatrix(lnI, 2)) Then
            MsgBox "Debe ingresar un movimiento valido.", vbInformation, "Aviso"
            Flex.Col = 2
            Flex.row = lnI
            Flex.SetFocus
            Exit Sub
        ElseIf Me.Flex.TextMatrix(lnI, 3) = "" Then
            MsgBox "Debe ingresar un bien valido.", vbInformation, "Aviso"
            Flex.Col = 3
            Flex.row = lnI
            Flex.SetFocus
            Exit Sub
        ElseIf Me.Flex.TextMatrix(lnI, 5) = "" Then
            MsgBox "Debe un numero de serie valido.", vbInformation, "Aviso"
            Flex.Col = 5
            Flex.row = lnI
            Flex.SetFocus
            Exit Sub
        ElseIf Not IsNumeric(Me.Flex.TextMatrix(lnI, 6)) Then
            MsgBox "Debe ingresar un monto valido.", vbInformation, "Aviso"
            Flex.Col = 6
            Flex.row = lnI
            Flex.SetFocus
            Exit Sub
        ElseIf Not IsDate(Me.Flex.TextMatrix(lnI, 10)) Then
            MsgBox "Debe una fecha valido.", vbInformation, "Aviso"
            Flex.Col = 10
            Flex.row = lnI
            Flex.SetFocus
            Exit Sub
        ElseIf CDate(Me.Flex.TextMatrix(lnI, 10)) < ldIni Or CDate(Me.Flex.TextMatrix(lnI, 10)) > ldFin Then
            MsgBox "Debe una fecha esta fuera del rango valido.", vbInformation, "Aviso"
            Flex.Col = 10
            Flex.row = lnI
            Flex.SetFocus
            Exit Sub
        End If
    Next lnI
        
    If MsgBox("Desea Grabar los cambios ??? ", vbQuestion + vbYesNo, "Aviso") = vbNo Then Exit Sub
    
    oDep.SetBSSalidasAFBND Flex.GetRsNew, ldIni, ldFin
    
    Activa False
End Sub

Private Sub cmdNuevo_Click()
    Me.Flex.AdicionaFila
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub cmdValida_Click()
    Dim oDep As DLogDeprecia
    Set oDep = New DLogDeprecia
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rsDep As ADODB.Recordset
    Set rsDep = New ADODB.Recordset
    
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    
    Dim lnItem As Integer
    
    Dim lsCadena As String
    
    Set rs = oDep.GetBSSalidasAFBNDValida(ldIni, ldFin)
    Set rsDep = oDep.GetBSSalidasAFBNDValidaDeprecia(ldIni, ldFin)
    
    lsCadena = ""
    lnItem = 0
    
    If Not (rs.EOF And rs.BOF) Then
        lsCadena = lsCadena & "---- DUPLICADO ----" & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & "   " & "Item  Año    Nro Mov    Cod.Bien        Serie         Cantidad " & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & "   " & "======================================================== " & oImpresora.gPrnSaltoLinea
        While Not rs.EOF
            lnItem = lnItem + 1
            lsCadena = lsCadena & "   " & Format(lnItem, "000") & "  " & Format(rs!nAnio, "0000") & "   " & Format(rs!nMovNro, "0000000000") & "    " & rs!cBSCod & "        " & rs!cSerie & "         " & rs!Num & " " & oImpresora.gPrnSaltoLinea
            rs.MoveNext
        Wend
    End If
    
    lnItem = 0
    If Trim(lsCadena) <> "" Then
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & oImpresora.gPrnSaltoLinea
    End If
    
    If Not (rsDep.EOF And rsDep.BOF) Then
        lsCadena = lsCadena & "---- BIENES DE ACTIVO QUE NO TIENEN INDICADO PERIODO O EL GRUPO DE DEPRECIACION ----" & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & "   " & " Item   Cod.Bien        Descricion                               " & oImpresora.gPrnSaltoLinea
        lsCadena = lsCadena & "   " & "======================================================== " & oImpresora.gPrnSaltoLinea
        While Not rsDep.EOF
            lnItem = lnItem + 1
            lsCadena = lsCadena & "   " & Format(lnItem, "000") & "  " & rsDep!cBSCod & "   " & rsDep!cBSDescripcion & oImpresora.gPrnSaltoLinea
            rsDep.MoveNext
        Wend
    End If
    
    
    If Trim(lsCadena) = "" Then
        MsgBox "Validacion OK.", vbInformation, "Aviso"
    Else
        oPrevio.Show lsCadena, Me.Caption, True, , gImpresora
    End If
    
    Set oDep = New DLogDeprecia
    Set rs = New ADODB.Recordset
    Set oPrevio = New Previo.clsPrevio
End Sub

Private Sub Flex_RowColChange()
    Dim oALmacen As DLogAlmacen
    Set oALmacen = New DLogAlmacen
    Dim oArea As DActualizaDatosArea
    Set oArea = New DActualizaDatosArea
            
    If Flex.Col = 3 Then
        Flex.TipoBusqueda = BuscaArbol
        Me.Flex.rsTextBuscar = oALmacen.GetBienesAlmacen(, gnLogBSTpoBienFijo)
    ElseIf Flex.Col = 11 Then
        Flex.TipoBusqueda = BuscaArbol
        Me.Flex.rsTextBuscar = oArea.GetAgenciasAreas
    ElseIf Flex.Col = 23 Then
        Flex.TipoBusqueda = BuscaPersona
        Me.Flex.rsTextBuscar = oArea.GetAgenciasAreas
    End If
    
End Sub

Private Sub Form_Load()
    Dim oDep As DLogDeprecia
    Set oDep = New DLogDeprecia
    
    Me.Flex.rsFlex = oDep.GetBSSalidasAFBND(ldIni, ldFin)

    Activa False
End Sub

Private Sub Activa(pbActiva As Boolean)
    Flex.lbEditarFlex = pbActiva
    Me.cmdEditar.Visible = Not pbActiva
    Me.cmdCancelar.Enabled = pbActiva
    Me.cmdEliminar.Enabled = pbActiva
    Me.cmdNuevo.Enabled = pbActiva
    Me.cmdGrabar.Visible = pbActiva
End Sub
