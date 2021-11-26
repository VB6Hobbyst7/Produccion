VERSION 5.00
Begin VB.Form frmNIIFNotasEstadoConfig 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Notas Estado"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10935
   Icon            =   "frmNIIFNotasEstadoConfig.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Notas Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3705
      Left            =   40
      TabIndex        =   0
      Top             =   40
      Width           =   10875
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "&Detalle"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9675
         TabIndex        =   10
         Top             =   1200
         Width           =   1050
      End
      Begin VB.CommandButton cmdQuitar 
         Caption         =   "&Quitar"
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
         Left            =   9675
         TabIndex        =   9
         Top             =   660
         Width           =   1050
      End
      Begin VB.CommandButton cmdBajar 
         Caption         =   "&Bajar Orden"
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
         Left            =   9675
         TabIndex        =   8
         Top             =   2040
         Width           =   1050
      End
      Begin VB.CommandButton cmdSubir 
         Caption         =   "&Subir Orden"
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
         Left            =   9675
         TabIndex        =   7
         Top             =   1680
         Width           =   1050
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9675
         TabIndex        =   6
         Top             =   3240
         Width           =   1050
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9675
         TabIndex        =   5
         Top             =   2880
         Width           =   1050
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   9675
         TabIndex        =   4
         Top             =   2520
         Width           =   1050
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "&Agregar"
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
         Left            =   9675
         TabIndex        =   3
         Top             =   300
         Width           =   1050
      End
      Begin Sicmact.FlexEdit feNotas 
         Height          =   3405
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   9480
         _ExtentX        =   16722
         _ExtentY        =   6006
         Cols0           =   7
         HighLight       =   1
         EncabezadosNombres=   "#-Descripción Nota-Config. Contable-Periodo-Comentario-nNotaEstado-Aux"
         EncabezadosAnchos=   "350-3000-3900-730-1080-0-0"
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
         ColumnasAEditar =   "X-1-2-3-4-X-X"
         ListaControles  =   "0-0-0-3-3-0-0"
         EncabezadosAlineacion=   "C-L-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         CantEntero      =   9
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         TipoBusqueda    =   0
         lbPuntero       =   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   345
         RowHeight0      =   300
      End
      Begin VB.Label Label1 
         BorderStyle     =   1  'Fixed Single
         Height          =   3405
         Left            =   9600
         TabIndex        =   2
         Top             =   240
         Width           =   1185
      End
   End
End
Attribute VB_Name = "frmNIIFNotasEstadoConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmNIIFNotasEstadoConfig
'** Descripción : Configuración del Reporte Notas Estado creado segun ERS052-2013
'** Creación : EJVG, 20130413 09:00:00 AM
'********************************************************************
Option Explicit
Dim fsOpeCod As String
Dim fsOpeDesc As String
Dim fMatNotasDet As Variant

Private Sub Form_Load()
    CentraForm Me
    IniciarControles
End Sub
Public Sub Inicio(ByVal psOpeCod As String, psOpeDesc As String)
    fsOpeCod = psOpeCod
    fsOpeDesc = psOpeDesc
    Caption = "CONFIGURACIÓN " & UCase(psOpeDesc)
    Show 1
End Sub
Private Sub cmdAgregar_Click()
    Dim MatDetalle As Variant
    ReDim MatDetalle(1 To 14, 0)
    
    If Not (feNotas.Rows - 1 = 1 And Len(Trim(feNotas.TextMatrix(1, 0))) = 0) Then 'Flex no esta Vacio
        If validarRegistroDatosNotasEstado = False Then Exit Sub
    End If
    
    feNotas.AdicionaFila
    ReDim Preserve fMatNotasDet(1 To UBound(fMatNotasDet) + 1)
    fMatNotasDet(UBound(fMatNotasDet)) = MatDetalle

    feNotas.SetFocus
    SendKeys "{Enter}"
End Sub
Private Sub cmdQuitar_Click()
    Dim i As Long, iTmp As Long, k As Long
    Dim MatNotasDetTmp As Variant
    Dim lnNroItemBorrado As Long
    
    lnNroItemBorrado = feNotas.row
    feNotas.EliminaFila lnNroItemBorrado

    If lnNroItemBorrado > 0 Then
        If Not (feNotas.Rows - 1 = 1 And Len(Trim(feNotas.TextMatrix(1, 0))) = 0) Then 'Flex no esta Vacio
            If lnNroItemBorrado > feNotas.Rows - 1 Then 'Borro la ultima fila
                ReDim Preserve fMatNotasDet(1 To UBound(fMatNotasDet) - 1)
            Else 'Borra fila <> ultima fila
                iTmp = 1
                ReDim MatNotasDetTmp(1 To 1)
                For i = lnNroItemBorrado + 1 To UBound(fMatNotasDet) 'A partir del registro borrado en adelante
                    ReDim Preserve MatNotasDetTmp(1 To iTmp)
                    MatNotasDetTmp(iTmp) = fMatNotasDet(i)
                    iTmp = iTmp + 1
                Next
                If lnNroItemBorrado = 1 Then 'Si borro el item 1 nueva matriz
                    ReDim Preserve fMatNotasDet(0)
                    fMatNotasDet = MatNotasDetTmp
                Else 'Si borro otro item
                    ReDim Preserve fMatNotasDet(1 To (lnNroItemBorrado - 1)) 'Redimenziono hasta item antes del borrado
                    For k = 1 To UBound(MatNotasDetTmp)
                        ReDim Preserve fMatNotasDet(1 To UBound(fMatNotasDet) + 1)
                        fMatNotasDet(UBound(fMatNotasDet)) = MatNotasDetTmp(k)
                    Next
                End If
            End If
        Else 'Flex esta vacio
            ReDim Preserve fMatNotasDet(0)
        End If
    End If
    Set MatNotasDetTmp = Nothing
End Sub
Private Sub cmdCancelar_Click()
    IniciarControles
End Sub
Private Sub cmdSubir_Click()
    Dim lsDesc1 As String, lsDesc2 As String
    Dim lsFormula1 As String, lsFormula2 As String
    Dim lsPeriodo1 As String, lsPeriodo2 As String
    Dim lsComentario1 As String, lsComentario2 As String
    Dim lnNotaEstado1 As Integer, lnNotaEstado2 As Integer 'EJVG20140102

    If validarRegistroDatosNotasEstado = False Then Exit Sub

    If feNotas.row > 1 Then
        'cambiamos las posiciones del flex
        lsDesc1 = feNotas.TextMatrix(feNotas.row - 1, 1)
        lsFormula1 = feNotas.TextMatrix(feNotas.row - 1, 2)
        lsPeriodo1 = feNotas.TextMatrix(feNotas.row - 1, 3)
        lsComentario1 = feNotas.TextMatrix(feNotas.row - 1, 4)
        lnNotaEstado1 = Val(feNotas.TextMatrix(feNotas.row - 1, 5))
        
        lsDesc2 = feNotas.TextMatrix(feNotas.row, 1)
        lsFormula2 = feNotas.TextMatrix(feNotas.row, 2)
        lsPeriodo2 = feNotas.TextMatrix(feNotas.row, 3)
        lsComentario2 = feNotas.TextMatrix(feNotas.row, 4)
        lnNotaEstado2 = Val(feNotas.TextMatrix(feNotas.row, 5))
        
        feNotas.TextMatrix(feNotas.row - 1, 1) = lsDesc2
        feNotas.TextMatrix(feNotas.row - 1, 2) = lsFormula2
        feNotas.TextMatrix(feNotas.row - 1, 3) = lsPeriodo2
        feNotas.TextMatrix(feNotas.row - 1, 4) = lsComentario2
        feNotas.TextMatrix(feNotas.row - 1, 5) = lnNotaEstado2
        
        feNotas.TextMatrix(feNotas.row, 1) = lsDesc1
        feNotas.TextMatrix(feNotas.row, 2) = lsFormula1
        feNotas.TextMatrix(feNotas.row, 3) = lsPeriodo1
        feNotas.TextMatrix(feNotas.row, 4) = lsComentario1
        feNotas.TextMatrix(feNotas.row, 5) = lnNotaEstado1
        
        'Cambiamos las posiciones del detalle
        Dim Det1 As Variant, Det2 As Variant
        Det1 = fMatNotasDet(feNotas.row - 1)
        Det2 = fMatNotasDet(feNotas.row)
        
        fMatNotasDet(feNotas.row - 1) = Det2
        fMatNotasDet(feNotas.row) = Det1
                
        feNotas.row = feNotas.row - 1
        feNotas.SetFocus
    End If
End Sub
Private Sub cmdBajar_Click()
    Dim lsDesc1 As String, lsDesc2 As String
    Dim lsFormula1 As String, lsFormula2 As String
    Dim lsPeriodo1 As String, lsPeriodo2 As String
    Dim lsComentario1 As String, lsComentario2 As String
    Dim lnNotaEstado1 As Integer, lnNotaEstado2 As Integer 'EJVG20140102

    If validarRegistroDatosNotasEstado = False Then Exit Sub

    If feNotas.row < feNotas.Rows - 1 Then
        'cambiamos las posiciones del flex
        lsDesc1 = feNotas.TextMatrix(feNotas.row + 1, 1)
        lsFormula1 = feNotas.TextMatrix(feNotas.row + 1, 2)
        lsPeriodo1 = feNotas.TextMatrix(feNotas.row + 1, 3)
        lsComentario1 = feNotas.TextMatrix(feNotas.row + 1, 4)
        lnNotaEstado1 = Val(feNotas.TextMatrix(feNotas.row + 1, 5))
        
        lsDesc2 = feNotas.TextMatrix(feNotas.row, 1)
        lsFormula2 = feNotas.TextMatrix(feNotas.row, 2)
        lsPeriodo2 = feNotas.TextMatrix(feNotas.row, 3)
        lsComentario2 = feNotas.TextMatrix(feNotas.row, 4)
        lnNotaEstado2 = Val(feNotas.TextMatrix(feNotas.row, 5))
        
        feNotas.TextMatrix(feNotas.row + 1, 1) = lsDesc2
        feNotas.TextMatrix(feNotas.row + 1, 2) = lsFormula2
        feNotas.TextMatrix(feNotas.row + 1, 3) = lsPeriodo2
        feNotas.TextMatrix(feNotas.row + 1, 4) = lsComentario2
        feNotas.TextMatrix(feNotas.row + 1, 5) = lnNotaEstado2
        
        feNotas.TextMatrix(feNotas.row, 1) = lsDesc1
        feNotas.TextMatrix(feNotas.row, 2) = lsFormula1
        feNotas.TextMatrix(feNotas.row, 3) = lsPeriodo1
        feNotas.TextMatrix(feNotas.row, 4) = lsComentario1
        feNotas.TextMatrix(feNotas.row, 5) = lnNotaEstado1
        
        'Cambiamos las posiciones del detalle
        Dim Det1 As Variant, Det2 As Variant
        Det1 = fMatNotasDet(feNotas.row + 1)
        Det2 = fMatNotasDet(feNotas.row)
        
        fMatNotasDet(feNotas.row + 1) = Det2
        fMatNotasDet(feNotas.row) = Det1
        
        feNotas.row = feNotas.row + 1
        feNotas.SetFocus
    End If
End Sub
Private Sub cmdDetalle_Click()
    If Not (feNotas.Rows - 1 = 1 And Len(Trim(feNotas.TextMatrix(1, 0))) = 0) Then 'Flex no esta Vacio
        If validarRegistroDatosNotasEstado = False Then Exit Sub
    End If
    fMatNotasDet(feNotas.row) = frmNIIFNotasEstadoConfigDet.Inicio(fsOpeCod, fsOpeDesc, fMatNotasDet(feNotas.row))
End Sub
Private Sub cmdGuardar_Click()
    Dim oRep As New NRepFormula
    Dim lsMovNro As String
    Dim lbExito As Boolean
    Dim MatNotas As Variant
    Dim i As Long
    
    If validarGrabar = False Then Exit Sub
    If validarRegistroDatosNotasEstado = False Then Exit Sub
    
    If MsgBox("¿Esta seguro de guardar la configuración de las Notas de Estado?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    
    'ReDim MatNotas(1 To 5, 0)
    ReDim MatNotas(1 To 6, 0) 'EJVG20140102
    For i = 1 To feNotas.Rows - 1
        'ReDim Preserve MatNotas(1 To 5, 1 To i)
        ReDim Preserve MatNotas(1 To 6, 1 To i) 'EJVG20140102
        MatNotas(1, i) = Trim(feNotas.TextMatrix(i, 1)) 'Descripcion
        MatNotas(2, i) = Trim(feNotas.TextMatrix(i, 2)) 'Formula
        MatNotas(3, i) = IIf(Trim(Right(feNotas.TextMatrix(i, 3), 10)) = "1", True, False) 'Periodo
        MatNotas(4, i) = IIf(Trim(Right(feNotas.TextMatrix(i, 4), 10)) = "1", True, False) 'Comentario
        MatNotas(5, i) = fMatNotasDet(i) 'Nota Estado Detalle
        MatNotas(6, i) = Val(Trim(feNotas.TextMatrix(i, 5))) 'Id Nota Estado
    Next
    
    lbExito = oRep.RegistrarNotasEstado(fsOpeCod, MatNotas, gdFecSis, Right(gsCodAge, 2), gsCodUser)
    
    If lbExito Then
        MsgBox "Se ha grabado satisfactoriamente los cambios de las Notas Estado", vbInformation, "Aviso"
        cmdCancelar_Click
    Else
        MsgBox "No se ha podido grabar los cambios realizados, vuelva a intentarlo, si persiste comuniquese con el Dpto. de TI", vbCritical, "Aviso"
    End If
    
    Set oRep = Nothing
    Set MatNotas = Nothing
End Sub
Private Sub cmdsalir_Click()
    If MsgBox("¿Esta seguro de salir de la configuración de las Notas de Estado?", vbInformation + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If
    Unload Me
End Sub
Private Sub IniciarControles()
    ListarConfiguracionNotas
    cmdSubir.Enabled = False
    cmdBajar.Enabled = False
    cmdDetalle.Enabled = False
End Sub
Private Sub ListarConfiguracionNotas()
    Dim oRep As New NRepFormula
    Dim rsNotas As New ADODB.Recordset
    Dim rsNotasDet As New ADODB.Recordset
    Dim Detalle As Variant
    Dim iCab As Long, iDet As Long
    
    Set rsNotas = oRep.RecuperaConfigRepNotasEstado(fsOpeCod)
    Call LimpiaFlex(feNotas)

    Set fMatNotasDet = Nothing
    ReDim fMatNotasDet(0)

    If Not RSVacio(rsNotas) Then
        ReDim fMatNotasDet(1 To rsNotas.RecordCount)
        For iCab = 1 To rsNotas.RecordCount
            feNotas.AdicionaFila
            'Notas
            feNotas.TextMatrix(feNotas.row, 1) = rsNotas!cDescripcion
            feNotas.TextMatrix(feNotas.row, 2) = rsNotas!cFormula
            feNotas.TextMatrix(feNotas.row, 3) = IIf(rsNotas!bPeriodo, "SI" & Space(75) & "1", "NO" & Space(75) & "2")
            feNotas.TextMatrix(feNotas.row, 4) = IIf(rsNotas!bComentario, "SI" & Space(75) & "1", "NO" & Space(75) & "2")
            feNotas.TextMatrix(feNotas.row, 5) = rsNotas!nNotaEstado 'EJVG20140102
            'Detalle
            Set rsNotasDet = oRep.RecuperaConfigRepNotasEstadoDetalle(fsOpeCod, rsNotas!nId, rsNotas!nNotaEstado)
            Set Detalle = Nothing
            ReDim Detalle(1 To 14, rsNotasDet.RecordCount)
            For iDet = 1 To rsNotasDet.RecordCount
                Detalle(1, iDet) = rsNotasDet!nTpoDetalle
                Detalle(2, iDet) = rsNotasDet!cDescripcion
                Detalle(3, iDet) = rsNotasDet!nNivel
                Detalle(4, iDet) = IIf(rsNotasDet!bNegrita, 1, 2)
                Detalle(5, iDet) = rsNotasDet!cFormula1
                Detalle(6, iDet) = rsNotasDet!cFormula1_2012
                Detalle(7, iDet) = rsNotasDet!cFormula2
                Detalle(8, iDet) = rsNotasDet!cFormula2_2012
                Detalle(9, iDet) = rsNotasDet!cFormula3
                Detalle(10, iDet) = rsNotasDet!cFormula3_2012
                Detalle(11, iDet) = rsNotasDet!cFormula4
                Detalle(12, iDet) = rsNotasDet!cFormula4_2012
                Detalle(13, iDet) = rsNotasDet!cFormula5
                Detalle(14, iDet) = rsNotasDet!cFormula5_2012
                rsNotasDet.MoveNext
            Next
            fMatNotasDet(iCab) = Detalle
            rsNotas.MoveNext
        Next
        feNotas.TopRow = 1
        feNotas.row = 1
    End If
    Set oRep = Nothing
    Set rsNotas = Nothing
    Set rsNotasDet = Nothing
    Set Detalle = Nothing
End Sub
Private Sub feNotas_RowColChange()
    Dim rsOpt As New ADODB.Recordset
    If feNotas.Col = 3 Or feNotas.Col = 4 Then
        With rsOpt
            .Fields.Append "desc", adVarChar, 10
            .Fields.Append "value", adVarChar, 1
        End With
        If feNotas.Col = 3 Or feNotas.Col = 4 Then
            With rsOpt
                .Open
                .AddNew
                .Fields("desc") = "SI"
                .Fields("value") = "1"
                .AddNew
                .Fields("desc") = "NO"
                .Fields("value") = "2"
            End With
        End If
        rsOpt.MoveFirst
        feNotas.CargaCombo rsOpt
    End If
    Set rsOpt = Nothing
End Sub
Private Sub feNotas_OnCellChange(pnRow As Long, pnCol As Long)
    If pnCol = 1 Or pnCol = 2 Then
        feNotas.TextMatrix(pnRow, pnCol) = UCase(feNotas.TextMatrix(pnRow, pnCol))
    End If
End Sub
Private Sub feNotas_Click()
    If feNotas.row > 0 Then
        If feNotas.TextMatrix(feNotas.row, 0) <> "" Then
            cmdSubir.Enabled = True
            cmdBajar.Enabled = True
            cmdDetalle.Enabled = True
        End If
    End If
End Sub
Private Function validarRegistroDatosNotasEstado() As Boolean
    validarRegistroDatosNotasEstado = True
    Dim i As Long, j As Long
    For i = 1 To feNotas.Rows - 1 'valida fila x fila
        For j = 1 To feNotas.Cols - 1
            If feNotas.ColWidth(j) > 0 And j <> 2 Then 'xq la plantilla contable es opcional
                If Trim(feNotas.TextMatrix(i, j)) = "" Then
                    validarRegistroDatosNotasEstado = False
                    MsgBox "Ud. debe de ingresar el dato '" & UCase(feNotas.TextMatrix(0, j)) & "'", vbInformation, "Aviso"
                    feNotas.row = i
                    feNotas.Col = j
                    feNotas.SetFocus
                    Exit Function
                End If
            End If
        Next
    Next
End Function
Private Function validarGrabar() As Boolean
    Dim i As Long
    validarGrabar = True
    If feNotas.Rows - 1 = 1 And feNotas.TextMatrix(1, 0) = "" Then
        MsgBox "Ud. debe de registrar las Notas de Estado", vbCritical, "Aviso"
        validarGrabar = False
        feNotas.SetFocus
        Exit Function
    End If
End Function
