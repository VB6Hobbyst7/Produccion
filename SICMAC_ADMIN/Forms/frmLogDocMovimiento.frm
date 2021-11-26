VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLogDocMovimiento 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7710
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7830
   Icon            =   "frmLogDocMovimiento.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7710
   ScaleWidth      =   7830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   360
      Left            =   5340
      TabIndex        =   7
      Top             =   7290
      Width           =   1185
   End
   Begin MSMask.MaskEdBox mskFecIni 
      Height          =   315
      Left            =   735
      TabIndex        =   3
      Top             =   45
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   360
      Left            =   6585
      TabIndex        =   2
      Top             =   7290
      Width           =   1185
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "&Procesar"
      Height          =   360
      Left            =   6570
      TabIndex        =   1
      Top             =   45
      Width           =   1185
   End
   Begin Sicmact.FlexEdit Flex 
      Height          =   6765
      Left            =   15
      TabIndex        =   0
      Top             =   450
      Width           =   7770
      _extentx        =   13705
      _extenty        =   11933
      cols0           =   10
      highlight       =   1
      allowuserresizing=   3
      encabezadosnombres=   "#-Documento-Fecha-Movimiento-Persona-Discripcion-Estado-Cod Estado-Estado Mov-Cod Est mov"
      encabezadosanchos=   "400-1200-1200-1200-3000-5000-2000-1200-2000-1200"
      font            =   "frmLogDocMovimiento.frx":030A
      font            =   "frmLogDocMovimiento.frx":0332
      font            =   "frmLogDocMovimiento.frx":035A
      font            =   "frmLogDocMovimiento.frx":0382
      font            =   "frmLogDocMovimiento.frx":03AA
      fontfixed       =   "frmLogDocMovimiento.frx":03D2
      columnasaeditar =   "X-X-X-X-X-X-X-X-X-X"
      textstylefixed  =   3
      listacontroles  =   "0-0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-L-L-L-L-L-L-L-L-L"
      formatosedit    =   "0-0-0-0-0-0-0-0-0-0"
      textarray0      =   "#"
      selectionmode   =   1
      lbpuntero       =   -1
      appearance      =   0
      colwidth0       =   405
      rowheight0      =   300
   End
   Begin MSMask.MaskEdBox mskFecFin 
      Height          =   315
      Left            =   2640
      TabIndex        =   4
      Top             =   45
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Fin :"
      Height          =   225
      Left            =   2190
      TabIndex        =   6
      Top             =   105
      Width           =   405
   End
   Begin VB.Label lblini 
      Caption         =   "Inicio :"
      Height          =   225
      Left            =   210
      TabIndex        =   5
      Top             =   90
      Width           =   735
   End
End
Attribute VB_Name = "frmLogDocMovimiento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsDocTpo As String
'ARLO 20170126******************
Dim objPista As COMManejador.Pista
'*******************************

Public Sub Inicio(psDocTpo As String, psCaption As String)
    lsDocTpo = psDocTpo
    Caption = psCaption
    Me.Show 1
End Sub

Private Sub CmdImprimir_Click()
    Dim lsItem As String * 5
    Dim lsDocNro As String * 15
    Dim lsFecha As String * 10
    Dim lsPersona As String * 30
    Dim lsComentario As String * 50
    Dim lsEstado As String * 17
    Dim lsFlag As String * 10
    Dim lnI As Integer
    Dim lsCadena As String
    Dim oPrevio As Previo.clsPrevio
    Set oPrevio = New Previo.clsPrevio
    Dim lsCabecera As String
    Dim lnItem As Long
    
    lnItem = 0
    For lnI = 0 To Me.Flex.Rows - 1
        lsItem = Format(Me.Flex.TextMatrix(lnI, 0), "0000")
        lsDocNro = Me.Flex.TextMatrix(lnI, 1)
        lsFecha = Me.Flex.TextMatrix(lnI, 2)
        lsPersona = Me.Flex.TextMatrix(lnI, 4)
        lsComentario = Me.Flex.TextMatrix(lnI, 5)
        lsEstado = Me.Flex.TextMatrix(lnI, 6)
        lsFlag = Me.Flex.TextMatrix(lnI, 8)
        
        lsCadena = lsCadena & lsItem & Space(2) & lsDocNro & Space(2) & lsFecha & Space(2) & lsPersona & Space(2) & lsComentario & Space(2) & lsEstado & Space(2) & lsFlag & oImpresora.gPrnSaltoLinea
        
        lnItem = lnItem + 1
        
        If lnI = 0 Then
            lsCadena = lsCadena & String(150, "=") & oImpresora.gPrnSaltoLinea
            lsCabecera = lsCadena
        End If
        
        If lnItem > 54 Then
            lsCadena = lsCadena & String(150, "=") & oImpresora.gPrnSaltoLinea
            lnItem = 2
            lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
            lsCadena = lsCadena & oImpresora.gPrnSaltoLinea & lsCabecera
        End If

    Next lnI
    lsCadena = lsCadena & String(150, "=") & oImpresora.gPrnSaltoLinea
    
    oPrevio.Show lsCadena, Caption, True, 66
        'ARLO 20160126 ***
        Dim lsPalabras As String
        If (gsOpeCod = 591501) Then
        lsPalabras = "Movimientos de Requerimientos"
        ElseIf (gsOpeCod = 591502) Then
        lsPalabras = "Notas de Ingreso"
        ElseIf (gsOpeCod = 591503) Then
        lsPalabras = "Guias de Salida"
        End If
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, "", "Se Imprimio el Reporte " & lsPalabras & " del " & mskFecIni & " al " & mskFecFin
        Set objPista = Nothing
        '**************
End Sub

Private Sub cmdProcesar_Click()
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim Sql As String
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    
    If Not IsDate(Me.mskFecIni.Text) Then
        MsgBox "Debe ingresar una fecha corecta.", vbInformation, "Aviso"
        Me.mskFecIni.SetFocus
        Exit Sub
    ElseIf Not IsDate(Me.mskFecFin.Text) Then
        MsgBox "Debe ingresar una fecha corecta.", vbInformation, "Aviso"
        Me.mskFecFin.SetFocus
        Exit Sub
    End If
    
    Sql = " Select MD.cDocNro, Convert(varchar(10),dDocFecha,103) Fecha , M.cMovNro, PE.cPersNombre, cMovDesc, CO.cConsDescripcion, nMovEstado, CO1.cConsDescripcion, nMovFlag" _
        & " From MovDoc MD" _
        & " Inner Join Mov M On MD.nMovNro =  M.nMovNro" _
        & " Left  Join MovGasto MG On MG.nMovNro = M.nMovNro " _
        & " Left  Join Persona  PE On MG.cPersCod = PE.cPersCod " _
        & " Left  Join Constante CO On CO.nConsCod = '4006' And nMovEstado = CO.nConsValor" _
        & " Left  Join Constante CO1 On CO1.nConsCod = '4007' And nMovFlag = CO1.nConsValor" _
        & " where nDocTpo = '" & lsDocTpo & "' And M.nMovNro In" _
        & " (Select Max(MMD.nMovNro) From MovDoc MMD Where MMD.cDocNro = MD.cDocNro And nDocTpo = '" & lsDocTpo & "')" _
        & " And dDocFecha Between '" & Format(Me.mskFecIni.Text, gcFormatoFecha) & "' And '" & Format(Me.mskFecFin.Text, gcFormatoFecha) & "'" _
        & " Order by MD.cDocNro"
    oCon.AbreConexion
    
    Set rs = oCon.CargaRecordSet(Sql)
    
    Flex.Clear
    Flex.Rows = 2
    Flex.FormaCabecera
    
    If rs.EOF And rs.BOF Then
        MsgBox "No se encontro informacion en las fecahs especificadas.", vbInformation, "Aviso"
        Me.mskFecIni.SetFocus
        Exit Sub
    End If
    
    Me.Flex.rsFlex = rs
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Flex_DblClick()
    If lsDocTpo = "70" Then 'Requerimiento
        frmLogSalAlmacen.InicioRepReq Me.Flex.TextMatrix(Me.Flex.row, 1)
    ElseIf lsDocTpo = "71" Then 'Guia de Remision
        frmLogSalAlmacen.InicioRep Me.Flex.TextMatrix(Me.Flex.row, 1)
    ElseIf lsDocTpo = "42" Then 'Nota de Ingreso
        If Me.Flex.TextMatrix(Me.Flex.row, 7) = gMovEstContabNoContable Then
            MsgBox "La nota de ingreso no ha sido confirmada. ", vbInformation
            Exit Sub
        End If
        frmLogIngAlmacen.InicioRep Me.Flex.TextMatrix(Me.Flex.row, 1)
    End If
End Sub

Private Sub Form_Activate()
    Me.mskFecIni.SetFocus
End Sub

Private Sub mskFecFin_GotFocus()
    Me.mskFecFin.SelStart = 0
    Me.mskFecFin.SelLength = 50
End Sub

Private Sub mskFecFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdProcesar.SetFocus
    End If
End Sub

Private Sub mskFecIni_GotFocus()
    Me.mskFecIni.SelStart = 0
    Me.mskFecIni.SelLength = 50
End Sub

Private Sub mskFecIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.mskFecFin.SetFocus
    End If
End Sub
