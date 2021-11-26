VERSION 5.00
Begin VB.Form FrmCompraVentaAut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizacion de Compra Venta de Moneda Extranjera"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   16290
   ForeColor       =   &H00000000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   16290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton CmdAprobar 
      Caption         =   "&Enviar Propuesta"
      Height          =   375
      Left            =   10080
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton CmdDenegar 
      BackColor       =   &H00FFFFFF&
      Caption         =   "&Anular"
      Height          =   375
      Left            =   11640
      TabIndex        =   3
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancelar 
      BackColor       =   &H80000016&
      Caption         =   "&Limpiar"
      Height          =   375
      Left            =   13200
      TabIndex        =   2
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame FraTC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   4935
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   16095
      Begin SICMACT.FlexEdit Flex 
         Height          =   4455
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   15735
         _ExtentX        =   27755
         _ExtentY        =   7858
         Cols0           =   12
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Fecha Hora-Check-Operacion-TC Actual-TC Nuevo-Monto-User-Cliente-IDAut-Agencia-Estado"
         EncabezadosAnchos=   "600-1900-800-1300-1200-1200-1200-800-3200-0-1500-1500"
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
         ColumnasAEditar =   "X-X-2-X-X-5-X-X-X-X-X-X"
         ListaControles  =   "0-0-4-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-C-L-R-R-R-C-L-C-C-C"
         FormatosEdit    =   "0-5-5-1-2-2-2-1-3-3-3-3"
         CantDecimales   =   4
         TextArray0      =   "Nro"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton CmdSalir 
      BackColor       =   &H00E0E0E0&
      Caption         =   "&Salir"
      Height          =   375
      Left            =   14760
      TabIndex        =   0
      Top             =   5040
      Width           =   1455
   End
End
Attribute VB_Name = "FrmCompraVentaAut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'By capi 21012009
Dim objPista As COMManejador.Pista
'End by
Dim valorCelda As String 'GIPO 14-01-2017


Private Sub cmdAprobar_Click()
Dim Opt As Integer
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim lblTitulo As String
Dim lsMovNro As String
Dim oGen  As COMNContabilidad.NCOMContFunciones
Dim ObjTcP As COMDConstSistema.DCOMTCEspPermiso
Dim i As Integer
Dim nTCNew As Currency
Dim nIDAprob As Long
Dim oImp As COMNCaptaGenerales.NCOMCaptaImpresion
Dim lbReimp As Boolean
Dim lsBoleta As String

If Not ValidaMatriz Then
    Exit Sub
End If
Set oGen = New COMNContabilidad.NCOMContFunciones
Set oImp = New COMNCaptaGenerales.NCOMCaptaImpresion
Set ObjTcP = New COMDConstSistema.DCOMTCEspPermiso
If MsgBox("Esta Seguro de Grabar", vbInformation + vbYesNo, "AVISO") = vbNo Then Exit Sub
For i = 1 To Me.Flex.rows - 1
    If Me.Flex.TextMatrix(i, 2) = "." Then
        lsMovNro = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        nTCNew = CCur(Flex.TextMatrix(i, 5))
        nIDAprob = CLng(Flex.TextMatrix(i, 9))
'        Call ObjTcP.Aprobar_TC_Especial(lsMovNro, nTCNew, gdFecSis, gsCodUser, nIDAprob)
        Call ObjTcP.CambioEstado_TC_Especial(lsMovNro, 2, nTCNew, gdFecSis, gsCodUser, nIDAprob) 'APRI20180201 MEJORA INC181005004
        '**************
        'By Capi 21012009
        objPista.InsertarPista gsOpeCod, lsMovNro, gsCodPersUser, GetMaquinaUsuario, gModificar, , nIDAprob, gNumeroAprobacion
        'End by


        lblTitulo = "Aprobar Tipo Cambio Especial"
    
        lsBoleta = ""
        
        'comentado por GIPO ERS069-2016  25-01-2017
        'lsBoleta = oImp.ImprimeBoletaAutorizacionTCE(LblTitulo, Flex.TextMatrix(i, 3), gsOpeCod, CStr(nIDAprob), _
                    CStr(Format(nTCNew, "#.00")), Format(CDbl(Flex.TextMatrix(i, 6)), "#.00"), gsNomAge, gsNomCmac, gdFecSis, gsCodUser)
        
'        lbReimp = True
'        Do While lbReimp
'             If Trim(lsBoleta) <> "" Then
'                nFicSal = FreeFile
'                Open sLpt For Output As nFicSal
'                    Print #nFicSal, lsBoleta
'                    Print #nFicSal, ""
'                Close #nFicSal
'             End If
'
'            If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
'                lbReimp = False
'            End If
'        Loop
    '**************
    End If
Next i
cmdCancelar_Click
End Sub
 Function ValidaMatriz(Optional ByVal pcVar As String = "A") As Boolean
Dim i As Integer
Dim ban As Boolean
If Trim(Flex.TextMatrix(1, 1)) = "" Then
    MsgBox "No hay Valores para Aprobar ò Rechazar", vbInformation, "Aviso"
    ValidaMatriz = False
    Exit Function
End If
ban = False
For i = 1 To Me.Flex.rows - 1
    If Me.Flex.TextMatrix(i, 2) = "." Then
        ban = True
    End If
Next i

If ban = False Then
    MsgBox "No se ha seleccionado ningún registro", vbInformation, "AVISO"
    ValidaMatriz = False
    Exit Function
End If

If pcVar = "A" Then
    For i = 1 To Me.Flex.rows - 1
        If Me.Flex.TextMatrix(i, 2) = "." And val(Me.Flex.TextMatrix(i, 5)) = 0 Then
            MsgBox "El Tipo de Cambio no Puede ser 0", vbInformation, "AVISO"
            ValidaMatriz = False
            Flex.row = i
            Flex.Col = 5
            Exit Function
        End If
    Next i
End If
ValidaMatriz = True
End Function

Private Sub cmdCancelar_Click()
CargaTC
End Sub

Private Sub cmdDenegar_Click()
Dim Opt As Integer
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim lblTitulo As String
Dim lsMovNro As String

Dim oGen  As COMNContabilidad.NCOMContFunciones
Dim ObjTcP As COMDConstSistema.DCOMTCEspPermiso

Dim i As Integer
Dim nTCNew As Currency
Dim nIDAprob As Long
Dim lsBoleta As String

Dim oImp As COMNCaptaGenerales.NCOMCaptaImpresion  'NCapImpBoleta



Dim lbReimp As Boolean

If Not ValidaMatriz("R") Then
    Exit Sub
End If
Set oGen = New COMNContabilidad.NCOMContFunciones
Set oImp = New COMNCaptaGenerales.NCOMCaptaImpresion
Set ObjTcP = New COMDConstSistema.DCOMTCEspPermiso
If MsgBox("Esta Seguro de Grabar", vbInformation + vbYesNo, "AVISO") = vbNo Then Exit Sub

For i = 1 To Me.Flex.rows - 1
    If Me.Flex.TextMatrix(i, 2) = "." Then
        lsMovNro = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        nTCNew = CCur(Flex.TextMatrix(i, 5))
        nIDAprob = CLng(Flex.TextMatrix(i, 9))
'        Call ObjTcP.Rechazar_TC_Especial(nTCNew, gdFecSis, gsCodUser, nIDAprob)
        Call ObjTcP.CambioEstado_TC_Especial(lsMovNro, 1, 0, gdFecSis, gsCodUser, nIDAprob) 'APRI20180201 MEJORA INC181005004
    '************** impresión comentada por GIPO según ERS069-2016  25-01-2017
'        LblTitulo = "Rechazo Tipo Cambio Especial"
'
'        lsBoleta = ""
        
'        lsBoleta = oImp.ImprimeBoletaAutorizacionTCE(LblTitulo, Flex.TextMatrix(i, 3), gsOpeCod, CStr(nIDAprob), _
'                    CStr(Format(nTCNew, "#.00")), Format(CDbl(Flex.TextMatrix(i, 6)), "#.00"), gsNomAge, gsNomCmac, gdFecSis, gsCodUser)
'
'        lbReimp = True
'        Do While lbReimp
'             If Trim(lsBoleta) <> "" Then
'                nFicSal = FreeFile
'                Open sLpt For Output As nFicSal
'                    Print #nFicSal, lsBoleta
'                    Print #nFicSal, ""
'                Close #nFicSal
'             End If
'
'            If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
'                lbReimp = False
'            End If
'        Loop
    '**************
    End If
Next i
cmdCancelar_Click
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub Flex_EnterCell()
   valorCelda = Flex.TextMatrix(Flex.row, Flex.Col)
End Sub

Private Sub Flex_OnCellChange(pnRow As Long, pnCol As Long)
    If (pnCol = 5 Or pnCol = 6 Or pnCol = 7) Then 'si corresponde a la columna de tipo de cambio nuevo
    
        Dim valor_anterior As Currency
        Dim valor_nuevo As Currency
        Dim estadoTipoCambio As String
        
        If (valorCelda = "") Then
            valorCelda = "0.00"
        End If
        
        valor_anterior = CCur(valorCelda)
       
        If (IsNumeric(Flex.TextMatrix(pnRow, pnCol))) Then
            valor_nuevo = CCur(Flex.TextMatrix(pnRow, pnCol))
            estadoTipoCambio = Flex.TextMatrix(pnRow, 11)
            
            If (valor_nuevo < 0) Then
                MsgBox "No se puede asignar un valor negativo a un TC especial", vbInformation, "Aviso"
                Flex.TextMatrix(pnRow, pnCol) = Format(valor_anterior, "###,##0.0000") 'GIPO CORRECCIÓN 08/02/2017
                Exit Sub
            End If
            
            If (estadoTipoCambio = "PENDIENTE") Then
                Flex.TextMatrix(pnRow, pnCol) = Format(valor_nuevo, "###,##0.0000") 'GIPO CORRECCIÓN 08/02/2017
            Else
                'no cambia el valor de la celda
                Flex.TextMatrix(pnRow, pnCol) = Format(valor_anterior, "###,##0.0000") 'GIPO CORRECCIÓN 08/02/2017
            End If
            
           
            
        Else
            MsgBox "Formato incorrecto!", vbInformation, "Aviso"
            Flex.TextMatrix(pnRow, pnCol) = Format(valor_anterior, "###,##0.0000") 'GIPO CORRECCIÓN 08/02/2017
        End If
        
        
        
    End If
   
End Sub

Private Sub Flex_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
    Dim i As Integer
    i = 11
     'For i = 1 To Me.Flex.Cols - 1
        If (Flex.TextMatrix(pnRow, i) <> "PENDIENTE") Then
            'Flex.row = i
            Flex.TextMatrix(pnRow, pnCol) = "" 'para el check habilitado es "."
        End If
        
    'Next i
End Sub

Public Sub Inicio()
'GIPO
Dim ObjTcP As COMDConstSistema.DCOMTCEspPermiso
Set ObjTcP = New COMDConstSistema.DCOMTCEspPermiso
Dim RS As ADODB.Recordset
Set RS = ObjTcP.GetTc_AccesoFormulario(gsCodUser)
If (RS!acceso = "PERMITIDO") Then
    'By Capi 20012009
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
    Set objPista = New COMManejador.Pista
    gsOpeCod = gCapAprobarTCEspecial
    CargaTC
    Me.Show 1
    'End By
Else
    MsgBox "Lo sentimos, su cargo no tiene permiso para acceder a éste formulario", vbInformation, "Aviso"
    Exit Sub
End If
End Sub
Sub CargaTC(Optional psCodAge As String = "00")
 Dim ObjTcP As COMDConstSistema.DCOMTCEspPermiso
 Dim RS As ADODB.Recordset
' Dim ban As Boolean
 Dim i As Integer
 Set ObjTcP = New COMDConstSistema.DCOMTCEspPermiso
 Set RS = New ADODB.Recordset
 Set RS = ObjTcP.Get_All(gdFecSis, psCodAge)
 Flex.Clear
 Flex.FormaCabecera
 Flex.rows = 2
 Dim Estado As String
 Estado = "NO DEFINIDA"
'ban = False
If Not (RS.EOF And Not RS.BOF) Then
    While Not RS.EOF
'        If ban Then
'            Flex.Rows = Flex.Rows + 1
'        End If
'        ban = True
        Flex.AdicionaFila
        Flex.TextMatrix(Flex.rows - 1, 1) = Format(RS!dFechaReg, "DD/MM/YYYY HH:MM:SS AMPM")
        'Flex.TextMatrix(Flex.Rows - 1, 3) = ""
        Flex.TextMatrix(Flex.rows - 1, 3) = RS!cOpedesc
        Flex.TextMatrix(Flex.rows - 1, 4) = Format(RS!nTCReg, "#0.0000") 'GIPO CORRECCIÓN 08/02/2017
        If (RS!nEstado <> 0) Then
            Flex.TextMatrix(Flex.rows - 1, 5) = Format(RS!nTCAprob, "#0.0000") 'GIPO CORRECCIÓN 08/02/2017
            Flex.BackColorRow (&H80000016) 'APRI20180201 MEJORA INC181005004
        Else
            Flex.TextMatrix(Flex.rows - 1, 5) = Format(0, "#0.0000")
        End If
        
        Flex.TextMatrix(Flex.rows - 1, 6) = Format(RS!nMontoReg, "###,##0.0000") 'GIPO CORRECCIÓN 08/02/2017
        Flex.TextMatrix(Flex.rows - 1, 7) = Right(RS!cMovNro, 4)
        Flex.TextMatrix(Flex.rows - 1, 8) = RS!cPersNombre
        Flex.TextMatrix(Flex.rows - 1, 9) = RS!nCodAut
        Flex.TextMatrix(Flex.rows - 1, 10) = RS!cAgeDescripcion
        
'        If (RS!nEstado = 0) Then
'            Estado = "PENDIENTE"
'        ElseIf (RS!nEstado = 1) Then
'            Estado = "PROP. ANULADA"
'            Flex.BackColorRow (&H80000016)
'        ElseIf (RS!nEstado = 2) Then
'            Estado = "PROP. ENVIADA"
'            Flex.BackColorRow (&H80000016)
'        ElseIf (RS!nEstado = 3) Then
'            Estado = "RECHAZADA" 'RECHAZADA POR EL CLIENTE
'            Flex.BackColorRow (&H80000016)
'        ElseIf (RS!nEstado = 4) Then
'            Estado = "APROBADA" 'APROBADA POR EL CLIENTE
'            Flex.BackColorRow (&H80000016)
'        End If
'        Flex.TextMatrix(Flex.rows - 1, 11) = Estado

        If (RS!nEstado <> 0) Then
            Flex.BackColorRow (&H80000016)
        End If
        
        Flex.TextMatrix(Flex.rows - 1, 11) = RS!cEstado 'APRI20180201 MEJORA INC181005004
        RS.MoveNext
    Wend
    Flex.Col = 5
    For i = 1 To Me.Flex.rows - 1
        If (Flex.TextMatrix(i, 11) = "PENDIENTE") Then
            Flex.row = i
            Flex.CellBackColor = &HFFFFC0
        End If
 
    Next i
    
End If
RS.Close
Set ObjTcP = Nothing
Set RS = Nothing
End Sub








