VERSION 5.00
Begin VB.Form FrmCompraVentaAut 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorizacion de Compra Venta de Moneda Extranjera"
   ClientHeight    =   3555
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9810
   Icon            =   "FrmCompraVentaAut1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3555
   ScaleWidth      =   9810
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton CmdAprobar 
      Caption         =   "&Aprobar"
      Height          =   375
      Left            =   1920
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton CmdDenegar 
      Caption         =   "&Denegar"
      Height          =   375
      Left            =   3480
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.CommandButton CmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5040
      TabIndex        =   2
      Top             =   3120
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
      Height          =   3015
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   9615
      Begin SICMACT.FlexEdit Flex 
         Height          =   2055
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   9255
         _ExtentX        =   16325
         _ExtentY        =   3625
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "Nro-Fecha Hora-Check-Operacion-TC Actual-TC Nuevo-Monto-User-IDAut"
         EncabezadosAnchos=   "600-1900-800-1300-1200-1200-1200-800-0"
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
         ColumnasAEditar =   "X-X-2-X-X-5-X-X-X"
         ListaControles  =   "0-0-4-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-R-C-L-R-R-R-C-R"
         FormatosEdit    =   "0-5-5-1-2-2-2-1-3"
         CantDecimales   =   3
         TextArray0      =   "Nro"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   600
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton CmdSalir 
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6600
      TabIndex        =   0
      Top             =   3120
      Width           =   1455
   End
End
Attribute VB_Name = "FrmCompraVentaAut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CmdAprobar_Click()
Dim opt As Integer
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
For i = 1 To Me.Flex.Rows - 1
    If Me.Flex.TextMatrix(i, 2) = "." Then
        lsMovNro = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        nTCNew = CCur(Flex.TextMatrix(i, 5))
        nIDAprob = CLng(Flex.TextMatrix(i, 8))
        Call ObjTcP.Aprobar_TC_Especial(lsMovNro, nTCNew, gdFecSis, gsCodUser, nIDAprob)
        '**************
        lblTitulo = "Aprobar Tipo Cambio Especial"
    
        lsBoleta = ""
        
        lsBoleta = oImp.ImprimeBoletaAutorizacionTCE(lblTitulo, Flex.TextMatrix(i, 3), gsOpeCod, CStr(nIDAprob), _
                    CStr(Format(nTCNew, "#.00")), Format(CDbl(Flex.TextMatrix(i, 6)), "#.00"), gsNomAge, gsNomCmac, gdFecSis, gsCodUser)
        
        lbReimp = True
        Do While lbReimp
             If Trim(lsBoleta) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, lsBoleta
                    Print #nFicSal, ""
                Close #nFicSal
             End If
                   
            If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbReimp = False
            End If
        Loop
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
For i = 1 To Me.Flex.Rows - 1
    If Me.Flex.TextMatrix(i, 2) = "." Then
        ban = True
    End If
Next i

If ban = False Then
    MsgBox "No se a seleccionado ninguna Operacion", vbInformation, "AVISO"
    ValidaMatriz = False
    Exit Function
End If

If pcVar = "A" Then
    For i = 1 To Me.Flex.Rows - 1
        If Me.Flex.TextMatrix(i, 2) = "." And Val(Me.Flex.TextMatrix(i, 5)) = 0 Then
            MsgBox "El Tipo de Cambio no Puede ser 0", vbInformation, "AVISO"
            ValidaMatriz = False
            Flex.Row = i
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

Private Sub CmdDenegar_Click()
Dim opt As Integer
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim lblTitulo As String
Dim lsMovNro As String

Dim oGen  As COMNContabilidad.NCOMContFunciones
Dim ObjTcP As COMDConstSistema.DCOMTCEspPermiso

Dim i As Integer
Dim nTCNew As Currency
Dim nIDAprob As Long

Dim oImp As COMNCaptaGenerales.NCOMCaptaImpresion  'NCapImpBoleta



Dim lbReimp As Boolean

If Not ValidaMatriz("R") Then
    Exit Sub
End If
Set oGen = New COMNContabilidad.NCOMContFunciones
Set oImp = New COMNCaptaGenerales.NCOMCaptaImpresion
Set ObjTcP = New COMDConstSistema.DCOMTCEspPermiso


For i = 1 To Me.Flex.Rows - 1
    If Me.Flex.TextMatrix(i, 2) = "." Then
        lsMovNro = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        nTCNew = CCur(Flex.TextMatrix(i, 5))
        nIDAprob = CLng(Flex.TextMatrix(i, 8))
        Call ObjTcP.Rechazar_TC_Especial(nTCNew, gdFecSis, gsCodUser, nIDAprob)
    '**************
        lblTitulo = "Rechazo Tipo Cambio Especial"
        'lsBoleta = oImp.ImprimeBoletaCompraVenta(lblTitulo, Flex(I, 3), "", "", "", _
                    nTCNew, gsOpeCod, CCur(Flex(I, 6)), 0, gsNomAge, lsMovNro, sLpt, gsCodCMAC, gsNomCmac)
        lsBoleta = ""
        lsBoleta = oImp.ImprimeBoleta(lblTitulo, Flex.TextMatrix(i, 3), gsOpeCod, CStr(nIDAprob), "", _
                    CStr(Format(nTCNew, "#.00")), gsOpeCod, Format(CDbl(Flex.TextMatrix(i, 6)), "#.00"), "", gsNomAge, 0, 0, False, False, , , , , gsNomCmac, , "", gdFecSis, , gsCodUser, sLpt, , False, , False)
                    
        lbReimp = True
        Do While lbReimp
             If Trim(lsBoleta) <> "" Then
                nFicSal = FreeFile
                Open sLpt For Output As nFicSal
                    Print #nFicSal, lsBoleta
                    Print #nFicSal, ""
                Close #nFicSal
             End If
                   
            If MsgBox("Desea Reimprimir boleta de Operación", vbYesNo + vbQuestion, "Aviso") = vbNo Then
                lbReimp = False
            End If
        Loop
    '**************
    End If
Next i
cmdCancelar_Click
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.path & gsRutaIcono)
CargaTC (gsCodAge)
End Sub
Sub CargaTC(Optional psCodAge As String = "00")
 Dim ObjTcP As COMDConstSistema.DCOMTCEspPermiso
 Dim rs As ADODB.Recordset
' Dim ban As Boolean
 Dim i As Integer
 Set ObjTcP = New COMDConstSistema.DCOMTCEspPermiso
 Set rs = New ADODB.Recordset
 Set rs = ObjTcP.Get_All(gdFecSis, psCodAge)
 Flex.Clear
 Flex.FormaCabecera
 Flex.Rows = 2
 
'ban = False
If Not (rs.EOF And Not rs.BOF) Then
    While Not rs.EOF
'        If ban Then
'            Flex.Rows = Flex.Rows + 1
'        End If
'        ban = True
        Flex.AdicionaFila
        Flex.TextMatrix(Flex.Rows - 1, 1) = Format(rs!dFechaReg, "DD/MM/YYYY HH:MM:SS AMPM")
        'Flex.TextMatrix(Flex.Rows - 1, 3) = ""
        Flex.TextMatrix(Flex.Rows - 1, 3) = rs!cOpedesc
        Flex.TextMatrix(Flex.Rows - 1, 4) = Format(rs!nTCReg, "#0.00")
        Flex.TextMatrix(Flex.Rows - 1, 5) = Format(0, "#0.00")
        Flex.TextMatrix(Flex.Rows - 1, 6) = Format(rs!nMontoReg, "#0.00")
        Flex.TextMatrix(Flex.Rows - 1, 7) = Right(rs!cMovNro, 4)
        Flex.TextMatrix(Flex.Rows - 1, 8) = rs!nCodAut
        rs.MoveNext
    Wend
    Flex.Col = 5
    For i = 1 To Me.Flex.Rows - 1
        Flex.Row = i
        Flex.CellBackColor = &HFFFFC0
    Next i
    
End If
rs.Close
Set ObjTcP = Nothing
Set rs = Nothing
End Sub
