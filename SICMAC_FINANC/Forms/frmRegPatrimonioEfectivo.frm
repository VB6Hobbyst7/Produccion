VERSION 5.00
Begin VB.Form frmRegPatrimonioEfectivo 
   Caption         =   "Patrimonio Efectivo"
   ClientHeight    =   2115
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6945
   Icon            =   "frmRegPatrimonioEfectivo.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2115
   ScaleWidth      =   6945
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdEliminarPatrimonio 
      Caption         =   "E&liminar"
      Default         =   -1  'True
      Height          =   350
      Left            =   5760
      TabIndex        =   2
      Top             =   840
      Width           =   1095
   End
   Begin VB.CommandButton cmdNuevoPatrimonio 
      Caption         =   "&Nuevo"
      Height          =   350
      Left            =   5760
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdEditarPatrimonio 
      Caption         =   "E&ditar"
      Height          =   350
      Left            =   5760
      TabIndex        =   1
      Top             =   480
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptarPatrimonio 
      Caption         =   "A&ceptar"
      Height          =   350
      Left            =   5760
      TabIndex        =   3
      Top             =   1200
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelarPatrimonio 
      Caption         =   "Ca&ncelar"
      Height          =   350
      Left            =   5760
      TabIndex        =   4
      Top             =   1560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin Sicmact.FlexEdit FEPatrimonio 
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   120
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   3201
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Nro-Patrimonio-Año-Mes-IdPatrimonio"
      EncabezadosAnchos=   "400-1500-1200-1700-3"
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
      ColumnasAEditar =   "X-1-2-3-X"
      ListaControles  =   "0-0-0-3-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-L-L-R"
      FormatosEdit    =   "0-4-0-0-3"
      TextArray0      =   "Nro"
      lbFlexDuplicados=   0   'False
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmRegPatrimonioEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'***Nombre:         frmRegPatrimonioEfectivo
'***Descripción:    Formulario que permite el registro del
'                   patrimonio efectivo.
'***Creación:       ELRO el 20111222 según Acta 352-2011/TI-D
'************************************************************

Private Enum Accion
gValorDefectoAccion = 0
gNuevoRegistro = 1
gEditarRegistro = 2
gEliminarRegistro = 3
End Enum

Private fnAccion, fnFilaNoEditar As Integer

Private fnPatrimonioEfectivo As Currency
Private fsMes, fsAnio As String

Public Property Get PnPatrimonioEfectivo() As Currency
PnPatrimonioEfectivo = fnPatrimonioEfectivo
End Property

Public Property Let PnPatrimonioEfectivo(ByVal vNewValue As Currency)
fnPatrimonioEfectivo = vNewValue
End Property

Public Property Get PsMes() As String
PsMes = fsMes
End Property

Public Property Let PsMes(ByVal vNewValue As String)
fsMes = vNewValue
End Property

Public Property Get PsAnio() As String
PsAnio = fsAnio
End Property

Public Property Let PsAnio(ByVal vNewValue As String)
 fsAnio = vNewValue
End Property

Private Sub CargarFEPatrimonio()
    Dim oDbalanceCont As DbalanceCont
    Set oDbalanceCont = New DbalanceCont
    Dim rsPatrimonio As ADODB.Recordset
    Set rsPatrimonio = New ADODB.Recordset

    Dim i As Integer
    fnAccion = gValorDefectoAccion
    fnFilaNoEditar = -1

    Set rsPatrimonio = oDbalanceCont.recuperarPatrimonioEfectivo()
   
    Call LimpiaFlex(FEPatrimonio)
    
    FEPatrimonio.lbEditarFlex = True
    
    If Not rsPatrimonio.BOF And Not rsPatrimonio.EOF Then
    i = 1
        Do While Not rsPatrimonio.EOF
            FEPatrimonio.AdicionaFila
            FEPatrimonio.TextMatrix(i, 1) = Format(rsPatrimonio!nSaldo, "#,#0.00")
            FEPatrimonio.TextMatrix(i, 3) = UCase(rsPatrimonio!sMES) & space(50) & rsPatrimonio!nMes
            FEPatrimonio.TextMatrix(i, 2) = rsPatrimonio!nAnio
            FEPatrimonio.TextMatrix(i, 4) = rsPatrimonio!IdPatrimonio
            If i = 1 Then
                PnPatrimonioEfectivo = Format(rsPatrimonio!nSaldo, "#,#0.00")
                PsMes = Trim(Left(FEPatrimonio.TextMatrix(i, 3), 10))
                PsAnio = rsPatrimonio!nAnio
            End If
            i = i + 1
            rsPatrimonio.MoveNext
        Loop
    Else
           MsgBox "No existe Montos de Patrimonios registrados", vbInformation, "Aviso"
    End If
    FEPatrimonio.lbEditarFlex = False
End Sub

Private Sub cargarMeses()
    Dim oDGeneral As DGeneral
    Set oDGeneral = New DGeneral
    Dim rsMeses As ADODB.Recordset
    Set rsMeses = New ADODB.Recordset
    
    Set rsMeses = oDGeneral.GetConstante(1010)
        
    FEPatrimonio.CargaCombo rsMeses
   
    Set rsMeses = Nothing
    Set oDGeneral = Nothing
    
End Sub

Private Function validarDatos() As Boolean

Dim i As Integer


    If FEPatrimonio.TextMatrix(FEPatrimonio.row, 1) = "" Then
        MsgBox "Falta ingresar el monto del Patrimonio Efectivo", vbInformation, "Aviso"
        validarDatos = False
        FEPatrimonio.SetFocus
        Exit Function
    End If


    If FEPatrimonio.Col = 3 Then
        If Trim(FEPatrimonio.TextMatrix(FEPatrimonio.row, 3)) = "" Then
            MsgBox "Falta elegir el mes", vbInformation, "Aviso"
            validarDatos = False
            FEPatrimonio.SetFocus
            Exit Function
        End If
    End If
    
    For i = 1 To CInt(FEPatrimonio.Rows) - 2
        If Trim(FEPatrimonio.TextMatrix(i, 2)) = Trim(FEPatrimonio.TextMatrix(FEPatrimonio.row, 2)) Then
            If i <> FEPatrimonio.row Then
                If Trim(Left(FEPatrimonio.TextMatrix(i, 3), 11)) = Trim(Left(FEPatrimonio.TextMatrix(FEPatrimonio.row, 3), 11)) Then
                    MsgBox "Mes " & Trim(Left(FEPatrimonio.TextMatrix(FEPatrimonio.row, 3), 11)) & " repetido en el año " & Trim(FEPatrimonio.TextMatrix(FEPatrimonio.row, 2)), vbInformation, "Aviso"
                    validarDatos = False
                    FEPatrimonio.SetFocus
                    Exit Function
                End If
            End If
        End If
    Next i
    
    If FEPatrimonio.Col = 2 Then
    
        If Trim(FEPatrimonio.TextMatrix(FEPatrimonio.row, 2)) = "" Then
            MsgBox "Falta ingresar el año", vbInformation, "Aviso"
            validarDatos = False
            FEPatrimonio.SetFocus
            Exit Function
        End If


        If Len(Trim(FEPatrimonio.TextMatrix(FEPatrimonio.row, 2))) <> 4 Then
            MsgBox "Año incorrecto", vbInformation, "Aviso"
            validarDatos = False
            FEPatrimonio.SetFocus
            Exit Function
        End If
    
        If Len(Trim(FEPatrimonio.TextMatrix(FEPatrimonio.row, 2))) = 4 Then
            If Not IsNumeric(FEPatrimonio.TextMatrix(FEPatrimonio.row, 2)) Then
                MsgBox "Ingrese un número", vbInformation, "Aviso"
                validarDatos = False
                FEPatrimonio.SetFocus
                Exit Function
            End If
        End If
    End If


validarDatos = True
End Function


Private Sub cmdAceptarPatrimonio_Click()

    Dim oDbalanceCont As DbalanceCont
    Set oDbalanceCont = New DbalanceCont
    Dim oNContFunciones As NContFunciones
    Set oNContFunciones = New NContFunciones
    Dim lsMovNro As String
    Dim lbConfirmarPatrimonio As Boolean
    
    If validarDatos = False Then Exit Sub

    If fnAccion = gNuevoRegistro Then
    
        Dim i, lnCodAnt, lnCodNue As Integer
        
        bConfirmarPatrimonio = False
        
        lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        For i = 1 To FEPatrimonio.Rows - 1
            If Trim(FEPatrimonio.TextMatrix(i, 4)) = "" Then
                    lbConfirmarPatrimonio = oDbalanceCont.registrarPatrimonioEfectivo(CCur(FEPatrimonio.TextMatrix(i, 1)), _
                                                                                 CInt(Trim(Right(FEPatrimonio.TextMatrix(i, 3), 2))), _
                                                                                 CInt(Trim(FEPatrimonio.TextMatrix(i, 2))), _
                                                                                 lsMovNro) 'NAGL 20170802 Cambio en CInt(Trim(Right(FEPatrimonio.TextMatrix(i, 3), 2)))
                If lbConfirmarPatrimonio Then
                    MsgBox "Se guardaron correctamente los datos", vbInformation, "Aviso"
                    Call cargarMeses
                    Call CargarFEPatrimonio
                Else
                    MsgBox "No se guardo el Monto del Patrimonio ", vbInformation, "Aviso"
                End If
            End If
        Next i
        
    End If

    If fnAccion = gEditarRegistro Then
    
        lbConfirmarPatrimonio = False
        
        lsMovNro = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        
        lbConfirmarPatrimonio = oDbalanceCont.ActualizarPatrimonioEfectivo(CInt(FEPatrimonio.TextMatrix(fnFilaNoEditar, 4)), _
                                                                          CCur(FEPatrimonio.TextMatrix(fnFilaNoEditar, 1)), _
                                                                          CInt(Trim(Right(FEPatrimonio.TextMatrix(fnFilaNoEditar, 3), 2))), _
                                                                          CInt(Trim(FEPatrimonio.TextMatrix(fnFilaNoEditar, 2))), _
                                                                          lsMovNro) 'NAGL Cambió de sMovNro a lsMovNro 20190118
        If lbConfirmarPatrimonio Then
            MsgBox "Se modificaron correctamente los datos", vbInformation, "Aviso"
            Call cargarMeses
            Call CargarFEPatrimonio
        Else
            MsgBox "No se modificó el Monto del Patrimonio", vbInformation, "Aviso"
        End If
    End If
    
    Set oNContFunciones = Nothing
    Set oDDocumento = Nothing
    
    cmdNuevoPatrimonio.Visible = True
    cmdEditarPatrimonio.Visible = True
    cmdEliminarPatrimonio.Visible = True
    cmdAceptarPatrimonio.Visible = False
    cmdCancelarPatrimonio.Visible = False
    FEPatrimonio.lbEditarFlex = False
    fnAccion = gValorDefectoAccion
End Sub

Private Sub cmdCancelarPatrimonio_Click()
Call cargarMeses
Call CargarFEPatrimonio
cmdNuevoPatrimonio.Visible = True
cmdEditarPatrimonio.Visible = True
cmdEliminarPatrimonio.Visible = True
cmdAceptarPatrimonio.Visible = False
cmdCancelarPatrimonio.Visible = False
 fnAccion = gValorDefectoAccion
End Sub

Private Sub cmdEditarPatrimonio_Click()
cmdNuevoPatrimonio.Visible = False
cmdEditarPatrimonio.Visible = False
cmdEliminarPatrimonio.Visible = False
cmdAceptarPatrimonio.Visible = True
cmdCancelarPatrimonio.Visible = True
fnAccion = gEditarRegistro
FEPatrimonio.lbEditarFlex = True
fnFilaNoEditar = FEPatrimonio.row
End Sub

Private Sub cmdEliminarPatrimonio_Click()
Dim oDbalanceCont As DbalanceCont
Set oDbalanceCont = New DbalanceCont
Dim lbConfirmarPatrimonio As Boolean

lbConfirmarPatrimonio = False

If MsgBox("¿Esta seguro que desea eliminar el Patrimonio " & Trim(FEPatrimonio.TextMatrix(FEPatrimonio.row, 1)) & "?", vbYesNo, "Aviso") = vbYes Then
    lbConfirmarPatrimonio = oDbalanceCont.eliminarPatrimonioEfectivo(CInt(Trim(FEPatrimonio.TextMatrix(FEPatrimonio.row, 4))))
    If lbConfirmarPatrimonio = False Then
        MsgBox "No pudo elimar el Patrimonio", vbInformation, "Aviso"
        Exit Sub
    End If
    Call cargarMeses
    Call CargarFEPatrimonio
End If

End Sub



Private Sub FEPatrimonio_OnCellChange(pnRow As Long, pnCol As Long)
    If fnFilaNoEditar > -1 Then
        If validarDatos() = False Then
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_Load()
    Call cargarMeses
    Call CargarFEPatrimonio
End Sub

Private Sub cmdNuevoPatrimonio_Click()
cmdNuevoPatrimonio.Visible = False
cmdEditarPatrimonio.Visible = False
cmdEliminarPatrimonio.Visible = False
cmdAceptarPatrimonio.Visible = True
cmdCancelarPatrimonio.Visible = True
fnAccion = gNuevoRegistro
FEPatrimonio.lbEditarFlex = True
FEPatrimonio.AdicionaFila
fnFilaNoEditar = FEPatrimonio.Rows - 1
End Sub

Private Sub FEPatrimonio_RowColChange()
    Dim oDGeneral As DGeneral
    Set oDGeneral = New DGeneral
    Dim rsMeses As ADODB.Recordset
    Set rsMeses = New ADODB.Recordset
   
    If FEPatrimonio.lbEditarFlex Then
        If fnFilaNoEditar <> -1 Then
            FEPatrimonio.row = fnFilaNoEditar
        End If
        Set rsMeses = oDGeneral.GetConstante(1010)
        Select Case FEPatrimonio.Col
           Case 3
                FEPatrimonio.CargaCombo rsMeses
           
        End Select
    End If
End Sub

