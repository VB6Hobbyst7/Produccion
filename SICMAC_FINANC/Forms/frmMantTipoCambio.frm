VERSION 5.00
Begin VB.Form frmMantTipoCambio 
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Tipo Cambio"
   ClientHeight    =   4635
   ClientLeft      =   2550
   ClientTop       =   2250
   ClientWidth     =   11055
   HelpContextID   =   250
   Icon            =   "frmMantTipoCambio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   11055
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9600
      TabIndex        =   2
      Top             =   4125
      Width           =   1275
   End
   Begin VB.Frame fratipocambio 
      Height          =   3945
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   10935
      Begin Sicmact.FlexEdit fgTipoCambio 
         Height          =   3240
         Left            =   90
         TabIndex        =   9
         Top             =   570
         Width           =   10725
         _extentx        =   18918
         _extenty        =   5715
         cols0           =   11
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "-Fecha-Venta-Compra-Venta Esp.-Compra Esp.-Fijo Dia-Fijo Mes-Pond...-Modi-Pond Vta"
         encabezadosanchos=   "0-1100-1100-1100-1200-1200-1100-1100-1100-0-1100"
         font            =   "frmMantTipoCambio.frx":030A
         font            =   "frmMantTipoCambio.frx":0332
         font            =   "frmMantTipoCambio.frx":035A
         font            =   "frmMantTipoCambio.frx":0382
         font            =   "frmMantTipoCambio.frx":03AA
         fontfixed       =   "frmMantTipoCambio.frx":03D2
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-1-2-3-4-5-6-7-8-X-10"
         textstylefixed  =   4
         listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-R-R-R-R-R-R-R-L-R"
         formatosedit    =   "0-5-2-2-2-2-2-2-2-0-2"
         cantdecimales   =   4
         lbbuscaduplicadotext=   -1
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "T I P O   D E  C A M B I O "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   255
         Left            =   2438
         TabIndex        =   4
         Top             =   210
         Width           =   5445
      End
   End
   Begin VB.Frame fraNuevo 
      Height          =   645
      Left            =   30
      TabIndex        =   5
      Top             =   3945
      Width           =   2580
      Begin VB.CommandButton CmdModificar 
         Caption         =   "&Modificar"
         Height          =   360
         Left            =   1275
         TabIndex        =   1
         Top             =   180
         Width           =   1155
      End
      Begin VB.CommandButton CmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   360
         Left            =   135
         TabIndex        =   0
         Top             =   180
         Width           =   1155
      End
   End
   Begin VB.Frame fraGrabar 
      Height          =   645
      Left            =   30
      TabIndex        =   6
      Top             =   3945
      Visible         =   0   'False
      Width           =   2580
      Begin VB.CommandButton CmdGrabar 
         Caption         =   "&Grabar"
         Height          =   360
         Left            =   105
         TabIndex        =   7
         Top             =   195
         Width           =   1170
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   360
         Left            =   1290
         TabIndex        =   8
         Top             =   195
         Width           =   1170
      End
   End
End
Attribute VB_Name = "frmMantTipoCambio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oTipoCambio As nTipoCambio
Dim lbNuevo As Boolean
'ARLO20170208****
Dim objPista As COMManejador.Pista
Dim lsAccion As String
'************


Private Sub cmdCancelar_Click()
    If lbNuevo Then
        fgTipoCambio.EliminaFila fgTipoCambio.row
    End If
    fgTipoCambio.BackColorRow vbWhite, False
    fgTipoCambio.lbEditarFlex = False
    fgTipoCambio.SoloFila = False
    fraGrabar.Visible = False
    fraNuevo.Visible = True
End Sub

Function ValidaInterfaz() As Boolean
    ValidaInterfaz = True
    If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 2)) = 0 Then
       MsgBox "Ingrese Valor de Venta", vbInformation, "Aviso"
       ValidaInterfaz = False
       fgTipoCambio.col = 2
       fgTipoCambio.SetFocus
       Exit Function
    End If
    If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 3)) = 0 Then
        MsgBox "Ingrese Valor de Compra", vbInformation, "Aviso"
        ValidaInterfaz = False
        fgTipoCambio.col = 3
        fgTipoCambio.SetFocus
        Exit Function
    End If
    
    If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 4)) = 0 Then
       MsgBox "Ingrese Valor de Venta Especial", vbInformation, "Aviso"
       ValidaInterfaz = False
       fgTipoCambio.col = 4
       fgTipoCambio.SetFocus
       Exit Function
    End If
    If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 5)) = 0 Then
        MsgBox "Ingrese Valor de Compra Especial", vbInformation, "Aviso"
        ValidaInterfaz = False
        fgTipoCambio.col = 3
        fgTipoCambio.SetFocus
        Exit Function
    End If
    
    If lbNuevo Then
        If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 6)) = 0 Then
           MsgBox "Ingrese Valor Fijo Diario", vbInformation, "Aviso"
           ValidaInterfaz = False
           fgTipoCambio.col = 6
           fgTipoCambio.SetFocus
           Exit Function
        End If
        If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 7)) = 0 Then
           MsgBox "Ingrese Valor Fijo Mensual", vbInformation, "Aviso"
           ValidaInterfaz = False
           fgTipoCambio.col = 7
           fgTipoCambio.SetFocus
           Exit Function
        End If
        If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 8)) = 0 Then
           MsgBox "Ingrese Valor de Tipo de cambio ponderado.", vbInformation, "Aviso"
           ValidaInterfaz = False
           fgTipoCambio.col = 8
           fgTipoCambio.SetFocus
           Exit Function
        End If
    End If
    If CCur(fgTipoCambio.TextMatrix(fgTipoCambio.row, 3)) > CCur(fgTipoCambio.TextMatrix(fgTipoCambio.row, 2)) Then
        MsgBox "El valor de compra no puede ser mayor que el valor de venta", vbInformation, "Aviso"
        ValidaInterfaz = False
        fgTipoCambio.col = 3
        fgTipoCambio.SetFocus
        Exit Function
    End If

    If CCur(fgTipoCambio.TextMatrix(fgTipoCambio.row, 5)) > CCur(fgTipoCambio.TextMatrix(fgTipoCambio.row, 4)) Then
        MsgBox "El valor de compra especial no puede ser mayor que el valor de venta especial ", vbInformation, "Aviso"
        ValidaInterfaz = False
        fgTipoCambio.col = 5
        fgTipoCambio.SetFocus
        Exit Function
    End If
End Function

Private Sub cmdGrabar_Click()
    Dim ldFecha As Date
    Dim lnVenta As Currency
    Dim lnCompra As Currency
    Dim lnVentaEsp As Currency
    Dim lnCompraEsp As Currency
    Dim lnFijoDia As Currency
    Dim lnFijoMensual As Currency
    Dim lnPonderado As Currency
    Dim lnPonderadoVenta As Currency
    
    Dim lsMovUltAct  As String
    On Error GoTo ErrorGrabar
    
    If ValidaInterfaz = False Then Exit Sub
    lnVenta = fgTipoCambio.TextMatrix(fgTipoCambio.row, 2)
    lnCompra = fgTipoCambio.TextMatrix(fgTipoCambio.row, 3)
    lnVentaEsp = fgTipoCambio.TextMatrix(fgTipoCambio.row, 4)
    lnCompraEsp = fgTipoCambio.TextMatrix(fgTipoCambio.row, 5)
    lnFijoDia = fgTipoCambio.TextMatrix(fgTipoCambio.row, 6)
    lnFijoMensual = fgTipoCambio.TextMatrix(fgTipoCambio.row, 7)
    lnPonderado = fgTipoCambio.TextMatrix(fgTipoCambio.row, 8)
    lnPonderadoVenta = fgTipoCambio.TextMatrix(fgTipoCambio.row, 10)
    
    lsMovUltAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
    If MsgBox("Desea grabar la información.?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        If lbNuevo Then
            ldFecha = gdFecSis
            oTipoCambio.GrabaTipoCambio gsFormatoFecha, ldFecha, lnVenta, lnCompra, lnVentaEsp, lnCompraEsp, lnFijoMensual, lnFijoDia, lnPonderado, lsMovUltAct, lnPonderadoVenta, False
        Else
            ldFecha = Format(fgTipoCambio.TextMatrix(fgTipoCambio.row, 9), "dd/mm/yyyy hh:mm:ss AMPM")
            oTipoCambio.ActualizaTipoCambio gsFormatoFecha, ldFecha, lnVenta, lnCompra, lnVentaEsp, lnCompraEsp, lnFijoMensual, lnFijoDia, lnPonderado, lsMovUltAct, lnPonderadoVenta, False
        End If
        lbNuevo = False
        fgTipoCambio.SoloFila = False
        fgTipoCambio.lbEditarFlex = False
        fgTipoCambio.BackColorRow vbWhite, False
        CargaTipoCambio
        fraNuevo.Visible = True
        fraGrabar.Visible = False
        
        ReLiberaGarantiaCreditoDiferenteMoneda ldFecha 'EJVG20151001
    End If
            'ARLO20170208
            Set objPista = New COMManejador.Pista
            If (ldFecha = gdFecSis) Then
            lsAccion = "1"
            Else
            lsAccion = "2"
            End If
            gsOpeCod = LogPistaTipoCambio
            objPista.InsertarPista gsOpeCod, glsMovNro, gsCodPersUser, GetMaquinaUsuario, lsAccion, "Se Grabo tipo Cambio | Venta : " & lnVenta & " | Compra : " & lnCompra & " La Fecha " & ldFecha
            Set objPista = Nothing
            '*******
    Exit Sub
ErrorGrabar:
      MsgBox "Error: [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
End Sub

Private Sub CmdModificar_Click()
    If CDate(fgTipoCambio.TextMatrix(fgTipoCambio.row, 1)) <> gdFecSis Then
        MsgBox "No se pueden modificar tipos de cambios de dias anteriores", vbInformation, "Aviso"
        Exit Sub
    End If
    lbNuevo = False
    fraGrabar.Visible = True
    fraNuevo.Visible = False
    fgTipoCambio.lbEditarFlex = True
    fgTipoCambio.SoloFila = True
    fgTipoCambio.BackColorRow &HC0C000, True
    If Day(CDate(fgTipoCambio.TextMatrix(fgTipoCambio.row, 1))) = 1 Then
        fgTipoCambio.ColumnasAEditar = "X-X-2-3-4-5-6-7-8-9-10"
    Else
        fgTipoCambio.ColumnasAEditar = "X-X-2-3-4-5-6-7-8-X-10"
    End If
    fgTipoCambio.SetFocus
End Sub

Private Sub cmdNuevo_Click()
    Dim ldFechaAnt As Date
    lbNuevo = True
    fraGrabar.Visible = True
    fraNuevo.Visible = False
    If fgTipoCambio.TextMatrix(fgTipoCambio.row, 1) <> "" Then
        ldFechaAnt = fgTipoCambio.TextMatrix(fgTipoCambio.row, 1)
    Else
        ldFechaAnt = gdFecSis
    End If
    fgTipoCambio.AdicionaFila
    fgTipoCambio.row = fgTipoCambio.Rows - 1
    fgTipoCambio.lbEditarFlex = True
    fgTipoCambio.SoloFila = True
    fgTipoCambio.BackColorRow &HC0C000, True
    fgTipoCambio.TextMatrix(fgTipoCambio.row, 1) = gdFecSis
    If ldFechaAnt <> gdFecSis Then
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 2) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCVenta), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 3) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCCompra), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 4) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCVentaEsp), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 5) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCCompraEsp), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 6) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCFijoDia), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 7) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCFijoMes), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 8) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCPonderado), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 10) = "0.000" 'Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCPonderadoVenta), "#,#0.000")
    Else
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 2) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 3) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCCompra), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 4) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCVentaEsp), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 5) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCCompraEsp), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 6) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCFijoDia), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 7) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 8) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCPonderado), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 10) = "0.000"
    End If
    fgTipoCambio.ColumnasAEditar = "X-X-2-3-4-5-6-7-8-X-10"
    If fgTipoCambio.TextMatrix(fgTipoCambio.row, 6) <> 0 And fgTipoCambio.TextMatrix(fgTipoCambio.row, 7) <> 0 Then
        fgTipoCambio.ColumnasAEditar = "X-X-2-3-4-5-6-7-8-X-10"
    End If
    If Day(CDate(fgTipoCambio.TextMatrix(fgTipoCambio.row, 1))) = 1 Then
        fgTipoCambio.ColumnasAEditar = "X-X-2-3-4-5-6-7-8-X-10"
    End If
    If fgTipoCambio.TextMatrix(fgTipoCambio.row, 7) = 0 Then
        MsgBox "No se ha ingresado tipo de Cambio de dias anteriores." & vbCrLf & "Verifique desde que fecha no se ha hecho y realice el ingreso por favor...", vbInformation, "Aviso"
    End If
    fgTipoCambio.col = 2
    fgTipoCambio.SetFocus
    SendKeys "{ENTER}"
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub fgTipoCambio_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Select Case pnCol
        Case 3
            If Val(fgTipoCambio.TextMatrix(pnRow, 3)) > Val(fgTipoCambio.TextMatrix(pnRow, 2)) Then
                MsgBox "Tipo de Cambio Compra debe ser menor que el tipo de cambio venta", vbInformation, "Aviso"
                Cancel = False
            End If
        Case Else
    End Select
End Sub

Private Sub Form_Activate()
    fgTipoCambio.DoColumnSort
    fgTipoCambio.col = 1
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub Form_Load()
    Set oTipoCambio = New nTipoCambio
    
    CentraForm Me
    CargaTipoCambio
End Sub
Sub CargaTipoCambio()
    Dim rs As ADODB.Recordset
    
    Set rs = New ADODB.Recordset
    Set rs = oTipoCambio.CargaTipoCambio
    
    fgTipoCambio.Clear
    fgTipoCambio.FormaCabecera
    fgTipoCambio.Rows = 2
    If Not rs.EOF And Not rs.BOF Then
        Set fgTipoCambio.Recordset = rs
        fgTipoCambio.row = rs.RecordCount - 1
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Me.fraGrabar.Visible Then
        If MsgBox("Desea Salir sin grabar la Información", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            Cancel = True
            Exit Sub
        End If
    End If
    Set oTipoCambio = Nothing
End Sub

'EJVG20151003 ***
Private Function ReLiberaGarantiaCreditoDiferenteMoneda(ByVal pdFecha As Date)
    Dim obj As DMov
    Dim bRequiere As Boolean
    
    On Error GoTo ErrorReLibera
    Set obj = New DMov
    
    bRequiere = obj.RequiereReLiberaGarantiasCreditoDiferenteMoneda(pdFecha)
    If bRequiere Then
        MsgBox "Se va a realizar la Liberación de de las Garantías con los créditos que tengan diferente moneda entre sí." & Chr(13) & Chr(13) & "Espere un momento por favor.", vbInformation, "Aviso"
        Screen.MousePointer = 11
        Call obj.ReLiberaGarantiasCreditoDiferenteMoneda(pdFecha)
        Screen.MousePointer = 0
        MsgBox "Se ha terminado con la ReLiberación de las Garantías", vbInformation, "Aviso"
    End If
    Set obj = Nothing
    Exit Function
ErrorReLibera:
    Screen.MousePointer = 0
    MsgBox Err.Description, vbCritical, "Aviso"
End Function
'END EJVG *******
