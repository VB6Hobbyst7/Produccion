VERSION 5.00
Begin VB.Form frmMantTipoCambio 
   BackColor       =   &H80000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento Tipo Cambio"
   ClientHeight    =   4635
   ClientLeft      =   2550
   ClientTop       =   2250
   ClientWidth     =   10395
   HelpContextID   =   250
   Icon            =   "frmMantTipoCambio.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4635
   ScaleWidth      =   10395
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9060
      TabIndex        =   2
      Top             =   4125
      Width           =   1275
   End
   Begin VB.Frame fratipocambio 
      Height          =   3945
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Width           =   10335
      Begin SICMACT.FlexEdit fgTipoCambio 
         Height          =   3240
         Left            =   90
         TabIndex        =   9
         Top             =   570
         Width           =   10140
         _ExtentX        =   14790
         _ExtentY        =   5715
         Cols0           =   16
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   $"frmMantTipoCambio.frx":030A
         EncabezadosAnchos=   "0-1200-1200-1200-1200-1200-1200-1200-1200-0-1200-1200-1200-1200-1200-0"
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
         ColumnasAEditar =   "X-1-2-3-4-5-6-7-8-X-10-X-X-X-X-X"
         TextStyleFixed  =   4
         ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-R-R-R-R-R-R-R-L-C-R-R-R-R-C"
         FormatosEdit    =   "0-5-2-2-2-2-2-2-2-0-2-2-0-0-0-0"
         CantDecimales   =   4
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
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
      Width           =   3780
      Begin VB.CommandButton CmdTipEsp 
         Caption         =   "TC Especial"
         Height          =   360
         Left            =   2520
         TabIndex        =   10
         Top             =   180
         Width           =   1155
      End
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
Dim oTipoCambio As COMDConstSistema.NCOMTipoCambio
Dim lbNuevo As Boolean

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
'ALPA 20081003*******************************************************************************
'Se agrego el parametro pnPonREU para el REU
'********************************************************************************************
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
    
    If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 10)) = 0 Then
           MsgBox "Ingrese Valor de Tipo de cambio ponderado venta.", vbInformation, "Aviso"
           ValidaInterfaz = False
           fgTipoCambio.col = 10
           fgTipoCambio.SetFocus
           Exit Function
    End If
    If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 11)) = 0 Then
           MsgBox "Ingrese Valor para el REU", vbInformation, "Aviso"
           ValidaInterfaz = False
           fgTipoCambio.col = 11
           fgTipoCambio.SetFocus
           Exit Function
    End If
    'ALPA20140225************************************************************************************
    If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 12)) = 0 Then
           MsgBox "Ingrese Valor para el tipo de cambio SBS día", vbInformation, "Aviso"
           ValidaInterfaz = False
           fgTipoCambio.col = 12
           fgTipoCambio.SetFocus
           Exit Function
    End If
    If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 13)) = 0 Then
           MsgBox "Ingrese Valor para el tipo de cambio Compra traider", vbInformation, "Aviso"
           ValidaInterfaz = False
           fgTipoCambio.col = 13
           fgTipoCambio.SetFocus
           Exit Function
    End If
    If Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 14)) = 0 Then
           MsgBox "Ingrese Valor para el tipo de cambio venta traider", vbInformation, "Aviso"
           ValidaInterfaz = False
           fgTipoCambio.col = 14
           fgTipoCambio.SetFocus
           Exit Function
    End If
    '*************************************************************************************************
End Function

Private Sub CmdGrabar_Click()
    Dim ldFecha As Date
    Dim lnVenta As Currency
    Dim lnCompra As Currency
    Dim lnVentaEsp As Currency
    Dim lnCompraEsp As Currency
    Dim lnFijoDia As Currency
    Dim lnFijoMensual As Currency
    Dim lnPonderado As Currency
    Dim lnPondVenta As Currency
    Dim lnPonREU As Currency
    
    'ALPA20140225***************************
    Dim lnSBSDia As Currency
    Dim lnComprTraider As Currency
    Dim lnVentaTraider As Currency
    '***************************************
    
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
    lnPondVenta = fgTipoCambio.TextMatrix(fgTipoCambio.row, 10)
    lnPonREU = Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 11))
    'ALPA20120225***********************************************************
    lnSBSDia = fgTipoCambio.TextMatrix(fgTipoCambio.row, 12)
    lnComprTraider = fgTipoCambio.TextMatrix(fgTipoCambio.row, 13)
    lnVentaTraider = Val(fgTipoCambio.TextMatrix(fgTipoCambio.row, 14))
    '***********************************************************************
    Set oTipoCambio = New COMDConstSistema.NCOMTipoCambio
    
    lsMovUltAct = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
    If MsgBox("Desea grabar la información.?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        If lbNuevo Then
            ldFecha = gdFecSis
            '*** PEAC 20130402
            'oTipoCambio.GrabaTipoCambio gsFormatoFecha, ldFecha, lnVenta, lnCompra, lnVentaEsp, lnCompraEsp, lnFijoMensual, lnFijoDia, lnPonderado, lsMovUltAct, False, lnPondVenta, lnPonREU
            GrabaTC gsFormatoFecha, ldFecha, lnVenta, lnCompra, lnVentaEsp, lnCompraEsp, lnFijoMensual, lnFijoDia, lnPonderado, lsMovUltAct, False, lnPondVenta, lnPonREU, lnSBSDia, lnComprTraider, lnVentaTraider
        Else
            ldFecha = Format(fgTipoCambio.TextMatrix(fgTipoCambio.row, 9), "dd/mm/yyyy hh:mm:ss AMPM")
            '*** PEAC 20130402
            'oTipoCambio.ActualizaTipoCambio gsFormatoFecha, ldFecha, lnVenta, lnCompra, lnVentaEsp, lnCompraEsp, lnFijoMensual, lnFijoDia, lnPonderado, lsMovUltAct, False, lnPondVenta, lnPonREU
            ActualizaTC gsFormatoFecha, ldFecha, lnVenta, lnCompra, lnVentaEsp, lnCompraEsp, lnFijoMensual, lnFijoDia, lnPonderado, lsMovUltAct, False, lnPondVenta, lnPonREU, lnSBSDia, lnComprTraider, lnVentaTraider
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
    Exit Sub
ErrorGrabar:
      MsgBox "Error: [" & Str(err.Number) & "] " & err.Description, vbInformation, "Aviso"
End Sub

'*** PEAC 20130402
Private Sub ActualizaTC(ByVal psFormatoFecha As String, ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
                                    ByVal pnValVentEnt As Currency, ByVal pnValCompEst As Currency, _
                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, ByVal psUltimaActualizacion As String, _
                                    Optional ByVal pbEjectBatch As Boolean, Optional ByVal pnValPondVenta As Currency = 0, Optional ByVal pnPonREU As Currency = 0, Optional ByVal pnSBSDia As Currency = 0, Optional ByVal pnCompraTraider As Currency = 0, Optional ByVal pnVentaTraider As Currency = 0)
                                    
Dim oTipoCambio As COMDConstSistema.DCOMTipoCambio
Set oTipoCambio = New COMDConstSistema.DCOMTipoCambio

oTipoCambio.Inicio psFormatoFecha
oTipoCambio.ActualizaTipoCambio pdFecCamb, psUltimaActualizacion, pnValVent, pnValComp, pnValVentEnt, pnValCompEst, pnValFijo, pnValFijoDia, pnValPonderado, pbEjectBatch, pnValPondVenta, pnPonREU, pnSBSDia, pnCompraTraider, pnVentaTraider
Set oTipoCambio = Nothing

End Sub
'*** PEAC 20130402
Private Sub GrabaTC(ByVal psFormatoFecha As String, ByVal pdFecCamb As Date, ByVal pnValVent As Currency, ByVal pnValComp As Currency, _
                                    ByVal pnValVentEsp As Currency, ByVal pnValCompEsp As Currency, _
                                    ByVal pnValFijo As Currency, ByVal pnValFijoDia As Currency, ByVal pnValPonderado As Currency, _
                                    ByVal psUltimaActualizacion As String, _
                                    Optional ByVal pbEjectBatch As Boolean, Optional ByVal pnPondVenta As Currency = 0, Optional ByVal pnPonREU As Currency = 0, Optional ByVal pnSBSDia As Currency = 0, Optional ByVal pnCompraTraider As Currency = 0, Optional ByVal pnVentaTraider As Currency = 0)

Dim oTipoCambio As COMDConstSistema.DCOMTipoCambio
Set oTipoCambio = New COMDConstSistema.DCOMTipoCambio

oTipoCambio.Inicio psFormatoFecha
oTipoCambio.ActualizaTipoCambioDiario pdFecCamb, psUltimaActualizacion, pnValFijoDia, pbEjectBatch
oTipoCambio.InsertaTipoCambio pdFecCamb, pnValVent, pnValComp, pnValVentEsp, pnValCompEsp, pnValFijo, pnValFijoDia, pnValPonderado, psUltimaActualizacion, pbEjectBatch, pnPondVenta, pnPonREU, pnSBSDia, pnCompraTraider, pnVentaTraider

Set oTipoCambio = Nothing

End Sub

Private Sub cmdModificar_Click()
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
    'ALPA20140225***********************************************************
    'Se agregó 12-13-14
    If Day(CDate(fgTipoCambio.TextMatrix(fgTipoCambio.row, 1))) = 1 Then
        fgTipoCambio.ColumnasAEditar = "X-X-2-3-4-5-6-7-8-9-10-11-12-13-14"
    Else
        fgTipoCambio.ColumnasAEditar = "X-X-2-3-4-5-6-7-8-X-10-11-12-13-14"
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
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 10) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCPondVenta), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 11) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCPondREU), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 12) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCSBSDia), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 13) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCCompraTraider), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 14) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis - 1, TCVentaTraider), "#,#0.000")
    Else
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 2) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCVenta), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 3) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCCompra), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 4) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCVentaEsp), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 5) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCCompraEsp), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 6) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCFijoDia), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 7) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCFijoMes), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 8) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCPonderado), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 10) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCPondVenta), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 11) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCPondREU), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 12) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCSBSDia), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 13) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCCompraTraider), "#,#0.000")
        fgTipoCambio.TextMatrix(fgTipoCambio.row, 14) = Format(oTipoCambio.EmiteTipoCambio(gdFecSis, TCVentaTraider), "#,#0.000")
    End If
    fgTipoCambio.ColumnasAEditar = "X-X-2-3-4-5-6-7-8-X-10-11-12-13-14"
    If fgTipoCambio.TextMatrix(fgTipoCambio.row, 6) <> 0 And fgTipoCambio.TextMatrix(fgTipoCambio.row, 7) <> 0 Then
        fgTipoCambio.ColumnasAEditar = "X-X-2-3-4-5-6-7-8-X-10-11-12-13-14"
    End If
    If Day(CDate(fgTipoCambio.TextMatrix(fgTipoCambio.row, 1))) = 1 Then
        fgTipoCambio.ColumnasAEditar = "X-X-2-3-4-5-6-7-8-X-10-11-12-13-14"
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

Private Sub CmdTipEsp_Click()
FrmMantTipoCambioNuevo.Show 1
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
    Me.Icon = LoadPicture(App.Path & gsRutaIcono)
End Sub

Private Sub Form_Load()
    Set oTipoCambio = New COMDConstSistema.NCOMTipoCambio
    
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
    Dim obj As DCOMCredActBD
    Dim bRequiere As Boolean
    
    On Error GoTo ErrorReLibera
    Set obj = New DCOMCredActBD
    
    bRequiere = obj.RequiereReLiberaGarantiasCreditoDiferenteMoneda(pdFecha)
    If bRequiere Then
        MsgBox "Se va a realizar la Liberación de de las Garantías con los créditos que tengan diferente moneda entre sí." & Chr(13) & Chr(13) & "Espere un momento por favor..", vbInformation, "Aviso"
        Screen.MousePointer = 11
        Call obj.ReLiberaGarantiasCreditoDiferenteMoneda(pdFecha)
        Screen.MousePointer = 0
        MsgBox "Se ha terminado con la ReLiberación de las Garantías", vbInformation, "Aviso"
    End If
    Set obj = Nothing
    Exit Function
ErrorReLibera:
    Screen.MousePointer = 0
    MsgBox err.Description, vbCritical, "Aviso"
End Function
'END EJVG *******
