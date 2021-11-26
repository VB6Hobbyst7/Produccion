VERSION 5.00
Begin VB.Form frmNivelesAprobacionCVAutorizar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Autorización de Compra/Venta ME"
   ClientHeight    =   3540
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   13485
   Icon            =   "frmNivelesAprobacionCVAutorizar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   13485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAutorizar 
      Caption         =   "Autorizar"
      Height          =   375
      Left            =   9960
      TabIndex        =   3
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton cmdDenegar 
      Caption         =   "Denegar"
      Height          =   375
      Left            =   11160
      TabIndex        =   2
      Top             =   3120
      Width           =   975
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   12240
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin SICMACT.FlexEdit feNivApr 
      Height          =   3015
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13335
      _extentx        =   23521
      _extenty        =   5318
      cols0           =   11
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-Movimiento-Agencia-Usuario-cNivelCod-cOpeCod-Operación-nMonto $-T/C Normal-T/C Solicitado-Tipo"
      encabezadosanchos=   "300-3000-3000-1200-0-0-1600-1600-1200-1300-0"
      font            =   "frmNivelesAprobacionCVAutorizar.frx":030A
      font            =   "frmNivelesAprobacionCVAutorizar.frx":0336
      font            =   "frmNivelesAprobacionCVAutorizar.frx":0362
      font            =   "frmNivelesAprobacionCVAutorizar.frx":038E
      font            =   "frmNivelesAprobacionCVAutorizar.frx":03BA
      fontfixed       =   "frmNivelesAprobacionCVAutorizar.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-X-X-X-X-X-X-X"
      listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-L-L-C-C-C-C-C"
      formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0"
      textarray0      =   "#"
      lbformatocol    =   -1
      colwidth0       =   300
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
End
Attribute VB_Name = "frmNivelesAprobacionCVAutorizar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Sub InicioAutorizarNiveles()
    Me.Caption = "Niveles de Aprobación C/V ME - Consulta"
    CargaDatosNiveles
    feNivApr.TopRow = 1
    feNivApr.row = 1
    Me.Show 1
End Sub

Private Sub cmdAutorizar_Click()
Dim lsMovNro As String
Dim oDN  As COMDCredito.DCOMNivelAprobacion
Set oDN = New COMDCredito.DCOMNivelAprobacion
Dim oGen  As COMNContabilidad.NCOMContFunciones
Set oGen = New COMNContabilidad.NCOMContFunciones
If Len(Trim(feNivApr.TextMatrix(feNivApr.row, 1))) > 0 Then
    lsMovNro = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Call oDN.AprobacionMovCompraVentaDetalle(feNivApr.TextMatrix(feNivApr.row, 1), feNivApr.TextMatrix(feNivApr.row, 5), feNivApr.TextMatrix(feNivApr.row, 4), lsMovNro, gsCodCargo, 1)
    MsgBox "Los datos se guardaron correctamente", vbInformation, "Aviso"
    Call CargaDatosNiveles
End If
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub CargaDatosNiveles()
    Dim lsMovNroFecha As String
    Dim oGen  As COMNContabilidad.NCOMContFunciones
    Set oGen = New COMNContabilidad.NCOMContFunciones

    lsMovNroFecha = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)

    Dim oCont As COMNContabilidad.NCOMContFunciones
    Set oCont = New COMNContabilidad.NCOMContFunciones
    Dim oNiv As COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer
    Set oNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oNiv.ObtenerAprobacionMovCompraVentaPendiente(Mid(lsMovNroFecha, 1, 8), gsCodCargo)
    Set oNiv = Nothing
    Call LimpiaFlex(feNivApr)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feNivApr.AdicionaFila
            lnFila = feNivApr.row
            feNivApr.TextMatrix(lnFila, 1) = rs!cMovNro
            feNivApr.TextMatrix(lnFila, 2) = rs!cAgeDescripcion
            feNivApr.TextMatrix(lnFila, 3) = rs!Usuario
            feNivApr.TextMatrix(lnFila, 4) = rs!cNivelCod '
            feNivApr.TextMatrix(lnFila, 5) = rs!cOpecod '
            feNivApr.TextMatrix(lnFila, 6) = rs!Operacion
            feNivApr.TextMatrix(lnFila, 7) = rs!nMonto
            feNivApr.TextMatrix(lnFila, 8) = rs!nTCNormal
            feNivApr.TextMatrix(lnFila, 9) = rs!nTCSolici
            feNivApr.TextMatrix(lnFila, 10) = rs!cTipoCargo
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdDenegar_Click()
Dim lsMovNro As String
Dim oDN  As COMDCredito.DCOMNivelAprobacion
Set oDN = New COMDCredito.DCOMNivelAprobacion
Dim oGen  As COMNContabilidad.NCOMContFunciones
Set oGen = New COMNContabilidad.NCOMContFunciones
If Len(Trim(feNivApr.TextMatrix(feNivApr.row, 1))) > 0 Then
    If feNivApr.TextMatrix(feNivApr.row, 10) = "O" Then
        If MsgBox("Usted denegará solo su registros,desea continuar?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    Else
        If MsgBox("Usted denegará todos los registros,desea continuar?", vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If
    lsMovNro = oGen.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    Call oDN.AprobacionMovCompraVentaDetalle(feNivApr.TextMatrix(feNivApr.row, 1), feNivApr.TextMatrix(feNivApr.row, 5), feNivApr.TextMatrix(feNivApr.row, 4), lsMovNro, gsCodCargo, IIf(feNivApr.TextMatrix(feNivApr.row, 10) = "O", 2, 3))
    MsgBox "Los datos se guardaron correctamente", vbInformation, "Aviso"
    Call CargaDatosNiveles
End If
End Sub
