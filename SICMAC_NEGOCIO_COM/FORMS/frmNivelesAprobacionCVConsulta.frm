VERSION 5.00
Begin VB.Form frmNivelesAprobacionCVConsulta 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Niveles de Aprobación C/V ME - Consulta"
   ClientHeight    =   7320
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12930
   Icon            =   "frmNivelesAprobacionCVConsulta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7320
   ScaleWidth      =   12930
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   375
      Left            =   11760
      TabIndex        =   0
      Top             =   6840
      Width           =   1095
   End
   Begin SICMACT.FlexEdit feNivApr 
      Height          =   6735
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   12855
      _extentx        =   22675
      _extenty        =   11880
      cols0           =   11
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-Nivel-CodCargos-Cargos-Tipo-Firmas-Ag?-Desde $-Hasta $-TCC Más-TCV Más"
      encabezadosanchos=   "300-1200-0-3000-0-800-800-1600-1600-1200-1200"
      font            =   "frmNivelesAprobacionCVConsulta.frx":030A
      font            =   "frmNivelesAprobacionCVConsulta.frx":0336
      font            =   "frmNivelesAprobacionCVConsulta.frx":0362
      font            =   "frmNivelesAprobacionCVConsulta.frx":038E
      font            =   "frmNivelesAprobacionCVConsulta.frx":03BA
      fontfixed       =   "frmNivelesAprobacionCVConsulta.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-X-X-X-X-X-X-X-X-X"
      listacontroles  =   "0-0-0-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "C-C-C-L-L-C-C-C-C-C-C"
      formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0"
      textarray0      =   "#"
      lbformatocol    =   -1
      colwidth0       =   300
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
End
Attribute VB_Name = "frmNivelesAprobacionCVConsulta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Sub InicioConsultaNiveles()
    Me.Caption = "Niveles de Aprobación C/V ME - Consulta"
    CargaDatosNiveles
    feNivApr.TopRow = 1
    feNivApr.row = 1
    Me.Show 1
End Sub
Private Sub cmdCerrar_Click()
    Unload Me
End Sub
Private Sub CargaDatosNiveles()
    Dim oConst As COMDConstantes.DCOMConstantes
    Set oConst = New COMDConstantes.DCOMConstantes
    Dim oCont As COMNContabilidad.NCOMContFunciones
    Set oCont = New COMNContabilidad.NCOMContFunciones


    Dim oNiv As COMDCredito.DCOMNivelAprobacion
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer
    Set oNiv = New COMDCredito.DCOMNivelAprobacion
    Set rs = oNiv.RecuperaNivAprCV()
    Set oNiv = Nothing
    Call LimpiaFlex(feNivApr)
    If Not rs.EOF Then
        Do While Not rs.EOF
            feNivApr.AdicionaFila
            lnFila = feNivApr.row
            feNivApr.TextMatrix(lnFila, 1) = rs!cNivelCod
            feNivApr.TextMatrix(lnFila, 2) = rs!cRHCargos
            feNivApr.TextMatrix(lnFila, 3) = rs!cRHCargoDescripcion
            feNivApr.TextMatrix(lnFila, 4) = rs!cTipoCargoDesc
            feNivApr.TextMatrix(lnFila, 5) = rs!nNroFirmas
            feNivApr.TextMatrix(lnFila, 6) = IIf(rs!bValidaAgencia = True, "Si", "No")
            feNivApr.TextMatrix(lnFila, 7) = rs!nMontoDesde
            feNivApr.TextMatrix(lnFila, 8) = rs!nMontoHasta
            feNivApr.TextMatrix(lnFila, 9) = rs!nTCCmas
            feNivApr.TextMatrix(lnFila, 10) = rs!nTCVmas
            rs.MoveNext
        Loop
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub feNivApr_DblClick()
    If feNivApr.TextMatrix(feNivApr.row, feNivApr.Col) <> "" Then
        MuestraDatosGridNivApr
    End If
End Sub
Private Sub MuestraDatosGridNivApr()
    Dim oLista As COMDCredito.DCOMNivelAprobacion
    Dim rsDatos As ADODB.Recordset
    Dim MatTitulos() As String
    ReDim MatTitulos(1, 2)
    MatTitulos(0, 0) = feNivApr.TextMatrix(feNivApr.row, 4)
    MatTitulos(0, 1) = "Tipo"
    If feNivApr.Col = 1 Then
        Set oLista = New COMDCredito.DCOMNivelAprobacion
            Set rsDatos = oLista.RecuperaNivAprValoresCV(feNivApr.TextMatrix(feNivApr.row, 1))
        Set oLista = Nothing
        frmCredListaDatos.Inicio feNivApr.TextMatrix(feNivApr.row, 4), rsDatos, , 2, MatTitulos
    End If
End Sub
