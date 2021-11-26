VERSION 5.00
Begin VB.Form frmCredNewNivAprAprobar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobar"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4545
   Icon            =   "frmCredNewNivAprAprobar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   4545
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SICMACT.FlexEdit feDatosCred 
      Height          =   1275
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4305
      _extentx        =   7594
      _extenty        =   2249
      cols0           =   3
      fixedcols       =   2
      highlight       =   1
      encabezadosnombres=   "Concepto-Valor Actual-Nuevo Valor"
      encabezadosanchos=   "1800-1200-1200"
      font            =   "frmCredNewNivAprAprobar.frx":030A
      font            =   "frmCredNewNivAprAprobar.frx":0336
      font            =   "frmCredNewNivAprAprobar.frx":0362
      font            =   "frmCredNewNivAprAprobar.frx":038E
      font            =   "frmCredNewNivAprAprobar.frx":03BA
      fontfixed       =   "frmCredNewNivAprAprobar.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      tipobusqueda    =   3
      columnasaeditar =   "X-X-2"
      listacontroles  =   "0-0-0"
      encabezadosalineacion=   "C-R-R"
      formatosedit    =   "0-2-2"
      cantdecimales   =   4
      textarray0      =   "Concepto"
      lbeditarflex    =   -1
      lbbuscaduplicadotext=   -1
      colwidth0       =   1800
      rowheight0      =   300
      cellbackcolor   =   -2147483633
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      TabIndex        =   3
      Top             =   960
      Width           =   1170
   End
   Begin VB.CommandButton cmdHistorial 
      Caption         =   "Historial de Aprobaciones"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2445
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "Aprobar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   1170
   End
End
Attribute VB_Name = "frmCredNewNivAprAprobar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredNewNivAprAprobar
'** Descripción : Formulario para aprobar los Creditos por nivel de aprobacion creado segun RFC110-2012
'** Creación : JUEZ, 20121211 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim oNNiv As COMNCredito.NCOMNivelAprobacion
Dim rs As ADODB.Recordset
Dim fsCtaCod As String, fsNivAprCod As String
Dim fsLineaCred As String, fsComent As String
Dim i As Integer
Dim fbAprobar As Boolean 'JUEZ 20160218

Public Function AprobarCredito(ByVal psCtaCod As String, ByVal psNivAprCod As String, ByVal pnCuotas As Double, ByVal pnTasa As Double, _
                               ByVal pnMonto As Double, ByVal psLineaCred As String, ByVal psComent As String) As Boolean 'JUEZ 20160218
    fbAprobar = False 'JUEZ 20160218
    fsCtaCod = psCtaCod
    fsNivAprCod = psNivAprCod
    fsComent = psComent
    fsLineaCred = psLineaCred
    Call LimpiaFlex(feDatosCred)
    For i = 1 To 3
        feDatosCred.AdicionaFila
        feDatosCred.TextMatrix(i, 0) = Choose(i, "Cuotas", "Tasa", "Monto")
        feDatosCred.TextMatrix(i, 1) = Choose(i, pnCuotas, pnTasa, pnMonto)
        feDatosCred.TextMatrix(i, 2) = Choose(i, pnCuotas, pnTasa, pnMonto)
    Next i
    feDatosCred.TextMatrix(1, 1) = Format(pnCuotas, "#0")
    feDatosCred.TextMatrix(2, 1) = Format(pnTasa, "#,##0.0000")
    feDatosCred.TextMatrix(3, 1) = Format(pnMonto, "#,##0.00")
    feDatosCred.TextMatrix(1, 2) = Format(pnCuotas, "#0")
    feDatosCred.TextMatrix(2, 2) = Format(pnTasa, "#,##0.0000")
    feDatosCred.TextMatrix(3, 2) = Format(pnMonto, "#,##0.00")
    Me.Show 1
    AprobarCredito = fbAprobar 'JUEZ 20160218
End Function

Private Sub cmdAprobar_Click()
    If ValidaLineaCredito Then
        Dim Aprobar As String
        If Not (ValidarTasaMaxima(fsCtaCod, CDbl(feDatosCred.TextMatrix(2, 2)), , , CCur(feDatosCred.TextMatrix(3, 2)))) Then Exit Sub 'FRHU 20170914 ERS049-2017
        If MsgBox("Se va a Aprobar el Credito, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
        Set oNNiv = New COMNCredito.NCOMNivelAprobacion
        Aprobar = oNNiv.AprobacionCreditoPorNivel(fsCtaCod, fsNivAprCod, CInt(feDatosCred.TextMatrix(1, 2)), CDbl(feDatosCred.TextMatrix(2, 2)), _
                                                  CDbl(feDatosCred.TextMatrix(3, 2)), fsComent, gdFecSis, gsCodAge, gsCodUser)
        If Aprobar <> "" Then
            MsgBox Aprobar, vbInformation, "Aviso"
        Else
            'RECO20161020 ERS060-2016 **********************************************************
            Dim oNCOMColocEval As New NCOMColocEval
            Dim lcMovNro As String
            
            If Not ValidaExisteRegProceso(fsCtaCod, gTpoRegCtrlNivAprobacion) Then
               lcMovNro = GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
               Call oNCOMColocEval.insEstadosExpediente(fsCtaCod, "Nivel de Aprobacion de Crédito", lcMovNro, "", "", "", 1, 2002, gTpoRegCtrlNivAprobacion)
               Set oNCOMColocEval = Nothing
            End If
            'RECO FIN **************************************************************************
            MsgBox "El credito ha sido aprobado por su nivel, Puede verificarlo en el Historial", vbInformation, "Aviso"
            fbAprobar = True 'JUEZ 20160218
            Unload Me
        End If
    End If
End Sub

Private Sub cmdHistorial_Click()
    Call frmCredNewNivAprHist.InicioCredito(fsCtaCod)
End Sub

Private Function ValidaLineaCredito() As Boolean
'ALPA 20150213*************************************
'    Dim oLineas As COMDCredito.DCOMLineaCredito
'    Dim RLinea As ADODB.Recordset
'    Set oLineas = New COMDCredito.DCOMLineaCredito
'    Set RLinea = oLineas.RecuperaLineadeCredito(fsLineaCred)
'
'    'JUEZ 20150105 ***************************************
'    If RLinea.EOF Or RLinea.BOF Then
'        MsgBox "La Linea de Credito seleccionada en la sugerencia no está disponible. Verificar", vbInformation, "Aviso"
'        ValidaLineaCredito = False
'        Exit Function
'    End If
'    'END JUEZ ********************************************
'
'    ValidaLineaCredito = True
'
'    'Valida Tasa Interes Comp.
'    If CDbl(feDatosCred.TextMatrix(2, 2)) < RLinea!nTasaIni Or CDbl(feDatosCred.TextMatrix(2, 2)) > RLinea!nTasafin Then
'        MsgBox "La Tasa de Interes No es Permitida por la Linea de Credito", vbInformation, "Aviso"
'        ValidaLineaCredito = False
'        Exit Function
'    End If
'
'    'Valida Monto Sugerido
'    If CDbl(feDatosCred.TextMatrix(3, 2)) < RLinea!nMontoMin Or CDbl(feDatosCred.TextMatrix(3, 2)) > RLinea!nMontoMax Then
'        MsgBox "El Monto del Credito, No es Permitido por la Linea de Credito", vbInformation, "Aviso"
'        ValidaLineaCredito = False
'        Exit Function
'    End If
'    Set RLinea = Nothing
    ValidaLineaCredito = True
End Function

Private Sub feDatosCred_OnCellChange(pnRow As Long, pnCol As Long)
    'If pnCol = 3 Then
        If pnRow = 1 Then
            feDatosCred.TextMatrix(pnRow, pnCol) = Format(feDatosCred.TextMatrix(pnRow, pnCol), "#0")
        ElseIf pnRow = 2 Then
            feDatosCred.TextMatrix(pnRow, pnCol) = Format(feDatosCred.TextMatrix(pnRow, pnCol), "#,##0.0000")
        ElseIf pnRow = 3 Then
            feDatosCred.TextMatrix(pnRow, pnCol) = Format(feDatosCred.TextMatrix(pnRow, pnCol), "#,##0.00")
        End If
    'End If
End Sub
