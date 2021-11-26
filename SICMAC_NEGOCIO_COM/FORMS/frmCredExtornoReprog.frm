VERSION 5.00
Begin VB.Form frmCredExtornoReprog 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Extorno de Reprogramacion"
   ClientHeight    =   4560
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11220
   Icon            =   "frmCredExtornoReprog.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   11220
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdLimpiar 
      Caption         =   "Limpiar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   12
      Top             =   960
      Width           =   1335
   End
   Begin VB.CommandButton cmdExtorno 
      Caption         =   "Extorno"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   7
      Top             =   360
      Width           =   1335
   End
   Begin VB.Frame frDatosMost 
      Caption         =   "Datos a Extornar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   0
      TabIndex        =   6
      Top             =   1560
      Width           =   11175
      Begin SICMACT.FlexEdit feDatosExtornoRepg 
         Height          =   2580
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   10965
         _ExtentX        =   19341
         _ExtentY        =   4551
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Nro Cuenta-Titular-Monedad-Saldo Cap.-Dias a Reprog.-Cuotas a Reprog."
         EncabezadosAnchos=   "400-1700-3800-1200-1200-1250-1300"
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
         ColumnasAEditar =   "X-X-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-L-L-R-R-C"
         FormatosEdit    =   "0-0-1-1-2-3-3"
         CantEntero      =   10
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.Frame frDatosBus 
      Height          =   1335
      Left            =   2880
      TabIndex        =   4
      Top             =   120
      Width           =   5535
      Begin VB.TextBox txtUsu 
         Height          =   285
         Left            =   1080
         MaxLength       =   4
         TabIndex        =   9
         Top             =   460
         Width           =   1095
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "Buscar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   4080
         TabIndex        =   5
         Top             =   400
         Width           =   1095
      End
      Begin SICMACT.ActXCodCta ActXCodCta 
         Height          =   495
         Left            =   250
         TabIndex        =   11
         Top             =   400
         Width           =   3735
         _ExtentX        =   6588
         _ExtentY        =   873
         Texto           =   "Crédito"
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
      End
      Begin VB.Label lblUsu 
         Caption         =   "Usuario:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   280
         TabIndex        =   10
         Top             =   450
         Width           =   735
      End
   End
   Begin VB.Frame fraBuscar 
      Caption         =   "Buscar Por"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2655
      Begin VB.OptionButton OpBuscar 
         Caption         =   "Cliente"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   1335
      End
      Begin VB.OptionButton OpBuscar 
         Caption         =   "Nro Cuenta"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton OpBuscar 
         Caption         =   "Usuario"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCredExtornoReprog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim nValorBus As Integer

Private Sub cmdExtorno_Click()

Dim nGrabaExtorno As Boolean
Dim obGrabarExtorno As COMNCredito.NCOMCredito
Dim rsTipPeriodo As ADODB.Recordset
Dim rsCred As ADODB.Recordset
Dim rsObtCalen As ADODB.Recordset

Dim rsResultadoGuardar As ADODB.Recordset 'PRueba Begin tras

Dim NewTCEA As Double
Dim nTipoPeriodo As Integer
Dim MatCalend As Variant
Dim nCuoPag As Integer
Dim i As Integer

Dim oCredito As COMDCredito.DCOMCredito
    Set oCredito = New COMDCredito.DCOMCredito
    Set obGrabarExtorno = New COMNCredito.NCOMCredito

        MatCalend = ""
        NewTCEA = 0
        nTipoPeriodo = 0
        nCuoPag = 0
        i = 0

If Not validaExtorno Then
    Exit Sub
End If

    Set rsCred = oCredito.RecuperaDatosComunes(feDatosExtornoRepg.TextMatrix(feDatosExtornoRepg.row, 1), False)
    Set rsObtCalen = oCredito.ObtieneCalendario(feDatosExtornoRepg.TextMatrix(feDatosExtornoRepg.row, 1))
    Set rsTipPeriodo = oCredito.IdentificarTipoPeriodo(feDatosExtornoRepg.TextMatrix(feDatosExtornoRepg.row, 1))

        If Not (rsObtCalen.EOF And rsObtCalen.BOF) Then
            nCuoPag = rsObtCalen.RecordCount
        Else
            MsgBox "Hubo error al Extornar el Credito, favor de comunicarce con TI-Desarrolla", vbInformation, "Aviso"
            Exit Sub
        End If

        If Not (rsTipPeriodo.EOF And rsTipPeriodo.BOF) Then
            nTipoPeriodo = rsTipPeriodo!nTpPeriodo
        Else
            MsgBox "Hubo error al Extornar el Credito, favor de comunicarce con TI-Desarrolla", vbInformation, "Aviso"
            Exit Sub
        End If

        ReDim MatCalend(nCuoPag, 11)

        For i = 0 To rsObtCalen.RecordCount - 1
            MatCalend(i, 0) = CDate(Format(rsObtCalen!dVenc, "dd/mm/yyyy")) 'Fecha Cuota
            MatCalend(i, 1) = rsObtCalen!nCuota 'Nr Cuota
            MatCalend(i, 2) = rsObtCalen!nCapital + rsObtCalen!nIntComp + rsObtCalen!nIntGracia + rsObtCalen!nGastos 'Suma Cuota
            rsObtCalen.MoveNext
        Next i
  
NewTCEA = obGrabarExtorno.GeneraTasaCostoEfectivoAnual(CDate(Format(rsCred!dVigencia, "dd/mm/yyyy")), CDbl(rsCred!nMontoCol), MatCalend, CDbl(rsCred!nTasaInteres), feDatosExtornoRepg.TextMatrix(feDatosExtornoRepg.row, 1), nTipoPeriodo) 'Para calcular la TCEA

    If feDatosExtornoRepg.TextMatrix(feDatosExtornoRepg.row, 1) <> "" Then
        
        If MsgBox("Se Procedera a Extornar el Credito: " & feDatosExtornoRepg.TextMatrix(feDatosExtornoRepg.row, 1) & ", Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
                
            Set rsResultadoGuardar = oCredito.ExtornoReprgActualizaMov(feDatosExtornoRepg.TextMatrix(feDatosExtornoRepg.row, 1), gsCodUser, gdFecSis, gsCodAge, NewTCEA) 'PRueba Begin tras
                
        If rsResultadoGuardar!nResultado = 1 Then
            MsgBox "Los Datos se Extornaron Correctamente, favor de verificar los datos del Cronograma Actual", vbInformation, "Aviso"
                     
            GenerarConstancia feDatosExtornoRepg.TextMatrix(feDatosExtornoRepg.row, 1), gsCodUser, gsCodAge 'JHCU 12-09-2020
                   
            If nValorBus = 1 Then
                Call Limpiar
            ElseIf nValorBus = 2 Then
                Call Limpiar
            Else
                Call CmdBuscar_Click
            End If
            
            
        Else
            MsgBox "Hubo error al Extornar el Credito, favor de comunicarce con TI-Desarrolla", vbInformation, "Aviso"
            Exit Sub
        End If
    Else
        MsgBox "Seleccione el Credito.", vbInformation, "Aviso"
        Exit Sub
    End If
    
RSClose rsCred
RSClose rsObtCalen
RSClose rsTipPeriodo
End Sub
Private Sub GenerarConstancia(ByVal pcCtaCod As String, ByVal pcCuser As String, ByVal pcAge As String)

 Dim oDCred As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
Dim nTasaCompAnual As Double
        
       On Error GoTo ErrorGenerarConstancia
  
    
    Set oDCred = New COMDCredito.DCOMCredito
    Set R = oDCred.varConsReversion(pcCtaCod, pcCuser, pcAge)
    Set oDCred = Nothing
        
    If R.State = 1 Then
    
        Set oWord = CreateObject("Word.Application")
        oWord.Visible = False

        Set oDoc = oWord.Documents.Open(App.Path & "\FormatoCarta\reprogramacion\formato.dot")
        
        With oWord.Selection.Find
            .Text = "{FECHA}"
            .Replacement.Text = Format(gdFecSis, "dd/mm/yyyy")
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "{COD}"
            .Replacement.Text = R!cCodSol
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "{DOI}"
            .Replacement.Text = R!cDOI
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "{NAME}"
            .Replacement.Text = R!cNombre
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "{CTA}"
            .Replacement.Text = R!cCtaCod
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "{CELULAR}"
            .Replacement.Text = R!cFono
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "{CORREO}"
            .Replacement.Text = R!cCorreo
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
    
        With oWord.Selection.Find
            .Text = "{CANAL}"
            .Replacement.Text = R!cCanal
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        With oWord.Selection.Find
            .Text = "{USER}"
            .Replacement.Text = R!cUser
            .Forward = True
            .Wrap = wdFindContinue
            .Format = False
            .Execute Replace:=wdReplaceAll
        End With
        oDoc.SaveAs (App.Path & "\spooler\CONSTANCIA_REVERSION" & pcCtaCod & ".doc")
        oDoc.Close
        Set oDoc = Nothing
    
        Set oWord = CreateObject("Word.Application")
            oWord.Visible = True
        Set oDoc = oWord.Documents.Open(App.Path & "\spooler\CONSTANCIA_REVERSION" & pcCtaCod & ".doc")
        Set oDoc = Nothing
        Set oWord = Nothing
        MsgBox "Por favor realizar la impresión de la constancia de Reversión", vbInformation, "Aviso"
       End If
    

       
 Exit Sub
ErrorGenerarConstancia:
   MsgBox "Error al generar la constancia de Reversión", vbInformation, "Aviso"
  Set oDoc = Nothing
  Set oWord = Nothing
End Sub

Private Function validaExtorno() As Boolean
    
    Dim obValidaExtorno As COMDCredito.DCOMCredito
    Dim rsValidaExtornoPrimePago As ADODB.Recordset
    Dim rsValidaExtornoPagoMes As ADODB.Recordset
    Set obValidaExtorno = New COMDCredito.DCOMCredito
    Set rsValidaExtornoPrimePago = obValidaExtorno.ValidaPagoPrimerExtoRep(feDatosExtornoRepg.TextMatrix(feDatosExtornoRepg.row, 1))
    Set rsValidaExtornoPagoMes = obValidaExtorno.ValidaPagoMesExtoRep(feDatosExtornoRepg.TextMatrix(feDatosExtornoRepg.row, 1))
    
        validaExtorno = True
        
    If Not (rsValidaExtornoPagoMes.EOF And rsValidaExtornoPagoMes.BOF) Then
        If rsValidaExtornoPagoMes!nValor = 1 Then
            MsgBox "Solo se podra Extornar en el Mes que se Reprogramo," & Chr(13) & "Fecha que se Reprogramo: " & rsValidaExtornoPagoMes!dPrdEstado & " ", vbInformation, "Aviso"
            validaExtorno = False
            Exit Function
        End If
    End If
    
    If Not (rsValidaExtornoPrimePago.EOF And rsValidaExtornoPrimePago.BOF) Then
        If rsValidaExtornoPrimePago!cCtaCod <> "" Then
            MsgBox "Cliente Realizo su Primer pago de su Crédito Reprogramado, No se podrá Extornar la Reprogramación", vbInformation, "Aviso"
            validaExtorno = False
            Exit Function
        End If
    End If
    
RSClose rsValidaExtornoPagoMes
RSClose rsValidaExtornoPrimePago

End Function

Private Sub Form_Load()
    ActXCodCta.Visible = False
    ActXCodCta.CMAC = "109"
    ActXCodCta.Age = gsCodAge
    cmdExtorno.Enabled = False
End Sub

Private Sub OpBuscar_Click(Index As Integer)
' 0 Busca por Usuario - 1 Busca por Cuenta - 2 Busca por Cliente
    nValorBus = Index

If Index = 0 Then
    ActXCodCta.Visible = False
    lblUsu.Visible = True
    txtUsu.Visible = True
    ActXCodCta.NroCuenta = ""
    
    Call Limpiar
    txtUsu.SetFocus
ElseIf Index = 1 Then
    lblUsu.Visible = False
    txtUsu.Visible = False
    ActXCodCta.Visible = True
    ActXCodCta.CMAC = "109"
    ActXCodCta.Age = gsCodAge
    lblUsu.Visible = False
    
    Call Limpiar
    ActXCodCta.SetFocus
ElseIf Index = 2 Then
    ActXCodCta.Visible = False
    txtUsu.Visible = False
    lblUsu.Visible = False
    ActXCodCta.NroCuenta = ""
    
    Call Limpiar
    cmdBuscar.SetFocus
End If

End Sub

Private Function Valida() As Boolean
    Valida = True
    
If nValorBus = 0 Then

    If txtUsu = "" Then
        MsgBox "Ingrese el Usuario", vbInformation, "Aviso"
        Valida = False
        Exit Function
    End If
    LimpiaFlex feDatosExtornoRepg
ElseIf nValorBus = 1 Then

    If ActXCodCta.Prod = "" Or ActXCodCta.Cuenta = "" Then
        MsgBox "Ingrese el Nro de Cuenta", vbInformation, "Aviso"
        Valida = False
        Exit Function
    End If
    LimpiaFlex feDatosExtornoRepg

End If

End Function

Private Sub CmdBuscar_Click()

Dim obExtornoRepg As COMDCredito.DCOMCredito
Dim rsObtieneEstado As ADODB.Recordset
Dim oPers As COMDpersona.UCOMPersona
Dim i As Integer
Dim nCodPers As String

Set obExtornoRepg = New COMDCredito.DCOMCredito

If Not Valida Then
    Exit Sub
End If

If nValorBus = 2 Then
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Call Limpiar
        nCodPers = oPers.sPersCod
    Else
        MsgBox "No Existe Datos ", vbInformation, "Aviso"
        Exit Sub
    End If
End If

Set rsObtieneEstado = obExtornoRepg.ObtenerEstadoExtoRepg(nValorBus, txtUsu.Text, ActXCodCta.NroCuenta, nCodPers)

If Not (rsObtieneEstado.EOF And rsObtieneEstado.BOF) Then
    
    If rsObtieneEstado!nPrdEstado = 205 Or rsObtieneEstado!nPrdEstado = 206 Then
        For i = 1 To rsObtieneEstado.RecordCount
            feDatosExtornoRepg.AdicionaFila
            feDatosExtornoRepg.TextMatrix(i, 1) = rsObtieneEstado!cCtaCod
            feDatosExtornoRepg.TextMatrix(i, 2) = rsObtieneEstado!cPersNombre
            feDatosExtornoRepg.TextMatrix(i, 3) = rsObtieneEstado!cMoneda
            feDatosExtornoRepg.TextMatrix(i, 4) = rsObtieneEstado!nSaldoCap
            feDatosExtornoRepg.TextMatrix(i, 5) = rsObtieneEstado!nDiasRepg
            feDatosExtornoRepg.TextMatrix(i, 6) = rsObtieneEstado!nCuotasReprog
            
            rsObtieneEstado.MoveNext
            
        Next i
        
    cmdExtorno.Enabled = True
    Else
        MsgBox "El Crédito se encuentra en Estado: " & rsObtieneEstado!cDesEstado & "", vbInformation, "Aviso"
    End If
Else
    MsgBox "No Existe Datos ", vbInformation, "Aviso"
End If

RSClose rsObtieneEstado

End Sub

Private Sub CmdLimpiar_Click()
    Call Limpiar
End Sub

Private Sub Limpiar()
    txtUsu.Text = ""
    ActXCodCta.Prod = ""
    ActXCodCta.Cuenta = ""
    cmdExtorno.Enabled = False
    
    LimpiaFlex feDatosExtornoRepg
    
End Sub

Private Sub txtUsu_Change()
    If IsNumeric(txtUsu.Text) Then
        txtUsu.Text = ""
    End If
End Sub

Private Sub txtUsu_KeyPress(KeyAscii As Integer)
    KeyAscii = fgIntfMayusculas(KeyAscii)
    If KeyAscii = 13 Then
        cmdBuscar.SetFocus
    End If
End Sub

Private Sub ActXCodCta_KeyPress(KeyAscii As Integer)
    cmdBuscar.SetFocus
End Sub

Private Sub txtUsu_LostFocus()
    txtUsu.Text = UCase(txtUsu.Text)
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
    If KeyCode = 113 And Shift = 0 Then
        KeyCode = 10
    End If
    If KeyCode = 27 And Shift = 0 Then
        Unload Me
    End If
End Sub

