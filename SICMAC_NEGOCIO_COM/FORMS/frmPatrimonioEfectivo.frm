VERSION 5.00
Begin VB.Form frmPatrimonioEfectivo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Patrimonio Efectivo Ajustado"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6750
   Icon            =   "frmPatrimonioEfectivo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   6750
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtParametroRCV_val 
      Height          =   285
      Left            =   3360
      TabIndex        =   7
      Text            =   " "
      Top             =   4560
      Width           =   3255
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   5640
      TabIndex        =   5
      Top             =   4920
      Width           =   990
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   4920
      Width           =   975
   End
   Begin VB.ComboBox cboAno 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      ItemData        =   "frmPatrimonioEfectivo.frx":030A
      Left            =   720
      List            =   "frmPatrimonioEfectivo.frx":030C
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin SICMACT.FlexEdit grdPatrimonioEfectivo 
      Height          =   3885
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   6853
      Cols0           =   5
      HighLight       =   1
      VisiblePopMenu  =   -1  'True
      EncabezadosNombres=   "#-Mes-Valor-Referencia-aux"
      EncabezadosAnchos=   "350-1200-1800-3000-0"
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
      ColumnasAEditar =   "X-X-2-3-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-R-L-C"
      FormatosEdit    =   "0-0-2-0-0"
      CantEntero      =   15
      TextArray0      =   "#"
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   345
      RowHeight0      =   300
   End
   Begin VB.CommandButton cmdEditar 
      Caption         =   "&Editar"
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   4920
      Width           =   975
   End
   Begin VB.Label lblParametroRCV_des 
      Caption         =   "RESERVA DE CRÉDITOS VINCULADOS"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   4560
      Width           =   3135
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Año:"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "frmPatrimonioEfectivo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objPista As COMManejador.Pista
Private Sub Form_Load()
    Call CargarAno
    Call CargaPatrimonioEfectivo
    Call EditionMode(False)
End Sub
Private Sub cboAno_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call CargaPatrimonioEfectivo
    End If
End Sub
Private Sub CmdEditar_Click()
    Call EditionMode(True)
End Sub
Private Sub cmdCancelar_Click()
    Call EditionMode(False)
    
    Call CargaPatrimonioEfectivo
    
    grdPatrimonioEfectivo.SetFocus
End Sub
Private Sub CmdGrabar_Click()
    If MsgBox("¿Desea grabar la información actualizada?", vbQuestion + vbYesNo, "Aviso") = vbNo Then
        Exit Sub
    End If

    Dim i As Integer
    
    Dim oPatrimonioEfectivo  As COMNCredito.NCOMPatrimonioEfectivo
    Set oPatrimonioEfectivo = New COMNCredito.NCOMPatrimonioEfectivo
    
    Dim ano  As Integer
    ano = Conversion.CInt(cboAno.Text)
    
    For i = 1 To grdPatrimonioEfectivo.Rows - 1
        Call oPatrimonioEfectivo.ActualizarPatrimonioEfectivo(ano, i, validarNumero(grdPatrimonioEfectivo.TextMatrix(i, 2)), grdPatrimonioEfectivo.TextMatrix(i, 3))
    Next i
    Set oPatrimonioEfectivo = Nothing
    
    
    Dim oPar As New COMDCredito.DCOMParametro
    Dim ParametroValor As Double
    
    ParametroValor = CDbl(validarNumero(txtParametroRCV_val.Text))
    Call oPar.ModificarParametro("102751", lblParametroRCV_des.Caption, ParametroValor, "Modificado desde Patrimonio Efectivo Ajustado")
        
    Call EditionMode(False)
    Call CargaPatrimonioEfectivo
    
    grdPatrimonioEfectivo.SetFocus
End Sub
Private Sub grdPatrimonioEfectivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cmdGrabar.Visible Then
            CmdEditar_Click
        End If
        
        If grdPatrimonioEfectivo.Col = 2 Then
            
            grdPatrimonioEfectivo.TextMatrix(grdPatrimonioEfectivo.row, 2) = validarNumero(grdPatrimonioEfectivo.TextMatrix(grdPatrimonioEfectivo.row, 2))
            
            If grdPatrimonioEfectivo.TextMatrix(grdPatrimonioEfectivo.row, 2) < 0 Then
                MsgBox "Ingrese solo valores positivos", vbInformation, "Aviso"
            End If
        End If
    End If
End Sub
Private Sub grdPatrimonioEfectivo_OnCellChange(pnRow As Long, pnCol As Long)
    If grdPatrimonioEfectivo.Col = 2 Then

        grdPatrimonioEfectivo.TextMatrix(grdPatrimonioEfectivo.row, 2) = validarNumero(grdPatrimonioEfectivo.TextMatrix(grdPatrimonioEfectivo.row, 2))

        If grdPatrimonioEfectivo.TextMatrix(grdPatrimonioEfectivo.row, 2) < 0 Then
            MsgBox "Ingrese solo valores positivos", vbInformation, "Aviso"
        End If
    End If
End Sub
Private Sub CargaPatrimonioEfectivo()
    Dim rsPar As ADODB.Recordset
    
    Dim oPatrimonioEfectivo  As New COMNCredito.NCOMPatrimonioEfectivo
    
    Dim ano  As Integer
    ano = Conversion.CInt(cboAno.Text)

    Set rsPar = oPatrimonioEfectivo.ListaPatrimonioEfectivo(ano)

    grdPatrimonioEfectivo.Clear
    grdPatrimonioEfectivo.FormaCabecera
    grdPatrimonioEfectivo.Rows = 2
    If Not (rsPar.EOF And rsPar.BOF) Then
        Do While Not rsPar.EOF
            grdPatrimonioEfectivo.AdicionaFila
            grdPatrimonioEfectivo.TextMatrix(grdPatrimonioEfectivo.row, 1) = rsPar!cMes
            grdPatrimonioEfectivo.TextMatrix(grdPatrimonioEfectivo.row, 2) = Format(rsPar!nValor, gsFormatoNumeroView)
            grdPatrimonioEfectivo.TextMatrix(grdPatrimonioEfectivo.row, 3) = rsPar!cReferencia
            rsPar.MoveNext
        Loop
        rsPar.Close
        
        Set rsPar = Nothing
        Set oPatrimonioEfectivo = Nothing
    End If
    
    Dim oPar As New COMDCredito.DCOMParametro
    txtParametroRCV_val.Text = Format(oPar.RecuperaValorParametro(102751), "###," & String(15, "#") & "#0.00") & " "
    
End Sub
Private Sub CargarAno()
    Dim i As Integer
    For i = -5 To 5
        cboAno.AddItem (Format(Now(), "YYYY") + i)
        cboAno.ItemData(cboAno.NewIndex) = i + 5
    Next i
    cboAno.ListIndex = 5
End Sub
Private Sub EditionMode(Estado As Boolean)
    cmdEditar.Enabled = Not Estado
    cmdCancelar.Enabled = Estado
    cmdGrabar.Enabled = Estado
    cboAno.Enabled = Not Estado
        
    txtParametroRCV_val.Enabled = Estado
        
    grdPatrimonioEfectivo.lbEditarFlex = Estado
End Sub
Private Function validarNumero(Cadena As String, Optional pbNegativos As Boolean = False) As String
    Dim cValidar As String
    Dim pos As Integer
    Dim vretorno As String
    Dim punto_pos As Integer
    
    'Inicio Eliminar exceso de puntos
    punto_pos = 0

    For pos = Len(Cadena) To 1 Step -1
        If InStr(".", Mid(Cadena, pos, 1)) <> 0 Then
            If punto_pos <> 0 Then
                Cadena = Mid(Cadena, 1, punto_pos - 1) + Mid(Cadena, punto_pos + 1, Len(Cadena))
            End If
            punto_pos = pos
        End If
    Next pos
    'Fin Eliminar exceso de puntos
    
    If pbNegativos = False Then
        cValidar = "0123456789."
    Else
        cValidar = "0123456789.-"
    End If
    
    vretorno = ""
    
    For pos = 1 To Len(Cadena)
        If InStr(cValidar, Mid(Cadena, pos, 1)) <> 0 Then
            vretorno = vretorno + Mid(Cadena, pos, 1)
        End If
    Next pos
    
    If vretorno = "" Then vretorno = "0"
    
    validarNumero = vretorno
End Function
