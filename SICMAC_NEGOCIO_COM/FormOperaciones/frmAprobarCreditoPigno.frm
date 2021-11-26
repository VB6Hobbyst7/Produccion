VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAprobarCreditoPigno 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aprobar / Rechazar Crèdito Pignoraticio"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11700
   Icon            =   "frmAprobarCreditoPigno.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   11700
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9840
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   8040
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton cmdRechazar 
      Caption         =   "Rechazar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   2280
      TabIndex        =   5
      Top             =   4440
      Width           =   1695
   End
   Begin VB.CommandButton cmdAprobar 
      Caption         =   "Aprobar"
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   4440
      Width           =   1695
   End
   Begin SICMACT.FlexEdit fCredPendientes 
      Height          =   2655
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   11175
      _extentx        =   19711
      _extenty        =   4683
      cols0           =   11
      highlight       =   1
      allowuserresizing=   3
      rowsizingmode   =   1
      encabezadosnombres=   "#-nFlag-Item-Fecha-Nº Crèdito-Cliente-V. Tasaciòn-V. Prestamo-Tot. Pieza-Tot. P. B.-Tot. P. N."
      encabezadosanchos=   "650-0-800-1800-1800-2500-1200-1200-1200-1200-1200"
      font            =   "frmAprobarCreditoPigno.frx":030A
      font            =   "frmAprobarCreditoPigno.frx":0336
      font            =   "frmAprobarCreditoPigno.frx":0362
      font            =   "frmAprobarCreditoPigno.frx":038E
      font            =   "frmAprobarCreditoPigno.frx":03BA
      fontfixed       =   "frmAprobarCreditoPigno.frx":03E6
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      backcolorcontrol=   -2147483643
      lbultimainstancia=   -1
      columnasaeditar =   "X-X-2-X-X-X-X-X-X-X-X"
      listacontroles  =   "0-0-4-0-0-0-0-0-0-0-0"
      encabezadosalineacion=   "L-L-C-C-C-C-C-C-C-C-C"
      formatosedit    =   "0-0-0-0-0-0-0-0-0-0-0"
      textarray0      =   "#"
      lbeditarflex    =   -1
      colwidth0       =   645
      rowheight0      =   300
      forecolorfixed  =   -2147483630
   End
   Begin VB.Frame Frame1 
      Caption         =   "Fecha"
      Height          =   735
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   11175
      Begin VB.CommandButton cmdProcesar 
         Caption         =   "Procesar"
         Height          =   375
         Left            =   9000
         TabIndex        =   1
         Top             =   240
         Width           =   1935
      End
      Begin MSMask.MaskEdBox txtfecha 
         Height          =   345
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Width           =   1395
         _ExtentX        =   2461
         _ExtentY        =   609
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
   End
   Begin VB.Label lblTipoForGarPigno 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   9
      Top             =   1080
      Width           =   3375
   End
   Begin VB.Label Label1 
      Caption         =   "Formalidad de Garantia: "
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
      Left            =   360
      TabIndex        =   8
      Top             =   1080
      Width           =   2175
   End
End
Attribute VB_Name = "frmAprobarCreditoPigno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RS As ADODB.Recordset
Dim oDNiv As COMDColocPig.DCOMColPContrato
Dim nTotFil As Integer
Dim nValorUIT As Currency

Private Sub Form_Load()
    txtfecha.Text = gdFecSis
End Sub

Private Sub cmdProcesar_Click()
Dim oDNiv As COMDCredito.DCOMNivelAprobacion
Set oDNiv = New COMDCredito.DCOMNivelAprobacion

If oDNiv.VerificaUsuarioSiTieneNivel(gsCodUser) Then
     If Not ListarCreditoPendientes Then
        MsgBox "No se encontró datos", vbInformation, "Aviso"
     End If
Else
    MsgBox "Ud. No cuenta con Nivel de Aprobaciòn", vbInformation, "Aviso"
End If

End Sub

Private Sub HabilitarComponentes(ByVal pActivar As Boolean)
    cmdAprobar.Enabled = pActivar
    cmdRechazar.Enabled = pActivar
    cmdCancelar.Enabled = pActivar
    lblTipoForGarPigno.Caption = ""
End Sub

Private Function ListarCreditoPendientes() As Boolean
Set oDNiv = New COMDColocPig.DCOMColPContrato
    Dim i As Integer
    ListarCreditoPendientes = False
    Set RS = oDNiv.RecuperaCreditosPorAprobarPigno(gsCodUser, gsCodAge, txtfecha.Text)
    
    fCredPendientes.Clear
    fCredPendientes.FormaCabecera
    LimpiaFlex fCredPendientes
    nTotFil = 0
    If Not (RS.BOF And RS.EOF) Then
        nTotFil = RS.RecordCount
        For i = 1 To nTotFil Step 1
            fCredPendientes.AdicionaFila
            fCredPendientes.TextMatrix(i, 1) = RS!nFlag
            fCredPendientes.TextMatrix(i, 3) = RS!dFechaSol
            fCredPendientes.TextMatrix(i, 4) = RS!cCtaCod
            fCredPendientes.TextMatrix(i, 5) = RS!cCliente
            fCredPendientes.TextMatrix(i, 6) = RS!nTasacion
            fCredPendientes.TextMatrix(i, 7) = RS!nPrestamo
            fCredPendientes.TextMatrix(i, 8) = RS!nTotPiezas
            fCredPendientes.TextMatrix(i, 9) = RS!nTotOroBruto
            fCredPendientes.TextMatrix(i, 10) = RS!nTotOroNeto
            RS.MoveNext
        Next i
    HabilitarComponentes (True)
    ListarCreditoPendientes = True
    End If
End Function

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    LimpiaFlex fCredPendientes
    HabilitarComponentes (False)
End Sub

Private Sub CmdAprobar_Click()
    Dim i As Integer
    Dim scCtaCod As String
    Dim lsMovNro As String
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    
    If Not ValidarFilaSeleccionada() Then
        MsgBox "Debe seleccionar al menos una solicitud", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Está seguro que desea aprobar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        Set oDNiv = New COMDColocPig.DCOMColPContrato
        
        For i = 1 To fCredPendientes.Rows - 1 Step 1
            If fCredPendientes.TextMatrix(i, 2) = "." Then
                scCtaCod = fCredPendientes.TextMatrix(i, 4)
                oDNiv.AprobarCreditoNivelAprobacionPigno scCtaCod, lsMovNro
            End If
        Next i
        
        MsgBox "Se ha aprobado satisfactoriamente " & Chr(10) & "la solicitud seleccionada", vbInformation, "Aviso"
        ListarCreditoPendientes
        
        If nTotFil < 1 Then
            LimpiaFlex fCredPendientes
            HabilitarComponentes (False)
        End If
        
    End If
End Sub

Private Function ValidarFilaSeleccionada() As Boolean
    Dim i As Integer
    Dim bValidar As Boolean

    bValidar = False
    For i = 1 To fCredPendientes.Rows - 1 Step 1
        If fCredPendientes.TextMatrix(i, 2) = "." Then
            bValidar = True
            Exit For
        End If
    Next i
    ValidarFilaSeleccionada = bValidar
End Function

Private Sub cmdRechazar_Click()
    Dim i As Integer
    Dim scCtaCod As String
    Dim lsMovNro As String
    Dim loContFunct As COMNContabilidad.NCOMContFunciones
    
    If Not ValidarFilaSeleccionada() Then
        MsgBox "Debe seleccionar al menos una solicitud", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If MsgBox("¿Está seguro de Rechazar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Set loContFunct = New COMNContabilidad.NCOMContFunciones
        lsMovNro = loContFunct.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        Set loContFunct = Nothing
        Set oDNiv = New COMDColocPig.DCOMColPContrato
        
        For i = 1 To fCredPendientes.Rows - 1 Step 1
            If fCredPendientes.TextMatrix(i, 2) = "." Then
                scCtaCod = fCredPendientes.TextMatrix(i, 4)
                oDNiv.RechazarCreditoNivelAprobacionPigno scCtaCod, lsMovNro
            End If
        Next i
        
        MsgBox "Se ha Rechazado satisfactoriamente " & Chr(10) & "la solicitud seleccionada", vbInformation, "Aviso"
        ListarCreditoPendientes
        
        If nTotFil < 1 Then
            LimpiaFlex fCredPendientes
            HabilitarComponentes (False)
        End If
        
    End If
End Sub

Private Sub fCredPendientes_Click()
    Dim i As Integer
    Dim Col As Integer
    Dim fila As Integer
    Dim bEstaChekeado As Boolean
    Dim nChekeado As String
    
    Col = fCredPendientes.Col
    fila = fCredPendientes.row
    
    If Col = 2 Then
        nChekeado = fCredPendientes.TextMatrix(fila, Col)
        
        If nChekeado = "." Then
            bEstaChekeado = VerificarSiExisteFilaChekeada(fila)
            If bEstaChekeado = False Then
                fCredPendientes.TextMatrix(fila, Col) = 1
                AsignarTipoFormalidadGarantia (fila)
            Else
                fCredPendientes.TextMatrix(fila, Col) = 0
            End If
        Else
            fCredPendientes.TextMatrix(fila, Col) = 0
            lblTipoForGarPigno.Caption = ""
        End If
    End If
End Sub

Private Function VerificarSiExisteFilaChekeada(ByVal nFila As Integer) As Boolean
    Dim i As Integer
    Dim bValidar As Boolean

    bValidar = False
    For i = 1 To fCredPendientes.Rows - 1 Step 1
        If fCredPendientes.TextMatrix(i, 2) = "." And nFila <> i Then
            bValidar = True
            Exit For
        End If
    Next i
    VerificarSiExisteFilaChekeada = bValidar
End Function

Private Sub AsignarTipoFormalidadGarantia(ByVal nFila As Integer)
    Dim nValorUIT_3 As Currency
    Dim nValorUIT_7 As Currency
    Dim nValorPrestamo As Currency
    
    If nValorUIT = 0 Then
        Set oDNiv = New COMDColocPig.DCOMColPContrato
        nValorUIT = oDNiv.ObtenerUITdelAnio(2001) 'Valor UIT
        If nValorUIT = 0 Then
            MsgBox "No existe UIT vigente", vbInformation, "Aviso"
            Exit Sub
        End If
    End If
    
    nValorUIT_3 = nValorUIT * 3
    nValorUIT_7 = nValorUIT * 7
    nValorPrestamo = fCredPendientes.TextMatrix(nFila, 7)
    
    If nValorPrestamo <= nValorUIT_3 Then
        lblTipoForGarPigno.Caption = "Firma simple"
    ElseIf nValorPrestamo > (nValorUIT_3 + 0.1) And nValorPrestamo <= nValorUIT_7 Then
        lblTipoForGarPigno.Caption = "Firma legalizada"
    ElseIf nValorPrestamo > nValorUIT_7 Then
        lblTipoForGarPigno.Caption = "Inscripciòn ante RRPP"
    End If
End Sub
