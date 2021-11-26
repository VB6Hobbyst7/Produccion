VERSION 5.00
Begin VB.Form frmOperacionesInterCMACInicio 
   Caption         =   "Operaciones Inter CMAC's"
   ClientHeight    =   3375
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7215
   LinkTopic       =   "Form1"
   ScaleHeight     =   3375
   ScaleWidth      =   7215
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir "
      Height          =   375
      Left            =   4920
      TabIndex        =   2
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton btnRecepcion 
      Caption         =   "Operaciones Inter CMAC's Recepción"
      Height          =   615
      Left            =   480
      TabIndex        =   1
      Top             =   1080
      Width           =   5535
   End
   Begin VB.CommandButton btnLlamada 
      Caption         =   "Operaciones Inter CMAC's Llamada"
      Height          =   615
      Left            =   480
      TabIndex        =   0
      Top             =   360
      Width           =   5535
   End
End
Attribute VB_Name = "frmOperacionesInterCMACInicio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsOpeLLama As ADODB.Recordset
Dim rsOpeRecep As ADODB.Recordset
Dim rsCMACs As ADODB.Recordset

'Dim oconecta As DConecta


Private Sub btnLlamada_Click()
    'Obtiene la impresora predeterminada
    Dim sImpresora As String
    Dim lnPos As Long
    sImpresora = Printer.DeviceName
    If Left(sImpresora, 2) <> "\\" Then
        lnPos = InStr(1, Printer.Port, ":", vbTextCompare)
        If lnPos > 0 Then
            sLpt = Mid(Printer.Port, 1, lnPos - 1)
        Else
            sLpt = "LPT1"
        End If
    Else
        sLpt = frmImpresora.EliminaEspacios(sImpresora)
    End If
    
    '    DeshabilitaOpeacionesPendientes

    MsgBox "Por favor Configure su Impresora antes de Empezar sus operaciones", vbInformation, "Aviso"
    frmImpresora.Show 1
    frmOperacionesInterCMAC.Inicia "Cajero - Operaciones CMACs Llamada", rsOpeLLama, rsCMACs
End Sub

Private Sub btnRecepcion_Click()
'    RetiroInterCMAC dFecSis, sCtaCod, nMonto, 0, nmoneda, sOpeCod, sOpeCodComision, nTipoCambio, slDTrama, _
'                    nOFFHost, PAN, Hora, MesDia, MontoEquiv
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    Call FuncionesIni
    
    'CargaVarSistema
    'fgITFParametros
    'obtenerRSOperaciones
End Sub

Private Sub obtenerRSOperaciones()
    Dim clsFun As DFunciones.dFuncionesNeg
    Set clsFun = New DFunciones.dFuncionesNeg
        
    clsFun.GetOperaciones rsOpeLLama, rsCMACs


End Sub
Private Sub FuncionesIni()
    Call CargaVarSistema
    Call fgITFParametros
    Call obtenerRSOperaciones
End Sub
