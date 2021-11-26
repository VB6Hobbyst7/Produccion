VERSION 5.00
Begin VB.Form frmPigRegHolograma 
   Caption         =   "Registro de Rango Holograma"
   ClientHeight    =   1980
   ClientLeft      =   9870
   ClientTop       =   5415
   ClientWidth     =   3840
   Icon            =   "frmPigRegHolograma.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   1980
   ScaleWidth      =   3840
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   2160
      TabIndex        =   5
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "Guardar"
      Height          =   375
      Left            =   360
      TabIndex        =   4
      Top             =   1440
      Width           =   1215
   End
   Begin VB.TextBox txtFin 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   3
      Top             =   720
      Width           =   1935
   End
   Begin VB.TextBox txtInicio 
      Alignment       =   1  'Right Justify
      Height          =   285
      Left            =   1680
      MaxLength       =   8
      TabIndex        =   2
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label2 
      Caption         =   "Holograma Fin :"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Holograma Inicio :"
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1335
   End
End
Attribute VB_Name = "frmPigRegHolograma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**************************
'* REGISTRO DE HOLOGRAMAS
'Archivo:  frmPigRegHolograma.frm
''JUCS   :  28/11/2017
'Resumen:  Nos permite registrar los hologramas para tener un inventario de ello

Option Explicit
Dim nInicio As Integer
Dim nCod As Integer
'APRI20190515 ERS005-2019
Dim gnHologIni As Long
Dim gnHologFin As Long
Dim gnMaxHolograma As Long
'END APRI
Private Sub cmdCancelar_Click()
    If nInicio = 1 Then
        Call Limpiar
    Else
        Unload Me
    End If
End Sub
Public Sub Inicio(ByVal pnInicio As Integer, Optional ByVal pnCod As Long, Optional ByVal pnHologIni As Long, Optional ByVal pnHologFin As Long, Optional ByVal pnContador As Integer = 0)
   'APRI20190515 ERS005-2019 Add parametro pnContador
    nInicio = pnInicio
    nCod = pnCod
    If nInicio = 1 Then
        Me.Show 1
    Else
        Me.Caption = "Edición de Rango Holograma"
        txtInicio.Text = pnHologIni
        txtFin.Text = pnHologFin
        'APRI20190515 ERS005-2019
        gnHologIni = pnHologIni
        gnHologFin = pnHologFin
        gnMaxHolograma = ObtenerMaxNroHolograma(pnCod)
        If pnContador > 0 Then
            txtInicio.Enabled = False
        End If
        'END APRI
        cmdCancelar.Caption = "Salir"
        Me.Show 1
    End If
End Sub
Private Sub Limpiar()
    txtInicio.Text = ""
    txtFin.Text = ""
End Sub
Private Sub cmdGuardar_Click()
Dim oNContFunc As New COMNContabilidad.NCOMContFunciones
Dim obj As New COMDColocPig.DCOMColPContrato
Dim sMovNro As String
      If Validacion Then
        If MsgBox("¿Esta Seguro de Guardar los datos? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
         
         If nInicio = 1 Then
            sMovNro = oNContFunc.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            obj.GuardarHologramas sMovNro, CLng(txtInicio.Text), CLng(txtFin.Text)
            MsgBox "Se han guardado los datos correctamente", vbInformation, "Aviso"
         Else
            obj.ActualizaMontosHologramas nCod, CLng(txtInicio.Text), CLng(txtFin.Text)
            MsgBox "Se han actualizado los datos correctamente", vbInformation, "Aviso"
         End If
         Unload Me
        End If
    End If
End Sub
Private Function Validacion() As Boolean
    Validacion = True
    If txtInicio.Text = "" Then
        MsgBox "Debe Ingresar en N° de Holograma Inicio", vbInformation, " Aviso "
        txtInicio.SetFocus
        Validacion = False
    ElseIf txtFin.Text = "" Then
        MsgBox "Debe Ingresar en N° de Holograma Fin", vbInformation, " Aviso "
        txtFin.SetFocus
        Validacion = False
    ElseIf CLng(txtInicio.Text) >= CLng(txtFin.Text) Then
        MsgBox "El N° de Holograma Inicio no puede ser mayor que el N° de Holograma Fin", vbInformation, " Aviso "
        txtInicio.SetFocus
        Validacion = False
    'COMENTADO APRI20190517 ERS005-2019
    'ElseIf VerificaRangoUtilizado(CLng(txtInicio.Text), CLng(txtFin.Text), nCod) Then
    'ElseIf nInicio = 1 And VerificaRangoUtilizado(CLng(txtInicio.Text), CLng(txtFin.Text), nCod, gsCodAge) Then
    '    MsgBox "El rango de holograma ya se encuentra registrado", vbInformation, " Aviso "
    '    Validacion = False
    'APRI20190515 ERS005-2019
    ElseIf nInicio = 2 And (CLng(txtFin.Text) < gnMaxHolograma) Then
        MsgBox "El rango de Holograma Fin no debe ser menor al N° de holograma máximo(" & gnMaxHolograma & ") utilizado", vbInformation, " Aviso "
        Validacion = False
    'END APRI
    End If
End Function
Private Sub txtFin_KeyPress(KeyAscii As Integer)
    KeyAscii = NumerosDecimales(txtFin, KeyAscii, 10, 2)
    If KeyAscii = 13 Then
        cmdGuardar.SetFocus
    End If
End Sub

Private Sub txtInicio_KeyPress(KeyAscii As Integer)
     KeyAscii = NumerosDecimales(txtInicio, KeyAscii, 10, 2)
     If KeyAscii = 13 Then
        txtFin.SetFocus
    End If
End Sub
