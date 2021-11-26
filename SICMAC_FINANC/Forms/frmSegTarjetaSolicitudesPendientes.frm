VERSION 5.00
Begin VB.Form frmSegTarjetaSolicitudesPendientes 
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6450
   Icon            =   "frmSegTarjetaSolicitudesPendientes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4050
   ScaleWidth      =   6450
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   3600
      Width           =   975
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   3480
      TabIndex        =   2
      Top             =   3600
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Caption         =   "Solicitudes"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.ListBox lstSolicitudes 
         Height          =   2790
         Left            =   240
         TabIndex        =   1
         Top             =   360
         Width           =   5775
      End
   End
End
Attribute VB_Name = "frmSegTarjetaSolicitudesPendientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsCuenta As UCapCuentas
Public Function Inicio() As UCapCuentas
    If lstSolicitudes.ListCount > 0 Then
        lstSolicitudes.ListIndex = 0
        Me.Show 1
        Set Inicio = clsCuenta
        Set clsCuenta = Nothing
    Else
        MsgBox "No hay ninguna solicitud pendiente", vbInformation, "Aviso"
    End If
End Function
Private Sub Form_Load()
    Me.Caption = "Depósito por Activación Seguro"
    Set clsCuenta = New UCapCuentas
End Sub
Private Sub cmdCancelar_Click()
    clsCuenta.CargaDatosSegTarj "", "", "", ""
    Unload Me
End Sub
Private Sub cmdAceptar_Click()
    Dim sCta As String, sNumSol As String, sMonto As String, sPersNom As String
    Dim sCadena As String, sRel As String
    Dim nPos As Integer
    sCta = ""
    sPersNom = ""
    sNumSol = ""
    sMonto = ""
    sCadena = lstSolicitudes.List(lstSolicitudes.ListIndex)
    nPos = InStr(1, sCadena, Space(2), vbTextCompare)

    If nPos > 0 Then
        sNumSol = Mid(sCadena, 1, nPos - 1)
        sCadena = Mid(sCadena, nPos + 2, Len(sCadena) - nPos - 1)
        nPos = InStr(1, sCadena, Space(2), vbTextCompare)
        If RecuperarCtaAhorro(sNumSol) = "" Then
            clsCuenta.CargaDatosSegTarj "", "", "", ""
            Unload Me
            Exit Sub
        Else
            sCta = RecuperarCtaAhorro(sNumSol)
        End If
        If nPos > 0 Then
            sPersNom = Mid(sCadena, 1, nPos - 1)
            sCadena = Mid(sCadena, nPos + 2, Len(sCadena) - nPos - 1)
            nPos = InStr(1, sCadena, Space(2), vbTextCompare)
            If nPos = 0 Then
                sMonto = Trim(sCadena)
            End If
            'If nPos > 0 Then
            '    sMonto = Mid(sCadena, 1, nPos - 1)
            '    sCadena = Mid(sCadena, nPos + 2, Len(sCadena) - nPos - 1)
            '    If sCadena <> "" Then
            '        sMonto = Trim(sCadena)
            '    End If
            'End If
        End If
    End If
    clsCuenta.CargaDatosSegTarj sNumSol, sPersNom, sCta, sMonto
    Unload Me
End Sub
Private Function RecuperarCtaAhorro(ByVal psNumSol As String) As String
    Dim oNSeguro As New NSeguros
    Dim rsSeguro As ADODB.Recordset
    Set rsSeguro = oNSeguro.RecuperarCtaAhorro(psNumSol)
    If Not rsSeguro.BOF And Not rsSeguro.EOF Then
        RecuperarCtaAhorro = rsSeguro!Cuenta
    Else
        RecuperarCtaAhorro = ""
    End If
    Set oNSeguro = Nothing
    Set rsSeguro = Nothing
End Function
