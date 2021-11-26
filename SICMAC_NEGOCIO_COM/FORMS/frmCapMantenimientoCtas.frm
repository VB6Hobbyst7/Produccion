VERSION 5.00
Begin VB.Form frmCapMantenimientoCtas 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5610
   Icon            =   "frmCapMantenimientoCtas.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   5610
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2865
      TabIndex        =   2
      Top             =   2640
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1725
      TabIndex        =   1
      Top             =   2625
      Width           =   1000
   End
   Begin VB.Frame fraCuentas 
      Caption         =   "Cuentas"
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
      Height          =   2505
      Left            =   90
      TabIndex        =   3
      Top             =   45
      Width           =   5430
      Begin VB.ListBox lstCuentas 
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2010
         Left            =   130
         TabIndex        =   0
         Top             =   300
         Width           =   5160
      End
   End
End
Attribute VB_Name = "frmCapMantenimientoCtas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsCuenta As UCapCuenta 'UCapCuentas
Dim vbTIpo As Integer

Public Function Inicia(Optional ByVal nTipo As Integer = 0) As UCapCuenta ' UCapCuentas
If lstCuentas.ListCount > 0 Then
    lstCuentas.ListIndex = 0
    vbTIpo = nTipo
    
    Me.Show 1
    Set Inicia = clsCuenta
    Set clsCuenta = Nothing
Else
    MsgBox "Cliente no posee cuentas adicionales para relacionar.", vbInformation, "Aviso"
End If
End Function

Private Sub CmdAceptar_Click()


If vbTIpo = 1 Then
  OtraConfiguracion
  Exit Sub
End If

Dim sCta As String, sProd As String, sMon As String
Dim sCadena As String, sRel As String
Dim nPos As Integer
sCta = ""
sProd = ""
sMon = ""
sCadena = lstCuentas.List(lstCuentas.ListIndex)
nPos = InStr(1, sCadena, Space(2), vbTextCompare)



If nPos > 0 Then
    sCta = Mid(sCadena, 1, nPos - 1)
    sCadena = Mid(sCadena, nPos + 2, Len(sCadena) - nPos - 1)
    nPos = InStr(1, sCadena, Space(2), vbTextCompare)
    If nPos > 0 Then
        sProd = Mid(sCadena, 1, nPos - 1)
        sCadena = Mid(sCadena, nPos + 2, Len(sCadena) - nPos - 1)
        nPos = InStr(1, sCadena, Space(2), vbTextCompare)
        If nPos > 0 Then
            sMon = Mid(sCadena, 1, nPos - 1)
            sCadena = Mid(sCadena, nPos + 2, Len(sCadena) - nPos - 1)
            If sCadena <> "" Then
                sRel = Trim(sCadena)
            End If
        End If
    End If
End If
clsCuenta.CargaDatos sCta, sProd, sMon, sRel
Unload Me
End Sub


Private Sub OtraConfiguracion()

Dim sCta As String, sProd As String, sMon As String
Dim sCadena As String, sEst As String, sPer As String
Dim nPos As Integer
sCta = ""
sProd = ""
sMon = ""
sEst = ""
sPer = ""

sCadena = lstCuentas.List(lstCuentas.ListIndex)
nPos = InStr(1, sCadena, Space(2), vbTextCompare)

If nPos > 0 Then
    sCta = Mid(sCadena, 1, nPos - 1)
    sCadena = Mid(sCadena, nPos + 2, Len(sCadena) - nPos - 1)
    nPos = InStr(1, sCadena, Space(2), vbTextCompare)
    If nPos > 0 Then
        sProd = Mid(sCadena, 1, nPos - 1)
        sCadena = Mid(sCadena, nPos + 2, Len(sCadena) - nPos - 1)
        nPos = InStr(1, sCadena, Space(2), vbTextCompare)
        If nPos > 0 Then
            sMon = Mid(sCadena, 1, nPos - 1)
            sCadena = Mid(sCadena, nPos + 2, Len(sCadena) - nPos - 1)
            nPos = InStr(1, sCadena, Space(2), vbTextCompare)
          If nPos > 0 Then
              sEst = Mid(sCadena, 1, nPos - 1)
              sCadena = Mid(sCadena, nPos + 2, Len(sCadena) - nPos - 1)
                If sCadena <> "" Then
                    sPer = Trim(sCadena)
                End If
          End If
        End If
     End If
End If

clsCuenta.CargaDatos2 sCta, sProd, sMon, sEst, sPer
Unload Me

End Sub

Private Sub cmdCancelar_Click()
clsCuenta.CargaDatos "", "", "", ""
Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Cuentas de la Persona"
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Set clsCuenta = New UCapCuenta 'UCapCuentas
End Sub
