VERSION 5.00
Begin VB.Form frmJoyGarPersona 
   Caption         =   "Credito/Joya de Persona"
   ClientHeight    =   3480
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4875
   Icon            =   "frmJoyGarPersona.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   4875
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame fraCuentas 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2445
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   4650
      Begin VB.ListBox lstCuentas 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1950
         Left            =   130
         TabIndex        =   3
         Top             =   300
         Width           =   4380
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1335
      TabIndex        =   1
      Top             =   3030
      Width           =   1000
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2550
      TabIndex        =   0
      Top             =   3030
      Width           =   1000
   End
   Begin VB.Label Label1 
      Caption         =   "Persona"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   150
      TabIndex        =   5
      Top             =   150
      Width           =   765
   End
   Begin VB.Label lblNombrePersona 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   930
      TabIndex        =   4
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "frmJoyGarPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsCuenta As New COMDPersona.UCOMProdPersona

Public Function Inicio(ByVal psPersNombre As String, ByVal prCtas As ADODB.Recordset, ByVal iTipo As Integer) As COMDPersona.UCOMProdPersona
    
    Me.lblNombrePersona.Caption = psPersNombre
    
    If Not (prCtas.EOF And prCtas.EOF) Then
        If iTipo = 1 Then
            Do While Not prCtas.EOF
                lstCuentas.AddItem prCtas("cCtaCod") & Space(2) & prCtas("cJoyGarCod") & Space(2) & Trim(prCtas("cConsDescripcion"))
                prCtas.MoveNext
            Loop
        Else
            Do While Not prCtas.EOF
                lstCuentas.AddItem prCtas("cJoyGarCod") & Space(2) & Trim(prCtas("cConsDescripcion"))
                prCtas.MoveNext
            Loop
        End If
    Else
        MsgBox "Persona no posee creditos con estas condiciones", vbInformation, "Aviso"
    End If

    Me.Show 1
    Set Inicio = clsCuenta
    Set clsCuenta = Nothing
End Function

Private Sub CmdAceptar_Click()
Dim sCta As String, sProd As String, sMon As String
Dim sCadena As String
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
        End If
    End If
End If
clsCuenta.CargaDatos sCta, sProd, sMon
Unload Me
End Sub

Private Sub cmdCancelar_Click()
    clsCuenta.CargaDatos "", "", ""
    Unload Me
End Sub

Private Sub Form_Load()
Me.Caption = "Cuentas de la Persona"
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Set clsCuenta = New COMDPersona.UCOMProdPersona
End Sub

