VERSION 5.00
Begin VB.Form frmProdPersona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Creditos de Persona"
   ClientHeight    =   3420
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4830
   Icon            =   "frmProdPersona.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3420
   ScaleWidth      =   4830
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2520
      TabIndex        =   2
      Top             =   3000
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1305
      TabIndex        =   1
      Top             =   3000
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
      ForeColor       =   &H80000002&
      Height          =   2445
      Left            =   90
      TabIndex        =   3
      Top             =   450
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
         TabIndex        =   0
         Top             =   300
         Width           =   4380
      End
   End
   Begin VB.Label lblNombrePersona 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   900
      TabIndex        =   5
      Top             =   90
      Width           =   3735
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
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   765
   End
End
Attribute VB_Name = "frmProdPersona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsCuenta As New COMDPersona.UCOMProdPersona

Public Function Inicio(ByVal psPersNombre As String, ByVal prCtas As ADODB.Recordset) As COMDPersona.UCOMProdPersona
    
    Me.lblNombrePersona.Caption = psPersNombre
    
    If Not (prCtas.EOF And prCtas.EOF) Then
        Do While Not prCtas.EOF
            lstCuentas.AddItem prCtas("cCtaCod") & Space(2) & prCtas("cRelacion") & Space(2) & Trim(prCtas("cEstado"))
            prCtas.MoveNext
        Loop
        lstCuentas.Selected(0) = True 'EJVG20151028
    Else
        'MsgBox "Persona no posee creditos con estas condiciones", vbInformation, "Aviso"
        MsgBox "Persona no posee cuentas con estas condiciones", vbInformation, "Aviso" 'FRHU 20150210 ERS048-2014
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
Else 'EJVG20151028
    MsgBox "Ud. debe seleccionar la cuenta", vbInformation, "Aviso"
    Exit Sub
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
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Set clsCuenta = New COMDPersona.UCOMProdPersona
End Sub
