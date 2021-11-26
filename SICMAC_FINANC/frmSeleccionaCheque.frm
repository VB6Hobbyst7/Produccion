VERSION 5.00
Begin VB.Form frmSeleccionaCheque 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Selecciona el Cheque"
   ClientHeight    =   3390
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3570
   Icon            =   "frmSeleccionaCheque.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3390
   ScaleWidth      =   3570
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCheque 
      Height          =   2790
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3375
   End
   Begin VB.CommandButton btnSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
      Height          =   350
      Left            =   2550
      TabIndex        =   1
      Top             =   3000
      Width           =   975
   End
   Begin VB.CommandButton btnAceptar 
      Caption         =   "&Aceptar"
      Height          =   350
      Left            =   1560
      TabIndex        =   0
      Top             =   3000
      Width           =   975
   End
End
Attribute VB_Name = "frmSeleccionaCheque"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim fsNroCheque As String

Private Sub Form_Load()
    Me.Move (Screen.Width - Me.Width) / 2, (Screen.Height - Me.Height) / 2, Me.Width, Me.Height
End Sub
Public Function Inicio(ByVal psIFTpo As String, ByVal psPersCod As String, ByVal psCtaIFCod As String) As String
    Dim oDDocRec As New DDocRec
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Set rs = oDDocRec.RecuperaTalonarioChequePaArbol(psIFTpo, psPersCod, psCtaIFCod)
    
    lstCheque.Clear
    fsNroCheque = ""

    If rs.BOF And rs.EOF Then
        MsgBox "No existe talonario de cheques registrados con esta cuenta" & Chr(10) & "Comuniquese con el Dpto. de Finanzas para el registro de cheques", vbInformation, "Aviso"
    Else
        Do While Not rs.EOF
            lstCheque.AddItem rs!CODIGO & " - " & Trim(rs!NOMBRE)
            rs.MoveNext
        Loop
        Me.Show 1
    End If
    Inicio = fsNroCheque
    
    Set oDDocRec = Nothing
    Set rs = Nothing
End Function
Private Sub btnAceptar_Click()
    Dim i As Integer
    fsNroCheque = ""
    For i = 0 To lstCheque.ListCount - 1
        If lstCheque.Selected(i) = True Then
            fsNroCheque = Trim(Mid(lstCheque.List(i), 1, 8))
        End If
    Next i
    If fsNroCheque = "" Then
        MsgBox "Ud. debe seleccionar el Cheque a utilizar", vbInformation, "Aviso"
        Exit Sub
    Else
        Unload Me
    End If
End Sub
Private Sub btnSalir_Click()
    fsNroCheque = ""
    Unload Me
End Sub
