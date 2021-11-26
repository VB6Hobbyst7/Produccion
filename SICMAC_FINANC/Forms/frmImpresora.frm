VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmImpresora 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   2355
   ClientLeft      =   3990
   ClientTop       =   2715
   ClientWidth     =   5955
   ControlBox      =   0   'False
   Icon            =   "frmImpresora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   5955
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4545
      TabIndex        =   2
      Top             =   1950
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impresora"
      Height          =   1725
      Left            =   180
      TabIndex        =   3
      Top             =   60
      Width           =   5520
      Begin Spinner.uSpinner txtcopias 
         Height          =   375
         Left            =   4440
         TabIndex        =   10
         Top             =   780
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   661
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin VB.ComboBox cboOrienta 
         Height          =   315
         ItemData        =   "frmImpresora.frx":030A
         Left            =   1170
         List            =   "frmImpresora.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1230
         Visible         =   0   'False
         Width           =   1515
      End
      Begin VB.ComboBox cboPrinters 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   4005
      End
      Begin VB.Label Label3 
         Caption         =   "Orientación"
         Height          =   195
         Left            =   210
         TabIndex        =   8
         Top             =   1320
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Copias :"
         Height          =   195
         Left            =   3495
         TabIndex        =   7
         Top             =   870
         Width           =   795
      End
      Begin VB.Label lbltipo 
         Height          =   210
         Left            =   900
         TabIndex        =   6
         Top             =   840
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   300
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo :"
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   825
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3360
      TabIndex        =   1
      Top             =   1950
      Width           =   1155
   End
End
Attribute VB_Name = "frmImpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Base 1
Dim lsPuerto As String
Dim indice As Integer
Dim lsPrintDefault As String
Dim lsPrintLocales() As String
Dim TotalLocales As Integer

Private Sub cboOrienta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub

Private Sub cboPrinters_Click()
  lbltipo = Buscatipo(Trim(cboPrinters))
End Sub
Private Sub cmdAceptar_Click()
Dim I As Integer
    If cboOrienta.ListIndex = 0 Then
      gnLinPage = gnLinHori
    Else
      gnLinPage = gnLinVert
    End If
    If Left(cboPrinters, 2) <> "\\" Then
      For I = 1 To TotalLocales
        If Trim(cboPrinters) = lsPrintLocales(1, I) Then
            sLPT = lsPrintLocales(2, I)
            Exit For
        End If
      Next I
    Else
        sLPT = EliminaEspacios(Trim(cboPrinters))
    End If
    lnNumCopias = Int(Val(txtcopias.Valor))
    lbCancela = False
    Unload Me
End Sub

Private Sub CmdCancelar_Click()
    lbCancela = True
    Unload Me
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    cmdCancelar.value = True
End If
End Sub

Private Sub Form_Load()
Dim X As Printer
Dim I As Integer
   ReDim lsPrintLocales(1 To 2, 1 To 2)
   lsPrintDefault = Printer.DeviceName
   Me.Caption = "Selecionar Impresora"
   CentraForm Me
   
   For Each X In Printers
     cboPrinters.AddItem Trim(X.DeviceName)
       If Left(X.DeviceName, 2) <> "\\" Then
        TotalLocales = TotalLocales + 1
        ReDim Preserve lsPrintLocales(1 To 2, 1 To TotalLocales)
        lsPrintLocales(1, TotalLocales) = Trim(X.DeviceName)
        lsPrintLocales(2, TotalLocales) = Left(X.Port, 4)
     End If
   Next
  txtcopias.Valor = 1
  If cboPrinters.ListCount <> 0 Then
    For I = 0 To cboPrinters.ListCount
       If Trim(cboPrinters.List(I)) = Trim(lsPrintDefault) Then
          indice = I
          Exit For
       End If
    Next
    cboPrinters.ListIndex = indice
  End If
  cboOrienta.ListIndex = IIf(gnLinPage = gnLinVert, 1, 0)
End Sub
Public Function Buscatipo(psNombreprint As String) As String
Dim indice As Integer
    indice = InStr(3, Trim(psNombreprint), "\")
    If indice <> 0 Then
       Buscatipo = Mid(psNombreprint, indice + 1, Len(Trim(psNombreprint)))
    Else
       Buscatipo = Trim(psNombreprint)
    End If
End Function

Private Sub txtcopias_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
    If txtcopias.Valor = 0 Then
        txtcopias.Valor = 1
    End If
    cboOrienta.SetFocus
  End If
End Sub
Private Sub txtcopias_LostFocus()
    If txtcopias.Valor = 0 Then
       txtcopias.Valor = 1
    End If
End Sub
Private Function EliminaEspacios(psCadena As String) As String
Dim Cadena1 As String
Dim Cadena2 As String
Dim I   As Long
Dim s   As String
Dim nxt As Integer
    Cadena2 = ""
    Cadena1 = ""
    I = 0
    s = Trim(psCadena)
    Do
       nxt = InStr(s, " ")
       If nxt Then
          Cadena1 = RTrim(Left(s, nxt - 1))
          s = LTrim(Mid(s, nxt + 1))
       Else
          Cadena1 = s
       End If
       Cadena2 = Trim(Cadena2) + Trim(Cadena1)
    Loop Until nxt = 0
    EliminaEspacios = Cadena2
End Function
