VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmImpresora 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1785
   ClientLeft      =   3990
   ClientTop       =   2715
   ClientWidth     =   5580
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   5580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4365
      TabIndex        =   2
      Top             =   1395
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   30
      TabIndex        =   3
      Top             =   0
      Width           =   5520
      Begin Spinner.uSpinner txtcopias 
         Height          =   345
         Left            =   4395
         TabIndex        =   7
         Top             =   810
         Width           =   705
         _ExtentX        =   1244
         _ExtentY        =   609
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
      Begin VB.ComboBox cboPrinters 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   4335
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Copias :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3495
         TabIndex        =   6
         Top             =   870
         Width           =   975
      End
      Begin VB.Label lbltipo 
         BackStyle       =   0  'Transparent
         Height          =   210
         Left            =   900
         TabIndex        =   5
         Top             =   840
         Width           =   1920
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Nombre :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   210
         TabIndex        =   4
         Top             =   300
         Width           =   780
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3180
      TabIndex        =   1
      Top             =   1395
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

''Private Sub cboOrienta_KeyPress(KeyAscii As Integer)
''If KeyAscii = 13 Then
''   cmdAceptar.SetFocus
''End If
''End Sub

Private Sub cboPrinters_Click()
  lbltipo = Buscatipo(Trim(cboPrinters))
End Sub
Private Sub CmdAceptar_Click()
Dim i As Integer
    'If cboOrienta.ListIndex = 0 Then
    '  gnLinPage = gnLinHori
    'Else
    '  gnLinPage = gnLinVert
    'End If
    If Left(cboPrinters, 2) <> "\\" Then
      For i = 1 To TotalLocales
        If Trim(cboPrinters) = lsPrintLocales(1, i) Then
            sLpt = lsPrintLocales(2, i)
            Exit For
        End If
      Next i
    Else
        sLpt = EliminaEspacios(Trim(cboPrinters))
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
Dim i As Integer
   ReDim lsPrintLocales(1 To 2, 1 To 2)
   lsPrintDefault = Printer.DeviceName
   Me.Caption = "Selecionar Impresora"
   'gcIntCentra = CentraSdi(Me)
   
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
    For i = 0 To cboPrinters.ListCount
       If Trim(cboPrinters.List(i)) = Trim(lsPrintDefault) Then
          indice = i
          Exit For
       End If
    Next
    cboPrinters.ListIndex = indice
  End If
  'cboOrienta.ListIndex = IIf(gnLinPage = gnLinVert, 1, 0)
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
    'cboOrienta.SetFocus
  End If
End Sub
Private Sub txtcopias_LostFocus()
    If txtcopias.Valor = 0 Then
       txtcopias.Valor = 1
    End If
End Sub
Public Function EliminaEspacios(psCadena As String) As String
Dim Cadena1 As String
Dim Cadena2 As String
Dim i   As Long
Dim s   As String
Dim nxt As Integer
    Cadena2 = ""
    Cadena1 = ""
    i = 0
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
