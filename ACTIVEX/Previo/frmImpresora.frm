VERSION 5.00
Begin VB.Form frmImpresora 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1590
   ClientLeft      =   2190
   ClientTop       =   2925
   ClientWidth     =   5520
   Icon            =   "frmImpresora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1590
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4335
      TabIndex        =   2
      Top             =   1215
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
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
      ForeColor       =   &H00800000&
      Height          =   1110
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   5415
      Begin VB.ComboBox cboCara 
         Height          =   315
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   675
         Width           =   2760
      End
      Begin Previo.uSpinner txtcopias 
         Height          =   360
         Left            =   4500
         TabIndex        =   6
         Top             =   675
         Width           =   810
         _ExtentX        =   1429
         _ExtentY        =   635
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
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   270
         Width           =   4500
      End
      Begin VB.Label lblTipoCarateres 
         Caption         =   "Caracter"
         Height          =   165
         Left            =   150
         TabIndex        =   7
         Top             =   750
         Width           =   555
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Nº Copias :"
         Height          =   195
         Left            =   3735
         TabIndex        =   5
         Top             =   758
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Nombre :"
         Height          =   195
         Left            =   150
         TabIndex        =   4
         Top             =   300
         Width           =   645
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3135
      TabIndex        =   1
      Top             =   1215
      Width           =   1155
   End
End
Attribute VB_Name = "frmImpresora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Option Base 1
Dim lsPuerto As String
Dim indice As Integer
Dim lsPrintDefault As String
Dim lsPrintLocales() As String
Dim TotalLocales As Integer

Private Sub cboCara_Change()

End Sub

Private Sub cboPrinters_Click()
  
  'lbltipo = Buscatipo(Trim(cboPrinters))
End Sub

Private Sub CmdAceptar_Click()
    Dim i As Integer
   
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
    If ImpreSensa = True Then
        lnNumCopias = Int(Val(txtcopias.Valor))
        lbCancela = False
        
        lTpoImpresora = Val(Right(Me.cboCara.Text, 5))
        
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    
    lbCancela = True
    Unload Me
End Sub

Private Sub Form_Load()
    Dim x As Printer
    Dim i As Integer
    Dim lResult As Long
    'lResult = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    
    Me.cboCara.AddItem "EPSON" & Space(100) & Impresoras.gEPSON
    Me.cboCara.AddItem "HEWLETT PACKARD" & Space(100) & Impresoras.gHEWLETT_PACKARD
    Me.cboCara.AddItem "IBM" & Space(100) & Impresoras.gIBM
    Me.cboCara.ListIndex = CInt(lTpoImpresora) - 1
    'cboCara.Enabled = False
    
    ReDim lsPrintLocales(1 To 2, 1 To 2)
    lsPrintDefault = Printer.DeviceName
    'lsPuerto = Printer.Port
    Me.Caption = "Selecionar Impresora"
    For Each x In Printers
        cboPrinters.AddItem Trim(x.DeviceName)
        
        If Left(x.DeviceName, 2) <> "\\" Then
            TotalLocales = TotalLocales + 1
            ReDim Preserve lsPrintLocales(1 To 2, 1 To TotalLocales)
            lsPrintLocales(1, TotalLocales) = Trim(x.DeviceName)
            lsPrintLocales(2, TotalLocales) = Left(x.Port, 4)
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
  KeyAscii = NumerosEnteros(KeyAscii)
  If KeyAscii = 13 Then
    If txtcopias.Valor = 0 Then
        txtcopias.Valor = 1
    End If
    cmdAceptar.SetFocus
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
