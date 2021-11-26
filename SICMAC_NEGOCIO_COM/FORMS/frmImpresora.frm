VERSION 5.00
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmImpresora 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1935
   ClientLeft      =   1860
   ClientTop       =   3270
   ClientWidth     =   5715
   Icon            =   "frmImpresora.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1935
   ScaleWidth      =   5715
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   345
      Left            =   4425
      TabIndex        =   2
      Top             =   1440
      Width           =   1155
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impresora"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   1365
      Left            =   60
      TabIndex        =   3
      Top             =   30
      Width           =   5580
      Begin VB.CheckBox ChkImpTMU 
         Caption         =   "Impresora TMU"
         Height          =   255
         Left            =   840
         TabIndex        =   9
         Top             =   1000
         Width           =   2175
      End
      Begin Spinner.uSpinner txtCopias 
         Height          =   435
         Left            =   4500
         TabIndex        =   8
         Top             =   780
         Width           =   795
         _ExtentX        =   1402
         _ExtentY        =   767
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
         Width           =   4005
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
         Caption         =   "hhh"
         Height          =   210
         Left            =   900
         TabIndex        =   6
         Top             =   750
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
         Top             =   750
         Width           =   405
      End
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   345
      Left            =   3225
      TabIndex        =   1
      Top             =   1440
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
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpbuffer As String, nSize As Long) As Long

'ALPA
Private Sub cboPrinters_Click()
  lbltipo = Buscatipo(Trim(cboPrinters))
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
    If ChkImpTMU.value = 1 Then
        gbImpTMU = True
    Else
        gbImpTMU = False
    End If
    If ImpreSensa = True Then
        lnNumCopias = Int(val(txtCopias.valor))
        lbCancela = False
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    lbCancela = True
    Unload Me
End Sub

'Private Sub Command1_Click()
'frmPosSBS.Show 1
'End Sub

Private Sub Form_Load()
Dim x As Printer
Dim i As Integer
Dim lsImpresora As String
'WIOR 20131014 ******************************************
Dim RS As New ADODB.Recordset
Dim oConecta As New COMConecta.DCOMConecta
Dim sServidores As String
Dim MatServidores() As String
Dim bEsServidor As Boolean

oConecta.AbreConexion
Set RS = oConecta.CargaRecordSet("SELECT nConsSisValor FROM ConstSistema WHERE nConsSisCod=443")
oConecta.CierraConexion
Set oConecta = Nothing

If Not (RS.EOF And RS.BOF) Then
    sServidores = Trim(RS!nConsSisValor)
Else
    sServidores = ""
End If
Set RS = Nothing
bEsServidor = False
MatServidores = Split(sServidores, ",")
For i = 0 To UBound(MatServidores)
    If Trim(MatServidores(i)) = Trim(GetMaquinaUsuarioImpresora) Then
        bEsServidor = True
    End If
Next i
'WIOR FIN ***********************************************
   Me.Icon = LoadPicture(App.path & gsRutaIcono)
   ReDim lsPrintLocales(1 To 2, 1 To 2)
   'ALPA20130923**************************
   'lsPrintDefault = Printer.DeviceName
   lsPrintDefault = IIf(bEsServidor, Printer.DeviceName, Trim(IIf(Left(Printer.DeviceName, 2) = "\\", Printer.DeviceName, "\\" & GetMaquinaUsuarioImpresora & "\" & Printer.DeviceName))) 'Printer.DeviceName'WIOR 20131014
   '*************************************
   'lsPuerto = Printer.Port
   Me.Caption = "Seleccionar Impresora"
   For Each x In Printers
    'ALPA20130923********************************
    lsImpresora = IIf(bEsServidor, Trim(x.DeviceName), Trim(IIf(Left(x.DeviceName, 2) = "\\", x.DeviceName, "\\" & GetMaquinaUsuarioImpresora & "\" & x.DeviceName))) 'WIOR 20131014
     cboPrinters.AddItem lsImpresora
     'cboPrinters.AddItem x.DeviceName
     'If Left(x.DeviceName, 2) <> "\\" Then
      If Left(lsImpresora, 2) <> "\\" Then
    '************************************
        TotalLocales = TotalLocales + 1
        ReDim Preserve lsPrintLocales(1 To 2, 1 To TotalLocales)
        lsPrintLocales(1, TotalLocales) = Trim(x.DeviceName)
        lsPrintLocales(2, TotalLocales) = Left(x.port, 4)
     End If
   Next
  txtCopias.valor = 1
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
  If KeyAscii = 13 Then
    If txtCopias.valor = 0 Then
        txtCopias.valor = 1
    End If
    cmdAceptar.SetFocus
  End If
End Sub

Private Sub txtcopias_LostFocus()
    If txtCopias.valor = 0 Then
        txtCopias.valor = 1
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
'ALPA20130924****************************************
Private Function GetMaquinaUsuarioImpresora() As String  'Para obtener la Maquina del Usuario
    Dim buffMaq As String
    Dim lSizeMaq As Long
    buffMaq = Space(255)
    lSizeMaq = Len(buffMaq)
    GetComputerName buffMaq, lSizeMaq
    GetMaquinaUsuarioImpresora = Trim(Left$(buffMaq, lSizeMaq))
End Function
