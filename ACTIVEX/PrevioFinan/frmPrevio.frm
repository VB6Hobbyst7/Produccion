VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form frmPrevio 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Previo de Impresión : "
   ClientHeight    =   3120
   ClientLeft      =   825
   ClientTop       =   3675
   ClientWidth     =   12375
   HelpContextID   =   280
   Icon            =   "frmPrevio.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   12375
   ShowInTaskbar   =   0   'False
   Begin MSComDlg.CommonDialog cmdprint 
      Left            =   5250
      Top             =   1065
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chknoEpson 
      Caption         =   "Impresoras Graficas NO EPSON"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   4995
      TabIndex        =   18
      Top             =   2235
      Width           =   1920
   End
   Begin VB.Frame fraPosicion 
      BorderStyle     =   0  'None
      Height          =   585
      Left            =   11580
      TabIndex        =   15
      Top             =   2055
      Width           =   1245
      Begin VB.Shape Shape1 
         BorderStyle     =   3  'Dot
         Height          =   225
         Left            =   150
         Shape           =   1  'Square
         Top             =   240
         Width           =   225
      End
      Begin VB.Line Line2 
         X1              =   60
         X2              =   225
         Y1              =   240
         Y2              =   240
      End
      Begin VB.Line Line1 
         X1              =   150
         X2              =   150
         Y1              =   150
         Y2              =   300
      End
      Begin VB.Label lblX 
         Height          =   225
         Left            =   420
         TabIndex        =   16
         Top             =   240
         Width           =   780
      End
   End
   Begin VB.Frame fraPrevio 
      Caption         =   "Previo "
      Height          =   585
      Left            =   10380
      TabIndex        =   12
      Top             =   2070
      Width           =   1050
      Begin VB.ComboBox cboPrevio 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmPrevio.frx":030A
         Left            =   120
         List            =   "frmPrevio.frx":0326
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   195
         Width           =   825
      End
   End
   Begin VB.Frame fraPagina 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   150
      TabIndex        =   7
      Top             =   2070
      Width           =   1890
      Begin VB.CheckBox chkPaginas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Páginas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   -15
         Width           =   1035
      End
      Begin VB.TextBox txtPagFin 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Left            =   1095
         TabIndex        =   9
         Top             =   195
         Width           =   675
      End
      Begin VB.TextBox txtPagIni 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   135
         TabIndex        =   8
         Top             =   210
         Width           =   660
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         Height          =   195
         Left            =   840
         TabIndex        =   10
         Top             =   255
         Width           =   135
      End
   End
   Begin VB.CommandButton cmdGuardar 
      Caption         =   "&Archivo..."
      Height          =   345
      Left            =   3960
      TabIndex        =   2
      Top             =   2235
      Width           =   1005
   End
   Begin VB.Frame fraArchivos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      ForeColor       =   &H80000008&
      Height          =   570
      Left            =   6915
      TabIndex        =   6
      Top             =   2070
      Width           =   3450
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   2475
         TabIndex        =   5
         Top             =   165
         Width           =   930
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "&Grabar"
         Height          =   345
         Left            =   1560
         TabIndex        =   4
         Top             =   165
         Width           =   930
      End
      Begin VB.TextBox txtArchivo 
         Height          =   315
         Left            =   75
         TabIndex        =   3
         Top             =   165
         Width           =   1440
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   345
      Left            =   3000
      TabIndex        =   1
      Top             =   2235
      Width           =   975
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   345
      Left            =   2085
      TabIndex        =   0
      Top             =   2235
      Width           =   930
   End
   Begin RichTextLib.RichTextBox rtfImpre 
      Height          =   180
      Left            =   15
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   195
      _ExtentX        =   344
      _ExtentY        =   318
      _Version        =   393217
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      RightMargin     =   25000
      TextRTF         =   $"frmPrevio.frx":035D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfTexto 
      Height          =   1770
      Left            =   225
      TabIndex        =   17
      Top             =   165
      Width           =   4230
      _ExtentX        =   7461
      _ExtentY        =   3122
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   3
      MousePointer    =   99
      RightMargin     =   25000
      TextRTF         =   $"frmPrevio.frx":03DD
      MouseIcon       =   "frmPrevio.frx":045D
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   "Opciones"
      Visible         =   0   'False
      Begin VB.Menu mnuOpcBusSig 
         Caption         =   "Buscar Siguiente"
      End
      Begin VB.Menu mnuOpcBuscar 
         Caption         =   "Buscar..."
      End
      Begin VB.Menu mnuSepara1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpcPagina 
         Caption         =   "Ir a Página ..."
      End
      Begin VB.Menu mnuSepara2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpcAnterior 
         Caption         =   "Página Anterior del Previo"
      End
      Begin VB.Menu mnuOpcSiguiente 
         Caption         =   "Página Siguiente del Previo"
      End
      Begin VB.Menu mnuOpcInicio 
         Caption         =   "Página Inicial del Previo"
      End
      Begin VB.Menu mnuOpcFinal 
         Caption         =   "Página Final del Previo"
      End
   End
End
Attribute VB_Name = "frmPrevio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lbCondensado As Boolean
Dim vLineas As Integer
Dim vPage As Double
Dim vPosCad As String
Dim vTipBus As Integer
'By Capi 01102008
Dim lnLineas As Integer

'Rutina para mostrar un previo de impresión
'Recepción de RTF para impresión previa
Public Sub PrevioFinan(ByVal pRTFImpresion As String, vTitulo As String, pbCondensado As Boolean, nLineas As Integer)
    Dim vSeparador As String
    Dim vTexto As String
    If Right(Trim(pRTFImpresion), 1) = oImpresora.gPrnSaltoPagina Then
        pRTFImpresion = Left(pRTFImpresion, Len(Trim(pRTFImpresion)) - 1)
    End If
    rtfImpre.Text = Trim(pRTFImpresion)
    vTexto = Trim(pRTFImpresion)
    'By Capi 01102008
    lnLineas = nLineas
    lbCondensado = pbCondensado
    vLineas = nLineas
    vSeparador = String(Round(rtfTexto.RightMargin / 155), "»") & oImpresora.gPrnSaltoLinea
    vPage = 0
    
    Dim sImpre As String
    Dim sPag   As String
    Dim n      As Currency
    vTexto = Replace(vTexto, Chr(13), "")
    Do While True And Len(Trim(vTexto)) > 0
       vPage = vPage + 1
       n = InStr(1, vTexto, oImpresora.gPrnSaltoPagina)
       If n > 0 Then
          sPag = sPag & Mid(vTexto, 1, n - 1) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & "<<Página :" & ImpreFormat(vPage, 7, 0) & ">>" & oImpresora.gPrnSaltoLinea & vSeparador
          'By Capi 01102008
          'vTexto = Mid(vTexto, n + 1)
          vTexto = Mid(vTexto, n + Len(oImpresora.gPrnSaltoPagina))
       Else
           sPag = sPag & vTexto & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & "<<Página :" & ImpreFormat(vPage, 7, 0) & ">>" & oImpresora.gPrnSaltoLinea
           Exit Do
       End If
       If vPage Mod 30 = 0 Then
           sImpre = sImpre & sPag
           sPag = ""
       End If
    Loop
    rtfTexto.Text = sImpre & sPag
    cboPrevio.ListIndex = 3
    txtPagIni.Text = 1
    txtPagFin.Text = vPage
    frmPrevio.Caption = frmPrevio.Caption & vTitulo
    frmPrevio.Show 1
End Sub

'Porcentaje en Pantalla
Private Sub cboPrevio_Click()
    rtfTexto.SelStart = 0
    rtfTexto.SelLength = Len(rtfTexto.Text)
    Select Case Trim(cboPrevio.Text)
    Case "25%": rtfTexto.SelFontSize = 4 '5
    Case "50%": rtfTexto.SelFontSize = 5 '6
    Case "75%": rtfTexto.SelFontSize = 7 '7.5
    Case "100%": rtfTexto.SelFontSize = 7.5 '8
    Case "125%": rtfTexto.SelFontSize = 8.5 '9.5
    Case "150%": rtfTexto.SelFontSize = 9.5 '11
    Case "175%": rtfTexto.SelFontSize = 11 '12.5
    Case "200%": rtfTexto.SelFontSize = 12.5 '14
    End Select
    rtfTexto.SelStart = 0
End Sub

Private Sub chkPaginas_Click()
    txtPagIni.Enabled = IIf(txtPagIni.Enabled, False, True)
    txtPagFin.Enabled = IIf(txtPagFin.Enabled, False, True)
    If chkPaginas.Value = 1 Then
        txtPagIni.BackColor = &H80000005
        txtPagFin.BackColor = &H80000005
    Else
        txtPagIni.BackColor = &HC0C0C0
        txtPagFin.BackColor = &HC0C0C0
    End If
End Sub

Private Sub cmdCancelar_Click()
    fraArchivos.Visible = False
End Sub

Private Sub cmdGrabar_Click()
    Dim fs As Scripting.FileSystemObject
    Dim carpeta As Scripting.Folder
    Dim File As Scripting.File
    Dim Ruta As String
    Dim Archivo As String
    Dim ArchBusc As String
    Dim cadenaCond As String
    On Error GoTo Errores
    
    If lbCondensado = True Then
        cadenaCond = "T"
    Else
        cadenaCond = "F"
    End If
    
    If Len(txtArchivo) = 0 Then
        MsgBox "Nombre de Archivo no Válido", vbInformation, "Aviso"
        txtArchivo.SetFocus
        Exit Sub
    End If
    cadenaCond = Trim(cadenaCond) & "," & Trim(Str(vLineas)) & Space(2)
    rtfImpre.Text = rtfImpre.Text & Space(50) & cadenaCond
    Set fs = New Scripting.FileSystemObject
    Ruta = App.Path & "\Spooler"
    If fs.FolderExists(Ruta) = False Then
         fs.CreateFolder (Ruta)
    End If
        Set carpeta = fs.GetFolder(Ruta)
        Archivo = Ruta & "\" & Trim(txtArchivo) & ".txt"
        For Each File In carpeta.Files
            ArchBusc = Trim(txtArchivo) & ".txt"
            If UCase(Trim(File.Name)) = UCase(Trim(ArchBusc)) Then
                If MsgBox("El Archivo ya Existe. Desea Sobreescribirlo?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                        rtfImpre.SaveFile Archivo, rtfText
                        txtArchivo = ""
                        fraArchivos.Visible = False
                        Exit Sub
                Else
                    txtArchivo.SetFocus
                    Exit Sub
                End If
            End If
        Next
        If MsgBox("Desea Grabar el Archivo : " & ArchBusc, vbYesNo + vbQuestion, "Aviso") = vbYes Then
              rtfImpre.SaveFile Archivo, rtfText
              txtArchivo = ""
              fraArchivos.Visible = False
        End If
    Exit Sub

Errores:
    Select Case Err.Number
        Case 75
            MsgBox "Nombre de Archivo no Válido", vbInformation, "Aviso"
            txtArchivo.SetFocus
            Exit Sub
        Case Else
            MsgBox "Error: [" & Str(Err.Number) & "]  " & Err.Description, vbInformation, "Aviso"
            Exit Sub
    End Select
End Sub

Private Sub cmdGuardar_Click()
fraArchivos.Visible = True
txtArchivo.SetFocus
txtArchivo = ""
End Sub

Private Sub cmdImprimir_Click()
Dim x As Double, y As Double
Dim vPosIni As Double, vPosFin As Double
'By Capi 01102008 para salto pagina IBM
Dim vPosSP As Double, vPosSL, lnCuantos As Integer, z As Double
Dim lsParcial As String, lsTotal As String
'
'By Capi 23102008
gsImpresoraElegida = ""
'
Dim i As Integer
Dim tmpText As String
On Error GoTo ERROR75
tmpText = rtfImpre.Text
If chkPaginas.Value = 1 Then
    
    
    tmpText = oImpresora.gPrnSaltoPagina & rtfImpre.Text
    
    x = 0: y = 0
    vPosIni = 0: vPosFin = 0
    Do While True
        x = x + 1
        vPosIni = InStr(vPosIni + 1, tmpText, oImpresora.gPrnSaltoPagina, vbTextCompare)
        If x > Val(txtPagIni.Text) - 1 Then Exit Do
    Loop
    Do While True
        y = y + 1
        vPosFin = InStr(vPosFin + 1, tmpText, oImpresora.gPrnSaltoPagina, vbTextCompare)
        If y > Val(txtPagFin.Text) Then Exit Do
    Loop
    If Val(txtPagFin.Text) = vPage Then
        'By Capi 01102008
        'tmpText = Mid(tmpText, vPosIni + 1)
        tmpText = Mid(tmpText, vPosIni + Len(oImpresora.gPrnSaltoPagina))
    Else
        'By Capi 01102008
        'tmpText = Mid(tmpText, vPosIni + 1, vPosFin - (vPosIni + 2))
        tmpText = Mid(tmpText, vPosIni + Len(oImpresora.gPrnSaltoPagina), vPosFin - (vPosIni + 2 * Len(oImpresora.gPrnSaltoPagina)))
    End If
   
    
End If

'by Capi 01102008 para controlar IBM
If lTpoImpresora = gIBM Then
    x = 0: y = 0: z = 0
    
    vPosSP = 0: vPosSL = 0
    z = Len(Trim(tmpText))
    Do While True
        lnCuantos = 0
        Do While True
            x = x + 1
            vPosSL = InStr(x, tmpText, oImpresora.gPrnSaltoLinea, vbTextCompare)
            vPosSP = InStr(x, tmpText, oImpresora.gPrnSaltoPagina, vbTextCompare)
            If vPosSP >= z Then Exit Do
            If vPosSL >= vPosSP Then Exit Do
            x = vPosSL
            lnCuantos = lnCuantos + 1
        Loop
        tmpText = Replace(tmpText, oImpresora.gPrnSaltoPagina, generarSaltosLineasIBM(lnLineas - lnCuantos), 1, 1)
        
        If vPosSP >= z Or vPosSP = 0 Then Exit Do
         x = vPosSL
    Loop
    'tmpText = generarSaltosLineasIBM(40) & tmpText
End If

If chknoEpson.Value = 0 Then
       
    frmImpresora.Show 1
    Inicia lTpoImpresora
    'By capi 23102008
    If gsImpresoraElegida = "Generic" Then
        tmpText = generarSaltosLineasIBM(40) & tmpText
    End If
    '
    If lbCancela = False Then
        For i = 1 To lnNumCopias
            'By Capi 01102008 para enviar el tipo de impresora
            'ImpreBegin lbCondensado, vLineas
            ImpreBegin lbCondensado, vLineas, lTpoImpresora
'                tmpText = Replace(tmpText, oImpresora.gPrnSaltoLineaDef, oImpresora.gPrnSaltoLinea)
'                tmpText = Replace(tmpText, oImpresora.gPrnSaltoPaginaDef, oImpresora.gPrnSaltoPagina)
'                tmpText = Replace(tmpText, oImpresora.gPrnInicializa, Space(Len(oImpresora.gPrnInicializa)))
'                tmpText = Replace(tmpText, oImpresora.gPrnCondensadaON, Space(Len(oImpresora.gPrnCondensadaON)))
'                tmpText = Replace(tmpText, oImpresora.gPrnCondensadaOFF, Space(Len(oImpresora.gPrnCondensadaOFF)))
'                tmpText = Replace(tmpText, oImpresora.gPrnEspaLineaN, Space(Len(oImpresora.gPrnEspaLineaN)))
'                tmpText = Replace(tmpText, oImpresora.gPrnTpoLetraCurier, Space(Len(oImpresora.gPrnTpoLetraCurier)))
'                tmpText = Replace(tmpText, oImpresora.gPrnTamLetra10CPI, Space(Len(oImpresora.gPrnTpoLetraRoman1P)))
'                tmpText = Replace(tmpText, oImpresora.gPrnTpoLetraRoman1P, Space(Len(oImpresora.gPrnTpoLetraRoman1P)))
                Print #ArcSal, GetCadenaFormateada(ImpreCarEsp(tmpText), lTpoImpresora)
            ImpreEnd
        Next i
    End If
Else
    If CargaPrintFile = True Then
        ImprimirPrintFile tmpText
    End If
End If

Exit Sub

ERROR75:
    Select Case Err.Number
         Case 75
             MsgBox "Impresora no encontrada", vbInformation, "Aviso"
             Exit Sub
         Case Else
             MsgBox "Error: [" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
             Exit Sub
    End Select
End Sub
'By Capi 01102008 genera saltos de linea
Private Function generarSaltosLineasIBM(ByVal pnLinea As Integer) As String
    Dim x As Integer
    Dim lsSaltosLineas As String
    lsSaltosLineas = ""
    For x = 1 To pnLinea
        lsSaltosLineas = lsSaltosLineas & oImpresora.gPrnSaltoLinea
    Next x
    generarSaltosLineasIBM = lsSaltosLineas
End Function

Sub ImprimirPrintFile(ByVal tmpText As String)
Dim BeginPage, EndPage, NumCopies, i
Dim lsArchivo As String
    cmdprint.CancelError = True
     
    On Error GoTo ErrHandler
    cmdprint.ShowPrinter
    ' Obtener los valores seleccionados por el usuario en el cuadro de diálogo
    BeginPage = cmdprint.FromPage
    EndPage = cmdprint.ToPage
    NumCopies = cmdprint.Copies
    For i = 1 To NumCopies
        lsArchivo = "C:\SpoolerPrt\prntfile" & Trim(Str(i)) & ".ps"
        'ImprimePrnFile lbCondensado, vLineas, lsArchivo, tmpText
        Open lsArchivo For Output As #1
        Print #1, tmpText
        Close #1
    Next i
    Exit Sub
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub
End Sub
Private Sub cmdSalir_Click()
    Unload frmPrevioBus
    Unload Me
End Sub

Private Sub Form_Load()
    Dim lResult As Long
    'lResult = SetWindowPos(Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, FLAGS)
    
    Me.Top = 1000
    Me.Left = 50
    Me.Width = Screen.Width - 200
    Me.Height = Screen.Height - 1400
    
    rtfTexto.Top = 0.2 * 567
    rtfTexto.Height = Me.Height - (2 * 567)
    rtfTexto.Left = 0.3 * 567
    rtfTexto.Width = Me.Width - (0.7 * 567)
    
    cmdSalir.Move cmdSalir.Left, Me.Height - (1.5 * 567), cmdSalir.Width, cmdSalir.Height
    cmdImprimir.Move cmdImprimir.Left, Me.Height - (1.5 * 567), cmdImprimir.Width, cmdImprimir.Height
    cmdGuardar.Move cmdGuardar.Left, Me.Height - (1.5 * 567), cmdGuardar.Width, cmdGuardar.Height
    
    chknoEpson.Move chknoEpson.Left, Me.Height - (1.5 * 567), chknoEpson.Width, chknoEpson.Height
    
    fraArchivos.Move fraArchivos.Left, Me.Height - (1.8 * 567), fraArchivos.Width, fraArchivos.Height
    fraPagina.Move fraPagina.Left, Me.Height - (1.8 * 567), fraPagina.Width, fraPagina.Height
    fraPrevio.Move fraPrevio.Left, Me.Height - (1.8 * 567), fraPrevio.Width, fraPrevio.Height
    fraPosicion.Move fraPosicion.Left, Me.Height - (1.8 * 567)
    fraArchivos.Visible = False
    CargaPrintFile
End Sub


Private Sub mnuOpcAnterior_Click()
    SendKeys "{PGUP}", True
    'rtfTexto.SelStart = rtfTexto.SelStart
End Sub

Private Sub mnuOpcBuscar_Click()
    frmPrevioBus.cmdCancelar.Enabled = True
    frmPrevioBus.Show 1
    If frmPrevioBus.cboDireccion.Text = "Todo" Then
        rtfTexto.SelStart = 0
    End If
    vTipBus = 0 ' rtfNoHighlight
    If frmPrevioBus.chkOpc1.Value = 1 Then
        vTipBus = 2
    End If
    If frmPrevioBus.chkOpc2.Value = 1 Then
        vTipBus = vTipBus + rtfMatchCase
    End If
    
    vPosCad = rtfTexto.Find(frmPrevioBus.cboBuscar.Text, rtfTexto.SelStart + Len(frmPrevioBus.cboBuscar.Text), , vTipBus)
    If frmPrevioBus.cmdCancelar.Enabled <> False And vPosCad = -1 Then
        MsgBox "Cadena no encontrada", vbInformation, " Aviso "
    End If
End Sub

Private Sub mnuOpcBusSig_Click()
    If frmPrevioBus.cboBuscar.Text <> "" Then
        vPosCad = rtfTexto.Find(frmPrevioBus.cboBuscar.Text, rtfTexto.SelStart + Len(frmPrevioBus.cboBuscar.Text), , vTipBus)
        If vPosCad = -1 Then
            MsgBox "Cadena no encontrada", vbInformation, " Aviso "
        End If
    Else
        MsgBox "Ingrese cadena a buscar", vbInformation, " Aviso "
    End If
End Sub

Private Sub mnuOpcFinal_Click()
    rtfTexto.SelStart = Len(rtfTexto.Text)
End Sub

Private Sub mnuOpcInicio_Click()
    rtfTexto.SelStart = 0
End Sub

Private Sub mnuOpcPagina_Click()
    Dim vPagina As String
    Dim vPosCad As Double
    vPagina = InputBox("Ingrese página", " Aviso ")
    If Val(vPagina) >= 1 And Val(vPagina) <= vPage Then
        rtfTexto.SelStart = 0
        If Val(vPagina) <> 1 Then
            vPosCad = rtfTexto.Find("<<Página :" & ImpreFormat(Val(vPagina) - 1, 7, 0) & ">>", 0, , rtfWholeWord)
            If vPosCad = -1 Then
                MsgBox "Página no encontrada", vbInformation, " Aviso "
            Else
                rtfTexto.SelStart = rtfTexto.SelStart
                SendKeys "{PGDN}", True
            End If
        End If
    Else
        MsgBox "Página no Reconocida", vbInformation, " Aviso "
    End If
End Sub

Private Sub mnuOpcSiguiente_Click()
    SendKeys "{PGDN}", True
    'rtfTexto.SelStart = rtfTexto.SelStart
End Sub

Private Sub rtfTexto_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    'rtfTexto.MouseIcon = LoadPicture(App.Path & cMouseMano)
    If Button = 2 Then PopupMenu mnuOpciones
End Sub

Private Sub rtfTexto_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    'rtfTexto.MouseIcon = LoadPicture(App.Path & cMousePunto)
End Sub

Private Sub rtfTexto_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    lblX.Caption = Str(Round(x \ 85, 0) + 1) & ", " & Str(Round(y \ 180, 0) + 1)
End Sub

Private Sub txtArchivo_GotFocus()
    fEnfoque txtArchivo
End Sub

Private Sub txtArchivo_KeyPress(KeyAscii As Integer)
    If KeyAscii >= 42 And KeyAscii < 48 Then
        KeyAscii = 0
    Else
        KeyAscii = Letras(KeyAscii)
    End If
End Sub

'Validación de txtPagIni
Private Sub txtPagIni_GotFocus()
    fEnfoque txtPagIni
End Sub

Private Sub txtPagIni_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtPagIni.Text) >= 1 And Val(txtPagIni.Text) <= vPage Then
            txtPagFin.Text = txtPagIni.Text
            txtPagFin.SetFocus
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub

Private Sub txtPagIni_Validate(Cancel As Boolean)
    If Not (Val(txtPagIni.Text) >= 1 And Val(txtPagIni.Text) <= vPage) Then
        Cancel = True
    End If
End Sub

'Validación de txtPagFin
Private Sub txtPagFin_GotFocus()
    fEnfoque txtPagFin
End Sub

Private Sub txtPagFin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Val(txtPagFin.Text) <= vPage And Val(txtPagFin) >= Val(txtPagIni) Then
            cmdImprimir.SetFocus
        End If
    Else
        KeyAscii = NumerosEnteros(KeyAscii)
    End If
End Sub

Private Sub txtPagFin_Validate(Cancel As Boolean)
    If Not (Val(txtPagFin.Text) <= vPage And Val(txtPagFin) >= Val(txtPagIni)) Then
        Cancel = True
    End If
End Sub

Private Function GetCadenaFormateada(ByRef psCadena As String, pImpresora As Impresoras) As String
    Dim lsCad As String
    
    lsCad = psCadena
    
    If pImpresora = gImpDfecto Then
            
            lsCad = Replace(lsCad, oImpresora.gPrnInicializaDef, oImpresora.gPrnInicializa)
            lsCad = Replace(lsCad, oImpresora.gPrnNegritaONDef, oImpresora.gPrnNegritaON)
            lsCad = Replace(lsCad, oImpresora.gPrnNegritaOFFDef, oImpresora.gPrnNegritaOFF)
            lsCad = Replace(lsCad, oImpresora.gPrnBoldONDef, oImpresora.gPrnBoldON)
            lsCad = Replace(lsCad, oImpresora.gPrnBoldOFFDef, oImpresora.gPrnBoldOFF)
            lsCad = Replace(lsCad, oImpresora.gPrnSaltoLineaDef, oImpresora.gPrnSaltoLinea)
            lsCad = Replace(lsCad, oImpresora.gPrnSaltoPaginaDef, oImpresora.gPrnSaltoPagina)
            lsCad = Replace(lsCad, oImpresora.gPrnEspaLineaValorDef, oImpresora.gPrnEspaLineaValor)
            lsCad = Replace(lsCad, oImpresora.gPrnEspaLineaNDef, oImpresora.gPrnEspaLineaN)
            lsCad = Replace(lsCad, oImpresora.gPrnTamPaginaCabDef, oImpresora.gPrnTamPaginaCab)
            lsCad = Replace(lsCad, oImpresora.gPrnTamPagina16Def, oImpresora.gPrnTamPagina16)
            lsCad = Replace(lsCad, oImpresora.gPrnTamPagina17Def, oImpresora.gPrnTamPagina17)
            lsCad = Replace(lsCad, oImpresora.gPrnTamPagina18Def, oImpresora.gPrnTamPagina18)
            lsCad = Replace(lsCad, oImpresora.gPrnTamPagina22Def, oImpresora.gPrnTamPagina22)
            lsCad = Replace(lsCad, oImpresora.gPrnTamPagina65Def, oImpresora.gPrnTamPagina65)
            lsCad = Replace(lsCad, oImpresora.gPrnTamPagina66Def, oImpresora.gPrnTamPagina66)
            lsCad = Replace(lsCad, oImpresora.gPrnTamPagina70Def, oImpresora.gPrnTamPagina70)
            lsCad = Replace(lsCad, oImpresora.gPrnTamLetra10CPIDef, oImpresora.gPrnTamLetra10CPI)
            lsCad = Replace(lsCad, oImpresora.gPrnTamLetra12CPIDef, oImpresora.gPrnTamLetra12CPI)
            lsCad = Replace(lsCad, oImpresora.gPrnTamLetra15CPIDef, oImpresora.gPrnTamLetra15CPI)
            lsCad = Replace(lsCad, oImpresora.gPrnTpoLetraRomanDef, oImpresora.gPrnTpoLetraRoman)
            lsCad = Replace(lsCad, oImpresora.gPrnTpoLetraSansSerifDef, oImpresora.gPrnTpoLetraSansSerif)
            lsCad = Replace(lsCad, oImpresora.gPrnTpoLetraCurierDef, oImpresora.gPrnTpoLetraCurier)
            lsCad = Replace(lsCad, oImpresora.gPrnMargenIzqCabDef, oImpresora.gPrnMargenIzqCab)
            lsCad = Replace(lsCad, oImpresora.gPrnMargenIzq00Def, oImpresora.gPrnMargenIzq00)
            lsCad = Replace(lsCad, oImpresora.gPrnMargenIzq01Def, oImpresora.gPrnMargenIzq01)
            lsCad = Replace(lsCad, oImpresora.gPrnMargenIzq02Def, oImpresora.gPrnMargenIzq02)
            lsCad = Replace(lsCad, oImpresora.gPrnMargenIzq06Def, oImpresora.gPrnMargenIzq06)
            lsCad = Replace(lsCad, oImpresora.gPrnMargenIzq44Def, oImpresora.gPrnMargenIzq44)
            lsCad = Replace(lsCad, oImpresora.gPrnMargenDerCabDef, oImpresora.gPrnMargenDerCab)
            lsCad = Replace(lsCad, oImpresora.gPrnTpoLetraRoman1PDef, oImpresora.gPrnTpoLetraRoman1P)
            lsCad = Replace(lsCad, oImpresora.gPrnTpoLetraSansSerif1PDef, oImpresora.gPrnTpoLetraSansSerif1P)
            lsCad = Replace(lsCad, oImpresora.gPrnTpoLetraCurier1PDef, oImpresora.gPrnTpoLetraCurier1P)
            lsCad = Replace(lsCad, oImpresora.gPrnTpoLetraRomanE2Def, oImpresora.gPrnTpoLetraRomanE2)
            lsCad = Replace(lsCad, oImpresora.gPrnTpoLetraSansSerifE2Def, oImpresora.gPrnTpoLetraSansSerifE2)
            lsCad = Replace(lsCad, oImpresora.gPrnTpoLetraCurierE2Def, oImpresora.gPrnTpoLetraCurierE2)
            lsCad = Replace(lsCad, oImpresora.gPrnTpoLetraCurierE240Def, oImpresora.gPrnTpoLetraCurierE240)
            lsCad = Replace(lsCad, oImpresora.gPrnCondensadaOFFDef, oImpresora.gPrnCondensadaOFF)
            lsCad = Replace(lsCad, oImpresora.gPrnCondensadaONDef, oImpresora.gPrnCondensadaON)
            lsCad = Replace(lsCad, oImpresora.gPrnUnderLineONOFFDef, oImpresora.gPrnUnderLineONOFF)
            lsCad = Replace(lsCad, oImpresora.gPrnItalicONDef, oImpresora.gPrnItalicON)
            lsCad = Replace(lsCad, oImpresora.gPrnItalicOFFDef, oImpresora.gPrnItalicOFF)
            lsCad = Replace(lsCad, oImpresora.gPrnDblAnchoONDef, oImpresora.gPrnDblAnchoON)
            lsCad = Replace(lsCad, oImpresora.gPrnDblAnchoOFFDef, oImpresora.gPrnDblAnchoOFF)
            lsCad = Replace(lsCad, oImpresora.gPrnUnoMedioEspacioDef, oImpresora.gPrnUnoMedioEspacio)
            lsCad = Replace(lsCad, oImpresora.gPrnSuperIdxOnDef, oImpresora.gPrnSuperIdxOn)
            lsCad = Replace(lsCad, oImpresora.gPrnSuperIdxOFFDef, oImpresora.gPrnSuperIdxOFF)
    'By Capi 01102008 para controlar cuando es IBM
    ElseIf pImpresora = gIBM Then
        Dim lsSaltoPaginaIBM As String
        lsCad = Replace(lsCad, oImpresora.gPrnSaltoLineaDef, oImpresora.gPrnSaltoLinea)
        
     
        End If
    'By Capi 17092008 para que devuelva la cadena formateada
    'GetCadenaFormateada = psCadena
    GetCadenaFormateada = lsCad
    '
End Function
Private Function CargaPrintFile() As Boolean
Dim fs As Scripting.FileSystemObject
Dim carpeta As Scripting.Folder
Dim File As Scripting.File
Dim Ruta As String
Dim lsArchivo As String
CargaPrintFile = True
lsArchivo = App.Path + "\prtfile\PrFile32.exe"
' verificamos si el archivo esta cargado
Set fs = New Scripting.FileSystemObject
Ruta = "C:\SpoolerPrt"
If fs.FolderExists(Ruta) = False Then
     fs.CreateFolder (Ruta)
End If
If fs.FileExists(lsArchivo) = True Then
    'cargamos el spooler de impresión en la carpeta SPOOLER.
    'If VerArchivoCargado(lsArchivo) = False Then
        Shell lsArchivo + " /s:" + Ruta + "\*.ps", vbMinimizedNoFocus
    'End If
Else
    MsgBox "Archivo de Emulación de Impresión no encontrado", vbInformation, "Aviso"
    CargaPrintFile = False
End If
Set fs = Nothing
End Function
