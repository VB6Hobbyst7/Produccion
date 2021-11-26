VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmCartaPrend 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Editor de Textos"
   ClientHeight    =   2475
   ClientLeft      =   1845
   ClientTop       =   1875
   ClientWidth     =   4170
   HelpContextID   =   280
   Icon            =   "frmCartaPrend.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   4170
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   10530
      TabIndex        =   0
      Top             =   60
      Width           =   975
   End
   Begin MSComDlg.CommonDialog dlgCarta 
      Left            =   9690
      Top             =   570
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imgBarra 
      Left            =   10245
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   21
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":041C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":052E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":0640
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":0752
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":0864
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":0976
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":0A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":0B9A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":0CAC
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":0DBE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":0ED0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":0FE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":10F4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":1206
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":1F78
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":208A
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":2DFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":2F0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":3F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCartaPrend.frx":52C2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tlbCarta 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   4170
      _ExtentX        =   7355
      _ExtentY        =   1905
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "imgBarra"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   27
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Nuevo"
            Object.ToolTipText     =   "Nuevo"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Abrir"
            Object.ToolTipText     =   "Abrir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Grabar"
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Imprimir"
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Cortar"
            Object.ToolTipText     =   "Cortar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Copiar"
            Object.ToolTipText     =   "Copiar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Pegar"
            Object.ToolTipText     =   "Pegar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Borrar"
            Object.ToolTipText     =   "Borrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Negrita"
            Object.ToolTipText     =   "Negrita"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Itálica"
            Object.ToolTipText     =   "Itálica"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Subrayado"
            Object.ToolTipText     =   "Subrayado"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Izquierda"
            Object.ToolTipText     =   "Alinear a la Izquierda"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Centrar"
            Object.ToolTipText     =   "Centrar"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Derecha"
            Object.ToolTipText     =   "Alinear a la Derecha"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TabIzq"
            Object.ToolTipText     =   "Tab Izquierdo"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UndoTabIzq"
            Object.ToolTipText     =   "Undo Tab Izquierdo"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "TabDer"
            Object.ToolTipText     =   "Tab Derecho"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "UndoTabDer"
            Object.ToolTipText     =   "Undo Tab Derecho"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tamaño-"
            Object.ToolTipText     =   "Tamaño menor"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Tamaño+"
            Object.ToolTipText     =   "Tamaño mayor"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "Proteger"
            Object.ToolTipText     =   "Proteger"
            ImageIndex      =   21
         EndProperty
      EndProperty
      Begin VB.ComboBox cboFonts 
         Height          =   315
         Left            =   8100
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   30
         Width           =   1365
      End
   End
   Begin RichTextLib.RichTextBox rtfCarta 
      Height          =   1065
      Left            =   150
      TabIndex        =   2
      Top             =   1095
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   1879
      _Version        =   393217
      ScrollBars      =   2
      TextRTF         =   $"frmCartaPrend.frx":5714
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmCartaPrend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'EDITOR DE CARTAS.
'Archivo:              frmCartaPrend.frm
'Fecha de creación     : -------------
'Fecha de modificación : 30/06/1999.
'Resumen:
'   Este formuario trabaja como un editor de textos y permite trabajar con archivos
'   RTF (Rich Text File - Archivo de texto enriquecido) y TXT (Archivo texto).

'permite abrir un archivo ya existente que se desee ver o actualizar
Private Sub Abrir()
With dlgCarta
    .CancelError = True
On Error GoTo ErrHandler
    .Filter = "Texto (*.txt)|*.txt|Rich TextBox (*.rtf)|*.rtf"
    .FilterIndex = 0
    '.InitDir = App.Path
    .ShowOpen
    rtfCarta.FileName = .FileName
    Exit Sub
ErrHandler:     ' El usuario ha hecho clic en el botón Cancelar
    MsgBox " Apertura cancelada ", vbInformation, " Aviso "
    Exit Sub
End With
End Sub

'Permite grabar el archivo creado o abierto por el editor; en formato RTF o TXT
Private Sub Grabar()
With dlgCarta
    .CancelError = True
On Error GoTo ErrHandler
    '.InitDir = App.Path
    .Filter = "Texto (*.txt)|*.txt|Rich TextBox (*.rtf)|*.rtf"
    .FilterIndex = 0
    .ShowSave
    If .FilterIndex = 1 Then
        rtfCarta.SaveFile .FileName, RTFTExt
    Else
        rtfCarta.SaveFile .FileName, rtfRTF
    End If
    Exit Sub
ErrHandler:     ' El usuario ha hecho clic en el botón Cancelar
    MsgBox " Grabación cancelada ", vbInformation, " Aviso "
    Exit Sub
End With
End Sub

'Permite salir del formulario actual
Private Sub cmdSalir_Click()
Unload Me
End Sub

'Inicializa el formulario actual
Private Sub Form_Load()
    'Centra el formulario
    Me.ScaleMode = vbTwips
    Me.Top = frmMdiMain.Top + 1075
    Me.Left = frmMdiMain.Left + 70
    Me.Width = frmMdiMain.Width - 170
    Me.Height = frmMdiMain.Height - 1420
    'Centra el RTF
    rtfCarta.Top = 500
    rtfCarta.Left = 0
    rtfCarta.Width = Me.Width - 110
    rtfCarta.Height = Me.Height - 850
    
    'llena el control CboFonts con los tipos de Fonts
    Dim vNroFont As Integer
    Dim I As Long
    For I = 1 To Screen.FontCount
      cboFonts.AddItem Screen.Fonts(I - 1)
      If Screen.Fonts(I - 1) = "Courier" Then vNroFont = I - 1
    Next I
    For I = 1 To cboFonts.ListCount
        If cboFonts.List(I) = "Courier" Or cboFonts.List(I) = "Courier New" Then
            vNroFont = I
            Exit For
        End If
    Next I
    cboFonts.ListIndex = vNroFont
End Sub

'cambia el tipo de letra del texto seleccionado
Private Sub cboFonts_Click()
On Error GoTo ErrHandler

rtfCarta.SelFontName = cboFonts.Text
Exit Sub

ErrHandler:     ' Errores obtenidos
Select Case Err.Number   ' Evalúa el número de error.
    Case 7
        MsgBox "Memoria insuficiente", vbInformation, " ! Aviso ! "
    Case 32011
        MsgBox "Texto Protegido", vbInformation, " ! Aviso ! "
    Case Else
        MsgBox "Operación no válida" & vbCr & _
            Err.Number & " : " & Err.Description, vbInformation, " ! Aviso ! "
End Select
End Sub

'Permite definir que opción se señalo, entre las opciones de la barra de herramientas
Private Sub tlbCarta_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo ErrHandler

With rtfCarta
    Select Case Button.Key
        Case "Nuevo"
            Me.Caption = "Editor de Textos"
            .Text = ""
        Case "Abrir"
            .SelFontName = "Courier New"
            Me.Caption = "Editor de Textos"
            Abrir
            Me.Caption = Me.Caption & " : " & dlgCarta.FileTitle
        Case "Grabar"
            Grabar
        Case "Imprimir"
            'Printer.PaperSize = 7
            .SelPrint (Printer.hdc)
        Case "Cortar"
            Clipboard.Clear                                 ' Borra el contenido del Portapapeles.
            Clipboard.SetText Screen.ActiveControl.SelText  ' Copia el texto seleccionado al Portapapeles.
            Screen.ActiveControl.SelText = ""               ' Elimina el texto seleccionado.
        Case "Copiar"
            Clipboard.Clear                                 ' Borra el contenido del Portapapeles.
            Clipboard.SetText Screen.ActiveControl.SelText  ' Copia el texto seleccionado al Portapapeles.
        Case "Pegar"
            Screen.ActiveControl.SelText = Clipboard.GetText() ' Lleva el texto del Portapapeles al control activo.
        Case "Borrar"
            Screen.ActiveControl.SelText = ""               ' Elimina el texto seleccionado.
        Case "Negrita"
            .SelBold = IIf(.SelBold = False, True, False)
        Case "Itálica"
            .SelItalic = IIf(.SelItalic = False, True, False)
        Case "Subrayado"
            .SelUnderline = IIf(.SelUnderline = False, True, False)
        Case "Izquierda"
            .SelAlignment = 0
        Case "Centrar"
            .SelAlignment = 2
        Case "Derecha"
            .SelAlignment = 1
        Case "TabIzq"
            If .SelIndent <= 13 * 567 Then
                .SelIndent = .SelIndent + (1 * 567)
            End If
        Case "UndoTabIzq"
            If .SelIndent <= 14 * 567 And .SelIndent > 0 Then
                .SelIndent = .SelIndent - (1 * 567)
            End If
        Case "TabDer"
            If .SelRightIndent <= 13 * 567 Then
                .SelRightIndent = .SelRightIndent + (1 * 567)
            End If
        Case "UndoTabDer"
            If .SelRightIndent <= 14 * 567 And .SelRightIndent > 1 * 567 Then
                .SelRightIndent = .SelRightIndent - (1 * 567)
            End If
        Case "Tamaño+"
            If .SelFontSize <= 48 Then
                .SelFontSize = .SelFontSize + 2
            End If
        Case "Tamaño-"
            If .SelFontSize >= 4 Then
                .SelFontSize = .SelFontSize - 2
            End If
        Case "Proteger"
            .SelProtected = IIf(.SelProtected = False, True, False)
        Case Else
            'Otro código.
    End Select
End With
Exit Sub

ErrHandler:     ' Errores obtenidos
Select Case Err.Number   ' Evalúa el número de error.
    Case 7
        MsgBox "Memoria insuficiente", vbInformation, " ! Aviso ! "
    Case 32011
        MsgBox "Texto Protegido", vbInformation, " ! Aviso ! "
    Case 438
        MsgBox "Operación no válida ", vbInformation, " ! Aviso ! "
    Case Else
        MsgBox "Operación no válida" & vbCr & _
            Err.Number & " : " & Err.Description, vbInformation, " ! Aviso ! "
End Select
End Sub
