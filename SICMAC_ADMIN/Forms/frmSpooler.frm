VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmSpooler 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   960
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HelpContextID   =   280
   Icon            =   "frmSpooler.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   11880
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox DrvUnidad 
      Height          =   315
      Left            =   30
      TabIndex        =   18
      Top             =   30
      Width           =   2505
   End
   Begin MSComctlLib.ProgressBar pgrBarra 
      Height          =   180
      Left            =   7080
      TabIndex        =   16
      Top             =   7140
      Width           =   4770
      _ExtentX        =   8414
      _ExtentY        =   318
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.StatusBar BarraEstado 
      Align           =   2  'Align Bottom
      Height          =   285
      Left            =   0
      TabIndex        =   14
      Top             =   7080
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   503
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   12347
            MinWidth        =   12347
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraPagina 
      Height          =   585
      Left            =   9120
      TabIndex        =   9
      Top             =   -45
      Width           =   1425
      Begin VB.TextBox txtPagIni 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   285
         Left            =   150
         TabIndex        =   12
         Top             =   240
         Width           =   435
      End
      Begin VB.TextBox txtPagFin 
         BackColor       =   &H80000004&
         Enabled         =   0   'False
         Height          =   300
         Left            =   900
         TabIndex        =   11
         Top             =   240
         Width           =   435
      End
      Begin VB.CheckBox chkPaginas 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Páginas"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   255
         TabIndex        =   10
         Top             =   30
         Width           =   1035
      End
      Begin VB.Label Label1 
         Caption         =   "A"
         Height          =   195
         Left            =   660
         TabIndex        =   13
         Top             =   255
         Width           =   135
      End
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6150
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpooler.frx":030A
            Key             =   "texto"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6840
      Top             =   75
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpooler.frx":084C
            Key             =   "grandes"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpooler.frx":095E
            Key             =   "detalles"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpooler.frx":0A70
            Key             =   "lista"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpooler.frx":0B82
            Key             =   "pequeños"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpooler.frx":0C94
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpooler.frx":11D6
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpooler.frx":12E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpooler.frx":13FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSpooler.frx":1554
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox rtfAux 
      Height          =   450
      Left            =   7635
      TabIndex        =   0
      Top             =   3240
      Visible         =   0   'False
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   794
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   3
      RightMargin     =   25000
      TextRTF         =   $"frmSpooler.frx":1A96
   End
   Begin VB.Frame fraZoom 
      Caption         =   "Zoom"
      Height          =   585
      Left            =   10590
      TabIndex        =   5
      Top             =   -45
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
         ItemData        =   "frmSpooler.frx":1B11
         Left            =   120
         List            =   "frmSpooler.frx":1B2D
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   195
         Width           =   825
      End
   End
   Begin MSComctlLib.Toolbar barra 
      Height          =   390
      Left            =   2580
      TabIndex        =   1
      Top             =   -30
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   11
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "previo"
            Object.ToolTipText     =   "Vista Preliminar"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "noprevio"
            Object.ToolTipText     =   "Cerrar Vista Preliminar"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar Archivo"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "imprimir"
            Object.ToolTipText     =   "Imprimir Archivo"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "grande"
            Object.ToolTipText     =   "Iconos Grandes"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pequeños"
            Object.ToolTipText     =   "Iconos Pequeños"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "listas"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "detalles"
            Object.ToolTipText     =   "Detalles"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   8
         EndProperty
      EndProperty
      Begin VB.CheckBox chkCondensado 
         Appearance      =   0  'Flat
         BackColor       =   &H80000000&
         Caption         =   "Condensado?"
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   4980
         TabIndex        =   15
         Top             =   60
         Width           =   1335
      End
   End
   Begin VB.Frame fraArchivos 
      Caption         =   "Lista de Archivos de Impresión"
      Height          =   6480
      Left            =   0
      TabIndex        =   7
      Top             =   450
      Width           =   11790
      Begin VB.DirListBox DirCarpetas 
         Appearance      =   0  'Flat
         Height          =   5940
         Left            =   150
         TabIndex        =   17
         Top             =   345
         Width           =   2805
      End
      Begin MSComctlLib.ListView lstArch 
         Height          =   6000
         Left            =   3015
         TabIndex        =   8
         ToolTipText     =   "Para Visualizar el Archivo Presione [Enter] o Doble Clic"
         Top             =   315
         Width           =   8580
         _ExtentX        =   15134
         _ExtentY        =   10583
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList2"
         SmallIcons      =   "ImageList2"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Tamaño"
            Object.Width           =   6174
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Tipo"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Modificación"
            Object.Width           =   3528
         EndProperty
      End
   End
   Begin VB.Frame fraPrevio 
      Height          =   6465
      Left            =   15
      TabIndex        =   2
      Top             =   465
      Visible         =   0   'False
      Width           =   11775
      Begin RichTextLib.RichTextBox rtfView 
         Height          =   6120
         Left            =   150
         TabIndex        =   3
         Top             =   270
         Width           =   11475
         _ExtentX        =   20241
         _ExtentY        =   10795
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   3
         RightMargin     =   3.00000e5
         TextRTF         =   $"frmSpooler.frx":1B64
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Vista Preliminar"
         Height          =   195
         Left            =   195
         TabIndex        =   4
         Top             =   60
         Width           =   1215
      End
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
      Begin VB.Menu Separa1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpcPagina 
         Caption         =   "Ir a Pagina..."
      End
      Begin VB.Menu Separa2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuOpcAnterior 
         Caption         =   "Página Anterior"
      End
      Begin VB.Menu mnuOpcSiguiente 
         Caption         =   "Página Siguiente"
      End
      Begin VB.Menu mnuOpcInicio 
         Caption         =   "Página Inicial"
      End
      Begin VB.Menu mnuOpcFinal 
         Caption         =   "Página Final"
      End
   End
End
Attribute VB_Name = "frmSpooler"
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
' Referenciar a Microsoft Scripting Runtime
Private drCurrent As Scripting.Drive
Private flCurrent As Scripting.Folder
Private fs As Scripting.FileSystemObject
Dim xt As ListItem
Dim fi As Scripting.File
Dim Condensado As Boolean
Dim NumLineas As Integer
Dim parametros As String
Dim Ruta As String
Dim vTipBus As Integer
Dim vPosCad As Double
Dim vPage As Double

Private Sub barra_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim tmpText As String
Dim x As Double, Y As Double
Dim vPosIni As Double, vPosFin As Double
Dim i As Integer
Dim Posi As Long

On Error GoTo ErrorPrint
Select Case Button.Key
    Case "grande"
        If lstArch.ListItems.Count <> 0 Then
             lstArch.View = lvwIcon
        End If
    Case "pequeños"
        If lstArch.ListItems.Count <> 0 Then
            lstArch.View = lvwSmallIcon
        End If
    Case "listas"
        If lstArch.ListItems.Count <> 0 Then
            lstArch.View = lvwList
        End If
    Case "detalles"
        If lstArch.ListItems.Count <> 0 Then
            lstArch.View = lvwReport
        End If
    Case "previo"
        If lstArch.ListItems.Count <> 0 Then
            CargaRTF
            Habilitar
        End If
    Case "noprevio"
        If lstArch.ListItems.Count <> 0 Then
            InHabilitar
        End If
    Case "eliminar"
      If lstArch.ListItems.Count <> 0 Then
        If MsgBox("Desea eliminar el Archivo : " & lstArch.SelectedItem & " ", vbYesNo + vbQuestion, "Aviso") = vbYes Then
            fs.DeleteFile (Ruta & "\" & lstArch.SelectedItem)
            Refresco
        Else
             lstArch.SetFocus
        End If
       End If
    Case "imprimir"
        If lstArch.ListItems.Count <> 0 Then
            'IMPRESION DIRECTA A IMPRESORA
            If Me.fraArchivos.Visible = True Then
                CargaRTFArchivo
                Exit Sub
            End If
            Posi = InStr(1, rtfAux.Text, "T,66")
            If Posi = 0 Then
                Posi = InStr(1, rtfAux.Text, "F,66")
            End If
            frmImpresora.Show 1
            If Posi <> 0 Then
                parametros = Trim(Right(Trim(rtfAux.Text), 6))
                ExtraeParam
                tmpText = Mid(rtfAux.Text, 1, Posi - 1)
            Else
                tmpText = rtfAux.Text
            End If
            If chkPaginas.Value = 1 Then
                If Posi <> 0 Then
                    tmpText = oImpresora.gPrnSaltoPagina & Mid(rtfAux.Text, 1, Posi - 1)
                Else
                    tmpText = oImpresora.gPrnSaltoPagina & rtfAux.Text
                End If
                x = 0: Y = 0
                vPosIni = 0: vPosFin = 0
                Do While True
                    x = x + 1
                    vPosIni = InStr(vPosIni + 1, tmpText, oImpresora.gPrnSaltoPagina, vbTextCompare)
                    If x > Val(txtPagIni.Text) - 1 Then Exit Do
                Loop
                Do While True
                    Y = Y + 1
                    vPosFin = InStr(vPosFin + 1, tmpText, oImpresora.gPrnSaltoPagina, vbTextCompare)
                    If Y > Val(txtPagFin.Text) Then Exit Do
                Loop
                If Val(txtPagFin.Text) = vPage Then
                    tmpText = Mid(tmpText, vPosIni + 1)
                Else
                    tmpText = Mid(tmpText, vPosIni + 1, vPosFin - (vPosIni + 2))
                End If
            End If
            If chkCondensado.Value = 1 Then
                Condensado = True
            Else
                Condensado = False
            End If
            If lbCancela = False Then
                For i = 1 To lnNumCopias
                    ImpreBegin Condensado, NumLineas
                        Print #ArcSal, ImpreCarEsp(tmpText)
                    ImpreEnd
                Next i
            End If
        End If
    Case "salir"
            Unload Me
End Select
Exit Sub

ErrorPrint:
    Select Case Err.Number
        Case 75
            MsgBox "Ruta de Impresora no válida", vbInformation, "Aviso"
        Case 53
            MsgBox "Archivo no Existe o ha cambiado el nombre" + Chr(13) + "Presione la tecla [F5]", vbInformation, "Aviso"
        Case Else
            MsgBox "Error :[" & Str(Err.Number) & "] " & Err.Description, vbInformation, "Aviso"
    End Select
End Sub
Sub ExtraeParam()
    Dim Cad As String
    Dim Pos As Integer
    
    Pos = InStr(1, parametros, ",")
    If Pos <> 0 Then
        Cad = Left(parametros, Pos - 1)
    End If
    If Cad = "T" Then
        Condensado = True
    Else
        If Cad = "F" Then
            Condensado = False
        End If
    End If
    Cad = Mid(parametros, Pos + 1, Len(parametros))
    NumLineas = Val(Trim(Cad))
End Sub

Private Sub cboPrevio_Click()
    Select Case Trim(cboPrevio.Text)
    Case "25%": rtfView.SelFontSize = 4 '5
    Case "50%": rtfView.SelFontSize = 5 '6
    Case "75%": rtfView.SelFontSize = 7 '7.5
    Case "100%": rtfView.SelFontSize = 7.5 '8
    Case "125%": rtfView.SelFontSize = 8.5 '9.5
    Case "150%": rtfView.SelFontSize = 9.5 '11
    Case "175%": rtfView.SelFontSize = 11 '12.5
    Case "200%": rtfView.SelFontSize = 12.5 '14
    End Select
    rtfView.Text = rtfView.Text
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

Private Sub DirCarpetas_Change()
    Me.BarraEstado.Panels(1).Text = "Contenido de " & DirCarpetas
    Ruta = DirCarpetas
    Set flCurrent = fs.GetFolder(Ruta)
    Refresco
End Sub

Private Sub DirCarpetas_Click()
   Ruta = DirCarpetas
   Set flCurrent = fs.GetFolder(Ruta)
   Refresco
End Sub

Private Sub DirCarpetas_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    Ruta = DirCarpetas
    Set flCurrent = fs.GetFolder(Ruta)
    Refresco
End If
End Sub

Private Sub DrvUnidad_Change()
  Set fs = New Scripting.FileSystemObject
    If fs.Drives(DrvUnidad.Drive).IsReady Then
        Me.DirCarpetas.path = DrvUnidad.Drive
    Else
        MsgBox "No se puede Tener Acceso a  " & UCase(DrvUnidad.Drive) & "\" + Chr(13) + Chr(13) + "El Dispositivo no Esta Listo", vbInformation, "Aviso"
    End If
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
       Refresco
    End If
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then
    Unload Me
End If
End Sub
Private Sub Form_Load()
Dim x As Printer
Dim i As Integer
'Me.Width = 5235
'Me.Height = 4500
On Error GoTo ErrorArchivos
   Set fs = New Scripting.FileSystemObject
    
   Ruta = App.path & "\Spooler"
   Me.DrvUnidad.Drive = Mid(App.path, 1, 2)
   DirCarpetas.path = Ruta
   
   If fs.FolderExists(Ruta) = False Then
        fs.CreateFolder (Ruta)
   End If
   Me.Caption = "Spooler de Archivos"
   
   InHabilitar
   Set flCurrent = fs.GetFolder(Ruta)
   Refresco
    
  Exit Sub
ErrorArchivos:
    Select Case Err.Number
        Case 76
            MsgBox "Archivo de Spooler no Existe consulte a su Administrador", vbInformation, "Aviso"
            Exit Sub
        Case Else
            MsgBox "Error :[" & Err.Number & "] " & Err.Description, vbInformation, "Aviso"
            Exit Sub
        
    End Select
    

End Sub
Private Sub Habilitar()
'PREVIO
fraPrevio.Visible = True
fraZoom.Visible = True
cboPrevio.ListIndex = 3
Me.fraPagina.Visible = True
Me.pgrBarra.Visible = False
Me.fraArchivos.Visible = False
Me.BarraEstado.Visible = False
Me.barra.Buttons(3).Enabled = False
Me.barra.Buttons(6).Enabled = False
Me.barra.Buttons(7).Enabled = False
Me.barra.Buttons(8).Enabled = False
Me.barra.Buttons(9).Enabled = False
barra.Buttons(1).Visible = False
barra.Buttons(2).Visible = True
Me.DrvUnidad.Enabled = False
'Me.Height = 7755
'Me.Width = 12030
'CentraSdi Me
End Sub
Private Sub InHabilitar()
'NOPREVIO
fraPrevio.Visible = False
fraZoom.Visible = False
Me.fraPagina.Visible = False
Me.BarraEstado.Visible = True
Me.pgrBarra.Visible = True
Me.BarraEstado.Panels(1).Text = "Spooler de Impresión"
Me.fraArchivos.Visible = True
Me.barra.Buttons(3).Enabled = True
Me.barra.Buttons(6).Enabled = True
Me.barra.Buttons(7).Enabled = True
Me.barra.Buttons(8).Enabled = True
Me.barra.Buttons(9).Enabled = True
Me.DrvUnidad.Enabled = True
'Me.Height = 7755
'Me.Width = 12030
'CentraSdi Me
barra.Buttons(1).Visible = True
barra.Buttons(2).Visible = False
End Sub
Sub Refresco()
    lstArch.ColumnHeaders.Item(2).Width = 700
    lstArch.ColumnHeaders.Item(3).Width = 2000
    lstArch.ColumnHeaders.Item(4).Width = 2000
    lstArch.Icons = ImageList2
    
    lstArch.ListItems.Clear
    For Each fi In flCurrent.Files
      If UCase(Right(fi.Name, 3)) = "TXT" Or UCase(Right(fi.Name, 3)) = "TXT" Then
        Set xt = lstArch.ListItems.Add(, , fi.Name, "texto", "texto")
        'xt.SmallIcon = 2
        xt.SubItems(1) = Format(fi.Size / 1024, "#,#0.00") & " KB"
        xt.SubItems(2) = fi.Type
        xt.SubItems(3) = fi.DateLastModified
        'xt.SubItems(4) = fi.Path
      End If
    Next
End Sub
Private Sub CargaRTF()
Dim fi As Scripting.File
Dim txt As Scripting.TextStream
Dim posicion As Long
Dim vRelle As String
Dim BuferArchivo
Dim MiRegistro As Long
Dim TamañoMax, NúmeroRegistro As Long
Dim Tamaño As Long
Dim lsCadenaPrint As String
Dim lsCadenaPrint1 As String

Me.BarraEstado.Panels(1).Text = "Cargando  Archivo para Vista Preliminar Por Favor Espere..."
Screen.MousePointer = 11
'On Error GoTo ErrHandler
If lstArch.ListItems.Count <> 0 Then
    Set fi = flCurrent.Files(lstArch.SelectedItem)
    If UCase(Right(fi.Name, 3)) = "TXT" Or UCase(Right(fi.Name, 3)) = "RTF" Then
        'Set Txt = fs.OpenTextFile(fi.Path, ForReading, False)
        rtfView.LoadFile fi.path
        rtfAux.LoadFile fi.path
        posicion = InStr(1, Me.rtfView.Text, "T,66")
        If posicion = 0 Then
            posicion = InStr(1, Me.rtfView.Text, "F,66")
            If posicion <> 0 Then
                rtfView.Text = Mid(rtfView.Text, 1, posicion - 1)
                parametros = Trim(Right(Trim(rtfAux.Text), 6))
                ExtraeParam
            End If
        Else
            rtfView.Text = Mid(rtfView.Text, 1, posicion - 1)
            parametros = Trim(Right(Trim(rtfAux.Text), 6))
            ExtraeParam
        End If
        rtfView.SelFontSize = 7.5
        rtfView.Font = "Courier New"
    End If
    If Condensado Then
        chkCondensado.Value = 1
    Else
        chkCondensado.Value = 0
    End If
    
    
    vRelle = String(Round(rtfView.RightMargin / 155), "»") & oImpresora.gPrnSaltoLinea
    vPage = 0
    Dim vTexto As String
    Dim sImpre As String
    Dim sPag   As String
    Dim N      As Currency
    
    vTexto = rtfView.Text
    Do While True And Len(Trim(vTexto)) > 0
       vPage = vPage + 1
       N = InStr(1, vTexto, oImpresora.gPrnSaltoPagina)
       If N > 0 Then
          sPag = sPag & Mid(vTexto, 1, N - 1) & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & "<<Página :" & ImpreFormat(vPage, 7, 0) & ">>" & oImpresora.gPrnSaltoLinea & vRelle
          vTexto = Mid(vTexto, N + 1)
       Else
           sPag = sPag & vTexto & oImpresora.gPrnSaltoLinea & oImpresora.gPrnSaltoLinea & "<<Página :" & ImpreFormat(vPage, 7, 0) & ">>" & oImpresora.gPrnSaltoLinea
           Exit Do
       End If
       If vPage Mod 30 = 0 Then
           sImpre = sImpre & sPag
           sPag = ""
       End If
    Loop
    'rtfTexto.Text = sImpre & sPag
    
    'Do While True And Len(Trim(rtfView.Text)) > 0
    '    vPage = vPage + 1
    '    If InStr(1, rtfView.Text, oImpresora.gPrnSaltoPagina ) > 0 Then
    '        rtfView.Text = Replace(rtfView.Text, oImpresora.gPrnSaltoPagina , oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & "<<Página :" & ImpreFormat(vPage, 7, 0) & ">>" & oImpresora.gPrnSaltoLinea  & vRelle, 1, 1, vbTextCompare)
    '    Else
    '        rtfView.Text = rtfView.Text & oImpresora.gPrnSaltoLinea  & oImpresora.gPrnSaltoLinea  & "<<Página :" & ImpreFormat(vPage, 7, 0) & ">>" & oImpresora.gPrnSaltoLinea
    '        Exit Do
    '    End If
    'Loop
    
    rtfView.Text = sImpre & sPag
    
    txtPagIni.Text = 1
    txtPagFin.Text = vPage
Else
    'MsgBox "No existen Archivos para Seleccionr", vbInformation, "Aviso"
    lstArch.ListItems.Clear
End If
Screen.MousePointer = 0
Me.BarraEstado.Panels(1).Text = "Spooler de Impresion"
Exit Sub
ErrHandler:
    Screen.MousePointer = 0
    Me.BarraEstado.Panels(1).Text = "Spooler de Impresion"
    Select Case Err.Number
        Case 7 ' fuera de memoria
            MsgBox "Archivo es muy grande para abrirlo", vbInformation, "Error"
        Case Else
            MsgBox Err.Description & " [" & Err.Number & "]", vbInformation, "Error"
    End Select
End Sub
Private Sub lstArch_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
  lstArch.Sorted = False
  lstArch.SortKey = ColumnHeader.Index - 1
  
  If lstArch.SortOrder = lvwAscending Then
    lstArch.SortOrder = lvwDescending
  Else
    lstArch.SortOrder = lvwAscending
  End If
   ' Asigna a Sorted el valor True para ordenar la lista.
  lstArch.Sorted = True
End Sub
Private Sub lstArch_DblClick()
If lstArch.ListItems.Count <> 0 Then
    CargaRTF
    Habilitar
End If
End Sub


Private Sub lstArch_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If lstArch.ListItems.Count <> 0 Then
        CargaRTF
        Habilitar
    End If
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
Public Function Buscatipo(psNombreprint As String) As String
Dim indice As Integer
    indice = InStr(3, Trim(psNombreprint), "\")
    If indice <> 0 Then
        Buscatipo = Mid(psNombreprint, indice + 1, Len(Trim(psNombreprint)))
    Else
        Buscatipo = Trim(psNombreprint)
    End If
End Function
Private Sub mnuOpcAnterior_Click()
rtfView.SetFocus
SendKeys "{PGUP}", True
End Sub

Private Sub mnuOpcBuscar_Click()

    frmPrevioBus.CmdCancelar.Enabled = True
    frmPrevioBus.Show 1
    If frmPrevioBus.cboDireccion.Text = "Todo" Then
        rtfView.SelStart = 0
    End If
    vTipBus = 0 ' rtfNoHighlight
    If frmPrevioBus.chkOpc1.Value = 1 Then
        vTipBus = 2
    End If
    If frmPrevioBus.chkOpc2.Value = 1 Then
        vTipBus = vTipBus + rtfMatchCase
    End If
    
    vPosCad = rtfView.Find(frmPrevioBus.cboBuscar.Text, rtfView.SelStart + Len(frmPrevioBus.cboBuscar.Text), , vTipBus)
    If frmPrevioBus.CmdCancelar.Enabled <> False And vPosCad = -1 Then
        MsgBox "Cadena no encontrada", vbInformation, " Aviso "
    End If
End Sub

Private Sub mnuOpcBusSig_Click()

If frmPrevioBus.cboBuscar.Text <> "" Then
    vPosCad = rtfView.Find(frmPrevioBus.cboBuscar.Text, rtfView.SelStart + Len(frmPrevioBus.cboBuscar.Text), , vTipBus)
    If vPosCad = -1 Then
        MsgBox "Cadena no encontrada", vbInformation, " Aviso "
    End If
Else
    MsgBox "Ingrese cadena a buscar", vbInformation, " Aviso "
End If
End Sub

Private Sub mnuOpcFinal_Click()
rtfView.SelStart = Len(rtfView.Text)
End Sub

Private Sub mnuOpcInicio_Click()
rtfView.SelStart = 0
End Sub

Private Sub mnuOpcPagina_Click()
Dim vPagina As String
Dim vPosCad As Double
vPagina = InputBox("Ingrese página", " Aviso ")
If Val(vPagina) >= 1 And Val(vPagina) <= vPage Then
    rtfView.SelStart = 0
    If Val(vPagina) <> 1 Then
        vPosCad = rtfView.Find("<<Página :" & ImpreFormat(Val(vPagina) - 1, 7, 0) & ">>", 0, , rtfWholeWord)
        If vPosCad = -1 Then
            MsgBox "Página no encontrada", vbInformation, " Aviso "
        Else
            rtfView.SelStart = rtfView.SelStart
            SendKeys "{PGDN}", True
        End If
    End If
Else
    MsgBox "Página no Reconocida", vbInformation, " Aviso "
End If

End Sub

Private Sub mnuOpcSiguiente_Click()
rtfView.SetFocus
SendKeys "{PGDN}", True
End Sub

Private Sub rtfView_MouseDown(Button As Integer, Shift As Integer, x As Single, Y As Single)
If Button = 2 Then PopupMenu mnuOpciones
End Sub

'Validación de txtPagFin
Private Sub txtPagFin_GotFocus()
fEnfoque txtPagFin
End Sub
Private Sub txtPagFin_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Val(txtPagFin.Text) <= vPage And Val(txtPagFin) >= Val(txtPagIni) Then
        cboPrevio.SetFocus
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
Private Sub CargaRTFArchivo()
Dim ArchivoImp As Integer
Dim fi As Scripting.File
Dim txt As Scripting.TextStream
Dim posicion As Long
Dim vRelle As String
Dim BuferArchivo
Dim MiRegistro As Long
Dim TamañoMax, NúmeroRegistro As Long
Dim Tamaño As Long
Dim lsCadenaPrint As String
Dim lsCadenaPrint1 As String
Dim i As Long

NumLineas = 66
Me.BarraEstado.Panels(1).Text = "Cargando  Archivo para Impresión Por Favor Espere..."
Screen.MousePointer = 0
ArchivoImp = FreeFile
If lstArch.ListItems.Count <> 0 Then
    Set fi = flCurrent.Files(lstArch.SelectedItem)
    If UCase(Right(fi.Name, 3)) = "TXT" Or UCase(Right(fi.Name, 3)) = "RTF" Then
        Open fi.path For Input As #1
        TamañoMax = LOF(1)
        lsCadenaPrint = ""
        Tamaño = 0
        NúmeroRegistro = 1000
        Me.pgrBarra.Max = TamañoMax
        Do While Tamaño < TamañoMax
            Tamaño = Tamaño + NúmeroRegistro
            If (Seek(1) + NúmeroRegistro) > TamañoMax Then
                NúmeroRegistro = TamañoMax - Seek(1)
                Me.pgrBarra.Value = TamañoMax
            Else
                Me.pgrBarra.Value = Tamaño
            End If
            BuferArchivo = Input(NúmeroRegistro, #1)  ' Establece la posición.
            lsCadenaPrint = lsCadenaPrint + BuferArchivo
            If Tamaño Mod 10000 = 0 Then
                lsCadenaPrint1 = lsCadenaPrint1 + lsCadenaPrint
                lsCadenaPrint = ""
            End If
            DoEvents
        Loop
        Close #1   ' Cierra el archivo.
        If Tamaño Mod 10000 <> 0 Then
            lsCadenaPrint1 = lsCadenaPrint1 + lsCadenaPrint
            lsCadenaPrint = ""
        End If
        posicion = InStr(1, lsCadenaPrint1, "T,66")
        If posicion = 0 Then
            posicion = InStr(1, lsCadenaPrint1, "F,66")
            If posicion <> 0 Then
                lsCadenaPrint1 = Mid(lsCadenaPrint1, 1, posicion - 1)
                parametros = Trim(Right(Trim(lsCadenaPrint1), 6))
                ExtraeParam
            End If
        Else
            parametros = Trim(Right(Trim(lsCadenaPrint1), 6))
            ExtraeParam
            lsCadenaPrint1 = Mid(lsCadenaPrint1, 1, posicion - 1)
        End If
        If Condensado Then
            chkCondensado.Value = 1
        Else
            chkCondensado.Value = 0
        End If
        frmImpresora.Show 1
        If lbCancela = False Then
            For i = 1 To lnNumCopias
                ImpreBegin Condensado, NumLineas
                    Print #ArcSal, ImpreCarEsp(lsCadenaPrint1)
                ImpreEnd
            Next i
        End If
        Me.pgrBarra.Value = 0
    End If
End If
End Sub

