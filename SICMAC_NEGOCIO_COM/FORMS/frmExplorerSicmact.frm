VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmExplorerSicmact 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Explorador de Archivos SICMACT"
   ClientHeight    =   6360
   ClientLeft      =   465
   ClientTop       =   1875
   ClientWidth     =   11160
   Icon            =   "frmExplorerSicmact.frx":0000
   LinkTopic       =   "Form2"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   11160
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.ImageList ImageArchivo1 
      Left            =   8280
      Top             =   375
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   41
      ImageHeight     =   39
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorerSicmact.frx":030A
            Key             =   "excel"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageArchivo 
      Left            =   6870
      Top             =   210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorerSicmact.frx":1640
            Key             =   "excel"
            Object.Tag             =   "excel"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorerSicmact.frx":1A92
            Key             =   "word"
            Object.Tag             =   "word"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6015
      Top             =   225
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
            Picture         =   "frmExplorerSicmact.frx":1EE4
            Key             =   "grandes"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorerSicmact.frx":1FF6
            Key             =   "detalles"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorerSicmact.frx":2108
            Key             =   "lista"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorerSicmact.frx":221A
            Key             =   "pequeños"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorerSicmact.frx":232C
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorerSicmact.frx":286E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorerSicmact.frx":2980
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorerSicmact.frx":2A92
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmExplorerSicmact.frx":2BEC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar barra 
      Align           =   1  'Align Top
      Height          =   390
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   688
      ButtonWidth     =   609
      ButtonHeight    =   582
      ImageList       =   "ImageList1"
      DisabledImageList=   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   5
            Style           =   4
            Object.Width           =   3200
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Key             =   "noprevio"
            Object.ToolTipText     =   "Cerrar Vista Preliminar"
            ImageIndex      =   9
            Style           =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "eliminar"
            Object.ToolTipText     =   "Eliminar Archivo"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "grande"
            Object.ToolTipText     =   "Iconos Grandes"
            ImageIndex      =   1
            Style           =   2
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "pequeños"
            Object.ToolTipText     =   "Iconos Pequeños"
            ImageIndex      =   4
            Style           =   2
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "listas"
            Object.ToolTipText     =   "Lista"
            ImageIndex      =   3
            Style           =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "detalles"
            Object.ToolTipText     =   "Detalles"
            ImageIndex      =   2
            Style           =   2
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "salir"
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   8
         EndProperty
      EndProperty
      Begin VB.DriveListBox DrvUnidad 
         Height          =   315
         Left            =   45
         TabIndex        =   6
         Top             =   30
         Width           =   3060
      End
   End
   Begin MSComctlLib.StatusBar BarDirectorio 
      Height          =   315
      Left            =   30
      TabIndex        =   3
      Top             =   390
      Width           =   2985
      _ExtentX        =   5265
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lstArchivos 
      Height          =   5295
      Left            =   3045
      TabIndex        =   2
      Top             =   720
      Width           =   8040
      _ExtentX        =   14182
      _ExtentY        =   9340
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      Icons           =   "ImageArchivo1"
      SmallIcons      =   "ImageArchivo"
      ColHdrIcons     =   "ImageArchivo"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Nombre"
         Object.Width           =   5292
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Tamaño"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Tipo"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Modificado"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Atributos"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSComctlLib.StatusBar BarraPrinc 
      Align           =   2  'Align Bottom
      Height          =   315
      Left            =   0
      TabIndex        =   1
      Top             =   6045
      Width           =   11160
      _ExtentX        =   19685
      _ExtentY        =   556
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14993
            MinWidth        =   14993
         EndProperty
      EndProperty
   End
   Begin VB.DirListBox dirExplorer 
      Height          =   5265
      Left            =   0
      TabIndex        =   0
      Top             =   720
      Width           =   3000
   End
   Begin MSComctlLib.StatusBar BarraArchivos 
      Height          =   330
      Left            =   3075
      TabIndex        =   4
      Top             =   390
      Width           =   7980
      _ExtentX        =   14076
      _ExtentY        =   582
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   14993
            MinWidth        =   14993
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuOpciones 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuVer 
         Caption         =   "&Ver"
         Begin VB.Menu mnuIcongrandes 
            Caption         =   "Iconos &Grandes"
         End
         Begin VB.Menu mnuIconsmall 
            Caption         =   "Iconos &Pequeños"
         End
         Begin VB.Menu mnuLista 
            Caption         =   "&Lista"
         End
         Begin VB.Menu mnuDetalle 
            Caption         =   "&Detalle"
         End
      End
      Begin VB.Menu mnuGuion 
         Caption         =   "-"
      End
      Begin VB.Menu mnueliminar 
         Caption         =   "&Eliminar"
      End
   End
End
Attribute VB_Name = "frmExplorerSicmact"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private drCurrent As Scripting.Drive
Private flCurrent As Scripting.Folder
Private fs As Scripting.FileSystemObject
Dim fi As Scripting.File
Dim Ruta As String
Dim lbLoad As Boolean
Private Sub barra_ButtonClick(ByVal Button As MSComctlLib.Button)
Select Case Button.Key
    Case "grande"
        If lstArchivos.ListItems.Count <> 0 Then
           lstArchivos.View = lvwIcon
        End If
    Case "pequeños"
        If lstArchivos.ListItems.Count <> 0 Then
            lstArchivos.View = lvwSmallIcon
        End If
    Case "listas"
        If Me.lstArchivos.ListItems.Count <> 0 Then
            lstArchivos.View = lvwList
        End If
    Case "detalles"
        If Me.lstArchivos.ListItems.Count <> 0 Then
            lstArchivos.View = lvwReport
        End If
    Case "salir"
        Unload Me
    Case "eliminar"
        If lstArchivos.ListItems.Count <> 0 Then
            If MsgBox("Desea eliminar el Archivo : " & lstArchivos.SelectedItem & " ", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                fs.DeleteFile (Ruta & "\" & lstArchivos.SelectedItem)
                Refresco
            Else
                lstArchivos.SetFocus
            End If
       End If
End Select
End Sub

Private Sub dirExplorer_Change()
Ruta = dirExplorer
Set flCurrent = fs.GetFolder(Ruta)
Refresco
BarDirectorio.Panels(1) = dirExplorer
Me.BarraArchivos.Panels(1).Text = "Contenido de " & dirExplorer
End Sub

Private Sub dirExplorer_Click()
Ruta = dirExplorer
Set flCurrent = fs.GetFolder(Ruta)
Refresco
BarDirectorio.Panels(1) = dirExplorer
Me.BarraArchivos.Panels(1).Text = "Contenido de " & dirExplorer
End Sub

Private Sub dirExplorer_KeyPress(KeyAscii As Integer)
'If KeyAscii = 13 Then
'    Ruta = dirExplorer
'    Set flCurrent = fs.GetFolder(Ruta)
'    Refresco
'End If
End Sub

Private Sub DrvUnidad_Change()
Set fs = New Scripting.FileSystemObject

    If fs.Drives(DrvUnidad.Drive).IsReady Then
        If lbLoad Then
            Me.dirExplorer.Path = Ruta
        Else
            Me.dirExplorer.Path = DrvUnidad.Drive
        End If
    Else
        MsgBox "No se puede Tener Acceso a  " & UCase(DrvUnidad.Drive) & "\" + Chr(13) + Chr(13) + "El Dispositivo no Esta Listo", vbInformation, "Aviso"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = vbKeyF5 Then
       Refresco
    End If
End Sub

Private Sub Form_Load()
CentraSdi Me
Set fs = New Scripting.FileSystemObject
lbLoad = True
Ruta = App.Path & "\Spooler"
DrvUnidad.Drive = Mid(Ruta, 1, 2)
lbLoad = False
End Sub
Private Sub Refresco()
Dim xt As ListItem
    lstArchivos.ListItems.Clear
    For Each fi In flCurrent.Files
      If UCase(Right(fi.Name, 3)) = "XLS" Or UCase(Right(fi.Name, 3)) = "XLS" Then
        Set xt = lstArchivos.ListItems.Add(, , fi.Name, "excel", "excel")
        xt.SubItems(1) = Format(fi.Size / 1024, "#,#0.00") & " KB"
        xt.SubItems(2) = fi.Type
        xt.SubItems(3) = fi.DateLastModified
        xt.SubItems(4) = fi.Attributes
      End If
    Next
End Sub
Private Sub CargaArchivo(lsArchivo As String, lsRutaArchivo As String)
Dim x As Long
Dim Temp As String
    Temp = GetActiveWindow()
    x = ShellExecute(Temp, "open", lsArchivo, "", lsRutaArchivo, 1)
    If x <= 32 Then
        If x = 2 Then
            MsgBox "No se encuentra el Archivo adjunto, " & vbCr & " verifique el servidor de archivos", vbInformation, " Aviso "
        ElseIf x = 8 Then
            MsgBox "Memoria insuficiente ", vbInformation, " Aviso "
        Else
            MsgBox "No se pudo abrir el Archivo adjunto", vbInformation, " Aviso "
        End If
    End If

End Sub
Private Sub lstArchivos_DblClick()
If Me.lstArchivos.ListItems.Count > 0 Then
    CargaArchivo Me.lstArchivos.SelectedItem, dirExplorer
End If
End Sub

Private Sub lstArchivos_ItemClick(ByVal Item As MSComctlLib.ListItem)
If Me.lstArchivos.ListItems.Count > 0 Then
    Me.BarraPrinc.Panels(1).Text = Item
    Me.BarraPrinc.Panels(2).Text = Item.SubItems(1)
End If
End Sub


Private Sub lstArchivos_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    If Me.lstArchivos.ListItems.Count > 0 Then
        CargaArchivo Me.lstArchivos.SelectedItem, dirExplorer
    End If
End If
End Sub

Private Sub lstArchivos_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then PopupMenu mnuOpciones
End Sub

Private Sub mnuDetalle_Click()
    If Me.lstArchivos.ListItems.Count <> 0 Then
        lstArchivos.View = lvwReport
        mnuIcongrandes.Checked = False
        mnuIconsmall.Checked = False
        mnuLista.Checked = False
        mnuDetalle.Checked = True
    End If
End Sub

Private Sub mnueliminar_Click()
If lstArchivos.ListItems.Count <> 0 Then
    If MsgBox("Desea eliminar el Archivo : " & lstArchivos.SelectedItem & " ", vbYesNo + vbQuestion, "Aviso") = vbYes Then
        fs.DeleteFile (Ruta & "\" & lstArchivos.SelectedItem)
        Refresco
    Else
        lstArchivos.SetFocus
    End If
End If
End Sub
Private Sub mnuIcongrandes_Click()
    If lstArchivos.ListItems.Count <> 0 Then
        lstArchivos.View = lvwIcon
        mnuIcongrandes.Checked = True
        mnuIconsmall.Checked = False
        mnuLista.Checked = False
        mnuDetalle.Checked = False
    End If
End Sub
Private Sub mnuIconsmall_Click()
    If lstArchivos.ListItems.Count <> 0 Then
        lstArchivos.View = lvwSmallIcon
        
        mnuIcongrandes.Checked = False
        mnuIconsmall.Checked = True
        mnuLista.Checked = False
        mnuDetalle.Checked = False
    End If
End Sub

Private Sub mnuLista_Click()
    If Me.lstArchivos.ListItems.Count <> 0 Then
        lstArchivos.View = lvwList
        mnuLista.Checked = True
        mnuIcongrandes.Checked = False
        mnuIconsmall.Checked = False
        mnuDetalle.Checked = False
    End If
End Sub
