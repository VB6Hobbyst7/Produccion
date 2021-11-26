VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMntContribuyeNoHabido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contribuyentes no Habidos: Mantenimiento"
   ClientHeight    =   3945
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   Icon            =   "frmMntContribuyeNoHabido.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3945
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid dg 
      Height          =   3135
      Left            =   90
      TabIndex        =   0
      Top             =   120
      Width           =   11355
      _ExtentX        =   20029
      _ExtentY        =   5530
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   2
      RowHeight       =   19
      FormatLocked    =   -1  'True
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   4
      BeginProperty Column00 
         DataField       =   "cRuc"
         Caption         =   "RUC"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   "cPersNombre"
         Caption         =   "RAZON SOCIAL/NOMBRES"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   "#,##0.00000000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column02 
         DataField       =   "cPersCod"
         Caption         =   "Codigo"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column03 
         DataField       =   "cMotivo"
         Caption         =   "MOTIVO"
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         MarqueeStyle    =   3
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         BeginProperty Column00 
            Alignment       =   2
            ColumnAllowSizing=   0   'False
            Locked          =   -1  'True
            ColumnWidth     =   2025.071
         EndProperty
         BeginProperty Column01 
            Locked          =   -1  'True
            ColumnWidth     =   4575.118
         EndProperty
         BeginProperty Column02 
            Alignment       =   2
            Locked          =   -1  'True
            ColumnWidth     =   0
         EndProperty
         BeginProperty Column03 
            Locked          =   -1  'True
            ColumnWidth     =   4185.071
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraControl 
      Height          =   555
      Left            =   90
      TabIndex        =   9
      Top             =   3210
      Width           =   11355
      Begin VB.CheckBox chkEliminar 
         Caption         =   "Eliminar los Existetes ?"
         Height          =   270
         Left            =   3345
         TabIndex        =   12
         Top             =   195
         Width           =   2055
      End
      Begin VB.CommandButton cmdImportar 
         Caption         =   "&Importar"
         Height          =   360
         Left            =   8880
         TabIndex        =   7
         Top             =   150
         Width           =   1065
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   360
         Left            =   120
         TabIndex        =   4
         Top             =   150
         Visible         =   0   'False
         Width           =   1065
      End
      Begin VB.CommandButton cmdSalir 
         Cancel          =   -1  'True
         Caption         =   "&Salir"
         CausesValidation=   0   'False
         Height          =   360
         Left            =   9990
         TabIndex        =   8
         Top             =   150
         Width           =   1215
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "&Eliminar"
         Height          =   360
         Left            =   1200
         TabIndex        =   6
         Top             =   150
         Width           =   1065
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   360
         Left            =   120
         TabIndex        =   10
         Top             =   150
         Width           =   1065
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         CausesValidation=   0   'False
         Height          =   360
         Left            =   1200
         TabIndex        =   5
         Top             =   150
         Visible         =   0   'False
         Width           =   1065
      End
   End
   Begin VB.Frame fraDatos 
      Height          =   585
      Left            =   90
      TabIndex        =   11
      Top             =   2610
      Visible         =   0   'False
      Width           =   11355
      Begin VB.TextBox txtMotivo 
         Appearance      =   0  'Flat
         Height          =   345
         Left            =   6870
         MaxLength       =   100
         TabIndex        =   3
         Tag             =   "txtnombre"
         Top             =   180
         Width           =   4335
      End
      Begin VB.TextBox lblProvNombre 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   345
         Left            =   2370
         TabIndex        =   2
         Tag             =   "txtnombre"
         Top             =   180
         Width           =   4485
      End
      Begin Sicmact.TxtBuscar txtBuscarProv 
         Height          =   360
         Left            =   240
         TabIndex        =   1
         Top             =   180
         Width           =   2115
         _ExtentX        =   3731
         _ExtentY        =   635
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         TipoBusPers     =   2
      End
   End
   Begin MSComDlg.CommonDialog Dialog 
      Left            =   4740
      Top             =   3390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "frmMntContribuyeNoHabido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sSql As String
Dim rs   As ADODB.Recordset
Dim lNuevo As Boolean
Dim oCon As DConecta
Dim oBarra  As clsProgressBar

Dim xlAplicacion As Excel.Application
Dim xlLibro As Excel.Workbook
Dim xlHoja1 As Excel.Worksheet

Private Sub MuestraBotones(plActiva As Boolean)
fraDatos.Visible = plActiva
cmdAceptar.Visible = plActiva
cmdCancelar.Visible = plActiva
cmdNuevo.Visible = Not plActiva
cmdEliminar.Visible = Not plActiva
If plActiva Then
   dg.Height = dg.Height - 600
Else
   dg.Height = dg.Height + 600
End If
End Sub

Private Function ValidaDatos() As Boolean
ValidaDatos = False
If Trim(txtBuscarProv) = "" Then
   MsgBox "Proveedor no válido", vbInformation, "¡Aviso!"
   txtBuscarProv.SetFocus
   Exit Function
End If
ValidaDatos = True
End Function

Private Sub cmdAceptar_Click()
Dim nPos
If Not ValidaDatos() Then
   Exit Sub
End If
If MsgBox(" ¿ Seguro que desea Grabar datos ? ", vbQuestion + vbYesNo, "¡Confirmación!") = vbNo Then
   Exit Sub
End If
If lNuevo Then
   sSql = "INSERT PersContribuyeNoHabido (cPersCod, cRuc, cPersNombre, cMotivo ) " _
        & "VALUES ('" & Me.txtBuscarProv.psCodigoPersona & "', '" & txtBuscarProv.Text & "','" & lblProvNombre.Text & "','" & Me.txtMotivo & "' )"
   oCon.Ejecutar sSql
End If
MuestraBotones False
CargaDatos
rs.Find "cRuc = '" & txtBuscarProv.Text & "'", , adSearchForward, 0
dg.SetFocus
End Sub

Private Sub cmdCancelar_Click()
MuestraBotones False
dg.SetFocus
End Sub

Private Sub cmdEliminar_Click()
Dim nPos
If rs Is Nothing Then
   Exit Sub
End If
If rs.EOF Then
   Exit Sub
End If

If MsgBox(" ¿ Seguro que desea Eliminar datos ? ", vbQuestion + vbYesNo, "¡Confirmación") = vbYes Then
   nPos = rs.Bookmark
   sSql = "DELETE PersContribuyeNoHabido WHERE cRUC = '" & rs!cRuc & "'"
   oCon.Ejecutar sSql
   
   CargaDatos
   If nPos > rs.Bookmark Then
      rs.MoveLast
   Else
      rs.Bookmark = nPos
   End If
   rs.MoveLast
   dg.SetFocus
End If
End Sub

'Private Sub cmdExportar_Click()
'Dim oArch As Object
'Dim lsArchivo As String
'Dim I As Integer, nFil As Integer
'Dim lbExcel As Boolean
'On Error GoTo ImportarErr
'lsArchivo = ""
'Dialog.FileName = ""
'Dialog.DialogTitle = "CONTRIBUYENTES NO HABIDOS: Importar Archivo"
'Dialog.FLAGS = 21
'Dialog.Filter = "*.xls"
'Dialog.DefaultExt = "*.xls"
'Dialog.ShowOpen
'
'
'lsArchivo = Dialog.FileName
'If lsArchivo = "" Then
'    MsgBox "Debe seleccionar un archivo Excel para Importar datos", vbInformation, "¡Aviso!"
'Else
'    lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
'    If lbExcel Then
'        Set xlHoja1 = xlLibro.Worksheets(1)
'        MousePointer = 11
'        sSql = "DELETE PersContribuyeNoHabido "
'        oCon.Ejecutar sSql
'         nFil = 6
'         Do While True
'             If xlHoja1.Cells(nFil, 1) = "" Then Exit Do
'             nFil = nFil + 1
'         Loop
'         Set oBarra = New clsProgressBar
'         BarraShow nFil
'         nFil = 6
'         Do While True
'             sSql = "INSERT PersContribuyeNoHabido (cPersCod, cRuc, cPersNombre, cMotivo ) " _
'                  & "VALUES ('', '" & xlHoja1.Cells(nFil, 1) & "','" & Replace(xlHoja1.Cells(nFil, 2), "'", "") & "','" & xlHoja1.Cells(nFil, 3) & "' )"
'             oCon.Ejecutar sSql
'             nFil = nFil + 1
'             BarraProgress nFil - 6, "CONTRIBUYENTES NO HABIDOS", "Importación de Datos", "", vbBlue
'             If xlHoja1.Cells(nFil, 1) = "" Then Exit Do
'         Loop
'        MousePointer = 0
'        BarraClose
'        MsgBox "Archivo " & lsArchivo & " importado Satisfactoriamente"
'    End If
'    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
'   CargaDatos
'
'End If
'Exit Sub
'ImportarErr:
'    MsgBox "Existieron problemas al Importar Archivo" & Chr(10) & TextErr(Err.Description), vbInformation, "¡Aviso!"
'    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
'End Sub

Private Sub cmdImpuesto_Click()
    frmRentaCuart.Show
End Sub

Private Sub cmdImportar_Click()
Dim oArch As Object
Dim lsArchivo As String
Dim I As Integer, nFil As Integer
Dim lbExcel As Boolean
On Error GoTo ImportarErr
lsArchivo = ""
Dialog.FileName = ""
Dialog.DialogTitle = "CONTRIBUYENTES NO HABIDOS: Importar Archivo"
Dialog.Flags = 21
Dialog.Filter = "*.xls"
Dialog.DefaultExt = "*.xls"
Dialog.ShowOpen


lsArchivo = Dialog.FileName
If lsArchivo = "" Then
    MsgBox "Debe seleccionar un archivo Excel para Importar datos", vbInformation, "¡Aviso!"
Else
    If Me.chkEliminar.value = 1 Then
        sSql = "Delete PersContribuyeNoHabido "
        oCon.Ejecutar sSql
    End If
    
    lbExcel = ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
    If lbExcel Then
        Set xlHoja1 = xlLibro.Worksheets(1)
        MousePointer = 11
        'sSql = "DELETE PersContribuyeNoHabido "
        'oCon.Ejecutar sSql
         nFil = 6
         Do While True
             If xlHoja1.Cells(nFil, 1) = "" Then Exit Do
             nFil = nFil + 1
         Loop
         Set oBarra = New clsProgressBar
         BarraShow nFil
         nFil = 6
         Do While True
             sSql = "INSERT PersContribuyeNoHabido (cPersCod, cRuc, cPersNombre, cMotivo ) " _
                  & "VALUES ('', '" & xlHoja1.Cells(nFil, 1) & "','" & Replace(xlHoja1.Cells(nFil, 2), "'", "") & "','" & xlHoja1.Cells(nFil, 3) & "' )"
             oCon.Ejecutar sSql
             nFil = nFil + 1
             BarraProgress nFil - 6, "CONTRIBUYENTES NO HABIDOS", "Importación de Datos", "", vbBlue
             If xlHoja1.Cells(nFil, 1) = "" Then Exit Do
         Loop
        MousePointer = 0
        BarraClose
        MsgBox "Archivo " & lsArchivo & " importado Satisfactoriamente", vbInformation, "Aviso"
    End If
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
   CargaDatos
    
End If
Exit Sub
ImportarErr:
    MsgBox "Existieron problemas al Importar Archivo" & Chr(10) & TextErr(Err.Description), vbInformation, "¡Aviso!"
    ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1, False
End Sub

Private Sub cmdNuevo_Click()
lNuevo = True
MuestraBotones True
Me.txtBuscarProv.SetFocus
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub Form_Load()
CentraForm Me
Set oCon = New DConecta
oCon.AbreConexion
CargaDatos
If Not rs.EOF Then
   rs.MoveLast
End If
End Sub

Private Sub CargaDatos()
sSql = "SELECT cRuc, cPersNombre, cPersCod, cMotivo FROM PersContribuyeNoHabido"
Set rs = oCon.CargaRecordSet(sSql)
Set dg.DataSource = rs
End Sub

Private Sub Form_Unload(Cancel As Integer)
oCon.CierraConexion
Set oCon = Nothing
End Sub

Private Sub txtBuscarProv_EmiteDatos()
   lblProvNombre = txtBuscarProv.psDescripcion
   If lblProvNombre <> "" Then
      txtMotivo.SetFocus
   End If
End Sub

Private Sub txtMotivo_KeyPress(KeyAscii As Integer)
KeyAscii = Letras(KeyAscii, True)
If KeyAscii = 13 Then
   cmdAceptar.SetFocus
End If
End Sub

Private Sub BarraClose()
oBarra.CloseForm Me
Set oBarra = Nothing
End Sub

Private Sub BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub BarraShow(pnMax As Variant)
Set oBarra = New clsProgressBar
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.ShowForm Me
oBarra.Max = pnMax
End Sub
