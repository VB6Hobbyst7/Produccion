VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMntCtasContBase 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cuentas Contables: Cuentas Base de SBS"
   ClientHeight    =   5175
   ClientLeft      =   510
   ClientTop       =   3105
   ClientWidth     =   10350
   Icon            =   "frmMntCtasContBase.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   10350
   Begin VB.CommandButton cmdExcel 
      Caption         =   "EXCEL"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   120
      TabIndex        =   19
      Top             =   4255
      Width           =   1395
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9000
      TabIndex        =   6
      Top             =   3780
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid grdCtas 
      Height          =   4005
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   8715
      _ExtentX        =   15372
      _ExtentY        =   7064
      _Version        =   393216
      AllowUpdate     =   0   'False
      AllowArrows     =   -1  'True
      HeadLines       =   2
      RowHeight       =   17
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
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   "cCtaContCod"
         Caption         =   "Código"
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
      BeginProperty Column01 
         DataField       =   "cCtaContDesc"
         Caption         =   "Descripción"
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
         ScrollBars      =   2
         AllowRowSizing  =   0   'False
         AllowSizing     =   0   'False
         Locked          =   -1  'True
         Size            =   2
         BeginProperty Column00 
            DividerStyle    =   6
            ColumnAllowSizing=   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   1950.236
         EndProperty
         BeginProperty Column01 
            DividerStyle    =   6
            ColumnAllowSizing=   -1  'True
            WrapText        =   -1  'True
            ColumnWidth     =   6210.142
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   9000
      Picture         =   "frmMntCtasContBase.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   2250
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   555
      Left            =   9000
      Picture         =   "frmMntCtasContBase.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   9000
      TabIndex        =   4
      Top             =   1830
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      DisabledPicture =   "frmMntCtasContBase.frx":0A76
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   1410
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   990
      Width           =   1215
   End
   Begin VB.Frame fraImprime 
      Caption         =   "Impresión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   2175
      Left            =   8940
      TabIndex        =   7
      Top             =   900
      Visible         =   0   'False
      Width           =   1335
      Begin VB.OptionButton optImpre 
         Caption         =   "&Todo"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   12
         Top             =   300
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton optImpre 
         Caption         =   "&Grupo"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   11
         Top             =   600
         Width           =   915
      End
      Begin VB.TextBox txtCtaCod 
         Height          =   315
         Left            =   60
         TabIndex        =   10
         Top             =   990
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   90
         TabIndex        =   9
         ToolTipText     =   "Cancelar Cambio de Fecha de Asiento"
         Top             =   1800
         Width           =   1155
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         Height          =   315
         Left            =   90
         TabIndex        =   8
         Top             =   1440
         Width           =   1155
      End
   End
   Begin VB.Frame fraDat 
      Height          =   1035
      Left            =   120
      TabIndex        =   13
      Top             =   4080
      Visible         =   0   'False
      Width           =   8745
      Begin VB.CommandButton cmdCancelarC 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   4320
         TabIndex        =   17
         Top             =   600
         Width           =   1275
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "&Guardar"
         Height          =   345
         Left            =   3000
         TabIndex        =   16
         Top             =   600
         Width           =   1275
      End
      Begin VB.TextBox txtCtaDesc 
         Height          =   345
         Left            =   2340
         MaxLength       =   100
         TabIndex        =   15
         Top             =   180
         Width           =   6165
      End
      Begin VB.TextBox txtCtaCod2 
         Height          =   345
         Left            =   300
         MaxLength       =   20
         TabIndex        =   14
         Top             =   180
         Width           =   1995
      End
   End
   Begin MSComctlLib.ProgressBar pgbCtaCont 
      Height          =   180
      Left            =   120
      TabIndex        =   18
      Top             =   4680
      Visible         =   0   'False
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmMntCtasContBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsCta As ADODB.Recordset
Dim clsCtaCont As DCtaCont

Dim EsNuevo As Boolean
Dim sCond
Dim nOrdenCta As Integer 'Para establecer el Orden de la tablas
Dim gvalBusca As String
Dim posIniSeek As Long
Dim lPressMouse As Boolean
Dim lConsulta   As Boolean
Dim sNomTabla   As String

Dim WithEvents oImp As NContImprimir
Attribute oImp.VB_VarHelpID = -1
Dim oBarra As clsProgressBar
Attribute oBarra.VB_VarHelpID = -1

Public Sub Inicio(plConsulta As Boolean, Optional psTabla As String = "")
lConsulta = plConsulta
If psTabla = "CtaCont" Then
   sNomTabla = gsCentralCom & psTabla
Else
   sNomTabla = psTabla
End If
Me.Show 0, frmMdiMain
End Sub

Private Sub ActivaControlVisible(plOpcion As Boolean)
cmdBuscar.Visible = plOpcion
If Not lConsulta Then
   cmdNuevo.Visible = plOpcion
   cmdModificar.Visible = plOpcion
   cmdEliminar.Visible = plOpcion
End If
cmdImprimir.Visible = plOpcion
fraImprime.Visible = Not plOpcion
End Sub

Private Sub ManejaBoton(plOpcion As Boolean)
cmdBuscar.Enabled = plOpcion
If Not lConsulta Then
   cmdNuevo.Enabled = plOpcion
   cmdModificar.Enabled = plOpcion
   cmdEliminar.Enabled = plOpcion
End If
cmdImprimir.Enabled = plOpcion
grdCtas.Enabled = plOpcion
End Sub

Private Sub RefrescaGrid()
Set rsCta = clsCtaCont.CargaCtaCont(, sNomTabla, adLockOptimistic)
Set grdCtas.DataSource = rsCta
End Sub

Private Sub cmdBuscar_Click()
Dim clsBusca As New DescObjeto.ClassDescObjeto
ManejaBoton False
clsBusca.BuscarDato rsCta, nOrdenCta, "Cuenta Contable"
nOrdenCta = clsBusca.gnOrdenBusca
Set clsBusca = Nothing
ManejaBoton True
grdCtas.SetFocus
End Sub

Private Sub cmdCancelar_Click()
ActivaControlVisible True
ManejaBoton True
End Sub

Private Sub cmdCancelarC_Click()
ActivaControlVisible True
ManejaBoton True
ManejaIngresoDatos False
cmdExcel.Visible = True
End Sub

Private Sub cmdEliminar_Click()
Dim dSQL As String
Dim Pos As Variant
On Error GoTo DelError
If Not rsCta.BOF And Not rsCta.EOF Then
   If Not clsCtaCont.CtaInstancia(rsCta!cCtaContCod, sNomTabla) Then
      MsgBox " Existen Cuentas en nivel Inferior ...!", vbExclamation, "Aviso de Eliminación"
      grdCtas.SetFocus
      Exit Sub
   End If
   If MsgBox(" ¿ Seguro de Eliminar Cuenta ? ", vbOKCancel, "Mensaje de Confirmación") = vbOk Then
      
      clsCtaCont.EliminaCtaCont Mid(rsCta!cCtaContCod, 1, 2) & IIf(Len(rsCta!cCtaContCod) > 2, "_", "") & Mid(rsCta!cCtaContCod, 4, 20), sNomTabla
      rsCta.Delete adAffectCurrent
      
   End If
   grdCtas.SetFocus
Else
   MsgBox "No existen Cuentas para eliminar...", vbInformation, "Error de Eliminación"
End If
Exit Sub
DelError:
 MsgBox TextErr(Err.Description), vbExclamation, "Error de Eliminación"
End Sub

Private Sub cmdExcel_Click()
'e = True
imprimirCtaContSBS
End Sub

Public Sub imprimirCtaContSBS()
    Dim fs              As Scripting.FileSystemObject
    Dim xlAplicacion    As Excel.Application
    Dim xlLibro         As Excel.Workbook
    Dim xlHoja1         As Excel.Worksheet
    Dim lbExisteHoja    As Boolean
    Dim liLineas        As Integer
    Dim i               As Integer
    Dim glsArchivo      As String
    Dim lsNomHoja       As String
    'Dim lnIndex         As Integer

    Dim rsTemp As New ADODB.Recordset
    
    Set rsTemp = clsCtaCont.CargaCtaCont("", "CtaContBase", adLockOptimistic)
    'Set RSTEMP = clsCtaCont.CargaCtaCont("SubString(cCtaContCod,3,1)='" & Trim(Str(lnIndex)) & "' or Len(rtrim(cCtaContCod))<3", "CtaCont", adLockOptimistic)
       
    If rsTemp Is Nothing Then
        MsgBox "No exite informacion para imprimir", vbInformation, "Aviso"
        Exit Sub
    End If

    glsArchivo = "Reporte_CuentasContablesSBS" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsArchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsArchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add

    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape

            lbExisteHoja = False
            lsNomHoja = "CuentasContablesSBS"
            For Each xlHoja1 In xlLibro.Worksheets
                If xlHoja1.Name = lsNomHoja Then
                    xlHoja1.Activate
                    lbExisteHoja = True
                    Exit For
                End If
            Next
            If lbExisteHoja = False Then
                Set xlHoja1 = xlLibro.Worksheets.Add
                xlHoja1.Name = lsNomHoja
            End If

            xlAplicacion.Range("A1:A1").ColumnWidth = 2
            xlAplicacion.Range("B1:B1").ColumnWidth = 20 '10
            xlAplicacion.Range("c1:c1").ColumnWidth = 70 '15
          
            xlAplicacion.Range("A1:Z100").Font.Size = 9

            xlHoja1.Cells(1, 1) = gsNomCmac
            xlHoja1.Cells(2, 2) = "L I S T A D O   D E   C U E N T A S   C O N T A B L E S   S B S"
            xlHoja1.Cells(3, 2) = "INFORMACION  AL  " & Format(gdFecSis, "dd/mm/yyyy")

            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 3)).Merge True
            xlHoja1.Range(xlHoja1.Cells(2, 2), xlHoja1.Cells(2, 3)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(3, 2), xlHoja1.Cells(3, 3)).HorizontalAlignment = xlCenter

            liLineas = 6

            xlHoja1.Cells(liLineas, 2) = "Codigo"
            xlHoja1.Cells(liLineas, 3) = "Descripcion"
         

            xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 3)).HorizontalAlignment = xlCenter
            xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 3)).Font.Bold = True
            xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 3)).Borders.LineStyle = 1
            xlHoja1.Range(xlHoja1.Cells(liLineas, 2), xlHoja1.Cells(liLineas, 3)).Interior.Color = RGB(159, 206, 238)
            liLineas = liLineas + 1

         Me.pgbCtaCont.Visible = True
         Me.pgbCtaCont.Min = 0
         Me.pgbCtaCont.Max = rsTemp.RecordCount
         
         Do Until rsTemp.EOF
            xlHoja1.Cells(liLineas, 2) = rsTemp(0)
            xlHoja1.Cells(liLineas, 3) = rsTemp(1)
            liLineas = liLineas + 1
            Me.pgbCtaCont.value = rsTemp.Bookmark
            rsTemp.MoveNext
            
         Loop
         Me.pgbCtaCont.Visible = False

        ExcelCuadro xlHoja1, 2, 6, 3, liLineas - 1
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsArchivo
        ExcelEnd App.path & "\Spooler\" & glsArchivo, xlAplicacion, xlLibro, xlHoja1
        'Cierra el libro de trabajo
        'xlLibro.Close
        ' Cierra Microsoft Excel con el método Quit.
        'xlAplicacion.Quit
        'Libera los objetos.
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsArchivo
        Call CargaArchivo(glsArchivo, App.path & "\SPOOLER\")
End Sub

Private Sub cmdGuardar_Click()
On Error GoTo ErrGuardar
If txtCtaCod2 = "" Then
    MsgBox "Falta Ingresar Cuenta Contable...", vbInformation, "¡Aviso!"
    Exit Sub
End If
If txtCtaDesc = "" Then
    MsgBox "Falta ingresar Descripción de la Cuenta...", vbInformation, "¡Aviso!"
    Exit Sub
End If

If MsgBox(" ¿ Seguro que desea guardar datos ? ", vbQuestion + vbYesNo, "!Confirmación!") = vbNo Then
   Exit Sub
End If
gsMovNro = GeneraMovNroActualiza(gdFecSis, gsCodUser, gsCodCMAC, gsCodAge)
If txtCtaCod2.Enabled Then
   'CtaCont
   clsCtaCont.InsertaCtaCont txtCtaCod2, txtCtaDesc, gsMovNro, gsCentralCom & "CtaCont", Mid(txtCtaCod2, 3, 1)
   'CtaContBase
   clsCtaCont.InsertaCtaCont txtCtaCod2, txtCtaDesc, gsMovNro, sNomTabla, Mid(txtCtaCod2, 3, 1)
Else
   clsCtaCont.ActualizaCtaCont txtCtaCod2, txtCtaDesc, gsMovNro, gsCentralCom & "CtaCont"
   clsCtaCont.ActualizaCtaCont txtCtaCod2, txtCtaDesc, gsMovNro, sNomTabla
End If
RefrescaGrid
rsCta.Find "cCtaContCod = '" & txtCtaCod2 & "'", 0, adSearchForward, 1
ManejaIngresoDatos False
ActivaControlVisible True
ManejaBoton True
Exit Sub
ErrGuardar:
   MsgBox TextErr(Err.Description), vbInformation, "!Aviso!"
   
End Sub

Private Sub cmdImprimir_Click()
ManejaBoton False
ActivaControlVisible False
If optImpre(0).value Then
   optImpre(0).SetFocus
Else
   optImpre(1).SetFocus
End If

End Sub

Private Sub cmdAplicar_Click()
Dim sSql As String
Dim rs   As New ADODB.Recordset
Dim N    As Integer
Dim sTexto As String
Dim sNomFile As String
Me.Enabled = False
If optImpre(0).value Then
   sSql = ""
   sNomFile = "PLANBASE_Todo.txt"
Else
   sSql = "cCtaContCod LIKE '" & txtCtaCod & "%' "
   sNomFile = "PLANBASE_Grupo.txt"
End If
Set rs = clsCtaCont.CargaCtaCont(sSql, sNomTabla, adLockReadOnly)
If rs.EOF Then
   MsgBox "No Existen Cuentas Contables Registradas", vbInformation, "Aviso"
   Exit Sub
End If
MousePointer = 11
   Set oImp = New NContImprimir
   oImp.Inicio gsNomCmac, gsCodAge, Format(gdFecSis, gsFormatoFechaView)
   sTexto = oImp.ImprimePlanContable(rs, sNomTabla, gnLinPage)
   Set oImp = Nothing
   RSClose rs
MousePointer = 0
EnviaPrevio sTexto, "Cuenta Contables : Reporte ", gnLinPage, False
ActivaControlVisible True
ManejaBoton True
Me.Enabled = True
End Sub

Private Sub CmdModificar_Click()
If rsCta.EOF Then
   Exit Sub
End If
ManejaBoton False
ManejaIngresoDatos True
txtCtaCod2 = rsCta(0)
txtCtaDesc = rsCta(1)

txtCtaCod2.Enabled = False
txtCtaDesc.SetFocus
cmdExcel.Visible = False
End Sub
Private Sub cmdNuevo_Click()
Dim Pos As Variant
txtCtaCod2 = ""
txtCtaDesc = ""
ManejaBoton False
ManejaIngresoDatos True
txtCtaCod2.SetFocus
cmdExcel.Visible = False
End Sub

Private Sub ManejaIngresoDatos(plOpcion As Boolean)
If plOpcion Then
   Me.Height = 5630
Else
   Me.Height = 5630 '5805
End If
fraDat.Visible = plOpcion
txtCtaCod2.Enabled = True
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub
Private Sub Form_Activate()
ManejaBoton True
grdCtas.SetFocus
End Sub

Private Sub Form_Load()
lPressMouse = False
frmMdiMain.Enabled = False
Me.Height = 5630 '5805
posIniSeek = 1
nOrdenCta = 0  'Inicialmente las cuentas se Ordenan por Codigo
'Me.Icon = LoadPicture(App.path & gsRutaIcono)
CentraForm Me

Set clsCtaCont = New DCtaCont
Set rsCta = clsCtaCont.CargaCtaCont("", "CtaContBase", adLockOptimistic)

Set grdCtas.DataSource = rsCta
EsNuevo = False
If lConsulta Then
   cmdNuevo.Visible = False
   cmdModificar.Visible = False
   cmdEliminar.Visible = False
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
frmMdiMain.Enabled = True
End Sub

Private Sub grdCtas_GotFocus()
grdCtas.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdCtas_LostFocus()
grdCtas.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub optImpre_Click(Index As Integer)
If Index = 1 Then
   txtCtaCod.Enabled = True
   txtCtaCod.SetFocus
Else
   txtCtaCod.Enabled = False
End If
End Sub

Private Sub txtCtaCod_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   cmdAplicar.SetFocus
End If

End Sub

Private Sub txtCtaCod2_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
   If Mid(txtCtaCod2, 3, 1) <> "0" Then
      MsgBox "La cuenta debe ser del tipo CONSOLIDADA", vbInformation, "!Aviso!"
      Exit Sub
   End If
   txtCtaDesc.SetFocus
End If
End Sub

Private Sub txtCtaDesc_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
If KeyAscii = 13 Then
   cmdGuardar.SetFocus
End If
End Sub

Private Sub oImp_BarraClose()
oBarra.CloseForm Me
End Sub

Private Sub oImp_BarraProgress(value As Variant, psTitulo As String, psSubTitulo As String, psTituloBarra As String, ColorLetras As ColorConstants)
oBarra.Progress value, psTitulo, psSubTitulo, psTituloBarra, ColorLetras
End Sub

Private Sub oImp_BarraShow(pnMax As Variant)
Set oBarra = New clsProgressBar
oBarra.ShowForm Me
oBarra.CaptionSyle = eCap_CaptionPercent
oBarra.Max = pnMax
End Sub


