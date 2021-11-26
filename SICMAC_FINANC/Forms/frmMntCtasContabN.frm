VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form frmMntCtasContabN 
   Caption         =   "Cuentas Contables: Mantenimiento"
   ClientHeight    =   5625
   ClientLeft      =   555
   ClientTop       =   3255
   ClientWidth     =   10350
   Icon            =   "frmMntCtasContabN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5625
   ScaleWidth      =   10350
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdExcelN 
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
      Height          =   375
      Left            =   120
      TabIndex        =   23
      Top             =   4920
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Height          =   3945
      Left            =   90
      TabIndex        =   13
      Top             =   930
      Width           =   8775
      Begin MSDataGridLib.DataGrid grdCtas 
         Height          =   3765
         Left            =   60
         TabIndex        =   14
         Top             =   120
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   6641
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
               ColumnWidth     =   1995.024
            EndProperty
            BeginProperty Column01 
               DividerStyle    =   6
               ColumnAllowSizing=   -1  'True
               WrapText        =   -1  'True
               ColumnWidth     =   6089.953
            EndProperty
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   9000
      TabIndex        =   5
      Top             =   4440
      Width           =   1215
   End
   Begin VB.Frame frmMoneda 
      Height          =   975
      Left            =   90
      TabIndex        =   6
      Top             =   -60
      Width           =   6405
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 0 ]  Dígito Integrador"
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
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   7
         Tag             =   "0"
         Top             =   180
         Value           =   -1  'True
         Width           =   2715
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 6 ] No Monetarias Ajustadas"
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
         Height          =   255
         Index           =   6
         Left            =   3000
         TabIndex        =   12
         Tag             =   "6"
         Top             =   660
         Width           =   3135
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 4 ] De Capital Reajustables"
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
         Height          =   255
         Index           =   4
         Left            =   3000
         TabIndex        =   11
         Tag             =   "4"
         Top             =   420
         Width           =   3285
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 3 ] De Actualización Constante"
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
         Height          =   255
         Index           =   3
         Left            =   3000
         TabIndex        =   10
         Tag             =   "3"
         Top             =   180
         Width           =   3345
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 1 ]  Moneda Naci&onal"
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
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   8
         Tag             =   "1"
         Top             =   420
         Width           =   2595
      End
      Begin VB.OptionButton Moneda 
         Caption         =   "[ 2 ]  Moneda E&xtranjera"
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
         Height          =   255
         Index           =   2
         Left            =   210
         TabIndex        =   9
         Tag             =   "2"
         Top             =   660
         Width           =   2895
      End
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   615
      Left            =   9000
      Picture         =   "frmMntCtasContabN.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton cmdBuscar 
      Caption         =   "&Buscar"
      Height          =   555
      Left            =   9000
      Picture         =   "frmMntCtasContabN.frx":0974
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   60
      Width           =   1215
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   375
      Left            =   9000
      TabIndex        =   3
      Top             =   1860
      Width           =   1215
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      DisabledPicture =   "frmMntCtasContabN.frx":0A76
      Height          =   375
      Left            =   9000
      TabIndex        =   2
      Top             =   1440
      Width           =   1215
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "&Nuevo"
      Height          =   375
      Left            =   9000
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1020
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
      Height          =   2775
      Left            =   8940
      TabIndex        =   15
      Top             =   930
      Visible         =   0   'False
      Width           =   1365
      Begin VB.OptionButton optImpre 
         Caption         =   "&Todo"
         Height          =   285
         Index           =   0
         Left            =   150
         TabIndex        =   21
         Top             =   300
         Width           =   1065
      End
      Begin VB.OptionButton optImpre 
         Caption         =   "&Grupo"
         Height          =   285
         Index           =   2
         Left            =   150
         TabIndex        =   20
         Top             =   930
         Width           =   915
      End
      Begin VB.TextBox txtCtaCod 
         Enabled         =   0   'False
         Height          =   315
         Left            =   60
         TabIndex        =   19
         Top             =   1500
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   315
         Left            =   90
         TabIndex        =   18
         ToolTipText     =   "Cancelar Cambio de Fecha de Asiento"
         Top             =   2310
         Width           =   1155
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "&Aplicar"
         Height          =   315
         Left            =   90
         TabIndex        =   17
         Top             =   1950
         Width           =   1155
      End
      Begin VB.OptionButton optImpre 
         Caption         =   "&Sin Agencia"
         Height          =   285
         Index           =   1
         Left            =   150
         TabIndex        =   16
         Top             =   615
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin MSComctlLib.ProgressBar pgbCtaContN 
      Height          =   180
      Left            =   120
      TabIndex        =   22
      Top             =   5400
      Visible         =   0   'False
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   318
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
End
Attribute VB_Name = "frmMntCtasContabN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sSqlCta As String
Dim rsCtaObj As ADODB.Recordset
Dim rsCta As ADODB.Recordset

Dim clsCtaCont As DCtaCont

Dim nOrdenCta   As Integer 'Para establecer el Orden de la tablas
Dim lPressMouse As Boolean
Dim lbConsulta  As Boolean
Dim vsTabla     As String

Dim WithEvents oImp As NContImprimir
Attribute oImp.VB_VarHelpID = -1
Dim oBarra As clsProgressBar

Public Sub Inicio(Optional psTabla As String = "CtaCont", Optional pbConsulta As Boolean = False)
lbConsulta = pbConsulta
vsTabla = psTabla
Me.Show 0, frmMdiMain
End Sub

Private Sub ManejaBotonVisible(plOpcion As Boolean)
cmdBuscar.Visible = plOpcion
If Not lbConsulta Then
   cmdNuevo.Visible = plOpcion
   cmdModificar.Visible = plOpcion
   cmdEliminar.Visible = plOpcion
End If
cmdImprimir.Visible = plOpcion
fraImprime.Visible = Not plOpcion
End Sub

Private Sub ManejaBoton(plOpcion As Boolean)
cmdBuscar.Enabled = plOpcion
If Moneda(0).value = True Then
   If ((Moneda(0).value Or Moneda(6).value) Or Not plOpcion) And Not lbConsulta Then
      cmdNuevo.Enabled = plOpcion
      cmdModificar.Enabled = plOpcion
   End If
End If
cmdEliminar.Enabled = plOpcion
cmdImprimir.Enabled = plOpcion
frmMoneda.Enabled = plOpcion
Moneda(0).Enabled = plOpcion
Moneda(1).Enabled = plOpcion
Moneda(2).Enabled = plOpcion
Moneda(3).Enabled = plOpcion
Moneda(4).Enabled = plOpcion
Moneda(6).Enabled = plOpcion

grdCtas.Enabled = plOpcion
'grdCtaObjs.Enabled = plOpcion
End Sub

Private Sub RefrescaGrid(npMoneda As Integer)
Set rsCta = clsCtaCont.CargaCtaCont("SubString(cCtaContCod,3,1)='" & Trim(Str(npMoneda)) & "' or Len(rtrim(cCtaContCod))<3", vsTabla, adLockOptimistic)
Set grdCtas.DataSource = rsCta
End Sub

Private Sub cmdbuscar_Click()
Dim clsBuscar As New ClassDescObjeto
ManejaBoton False
clsBuscar.BuscarDato rsCta, nOrdenCta, "Cuenta Contable"
nOrdenCta = clsBuscar.gnOrdenBusca
Set clsBuscar = Nothing
ManejaBoton True
grdCtas.SetFocus
End Sub

Private Sub cmdCancelar_Click()
ManejaBotonVisible True
ManejaBoton True
End Sub

Private Sub cmdEliminar_Click()
Dim dSQL As String
Dim Pos As Variant
On Error GoTo DelError
If Not rsCta.BOF And Not rsCta.EOF Then
   If Not clsCtaCont.CtaInstancia(rsCta!cCtaContCod, vsTabla) Then
      MsgBox " Existen Cuentas en nivel Inferior ...!", vbExclamation, "Aviso de Eliminación"
      grdCtas.SetFocus
      Exit Sub
   End If
   If rsCtaObj.RecordCount > 0 Then
      MsgBox " Existen Objetos asignados a Cuenta ...!", vbExclamation, "Aviso de Eliminación"
      grdCtas.SetFocus
      Exit Sub
   End If
   If MsgBox(" ¿ Seguro de Eliminar Cuenta ? ", vbOKCancel, "Mensaje de Confirmación") = vbOk Then
      
      clsCtaCont.EliminaCtaCont Mid(rsCta!cCtaContCod, 1, 2) & IIf(Len(rsCta!cCtaContCod) > 2, "_", "") & Mid(rsCta!cCtaContCod, 4, 22), vsTabla
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

Private Sub cmdExcelN_Click()
E = True
imprimirCtaContN
End Sub
Public Sub imprimirCtaContN()
    Dim fs              As Scripting.FileSystemObject
    Dim xlAplicacion    As Excel.Application
    Dim xlLibro         As Excel.Workbook
    Dim xlHoja1         As Excel.Worksheet
    Dim lbExisteHoja    As Boolean
    'ALPA 20101117**************************
    'Dim liLineas        As Integer
    Dim liLineas        As Long
    '***************************************
    Dim i               As Integer
    Dim glsarchivo      As String
    Dim lsNomHoja       As String

    Dim RSTEMP As New ADODB.Recordset
    'Dim clsBuscar As New ClassDescObjeto

    Set RSTEMP = clsCtaCont.CargaCtaCont("SubString(cCtaContCod,3,1)='" & Trim(Str(lnIndex)) & "' or Len(rtrim(cCtaContCod))<3", "CtaCont", adLockOptimistic)
       
    If RSTEMP Is Nothing Then
        MsgBox "No exite informacion para imprimir", vbInformation, "Aviso"
        Exit Sub
    End If

    glsarchivo = "Reporte_CuentasContables" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".XLS"
    Set fs = New Scripting.FileSystemObject

    Set xlAplicacion = New Excel.Application
    If fs.FileExists(App.path & "\SPOOLER\" & glsarchivo) Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\SPOOLER\" & glsarchivo)
    Else
        Set xlLibro = xlAplicacion.Workbooks.Add
    End If
    Set xlHoja1 = xlLibro.Worksheets.Add
     
    xlHoja1.PageSetup.CenterHorizontally = True
    xlHoja1.PageSetup.Zoom = 60
    xlHoja1.PageSetup.Orientation = xlLandscape

            lbExisteHoja = False
            lsNomHoja = "CuentasContables"
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
            xlHoja1.Cells(2, 2) = "L I S T A D O   D E   C U E N T A S   C O N T A B L E S  "
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

         Me.pgbCtaContN.Visible = True
         Me.pgbCtaContN.Min = 0
         Me.pgbCtaContN.Max = RSTEMP.RecordCount
         
         Do Until RSTEMP.EOF
            xlHoja1.Cells(liLineas, 2) = RSTEMP(0)
            xlHoja1.Cells(liLineas, 3) = RSTEMP(1)
            liLineas = liLineas + 1
            Me.pgbCtaContN.value = RSTEMP.Bookmark
            RSTEMP.MoveNext
            
         Loop
         Me.pgbCtaContN.Visible = False

        ExcelCuadro xlHoja1, 2, 6, 3, liLineas - 1
        xlHoja1.SaveAs App.path & "\SPOOLER\" & glsarchivo
        ExcelEnd App.path & "\Spooler\" & glsarchivo, xlAplicacion, xlLibro, xlHoja1
        
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing
        MsgBox "Se ha generado el Archivo en " & App.path & "\SPOOLER\" & glsarchivo
        Call CargaArchivo(glsarchivo, App.path & "\SPOOLER\")

End Sub

Private Sub cmdImprimir_Click()
ManejaBoton False
ManejaBotonVisible False
If optImpre(0).value Then
   optImpre(0).SetFocus
ElseIf optImpre(1).value Then
   optImpre(1).SetFocus
Else
   optImpre(2).SetFocus
End If
End Sub

Private Sub cmdAplicar_Click()
Dim sSql As String
Dim rs   As New ADODB.Recordset
Dim N    As Integer
Dim sTexto As String, lsImpre As String
Dim sNomFile As String
MousePointer = 11
Me.Enabled = False
For N = 0 To Moneda.Count
   If Moneda(IIf(N = 5, 6, N)).value Then
      If optImpre(0).value Then
         sSql = "SubString(cCtaContCod,3,1)='" & Moneda(IIf(N = 5, 6, N)).Tag & "' " _
                   & "or Len(rtrim(cCtaContCod))<3 Order By cCtaContCod"
         sNomFile = "PLAN_Todo.txt"
      ElseIf optImpre(2).value Then
         sSql = "cCtaContCod LIKE '" & txtCtaCod & "%' and SubString(cCtaContCod,3,1) = '" & Moneda(IIf(N = 5, 6, N)).Tag & "'"
         sNomFile = "PLAN_Grupo.txt"
      
      Else
         sSql = "cCtaContCod LIKE '__" & Moneda(IIf(N = 5, 6, N)).Tag & "%' AND NOT CCTACONTDESC LIKE 'AGENCIA %'"
         sSql = sSql & " And NOT CCTACONTDESC LIKE 'AGENCIA %' and NOT cCtaContDesc LIKE 'OFICINA ESPECIAL%' " _
                     & " And NOT cCtaContDesc LIKE 'OFICINA ESP.%' and NOT cCtaContDesc LIKE '%OFIC. ESPEC.%' " _
                     & " And NOT cCtaContDesc LIKE 'SEDE INSTITUCIONAL%' and NOT cCtaContDesc LIKE 'OFIC. ESPECIAL%' " _
                     & " And NOT cCtaContDesc LIKE 'OFIC.ESPECIAL%' and NOT cCtaContDesc LIKE 'OFICINA%'" _
                     & " And NOT cCtaContDesc LIKE 'CHICLAYO%' and NOT cCtaContDesc LIKE 'AG. CAJAMARCA%' "
         sNomFile = "PLAN_SAgencia.txt"
      End If
      Exit For
   End If
Next
Set rs = clsCtaCont.CargaCtaCont(sSql, vsTabla)
If rs.EOF Then
   MsgBox "No Existen Cuentas Contables Registradas", vbInformation, "Aviso"
   Me.Enabled = True
   MousePointer = 0
   Exit Sub
End If
MousePointer = 11
   Set oImp = New NContImprimir
   lsImpre = oImp.ImprimePlanContable(rs, vsTabla, gnLinPage)
   Set oImp = Nothing
   RSClose rs
MousePointer = 0
EnviaPrevio lsImpre, "Cuenta Contables : Reporte ", gnLinPage, False
ManejaBotonVisible True
ManejaBoton True
Me.Enabled = True
End Sub

Private Sub cmdModificar_Click()
Dim lsCtaCod As String
ManejaBoton False
lsCtaCod = rsCta!cCtaContCod
frmMntCtasContNuevo.Inicia False, rsCta!cCtaContCod, rsCta!cCtaContDesc, , Moneda(0).value
If frmMntCtasContNuevo.OK Then
   RefrescaGrid IIf(Moneda(0).value, 0, 6)
   rsCta.Find "cCtaContCod = '" & lsCtaCod & "'"
End If
ManejaBoton True
End Sub
Private Sub cmdNuevo_Click()
Dim Pos As Variant
ManejaBoton False
frmMntCtasContNuevo.Inicia True, "", "", , Moneda(0).value
If frmMntCtasContNuevo.OK Then
   RefrescaGrid IIf(Moneda(0).value, 0, 6)
   rsCta.Find "cCtaContCod = '" & frmMntCtasContNuevo.cCtaContCod & "'", 0, adSearchForward, 1
End If
ManejaBoton True
End Sub
Private Sub cmdSalir_Click()

Unload Me
End Sub
Private Sub Form_Activate()
ManejaBoton True
grdCtas.SetFocus
End Sub

Private Sub Form_Load()
frmMdiMain.Enabled = False
CentraForm Me
lPressMouse = False
nOrdenCta = 0  'Inicialmente las cuentas se Ordenan por Codigo

Set clsCtaCont = New DCtaCont
RefrescaGrid 0
Set grdCtas.DataSource = rsCta

EsNuevo = False
lTransActiva = False
If lbConsulta Then
   cmdNuevo.Visible = False
   cmdModificar.Visible = False
   cmdEliminar.Visible = False
   cmdImprimir.Top = cmdNuevo.Top
End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
Set clsCtaCont = Nothing
Set rsCtaObj = Nothing
frmMdiMain.Enabled = True
End Sub

Private Sub grdCtas_GotFocus()
grdCtas.MarqueeStyle = dbgHighlightRow
End Sub

Private Sub grdCtas_HeadClick(ByVal ColIndex As Integer)
If Not rsCta Is Nothing Then
   If Not rsCta.EOF Then
      rsCta.Sort = grdCtas.Columns(ColIndex).DataField
   End If
End If
End Sub

Private Sub grdCtas_LostFocus()
grdCtas.MarqueeStyle = dbgNoMarquee
End Sub

Private Sub Moneda_Click(Index As Integer)
RefrescaGrid Index

If Index = 0 Or Index = 6 Then
   cmdNuevo.Enabled = True
   cmdModificar.Enabled = True
Else
   cmdNuevo.Enabled = False
End If
End Sub

Private Sub optImpre_Click(Index As Integer)
If Index = 2 Then
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

