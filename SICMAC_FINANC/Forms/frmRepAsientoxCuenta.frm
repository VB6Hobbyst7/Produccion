VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmRepAsientoxCuenta 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Reporte de Asientos Contables por Cuenta"
   ClientHeight    =   4500
   ClientLeft      =   11010
   ClientTop       =   7455
   ClientWidth     =   13920
   Icon            =   "frmRepAsientoxCuenta.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   13920
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame frmResultado 
      Caption         =   "Resultado"
      Height          =   3030
      Left            =   90
      TabIndex        =   8
      Top             =   945
      Width           =   13740
      Begin Sicmact.FlexEdit fgDATA 
         Height          =   2535
         Left            =   90
         TabIndex        =   3
         Top             =   405
         Width           =   13560
         _ExtentX        =   23918
         _ExtentY        =   4471
         Cols0           =   11
         FixedCols       =   0
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Fecha-Hora-Cod. Ope.-Operación-Cuenta Contable-Debe-Haber-Cod. Ag.-Mov.-Cuenta"
         EncabezadosAnchos=   "400-1150-950-950-2400-1450-1150-1150-1100-1150-1550"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X-X-X-X-X-X-X"
         ListaControles  =   "-0-0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "-C-C-C-C-C-C-C-C-C-C"
         FormatosEdit    =   "-0-0-0-0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   6660
      TabIndex        =   5
      Top             =   4050
      Width           =   1455
   End
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Excel"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   4050
      Width           =   1590
   End
   Begin VB.Frame frmFiltro 
      Caption         =   "Filtro"
      Height          =   870
      Left            =   90
      TabIndex        =   6
      Top             =   90
      Width           =   13740
      Begin Sicmact.ActXCodCta actCodigoCuenta 
         Height          =   465
         Left            =   135
         TabIndex        =   0
         Top             =   270
         Width           =   3795
         _ExtentX        =   6694
         _ExtentY        =   820
         Texto           =   "Cuenta:"
         EnabledCMAC     =   -1  'True
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "109"
      End
      Begin VB.CommandButton cmdMostrar 
         Caption         =   "Mostrar"
         Height          =   375
         Left            =   12150
         TabIndex        =   2
         Top             =   270
         Width           =   1365
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   345
         Left            =   6390
         TabIndex        =   1
         Top             =   315
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "dd/mm/yyyy"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Fecha:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   5625
         TabIndex        =   7
         Top             =   360
         Width           =   555
      End
   End
End
Attribute VB_Name = "frmRepAsientoxCuenta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sTipoCod As String
Dim oDcreditos As DCreditos
Dim rs As ADODB.Recordset
Private Sub generarReporte()
    On Error GoTo ErrorGenerar
    Dim oAsiento As New ncontasientos
    Dim row As Long
        
        If Len(actCodigoCuenta.Age) <> 2 Or Len(actCodigoCuenta.Prod) <> 3 Or Len(actCodigoCuenta.Cuenta) <> 10 Then
        MsgBox "Por favor, ingresar un número de cuenta válido.", vbInformation, "Aviso"
       
        Exit Sub
        End If
        
        LimpiaFlex fgDATA
    Set rs = oAsiento.getasientocontablexcuenta(actCodigoCuenta.NroCuenta, txtFecha.Text)
    
    If rs.EOF And rs.BOF Then
        MsgBox "No hay Asientos Contables para mostrar en esta fecha."
        cmdExportar.Enabled = False
        Else
        cmdExportar.Enabled = True
        
    End If

    Do While Not rs.EOF
        fgDATA.AdicionaFila
        row = fgDATA.row
        fgDATA.TextMatrix(row, 1) = rs!Fecha
        fgDATA.TextMatrix(row, 2) = rs!hora
        fgDATA.TextMatrix(row, 3) = rs!CodOperacion
        fgDATA.TextMatrix(row, 4) = rs!cOpeDesc
        fgDATA.TextMatrix(row, 5) = rs!CtaContable
        fgDATA.TextMatrix(row, 6) = rs!Debe
        fgDATA.TextMatrix(row, 7) = rs!Haber
        fgDATA.TextMatrix(row, 8) = rs!CodAgencia
        fgDATA.TextMatrix(row, 9) = rs!Num_Mov
        fgDATA.TextMatrix(row, 10) = rs!Cuenta
        rs.MoveNext
    Loop
    Exit Sub
    
ErrorGenerar:
    MsgBox "Por favor, ingresar una fecha correcta."
End Sub
Private Sub actCodigoCuenta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
        txtFecha.SetFocus
    End If
End Sub
Private Sub cmdExportar_Click()
    CrearExcelReporte
End Sub
Private Sub cmdMostrar_Click()
    generarReporte
End Sub
Public Sub CrearExcelReporte()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String
    Dim lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim i As Integer: Dim IniTablas As Integer
    Dim oPersona As UPersona
    
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    Set oPersona = New UPersona
    
    lsNomHoja = "Hoja1"
    lsFile = "FormatoAsientoxCuenta"
    
    lsArchivo = "\spooler\" & "RepAsientoxCuenta" & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xlsx"
    If fs.FileExists(App.path & "\FormatoCarta\" & lsFile & ".xlsx") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsFile & ".xlsx")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xlsx), Consulte con el Area de TI", vbInformation, "Advertencia"
        Exit Sub
    End If

    For Each xlHoja1 In xlsLibro.Worksheets
       If xlHoja1.Name = lsNomHoja Then
            xlHoja1.Activate
         lbExisteHoja = True
        Exit For
       End If
    Next
    
    If lbExisteHoja = False Then
        Set xlHoja1 = xlsLibro.Worksheets
        xlHoja1.Name = lsNomHoja
    End If
    
    xlHoja1.Cells(2, 5) = actCodigoCuenta.NroCuenta
    xlHoja1.Cells(3, 5) = txtFecha.Text
    
    IniTablas = 5
    For i = 1 To fgDATA.Rows - 1
    xlHoja1.Cells(IniTablas + i, 3) = fgDATA.TextMatrix(i, 0)
        xlHoja1.Cells(IniTablas + i, 4) = fgDATA.TextMatrix(i, 1)
        xlHoja1.Cells(IniTablas + i, 5) = fgDATA.TextMatrix(i, 2)
        xlHoja1.Cells(IniTablas + i, 6) = fgDATA.TextMatrix(i, 3)
        xlHoja1.Cells(IniTablas + i, 7) = fgDATA.TextMatrix(i, 4)
        xlHoja1.Cells(IniTablas + i, 8) = fgDATA.TextMatrix(i, 5)
        xlHoja1.Cells(IniTablas + i, 9) = fgDATA.TextMatrix(i, 6)
        xlHoja1.Cells(IniTablas + i, 10) = fgDATA.TextMatrix(i, 7)
        xlHoja1.Cells(IniTablas + i, 11) = fgDATA.TextMatrix(i, 8)
        xlHoja1.Cells(IniTablas + i, 12) = fgDATA.TextMatrix(i, 9)
        xlHoja1.Cells(IniTablas + i, 13) = fgDATA.TextMatrix(i, 10)
    Next i
    
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.path & lsArchivo
    psArchivoAGrabarC = App.path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
End Sub
Private Sub cmdSalir_Click()
    Unload Me
End Sub
Private Sub Form_Activate()
    actCodigoCuenta.SetFocusAge
    cmdExportar.Enabled = False
End Sub
Private Sub Form_Load()
    SetDatosControl
End Sub
Private Sub SetDatosControl()
   txtFecha.Text = gdFecSis
End Sub
Private Sub txtFecha_GotFocus()
    fEnfoque txtFecha
End Sub
Private Sub txtFecha_KeyPress(KeyAscii As Integer)
  If KeyAscii = 13 Then
        cmdMostrar.SetFocus
    End If
End Sub
