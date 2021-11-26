VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmOpeRepActDatosCamp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Actualización de Datos por Campaña"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4245
   Icon            =   "frmOpeRepActDatosCamp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   0
      Width           =   3975
      Begin VB.CommandButton cmdGenerar 
         Caption         =   "Generar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   380
         Left            =   2520
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin MSMask.MaskEdBox txtFecha 
         Height          =   350
         Left            =   840
         TabIndex        =   0
         Top             =   260
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   635
         _Version        =   393216
         BackColor       =   16777215
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha :"
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   315
         Width           =   615
      End
   End
   Begin ComctlLib.ProgressBar pgbExcel 
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Visible         =   0   'False
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   450
      _Version        =   327682
      Appearance      =   1
   End
End
Attribute VB_Name = "frmOpeRepActDatosCamp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'** Nombre : frmOpeRepActDatosCamp
'** Descripción : Formulario para la generación de reporte sobre la Campaña "Actualiza tus Datos" según TI-ERS134-2013
'** Creación : JUEZ, 20131021 09:00:00 AM
'*****************************************************************************************************

Option Explicit

Private Sub cmdGenerar_Click()
    Dim sMsjValida As String
    sMsjValida = ValidaFecha(txtFecha.Text)
    If Len(sMsjValida) > 0 Then
        MsgBox sMsjValida, vbInformation, "Aviso"
        txtFecha.SetFocus
    End If
    
    Dim oDPers As COMDPersona.DCOMPersonas
    Dim rsRep As ADODB.Recordset
    
    Set oDPers = New COMDPersona.DCOMPersonas
    Set rsRep = oDPers.ReporteArqueoCampActualizacionDatos(txtFecha.Text, gsCodAge)
    Set oDPers = Nothing
    If rsRep.RecordCount > 0 Then
        Dim xlAplicacion As Excel.Application
        Dim xlLibro As Excel.Workbook
        Dim lbLibroOpen As Boolean
        Dim lsArchivo As String
        Dim lsHoja As String
        Dim xlHoja1 As Excel.Worksheet
        Dim xlHoja2 As Excel.Worksheet
        Dim nLin As Long
        Dim nItem As Long
              
        pgbExcel.Visible = True
        pgbExcel.Min = 0
        pgbExcel.value = 0
    
        lsArchivo = App.path & "\SPOOLER\ReporteArqueoActualizacionDatos_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS") & ".xls"
        lbLibroOpen = gFunGeneral.ExcelBegin(lsArchivo, xlAplicacion, xlLibro, False)
        If Not lbLibroOpen Then
            Exit Sub
        End If
        nLin = 1
        lsHoja = "Hoja1"
        gFunGeneral.ExcelAddHoja lsHoja, xlLibro, xlHoja1
        
        xlHoja1.Range("A1:Y1").EntireColumn.Font.FontStyle = "Arial"
        xlHoja1.PageSetup.CenterHorizontally = True
        xlHoja1.PageSetup.Zoom = 75
        xlHoja1.PageSetup.TopMargin = 2
        
        xlHoja1.Range("A1:A1").RowHeight = 18
        xlHoja1.Range("A1:A1").ColumnWidth = 8
        xlHoja1.Range("B1:B1").ColumnWidth = 18
        xlHoja1.Range("C1:C1").ColumnWidth = 14
        xlHoja1.Range("D1:D1").ColumnWidth = 50
        xlHoja1.Range("E1:F1").ColumnWidth = 12
        
        xlHoja1.Cells(nLin, 1) = "Número"
        xlHoja1.Cells(nLin, 2) = "Agencia"
        xlHoja1.Cells(nLin, 3) = "Codigo Cliente"
        xlHoja1.Cells(nLin, 4) = "Cliente"
        xlHoja1.Cells(nLin, 5) = "Usuario Reg."
        xlHoja1.Cells(nLin, 6) = "Usuario Resp."
        
        xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Bold = True
        xlHoja1.Range("A" & nLin & ":F" & nLin).HorizontalAlignment = xlHAlignCenter
        xlHoja1.Range("A" & nLin & ":F" & nLin).BorderAround xlContinuous, xlThin, xlColorIndexAutomatic, 0
        xlHoja1.Range("A" & nLin & ":F" & nLin).Borders(xlInsideVertical).LineStyle = xlContinuous
        xlHoja1.Range("A" & nLin & ":F" & nLin).Borders(xlInsideVertical).Color = vbBlack
        xlHoja1.Range("A" & nLin & ":F" & nLin).Interior.Color = RGB(255, 50, 50)
        xlHoja1.Range("A" & nLin & ":F" & nLin).Font.Color = RGB(255, 255, 255)
        
        With xlHoja1.PageSetup
            .LeftHeader = ""
            .CenterHeader = ""
            .RightHeader = ""
            .LeftFooter = ""
            .CenterFooter = ""
            .RightFooter = ""
        
            .PrintHeadings = False
            .PrintGridlines = False
            .PrintComments = xlPrintNoComments
            .CenterHorizontally = True
            .CenterVertically = False
            .Draft = False
            .FirstPageNumber = xlAutomatic
            .Order = xlDownThenOver
            .BlackAndWhite = False
            .Zoom = 55
        End With
        
        nItem = 1
        nLin = nLin + 1
        pgbExcel.Max = rsRep.RecordCount
        For nItem = 1 To rsRep.RecordCount
'            xlHoja1.Range("A" & nLin & ":B" & nLin).HorizontalAlignment = xlHAlignLeft
            xlHoja1.Cells(nLin, 1) = rsRep!nItem
            xlHoja1.Cells(nLin, 2) = rsRep!cAgeDescripcion
            xlHoja1.Cells(nLin, 3) = rsRep!cPersCod
            xlHoja1.Cells(nLin, 4) = rsRep!cPersNombre
            xlHoja1.Cells(nLin, 5) = rsRep!cUserReg
            xlHoja1.Cells(nLin, 6) = rsRep!cCodUserResp
            rsRep.MoveNext
            pgbExcel.value = pgbExcel.value + 1
            nLin = nLin + 1
        Next nItem
        
        gFunGeneral.ExcelEnd lsArchivo, xlAplicacion, xlLibro, xlHoja1
        gFunGeneral.CargaArchivo lsArchivo, App.path & "\SPOOLER\"
        pgbExcel.Min = 0
        pgbExcel.value = 0
        pgbExcel.Visible = False
    Else
        MsgBox "No existen datos para el reporte", vbInformation, "Aviso"
    End If
End Sub

Private Sub txtFecha_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdGenerar.SetFocus
    End If
End Sub
