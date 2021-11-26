VERSION 5.00
Begin VB.Form frmColPTarifarioNotificHis 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Historial de Cambios"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6855
   Icon            =   "frmColPTarifarioNotificHis.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4500
   ScaleWidth      =   6855
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExportar 
      Caption         =   "Exportar"
      Height          =   310
      Left            =   4305
      TabIndex        =   4
      Top             =   4080
      Width           =   1215
   End
   Begin VB.CommandButton cmdCerrar 
      Caption         =   "Cerrar"
      Height          =   310
      Left            =   5520
      TabIndex        =   3
      Top             =   4080
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   " Agencia:  "
      ForeColor       =   &H00FF0000&
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6615
      Begin SICMACT.FlexEdit feHistorial 
         Height          =   3375
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   6375
         _ExtentX        =   11245
         _ExtentY        =   5953
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Fecha Registro-Fecha Vigencia-Costo-Usuario"
         EncabezadosAnchos=   "400-2600-1400-800-800"
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
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label lblAgeNom 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Left            =   840
         TabIndex        =   1
         Top             =   0
         Width           =   3015
      End
   End
End
Attribute VB_Name = "frmColPTarifarioNotificHis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre      : frmColPTarifarioNotificHis
'** Descripción : Formulario que muestra el historial del tarifario de costo por carta notarial en las agencias.
'** Creación    : RECO, 20160229 - ERS056-2015
'**********************************************************************************************

Option Explicit

Public Sub Inicia(ByVal pnTarifarioID As Integer, ByVal psAgeNombre As String)
    Call CargarHistorial(pnTarifarioID)
    lblAgeNom.Caption = Space(4) & psAgeNombre
    Me.Show 1
End Sub

Private Sub cmdCerrar_Click()
    Unload Me
End Sub

Private Sub CargarHistorial(ByVal pnTarifarioID As Integer)
    Dim oPig As New COMDColocPig.DCOMColPActualizaBD
    Dim rs As New ADODB.Recordset
    Dim nIndice As Integer
    
    Set rs = oPig.PignoHistorialTarifarioNotific(pnTarifarioID)
    
    If Not (rs.EOF And rs.BOF) Then
        feHistorial.Clear
        FormateaFlex feHistorial
        For nIndice = 1 To rs.RecordCount
            feHistorial.AdicionaFila
            feHistorial.TextMatrix(nIndice, 1) = rs!dFecReg
            feHistorial.TextMatrix(nIndice, 2) = Format(rs!dFecIni, "dd/MM/yyyy")
            feHistorial.TextMatrix(nIndice, 3) = Format(rs!nValor, gsFormatoNumeroView)
            feHistorial.TextMatrix(nIndice, 4) = rs!cUser
            rs.MoveNext
        Next
    End If
End Sub

Public Sub ImprimeHojaRuta()
    Dim fs As Scripting.FileSystemObject
    Dim xlsAplicacion As Excel.Application
    Dim lsArchivo As String, lsFile As String
    Dim lsNomHoja As String
    Dim xlsLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lbExisteHoja As Boolean
    Dim I As Integer: Dim IniTablas As Integer
    Set fs = New Scripting.FileSystemObject
    Set xlsAplicacion = New Excel.Application
    
    
    lsNomHoja = "Hoja1"
    lsFile = "FormtatoTarifarioPignpAdj"
    
    lsArchivo = "\spooler\" & lsFile & "_" & gsCodUser & "_" & Format(gdFecSis, "yyyymmdd") & "_" & Format(Time, "hhmmss") & ".xls"
    If fs.FileExists(App.Path & "\FormatoCarta\" & lsFile & ".xls") Then
        Set xlsLibro = xlsAplicacion.Workbooks.Open(App.Path & "\FormatoCarta\" & lsFile & ".xls")
    Else
        MsgBox "No Existe Plantilla en Carpeta FormatoCarta (" & lsFile & ".xls), Consulte con el Area de TI", vbInformation, "Advertencia"
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
     
    IniTablas = 5
    
    For I = 1 To feHistorial.Rows - 1
        xlHoja1.Cells(IniTablas + I, 1) = feHistorial.TextMatrix(I, 0)
        xlHoja1.Cells(IniTablas + I, 2) = Format(feHistorial.TextMatrix(I, 1), "dd/MM/yyyy")
        xlHoja1.Cells(IniTablas + I, 3) = Format(feHistorial.TextMatrix(I, 2), "dd/MM/yyyy")
        xlHoja1.Cells(IniTablas + I, 4) = feHistorial.TextMatrix(I, 3)
        xlHoja1.Cells(IniTablas + I, 5) = feHistorial.TextMatrix(I, 4)
    Next I
    
    xlHoja1.Range(xlHoja1.Cells(6, 1), xlHoja1.Cells(IniTablas + I, 5)).Borders.LineStyle = 1
                
    Dim psArchivoAGrabarC As String
    
    xlHoja1.SaveAs App.Path & lsArchivo
    psArchivoAGrabarC = App.Path & lsArchivo
    xlsAplicacion.Visible = True
    xlsAplicacion.Windows(1).Visible = True
    Set xlsAplicacion = Nothing
    Set xlsLibro = Nothing
    Set xlHoja1 = Nothing
End Sub

Private Sub cmdExportar_Click()
    Call ImprimeHojaRuta
End Sub
