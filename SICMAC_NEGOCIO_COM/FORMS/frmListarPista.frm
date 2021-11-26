VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmListarPista 
   Caption         =   "Listar Pistas"
   ClientHeight    =   6570
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11640
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmListarPista.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   11640
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Búsqueda"
      Height          =   1695
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   11415
      Begin VB.CommandButton Command2 
         Caption         =   "Excel"
         Height          =   375
         Left            =   7080
         TabIndex        =   9
         Top             =   840
         Visible         =   0   'False
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   315
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61210625
         CurrentDate     =   39678
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   7080
         TabIndex        =   4
         Top             =   360
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   315
         Left            =   2040
         TabIndex        =   7
         Top             =   840
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61210625
         CurrentDate     =   39678
      End
      Begin VB.Label Label1 
         Caption         =   "Hasta:"
         Height          =   255
         Left            =   1080
         TabIndex        =   8
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Desde:"
         Height          =   255
         Left            =   1080
         TabIndex        =   5
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   1800
      Width           =   11415
      Begin MSDataGridLib.DataGrid dgBuscar 
         Height          =   4215
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   11145
         _ExtentX        =   19659
         _ExtentY        =   7435
         _Version        =   393216
         AllowUpdate     =   0   'False
         ColumnHeaders   =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         RowDividerStyle =   1
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   8
         BeginProperty Column00 
            DataField       =   "NombreUsuario"
            Caption         =   "Id"
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
            DataField       =   "NombreUsuario"
            Caption         =   "Nombre Usuario"
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
         BeginProperty Column02 
            DataField       =   "Usuario"
            Caption         =   "Usuario"
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
            DataField       =   "Operacion"
            Caption         =   "Accion"
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
         BeginProperty Column04 
            DataField       =   "Fecha"
            Caption         =   "F. Reg"
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
         BeginProperty Column05 
            DataField       =   "Hora"
            Caption         =   "H. Reg"
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
         BeginProperty Column06 
            DataField       =   "Maquina"
            Caption         =   "Terminal"
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   "0%"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   10250
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column07 
            DataField       =   "Observacion"
            Caption         =   "Observacion"
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
            Size            =   182
            BeginProperty Column00 
               ColumnWidth     =   0
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   900.284
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   3000.189
            EndProperty
            BeginProperty Column04 
               ColumnWidth     =   1305.071
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   1200.189
            EndProperty
            BeginProperty Column06 
               ColumnWidth     =   1500.095
            EndProperty
            BeginProperty Column07 
               ColumnWidth     =   1500.095
            EndProperty
         EndProperty
      End
      Begin VB.Label lblMensaje 
         Caption         =   "NO EXISTEN DATOS"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3960
         TabIndex        =   2
         Top             =   240
         Visible         =   0   'False
         Width           =   2175
      End
   End
End
Attribute VB_Name = "frmListarPista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'** Nombre : frmListarPista
'** Descripción : Formulario para visualizar la Relacion de Pistas.
'** Creación : MAVM, 20080905 10:17:15 AM
'** Modificación:
'********************************************************************

Option Explicit
Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision
Dim FechaFinMes As Date
Dim lsmensaje As String

Public Sub BuscarDatos()
    Dim rs As Recordset
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    lsmensaje = ""
    Set rs = objCOMNAuditoria.ObtenerDatosPistas(Format(DTPicker1.value, "yyyymmdd"), Format(DTPicker2.value, "yyyymmdd"), lsmensaje)
    
    If lsmensaje = "" Then
    lblMensaje.Visible = False
    dgBuscar.Visible = True
    Set dgBuscar.DataSource = rs
    dgBuscar.Refresh
    Screen.MousePointer = 0
    dgBuscar.SetFocus
    Command2.Visible = True
'    Command3.Visible = True
'    Command4.Visible = True

    Else
     Set dgBuscar.DataSource = Nothing
    dgBuscar.Refresh
    lblMensaje.Visible = True
    dgBuscar.Visible = False
    Command2.Visible = False
'    Command3.Visible = False
'    Command4.Visible = False
    End If
    
    Set rs = Nothing
    Set objCOMNAuditoria = Nothing

End Sub

Private Sub Command1_Click()
    BuscarDatos
End Sub

Private Sub Command2_Click()
MostrarReportePistas
End Sub

Public Sub MostrarReportePistas()
Dim rs As Recordset
Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
Dim lMatCabecera As Variant
Dim lsNombreArchivo As String

    lsNombreArchivo = "Pistas"
    
    ReDim lMatCabecera(6, 2)

    lMatCabecera(0, 0) = "NombreUsuario": lMatCabecera(0, 1) = ""
    lMatCabecera(1, 0) = "Fecha": lMatCabecera(1, 1) = ""
    lMatCabecera(2, 0) = "Hora": lMatCabecera(2, 1) = "N"
    lMatCabecera(3, 0) = "Usuario": lMatCabecera(3, 1) = ""
    lMatCabecera(4, 0) = "Operacion": lMatCabecera(4, 1) = ""
    lMatCabecera(5, 0) = "Maquina": lMatCabecera(5, 1) = ""
    
        
    Set rs = objCOMNAuditoria.ObtenerDatosPistasReporte(Format(DTPicker1.value, "yyyymmdd"), Format(DTPicker2.value, "yyyymmdd"))
    Set objCOMNAuditoria = Nothing
           
    Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Pistas de la Evaluación de Cartera", "", lsNombreArchivo, lMatCabecera, rs, 2, , , True)
End Sub

Public Sub GeneraReporteEnArchivoExcel(ByVal psNomCmac As String, ByVal psNomAge As String, ByVal psCodUser As String, ByVal pdFecSis As Date, ByVal psTitulo As String, ByVal psSubTitulo As String, _
                                    ByVal psNomArchivo As String, ByVal pMatCabeceras As Variant, ByVal prRegistros As ADODB.Recordset, _
                                    Optional pnNumDecimales As Integer, Optional Visible As Boolean = False, Optional psNomHoja As String = "", _
                                    Optional pbSinFormatDeReg As Boolean = False, _
                                    Optional pbUsarCabecerasDeRS As Boolean = False)
    Dim rs As ADODB.Recordset
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim liLineas As Integer, i As Integer
    Dim fs As Scripting.FileSystemObject
    Dim lnNumColumns As Integer


    If Not (prRegistros.EOF And prRegistros.BOF) Then
        If pbUsarCabecerasDeRS = True Then
            lnNumColumns = prRegistros.Fields.Count
        Else
            lnNumColumns = UBound(pMatCabeceras)
            lnNumColumns = IIf(prRegistros.Fields.Count < lnNumColumns, prRegistros.Fields.Count, prRegistros.Fields.Count)
        End If

        If psNomHoja = "" Then psNomHoja = psNomArchivo
        psNomArchivo = psNomArchivo & "_" & psCodUser & ".xls"

        Set fs = New Scripting.FileSystemObject
        Set xlAplicacion = New Excel.Application
        If fs.FileExists(App.path & "\Spooler\" & psNomArchivo) Then
            fs.DeleteFile (App.path & "\Spooler\" & psNomArchivo)
        End If
        Set xlLibro = xlAplicacion.Workbooks.Add
        Set xlHoja1 = xlLibro.Worksheets.Add

        xlHoja1.Name = psNomHoja
        xlHoja1.Cells.Select
        'xlHoja1.Cells.NumberFormat = "@"

        'Cabeceras
        xlHoja1.Cells(1, 1) = psNomCmac
        xlHoja1.Cells(1, lnNumColumns) = Trim(Format(pdFecSis, "dd/mm/yyyy hh:mm:ss"))
        xlHoja1.Cells(2, 1) = psNomAge
        xlHoja1.Cells(2, lnNumColumns) = psCodUser
        xlHoja1.Cells(4, 1) = psTitulo
        xlHoja1.Cells(5, 1) = psSubTitulo
        xlHoja1.Range(xlHoja1.Cells(1, 1), xlHoja1.Cells(5, lnNumColumns)).Font.Bold = True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(4, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(5, 1), xlHoja1.Cells(5, lnNumColumns)).Merge True
        xlHoja1.Range(xlHoja1.Cells(4, 1), xlHoja1.Cells(5, lnNumColumns)).HorizontalAlignment = xlCenter

        liLineas = 6
        If pbUsarCabecerasDeRS = True Then
            For i = 0 To prRegistros.Fields.Count - 1
                xlHoja1.Cells(liLineas, i + 1) = prRegistros.Fields(i).Name
            Next i
        Else
            For i = 0 To lnNumColumns - 1
                If (i + 1) > UBound(pMatCabeceras) Then
                    xlHoja1.Cells(liLineas, i + 1) = prRegistros.Fields(i).Name
                Else
                    xlHoja1.Cells(liLineas, i + 1) = pMatCabeceras(i, 0)
                End If
            Next i
        End If
        
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).Cells.Interior.Color = RGB(220, 220, 220)
        xlHoja1.Range(xlHoja1.Cells(liLineas, 1), xlHoja1.Cells(liLineas, lnNumColumns)).HorizontalAlignment = xlCenter

        If pbSinFormatDeReg = False Then
            liLineas = liLineas + 1
            While Not prRegistros.EOF
                For i = 0 To lnNumColumns - 1
                    If pMatCabeceras(i, 1) = "" Then  'Verificamos si tiene tipo
                        xlHoja1.Cells(liLineas, i + 1) = prRegistros(i)
                    Else
                        Select Case pMatCabeceras(i, 1)
                            Case "S"
                                xlHoja1.Cells(liLineas, i + 1) = prRegistros(i)
                            Case "N"
                                xlHoja1.Cells(liLineas, i + 1) = Format(prRegistros(i), "#0.00")
                            Case "D"
                                xlHoja1.Cells(liLineas, i + 1) = IIf(Format(prRegistros(i), "yyyymmdd") = "19000101", "", Format(prRegistros(i), "dd/mm/yyyy"))
                        End Select
                    End If
                Next i
                liLineas = liLineas + 1
                prRegistros.MoveNext
            Wend
        Else
            xlHoja1.Range("A7").CopyFromRecordset prRegistros 'Copia el contenido del recordset a excel
        End If

        xlHoja1.SaveAs App.path & "\Spooler\" & psNomArchivo
        MsgBox "Se ha generado el Archivo en " & App.path & "\Spooler\" & psNomArchivo

        If Visible Then
            xlAplicacion.Visible = True
            xlAplicacion.Windows(1).Visible = True
        End If

            xlLibro.Close
            xlAplicacion.Quit
        Set xlAplicacion = Nothing
        Set xlLibro = Nothing
        Set xlHoja1 = Nothing

    End If
End Sub

Private Sub Form_Load()
    DTPicker1.value = Date
    DTPicker2.value = Date
End Sub
