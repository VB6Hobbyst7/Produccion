VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmAuditReporteCreditosDC 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Reporte de Créditos Desembolsados - Cancelados"
   ClientHeight    =   1890
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5100
   Icon            =   "frmAuditReporteCreditosDC.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   5100
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      Caption         =   "Parámetros de Búsqueda"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      Begin VB.CommandButton Command1 
         Caption         =   "Aceptar"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   300
         Left            =   720
         TabIndex        =   1
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   69402625
         CurrentDate     =   39743
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   300
         Left            =   3240
         TabIndex        =   2
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   69402625
         CurrentDate     =   39743
      End
      Begin VB.Label Label4 
         Caption         =   "De:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   480
         Width           =   375
      End
      Begin VB.Label Label5 
         Caption         =   "Hasta:"
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   3
         Top             =   480
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmAuditReporteCreditosDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCOMNAuditoria As COMNAuditoria.NCOMRevision

Private Sub Form_Load()
    CentraForm Me
    DTPicker1.value = Date
    DTPicker2.value = Date
End Sub

Private Sub Command1_Click()
    MostrarReporte
End Sub

Public Sub MostrarReporte()
    Dim rs As Recordset
    Set objCOMNAuditoria = New COMNAuditoria.NCOMRevision
    Dim lMatCabecera As Variant
    Dim lsNombreArchivo As String
    
        lsNombreArchivo = "CREDITOS_DC"
        
        ReDim lMatCabecera(20, 2)
    
        lMatCabecera(0, 0) = "Pers Cod": lMatCabecera(0, 1) = ""
        lMatCabecera(1, 0) = "Titular": lMatCabecera(1, 1) = ""
        lMatCabecera(2, 0) = "Cta Cod1": lMatCabecera(2, 1) = "N"
        lMatCabecera(3, 0) = "Descripcion1": lMatCabecera(3, 1) = ""
        lMatCabecera(4, 0) = "Moneda": lMatCabecera(4, 1) = ""
        lMatCabecera(5, 0) = "Condicion": lMatCabecera(5, 1) = ""
        
        lMatCabecera(6, 0) = "DA1": lMatCabecera(6, 1) = ""
        lMatCabecera(7, 0) = "Nro Cuota Apr1": lMatCabecera(7, 1) = ""
        lMatCabecera(8, 0) = "Nro Cuota Can1": lMatCabecera(8, 1) = ""
        lMatCabecera(9, 0) = "Calificacion1": lMatCabecera(9, 1) = ""
        lMatCabecera(10, 0) = "Monto Pagado": lMatCabecera(10, 1) = ""
        lMatCabecera(11, 0) = "User1": lMatCabecera(11, 1) = ""
        lMatCabecera(12, 0) = "Agencia1": lMatCabecera(12, 1) = ""
        lMatCabecera(13, 0) = "Cta Cod2": lMatCabecera(13, 1) = ""
        lMatCabecera(14, 0) = "Descripcion2": lMatCabecera(14, 1) = ""
        lMatCabecera(15, 0) = "Moneda2": lMatCabecera(15, 1) = ""
        lMatCabecera(16, 0) = "Desembolsado": lMatCabecera(16, 1) = ""
        lMatCabecera(17, 0) = "Nro Cuota Apr": lMatCabecera(17, 1) = ""
        lMatCabecera(18, 0) = "User2": lMatCabecera(18, 1) = ""
        lMatCabecera(19, 0) = "Agencia2": lMatCabecera(19, 1) = ""
        
        Set rs = objCOMNAuditoria.ObtenerCreditosDesembolsados(Format(DTPicker1.value, "yyyymmdd"), Format(DTPicker2.value, "yyyymmdd"))
        Set objCOMNAuditoria = Nothing
               
        Call GeneraReporteEnArchivoExcel(gsNomCmac, gsNomAge, gsCodUser, gdFecSis, "Reporte de Créditos Desembolsados - Cancelados", "", lsNombreArchivo, lMatCabecera, rs, 2, , , True)
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
        xlHoja1.Cells(4, 1) = psTitulo & " " & "Del:" & " " & DTPicker1.value & " " & "Al:" & " " & DTPicker2.value
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
