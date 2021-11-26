VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmReqSunat 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Requerimientos SUNAT"
   ClientHeight    =   2055
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6615
   Icon            =   "frmReqSunat.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2055
   ScaleWidth      =   6615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PB1 
      Height          =   255
      Left            =   1440
      TabIndex        =   4
      Top             =   1680
      Width           =   5055
      _ExtentX        =   8916
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdProcesar 
      Caption         =   "Procesar"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Frame frmReq 
      Caption         =   "Requerimiento"
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
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   6375
      Begin VB.TextBox txtArchivo 
         Height          =   375
         Left            =   1200
         TabIndex        =   7
         Top             =   360
         Width           =   3975
      End
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "..."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   5280
         TabIndex        =   6
         Top             =   360
         Width           =   495
      End
      Begin MSComDlg.CommonDialog dlgArchivo 
         Left            =   5760
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.ComboBox cmbTipoReqSunat 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   840
         Width           =   4575
      End
      Begin VB.Label lblTipo 
         Caption         =   "Tipo:"
         Height          =   255
         Left            =   360
         TabIndex        =   2
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblArchivo 
         Caption         =   "Archivo:"
         Height          =   255
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmReqSunat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'********************************************************************
'*** Nombre : frmReqSunat
'*** Descripción : Formulario para agregar archivo excel de Asesoria Legal.
'*** Creación : MIOL el 20120919, según OYP-RFC086-2012
'********************************************************************
Option Explicit
Dim rsTipoArchivo As Recordset
Dim rsClienteSunat As Recordset

Private Sub cmdBuscar_Click()
txtArchivo.Text = Empty

    dlgArchivo.InitDir = "C:\"
    dlgArchivo.Filter = "Archivos de Excel (*.xls)|*.xls| Archivos de Excel (*.xlsx)|*.xlsx"
    dlgArchivo.ShowOpen
    If dlgArchivo.FileName <> Empty Then
        txtArchivo.Text = dlgArchivo.FileName
    Else
        txtArchivo.Text = "NO SE ABRIO NINGUN ARCHIVO"
    End If
End Sub

Private Sub cmdProcesar_Click()
    Dim oNCOMColocEval As NCOMColocEval
    Set oNCOMColocEval = New NCOMColocEval
    Set rsClienteSunat = New ADODB.Recordset
   'Variable de tipo Aplicación de Excel
    Dim objExcel As Excel.Application

   'Una variable de tipo Libro de Excel
    Dim xLibro As Excel.Workbook
    Dim xlHoja As Excel.Worksheet
    Dim Col As Integer, Fila As Integer
    Dim NroDoc As String
   'creamos un nuevo objeto excel
    Set objExcel = New Excel.Application

   'Usamos el método open para abrir el archivo que está en el directorio del programa llamado archivo.xls
    Set xLibro = objExcel.Workbooks.Open(txtArchivo.Text)

    PB1.Min = 0
    PB1.Max = 3
    PB1.value = 0
    PB1.Visible = True

   'Hacemos el Excel Visible
    objExcel.Visible = False

    With xLibro
    PB1.value = 1
       ' Hacemos referencia a la Hoja
        With .Sheets(1)
    Select Case cmbTipoReqSunat.ListIndex
        Case 0 'AMPLIACION/REDUCCION DE EMBARGO
           'Recorremos la fila desde la 1 hasta la 1000
            .Cells(1, 12) = "Tiene Cuenta"
            .Cells(1, 12).Borders.LineStyle = 1
            For Fila = 2 To 1000
           'Agregamos el valor de la fila que corresponde a la columna 12
            NroDoc = .Cells(Fila, 5)
                If NroDoc <> "" Then
                    If oNCOMColocEval.ValidaCuentaCliente(NroDoc) Then
                        .Cells(Fila, 12) = "SI"
                    Else
                        .Cells(Fila, 12) = "No"
                    End If
                    NroDoc = 0
                End If
            Next
        Case 1 'RETENCION BANCARIA
           'Recorremos la fila desde la 1 hasta la 1000 - "Tiene Cuenta"
           .Cells(1, 10) = "Tiene Cuenta"
           .Cells(1, 10).Borders.LineStyle = 1
            For Fila = 2 To 1000
           'Agregamos el valor de la fila que corresponde a la columna 10
            NroDoc = .Cells(Fila, 4)
                If NroDoc <> "" Then
                    If oNCOMColocEval.ValidaCuentaCliente(NroDoc) Then
                        .Cells(Fila, 10) = "Si"
                    Else
                        .Cells(Fila, 10) = "No"
                    End If
                    NroDoc = 0
                End If
            Next
           'Recorremos la fila desde la 1 hasta la 1000 - "Sum.CTA.Ahorros"
           .Cells(1, 11) = "Sum.CTA.Ahorros"
           .Cells(1, 11).Borders.LineStyle = 1
            For Fila = 2 To 1000
           'Agregamos el valor de la fila que corresponde a la columna 11
            NroDoc = .Cells(Fila, 4)
                If NroDoc <> "" Then
                    Set rsClienteSunat = oNCOMColocEval.obtenerSaldoClienteSunat(NroDoc, 1)
                        If rsClienteSunat.RecordCount > 0 Then
                            .Cells(Fila, 11) = rsClienteSunat!Saldo
                        Else
                            .Cells(Fila, 11) = "0.00"
                        End If
                    NroDoc = 0
                    Set rsClienteSunat = Nothing
                End If
            Next
           'Recorremos la fila desde la 1 hasta la 1000 - "Sum.CTA.PF"
           .Cells(1, 12) = "Sum.CTA.PF"
           .Cells(1, 12).Borders.LineStyle = 1
            For Fila = 2 To 1000
           'Agregamos el valor de la fila que corresponde a la columna 12
            NroDoc = .Cells(Fila, 4)
                If NroDoc <> "" Then
                    Set rsClienteSunat = oNCOMColocEval.obtenerSaldoClienteSunat(NroDoc, 2)
                        If rsClienteSunat.RecordCount > 0 Then
                            .Cells(Fila, 12) = rsClienteSunat!Saldo
                        Else
                            .Cells(Fila, 12) = "0.00"
                        End If
                    NroDoc = 0
                    Set rsClienteSunat = Nothing
                End If
            Next
           'Recorremos la fila desde la 1 hasta la 1000 - "Sum.CTA.PF"
           .Cells(1, 13) = "Garantiza"
           .Cells(1, 14) = "Saldo"
           .Range("M1", "N1").Borders.LineStyle = 1
            For Fila = 2 To 1000
           'Agregamos el valor de la fila que corresponde a la columna 13 Y 14
            NroDoc = .Cells(Fila, 4)
                If NroDoc <> "" Then
                    Set rsClienteSunat = oNCOMColocEval.obtenerGaranSaldoClienteSunat(NroDoc)
                        If rsClienteSunat.RecordCount > 0 Then
                            .Cells(Fila, 13) = rsClienteSunat!nGravament
                            .Cells(Fila, 14) = rsClienteSunat!Disponible
                        Else
                            .Cells(Fila, 13) = "0.00"
                            .Cells(Fila, 14) = "0.00"
                        End If
                    NroDoc = 0
                    Set rsClienteSunat = Nothing
                End If
            Next
        Case 2 'LEVANTAMIENTO
           'Recorremos la fila desde la 1 hasta la 1000
           .Cells(1, 10) = "Tiene Cuenta"
           .Cells(1, 10).Borders.LineStyle = 1
            For Fila = 2 To 1000
           'Agregamos el valor de la fila que corresponde a la columna 10
                NroDoc = .Cells(Fila, 5)
                If NroDoc <> "" Then
                    If oNCOMColocEval.ValidaCuentaCliente(NroDoc) Then
                        .Cells(Fila, 10) = "Si"
                    Else
                        .Cells(Fila, 10) = "No"
                    End If
                    NroDoc = 0
                End If
            Next

           .Cells(1, 11) = "Bloqueado"
           .Cells(1, 11).Borders.LineStyle = 1
            For Fila = 2 To 1000
           'Agregamos el valor de la fila que corresponde a la columna 11
                NroDoc = .Cells(Fila, 5)
                If NroDoc <> "" Then
                    If oNCOMColocEval.ValidaClienteBloqueado(NroDoc) Then
                        .Cells(Fila, 11) = "Si"
                    Else
                        .Cells(Fila, 11) = "No"
                    End If
                    NroDoc = 0
                End If
            Next
    End Select
    PB1.value = 2
        End With
    End With
    MsgBox "La comprobación se realizo en forma correcta ", vbInformation, "Aviso"
    PB1.value = 3
    If MsgBox("Desea Realizar otro proceso?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        Me.txtArchivo.Text = ""
        PB1.Visible = False
        cmbTipoReqSunat.ListIndex = 0
        objExcel.Visible = True
        Set objExcel = Nothing
        Set xLibro = Nothing
    Else
        objExcel.Visible = True
        Set objExcel = Nothing
        Set xLibro = Nothing
        Unload Me
    End If
End Sub

Private Sub Form_Load()
CargarComboTipoArchivo
End Sub

Private Sub CargarComboTipoArchivo()
Dim oConstante As COMDConstantes.DCOMConstantes
Set oConstante = New COMDConstantes.DCOMConstantes
Set rsTipoArchivo = New ADODB.Recordset

Set rsTipoArchivo = oConstante.ObtenerConstReqSUNAT()
    If Not rsTipoArchivo.BOF And Not rsTipoArchivo.EOF Then
        Do While Not rsTipoArchivo.EOF
            cmbTipoReqSunat.AddItem (rsTipoArchivo!cConsDescripcion)
            rsTipoArchivo.MoveNext
        Loop
        cmbTipoReqSunat.ListIndex = 0
    End If
Set rsTipoArchivo = Nothing
End Sub
