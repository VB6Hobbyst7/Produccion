VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmFSDTiposClientes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IDENTIFICACION DE TIPOS DE CLIENTES"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9165
   Icon            =   "frmFSDTiposClientes.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   9165
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExportar 
      Cancel          =   -1  'True
      Caption         =   "&Exportar"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   120
      TabIndex        =   12
      Top             =   6960
      Width           =   1410
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   7680
      TabIndex        =   2
      Top             =   6960
      Width           =   1410
   End
   Begin VB.Frame Frame2 
      Caption         =   "Listado de Registros con Información TIPO 2 Generados en Tabla"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   3600
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   8985
      Begin MSDataGridLib.DataGrid dgPersonas 
         Height          =   2805
         Left            =   165
         TabIndex        =   1
         Top             =   300
         Width           =   8685
         _ExtentX        =   15319
         _ExtentY        =   4948
         _Version        =   393216
         AllowUpdate     =   0   'False
         HeadLines       =   1
         RowHeight       =   17
         FormatLocked    =   -1  'True
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   6
         BeginProperty Column00 
            DataField       =   "cIdentificacion"
            Caption         =   "Tipo Cliente"
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
            DataField       =   "cTipo"
            Caption         =   "Clasif."
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
            DataField       =   "cCodPers"
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
         BeginProperty Column03 
            DataField       =   "cNomPers"
            Caption         =   "Nombre"
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
            DataField       =   "nPersMotivo"
            Caption         =   "nPersMotivo"
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
            DataField       =   "cAgencia"
            Caption         =   "Agencia"
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
            MarqueeStyle    =   4
            BeginProperty Column00 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column01 
               ColumnWidth     =   599.811
            EndProperty
            BeginProperty Column02 
               ColumnWidth     =   1470.047
            EndProperty
            BeginProperty Column03 
               ColumnWidth     =   4020.095
            EndProperty
            BeginProperty Column04 
               Object.Visible         =   0   'False
            EndProperty
            BeginProperty Column05 
               ColumnWidth     =   2009.764
            EndProperty
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Registro(s) Encontrado(s)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   1350
         TabIndex        =   7
         Top             =   3210
         Width           =   2130
      End
      Begin VB.Label lblCantidad 
         Alignment       =   1  'Right Justify
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   165
         TabIndex        =   6
         Top             =   3165
         Width           =   1080
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3075
      Left            =   105
      TabIndex        =   3
      Top             =   30
      Width           =   9000
      Begin VB.OptionButton optOpcion 
         Caption         =   "&Ninguno"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   0
         Left            =   2760
         TabIndex        =   11
         Top             =   600
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "&Todos"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Index           =   1
         Left            =   4200
         TabIndex        =   10
         Top             =   600
         Width           =   825
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Generar Tabla &Información Tipo 2"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   2760
         TabIndex        =   9
         Top             =   2520
         Width           =   3120
      End
      Begin VB.ListBox lstOpciones 
         Height          =   1635
         ItemData        =   "frmFSDTiposClientes.frx":030A
         Left            =   120
         List            =   "frmFSDTiposClientes.frx":0323
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   840
         Width           =   8760
      End
      Begin MSComCtl2.Animation Logo 
         Height          =   645
         Left            =   195
         TabIndex        =   4
         Top             =   165
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   1138
         _Version        =   393216
         FullWidth       =   45
         FullHeight      =   43
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Registros con Información TIPO 2"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1755
         TabIndex        =   8
         Top             =   195
         Width           =   3165
      End
   End
End
Attribute VB_Name = "frmFSDTiposClientes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oCon As NAnx_FSD
Dim frsTipoCliente As New ADODB.Recordset

Private Sub cmdAceptar_Click()
Dim rs As ADODB.Recordset
Dim i As Integer
Dim nCont As Integer
Dim lsCondElimina As String
On Error GoTo ErrorAceptar
nCont = 0
Set dgPersonas.DataSource = Nothing
For i = 0 To lstOpciones.ListCount - 1
    If lstOpciones.Selected(i) = True Then
        lsCondElimina = lsCondElimina & (i + 1) & ","
        nCont = nCont + 1
    End If
Next
If nCont > 0 Then
    Screen.MousePointer = vbHourglass
Else
    MsgBox "Seleccione un criterio para mostrar", vbExclamation, "Aviso!!!"
    Exit Sub
End If
lsCondElimina = Left(lsCondElimina, Len(lsCondElimina) - 1)

Set oCon = New NAnx_FSD
oCon.GetClientesTipo2 lstOpciones.Selected(0), lstOpciones.Selected(1), lstOpciones.Selected(2), lstOpciones.Selected(3), lstOpciones.Selected(4), lstOpciones.Selected(5), lstOpciones.Selected(6), lsCondElimina
CargaDatos lstOpciones.ListIndex + 1
Set oCon = Nothing
Screen.MousePointer = vbDefault

If Val(lblCantidad.Caption) = 0 Then
    MsgBox "No se encontraron registros", vbInformation, "Aviso!!!"
Else
    MsgBox "Se generó tabla con Informacion Tipo 2 satisfactoriamente", vbInformation, "Aviso!!!"
End If

Exit Sub
ErrorAceptar:
    Screen.MousePointer = vbDefault
    MsgBox Err.Description, vbExclamation, "Aviso!!!"
End Sub
 
Private Sub cmdExportar_Click()
generarTipoCliente frsTipoCliente
End Sub

Private Sub cmdSalir_Click()
    Set frsTipoCliente = Nothing
    Unload Me
End Sub

Public Sub CargaDatos(psMotivo As Integer)
Dim rs As ADODB.Recordset
    Set oCon = New NAnx_FSD
    Set rs = New ADODB.Recordset
    '***Modificado por ELRO el 20130507, según TI-ERS019-2013****
    'Set rs = ocon.CargaClientesFSD(psMotivo)
    Set rs = oCon.CargaClientesFSD(psMotivo, optOpcion.Item(1))
    Set frsTipoCliente = Nothing
    Set frsTipoCliente = rs
    '***Fin Modificado por ELRO el 20130507, según TI-ERS019-2013
    Set dgPersonas.DataSource = rs
    lblCantidad.Caption = rs.RecordCount
    Set rs = Nothing
    Set oCon = Nothing
    Screen.MousePointer = vbDefault
    
    '***Modificado por ELRO el 20130507, según TI-ERS019-2013****
    'Label2.Caption = "Registro(s) Encontrado(s) con Clasificación " & psMotivo & " para el Tipo 2"
    If optOpcion.Item(1) Then
        Label2.Caption = "Registro(s) Encontrado(s) para el Tipo 2"
    Else
        Label2.Caption = "Registro(s) Encontrado(s) con Clasificación " & psMotivo & " para el Tipo 2"
    End If
    '***Fin Modificado por ELRO el 20130507, según TI-ERS019-2013

End Sub

Private Sub Form_Load()
CentraForm Me
frmFondoSeguroDep.Enabled = False
Logo.AutoPlay = True
Logo.Open App.path & "\videos\LogoA.avi"
CargaDatos 1
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmFondoSeguroDep.Enabled = True
End Sub

Private Sub optSeleccion_Click(Index As Integer)
Dim i As Integer
    For i = 0 To lstOpciones.ListCount - 1
        If Index = 0 Then
            lstOpciones.Selected(i) = True
        Else
            lstOpciones.Selected(i) = False
        End If
    Next
End Sub

Private Sub lstOpciones_Click()
    CargaDatos lstOpciones.ListIndex + 1
End Sub

Private Sub lstOpciones_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    cmdAceptar.SetFocus
End If
End Sub
 
Private Sub optOpcion_Click(Index As Integer)

    Dim i As Integer
    
    For i = 0 To lstOpciones.ListCount - 1
        lstOpciones.Selected(i) = IIf(Index = 0, False, True)
    Next
    
End Sub

Private Sub generarTipoCliente(ByVal prsTipoCliente As ADODB.Recordset)
    Dim xlAplicacion As Excel.Application
    Dim xlLibro As Excel.Workbook
    Dim xlHoja1 As Excel.Worksheet
    Dim lsArchivo As String
    Dim lsArchivo1, lsArchivo2 As String
    Dim lnFilaFin As Integer
    Dim lsHoja As String
    Dim fs As Scripting.FileSystemObject
    Dim lsNombreAgencia As String
    Dim i As Integer
    
    Set fs = New Scripting.FileSystemObject
    Set xlAplicacion = New Excel.Application
    
    lsArchivo = "ClienteTipo2"
    lsArchivo1 = Format(gdFecSis, "yyyymmdd") & "_" & Format(Time(), "HHMMSS")
    lsArchivo2 = "\SPOOLER\" & lsArchivo & "_" & lsArchivo1 & ".xls"
    
    If fs.FileExists(App.path & "\FormatoCarta\" & lsArchivo & ".xls") Then
        Set xlLibro = xlAplicacion.Workbooks.Open(App.path & "\FormatoCarta\" & lsArchivo & ".xls")
    
    Else
        MsgBox "No existe la plantilla ClienteTipo2.xls en la carpeta FormatoCarta, Consulte con el Área de  TI", vbInformation, "Advertencia"
        Exit Sub
    End If
    
        
    For Each xlHoja1 In xlLibro.Worksheets
       
        xlHoja1.Activate
        
        lnFilaFin = 4
        prsTipoCliente.MoveFirst
        
        If Not (prsTipoCliente.BOF And prsTipoCliente.EOF) Then
            Do While Not prsTipoCliente.EOF
                xlHoja1.Cells(lnFilaFin, 1).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 1) = UCase(prsTipoCliente("cTipo"))
                xlHoja1.Cells(lnFilaFin, 2).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 2) = UCase(prsTipoCliente("cCodPers"))
                xlHoja1.Cells(lnFilaFin, 3).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 3) = UCase(prsTipoCliente("cNomPers"))
                xlHoja1.Cells(lnFilaFin, 4).Borders.LineStyle = 1
                xlHoja1.Cells(lnFilaFin, 4) = UCase(prsTipoCliente("cAgencia"))
                prsTipoCliente.MoveNext
                lnFilaFin = lnFilaFin + 1
            Loop
        End If
        
        lnFilaFin = 4
        For i = 0 To lstOpciones.ListCount - 1
            xlHoja1.Cells(lnFilaFin, 6).Borders.LineStyle = 1
            xlHoja1.Cells(lnFilaFin, 6) = UCase(lstOpciones.List(i))
            lnFilaFin = lnFilaFin + 1
        Next i
    Next

    Set prsTipoCliente = Nothing
    
    xlLibro.SaveAs (App.path & lsArchivo2)
    xlLibro.Close
    xlAplicacion.Quit
    Set xlAplicacion = Nothing
    Set xlLibro = Nothing
    Set xlHoja1 = Nothing
    
    CargaArchivo lsArchivo & "_" & lsArchivo1 & ".xls", App.path & "\SPOOLER\"
    Exit Sub
ErrImprime:
     MsgBox TextErr(Err.Description), vbInformation, "¡Aviso!"
          xlLibro.Close
          xlAplicacion.Quit
       Set xlAplicacion = Nothing
       Set xlLibro = Nothing
       Set xlHoja1 = Nothing

End Sub

