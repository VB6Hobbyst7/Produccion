VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredSolicAfp 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Registro de Solicitud del 25% AFP"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   13815
   Icon            =   "frmCredSolicAfp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   13815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin TabDlg.SSTab SSTab1 
      Height          =   6375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   13575
      _ExtentX        =   23945
      _ExtentY        =   11245
      _Version        =   393216
      Tabs            =   1
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "Solicitud: 25% del AFP"
      TabPicture(0)   =   "frmCredSolicAfp.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "feCredSoliAfp"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtBuscar"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "cmdCerrar"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "cmdActualizar"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "cmdEditar"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdNuevo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Command1"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Command2"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdeliminar"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton cmdeliminar 
         Caption         =   "&Eliminar"
         Height          =   375
         Left            =   2520
         TabIndex        =   10
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton Command2 
         Caption         =   "E&xportar a Excel"
         Height          =   375
         Left            =   5880
         TabIndex        =   9
         Top             =   5880
         Width           =   1455
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Buscar"
         Height          =   255
         Left            =   4320
         TabIndex        =   2
         Top             =   480
         Width           =   855
      End
      Begin VB.CommandButton cmdNuevo 
         Caption         =   "&Nuevo"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "&Editar"
         Height          =   375
         Left            =   1320
         TabIndex        =   5
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdActualizar 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   10920
         TabIndex        =   6
         Top             =   5880
         Width           =   1215
      End
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "Cerrar"
         Height          =   375
         Left            =   12120
         TabIndex        =   7
         Top             =   5880
         Width           =   1215
      End
      Begin VB.TextBox txtBuscar 
         Height          =   285
         Left            =   960
         TabIndex        =   1
         Top             =   480
         Width           =   3255
      End
      Begin SICMACT.FlexEdit feCredSoliAfp 
         Height          =   4980
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   13290
         _ExtentX        =   23442
         _ExtentY        =   8784
         Cols0           =   9
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Agencia-Crédito-Titular-AFP-Imp. Disponible-F. Carta-Destino-F. Abono"
         EncabezadosAnchos=   "400-2000-1800-2900-1200-1400-950-1400-950"
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
         ColumnasAEditar =   "X-1-2-3-X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-C-R-C-C-C"
         FormatosEdit    =   "0-0-0-0-0-2-5-0-5"
         CantEntero      =   10
         TextArray0      =   "#"
         SelectionMode   =   1
         lbUltimaInstancia=   -1  'True
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   405
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Titular:"
         Height          =   255
         Left            =   360
         TabIndex        =   8
         Top             =   480
         Width           =   495
      End
   End
End
Attribute VB_Name = "frmCredSolicAfp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fBotonPulsado As Integer

Private rsCreditoTotal As ADODB.Recordset

Private Sub cmdActualizar_Click()
txtBuscar.Text = ""
Call LlenarGrilla
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub CmdEditar_Click()
  fBotonPulsado = 2
  frmCredSolicAfpClie.lblvalor.Caption = 2
  
  If feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 1) = "" Then
  
    MsgBox "Seleccione Correctamente un Crédito", vbInformation, "AVISO"
    
   
  ElseIf feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 1) <> "" Then
  
  frmCredSolicAfpClie.txtCuenta.CMAC = Mid(feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 2), 1, 3)
  frmCredSolicAfpClie.txtCuenta.Age = Mid(feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 2), 4, 5)
  frmCredSolicAfpClie.txtCuenta.Prod = Mid(feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 2), 6, 8)
  frmCredSolicAfpClie.txtCuenta.Cuenta = Mid(feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 2), 9, 18)
  frmCredSolicAfpClie.cmbDestino.Text = feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 7)
  frmCredSolicAfpClie.LblCliente.Caption = feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 3)
  frmCredSolicAfpClie.cmbAfp.Text = feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 4)
  frmCredSolicAfpClie.txtFechaCarta.Text = Format(feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 6), "dd/mm/yyyy")
  frmCredSolicAfpClie.edtDisponible.Text = feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 5)
  frmCredSolicAfpClie.txtFechaAbono.Text = Format(feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 8), "dd/mm/yyyy")
      
  frmCredSolicAfpClie.Show 1
  
  
  End If
  
  
  
  
End Sub

Private Sub CmdEliminar_Click()
Dim oBuscarafp As COMDCredito.DCOMCredito
Dim lrsBusca As ADODB.Recordset
Dim sCtaCod As String


If feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 1) = "" Then
  
    MsgBox "Seleccione Correctamente un Crédito", vbInformation, "AVISO"
    
   
  ElseIf feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 1) <> "" Then
  
        sCtaCod = feCredSoliAfp.TextMatrix(feCredSoliAfp.row, 2)
        Set oBuscarafp = New COMDCredito.DCOMCredito
        Set lrsBusca = oBuscarafp.OperacionesVarios(sCtaCod, , , , , , 3)
        Set oBuscarafp = Nothing
        
        txtBuscar.Text = ""
        Call LlenarGrilla
        
  End If

End Sub

Private Sub cmdNuevo_Click()
  fBotonPulsado = 1
  frmCredSolicAfpClie.lblvalor = 1
  frmCredSolicAfpClie.Show 1
End Sub

Private Sub Command1_Click()

If txtBuscar.Text <> "" Then

Call LlenarGrilla
Else
  MsgBox "Ingrese datos del cliente a Buscar", vbInformation, "AVISO"
  txtBuscar.SetFocus
End If
End Sub

Private Sub Command2_Click()
Call exportar_Datagrid(feCredSoliAfp, feCredSoliAfp.rows)
End Sub

Public Sub inicio()
Dim validarUsuario As COMDCredito.DCOMCredito
Dim rsvalida As ADODB.Recordset

Set validarUsuario = New COMDCredito.DCOMCredito
Set rsvalida = validarUsuario.validausuario(gsCodUser)

If rsvalida.RecordCount = 0 Then MsgBox "Ud. No cuenta con Permisos para este Módulo", vbInformation, "AVISO": Exit Sub

Me.Show 1

End Sub

Private Sub Form_Load()

fBotonPulsado = 0
Call LlenarGrilla

End Sub

Sub LlenarGrilla()

Dim oCredito As COMDCredito.DCOMCredito
Dim rsCredito As ADODB.Recordset

Set oCredito = New COMDCredito.DCOMCredito

Set rsCredito = oCredito.MostrarTodo(txtBuscar.Text, 2)
Set rsCreditoTotal = rsCredito.Clone

Call MostrarLista(rsCredito)

End Sub

Sub MostrarLista(ByVal pRs As ADODB.Recordset)
Dim i As Integer

  Call LimpiaFlex(feCredSoliAfp)
  
  If pRs.RecordCount > 0 Then
  
        For i = 1 To pRs.RecordCount
        
           feCredSoliAfp.AdicionaFila
           feCredSoliAfp.TextMatrix(i, 0) = i
           feCredSoliAfp.TextMatrix(i, 1) = pRs!Agencia
           feCredSoliAfp.TextMatrix(i, 2) = pRs!cCtaCod
           feCredSoliAfp.TextMatrix(i, 3) = pRs!cPersNombre
           feCredSoliAfp.TextMatrix(i, 4) = pRs!cAfp
           feCredSoliAfp.TextMatrix(i, 5) = FormatNumber(pRs!nImpDisp, 2)
           feCredSoliAfp.TextMatrix(i, 6) = pRs!dFecCarta
           feCredSoliAfp.TextMatrix(i, 7) = pRs!cDestino
           feCredSoliAfp.TextMatrix(i, 8) = pRs!dFecAbono
           
           pRs.MoveNext
        Next i
 ' Else
   '  MsgBox "No se encontro información para mostrar", vbInformation, "AVISO"
   '  txtBuscar.SetFocus
  End If
  
Set pRs = Nothing

End Sub

Private Sub exportar_Datagrid(Datagrid As FlexEdit, n_Filas As Long)
  
  ' -- Variables para Excel
Dim Obj_Excel   As Object
Dim Obj_Libro   As Object
Dim Obj_Hoja    As Object
Dim iCol As Integer
Dim Ruta, RutaSave As String

Ruta = App.Path & "\FormatoCarta\" & "lista_AFP_25porc.xls"
RutaSave = App.Path & "\spooler\" & "lista_AFP_25porc" & Format(Date, "ddMMyyyy") & ".xls"
    On Error GoTo error_handler
      
    Dim i   As Integer
    Dim j   As Integer
      
    ' -- Colocar el cursor de espera mientras se exportan los datos
    Me.MousePointer = vbHourglass
      
    If n_Filas = 0 Then
        MsgBox "No hay datos para exportar a excel ": Exit Sub
    Else
          
        ' -- Crear nueva instancia de Excel
        Set Obj_Excel = CreateObject("Excel.Application")
        'Obj_Excel.Visible = False
        ' -- Agregar nuevo libro
        Set Obj_Libro = Obj_Excel.Workbooks.Open(Ruta)
        Obj_Libro.SaveAs (RutaSave)
      
        ' -- Referencia a la Hoja activa ( la que añade por defecto Excel )
        Set Obj_Hoja = Obj_Excel.ActiveSheet
     
        iCol = 0
        ' --  Recorrer el Datagrid ( Las columnas )
        For i = 0 To Datagrid.cols - 1
                ' -- Incrementar índice de columna
                iCol = iCol + 1
                ' -- Obtener el caption de la columna
                'Obj_Hoja.Cells(1, iCol) = Datagrid.TextMatrix(1, i)
                ' -- Recorrer las filas
                For j = 1 To n_Filas - 1
                    ' -- Asignar el valor a la celda del Excel
                    Obj_Hoja.Cells(j + 3, iCol) = _
                    Datagrid.TextMatrix(j, i)
                    'Datagrid.Columns(i).CellValue (Datagrid.GetBookmark(j))
                Next
        Next
          
        ' -- Hacer excel visible
        Obj_Excel.Visible = True
          
        ' -- Opcional : colocar en negrita y de color rojo los enbezados en la hoja
    '    With Obj_Hoja
    '        .rows(1).Font.Bold = True
    '        .rows(1).Font.Color = vbRed
    '        ' -- Autoajustar las cabeceras
    '        .Columns("A:Z").AutoFit
    '    End With
    End If
  
    'obj_Libro.saveas()
  
    ' -- Eliminar las variables de objeto excel
    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing
      
    ' -- Restaurar cursor
    Me.MousePointer = vbDefault
      
Exit Sub
  
' -- Error
error_handler:
  
If Err.Number = 1004 Then MsgBox "La Operación fue Cancelada por el Usuario", vbInformation, "AVISO"
  '  MsgBox Err.Description, vbCritical
    On Error Resume Next
  
    Set Obj_Hoja = Nothing
    Set Obj_Libro = Nothing
    Set Obj_Excel = Nothing
    Me.MousePointer = vbDefault
  
End Sub

