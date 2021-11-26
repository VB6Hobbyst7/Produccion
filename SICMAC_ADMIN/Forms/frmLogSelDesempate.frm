VERSION 5.00
Begin VB.Form frmLogSelDesempate 
   Caption         =   "Desempate "
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10710
   Icon            =   "frmLogSelDesempate.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   3870
   ScaleWidth      =   10710
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdcancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   5880
      TabIndex        =   2
      Top             =   3480
      Width           =   1575
   End
   Begin VB.CommandButton cmdaceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3480
      Width           =   1575
   End
   Begin Sicmact.FlexEdit FlexEmpates 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   240
      Width           =   10575
      _ExtentX        =   18653
      _ExtentY        =   5741
      Cols0           =   10
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Item-Codigo Bien-Descripcion Bien-Codigo Persona-Descripcion Proveedor-Punt.Econ.-Punt.Tecn.-Punt. Total-Ganador-Comentarios"
      EncabezadosAnchos=   "450-0-2000-0-2000-1000-1000-1000-900-2100"
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
      ColumnasAEditar =   "X-X-X-X-X-X-X-X-8-9"
      ListaControles  =   "0-0-0-0-0-0-0-0-4-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-L-L-L-L-R-R-C-L-L"
      FormatosEdit    =   "0-0-0-0-0-2-2-2-0-0"
      TextArray0      =   "Item"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      ColWidth0       =   450
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmLogSelDesempate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim clsDGAdqui As DLogAdquisi


Private Sub cmdAceptar_Click()
Dim nTotalReg As Integer
Dim nTotalMarcados As Integer
Dim nTotalSinMarca As Integer
nTotalReg = FlexEmpates.Rows - 1
nTotalMarcados = 0
nTotalSinMarca = 0
Dim sActualiza As String
Dim nestadoProc As Integer

nestadoProc = clsDGAdqui.CargaLogSelEstadoProceso(frmLogSelEvalTecResumen.txtSeleccionA.Text)
If nestadoProc = SelEstProcesoCerrado Then
        MsgBox "El Procesos de Seleccion  " + frmLogSelEvalTecResumen.txtSeleccionA.Text + " ya esta Cerrado", vbInformation, "Estado del proceso " + frmLogSelEvalTecResumen.txtSeleccionA.Text + " ya esta Cerrado"
        Exit Sub
End If


sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)

For i = 1 To FlexEmpates.Rows - 1
        If FlexEmpates.TextMatrix(i, 8) = "." Then 'con check
            nTotalMarcados = nTotalMarcados + 1
        End If
        
        If FlexEmpates.TextMatrix(i, 8) = "" Then 'con check
        nTotalSinMarca = nTotalSinMarca + 1
        End If
        

Next

If nTotalSinMarca = nTotalReg / 2 And nTotalMarcados = nTotalReg / 2 Then   'ok
   Else
   MsgBox "Debe Existir un  Ganador por Bien", vbInformation, "Debe Existir un Solo Ganador por Bien"
   Exit Sub
End If

For i = 1 To FlexEmpates.Rows - 1
        
        If i = FlexEmpates.Rows - 1 Then
        Exit For
        End If
        
        If FlexEmpates.TextMatrix(i, 1) = FlexEmpates.TextMatrix(i + 1, 1) Then 'con check
        If FlexEmpates.TextMatrix(i, 8) = FlexEmpates.TextMatrix(i + 1, 8) Then 'con check
            MsgBox "Debe existir  ganador unico por codigo de bien ", vbInformation, "Debe existir  ganador unico por codigo de bien"
            Exit Sub
            
        End If
        
        End If
Next

For i = 1 To FlexEmpates.Rows - 1
        If FlexEmpates.TextMatrix(i, 8) = "." Then 'con check
                If FlexEmpates.TextMatrix(i, 9) = "" Then
                    MsgBox "Debe Ingresar un Comentario sobre la Eleccion", vbInformation, "Debe ingresar un comentario sobre la eleccion"
                    Exit Sub
                End If
        End If
Next


'sActualiza
'Registrar Ganador
'Marcar ambos para el Log
'pnLogSelNro As Long, pscBSCod As String, psCodPer As String, psDescripcion As String, psActualiza As String

clsDGAdqui.EliminaLogDesenmpate frmLogSelEvalTecResumen.txtSeleccionA.Text
For i = 1 To FlexEmpates.Rows - 1
        clsDGAdqui.InsertaLogDesenmpate frmLogSelEvalTecResumen.txtSeleccionA.Text, FlexEmpates.TextMatrix(i, 1), FlexEmpates.TextMatrix(i, 3), FlexEmpates.TextMatrix(i, 9), sActualiza
        If FlexEmpates.TextMatrix(i, 8) = "." Then 'con check
           clsDGAdqui.ActualizaLogDesenmpate frmLogSelEvalTecResumen.txtSeleccionA.Text, FlexEmpates.TextMatrix(i, 1), FlexEmpates.TextMatrix(i, 3), "SI", sActualiza
        End If
        If FlexEmpates.TextMatrix(i, 8) = "" Then 'sin check
           clsDGAdqui.ActualizaLogDesenmpate frmLogSelEvalTecResumen.txtSeleccionA.Text, FlexEmpates.TextMatrix(i, 1), FlexEmpates.TextMatrix(i, 3), "NO", sActualiza
        End If
        
Next

frmLogSelEvalTecResumen.Mostrar_Resumen_Cuadro frmLogSelEvalTecResumen.txtSeleccionA.Text


Unload Me

End Sub

Private Sub cmdCancelar_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim rs As ADODB.Recordset
Set rs = New ADODB.Recordset
Set clsDGAdqui = New DLogAdquisi
'nvalida = clsDGAdqui.ValidaLogSelEmpate(txtSeleccionA.Text)
Set rs = clsDGAdqui.CargalogSelEmpates(frmLogSelEvalTecResumen.txtSeleccionA.Text, 1)

If rs.EOF = True Then
    Set rs = clsDGAdqui.CargalogSelEmpates(frmLogSelEvalTecResumen.txtSeleccionA.Text, 0)
End If

Set FlexEmpates.Recordset = rs
Me.Width = 10830
Me.Height = 4380

End Sub

