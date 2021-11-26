VERSION 5.00
Begin VB.Form frmCapParamComisionOtrasAge 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tarifario de Comisiones de Operaciones en Otras Agencias"
   ClientHeight    =   8445
   ClientLeft      =   2925
   ClientTop       =   2250
   ClientWidth     =   9750
   Icon            =   "frmCapParamComisionOtrasAge.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   9750
   Begin VB.Frame frmAplicar 
      Caption         =   "Datos a Aplicar a la Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1095
      Left            =   120
      TabIndex        =   13
      Top             =   3240
      Width           =   7335
      Begin VB.TextBox txtComision 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   120
         TabIndex        =   17
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtMontoMin 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1920
         TabIndex        =   16
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtMontoMax 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3720
         TabIndex        =   15
         Top             =   600
         Width           =   1695
      End
      Begin VB.CommandButton cmdAplicar 
         Caption         =   "Aplicar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         TabIndex        =   14
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Comision %"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label2 
         Caption         =   "Monto Min"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   1920
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Monto Maximo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   3720
         TabIndex        =   18
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Tarifario"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3855
      Left            =   120
      TabIndex        =   11
      Top             =   4440
      Width           =   9495
      Begin SICMACT.FlexEdit fgTarifario 
         Height          =   3495
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   9255
         _extentx        =   16325
         _extenty        =   6165
         cols0           =   9
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-OK-Agencia-Tipo Ahorro-Comision %-Monto Min-Monto Max-cAgeCod-nTpoPrograma"
         encabezadosanchos=   "350-350-2000-3000-900-1100-1100-0-0"
         font            =   "frmCapParamComisionOtrasAge.frx":030A
         font            =   "frmCapParamComisionOtrasAge.frx":0336
         font            =   "frmCapParamComisionOtrasAge.frx":0362
         font            =   "frmCapParamComisionOtrasAge.frx":038E
         fontfixed       =   "frmCapParamComisionOtrasAge.frx":03BA
         backcolorcontrol=   12058623
         backcolorcontrol=   12058623
         backcolorcontrol=   12058623
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-1-X-X-4-5-6-X-X"
         listacontroles  =   "0-4-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-L-L-R-R-R-C-C"
         formatosedit    =   "0-0-0-0-2-2-2-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1  'True
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   345
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Opcion  Grabar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4215
      Left            =   7560
      TabIndex        =   6
      Top             =   120
      Width           =   2055
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   1320
         Width           =   1575
      End
      Begin VB.CommandButton Cancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   240
         TabIndex        =   8
         Top             =   840
         Width           =   1575
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   1575
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opciones de Busqueda"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   3015
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
      Begin VB.ComboBox cboMoneda 
         Height          =   315
         ItemData        =   "frmCapParamComisionOtrasAge.frx":03E8
         Left            =   960
         List            =   "frmCapParamComisionOtrasAge.frx":03F2
         TabIndex        =   5
         Text            =   "--Moneda--"
         Top             =   360
         Width           =   2535
      End
      Begin VB.CheckBox chkTipoAhorro 
         Caption         =   "Todos los Tipos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   3600
         TabIndex        =   4
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkAgencias 
         Caption         =   "Todas las Agencias"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   2295
      End
      Begin SICMACT.FlexEdit fgAgencias 
         Height          =   1815
         Left            =   120
         TabIndex        =   1
         Top             =   1080
         Width           =   3375
         _extentx        =   5953
         _extenty        =   3201
         cols0           =   4
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-OK-Agencia-cAgeCod"
         encabezadosanchos=   "300-350-2300-0"
         font            =   "frmCapParamComisionOtrasAge.frx":040C
         font            =   "frmCapParamComisionOtrasAge.frx":0438
         font            =   "frmCapParamComisionOtrasAge.frx":0464
         font            =   "frmCapParamComisionOtrasAge.frx":0490
         fontfixed       =   "frmCapParamComisionOtrasAge.frx":04BC
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-1-X-X"
         listacontroles  =   "0-4-0-0"
         encabezadosalineacion=   "C-C-L-C"
         formatosedit    =   "0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1  'True
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin SICMACT.FlexEdit fgTipoAhorro 
         Height          =   1815
         Left            =   3600
         TabIndex        =   2
         Top             =   1080
         Width           =   3375
         _extentx        =   5953
         _extenty        =   3201
         cols0           =   4
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-OK-Tipo Ahorro-cTpoPrograma"
         encabezadosanchos=   "300-350-2300-0"
         font            =   "frmCapParamComisionOtrasAge.frx":04EA
         font            =   "frmCapParamComisionOtrasAge.frx":0516
         font            =   "frmCapParamComisionOtrasAge.frx":0542
         font            =   "frmCapParamComisionOtrasAge.frx":056E
         fontfixed       =   "frmCapParamComisionOtrasAge.frx":059A
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-1-X-X"
         listacontroles  =   "0-4-0-0"
         encabezadosalineacion=   "C-C-L-C"
         formatosedit    =   "0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1  'True
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.Label Label4 
         Caption         =   "Moneda:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   390
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmCapParamComisionOtrasAge"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lnOpcion As Integer 'Registrar=1;Consulta=2
Public Sub Inicio(ByVal pnOpcion As Integer)
    lnOpcion = pnOpcion
    Me.Show 1
End Sub
Private Sub Cancelar_Click()
    Me.txtComision.Text = ""
    Me.txtMontoMin.Text = ""
    Me.txtMontoMax.Text = ""
    Me.chkAgencias.value = 1
    Me.chkTipoAhorro.value = 1
    cargarControles
End Sub
Private Sub cboMoneda_Click()
    Me.txtComision.Text = ""
    Me.txtMontoMin.Text = ""
    Me.txtMontoMax.Text = ""
    Me.chkAgencias.value = 1
    Me.chkTipoAhorro.value = 1
    cargarControles
End Sub

Private Sub chkAgencias_Click()
    checkeaDatos Me.fgAgencias, Me.chkAgencias.value
    
    Me.txtComision.Text = ""
    Me.txtMontoMin.Text = ""
    Me.txtMontoMax.Text = ""
        
    If Not verificarDatoChecked(Me.fgAgencias) Then
        Call LimpiaFlex(Me.fgTarifario)
        Exit Sub
    End If

    If Not verificarDatoChecked(Me.fgTipoAhorro) Then
        Call LimpiaFlex(Me.fgTarifario)
        Exit Sub
    End If
        
   cargarTarifario
   
End Sub
Private Sub chkTipoAhorro_Click()
    checkeaDatos Me.fgTipoAhorro, Me.chkTipoAhorro.value
    Me.txtComision.Text = ""
    Me.txtMontoMin.Text = ""
    Me.txtMontoMax.Text = ""
        
    If Not verificarDatoChecked(Me.fgAgencias) Then
        Call LimpiaFlex(Me.fgTarifario)
        Exit Sub
    End If

    If Not verificarDatoChecked(Me.fgTipoAhorro) Then
        Call LimpiaFlex(Me.fgTarifario)
        Exit Sub
    End If
        
   cargarTarifario
End Sub

Private Sub cmdGrabar_Click()
    Me.txtComision.Text = ""
    Me.txtMontoMin.Text = ""
    Me.txtMontoMax.Text = ""
    
    If Not verificarDatoChecked(Me.fgTarifario) Then
         MsgBox "Debe Seleccionar al menos un Item del Tarifario", vbInformation, "AVISO!"
         Exit Sub
    End If
    
    If MsgBox("Se va ha Proceder a Guardar el Tarifario", vbYesNo, "AVISO!") = vbYes Then
        Dim objCap As New COMNCaptaGenerales.NCOMCaptaGenerales
        Dim objMov As New COMNContabilidad.NCOMContFunciones
        Dim sMovNro As String
        Dim sCadAgencia As String
        Dim sCadTpoPrograma As String
        
        Me.MousePointer = vbHourglass
        
        sMovNro = objMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
        sCadAgencia = obtenerCadenaDatos(Me.fgAgencias)
        sCadTpoPrograma = obtenerCadenaDatos(Me.fgTipoAhorro)
        
        If objCap.guardarTarifaOpeEnOtrasAge(Me.fgTarifario.GetRsNew, sCadAgencia, sCadTpoPrograma, Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex), sMovNro) Then
            If MsgBox("Se han Guardado los Datos con Exito!,Desea Realizar Otra Operacion", vbYesNo, "AVISO!") = vbYes Then
                cargarControles
            Else
                Unload Me
            End If
        Else
            MsgBox "No se han Podido Guardar los Datos, Verifique los Datos", vbInformation, "AVISO!"
        End If
        
        Me.MousePointer = vbDefault
    End If
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub fgAgencias_Click()
    
    If Me.fgAgencias.Col = 1 Then
        Me.txtComision.Text = ""
        Me.txtMontoMin.Text = ""
        Me.txtMontoMax.Text = ""
        
        If Not verificarDatoChecked(Me.fgAgencias) Then
           Call LimpiaFlex(Me.fgTarifario)
            Exit Sub
        End If
    
        If Not verificarDatoChecked(Me.fgTipoAhorro) Then
            Call LimpiaFlex(Me.fgTarifario)
            Exit Sub
        End If
        
        cargarTarifario
    End If
End Sub

Private Sub fgTipoAhorro_Click()
    
    
    If Me.fgTipoAhorro.Col = 1 Then
        Me.txtComision.Text = ""
        Me.txtMontoMin.Text = ""
        Me.txtMontoMax.Text = ""
        
        If Not verificarDatoChecked(Me.fgAgencias) Then
            Call LimpiaFlex(Me.fgTarifario)
            Exit Sub
        End If
    
        If Not verificarDatoChecked(Me.fgTipoAhorro) Then
            Call LimpiaFlex(Me.fgTarifario)
            Exit Sub
        End If
        
        cargarTarifario
    End If
End Sub
Private Sub Form_Load()
    Me.cboMoneda.ListIndex = 0
    cargarControles
    If lnOpcion = 2 Then
        Me.cmdGrabar.Enabled = False
        Me.frmAplicar.Enabled = False
         Me.Caption = "Consulta Tarifario por Operaciones en Otras Agencias"
    End If
End Sub
Private Sub cargarControles()
    Dim i As Integer



'Agencias
    Me.chkAgencias.value = 1 'Cargar por defecto con todos
    Dim rs As New ADODB.Recordset
    Dim objCOMNCredito As New COMNCredito.NCOMBPPR
    Set rs = objCOMNCredito.getCargarAgencias
    If Not (rs.BOF And rs.EOF) Then
       rs.MoveNext
       Me.fgAgencias.Rows = rs.RecordCount  'igualamos numero de filas
       For i = 1 To rs.RecordCount - 1
          Me.fgAgencias.TextMatrix(i, 0) = i
          Me.fgAgencias.TextMatrix(i, 1) = 1
          Me.fgAgencias.TextMatrix(i, 2) = rs!cAgeDescripcion
          Me.fgAgencias.TextMatrix(i, 3) = rs!cAgeCod
          rs.MoveNext
       Next i
    
    End If
    Set rs = Nothing
'Tipos de Ahorros
   
    Me.chkTipoAhorro.value = 1 'Cargar por defecto con todos
    Dim objGen As New COMDConstSistema.DCOMGeneral
    Set rs = objGen.GetConstante("2030", , , "1")
    
    If Not (rs.BOF And rs.EOF) Then
       
       Me.fgTipoAhorro.Rows = rs.RecordCount + 1 'igualamos numero de filas
       For i = 1 To rs.RecordCount
          Me.fgTipoAhorro.TextMatrix(i, 0) = i
          Me.fgTipoAhorro.TextMatrix(i, 1) = 1
          Me.fgTipoAhorro.TextMatrix(i, 2) = rs!cDescripcion
          Me.fgTipoAhorro.TextMatrix(i, 3) = rs!nConsValor
          rs.MoveNext
       Next i
    
    End If
    Set rs = Nothing
'Tarifario
    cargarTarifario
    
End Sub
Private Sub cargarTarifario()
    Dim sCadAgencia As String
    Dim sCadTpoPrograma As String
    Dim rs As New Recordset
    Dim objCapta As New COMNCaptaGenerales.NCOMCaptaGenerales
    
    'se obtienen las Agencias y Tipos de Ahorros Seleccionadas
    sCadAgencia = obtenerCadenaDatos(Me.fgAgencias)
    sCadTpoPrograma = obtenerCadenaDatos(Me.fgTipoAhorro)
    
    Set rs = objCapta.obtenerTarifasOpeEnOtrasAgencias(sCadAgencia, sCadTpoPrograma, Me.cboMoneda.ItemData(Me.cboMoneda.ListIndex))

    If Not (rs.BOF And rs.EOF) Then
       
      Set Me.fgTarifario.Recordset = rs
       
    End If
    Set rs = Nothing
End Sub
Private Sub cmdAplicar_Click()
        
    AplicarDatosComision
    Me.cmdGrabar.SetFocus
End Sub
Private Function obtenerCadenaDatos(ByVal pfgFlex As FlexEdit) As String
    Dim i As Integer
        
    For i = 1 To pfgFlex.Rows - 1
        If pfgFlex.TextMatrix(i, 1) = "." Then
            obtenerCadenaDatos = obtenerCadenaDatos + pfgFlex.TextMatrix(i, 3) + ","
        End If
    Next i
    obtenerCadenaDatos = Mid(obtenerCadenaDatos, 1, Len(obtenerCadenaDatos) - 1)
End Function
Private Function verificarDatoChecked(pfgFlex As FlexEdit) As Boolean
    Dim i As Integer
    verificarDatoChecked = False
    For i = 1 To pfgFlex.Rows - 1
        If pfgFlex.TextMatrix(i, 1) = "." Then
            verificarDatoChecked = True
            Exit For
        End If
    Next i
End Function
Private Sub AplicarDatosComision()
    Dim i As Integer
    Dim rs As New Recordset
    Set rs = Me.fgTarifario.Recordset
    For i = 1 To rs.RecordCount
         If Me.fgTarifario.TextMatrix(i, 1) = "." Then
            
            If Me.txtComision.Text <> "" Then Me.fgTarifario.TextMatrix(i, 4) = Format(Me.txtComision, "#0.00")
            If Me.txtMontoMin.Text <> "" Then Me.fgTarifario.TextMatrix(i, 5) = Format(Me.txtMontoMin, "##,##0.00")
            If Me.txtMontoMax.Text <> "" Then Me.fgTarifario.TextMatrix(i, 6) = Format(Me.txtMontoMax, "##,##0.00")
          End If
    Next i
End Sub

Private Sub txtComision_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMontoMin.SetFocus
    ElseIf KeyAscii = 8 Then 'si es retroceso
        If Len(txtComision.Text) > 0 Then
            txtComision.Text = Mid(txtComision.Text, 1, Len(txtComision.Text) - 1)
            txtComision.SelStart = Len(txtComision.Text)
        End If
    ElseIf InStr("0123456789.", Chr(KeyAscii)) = 0 Or Len(txtComision.Text) = 6 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtMontoMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAplicar.SetFocus
    ElseIf KeyAscii = 8 Then 'si es retroceso
        If Len(txtMontoMax.Text) > 0 Then
            txtMontoMax.Text = Mid(txtMontoMax.Text, 1, Len(txtMontoMax.Text) - 1)
            txtMontoMax.SelStart = Len(txtMontoMax.Text)
        End If
    ElseIf InStr("0123456789.", Chr(KeyAscii)) = 0 Or Len(txtMontoMax.Text) = 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtMontoMin_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMontoMax.SetFocus
    ElseIf KeyAscii = 8 Then 'si es retroceso
        If Len(txtMontoMin.Text) > 0 Then
            txtMontoMin.Text = Mid(txtMontoMin.Text, 1, Len(txtMontoMin.Text) - 1)
            txtMontoMin.SelStart = Len(txtMontoMin.Text)
        End If
    ElseIf InStr("0123456789.", Chr(KeyAscii)) = 0 Or Len(txtMontoMin.Text) = 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub checkeaDatos(ByRef pfgFlex As FlexEdit, pbValor As Integer)
    Dim i As Integer
    
    For i = 1 To pfgFlex.Rows - 1
        pfgFlex.TextMatrix(i, 1) = pbValor
    Next i

End Sub
