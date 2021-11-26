VERSION 5.00
Begin VB.Form frmCapTarifaOperaciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tarifario por Nro. Operaciones Max. de Retiro"
   ClientHeight    =   7410
   ClientLeft      =   1965
   ClientTop       =   2730
   ClientWidth     =   12690
   Icon            =   "frmCapTarifaOperaciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   12690
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
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
      Height          =   3135
      Left            =   0
      TabIndex        =   17
      Top             =   4200
      Width           =   12615
      Begin SICMACT.FlexEdit fgTarifario 
         Height          =   2775
         Left            =   120
         TabIndex        =   18
         Top             =   240
         Width           =   12375
         _extentx        =   21828
         _extenty        =   4895
         cols0           =   11
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-OK-Operacion-Agencia-Tipo Ahorro-Monto MN-Monto ME-Nro Ope-cAgeCod-nTpoPrograma-cOpeCod"
         encabezadosanchos=   "350-350-3000-3000-2000-1100-1100-900-0-0-0"
         font            =   "frmCapTarifaOperaciones.frx":030A
         font            =   "frmCapTarifaOperaciones.frx":0336
         font            =   "frmCapTarifaOperaciones.frx":0362
         font            =   "frmCapTarifaOperaciones.frx":038E
         fontfixed       =   "frmCapTarifaOperaciones.frx":03BA
         backcolorcontrol=   12058623
         backcolorcontrol=   12058623
         backcolorcontrol=   12058623
         lbultimainstancia=   -1  'True
         columnasaeditar =   "X-1-X-X-X-5-6-7-X-X-X"
         listacontroles  =   "0-4-0-0-0-0-0-0-0-0-0"
         encabezadosalineacion=   "C-C-L-L-L-R-R-R-C-C-C"
         formatosedit    =   "0-0-0-0-0-2-2-3-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1  'True
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   345
         rowheight0      =   300
         forecolor       =   -2147483630
         forecolorfixed  =   -2147483630
         cellforecolor   =   -2147483630
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
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
      Height          =   4095
      Left            =   11040
      TabIndex        =   13
      Top             =   0
      Width           =   1575
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Height          =   375
         Left            =   120
         TabIndex        =   16
         Top             =   360
         Width           =   1215
      End
      Begin VB.CommandButton Cancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   120
         TabIndex        =   15
         Top             =   840
         Width           =   1215
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1320
         Width           =   1215
      End
   End
   Begin VB.Frame frmAplicar 
      Appearance      =   0  'Flat
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
      Left            =   0
      TabIndex        =   7
      Top             =   3000
      Width           =   10935
      Begin VB.TextBox txtNroOpeMax 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   3720
         TabIndex        =   19
         Top             =   600
         Width           =   1575
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
         Left            =   5880
         TabIndex        =   10
         Top             =   600
         Width           =   1455
      End
      Begin VB.TextBox txtMontoME 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   1920
         TabIndex        =   9
         Top             =   600
         Width           =   1695
      End
      Begin VB.TextBox txtMontoMN 
         Alignment       =   1  'Right Justify
         Height          =   330
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   1695
      End
      Begin VB.Label Label4 
         Caption         =   "Nro Ope."
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
         TabIndex        =   20
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Monto ME"
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
         TabIndex        =   12
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Monto MN"
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
         TabIndex        =   11
         Top             =   360
         Width           =   1695
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
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
      Height          =   2895
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10935
      Begin VB.CheckBox chkOperacion 
         Caption         =   "Todas las Operaciones"
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
         TabIndex        =   5
         Top             =   360
         Width           =   2295
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
         Left            =   3960
         TabIndex        =   2
         Top             =   360
         Width           =   2295
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
         Left            =   7440
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin SICMACT.FlexEdit fgAgencias 
         Height          =   2175
         Left            =   3960
         TabIndex        =   3
         Top             =   600
         Width           =   3375
         _extentx        =   5953
         _extenty        =   3836
         cols0           =   4
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-OK-Agencia-cAgeCod"
         encabezadosanchos=   "300-350-2300-0"
         font            =   "frmCapTarifaOperaciones.frx":03E8
         font            =   "frmCapTarifaOperaciones.frx":0414
         font            =   "frmCapTarifaOperaciones.frx":0440
         font            =   "frmCapTarifaOperaciones.frx":046C
         fontfixed       =   "frmCapTarifaOperaciones.frx":0498
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
         Height          =   2175
         Left            =   7440
         TabIndex        =   4
         Top             =   600
         Width           =   3375
         _extentx        =   5953
         _extenty        =   3836
         cols0           =   4
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-OK-Tipo Ahorro-cTpoPrograma"
         encabezadosanchos=   "300-350-2300-0"
         font            =   "frmCapTarifaOperaciones.frx":04C6
         font            =   "frmCapTarifaOperaciones.frx":04F2
         font            =   "frmCapTarifaOperaciones.frx":051E
         font            =   "frmCapTarifaOperaciones.frx":054A
         fontfixed       =   "frmCapTarifaOperaciones.frx":0576
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
      Begin SICMACT.FlexEdit fgOperacion 
         Height          =   2175
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   3735
         _extentx        =   6588
         _extenty        =   3836
         cols0           =   4
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-OK-Operacion-cOpeCod"
         encabezadosanchos=   "300-350-3000-0"
         font            =   "frmCapTarifaOperaciones.frx":05A4
         font            =   "frmCapTarifaOperaciones.frx":05D0
         font            =   "frmCapTarifaOperaciones.frx":05FC
         font            =   "frmCapTarifaOperaciones.frx":0628
         fontfixed       =   "frmCapTarifaOperaciones.frx":0654
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
         lbformatocol    =   -1  'True
         lbbuscaduplicadotext=   -1  'True
         colwidth0       =   300
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmCapTarifaOperaciones"
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
Private Sub cmdGrabar_Click()
    
    Me.txtMontoMN.Text = ""
    Me.txtMontoME.Text = ""
    
    If Not verificarDatoChecked(Me.fgTarifario) Then
         MsgBox "Debe Seleccionar al menos un Item del Tarifario", vbInformation, "AVISO!"
         Exit Sub
    End If
    
    Me.MousePointer = vbHourglass
    Dim objMov As New COMNContabilidad.NCOMContFunciones
    Dim objCap As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim sMovNro As String
    
    sMovNro = objMov.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    If MsgBox("Desea Guardar el Tarifario?", vbYesNo, "AVISO!") = vbYes Then
        
        
        Dim sCadOperacion As String
        Dim sCadAgencia As String
        Dim sCadTpoPrograma As String
        
       
        sCadOperacion = obtenerCadenaDatos(Me.fgOperacion)
        sCadAgencia = obtenerCadenaDatos(Me.fgAgencias)
        sCadTpoPrograma = obtenerCadenaDatos(Me.fgTipoAhorro)
        
        If objCap.guardarTarifasOperaciones(Me.fgTarifario.GetRsNew, sCadOperacion, sCadAgencia, sCadTpoPrograma, sMovNro) Then
            If MsgBox("Se ha Guardado el Tarifario con Exito!,Desea Realizar Otra Operacion", vbYesNo, "AVISO!") = vbYes Then
                cargarControles
            Else
                Unload Me
            End If
        Else
            MsgBox "No se han Podido Guardar los Datos, Verifique los Datos", vbInformation, "AVISO!"
        End If
        
    End If
    Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    cargarControles
    If lnOpcion = 2 Then
        Me.cmdGrabar.Enabled = False
        Me.frmAplicar.Enabled = False
        Me.Caption = "Consulta Tarifario por Nro Operaciones Max. de Retiro"
    End If
End Sub
Private Sub cargarControles()
    Dim i As Integer
    Dim rs As New ADODB.Recordset

'Operaciones
    Me.chkOperacion.value = 1 'Cargar por defecto con todos
    Dim objCap As New COMNCaptaGenerales.NCOMCaptaGenerales
    Set rs = objCap.obtenerOperaciones
    If Not (rs.BOF And rs.EOF) Then
       
       Me.fgOperacion.Rows = rs.RecordCount + 1 'igualamos numero de filas
       For i = 1 To rs.RecordCount
          Me.fgOperacion.TextMatrix(i, 0) = i
          Me.fgOperacion.TextMatrix(i, 1) = 1
          Me.fgOperacion.TextMatrix(i, 2) = rs!Operacion
          Me.fgOperacion.TextMatrix(i, 3) = rs!cOpecod
          rs.MoveNext
       Next i
       
    End If
    Set rs = Nothing
    
'Agencias
    Me.chkAgencias.value = 1 'Cargar por defecto con todos
   
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

Private Sub chkOperacion_Click()
    checkeaDatos Me.fgOperacion, Me.chkOperacion.value
    
    Me.txtMontoMN.Text = ""
    Me.txtMontoME.Text = ""
    
    If Not verificarDatoChecked(Me.fgOperacion) Then
            Call LimpiaFlex(Me.fgTarifario)
            Exit Sub
    End If
        
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
Private Sub chkAgencias_Click()
    checkeaDatos Me.fgAgencias, Me.chkAgencias.value
    
    Me.txtMontoMN.Text = ""
    Me.txtMontoME.Text = ""
    
    If Not verificarDatoChecked(Me.fgOperacion) Then
            Call LimpiaFlex(Me.fgTarifario)
            Exit Sub
    End If
        
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
   
    Me.txtMontoMN.Text = ""
    Me.txtMontoME.Text = ""
        
    If Not verificarDatoChecked(Me.fgOperacion) Then
            Call LimpiaFlex(Me.fgTarifario)
            Exit Sub
    End If
        
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
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub fgOperacion_Click()
    If Me.fgOperacion.Col = 1 Then
        
        Me.txtMontoMN.Text = ""
        Me.txtMontoME.Text = ""
        
        
        If Not verificarDatoChecked(Me.fgOperacion) Then
            Call LimpiaFlex(Me.fgTarifario)
            Exit Sub
        End If
        
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
Private Sub fgAgencias_Click()
    
    If Me.fgAgencias.Col = 1 Then
        
        Me.txtMontoMN.Text = ""
        Me.txtMontoME.Text = ""
        
        If Not verificarDatoChecked(Me.fgOperacion) Then
            Call LimpiaFlex(Me.fgTarifario)
            Exit Sub
        End If
        
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
        
        Me.txtMontoMN.Text = ""
        Me.txtMontoME.Text = ""
        
        
        If Not verificarDatoChecked(Me.fgOperacion) Then
            Call LimpiaFlex(Me.fgTarifario)
            Exit Sub
        End If
        
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
Private Sub cargarTarifario()
    
    Dim sCadOperacion As String
    Dim sCadAgencia As String
    Dim sCadTpoPrograma As String
    Dim rs As New Recordset
    Dim objCapta As New COMNCaptaGenerales.NCOMCaptaGenerales
    
    'se obtienen las Agencias y Tipos de Ahorros Seleccionadas
    sCadOperacion = obtenerCadenaDatos(Me.fgOperacion)
    sCadAgencia = obtenerCadenaDatos(Me.fgAgencias)
    sCadTpoPrograma = obtenerCadenaDatos(Me.fgTipoAhorro)
    
    Set rs = objCapta.obtenerTarifasOperaciones(sCadOperacion, sCadAgencia, sCadTpoPrograma)

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
    Dim nFilas As Integer
    Dim rs As New Recordset
    Set rs = Me.fgTarifario.Recordset
        
    If rs.RecordCount > 2048 Then
        nFilas = 2048
    Else
        nFilas = rs.RecordCount
    End If
    
    For i = 1 To nFilas
         If Me.fgTarifario.TextMatrix(i, 1) = "." Then
         
            If Me.txtMontoMN.Text <> "" Then Me.fgTarifario.TextMatrix(i, 5) = Format(Me.txtMontoMN, "##,##0.00")
            If Me.txtMontoME.Text <> "" Then Me.fgTarifario.TextMatrix(i, 6) = Format(Me.txtMontoME, "##,##0.00")
            If Me.txtNroOpeMax.Text <> "" Then Me.fgTarifario.TextMatrix(i, 7) = Me.txtNroOpeMax
            
          End If

    Next i

End Sub

Private Sub txtNroOpeMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdAplicar.SetFocus
    ElseIf KeyAscii = 8 Then 'si es retroceso
        If Len(txtNroOpeMax.Text) > 0 Then
            txtNroOpeMax.Text = Mid(txtNroOpeMax.Text, 1, Len(txtNroOpeMax.Text) - 1)
            txtNroOpeMax.SelStart = Len(txtNroOpeMax.Text)
        End If
    ElseIf InStr("0123456789", Chr(KeyAscii)) = 0 Or Len(txtNroOpeMax.Text) = 6 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtMontoME_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtNroOpeMax.SetFocus
    ElseIf KeyAscii = 8 Then 'si es retroceso
        If Len(txtMontoME.Text) > 0 Then
            txtMontoME.Text = Mid(txtMontoME.Text, 1, Len(txtMontoME.Text) - 1)
            txtMontoME.SelStart = Len(txtMontoME.Text)
        End If
    ElseIf InStr("0123456789.", Chr(KeyAscii)) = 0 Or Len(txtMontoME.Text) = 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub txtMontoMN_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtMontoME.SetFocus
    ElseIf KeyAscii = 8 Then 'si es retroceso
        If Len(txtMontoMN.Text) > 0 Then
            txtMontoMN.Text = Mid(txtMontoMN.Text, 1, Len(txtMontoMN.Text) - 1)
            txtMontoMN.SelStart = Len(txtMontoMN.Text)
        End If
    ElseIf InStr("0123456789.", Chr(KeyAscii)) = 0 Or Len(txtMontoMN.Text) = 8 Then
        KeyAscii = 0
    End If
End Sub
Private Sub checkeaDatos(ByRef pfgFlex As FlexEdit, pbValor As Integer)
    Dim i As Integer
    
    For i = 1 To pfgFlex.Rows - 1
        pfgFlex.TextMatrix(i, 1) = pbValor
    Next i

End Sub

