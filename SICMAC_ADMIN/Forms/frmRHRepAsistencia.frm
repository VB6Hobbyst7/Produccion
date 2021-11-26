VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{DFDE2506-090D-11D5-BEF8-C11EAA34970C}#2.0#0"; "Spinner.ocx"
Begin VB.Form frmRHReporteAsistencia 
   Caption         =   " "
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11985
   Icon            =   "frmRHRepAsistencia.frx":0000
   LinkTopic       =   "Form2"
   ScaleHeight     =   7200
   ScaleWidth      =   11985
   StartUpPosition =   1  'CenterOwner
   Begin MSComCtl2.Animation logo 
      Height          =   375
      Left            =   9600
      TabIndex        =   22
      Top             =   1080
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   661
      _Version        =   393216
      FullWidth       =   57
      FullHeight      =   25
   End
   Begin Spinner.uSpinner uSMin 
      Height          =   255
      Left            =   5040
      TabIndex        =   20
      Top             =   1320
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
   End
   Begin VB.Frame FMarca 
      Caption         =   "No Marcaron"
      Height          =   1095
      Left            =   3720
      TabIndex        =   15
      Top             =   1200
      Visible         =   0   'False
      Width           =   2895
      Begin VB.ComboBox cmbturno 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   240
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.ComboBox cmbingreso 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   720
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Label lblturno 
         AutoSize        =   -1  'True
         Caption         =   "Turno"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Visible         =   0   'False
         Width           =   420
      End
      Begin VB.Label lblingreso 
         AutoSize        =   -1  'True
         Caption         =   "Ing/Sal"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   840
         Visible         =   0   'False
         Width           =   525
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Mostrar Por"
      Height          =   1455
      Left            =   240
      TabIndex        =   12
      Top             =   1200
      Width           =   3255
      Begin VB.OptionButton optvertodo 
         Caption         =   "Ver Todo"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   1815
      End
      Begin VB.OptionButton optMarca 
         Caption         =   "Marcaciones"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   600
         Width           =   2415
      End
      Begin VB.OptionButton optminutos 
         Caption         =   "Minutos Acumulados"
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   240
         Value           =   -1  'True
         Width           =   1935
      End
   End
   Begin VB.CommandButton cmdimprimir 
      Caption         =   "Imprimir"
      Height          =   375
      Left            =   5160
      TabIndex        =   11
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   9480
      TabIndex        =   10
      Top             =   2400
      Width           =   1335
   End
   Begin VB.CommandButton cmdver 
      Caption         =   "Ver"
      Height          =   375
      Left            =   3720
      TabIndex        =   8
      Top             =   2400
      Width           =   1335
   End
   Begin MSComCtl2.DTPicker DTPicker2 
      Height          =   285
      Left            =   4320
      TabIndex        =   4
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      Format          =   57671681
      CurrentDate     =   38286
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   285
      Left            =   1440
      TabIndex        =   3
      Top             =   720
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   503
      _Version        =   393216
      Format          =   57671681
      CurrentDate     =   38286
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHListado 
      Height          =   4215
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7435
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   16777215
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.TextBox txtdescripcion 
      Enabled         =   0   'False
      Height          =   285
      Left            =   2760
      TabIndex        =   1
      Top             =   240
      Width           =   5415
   End
   Begin Sicmact.TxtBuscar txtAgencia 
      Height          =   285
      Left            =   1440
      TabIndex        =   9
      Top             =   240
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   503
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      sTitulo         =   ""
   End
   Begin VB.Label lblMinutos 
      AutoSize        =   -1  'True
      Caption         =   "Min Acumulados"
      Height          =   195
      Left            =   3720
      TabIndex        =   7
      Top             =   1320
      Width           =   1170
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Fecha  Final"
      Height          =   195
      Left            =   3360
      TabIndex        =   6
      Top             =   720
      Width           =   870
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Fecha Inicial"
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   720
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Cod Agencia"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   915
   End
End
Attribute VB_Name = "frmRHReporteAsistencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim oArea As DActualizaDatosArea
Dim oAsisRep As DActualizaDatosHorarios
Dim rs As ADODB.Recordset

Private Sub cmdImprimir_Click()
Dim clsRep As DRHReportes
Dim lsCad As String
Dim oPrevio As Previo.clsPrevio
Dim lsCadena As String

Set oPrevio = New Previo.clsPrevio
Set clsRep = New DRHReportes
'gsNomAge, gsEmpresa, gdFecSis



If txtAgencia.Text = "" Then Exit Sub

If optminutos.value = True Then
    lsCad = clsRep.GetReporteRHMinAcumulados(txtdescripcion.Text, gsEmpresa, gdFecSis, txtAgencia.Text, Format(DTPicker1, "mm/dd/yyyy"), Format(DTPicker2, "mm/dd/yyyy"), uSMin.Valor)

    oPrevio.Show lsCad, Me.Caption, True, 66
    Set oPrevio = Nothing
End If

End Sub


Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub cmdver_Click()
Dim STurno As String
Dim SIngreso As String
MSHListado.Clear
If txtAgencia.Text = "" Then
    MsgBox "Debe Seleccionar una Agencia", vbInformation, "Seleccione una Agencia"
    Exit Sub
End If

If DTPicker1.value > DTPicker2.value Then
   MsgBox "la Fecha Fin no puede ser menor que la fecha minima", vbInformation, "Fceha Final es Menor que fecha Inicial.. Corrija"
   Exit Sub
End If
If optminutos.value = True Then
          
             STurno = "S"
            SIngreso = "S"
            
            MSHListado.ColWidth(0) = 0
            MSHListado.ColWidth(1) = 3500
            MSHListado.ColWidth(2) = 2000
            MSHListado.ColWidth(3) = 2000
            
            Set rs = oAsisRep.GetListadofaltas(Format(DTPicker1, "mm/dd/yyyy"), Format(DTPicker2, "mm/dd/yyyy"), txtAgencia.Text, STurno, SIngreso, uSMin.Valor)
            If rs.EOF Then
                    MsgBox "No existen registros que mostrar", vbInformation, "No existen registros que mostrar"
                    Exit Sub
            End If
    Set MSHListado.DataSource = rs
        
End If

    If optMarca.value = True Then
        STurno = Right(cmbturno.Text, 1)
        SIngreso = Right(cmbingreso.Text, 1)
        Set rs = oAsisRep.GetListadofaltas(Format(DTPicker1, "mm/dd/yyyy"), Format(DTPicker2, "mm/dd/yyyy"), txtAgencia.Text, STurno, SIngreso, uSMin.Valor)
            If rs.EOF Then
                    MsgBox "No existen registros que mostrar", vbInformation, "No existen registros que mostrar"
                    Exit Sub
            End If
        
        Set MSHListado.DataSource = rs
        
        MSHListado.ColWidth(0) = 0
        MSHListado.ColWidth(1) = 3000
        MSHListado.ColWidth(2) = 1000
        MSHListado.ColWidth(3) = 800
        MSHListado.ColWidth(4) = 1000
        MSHListado.ColWidth(5) = 1000
        MSHListado.ColWidth(6) = 1000
        MSHListado.ColWidth(7) = 1000
        
    End If
    If optvertodo.value = True Then
        STurno = "T"
        SIngreso = "S"
        Set rs = oAsisRep.GetListadofaltas(Format(DTPicker1, "mm/dd/yyyy"), Format(DTPicker2, "mm/dd/yyyy"), txtAgencia.Text, STurno, SIngreso, uSMin.Valor)
        If rs.EOF Then
            MsgBox "No existen registros que mostrar", vbInformation, "No existen registros que mostrar"
            Exit Sub
        End If
        
        If rs.EOF Then
                    MsgBox "No existen registros que mostrar", vbInformation, "No existen registros que mostrar"
                    Exit Sub
        End If
        
        Set MSHListado.DataSource = rs
        
        MSHListado.ColWidth(0) = 0
        MSHListado.ColWidth(1) = 2500
        MSHListado.ColWidth(2) = 1000
        MSHListado.ColWidth(3) = 800
        MSHListado.ColWidth(4) = 950
        MSHListado.ColWidth(5) = 950
        MSHListado.ColWidth(6) = 950
        MSHListado.ColWidth(7) = 950
        
        MSHListado.ColWidth(8) = 800
        MSHListado.ColWidth(9) = 800
        MSHListado.ColWidth(10) = 800
        MSHListado.ColWidth(11) = 800
        
        
        

    End If


End Sub



Private Sub Form_Load()
Dim oCon As DConstantes
Set rs = New ADODB.Recordset
Set oCon = New DConstantes
Me.Width = 12100
Set oArea = New DActualizaDatosArea
Me.txtAgencia.rs = oArea.GetAgencias
Set rs = New ADODB.Recordset
Set oAsisRep = New DActualizaDatosHorarios
MSHListado.ColWidth(0) = 1300
MSHListado.ColWidth(1) = 2800
MSHListado.ColWidth(2) = 1000
MSHListado.Cols = 3
MSHListado.TextMatrix(0, 0) = "Cod Persona"
MSHListado.TextMatrix(0, 1) = "Descripcion"
MSHListado.TextMatrix(0, 2) = "Min Acumu."

DTPicker1.value = Date
DTPicker2.value = Date
uSMin.Valor = 20



Set rs = oCon.GetConstante(gRHEmpleadoTurno)
CargaCombo rs, cmbturno
cmbturno.AddItem "[Todos]                                       T"
cmbturno.ListIndex = cmbturno.ListCount - 1


cmbingreso.AddItem "Ingreso                                     I"
cmbingreso.AddItem "Salida                                      S"
cmbingreso.AddItem "[Todos]                                     T"
cmbingreso.ListIndex = cmbingreso.ListCount - 1

 logo.AutoPlay = True
 logo.Open App.path & "\videos\LogoA.avi"

End Sub






Private Sub MSHListado_DblClick()

Dim lsCadena As String
Dim lsCad As String
Dim lsCodEmp As String
Dim oRepEvento As NRHReportes
Dim oPrevio As Previo.clsPrevio
Set oPrevio = New Previo.clsPrevio
Set oRepEvento = New NRHReportes

Set oRepEvento = New NRHReportes
lsCodEmp = ""


If optminutos.value = True Then
    
    lsCodEmp = MSHListado.Text
    If lsCodEmp = "" Then Exit Sub
    
    

    lsCad = oRepEvento.GetTardanzasEmpleadosDetalle(CDate(Me.DTPicker1.value), CDate(Me.DTPicker2.value), lsCodEmp, txtdescripcion.Text, gsEmpresa, gdFecSis)

    If lsCad <> "" Then
       If lsCadena <> "" Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
          lsCadena = lsCadena & lsCad
       
    End If
    
    
    
    If lsCadena <> "" Then
            oPrevio.Show lsCadena, "Reportes de RRHH", 1, 66, gImpresora
    End If

End If



If optMarca.value = True Or optvertodo.value = True Then
    
   
    lsCodEmp = MSHListado.Text
    If lsCodEmp = "" Then Exit Sub
  
    lsCad = oRepEvento.GetTardanzasEmpleadosDetalle(CDate(Me.DTPicker1.value), CDate(Me.DTPicker2.value), lsCodEmp, txtdescripcion.Text, gsEmpresa, gdFecSis)

    If lsCad <> "" Then
       If lsCadena <> "" Then lsCadena = lsCadena & oImpresora.gPrnSaltoPagina
          lsCadena = lsCadena & lsCad
       
    End If
    
    
    
    If lsCadena <> "" Then
            oPrevio.Show lsCadena, "Reportes de RRHH", 1, 66, gImpresora
    End If

End If






End Sub

Private Sub optMarca_Click()

If optMarca.value = True Then
    lblturno.Visible = True
    cmbturno.Visible = True
    lblingreso.Visible = True
    cmbingreso.Visible = True
    FMarca.Visible = True
    lblMinutos.Visible = False
    uSMin.Visible = False
    
    Else
    lblturno.Visible = False
    cmbturno.Visible = False
    lblingreso.Visible = False
    cmbingreso.Visible = False
    FMarca.Visible = False
    lblMinutos.Visible = True
    uSMin.Visible = True
    


End If

End Sub

Private Sub optminutos_Click()


If optminutos.value = True Then
    lblturno.Visible = False
    cmbturno.Visible = False
    lblingreso.Visible = False
    cmbingreso.Visible = False
    FMarca.Visible = False
    lblMinutos.Visible = True
    uSMin.Visible = True
    
Else
    lblturno.Visible = True
    cmbturno.Visible = True
    lblingreso.Visible = True
    cmbingreso.Visible = True
    FMarca.Visible = True
    lblMinutos.Visible = False
    uSMin.Visible = False
    


End If

End Sub

Private Sub optvertodo_Click()
If optvertodo.value = True Then
    lblturno.Visible = False
    cmbturno.Visible = False
    lblingreso.Visible = False
    cmbingreso.Visible = False
    FMarca.Visible = False
    lblMinutos.Visible = False
    uSMin.Visible = False
End If



End Sub

Private Sub txtAgencia_EmiteDatos()
txtdescripcion = txtAgencia.psDescripcion
 If txtAgencia.Text <> gsCodAge Then
        If gsCodArea = "022" Or gsCodArea = "044" Then
           Else
               'mientras
                If gsCodUser = "JACM" And (txtAgencia.Text = "13" Or txtAgencia.Text = "06" Or txtAgencia.Text = "03") Then
                Else
                    MsgBox "Solo esta permitido ver asistencia de la agencia asignada", vbInformation, "no esta permitido"
                    Me.txtdescripcion.Text = ""
                    txtAgencia.Text = ""
                    Exit Sub
                End If
        End If
  End If

End Sub

Sub formato(psNumero As Integer)
Select Case psNumero
    Case 0
        MSHListado.ColWidth(0) = 1300
        MSHListado.ColWidth(1) = 2800
        MSHListado.ColWidth(2) = 1000
        MSHListado.Cols = 3
        MSHListado.TextMatrix(0, 0) = "Cod Persona"
        MSHListado.TextMatrix(0, 1) = "Descripcion"
        MSHListado.TextMatrix(0, 2) = "Min Acumu."
     Case 1
        MSHListado.ColWidth(0) = 1300
        MSHListado.ColWidth(1) = 2800
        MSHListado.ColWidth(2) = 1000
        MSHListado.Cols = 3
        MSHListado.TextMatrix(0, 0) = "Cod Persona"
        MSHListado.TextMatrix(0, 1) = "Descripcion"
        MSHListado.TextMatrix(0, 2) = "Min Acumu."

     
     
        
        
End Select
End Sub
