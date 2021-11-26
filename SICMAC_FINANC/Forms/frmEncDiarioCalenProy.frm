VERSION 5.00
Begin VB.Form frmEncDiarioCalenProy 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calendario de Proyecciones - Encaje"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   Icon            =   "frmEncDiarioCalenProy.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   9690
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboMes 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmEncDiarioCalenProy.frx":030A
      Left            =   2160
      List            =   "frmEncDiarioCalenProy.frx":0332
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   360
      Width           =   2055
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      Height          =   375
      Left            =   7320
      TabIndex        =   6
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   8400
      TabIndex        =   5
      Top             =   4200
      Width           =   1095
   End
   Begin Sicmact.FlexEdit FeCalendario 
      Height          =   3015
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   5318
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "#-Codigo-Fecha-Depósitos a Plazo-Depósitos de Ahorro-Depósitos BCRP-Obligaciones Inmediatas-Efectivo Caja"
      EncabezadosAnchos=   "300-0-0-1800-1800-1800-1800-1800"
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
      ColumnasAEditar =   "X-X-X-3-4-5-6-7"
      ListaControles  =   "0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-R-L-R-C-C-R-C"
      FormatosEdit    =   "0-3-1-2-2-2-2-0"
      TextArray0      =   "#"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   300
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Frame frmPeriodo 
      Caption         =   "Periodo"
      Height          =   855
      Left            =   1920
      TabIndex        =   0
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton cmdSeleccionar 
         Caption         =   "Seleccionar"
         Height          =   375
         Left            =   4800
         TabIndex        =   3
         Top             =   240
         Width           =   1095
      End
      Begin VB.ComboBox CboMoneda 
         Height          =   315
         ItemData        =   "frmEncDiarioCalenProy.frx":039A
         Left            =   3600
         List            =   "frmEncDiarioCalenProy.frx":03A4
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtAnio 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2400
         MaxLength       =   4
         TabIndex        =   1
         Top             =   240
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmEncDiarioCalenProy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**-------------------------------------------------------------------------------------**'
'** Formulario : frmEncDiarioCalenProy                                                  **'
'** Finalidad  : Este formulario permite ingresar valores que serán                     **'
'**            sumados a los montos de las columnas correspondientes                    **'
'**                en las hojas de proyeccion en el formato de                          **'
'**                generacion del encaje legal diario                                   **'
'**Programador: Paolo Hector Sinti Cabrera - PASI                                       **'
'** Fecha/Hora : 20140305 11:50 AM                                                      **'
'**-------------------------------------------------------------------------------------**'
Dim ldFecIni As Date
Dim ldFecFin As Date
Dim lnMoneda As Integer
Dim nEncaje As nEncajeBCR
Dim I As Integer
Dim lnInserModif As Integer

Private Sub cmdAceptar_Click()
    Dim ldFechaRef As Date
    Set nEncaje = New nEncajeBCR
    If FeCalendario.TextMatrix(1, 1) <> "" Then
        ldFechaRef = ldFecIni
        If lnInserModif = 0 Then
            'Codigo para un nuevo registro
            For I = 1 To FeCalendario.Rows - 1
                nEncaje.InsertaCalenProyEncaje ldFechaRef, FeCalendario.TextMatrix(I, 3), FeCalendario.TextMatrix(I, 4), FeCalendario.TextMatrix(I, 5), FeCalendario.TextMatrix(I, 6), FeCalendario.TextMatrix(I, 7), lnMoneda
                ldFechaRef = DateAdd("D", 1, ldFechaRef)
            Next I
            MsgBox "los montos fueron registrados correctamente.", vbInformation, "Aviso!!!"
        Else
            'Codigo para Actualizar un registro
            For I = 1 To FeCalendario.Rows - 1
                nEncaje.ActualizaCalenProyEncaje FeCalendario.TextMatrix(I, 1), FeCalendario.TextMatrix(I, 3), FeCalendario.TextMatrix(I, 4), FeCalendario.TextMatrix(I, 5), FeCalendario.TextMatrix(I, 6), FeCalendario.TextMatrix(I, 7)
            Next I
            MsgBox "los montos fueron actualizados correctamente.", vbInformation, "Aviso!!!"
        End If
        CargarDatos
    Else
        MsgBox "No existen registros para guardar.", vbInformation, "Aviso!!!"
    End If
End Sub

Private Sub cmdCancelar_Click()
    Unload Me
End Sub
Private Sub CargarDatos()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim ldFechaRef As Date
    If CboMoneda.ListIndex = -1 Then
        MsgBox "No se ha seleccionado el tipo de moneda.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If cboMes.ListIndex = -1 Then
        MsgBox "No se ha seleccionado ningun mes.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    ldFecIni = CDate("01/" & Format(cboMes.ListIndex + 1, "00") & "/" & txtAnio.Text)
    ldFecFin = DateAdd("M", 1, ldFecIni) - 1
    lnMoneda = IIf(CboMoneda.ListIndex + 1 = 1, 1, 2)
    
    Set nEncaje = New nEncajeBCR
    Set rs = nEncaje.ObtenerCalenProyEncaje(ldFecIni, ldFecFin, lnMoneda)
    FormateaFlex FeCalendario
    If Not rs.EOF And Not rs.BOF Then
        lnInserModif = 1
        Do While Not rs.EOF
            FeCalendario.AdicionaFila
            I = FeCalendario.row
            FeCalendario.TextMatrix(I, 1) = rs!Codigo
            FeCalendario.TextMatrix(I, 2) = rs!Fecha
            FeCalendario.TextMatrix(I, 3) = Format(rs!DepPla, "#,#0.00")
            FeCalendario.TextMatrix(I, 4) = Format(rs!DepAho, "#,#0.00")
            FeCalendario.TextMatrix(I, 5) = Format(rs!DepBcrp, "#,#0.00")
            FeCalendario.TextMatrix(I, 6) = Format(rs!OblInm, "#,#0.00")
            FeCalendario.TextMatrix(I, 7) = Format(rs!EfecCaja, "#,#0.00")
            rs.MoveNext
        Loop
    Else
        lnInserModif = 0
        I = 0
        ldFechaRef = ldFecIni
        FeCalendario.lbEditarFlex = True
        Do While ldFechaRef <= ldFecFin
            I = I + 1
            FeCalendario.AdicionaFila
            FeCalendario.TextMatrix(I, 1) = I
            FeCalendario.TextMatrix(I, 2) = ldFechaRef
            FeCalendario.TextMatrix(I, 3) = Format(0, "#,#0.00")
            FeCalendario.TextMatrix(I, 4) = Format(0, "#,#0.00")
            FeCalendario.TextMatrix(I, 5) = Format(0, "#,#0.00")
            FeCalendario.TextMatrix(I, 6) = Format(0, "#,#0.00")
            FeCalendario.TextMatrix(I, 7) = Format(0, "#,#0.00")
            ldFechaRef = DateAdd("D", 1, ldFechaRef)
        Loop
    End If
End Sub
Private Sub cmdSeleccionar_Click()
    CargarDatos
End Sub
Private Sub Form_Load()
cboMes.ListIndex = Month(gdFecSis) - 1
txtAnio.Text = Year(gdFecSis)
CboMoneda.ListIndex = 0
End Sub
