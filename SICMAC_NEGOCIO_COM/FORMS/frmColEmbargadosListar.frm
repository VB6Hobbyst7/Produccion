VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Begin VB.Form frmColEmbargadosListar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Listar Embargados"
   ClientHeight    =   6375
   ClientLeft      =   2145
   ClientTop       =   2280
   ClientWidth     =   12420
   Icon            =   "frmColEmbargadosListar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6375
   ScaleWidth      =   12420
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox chkFiltroMul 
      Caption         =   "Filtro Múltiple"
      ForeColor       =   &H00000080&
      Height          =   255
      Left            =   4320
      TabIndex        =   18
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdListar 
      Caption         =   "Listar"
      Height          =   375
      Left            =   11520
      TabIndex        =   15
      Top             =   240
      Width           =   855
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgEmbargos 
      Height          =   4335
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   12135
      _ExtentX        =   21405
      _ExtentY        =   7646
      _Version        =   393216
      Cols            =   12
      _NumberOfBands  =   1
      _Band(0).Cols   =   12
      _Band(0).GridLinesBand=   1
      _Band(0).TextStyleBand=   0
      _Band(0).TextStyleHeader=   0
   End
   Begin VB.Frame Frame3 
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
      Height          =   855
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   7215
      Begin VB.TextBox txtNombreColaborador 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   3795
         Locked          =   -1  'True
         TabIndex        =   21
         Top             =   360
         Width           =   3375
      End
      Begin VB.ComboBox cmbBuscar 
         Height          =   315
         ItemData        =   "frmColEmbargadosListar.frx":030A
         Left            =   600
         List            =   "frmColEmbargadosListar.frx":031A
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   345
         Width           =   1815
      End
      Begin VB.TextBox txtBuscar 
         Height          =   285
         Left            =   2400
         TabIndex        =   7
         Top             =   360
         Width           =   4695
      End
      Begin SICMACT.ActXCodCta AXCodCta 
         Height          =   375
         Left            =   2445
         TabIndex        =   19
         Top             =   360
         Width           =   3615
         _ExtentX        =   6376
         _ExtentY        =   661
         Texto           =   ""
         EnabledCta      =   -1  'True
         EnabledProd     =   -1  'True
         EnabledAge      =   -1  'True
         CMAC            =   "109"
      End
      Begin SICMACT.TxtBuscar txtCodigoCliente 
         Height          =   345
         Left            =   2400
         TabIndex        =   20
         Top             =   360
         Width           =   1380
         _ExtentX        =   2646
         _ExtentY        =   609
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Tipo:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   5520
      Width           =   12135
      Begin VB.CommandButton cmdExpExcel 
         Caption         =   "Exportar"
         Height          =   375
         Left            =   5280
         TabIndex        =   22
         ToolTipText     =   "Permite Exportar el contenido del Grid a Excel."
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   10680
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdSalidaBienes 
         Caption         =   "Salida de Bienes"
         Height          =   375
         Left            =   3480
         TabIndex        =   5
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdModificarEmbargo 
         Caption         =   "Modificar Embargo"
         Height          =   375
         Left            =   1680
         TabIndex        =   4
         Top             =   240
         Width           =   1695
      End
      Begin VB.CommandButton cmdDetalle 
         Caption         =   "Ver Detalle"
         Height          =   375
         Left            =   9240
         TabIndex        =   3
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar Embargo"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1455
      End
   End
   Begin VB.Frame Frame1 
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
      Height          =   855
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   3975
      Begin VB.CheckBox chkFechas 
         Caption         =   "Fechas"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   0
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtFecDel 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   720
         TabIndex        =   11
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txtFecAl 
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd/MM/yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   10250
            SubFormatType   =   3
         EndProperty
         Height          =   255
         Left            =   2520
         TabIndex        =   12
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   510
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2040
         TabIndex        =   13
         Top             =   360
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmColEmbargadosListar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Const consSalidaEmbargo = 1
Const consModificarEmbargo = 2
Const consConsultarEmbargo = 3
Dim lnColumna As Integer
Dim Conecta As DConecta '*** PEAC 20130618
Dim ApExcel As Variant '*** PEAC 20130619
Private WithEvents lrsBuscar As ADODB.Recordset
Attribute lrsBuscar.VB_VarHelpID = -1

Private Sub AXCodCta_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdListar.SetFocus
    End If
End Sub

Private Sub chkFechas_Click()
    If chkFechas.value = vbUnchecked Then
        Me.txtFecDel.Enabled = False
        Me.txtFecAl.Enabled = False
        Me.txtFecDel.Text = "__/__/____"
        Me.txtFecAl.Text = "__/__/____"
    Else
        Me.txtFecDel.Enabled = True
        Me.txtFecAl.Enabled = True
    End If
End Sub

Private Sub chkFiltroMul_Click()
    If chkFiltroMul.value = vbUnchecked Then
        
        Me.cmbBuscar.Enabled = False
        
        AXCodCta.Age = ""
        AXCodCta.Prod = ""
        AXCodCta.Cuenta = ""
        txtBuscar.Text = ""
        txtNombreColaborador.Text = ""
        
        txtCodigoCliente.Visible = False
        txtNombreColaborador.Visible = False
        txtBuscar.Visible = False
        AXCodCta.Visible = False

    Else
        Me.cmbBuscar.Enabled = True
    End If
End Sub

Private Sub cmbBuscar_Click()
    If CInt(Left(cmbBuscar.Text, 2)) = 1 Then 'nro credito
        
        AXCodCta.Visible = True
        AXCodCta.Age = ""
        AXCodCta.Prod = ""
        AXCodCta.Cuenta = ""
        
        txtCodigoCliente.Visible = False
        txtNombreColaborador.Visible = False
        txtBuscar.Visible = False
        
    ElseIf CInt(Left(cmbBuscar.Text, 2)) = 2 Then 'cliente
        
        txtCodigoCliente.Visible = True
        txtNombreColaborador.Visible = True
        txtNombreColaborador.Text = ""
        
        AXCodCta.Visible = False
        txtBuscar.Visible = False
        
    ElseIf CInt(Left(cmbBuscar.Text, 2)) = 3 Then 'nro exp
        
        Me.txtBuscar.Enabled = True
        txtBuscar.Visible = True
        txtBuscar.Text = ""
        
        AXCodCta.Visible = False
        txtCodigoCliente.Visible = False
        txtNombreColaborador.Visible = False
        
    ElseIf CInt(Left(cmbBuscar.Text, 2)) = 4 Then 'nro resolu
        
        Me.txtBuscar.Enabled = True
        txtBuscar.Visible = True
        txtBuscar.Text = ""
        
        AXCodCta.Visible = False
        txtCodigoCliente.Visible = False
        txtNombreColaborador.Visible = False
        
    End If
    

End Sub

Private Sub cmdAgregar_Click()
       If frmColEmbargado.Inicio Then
           CargarEmbargos
       End If
       
End Sub
Private Sub CargarEmbargos()
    Dim oColRec As COMNColocRec.NCOMColRecCredito
    Dim dFecha As Date
    Dim lcFecDel, lcFecAl, lcNumCta, lcCodCli, lcNroExp, lcNroRes As String '*** PEAC 20130619
   
    Set oColRec = New COMNColocRec.NCOMColRecCredito
   
    '*** PEAC 20130619
    'dFecha = DateAdd("d", -1, DateAdd("M", 1, CDate(Me.mskAnio.Text + "/" + Format(Me.cmbMes.ListIndex + 1, "00") + "/" + "01")))
    'Set lrsBuscar = oColRec.obtenerEmbargosListar(dFecha)
    
    lcFecDel = IIf(Me.chkFechas.value = 1, Format(Me.txtFecDel.Text, "yyyymmdd"), "")
    lcFecAl = IIf(Me.chkFechas.value = 1, Format(Me.txtFecAl.Text, "yyyymmdd"), "")
    
    If Me.chkFiltroMul.value = 1 Then
        lcNumCta = IIf(CInt(Left(Me.cmbBuscar.Text, 2)) = 1, Me.AXCodCta.NroCuenta, "")
        lcCodCli = IIf(CInt(Left(Me.cmbBuscar.Text, 2)) = 2, Me.txtCodigoCliente.Text, "")
        lcNroExp = IIf(CInt(Left(Me.cmbBuscar.Text, 2)) = 3, Trim(Me.txtBuscar.Text), "")
        lcNroRes = IIf(CInt(Left(Me.cmbBuscar.Text, 2)) = 4, Trim(Me.txtBuscar.Text), "")
    Else
        lcNumCta = "": lcCodCli = "": lcNroExp = "": lcNroRes = ""
    End If
    
    Set lrsBuscar = oColRec.obtenerEmbargosListar(lcFecDel, lcFecAl, lcNumCta, lcCodCli, lcNroExp, lcNroRes)
    '*** FIN PEAC
   
    If Not (lrsBuscar.EOF And lrsBuscar.BOF) Then
   
       Set fgEmbargos.Recordset = lrsBuscar
       'Me.txtBuscar.Enabled = True
    Else
        MsgBox "No existe datos para mostrar.", vbOKOnly + vbInformation, "Atención"
        Call LimpiaFlex(Me.fgEmbargos)
         'Me.txtBuscar.Enabled = False
    End If
    ConfigFlexCabecera
End Sub

Private Sub cmdDetalle_Click()
    With Me.fgEmbargos
            If .TextMatrix(.Row, 1) <> "" Then
                If frmColEmbargado.Inicio(.TextMatrix(.Row, 1), .TextMatrix(.Row, 3), consConsultarEmbargo) Then
                   CargarEmbargos
                End If
            End If
      End With
End Sub

Private Sub cmdExpExcel_Click()
    Dim i As Integer
    Dim rs As ADODB.Recordset
    Dim lnFila As Integer
    
    If Me.fgEmbargos.Row >= 1 And fgEmbargos.TextMatrix(fgEmbargos.Row, 2) <> "" Then
    
        Set ApExcel = CreateObject("Excel.application")
        '-------------------------------
        'Agrega un nuevo Libro
        ApExcel.Workbooks.Add
        'Poner Titulos
        
        ApExcel.Cells(1, 1) = "CAJA MAYNAS S.A."
        ApExcel.Cells(2, 1) = "BIENES EMBARGADOS"
        ApExcel.Cells(4, 1) = "FECHA"
        ApExcel.Cells(4, 2) = "NUM. CREDITO"
        ApExcel.Cells(4, 3) = "CLIENTE"
        ApExcel.Cells(4, 4) = "NUM. EXPEDIENTE"
        ApExcel.Cells(4, 5) = "NUM. RESOLUCION"
        ApExcel.Cells(4, 6) = "MONEDA"
        ApExcel.Cells(4, 7) = "CAPITAL"
        ApExcel.Cells(4, 8) = "INTERES"
        ApExcel.Cells(4, 9) = "GASTOS"
        ApExcel.Cells(4, 10) = "BIENES"

        ApExcel.Range("A1:J4").Font.Bold = True
        ApExcel.Range("A4:J4").Interior.ColorIndex = 42

        lnFila = 5
        For i = 1 To fgEmbargos.Rows - 1
            ApExcel.Cells(lnFila, 1).NumberFormat = "@"
            ApExcel.Cells(lnFila, 1) = fgEmbargos.TextMatrix(i, 2)
            ApExcel.Cells(lnFila, 2).NumberFormat = "@"
            ApExcel.Cells(lnFila, 2) = fgEmbargos.TextMatrix(i, 3)
            ApExcel.Cells(lnFila, 3) = fgEmbargos.TextMatrix(i, 4)
            ApExcel.Cells(lnFila, 4).NumberFormat = "@"
            ApExcel.Cells(lnFila, 4) = "'" + fgEmbargos.TextMatrix(i, 5)
            ApExcel.Cells(lnFila, 5).NumberFormat = "@"
            ApExcel.Cells(lnFila, 5) = fgEmbargos.TextMatrix(i, 6)
            ApExcel.Cells(lnFila, 6) = fgEmbargos.TextMatrix(i, 7)
            ApExcel.Cells(lnFila, 7).NumberFormat = "#,##0.00"
            ApExcel.Cells(lnFila, 7) = fgEmbargos.TextMatrix(i, 8)
            ApExcel.Cells(lnFila, 8).NumberFormat = "#,##0.00"
            ApExcel.Cells(lnFila, 8) = fgEmbargos.TextMatrix(i, 9)
            ApExcel.Cells(lnFila, 9).NumberFormat = "#,##0.00"
            ApExcel.Cells(lnFila, 9) = fgEmbargos.TextMatrix(i, 10)
            ApExcel.Cells(lnFila, 10) = fgEmbargos.TextMatrix(i, 11)
            
            lnFila = lnFila + 1
        Next i

        ApExcel.Cells.Select
        ApExcel.Cells.EntireColumn.AutoFit

        ApExcel.Cells.Select
        ApExcel.Cells.Font.Size = 8
        ApExcel.Cells.Range("A1").Select

        ApExcel.Cells.Columns("A:A").ColumnWidth = 10
        ApExcel.Cells.Columns("B:B").ColumnWidth = 17
        ApExcel.Cells.Columns("C:C").ColumnWidth = 40
        ApExcel.Cells.Columns("D:D").ColumnWidth = 16
        
                
        '-------------------------------
        ApExcel.Visible = True
        Set ApExcel = Nothing
    
    Else
        MsgBox "No existen datos para mostrar.", vbInformation + vbOKOnly, "SICMACM"
    End If

End Sub

Private Sub cmdListar_Click()
    '*** PEAC 20130619
    If Me.chkFiltroMul.value = 0 And Me.chkFechas.value = 0 Then
        MsgBox "Seleccione por lo menos un filtro.", vbOKOnly + vbInformation, "Atención"
        Call LimpiaFlex(Me.fgEmbargos)
        Exit Sub
    End If

    If Me.chkFiltroMul.value = 1 Then
        If cmbBuscar.Text = "" Then
            MsgBox "Seleccione un tipo de Buscqueda", vbInformation + vbOKOnly, "Atención"
            Call LimpiaFlex(Me.fgEmbargos)
            Exit Sub
        Else
            If CInt(Left(cmbBuscar.Text, 2)) = 1 And Len(Trim(AXCodCta.NroCuenta)) < 18 Then 'nro credito
                
                MsgBox "Ingrese un numero de credito valido, por favor.", vbExclamation + vbOKOnly, "Atención"
                If AXCodCta.Visible = True Then
                    AXCodCta.SetFocus
                End If
                Exit Sub
                
            ElseIf CInt(Left(cmbBuscar.Text, 2)) = 2 And txtCodigoCliente.Text = "" Then 'cliente
                
                MsgBox "Ingrese un cliente, por favor.", vbExclamation + vbOKOnly, "Atención"
                If txtCodigoCliente.Visible = True Then
                    txtCodigoCliente.SetFocus
                End If
                Exit Sub
                                            
            ElseIf CInt(Left(cmbBuscar.Text, 2)) = 3 And txtBuscar.Text = "" Then 'nro exp
                
                MsgBox "Ingrese un número de expediente, por favor.", vbExclamation + vbOKOnly, "Atención"
                
                If Me.txtBuscar.Visible = True Then
                    Me.txtBuscar.SetFocus
                End If
                Exit Sub
                                                        
            ElseIf CInt(Left(cmbBuscar.Text, 2)) = 4 And txtBuscar.Text = "" Then 'nro resolu
                
                MsgBox "Ingrese un número de resolución, por favor.", vbExclamation + vbOKOnly, "Atención"
                If txtBuscar.Visible = True Then
                    txtBuscar.SetFocus
                End If
                Exit Sub
                            
            End If
        End If
    ElseIf Me.chkFechas.value = 1 Then
        If (IsDate(Me.txtFecDel.Text) = False Or IsDate(Me.txtFecAl.Text) = False) Then
            MsgBox "Ingrese un rango de fechas correctas.", vbInformation, "Aviso"
            Exit Sub
        End If
        
        If CDate(Me.txtFecDel.Text) > CDate(Me.txtFecAl.Text) Then
            MsgBox "La fecha inicial no puede ser mayor a la fecha final.", vbInformation, "Aviso"
            Exit Sub
        End If
    End If

    '*** FIN PEAC

    CargarEmbargos
    
End Sub


Private Sub cmdModificarEmbargo_Click()
    With Me.fgEmbargos
    If .TextMatrix(.Row, 1) <> "" Then
        If frmColEmbargado.Inicio(.TextMatrix(.Row, 1), .TextMatrix(.Row, 3), consModificarEmbargo) Then
           CargarEmbargos
        End If
    End If
  End With
End Sub

Private Sub cmdSalidaBienes_Click()
    With Me.fgEmbargos
     If .TextMatrix(.Row, 1) <> "" Then
        If frmColEmbargado.Inicio(.TextMatrix(.Row, 1), .TextMatrix(.Row, 3), consSalidaEmbargo) Then
           CargarEmbargos
        End If
     End If
  End With
End Sub

Private Sub cmdsalir_Click()
 Unload Me
End Sub

Private Sub fgEmbargos_Click()
    ColoreaCelda &HC0FFC0, vbBlack, Me.fgEmbargos.Col
End Sub

Private Sub fgEmbargos_DblClick()
    With Me.fgEmbargos
        If .TextMatrix(.Row, 1) <> "" Then
            If frmColEmbargado.Inicio(.TextMatrix(.Row, 1), .TextMatrix(.Row, 3), consConsultarEmbargo) Then
               CargarEmbargos
            End If
        End If
  End With
End Sub

Private Sub fgEmbargos_KeyPress(KeyAscii As Integer)
  With Me.fgEmbargos
     If frmColEmbargado.Inicio(.TextMatrix(.Row, 1), .TextMatrix(.Row, 3)) Then
        CargarEmbargos
     End If
  End With
End Sub

Private Sub fgEmbargos_LeaveCell()
    ColoreaCelda vbWhite, vbBlack, Me.fgEmbargos.Col
End Sub

Private Sub fgEmbargos_RowColChange()
    ColoreaCelda &HC0FFC0, vbBlack, Me.fgEmbargos.Col
End Sub

Private Sub Form_Load()
    ConfigFlexCabecera
'    Me.mskAnio.Text = Format(gdFecSis, "yyyy")
'    Me.cmbMes.ListIndex = Month(gdFecSis) - 1
    Set lrsBuscar = New Recordset
    Set Conecta = New DConecta
    '*** PEAC 20130619
    Me.txtFecDel.Enabled = False
    Me.txtFecAl.Enabled = False
    Me.txtFecDel.Text = "__/__/____"
    Me.txtFecAl.Text = "__/__/____"
   
    Me.cmbBuscar.Enabled = False

    AXCodCta.Age = ""
    AXCodCta.Prod = ""
    AXCodCta.Cuenta = ""
    txtBuscar.Text = ""
    txtNombreColaborador.Text = ""
    
    txtCodigoCliente.Visible = False
    txtNombreColaborador.Visible = False
    txtBuscar.Visible = False
    AXCodCta.Visible = False
    '*** FIN PEAC
End Sub

Private Sub txtBuscar_Change()
'    If Me.cmbBuscar.ListIndex <> -1 Then
'        BuscarEmbargo
'    End If
End Sub
Private Sub BuscarEmbargo()
        
        If lrsBuscar Is Nothing Then
            MsgBox " No se ha creado el recordset", vbCritical
            Exit Sub
        End If

     ' verifica que el recordset se encuentre abierto
        If Not lrsBuscar.State = adStateOpen Then
                'MsgBox " El recordset no se encuentra abierto", vbCritical
                MsgBox " Debe Listar los Embargos a Buscar", vbInformation
                Exit Sub
        End If
        
        If Me.txtBuscar.Text <> "" Then
           lrsBuscar.Filter = Me.cmbBuscar.Text & " LIKE '" + Me.txtBuscar.Text + "*'"
        ElseIf Me.txtBuscar.Text = "" Then
            lrsBuscar.Filter = ""
        End If

        If Not lrsBuscar.State = adStateOpen Then
         lrsBuscar.Requery
        End If
       
        Set fgEmbargos.Recordset = lrsBuscar
        ConfigFlexCabecera
     
End Sub
Private Sub lsrBuscar_Error( _
         ByVal ErrorNumber As Long, _
         Description As String, _
         ByVal Scode As Long, _
         ByVal Source As String, _
         ByVal HelpFile As String, _
         ByVal HelpContext As Long, _
         fCancelDisplay As Boolean)
        

   ' Mostramos el error
   MsgBox " Descripción del Error :" & Description, vbCritical
 End Sub
 
  Private Sub lrsBuscar_WillChangeRecord( _
     ByVal adReason As ADODB.EventReasonEnum, _
     ByVal cRecords As Long, _
     adStatus As ADODB.EventStatusEnum, _
     ByVal pRecordset As ADODB.Recordset)

   'Aquí se coloca el código de validación
   'Se llama a este evento cuando ocurre la siguiente acción
       Dim bCancel As Boolean
    
       Select Case adReason
       Case adRsnAddNew
       Case adRsnClose
       Case adRsnDelete
       Case adRsnFirstChange
       Case adRsnMove
       Case adRsnRequery
       Case adRsnResynch
       Case adRsnUndoAddNew
       Case adRsnUndoDelete
       Case adRsnUndoUpdate
       Case adRsnUpdate
       End Select
    
       If bCancel Then adStatus = adStatusCancel
 End Sub
Private Sub ConfigFlexCabecera()
    
    With Me.fgEmbargos
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = "nMovNro"
        .TextMatrix(0, 2) = "Fecha"
        .TextMatrix(0, 3) = "Nro_Credito"
        .TextMatrix(0, 4) = "Cliente"
        .TextMatrix(0, 5) = "Nro_Expediente"
        .TextMatrix(0, 6) = "Nro_Resolucion"
        .TextMatrix(0, 7) = "Moneda"
        .TextMatrix(0, 8) = "Capital"
        .TextMatrix(0, 9) = "Interes"
        .TextMatrix(0, 10) = "Gastos"
        .TextMatrix(0, 11) = "Bienes"
    
         'Ancho de las Columnas
        .ColWidth(0) = 300
        .ColWidth(1) = 0 'nMovNro
        .ColWidth(2) = 1000 'Fecha
        .ColWidth(3) = 1800 'Nro_Credito
        .ColWidth(4) = 3300 'Cliente
        .ColWidth(5) = 1350 'Nro_Expediente
        .ColWidth(6) = 1350 'Nro_Resolucion
        .ColWidth(7) = 1000 'Moneda
        .ColWidth(8) = 1200 'Capital
        .ColWidth(9) = 1200 'Interes
        .ColWidth(10) = 1000 'Gastos
        .ColWidth(11) = 700 'Bienes_Embargados
    
        'Alineacion del contenido en las Celdas
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignCenterCenter
        .ColAlignment(4) = flexAlignLeftCenter
        .ColAlignment(5) = flexAlignLeftCenter
        .ColAlignment(6) = flexAlignLeftCenter
         .ColAlignment(7) = flexAlignLeftCenter
        .ColAlignment(8) = flexAlignRightCenter
        .ColAlignment(9) = flexAlignRightCenter
        .ColAlignment(10) = flexAlignRightCenter
       .ColAlignment(11) = flexAlignCenterCenter
    
    End With
End Sub
Private Sub ColoreaCelda(ByVal colorCelda As OLE_COLOR, ByVal colorFuente As OLE_COLOR, ByVal nCol As Integer)
    lnColumna = nCol
    With Me.fgEmbargos
        Dim i As Integer
        For i = 1 To .Cols - 1
            .Col = i
            .CellBackColor = colorCelda
            .CellForeColor = colorFuente
        Next i
        .Col = lnColumna
    End With
    
End Sub

Private Sub txtBuscar_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyAscii = 13 Then
'        Me.cmdListar.SetFocus
'    End If
End Sub

Private Sub txtBuscar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdListar.SetFocus
    End If
End Sub

Private Sub txtCodigoCliente_EmiteDatos()
    Call CargarNombreCliente
End Sub

'*** PEAC 20130618
Private Sub CargarNombreCliente()
    Dim RSCli As New ADODB.Recordset
    Dim lsSQL As String
    lsSQL = "exec stp_sel_BuscaClienteEmbargos '" & Trim(txtCodigoCliente.Text) & "' "

    Conecta.AbreConexion
    Set RSCli = Conecta.CargaRecordSet(lsSQL)
    Conecta.CierraConexion

    If (RSCli.EOF And RSCli.BOF) Then
        MsgBox "Código de Cliente no existe.", vbInformation + vbOKOnly, "Atención"
        Me.txtNombreColaborador.Text = ""
        'Exit Sub
    Else
        txtNombreColaborador.Text = RSCli!cPersNombre
    End If
    RSCli.Close
    Set RSCli = Nothing
End Sub

Private Sub txtCodigoCliente_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdListar.SetFocus
    End If
End Sub

Private Sub txtFecAl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.cmdListar.SetFocus
    End If
End Sub

Private Sub txtFecDel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Me.txtFecAl.SetFocus
    End If
End Sub
