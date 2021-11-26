VERSION 5.00
Begin VB.Form frmBuscaPersonaNegativaServiciosLista 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buscar Persona Pago Servicios"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9255
   Icon            =   "frmBuscaPersonaNegativaServiciosLista.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   9255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtcodcli 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      MaxLength       =   30
      TabIndex        =   17
      Tag             =   "3"
      Top             =   480
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox txtNomPer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1920
      TabIndex        =   13
      Tag             =   "1"
      Top             =   480
      Width           =   3990
   End
   Begin VB.TextBox txtDocPer 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1950
      MaxLength       =   15
      TabIndex        =   12
      Tag             =   "3"
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.CommandButton cmdClose 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
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
      Left            =   7920
      TabIndex        =   3
      Top             =   480
      Width           =   1230
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
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
      Left            =   6600
      TabIndex        =   2
      Top             =   480
      Width           =   1230
   End
   Begin VB.Frame frabusca 
      Caption         =   "Buscar por ...."
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
      Height          =   1410
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1800
      Begin VB.OptionButton optOpcion 
         Caption         =   "Cod Cliente"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   960
         Width           =   1635
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "Nº Docu&mento"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   120
         TabIndex        =   10
         Top             =   600
         Width           =   1635
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "A&pellido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   1
         Top             =   360
         Value           =   -1  'True
         Width           =   1200
      End
   End
   Begin SICMACT.FlexEdit FEPagoServicio 
      Height          =   2925
      Left            =   1920
      TabIndex        =   5
      Top             =   2520
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   5159
      Cols0           =   7
      HighLight       =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      EncabezadosNombres=   "Nº-OK-Concepto-nDeuNro-cCodServicio-Importe-Periodo"
      EncabezadosAnchos=   "400-400-3500-0-0-1200-1200"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-1-X-X-X-X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-4-0-0-0-0-0"
      BackColorControl=   65535
      BackColorControl=   65535
      BackColorControl=   65535
      EncabezadosAlineacion=   "C-C-C-C-C-R-C"
      FormatosEdit    =   "0-0-0-0-0-2-2"
      AvanceCeldas    =   1
      TextArray0      =   "Nº"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483635
   End
   Begin SICMACT.FlexEdit FECliente 
      Height          =   1365
      Left            =   1920
      TabIndex        =   15
      Top             =   960
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   2408
      HighLight       =   1
      AllowUserResizing=   1
      RowSizingMode   =   1
      EncabezadosNombres=   "Nº-Nombre"
      EncabezadosAnchos=   "400-5000"
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
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X"
      TextStyleFixed  =   4
      ListaControles  =   "0-0"
      BackColorControl=   65535
      BackColorControl=   65535
      BackColorControl=   65535
      EncabezadosAlineacion=   "L-L"
      FormatosEdit    =   "0-0"
      AvanceCeldas    =   1
      TextArray0      =   "Nº"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   405
      RowHeight0      =   300
      ForeColorFixed  =   -2147483635
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Para la Búsqueda por apellido, ingrese el Apellido seguido de un espacio en blanco y luego el Nombre"
      Height          =   195
      Left            =   1920
      TabIndex        =   11
      Top             =   0
      Width           =   7200
   End
   Begin VB.Label LblDoc 
      Height          =   195
      Left            =   3975
      TabIndex        =   14
      Top             =   555
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "A Pagar :"
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
      Height          =   195
      Left            =   240
      TabIndex        =   9
      Top             =   2760
      Width           =   810
   End
   Begin VB.Label LblTotPag 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0.00"
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
      Height          =   300
      Left            =   120
      TabIndex        =   8
      Top             =   3000
      Width           =   1245
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Total Registros:"
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
      Height          =   195
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   1365
   End
   Begin VB.Label LblTotCredPag 
      Alignment       =   1  'Right Justify
      BackColor       =   &H8000000E&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "0"
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
      Height          =   300
      Left            =   600
      TabIndex        =   6
      Top             =   2400
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ingrese Dato a Buscar :"
      Height          =   195
      Left            =   1920
      TabIndex        =   4
      Top             =   240
      Width           =   1680
   End
End
Attribute VB_Name = "frmBuscaPersonaNegativaServiciosLista"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Persona As COMDPersona.UCOMPersona
Dim R As ADODB.Recordset
Dim bInstit As String
Private sNumdeuSist() As String

Public Function Inicio(Optional ByVal pbInstitucion As String = "") As Variant 'COMDPersona.UCOMPersona
   Dim xValor As Integer
   bInstit = pbInstitucion
   ReDim sNumdeuSist(0)
   Me.Show 1
   Inicio = sNumdeuSist
End Function

Private Sub FECliente_RowColChange()
  Dim ClsPersona As COMDPersona.DCOMPersonas
  Set ClsPersona = New COMDPersona.DCOMPersonas
  If FECliente.TextMatrix(1, 1) <> "" Then
     If txtNomPer.Visible Then
           txtNomPer.Text = FECliente.TextMatrix(FECliente.Row, 1)
           Set R = ClsPersona.BuscaClienteServicios(txtNomPer.Text, bInstit, BusquedaNombre)
           LimpiaFlex FEPagoServicio
      
          If Not R.EOF And Not R.BOF Then
            For i = 0 To R.RecordCount - 1
                FEPagoServicio.AdicionaFila , , True
                FEPagoServicio.TextMatrix(i + 1, 2) = Trim(R!cConcepto)
                FEPagoServicio.TextMatrix(i + 1, 3) = Trim(R!nDeuNro)
                FEPagoServicio.TextMatrix(i + 1, 4) = Trim(R!cCodServicio)
                FEPagoServicio.TextMatrix(i + 1, 5) = Format(R!nImporteCuota, "#0.00")
                FEPagoServicio.TextMatrix(i + 1, 6) = R!cPeriodo
                R.MoveNext
            Next i
        End If
        FEPagoServicio.lbEditarFlex = True
       
       ElseIf txtDocPer.Visible Then
               'txtDocPer.Text = FECliente.TextMatrix(1, 1)
                  Set R = ClsPersona.BuscaClienteServicios(txtDocPer.Text, bInstit, BusquedaDocumento)
                  LimpiaFlex FEPagoServicio
      
                  If Not R.EOF And Not R.BOF Then
                    For i = 0 To R.RecordCount - 1
                        FEPagoServicio.AdicionaFila , , True
                        FEPagoServicio.TextMatrix(i + 1, 2) = Trim(R!cConcepto)
                        FEPagoServicio.TextMatrix(i + 1, 3) = Trim(R!nDeuNro)
                        FEPagoServicio.TextMatrix(i + 1, 4) = Trim(R!cCodServicio)
                        FEPagoServicio.TextMatrix(i + 1, 5) = Format(R!nImporteCuota, "#0.00")
                        FEPagoServicio.TextMatrix(i + 1, 6) = R!cPeriodo
                        R.MoveNext
                    Next i
                  End If
                    FEPagoServicio.lbEditarFlex = True
        Else
                  Set R = ClsPersona.BuscaClienteServicios(Me.txtcodcli.Text, bInstit, 4)
                  LimpiaFlex FEPagoServicio
      
                  If Not R.EOF And Not R.BOF Then
                    For i = 0 To R.RecordCount - 1
                        FEPagoServicio.AdicionaFila , , True
                        FEPagoServicio.TextMatrix(i + 1, 2) = Trim(R!cConcepto)
                        FEPagoServicio.TextMatrix(i + 1, 3) = Trim(R!nDeuNro)
                        FEPagoServicio.TextMatrix(i + 1, 4) = Trim(R!cCodServicio)
                        FEPagoServicio.TextMatrix(i + 1, 5) = Format(R!nImporteCuota, "#0.00")
                        FEPagoServicio.TextMatrix(i + 1, 6) = R!cPeriodo
                        R.MoveNext
                    Next i
                    FEPagoServicio.lbEditarFlex = True
            End If
     End If
'    Set R = Nothing
 End If
End Sub

Private Sub Form_Load()
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

'20110601
Private Sub cmdAceptar_Click()
    If FEPagoServicio.Rows >= 2 And (FEPagoServicio.TextMatrix(0, 1) <> "") And (Me.txtDocPer.Text <> "" Or Me.txtNomPer.Text <> "" Or Me.txtcodcli.Text <> "") Then
        If Not R.EOF And Not R.BOF Then
            Set FEPagoServicio.Recordset = R
        End If
        FEPagoServicio.lbEditarFlex = True
        FEPagoServicio.Enabled = True
        If LblTotCredPag.Caption > 0 Then
            ReDim Preserve sNumdeuSist(LblTotCredPag.Caption)
        End If
    Else
        limpia_array
    End If
   Unload Me
End Sub
Sub limpia_array()
ReDim sNumdeuSist(0)
sNumdeuSist(0) = ""
End Sub
Private Sub cmdClose_Click()
    Set R = Nothing
    Set Persona = Nothing
    limpia_array
    Unload Me
End Sub

Private Sub optOpcion_Click(Index As Integer)
    Select Case Index
        Case 0 'Busqueda por Nombre
            txtNomPer.Text = ""
            txtNomPer.Visible = True
            txtNomPer.SetFocus
            txtDocPer.Text = ""
            txtDocPer.Visible = False
            lblDoc.Visible = False
            txtcodcli.Text = ""
            txtcodcli.Visible = False
        Case 2 'Busqueda por Documento
            txtDocPer.Text = ""
            txtDocPer.Visible = True
            lblDoc.Visible = True
            txtDocPer.SetFocus
            txtNomPer.Text = ""
            txtNomPer.Visible = False
            txtcodcli.Text = ""
            txtcodcli.Visible = False
        Case 1 'Busqueda por Documento
            txtcodcli.Text = ""
            txtcodcli.Visible = True
            txtDocPer.Text = ""
            txtDocPer.Visible = False
            lblDoc.Visible = False
            txtNomPer.Text = ""
            txtNomPer.Visible = False
            txtcodcli.SetFocus
    End Select
    LimpiaFlex FECliente
End Sub

Private Sub txtcodcli_GotFocus()
    fEnfoque txtcodcli
End Sub

Private Sub txtcodcli_KeyPress(KeyAscii As Integer)
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim R1 As ADODB.Recordset

    If Left(txtcodcli.Text, 1) = "%" Then
        txtcodcli.Text = ""
    End If

   If KeyAscii = 13 Then
      If Len(Trim(txtcodcli.Text)) = 0 Then
        MsgBox "Falta Ingresar el Codigo de la Persona", vbInformation, "Aviso"
        Exit Sub
      End If

      Screen.MousePointer = 11
      Set ClsPersona = New COMDPersona.DCOMPersonas
      If bInstit Then
        Set R1 = ClsPersona.BuscaClienteServiciosNombre(txtcodcli.Text, bInstit, 4)
      Else
         Set R1 = ClsPersona.BuscaClienteServiciosNombre(txtcodcli.Text, bInstit, 4)
      End If
           
      LimpiaFlex FECliente
      If Not R1.EOF And Not R1.BOF Then
        For i = 0 To R1.RecordCount - 1
            FECliente.AdicionaFila
            FECliente.TextMatrix(i + 1, 1) = R1!valor
            R1.MoveNext
        Next i
    End If

      Screen.MousePointer = 0
       
      If R1.RecordCount = 0 Then
        MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
        txtcodcli.SetFocus
        cmdAceptar.Default = False
      Else
        FECliente.lbEditarFlex = True
        Screen.MousePointer = 0
        FECliente.SetFocus
      End If
   End If
End Sub

Private Sub txtDocPer_GotFocus()
    fEnfoque txtDocPer
End Sub

Private Sub txtDocPer_KeyPress(KeyAscii As Integer)
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim R1 As ADODB.Recordset

    If Left(txtDocPer.Text, 1) = "%" Then
        txtDocPer.Text = ""
    End If

   If KeyAscii = 13 Then
      If Len(Trim(txtDocPer.Text)) = 0 Then
        MsgBox "Falta Ingresar el Nombre de la Persona", vbInformation, "Aviso"
        Exit Sub
      End If

      Screen.MousePointer = 11
      Set ClsPersona = New COMDPersona.DCOMPersonas
      If bInstit Then
        Set R1 = ClsPersona.BuscaClienteServiciosNombre(txtDocPer.Text, bInstit, BusquedaDocumento)
      Else
         Set R1 = ClsPersona.BuscaClienteServiciosNombre(txtDocPer.Text, bInstit, BusquedaDocumento)
      End If
           
      LimpiaFlex FECliente
      If Not R1.EOF And Not R1.BOF Then
        For i = 0 To R1.RecordCount - 1
            FECliente.AdicionaFila
            FECliente.TextMatrix(i + 1, 1) = R1!valor
            R1.MoveNext
        Next i
    End If

      Screen.MousePointer = 0
       
      If R1.RecordCount = 0 Then
        MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
        txtDocPer.SetFocus
        cmdAceptar.Default = False
      Else
        FECliente.lbEditarFlex = True
        Screen.MousePointer = 0
        FECliente.SetFocus
      End If
   End If
End Sub

Private Sub txtNomPer_GotFocus()
    fEnfoque txtNomPer
End Sub

Private Sub txtNomPer_KeyPress(KeyAscii As Integer)
Dim ClsPersona As COMDPersona.DCOMPersonas
Dim R1 As ADODB.Recordset

    If Left(txtNomPer.Text, 1) = "%" And Len(txtNomPer.Text) = "1" Then
        txtNomPer.Text = ""
    End If

   If KeyAscii = 13 Then
      If Len(Trim(txtNomPer.Text)) = 0 Then
        MsgBox "Falta Ingresar el Nombre de la Persona", vbInformation, "Aviso"
        Exit Sub
      End If

      Screen.MousePointer = 11
      Set ClsPersona = New COMDPersona.DCOMPersonas
      If bInstit Then
        Set R1 = ClsPersona.BuscaClienteServiciosNombre(txtNomPer.Text, bInstit, BusquedaNombre)
      Else
         Set R1 = ClsPersona.BuscaClienteServiciosNombre(txtNomPer.Text, bInstit, BusquedaNombre)
      End If
           
      LimpiaFlex FECliente
      If Not R1.EOF And Not R1.BOF Then
        For i = 0 To R1.RecordCount - 1
            FECliente.AdicionaFila
            FECliente.TextMatrix(i + 1, 1) = R1!valor
            R1.MoveNext
        Next i
    End If

      Screen.MousePointer = 0
       
      If R1.RecordCount = 0 Then
        MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
        txtNomPer.SetFocus
        cmdAceptar.Default = False
      Else
            FECliente.lbEditarFlex = True
            FECliente.Enabled = True
            Screen.MousePointer = 0
            FECliente.SetFocus
      End If
   End If

''''''Dim ClsPersona As COMDPersona.DCOMPersonas
''''''   If KeyAscii = 13 Then
''''''      If Len(Trim(txtNomPer.Text)) = 0 Then
''''''        MsgBox "Falta Ingresar el Nombre de la Persona", vbInformation, "Aviso"
''''''        Exit Sub
''''''      End If
''''''
''''''      Screen.MousePointer = 11
''''''      Set ClsPersona = New COMDPersona.DCOMPersonas
''''''      If bInstit Then
''''''        Set R = ClsPersona.BuscaClienteServicios(txtNomPer.Text, bInstit, BusquedaEmpleadoNombre)
''''''      Else
''''''         Set R = ClsPersona.BuscaClienteServicios(txtNomPer.Text, bInstit, BusquedaEmpleadoNombre)
''''''      End If
''''''
''''''      LimpiaFlex FEPagoServicio
''''''      If Not R.EOF And Not R.BOF Then
''''''        For i = 0 To R.RecordCount - 1
''''''            FEPagoServicio.AdicionaFila , , True
''''''            FEPagoServicio.TextMatrix(i + 1, 2) = PstaNombre(R!cNombre)
''''''            FEPagoServicio.TextMatrix(i + 1, 3) = Trim(R!cConcepto)
''''''            FEPagoServicio.TextMatrix(i + 1, 4) = Trim(R!nDeuNro)
''''''            FEPagoServicio.TextMatrix(i + 1, 5) = Trim(R!cCodServicio)
''''''            FEPagoServicio.TextMatrix(i + 1, 6) = Format(R!nImporteCuota, "#0.00")
'''''''            FEPagoServicio.TextMatrix(i + 1, 7) = str(R!cPeriodo)
''''''            R.MoveNext
''''''        Next i
''''''    End If
''''''    FEPagoServicio.lbEditarFlex = True
''''''
''''''      Screen.MousePointer = 0
''''''      If R.RecordCount = 0 Then
''''''        MsgBox "No se Encontraron Datos", vbInformation, "Aviso"
''''''        txtNomPer.SetFocus
''''''        cmdAceptar.Default = False
''''''      Else
''''''        cmdAceptar.Default = True
''''''        txtNomPer.Text = FEPagoServicio.TextMatrix(1, 2) 'Trim(R!cNombre)
''''''      End If
''''''   Else
''''''        KeyAscii = Letras(KeyAscii)
''''''        cmdAceptar.Default = False
''''''   End If
End Sub

Private Sub FEPagoServicio_OnCellChange(pnRow As Long, pnCol As Long)
    Call FEPagoServicio_OnCellCheck(1, 1)
End Sub

Private Sub FEPagoServicio_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
Dim i As Integer
Dim iint As Integer
    nTotCred = 0
    nMontoAPagar = 0
    LblTotCredPag.Caption = ""
    LblTotPag.Caption = ""
    For i = 0 To FEPagoServicio.Rows - 2
        If i = 0 Then
            ReDim Preserve sNumdeuSist(FEPagoServicio.Rows - 2)
        End If
        If FEPagoServicio.TextMatrix(i + 1, 1) = "." Then
            nMontoAPagar = nMontoAPagar + CDbl(FEPagoServicio.TextMatrix(i + 1, 5))
            sNumdeuSist(nTotCred) = Trim(FEPagoServicio.TextMatrix(i + 1, 3))
            nTotCred = nTotCred + 1
        End If
    Next i
    LblTotCredPag.Caption = nTotCred
    LblTotPag.Caption = Format(nMontoAPagar, "#0.00")
End Sub
