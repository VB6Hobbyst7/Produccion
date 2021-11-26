VERSION 5.00
Begin VB.Form frmCoberturaSeguroAgencia 
   Caption         =   "Cobertura del Seguro por Agencia"
   ClientHeight    =   6105
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11085
   Icon            =   "frmCoberturaSeguroAgencia.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6105
   ScaleWidth      =   11085
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   3720
      TabIndex        =   7
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox txtAnio 
      Height          =   375
      Left            =   9840
      MaxLength       =   4
      TabIndex        =   5
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Height          =   375
      Left            =   6120
      TabIndex        =   3
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdModificar 
      Caption         =   "&Modificar"
      Height          =   375
      Left            =   4920
      TabIndex        =   1
      Top             =   5640
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton cmdGenerarAnio 
      Caption         =   "&Generar Año"
      Height          =   375
      Left            =   8520
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cboAnio 
      Height          =   315
      Left            =   4680
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin Sicmact.FlexEdit FECoberturaME 
      Height          =   4935
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   8705
      Cols0           =   14
      FixedCols       =   0
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "Agencia-Ene-Feb-Mar-Abr-May-Jun-Jul-Ago-Sep-Oct-Nov-Dic-Salto"
      EncabezadosAnchos=   "2500-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1200-1"
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
      ColumnasAEditar =   "X-1-2-3-4-5-6-7-8-9-10-11-12-X"
      ListaControles  =   "0-0-0-0-0-0-0-0-0-0-0-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "L-R-R-R-R-R-R-R-R-R-R-R-R-C"
      FormatosEdit    =   "0-4-4-4-4-4-4-4-4-4-4-4-4-0"
      TextArray0      =   "Agencia"
      lbUltimaInstancia=   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      ColWidth0       =   2505
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
   Begin VB.Label lblAnno 
      Alignment       =   2  'Center
      Caption         =   "Año"
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   120
      Width           =   375
   End
End
Attribute VB_Name = "frmCoberturaSeguroAgencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************
'***Nombre:         frmCoberturaSeguroAgencia
'***Descripción:    Formulario que permite el registro de la
'                   cobertura del seguro por Agencia.
'***Creación:       ELRO el 20111026 según Acta 277-2011/TI-D
'************************************************************
Option Explicit

Private Sub cboAnio_Click()
    If Me.cboAnio.ListIndex > -1 Then
        Call llenarFECoberturaME(CInt(Me.cboAnio))
    End If
End Sub

Private Sub cmdAceptar_Click()
    Dim oDAgencia As DAgencia
    Set oDAgencia = New DAgencia
    Dim oNContFunciones As NContFunciones
    Set oNContFunciones = New NContFunciones
    Dim nFila, i As Integer
    Dim lsMov As String
    
    lsMov = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
    
    nFila = Me.FECoberturaME.Rows - 1
    
    For i = 1 To nFila
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesEnero, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesEnero), lsMov)
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesFebrero, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesFebrero), lsMov)
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesMarzo, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesMarzo), lsMov)
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesAbril, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesAbril), lsMov)
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesMayo, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesMayo), lsMov)
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesJunio, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesJunio), lsMov)
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesJulio, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesJulio), lsMov)
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesAgosto, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesAgosto), lsMov)
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesSeptiembre, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesSeptiembre), lsMov)
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesOctubre, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesOctubre), lsMov)
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesNoviembre, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesNoviembre), lsMov)
        Call oDAgencia.modificarCoberturaAgencia(Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)), gMesDiciembre, Me.cboAnio, gMonedaExtranjera, Me.FECoberturaME.TextMatrix(i, gMesDiciembre), lsMov)
        
    Next i

MsgBox "Se modificó correctamente los datos", vbInformation, "Aviso"

Me.cmdAceptar.Visible = False
Me.cmdCancelar.Visible = False
Me.cmdModificar.Visible = True
Me.cmdGenerarAnio.Enabled = True
Me.txtAnio.Enabled = True
Me.cboAnio.Enabled = True
Me.FECoberturaME.lbEditarFlex = False
Set oDAgencia = Nothing
Set oNContFunciones = Nothing
End Sub

Private Sub cmdCancelar_Click()
Me.cmdAceptar.Visible = False
Me.cmdCancelar.Visible = False
Me.cmdModificar.Visible = True
Me.cmdGenerarAnio.Enabled = True
Me.txtAnio.Enabled = True
Me.cboAnio.Enabled = True
Me.FECoberturaME.lbEditarFlex = False
Call llenarFECoberturaME(CInt(Me.cboAnio))
End Sub

Private Sub cmdGenerarAnio_Click()
    If Me.txtAnio <> "" Then
        If verificarAnioGenerado(CInt(Me.txtAnio)) Then
            MsgBox "El Año " & Me.txtAnio & " ya fue generado", vbInformation, "Aviso"
        Else
            Dim oDAgencia As DAgencia
            Set oDAgencia = New DAgencia
            Dim rsGeneraCoberturaAgencia As ADODB.Recordset
            Set rsGeneraCoberturaAgencia = New ADODB.Recordset
            Dim oNContFunciones As NContFunciones
            Set oNContFunciones = New NContFunciones
            Dim lsMov As String
            
            lsMov = oNContFunciones.GeneraMovNro(gdFecSis, gsCodAge, gsCodUser)
            
            If MsgBox("¿Desea que al generar copie los datos de un Año en especifico?", vbYesNo, "Aviso") = vbYes Then
            Dim nAnioEspecifico  As Integer
                nAnioEspecifico = InputBox("Por favor ingrese el Año", "Copiar datos de un Año especifico")
                
                If verificarAnioGenerado(nAnioEspecifico) = False Then
                    MsgBox "El Año " & nAnioEspecifico & " aún no se genera, ingrese un Año que ya esta generado", vbInformation, "Aviso"
                    Exit Sub
                End If
                
                Set rsGeneraCoberturaAgencia = oDAgencia.recuperarCoberturaAgencia(nAnioEspecifico, gMonedaExtranjera)
                
                If Not rsGeneraCoberturaAgencia Is Nothing Then
                    If Not rsGeneraCoberturaAgencia.EOF Then
                         Do While Not rsGeneraCoberturaAgencia.EOF
                            Call oDAgencia.generarCoberturaAgencia(rsGeneraCoberturaAgencia!cAgecod, _
                                                                   rsGeneraCoberturaAgencia!nMES, _
                                                                   Me.txtAnio, _
                                                                   rsGeneraCoberturaAgencia!nMoneda, _
                                                                   rsGeneraCoberturaAgencia!nCobertura, _
                                                                   lsMov)
                                                    

                           
                            rsGeneraCoberturaAgencia.MoveNext
                        Loop
                    End If
                End If
            
            Else
                Dim lnAnio As Integer
                lnAnio = CInt(Me.txtAnio) - 1
                
                If verificarAnioGenerado(lnAnio) = False Then
                    MsgBox "Primero genere el Año " & lnAnio & " para generar el Año " & Me.txtAnio, vbInformation, "Aviso"
                    Exit Sub
                End If
                
                Set rsGeneraCoberturaAgencia = oDAgencia.recuperarCoberturaAgencia(lnAnio, gMonedaExtranjera)
                
                If Not rsGeneraCoberturaAgencia Is Nothing Then
                    If Not rsGeneraCoberturaAgencia.EOF Then
                         Do While Not rsGeneraCoberturaAgencia.EOF
                            Call oDAgencia.generarCoberturaAgencia(rsGeneraCoberturaAgencia!cAgecod, _
                                                                   rsGeneraCoberturaAgencia!nMES, _
                                                                   Me.txtAnio, _
                                                                   rsGeneraCoberturaAgencia!nMoneda, _
                                                                   rsGeneraCoberturaAgencia!nCobertura, _
                                                                   lsMov)
                                                    

                           
                            rsGeneraCoberturaAgencia.MoveNext
                        Loop
                    End If
                End If
            End If
                        
            MsgBox "Se genero correctamento el Año " & Me.txtAnio, vbInformation, "Aviso"
            Call llenarComboAnio
            Call verificarAnioGenerado(Me.txtAnio)
            Set oDAgencia = Nothing
            Set rsGeneraCoberturaAgencia = Nothing
            lsMov = ""
        End If
    Else
        MsgBox "Ingrese un Año", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdModificar_Click()
Me.cmdModificar.Visible = False
Me.cmdAceptar.Visible = True
Me.cmdCancelar.Visible = True
Me.cmdGenerarAnio.Enabled = False
Me.txtAnio.Enabled = False
Me.cboAnio.Enabled = False
Me.FECoberturaME.lbEditarFlex = True
End Sub

Private Sub Form_Load()
    Me.txtAnio = ""
    Me.txtAnio = Right(CStr(gdFecSis), 4)
    Call llenarComboAnio
    If Me.cboAnio.ListCount > 0 Then
        Me.cmdModificar.Visible = True
    End If
End Sub

Private Sub llenarComboAnio()
    Dim oDAgencia As DAgencia
    Set oDAgencia = New DAgencia
    Dim rsListaAnios As ADODB.Recordset
    Set rsListaAnios = New ADODB.Recordset
    
    Set rsListaAnios = oDAgencia.listarAniosCobertura
    
    If Not rsListaAnios Is Nothing Then
        If Not rsListaAnios.EOF Then
            Me.cboAnio.Clear
            Do While Not rsListaAnios.EOF
                Me.cboAnio.AddItem Trim(rsListaAnios!nAnio)
                rsListaAnios.MoveNext
            Loop
        Else
            MsgBox "Genere el Año " & Right(CStr(gdFecSis), 4), vbInformation, "Aviso"
            Exit Sub
        End If
    Else
        MsgBox "Genere el Año " & Right(CStr(gdFecSis), 4), vbInformation, "Aviso"
        Exit Sub
    End If
    
    Set oDAgencia = Nothing
    Set rsListaAnios = Nothing
    Me.cboAnio.ListIndex = -1
End Sub

Private Sub llenarFECoberturaME(ByVal pnAnio As Integer)
    Dim oDAgencia As DAgencia
    Set oDAgencia = New DAgencia
    Dim rsCobertutasAgencias As ADODB.Recordset
    Set rsCobertutasAgencias = New ADODB.Recordset
    Dim nRows, i As Integer
    
         
    Set rsCobertutasAgencias = oDAgencia.recuperarCoberturaAgencia(pnAnio, gMonedaExtranjera)
            
    If Not rsCobertutasAgencias Is Nothing Then
        
        If Not rsCobertutasAgencias.EOF Then
        
            Call LimpiaFlex(Me.FECoberturaME)
            Me.FECoberturaME.lbEditarFlex = True
            
            Do While Not rsCobertutasAgencias.EOF
                       
                If Trim(rsCobertutasAgencias!cAgecod) <> Trim(Right(Me.FECoberturaME.TextMatrix(i, 0), 3)) Then
                     i = i + 1
                     Me.FECoberturaME.AdicionaFila
                End If
            
                Me.FECoberturaME.TextMatrix(i, 0) = rsCobertutasAgencias!cAgeDescripcion
                Me.FECoberturaME.BackColorRow (&HC0FFC0)
                
                Select Case rsCobertutasAgencias!nMES
                    Case gMesEnero
                        Me.FECoberturaME.TextMatrix(i, gMesEnero) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                    Case gMesFebrero
                        Me.FECoberturaME.TextMatrix(i, gMesFebrero) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                    Case gMesMarzo
                        Me.FECoberturaME.TextMatrix(i, gMesMarzo) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                    Case gMesAbril
                        Me.FECoberturaME.TextMatrix(i, gMesAbril) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                    Case gMesMayo
                        Me.FECoberturaME.TextMatrix(i, gMesMayo) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                    Case gMesJunio
                        Me.FECoberturaME.TextMatrix(i, gMesJunio) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                    Case gMesJulio
                        Me.FECoberturaME.TextMatrix(i, gMesJulio) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                    Case gMesAgosto
                        Me.FECoberturaME.TextMatrix(i, gMesAgosto) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                    Case gMesSeptiembre
                        Me.FECoberturaME.TextMatrix(i, gMesSeptiembre) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                    Case gMesOctubre
                        Me.FECoberturaME.TextMatrix(i, gMesOctubre) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                    Case gMesNoviembre
                        Me.FECoberturaME.TextMatrix(i, gMesNoviembre) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                    Case gMesDiciembre
                        Me.FECoberturaME.TextMatrix(i, gMesDiciembre) = Format(rsCobertutasAgencias!nCobertura, "#,#0.00")
                End Select
                                   
                rsCobertutasAgencias.MoveNext
                
            Loop
           
        End If
    
    End If
    
    Me.FECoberturaME.lbEditarFlex = False
    Set rsCobertutasAgencias = Nothing
    Me.cmdModificar.SetFocus
End Sub

Private Function verificarAnioGenerado(ByVal pnAnio As Integer) As Boolean
    verificarAnioGenerado = False
    Dim i As Integer
    
     For i = 0 To Me.cboAnio.ListCount - 1
     
        If Me.cboAnio.List(i) = pnAnio Then
            Me.cboAnio.ListIndex = i
            verificarAnioGenerado = True
            Exit Function
        End If
     
     Next

End Function

Private Sub txtAnio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
If KeyAscii = 13 Then
    Me.cmdGenerarAnio.SetFocus
End If
End Sub
