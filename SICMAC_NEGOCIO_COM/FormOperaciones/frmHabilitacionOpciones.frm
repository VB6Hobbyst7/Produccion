VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmHabilitacionOpciones 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Habilitación de opciones especiales"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8700
   Icon            =   "frmHabilitacionOpciones.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   8700
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ProgressBar progreso 
      Height          =   375
      Left            =   120
      TabIndex        =   10
      Top             =   5040
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      Height          =   375
      Left            =   7560
      TabIndex        =   2
      Top             =   5040
      Width           =   1000
   End
   Begin VB.CommandButton cmdNuevo 
      Caption         =   "Nuevo"
      Height          =   375
      Left            =   6480
      TabIndex        =   1
      Top             =   5040
      Width           =   1000
   End
   Begin VB.Frame IFrame 
      Caption         =   "Datos"
      Height          =   4695
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
      Begin VB.CommandButton cmdOtrasOpe 
         Caption         =   "Otras Ope."
         Height          =   375
         Left            =   7200
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   1095
      End
      Begin SICMACT.FlexEdit flxOperaciones 
         Height          =   2175
         Left            =   360
         TabIndex        =   11
         Top             =   2280
         Width           =   7815
         _extentx        =   13785
         _extenty        =   3836
         cols0           =   7
         highlight       =   1
         allowuserresizing=   3
         rowsizingmode   =   1
         encabezadosnombres=   "#-Item-Operación-Usuario-Cant. Asignada-Cant. Usada-cOpeCod"
         encabezadosanchos=   "700-1200-4000-1200-1200-1200-0"
         font            =   "frmHabilitacionOpciones.frx":030A
         font            =   "frmHabilitacionOpciones.frx":0336
         font            =   "frmHabilitacionOpciones.frx":0362
         font            =   "frmHabilitacionOpciones.frx":038E
         font            =   "frmHabilitacionOpciones.frx":03BA
         fontfixed       =   "frmHabilitacionOpciones.frx":03E6
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         backcolorcontrol=   -2147483643
         lbultimainstancia=   -1
         columnasaeditar =   "X-1-X-X-4-5-X"
         listacontroles  =   "0-4-0-0-0-0-0"
         encabezadosalineacion=   "C-C-L-C-C-C-C"
         formatosedit    =   "0-0-0-0-0-0-0"
         textarray0      =   "#"
         lbeditarflex    =   -1
         lbbuscaduplicadotext=   -1
         colwidth0       =   705
         rowheight0      =   300
         forecolorfixed  =   -2147483630
      End
      Begin VB.TextBox txtUsuario 
         Height          =   285
         Left            =   1920
         MaxLength       =   4
         TabIndex        =   6
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label lblCargo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   8
         Top             =   1440
         Width           =   3615
      End
      Begin VB.Label Label6 
         Caption         =   "Lista de Operaciones"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   1920
         Width           =   1935
      End
      Begin VB.Label lblNameComplete 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   1920
         TabIndex        =   7
         Top             =   960
         Width           =   3615
      End
      Begin VB.Label Label3 
         Caption         =   "Cargo"
         Height          =   255
         Left            =   360
         TabIndex        =   5
         Top             =   1440
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Nombre completo"
         Height          =   255
         Left            =   360
         TabIndex        =   4
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Usuario"
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   480
         Width           =   855
      End
   End
End
Attribute VB_Name = "frmHabilitacionOpciones"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim l_nCodPersona As String
Dim l_cUsuario As String
Dim l_bCargaOperacionesAsignadas As Boolean
Dim l_deselHabilitacion As Boolean

Private Sub cmdCancelar_Click()
    Unload Me
End Sub

Private Sub Guardar(ByVal pnRow As Integer, Optional ByVal pnEstado As Integer = 0)
    If GuardarAsignacion(pnRow, pnEstado) Then ' guardar asignación
        MsgBox "Configuración guardada", vbOKOnly + vbInformation, "Aviso"
        If MsgBox("¿Desea continuar configurando los permisos de operaciones para el usuario " & l_cUsuario & "?", vbYesNo + vbQuestion, "Aviso") = vbNo Then
            Call NuevaAsignacion
        Else
            l_bCargaOperacionesAsignadas = True
            Call CargarOperacionesHabilitadasXUser
            l_bCargaOperacionesAsignadas = False
        End If
    Else
        MsgBox "No se pudo asignar la operación por favor, intente de nuevo más tarde. Si el error persiste comuniquese con TI-Desarrollo.", vbError + vbOKOnly, "Aviso"
    End If
End Sub

Private Sub cmdNuevo_Click()
    Call NuevaAsignacion
    cmdOtrasOpe.Visible = False 'Add TORE 20200710: Adecuacion por HelpDesk
End Sub

'Add TORE 20200710: Adecuacion por HelpDesk
Private Sub cmdOtrasOpe_Click()
    frmMantPermisos.InicioOpeEspecial (Trim(txtUsuario.Text))
End Sub
'******************************************

Private Sub flxOperaciones_DblClick()
    If l_nCodPersona <> "" Then
        Dim row As Integer
        row = flxOperaciones.row
        If (flxOperaciones.TextMatrix(row, 1) = ".") Then
            If MsgBox("¿Desea aumentar la cantidad de operaciones asignadas?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                Call CargarDatosDeAsignacion(row, 1)
            End If
        Else
            MsgBox "No es posible aumentar la cantidad de operaciones a una operación que no se le fue asignada al usuario.", vbOKOnly + vbInformation, "Aviso"
            Exit Sub
        End If
    Else
        MsgBox "Primero debe ingresar un usuario", vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
End Sub

Private Sub flxOperaciones_OnCellCheck(ByVal pnRow As Long, ByVal pnCol As Long)
  If l_nCodPersona <> "" Then
    If Not l_bCargaOperacionesAsignadas Then
        If (flxOperaciones.TextMatrix(pnRow, 1) = ".") Then
            If Not l_deselHabilitacion Then
                If MsgBox("¿Desea asignar la opción al usuario?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                    Call CargarDatosDeAsignacion(pnRow)
                Else
                    flxOperaciones.TextMatrix(pnRow, 1) = ""
                End If
            Else
                l_deselHabilitacion = False
            End If
        Else
            If MsgBox("¿Está seguro que desea deshabilitar la opción al usuario?", vbYesNo + vbQuestion, "Aviso") = vbYes Then
                Call ActualizarDatosDeAsignacion(pnRow)
            Else
                l_deselHabilitacion = True
                With flxOperaciones
                    .TextMatrix(pnRow, 1) = "."
                    .SeleccionaChekTecla
                End With
            End If
        End If
    End If
Else
    flxOperaciones.TextMatrix(pnRow, 1) = ""
    MsgBox "Primero debe ingresar un usuario", vbOKOnly + vbInformation, "Aviso"
    Exit Sub
End If
End Sub

Private Sub Form_Load()
    Call Controles(False)
    Call CargarOperaciones
End Sub

Private Sub Controles(ByVal bEstado As Boolean, Optional ByVal pbNuevo As Boolean = True)
    txtUsuario.Enabled = Not bEstado
    progreso.value = 0
    
    If pbNuevo Then
        txtUsuario.Text = ""
        l_cUsuario = ""
        lblNameComplete.Caption = ""
        lblCargo.Caption = ""
        l_nCodPersona = ""
    End If
End Sub

Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    If (KeyAscii = 13) Then
        l_bCargaOperacionesAsignadas = True
        Call CargarDatos
        
        'Add TORE 20201007: Adecuacion por HelpDesk
        If (gsCodCargo = "005018" Or gsCodCargo = "006005") Then 'Solo SUPERVISOR, COORDINADOR DE OPERACIONES
            cmdOtrasOpe.Visible = True
        End If
        '******************************************
    End If
End Sub

Private Sub CargarOperaciones()
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim R As New ADODB.Recordset
    
    Set R = oCaptaLN.ObtenerOperacionesH
        
    If Not (R.BOF And R.EOF) Then
        Dim i As Integer
        
        With flxOperaciones
            .AdicionaFila , , True
            Do While (Not R.EOF)
                i = flxOperaciones.Rows - 1
                .TextMatrix(i, 2) = R!Operacion
                .TextMatrix(i, 4) = "0"
                .TextMatrix(i, 5) = "0"
                .TextMatrix(i, 6) = R!cOpecod
                .AdicionaFila , , True
                R.MoveNext
            Loop
            
            .EliminaFila (flxOperaciones.Rows - 1)
        End With
        
    End If
End Sub

Private Sub CargarOperacionesHabilitadasXUser()
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim R As New ADODB.Recordset
    
    Set R = oCaptaLN.ObtenerOperacionesHabilitadasxUsuario(l_cUsuario, gdFecSis)
    
    If (R.State = 0) Then
        Dim cMensaje As String
        cMensaje = "No se ha podido verificar que el usuario tiene operaciones especiales asignadas para el día de hoy"
        MsgBox cMensaje, vbOKOnly + vbInformation, "Aviso"
        Exit Sub
    End If
    
    progreso.value = 0
    
    If Not (R.BOF And R.EOF) Then
        Dim cOpecod As String
        
        Dim nRows As Integer, i As Integer
        nRows = flxOperaciones.Rows - 1
        progreso.Max = nRows
        
        For i = 1 To nRows
            progreso.value = progreso.value + 1
            cOpecod = flxOperaciones.TextMatrix(i, 6)
            Do While (Not R.EOF)
                If (cOpecod = R!cOpecod) Then
                    With flxOperaciones
                        .TextMatrix(i, 1) = "." 'check
                        .SeleccionaChekTecla
                        .TextMatrix(i, 3) = R!cUserQH 'user
                        .TextMatrix(i, 4) = R!nCantOpe 'cant operaciones
                        .TextMatrix(i, 5) = R!nCantUsadas ' cant usadas
                    End With
                End If
                R.MoveNext
            Loop
            R.MoveFirst
        Next i
        
        MsgBox "Se ha cargado las operaciones asignadas vigentes", vbOKOnly + vbInformation, "Aviso"
        progreso.value = 0
    End If
End Sub

Private Sub CargarDatos()
    
    l_cUsuario = Trim(txtUsuario.Text)
    
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim R As New ADODB.Recordset
    
    Set R = oCaptaLN.ObtenerUsuarios_RF_AS_SUP(l_cUsuario)
    If Not (R.BOF And R.EOF) Then
        Call Controles(True, False)
    
        l_nCodPersona = R!cPersCod
        lblNameComplete.Caption = R!NombreCompleto
        lblCargo.Caption = R!Cargo
        
        Call CargarOperacionesHabilitadasXUser
        l_bCargaOperacionesAsignadas = False
    Else
        Dim cMensaje As String
        cMensaje = "Usuario no válido. Solo permite la búsqueda de los usuarios con los cargos de: " & Chr$(10)
        cMensaje = cMensaje & "- SUPERVIDOR DE OPERACIONES " & Chr$(10)
        cMensaje = cMensaje & "- ASESOR DE CLIENTES " & Chr$(10)
        cMensaje = cMensaje & "- REPRESENTANTE FINANCIERO " & Chr$(10)
        cMensaje = cMensaje & "- REPRESENTANTE FINANCIERO III " & Chr$(10)
        cMensaje = cMensaje & "- TASADOR "
        
        MsgBox cMensaje, vbOKOnly + vbInformation, "Aviso"
    End If
End Sub

Private Sub limpiarflex()
    Dim nRows As Integer, i As Integer, iFila As Integer
    nRows = flxOperaciones.Rows - 1
    iFila = nRows
    For i = 0 To nRows
        flxOperaciones.EliminaFila iFila
        iFila = iFila - 1
    Next
End Sub

Private Sub CargarDatosDeAsignacion(ByVal pnRow As Integer, Optional nMomento As Integer = 0)
    Dim cValorIngresado As String, nValorIngresado As Integer, cMensaje As String
    
    If nMomento = 0 Then 'asignar
        cMensaje = "Ingrese la cantidad de operaciones a asignar:"
    Else 'aumentar
        cMensaje = "Ingrese la cantidad de operaciones a aumentar:"
    End If
    
    cValorIngresado = InputBox(cMensaje, "N° de operaciones")
    If (ValorIngresadoValido(cValorIngresado) = True) Then
        nValorIngresado = CInt(cValorIngresado)
        With flxOperaciones
            .TextMatrix(pnRow, 3) = gsCodUser
            If nMomento = 0 Then
                .TextMatrix(pnRow, 4) = nValorIngresado
            Else
                .TextMatrix(pnRow, 4) = CInt(.TextMatrix(pnRow, 4)) + nValorIngresado
            End If
            '.TextMatrix(pnRow, 5) = IIf(.TextMatrix(pnRow, 5) = "0", 0, CInt(.TextMatrix(pnRow, 5)) + nValorIngresado)
        End With
        Call Guardar(pnRow) 'estado 0 activa
    Else
        MsgBox "Valor no válido, solo se permite valor numérico.", vbOKOnly + vbInformation, "Aviso"
        flxOperaciones.TextMatrix(pnRow, 1) = ""
    End If
End Sub

Private Function ValorIngresadoValido(ByVal pValor As String) As Boolean

    Dim cadenaBlanca As String, i As Integer, valido As Boolean
    cadenaBlanca = "0123456789"
    
    For i = 1 To Len(pValor)
        If InStr(1, cadenaBlanca, Mid(pValor, i, 1)) > 0 Then
            valido = True
        Else
            valido = False
            Exit For
        End If
    Next i
    
    If valido Then
        If (CInt(pValor) = 0) Then
            valido = False
        End If
    End If
    
    ValorIngresadoValido = valido
End Function
Private Sub ActualizarDatosDeAsignacion(ByVal pnRow As Integer)
    With flxOperaciones
        .TextMatrix(pnRow, 3) = ""
        .TextMatrix(pnRow, 4) = 0
        .TextMatrix(pnRow, 5) = 0
    End With
    Call Guardar(pnRow, 1) 'estado 1 desactiva
End Sub

Private Function GuardarAsignacion(ByVal pnRow As Integer, ByVal pnEstado As Integer) As Boolean
    Dim oCaptaLN As New COMNCaptaGenerales.NCOMCaptaGenerales
    Dim i As Integer, nResultado As Integer, nCantAsig As Integer, nCantUsada As Integer
    'nRows As Integer, nTotalErrores As Integer
    Dim cUserH As String, cUserQH As String, cOpecod As String, cOpeCod_NoGuardadas As String, cFechaHoraSist As String
    
On Error GoTo ErrorGuardarAsignacion
    i = pnRow
    GuardarAsignacion = True 'estado normal tudu ok
    cUserH = l_cUsuario 'usuario al que se ha habilitado la operación
    cUserQH = gsCodUser 'usuario que asigna
    'nRows = flxOperaciones.Rows - 1
    'nTotalErrores = 0
    'progreso.Max = nRows
    
    'progreso.value = 0
    
    'For i = 1 To nRows
    '   progreso.value = progreso.value + 1
        With flxOperaciones
            'If .TextMatrix(i, 1) = "." Then
                nCantAsig = .TextMatrix(i, 4) 'cant asignada
                nCantUsada = .TextMatrix(i, 5) 'cant usada
                cOpecod = .TextMatrix(i, 6) 'cod Operación
                cFechaHoraSist = Format(gdFecSis, "yyyy-MM-dd") & " " & CStr(Time)
                nResultado = oCaptaLN.RegistrarHabilitacionDeOperacion(cUserH, cOpecod, cFechaHoraSist, cUserQH, nCantAsig, pnEstado)
                'If nResultado = 1 Then
                '    nTotalErrores = nTotalErrores + 1
                '    cOpeCod_NoGuardadas = cOpeCod_NoGuardadas & "," & cOpeCod
                'End If
            'End If
        End With
    'Next i
    
'    If nTotalErrores > 0 Then
'        If nTotalErrores = nRows Then
'            GuardarAsignacion = False
'        End If
'        cOpeCod_NoGuardadas = Mid(cOpeCod_NoGuardadas, 2, Len(cOpeCod_NoGuardadas) - 2)
'        Call DesglosarOperacionesNoAsignadas(cOpeCod_NoGuardadas)
'    End If
    
    If nResultado = 1 Then
        GuardarAsignacion = False
    End If
    
    Exit Function
ErrorGuardarAsignacion:
    GuardarAsignacion = False
End Function

Private Sub DesglosarOperacionesNoAsignadas(ByVal cOperaciones As String)
    Dim aOperaciones As Variant
    Dim cOperacionDescripcion As String
    Dim i As Integer, j As Integer, nRows As Integer
    
    aOperaciones = Split(cOperaciones, ",")
    nRows = flxOperaciones.Rows - 1
    
    For i = 0 To UBound(aOperaciones)
        For j = 1 To nRows
            If aOperaciones(i) = flxOperaciones.TextMatrix(j, 6) Then 'entrando operación
                cOperacionDescripcion = Chr$(10) & "-" & cOperacionDescripcion & flxOperaciones.TextMatrix(j, 2)
            End If
        Next j
    Next i
    
    MsgBox "El proceso de asignación terminó con error, las operaciones que no se registraron fueron: " & cOperacionDescripcion, vbOKOnly + vbInformation, "Aviso"
End Sub

Private Sub NuevaAsignacion()
    Call Controles(False)
    Call limpiarflex
    Call CargarOperaciones
End Sub
