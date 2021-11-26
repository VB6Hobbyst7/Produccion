VERSION 5.00
Begin VB.Form frmGeneraRelacionArqueo 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   3330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7290
   Icon            =   "frmGeneraRelacionArqueo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3330
   ScaleWidth      =   7290
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "Salir"
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
      Left            =   4110
      TabIndex        =   4
      Top             =   2970
      Width           =   1005
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
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
      Left            =   5190
      TabIndex        =   3
      Top             =   2970
      Width           =   1005
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar"
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
      Left            =   6300
      TabIndex        =   2
      Top             =   2970
      Width           =   1005
   End
   Begin VB.Frame fraUsuarios 
      Caption         =   "Exclusión"
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
      Height          =   2865
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   7245
      Begin SICMACT.FlexEdit grdUsuarios 
         Height          =   2595
         Left            =   60
         TabIndex        =   1
         Top             =   210
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   4577
         Cols0           =   7
         HighLight       =   1
         AllowUserResizing=   3
         RowSizingMode   =   1
         EncabezadosNombres=   "#-Usuario-Nombre-Cargo-Excluir-nEstado-UserArqueador"
         EncabezadosAnchos=   "450-800-2500-1900-850-0-0"
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
         ColumnasAEditar =   "X-X-X-X-4-X-X"
         ListaControles  =   "0-0-0-0-4-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-0"
         TextArray0      =   "#"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
         ForeColorFixed  =   -2147483630
      End
   End
End
Attribute VB_Name = "frmGeneraRelacionArqueo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdCancelar_Click()
    If MsgBox("El formulario retornará a su estado inicial, ¿desea continuar?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        grdUsuarios.SetFocus
        RefrescarGrid
    End If
End Sub
Private Sub cmdGenerar_Click()

Dim nTop As Integer, nAleatorio As Integer, nRes As Integer, i As Integer, Y As Integer, j As Integer
Dim sRelaciones As String, sUsuarioSinPar As String, sUsuariosExcluidos As String
Dim sMovNro As String, sMensaje As String
Dim sUsers() As String
Dim nEstado As Integer
Dim clsMov As COMNContabilidad.NCOMContFunciones
Dim ClsArq As COMDConstSistema.DCOMGeneral
Dim rsTemp As New ADODB.Recordset
Dim bContinuar As Boolean

If grdUsuarios.Rows = 2 And Trim(grdUsuarios.TextMatrix(1, 1)) = "" Then
    MsgBox "No es posible generar las parejas por que no hay usuarios disponibles.", vbInformation, "Aviso"
    Exit Sub
End If
For nTop = 1 To grdUsuarios.Rows - 1
    If grdUsuarios.TextMatrix(nTop, 4) = "." Then
        sUsuariosExcluidos = sUsuariosExcluidos & grdUsuarios.TextMatrix(nTop, 1) & ","
    ElseIf IIf(IsNumeric(grdUsuarios.TextMatrix(nTop, 5)), grdUsuarios.TextMatrix(nTop, 5), 0) > 1 _
           And IIf(IsNumeric(grdUsuarios.TextMatrix(nTop, 5)), grdUsuarios.TextMatrix(nTop, 5), 0) <> 6 Then
    Else
        i = i + 1
    End If
Next
If i <= 3 Then
    MsgBox "Solo pueden generarse parejas con un numero mayor a tres usuarios", vbInformation, "Aviso"
    Exit Sub
End If

'Proceso de verfificacion de estado de usuarios ********
Y = 0
sMensaje = " Se procederá a crear las parejas de arqueo " & vbNewLine & "  ¿Desea continuar?"
For i = 1 To grdUsuarios.Rows - 1
    If IIf(IsNumeric(grdUsuarios.TextMatrix(i, 5)), grdUsuarios.TextMatrix(i, 5), 0) > 0 And Len(Trim(grdUsuarios.TextMatrix(i, 6))) > 0 Then
        Y = Y + 1
        Exit For
    End If
Next
If Y > 0 Then
        sMensaje = "Verificar que las parejas de arqueo que hayan iniciado el proceso " & vbNewLine & _
                   "lo concluyan, caso contrario indicarles que detengan dicho proceso." & vbNewLine & vbNewLine & _
                   "Se procederá a crear las parejas de arqueo " & vbNewLine & "  ¿Desea continuar?"
End If
Y = 0
'Fin de proceso de verificacion ************************

' Obteniendo parejas del dia anterior
Set ClsArq = New COMDConstSistema.DCOMGeneral
Set rstemp = ClsArq.GetUserUltimoArqueo(gdFecSis, gsCodArea, gsCodAge, IIf(gsCodAge = "01", 1, 2)) 'ADD BY JATO 20210105
'Set rstemp = ClsArq.GetUserAreaAgenciaRelacion(Format(DateAdd("d", -1, CDate(gdFecSis)), "dd/MM/yyyy"), gsCodArea, gsCodAge, IIf(gsCodAge = "01", 1, 2)) 'COM BY JATO 20210106

If MsgBox(sMensaje, vbQuestion + vbYesNo, "Aviso") = vbYes Then
    grdUsuarios.SetFocus
    RefrescarGrid True
        
    bContinuar = True
    Do While bContinuar
        bContinuar = False
        ReDim sUsers(0)
        i = 0
        For nTop = 1 To grdUsuarios.Rows - 1
            If grdUsuarios.TextMatrix(nTop, 4) = "." Then
                sUsuariosExcluidos = sUsuariosExcluidos & grdUsuarios.TextMatrix(nTop, 1) & ","
            ElseIf IIf(IsNumeric(grdUsuarios.TextMatrix(nTop, 5)), grdUsuarios.TextMatrix(nTop, 5), 0) > 1 _
                   And IIf(IsNumeric(grdUsuarios.TextMatrix(nTop, 5)), grdUsuarios.TextMatrix(nTop, 5), 0) <> 6 Then
            Else
                i = i + 1
                ReDim Preserve sUsers(i)
                sUsers(i) = grdUsuarios.TextMatrix(nTop, 1)
            End If
        Next
        If i <= 3 Then
            MsgBox "Solo pueden generarse parejas con un numero mayor a tres usuarios", vbInformation, "Aviso"
            Exit Sub
        End If
        ReDim Preserve sUsers(i + 1)
        
        ' Iniciando la creacion de parejas aleatorias.
        Randomize
        i = 0: Y = 0
        nTop = UBound(sUsers) - 1
        nRes = nTop Mod 2 ' "0" si es par, "1" si es impar.
        Do While nTop > nRes
            nAleatorio = 1 + Int(Rnd() * nTop)
            i = i + 1
            If i Mod 2 = 1 Then
                sRelaciones = sRelaciones & sUsers(nAleatorio) & "-"
            Else
                sRelaciones = sRelaciones & sUsers(nAleatorio) & ","
            End If
            For Y = nAleatorio To nTop - 1
                sUsers(Y) = sUsers(Y + 1)
            Next
            nTop = nTop - 1
        Loop
            'JATO 20210106
            If Not rsTemp Is Nothing Then
                If rsTemp.State = 1 Then
                    If rsTemp.RecordCount > 0 Then

                        rsTemp.MoveFirst
                        Do While Not rsTemp.EOF And Not rsTemp.BOF
                            If InStr(1, sRelaciones, (rsTemp!cUserArqueado & "-" & IIf(Trim(rsTemp!cUserArqueador) = "", "xxxxxx", rsTemp!cUserArqueador))) > 0 Then
                                bContinuar = True
                                sRelaciones = ""
                                rsTemp.MoveLast
                            End If
                            rsTemp.MoveNext
                        Loop

                    End If
                End If
            End If
            'JATO 20210106
        Loop
    ' fin de parejas aleatorias.
    If nRes > 0 Then
        sUsuarioSinPar = sUsers(1)
    End If
    sRelaciones = Mid(sRelaciones, 1, Len(sRelaciones) - 1)
    If Len(sUsuariosExcluidos) > 0 Then sUsuariosExcluidos = Mid(sUsuariosExcluidos, 1, Len(sUsuariosExcluidos) - 1)
    
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Set clsMov = Nothing
    nEstado = 1 'GENERADO
    
    If ClsArq.SetRelacionArqueo(sRelaciones, sUsuarioSinPar, nEstado, sMovNro, Trim(sUsuariosExcluidos)) Then
        MsgBox "Las parejas de arqueo se generaron correctamente", vbInformation, "Aviso"
        Unload Me
    Else
        MsgBox "Se presentaron inconvenientes durante la creación de las parejas de arqueo, coordinar con el area de T.I.", vbExclamation, "Aviso"
    End If
End If

End Sub
Private Sub cmdsalir_Click()
    Unload Me
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub
Private Sub Form_Load()
    Me.Caption = "Arqueos: Entre Ventanillas"
    RefrescarGrid
End Sub
Private Sub RefrescarGrid(Optional bGuardarExcluidos As Boolean = False)
    Dim oUsuariosArea As COMDConstSistema.DCOMGeneral
    Dim rs As New ADODB.Recordset
    Dim sRelacionExcluidos() As String
    Dim i As Integer, j As Integer
    Dim nTipoAge As Integer
    If gsCodAge = "01" Then
        nTipoAge = 1
    Else
        nTipoAge = 2
    End If
    Set oUsuariosArea = New COMDConstSistema.DCOMGeneral
    Set rs = oUsuariosArea.GetUserAreaAgenciaRelacion(gdFecSis, gsCodArea, gsCodAge, nTipoAge)
    'Guardando los excluidos
    ReDim Preserve sRelacionExcluidos(0)
    If bGuardarExcluidos Then
        For i = 1 To grdUsuarios.Rows - 1
            If grdUsuarios.TextMatrix(i, 4) = "." Then
                j = j + 1
                ReDim Preserve sRelacionExcluidos(j)
                sRelacionExcluidos(j) = grdUsuarios.TextMatrix(i, 1)
            End If
        Next i
    End If
    cargar rs, IIf(bGuardarExcluidos, False, True)
    'Asignando los excluidos
    If bGuardarExcluidos Then
        For i = 1 To UBound(sRelacionExcluidos)
            For j = 1 To grdUsuarios.Rows - 1
                If sRelacionExcluidos(i) = Trim(grdUsuarios.TextMatrix(j, 1)) Then
                    grdUsuarios.TextMatrix(j, 4) = "1"
                    Exit For
                End If
            Next j
        Next i
    End If
    grdUsuarios.row = 1
    grdUsuarios.Col = 1
    SendKeys "{left}"
End Sub
Private Sub cargar(ByVal rs As ADODB.Recordset, Optional bCargarExcluidos As Boolean = True)
    Dim i As Integer, j As Integer, nEstado As Integer
    grdUsuarios.Clear
    grdUsuarios.Rows = 2
    grdUsuarios.FormaCabecera
    
    If Not rs.BOF And Not rs.EOF Then
        grdUsuarios.AdicionaFila
        For i = 1 To rs.RecordCount
            grdUsuarios.TextMatrix(i, 0) = i
            grdUsuarios.TextMatrix(i, 1) = rs!cUserArqueado
            grdUsuarios.TextMatrix(i, 2) = rs!cPersNombreArqueado
            grdUsuarios.TextMatrix(i, 3) = rs!cRHCargoArqueado
            If bCargarExcluidos Then grdUsuarios.TextMatrix(i, 4) = IIf(rs!nEstado = 6, 1, 0)
            grdUsuarios.TextMatrix(i, 5) = rs!nEstado
            grdUsuarios.TextMatrix(i, 6) = rs!cUserArqueador
            If rs.RecordCount > i Then
                grdUsuarios.AdicionaFila
                rs.MoveNext
            End If
        Next
    End If
    For i = 1 To grdUsuarios.Rows - 1
        If IIf(IsNumeric(grdUsuarios.TextMatrix(i, 5)), grdUsuarios.TextMatrix(i, 5), 0) > 1 And _
           IIf(IsNumeric(grdUsuarios.TextMatrix(i, 5)), grdUsuarios.TextMatrix(i, 5), 0) <> 6 Then
            For j = 1 To grdUsuarios.Cols - 1
                grdUsuarios.row = i
                grdUsuarios.Col = j
                grdUsuarios.CellBackColor = &H80000013
            Next
        End If
    Next i
    i = 0: j = 0
    For i = 1 To grdUsuarios.Rows - 1
        If grdUsuarios.TextMatrix(i, 4) = "." Or _
           IIf(IsNumeric(grdUsuarios.TextMatrix(i, 5)), grdUsuarios.TextMatrix(i, 5), 0) > 1 Then
        Else
            j = j + 1
        End If
    Next
    'FRHU 20140819
    'If j <= 4 Then
    '    cmdGenerar.Enabled = False
    'End If
    If (grdUsuarios.Rows - 1) <= 3 Then
        cmdGenerar.Enabled = False
    End If
    'FIN FRHU
End Sub

Private Sub grdUsuarios_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(grdUsuarios.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub
Private Sub grdUsuarios_RowColChange()
    If IIf(IsNumeric(grdUsuarios.TextMatrix(grdUsuarios.row, 5)), grdUsuarios.TextMatrix(grdUsuarios.row, 5), 0) > 1 And _
       IIf(IsNumeric(grdUsuarios.TextMatrix(grdUsuarios.row, 5)), grdUsuarios.TextMatrix(grdUsuarios.row, 5), 0) <> 6 Then
        
        grdUsuarios.lbEditarFlex = False
    Else
        grdUsuarios.lbEditarFlex = True
    End If
End Sub
