VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmCredPagoCuotasPersParam 
   Caption         =   "Parametros de Pago de Cuotas por Agencia"
   ClientHeight    =   4965
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7170
   Icon            =   "frmCredPagoCuotasPersParam.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7170
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Caption         =   "Montos por Agencia"
      Height          =   3975
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   7095
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdAgencia 
         Height          =   3615
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   6735
         _ExtentX        =   11880
         _ExtentY        =   6376
         _Version        =   393216
         Cols            =   5
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
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
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   4080
      Width           =   7095
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "Cancelar"
         Height          =   375
         Left            =   2760
         TabIndex        =   4
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         Enabled         =   0   'False
         Height          =   375
         Left            =   1440
         TabIndex        =   3
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cdmSalir 
         Caption         =   "Salir"
         Height          =   375
         Left            =   5640
         TabIndex        =   1
         Top             =   240
         Width           =   1215
      End
   End
End
Attribute VB_Name = "frmCredPagoCuotasPersParam"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim objCred As COMNCredito.NCOMCredito
Dim matris() As String

Private Sub cdmSalir_Click()
    Unload Me
End Sub

Private Sub cmdCancelar_Click()
    grdAgencia.Enabled = False
    cmdGrabar.Enabled = False
    cmdEditar.Enabled = True
    grdAgencia_LeaveCell
   
End Sub

Private Sub CmdEditar_Click()
    grdAgencia.Enabled = True
    cmdGrabar.Enabled = True
    cmdEditar.Enabled = False
    grdAgencia.Row = 1
    grdAgencia.Col = 3
    grdAgencia.SetFocus
    grdAgencia_EnterCell
End Sub

Private Sub CmdGrabar_Click()
    Set objCred = New COMNCredito.NCOMCredito
    Dim clsMov As COMNContabilidad.NCOMContFunciones
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    Dim sMovNro As String
    
     
    If MsgBox("Esta Seguro de Registrar los Parametros para el Pago de Créditos?", vbInformation + vbYesNo) = vbYes Then
        sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
        Dim i As Integer
        objCred.actualizaCredPagoCuotasAgeParam
        For i = 1 To grdAgencia.Rows - 1
            objCred.insertarCredPagoCuotasAgeParam grdAgencia.TextMatrix(i, 1), matris(i, 3), matris(i, 4), sMovNro
        Next
        MsgBox "Se registraron los Parametros con Exito!", vbInformation
        cmdGrabar.Enabled = False
        cmdEditar.Enabled = True
        grdAgencia.Enabled = False
        grdAgencia.CellBackColor = &H8000000E
    End If
End Sub

Private Sub Form_Load()
grdAgencia.Enabled = False
ConfigGridAgencia
CargarAgenciasParametro

End Sub
Private Sub CargarAgenciasParametro()
    Dim rsAgencia As New ADODB.Recordset
       
    Set objCred = New COMNCredito.NCOMCredito
    Set rsAgencia.DataSource = objCred.obtieneListaCredPagoCuotasAgeParam
    Dim i As Integer
    If Not rsAgencia.EOF Or rsAgencia.BOF Then
            grdAgencia.Rows = grdAgencia.Rows + rsAgencia.RecordCount - 1
            ReDim matris(grdAgencia.Rows, grdAgencia.Cols)
            For i = 0 To rsAgencia.RecordCount - 1
                grdAgencia.TextMatrix(i + 1, 0) = i + 1
                grdAgencia.TextMatrix(i + 1, 1) = rsAgencia!cAgeCod
                grdAgencia.TextMatrix(i + 1, 2) = rsAgencia!cAgeDescripcion
                grdAgencia.TextMatrix(i + 1, 3) = "S/." & Format(rsAgencia!nMontoMin, "#,###,##0.00")
                matris(i + 1, 3) = rsAgencia!nMontoMin
                grdAgencia.TextMatrix(i + 1, 4) = "S/." & Format(rsAgencia!nMontoMax, "#,###,##0.00")
                matris(i + 1, 4) = rsAgencia!nMontoMax
                rsAgencia.MoveNext
            Next
    End If
   
End Sub
Private Sub ConfigGridAgencia()
    grdAgencia.Clear
    grdAgencia.Rows = 2
    
    With grdAgencia
        .TextMatrix(0, 0) = "#"
        .TextMatrix(0, 1) = "Cod"
        .TextMatrix(0, 2) = "Agencia"
        .TextMatrix(0, 3) = "Monto Min"
        .TextMatrix(0, 4) = "Monto Max"
                
        .ColWidth(0) = 300
        .ColWidth(1) = 500
        .ColWidth(2) = 3000
        .ColWidth(3) = 1300
        .ColWidth(4) = 1300
               
        .ColAlignment(1) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignRightCenter
        .ColAlignment(4) = flexAlignRightCenter
        .ColAlignmentFixed(1) = flexAlignCenterCenter
        .ColAlignmentFixed(3) = flexAlignCenterCenter
        .ColAlignmentFixed(4) = flexAlignCenterCenter
 
        
    End With
End Sub

Private Sub grdAgencia_EnterCell()
    With grdAgencia
        If (.Col > 2 And .Col < (.Cols)) Then
        .CellBackColor = &H80000018
        .Tag = ""
        End If
    End With
End Sub

Private Sub grdAgencia_KeyDown(KeyCode As Integer, Shift As Integer)
    With grdAgencia
     Select Case KeyCode
        Case 46 'si presiona tecla del
            .Tag = grdAgencia
           
            If (.Col > 2 And .Col < (.Cols)) Then
                 matris(.Row, .Col) = "0"
                 '.TextMatrix(.Row, .Col) = Format("0", "S/.#,###,##0.00")
                 grdAgencia = "S/." & Format("0", "#,###,##0.00")
            End If
        

      End Select
    End With
End Sub

Private Sub grdAgencia_KeyPress(KeyAscii As Integer)
 
With grdAgencia
    'si es enter
    If KeyAscii = 13 Then
        'If Len(matris(.Row, .Col)) >= 12 Then
            'si esta en las columnas 5 de edicion
            If .Col > 2 And .Col < 4 And .TextMatrix(.Row, .Col) <> "" Then
                    If .Col = 3 Then
                        'If CLng(matris(.Row, .Col)) > 9999 Then
                        '    .TextMatrix(.Row, .Col) = ""
                        '    MsgBox "Solo esta permitido hasta S/.9999", vbInformation
                        '    Exit Sub
                        'End If
                    End If
                                    
                    .CellBackColor = &H8000000E
                    .Row = .Row
                    .Col = .Col + 1
                    .CellBackColor = &H80000018
            'si esta en las columnas 5 de edicion
            ElseIf .Row <= .Rows - 1 And .TextMatrix(.Row, .Col) <> "" Then
                    If .Col = 4 Then
                        If CLng(matris(.Row, .Col)) > 9999 Then
                          'MsgBox "Solo esta permitido hasta S/.9999", vbInformation
                            'Exit Sub
                        ElseIf CLng(matris(.Row, .Col)) < CLng(matris(.Row, .Col - 1)) Then
                          MsgBox "El Monto Maximo debe ser Mayor que el Monto Minimo", vbInformation
                            Exit Sub
                        End If
                    End If
                    .CellBackColor = &H8000000E
                    
                    If .Row < .Rows - 1 Then 'si no esta en la ultima fila y columna
                        .Row = .Row + 1
                        .Col = 3
                        .CellBackColor = &H80000018
                    Else
                         Me.cmdGrabar.SetFocus
                    End If
                   
            End If
    
    ElseIf KeyAscii = 8 Then 'si es retroceso
            
                If Len(.TextMatrix(.Row, .Col)) > 0 Then
                        If Len(matris(.Row, .Col)) > 1 Then
                            matris(.Row, .Col) = Mid(matris(.Row, .Col), 1, Len(matris(.Row, .Col)) - 1) 'retrocede 1 digito
                        Else
                            matris(.Row, .Col) = "0"
                        End If
                        grdAgencia = "S/." & Format(matris(.Row, .Col), "#,###,##0.00")
                End If
            
    ElseIf (.Col > 2 And .Col < (.Cols)) Then
           
                If InStr("0123456789.", Chr(KeyAscii)) = 0 Then 'compara si no es un numero
                        KeyAscii = 0
                        .TextMatrix(.Row, .Col) = ""
                Else
                    
                    matris(.Row, .Col) = matris(.Row, .Col) + Chr(KeyAscii)
                    'If CLng(matris(.Row, .Col)) < 10000 Then ' si es menor al monto del REU
                    If Len(matris(.Row, .Col)) < 8 Then
                        .TextMatrix(.Row, .Col) = "S/." & Format(matris(.Row, .Col), "#,###,##0.00")
                        grdAgencia = .TextMatrix(.Row, .Col)
                    Else
                        MsgBox "Ud. pasó el limite de digitos permitidos en la celda. Hay demadiados numeros", vbExclamation, "Mensaje"
                    End If
                    'Else
                    '    .TextMatrix(.Row, .Col) = "S/." & Format(matris(.Row, .Col), "#,###,##0.00")
                    '    MsgBox "Solo esta permitido ingresar hasta S/.9,999"
                    '    matris(.Row, .Col) = Mid(matris(.Row, .Col), 1, Len(matris(.Row, .Col)) - 1)
                    '    .TextMatrix(.Row, .Col) = "S/." & Format(matris(.Row, .Col), "#,###,##0.00")
                    'End If
                    
                End If
        'Else
        '    MsgBox "Demasiados numeros"
        'End If
    End If
                    
End With
End Sub

Private Sub grdAgencia_LeaveCell()
    With grdAgencia
        Dim s As String
        If .Col > 2 And .Col < (.Cols) Then
              
              If .Col = 4 Then
              
                If CLng(matris(.Row, .Col)) < CLng(matris(.Row, .Col - 1)) Then
                   MsgBox "El Monto Maximo debe ser Mayor que el Monto Minimo", vbInformation
                   matris(.Row, .Col) = "0"
                   matris(.Row, .Col - 1) = "0"
                   .TextMatrix(.Row, .Col - 1) = "S/." & Format("0", "#,###,##0.00")
                   grdAgencia = "S/." & Format("0", "#,###,##0.00")
                End If
              End If
              .CellBackColor = &H8000000E
        End If
        
    End With
End Sub


