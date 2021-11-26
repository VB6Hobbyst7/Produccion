VERSION 5.00
Begin VB.Form frmCredEvalParamEspecializacion 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Parámetros de Especialización de Créditos"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   15510
   Icon            =   "frmCredEvalParamEspecializacion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5985
   ScaleWidth      =   15510
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   5775
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   15375
      Begin VB.CommandButton cmdEditar 
         Caption         =   "Editar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   5
         Top             =   5160
         Width           =   1155
      End
      Begin VB.CommandButton cmdAgregar 
         Caption         =   "Agregar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   240
         TabIndex        =   4
         Top             =   5160
         Visible         =   0   'False
         Width           =   1245
      End
      Begin VB.CommandButton cmdEliminar 
         Caption         =   "Eliminar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   1680
         TabIndex        =   3
         Top             =   5160
         Visible         =   0   'False
         Width           =   1170
      End
      Begin VB.CommandButton cmdGrabar 
         Caption         =   "Grabar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   12840
         TabIndex        =   2
         Top             =   5280
         Width           =   1170
      End
      Begin VB.CommandButton cmdSalir 
         Caption         =   "Salir"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   14040
         TabIndex        =   1
         Top             =   5280
         Width           =   1170
      End
      Begin SICMACT.FlexEdit feEspecializacion 
         Height          =   4815
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   15135
         _ExtentX        =   26696
         _ExtentY        =   8493
         Cols0           =   10
         HighLight       =   1
         EncabezadosNombres=   "-Especializacion-Producto-Sub Producto-Tipo Credito-Sub Tipo Credito-Ultimo End(Mes)-End+Cred(Min S/.)-End+Cred(Max S/.)-Aux"
         EncabezadosAnchos=   "450-1900-1900-1800-2200-1900-1400-1560-1570-0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-1-2-3-4-5-6-7-8-X"
         ListaControles  =   "0-3-3-3-3-3-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-L-L-L-L-L-R-R-C-C"
         FormatosEdit    =   "0-0-0-0-0-0-3-2-2-2"
         lbEditarFlex    =   -1  'True
         lbUltimaInstancia=   -1  'True
         TipoBusqueda    =   3
         lbBuscaDuplicadoText=   -1  'True
         ColWidth0       =   450
         RowHeight0      =   300
      End
   End
End
Attribute VB_Name = "frmCredEvalParamEspecializacion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'**********************************************************************************************
'** Nombre : frmCredEvalParamEspecializacion
'** Descripción : Administración de parametros de especialización de créditos creado
'**               segun RFC090-2012
'** Creación : WIOR, 20120903 09:00:00 AM
'**********************************************************************************************

Option Explicit
Dim fnRow As Integer
Dim fnCol As Integer
Private Sub cmdAgregar_Click()
    If ValidaRegistroRepetido Then
        feEspecializacion.lbEditarFlex = True
        feEspecializacion.AdicionaFila
        feEspecializacion.SetFocus
        SendKeys "{Enter}"
        Call feEspecializacion_RowColChange
    End If
End Sub

Private Sub CmdEditar_Click()
    cmdAgregar.Visible = True
    cmdEliminar.Visible = True
    feEspecializacion.lbEditarFlex = True
    cmdEditar.Visible = False
    Call feEspecializacion_RowColChange
End Sub

Private Sub CmdEliminar_Click()
    If MsgBox("¿Está seguro de eliminar los datos de la fila " + CStr(feEspecializacion.Row) + "?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
        feEspecializacion.EliminaFila feEspecializacion.Row
    End If
End Sub

Private Sub cmdGrabar_Click()
    Dim oNCred As COMDCredito.DCOMCredActBD
    Dim i As Integer
    Dim nId As String
    Set oNCred = New COMDCredito.DCOMCredActBD
    
    If ValidaRegistroRepetido Then
        If ValidaGrilla Then
            If MsgBox("Los Datos seran Grabados, Desea Continuar ?", vbInformation + vbYesNo, "Aviso") = vbNo Then Exit Sub
            
            Call oNCred.dEliminaCredEvalParametrosEspecializacion
            For i = 1 To feEspecializacion.Rows - 1
                If i < 10 Then
                    nId = "000" & CStr(i)
                ElseIf i < 100 Then
                    nId = "00" & CStr(i)
                ElseIf i < 1000 Then
                    nId = "0" & CStr(i)
                Else
                    nId = CStr(i)
                End If
                nId = "ESP" & nId
                Call oNCred.dInsertaCredEvalParametrosEspecializacion(nId, CInt(Trim(Right(Trim(feEspecializacion.TextMatrix(i, 1)), 3))), _
                                            Trim(Right(Trim(feEspecializacion.TextMatrix(i, 2)), 3)), Trim(Right(Trim(feEspecializacion.TextMatrix(i, 3)), 3)), _
                                            Trim(Right(Trim(feEspecializacion.TextMatrix(i, 4)), 3)), Trim(Left(Trim(Mid(feEspecializacion.TextMatrix(i, 5), 70, 300)), 3)), _
                                            feEspecializacion.TextMatrix(i, 6), feEspecializacion.TextMatrix(i, 7), feEspecializacion.TextMatrix(i, 8))
            Next i
            feEspecializacion.lbEditarFlex = False
            cmdAgregar.Visible = False
            cmdEliminar.Visible = False
            cmdEditar.Visible = True
            MsgBox "Los parámetros se grabaron con exito", vbInformation, "Aviso"
        Else
            MsgBox "Faltan datos en la lista de parametros", vbInformation, "Aviso"
        End If
    End If
End Sub

Private Sub cmdSalir_Click()
If cmdEditar.Visible Then
    Unload Me
Else
    MsgBox "Primero Graba los datos Antes de Salir", vbInformation, "Aviso"
    cmdGrabar.SetFocus
End If
End Sub


Private Sub feEspecializacion_OnCellChange(pnRow As Long, pnCol As Long)
If pnCol = 2 Or pnCol = 4 Then
    feEspecializacion.TextMatrix(pnRow, pnCol + 1) = ""
End If
End Sub

Private Sub feEspecializacion_RowColChange()
Dim oDCred As COMDCredito.DCOMCredito
Dim oCons As COMDConstantes.DCOMConstantes
   If feEspecializacion.lbEditarFlex Then
        Set oDCred = New COMDCredito.DCOMCredito
        Set oCons = New COMDConstantes.DCOMConstantes
        Select Case feEspecializacion.Col
            Case 1 'Especializacion
                feEspecializacion.CargaCombo oCons.RecuperaConstantes(7052)
            Case 2 'Producto
                feEspecializacion.CargaCombo oDCred.RecuperaProductosCrediticios
            Case 3 'SubProducto
                feEspecializacion.CargaCombo oDCred.RecuperaSubProductosCrediticios(Trim(Right(feEspecializacion.TextMatrix(feEspecializacion.Row, 2), 3)))
            Case 4 'Tipo Credito
                feEspecializacion.CargaCombo oDCred.RecuperaTipoCreditos
            Case 5 'Sub Tipo Credito
                feEspecializacion.CargaCombo oDCred.RecuperaSubTipoCrediticios(Trim(Right(feEspecializacion.TextMatrix(feEspecializacion.Row, 4), 3)))
        End Select
        Set oDCred = Nothing
        Set oCons = Nothing
    End If
End Sub

Private Sub Form_Load()
    Dim oCred As COMDCredito.DCOMCredito
    Dim rsParam As ADODB.Recordset
    Dim i As Integer
    Set oCred = New COMDCredito.DCOMCredito
    
   Set rsParam = oCred.ListaCredEvalParamEspecializacion

     If Not (rsParam.BOF And rsParam.EOF) Then
        feEspecializacion.lbEditarFlex = True
        Call LimpiaFlex(feEspecializacion)
            For i = 0 To rsParam.RecordCount - 1
                feEspecializacion.AdicionaFila
                feEspecializacion.TextMatrix(i + 1, 0) = i + 1
                feEspecializacion.TextMatrix(i + 1, 1) = rsParam!cTpoEspDesc & Space(75) & rsParam!nTpoEsp
                feEspecializacion.TextMatrix(i + 1, 2) = rsParam!cTpoProdDesc & Space(75) & rsParam!cTpoProdCod
                feEspecializacion.TextMatrix(i + 1, 3) = rsParam!cTpoSubProdDesc & Space(75) & rsParam!cTpoSubProdCod
                feEspecializacion.TextMatrix(i + 1, 4) = rsParam!cTpoCredDesc & Space(75) & rsParam!cTpoCredCod
                feEspecializacion.TextMatrix(i + 1, 5) = rsParam!cTpoSubCredDesc & Space(75) & rsParam!cTpoSubCredCod
                feEspecializacion.TextMatrix(i + 1, 6) = rsParam!nUltEndeud
                feEspecializacion.TextMatrix(i + 1, 7) = Format(rsParam!nEndeudCredMin, "#,##0.00")
                feEspecializacion.TextMatrix(i + 1, 8) = Format(rsParam!nEndeudCredMax, "#,##0.00")
                rsParam.MoveNext
            Next i
    End If
    feEspecializacion.lbEditarFlex = False
End Sub

Public Function ValidaRegistroRepetido() As Boolean
    Dim i As Integer, j As Integer, UltRegistro As Integer
    ValidaRegistroRepetido = False
    UltRegistro = feEspecializacion.Rows - 1
    For i = 1 To feEspecializacion.Rows - 1
        If feEspecializacion.TextMatrix(i, 0) <> "" Then
            For j = 1 To feEspecializacion.Rows - 1
                If i <> j Then
                    If Trim(feEspecializacion.TextMatrix(i, 1)) = Trim(feEspecializacion.TextMatrix(j, 1)) And _
                    Trim(feEspecializacion.TextMatrix(i, 2)) = Trim(feEspecializacion.TextMatrix(j, 2)) And _
                    Trim(feEspecializacion.TextMatrix(i, 3)) = Trim(feEspecializacion.TextMatrix(j, 3)) And _
                    Trim(feEspecializacion.TextMatrix(i, 4)) = Trim(feEspecializacion.TextMatrix(j, 4)) And _
                    Trim(Left(Trim(Mid(feEspecializacion.TextMatrix(i, 5), 70, 300)), 3)) = Trim(Left(Trim(Mid(feEspecializacion.TextMatrix(j, 5), 70, 300)), 3)) And _
                    Trim(feEspecializacion.TextMatrix(i, 6)) = Trim(feEspecializacion.TextMatrix(j, 6)) And _
                    Trim(feEspecializacion.TextMatrix(i, 7)) = Trim(feEspecializacion.TextMatrix(j, 7)) And _
                    Trim(feEspecializacion.TextMatrix(i, 8)) = Trim(feEspecializacion.TextMatrix(j, 8)) Then
                        ValidaRegistroRepetido = False
                        MsgBox "El registro de la fila " & j & " ya existe en la fila " & i & ", favor de modificar", vbInformation, "Aviso"
                        Exit Function
                    End If
                End If
            Next j
        End If
    Next i
    ValidaRegistroRepetido = True
End Function

Public Function ValidaGrilla() As Boolean
    Dim i As Integer
    
    ValidaGrilla = False
    For i = 1 To feEspecializacion.Rows - 1
        If feEspecializacion.TextMatrix(i, 0) <> "" Then
            If Trim(feEspecializacion.TextMatrix(i, 1)) = "" Or Trim(feEspecializacion.TextMatrix(i, 2)) = "" Or _
               Trim(feEspecializacion.TextMatrix(i, 3)) = "" Or Trim(feEspecializacion.TextMatrix(i, 4)) = "" Or _
               Trim(feEspecializacion.TextMatrix(i, 5)) = "" Or Trim(feEspecializacion.TextMatrix(i, 6)) = "" Or _
               Trim(feEspecializacion.TextMatrix(i, 8)) = "" Or Trim(feEspecializacion.TextMatrix(i, 8)) = "" Then
                ValidaGrilla = False
                Exit Function
            End If
        End If
    Next i
    ValidaGrilla = True
End Function
