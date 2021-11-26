VERSION 5.00
Begin VB.Form frmCapTarifarioGrupo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuracion Grupos de Agencia"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7050
   Icon            =   "frmCapTarifarioGrupoTasa.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton btnCerrar 
      Cancel          =   -1  'True
      Caption         =   "Cerrar"
      Height          =   435
      Left            =   5805
      TabIndex        =   3
      Top             =   5760
      Width           =   1110
   End
   Begin VB.CommandButton btnBitacora 
      Caption         =   "Bitacora"
      Height          =   435
      Left            =   4665
      TabIndex        =   2
      Top             =   5760
      Width           =   1110
   End
   Begin VB.CommandButton btnActualizar 
      Caption         =   "Actualizar"
      Height          =   435
      Left            =   3525
      TabIndex        =   1
      Top             =   5760
      Width           =   1110
   End
   Begin SICMACT.FlexEdit grdGrupos 
      Height          =   5595
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   6825
      _ExtentX        =   12039
      _ExtentY        =   9869
      Cols0           =   8
      HighLight       =   1
      AllowUserResizing=   3
      RowSizingMode   =   1
      EncabezadosNombres=   "nro-nIdGrupo-Agencias-Grupo Tasas-Grupo Comis.-Ult. Actualizacion-cAgeCod-tmp"
      EncabezadosAnchos=   "0-0-2400-1200-1200-1600-0-0"
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
      ColumnasAEditar =   "X-X-X-3-4-X-X-X"
      ListaControles  =   "0-0-0-3-3-0-0-0"
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      BackColorControl=   -2147483643
      EncabezadosAlineacion=   "C-C-L-C-C-C-C-C"
      FormatosEdit    =   "0-0-0-0-0-0-0-0"
      TextArray0      =   "nro"
      lbEditarFlex    =   -1  'True
      lbUltimaInstancia=   -1  'True
      RowHeight0      =   300
      ForeColorFixed  =   -2147483630
   End
End
Attribute VB_Name = "frmCapTarifarioGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'************************************************************************************************************
'* NOMBRE         : frmCapTarifarioGrupo
'* DESCRIPCION    : Proyecto - Tarifario Versionado - Creacion de Grupos de Agencias
'* CREACION       : RIRO, 20160420 10:00 AM
'************************************************************************************************************

Option Explicit

Dim oCon As COMNCaptaGenerales.NCOMCaptaDefinicion
Dim rsInicial As ADODB.Recordset

Private Sub btnActualizar_Click()
    Dim rs As ADODB.Recordset
    Dim bVerificacion As Boolean
    Set rs = ObtenerGruposActualizar
    bVerificacion = False
    
    If VerificaDuplicidadAgencias Then
        MsgBox "Existe duplicidad en las agencias, favor de comunicarse con el area de T.I.", vbInformation, "Aviso"
        Exit Sub
    End If
    If Not rs Is Nothing Then
        If Not rs.EOF And Not rs.BOF Then
            If rsInicial.RecordCount > 0 Then
                bVerificacion = True
            End If
        End If
    End If
    If bVerificacion Then
        bVerificacion = False
        If MsgBox("¿Desea actualizar los grupos?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
            Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion
            bVerificacion = oCon.ActualizaTarifarioGrupo(rs, gsCodUser, gsCodAge)
            If bVerificacion Then
                CargarControles
                MsgBox "Los grupos fueron actualizados correctamente", vbInformation, "Aviso"
            Else
                MsgBox "Se presentaron inconvenientes durante la actualización de los grupos, comunicarse con el Area de T.I.", vbInformation, "Aviso"
            End If
            Set oCon = Nothing
        End If
    Else
        MsgBox "No existen registros para la actualización", vbInformation, "Aviso"
    End If
End Sub
Private Function ObtenerGruposActualizar() As ADODB.Recordset
    Dim i As Integer
    Dim rsTmp As ADODB.Recordset
    
    Set rsTmp = New ADODB.Recordset
    
    rsTmp.Fields.Append "nIdGrupo", adInteger
    rsTmp.Fields.Append "cAgeCod", adVarChar, 2
    rsTmp.Fields.Append "Tasa", adVarChar, 1
    rsTmp.Fields.Append "Comision", adVarChar, 1
    rsTmp.CursorLocation = adUseClient
    rsTmp.CursorType = adOpenStatic
    rsTmp.Open
    
    If Not rsInicial Is Nothing Then
        If Not rsInicial.EOF And Not rsInicial.BOF Then
            If rsInicial.RecordCount > 0 Then
                Do While Not rsInicial.EOF And Not rsInicial.BOF
                    For i = 1 To grdGrupos.Rows - 1
                        If rsInicial!cAgeCod = grdGrupos.TextMatrix(i, 6) Then
                            If rsInicial!cGrupoTasa <> grdGrupos.TextMatrix(i, 3) Or rsInicial!cGrupoComi <> grdGrupos.TextMatrix(i, 4) Then
                                rsTmp.AddNew
                                rsTmp.Fields("nIdGrupo") = grdGrupos.TextMatrix(i, 1)
                                rsTmp.Fields("cAgeCod") = grdGrupos.TextMatrix(i, 6)
                                rsTmp.Fields("Tasa") = grdGrupos.TextMatrix(i, 3)
                                rsTmp.Fields("Comision") = grdGrupos.TextMatrix(i, 4)
                            End If
                            i = grdGrupos.Rows
                        End If
                    Next i
                    rsInicial.MoveNext
                Loop
                rsInicial.MoveFirst
                If Not rsTmp.EOF Then rsTmp.MoveFirst
            End If
        End If
    End If
    Set ObtenerGruposActualizar = rsTmp
End Function
Private Function VerificaDuplicidadAgencias() As Boolean
    Dim i As Integer, j As Integer
    Dim nRepite As Integer
    
    For i = 1 To grdGrupos.Rows - 1
        nRepite = 0
        For j = 1 To grdGrupos.Rows - 1
            If grdGrupos.TextMatrix(i, 6) = grdGrupos.TextMatrix(j, 6) Then 'comparando agencias
                nRepite = nRepite + 1
            End If
        Next j
        If nRepite > 1 Then
            VerificaDuplicidadAgencias = True
            Exit Function
        End If
    Next i
    VerificaDuplicidadAgencias = False
End Function

Private Sub btnBitacora_Click()
    frmCapTarifarioGroupBitacora.Show
End Sub

Private Sub btnCerrar_Click()
    If MsgBox("Desea salir del formulario de Grupos?", vbInformation + vbYesNo + vbDefaultButton1, "Aviso") = vbYes Then
        Unload Me
    End If
End Sub
Private Sub CargarControles()

Dim rsTmp As ADODB.Recordset
Set oCon = New COMNCaptaGenerales.NCOMCaptaDefinicion

Set rsInicial = Nothing

'cargando los grupos
Set rsTmp = oCon.ObtenerAgenciaGrupoTasaComision
CargarGrid rsTmp
Set oCon = Nothing

End Sub
Private Sub CargarGrid(ByVal prsGrupos As ADODB.Recordset)
Dim i As Integer

'Cargando el los combos de los grupos
Dim rsGrupos As New ADODB.Recordset
rsGrupos.Fields.Append "Grupo", adVarChar, 10
rsGrupos.CursorLocation = adUseClient
rsGrupos.CursorType = adOpenStatic
rsGrupos.Open
For i = 65 To 90
    rsGrupos.AddNew
    rsGrupos.Fields("Grupo") = Chr(i)
Next i
rsGrupos.MoveFirst
grdGrupos.CargaCombo rsGrupos

'Cargando los valores de las agencias, grupos, tasas
i = 1
If Not prsGrupos Is Nothing Then
    If Not prsGrupos.EOF And Not prsGrupos.BOF Then
        If prsGrupos.RecordCount > 0 Then
            Do While Not prsGrupos.EOF And Not prsGrupos.BOF
                Set rsInicial = prsGrupos.Clone
                grdGrupos.AdicionaFila
                grdGrupos.TextMatrix(i, 0) = i
                grdGrupos.TextMatrix(i, 1) = prsGrupos!nIdGrupo
                grdGrupos.TextMatrix(i, 2) = prsGrupos!cAgeDescripcion
                grdGrupos.TextMatrix(i, 3) = prsGrupos!cGrupoTasa
                grdGrupos.TextMatrix(i, 4) = prsGrupos!cGrupoComi
                grdGrupos.TextMatrix(i, 5) = prsGrupos!cFecha
                grdGrupos.TextMatrix(i, 6) = prsGrupos!cAgeCod
                prsGrupos.MoveNext
                i = i + 1
            Loop
        End If
    End If
End If
End Sub
  
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    'controlando el Ctrl + V
    If KeyCode = 86 And Shift = 2 Then
        KeyCode = 10
    End If
End Sub

Private Sub Form_Load()
CargarControles
End Sub
Private Sub grdGrupos_Click()
    If Len(Trim(grdGrupos.TextMatrix(1, 1))) > 0 Then
        If (grdGrupos.Col = 3 Or grdGrupos.Col = 4) And grdGrupos.row >= 1 Then
            SendKeys "{Enter}"
        End If
    End If
End Sub

Private Sub grdGrupos_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
    Dim sColumnas() As String
    sColumnas = Split(grdGrupos.ColumnasAEditar, "-")
    If sColumnas(pnCol) = "X" Then
        Cancel = False
        MsgBox "Esta celda no es editable", vbInformation, "Aviso"
        SendKeys "{Tab}", True
        Exit Sub
    End If
End Sub

Public Sub Inicia(ByVal psProd As String, Optional ByVal bTipo As Boolean)
    If bTipo = False Then
        frmCapTarifarioGrupo.Show
    Else
        grdGrupos.ColumnasAEditar = "X-X-X-X-X-X-X-X"
        btnActualizar.Visible = False
        frmCapTarifarioGrupo.Show
    End If
End Sub
