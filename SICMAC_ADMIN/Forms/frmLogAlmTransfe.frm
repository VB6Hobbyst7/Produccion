VERSION 5.00
Begin VB.Form frmLogAlmTransfe 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Almacen : Tranferencias"
   ClientHeight    =   5880
   ClientLeft      =   375
   ClientTop       =   1845
   ClientWidth     =   10110
   Icon            =   "frmLogAlmTransfe.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5880
   ScaleWidth      =   10110
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdAlm 
      Caption         =   "&Transferir"
      Enabled         =   0   'False
      Height          =   390
      Index           =   2
      Left            =   6090
      TabIndex        =   8
      Top             =   5370
      Width           =   1260
   End
   Begin VB.CommandButton cmdAlm 
      Caption         =   "&Cancelar"
      Enabled         =   0   'False
      Height          =   390
      Index           =   1
      Left            =   3795
      TabIndex        =   7
      Top             =   5370
      Width           =   1260
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   390
      Left            =   8325
      TabIndex        =   6
      Top             =   5355
      Width           =   1305
   End
   Begin VB.ComboBox cboAlmDes 
      Height          =   315
      Left            =   7200
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   165
      Width           =   1755
   End
   Begin VB.ComboBox cboAlmFue 
      Height          =   315
      Left            =   1875
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   165
      Width           =   1755
   End
   Begin Sicmact.FlexEdit fgeAlmFue 
      Height          =   4560
      Left            =   225
      TabIndex        =   4
      Top             =   660
      Width           =   5190
      _ExtentX        =   9155
      _ExtentY        =   8043
      Cols0           =   6
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad-Transferir"
      EncabezadosAnchos=   "400-0-1800-800-900-900"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X-5"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-L-R-R"
      FormatosEdit    =   "0-0-0-0-2-2"
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin Sicmact.FlexEdit fgeAlmDes 
      Height          =   4560
      Left            =   5580
      TabIndex        =   5
      Top             =   660
      Width           =   4275
      _ExtentX        =   7541
      _ExtentY        =   8043
      Cols0           =   5
      HighLight       =   1
      AllowUserResizing=   3
      EncabezadosNombres=   "Item-cBSCod-Bien/Servicio-Unidad-Cantidad"
      EncabezadosAnchos=   "400-0-1800-800-900"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnasAEditar =   "X-X-X-X-X"
      TextStyleFixed  =   3
      ListaControles  =   "0-0-0-0-0"
      EncabezadosAlineacion=   "C-L-L-L-R"
      FormatosEdit    =   "0-0-0-0-2"
      AvanceCeldas    =   1
      TextArray0      =   "Item"
      lbUltimaInstancia=   -1  'True
      lbFormatoCol    =   -1  'True
      lbBuscaDuplicadoText=   -1  'True
      Appearance      =   0
      ColWidth0       =   405
      RowHeight0      =   300
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Almacén Destino"
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
      Height          =   180
      Index           =   0
      Left            =   5655
      TabIndex        =   3
      Top             =   210
      Width           =   1500
   End
   Begin VB.Label lblEtiqueta 
      Caption         =   "Almacén Fuente"
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
      Height          =   180
      Index           =   5
      Left            =   330
      TabIndex        =   1
      Top             =   210
      Width           =   1500
   End
End
Attribute VB_Name = "frmLogAlmTransfe"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cboAlmDes_Click()
    Dim Rs As ADODB.Recordset
    Dim clsDAlm As DLogAlmacen
    
    Set Rs = New ADODB.Recordset
    Set clsDAlm = New DLogAlmacen
    
    fgeAlmDes.Clear
    fgeAlmDes.FormaCabecera
    fgeAlmDes.Rows = 2
    
    'Detalle de bienes/servicios del almacen destino
    Set Rs = clsDAlm.CargaAlmacenBS(Val(Right(cboAlmDes.Text, 2)), 0, Date)
    If Rs.RecordCount > 0 Then
        Set fgeAlmDes.Recordset = Rs
        fgeAlmDes.lbEditarFlex = True
    End If
    
    Set clsDAlm = Nothing
    Set Rs = Nothing
End Sub

Private Sub cboAlmFue_Click()
    Dim Rs As ADODB.Recordset
    Dim clsDAlm As DLogAlmacen
    
    Set Rs = New ADODB.Recordset
    Set clsDAlm = New DLogAlmacen
    
    fgeAlmFue.Clear
    fgeAlmFue.FormaCabecera
    fgeAlmFue.Rows = 2
    fgeAlmDes.Clear
    fgeAlmDes.FormaCabecera
    fgeAlmDes.Rows = 2
    
    'Detalle de bienes/servicios del almacen seleccionado
    Set Rs = clsDAlm.CargaAlmacenBS(Val(Right(cboAlmFue.Text, 2)), False, Date)
    If Rs.RecordCount > 0 Then
        Set fgeAlmFue.Recordset = Rs
        fgeAlmFue.lbEditarFlex = True
    End If
    
    'Almacen Destino
    Set Rs = clsDAlm.CargaAlmacen(ATodosMenosUno, Val(Right(cboAlmFue.Text, 2)))
    If Rs.RecordCount > 0 Then
        CargaCombo Rs, cboAlmDes
    End If
    
    cmdAlm(1).Enabled = True
    cmdAlm(2).Enabled = True
    
    Set clsDAlm = Nothing
    Set Rs = Nothing
End Sub

Private Sub cmdAlm_Click(Index As Integer)
    Dim clsDAlm As DLogAlmacen
    Dim clsDGnral As DLogGeneral
    Dim clsDMov As DLogMov
    Dim sMovAlm As String, sBSCod As String, sActualiza As String
    Dim nMovAlm As Long
    Dim nCont As Integer, nItem As Integer, nResult As Integer
    Dim nCantid As Currency
    Dim nAlmFue As Integer, nAlmDes As Integer, nPosAlmDes As Integer
    
    If Index = 1 Then
        'CANCELAR
        If MsgBox("¿ Estás seguro de cancelar toda la operación ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            cmdAlm(1).Enabled = False
            cmdAlm(2).Enabled = False
            Call Limpiar
        End If
    ElseIf Index = 2 Then
        'GRABAR - TRANSFERIR
        'Validar
        nAlmFue = Val(Right(cboAlmFue.Text, 2))
        nAlmDes = Val(Right(cboAlmDes.Text, 2))
        nPosAlmDes = cboAlmDes.ListIndex
        If cboAlmDes.Text = "" Then
            MsgBox "Falta determinar el Almacén a recibir la Transferencia", vbInformation, " Aviso"
            Exit Sub
        End If
        nItem = 0
        For nCont = 1 To fgeAlmFue.Rows - 1
            If CCur(IIf(fgeAlmFue.TextMatrix(nCont, 5) = "", 0, fgeAlmFue.TextMatrix(nCont, 5))) > 0 Then
                nItem = nItem + 1
            End If
        Next
        If nItem = 0 Then
            MsgBox "Falta determinar que Item(s) van a trasladarse", vbInformation, " Aviso"
            Exit Sub
        End If
        
        'Trasferir
        If MsgBox("¿ Estás seguro de Transferir " & nItem & " item(s) ? ", vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            Set clsDAlm = New DLogAlmacen
            Set clsDGnral = New DLogGeneral
            Set clsDMov = New DLogMov
            
            sMovAlm = clsDGnral.GeneraMov(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)
            sActualiza = GeneraActualizacion(gdFecSis, gsCodCMAC, gsCodAge, gsCodUser)

            'Grabación de MOV -MOVREF
            clsDMov.InsertaMov sMovAlm, Trim(Str(gLogOpeAlmTramite)), "", 0
            nMovAlm = clsDMov.GetnMovNro(sMovAlm)
            
            'Traslado a Almacén Destino y Actualización Fuente
            For nCont = 1 To fgeAlmFue.Rows - 1
                If CCur(IIf(fgeAlmFue.TextMatrix(nCont, 5) = "", 0, fgeAlmFue.TextMatrix(nCont, 5))) > 0 Then
                    sBSCod = CCur(fgeAlmFue.TextMatrix(nCont, 1))
                    nCantid = CCur(fgeAlmFue.TextMatrix(nCont, 5))
                    'Inserta o Actualiza AlmacenBS Destino
                    If clsDAlm.ExisAlmacenBS(nAlmDes, sBSCod) Then
                        clsDMov.ActualizaAlmacenBS nAlmDes, sBSCod, nCantid, True
                    Else
                        clsDMov.InsertaAlmacenBS nAlmDes, sBSCod, nCantid
                    End If
                    'Actualiza AlmacenBS Fuente
                    clsDMov.ActualizaAlmacenBS nAlmFue, sBSCod, nCantid, False
                End If
            Next
            
            'Ejecuta todos los querys en una transacción
            'nResult = clsDMov.EjecutaBatch
            Set clsDMov = Nothing
            Set clsDGnral = Nothing
            Set clsDAlm = Nothing
            
            If nResult = 0 Then
                cmdAlm(1).Enabled = False
                cmdAlm(2).Enabled = False
                Call cboAlmFue_Click
                cboAlmDes.ListIndex = nPosAlmDes
                MsgBox "Transfencia concluida", vbInformation, " Aviso "
            Else
                MsgBox "Error al grabar la información", vbInformation, " Aviso "
            End If
        End If
    Else
        MsgBox "Tipo comando no reconocido", vbInformation, " Aviso "
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub fgeAlmFue_OnValidate(ByVal pnRow As Long, ByVal pnCol As Long, Cancel As Boolean)
If pnCol = 5 Then
    If CCur(fgeAlmFue.TextMatrix(pnRow, pnCol)) > CCur(fgeAlmFue.TextMatrix(pnRow, 4)) Then
        Cancel = False
    End If
End If
End Sub

Private Sub Form_Load()
    Dim Rs As ADODB.Recordset
    Dim clsDAlm As DLogAlmacen
    
    Call CentraForm(Me)
    
    Set Rs = New ADODB.Recordset
    Set clsDAlm = New DLogAlmacen
    
    Set Rs = clsDAlm.CargaAlmacen(ATodos)
    If Rs.RecordCount > 0 Then
        CargaCombo Rs, cboAlmFue
    End If
    
    Set clsDAlm = Nothing
    Set Rs = Nothing
End Sub

Private Sub Limpiar()
    cboAlmFue.ListIndex = -1
    cboAlmDes.Clear
    fgeAlmFue.Clear
    fgeAlmFue.FormaCabecera
    fgeAlmFue.Rows = 2
    fgeAlmDes.Clear
    fgeAlmDes.FormaCabecera
    fgeAlmDes.Rows = 2
End Sub
