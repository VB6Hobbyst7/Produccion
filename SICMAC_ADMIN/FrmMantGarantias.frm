VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form FrmMantGarantias 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento Garantias"
   ClientHeight    =   6285
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6285
   ScaleWidth      =   8130
   StartUpPosition =   3  'Windows Default
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSH 
      Height          =   3165
      Left            =   0
      TabIndex        =   8
      Top             =   2310
      Width           =   8115
      _ExtentX        =   14314
      _ExtentY        =   5583
      _Version        =   393216
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Height          =   2235
      Left            =   0
      TabIndex        =   4
      Top             =   30
      Width           =   8115
      Begin MSComctlLib.ListView LstTpoGarantia 
         Height          =   1725
         Left            =   270
         TabIndex        =   7
         Top             =   420
         Width           =   5955
         _ExtentX        =   10504
         _ExtentY        =   3043
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Descripcion"
            Object.Width           =   7937
         EndProperty
      End
      Begin VB.CommandButton CmdBuscar 
         Caption         =   "Buscar"
         Height          =   345
         Left            =   6540
         TabIndex        =   6
         Top             =   510
         Width           =   1305
      End
      Begin VB.Label Label1 
         Caption         =   "Tipo Garantia"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   270
         TabIndex        =   5
         Top             =   180
         Width           =   1305
      End
   End
   Begin VB.Frame Frame3 
      Height          =   735
      Left            =   -30
      TabIndex        =   0
      Top             =   5520
      Width           =   8085
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   6615
         TabIndex        =   3
         Top             =   240
         Width           =   1185
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   390
         Left            =   6615
         TabIndex        =   2
         Top             =   225
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   390
         Left            =   5400
         TabIndex        =   1
         Top             =   225
         Width           =   1185
      End
   End
End
Attribute VB_Name = "FrmMantGarantias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rsGarantDoc As ADODB.Recordset
        
Private Sub CmdAceptar_Click()
    Dim objMantGarant As DMantGarantia
    If MSH.Rows > 1 Then
        CargarDataAActualizar
        Set objMantGarant = New DMantGarantia
         If objMantGarant.ActualizarGarantDoc(rsGarantDoc, Mid(LstTpoGarantia.SelectedItem.Key, 2, Len(LstTpoGarantia.SelectedItem.Key) - 1)) = True Then
            MsgBox "Se proceso satisfactoriamente", vbInformation, "AVISO"
         Else
            MsgBox "Hubo error en el  proceso", vbInformation, "AVISO"
         End If
         Set objMantGarant = Nothing
         Set rsGarantDoc = Nothing
    Else
        MsgBox "Debe existir registros a actualizar", vbInformation, "AVISO"
    End If
End Sub
Sub CargarDataAActualizar()
    Dim i As Integer
    Dim j As Integer
    Dim rsX As ADODB.Recordset
    ConfigurarGarantDoc
    
    For i = 1 To MSH.Rows - 2
        rsGarantDoc.AddNew
        rsGarantDoc(0) = IIf(MSH.TextMatrix(i, 0) = "SI", 1, 0)
        rsGarantDoc(1) = Mid(LstTpoGarantia.SelectedItem.Key, 2, Len(LstTpoGarantia.SelectedItem.Key) - 1)
        rsGarantDoc(2) = MSH.TextMatrix(i, 1)
        rsGarantDoc.Update
    Next i
End Sub

Private Sub CmdBuscar_Click()
    If Not LstTpoGarantia.SelectedItem Is Nothing Then
       MSH.Clear
       ConfigurarMSH
       CargarDatos
    Else
        MsgBox "Debe seleccionar un tipo de garantia", vbInformation, "AVISO"
    End If
End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
    CargarLstTpoGarantia
    ConfigurarMSH
End Sub

Sub CargarDatos()
    Dim nTpoGarantia
    Dim oMantGarantia As DMantGarantia
    Dim rs As ADODB.Recordset
    Dim nTpoDocmento As Integer
    Dim cDesDocumento As String
    Dim nAsignacion As Integer
    On Error GoTo ErrHandler
         nTpoGarantia = Mid(LstTpoGarantia.SelectedItem.Key, 2, Len(LstTpoGarantia.SelectedItem.Key) - 1)
         Set oMantGarantia = New DMantGarantia
         Set rs = oMantGarantia.ConsultarTpoGarantia(nTpoGarantia)
         Do Until rs.EOF
              MSH.Rows = MSH.Rows + 1
                With MSH
                    .TextMatrix(.Rows - 2, 0) = IIf(rs!Asignacion = 1, "SI", "NO")
                    .Row = .Rows - 2
                    .Col = 0
                    .CellFontBold = True
                    
                    .TextMatrix(.Rows - 2, 1) = rs!nDocTpo
                    .TextMatrix(.Rows - 2, 2) = rs!cDocDesc
                End With
            rs.MoveNext
         Loop
    Exit Sub
ErrHandler:
    If Not rs Is Nothing Then Set rs = Nothing
    If Not oMantGarantia Is Nothing Then Set oMantGarantia = Nothing
    MsgBox "Error al cargar los datos", vbInformation, "AVISO"
End Sub


Sub CargarLstTpoGarantia()
    Dim objTpoGarantia As DMantGarantia
    Dim rs As ADODB.Recordset
    Dim iListItem As ListItem
    
        Set objTpoGarantia = New DMantGarantia
        Set rs = objTpoGarantia.CargarTpoGarantia
        Set objTpoGarantia = Nothing
        LstTpoGarantia.ListItems.Clear
        Do Until rs.EOF
               With LstTpoGarantia
                Set iListItem = .ListItems.Add(, "C" & CStr(rs!nConsValor), rs!nConsValor)
                iListItem.SubItems(1) = rs!cConsDescripcion
            End With
            rs.MoveNext
        Loop
End Sub

Sub ConfigurarGarantDoc()
    Set rsGarantDoc = New ADODB.Recordset
    With rsGarantDoc.Fields
        .Append "Asignacion", adInteger
        .Append "TpoGarantia", adInteger
        .Append "TpoDoc", adInteger
    End With
    rsGarantDoc.Open
End Sub

Sub ConfigurarMSH()
    With MSH
        .Cols = 3
        .Rows = 2
        MSH.TextMatrix(0, 0) = "ASIGNACION"
        MSH.TextMatrix(0, 1) = "NRO"
        MSH.TextMatrix(0, 2) = "DOCUMENTO"
        MSH.ColWidth(0) = 1300
        MSH.ColWidth(1) = 800
        MSH.ColWidth(2) = 5500
        
    End With
End Sub



Private Sub MSH_DblClick()
        Dim nRow As Integer
        On Error GoTo ErrHandler
        nRow = MSH.Row
    If MSH.Rows > 1 Then
            If MSH.TextMatrix(nRow, 0) <> "" And MSH.TextMatrix(nRow, 0) <> "" And MSH.TextMatrix(nRow, 0) <> "" Then
                If MSH.TextMatrix(nRow, 0) = "SI" Then
                    MSH.TextMatrix(nRow, 0) = "NO"
                Else
                    MSH.TextMatrix(nRow, 0) = "SI"
                End If
            Else
                MsgBox "Debe seleccionar una region con data", vbInformation, "AVISO"
            End If
    Else
        MsgBox "Debe existir data", vbInformation, "AVISO"
    End If
        Exit Sub
ErrHandler:
    MsgBox "Selecciono region equivocada", vbInformation, "AVISO"
End Sub
