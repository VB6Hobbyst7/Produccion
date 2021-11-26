VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmCredCrediPagoArchivoResultado 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Informe RCD - Carga Archivo de Error TXT"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10335
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   10335
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtRuta 
      Height          =   330
      Left            =   615
      TabIndex        =   5
      Top             =   60
      Width           =   4905
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "..."
      Height          =   330
      Left            =   5535
      TabIndex        =   4
      Top             =   60
      Width           =   405
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "&Transformar Archivo"
      Height          =   360
      Left            =   5160
      TabIndex        =   3
      Top             =   4500
      Width           =   1725
   End
   Begin VB.CommandButton cmdPasar 
      Caption         =   "Pasar al SICMACT"
      Height          =   360
      Left            =   6915
      TabIndex        =   2
      Top             =   4500
      Width           =   1830
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   360
      Left            =   8895
      TabIndex        =   0
      Top             =   4500
      Width           =   1320
   End
   Begin MSComctlLib.ProgressBar barra 
      Height          =   285
      Left            =   75
      TabIndex        =   1
      Top             =   4560
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   503
      _Version        =   393216
      Appearance      =   1
   End
   Begin MSComDlg.CommonDialog cmdlOpen 
      Left            =   6435
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView lstCampos 
      Height          =   3975
      Left            =   0
      TabIndex        =   6
      Top             =   480
      Width           =   10290
      _ExtentX        =   18150
      _ExtentY        =   7011
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   7
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Credito"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "Fec.Pago"
         Object.Width           =   2117
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   2
         Text            =   "Monto"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   3
         Text            =   "Mora"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   1
         SubItemIndex    =   4
         Text            =   "Monto Pagado"
         Object.Width           =   2293
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Sucursal"
         Object.Width           =   1764
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Referencia"
         Object.Width           =   2117
      EndProperty
   End
   Begin MSMask.MaskEdBox txtFechaArchivoTXT 
      Height          =   315
      Left            =   8835
      TabIndex        =   7
      Top             =   120
      Width           =   1230
      _ExtentX        =   2170
      _ExtentY        =   556
      _Version        =   393216
      MaxLength       =   10
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Ruta :"
      Height          =   195
      Left            =   75
      TabIndex        =   9
      Top             =   120
      Width           =   435
   End
   Begin VB.Label Label2 
      Caption         =   "Fecha Archivo TXT"
      Height          =   255
      Left            =   7215
      TabIndex        =   8
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmCredCrediPagoArchivoResultado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Dim lsArchivo As String

Private Sub cmdOpen_Click()
    ' Establecer CancelError a True
    cmdlOpen.CancelError = True
    On Error GoTo ErrHandler
    ' Establecer los indicadores
    cmdlOpen.Flags = cdlOFNHideReadOnly
    cmdlOpen.InitDir = App.path
    ' Establecer los filtros
    cmdlOpen.Filter = "Todos los archivos (*.*)|*.*|Archivos de texto" & _
    "(*.txt)|*.txt|Archivos por lotes (*.bat)|*.bat|Archivos 112 (*.112)|*.112"
    ' Especificar el filtro predeterminado
    cmdlOpen.FilterIndex = 2
    ' Presentar el cuadro de diálogo Abrir
    cmdlOpen.ShowOpen
    ' Presentar el nombre del archivo seleccionado
    txtRuta = cmdlOpen.FileName
    Exit Sub
    
ErrHandler:
    ' El usuario ha hecho clic en el botón Cancelar
    Exit Sub

End Sub

Private Sub cmdPasar_Click()
Dim I As Integer
Dim lsSQL As String
Dim lnTotal As Long
Dim ObjCredi As COMNCredito.NCOMCrediPago
Set ObjCredi = New COMNCredito.NCOMCrediPago
    If Me.lstCampos.ListItems.Count > 0 Then
        Call ObjCredi.DeleteColocCrediPagoArcResultado(CDate(txtFechaArchivoTXT))
        For I = 1 To lstCampos.ListItems.Count
            Call ObjCredi.InsertColocCrediPagoArcResultado(txtFechaArchivoTXT, lstCampos.ListItems(I), _
             lstCampos.ListItems(I).SubItems(1), lstCampos.ListItems(I).SubItems(2), lstCampos.ListItems(I).SubItems(3), _
             lstCampos.ListItems(I).SubItems(4), lstCampos.ListItems(I).SubItems(5), lstCampos.ListItems(I).SubItems(6))
            
            barra.value = Int(I / lstCampos.ListItems.Count * 100)
        Next
        MsgBox "Actualiza Terminada con exito", vbInformation, "Aviso"
    Else
        MsgBox "NO SE TIENEN DATOS PARA ", vbInformation, "Aviso"
    End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub


Private Sub cmdTrans_Click()
Dim lsLinea As String
Dim nItem As ListItem
Dim lsFechaArchivoTXT As String
Dim lnTotalRegis As Long, lnTotalPagad As Double
Dim lsCodCta As String, lsFecPago As String
Dim lnMonto As Double, lnMora As Double
Dim lnMontoPag As Double
Dim lsSucursal As String
Dim lsReferencia As String

'Dim loBase As COMConecta.DCOMConecta

If Len(Trim(txtRuta)) = 0 Then
    MsgBox "Ruta no válida", vbInformation, "Aviso"
    Exit Sub
End If
lsArchivo = Trim(txtRuta)
If Dir(lsArchivo) = "" Then Exit Sub

lstCampos.ListItems.Clear

Open lsArchivo For Input As #1   ' Abre el archivo.

'Set loBase = New COMConecta.DCOMConecta
'    loBase.AbreConexion
    Do While Not EOF(1)   ' Repite el bucle hasta el final del archivo.
        Input #1, lsLinea
        If Len(Trim(lsLinea)) > 0 Then ' Linea tiene datos
            If Mid(lsLinea, 1, 2) = "CC" Then  ' Cabecera
                lsFechaArchivoTXT = Mid(lsLinea, 55, 8) '15,8
                lnTotalRegis = Mid(lsLinea, 63, 9) '23,9
                lnTotalPagad = CDbl(Mid(lsLinea, 72, 13)) + CDbl(Mid(lsLinea, 85, 2)) / 100 '32,13--45,2
                Me.txtFechaArchivoTXT = Mid(lsFechaArchivoTXT, 7, 4) & "/" & Mid(lsFechaArchivoTXT, 5, 2) & "/" & Mid(lsFechaArchivoTXT, 1, 4)
            ElseIf Mid(lsLinea, 1, 2) = "DD" Then
                lsCodCta = Mid(Mid(lsLinea, 15, 18), 3, 15) '14,14,3,12
                lsFecPago = Mid(lsLinea, 102, 8) '58,8
                lnMonto = CDbl(Mid(lsLinea, 118, 13)) + CDbl(Mid(lsLinea, 131, 2)) / 100 '74,13-87,2
                lnMora = CDbl(Mid(lsLinea, 133, 13)) + CDbl(Mid(lsLinea, 146, 2)) / 100  '89-102
                lnMontoPag = CDbl(Mid(lsLinea, 148, 13)) + CDbl(Mid(lsLinea, Len(lsLinea) - 2, 2)) / 100 '104-117
                lsSucursal = Mid(lsLinea, 119, 6)
                lsReferencia = Mid(lsLinea, 131, 22)
                ' Cargo la lista
                Set nItem = Me.lstCampos.ListItems.Add(, , lsCodCta)
                nItem.SubItems(1) = Mid(lsFecPago, 7, 2) & "/" & Mid(lsFecPago, 5, 2) & "/" & Mid(lsFecPago, 1, 4)
                nItem.SubItems(2) = Format(lnMonto, "###0.00")
                nItem.SubItems(3) = Format(lnMora, "###0.00")
                nItem.SubItems(4) = Format(lnMontoPag, "###0.00")
                nItem.SubItems(5) = lsSucursal
                nItem.SubItems(6) = lsReferencia
                
            End If
            ''Tamano = LOF(1)  'Tamaño de la cadena
        End If

        DoEvents
    Loop
'Set loBase = Nothing
Close #1   ' Cierra el archivo.
'Set nItem = Me.lstCampos.ListItems.Add(, , lsCorrelativo)

End Sub

Private Sub Form_Load()

Dim lsAgConsolida As String
txtRuta.Text = App.path
Me.Icon = LoadPicture(App.path & gsRutaIcono)

End Sub
