VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmColRecComisionSelecciona 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Tabla de Comisiones de Abogado"
   ClientHeight    =   3435
   ClientLeft      =   1875
   ClientTop       =   2910
   ClientWidth     =   5355
   Icon            =   "frmColRecComisionSelecciona.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3435
   ScaleWidth      =   5355
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   2850
      TabIndex        =   1
      Top             =   2970
      Width           =   1000
   End
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "&Aceptar"
      Default         =   -1  'True
      Height          =   375
      Left            =   1470
      TabIndex        =   0
      Top             =   2970
      Width           =   1000
   End
   Begin VB.Frame fraCuentas 
      Caption         =   "Comisión"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000002&
      Height          =   2445
      Left            =   90
      TabIndex        =   2
      Top             =   450
      Width           =   5160
      Begin MSComctlLib.ListView lstComisiones 
         Height          =   2115
         Left            =   75
         TabIndex        =   5
         Top             =   270
         Width           =   5010
         _ExtentX        =   8837
         _ExtentY        =   3731
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   7
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Nro"
            Object.Width           =   265
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Rango Inicial"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "Rango Final"
            Object.Width           =   1587
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Tipo Comision"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "Valor"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "CodComision"
            Object.Width           =   0
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Categ."
            Object.Width           =   1587
         EndProperty
      End
   End
   Begin VB.Label lblNombrePersona 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1620
      TabIndex        =   4
      Top             =   90
      Width           =   3630
   End
   Begin VB.Label Label1 
      Caption         =   "Estudio Juridico"
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
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   1395
   End
End
Attribute VB_Name = "frmColRecComisionSelecciona"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim clsComision As UColRecComisionSelecciona

Public Function Inicio(ByVal psPersCod As String, ByVal psPersNombre As String) As UColRecComisionSelecciona
    
    Me.lblNombrePersona.Caption = psPersNombre
    CargaListaComision (psPersCod)
    
    
    'If Not (prCtas.EOF And prCtas.EOF) Then
    '    Do While Not prCtas.EOF
    '        lstComisiones.AddItem prCtas("cCtaCod") & Space(2) & prCtas("cRelacion") & Space(2) & Trim(prCtas("cEstado"))
    '        prCtas.MoveNext
    '    Loop
    'Else
    '    MsgBox "Persona no posee Comisiones de Abogados ", vbInformation, "Aviso"
    'End If

    Me.Show 1
    Set Inicio = clsComision
    Set clsComision = Nothing
End Function

'20060429
'modificado para nueva estructura.
'Autor: Pedro Mucha
Function CargaListaComision(ByVal psCodAbogado As String) As Boolean
Dim loRegComision As COMNColocRec.NComColRecComision
Dim lrComis As Recordset
Dim litmX As ListItem
Dim lnItem As Integer

Set loRegComision = New COMNColocRec.NComColRecComision
    Set lrComis = loRegComision.nObtieneListaComisionAbogado(psCodAbogado, "")
        
        '20060505
        'modificado para soportar recorsets nulos
        'Autor : Pedro Mucha
        If lrComis Is Nothing Then
            MsgBox "Estudio Juridico no tiene comisiones asignadas ", vbInformation, "Aviso"
            'lstComision.ListItems.Clear
            Exit Function
        Else
            Do While Not lrComis.EOF
                lnItem = lnItem + 1
                Set litmX = lstComisiones.ListItems.Add(, , lnItem)
                    litmX.SubItems(1) = Format(lrComis!nRangIni, "#0.00")
                    litmX.SubItems(2) = Format(lrComis!nRangFin, "#0.00")
                    litmX.SubItems(3) = lrComis!nTipComis
                    litmX.SubItems(4) = Format(lrComis!nValor, "#0.00")
                    litmX.SubItems(5) = Format(lrComis!nComisionCod)
                    litmX.SubItems(6) = Format(lrComis!nCategoria)
                    
                lrComis.MoveNext
            Loop
        End If
    Set lrComis = Nothing
Set loRegComision = Nothing

End Function
Private Sub CmdAceptar_Click()
Dim lsAbogCod As String, lsAbogNom   As String
Dim lnComCod As Integer, lnComTipo As Integer, lnComVal As Double

Dim sCadena As String
Dim nPos As Integer

    If lstComisiones.ListItems.Count > 0 Then
        If lstComisiones.SelectedItem.SubItems(3) = "Moneda" Then
            lnComTipo = 1
        Else
            lnComTipo = 2
        End If
        lnComCod = Trim(lstComisiones.SelectedItem.SubItems(5))
        lnComVal = Trim(lstComisiones.SelectedItem.SubItems(4))

    End If
    
clsComision.CargaDatos lsAbogCod, lnComTipo, lnComCod, lnComVal
Unload Me
End Sub

Private Sub cmdcancelar_Click()
    clsComision.CargaDatos "", 0, 0, 0
    Unload Me
End Sub

Private Sub Form_Load()
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
Me.Caption = "Comisiones de Estudios Juridicos"
Set clsComision = New UColRecComisionSelecciona
End Sub
