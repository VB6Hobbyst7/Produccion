VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdmCredAutoMant 
   Caption         =   "Mantenimiento de Autorizaciones"
   ClientHeight    =   4410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4830
   LinkTopic       =   "Form2"
   ScaleHeight     =   4410
   ScaleWidth      =   4830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3480
      TabIndex        =   5
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtautorizacion 
         Height          =   285
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   3495
      End
      Begin MSComctlLib.ListView lvwNiveles 
         Height          =   3000
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   5292
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         HoverSelection  =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Reference Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Codigo"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Nombre"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "Autorizaciones"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   0
         TabIndex        =   4
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAdmCredAutoMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub cmdActualiza_Click()
Dim oNCredito As COMNCredito.NCOMCredito
Set oNCredito = New COMNCredito.NCOMCredito
Dim lbAprobado As Boolean
Dim pValor As Integer
Dim i As Integer
 If Me.txtautorizacion.Text <> "" Then
        If oNCredito.ValidaConstanteAdm("9013", Trim(Me.txtautorizacion.Text)) Then
          MsgBox "Autorizacion Duplicada, Verifique ", vbExclamation, "Aviso"
          Exit Sub
        End If
        
        Call oNCredito.InsertaConstanteAdm("9013", Trim(Me.txtautorizacion.Text))
        CargaDatos
        Call limpiaTexts
    Else
             
       If MsgBox("¿Desea Actualizar las Autorizaciones?.", vbInformation + vbYesNo, "Atención") = vbYes Then
            For i = 1 To lvwNiveles.ListItems.Count
                  pValor = CDate(lvwNiveles.ListItems.iTem(i).Text)
                  'pCodUser = lvwNiveles.ListItems.iTem(i).SubItems(1)
                  lbAprobado = IIf(lvwNiveles.ListItems.iTem(i).Checked, True, False)
                              
                  If lbAprobado = True Then
                      Call oNCredito.ActualizaConstanteAdm("9013", pValor, True)
                  Else
                      Call oNCredito.ActualizaConstanteAdm("9013", pValor, False)
                  End If
             Next
        CargaDatos
        End If
    End If
    Set oNCredito = Nothing
    
End Sub

Public Sub CargaDatos(Optional ByVal ind As Integer = 0)
Dim rs As ADODB.Recordset
Dim oNCredito As COMNCredito.NCOMCredito
Dim i As Integer
Dim lista As ListItem
Set rs = New ADODB.Recordset
Set oNCredito = New COMNCredito.NCOMCredito
Set rs = oNCredito.obtenerConstanteAdm("9013")
Set oNCredito = Nothing
    
    i = 1
    If Not (rs.EOF And rs.BOF) Then
       lvwNiveles.ListItems.Clear
       Do Until rs.EOF
         Set lista = lvwNiveles.ListItems.Add(, , rs!nConsValor)
         lvwNiveles.ListItems.iTem(i).Checked = IIf(rs!bEstado, True, False)
         lista.SubItems(1) = IIf(rs!cConsDescripcion = "", "", rs!cConsDescripcion)
'         lista.SubItems(2) = IIf(rs!cConsDescripcion = "", "", rs!cConsDescripcion)
         i = i + 1
         rs.MoveNext
       Loop
    Else
       MsgBox "No Existen Datos", vbInformation, "Aviso"
    End If
    rs.Close
    Set rs = Nothing
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
limpiaTexts
CargaDatos
End Sub

Sub limpiaTexts()
    Me.txtautorizacion.Text = ""
    Me.txtautorizacion.Enabled = True
End Sub
