VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmAdmCredExoMant 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mantenimiento de Exoneraciones"
   ClientHeight    =   7515
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4875
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7515
   ScaleWidth      =   4875
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdsalir 
      Caption         =   "Salir"
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   7080
      Width           =   1215
   End
   Begin VB.Frame Frame2 
      Caption         =   "Tipos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   4
      Top             =   3840
      Width           =   4695
      Begin VB.TextBox txtexoneraciones1 
         Height          =   285
         Left            =   120
         TabIndex        =   8
         Top             =   600
         Width           =   3495
      End
      Begin VB.CommandButton cmdAnadir 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   3720
         TabIndex        =   7
         Top             =   240
         Width           =   855
      End
      Begin VB.ComboBox cbo9005 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   240
         Width           =   3495
      End
      Begin MSComctlLib.ListView lvwNiveles1 
         Height          =   2040
         Left            =   120
         TabIndex        =   10
         Top             =   960
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   3598
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
   End
   Begin VB.Frame Frame1 
      Height          =   3855
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   4695
      Begin VB.TextBox txtExoneraciones 
         Height          =   285
         Left            =   120
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
      Begin VB.CommandButton cmdActualiza 
         Caption         =   "Actualizar"
         Height          =   375
         Left            =   3720
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin MSComctlLib.ListView lvwNiveles 
         Height          =   3000
         Left            =   120
         TabIndex        =   9
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
         Caption         =   "Exoneraciones"
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
         TabIndex        =   1
         Top             =   0
         Width           =   1455
      End
   End
End
Attribute VB_Name = "frmAdmCredExoMant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Sub llenar_exo()
Dim rs As ADODB.Recordset
Dim oNCredito As COMNCredito.NCOMCredito

Set rs = New ADODB.Recordset
Set oNCredito = New COMNCredito.NCOMCredito
Set rs = oNCredito.obtenerConstanteAdm("9005")
Set oNCredito = Nothing

    If Not rs.EOF Or rs.BOF Then
        Call llenar_cbo(rs)
    End If
End Sub

Private Sub cbo9005_Click()
    Dim rs1 As ADODB.Recordset
    Dim oNCredito As COMNCredito.NCOMCredito
    Dim pValor As Integer
    Dim i As Integer
    Set rs1 = New ADODB.Recordset
    Set oNCredito = New COMNCredito.NCOMCredito
    Dim Lista As ListItem
    If Me.cbo9005.ListIndex <> -1 Then
        If CInt(Right(Me.cbo9005.Text, 3)) = 4 Then
            Set rs1 = oNCredito.obtenerConstanteAdm("9012")
           
            i = 1
            If Not (rs1.EOF And rs1.BOF) Then
               lvwNiveles1.ListItems.Clear
               Do Until rs1.EOF
                 Set Lista = lvwNiveles1.ListItems.Add(, , rs1!nConsValor)
                 lvwNiveles1.ListItems.iTem(i).Checked = IIf(rs1!bEstado, True, False)
                 Lista.SubItems(1) = IIf(rs1!cConsDescripcion = "", "", rs1!cConsDescripcion)
                 i = i + 1
                 rs1.MoveNext
               Loop
               
               Set rs1 = Nothing
            Else
               MsgBox "No Existen Datos", vbInformation, "Aviso"
            End If
            
            txtexoneraciones1.Visible = True
            txtexoneraciones1.Enabled = True
            cmdAnadir.Visible = True
        Else
            lvwNiveles1.ListItems.Clear
            txtexoneraciones1.Visible = False
            cmdAnadir.Visible = False
        End If
    End If
End Sub

Private Sub cmdActualiza_Click()
Dim oNCredito As COMNCredito.NCOMCredito
Set oNCredito = New COMNCredito.NCOMCredito
Dim lbAprobado As Boolean
Dim pValor As Integer
Dim i As Integer
 If Me.txtExoneraciones.Text <> "" Then
        If oNCredito.ValidaConstanteAdm("9005", Trim(Me.txtExoneraciones.Text)) Then
          MsgBox "Exoneracion Duplicada, Verifique ", vbExclamation, "Aviso"
          Exit Sub
        End If
        
        Call oNCredito.InsertaConstanteAdm("9005", Trim(Me.txtExoneraciones.Text))
        Call limpiaTexts
        
    Else
             
       If MsgBox("¿Desea Actualizar las Exoneraciones?.", vbInformation + vbYesNo, "Atención") = vbYes Then
            For i = 1 To lvwNiveles.ListItems.Count
                  pValor = lvwNiveles.ListItems.iTem(i).Text
                  'pCodUser = lvwNiveles.ListItems.iTem(i).SubItems(1)
                  lbAprobado = IIf(lvwNiveles.ListItems.iTem(i).Checked, True, False)
                              
                  If lbAprobado = True Then
                      Call oNCredito.ActualizaConstanteAdm("9005", pValor, True)
                  Else
                      Call oNCredito.ActualizaConstanteAdm("9005", pValor, False)
                  End If
             Next
        End If
    End If
    Set oNCredito = Nothing
    CargaDatos
End Sub

Private Sub cmdAnadir_Click()
Dim oNCredito As COMNCredito.NCOMCredito
Set oNCredito = New COMNCredito.NCOMCredito
Dim lbAprobado As Boolean
Dim pValor As Integer
Dim i As Integer
If Me.txtexoneraciones1.Text <> "" Then
        If oNCredito.ValidaConstanteAdm("9012", Trim(Me.txtexoneraciones1.Text)) Then
            MsgBox "Exoneracion Duplicada, Verifique ", vbExclamation, "Aviso"
            Exit Sub
        End If
                
        Call oNCredito.InsertaConstanteAdm("9012", Trim(Me.txtexoneraciones1.Text))
        Call CargaDatos1
        Call limpiaTexts
        
Else
        
       If MsgBox("¿Desea Actualizar las Exoneraciones?.", vbInformation + vbYesNo, "Atención") = vbYes Then
            For i = 1 To lvwNiveles1.ListItems.Count
                pValor = CInt(lvwNiveles1.ListItems.iTem(i).Text)
                'pValor = lvwNiveles.ListItems.iTem(i).SubItems(1)
                lbAprobado = IIf(lvwNiveles1.ListItems.iTem(i).Checked, True, False)
                
                If lbAprobado = True Then
                    Call oNCredito.ActualizaConstanteAdm("9005", pValor, True)
                Else
                    Call oNCredito.ActualizaConstanteAdm("9005", pValor, False)
                End If
            Next
            CargaDatos1
        End If
End If
    Set oNCredito = Nothing
    'llenar_exo
End Sub

Sub limpiaTexts()
    Me.txtExoneraciones.Text = ""
    Me.txtExoneraciones.Enabled = True
    Me.txtexoneraciones1.Text = ""
    Me.txtexoneraciones1.Enabled = False
    txtexoneraciones1.Visible = False
    cmdAnadir.Visible = False
    Me.cbo9005.ListIndex = -1
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub

Private Sub Form_Load()
CentraForm Me
Me.Icon = LoadPicture(App.Path & gsRutaIcono)
limpiaTexts
llenar_exo
CargaDatos
End Sub

Sub llenar_cbo(rs As ADODB.Recordset)
    Me.cbo9005.Enabled = True
    rs.MoveFirst
    Do While Not rs.EOF
        Me.cbo9005.AddItem Trim(rs!cConsDescripcion) & Space(100) & Trim(str(rs!nConsValor))
        rs.MoveNext
    Loop
End Sub

Public Sub CargaDatos(Optional ByVal ind As Integer = 0)
Dim rs As ADODB.Recordset
Dim oNCredito As COMNCredito.NCOMCredito
Dim i As Integer
Dim Lista As ListItem
Set rs = New ADODB.Recordset
Set oNCredito = New COMNCredito.NCOMCredito
Set rs = oNCredito.obtenerConstanteAdm("9005")
Set oNCredito = Nothing
    
    i = 1
    If Not (rs.EOF And rs.BOF) Then
       lvwNiveles.ListItems.Clear
       Do Until rs.EOF
         Set Lista = lvwNiveles.ListItems.Add(, , rs!nConsValor)
         lvwNiveles.ListItems.iTem(i).Checked = IIf(rs!bEstado, True, False)
         Lista.SubItems(1) = IIf(rs!cConsDescripcion = "", "", rs!cConsDescripcion)
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

Public Sub CargaDatos1(Optional ByVal ind As Integer = 0)
Dim rs As ADODB.Recordset
Dim oNCredito As COMNCredito.NCOMCredito
Dim i As Integer
Dim Lista As ListItem

Set rs = New ADODB.Recordset
Set oNCredito = New COMNCredito.NCOMCredito
Set rs = oNCredito.obtenerConstanteAdm("9012")
Set oNCredito = Nothing

    i = 1
    If Not (rs.EOF And rs.BOF) Then
       lvwNiveles1.ListItems.Clear
       Do Until rs.EOF
         Set Lista = lvwNiveles1.ListItems.Add(, , rs!nConsValor)
         lvwNiveles1.ListItems.iTem(i).Checked = IIf(rs!bEstado, True, False)
         Lista.SubItems(1) = IIf(rs!cConsDescripcion = "", "", rs!cConsDescripcion)
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

