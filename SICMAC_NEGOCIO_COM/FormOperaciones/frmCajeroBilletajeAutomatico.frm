VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmCajeroBilletajeAutomatico 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Form1"
   ClientHeight    =   5325
   ClientLeft      =   2625
   ClientTop       =   2790
   ClientWidth     =   8100
   Icon            =   "frmCajeroBilletajeAutomatico.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5325
   ScaleWidth      =   8100
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   840
      Left            =   105
      TabIndex        =   8
      Top             =   5355
      Width           =   7845
      Begin VB.OptionButton optOpcion 
         Caption         =   "EXTORNO BILLETAJE AUTOMATICO"
         Height          =   450
         Index           =   1
         Left            =   4080
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   225
         Width           =   3420
      End
      Begin VB.OptionButton optOpcion 
         Caption         =   "BILLETAJE AUTOMÁTICO"
         Height          =   480
         Index           =   0
         Left            =   360
         Style           =   1  'Graphical
         TabIndex        =   0
         Top             =   210
         Value           =   -1  'True
         Width           =   3420
      End
   End
   Begin VB.CommandButton cmdExtornar 
      Caption         =   "Extornar"
      Height          =   375
      Left            =   5175
      TabIndex        =   5
      Top             =   4860
      Width           =   1335
   End
   Begin VB.CommandButton cmdRegistrar 
      Caption         =   "Registrar"
      Height          =   375
      Left            =   5175
      TabIndex        =   4
      Top             =   4860
      Width           =   1335
   End
   Begin VB.Frame fraAgencias 
      Caption         =   "Agencias"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4695
      Left            =   105
      TabIndex        =   7
      Top             =   75
      Width           =   7890
      Begin VB.CommandButton cmdBuscar 
         Caption         =   "&Buscar"
         Height          =   375
         Left            =   6360
         TabIndex        =   2
         Top             =   225
         Width           =   1335
      End
      Begin MSComctlLib.ListView lvwAgencia 
         Height          =   3810
         Left            =   135
         TabIndex        =   3
         Top             =   720
         Width           =   7590
         _ExtentX        =   13388
         _ExtentY        =   6720
         View            =   3
         LabelEdit       =   1
         MultiSelect     =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Verdana"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Cod"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Agencia"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Registró?"
            Object.Width           =   5644
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6645
      TabIndex        =   6
      Top             =   4860
      Width           =   1335
   End
End
Attribute VB_Name = "frmCajeroBilletajeAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Sub Inicio(ByVal nOpcion As Integer)
    optOpcion(nOpcion - 1).value = True
    
    If nOpcion = 1 Then
        cmdExtornar.Visible = False
        cmdRegistrar.Visible = True
    ElseIf nOpcion = 2 Then
        cmdExtornar.Visible = True
        cmdRegistrar.Visible = False
    End If
    Me.Show 1
End Sub

Private Sub CmdBuscar_Click()
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim L As MSComctlLib.ListItem
Dim bValorRetorno As Integer
Dim i As Integer

For Each L In lvwAgencia.ListItems
    L.Checked = True
Next

Set oCajero = New COMNCajaGeneral.NCOMCajero
For Each L In lvwAgencia.ListItems
    If L.Checked Then
        bValorRetorno = oCajero.YaRegistroEfectivoAutomatico(L.Text, gdFecSis)
        L.SubItems(3) = bValorRetorno
        
        If bValorRetorno = 0 Then
            L.SubItems(2) = "Ningún tipo de registro"
        ElseIf bValorRetorno = 1 Then
            L.SubItems(2) = "Registro Automático ya Efectuado"
        ElseIf bValorRetorno = 2 Then
            L.SubItems(2) = "Cierre de Agencia Efectuado"
        ElseIf bValorRetorno = 3 Then
            L.SubItems(2) = "Algún usuario ya registro su efectivo"
        End If
        
        If optOpcion(0).value = True Then
            If bValorRetorno <> 0 Then
                L.ForeColor = vbBlue
                For i = 1 To 3
                    L.ListSubItems(i).ForeColor = vbBlue
                Next
                L.Checked = False
            Else
                L.ForeColor = vbBlack
                For i = 1 To 3
                    L.ListSubItems(i).ForeColor = vbBlack
                Next
            End If
        ElseIf optOpcion(1).value = True Then
            If bValorRetorno <> 1 Then
                L.ForeColor = vbBlue
                For i = 1 To 3
                    L.ListSubItems(i).ForeColor = vbBlue
                Next
                L.Checked = False
            Else
                L.ForeColor = vbBlack
                For i = 1 To 3
                    L.ListSubItems(i).ForeColor = vbBlack
                Next
            End If
        End If
    Else
        L.SubItems(2) = ""
        L.SubItems(3) = ""
        L.ForeColor = vbBlack
        For i = 1 To 3
            L.ListSubItems(i).ForeColor = vbBlack
        Next
    End If
Next
Set oCajero = Nothing
End Sub

Private Sub cmdExtornar_Click()
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim i As Integer
Dim L As MSComctlLib.ListItem
Dim nCant As Integer

nCant = 0

For Each L In lvwAgencia.ListItems
    If L.Checked Then
        If Len(L.ListSubItems(3).Text) > 0 Then
            nCant = nCant + 1
        End If
    End If
Next

If nCant = 0 Then
    MsgBox "Debe presionar buscar...", vbInformation, "Aviso"
    Exit Sub
End If

nCant = 0

For Each L In lvwAgencia.ListItems
    If L.Checked Then
        If Val(L.ListSubItems(3).Text) = 1 Then
            nCant = nCant + 1
        End If
    End If
Next

If nCant = 0 Then
    MsgBox "Debe seleccionar al menos una opción para extornar", vbInformation, "Aviso"
    Exit Sub
Else
    If MsgBox("Desea efectuar el extorno?", vbQuestion + vbYesNo, "Confirmación!!!") = vbNo Then
        Exit Sub
    End If
End If

Set oCajero = New COMNCajaGeneral.NCOMCajero
nCant = 0

For Each L In lvwAgencia.ListItems
    If L.Checked Then
        If Val(L.ListSubItems(3).Text) = 1 Then
            oCajero.RegistraExtornoAutomatico L.Text, gdFecSis, gsCodUser
            nCant = nCant + 1
            L.ForeColor = vbBlack
            For i = 1 To 3
                L.ListSubItems(i).ForeColor = vbBlack
            Next
            L.SubItems(2) = "Ningun tipo de registro"
            L.SubItems(3) = 0
            L.Checked = False
            
        End If
    End If
Next
If nCant > 0 Then
    MsgBox "Registros de Billetaje de " & nCant & " agencia(s) extornado(s) satisfactoriamente", vbInformation, "Aviso"
End If
End Sub

Private Sub cmdRegistrar_Click()
Dim oCajero As COMNCajaGeneral.NCOMCajero
Dim i As Integer
Dim L As MSComctlLib.ListItem
Dim nCant As Integer

nCant = 0

For Each L In lvwAgencia.ListItems
    If L.Checked Then
        If Len(L.ListSubItems(3).Text) > 0 Then
            nCant = nCant + 1
        End If
    End If
Next

If nCant = 0 Then
    MsgBox "Debe presionar buscar...", vbInformation, "Aviso"
    Exit Sub
End If

nCant = 0

For Each L In lvwAgencia.ListItems
    If L.Checked Then
        If Val(L.ListSubItems(3).Text) = 0 Then
            nCant = nCant + 1
        End If
    End If
Next

If nCant = 0 Then
    MsgBox "Debe seleccionar al menos una opción para extornar", vbInformation, "Aviso"
    Exit Sub
Else
    If MsgBox("Desea efectuar el registro?", vbQuestion + vbYesNo, "Confirmación!!!") = vbNo Then
        Exit Sub
    End If
End If

Set oCajero = New COMNCajaGeneral.NCOMCajero
nCant = 0

For Each L In lvwAgencia.ListItems
    If L.Checked Then
        If Val(L.ListSubItems(3).Text) = 0 Then
            oCajero.RegistraEfectivoAutomatico L.Text, gdFecSis, gsCodUser, L.ListSubItems(1).Text
            nCant = nCant + 1
            L.ForeColor = vbRed
            For i = 1 To 3
                L.ListSubItems(i).ForeColor = vbRed
            Next
            L.SubItems(2) = "Registro Automático ya Efectuado"
            L.SubItems(3) = 1
            L.Checked = False
        End If
    End If
Next

If nCant > 0 Then
    MsgBox "Registros de Billetajes de " & nCant & " agencia(s) efectuado(s) satisfactoriamente", vbInformation, "Aviso"
End If

End Sub

Private Sub CmdSalir_Click()
Unload Me
End Sub

 

 
Private Sub Form_Load()
Dim oGen As COMDConstSistema.DCOMGeneral
Dim rs As ADODB.Recordset
Dim L As MSComctlLib.ListItem

Set oGen = New COMDConstSistema.DCOMGeneral
Set rs = oGen.GetAgencias()
Set oGen = Nothing

Do While Not rs.EOF
    Set L = lvwAgencia.ListItems.Add(, , rs("cAgeCod"))
    L.SubItems(1) = rs("cAgeDescripcion")
    rs.MoveNext
Loop
rs.Close
Set rs = Nothing
Me.Caption = "Efectuar Registro Automático de Billetaje de Agencias"
Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub
 
Private Sub lvwAgencia_ItemCheck(ByVal Item As MSComctlLib.ListItem)
If optOpcion(0).value = True Then
    If Val(Item.SubItems(3)) <> 0 Then
        Item.Checked = False
    End If
ElseIf optOpcion(1).value = True Then
    If Val(Item.SubItems(3)) <> 1 Then
        Item.Checked = False
    End If
End If
End Sub




