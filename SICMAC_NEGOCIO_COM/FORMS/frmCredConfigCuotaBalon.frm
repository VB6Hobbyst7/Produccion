VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmCredConfigCuotaBalon 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Configuración Cuotas con periodo de gracia con pago de intereses"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12645
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredConfigCuotaBalon.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4590
   ScaleWidth      =   12645
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSalir 
      Caption         =   "Salir"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   10800
      TabIndex        =   7
      Top             =   4080
      Width           =   1170
   End
   Begin VB.Frame fraTpoProductos 
      Caption         =   "Aplicable a:"
      ForeColor       =   &H8000000D&
      Height          =   3855
      Left            =   4080
      TabIndex        =   9
      Top             =   120
      Width           =   8415
      Begin VB.ListBox lstTpoProducto 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3210
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   480
         Width           =   4695
      End
      Begin VB.CheckBox chkTodosSubProd 
         Caption         =   "Todos"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   240
         Width           =   1215
      End
      Begin VB.CommandButton cmdGuardar 
         Caption         =   "Guardar"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   6720
         TabIndex        =   6
         Top             =   3360
         Width           =   1170
      End
      Begin VB.Frame fraRangos 
         Caption         =   "Rango Montos(Min/Max)"
         ForeColor       =   &H8000000D&
         Height          =   2175
         Left            =   4920
         TabIndex        =   10
         Top             =   480
         Width           =   3375
         Begin VB.TextBox txtMontoDolaresB 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   5
            Top             =   1680
            Width           =   1740
         End
         Begin VB.TextBox txtMontoDolaresA 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   4
            Top             =   1320
            Width           =   1740
         End
         Begin VB.TextBox txtMontoSolesB 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   3
            Top             =   720
            Width           =   1740
         End
         Begin VB.TextBox txtMontoSolesA 
            Alignment       =   1  'Right Justify
            Height          =   330
            Left            =   1440
            MaxLength       =   15
            TabIndex        =   2
            Top             =   360
            Width           =   1740
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Min:"
            Height          =   195
            Left            =   960
            TabIndex        =   17
            Top             =   1440
            Width           =   375
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Máx:"
            Height          =   195
            Left            =   960
            TabIndex        =   16
            Top             =   840
            Width           =   420
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Máx:"
            Height          =   195
            Left            =   960
            TabIndex        =   15
            Top             =   1800
            Width           =   420
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Min:"
            Height          =   195
            Left            =   960
            TabIndex        =   14
            Top             =   480
            Width           =   375
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Dolares:"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   720
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Soles:"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   480
            Width           =   540
         End
      End
   End
   Begin VB.Frame fraAgencia 
      Caption         =   "Agencia:"
      ForeColor       =   &H8000000D&
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   3855
      Begin MSComctlLib.ListView lstAgencia 
         Height          =   3465
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Width           =   3585
         _ExtentX        =   6324
         _ExtentY        =   6112
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Agencia"
            Object.Width           =   6174
         EndProperty
      End
   End
End
Attribute VB_Name = "frmCredConfigCuotaBalon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************
'***     Rutina                 :   frmCredConfigCuotaBalon
'***     Descripcion            :   Realiza la configuración de la Cuota Balon
'***     Creado por             :   WIOR
'***     Maquina                :   TIF-1-19
'***     Fecha-Tiempo           :   06/11/2013 05:56:29 PM
'***     Ultima Modificacion    :   Creacion de la Opcion
'*****************************************************************************************
Option Explicit
Private fsAgencia As String
Private Sub chkTodosSubProd_Click()
Call CheckLista(IIf(chkTodosSubProd.value = 1, True, False), lstTpoProducto)
End Sub

Private Sub CmdGuardar_Click()
If Not ValidaDatos Then Exit Sub

If MsgBox("Estas seguro de Guardar la configuracion?", vbInformation + vbYesNo, "Aviso") = vbYes Then
    Dim i As Integer
    Dim oCredito As COMNCredito.NCOMCredito
    Set oCredito = New COMNCredito.NCOMCredito
    
    Call oCredito.OperacionConfigCuotaBalonSubProd(2, fsAgencia)
    Call oCredito.OperacionConfigCuotaBalonRangosMontos(2, fsAgencia)
    
    Call oCredito.OperacionConfigCuotaBalonRangosMontos(1, fsAgencia, CDbl(txtMontoSolesA.Text), CDbl(txtMontoSolesB.Text), CDbl(txtMontoDolaresA.Text), CDbl(txtMontoDolaresB.Text))
    
    For i = 0 To lstTpoProducto.ListCount - 1
        If lstTpoProducto.Selected(i) = True Then
            Call oCredito.OperacionConfigCuotaBalonSubProd(1, fsAgencia, Trim(Left(lstTpoProducto.List(i), 3)))
        End If
    Next i
    
    MsgBox "Se grabaron correctamente los Datos", vbInformation, "Aviso"
    
    LimpiaDatos
End If
End Sub

Private Sub cmdsalir_Click()
Unload Me
End Sub
Private Sub CheckLista(ByVal bCheck As Boolean, ByVal lstLista As ListBox)
Dim i As Integer
For i = 0 To lstLista.ListCount - 1
    lstLista.Selected(i) = bCheck
Next i
End Sub

Private Sub LlenaListasSubProd(ByRef lista As ListBox, Optional psAgencia As String = "00")
    Dim oCredito As COMDCredito.DCOMCredito
    Set oCredito = New COMDCredito.DCOMCredito
    
    Dim rs As ADODB.Recordset
    Dim rsLista As ADODB.Recordset
    Dim i As Integer, J As Integer
    
    If Trim(psAgencia) = "00" Then
        Set rs = Nothing
    Else
        Set rs = oCredito.ObtenerConfigCuotaBalon(1, psAgencia)
    End If
    
    Call CheckLista(False, lista)
    Set rsLista = oCredito.RecuperaSubProductosCrediticios
    
    For i = 0 To rsLista.RecordCount - 1
        lista.AddItem rsLista!nConsValor & " " & Trim(rsLista!cConsDescripcion)
        If Trim(psAgencia) <> "00" Then
            If Not (rs.EOF And rs.BOF) Then
                rs.MoveFirst
                For J = 0 To rs.RecordCount - 1
                    If CStr(Trim(rsLista!nConsValor)) = Trim(rs!cTpoProdCod) Then
                        lista.Selected(i) = True
                    End If
                    rs.MoveNext
                Next J
            End If
        End If
        rsLista.MoveNext
    Next i
    Set oCredito = Nothing
End Sub

Private Sub Form_Load()
LimpiaDatos
End Sub

Private Sub LimpiaDatos()
fsAgencia = ""
LlenaListasSubProd lstTpoProducto
CargaAgencias
txtMontoSolesA.Text = "0.00"
txtMontoSolesB.Text = "0.00"
txtMontoDolaresA.Text = "0.00"
txtMontoDolaresB.Text = "0.00"
End Sub

Private Sub CargaAgencias()
Dim oConst As COMDConstantes.DCOMAgencias
Dim rsAgencias As ADODB.Recordset
Dim sOpeValor As String, sOpeDescrip As String
Dim L As ListItem

Set oConst = New COMDConstantes.DCOMAgencias
Set rsAgencias = oConst.ObtieneAgencias()

lstAgencia.ListItems.Clear
Do While Not rsAgencias.EOF
    sOpeValor = rsAgencias("nConsValor")
    sOpeDescrip = rsAgencias("cConsDescripcion")

    Set L = lstAgencia.ListItems.Add(, "p" & CStr(sOpeValor), sOpeDescrip)
    rsAgencias.MoveNext
Loop
End Sub


Private Sub lstAgencia_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim rsMontos As ADODB.Recordset

Set oCredito = New COMDCredito.DCOMCredito
fsAgencia = Mid(lstAgencia.SelectedItem.Key, 2, 3)

Set rsMontos = oCredito.ObtenerConfigCuotaBalon(2, fsAgencia)

If Not (rsMontos.EOF And rsMontos.BOF) Then
    txtMontoSolesA.Text = Format(rsMontos!nMontoASoles, "###," & String(15, "#") & "#0.00")
    txtMontoSolesB.Text = Format(rsMontos!nMontoBSoles, "###," & String(15, "#") & "#0.00")
    txtMontoDolaresA.Text = Format(rsMontos!nMontoADolares, "###," & String(15, "#") & "#0.00")
    txtMontoDolaresB.Text = Format(rsMontos!nMontoBDolares, "###," & String(15, "#") & "#0.00")
Else
    txtMontoSolesA.Text = "0.00"
    txtMontoSolesB.Text = "0.00"
    txtMontoDolaresA.Text = "0.00"
    txtMontoDolaresB.Text = "0.00"
End If

lstTpoProducto.Clear
Call LlenaListasSubProd(lstTpoProducto, fsAgencia)

End Sub

Private Sub txtMontoDolaresA_GotFocus()
fEnfoque txtMontoDolaresA
End Sub

Private Sub txtMontoDolaresA_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoDolaresA, KeyAscii)
If KeyAscii = 13 Then
    txtMontoDolaresB.SetFocus
End If
End Sub

Private Sub txtMontoDolaresA_LostFocus()
    If Len(Trim(txtMontoDolaresA.Text)) = 0 Or Trim(txtMontoDolaresA.Text) = "." Then
         txtMontoDolaresA.Text = "0.00"
    End If
    txtMontoDolaresA.Text = Format(txtMontoDolaresA.Text, "###," & String(15, "#") & "#0.00")
End Sub

Private Sub txtMontoDolaresB_GotFocus()
fEnfoque txtMontoDolaresB
End Sub

Private Sub txtMontoDolaresB_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoDolaresB, KeyAscii)
If KeyAscii = 13 Then
    cmdGuardar.SetFocus
End If
End Sub

Private Sub txtMontoDolaresB_LostFocus()
    If Len(Trim(txtMontoDolaresB.Text)) = 0 Or Trim(txtMontoDolaresB.Text) = "." Then
         txtMontoDolaresB.Text = "0.00"
    End If
    txtMontoDolaresB.Text = Format(txtMontoDolaresB.Text, "###," & String(15, "#") & "#0.00")
End Sub

Private Sub txtMontoSolesA_GotFocus()
fEnfoque txtMontoSolesA
End Sub

Private Sub txtMontoSolesA_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoSolesA, KeyAscii)
If KeyAscii = 13 Then
    txtMontoSolesB.SetFocus
End If
End Sub

Private Sub txtMontoSolesA_LostFocus()
    If Len(Trim(txtMontoSolesA.Text)) = 0 Or Trim(txtMontoSolesA.Text) = "." Then
         txtMontoSolesA.Text = "0.00"
    End If
    txtMontoSolesA.Text = Format(txtMontoSolesA.Text, "###," & String(15, "#") & "#0.00")
End Sub

Private Sub txtMontoSolesB_GotFocus()
fEnfoque txtMontoSolesB
End Sub

Private Sub txtMontoSolesB_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosDecimales(txtMontoSolesB, KeyAscii)
If KeyAscii = 13 Then
    txtMontoDolaresA.SetFocus
End If
End Sub

Private Sub txtMontoSolesB_LostFocus()
    If Len(Trim(txtMontoSolesB.Text)) = 0 Or Trim(txtMontoSolesB.Text) = "." Then
         txtMontoSolesB.Text = "0.00"
    End If
    txtMontoSolesB.Text = Format(txtMontoSolesB.Text, "###," & String(15, "#") & "#0.00")
End Sub

Private Function ValidaDatos() As Boolean
Dim i As Integer, Cont As Integer

ValidaDatos = True

If Trim(fsAgencia) = "" Then
    MsgBox "Selecciona una Agencia para realizar la Configuración", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

Cont = 0
For i = 0 To lstTpoProducto.ListCount - 1
    If lstTpoProducto.Selected(i) = True Then
        Cont = Cont + 1
    End If
Next i

If Cont = 0 Then
    MsgBox "Seleccione Por lo menos 1 Sub-Producto", vbInformation, "Aviso"
    ValidaDatos = False
    Exit Function
End If

If Trim(txtMontoSolesA.Text) = "" Then
    MsgBox "Ingrese Monto Incial en Soles", vbInformation, "Aviso"
    ValidaDatos = False
    txtMontoSolesA.SetFocus
    Exit Function
End If

If Trim(txtMontoSolesB.Text) = "" Or Trim(txtMontoSolesB.Text) = "0.00" Then
    MsgBox "Ingrese Monto Final en Soles", vbInformation, "Aviso"
    ValidaDatos = False
    txtMontoSolesB.SetFocus
    Exit Function
End If


If Trim(txtMontoDolaresA.Text) = "" Then
    MsgBox "Ingrese Monto Incial en Dolares", vbInformation, "Aviso"
    ValidaDatos = False
    txtMontoDolaresA.SetFocus
    Exit Function
End If


If Trim(txtMontoDolaresB.Text) = "" Or Trim(txtMontoDolaresB.Text) = "0.00" Then
    MsgBox "Ingrese Monto Final en Dolares", vbInformation, "Aviso"
    ValidaDatos = False
    txtMontoDolaresB.SetFocus
    Exit Function
End If

If Not (Trim(txtMontoSolesA.Text) = "") Then
    If Not (Trim(txtMontoSolesB.Text) = "" Or Trim(txtMontoSolesB.Text) = "0.00") Then
        If CDbl(Trim(txtMontoSolesA.Text)) > CDbl(Trim(txtMontoSolesB.Text)) Then
            MsgBox "El Monto en Inicial en Soles no puede ser Mayor al Monto Final", vbInformation, "Aviso"
            ValidaDatos = False
            txtMontoSolesA.SetFocus
            Exit Function
        End If
    End If
End If

If Not (Trim(txtMontoDolaresA.Text) = "") Then
    If Not (Trim(txtMontoDolaresB.Text) = "" Or Trim(txtMontoDolaresB.Text) = "0.00") Then
        If CDbl(Trim(txtMontoDolaresA.Text)) > CDbl(Trim(txtMontoDolaresB.Text)) Then
            MsgBox "El Monto en Inicial en Dolares no puede ser Mayor al Monto Final", vbInformation, "Aviso"
            ValidaDatos = False
            txtMontoDolaresA.SetFocus
            Exit Function
        End If
    End If
End If
End Function

