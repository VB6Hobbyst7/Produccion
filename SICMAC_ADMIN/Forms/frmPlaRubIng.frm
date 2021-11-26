VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPlaRubIng 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presupuesto: Rubros: "
   ClientHeight    =   1695
   ClientLeft      =   3150
   ClientTop       =   1995
   ClientWidth     =   6885
   Icon            =   "frmPlaRubIng.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1695
   ScaleWidth      =   6885
   ShowInTaskbar   =   0   'False
   Begin VB.Frame fraContene 
      Height          =   1080
      Left            =   210
      TabIndex        =   6
      Top             =   15
      Width           =   6435
      Begin VB.TextBox txtCodRub 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1335
         MaxLength       =   20
         TabIndex        =   0
         Top             =   240
         Width           =   2445
      End
      Begin VB.TextBox txtDescri 
         Height          =   300
         Left            =   1335
         MaxLength       =   250
         TabIndex        =   1
         Top             =   615
         Width           =   4890
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Código :"
         Height          =   255
         Index           =   0
         Left            =   195
         TabIndex        =   8
         Top             =   255
         Width           =   1140
      End
      Begin VB.Label lblEtiqueta 
         Caption         =   "Descripción :"
         Height          =   255
         Index           =   1
         Left            =   195
         TabIndex        =   7
         Top             =   660
         Width           =   1125
      End
   End
   Begin VB.CommandButton cmdEliminar 
      Caption         =   "&Eliminar"
      Height          =   390
      Left            =   5460
      TabIndex        =   5
      Top             =   2430
      Width           =   1020
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   390
      Left            =   5310
      TabIndex        =   3
      Top             =   1215
      Width           =   1110
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   390
      Left            =   4140
      TabIndex        =   2
      Top             =   1215
      Width           =   1110
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgCuenta 
      Height          =   1845
      Left            =   315
      TabIndex        =   4
      Top             =   1740
      Width           =   4995
      _ExtentX        =   8811
      _ExtentY        =   3254
      _Version        =   393216
      Cols            =   4
      BackColorBkg    =   -2147483643
      GridColor       =   -2147483638
      FocusRect       =   2
      HighLight       =   2
      FillStyle       =   1
      RowSizingMode   =   1
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   4
   End
End
Attribute VB_Name = "frmPlaRubIng"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Resul As Boolean
Dim pPresu As String, pAno As String, pCodMae As String, pTipo As String

Public Function Inicio(ByVal cAno As String, ByVal cPresu As String, _
    ByVal cMaestro As String, cTipo As String) As Boolean
pAno = cAno
pPresu = cPresu
pCodMae = cMaestro
pTipo = cTipo
Me.Show 1
Inicio = Resul
End Function

Private Sub cmdCancelar_Click()
Resul = False
Unload Me
End Sub

Private Sub cmdEliminar_Click()
    Dim tmpSql As String
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    If Len(fgCuenta.TextMatrix(fgCuenta.Row, 2)) > 0 Then
        If MsgBox("Estas seguro de eliminar la cuenta : " & Trim(fgCuenta.TextMatrix(fgCuenta.Row, 2)) & " ?", vbQuestion + vbOKCancel, "Aviso") = vbOk Then
            oPP.EliminaRubroCta pAno, pPresu, txtCodRub.Text, fgCuenta.TextMatrix(fgCuenta.Row, 2)
        End If
        Call CargaCuenta
    End If
End Sub

Private Sub CmdGrabar_Click()
    Dim tmpSql As String
    Dim N As Integer
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    On Error GoTo cmdGrabarErr
    
    txtDescri.Text = Trim(Replace(txtDescri.Text, "'", "", , , vbTextCompare))
    If Len(txtDescri.Text) > 0 Then
        If pTipo = "3" Then
            For N = 1 To fgCuenta.Row
                If fgCuenta.TextMatrix(N, 2) = txtDescri.Text Then
                    MsgBox "Cuenta ya se encuentra ingresada", vbInformation, " Aviso"
                    txtDescri.SetFocus
                    Exit Sub
                End If
            Next
        End If
        If MsgBox("Esta seguro de Grabar : " & txtDescri.Text, vbQuestion + vbYesNo, " Aviso ") = vbYes Then
            If pTipo = "1" Then
                Resul = oPP.AgregaRubro(pPresu, pAno, txtCodRub.Text, txtDescri.Text)
                Unload Me
            ElseIf pTipo = "2" Then
                Resul = oPP.ModificaRubro(pPresu, pAno, txtCodRub.Text, txtDescri.Text)
                Unload Me
            ElseIf pTipo = "3" Then
                oPP.AgregaRubroCta pAno, pPresu, txtCodRub.Text, txtDescri.Text
                txtDescri.Text = ""
                Call CargaCuenta
                Resul = True
    
            Else
                MsgBox "Tipo no especificado", vbInformation, " Aviso"
            End If
        Else
            MsgBox "Grabación Cancelada", vbInformation, " Aviso "
        End If
    Else
        MsgBox "Falta ingresar la Descripción", vbInformation, " Aviso "
    End If
    Exit Sub
cmdGrabarErr:
      MsgBox Err.Description, vbInformation, "¡Aviso!"
End Sub

Private Sub Form_Load()
    Dim tmpReg As ADODB.Recordset
    Set tmpReg = New ADODB.Recordset
    Dim tmpSql As String
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    If pTipo = "1" Then
        Me.Caption = Me.Caption & "Ingreso"
        Set tmpReg = oPP.GetPresupRubro(pAno, pPresu, pCodMae & "__")
        If (tmpReg.BOF Or tmpReg.EOF) Then
            txtCodRub.Text = pCodMae & "01"
        Else
            tmpReg.MoveLast
            txtCodRub.Text = pCodMae & FillNum(Str(Val(Right(tmpReg!cPresuRubCod, 2)) + 1), 2, "0")
        End If
        RSClose tmpReg
        txtCodRub.Enabled = True
    ElseIf pTipo = "2" Then
        Me.Caption = Me.Caption & "Modificación"
        Set tmpReg = oPP.GetPresupRubro(pAno, pPresu, pCodMae)
        If (tmpReg.BOF Or tmpReg.EOF) Then
        Else
            txtCodRub.Text = pCodMae
            txtDescri.Text = Trim(tmpReg!cPresuRubDescripcion)
        End If
        RSClose tmpReg
        
    ElseIf pTipo = "3" Then
        Me.Caption = "Ingreso de Cuentas Contables"
        txtCodRub.Text = pCodMae
        lblEtiqueta(0).Caption = "Código Rubro :"
        lblEtiqueta(1).Caption = "Cta. Contable :"
        Call CargaCuenta
        Me.Height = Me.Height + 2200
    Else
        MsgBox "Tipo no especificado", vbInformation, " Aviso"
        txtDescri.Enabled = False
        cmdGrabar.Enabled = False
    End If
    
    CentraForm Me
End Sub

Private Sub txtDescri_GotFocus()
    fEnfoque txtDescri
End Sub

Private Sub txtDescri_KeyPress(KeyAscii As Integer)
KeyAscii = Asc(UCase(Chr(KeyAscii)))
   If KeyAscii = 13 Then
       cmdGrabar.SetFocus
   Else
       If pTipo = "3" Then
           If InStr("0123456789M", Chr(KeyAscii)) = 0 Then
               KeyAscii = 0
               Beep
           End If
       End If
   End If
End Sub

Private Sub CargaCuenta()
    Dim tmpReg As ADODB.Recordset
    Set tmpReg = New ADODB.Recordset
    Dim tmpSql As String
    Dim oPP As DPresupuesto
    Set oPP = New DPresupuesto
    Dim x As Integer, N As Integer
    Call MSHFlex(fgCuenta, 3, "Item-Código-Cuenta", "450-1300-2500", "R-L-L")
    Set tmpReg = oPP.GetRubroCta(pAno, pPresu, pCodMae)
    If Not (tmpReg.BOF Or tmpReg.EOF) Then
        With tmpReg
            Do While Not .EOF
                x = x + 1
                AdicionaRow fgCuenta, x
                fgCuenta.Row = fgCuenta.Rows - 1
                fgCuenta.TextMatrix(x, 0) = x
                fgCuenta.TextMatrix(x, 1) = !cCodRub
                fgCuenta.TextMatrix(x, 2) = !cCtaCnt
                .MoveNext
            Loop
        End With
    End If
    tmpReg.Close
    Set tmpReg = Nothing
End Sub
