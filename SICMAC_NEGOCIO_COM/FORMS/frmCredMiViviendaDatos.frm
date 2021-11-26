VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmCredMiViviendaDatos 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Datos del Crédito"
   ClientHeight    =   4110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5040
   Icon            =   "frmCredMiViviendaDatos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   5040
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdAceptar 
      Caption         =   "Aceptar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   3600
      Width           =   1335
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "Cancelar"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   1
      Top             =   3600
      Width           =   1335
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   5953
      _Version        =   393216
      Tabs            =   1
      TabHeight       =   520
      BackColor       =   -2147483643
      TabCaption(0)   =   "MIVIVIENDA"
      TabPicture(0)   =   "frmCredMiViviendaDatos.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblValorInmuebleDesc"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label4"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtCuotaInicial"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtGastoCierre"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtMontoVivienda"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "uspPeriodoLimite"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtBono"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtMontoFinal"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).ControlCount=   13
      Begin VB.TextBox txtMontoFinal 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   2640
         TabIndex        =   14
         Text            =   "0"
         Top             =   2280
         Width           =   1815
      End
      Begin VB.TextBox txtBono 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   285
         Left            =   2640
         TabIndex        =   13
         Text            =   "0"
         Top             =   1920
         Width           =   1815
      End
      Begin SICMACT.uSpinner uspPeriodoLimite 
         Height          =   375
         Left            =   2640
         TabIndex        =   11
         Top             =   2640
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   661
         Min             =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontName        =   "MS Sans Serif"
         FontSize        =   8.25
      End
      Begin SICMACT.EditMoney txtMontoVivienda 
         Height          =   255
         Left            =   2640
         TabIndex        =   9
         Top             =   840
         Width           =   1815
         _extentx        =   3201
         _extenty        =   450
         font            =   "frmCredMiViviendaDatos.frx":0326
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtGastoCierre 
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Top             =   1200
         Width           =   1815
         _extentx        =   3201
         _extenty        =   450
         font            =   "frmCredMiViviendaDatos.frx":0352
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin SICMACT.EditMoney txtCuotaInicial 
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Top             =   1560
         Width           =   1815
         _extentx        =   3201
         _extenty        =   450
         font            =   "frmCredMiViviendaDatos.frx":037E
         text            =   "0"
         enabled         =   -1  'True
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Gastos de cierre:"
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   1200
         Width           =   1200
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Periodo Limite Pérdida del Bono:"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   2640
         Width           =   2295
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "años"
         Height          =   195
         Left            =   3360
         TabIndex        =   7
         Top             =   2760
         Width           =   345
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Monto Final Crédito:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   2280
         Width           =   1410
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Bono:"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1920
         Width           =   420
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Cuota Inicial (Aporte y Otros):"
         Height          =   195
         Left            =   240
         TabIndex        =   4
         Top             =   1560
         Width           =   2055
      End
      Begin VB.Label lblValorInmuebleDesc 
         AutoSize        =   -1  'True
         Caption         =   "Valor Venta / Valor  de Inmueble:"
         Height          =   195
         Left            =   240
         TabIndex        =   3
         Top             =   840
         Width           =   2355
      End
   End
End
Attribute VB_Name = "frmCredMiViviendaDatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private fDCredito As COMDCredito.DCOMCredito
Private fsCtaCod As String
Private fArrDatos As Variant
Private fArrDatos_temporal As Variant
Private psSubTipoV As String
Private psDestinoV As String
Private pnBono As Currency
Private pbError As Boolean
Private pnMontoCuotas As Currency
Private fbEditar As Boolean

Private Sub cmdAceptar_Click()
    If Not ValidaDatos(1) Then
        EnfocaControl txtMontoVivienda
        Exit Sub
    End If
    If Not ValidaDatos(2) Then
        EnfocaControl txtGastoCierre
        Exit Sub
    End If
    If Not ValidaDatos(3) Then
        EnfocaControl uspPeriodoLimite
        Exit Sub
    End If
    If Not ValidaDatos(4) Then
        EnfocaControl txtCuotaInicial
        Exit Sub
    End If
    
    CalcularValores (5)

    If psSubTipoV = "854" Then
        fArrDatos(10) = 999
    Else
        fArrDatos(10) = CInt(uspPeriodoLimite.valor)
    End If
    fArrDatos(11) = CDbl(txtGastoCierre.Text)
    
    If Not pbError Then
        Unload Me
    End If
End Sub

Private Sub cmdCancelar_Click()
    fArrDatos = fArrDatos_temporal
    Unload Me
End Sub

Private Sub Form_Load()
    uspPeriodoLimite.valor = 5
    If IsArray(fArrDatos) Then
        If Trim(fArrDatos(0)) <> "" Then
            txtGastoCierre.Text = Format(CDbl(fArrDatos(11)), "###," & String(15, "#") & "#0.00")
            txtMontoVivienda.Text = Format(CDbl(fArrDatos(0)), "###," & String(15, "#") & "#0.00")
            txtCuotaInicial.Text = Format(CDbl(fArrDatos(1)), "###," & String(15, "#") & "#0.00")
            txtBono.Text = Format(CDbl(fArrDatos(2)), "###," & String(15, "#") & "#0.00")
            txtMontoFinal.Text = Format(CDbl(fArrDatos(3)), "###," & String(15, "#") & "#0.00")
            uspPeriodoLimite.valor = CInt(fArrDatos(10))
        End If
    End If
    
    txtMontoVivienda.Enabled = fbEditar
    txtGastoCierre.Enabled = fbEditar
    txtCuotaInicial.Enabled = fbEditar
    uspPeriodoLimite.Enabled = fbEditar
    cmdAceptar.Enabled = fbEditar
End Sub

Private Sub txtCuotaInicial_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CalcularValores (4)
        If Not pbError Then
            If psSubTipoV = "854" Then
                EnfocaControl cmdAceptar
            Else
                EnfocaControl uspPeriodoLimite
            End If
        End If
    End If
End Sub

Private Sub txtCuotaInicial_LostFocus()
    txtGastoCierre.Text = Format(txtGastoCierre.Text, "#,##0.00")
End Sub

Private Sub txtGastoCierre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CalcularValores (2)
        If Not pbError Then
            EnfocaControl txtCuotaInicial
        End If
    End If
End Sub

Private Sub txtGastoCierre_LostFocus()
    txtGastoCierre.Text = Format(txtGastoCierre.Text, "#,##0.00")
End Sub

Private Sub txtMontoVivienda_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        CalcularValores (1)
        If Not pbError Then
            EnfocaControl txtGastoCierre
        End If
    End If
End Sub

Private Function ValidarControles(ByVal pnTipo As Integer, ByVal pnCuotaInicial As Double, ByVal pnMontoVivienda As Double) As Boolean
    ValidarControles = False
    pnMontoCuotas = CDbl(pnCuotaInicial)
    
    If pnTipo = 1 Then
        If CDbl(txtCuotaInicial.Text) = 0 Then
            txtCuotaInicial.Text = Format(CDbl(pnMontoCuotas), "###," & String(15, "#") & "#0.00")
        Else
            If CDbl(txtCuotaInicial.Text) < pnMontoCuotas Then
                txtCuotaInicial.Text = pnMontoCuotas
            End If
        End If
        
        If CDbl(txtGastoCierre.Text) > 0 Then
            If CDbl(txtGastoCierre.Text) > (CDbl(txtMontoVivienda.Text) * 0.05) Then
                txtGastoCierre.Text = Format((CDbl(txtMontoVivienda.Text) * 0.05), "###," & String(15, "#") & "#0.00")
            End If
        End If
        
    ElseIf pnTipo = 2 Then
        If CDbl(txtGastoCierre.Text) > (CDbl(txtMontoVivienda.Text) * 0.05) Then
            MsgBox "El monto del Gastos de cierre debe ser menor o igual al 5% del monto Valor Venta/Valor.", vbInformation, "Aviso"
            ValidarControles = True
        End If
    ElseIf pnTipo = 4 Then
        If CDbl(txtCuotaInicial.Text) < pnMontoCuotas Then
            MsgBox "El monto de la Cuota Inicial (Aporte y Otros) debe ser mayor o igual a: " & CStr(pnMontoCuotas) & ".", vbInformation, "Aviso"
            ValidarControles = True
        ElseIf CDbl(txtCuotaInicial.Text) > CDbl(pnMontoVivienda) Then
            MsgBox "El monto de Cuota Inicial (Aporte y Otros) debe ser menor al monto Valor Venta/Valor.", vbInformation, "Aviso"
            ValidarControles = True
        End If
    ElseIf pnTipo = 5 Then
        If CDbl(txtGastoCierre.Text) > 0 Then
            If CDbl(txtGastoCierre.Text) > (CDbl(txtMontoVivienda.Text) * 0.05) Then
                MsgBox "El monto del Gastos de cierre debe ser menor o igual al 5% del monto Valor Venta/Valor.", vbInformation, "Aviso"
                ValidarControles = True
                EnfocaControl txtGastoCierre
            End If
        End If
        
        If CDbl(txtCuotaInicial.Text) < pnMontoCuotas Then
            MsgBox "El monto de la Cuota Inicial (Aporte y Otros) debe ser mayor o igual a: " & CStr(pnMontoCuotas) & ".", vbInformation, "Aviso"
            ValidarControles = True
            EnfocaControl txtCuotaInicial
        ElseIf CDbl(txtCuotaInicial.Text) > CDbl(pnMontoVivienda) Then
            MsgBox "El monto de Cuota Inicial (Aporte y Otros) debe ser menor al monto Valor Venta/Valor.", vbInformation, "Aviso"
            ValidarControles = True
            EnfocaControl txtCuotaInicial
        End If
    End If
End Function
Private Sub CalcularValores(ByVal pnTipo As Integer)
    cmdAceptar.Enabled = False
    If Not ValidaDatos(pnTipo) Then Exit Sub
    
    Set fDCredito = New COMDCredito.DCOMCredito
    Dim rs As ADODB.Recordset
    
    If psSubTipoV = "854" Then 'HIPOTECARIO TECHO PROPIO
        Set rs = fDCredito.ObtenerValoresNuevoMIVIVIENDA_NEW(Mid(fsCtaCod, 9, 1), gdFecSis, CDbl(txtMontoVivienda.Text), CDbl(txtCuotaInicial.Text), CStr(psSubTipoV), CInt(psDestinoV), CDbl(txtGastoCierre.Text))
        ReDim fArrDatos(12)
        If Not (rs.EOF And rs.BOF) Then
    
            If (CInt(rs!bRango) = 1) Then
                If (CInt(rs!nValida) = 1) Then
                    pbError = ValidarControles(pnTipo, rs!nCuotaInicial, CDbl(rs!nMontoVivienda))
                    txtBono.Text = Format(CDbl(rs!nBonoOtorgado), "###," & String(15, "#") & "#0.00")
                    Me.txtMontoFinal.Text = Format(CDbl(CDbl(txtMontoVivienda.Text) + CDbl(txtGastoCierre.Text) - (CDbl(txtCuotaInicial.Text) + CDbl(txtBono.Text))), "###," & String(15, "#") & "#0.00")
                    
                    fArrDatos(0) = CDbl(rs!nMontoVivienda)
                    fArrDatos(1) = CDbl(txtCuotaInicial.Text)
                    fArrDatos(2) = CDbl(txtBono.Text)
                    fArrDatos(3) = CDbl(CDbl(txtMontoVivienda.Text) + CDbl(txtGastoCierre.Text) - (CDbl(txtCuotaInicial.Text) + CDbl(rs!nBonoOtorgado)))
                    fArrDatos(4) = CDbl(rs!nUIT)
                    fArrDatos(5) = CLng(rs!nDesde)
                    fArrDatos(6) = CLng(rs!nHasta)
                    fArrDatos(7) = CDbl(rs!nBono)
                    fArrDatos(8) = CDbl(rs!nMinCredUIT)
                    fArrDatos(9) = CInt(rs!nValida)
                Else
                    MsgBox "El valor del crédito (" & IIf(Mid(fsCtaCod, 9, 1) = "1", "S/. ", "$. ") & Format(CDbl(rs!nMOntoCred), "###," & String(15, "#") & "#0.00") & ") tiene que se mayor o igual al " & CDbl(rs!nMinCredUIT) & " de la UIT(S/. " & Format(CDbl(rs!nUIT), "###," & String(15, "#") & "#0.00") & ").", vbInformation, "Aviso"
                    txtCuotaInicial.Text = "0.00"
                    txtBono.Text = "0.00"
                    txtMontoFinal.Text = "0.00"
                    pbError = True
                End If
            Else
                MsgBox "El monto de Valor Venta/Valor de Inmueble no se encuentra en ningún rango configurado.", vbInformation, "Aviso"
                pbError = True
            End If
        End If
    Else
        Set rs = fDCredito.ObtenerValoresNuevoMIVIVIENDA_NEW(Mid(fsCtaCod, 9, 1), gdFecSis, CDbl(txtMontoVivienda.Text), CDbl(txtCuotaInicial.Text), CStr(psSubTipoV), CInt(psDestinoV), CDbl(txtGastoCierre.Text))
        ReDim fArrDatos(12)
        If Not (rs.EOF And rs.BOF) Then
            If (CInt(rs!bRango) = 1) Then
                If (CInt(rs!nValida) = 1) Then
                    pbError = ValidarControles(pnTipo, rs!nCuotaInicial, CDbl(rs!nMontoVivienda))
                    txtBono.Text = Format(CDbl(rs!nBonoOtorgado), "###," & String(15, "#") & "#0.00")
                    Me.txtMontoFinal.Text = Format(CDbl(CDbl(txtMontoVivienda.Text) + CDbl(txtGastoCierre.Text) - (CDbl(txtCuotaInicial.Text) + CDbl(txtBono.Text))), "###," & String(15, "#") & "#0.00")
        
                    fArrDatos(0) = CDbl(rs!nMontoVivienda)
                    fArrDatos(1) = CDbl(txtCuotaInicial.Text)
                    fArrDatos(2) = CDbl(rs!nBonoOtorgado)
                    fArrDatos(3) = CDbl(CDbl(txtMontoVivienda.Text) + CDbl(txtGastoCierre.Text) - (CDbl(txtCuotaInicial.Text) + CDbl(txtBono.Text))) 'CDbl(rs!nMOntoCred)
                    fArrDatos(4) = CDbl(rs!nUIT)
                    fArrDatos(5) = CLng(rs!nDesde)
                    fArrDatos(6) = CLng(rs!nHasta)
                    fArrDatos(7) = CDbl(rs!nBono)
                    fArrDatos(8) = CDbl(rs!nMinCredUIT)
                    fArrDatos(9) = CInt(rs!nValida)
                Else
                    MsgBox "El valor del crédito (" & IIf(Mid(fsCtaCod, 9, 1) = "1", "S/. ", "$. ") & Format(CDbl(rs!nMOntoCred), "###," & String(15, "#") & "#0.00") & ") tiene que se mayor o igual al " & CDbl(rs!nMinCredUIT) & " de la UIT(S/. " & Format(CDbl(rs!nUIT), "###," & String(15, "#") & "#0.00") & ").", vbInformation, "Aviso"
                    txtCuotaInicial.Text = "0.00"
                    txtBono.Text = "0.00"
                    txtMontoFinal.Text = "0.00"
                    pbError = True
                End If
            Else
                MsgBox "El monto de Valor Venta/Valor de Inmueble no se encuentra en ningún rango configurado.", vbInformation, "Aviso"
                pbError = True
            End If
        End If
    End If
    
    cmdAceptar.Enabled = Not pbError
End Sub

Private Sub txtMontoVivienda_LostFocus()
    txtMontoVivienda.Text = Format(txtMontoVivienda.Text, "#,##0.00")
End Sub

Public Sub Inicio(ByVal psCtaCod As String, Optional ByRef pArrDatos As Variant, Optional psSubTipo As String = "", Optional psDestino As String = "", Optional pbEditar As Boolean = True)
    ReDim fArrDatos_temporal(11)
    fsCtaCod = psCtaCod
    fArrDatos = pArrDatos
    psSubTipoV = psSubTipo 'JGPA20201117
    psDestinoV = psDestino
    fbEditar = pbEditar
    
    If Not IsArray(fArrDatos) Then
        ReDim fArrDatos(11)
    End If
    fArrDatos_temporal = fArrDatos
    If psSubTipoV = "854" Then
        uspPeriodoLimite.Visible = False
        Me.Caption = "Datos del Crédito: Techo Propio"
        Label4.Visible = False
        Label5.Visible = False
    Else
        uspPeriodoLimite.Visible = True
        Me.Caption = "Datos del Crédito"
        Label4.Visible = True
        Label5.Visible = True
    End If
    
    Me.Show 1
    pArrDatos = fArrDatos
End Sub

Private Function ValidaDatos(ByVal pnTipo As Integer) As Boolean
    ValidaDatos = True
    
    If pnTipo = 1 Then
        If Not IsNumeric(txtMontoVivienda.Text) Then
            MsgBox "Ingrese el valor de la Vivienda correctamente.", vbInformation, "Aviso"
            ValidaDatos = False
        End If
        If CDbl(txtMontoVivienda.Text) <= 0 Then
            MsgBox "El valor de la Vivienda debe ser mayor a cero.", vbInformation, "Aviso"
            ValidaDatos = False
        End If
     ElseIf pnTipo = 2 Then
        If Not IsNumeric(txtGastoCierre.Text) Then
            MsgBox "Ingrese el valor del Gasto de cierre correctamente.", vbInformation, "Aviso"
            ValidaDatos = False
        End If
     ElseIf pnTipo = 3 And psSubTipoV <> "854" Then
        If Not IsNumeric(uspPeriodoLimite.valor) Then
            MsgBox "Ingrese el valor del periodo límite de pérdida de bono correctamente.", vbInformation, "Aviso"
            ValidaDatos = False
        End If
        If CDbl(uspPeriodoLimite.valor) <= 0 Then
            MsgBox "El valor del periodo límite de pérdida de bono debe ser mayor a cero.", vbInformation, "Aviso"
            ValidaDatos = False
        End If
    ElseIf pnTipo = 4 Then
        If Not IsNumeric(txtCuotaInicial.Text) Then
            MsgBox "Ingrese el valor de la Cuota Inicial correctamente.", vbInformation, "Aviso"
            ValidaDatos = False
        End If
    End If
End Function

Private Sub uspPeriodoLimite_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        EnfocaControl cmdAceptar
    End If
End Sub
