VERSION 5.00
Begin VB.Form frmCredMntMetLiquid 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mantenimiento de Metodo de Liquidacion"
   ClientHeight    =   4560
   ClientLeft      =   2460
   ClientTop       =   2145
   ClientWidth     =   7245
   Icon            =   "frmCredMntMetLiquid.frx":0000
   LinkTopic       =   "Form3"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   7245
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   4470
      Left            =   15
      TabIndex        =   0
      Top             =   -15
      Width           =   7170
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   405
         Left            =   3765
         TabIndex        =   23
         Top             =   3855
         Width           =   1440
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Enabled         =   0   'False
         Height          =   405
         Left            =   2130
         TabIndex        =   22
         Top             =   3855
         Width           =   1440
      End
      Begin VB.Frame Frame3 
         Caption         =   "Credito"
         Height          =   1245
         Left            =   90
         TabIndex        =   17
         Top             =   165
         Width           =   6990
         Begin SICMACT.ActXCodCta ActxCta 
            Height          =   435
            Left            =   195
            TabIndex        =   21
            Top             =   495
            Width           =   3660
            _ExtentX        =   6456
            _ExtentY        =   767
            Texto           =   "Credito :"
            EnabledCMAC     =   -1  'True
            EnabledCta      =   -1  'True
            EnabledProd     =   -1  'True
            EnabledAge      =   -1  'True
         End
         Begin VB.CommandButton CmdBuscar 
            Caption         =   "&Buscar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4035
            TabIndex        =   20
            Top             =   525
            Width           =   900
         End
         Begin VB.Frame FraListaCred 
            Caption         =   "&Lista Creditos"
            Height          =   960
            Left            =   5025
            TabIndex        =   18
            Top             =   180
            Width           =   1875
            Begin VB.ListBox LstCred 
               Height          =   645
               ItemData        =   "frmCredMntMetLiquid.frx":030A
               Left            =   75
               List            =   "frmCredMntMetLiquid.frx":030C
               TabIndex        =   19
               Top             =   225
               Width           =   1725
            End
         End
      End
      Begin VB.Frame Frame2 
         Height          =   1005
         Left            =   90
         TabIndex        =   8
         Top             =   1410
         Width           =   7005
         Begin VB.Label LblMonCred 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1635
            TabIndex        =   16
            Top             =   255
            Width           =   915
         End
         Begin VB.Label lblmonto 
            AutoSize        =   -1  'True
            Caption         =   "&Prestamo : "
            Height          =   195
            Left            =   570
            TabIndex        =   15
            Top             =   285
            Width           =   795
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Saldo Capital"
            Height          =   195
            Left            =   555
            TabIndex        =   14
            Top             =   615
            Width           =   930
         End
         Begin VB.Label LblSalCap 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   1635
            TabIndex        =   13
            Top             =   585
            Width           =   915
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Metodo de Liquidacion"
            Height          =   195
            Left            =   3390
            TabIndex        =   12
            Top             =   285
            Width           =   1620
         End
         Begin VB.Label LblMetLiq 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5220
            TabIndex        =   11
            Top             =   255
            Width           =   1020
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "Interes del Credito"
            Height          =   195
            Left            =   3405
            TabIndex        =   10
            Top             =   615
            Width           =   1275
         End
         Begin VB.Label LblIntCred 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000E&
            BorderStyle     =   1  'Fixed Single
            Height          =   255
            Left            =   5220
            TabIndex        =   9
            Top             =   585
            Width           =   1020
         End
      End
      Begin VB.Frame fraNuevoMet 
         Enabled         =   0   'False
         Height          =   1035
         Left            =   3690
         TabIndex        =   5
         Top             =   2655
         Width           =   3405
         Begin VB.TextBox txtMetLiq 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   1800
            MaxLength       =   4
            TabIndex        =   6
            Top             =   360
            Width           =   780
         End
         Begin VB.Label Label9 
            Caption         =   "Nuevo Metodo de Liquidacion"
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
            Height          =   480
            Left            =   270
            TabIndex        =   7
            Top             =   270
            Width           =   1425
         End
      End
      Begin VB.Frame fraTipoMet 
         Enabled         =   0   'False
         Height          =   1035
         Left            =   90
         TabIndex        =   1
         Top             =   2655
         Width           =   3555
         Begin VB.OptionButton OptMetLiq 
            Caption         =   "Met. Cuota Adelantda"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   4
            Top             =   720
            Width           =   2085
         End
         Begin VB.OptionButton OptMetLiq 
            Caption         =   "Met. Refinanciado"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   3
            Top             =   450
            Width           =   1905
         End
         Begin VB.OptionButton OptMetLiq 
            Caption         =   "Met. Configurable"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   2
            Top             =   180
            Value           =   -1  'True
            Width           =   1995
         End
      End
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   405
         Left            =   3765
         TabIndex        =   24
         Top             =   3855
         Visible         =   0   'False
         Width           =   1440
      End
   End
End
Attribute VB_Name = "frmCredMntMetLiquid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub LimpiaPantalla()
    Call LimpiaControles(Me)
    ActxCta.NroCuenta = ""
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
End Sub

Private Function ValidaDatos() As Boolean
Dim i As Integer
Dim J As Integer
Dim Cont As Integer

    ValidaDatos = True
    If Trim(txtMetLiq.Text) = "" Then
        MsgBox "Ingrese El Metodo de Liquidacion", vbInformation, "Aviso"
        If txtMetLiq.Enabled Then
            txtMetLiq.SetFocus
        End If
        ValidaDatos = False
        Exit Function
    End If
    
    If InStr("GMICiY", Mid(txtMetLiq.Text, 1, 1)) <= 0 Then
        MsgBox "La Primera letra del metodo de Liquidacion No es Valida, " & Chr(10) & " Solo Puede Ingresar Una de estas Letras GMIYiC", vbInformation, "Aviso"
        If txtMetLiq.Enabled Then
            txtMetLiq.SetFocus
        End If
        ValidaDatos = False
        Exit Function
    End If
    
    If InStr("GMICiY", Mid(txtMetLiq.Text, 2, 1)) <= 0 Then
        MsgBox "La Segunda letra del metodo de Liquidacion No es Valida, " & Chr(10) & " Solo Puede Ingresar Una de estas Letras GMIYiC", vbInformation, "Aviso"
        If txtMetLiq.Enabled Then
            txtMetLiq.SetFocus
        End If
        ValidaDatos = False
        Exit Function
    End If
    
    If InStr("GMICiY", Mid(txtMetLiq.Text, 3, 1)) <= 0 Then
        MsgBox "La Tercera letra del metodo de Liquidacion No es Valida, " & Chr(10) & " Solo Puede Ingresar Una de estas Letras GMIYiC", vbInformation, "Aviso"
        If txtMetLiq.Enabled Then
            txtMetLiq.SetFocus
        End If
        ValidaDatos = False
        Exit Function
    End If
    
    If InStr("GMICiY", Mid(txtMetLiq.Text, 4, 1)) <= 0 Then
        MsgBox "La Cuarta letra del metodo de Liquidacion No es Valida, " & Chr(10) & " Solo Puede Ingresar Una de estas Letras GMIYiC", vbInformation, "Aviso"
        If txtMetLiq.Enabled Then
            txtMetLiq.SetFocus
        End If
        ValidaDatos = False
        Exit Function
    End If
    
    'Valida Duplicidad
    Cont = 0
    txtMetLiq.Text = Trim(txtMetLiq.Text)
    For i = 1 To Len(txtMetLiq.Text)
        If Mid(txtMetLiq.Text, i, 1) = "G" Then
            Cont = Cont + 1
        End If
    Next i
    If Cont > 1 Then
        MsgBox "No Puede haber Letras Duplicadas en el Metodo de Liquidacion", vbInformation, "Aviso"
        If txtMetLiq.Enabled Then
            txtMetLiq.SetFocus
        End If
        ValidaDatos = False
        Exit Function
    End If
    Cont = 0
    For i = 1 To Len(txtMetLiq.Text)
        If Mid(txtMetLiq.Text, i, 1) = "I" Or Mid(txtMetLiq.Text, i, 1) = "Y" Or Mid(txtMetLiq.Text, i, 1) = "i" Then
            Cont = Cont + 1
        End If
    Next i
    If Cont > 1 Then
        MsgBox "No Puede haber Letras Duplicadas en el Metodo de Liquidacion", vbInformation, "Aviso"
        If txtMetLiq.Enabled Then
            txtMetLiq.SetFocus
        End If
        ValidaDatos = False
        Exit Function
    End If
    
    Cont = 0
    For i = 1 To Len(txtMetLiq.Text)
        If Mid(txtMetLiq.Text, i, 1) = "M" Then
            Cont = Cont + 1
        End If
    Next i
    If Cont > 1 Then
        MsgBox "No Puede haber Letras Duplicadas en el Metodo de Liquidacion", vbInformation, "Aviso"
        If txtMetLiq.Enabled Then
            txtMetLiq.SetFocus
        End If
        ValidaDatos = False
        Exit Function
    End If
    
    Cont = 0
    For i = 1 To Len(txtMetLiq.Text)
        If Mid(txtMetLiq.Text, i, 1) = "C" Then
            Cont = Cont + 1
        End If
    Next i
    If Cont > 1 Then
        MsgBox "No Puede haber Letras Duplicadas en el Metodo de Liquidacion", vbInformation, "Aviso"
        If txtMetLiq.Enabled Then
            txtMetLiq.SetFocus
        End If
        ValidaDatos = False
        Exit Function
    End If
End Function

Private Sub HabilitaActualizacion(ByVal pbHabilita As Boolean)
    Frame3.Enabled = Not pbHabilita
    fraTipoMet.Enabled = pbHabilita
    fraNuevoMet.Enabled = pbHabilita
    cmdAceptar.Enabled = pbHabilita
    CmdCancelar.Visible = pbHabilita
    CmdSalir.Visible = Not pbHabilita
End Sub


Private Function CargaDatos() As Boolean
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset

    On Error GoTo ErrorCargaDatos
    Set oCredito = New COMDCredito.DCOMCredito
    Set R = oCredito.RecuperaDatosComunes(ActxCta.NroCuenta)
    Set oCredito = Nothing
    
    If Not R.BOF And Not R.EOF Then
        CargaDatos = True
        LblMonCred.Caption = Format(R!nMontoCol, "#0.00")
        lblMetLiq.Caption = R!cMetLiquidacion
        LblSalCap.Caption = Format(R!nSaldo, "#0.00")
        LblIntCred.Caption = Format(R!nTasaInteres, "#0.00")
        txtMetLiq.Text = R!cMetLiquidacion
    Else
        CargaDatos = False
        Frame3.Enabled = True
    End If
    R.Close
    Set R = Nothing
    Exit Function

ErrorCargaDatos:
        MsgBox err.Description, vbCritical, "Aviso"

End Function

Private Sub ActxCta_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrorActxCta_KeyPress
    If KeyAscii = 13 Then
        If Not CargaDatos() Then
            HabilitaActualizacion False
            MsgBox "No se pudo encontrar el Credito, el Credito No esta Vigente o es de Tipo Calendario Dinamico", vbInformation, "Aviso"
        Else
            HabilitaActualizacion True
        End If
    End If
    Exit Sub

ErrorActxCta_KeyPress:
        MsgBox err.Description, vbCritical, "Aviso"
End Sub

Private Sub CmdAceptar_Click()
Dim oCredito As COMDCredito.DCOMCredActBD
    
    On Error GoTo ErrorCmdAceptar_Click
    If Not ValidaDatos Then
        Exit Sub
    End If
    If MsgBox("Se va a Grabar el Metodo de Liquidacion, Desea Continuar?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oCredito = New COMDCredito.DCOMCredActBD
        Call oCredito.dUpdateColocacCred(ActxCta.NroCuenta, , , , , , txtMetLiq.Text, , , , , , , , , False)
        
        'MAVM 20120328 ***
        Dim objPista As COMManejador.Pista
        Set objPista = New COMManejador.Pista
        objPista.InsertarPista gCredActualizarMetodoLiquid, GeneraMovNro(gdFecSis, gsCodAge, gsCodUser), gsCodPersUser, GetMaquinaUsuario, gModificar, "Cambio de Metodo Liquidacion " & txtMetLiq.Text, ActxCta.NroCuenta, gCodigoCuenta
        Set objPista = Nothing
        '***
        
        Set oCredito = Nothing
        HabilitaActualizacion False
        Call LimpiaPantalla
    End If
Exit Sub
ErrorCmdAceptar_Click:
        MsgBox err.Description, vbCritical, "Aviso"
    
End Sub

Private Sub cmdBuscar_Click()
Dim oCredito As COMDCredito.DCOMCredito
Dim R As ADODB.Recordset
Dim oPers As COMDPersona.UCOMPersona

    On Error GoTo ErrorCmdBuscar_Click
    LstCred.Clear
    Set oPers = frmBuscaPersona.Inicio()
    If Not oPers Is Nothing Then
        Set oCredito = New COMDCredito.DCOMCredito
        Set R = oCredito.RecuperaCreditosVigentes(oPers.sPersCod, , Array(gColocEstVigNorm, gColocEstVigVenc, gColocEstVigMor, gColocEstRefVenc, gColocEstRefNorm, gColocEstRefMor))
        Do While Not R.EOF
            LstCred.AddItem R!cCtaCod
            R.MoveNext
        Loop
        R.Close
        Set R = Nothing
        Set oCredito = Nothing
    End If
    If LstCred.ListCount = 0 Then
        MsgBox "El Cliente No Tiene Creditos Vigentes", vbInformation, "Aviso"
    End If
    
    Exit Sub

ErrorCmdBuscar_Click:
    MsgBox err.Description, vbCritical, "Aviso"

End Sub

Private Sub cmdCancelar_Click()
    Call LimpiaPantalla
    HabilitaActualizacion False
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ActxCta.CMAC = gsCodCMAC
    ActxCta.Age = gsCodAge
End Sub

Private Sub LstCred_Click()
    If LstCred.ListCount > 0 And LstCred.ListIndex <> -1 Then
        ActxCta.NroCuenta = LstCred.Text
        ActxCta.SetFocusCuenta
    End If
End Sub

Private Sub OptMetLiq_Click(Index As Integer)
    If Index = 0 Then
        txtMetLiq.Text = "GMIC"
        txtMetLiq.Enabled = True
    Else
        txtMetLiq.Enabled = False
        If Index = 1 Then
            txtMetLiq.Text = "GMiC"
        Else
            txtMetLiq.Text = "GMYC"
        End If
    End If
End Sub
