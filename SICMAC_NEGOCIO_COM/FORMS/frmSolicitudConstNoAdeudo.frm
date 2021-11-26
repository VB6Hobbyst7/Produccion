VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmSolicitudConstNoAdeudo 
   Caption         =   "Solicitud de Constancia de No Adeudo"
   ClientHeight    =   3915
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7470
   Icon            =   "frmSolicitudConstNoAdeudo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3915
   ScaleWidth      =   7470
   StartUpPosition =   1  'CenterOwner
   Begin VB.Frame Frame1 
      Caption         =   "Datos de la Solicitud:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7215
      Begin VB.CommandButton cmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Top             =   2570
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   375
         Left            =   4440
         TabIndex        =   8
         Top             =   2570
         Width           =   1215
      End
      Begin VB.TextBox txtMotivo 
         Height          =   1095
         Left            =   1080
         TabIndex        =   7
         Top             =   1320
         Width           =   5895
      End
      Begin VB.TextBox txtNombre 
         Height          =   315
         Left            =   3600
         TabIndex        =   2
         Top             =   360
         Width           =   3375
      End
      Begin SICMACT.TxtBuscar txtCodigo 
         Height          =   315
         Left            =   960
         TabIndex        =   3
         Top             =   360
         Width           =   1815
         _ExtentX        =   2990
         _ExtentY        =   556
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
         TipoBusqueda    =   3
         sTitulo         =   ""
      End
      Begin MSMask.MaskEdBox txtFSolicitud 
         Height          =   300
         Left            =   1470
         TabIndex        =   10
         Top             =   840
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         BackColor       =   12648447
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Motivo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   6
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Nombre:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2880
         TabIndex        =   5
         Top             =   390
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Código:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   390
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "F. de Solicitud:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   880
         Width           =   1095
      End
   End
   Begin VB.Label lblConst 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3480
      Width           =   7215
   End
End
Attribute VB_Name = "frmSolicitudConstNoAdeudo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private oPersona As UPersona_Cli

Private Sub cmdAceptar_Click()
    Dim objCOMNCredito As COMNCredito.NCOMCredito
    Dim oCredAct As COMDCredito.DCOMCredActBD 'JUEZ 20151229
    Set objCOMNCredito = New COMNCredito.NCOMCredito
    Set oCredAct = New COMDCredito.DCOMCredActBD 'JUEZ 20151229
    
        If Len(Trim(txtCodigo.Text)) <> "0" And txtFSolicitud.Text <> "__/__/____" Then
            If MsgBox("Esta Seguro de Registar los Datos?", vbYesNo + vbQuestion, Me.Caption) = vbYes Then
                objCOMNCredito.InsertarSolicitudConstNoAdeudo txtCodigo.Text, txtFSolicitud.Text, txtMotivo.Text
                Call oCredAct.dUpdateComision(txtCodigo.Text, gComiCredConstNoAdeudo) 'JUEZ 20151229
                Limpiar
            End If
        Else
            MsgBox "Debe Completar los datos", vbCritical, "Aviso"
        End If
    
End Sub

Private Sub cmdCancelar_Click()
     Limpiar
End Sub

Private Sub Limpiar()
     txtCodigo.Text = ""
     txtNombre.Text = ""
     txtFSolicitud.Text = "__/__/____"
     txtMotivo.Text = ""
     lblConst.Caption = ""
End Sub

Private Sub txtCodigo_EmiteDatos()
    Dim sCodigo As String
    Dim odRFa As COMDCredito.DCOMRFA
    Dim rs As ADODB.Recordset
    
    If txtCodigo.Text = "" Then
        Limpiar
    Else
        txtFSolicitud.Text = "__/__/____"
        txtMotivo.Text = ""
        lblConst.Caption = ""
    End If
    
    If txtCodigo.Text <> "" Then
        sCodigo = txtCodigo.Text
        Set odRFa = New COMDCredito.DCOMRFA
        Set rs = odRFa.BuscarPersona(sCodigo)
        Set odRFa = Nothing
    
        If Not rs.EOF And Not rs.BOF Then
            txtNombre.Text = rs!cPersNombre
        Else
            Limpiar
        End If
        Set rs = Nothing
        'JUEZ 20151229 ******************************
        If ValidarConstNoAdeudo = True Then
            'MsgBox "La Primera Constancia es Gratis", vbInformation, "Aviso"
            lblConst.Caption = "La primera constancia es gratis"
            txtFSolicitud.SetFocus
        Else
            'MsgBox "Debe Realizar el Pago en la Ventanilla", vbInformation, "Aviso"
            'lblConst.Caption = "Realizar el Pago en la Ventanilla"
            Dim oCred As New COMDCredito.DCOMCredito
            If oCred.ExisteComisionVigente(sCodigo, gComiCredConstNoAdeudo) = False Then
               MsgBox "Debe realizar el pago de la comisión antes de realizar la solicitud", vbInformation, "Aviso"
               Limpiar
            Else
                lblConst.Caption = "Comisión por constancia pagada"
                txtFSolicitud.SetFocus
            End If
        End If
        
        'JUEZ 20151229 ******************************
    End If
End Sub

Private Sub txtFSolicitud_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
       txtMotivo.SetFocus
    End If
End Sub

Private Sub txtFSolicitud_LostFocus()
    Dim sCad As String
    sCad = ValidaFecha(txtFSolicitud.Text)
        If Not Trim(sCad) = "" Then
            MsgBox sCad, vbInformation, "Aviso"
            Exit Sub
        End If
        If CDate(txtFSolicitud.Text) > gdFecSis Then
            MsgBox "Fecha No Puede Ser Mayor o Igual que la Fecha del Sistema", vbInformation, "Aviso"
            txtFSolicitud.SetFocus
            Exit Sub
        End If
End Sub

Private Function ValidarConstNoAdeudo() As Boolean
    Dim objCOMNCredito As COMNCredito.NCOMCredito
    Set objCOMNCredito = New COMNCredito.NCOMCredito
    Dim rs As ADODB.Recordset
    
    Set rs = objCOMNCredito.ValidarConstNoAdeudoXPersona(txtCodigo.Text)
    If rs.RecordCount = 0 Then
        ValidarConstNoAdeudo = True
    Else
        ValidarConstNoAdeudo = False
    End If
End Function

Private Sub txtMotivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        cmdAceptar.SetFocus
    End If
End Sub
