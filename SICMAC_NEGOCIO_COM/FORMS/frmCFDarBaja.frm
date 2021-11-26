VERSION 5.00
Begin VB.Form frmCFDarBaja 
   Caption         =   "Carta Fianza - Dar de Baja"
   ClientHeight    =   4455
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4605
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   4605
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Height          =   735
      Left            =   120
      TabIndex        =   9
      Top             =   3600
      Width           =   4455
      Begin VB.CommandButton cmdCerrar 
         Caption         =   "&Cerrar"
         Height          =   375
         Left            =   3120
         TabIndex        =   10
         Top             =   240
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Dar de Baja y Vetar CF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      Begin VB.ComboBox cmbAgencia 
         Height          =   315
         ItemData        =   "frmCFDarBaja.frx":0000
         Left            =   1080
         List            =   "frmCFDarBaja.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2520
         Width           =   2535
      End
      Begin VB.TextBox txtCantSFolio 
         Height          =   320
         Left            =   1080
         TabIndex        =   14
         Top             =   3000
         Width           =   1215
      End
      Begin VB.CommandButton cmdSinFolio 
         Caption         =   "&Actualizar"
         Height          =   375
         Left            =   2880
         TabIndex        =   13
         Top             =   3000
         Width           =   1095
      End
      Begin VB.CommandButton cmdDarBaja 
         Caption         =   "&Dar de Baja"
         Height          =   375
         Left            =   2880
         TabIndex        =   4
         Top             =   720
         Width           =   1095
      End
      Begin VB.CommandButton cmdVetar 
         Caption         =   "&Vetar"
         Height          =   375
         Left            =   2880
         TabIndex        =   3
         Top             =   1680
         Width           =   1095
      End
      Begin VB.TextBox txtFDarBaja 
         Height          =   320
         Left            =   960
         TabIndex        =   2
         Top             =   720
         Width           =   1695
      End
      Begin VB.TextBox txtFVetar 
         Height          =   320
         Left            =   960
         TabIndex        =   1
         Top             =   1680
         Width           =   1695
      End
      Begin VB.Label lblAgencia 
         AutoSize        =   -1  'True
         Caption         =   "Agencia:"
         Height          =   195
         Left            =   360
         TabIndex        =   16
         Top             =   2640
         Width           =   630
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Cant. Folios:"
         Height          =   195
         Left            =   120
         TabIndex        =   12
         Top             =   3000
         Width           =   870
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Folios sin numeración Errados:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   2280
         Width           =   2145
      End
      Begin VB.Label lbRegistro 
         AutoSize        =   -1  'True
         Caption         =   "Dar de Baja CF emitidas"
         Height          =   195
         Left            =   240
         TabIndex        =   8
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label lblFdesde 
         AutoSize        =   -1  'True
         Caption         =   "Folio CF:"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label lblFsin 
         AutoSize        =   -1  'True
         Caption         =   "Nº Folio:"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   1680
         Width           =   600
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Vetar Número de Folio"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   1320
         Width           =   1575
      End
   End
End
Attribute VB_Name = "frmCFDarBaja"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public fbConfimacion As Boolean
Dim fsCodEnvio As String
Dim fsCodAgencias As String
Private Sub cmbAgencia_Click()
fsCodAgencias = Trim(Right(Me.cmbAgencia.Text, 4))
fsCodAgencias = IIf(Len(fsCodAgencias) < 2, "0" & fsCodAgencias, fsCodAgencias)
End Sub

Private Sub cmdCerrar_Click()
Unload Me
End Sub

Private Sub cmdDarBaja_Click()
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Dim oBase As COMDCredito.DCOMCredActBD
Dim psCtaCod As String
Dim rsCartaFianza As ADODB.Recordset
    If ValidaDatos(1) Then
        Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
        psCtaCod = oCartaFianza.ObtenerCtaCod(CLng(Me.txtFDarBaja.Text))
        frmCFConfirmacion.Inicio (psCtaCod)
        If fbConfimacion Then
            Set oBase = New COMDCredito.DCOMCredActBD
            Call oBase.dUpdateProducto(psCtaCod, , , gColocEstAprob, gdFecSis, -2, False)
            Call oCartaFianza.ActualizarFolio(CLng(Me.txtFDarBaja.Text), 2, , gdFecSis)
            Call oCartaFianza.QuitarEmisionFolio(psCtaCod)
            Set rsCartaFianza = oCartaFianza.UltimoRegistroEnvio(, , CLng(txtFDarBaja.Text))
            If rsCartaFianza.RecordCount > 0 Then
                Call oCartaFianza.ActualizarEnvioFolios(Trim(rsCartaFianza!nCodEnvio), 2)
            End If
            MsgBox "Nº de Folio dado de baja satisfactoriamente.", vbInformation, "Aviso"
            fbConfimacion = False
            Me.txtFDarBaja.Text = ""
            Set oBase = Nothing
            Set oCartaFianza = Nothing
            Set rsCartaFianza = Nothing
        End If
    End If
End Sub

Private Sub cmdSinFolio_Click()
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Dim rsCartaFianza As ADODB.Recordset
If ValidaDatos(3) Then
    If MsgBox("Esta seguro de Actualizar Cantidad de Folios sin numeración?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
        Set rsCartaFianza = oCartaFianza.UltimoRegistroEnvio(fsCodAgencias, 1, , CLng(Me.txtCantSFolio.Text))
        If rsCartaFianza.RecordCount > 0 Then
            Call oCartaFianza.ActualizarEnvioFolios(Trim(rsCartaFianza!nCodEnvio), 2)
        End If
        Call oCartaFianza.ActualizarFolioSinNumero(fsCodEnvio, CLng(Me.txtCantSFolio.Text))
        MsgBox "Folios sin numeración actualizados satisfactoriamente.", vbInformation, "Aviso"
    End If
End If
End Sub

Private Sub cmdVetar_Click()
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Dim rsCartaFianza As ADODB.Recordset
    If ValidaDatos(2) Then
        Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida
        If MsgBox("Esta seguro de Vetar la Carta Fianza?", vbInformation + vbYesNo, "Aviso") = vbYes Then
            Call oCartaFianza.ActualizarFolio(CLng(Trim(Me.txtFVetar.Text)), 3)
            Set rsCartaFianza = oCartaFianza.UltimoRegistroEnvio(, , CLng(txtFVetar.Text))
            If rsCartaFianza.RecordCount > 0 Then
                Call oCartaFianza.ActualizarEnvioFolios(Trim(rsCartaFianza!nCodEnvio), 2)
            End If
            MsgBox "Nº de Folio vetado satisfactoriamente.", vbInformation, "Aviso"
            Me.txtFVetar.Text = ""
            Set oCartaFianza = Nothing
            Set rsCartaFianza = Nothing
        End If
    End If
End Sub

Private Sub Form_Load()
Call CentraForm(Me)
Me.Icon = LoadPicture(App.path & gsRutaIcono)
fbConfimacion = False
Cargar_Objetos_Controles
End Sub

Private Function ValidaDatos(ByVal pnTipo As Integer) As Boolean
Dim oCartaFianza As COMNCartaFianza.NCOMCartaFianzaValida
Dim rsCartaFianza As ADODB.Recordset
Set oCartaFianza = New COMNCartaFianza.NCOMCartaFianzaValida

If pnTipo = 1 Then
    If Trim(Me.txtFDarBaja.Text) = "" Then
    MsgBox "Ingrese el Nº de Folio", vbInformation, "Aviso"
    ValidaDatos = False
    txtFDarBaja.SetFocus
    Exit Function
    End If
    
    Set rsCartaFianza = oCartaFianza.ObtenerFoliosEmision(CLng(Trim(Me.txtFDarBaja.Text)))
    
    If rsCartaFianza.RecordCount > 0 Then
        If Trim(rsCartaFianza!nEstado) = "2" Then
            MsgBox "Nº Folio ya fue dado de baja.", vbInformation, "Aviso"
            ValidaDatos = False
            txtFDarBaja.SetFocus
            Exit Function
        End If
        
        If Trim(rsCartaFianza!nEstado) = "3" Then
            MsgBox "Nº Folio esta vetado.", vbInformation, "Aviso"
            ValidaDatos = False
            txtFDarBaja.SetFocus
            Exit Function
        End If
        
        If Trim(rsCartaFianza!nEstado) = "0" Then
            MsgBox "Nº Folio aun no fue emitido.", vbInformation, "Aviso"
            ValidaDatos = False
            txtFDarBaja.SetFocus
            Exit Function
        End If
        
        If Trim(rsCartaFianza!nEstado) = "4" Then
            MsgBox "Nº Folio ya se realizo una renovacion.", vbInformation, "Aviso"
            ValidaDatos = False
            txtFDarBaja.SetFocus
            Exit Function
        End If
    Else
        MsgBox "No Existe Nº Folio.", vbInformation, "Aviso"
        ValidaDatos = False
        txtFDarBaja.SetFocus
        Exit Function
    End If
    
End If
If pnTipo = 2 Then
    If Trim(Me.txtFVetar.Text) = "" Then
    MsgBox "Ingrese el Nº de Folio", vbInformation, "Aviso"
    ValidaDatos = False
    txtFVetar.SetFocus
    Exit Function
    End If
    
    Set rsCartaFianza = oCartaFianza.ObtenerFoliosEmision(CLng(Trim(Me.txtFVetar.Text)))
    
    If rsCartaFianza.RecordCount > 0 Then
        If Trim(rsCartaFianza!nEstado) = "2" Then
            MsgBox "Nº Folio esta de baja.", vbInformation, "Aviso"
            ValidaDatos = False
            txtFVetar.SetFocus
            Exit Function
        End If
        
        If Trim(rsCartaFianza!nEstado) = "3" Then
            MsgBox "Nº Folio esta ya fue vetado.", vbInformation, "Aviso"
            ValidaDatos = False
            txtFVetar.SetFocus
            Exit Function
        End If
        
        If Trim(rsCartaFianza!nEstado) = "1" Then
            MsgBox "Nº Folio esta emitido.", vbInformation, "Aviso"
            ValidaDatos = False
            txtFVetar.SetFocus
            Exit Function
        End If
        
        If Trim(rsCartaFianza!nEstado) = "4" Then
            MsgBox "Nº Folio ya se realizo una renovacion.", vbInformation, "Aviso"
            ValidaDatos = False
            txtFVetar.SetFocus
            Exit Function
        End If
    Else
        MsgBox "No Existe Nº Folio.", vbInformation, "Aviso"
        ValidaDatos = False
        txtFVetar.SetFocus
        Exit Function
    End If
    
End If

If pnTipo = 3 Then
    If Me.cmbAgencia.Text = "" Then
        MsgBox "Seleccione Agencia", vbInformation, "Aviso"
        ValidaDatos = False
        Exit Function
    End If
    
    If Me.txtCantSFolio.Text = "" Or Me.txtCantSFolio.Text = "0" Then
        MsgBox "Ingrese Cant. Folios sin numeración.", vbInformation, "Aviso"
        ValidaDatos = False
        txtCantSFolio.SetFocus
        Exit Function
    End If
    
    Set rsCartaFianza = oCartaFianza.ObtenerNumFolioAEmitir(fsCodAgencias, 1, True)
    
    If rsCartaFianza.RecordCount > 0 Then
        If CLng(rsCartaFianza!nSFolio) < (CLng(rsCartaFianza!nSFolioUtl) + CLng(Trim(Me.txtCantSFolio.Text))) Then
            MsgBox "Cant. Folios utilizados no puede ser mayor a lo enviado. " & Chr(10) & "Folio Env.:" & rsCartaFianza!nSFolio & Chr(10) & "Folio Utilizados: " & rsCartaFianza!nSFolioUtl, vbInformation, "Aviso"
            ValidaDatos = False
            txtCantSFolio.SetFocus
            Exit Function
        Else
            fsCodEnvio = rsCartaFianza!nCodEnvio
        End If
    Else
        MsgBox "Agencia no cuenta con Folios sin numeración.", vbInformation, "Aviso"
        ValidaDatos = False
        txtCantSFolio.SetFocus
        Exit Function
    End If
End If

ValidaDatos = True
End Function


Private Sub txtCantSFolio_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub txtFDarBaja_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub
Private Sub txtFVetar_KeyPress(KeyAscii As Integer)
KeyAscii = NumerosEnteros(KeyAscii)
End Sub

Private Sub Cargar_Objetos_Controles()
Dim oAgencia  As COMDConstantes.DCOMAgencias
Dim rsAgencias As ADODB.Recordset
Set oAgencia = New COMDConstantes.DCOMAgencias
Set rsAgencias = oAgencia.ObtieneAgencias()
Call Llenar_Combo_con_Recordset(rsAgencias, cmbAgencia)
End Sub

