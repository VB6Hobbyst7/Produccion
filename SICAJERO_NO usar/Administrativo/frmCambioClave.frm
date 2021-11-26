VERSION 5.00
Object = "{F9AB04EF-FCD4-4161-99E1-9F65F8191D72}#12.0#0"; "OCXTarjeta.ocx"
Begin VB.Form frmCambioClave 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cambio de Clave - F12 para Digitar Tarjeta"
   ClientHeight    =   3285
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmCambioClave.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3285
   ScaleWidth      =   5520
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin OCXTarjeta.CtrlTarjeta Tarjeta 
      Height          =   555
      Left            =   5970
      TabIndex        =   14
      Top             =   780
      Width           =   870
      _extentx        =   1535
      _extenty        =   979
   End
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   30
      TabIndex        =   11
      Top             =   2400
      Width           =   5490
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   420
         Left            =   4125
         TabIndex        =   13
         Top             =   195
         Width           =   1215
      End
      Begin VB.CommandButton cmdAceptar 
         Caption         =   "Aceptar"
         Enabled         =   0   'False
         Height          =   405
         Left            =   105
         TabIndex        =   12
         Top             =   225
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Claves"
      Enabled         =   0   'False
      Height          =   1545
      Left            =   15
      TabIndex        =   4
      Top             =   855
      Width           =   5490
      Begin VB.CommandButton CmdPedirClaNewC 
         Caption         =   "Pedir Clave"
         Height          =   360
         Left            =   3735
         TabIndex        =   15
         Top             =   1000
         Width           =   1305
      End
      Begin VB.CommandButton CmdPedirClaNew 
         Caption         =   "Pedir Clave"
         Height          =   360
         Left            =   3735
         TabIndex        =   10
         Top             =   600
         Width           =   1305
      End
      Begin VB.CommandButton CmdPedClaveAnt 
         Caption         =   "Pedir Clave"
         Height          =   360
         Left            =   3750
         TabIndex        =   7
         Top             =   225
         Width           =   1305
      End
      Begin VB.Label Label5 
         Caption         =   "Confirmar Clave :"
         Height          =   315
         Left            =   210
         TabIndex        =   17
         Top             =   1080
         Width           =   1230
      End
      Begin VB.Label lblClaveNewC 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NO INGRESADO"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1485
         TabIndex        =   16
         Top             =   1080
         Width           =   2085
      End
      Begin VB.Label lblClaveNew 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NO INGRESADO"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1485
         TabIndex        =   9
         Top             =   660
         Width           =   2085
      End
      Begin VB.Label Label3 
         Caption         =   "Clave Nueva :"
         Height          =   315
         Left            =   210
         TabIndex        =   8
         Top             =   705
         Width           =   1230
      End
      Begin VB.Label LblClaveAnt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "NO INGRESADO"
         ForeColor       =   &H00400000&
         Height          =   285
         Left            =   1485
         TabIndex        =   6
         Top             =   285
         Width           =   2085
      End
      Begin VB.Label Label1 
         Caption         =   "Clave Anterior :"
         Height          =   315
         Left            =   225
         TabIndex        =   5
         Top             =   330
         Width           =   1230
      End
   End
   Begin VB.Frame Frame1 
      Height          =   855
      Left            =   15
      TabIndex        =   0
      Top             =   0
      Width           =   5505
      Begin VB.CommandButton CmdLecTarj 
         Caption         =   "Leer Tarjeta"
         Height          =   345
         Left            =   4065
         TabIndex        =   1
         Top             =   255
         Width           =   1290
      End
      Begin VB.Label Lblnumtarjeta 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   375
         Left            =   795
         TabIndex        =   3
         Top             =   240
         Width           =   3225
      End
      Begin VB.Label Label2 
         Caption         =   "Tarjeta :"
         Height          =   255
         Left            =   60
         TabIndex        =   2
         Top             =   300
         Width           =   735
      End
   End
End
Attribute VB_Name = "frmCambioClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

 Private Declare Function GetTokenInfo _
    Lib "RQxDFTk.dll" _
                 (ByVal file As String, _
                  ByVal info As String, _
                  ByVal subinfo As String, _
                  ByVal tokenitem As String _
                 ) As Long
                 
    Private Declare Function pinverify _
    Lib "PINVerify.dll" _
                 (ByVal ippuerto As String, _
                  ByVal key As String, _
                  ByVal PAN As String, _
                  ByVal pvki As String, _
                  ByVal pIn As String, _
                  ByVal pvv As String _
                 ) As Integer
    
    Private Declare Function changepin _
    Lib "PINVerify.dll" _
                 (ByVal ippuerto As String, _
                  ByVal key As String, _
                  ByVal PAN As String, _
                  ByVal pvki As String, _
                  ByVal pIn As String, _
                  ByVal pvv As String, _
                  ByVal npin As String _
                 ) As Long
    
    Private Declare Function genpckey Lib "PINVerify.dll" () As Long
    
    Private Declare Function lstrlenW Lib "kernel32.dll" (ByVal lpString As Long) As Long
    
    Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)




Dim sTrack2 As String
Dim sPinAnt As String
Dim sPinNew As String

'**DAOR 20081125 ***********************
Dim sPinNewC As String
Dim nValidaPIN As Integer
Dim sPVVNew As String
Dim sPVVNewC As String
Dim bClavesIguales As Boolean
'***************************************

Dim sPVV As String


Private Sub CmdAceptar_Click()
Dim sResp As String
Dim sTramaResp As String
Dim i As Integer, X As Long

On Error GoTo ManejadorError
'sTramaResp = "CP04NONE"
'OJO  con tipoTxt 86 funciona
'VP04NONE es para valiaar el pin y CP04NONE es para cambiar PIN
'sResp = InstanciaProxy("86", Lblnumtarjeta.Caption, sTramaResp, sTramaResp, sTrack2, sPinAnt, sPinNew)

'sPVV = RecuperaPVV(Me.Lblnumtarjeta.Caption)


'**Modificado por DAOR 20081125 **********************************************************************
' i = pinverify("192.168.0.9:81", "444", Me.Lblnumtarjeta.Caption, "1", sPinAnt, sPVV)
'sPVVNew = changepin("192.168.0.9:81", "444", Me.Lblnumtarjeta.Caption, "1", sPinAnt, sPVV, sPinNew)
            
'i = pinverify(gIpPuertoPinVerifyPOS, "444", Me.Lblnumtarjeta.Caption, "1", "t" & sPinAnt & gCanalIdPOS, sPVV)
i = nValidaPIN
'*****************************************************************************************************


If i = 1 Then 'Si la clave es correcta
    'Verificar que las nuevas claves coinciden
    If gnTipoPinPad = 2 Then
        If sPVVNew = sPVVNewC Then bClavesIguales = True Else bClavesIguales = False
    Else
        If sPinNew = sPinNewC Then bClavesIguales = True Else bClavesIguales = False
    End If
    
    If Not bClavesIguales Then
        Call MsgBox("Las nuevas claves no coinciden, intentelo nuevamente", vbInformation, "Aviso")
        Call ReiniciarForm
        Exit Sub
    End If
    
    If gnTipoPinPad <> 2 Then
        X = changepin(gIpPuertoPinVerifyPOS, "444", Me.Lblnumtarjeta.Caption, "1", "t" & sPinAnt & gCanalIdPOS, sPVV, "t" & sPinNew & gCanalIdPOS)
        sPVVNew = DevuelveParametro(X)
    End If
    
    If sPVVNew = sPVV Then 'si la clave actual es igual a la nueva
        Call MsgBox("No se pudo realizar el cambio de clave, debido a que la clave actual y la nueva son iguales, intentelo nuevamente", vbInformation, "Aviso")
        Call ReiniciarForm
        Exit Sub
    End If
    
    If Left(sPVVNew, 3) = "ERR" Then
        Call MsgBox("Ocurrio un error al intentar generar la nueva clave, intentelo nuevamente", vbInformation, "Aviso")
        Call ReiniciarForm
        Exit Sub
    End If
    
    Call ActualizaPVV(sPVVNew, Me.Lblnumtarjeta.Caption)
    sResp = "00"
        
Else
    sResp = "99"
End If


If sResp = "00" Then
   Call MsgBox("Cambio de Clave Realizado Correctamente", vbInformation, "Aviso")

Else
   Call MsgBox("Error en la Clave, Reintente por favor, y si persiste el Error Avisar al Area de Sistemas", vbInformation, "Aviso")
End If
 
Call ReiniciarForm

Exit Sub

ManejadorError:
     Call MsgBox("Ocurrio un error inesperado, si el problema persiste comuniquese con el Area de Sistemas", vbInformation, "Aviso")
End Sub

Private Sub CmdLecTarj_Click()

    Me.Caption = "Cambio de Clave - PASE LA TARJETA"
    sTrack2 = Mid(Tarjeta.LeerTarjeta("PASE LA TARJETA", gnTipoPinPad, gnPinPadPuerto), 2, 32)
    Lblnumtarjeta.Caption = Mid(sTrack2, 1, 16)
    Frame2.Enabled = True
    cmdAceptar.Enabled = True
    CmdPedClaveAnt.SetFocus
    Me.Caption = "Cambio de Clave"
    
    '**DAOR 20081127 ******************************************
    If Left(Me.Lblnumtarjeta.Caption, 3) <> "ERR" Then
        sPVV = RecuperaPVV(Me.Lblnumtarjeta.Caption)
    End If
    '**********************************************************
 
End Sub

Private Sub CmdPedClaveAnt_Click()
'sPinAnt = Tarjeta.PedirPinDes
'17/06/2008

'**Modificado por DAOR 20081125**********************************
'sPinAnt = Tarjeta.PedirPinEnc(Me.Lblnumtarjeta.Caption, "0", "0123456789123456")
'sPinAnt = Tarjeta.PedirPinEnc(Me.Lblnumtarjeta.Caption, gNMKPOS, gWKPOS, gnTipoPinPad, gnPinPadPuerto)
'****************************************************************

'**DAOR 20081127, Nuevo, para trabajar con los dos tipos de PinPad *******
nValidaPIN = 0
nValidaPIN = Tarjeta.PedirPinYValida(Me.Lblnumtarjeta.Caption, gNMKPOS, gWKPOS, gIpPuertoPinVerifyPOS, sPVV, gCanalIdPOS, gnTipoPinPad, gnPinPadPuerto, sPinAnt)
'*************************************************************************


'**Modificado por DAOR 20081127 *********************************
'If Len(Trim(sPinAnt)) > 0 Then
'    Me.LblClaveAnt.Caption = "CLAVE INGRESADA"
'Else
'    Me.LblClaveAnt.Caption = "NO INGRESADO"
'End If
If nValidaPIN <> 0 Then
    Me.LblClaveAnt.Caption = "CLAVE INGRESADA"
Else
    Me.LblClaveAnt.Caption = "NO INGRESADO"
End If
'****************************************************************

End Sub

Private Sub CmdPedirClaNew_Click()
'sPinNew = Tarjeta.PedirPinDes
'17/06/2008

'**Modificado por DAOR 20081125**********************************
'sPinNew = Tarjeta.PedirPinEnc(Me.Lblnumtarjeta.Caption, "0", "0123456789123456")
If gnTipoPinPad = 2 Then
    sPVVNew = Tarjeta.PedirPinDevPvvACS(Me.Lblnumtarjeta.Caption, gIpPuertoPinVerifyPOS, gnTipoPinPad, gnPinPadPuerto)
Else
    sPinNew = Tarjeta.PedirPinEnc(Me.Lblnumtarjeta.Caption, gNMKPOS, gWKPOS, gnTipoPinPad, gnPinPadPuerto)
End If
'****************************************************************


If Len(Trim(sPinNew)) > 0 Or Len(Trim(sPVVNew)) > 0 Then
    Me.lblClaveNew.Caption = "CLAVE INGRESADA"
Else
    Me.lblClaveNew.Caption = "NO INGRESADO"
End If

End Sub

'**DAOR 20081125 *******************************
Private Sub CmdPedirClaNewC_Click()
    If gnTipoPinPad = 2 Then
        sPVVNewC = Tarjeta.PedirPinDevPvvACS(Me.Lblnumtarjeta.Caption, gIpPuertoPinVerifyPOS, gnTipoPinPad, gnPinPadPuerto)
    Else
        sPinNewC = Tarjeta.PedirPinEnc(Me.Lblnumtarjeta.Caption, gNMKPOS, gWKPOS, gnTipoPinPad, gnPinPadPuerto)
    End If

    
    If Len(Trim(sPinNewC)) > 0 Or Len(Trim(sPVVNewC)) > 0 Then
        Me.lblClaveNewC.Caption = "CLAVE INGRESADA"
    Else
        Me.lblClaveNewC.Caption = "NO INGRESADO"
    End If
End Sub
'*************************************************

Private Sub CmdSalir_Click()
    Unload Me
End Sub


Private Function DevuelveParametro(ByVal X As Long) As String
Dim buffer() As Byte
Dim nLen As Long
Dim res As String
Dim i As Integer
 nLen = lstrlenW(X) * 2
     ReDim buffer(0 To (nLen - 1)) As Byte
    CopyMemory buffer(0), ByVal X, nLen
    res = ""
    For i = 0 To nLen - 1
    If (buffer(i) = 0) Then
    Exit For
    End If
    res = res + Chr(buffer(i))
    Next
    DevuelveParametro = res
End Function


Private Sub ReiniciarForm()
    Me.Lblnumtarjeta.Caption = ""
    Me.LblClaveAnt.Caption = ""
    Me.lblClaveNew.Caption = ""
    Me.lblClaveNewC.Caption = "" 'DAOR 20081125
    Me.Frame2.Enabled = False
    Me.cmdAceptar.Enabled = False
End Sub
