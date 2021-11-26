VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "mshflxgd.ocx"
Object = "{A7C47A80-96CC-11CF-8B85-0020AFE89883}#4.0#0"; "SigBox.OCX"
Begin VB.Form frmCapTarjetaCambioClave 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7380
   Icon            =   "frmCapTarjetaCambioClave.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6060
   ScaleWidth      =   7380
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin SigBoxLib.SigBox boxFirma 
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   5640
      Width           =   1095
      _Version        =   262144
      _ExtentX        =   1931
      _ExtentY        =   450
      _StockProps     =   233
      Appearance      =   1
      TitleText       =   ""
      PromptText      =   ""
      Picture         =   "frmCapTarjetaCambioClave.frx":030A
      DebugFileName   =   "SigBox1.TXT"
   End
   Begin VB.CommandButton cmdGrabar 
      Caption         =   "&Grabar"
      Height          =   375
      Left            =   5250
      TabIndex        =   19
      Top             =   5565
      Width           =   960
   End
   Begin VB.Frame fraPersona 
      Caption         =   "Persona"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1170
      Left            =   105
      TabIndex        =   12
      Top             =   840
      Width           =   7155
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCliente 
         Height          =   855
         Left            =   105
         TabIndex        =   13
         Top             =   210
         Width           =   6945
         _ExtentX        =   12250
         _ExtentY        =   1508
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   6300
      TabIndex        =   11
      Top             =   5565
      Width           =   960
   End
   Begin VB.Frame fraRelacion 
      Caption         =   "Cuentas Relacionadas"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1905
      Left            =   105
      TabIndex        =   9
      Top             =   2100
      Width           =   3480
      Begin SICMACT.FlexEdit grdClienteTarj 
         Height          =   1380
         Left            =   105
         TabIndex        =   10
         Top             =   315
         Width           =   3270
         _ExtentX        =   5768
         _ExtentY        =   2434
         Cols0           =   3
         HighLight       =   1
         AllowUserResizing=   3
         VisiblePopMenu  =   -1  'True
         EncabezadosNombres=   "#-Cuenta-Relación"
         EncabezadosAnchos=   "250-1800-1100"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X"
         ListaControles  =   "0-0-0"
         EncabezadosAlineacion=   "C-C-C"
         FormatosEdit    =   "0-0-0"
         TextArray0      =   "#"
         lbFlexDuplicados=   0   'False
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   285
      End
   End
   Begin VB.Frame fraTarjeta 
      Height          =   750
      Left            =   105
      TabIndex        =   6
      Top             =   0
      Width           =   3405
      Begin MSMask.MaskEdBox txtTarjeta 
         Height          =   375
         Left            =   945
         TabIndex        =   7
         Top             =   210
         Width           =   2325
         _ExtentX        =   4101
         _ExtentY        =   661
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   19
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "####-####-####-####"
         Mask            =   "####-####-####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Tarjeta :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   105
         TabIndex        =   8
         Top             =   300
         Width           =   735
      End
   End
   Begin VB.Frame fraEstado 
      Caption         =   "Estado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1485
      Left            =   105
      TabIndex        =   3
      Top             =   3990
      Width           =   7155
      Begin VB.TextBox lblpassw2 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   5760
         PasswordChar    =   "*"
         TabIndex        =   15
         Top             =   225
         Width           =   810
      End
      Begin VB.TextBox lblpassw1 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   4710
         PasswordChar    =   "*"
         TabIndex        =   14
         Top             =   225
         Width           =   810
      End
      Begin VB.TextBox txtGlosa 
         Appearance      =   0  'Flat
         Height          =   540
         Left            =   945
         TabIndex        =   4
         Top             =   735
         Width           =   6000
      End
      Begin VB.Label lblTrack1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Verdana"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   330
         Left            =   945
         TabIndex        =   18
         Top             =   240
         Width           =   2565
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Número :"
         Height          =   195
         Left            =   210
         TabIndex        =   17
         Top             =   315
         Width           =   645
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Password :"
         Height          =   165
         Left            =   3780
         TabIndex        =   16
         Top             =   330
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Glosa :"
         Height          =   195
         Left            =   210
         TabIndex        =   5
         Top             =   735
         Width           =   495
      End
   End
   Begin VB.Frame fraHistoria 
      Caption         =   "Historia"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1905
      Left            =   3675
      TabIndex        =   1
      Top             =   2100
      Width           =   3585
      Begin SICMACT.FlexEdit grdTarjetaEstado 
         Height          =   1380
         Left            =   105
         TabIndex        =   2
         Top             =   315
         Width           =   3375
         _ExtentX        =   5953
         _ExtentY        =   2434
         Cols0           =   5
         HighLight       =   1
         AllowUserResizing=   3
         EncabezadosNombres=   "#-Fecha-Estado-Comentario-Usu"
         EncabezadosAnchos=   "250-1000-2000-3500-600"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnasAEditar =   "X-X-X-X-X"
         ListaControles  =   "0-0-0-0-0"
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         BackColorControl=   -2147483643
         EncabezadosAlineacion=   "C-C-C-C-C"
         FormatosEdit    =   "0-0-0-0-0"
         TextArray0      =   "#"
         lbUltimaInstancia=   -1  'True
         ColWidth0       =   255
         RowHeight0      =   285
      End
   End
   Begin VB.CommandButton cmdCancelar 
      Caption         =   "&Cancelar"
      Height          =   375
      Left            =   105
      TabIndex        =   0
      Top             =   5565
      Width           =   960
   End
   Begin VB.Label lblMensaje 
      Caption         =   "Presiones <F11> para activar la lectura de la tarjeta...."
      Height          =   495
      Left            =   3780
      TabIndex        =   20
      Top             =   128
      Width           =   3375
   End
End
Attribute VB_Name = "frmCapTarjetaCambioClave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim lsCodTar As String



Private Sub MuestraPantalla()
boxFirma.Clear
boxFirma.TitleText = "CMAC TRUJILLO"
boxFirma.PromptText = "Pase su Tarjeta"
boxFirma.PutBitmap 2, 2, App.path & "\Logo5.bmp"
End Sub

'Private Function GetNumTarjeta() As String
'Dim Result As Long
'Dim lsTarjeta As String, sSQL As String
'
'Result = McrRead(lsTarjeta, 76, 0, sSQL, 0, 0, "", 0, 0)
'If Result <= 0 Then
'   GetNumTarjeta = GetCardNumber(lsTarjeta)
'Else
'    MsgBox GetErrorPINPAD(Result) & " Consulte con Servicio Tecnico.", vbInformation, "Aviso"
'    GetNumTarjeta = ""
'End If
'End Function

Private Function GetCardNumber(ByVal psTrack As String) As String
    GetCardNumber = Mid(psTrack, 3, 16)
End Function

Private Sub UnSetupPad()
boxFirma.Clear
boxFirma.MemoryClear
boxFirma.MagCardEnabled = False
If boxFirma.ConnectedToPad = True Then
'    boxFirma.LoadLogoPicture LogoPict
End If
boxFirma.ConnectToPad = Never
End Sub

Private Function ConectaPad() As Boolean
On Error GoTo ErrPad
boxFirma.ConnectToPad = Always
ConectaPad = boxFirma.ConnectedToPad
If Not ConectaPad Then
    MsgBox "Error de Conexión Con PENWARE." & Chr$(13) & "Verifique la Conexión o Consulte con su Administrador", vbInformation, "Aviso "
End If
Exit Function
ErrPad:
    MsgBox Err.Description, vbCritical, "Error"
End Function

'
Private Function GrabaTarjetaPINPADV_5000(ByVal psCodTarj As String, ByVal pnCom As COMDConstantes.TipoPuertoSerial)
Dim sNumTar As String
Dim sClaveTar As String
Dim lnErr As Long
Dim lnNumOp As Integer
Dim sTitulo As String

Dim lnNumTar As String
Dim lnClaveTar As String
'Dim lnErr As Long
'Dim lnNumOp As Integer
Dim sTarjeta As String, sCaption As String

Dim clsGen As COMDConstSistema.DCOMGeneral
Set clsGen = New COMDConstSistema.DCOMGeneral


'MsgBox "PROCEDA A INGRESAR SU CLAVE...", vbInformation, "AVISO"

Me.Caption = "Ingrese la Clave de la Tarjeta."
sClaveTar = GetClaveTarjeta_Vrf5000("INGRESE CLAVE ")
If sClaveTar = "" Then
    MsgBox "Debe Ingresar una Clave Valida.", vbInformation, "Aviso"
    lblTrack1 = ""
    Exit Function
End If

Dim lnResult As ResultVerificacionTarjeta 'verificar
Set clsGen = New COMDConstSistema.DCOMGeneral
'*************************************
Select Case clsGen.ValidaTarjeta(psCodTarj, sClaveTar)
            Case gClaveValida
                lblpassw1 = sClaveTar
                lnClaveTar = ""
                lnNumOp = 0
                While lnNumOp < 3 And lblpassw1 <> lnClaveTar
                    MsgBox "Ingrese NUEVA Clave de Tarjeta.", vbInformation, "Aviso"
                    Me.Caption = "Ingrese NUEVA Clave de Tarjeta."
                    lnClaveTar = GetClaveTarjeta_Vrf5000("Nueva Clave")
                    MsgBox "Confirme NUEVA Clave de Tarjeta.", vbInformation, "Aviso"
                    Me.Caption = "Confirme NUEVA Clave de Tarjeta."
                    lblpassw1 = GetClaveTarjeta_Vrf5000("Confirme Clave")
                    lnNumOp = lnNumOp + 1
                    If lblpassw1 <> lnClaveTar Then  'And lnNumOp < 3 Then
                        MsgBox "Ingreso de Nueva Clave Erroneo" + Chr(13) + "Reintentelo (" + Str(lnNumOp) + ")", vbInformation, "Aviso"
                    End If
                Wend
                lblpassw2 = lnClaveTar
                If lblpassw1 = lblpassw2 Then
                    txtGlosa.SetFocus
                    cmdGrabar.Enabled = True
                    cmdCancelar.Enabled = True
                    GrabaTarjetaPINPADV_5000 = True
                Else
                    MsgBox "La clave no ha sido reconocida, el proceso sera Cancelado.", vbInformation, "Aviso"
                    'FinalizaPinPad
                    Exit Function
                End If
                Me.Caption = sCaption
                'FinalizaPinPad
                cmdCancelar.Enabled = False
            Case gTarjNoRegistrada
                'ppoa Modificacion
                If Not WriteToLcd("Espere Por Favor") Then
                    FinalizaPinPad
                    MsgBox "No se Realizó Envío", vbInformation, "Aviso"
                    Exit Function
                End If
                MsgBox "Tarjeta no Registrada", vbInformation, "Aviso"
            Case gClaveNOValida
                'ppoa Modificacion
                If Not WriteToLcd("Clave Incorrecta") Then
                    'FinalizaPinPad
                    MsgBox "No se Realizó Envío", vbInformation, "Aviso"
                    lblTrack1 = ""
                    Exit Function
                End If
                MsgBox "Clave Incorrecta", vbInformation, "Aviso"
End Select
'Else
'                lblpassw1 = " "
'                lnClaveTar = "  "
'                lnNumOp = 0
'                While lnNumOp < 3 And lblpassw1 <> lnClaveTar
'                    MsgBox "Ingrese NUEVA Clave de Tarjeta.", vbInformation, "Aviso"
'                    Me.Caption = "Ingrese NUEVA Clave de Tarjeta."
'                    lnClaveTar = GetClaveTarjeta("*")
'                    MsgBox "Confirme NUEVA Clave de Tarjeta.", vbInformation, "Aviso"
'                    Me.Caption = "Confirme NUEVA Clave de Tarjeta."
'                    lblpassw1 = GetClaveTarjeta("**")
'                    lnNumOp = lnNumOp + 1
'                    If lblpassw1 <> lnClaveTar Then  'And lnNumOp < 3 Then
'                        MsgBox "Ingreso de Nueva Clave Erroneo" + Chr(13) + "Reintentelo (" + Str(lnNumOp) + ")", vbInformation, "Aviso"
'                    End If
'                Wend
'                lblpassw2 = lnClaveTar
'                If lblpassw1 = lblpassw2 Then
'                    txtGlosa.Enabled = True
'                    txtGlosa.SetFocus
'                    cmdGrabar.Enabled = True
'                    cmdCancelar.Enabled = True
'                    GrabaTarjetaPINPADV_5000 = True
'                Else
'                    MsgBox "La clave no ha sido reconocida, el proceso sera Cancelado.", vbInformation, "Aviso"
'                    FinalizaPinPad
'                    Exit Function
'                End If
'                Me.Caption = sCaption
'                FinalizaPinPad
'                cmdCancelar.Enabled = False
'End Select


'*************************************


'lblpassw1 = sClaveTar
'sClaveTar = ""
'lnNumOp = 0
'
'While lnNumOp < 3 And lblpassw1 <> sClaveTar
'    sClaveTar = GetClaveTarjeta("Conf.Clave [" & Trim(Str(lnNumOp + 1)) & "]")
'    lnNumOp = lnNumOp + 1
'    If lblpassw1 <> sClaveTar And lnNumOp < 3 Then
'        MsgBox "La clave es errada. Re-Ingrese su Clave.", vbInformation, "Aviso"
'    End If
'Wend
'
'lblpassw2 = sClaveTar
If lblpassw1 = lblpassw2 Then
'    If MsgBox("Desea Registrar la Tarjeta ? ", vbQuestion + vbYesNo, "Aviso") = vbYes Then
'        AgregaTarjeta
'    End If
Else
    MsgBox "La clave no ha sido reconocida, el proceso sera Cancelado.", vbInformation, "Aviso"
End If
'cmdCancelaCard.Enabled = False
Me.Caption = sTitulo
End Function

Private Function GrabaTarjetaPINPAD() As Boolean
Dim lnNumTar As String
Dim lnClaveTar As String
Dim lnErr As Long
Dim lnNumOp As Integer
Dim sTarjeta As String, sCaption As String

Dim clsGen As COMDConstSistema.DCOMGeneral
Set clsGen = New COMDConstSistema.DCOMGeneral
sCaption = Me.Caption

'****************************

'****************************
'ppoa Modificacion
If Not WriteToLcd("Pase su Tarjeta por la Lectora.") Then
    FinalizaPinPad
    Me.Caption = sCaption
    MsgBox "No se Realizó Envío", vbInformation, "Aviso"
    Exit Function
End If

Me.Caption = "Lectura de Tarjeta Activada. Pase la tarjeta por la Lectora."


'lnNumTar = Trim(Mid("0" & GetNumTarjeta(), 1, 16))


'ppoa Modificacion
lnNumTar = GetNumTarjeta

sTarjeta = Replace(txtTarjeta, "-", "", 1, , vbTextCompare)
sTarjeta = Replace(sTarjeta, "_", "", 1, , vbTextCompare)
sTarjeta = Trim(sTarjeta)
lsCodTar = lnNumTar


If Len(lnNumTar) <> 16 Then
        MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
        FinalizaPinPad
        Me.Caption = sCaption
        Exit Function
End If

If Trim(sTarjeta) = "" Then
    sTarjeta = Format$(lsCodTar, "@@@@-@@@@-@@@@-@@@@")
    txtTarjeta.Text = Format$(lsCodTar, "@@@@-@@@@-@@@@-@@@@")
    ObtieneDatosTarjeta
End If


If lsCodTar <> sTarjeta Then

    lblTrack1 = Format$(lnNumTar, "@@@@ @@@@ @@@@ @@@@")
    Me.Caption = "Ingrese la Clave de la Tarjeta."
    
'    'ppoa Modificacion
'    If Not WriteToLcd("Ingrese Clave") Then
'        FinalizaPinPad
'        MsgBox "No se Realizó Envío", vbInformation, "Aviso"
'        Exit Function
'    End If
    
    
    'ppoa Modificacion
    lnClaveTar = GetClaveTarjeta
    
        Dim lnResult As ResultVerificacionTarjeta 'verificar
        
        Set clsGen = New COMDConstSistema.DCOMGeneral
        Select Case clsGen.ValidaTarjeta(lnNumTar, lnClaveTar)
            Case gClaveValida
                lblpassw1 = lnClaveTar
                lnClaveTar = ""
                lnNumOp = 0
                While lnNumOp < 3 And lblpassw1 <> lnClaveTar
                    MsgBox "Ingrese NUEVA Clave de Tarjeta.", vbInformation, "Aviso"
                    Me.Caption = "Ingrese NUEVA Clave de Tarjeta."
                    lnClaveTar = GetClaveTarjeta("*")
                    MsgBox "Confirme NUEVA Clave de Tarjeta.", vbInformation, "Aviso"
                    Me.Caption = "Confirme NUEVA Clave de Tarjeta."
                    lblpassw1 = GetClaveTarjeta("**")
                    lnNumOp = lnNumOp + 1
                    If lblpassw1 <> lnClaveTar Then  'And lnNumOp < 3 Then
                        MsgBox "Ingreso de Nueva Clave Erroneo" + Chr(13) + "Reintentelo (" + Str(lnNumOp) + ")", vbInformation, "Aviso"
                    End If
                Wend
                                 
                lblpassw2 = lnClaveTar
                If lblpassw1 = lblpassw2 Then
                    txtGlosa.SetFocus
                    cmdGrabar.Enabled = True
                    cmdCancelar.Enabled = True
                    GrabaTarjetaPINPAD = True
                Else
                    MsgBox "La clave no ha sido reconocida, el proceso sera Cancelado.", vbInformation, "Aviso"
                    FinalizaPinPad
                    Exit Function
                End If
                                
                Me.Caption = sCaption
                FinalizaPinPad
                cmdCancelar.Enabled = False
                
            Case gTarjNoRegistrada
                'ppoa Modificacion
                If Not WriteToLcd("Espere Por Favor") Then
                    FinalizaPinPad
                    MsgBox "No se Realizó Envío", vbInformation, "Aviso"
                    Exit Function
                End If
                MsgBox "Tarjeta no Registrada", vbInformation, "Aviso"
                
            Case gClaveNOValida
                'ppoa Modificacion
                If Not WriteToLcd("Clave Incorrecta") Then
                    FinalizaPinPad
                    MsgBox "No se Realizó Envío", vbInformation, "Aviso"
                    lblTrack1 = ""
                    Exit Function
                End If
                MsgBox "Clave Incorrecta", vbInformation, "Aviso"
                
        End Select
    
Else

                lblpassw1 = " "
                lnClaveTar = "  "
                
                lnNumOp = 0
                
                While lnNumOp < 3 And lblpassw1 <> lnClaveTar
                
                    MsgBox "Ingrese NUEVA Clave de Tarjeta.", vbInformation, "Aviso"
                    Me.Caption = "Ingrese NUEVA Clave de Tarjeta."
                    lnClaveTar = GetClaveTarjeta("*")
                    
                    MsgBox "Confirme NUEVA Clave de Tarjeta.", vbInformation, "Aviso"
                    Me.Caption = "Confirme NUEVA Clave de Tarjeta."
                    lblpassw1 = GetClaveTarjeta("**")
                    
                    lnNumOp = lnNumOp + 1
                    If lblpassw1 <> lnClaveTar Then  'And lnNumOp < 3 Then
                        MsgBox "Ingreso de Nueva Clave Erroneo" + Chr(13) + "Reintentelo (" + Str(lnNumOp) + ")", vbInformation, "Aviso"
                    End If
                Wend
                                 
                                 
                lblpassw2 = lnClaveTar
                If lblpassw1 = lblpassw2 Then
                    txtGlosa.Enabled = True
                
                    txtGlosa.SetFocus
                    cmdGrabar.Enabled = True
                    cmdCancelar.Enabled = True
                    
                    GrabaTarjetaPINPAD = True
                Else
                    MsgBox "La clave no ha sido reconocida, el proceso sera Cancelado.", vbInformation, "Aviso"
                    FinalizaPinPad
                    Exit Function
                End If
                                
                                
                Me.Caption = sCaption
                FinalizaPinPad
                cmdCancelar.Enabled = False
End If
Set clsGen = Nothing

End Function

'Private Function GetClaveTarjeta() As String
'Dim Result As Long
'Dim lsClave As String
'Result = McrReadPin(lsClave, 76, 0, "", 0, 0, "", 0, 0)
'If Result <= 0 Then
'   GetClaveTarjeta = lsClave
'Else
'    MsgBox GetErrorPINPAD(Result) & " Consulte con Servicio Tecnico.", vbInformation, "Aviso"
'    GetClaveTarjeta = ""
'End If
'End Function

Private Function GetErrorPINPAD(pnNumber As Long) As String
    GetErrorPINPAD = "Error de PINPAN, Verifique si el Programa DMONNT.EXE esta en la Barra de Tareas. Verifique la Conexión del PINPAD."
End Function

Private Sub AgregaTarjeta()
If Trim(txtGlosa) = "" Then
    MsgBox "Debe colocar la glosa correspondiente.", vbInformation, "Aviso"
    txtGlosa.SetFocus
    Exit Sub
End If
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim sTarjeta As String, sMovNro As String
Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones

Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
Set clsMov = Nothing

sTarjeta = Replace(txtTarjeta, "-", "", 1, , vbTextCompare)
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales

If clsMant.ActualizaTarjetaEstado(sTarjeta, sMovNro, gCapTarjEstCmbioClave, Trim(txtGlosa)) Then
    cmdCancelar_Click
End If
Set clsMant = Nothing
End Sub

Public Sub LimpiaPantalla()
grdCliente.Clear
grdCliente.Rows = 2
SetupGridCliente
grdClienteTarj.Clear
grdClienteTarj.Rows = 2
grdClienteTarj.FormaCabecera
grdTarjetaEstado.Clear
grdTarjetaEstado.Rows = 2
grdTarjetaEstado.FormaCabecera
txtTarjeta.Mask = "____-____-____-____"
fraTarjeta.Enabled = True
fraEstado.Enabled = False
cmdCancelar.Enabled = False
lblpassw1 = ""
lblpassw2 = ""
lblTrack1 = ""
txtGlosa = ""
cmdGrabar.Enabled = False
End Sub

Public Sub SetupGridCliente()
Dim i As Integer
For i = 1 To grdCliente.Rows - 1
    grdCliente.MergeCol(i) = True
Next i
grdCliente.MergeCells = flexMergeFree
grdCliente.BandExpandable(0) = True
grdCliente.Cols = 9
grdCliente.ColWidth(0) = 100
grdCliente.ColWidth(1) = 3500
grdCliente.ColWidth(2) = 3500
grdCliente.ColWidth(3) = 1500
grdCliente.ColWidth(4) = 1000
grdCliente.ColWidth(5) = 600
grdCliente.ColWidth(6) = 1500
grdCliente.ColWidth(7) = 0
grdCliente.ColWidth(8) = 0
grdCliente.TextMatrix(0, 1) = "Nombre"
grdCliente.TextMatrix(0, 2) = "Dirección"
grdCliente.TextMatrix(0, 3) = "Zona"
grdCliente.TextMatrix(0, 4) = "Fono"
grdCliente.TextMatrix(0, 5) = "ID"
grdCliente.TextMatrix(0, 6) = "ID N°"
End Sub

Public Function ObtieneDatosTarjeta() As Boolean
Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales 'NCapMantenimiento
Dim rsTarj As New ADODB.Recordset
Dim sTarjeta As String, sPersona As String
Dim nEstado As COMDConstantes.CaptacTarjetaEstado

sTarjeta = Replace(txtTarjeta, "-", "", 1, , vbTextCompare)
Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
Set rsTarj = clsMant.GetTarjetaCuentas(sTarjeta)
If rsTarj.EOF And rsTarj.BOF Then
    MsgBox "Tarjeta no posee ninguna relación con cuentas activas o Tarjeta no activa.", vbInformation, "Aviso"
    ObtieneDatosTarjeta = False
    Set rsTarj = Nothing
    Set clsMant = Nothing
    Exit Function
Else
    Dim nItem As Integer
    sPersona = rsTarj("cPersCod")
    Do While Not rsTarj.EOF
        grdClienteTarj.AdicionaFila
        nItem = grdClienteTarj.Rows - 1
        grdClienteTarj.TextMatrix(nItem, 1) = rsTarj("cCtaCod")
        grdClienteTarj.TextMatrix(nItem, 2) = rsTarj("Relacion")
        rsTarj.MoveNext
    Loop
    
    rsTarj.Close
    Set rsTarj = clsMant.GetDatosPersona(sPersona)
    Set grdCliente.Recordset = rsTarj
    SetupGridCliente
    
    rsTarj.Close
    Set rsTarj = clsMant.GetTarjetaEstadoHist(sTarjeta)
    Set grdTarjetaEstado.Recordset = rsTarj
    cmdCancelar.Enabled = True
    fraTarjeta.Enabled = False
    fraEstado.Enabled = True
    ObtieneDatosTarjeta = True
End If
Set rsTarj = Nothing
Set clsMant = Nothing
End Function

Private Sub BoxFirma_MagCard(ByVal timedOut As Boolean)
Dim lsPassw2 As String, lsFecha As String
Dim lsAnio As String, lsMes As String
Dim lsDia As String, sTarjeta As String
Dim i As Integer
Dim lsPassw As String

sTarjeta = Replace(txtTarjeta, "-", "", 1, , vbTextCompare)
If Len(boxFirma.MagCardTrack1) > 0 Then
    lsCodTar = GetCardNumber(boxFirma.MagCardTrack1)
    If sTarjeta = "" Then
        sTarjeta = lsCodTar
        txtTarjeta = Format$(lsCodTar, "@@@@-@@@@-@@@@-@@@@")
        ObtieneDatosTarjeta
    End If
    lblTrack1 = Format(lsCodTar, "@@@@ @@@@ @@@@ @@@@")
    lsFecha = Mid(Me.boxFirma.MagCardTrack1, 60, 6)
    lsAnio = Left(lsFecha, 4)
    lsMes = Right(lsFecha, 2)
    If lsMes = "02" Then
        lsDia = "29"
    Else
        lsDia = "30"
    End If
    DoEvents
    lsPassw = Trim(boxFirma.GetNumber("Ingrese Password", 4, 0, 60))
    Do While Len(lsPassw) < 4
        If lsPassw = "" Then
            MsgBox "Operación Cancelada por el Usuario", vbInformation, "Aviso"
            'InHabilitar
            UnSetupPad
            Exit Sub
        Else
            MsgBox "Password no Válido. Debe poseer Cuatro digitos", vbInformation, "Aviso"
            lsPassw = ""
            lsPassw = Trim(boxFirma.GetNumber("Ingrese Password", 4, 0, 60))
        End If
    Loop
    lblpassw1 = lsPassw
    DoEvents
    lsPassw2 = boxFirma.GetNumber("Confirme su Pasword", 4, 0, 60)
    i = 1
    Do While i <= 3
        If lsPassw <> lsPassw2 Then
            If i = 3 Then
                MsgBox "Numero de Reintentos Agotados. Vuelva a Realizar la Operación", vbInformation, "Aviso"
                'InHabilitar
                UnSetupPad
                Exit Sub
            Else
                i = i + 1
                MsgBox "Confirmación de Password Incorrecta. Por Favor Reintente", vbInformation, "Aviso"
                lsPassw2 = boxFirma.GetNumber("Confirme su Pasword", 4, 0, 60)
                lblpassw2 = lsPassw2
            End If
        Else
            lblpassw2 = lsPassw2
            MsgBox "Password Confirmado Correctamente.", vbInformation, "Aviso"
            'InHabilitar
            UnSetupPad
            txtGlosa.SetFocus
            Exit Do
        End If
    Loop
Else
    MsgBox "Error de Lectura de Tarjeta", vbInformation, "Aviso"
    UnSetupPad
    'InHabilitar
End If
End Sub

Private Sub cmdCancelar_Click()
boxFirma.MagCardEnabled = False
UnSetupPad
LimpiaPantalla
cmdCancelar.Enabled = False
End Sub

Private Sub cmdGrabar_Click()

If MsgBox("¿Desea grabar los cambios de la nueva clave?", vbQuestion + vbYesNo, "Aviso") = vbYes Then
    If Trim(txtGlosa) = "" Then
        MsgBox "Debe colocar la glosa correspondiente.", vbInformation, "Aviso"
        txtGlosa.SetFocus
        Exit Sub
    End If
    Dim clsMant As COMNCaptaGenerales.NCOMCaptaGenerales
    Dim sTarjeta As String, sMovNro As String
    Dim clsMov As COMNContabilidad.NCOMContFunciones 'NContFunciones
    Dim sClave As String
    Dim CLSSERV As COMNCaptaServicios.NCOMCaptaServicios 'NCapServicios
    Set CLSSERV = New COMNCaptaServicios.NCOMCaptaServicios
    Dim lscadimp As String
    Dim loPrevio As previo.clsPrevio
    
    sClave = Encripta(Trim(lblpassw1.Text), True)
    Set clsMov = New COMNContabilidad.NCOMContFunciones
    sMovNro = clsMov.GeneraMovNro(gdFecSis, Right(gsCodAge, 2), gsCodUser)
    Set clsMov = Nothing
    
    sTarjeta = Replace(txtTarjeta, "-", "", 1, , vbTextCompare)
    Set clsMant = New COMNCaptaGenerales.NCOMCaptaGenerales
    
    If clsMant.ActualizaTarjetaEstado(sTarjeta, sMovNro, gCapTarjEstCmbioClave, Trim(txtGlosa), sClave) Then
        lscadimp = CLSSERV.ImprimeBolTarjeta("CAMBIO DE CLAVE TARJETA", _
                                            Trim(grdCliente.TextMatrix(1, 1)), txtTarjeta.Text, _
                                            "TARJEA MAGNETICA", gdFecSis, gsNomAge, _
                                            gsCodUser, sLpt)
        Do
           Set loPrevio = New previo.clsPrevio
             loPrevio.PrintSpool sLpt, lscadimp, False
             loPrevio.PrintSpool sLpt, Chr(10) & Chr(10) & Chr(10) & Chr(10) & lscadimp, False
           Set loPrevio = Nothing
           
        Loop Until MsgBox("DESEA REIMPRIMIR BOLETA?", vbYesNo, "AVISO") = vbNo

        cmdCancelar_Click
    End If
    Set clsMant = Nothing
    Set CLSSERV = Nothing
End If
End Sub

Private Sub cmdSalir_Click()
Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
Me.Icon = LoadPicture(App.path & gsRutaIcono)
Dim lsNumTar As String
If KeyCode = vbKeyF11 Then 'F11
    Dim nCOM As TipoPuertoSerial
    Dim clsGen As COMDConstSistema.DCOMGeneral
    Dim sMaquina As String
    sMaquina = GetComputerName
    cmdCancelar.Enabled = True
    'opciones de validacion
    Set clsGen = New COMDConstSistema.DCOMGeneral
    nCOM = clsGen.GetPuertoPeriferico(gPerifPENWARE, sMaquina)
    If nCOM = -1 Then
        '*********
       ' GnTipoPinPad = ObtieneTipoPinPad()
        Set clsGen = New COMDConstSistema.DCOMGeneral
            nCOM = clsGen.GetPuertoPeriferico(COMDConstantes.gPerifPINPAD, sMaquina)
        Set clsGen = Nothing
        '*********
            If Not GmyPSerial Is Nothing Then
                GmyPSerial.Disconnect
                Set GmyPSerial = Nothing
            End If
            Set GmyPSerial = CreateObject("HComPinpad.Pinpad")
            If GmyPSerial.ConnectionTest = 0 Then
                Call GmyPSerial.Connect(CInt(nCOM), 9600)
                If GmyPSerial.ConnectionTest = 1 Then
                     If GmyPSerial.ReadCardIniConf("PASE SU TARJETA") = 1 Then
                            lsNumTar = GetNumTarjeta_Vrf5000
                            'lsNumTarTemp = lsNumTar
                            If Len(lsNumTar) <> 16 Then
                                MsgBox "Error en la Lectura de Tarjeta.", vbInformation, "Aviso"
                                GmyPSerial.Disconnect
                                Set GmyPSerial = Nothing
                                Exit Sub
                            End If
                            'txtTarjeta.Text = Format(lsNumTar, "####-####-####-####")
                            txtTarjeta.Mask = LTrim(lsNumTar)
                            If ObtieneDatosTarjeta = False Then
                                MsgBox "Tarjeta no Valida", vbInformation, "AVISO"
                            Else
                                'ObtieneDatosTarjeta
                                MsgBox "PROCEDA A INGRESAR SU CLAVE...", vbInformation, "AVISO"
                                Me.Caption = "Ingrese la Clave de la Tarjeta."
                                Call GrabaTarjetaPINPADV_5000(lsNumTar, nCOM)
                            End If
                     End If
                End If
            End If
            If Not GmyPSerial Is Nothing Then
                GmyPSerial.Disconnect
                Set GmyPSerial = Nothing
            End If
                
        '*****************************
'        'ppoa Modificacion
'        If Not IniciaPinPad(nCOM) Then
'            MsgBox "No Inicio Dispositivo" & ". Consulte con Servicio Tecnico.", vbInformation, "Aviso"
'            FinalizaPinPad
'            Exit Sub
'        End If
'         If Not GrabaTarjetaPINPAD Then
'            cmdcancelar_Click
'        End If
'        FinalizaPinPad
        
    Else
        If ConectaPad Then
            boxFirma.MagCardEnabled = False
            MuestraPantalla
        End If
    End If
End If
End Sub

Public Function GetNumTarjeta_Vrf5000_2() As String

    Dim lsNumTarTemp As String
    Dim lsNumTar As String
    
    
    GmyPSerial.ReadCardIni
    While lsNumTar = ""
        lsNumTar = GmyPSerial.ReadCard
        If GmyPSerial.ReadCardIni = 1 Then
            DoEvents
            lsNumTar = GmyPSerial.ReadCard
        End If
        DoEvents
    Wend
    'Debug.Print lsNumTar
    lsNumTarTemp = lsNumTar
    
    If IsNumeric(Left(lsNumTarTemp, 1)) Then
        lsNumTar = Trim(Mid(lsNumTarTemp, 1, 16))
    ElseIf Not IsNumeric(Left(lsNumTarTemp, 1)) And IsNumeric(Mid(lsNumTarTemp, 2, 1)) Then
        lsNumTar = Trim(Mid(lsNumTarTemp, 2, 16))
    End If
    
    If Not IsNumeric(lsNumTar) Then
      lsNumTar = Trim(Mid(lsNumTarTemp, 3, 16))
    End If
    
    GetNumTarjeta_Vrf5000_2 = lsNumTar
    
End Function
Private Sub Form_Load()
Me.Caption = "Captaciones - Tarjeta - Cambio de Clave"
LimpiaPantalla
End Sub


Private Sub txtGlosa_KeyPress(KeyAscii As Integer)
KeyAscii = fgIntfMayusculas(KeyAscii)
If KeyAscii = 13 Then
    cmdGrabar.SetFocus
End If
End Sub

Private Sub txtTarjeta_GotFocus()
With txtTarjeta
    .SelStart = 0
    .SelLength = Len(.Text)
End With
End Sub

Private Sub txtTarjeta_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
    ObtieneDatosTarjeta
End If
End Sub


