VERSION 5.00
Begin VB.Form frmMuestraFirma 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Firma del Cliente"
   ClientHeight    =   5280
   ClientLeft      =   600
   ClientTop       =   2085
   ClientWidth     =   9555
   Icon            =   "frmMuestraFirma.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   9555
   ShowInTaskbar   =   0   'False
   Begin SICMACT.ImageDB ImageDB1 
      Height          =   4290
      Left            =   225
      TabIndex        =   5
      Top             =   600
      Width           =   8940
      _ExtentX        =   15769
      _ExtentY        =   7567
      Enabled         =   0   'False
   End
   Begin VB.PictureBox IDBFirma 
      Enabled         =   0   'False
      Height          =   1680
      Left            =   225
      ScaleHeight     =   1620
      ScaleWidth      =   9180
      TabIndex        =   4
      Top             =   5400
      Width           =   9240
   End
   Begin VB.CommandButton cmdImprimir 
      Caption         =   "&Imprimir"
      Height          =   375
      Left            =   4020
      TabIndex        =   3
      Top             =   75
      Visible         =   0   'False
      Width           =   1020
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   375
      Left            =   2940
      TabIndex        =   0
      Top             =   75
      Width           =   1020
   End
   Begin VB.OptionButton OptFirma 
      Caption         =   "Firma y Foto"
      Enabled         =   0   'False
      Height          =   285
      Index           =   1
      Left            =   1395
      TabIndex        =   2
      Top             =   120
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.OptionButton OptFirma 
      Caption         =   "DNI"
      Height          =   285
      Index           =   0
      Left            =   135
      TabIndex        =   1
      Top             =   120
      Value           =   -1  'True
      Width           =   1125
   End
   Begin VB.Image ImgPrueba 
      Height          =   165
      Left            =   300
      Top             =   5025
      Width           =   1065
   End
End
Attribute VB_Name = "frmMuestraFirma"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public psCodCli As String
Public rs As ADODB.Recordset

Private Sub CambiaAFormato(ByVal pbSoloFirma As Boolean)
    If pbSoloFirma Then
        Me.Height = 4905
        Me.Width = 5820
        IDBFirma.Height = 3945
        IDBFirma.Width = 5550
    Else
        Me.Height = 7000
        Me.Width = 4110
        IDBFirma.Height = 6120
        IDBFirma.Width = 3855
    End If
    CentraSdi Me
End Sub

Private Sub cmdImprimir_Click()
'    Dim x As Printer
'    Dim sql As String
'    Dim rs As ADODB.Recordset
'    Set rs = New ADODB.Recordset
'
'    frmImpresora.Show 1
'
'    sql = " Select TCNat.cNomTab Nat, TCJur.cNomTab Jur, cNudoci, cNudotr From " & gcCentralPers & "Persona PE Left Join " & gcCentralCom & "Tablacod TCNat On PE.cTidoci = TCNat.cValor And TCNat.cCodTab Like '04__' Left Join " & gcCentralCom & "Tablacod TCJur On PE.cTidotr = TCJur.cValor And TCJur.cCodTab Like '05__' Where PE.cCodPers = '" & lsCodPers & "'"
'    rs.Open sql, dbCmact, adOpenForwardOnly, adLockReadOnly, adCmdText
'
'    For Each x In Printers
'        If x.DeviceName = sLpt Then
'            Set Printer = x
'        End If
'    Next
'
'    Printer.ScaleMode = 6 'milimetro
'
'    Printer.PaintPicture PicFirma.Picture, 30, 0
'    Printer.Line (10, 50)-(100, 50)
'    Printer.Print ""
'    Printer.Print ""
'    Printer.Print "             CLIENTE    : " & lblCliente.Caption
'    Printer.Print "             DOCUMENTO    : " & IIf(IsNull(rs!cNudoci), rs!cNudoTr & "", rs!cNudoci)
'
'    rs.Close
'
'    Printer.EndDoc
End Sub

Private Sub cmdsalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraSdi Me
    Call MuestraFirma(rs)
    
    Me.Icon = LoadPicture(App.path & gsRutaIcono)
End Sub

Private Sub MuestraFirma(ByVal oRs As ADODB.Recordset)
   'Dim oRs As New Recordset
   Dim oST As Stream
    Const sFile1 = "\Temp1.jpg"
    Screen.MousePointer = vbArrowHourglass
   ' oRs.Open "SELECT * FROM DAT_LogosEmpresas WHERE IdRowEmpresa=" &
    'pIdRowEmp , oCon
    If Not oRs.EOF Then
        If Not IsNull(oRs("iPersFirma")) Then
            Set oST = New Stream
            oST.Type = adTypeBinary
            oST.Open
            oST.Write oRs.Fields("iPersFirma").value
            oST.SaveToFile App.path & sFile1, adSaveCreateOverWrite
            'Me.iEnc.Image = LoadPicture(App.path & sFile1)
            On Error GoTo err2
            Me.ImgPrueba.Picture = LoadPicture(App.path & sFile1)
            
            Kill (App.path & sFile1)
            oST.Close
            Set oST = Nothing
        End If
    End If
    oRs.Close
    Screen.MousePointer = vbDefault
    Exit Sub
err2:
    Screen.MousePointer = vbDefault
    MsgBox "La visualización del DNI no esta Disponible", vbOKOnly + vbInformation, "AVISO"

End Sub

Private Sub OptFirma_Click(Index As Integer)
Dim R As ADODB.Recordset
Dim sSql As String
        
'    If Index = 0 Then
'        CambiaAFormato True
'    Else
'        CambiaAFormato False
'    End If
'    AbreConexionFirmas
'    sSql = "Select cCodPers, iFirma From " & gcDBImg & ".dbo.Firma where cCodPers = '" & psCodCli & "'"
'    Set R = New ADODB.Recordset
'    R.Open sSql, DBCmactFirmas, adOpenKeyset, adLockOptimistic, adCmdText
'    If R.RecordCount > 0 Then
'        Call IDBFirma.CargarFirma(R)
'    End If
'    R.Close
'    Set R = Nothing
'    CierraConexionFirmas
'    IDBFirma.SetFocus


End Sub
