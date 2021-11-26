VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmDaot 
   Caption         =   "DAOT"
   ClientHeight    =   2595
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   3405
   LinkTopic       =   "Form1"
   ScaleHeight     =   2595
   ScaleWidth      =   3405
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkIngresos 
      Caption         =   "Ingresos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1755
      TabIndex        =   6
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CheckBox chkCostos 
      Caption         =   "Costos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   555
      TabIndex        =   5
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "&Generar"
      Height          =   400
      Left            =   422
      TabIndex        =   4
      Top             =   2040
      Width           =   1000
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      Height          =   1335
      Left            =   375
      TabIndex        =   1
      Top             =   45
      Width           =   2655
      Begin VB.TextBox txtTipCamb 
         Height          =   315
         Left            =   1320
         TabIndex        =   8
         Top             =   780
         Width           =   975
      End
      Begin MSMask.MaskEdBox txtAnio 
         Height          =   315
         Left            =   1320
         TabIndex        =   2
         Top             =   255
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         MaxLength       =   4
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label2 
         Caption         =   "Tipo Cambio"
         Height          =   375
         Left            =   120
         TabIndex        =   7
         Top             =   720
         Width           =   855
      End
      Begin VB.Label Label1 
         Caption         =   "Año :"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   3
         Top             =   285
         Width           =   615
      End
   End
   Begin VB.CommandButton cdmSalir 
      Caption         =   "&Salir"
      Height          =   400
      Left            =   1982
      TabIndex        =   0
      Top             =   2040
      Width           =   1000
   End
End
Attribute VB_Name = "frmDaot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cdmSalir_Click()
    Unload Me
End Sub

Private Sub cmdGenerar_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim oReg As NContImpreReg
    Set oReg = New NContImpreReg
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim lsAnio As String
    Dim lnTipCamb As Double
    Dim i As String
    Dim lcText As String
    Dim lsTipPer As String
    Dim lsTipDoc As String
    Dim lsNumDoc As String
    Dim lsApePat As String
    Dim lsApeMat As String
    Dim lsNombre1 As String
    Dim lsNombre2 As String
    Dim lsRazSoc As String
    Dim lnOpc As Integer
    Dim lnOpc1 As Integer
    Dim lsNomcli As String
    Dim oSys As Scripting.FileSystemObject
    Dim oText As Scripting.TextStream

    'If oSys.FileExists(App.path & "/SPOOLER/") = False Then
    Set oSys = New Scripting.FileSystemObject
    
    'End If

    
'En lugar de lo de arriba puedo cambiar el último parametro para modificarlo.
    'Set oText = oSys.OpenTextFile(rutafichero, ForWriting)

    

    lsAnio = Me.txtAnio.Text
    lnTipCamb = Me.txtTipCamb.Text
    i = 0
    If Me.chkCostos.value = 1 Then
        oSys.CreateTextFile (App.path & "\SPOOLER\Costos.txt")
        Set oText = oSys.OpenTextFile(App.path & "\SPOOLER\Costos.txt", ForWriting)
        oCon.AbreConexion
        Set rs = oReg.GetDaotCostos(lsAnio, lnTipCamb)
        oCon.CierraConexion
        Do While Not rs.EOF
            i = i + 1
            lsTipPer = rs!Personeria
            lsTipDoc = rs!TpoDoc
            lsNumDoc = Trim(rs!cPersIdNro)
            lsMonto = Trim(Str(Round(rs!Monto, 0)))
            If (rs!TpoDoc = "6" Or rs!TpoDoc = "-" Or rs!TpoDoc = "0") And InStr(rs!cPersNombre, "/") = 0 Then
                lsRazSoc = Trim(rs!cPersNombre)
                lsApePat = ""
                lsApeMat = ""
                lsNombre1 = ""
                lsNombre2 = ""
            Else
                lsApePat = Mid(rs!cPersNombre, 1, InStr(rs!cPersNombre, "/") - 1)
                lnOpc = InStr(rs!cPersNombre, ",") - InStr(rs!cPersNombre, "/")
                lsApeMat = Mid(rs!cPersNombre, InStr(rs!cPersNombre, "/") + 1, lnOpc - 1)
                lsNomcli = Mid(rs!cPersNombre, InStr(rs!cPersNombre, ",") + 1)
                lnOpc1 = Len(lsNomcli) - InStr(lsNomcli, " ")
                If InStr(lsNomcli, " ") = 0 Then
                    lsNombre1 = Right(lsNomcli, lnOpc1)
                    lsNombre2 = ""
                 Else
                    lsNombre1 = Mid(lsNomcli, 1, InStr(lsNomcli, " ") - 1)
                    lsNombre2 = Trim(Right(lsNomcli, lnOpc1))
                End If
                lsRazSoc = ""
            End If
            lcText = lcText & GeneraArchivotxt(i, lsAnio, lsTipPer, lsTipDoc, lsNumDoc, lsMonto, lsApePat, lsApeMat, lsNombre1, lsNombre2, lsRazSoc)
            rs.MoveNext
        Loop
        oText.WriteLine lcText
        oText.Close
    End If
    
    i = 0
    If Me.chkIngresos.value = 1 Then
        oSys.CreateTextFile (App.path & "\SPOOLER\Ingresos.txt")
        Set oTextI = oSys.OpenTextFile(App.path & "\SPOOLER\Ingresos.txt", ForWriting)
        i = i + 1
        oCon.AbreConexion
        Set rs = oReg.GetDaotIngresos(lsAnio, lnTipCamb)
        oCon.CierraConexion
        Do While Not rs.EOF
            i = i + 1
            lsTipPer = rs!Personeria
            lsTipDoc = rs!TpoDoc
            lsNumDoc = Trim(rs!cPersIdNro)
            lsMonto = Trim(Str(Round(rs!Monto, 0)))
            If (rs!TpoDoc = "6" Or rs!TpoDoc = "-" Or rs!TpoDoc = "0") And InStr(rs!cPersNombre, "/") = 0 Then
                lsRazSoc = Trim(rs!cPersNombre)
                lsApePat = ""
                lsApeMat = ""
                lsNombre1 = ""
                lsNombre2 = ""
            Else
                lsApePat = Mid(rs!cPersNombre, 1, InStr(rs!cPersNombre, "/") - 1)
                lnOpc = InStr(rs!cPersNombre, ",") - InStr(rs!cPersNombre, "/")
                lsApeMat = Mid(rs!cPersNombre, InStr(rs!cPersNombre, "/") + 1, lnOpc - 1)
                lsNomcli = Mid(rs!cPersNombre, InStr(rs!cPersNombre, ",") + 1)
                lnOpc1 = Len(lsNomcli) - InStr(lsNomcli, " ")
                If InStr(lsNomcli, " ") = 0 Then
                    lsNombre1 = Right(lsNomcli, lnOpc1)
                    lsNombre2 = ""
                 Else
                    lsNombre1 = Mid(lsNomcli, 1, InStr(lsNomcli, " ") - 1)
                    lsNombre2 = Trim(Right(lsNomcli, lnOpc1))
                End If
                lsRazSoc = ""
            End If
            lcText = lcText & GeneraArchivotxt(i, lsAnio, lsTipPer, lsTipDoc, lsNumDoc, lsMonto, lsApePat, lsApeMat, lsNombre1, lsNombre2, lsRazSoc)
            rs.MoveNext
        Loop
        oTextI.WriteLine lcText
        oTextI.Close
    End If

    
End Sub

Private Sub Form_Load()
   CentraForm Me
   txtAnio = Year(gdFecSis)
End Sub
Public Function GeneraArchivotxt(ByVal i As String, ByVal psAnio As String, ByVal lsTipPer As String, ByVal lsTipoDoc As String, ByVal lsNumDoc As String, _
                                 ByVal lsMonto As String, ByVal lsApePat As String, ByVal lsApeMat As String, ByVal lsNombre1 As String, _
                                 ByVal lsNombre2 As String, ByVal lsRazSoc As String) As String
GeneraArchivotxt = i & "|" & "6" & "|" & gsRUC & "|" & psAnio & "|" & lsTipPer & "|" & _
                   lsTipoDoc & "|" & lsNumDoc & "|" & lsMonto & "|" & lsApePat & "|" & _
                   lsApeMat & "|" & lsNombre1 & "|" & lsNombre2 & "|" & lsRazSoc & "|" & Chr(10)
End Function

