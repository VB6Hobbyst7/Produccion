VERSION 5.00
Begin VB.Form frmAfiliaLote 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generacion de Archivo de Afiliacion en Lote"
   ClientHeight    =   2130
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4455
   Icon            =   "frmAfiliaLote.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2130
   ScaleWidth      =   4455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      Height          =   555
      Left            =   75
      TabIndex        =   10
      Top             =   990
      Width           =   4335
      Begin VB.Label LblNumReg 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "0"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2235
         TabIndex        =   12
         Top             =   165
         Width           =   780
      End
      Begin VB.Label Label4 
         Caption         =   "Numero de Registros :"
         Height          =   255
         Left            =   105
         TabIndex        =   11
         Top             =   180
         Width           =   1725
      End
   End
   Begin VB.Frame Frame2 
      Height          =   645
      Left            =   75
      TabIndex        =   7
      Top             =   1485
      Width           =   4335
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   330
         Left            =   3030
         TabIndex        =   9
         Top             =   180
         Width           =   1215
      End
      Begin VB.CommandButton CmdGenerar 
         Caption         =   "Generar Archivo"
         Height          =   330
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   1620
      End
   End
   Begin VB.Frame Frame1 
      Height          =   945
      Left            =   75
      TabIndex        =   0
      Top             =   45
      Width           =   4335
      Begin VB.TextBox TxtPanFin 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3000
         MaxLength       =   9
         TabIndex        =   6
         Top             =   555
         Width           =   1185
      End
      Begin VB.TextBox TxtPanIni 
         Appearance      =   0  'Flat
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   3000
         MaxLength       =   9
         TabIndex        =   3
         Top             =   195
         Width           =   1185
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "810900"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2235
         TabIndex        =   5
         Top             =   555
         Width           =   780
      End
      Begin VB.Label Label2 
         Caption         =   "Numero de Tarjeta de Final :"
         Height          =   300
         Left            =   135
         TabIndex        =   4
         Top             =   570
         Width           =   2055
      End
      Begin VB.Label LblID 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "810900"
         ForeColor       =   &H00800000&
         Height          =   285
         Left            =   2235
         TabIndex        =   2
         Top             =   195
         Width           =   780
      End
      Begin VB.Label Label1 
         Caption         =   "Numero de Tarjeta de Inicio :"
         Height          =   300
         Left            =   135
         TabIndex        =   1
         Top             =   210
         Width           =   2055
      End
   End
End
Attribute VB_Name = "frmAfiliaLote"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim oConec As DConecta

Private Sub CmdGenerar_Click()
Dim sCadLote As String
Dim sCadHeader As String
Dim sCadTrailer As String
Dim sCadDatos As String
Dim i As Integer
Dim PAN As String
Dim Cmd As New Command
Dim Prm As New ADODB.Parameter


    If CInt(Me.LblNumReg.Caption) <= 0 Then
        Call MsgBox("El Numero de Registros a Generar debe ser mayor que Cero", vbCritical)
        Exit Sub
    End If

    sCadHeader = "A"
    sCadHeader = sCadHeader & Right("00000000" & Me.LblNumReg.Caption, 8)
    sCadHeader = sCadHeader & "00000426154"
    sCadHeader = sCadHeader & "0000"
    sCadHeader = sCadHeader & Right(Str(Year(Now)), 2) & Right("00" & Trim(Str(Month(Now))), 2) & Right("00" & Trim(Str(Day(Now))), 2)
    sCadHeader = sCadHeader & "00000000"
    sCadHeader = sCadHeader & String(1125 - Len(sCadHeader), " ")


    sCadLote = ""

    For i = CInt(Me.TxtPanIni.Text) To CInt(Me.TxtPanFin.Text)
        sCadDatos = "B0302C2200001800100000000003008000001"
        'LSDO 20080825
        PAN = gsBIN & Right("000000000" & Trim(Str(i)), 9)
        PAN = PAN + DigitoChequeo(PAN)
        'sCadDatos = sCadDatos & "16426154" & Right("0000000000" & Trim(Str(i)), 10)
        sCadDatos = sCadDatos & "16" & PAN
        'FIN LSDO
        sCadDatos = sCadDatos & Right("00" & Trim(Str(Month(Now))), 2) & Right("00" & Trim(Str(Day(Now))), 2) & Right("00" & Trim(Str(Hour(Now))), 2) & Right("00" & Trim(Str(Minute(Now))), 2) & Right("00" & Trim(Str(Second(Now))), 2)
        sCadDatos = sCadDatos & Right("00" & Trim(Str(Hour(Now))), 2) & Right("00" & Trim(Str(Minute(Now))), 2) & Right("00" & Trim(Str(Second(Now))), 2)
        sCadDatos = sCadDatos & "0642615406426154413"
        sCadDatos = sCadDatos & GeneraTrama0200(Now, "INNOMINADA", "INNOMINADA ", "INNOMINADA ", "12345678", "12345678", "M", CDate("01/01/1900"), "S", gsCodAge, "12345678", gsCodCiudad)
        sCadDatos = sCadDatos & "000000000000000000000000000000000000"
        sCadDatos = sCadDatos & String(1125 - Len(sCadDatos), " ")
        
        sCadLote = sCadLote & sCadDatos
        
        Set Cmd = New ADODB.Command
        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cNumTarjeta", adVarChar, adParamInput, 50, PAN)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nCondicion", adInteger, adParamInput, , -1)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nRetenerTarjeta", adInteger, adParamInput, , 0)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@nCodAge", adInteger, adParamInput, , gsCodAge)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cUserAfil", adVarChar, adParamInput, 50, gsCodUser)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@dFecAfil", adDate, adParamInput, , Now)
        Cmd.Parameters.Append Prm

        Set Prm = New ADODB.Parameter
        Set Prm = Cmd.CreateParameter("@cPersCod", adVarChar, adParamInput, 50, "")
        Cmd.Parameters.Append Prm

        oConec.AbreConexion
        Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
        Cmd.CommandType = adCmdStoredProc
        Cmd.CommandText = "ATM_RegistraTarjeta"
        Cmd.Execute
        
        oConec.CierraConexion
        'MsgBox "Orden de Afiliación de Tarjeta Con Exito"
    Next i

    sCadTrailer = "Z" & Right("00000000" & Me.LblNumReg.Caption, 8)
    sCadTrailer = sCadTrailer & String(1125 - Len(sCadTrailer), " ")
    
    sCadLote = sCadHeader & sCadLote & sCadTrailer
    
    Dim X As Integer
    X = FreeFile
    Open App.Path & "\AfiliacionLote_" & Format(Now, "ddmmyyyyhhmmss") & ".txt" For Output As X
    Print #X, sCadLote
    Close X
    
    MsgBox "Archivo Generado"

End Sub

Private Sub CmdSalir_Click()
        Unload Me
End Sub


Private Sub Form_Load()
    Set oConec = New DConecta
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub

Private Sub TxtPanFin_Change()
If Not IsNumeric(Me.TxtPanIni.Text) Then
        Me.TxtPanIni.Text = 0
    End If
    If Not IsNumeric(Me.TxtPanFin.Text) Then
        Me.TxtPanFin.Text = 0
    End If
    
    Me.LblNumReg.Caption = (CInt(Me.TxtPanFin.Text) - CInt(Me.TxtPanIni.Text)) + 1
    
End Sub

Private Sub TxtPanIni_Change()
    If Not IsNumeric(Me.TxtPanIni.Text) Then
        Me.TxtPanIni.Text = 0
    End If
    If Not IsNumeric(Me.TxtPanFin.Text) Then
        Me.TxtPanFin.Text = 0
    End If
    
    Me.LblNumReg.Caption = (CInt(Me.TxtPanFin.Text) - CInt(Me.TxtPanIni.Text)) + 1
    
End Sub
