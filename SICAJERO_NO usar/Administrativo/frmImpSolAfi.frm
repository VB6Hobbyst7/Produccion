VERSION 5.00
Begin VB.Form frmImpSolAfi 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Impresion de Solicitud de Vinculacion"
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "frmImpSolAfi.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5520
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   660
      Left            =   0
      TabIndex        =   3
      Top             =   1170
      Width           =   5505
      Begin VB.CommandButton CmdSalir 
         Caption         =   "Salir"
         Height          =   360
         Left            =   4110
         TabIndex        =   5
         Top             =   180
         Width           =   1275
      End
      Begin VB.CommandButton CmdImprimir 
         Caption         =   "Imprimir"
         Height          =   360
         Left            =   75
         TabIndex        =   4
         Top             =   180
         Width           =   1275
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1170
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5505
      Begin VB.TextBox TxtFecAprob 
         Height          =   300
         Left            =   1800
         TabIndex        =   7
         Top             =   750
         Width           =   1275
      End
      Begin VB.TextBox TxtNumTarj 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   405
         Left            =   780
         MaxLength       =   16
         TabIndex        =   1
         Top             =   225
         Width           =   3240
      End
      Begin VB.Label Label1 
         Caption         =   "Fecha de Aprobacion :"
         Height          =   285
         Left            =   75
         TabIndex        =   6
         Top             =   795
         Width           =   1710
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
Attribute VB_Name = "frmImpSolAfi"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim R1 As ADODB.Recordset
Dim R2 As ADODB.Recordset
Dim oConec As DConecta

Private Sub ReemplazaValores(ByRef mWord As Word.Application, ByVal CadBuscar As String, ByVal CadReemplazar As String)
   Call mWord.Selection.Find.ClearFormatting
   Call mWord.Selection.Find.Replacement.ClearFormatting
   
    With mWord.Selection.Find
        .Text = CadBuscar
        .Replacement.Text = CadReemplazar
        .Forward = True
        .Wrap = 1
        .Format = False
        .MatchCase = False
        .MatchWholeWord = True
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With

    'If Len(Trim(CadReemplazar)) > 0 Then
        Call mWord.Selection.Find.Execute(Replace:=2)
    'End If
End Sub

Private Sub ImprimirDocumentos()

        Dim mWord As Word.Application
        Dim mDoc As Word.Document
        
        Set mWord = New Word.Application

        'mDoc = mWord.Documents.Open(System.Windows.Forms.Application.StartupPath & "\DocumentosWIN.doc")
        Set mDoc = mWord.Documents.Open(App.Path & "\docs\Anexo1.doc")

        mWord.Visible = False
        mDoc.SaveAs (App.Path & "\" & Me.TxtNumTarj.Text & ".doc")
        
        Call CargaDatos

        Call ReemplazaValores(mWord, "<<NOMBRE>>", R1!cPersNombre)
        Call ReemplazaValores(mWord, "<<DNI>>", R1!cDNI)
        Call ReemplazaValores(mWord, "<<DIRECCION>>", R1!cPersDireccDomicilio & " - " & R1!cUbigeoDescripcion)
        Call ReemplazaValores(mWord, "<<FECHA>>", Me.TxtFecAprob.Text)
        Call ReemplazaValores(mWord, "<<FECHAN>>", ArmaFecha(TxtFecAprob.Text))
        Call ReemplazaValores(mWord, "<<NUMTARJETA>>", Me.TxtNumTarj.Text)
        
        Dim i As Integer
        i = 1
        Do While Not R2.EOF
            Call ReemplazaValores(mWord, "<<CUENTA" & Trim(Str(i)) & ">>", R2!Cuenta & Space(4))
            Call ReemplazaValores(mWord, "<<TIPO" & Trim(Str(i)) & ">>", Left(R2!TipoCta & Space(20), 20))
            Call ReemplazaValores(mWord, "<<MONEDA" & Trim(Str(i)) & ">>", Left(R2!cMoneda & Space(10), 10))
            R2.MoveNext
            i = i + 1
        Loop

        Do While i <= 8
            Call ReemplazaValores(mWord, "<<CUENTA" & Trim(Str(i)) & ">>", " " & Space(4))
            Call ReemplazaValores(mWord, "<<TIPO" & Trim(Str(i)) & ">>", " ")
            Call ReemplazaValores(mWord, "<<MONEDA" & Trim(Str(i)) & ">>", " ")
            i = i + 1
        Loop

       
        Call mDoc.Close

        Set mWord = New Word.Application

        mWord.Visible = True
        Set mDoc = mWord.Documents.Open(App.Path & "\" & Me.TxtNumTarj.Text & ".doc")

        Call EliminaObjeto(mDoc)
        Call EliminaObjeto(mWord)
        R1.Close
        R2.Close
        Set R1 = Nothing
        Set R2 = Nothing

    End Sub
        Public Sub EliminaObjeto(ByRef Objeto As Object)
        On Error Resume Next
            'Bucle de eliminacion
'            Do Until _
'                 System.Runtime.InteropServices.Marshal.ReleaseComObject(Objeto) <= 0
'            Loop

            Objeto = Nothing
    End Sub
    
Private Sub CargaDatos()
Dim Cmd As New Command
Dim Cmd2 As New Command
Dim Prm As New ADODB.Parameter

    Set R1 = New ADODB.Recordset
    
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd.CreateParameter("@psNumTarj", adVarChar, adParamInput, 20, Me.TxtNumTarj.Text)
    Cmd.Parameters.Append Prm
    
    oConec.AbreConexion
    Cmd.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd.CommandType = adCmdStoredProc
    
    Cmd.CommandText = "ATM_RecuperaDatosImpSolAfil"
    
    R1.CursorType = adOpenStatic
    R1.LockType = adLockReadOnly
    Set R1 = Cmd.Execute
    
    
    Set R2 = New ADODB.Recordset
    Set Prm = New ADODB.Parameter
    Set Prm = Cmd2.CreateParameter("@psNumTarjeta", adVarChar, adParamInput, 20, Me.TxtNumTarj.Text)
    Cmd2.Parameters.Append Prm
    
    Cmd2.ActiveConnection = oConec.ConexionActiva ' AbrirConexion
    Cmd2.CommandType = adCmdStoredProc
    
    Cmd2.CommandText = "ATM_RecuperaCuentasReporteAfil"
    
    R2.CursorType = adOpenStatic
    R2.LockType = adLockReadOnly
    Set R2 = Cmd2.Execute
    
    
End Sub


Private Sub CmdImprimir_Click()
If Len(Trim(Me.TxtNumTarj.Text)) = 0 Then
    MsgBox ("Numero de Tarjeta Invalida")
    Exit Sub
End If

If Len(Trim(Me.TxtFecAprob.Text)) = 0 Then
    MsgBox ("Fecha de Aprobacion Invalida")
    Exit Sub
End If

 Screen.MousePointer = 11

 Call ImprimirDocumentos
 
 Screen.MousePointer = 0
 MsgBox ("Por favor Cerrar Microsoft Word despues de imprimir la Solicitud")
 
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Set oConec = New DConecta
    Me.TxtFecAprob.Text = Format(Now, "dd/mm/yyyy")
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set oConec = Nothing
End Sub
