VERSION 5.00
Begin VB.Form frmCredEduacionGeneraTxt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Generacion de Archivo de Texto para Educacion"
   ClientHeight    =   2715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8070
   Icon            =   "frmCredEduacionGeneraTxt.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2715
   ScaleWidth      =   8070
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox CboAgencia 
      Height          =   315
      Left            =   1215
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   630
      Width           =   4785
   End
   Begin VB.CommandButton CmdSalir 
      Cancel          =   -1  'True
      Caption         =   "&Salir"
      Height          =   315
      Left            =   6330
      TabIndex        =   4
      Top             =   2160
      Width           =   1440
   End
   Begin VB.CommandButton CmdGeneraTxt 
      Caption         =   "Generar Archivo de Texto"
      Height          =   345
      Left            =   990
      TabIndex        =   3
      Top             =   2160
      Width           =   4995
   End
   Begin VB.CommandButton CmdAplicar 
      Caption         =   "Aplicar"
      CausesValidation=   0   'False
      Height          =   420
      Left            =   6225
      TabIndex        =   2
      Top             =   180
      Width           =   1440
   End
   Begin VB.ComboBox CmbInstitucion 
      Height          =   315
      Left            =   1230
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   225
      Width           =   4785
   End
   Begin VB.Label LblAgencia 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "AGENCIA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   1215
      TabIndex        =   8
      Top             =   1545
      Width           =   4770
   End
   Begin VB.Label LblInst 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "INSTITUCION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   390
      Left            =   1230
      TabIndex        =   7
      Top             =   1095
      Width           =   4770
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Agencia :"
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
      Height          =   195
      Left            =   150
      TabIndex        =   6
      Top             =   690
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Institucion : "
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
      Height          =   195
      Left            =   165
      TabIndex        =   1
      Top             =   285
      Width           =   1080
   End
End
Attribute VB_Name = "frmCredEduacionGeneraTxt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CargaControles()
 Dim oCredito As COMDCredito.DCOMCredito
 Set oCredito = New COMDCredito.DCOMCredito
 Dim rsInstit As ADODB.Recordset
 Dim lrAgenc As ADODB.Recordset
 Dim loCargaAg As COMDColocPig.DCOMColPFunciones
 Dim rsTipoPago As ADODB.Recordset
 
    Call oCredito.CargarControlesPagoLote(rsInstit, rsTipoPago)
    Set oCredito = Nothing
    
    Set loCargaAg = New COMDColocPig.DCOMColPFunciones
    Set lrAgenc = loCargaAg.dObtieneAgencias(True)
    Set loCargaAg = Nothing
    
    CmbInstitucion.Clear
    Do While Not rsInstit.EOF
        CmbInstitucion.AddItem PstaNombre(rsInstit!cPersNombre) & Space(250) & rsInstit!cPersCod
        rsInstit.MoveNext
    Loop

    CboAgencia.Clear
    Do While Not lrAgenc.EOF
        CboAgencia.AddItem lrAgenc!cAgeDescripcion & Space(250) & lrAgenc!cAgeCod
        lrAgenc.MoveNext
    Loop

    
End Sub

Private Sub CmdAplicar_Click()
Dim sSql As String
Dim C As COMConecta.DCOMConecta
    
    If Me.CboAgencia.Text = "" Then
        MsgBox "Seleccione una Agencia"
        Exit Sub
    End If
        
    If Me.CmbInstitucion.Text = "" Then
        MsgBox "Seleccione una Institucion"
        Exit Sub
    End If
    
    
    Set C = New COMConecta.DCOMConecta
    C.AbreConexion
    sSql = "DELETE CredTxtEducacion"
    C.Ejecutar sSql
    
    sSql = "INSERT INTO CredTxtEducacion(cPersCodInst,cAgeCod) VALUES('" & Trim(Right(CmbInstitucion.Text, 13)) & "','" & Right(Me.CboAgencia.Text, 2) & "')"
    C.Ejecutar sSql
    
    C.CierraConexion
    Set C = Nothing
    
    LblInst.Caption = Left(CmbInstitucion.Text, 100)
    LblAgencia.Caption = Left(CboAgencia.Text, 100)
    
   
End Sub

Private Sub CmdGeneraTxt_Click()
Dim sCad As String
Dim sArchivo As String
Dim NumeroArchivo As Integer
Dim oCred As COMNCredito.NCOMCredDoc
Set oCred = New COMNCredito.NCOMCredDoc

If Trim(Left(Me.CmbInstitucion.Text, 100)) <> Trim(Left(Me.LblInst.Caption, 100)) Then
    MsgBox "La Institucion Seleccionada es diferente a la Institucion que se Aplicó"
    Exit Sub
End If

If Trim(Left(Me.CboAgencia.Text, 100)) <> Trim(Left(Me.LblAgencia.Caption, 100)) Then
    MsgBox "La Agencia Seleccionada es diferente a la Agencia que se Aplicó"
    Exit Sub
End If

sArchivo = App.path & "\Spooler\" & "A" & Mid(CStr(Year(gdFecSis)), 3, 2) & Format(CStr(Month(gdFecSis)), "00") & "31.151"
sCad = oCred.GeneraArchivoTXT_Educacion(CStr(Year(gdFecSis)) & Format(CStr(Month(gdFecSis)), "00"), "A")
                    
NumeroArchivo = FreeFile
Open sArchivo For Output As #NumeroArchivo
Print #NumeroArchivo, sCad
Close #NumeroArchivo
MsgBox "Se ha generado el Archivo " & "A" & Mid(CStr(Year(gdFecSis)), 3, 2) & Format(CStr(Month(gdFecSis)), "00") & "31.151" & " Satisfactoriamente", vbInformation, "Mensaje"
                                               
sArchivo = App.path & "\Spooler\" & "C" & Mid(CStr(Year(gdFecSis)), 3, 2) & Format(CStr(Month(gdFecSis)), "00") & "31.151"
                
sCad = oCred.GeneraArchivoTXT_Educacion(CStr(Year(gdFecSis)) & Format(CStr(Month(gdFecSis)), "00"), "C")
                
NumeroArchivo = FreeFile
Open sArchivo For Output As #NumeroArchivo
Print #NumeroArchivo, sCad
Close #NumeroArchivo
MsgBox "Se ha generado el Archivo " & "C" & Mid(CStr(Year(gdFecSis)), 3, 2) & Format(CStr(Month(gdFecSis)), "00") & "31.151" & " Satisfactoriamente", vbInformation, "Mensaje"
        
Set oCred = Nothing

    
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    CentraForm Me
    Call CargaControles
End Sub
