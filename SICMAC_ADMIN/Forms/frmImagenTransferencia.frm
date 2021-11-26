VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmImagenTransferencia 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Migra Imagenes"
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7035
   Icon            =   "frmImagenTransferencia.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   7035
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdConvertir 
      Caption         =   "Convertir"
      Height          =   345
      Left            =   1155
      TabIndex        =   5
      Top             =   1020
      Width           =   1080
   End
   Begin VB.TextBox txtCampo 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   3225
      TabIndex        =   4
      Top             =   240
      Width           =   3720
   End
   Begin VB.TextBox txtTabla 
      Appearance      =   0  'Flat
      Height          =   330
      Left            =   15
      TabIndex        =   3
      Top             =   255
      Width           =   3045
   End
   Begin MSComctlLib.ProgressBar pB 
      Height          =   315
      Left            =   45
      TabIndex        =   2
      Top             =   660
      Width           =   6900
      _ExtentX        =   12171
      _ExtentY        =   556
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.CommandButton cmdProceso 
      Caption         =   "&Proceso"
      Height          =   345
      Left            =   45
      TabIndex        =   1
      Top             =   1020
      Width           =   1080
   End
   Begin VB.CommandButton cmdSalir 
      Caption         =   "&Salir"
      Height          =   345
      Left            =   5910
      TabIndex        =   0
      Top             =   1050
      Width           =   1080
   End
   Begin VB.Label lblCampo 
      Caption         =   "Campo"
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
      Height          =   255
      Left            =   3225
      TabIndex        =   7
      Top             =   30
      Width           =   1905
   End
   Begin VB.Label lnlTabla 
      Caption         =   "Tabla"
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
      Height          =   255
      Left            =   15
      TabIndex        =   6
      Top             =   30
      Width           =   1950
   End
End
Attribute VB_Name = "frmImagenTransferencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdConvertir_Click()
    On Error GoTo Error
    
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rsIni As ADODB.Recordset
    Set rsIni = New ADODB.Recordset
    Dim rsFin As ADODB.Recordset
    Set rsFin = New ADODB.Recordset
    Dim sql As String
    
    oCon.AbreConexion
    
    sql = "Select cPersCod From " & Me.txtTabla
    Set rs = oCon.CargaRecordSet(sql)
    
    Me.pB.Min = 0
    Me.pB.Max = rs.RecordCount + 2
    
    While Not rs.EOF
        sql = "Select " & Me.txtCampo.Text & " From " & Me.txtTabla.Text & " Where cPersCod = '" & rs.Fields(0) & "'"
        Set rsIni = oCon.CargaRecordSet(sql)
        
        sql = "Select " & Me.txtCampo.Text & " From " & Me.txtTabla.Text & " Where cPersCod = '" & rs.Fields(0) & "'"
        Set rsFin = oCon.CargaRecordSet(sql, adLockOptimistic, , False)
        
        If IsNull(rsIni.Fields(0)) Then
            If rsFin.EOF And rsFin.BOF Then
                'sql = " Insert DBImagenes.dbo.FirmaCom (cPersCod, iFirma)" _
                    & " Values ('" & rs.Fields(0) & "',Null)"
            Else
                sql = " Update " & Me.txtTabla.Text & " " _
                    & " Set " & Me.txtCampo.Text & " = Null Where cPersCod = '" & rs.Fields(0) & "'"
            End If
            oCon.Ejecutar sql
        Else
            If rsFin.EOF And rsFin.BOF Then
                sql = " Insert DBImagenes.dbo.FirmaCom (cPersCod, iFirma)" _
                    & " Values ('" & rs.Fields(0) & "',Null)"
                oCon.Ejecutar sql
            
                rsFin.Close
                
                sql = "Select iFirma From DBImagenes.dbo.FirmaCom Where cPersCod = '" & rs.Fields(0) & "'"
                Set rsFin = oCon.CargaRecordSet(sql, adLockOptimistic, , False)
            End If
            
            'LetPictureActualiza rsIni.Fields(0), rsFin.Fields(0)
            rsFin.Update
        End If
        Caption = Trim(Str(rs.Bookmark)) & "/" & Trim(Str(rs.RecordCount))
        Me.pB.value = rs.Bookmark
        rs.MoveNext
    Wend
Error:
    MsgBox Err.Description, vbInformation, "Aviso"
End Sub

Private Sub cmdProceso_Click()
    On Error GoTo Error
    Dim oCon As DConecta
    Set oCon = New DConecta
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset
    Dim rsIni As ADODB.Recordset
    Set rsIni = New ADODB.Recordset
    Dim rsFin As ADODB.Recordset
    Set rsFin = New ADODB.Recordset
    Dim sql As String
    
    oCon.AbreConexion
    
    sql = "Select cCodPers From " & Me.txtTabla
    Set rs = oCon.CargaRecordSet(sql)
    
    Me.pB.Min = 0
    Me.pB.Max = rs.RecordCount + 2
    
    While Not rs.EOF
        sql = "Select iFirma From DBImagenes.dbo.Firma Where cCodPers = '" & rs.Fields(0) & "'"
        Set rsIni = oCon.CargaRecordSet(sql)
        
        sql = "Select iFirma From DBImagenes.dbo.FirmaCom Where cCodPers = '" & rs.Fields(0) & "'"
        Set rsFin = oCon.CargaRecordSet(sql, adLockOptimistic, , False)
        
        If IsNull(rsIni.Fields(0)) Then
            If rsFin.EOF And rsFin.BOF Then
                sql = " Insert DBImagenes.dbo.FirmaCom (cCodPers, iFirma)" _
                    & " Values ('" & rs.Fields(0) & "',Null)"
            Else
                sql = " Update DBImagenes.dbo.FirmaCom (cCodPers, )" _
                    & " Set iFirma = Null Where cCodPers = '" & rs.Fields(0) & "'"
            End If
            oCon.Ejecutar sql
        Else
            If rsFin.EOF And rsFin.BOF Then
                sql = " Insert DBImagenes.dbo.FirmaCom (cCodPers, iFirma)" _
                    & " Values ('" & rs.Fields(0) & "',Null)"
                oCon.Ejecutar sql
            
                rsFin.Close
                
                sql = "Select iFirma From DBImagenes.dbo.FirmaCom Where cCodPers = '" & rs.Fields(0) & "'"
                Set rsFin = oCon.CargaRecordSet(sql, adLockOptimistic, , False)
            End If
            
            'LetPictureActualiza rsIni.Fields(0), rsFin.Fields(0)
            rsFin.Update
        End If
        Caption = Trim(Str(rs.Bookmark)) & "/" & Trim(Str(rs.RecordCount))
        Me.pB.value = rs.Bookmark
        rs.MoveNext
    Wend
Error:
    MsgBox Err.Description, vbInformation, "Aviso"
    
End Sub

Private Sub cmdSalir_Click()
    Unload Me
End Sub
