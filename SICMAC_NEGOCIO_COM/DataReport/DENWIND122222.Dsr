VERSION 5.00
Begin {C0E45035-5775-11D0-B388-00A0C9055D8E} ConvenioOP 
   ClientHeight    =   8115
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10380
   _ExtentX        =   18309
   _ExtentY        =   14314
   FolderFlags     =   7
   TypeLibGuid     =   "{BC4468AF-763D-11D1-AB28-00A0C9054348}"
   TypeInfoGuid    =   "{BC4468B0-763D-11D1-AB28-00A0C9054348}"
   TypeInfoCookie  =   0
   Version         =   4
   NumConnections  =   1
   BeginProperty Connection1 
      ConnectionName  =   "Cnn1"
      ConnDispId      =   1004
      SourceOfData    =   3
      ConnectionSource=   "Provider=SQLOLEDB.1;Persist Security Info=False"
      Expanded        =   -1  'True
      IsSQL           =   -1  'True
      QuoteChar       =   34
      SeparatorChar   =   46
   EndProperty
   NumRecordsets   =   2
   BeginProperty Recordset1 
      CommandName     =   "ConvenioOP"
      CommDispId      =   1018
      RsDispId        =   1020
      CommandText     =   "dbo.tmpConvenioOP"
      ActiveConnectionName=   "Cnn1"
      CommandType     =   2
      dbObjectType    =   1
      Prepared        =   -1  'True
      IsRSReturning   =   -1  'True
      NumFields       =   15
      BeginProperty Field1 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "dia"
         Caption         =   "dia"
      EndProperty
      BeginProperty Field2 
         Precision       =   0
         Size            =   10
         Scale           =   0
         Type            =   200
         Name            =   "Mes"
         Caption         =   "Mes"
      EndProperty
      BeginProperty Field3 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "a�o"
         Caption         =   "a�o"
      EndProperty
      BeginProperty Field4 
         Precision       =   0
         Size            =   200
         Scale           =   0
         Type            =   200
         Name            =   "Nombre"
         Caption         =   "Nombre"
      EndProperty
      BeginProperty Field5 
         Precision       =   0
         Size            =   180
         Scale           =   0
         Type            =   200
         Name            =   "Relacion"
         Caption         =   "Relacion"
      EndProperty
      BeginProperty Field6 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "Direccion"
         Caption         =   "Direccion"
      EndProperty
      BeginProperty Field7 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "Zona"
         Caption         =   "Zona"
      EndProperty
      BeginProperty Field8 
         Precision       =   0
         Size            =   100
         Scale           =   0
         Type            =   200
         Name            =   "Fono"
         Caption         =   "Fono"
      EndProperty
      BeginProperty Field9 
         Precision       =   0
         Size            =   120
         Scale           =   0
         Type            =   200
         Name            =   "ID"
         Caption         =   "ID"
      EndProperty
      BeginProperty Field10 
         Precision       =   0
         Size            =   20
         Scale           =   0
         Type            =   200
         Name            =   "NroID"
         Caption         =   "NroID"
      EndProperty
      BeginProperty Field11 
         Precision       =   0
         Size            =   13
         Scale           =   0
         Type            =   200
         Name            =   "cPersCod"
         Caption         =   "cPersCod"
      EndProperty
      BeginProperty Field12 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "nPersPersoneria"
         Caption         =   "nPersPersoneria"
      EndProperty
      BeginProperty Field13 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "cPersIDtpo"
         Caption         =   "cPersIDtpo"
      EndProperty
      BeginProperty Field14 
         Precision       =   10
         Size            =   4
         Scale           =   0
         Type            =   3
         Name            =   "nPrdPersRelac"
         Caption         =   "nPrdPersRelac"
      EndProperty
      BeginProperty Field15 
         Precision       =   0
         Size            =   18
         Scale           =   0
         Type            =   200
         Name            =   "cCtaCod"
         Caption         =   "cCtaCod"
      EndProperty
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
   BeginProperty Recordset2 
      CommandName     =   "ColSugerencia"
      CommDispId      =   1021
      RsDispId        =   -1
      ActiveConnectionName=   "Cnn1"
      NumFields       =   0
      NumGroups       =   0
      ParamCount      =   0
      RelationCount   =   0
      AggregateCount  =   0
   EndProperty
End
Attribute VB_Name = "ConvenioOP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub DataEnvironment_Initialize()
    Dim sConn As String
    Dim ClsIni As ClsIni.ClasIni
    Set ClsIni = New ClsIni.ClasIni
    sConn = ClsIni.CadenaConexion
    sConn = Mid(sConn, InStr(1, sConn, ";") + 1)
    'MsgBox sConn
    'Cnn1.ConnectionString = sConn
    Cnn1.Open sConn


End Sub

