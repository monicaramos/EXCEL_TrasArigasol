VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmTrasArigasol 
   Caption         =   "Traspaso de Poste Arigasol"
   ClientHeight    =   2355
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7275
   Icon            =   "frmTrasArigasol.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   2355
   ScaleWidth      =   7275
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameImportar 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   7335
      Begin VB.CommandButton Command3 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   5760
         TabIndex        =   9
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Importar"
         Height          =   375
         Left            =   4500
         TabIndex        =   8
         Top             =   1920
         Width           =   1095
      End
      Begin VB.TextBox Text5 
         Height          =   315
         Left            =   180
         TabIndex        =   6
         Top             =   690
         Width           =   6735
      End
      Begin VB.Label Label1 
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   23
         Top             =   1230
         Visible         =   0   'False
         Width           =   4695
      End
      Begin VB.Label Label1 
         Caption         =   "Fichero"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   450
         Width           =   615
      End
      Begin VB.Image Image3 
         Height          =   240
         Left            =   900
         Picture         =   "frmTrasArigasol.frx":1782
         Top             =   420
         Width           =   240
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   7260
      Top             =   1620
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FrameEscribir 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7395
      Begin ComctlLib.ProgressBar Pb1 
         Height          =   195
         Left            =   90
         TabIndex        =   22
         Top             =   1170
         Visible         =   0   'False
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.TextBox Text2 
         Height          =   315
         Left            =   90
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   690
         Width           =   7155
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Aceptar"
         Height          =   375
         Index           =   0
         Left            =   4920
         TabIndex        =   2
         Top             =   1800
         Width           =   1035
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Salir"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   1
         Top             =   1800
         Width           =   1035
      End
      Begin VB.Label Label2 
         Caption         =   "Fichero generado"
         Height          =   195
         Left            =   90
         TabIndex        =   4
         Top             =   450
         Width           =   1395
      End
      Begin VB.Image Image2 
         Height          =   240
         Left            =   1530
         Picture         =   "frmTrasArigasol.frx":1884
         Top             =   450
         Width           =   240
      End
   End
   Begin VB.Frame FrameConfig 
      BorderStyle     =   0  'None
      Height          =   2415
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   7335
      Begin VB.TextBox Text8 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   2790
         PasswordChar    =   "*"
         TabIndex        =   13
         Text            =   "Text8"
         Top             =   1290
         Width           =   1485
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   2
         Left            =   2790
         TabIndex        =   12
         Text            =   "Text8"
         Top             =   930
         Width           =   1485
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   1
         Left            =   2790
         TabIndex        =   11
         Text            =   "Text8"
         Top             =   570
         Width           =   1515
      End
      Begin VB.TextBox Text8 
         Height          =   285
         Index           =   0
         Left            =   2790
         TabIndex        =   15
         Text            =   "Text8"
         Top             =   1650
         Width           =   1485
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Salir"
         Height          =   375
         Index           =   1
         Left            =   6120
         TabIndex        =   19
         Top             =   1920
         Width           =   1095
      End
      Begin VB.CommandButton cmdConfig 
         Caption         =   "Guardar"
         Height          =   375
         Index           =   0
         Left            =   4800
         TabIndex        =   17
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Server"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   240
         TabIndex        =   21
         Top             =   600
         Width           =   570
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Usuario"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   20
         Top             =   990
         Width           =   660
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nro.Máximo de Calidades"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   18
         Top             =   1710
         Width           =   2145
      End
      Begin VB.Label Label7 
         Caption         =   "CLASIFICACION"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   240
         Width           =   2010
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   165
         Index           =   12
         Left            =   240
         TabIndex        =   14
         Top             =   1350
         Width           =   825
      End
   End
End
Attribute VB_Name = "frmTrasArigasol"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'Private WithEvents frmC As frmCal
Private NoEncontrados As String




Dim SQL As String
Dim VariasEntradas As String


Dim Albaran As Long
Dim FecAlbaran As String
Dim Socio As String
Dim Campo As String
Dim Variedad As String
Dim TipoEntr As String
Dim KilosNet As String
Dim Cajones As String
Dim Calidad(20) As String

Private WithEvents frmMens As frmMensajes 'Registros que no ha entrado con error
Attribute frmMens.VB_VarHelpID = -1



Private Sub cmdConfig_Click(Index As Integer)
Dim I As Integer

    If Index = 1 Then
        Unload Me
    Else
        SQL = ""
        For I = 0 To Text8.Count - 1
            If Text8(I).Text = "" Then SQL = SQL & "Campo: " & I & vbCrLf
        Next I
        If SQL <> "" Then
            SQL = "No pueden haber campos vacios: " & vbCrLf & vbCrLf & SQL
            MsgBox SQL, vbExclamation
            Text8(0).SetFocus
            Exit Sub
        End If
        
        mConfig.MaxCalidades = Text8(0).Text
        mConfig.SERVER = Text8(1).Text
        mConfig.User = Text8(2).Text
        mConfig.password = Text8(3).Text
        
        mConfig.Guardar
        
        vConfiguracion False
'        If varConfig.Grabar = 0 Then End
    End If
End Sub

Private Sub Command1_Click(Index As Integer)
Dim Rc As Byte
Dim Mens As String

    If Index = 1 Then
        Unload Me
        Exit Sub
    End If
        
    If Text2.Text <> "" Then
        If Dir(Text2.Text) <> "" Then
            MsgBox "Fichero ya existe", vbExclamation
            Exit Sub
        Else
            FileCopy App.Path & "\" & mConfig.Plantilla, Text2.Text
            NombreHoja = Text2.Text
        End If
    End If
    
    'Abrimos excel
    Rc = AbrirEXCEL
    
    If Rc = 0 Then
    
        'Si queremos que se vea descomentamos  esto
        MiXL.Application.visible = False
'        MiXL.Parent.Windows(1).Visible = False
    
        'Realizamos todos los datos
        'abrimos conexion
        If AbrirConexion(BaseDatos) Then
        
            Screen.MousePointer = vbHourglass
            
            'Vamos linea a linea
            Mens = "Error insertando en Excel"
            
            If EsImportaci = 3 Then
                If Not RecorremosLineasInformes(Mens) Then
                    MsgBox Mens, vbExclamation
                End If
            Else ' esimportaci = 2
                If Not RecorremosLineas(Mens) Then
                    MsgBox Mens, vbExclamation
                End If
            End If
            
            Screen.MousePointer = vbDefault
            
        End If
    
        'Cerramos el excel
        CerrarExcel
                
        MsgBox "Proceso finalizado", vbExclamation


    End If
    
    
End Sub

Private Sub Command2_Click()
Dim Rc As Byte
Dim I As Integer
Dim Rs1 As ADODB.Recordset
Dim KilosI As Long
Dim b As Boolean
Dim Notas As String

    'IMPORTAR
    If Text5.Text = "" Then
        MsgBox "Escriba el nombre del fichero excel", vbExclamation
        Exit Sub
    End If
        
    If Dir(Text5.Text) = "" Then
        MsgBox "Fichero no existe"
        Exit Sub
    End If
    
    NombreHoja = Text5.Text
    'Abrimos excel
    
    Rc = AbrirEXCEL
    
    If Rc = 0 Then
    
        If EsImportaci = 1 Then
            If AbrirConexion(BaseDatos) Then
            
                
                'Vamos linea a linea, buscamos su trabajador
                RecorremosLineasFicheroTraspaso
                
            End If
        
            'Cerramos el excel
            CerrarExcel
        
        
        End If
        
        
        MsgBox "FIN", vbInformation
        
    End If
    
End Sub

Private Sub Command3_Click()
    Unload Me
End Sub


Private Sub Form_Load()
    
'    Combo1.ListIndex = Month(Now) - 1
'    Text3.Text = Year(Now)
    FrameEscribir.visible = False
    FrameImportar.visible = False
    Me.FrameConfig.visible = False
'    Limpiar
'    Select Case EsImportaci
'    Case 1
'        Caption = "Cargar Traspaso de Poste desde fichero excel"
'        FrameImportar.visible = True
'
'    End Select
    
    
 
End Sub

Private Sub Limpiar()
Dim T As Control
    For Each T In Me.Controls
        If TypeOf T Is TextBox Then
            T.Text = ""
        End If
    Next
        
End Sub
Private Function TransformaComasPuntos(CADENA) As String
Dim cad As String
Dim J As Integer
    
    J = InStr(1, CADENA, ",")
    If J > 0 Then
        cad = Mid(CADENA, 1, J - 1) & "." & Mid(CADENA, J + 1)
    Else
        cad = CADENA
    End If
    TransformaComasPuntos = cad
End Function

Private Sub frmC_Selec(vFecha As Date)
'    Text4.Text = Format(vFecha, "dd/mm/yyyy")
End Sub

Private Sub Image1_Click()
    AbrirDialogo 0
End Sub

Private Sub Image2_Click()
    AbrirDialogo 1
End Sub


Private Sub AbrirDialogo(Opcion As Integer)

    On Error GoTo EA
    
    With Me.CommonDialog1
        Select Case Opcion
        Case 0, 2
            .DialogTitle = "Archivo origen de datos"
        Case 1
            .DialogTitle = "Archivo destino de datos"
        End Select
        .Filter = "EXCEL (*.xls)|*.xls"
        .CancelError = True
        If Opcion <> 1 Then
            .ShowOpen
            If Opcion = 0 Then
                Text2.Text = .FileName
            Else
                Text5.Text = .FileName
            End If
        Else
            .ShowSave
            Text2.Text = .FileName
        End If
        
        
        
    End With
EA:
End Sub

Public Sub IncrementarProgresNew(ByRef PBar As ProgressBar, Veces As Integer)
On Error Resume Next
'    PBar.Value = PBar.Value + ((Veces * PBar.Max) / CInt(PBar.Tag))
    PBar.Value = PBar.Value + Veces
    If Err.Number <> 0 Then Err.Clear
End Sub



Private Function RecorremosLineas(Mens As String) As Boolean
Dim I As Integer
Dim J As Integer
Dim JJ As Integer
Dim F As Date
Dim Cod As String
Dim FE As String
Dim RT As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim Calidad As Integer
Dim NFic As Integer
Dim Lin As String
Dim Sql1 As String
Dim NFile As Integer

    On Error GoTo eRecorremosLineas

    RecorremosLineas = False


    SQL = "select * from rhisfruta "
    Sql1 = "select count(*) from rhisfruta "
    
    '[Monica] 19/04/2010: añadida la condicion del sql en el fichero condicionsql.txt
    If Dir(App.Path & "\condicionsql.txt", vbArchive) <> "" Then
    
        NFile = FreeFile
    
        Open App.Path & "\condicionsql.txt" For Input As #NFile
 
        If Not EOF(NFile) Then
            Line Input #NFile, Lin
    
            SQL = SQL & " where numalbar in (" & Lin & ")"
            Sql1 = Sql1 & " where numalbar in (" & Lin & ")"
        End If
    End If
    '[Monica] 19/04/2010
    

    Set RT = New ADODB.Recordset
    RT.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RT.EOF Then
        Me.Pb1.visible = True
        Me.Pb1.Max = RT.Fields(0).Value
        Me.Pb1.Value = 0
        Me.Refresh
    End If
    
    Set RT = Nothing
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    
    I = 1
    While Not RT.EOF
        I = I + 1
    
        IncrementarProgresNew Pb1, 1
    
        ExcelSheet.Cells(I, 1).Value = RT!numalbar ' numero de albaran
        ExcelSheet.Cells(I, 2).Value = Format(RT!fecalbar, "yyyy/mm/dd") ' fecha de albaran
        ExcelSheet.Cells(I, 3).Value = RT!codsocio ' codigo de socio
        
        SQL = "select nomsocio from rsocios where codsocio = " & RT!codsocio
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RS.EOF Then
            ExcelSheet.Cells(I, 4).Value = RS.Fields(0).Value ' nombre de socio
        Else
            ExcelSheet.Cells(I, 4).Value = "" ' nombre de socio
        End If
        
        Set RS = Nothing
        
        ExcelSheet.Cells(I, 5).Value = RT!codcampo ' codigo de campo
        ExcelSheet.Cells(I, 6).Value = RT!codvarie ' codigo de variedad
        
        SQL = "select nomvarie from variedades where codvarie = " & RT!codvarie
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        If Not RS.EOF Then
            ExcelSheet.Cells(I, 7).Value = RS.Fields(0).Value  ' nombre de variedad
        Else
            ExcelSheet.Cells(I, 7).Value = "" ' nombre de variedad
        End If
        
        Set RS = Nothing
        
        ExcelSheet.Cells(I, 8).Value = RT!TipoEntr ' tipo de entrada
        ExcelSheet.Cells(I, 9).Value = RT!KilosNet ' kilos netos
        
        ' cargamos las calidades
        SQL = "select * from rhisfruta_clasif where numalbar = " & RT!numalbar
        Set RS = New ADODB.Recordset
        RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
        
        While Not RS.EOF
            Calidad = RS!codcalid
            
            ExcelSheet.Cells(I, Calidad + 9).Value = RS!KilosNet ' kilos netos
            
            RS.MoveNext
        Wend
        Set RS = Nothing
        
        ' si no hay kilos de algunas calidades las rellenamos a cero
        For JJ = 10 To 29
            If ExcelSheet.Cells(I, JJ).Value = "" Then ExcelSheet.Cells(I, JJ).Value = 0
        Next JJ
    
'        ExcelSheet.Cells(I, 23).Value = 0
'        ExcelSheet.Cells(I, 24).Value = 0
'
    
        RT.MoveNext
    Wend
    
    RT.Close
    Set RT = Nothing
    
    RecorremosLineas = True
    
    Exit Function
    
eRecorremosLineas:
    Mens = Mens & vbCrLf & "Recorriendo lineas: " & Err.Description
End Function


Private Sub Image3_Click()
 AbrirDialogo 2
End Sub


Private Sub Image4_Click()
'    Set frmC = New frmCal
'    frmC.Fecha = Now
'    If Text4.Text <> "" Then
'        If IsDate(Text4.Text) Then frmC.Fecha = CDate(Text4.Text)
'    End If
'    frmC.Show vbModal
'    Set frmC = Nothing
End Sub

Private Sub Image5_Click()
    MsgBox "Formato importe:   SOLO el punto decimal: 1.49", vbExclamation
End Sub

'Private Sub Text4_LostFocus()
'    Text4.Text = Trim(Text4.Text)
'    If Text4.Text <> "" Then
'        If IsDate(Text4.Text) Then
'            Text4.Text = Format(Text4.Text, "dd/mm/yyyy")
'        Else
'            MsgBox "Fecha incorrecta", vbExclamation
'            Text4.Text = ""
'        End If
'    End If
'End Sub
'
'

'-------------------------------------
Private Function RecorremosLineasLiquidacion()
Dim FIN As Boolean
Dim I As Long
Dim JJ As Integer
Dim LineasEnBlanco As Integer

    'Desde la fila donde empieza los trabajadores
    'Hasta k este vacio
    'Iremos insertando en tmpHoras
    ' Con trbajador, importe, 0 , 1 ,2
    '             Existe, No existe, IMPORTE negativo
    '
    
    SQL = "DELETE FROM tmpExcel where codusu = " & Usuario
    Conn.Execute SQL
    FIN = False
    I = 2
    LineasEnBlanco = 0
    While Not FIN
        'Debug.Print "L: " & i
        If Trim(CStr(ExcelSheet.Cells(I, 1).Value)) <> "" Then
            LineasEnBlanco = 0
            If IsNumeric((ExcelSheet.Cells(I, 1).Value)) Then
                If Val(ExcelSheet.Cells(I, 1).Value) > 0 Then
                        'albaran
                        Albaran = Val(ExcelSheet.Cells(I, 1).Value)
                        
                        'Importe
                        FecAlbaran = Format(ExcelSheet.Cells(I, 2).Value, "yyyy/mm/dd")
                        Socio = ExcelSheet.Cells(I, 3).Value
                        Campo = ExcelSheet.Cells(I, 5).Value
                        Variedad = ExcelSheet.Cells(I, 6).Value
                        TipoEntr = ExcelSheet.Cells(I, 8).Value
                        KilosNet = ExcelSheet.Cells(I, 9).Value
                        
                        
                        For JJ = 1 To 20
                            Calidad(JJ) = Val(ExcelSheet.Cells(I, 9 + JJ).Value)
                        Next JJ
                        
                        'InsertartmpLiquida
                        InsertaTmpExcel
                    
                    End If
            End If
        Else
            LineasEnBlanco = LineasEnBlanco + 1
            If LineasEnBlanco < 30 Then
               ' FIN = False
            Else
                FIN = True
               
            End If
        End If
        'Siguiente
        'Siguiente
        I = I + 1
    Wend
End Function




Private Sub InsertaTmpExcel()
Dim vSql As String
Dim vSql2 As String
Dim RT As ADODB.Recordset
Dim RT1 As ADODB.Recordset
Dim RT2 As ADODB.Recordset
Dim Existe As Boolean
Dim ExisteCalidad As Boolean
Dim ExisteEnTemporal As Boolean
Dim TotalKilos As Long
Dim Cuadra As Boolean
Dim JJ As Integer

    On Error GoTo EInsertaTmpExcel
    
    vSql = "Select * from rhisfruta "
    vSql = vSql & " WHERE numalbar = " & Albaran
    vSql = vSql & " and fecalbar = '" & Format(CDate(FecAlbaran), "yyyy-mm-dd") & "'"
    vSql = vSql & " and codsocio = " & Socio
    vSql = vSql & " and codcampo = " & Campo
    vSql = vSql & " and codvarie = " & Variedad
    vSql = vSql & " and tipoentr = " & TipoEntr

    Set RT = New ADODB.Recordset
    RT.Open vSql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If RT.EOF Then
        Existe = False
    Else
        Existe = True
    End If
    
    ' si existe la entrada vemos si podemos actualizarla
    If Existe Then
        ExisteCalidad = True
        
        For JJ = 1 To mConfig.MaxCalidades
            If Calidad(JJ) <> 0 Then  ' solo si hay kilos
'                vSQL = "select * from rhisfruta_clasif where numalbar = " & Albaran
'                vSQL = vSQL & " and codvarie = " & Variedad
'                vSQL = vSQL & " and codcalid = " & JJ
'
'                Set RT1 = New ADODB.Recordset
'                RT1.Open vSQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
'
'                If RT1.EOF Then
                    vSql2 = "select * from rcalidad where codvarie = " & Variedad
                    vSql2 = vSql2 & " and codcalid = " & JJ
                    
                    Set RT2 = New ADODB.Recordset
                    RT2.Open vSql2, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
                    
                    If RT2.EOF Then
                        ExisteCalidad = False
                        Set RT2 = Nothing
                        Exit For
                    Else
                        ExisteCalidad = True
                        Set RT2 = Nothing
                    End If
'                End If
                
            End If
        Next JJ
    
    
        If ExisteCalidad Then ' comprobamos que la suma de calidades da kilosnetos
            TotalKilos = 0
            For JJ = 1 To 20
                TotalKilos = TotalKilos + Calidad(JJ)
            Next JJ
            If TotalKilos <> RT!KilosNet Then
                Cuadra = False
            Else
                Cuadra = True
            End If
        End If
    
    End If
    
    If Existe And ExisteCalidad And Cuadra Then
        
        ExisteEnTemporal = False
        vSql = "select * from tmpexcel where numalbar = " & Albaran & " and codusu = " & Usuario
    
        Set RT2 = New ADODB.Recordset
        RT2.Open vSql, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
            
        If Not RT2.EOF Then
            ExisteEnTemporal = True
        End If
    
        SQL = "insert into tmpexcel (codusu, numalbar, fecalbar, codsocio, codcampo, codvarie, tipoentr, kilosnet, "
        SQL = SQL & "calidad1, calidad2, calidad3, calidad4, calidad5, calidad6, calidad7, calidad8, calidad9, "
        SQL = SQL & "calidad10, calidad11, calidad12, calidad13, calidad14, calidad15, calidad16, calidad17, "
        SQL = SQL & "calidad18, calidad19, calidad20, situacion) values ("
        SQL = SQL & Usuario & ","
        SQL = SQL & Albaran & ","
        SQL = SQL & "'" & Format(CDate(FecAlbaran), "yyyy-mm-dd") & "',"
        SQL = SQL & Socio & ","
        SQL = SQL & Campo & ","
        SQL = SQL & Variedad & ","
        SQL = SQL & TipoEntr & ","
        SQL = SQL & KilosNet & ","
        
        For JJ = 1 To mConfig.MaxCalidades
            SQL = SQL & Calidad(JJ) & ","
        Next JJ
    
        If ExisteEnTemporal Then
            SQL = SQL & "2)"
        Else
            SQL = SQL & "0)"
        End If
        
    Else
        SQL = "insert into tmpexcel (codusu, numalbar, fecalbar, codsocio, codcampo, codvarie, tipoentr, kilosnet, "
        SQL = SQL & "calidad1, calidad2, calidad3, calidad4, calidad5, calidad6, calidad7, calidad8, calidad9, "
        SQL = SQL & "calidad10, calidad11, calidad12, calidad13, calidad14, calidad15, calidad16, calidad17,"
        SQL = SQL & "calidad18, calidad19, calidad20, situacion) values ("
        SQL = SQL & Usuario & ","
        SQL = SQL & Albaran & ",'"
        SQL = SQL & Format(CDate(FecAlbaran), "yyyy-mm-dd") & "',"
        SQL = SQL & Socio & ","
        SQL = SQL & Campo & ","
        SQL = SQL & Variedad & ","
        SQL = SQL & TipoEntr & ","
        SQL = SQL & KilosNet & ","
        
        For JJ = 1 To mConfig.MaxCalidades
            SQL = SQL & Calidad(JJ) & ","
        Next JJ
    
        If Not Existe Then
            SQL = SQL & "1)" ' no existe el albaran
        Else
            If Not ExisteCalidad Then ' no existe la calidad
                SQL = SQL & "11)"
            Else
                SQL = SQL & "12)" ' no cuadran kilos
            End If
        End If
        
    End If
    
    
    If SQL <> "" Then Conn.Execute SQL
        
    RT.Close
    
    Exit Sub
EInsertaTmpExcel:
    MsgBox Err.Description
End Sub



Private Sub vConfiguracion(Leer As Boolean)

'    With varConfig
'        If Leer Then
'            Text8(0).Text = .IniLinNomina
'            Text8(1).Text = .FinLinNominas
'            Text8(2).Text = .ColTrabajadorNom
'            Text8(3).Text = .hc
'            Text8(4).Text = .HPLUS
'            Text8(5).Text = .DIAST
'            Text8(6).Text = .Anticipos
'            Text8(7).Text = .ColTrabajadoresLIQ
'            Text8(8).Text = .ColumnaLiquidacion
'            Text8(9).Text = .FilaLIQ
'            Text8(10).Text = .HN
'        Else
'            .IniLinNomina = Val(Text8(0).Text)
'            .FinLinNominas = Val(Text8(1).Text)
'            .ColTrabajadorNom = Val(Text8(2).Text)
'            .hc = Val(Text8(3).Text)
'            .HPLUS = Val(Text8(4).Text)
'            .DIAST = Val(Text8(5).Text)
'            .Anticipos = Val(Text8(6).Text)
'            .ColTrabajadoresLIQ = Val(Text8(7).Text)
'            .ColumnaLiquidacion = Val(Text8(8).Text)
'            .FilaLIQ = Val(Text8(9).Text)
'            .HN = Val(Text8(10).Text)
'        End If
'    End With
End Sub

Private Sub Text8_GotFocus(Index As Integer)
    With Text8(Index)
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub Text8_LostFocus(Index As Integer)
    With Text8(Index)
        .Text = Trim(.Text)
        If .Text = "" Then Exit Sub
        
        Select Case Index
            Case 0 ' numero de calidades
                If Not IsNumeric(.Text) Then
                    MsgBox "Campo debe ser numérico", vbExclamation
                    .Text = ""
                    .SetFocus
                    Exit Sub
                End If
                .Text = Val(.Text)
            
            Case 2, 3 ' usuario y password deben de estar encriptados
            
            
        End Select
            
            
    End With
End Sub



Private Function RecorremosLineasInformes(Mens As String) As Boolean
Dim I As Integer
Dim J As Integer
Dim JJ As Integer
Dim F As Date
Dim Cod As String
Dim FE As String
Dim RT As ADODB.Recordset
Dim RS As ADODB.Recordset
Dim Calidad As Integer
Dim NFic As Integer
Dim Lin As String
Dim Sql1 As String
Dim NFile As Integer
Dim AlbaranAnt As Long
Dim Primero As Boolean

    On Error GoTo eRecorremosLineas

    RecorremosLineasInformes = False

    Select Case TipoListado
        Case 1
            SQL = "select tmpinformes.*, variedades.nomvarie from tmpinformes INNER JOIN variedades ON tmpinformes.importe1 = variedades.codvarie where codusu = " & Usuario
            Sql1 = "select count(*) from tmpinformes where codusu = " & Usuario
            
            SQL = SQL & " order by campo1, codigo1, importe1, fecha1, importe2"
    End Select

    Set RT = New ADODB.Recordset
    RT.Open Sql1, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RT.EOF Then
        Me.Pb1.visible = True
        Me.Pb1.Max = RT.Fields(0).Value
        Me.Pb1.Value = 0
        Me.Refresh
    End If
    
    Set RT = Nothing
    Set RT = New ADODB.Recordset
    RT.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    Select Case TipoListado
        Case 1 ' informe de comprobacion de venta fruta
            ExcelSheet.Cells(1, 1).Value = "Socio/Cliente"
            ExcelSheet.Cells(1, 2).Value = "Código"
            ExcelSheet.Cells(1, 3).Value = "Nombre"
            ExcelSheet.Cells(1, 4).Value = "Cod.Var"
            ExcelSheet.Cells(1, 5).Value = "Variedad"
            ExcelSheet.Cells(1, 6).Value = "Fecha"
            ExcelSheet.Cells(1, 7).Value = "Albarán"
            ExcelSheet.Cells(1, 8).Value = "Palot"
            ExcelSheet.Cells(1, 9).Value = "Tara Palot"
            ExcelSheet.Cells(1, 10).Value = "Calibre"
            ExcelSheet.Cells(1, 11).Value = "Cajas"
            ExcelSheet.Cells(1, 12).Value = "Palets"
            ExcelSheet.Cells(1, 13).Value = "Peso Neto"
            ExcelSheet.Cells(1, 14).Value = "Tipo Alb"
            
            For I = 1 To 15
                ExcelSheet.Cells(1, I + 14).Value = ""
            Next I
            
    End Select
            
    I = 1
    
    While Not RT.EOF
        I = I + 1
    
        IncrementarProgresNew Pb1, 1
    
        If RT!campo1 = 0 Then
            ExcelSheet.Cells(I, 1).Value = "Socio" 'tipo
        Else
            ExcelSheet.Cells(I, 1).Value = "Cliente" 'tipo
        End If
        
        ExcelSheet.Cells(I, 2).Value = RT!codigo1 ' codigo de socio o de cliente
        ExcelSheet.Cells(I, 3).Value = RT!nombre1 ' nombre de socio o de cliente
        ExcelSheet.Cells(I, 4).Value = RT!importe1 ' codigo de la variedad
        ExcelSheet.Cells(I, 5).Value = RT!nomvarie ' nombre de la variedad
        ExcelSheet.Cells(I, 6).Value = Format(RT!fecha1, "dd/mm/yyyy") ' fecha de albaran
        ExcelSheet.Cells(I, 7).Value = RT!importe2 ' numero del albaran
        ExcelSheet.Cells(I, 8).Value = RT!importeb1 ' numero de palots
        ExcelSheet.Cells(I, 9).Value = RT!importeb2 ' tara de palots
        ExcelSheet.Cells(I, 10).Value = RT!nombre2 ' calibre
        ExcelSheet.Cells(I, 11).Value = RT!importe3 ' numero de cajas
        ExcelSheet.Cells(I, 12).Value = RT!importe4 ' numero de palets
        ExcelSheet.Cells(I, 13).Value = RT!importe5 ' peso neto
    
        If RT!importeb3 = 0 Then
            ExcelSheet.Cells(I, 14).Value = "Vta.Fruta" 'tipo de albaran
        Else
            ExcelSheet.Cells(I, 14).Value = "Precalibrado" 'tipo de albaran
        End If
    
        RT.MoveNext
    Wend
    
    RT.Close
    Set RT = Nothing
    
    RecorremosLineasInformes = True
    
    Exit Function
    
eRecorremosLineas:
    Mens = Mens & vbCrLf & "Recorriendo lineas: " & Err.Description
End Function



Private Function RecorremosLineasFicheroTraspaso()
Dim FIN As Boolean
Dim I As Long
Dim JJ As Integer
Dim LineasEnBlanco As Integer
Dim RS As ADODB.Recordset
Dim SQL As String
Dim vSql As String
Dim CADENA As String
Dim Posicion As Integer

Dim turno As Integer
Dim Albaran As Long
Dim factura As String
Dim Fecha As String
Dim cliente As String
Dim nomclien As String
Dim tarjeta As Double
Dim matricula As String
Dim km As Long
Dim producto As Integer
Dim nomprodu As String
Dim surtidor As Integer
Dim manguera As Integer
Dim NSuministro As Integer
Dim precio As Currency
Dim descuento As Currency
Dim descuentoporc As Currency
Dim iva As Currency
Dim cantidad As Currency
Dim Importe As Currency
Dim idtipopago As Integer
Dim desctipopago As String
Dim nif As String


    vSql = "delete from tmptraspaso where codusu = " & Usuario
    Conn.Execute vSql

    FIN = False
    I = 3
    LineasEnBlanco = 0
    While Not FIN
        'Debug.Print "L: " & i
        If Trim(CStr(ExcelSheet.Cells(I, 1).Value)) <> "" Then
            LineasEnBlanco = 0
            
            turno = ExcelSheet.Cells(I, 4).Value
            Albaran = 0
            If ExcelSheet.Cells(I, 5).Value <> "" Then
                Albaran = Mid(ExcelSheet.Cells(I, 5).Value, 5, Len(ExcelSheet.Cells(I, 5).Value) - 4) ' le quito el TIC1
            End If
            factura = ExcelSheet.Cells(I, 6).Value
            Fecha = ""
            If ExcelSheet.Cells(I, 8).Value <> "" Then
                Fecha = Mid(ExcelSheet.Cells(I, 8).Value, 2, Len(ExcelSheet.Cells(I, 8).Value) - 1) ' quito la comilla
            End If
'            fechaF = Mid(Fecha, 1, 4) & "-" & Mid(Fecha, 5, 2) & "-" & Mid(Fecha, 7, 2)
'            fechaH = Mid(Fecha, 9, 2) & ":" & Mid(Fecha, 11, 2) & ":" & Mid(Fecha, 13, 2)
            cliente = ExcelSheet.Cells(I, 9).Value
            nomclien = ExcelSheet.Cells(I, 10).Value
            tarjeta = 0
            If ExcelSheet.Cells(I, 11).Value <> "'" Then
                tarjeta = Mid(ExcelSheet.Cells(I, 11).Value, 2, Len(ExcelSheet.Cells(I, 11).Value) - 1)
            End If
            matricula = ExcelSheet.Cells(I, 13).Value
            km = ExcelSheet.Cells(I, 14).Value
            producto = ExcelSheet.Cells(I, 15).Value
            nomprodu = ExcelSheet.Cells(I, 16).Value
            surtidor = ExcelSheet.Cells(I, 17).Value
            manguera = ExcelSheet.Cells(I, 18).Value
            NSuministro = 0
            If ExcelSheet.Cells(I, 19).Value <> "" Then
                NSuministro = Mid(ExcelSheet.Cells(I, 19).Value, 2, Len(ExcelSheet.Cells(I, 19).Value) - 1)
            End If
            precio = ExcelSheet.Cells(I, 20).Value
            descuento = ExcelSheet.Cells(I, 21).Value
            descuentoporc = ExcelSheet.Cells(I, 22).Value
            iva = ExcelSheet.Cells(I, 23).Value
            cantidad = ExcelSheet.Cells(I, 24).Value
            Importe = ExcelSheet.Cells(I, 25).Value
            idtipopago = Mid(ExcelSheet.Cells(I, 26).Value, 2, Len(ExcelSheet.Cells(I, 26).Value) - 1)
            desctipopago = ExcelSheet.Cells(I, 27).Value
            nif = ExcelSheet.Cells(I, 28).Value
            
            vSql = "insert into tmptraspaso (codusu,turno,albaran,factura,fecha,cliente,nomclien,tarjeta,matricula,km,producto,nomprodu, "
            vSql = vSql & "surtidor,manguera,nsuministro,precio,descuento,descuentoporc,iva,cantidad,importe,idtipopago,desctipopago,nif) values ("
            vSql = vSql & Usuario & "," & DBSet(turno, "N") & "," & DBSet(Albaran, "N") & "," & DBSet(factura, "T") & "," & DBSet(Fecha, "T") & ","
            vSql = vSql & DBSet(cliente, "T") & "," & DBSet(nomclien, "T") & "," & DBSet(tarjeta, "N", "S") & "," & DBSet(matricula, "T") & ","
            vSql = vSql & DBSet(km, "N") & "," & DBSet(producto, "N") & "," & DBSet(nomprodu, "T") & "," & DBSet(surtidor, "N") & ","
            vSql = vSql & DBSet(manguera, "N") & "," & DBSet(NSuministro, "T") & "," & DBSet(precio, "N") & "," & DBSet(descuento, "N") & ","
            vSql = vSql & DBSet(descuentoporc, "N") & "," & DBSet(iva, "N") & "," & DBSet(cantidad, "N") & "," & DBSet(Importe, "N") & ","
            vSql = vSql & DBSet(idtipopago, "N") & "," & DBSet(desctipopago, "T") & "," & DBSet(nif, "T") & ")"
            
            Conn.Execute vSql
            
        Else
            
            FIN = True

        End If
        'Siguiente
        'Siguiente
        I = I + 1
    Wend
End Function



