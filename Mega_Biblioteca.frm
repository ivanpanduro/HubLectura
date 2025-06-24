VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "mscomctl.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6585
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   14085
   LinkTopic       =   "Form1"
   ScaleHeight     =   6585
   ScaleWidth      =   14085
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView ListView1 
      Height          =   5415
      Left            =   2520
      TabIndex        =   6
      Top             =   120
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   9551
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton btn_Eliminar 
      Caption         =   "Eliminar"
      Height          =   495
      Left            =   9000
      TabIndex        =   5
      Top             =   5640
      Width           =   1575
   End
   Begin VB.CommandButton btn_Modificar 
      Caption         =   "Modificar"
      Height          =   375
      Left            =   6120
      TabIndex        =   4
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton btn_Agregar 
      Caption         =   "Agregar"
      Height          =   375
      Left            =   3000
      TabIndex        =   3
      Top             =   5760
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   5895
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   2295
      Begin VB.CommandButton btn_Calificacion 
         Caption         =   "Calificacion"
         Height          =   975
         Left            =   240
         TabIndex        =   8
         Top             =   4080
         Width           =   1815
      End
      Begin VB.CommandButton btn_Porleer 
         Caption         =   "Por Leer "
         Height          =   1215
         Left            =   240
         TabIndex        =   7
         Top             =   2640
         Width           =   1815
      End
      Begin VB.CommandButton btn_Leiste 
         Caption         =   "Leidos"
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   1440
         Width           =   1815
      End
      Begin VB.CommandButton btn_Catalogo 
         Caption         =   "Catalogo MEGA"
         Height          =   855
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   360
         Width           =   1815
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

' Variables globales del formulario
Dim conn As ADODB.Connection
Dim rs As ADODB.Recordset
Private Sub CargarLibrosBD()
    Dim sql As String
    Dim item As ListItem
    
    On Error GoTo errorCarga
    
    ' Verificar que la conexión esté disponible
    If conn Is Nothing Then
        MsgBox "No hay conexión a la base de datos.", vbCritical
        Exit Sub
    End If
    
    If conn.State <> adStateOpen Then
        MsgBox "La conexión a la base de datos está cerrada.", vbCritical
        Exit Sub
    End If
    
    ' Crear recordset
    Set rs = New ADODB.Recordset
    
    ' Consulta SQL - Basada en los campos que veo en tu captura
    ' Nota: He agregado Id para referencia, puedes quitarlo si no lo necesitas
    sql = "SELECT Id, Titulo, Autor, Genero, Calificacion, Estado, UsuarioId FROM Libros ORDER BY Id"
    
    ' Abrir recordset
    rs.Open sql, conn, adOpenStatic, adLockReadOnly
    
    ' Limpiar ListView
    ListView1.ListItems.Clear
    
    ' Verificar si hay registros
    If rs.EOF Then
        MsgBox "No hay libros registrados en la base de datos.", vbInformation
        GoTo Cleanup
    End If
    
    ' Cargar datos en ListView
    Do While Not rs.EOF
        ' Agregar item principal (Título)
        Set item = ListView1.ListItems.Add(, "ID_" & rs!Id, IIf(IsNull(rs!Titulo), "", rs!Titulo))
        
        ' Agregar subitems (manejar valores nulos)
        item.SubItems(1) = IIf(IsNull(rs!Autor), "", rs!Autor)
        item.SubItems(2) = IIf(IsNull(rs!Genero), "", rs!Genero)
        item.SubItems(3) = IIf(IsNull(rs!Calificacion), "", rs!Calificacion)
        item.SubItems(4) = IIf(IsNull(rs!Estado), "", rs!Estado)
        item.SubItems(5) = IIf(IsNull(rs!UsuarioId), "", rs!UsuarioId)
        
        ' Guardar el ID en el Tag para referencia futura
        item.Tag = rs!Id
        
        rs.MoveNext
    Loop
    
    MsgBox "Libros cargados correctamente. Total: " & ListView1.ListItems.Count, vbInformation

Cleanup:
    ' Cerrar recordset
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
    Exit Sub

errorCarga:
    MsgBox "Error al cargar libros: " & Err.Description, vbCritical
    GoTo Cleanup
End Sub
Private Sub Command1_Click()

End Sub

Private Sub Command3_Click()

End Sub
Dim conn As ADODB.Connection

Private Sub btn_Agregar_Click()
    Dim rs As ADODB.Recordset
    Set rs = New ADODB.Recordset

    On Error GoTo errorInsertar

    conn.Execute "INSERT INTO Libros (Titulo, Autor, Genero, Calificacion, Estado, UsuarioId) " & _
                 "VALUES ('1984', 'George Orwell', 'Distopía', 5, 'Leído', 1)"

    MsgBox "Libro agregado correctamente.", vbInformation
    Exit Sub

errorInsertar:
    MsgBox "Error al insertar: " & Err.Description, vbCritical
End Sub






Private Sub btnCatalogo_Click()
    Call CargarLibros
End Sub

Private Sub btn_Catalogo_Click()

        Call CargarLibrosBD

End Sub

Private Sub btnLeidos_Click()
    CargarLibrosEstado "Leído"
End Sub

Private Sub btnPorLeer_Click()
    CargarLibrosEstado "Quiero leer"
End Sub

Private Sub Form_Load()
    ' Inicializar conexión
    Set conn = New ADODB.Connection
    
    On Error GoTo errorConexion
    
    ' CONEXIÓN CORREGIDA - Cambiado PANDURO_LAF por PANDURO_LAP
    conn.Open "Provider=SQLOLEDB;Data Source=PANDURO_LAP\SQLEXPRESS;Initial Catalog=HubLectura;Integrated Security=SSPI;"
    
    MsgBox "Conexión exitosa a la base de datos.", vbInformation
    
    ' Configurar ListView
    With ListView1
        .View = lvwReport
        .FullRowSelect = True
        .GridLines = True
        .ColumnHeaders.Clear
        .ColumnHeaders.Add 1, "colTitulo", "Título", 2000
        .ColumnHeaders.Add 2, "colAutor", "Autor", 2000
        .ColumnHeaders.Add 3, "colGenero", "Género", 1500
        .ColumnHeaders.Add 4, "colCalificacion", "Calificación", 1000
        .ColumnHeaders.Add 5, "colEstado", "Estado", 2000
        .ColumnHeaders.Add 6, "colUsuario", "Usuario ID", 1000
    End With
    
    ' Cargar datos desde la base de datos
    Call CargarLibrosBD
    
    Exit Sub

errorConexion:
    MsgBox "Error al conectar con la base de datos: " & Err.Description, vbCritical
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ' Cerrar recordset si está abierto
    If Not rs Is Nothing Then
        If rs.State = adStateOpen Then rs.Close
        Set rs = Nothing
    End If
    
    ' Cerrar conexión si está abierta
    If Not conn Is Nothing Then
        If conn.State = adStateOpen Then conn.Close
        Set conn = Nothing
    End If
End Sub


