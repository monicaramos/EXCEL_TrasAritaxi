VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMensajes 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Mensajes"
   ClientHeight    =   4740
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8730
   Icon            =   "frmMensajes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   8730
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FrameAlbaranesErroneos 
      Height          =   4620
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   8655
      Begin MSComctlLib.ListView ListView7 
         Height          =   3135
         Left            =   240
         TabIndex        =   2
         Top             =   540
         Width           =   8100
         _ExtentX        =   14288
         _ExtentY        =   5530
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7230
         TabIndex        =   1
         Top             =   3960
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Registros con errores:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   5
         Left            =   270
         TabIndex        =   3
         Top             =   210
         Width           =   7215
      End
   End
End
Attribute VB_Name = "frmMensajes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'====================== VBLES PUBLICAS ================================

Public Event DatoSeleccionado(CadenaSeleccion As String)

Public OpcionMensaje As Byte
'1 .- Entradas de bascula que no tienen CRFID

Public cadWHERE As String 'Cadena para pasarle la WHERE de la SELECT de los cobros pendientes o de Pedido(para comp. stock)
                          'o CodArtic para seleccionar los Nº Series
                          'para cargar el ListView
                          
Public cadWHERE2 As String
Public Campo As String
Public Cadena As String ' sql para cargar el listview
Public vCampos As String 'Articulo y cantidad Empipados para Nº de Series
                         'Tambien para pasar el nombre de la tabla de lineas (sliped, slirep,...)
                         'Dependiendo desde donde llamemos, de Pedidos o Reparaciones


'====================== VBLES LOCALES ================================

Dim PulsadoSalir As Boolean 'Solo salir con el boton de Salir no con aspa del form
Dim PrimeraVez As Boolean

'Para los Nº de Serie
Dim TotalArray As Integer
Dim codArtic() As String
Dim cantidad() As Integer


Private Sub cmdSalir_Click()
    Unload Me
End Sub





Private Sub Form_Load()
Dim H As Integer, W As Integer
Dim cad As String
On Error Resume Next

    Me.FrameAlbaranesErroneos.visible = False
    PulsadoSalir = True
    PrimeraVez = True
    
    Select Case OpcionMensaje
        Case 1 ' Entradas de excel erroneas
            PonerFrameAlbaranesErroneosVisible True, H, W
            CargarListaAlbaranesErroneos Cadena
            Me.Label1(3).Caption = "Registros erróneos: "
            Me.CmdSalir.SetFocus
        
    
    End Select
    'Me.cmdCancel(indFrame).Cancel = True
    Me.Height = H + 350
    Me.Width = W + 70
    
    If Err.Number <> 0 Then Err.Clear
End Sub


Private Sub PonerFrameAlbaranesErroneosVisible(visible As Boolean, ByRef H As Integer, ByRef W As Integer)
'Pone el Frame de Cobros Pendientes Visible y Ajustado al Formulario, y visualiza los controles
'necesario para el Informe

        
    H = 4600
    W = 8655
        
    PonerFrameVisible Me.FrameAlbaranesErroneos, visible, H, W

End Sub


Private Sub PonerFrameVisible(ByRef vFrame As Frame, visible As Boolean, H As Integer, W As Integer)
'Pone el Frame Visible y Ajustado al Formulario, y visualiza los controles
    
        vFrame.visible = visible
        If visible = True Then
            'Ajustar Tamaño del Frame para ajustar tamaño de Formulario al del Frame
            vFrame.Top = -90
            vFrame.Left = 0
            vFrame.Width = W
            vFrame.Height = H
        End If
End Sub

Private Sub CargarListaAlbaranesErroneos(SQL As String)
'Muestra la lista Detallada de entradas que no tienen CRFID
'en un ListView
Dim RS As ADODB.Recordset
Dim ItmX

    On Error GoTo ECargarList


    Set RS = New ADODB.Recordset
    RS.Open SQL, Conn, adOpenForwardOnly, adLockPessimistic, adCmdText
    
    If Not RS.EOF Then
        'Los encabezados
        ListView7.ColumnHeaders.Clear

        ListView7.ColumnHeaders.Add , , "Nº Albaran", 1000
        ListView7.ColumnHeaders.Add , , "Código", 1000, 2
        ListView7.ColumnHeaders.Add , , "Socio", 2000, 0
        ListView7.ColumnHeaders.Add , , "Error", 3000, 0
    
        While Not RS.EOF
            Set ItmX = ListView7.ListItems.Add
            ItmX.Text = Format(RS!numalbar, "000000")
            ItmX.SubItems(1) = Format(RS!codvarie, "000000")
            ItmX.SubItems(2) = Format(RS!codsocio, "000000")
            
            
            Select Case RS!situacion
                Case 1
                    ItmX.SubItems(3) = "No existe Albarán."
                Case 2
                    ItmX.SubItems(3) = "Entrada duplicada."
                Case 11
                    ItmX.SubItems(3) = "No existe Calidad."
                Case 12
                    ItmX.SubItems(3) = "No cuadran kilos."
            End Select
            
            RS.MoveNext
        Wend
    End If
    RS.Close
    Set RS = Nothing

ECargarList:
    If Err.Number <> 0 Then
        Err.Clear
    End If
End Sub


Private Sub Form_Unload(Cancel As Integer)
    If PulsadoSalir = False Then Cancel = 1
End Sub



Private Function ObtenerTamanyosArray() As Boolean
'Para el frame de los Nº de Serie de los Articulos
'En cada indice pone en CodArtic(i) el codigo del articulo
'y en Cantidad(i) la cantidad solicitada de cada codartic
Dim I As Integer, J As Integer

    ObtenerTamanyosArray = False
    'Primero a los campos de la tabla
    TotalArray = -1
    J = 0
    Do
        I = J + 1
        J = InStr(I, vCampos, "·")
        If J > 0 Then TotalArray = TotalArray + 1
    Loop Until J = 0
    
    If TotalArray < 0 Then Exit Function
    
    'Las redimensionaremos
    ReDim codArtic(TotalArray)
    ReDim cantidad(TotalArray)
    
    ObtenerTamanyosArray = True
End Function


Private Function SeparaCampos() As Boolean
'Para el frame de los Nº de Serie de los Articulos
Dim Grupo As String
Dim I As Integer
Dim J As Integer
Dim C As Integer 'Contador dentro del array

    SeparaCampos = False
    I = 0
    C = 0
    Do
        J = I + 1
        I = InStr(J, vCampos, "·")
        If I > 0 Then
            Grupo = Mid(vCampos, J, I - J)
            'Y en la martriz
            InsertaGrupo Grupo, C
            C = C + 1
        End If
    Loop Until I = 0
    SeparaCampos = True
End Function


Private Sub InsertaGrupo(Grupo As String, Contador As Integer)
Dim J As Integer
Dim cad As String

    J = 0
    cad = ""
    
    'Cod Artic
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
        J = 1
    End If
    codArtic(Contador) = cad
    
    'Cantidad
    J = InStr(1, Grupo, "|")
    If J > 0 Then
        cad = Mid(Grupo, 1, J - 1)
        Grupo = Mid(Grupo, J + 1)
    Else
        cad = Grupo
        Grupo = ""
    End If
    cantidad(Contador) = cad
End Sub





