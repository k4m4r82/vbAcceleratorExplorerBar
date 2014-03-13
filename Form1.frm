VERSION 5.00
Object = "{77EBD0B1-871A-4AD1-951A-26AEFE783111}#2.1#0"; "vbalExpBar6.ocx"
Object = "{396F7AC0-A0DD-11D3-93EC-00C0DFE7442A}#1.0#0"; "vbalIml6.ocx"
Begin VB.Form Form1 
   Caption         =   "Demo membuat menu dengan komponen vbAccelerator Explorer Bar"
   ClientHeight    =   7890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9315
   LinkTopic       =   "Form1"
   ScaleHeight     =   7890
   ScaleWidth      =   9315
   StartUpPosition =   2  'CenterScreen
   Begin vbalExplorerBarLib6.vbalExplorerBarCtl vbalExplorerBarCtl1 
      Align           =   4  'Align Right
      Height          =   7890
      Left            =   5775
      TabIndex        =   0
      Top             =   0
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   13917
      BackColorEnd    =   0
      BackColorStart  =   0
   End
   Begin vbalIml6.vbalImageList barIcons 
      Left            =   3600
      Top             =   1560
      _ExtentX        =   953
      _ExtentY        =   953
      IconSizeX       =   32
      IconSizeY       =   32
      ColourDepth     =   32
      Size            =   8824
      Images          =   "Form1.frx":0000
      Version         =   131072
      KeyCount        =   2
      Keys            =   "ÿ"
   End
   Begin vbalIml6.vbalImageList itemIcons 
      Left            =   3600
      Top             =   2280
      _ExtentX        =   953
      _ExtentY        =   953
      ColourDepth     =   32
      Size            =   2296
      Images          =   "Form1.frx":2298
      Version         =   131072
      KeyCount        =   2
      Keys            =   "ÿ"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'***************************************************************************
' MMMM  MMMMM  OMMM   MMMO    OMMM    OMMM    OMMMMO     OMMMMO    OMMMMO  '
'  MM    MM   MM MM    MMMO  OMMM    MM MM    MM   MO   OM    MO  OM    MO '
'  MM  MM    MM  MM    MM  OO  MM   MM  MM    MM   MO   OM    MO       OMO '
'  MMMM     MMMMMMMM   MM  MM  MM  MMMMMMMM   MMMMMO     OMMMMO      OMO   '
'  MM  MM        MM    MM      MM       MM    MM   MO   OM    MO   OMO     '
'  MM    MM      MM    MM      MM       MM    MM    MO  OM    MO  OM   MM  '
' MMMM  MMMM    MMMM  MMMM    MMMM     MMMM  MMMM  MMMM  OMMMMO   MMMMMMM  '
'                                                                          '
' K4m4r82's Laboratory                                                     '
' http://coding4ever.wordpress.com                                         '
'***************************************************************************

Private Const WARNA_BIRU_TUA    As Long = &H800000
Private Const WARNA_BIRU        As Long = &HED9564
Private Const WARNA_ABU_ABU     As Long = &HDEC4B0
Private Const WARNA_PUTIH       As Long = &H80000005

Private Function setBarMenu(ByVal explorerBar As Object, ByVal menuName As String, _
                            ByVal menuCaption As String, ByVal iconIndex As Long) As Object
                           
    Dim cBar As Object
    
    Set cBar = explorerBar.Bars.Add(, menuName, menuCaption)
    cBar.IsSpecial = True
    cBar.iconIndex = iconIndex
    cBar.TitleForeColor = WARNA_BIRU_TUA
    cBar.TitleForeColorOver = WARNA_BIRU_TUA
    cBar.TitleBackColorLight = WARNA_BIRU
    cBar.TitleBackColorDark = RGB(234, 241, 253)
    cBar.BackColor = WARNA_ABU_ABU
    
    Set setBarMenu = cBar
End Function

Private Sub setItemMenu(ByVal cBar As Object, ByVal menuName As String, ByVal menuCaption As String, ByVal iconIndex As Long)
    Dim cItem   As Object
    
    Set cItem = cBar.Items.Add(, menuName, menuCaption)
    With cItem
        .iconIndex = iconIndex
        .TextColor = WARNA_BIRU_TUA
        .TextColorOver = WARNA_PUTIH
    End With
End Sub

Private Sub addMenu(ByVal explorerBar As Object, ByVal barIcons As Object, ByVal itmIcons As Object)
    Dim rsMenuInduk As ADODB.Recordset
    Dim rsMenuAnak  As ADODB.Recordset
    Dim cBar        As Object
    
    Dim i           As Long
    Dim x           As Long
    Dim rowCount(1) As Long
    
    With explorerBar
        .UseExplorerStyle = False
        
        .Redraw = False
        
        .BackColorStart = WARNA_BIRU
        .BackColorEnd = WARNA_BIRU
         
        .ImageList = itmIcons.hIml
        .BarTitleImageList = barIcons.hIml
        
        'menampilkan menu induk
        strSql = "SELECT id, menu_name, menu_caption " & _
                 "FROM menu_induk " & _
                 "ORDER BY id"
        Set rsMenuInduk = openRecordset(strSql)
        If Not rsMenuInduk.EOF Then
            rowCount(0) = getRecordCount(rsMenuInduk)
            
            For i = 1 To rowCount(0)
                Set cBar = setBarMenu(explorerBar, rsMenuInduk("menu_name").Value, rsMenuInduk("menu_caption").Value, 0)
                
                'menampilkan menu anak
                strSql = "SELECT menu_name, menu_caption " & _
                         "FROM menu_anak " & _
                         "WHERE menu_induk_id = " & rsMenuInduk("id").Value & " " & _
                         "ORDER BY id"
                Set rsMenuAnak = openRecordset(strSql)
                If Not rsMenuAnak.EOF Then
                    rowCount(1) = getRecordCount(rsMenuAnak)

                    For x = 1 To rowCount(1)
                        Call setItemMenu(cBar, rsMenuAnak("menu_name").Value, rsMenuAnak("menu_caption").Value, 0)

                        rsMenuAnak.MoveNext
                    Next x
                End If
                Call closeRecordset(rsMenuAnak)
                
                rsMenuInduk.MoveNext
            Next i
        End If
        Call closeRecordset(rsMenuInduk)
        
        Set cBar = setBarMenu(explorerBar, "mnuKeluar", "Keluar", 1)
        Call setItemMenu(cBar, "mnuKeluarDrProgram", "Keluar dari Program", 1)
        
        .Redraw = True
    End With
End Sub

Private Sub Form_Load()
    Dim ret As Boolean
    
    ret = KonekToServer
    
    Me.BackColor = WARNA_BIRU
    
    Call addMenu(vbalExplorerBarCtl1, barIcons, itemIcons)
End Sub

Private Sub vbalExplorerBarCtl1_ItemClick(itm As vbalExplorerBarLib6.cExplorerBarItem)
    Select Case itm.Key
        Case "mnuBarang": 'TODO : tampilkan frmBarang disini
        Case "mnuCustomer"
        Case "mnuSupplier"
        Case "mnuPembelian"
        Case "mnuReturPembelian"
        Case "mnuPenjualan"
        Case "mnuBiayaOperasional"
        Case "mnuGajiKaryawan"
        Case "mnuLapPembelian"
        Case "mnuLapJthTempo"
        Case "mnuLapPenjualan"
    End Select
End Sub
