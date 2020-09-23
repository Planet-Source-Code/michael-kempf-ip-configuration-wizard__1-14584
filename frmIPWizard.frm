VERSION 5.00
Begin VB.Form frmIPWizard 
   Caption         =   "IP Configuration Wizard"
   ClientHeight    =   6360
   ClientLeft      =   3945
   ClientTop       =   2265
   ClientWidth     =   5955
   Icon            =   "frmIPWizard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6360
   ScaleWidth      =   5955
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
   Begin VB.CommandButton cmdRenew 
      Caption         =   "Re&new"
      Height          =   390
      Left            =   3495
      TabIndex        =   27
      Top             =   5850
      Width           =   915
   End
   Begin VB.CommandButton cmdRelease 
      Caption         =   "Relea&se"
      Height          =   390
      Left            =   2520
      TabIndex        =   26
      Top             =   5850
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "&OK"
      Height          =   390
      Left            =   1545
      TabIndex        =   25
      Top             =   5850
      Width           =   915
   End
   Begin VB.Frame Frame2 
      Caption         =   "Host Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1590
      Left            =   75
      TabIndex        =   20
      Top             =   4120
      Width           =   5790
      Begin VB.CheckBox chkDHCPEnabled 
         Alignment       =   1  'Right Justify
         Caption         =   "DHCP Enabled"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   900
         TabIndex        =   30
         Top             =   1125
         Width           =   1590
      End
      Begin VB.CommandButton cmdChgDNSServer 
         Caption         =   "..."
         Height          =   315
         Left            =   5325
         TabIndex        =   28
         Top             =   750
         Width           =   315
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "DNS Servers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   825
         TabIndex        =   24
         Tag             =   "NOCLEAR"
         Top             =   750
         Width           =   1290
      End
      Begin VB.Label lblDNSServers 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2175
         TabIndex        =   23
         Top             =   750
         Width           =   3090
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Host Name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   825
         TabIndex        =   22
         Tag             =   "NOCLEAR"
         Top             =   375
         Width           =   1290
      End
      Begin VB.Label lblHostName 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2175
         TabIndex        =   21
         Top             =   375
         Width           =   3465
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Ethernet Adapter Information"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   4140
      Left            =   75
      TabIndex        =   0
      Top             =   0
      Width           =   5790
      Begin VB.ComboBox cboAdapter 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         ItemData        =   "frmIPWizard.frx":08CA
         Left            =   2175
         List            =   "frmIPWizard.frx":08CC
         TabIndex        =   19
         Top             =   300
         Width           =   3465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Adapter"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   180
         TabIndex        =   29
         Tag             =   "NOCLEAR"
         Top             =   360
         Width           =   1965
      End
      Begin VB.Label lblLeaseExpires 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2175
         TabIndex        =   18
         Top             =   3675
         Width           =   3465
      End
      Begin VB.Label lblLeaseObtained 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2175
         TabIndex        =   17
         Top             =   3298
         Width           =   3465
      End
      Begin VB.Label lblSecondaryWins 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2175
         TabIndex        =   16
         Top             =   2927
         Width           =   3465
      End
      Begin VB.Label lblPrimaryWins 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2175
         TabIndex        =   15
         Top             =   2556
         Width           =   3465
      End
      Begin VB.Label lblDHCPServer 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2175
         TabIndex        =   14
         Top             =   2185
         Width           =   3465
      End
      Begin VB.Label lblDefGateway 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2175
         TabIndex        =   13
         Top             =   1814
         Width           =   3465
      End
      Begin VB.Label lblSubnetMask 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2175
         TabIndex        =   12
         Top             =   1443
         Width           =   3465
      End
      Begin VB.Label lblIPAddress 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2175
         TabIndex        =   11
         Top             =   1072
         Width           =   3465
      End
      Begin VB.Label lblAdapterAddress 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2160
         TabIndex        =   10
         Tag             =   "NOCLEAR"
         Top             =   705
         Width           =   3465
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Adapter Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   9
         Tag             =   "NOCLEAR"
         Top             =   750
         Width           =   1965
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Lease Expires"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   450
         TabIndex        =   8
         Tag             =   "NOCLEAR"
         Top             =   3645
         Width           =   1665
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Lease Obtained"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   450
         TabIndex        =   7
         Tag             =   "NOCLEAR"
         Top             =   3270
         Width           =   1665
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Secondary WINS Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   6
         Tag             =   "NOCLEAR"
         Top             =   2910
         Width           =   2040
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Primary WINS Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   75
         TabIndex        =   5
         Tag             =   "NOCLEAR"
         Top             =   2550
         Width           =   2040
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "DHCP Server"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   825
         TabIndex        =   4
         Tag             =   "NOCLEAR"
         Top             =   2190
         Width           =   1290
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Default Gateway"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         TabIndex        =   3
         Tag             =   "NOCLEAR"
         Top             =   1830
         Width           =   1890
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Subnet Mask"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   825
         TabIndex        =   2
         Tag             =   "NOCLEAR"
         Top             =   1470
         Width           =   1290
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "IP Address"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   150
         TabIndex        =   1
         Tag             =   "NOCLEAR"
         Top             =   1110
         Width           =   1965
      End
   End
End
Attribute VB_Name = "frmIPWizard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************
'*                IP Wizard                    *
'*  Copyright © 2000-2001 , Kemtech Software   *
'*             Michael J. Kempf                *
'***********************************************


Dim gintNetworkAdapter As Integer

Private Namespace As SWbemServices
Private Method As SWbemMethod
'Default System Address
Private Const SystemAddress = "127.0.0.0"


Private Sub cboAdapter_Click()
'ClearAll Controls
    Call clearall
'Get Configuration for the selected adapter
    Call GetIPConfig(cboAdapter.ItemData(cboAdapter.ListIndex))
'Set a Global Variable for the selected adapter
    gintNetworkAdapter = cboAdapter.ItemData(cboAdapter.ListIndex)
End Sub

Private Sub cmdChgDNSServer_Click()
On Error Resume Next
Dim Adapter As SWbemObject
'Change DNS Severs
Set Adapter = GetObject("winmgmts:Win32_NetworkAdapterConfiguration=" & gintNetworkAdapter & "")
If lblDNSServers.Tag = 0 Then
    lblDNSServers = Adapter.DNSServerSearchOrder(1)
    lblDNSServers.Tag = 1
Else
     lblDNSServers = Adapter.DNSServerSearchOrder(0)
     lblDNSServers.Tag = 0
End If

Set Adapter = Nothing
End Sub

Private Sub cmdOK_Click()
    End
End Sub

Private Sub cmdRelease_Click()
On Error GoTo ErrorHandler

    Dim ReleaseLease As SWbemObject
    Set ReleaseLease = GetObject("winmgmts:Win32_NetworkAdapterConfiguration=" & gintNetworkAdapter & "")
'Release DHCP Address
    ReleaseLease.ReleaseDHCPLease
'Clear Info
    Call clearall
'Get Adapter Congiguration
    Call GetIPConfig(gintNetworkAdapter)
    
    Set ReleaseLease = Nothing
    
Exit Sub
ErrorHandler:
MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & "Description: " & Err.Description, vbCritical, "IP Wizard"
 Set ReleaseLease = Nothing
End Sub
 Sub GetAdapterData()
On Error Resume Next
    
    Dim Adapter As SWbemObject
    Dim NIC As SWbemObject
   
'Enumerate the instances
    Set NICS = Namespace.InstancesOf("Win32_NetworkAdapterConfiguration")
        For Each NIC In NICS
        ' Use the RelPath property of the instance path
            Set NIC = GetObject("winmgmts:Win32_NetworkAdapterConfiguration=Adapter.Path_.RelPath")
        'Fill the Adapter with all the adapters
            cboAdapter.AddItem NIC.Description
            cboAdapter.ItemData(cboAdapter.NewIndex) = NIC.Index
        Next NIC
         cboAdapter.Text = cboAdapter.List(0)
         
    Set NICS = Nothing
End Sub

Private Sub cmdRenew_Click()
Me.MousePointer = vbHourglass

On Error GoTo RenewLease

    Dim RenewLease As SWbemObject
    Set RenewLease = GetObject("winmgmts:Win32_NetworkAdapterConfiguration=" & gintNetworkAdapter & "")
'Renew DHCP Lease for the selected adapter
      RenewLease.RenewDHCPLease
'Clear Info
    Call clearall
'Get Adapter Congiguration
    Call GetIPConfig(gintNetworkAdapter)
    
    Set RenewLease = Nothing
Me.MousePointer = vbDefault

Exit Sub
RenewLease:
MsgBox "Error: " & Err.Number & vbCrLf & vbCrLf & "Description: " & Err.Description, vbCritical, "IP Wizard"
 Set RenewLease = Nothing
 Me.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Set Namespace = GetObject("winmgmts:")
'Get Machine Adapters
    GetAdapterData
'Load Default Adapter Information
    Call GetIPConfig(0)
End Sub

Sub GetIPConfig(intNetworkAdapter As Integer)
On Error Resume Next
Dim Adapter As SWbemObject
'Clear Adapter Address
    lblAdapterAddress.Caption = ""
Set Adapter = GetObject("winmgmts:Win32_NetworkAdapterConfiguration=" & intNetworkAdapter & "")
'Fill Adapter and Host Information
   With Adapter
'MAC Address
        lblAdapterAddress.Caption = .MACAddress
'TCP/IP Address
        If .IPADDRESS(intNetworkAdapter) = SystemAddress Then
            lblIPAddress.Caption = "0.0.0.0"
        ElseIf .IPADDRESS(intNetworkAdapter) = "" Then
            lblIPAddress.Caption = "0.0.0.0"
        Else
            lblIPAddress.Caption = .IPADDRESS(intNetworkAdapter)
        End If
'Subnet Mask
        If .IPSubnet(intNetworkAdapter) = SystemAddress Then
            lblSubnetMask.Caption = "0.0.0.0"
        ElseIf .IPSubnet(intNetworkAdapter) = "" Then
            lblSubnetMask.Caption = "0.0.0.0"
        Else
           lblSubnetMask.Caption = .IPSubnet(intNetworkAdapter)
        End If
'Default Gateway
        If Not .DefaultIPGateway(intNetworkAdapter) = SystemAddress Then lblDefGateway.Caption = .DefaultIPGateway(intNetworkAdapter) _
        Else lblDefGateway.Caption = ""
'DHCP Server
        lblDHCPServer.Caption = .DHCPServer
'Primary WINS Server
        If Not .WINSPrimaryServer = SystemAddress Then lblPrimaryWins.Caption = .WINSPrimaryServer _
        Else lblPrimaryWins.Caption = ""
'Secondary WINS Server
        If Not .WINSSecondaryServer = SystemAddress Then lblSecondaryWins.Caption = .WINSSecondaryServer _
        Else lblSecondaryWins.Caption = ""
'DHCP Lease Obtained Day and Time
        If lblIPAddress.Caption = "0.0.0.0" Then lblLeaseObtained.Caption = "" _
        Else lblLeaseObtained.Caption = ParseDate(.DHCPLeaseObtained)
'DHCP Lease Expire Day and Time
        If lblIPAddress.Caption = "0.0.0.0" Then lblLeaseExpires.Caption = "" _
        Else lblLeaseExpires.Caption = ParseDate(.DHCPLeaseExpires)
'Host Name and Domain
        lblHostName.Caption = .DNSHostName & "." & .DNSdomain
'DNS Server(s)
        lblDNSServers.Caption = .DNSServerSearchOrder(intNetworkAdapter)
        lblDNSServers.Tag = 0

'DHCP Enabled
        If .DHCPEnabled = True Then
            chkDHCPEnabled.Value = 1
        Else
            chkDHCPEnabled.Value = 0
        End If
'Disable / Enable Buttons
    If lblIPAddress.Caption = "0.0.0.0" Then
        cmdRenew.Enabled = True
        cmdRelease.Enabled = False
    ElseIf chkDHCPEnabled.Value = 0 Then
        cmdRenew.Enabled = False
        cmdRelease.Enabled = False
    Else
        cmdRenew.Enabled = False
        cmdRelease.Enabled = True
    End If
    
    End With
Set Adapter = Nothing
End Sub
Function ParseDate(strDateString As String) As String

Dim strParseDate As String
Dim strParseTime As String

If strDateString = "" Then
    ParseDate = ""
Else
'Parse Date and Time from string
    strParseDate = Left(strDateString, 8)
    strParseTime = Right(Left(strDateString, 14), 6)
'Parse to readable Date and time
    ParseDate = Mid(strParseDate, 5, 2) & "/" & Mid(strParseDate, 7, 2) & "/" & Mid(strParseDate, 1, 4) & "   " & Format(Mid(strParseTime, 1, 2) & ":" & Mid(strParseTime, 3, 2) & ":" & Mid(strParseTime, 5, 2), "hh:mm:ss AMPM")
End If
End Function

Sub clearall()
Dim Control
'Loop through each contol on the form and clear it
    For Each Control In Me.Controls
        If TypeOf Control Is Label Then
            If Not Control.Tag = "NOCLEAR" Then
               Control.Caption = ""
            End If
        End If
    If TypeOf Control Is CheckBox Then Control.Value = 0
    Next Control
        
    lblIPAddress.Caption = "0.0.0.0"
    lblSubnetMask.Caption = "0.0.0.0"
End Sub


