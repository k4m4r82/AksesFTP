VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Demo Load Gambar via FTP"
   ClientHeight    =   4290
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4275
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4290
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdLoadFoto 
      Caption         =   "Tampilkan Foto"
      Height          =   495
      Left            =   2055
      TabIndex        =   4
      Top             =   2520
      Width           =   1335
   End
   Begin VB.PictureBox picSiswa 
      Height          =   1575
      Left            =   120
      ScaleHeight     =   1515
      ScaleWidth      =   1755
      TabIndex        =   3
      Top             =   2520
      Width           =   1815
   End
   Begin VB.CommandButton cmdDaftarFile 
      Caption         =   "Daftar File"
      Height          =   495
      Left            =   2055
      TabIndex        =   2
      Top             =   480
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Daftar File"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDaftarFile_Click()
    Dim objFTP      As FTPClass
    Dim objFTPFile  As FTPFileClass
    
    Dim serverName  As String
    Dim userName    As String
    Dim password    As String
    
    serverName = "192.168.0.6"
    userName = "k4m4r82"
    password = "rahasia"
    
    Set objFTP = New FTPClass
    If objFTP.OpenFTP(serverName, userName, password) Then
        If objFTP.SetCurrentFolder("/") Then
            For Each objFTPFile In objFTP.Files
                List1.AddItem objFTPFile.FileName
            Next
        End If
        objFTP.CloseFTP
        
    Else
        'TODO : tampilkan pesan gagal membuka port FTP
    End If
    Set objFTP = Nothing
End Sub

Private Sub cmdLoadFoto_Click()
    Dim objFTP      As FTPClass
    
    Dim serverName  As String
    Dim userName    As String
    Dim password    As String
    Dim foto        As String
    
    serverName = "192.168.0.6"
    userName = "k4m4r82"
    password = "rahasia"
    
    foto = "02024112.jpg" 'contoh nama file gambar yang ingin ditampilkan
    
    Set objFTP = New FTPClass
    If objFTP.OpenFTP(serverName, userName, password) Then
        If objFTP.SetCurrentFolder("/") Then
            If objFTP.FileExists(foto) Then 'cek dulu file gambarnya dan jika ada...
                Call objFTP.GetFile(foto, App.Path & "\" & foto, True) 'download file gambarnya ke komputer lokal
                picSiswa.Picture = LoadPicture(App.Path & "\" & foto) 'baru ditampilkan
            End If
        End If
        objFTP.CloseFTP
        
    Else
        'TODO : tampilkan pesan gagal membuka port FTP
    End If
    Set objFTP = Nothing
End Sub
