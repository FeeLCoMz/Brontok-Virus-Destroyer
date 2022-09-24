Attribute VB_Name = "ModulUtama"
Option Explicit

' *** Modul Utama ***

Public NamaUser As String
Public WinDir As String
Public UserProfile As String
Public FolderVirus As String
Public SeluruhInfo As String

' *** Fungsi Umum ***

Public Function Nama_Aplikasi() As String

    Nama_Aplikasi = App.Title & " v" & App.Major & "." & App.Minor & "." & App.Revision

End Function

Function FileExist(strPath As String) As Integer

    Dim lngRetVal As Long
    
    On Error Resume Next
    lngRetVal = Len(Dir$(strPath))
    
    If Err Or lngRetVal = 0 Then
        FileExist = False
    Else
        FileExist = True
    End If
    
End Function


Public Sub Bunuh(ByVal Virusnya As String)

Dim Bersih As Boolean

    Bersih = False

    Informasi "  » Menghapus " & Virusnya
    
    On Error Resume Next
    
    SetAttr Virusnya, GetAttr(Virusnya) And Not vbHidden
    SetAttr Virusnya, GetAttr(Virusnya) And Not vbReadOnly
    SetAttr Virusnya, GetAttr(Virusnya) And Not vbSystem
    
    If FileExist(Virusnya) Then
        
        If InStr(Virusnya, "*") <> 0 Then
            Kill Virusnya
            Bersih = True
        End If
       
        If Err.Number <> 0 Then
            If Err.Number = 5 And Err.Number <> 52 Then
                Kill Virusnya
                Bersih = True
            Else
                Informasi "     » " & Err.Number & " : " & Err.Description
                Informasi "     » File tidak dapat dihapus! Coba cek manual file tsb!"
            End If
        Else
            Kill Virusnya
            Bersih = True
        End If
    Else
        Informasi "     » File tidak ada!"
    End If
    
    If Bersih Then
        Informasi "     » Virus telah dihapus!"
    End If
        
End Sub

Public Sub HapusReg(SubKey As String, Entry As String)

    Informasi "      » Menghapus Entry Registry " & SubKey & "\" & Entry
    DeleteValue SubKey, Entry
    
End Sub

Public Sub UbahRegDWORD(SubKey As String, Entry As String, Nilai As Long)

    'Informasi "      » Mengubah Entry Registry " & SubKey & "\" & Entry
    SetDWORDValue SubKey, Entry, Nilai
    
End Sub

Public Sub UbahRegString(SubKey As String, Entry As String, Nilai As String)

    'Informasi "      » Mengubah Entry Registry " & Subkey & "\" & Entry
    SetStringValue SubKey, Entry, Nilai
    
End Sub

Public Sub Informasi(ByVal Infonya As String)
    
    SeluruhInfo = SeluruhInfo & Infonya & vbCrLf
    FDestroyer.Info.Text = SeluruhInfo
    FDestroyer.Info.SelStart = Len(SeluruhInfo)
    FDestroyer.Info.Refresh
    
End Sub

Public Sub Hajar_Virus()

    Dim Status As Boolean
    Dim JmlVirMem As Integer

    NamaUser = Environ("UserName")
    WinDir = Environ("Windir")
    UserProfile = Environ("UserProfile")
    FolderVirus = UserProfile & AppData

    Status = False
    Informasi ""
    SeluruhInfo = ""
    JmlVirMem = 0
    
    ' *** Mulai ***
    ' *** Penghapusan Virus dari Memory ***
    
    Informasi "*** " & Now & " ***"
    Informasi ""
    Informasi "Pencarian dan Penghapusan Virus di Memory..."
    
    Do While FindWindow(vbNullString, NamaProsesVirus) <> 0
        JmlVirMem = JmlVirMem + 1
        WindowHandle FindWindow(vbNullString, NamaProsesVirus), 0
        Status = True
    Loop
    
    Do While FindWindow(vbNullString, NamaProsesVirus2) <> 0
        JmlVirMem = JmlVirMem + 1
        WindowHandle FindWindow(vbNullString, NamaProsesVirus2), 0
        Status = True
    Loop
    
    If Status = True Then
        Informasi "  » " & JmlVirMem & " Virus Brontok ditemukan di memory dan telah dibersihkan (Get Out, Bro!)"
        MsgBox JmlVirMem & " Virus Brontok ditemukan di memory dan telah dibersihkan (Get Out, Bro!)", vbInformation
    Else
        Informasi "  » Virus Brontok tidak ada di memory"
        MsgBox "Virus Brontok tidak ada di memory", vbInformation
    End If
    
    ' *** Perbaikan Registry yang dimodifikasi oleh Virus ***
    
    Informasi ""
    Informasi "Perbaikan Registry..."
    Informasi "  » Mengembalikan menu Folder Option pada Explorer..."
    UbahRegDWORD ExPol, "NoFolderOptions", 0
    Informasi "  » Mengaktifkan kembali Registry Editor..."
    UbahRegDWORD SysPol, "DisableRegistryTools", 0
    Informasi "  » Menyembuhkan Shell Windows..."
    UbahRegString WinLogon, "Shell", "Explorer.exe"
    Informasi "  » Menghapus Shell Alternatif Virus pada Safe Mode... "
    HapusReg CtrSet1, "AlternateShell"
    HapusReg CtrSet2, "AlternateShell"
    Informasi "  » Menghapus Entry Startup Virus ..."
    HapusReg StartupRun, "br11013on"
    HapusReg StartupRun, "br6657on"
    HapusReg StartupRun, "RakyatKelaparan"
    HapusReg StartupRun, "KesenjanganSosial"
    HapusReg StartupRun, "Bron-Spizaetus"
    HapusReg StartupRun, "Tok-Cirrhatus"
    HapusReg StartupRun, "Tok-Cirrhatus-4995"
    HapusReg StartupRun, "smss" 'Brontok.A[16]
    HapusReg StartupRun2, "Tok-Cirrhatus"
    HapusReg StartupRun2, "Tok-Cirrhatus-4995"
    
    ' *** Pembunuhan Virus ***
    
    Informasi ""
    Informasi "Pembersihan Master Virus..."
    Bunuh WinDir & "\shellnew\RakyatKelaparan.exe"
    Bunuh WinDir & "\shellnew\ElnorB.exe" 'Brontok.A[16]
    Bunuh WinDir & "\Eksplorasi.exe" 'Brontok.A[16]
    Bunuh WinDir & "\System32\" & NamaUser & "'s Setting.scr" 'Brontok.A[16]
    Bunuh WinDir & "\system32\cmd-brontok.exe"
    Bunuh WinDir & "\KesenjanganSosial.exe"
    Bunuh UserProfile & "\Start Menu\Programs\Startup\Empty.pif"
    Bunuh UserProfile & "\Templates\*nendangbro*.*"
    Bunuh FolderVirus & "\smss.exe"
    Bunuh FolderVirus & "\services.exe"
    Bunuh FolderVirus & "\lsass.exe"
    Bunuh FolderVirus & "\inetinfo.exe"
    Bunuh FolderVirus & "\csrss.exe"
    Bunuh FolderVirus & "\br11013on.exe"
    Bunuh FolderVirus & "\br6657on.exe"
    Bunuh FolderVirus & "\br*on.exe"
    Bunuh FolderVirus & "\svchost.exe"
    Bunuh FolderVirus & "\winlogon.exe"
    
    ' *** Penghapusan jejak Virus ***
    
    Informasi ""
    Informasi "Penghapusan Jejak Virus..."
    Bunuh FolderVirus & "\BronFoldNetDomList.txt"
    Bunuh FolderVirus & "\Kosong.Bron.Tok.txt"
    On Error Resume Next
    RmDir FolderVirus & "\Loc.Mail.Bron.Tok"
    RmDir FolderVirus & "\Ok-SendMail-Bron-tok"
    Informasi ""
    Informasi "*** Selesai ***"
    Informasi ""
    Informasi "*** " & Now & " ***"
    Informasi ""
    
    MsgBox "Pembersihan Master Virus Brontok telah selesai." & vbCrLf & _
            "Silahkan lakukan Scanning untuk mencari folder-folder yang terinfeksi Virus." & vbCrLf & _
            "Untuk melakukan Scanning, klik tombol Scan di Program!", vbInformation
    
    ' *** Selesai ***
    
End Sub
