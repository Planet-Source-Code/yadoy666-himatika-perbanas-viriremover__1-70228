Attribute VB_Name = "Database"
 Public namaVirus As String, CrcVirus As String 'deklarasi variabel global untuk nama dan CRC virus


 Public Function cariDatabase(Crc As String, namaFileDB As String) As Boolean
  Dim lineStr As String, tmp() As String 'variabel penampung untuk isi file
   On Error Resume Next
   Open namaFileDB For Input As #1 'buka file dengan mode input
    Do
     Line Input #1, lineStr
     tmp = Split(lineStr, "=") 'pisahkan isi file bedasarkan pemisah karakter '='
     namaVirus = tmp(1) 'masukkan namavirus ke variabel dari array
     CrcVirus = tmp(0) 'masukkan Crcvirus ke variabel dari array
     If CrcVirus = Crc Then 'bila CRC perhitungan cocok/match dengan database
      cariDatabase = True 'kembalikan nilai TRUE
      Exit Do 'keluar dari perulangan
     End If
    Loop Until EOF(1)
   Close #1
 End Function




