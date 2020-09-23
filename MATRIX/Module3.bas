Attribute VB_Name = "Module3"
Public Type tData
    Key As String
    Value As String
    End Type


Public Function FindArtist(theArtist As String)
    Dim TheData As String, TheSize As Integer, X As Integer, tmpVar As String

    TheSize = Len("a=search&p=1&s=" + theArtist + "&l=artist")
    tmpVar = "POST " & "/cgi-exe/am.cgi" & " HTTP/1.1" & vbCrLf & _
    "Host: www.letssingit.com" & vbCrLf & _
    "Content-Type: application/x-www-form-urlencoded" & vbCrLf & _
    "Accept-Encoding: gzip, deflate" & vbCrLf & _
    "Content-Length: " & TheSize & vbCrLf & _
    "Connection: Keep-Alive" & vbCrLf & vbCrLf & _
    "a=search&p=1&s=" + theArtist + "&l=artist" & vbCrLf

    FindArtist = tmpVar
    End Function


Public Function FormGET(File As String, Data() As tData)
    Dim TheData As String, X As Integer, tmpVar As String


    For X = 0 To UBound(Data())
        TheData = TheData & Data(X).Key & "=" & Data(X).Value & "&"
    Next
    TheData = Mid$(TheData, 1, Len(TheData) - 1)
    tmpVar = "GET " & File & "?" & TheData & " HTTP/1.1" & vbCrLf & _
    "Accept-Encoding: gzip, deflate" & vbCrLf & _
    "Connection: Keep-Alive" & vbCrLf


    FormGET = tmpVar
    End Function


Public Function FormData(Key As String, Value As String) As tData 'saves a few lines....


    FormData.Key = Key


        FormData.Value = Value
End Function


