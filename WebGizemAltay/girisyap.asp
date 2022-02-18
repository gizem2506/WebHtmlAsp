<!DOCTYPE html>
<html style="font-size: 16px;">
  <head>
  </head>
    <%
Set baglanti = server.createobject("adodb.connection")
baglanti.open "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & Server.MapPath("Sporsalonu.mdb")
%>
 
 <%

UserName = Server.HTMLEncode(Request.Form("UserName"))
UserName = Replace(UserName, "'", "", 1, -1, 1) 

Password = Server.HTMLEncode(Request.Form("Password"))
Password = Replace(Password, "'", "", 1, -1, 1)


If Len(UserName) = 0 or Len(Password) = 0 Then

    Response.Write "Giris Yapabilmek icin Kullanici Adi ve Sifre Girmeniz Gerekmektedir.<br>"

    Response.Write ("<a href=""anasayfa.html"">Giris Ekranina Geri Donun...</a>")

    Response.End()

Else

Set stDosya = Server.CreateObject("ADODB.Connection")
stDosya.Open ("DRIVER={Microsoft Access Driver (*.mdb)};DBQ=" & Server.MapPath("Sporsalonu.mdb"))

Set rsUser = Server.CreateObject("ADODB.Recordset")
qUser ="SELECT * FROM uyegirisi WHERE gmail='"&UserName&"'"
rsUser.Open qUser, stDosya

    If rsUser.EOF Then

        stDosya.Close()
    Set stDosya = Nothing

        Response.Write ("Kullanici Adi Bulunamadi...")
    Response.Write ("<br><br><a href=""index.html"">Giris Ekranina Geri Donun...</a>")

    Else

        If Password = rsUser("sifre") Then

            Login        = 1
        User         = rsUser("gmail")
           
        stDosya.Close()
    Set stDosya = Nothing
        
           Session("Active")   = Login
             Session("gmail")     = User
   
        Response.Redirect("index.html")
        Else

       stDosya.Close()
        Set stDosya = Nothing

            Response.Write "Sifreyi Hatali Girdiniz..."
        Response.Write ("<br><br><a href=""index.html"">Giris Ekranina Geri Donun...</a>")

        End If

    End If

End If

%>
  <body>
 
 
<%
Response.Write " <a href='index.html'>Anasayfaya Donmek icin Tıklayın</a>"
%>
  </body>
  </html>
