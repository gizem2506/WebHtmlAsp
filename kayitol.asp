<!DOCTYPE html>
<html style="font-size: 16px;">
  <head>
  </head>
    <%
Set baglanti = server.createobject("adodb.connection")
baglanti.open "Provider=Microsoft.Jet.oledb.4.0;Data Source=" & Server.MapPath("Sporsalonu.mdb")
%>
 
 
<%
 
email = Request.Form("email")
sifre = Request.Form("sifre")
 
Set Kaydet = Server.CreateObject("adodb.recordset")
sql="Select * From uyegirisi"
Kaydet.open sql , Baglanti ,1,3
 
 
Kaydet.AddNew
Kaydet("gmail") = email
Kaydet("sifre") = sifre
 
 
 
Kaydet.Update
Kaydet.Close
Set Kaydet = Nothing
Baglanti.Close
Set Baglanti = Nothing
%>
 
 
<%
Response.Write "<script language='JavaScript'>alert('Başarı İle Kaydedildi...');</script>"
%>

    
    
 
  <body>
 <a href="index.html">Anasayfa'ya geri dön</a>
 <a href="index.html">Giriş Yap</a>
  </body>
  </html>