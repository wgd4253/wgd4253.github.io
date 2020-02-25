<%
      '직접 접근.. 즉 실행하면 에러가 나.. add.asp -> addOK.asp 순으로 접근이 되어야 에러가 안나.

      txtname = Request.Form("txtname")
      txtemail = Request.Form("txtemail")
      txtcomment = Request.Form("txtcomment")

 

      '연결문자열.. add.asp 페이지에서 데이터를 받아와서 DB에 넣는 작업을 할꺼야.
      Set conn = Server.CreateObject("ADODB.Connection")
      conn.Open Application("conStr")

 

      'Insert 쿼리구문을 만들자
      '나이는 value ( ) 에서 홑따음표를 넣으면 안됨.. int type이니깐..
      query = "Insert into wed(txtname,txtemail,txtcomment) values('" & txtName & "', '" & txtemail & "', '" & txtcomment & "')"
      'Response.Write query
      'Response.End

 

      '반환값이 없으면 이렇게 해도 실행이 된다.
      conn.Execute query

 

      conn.Close
      Set conn = Nothing

%>

 

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
 <head>
  <title> new document </title>
        <meta name="generator" content="editplus" />
        <meta name="author" content="" />
        <meta name="keywords" content="" />
        <meta name="description" content="" />
 </head>


 <body>
  <script type="text/javascript">
  <!--
       alert("입력성공!!");


       // 페이지 이동!!
      <!-- location.href = "member.asp";  // 회원의 목록을 보여줄 페이지야. -->
      //-->
  </script>   
 </body>
</html>

