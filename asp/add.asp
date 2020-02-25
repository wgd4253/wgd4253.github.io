[add.asp]

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
 <head>
  <title> 회원 정보 입력 </title>
       <meta name="generator" content="editplus" />
       <meta name="author" content="" />
       <meta name="keywords" content="" />
       <meta name="description" content="" />


       <link type="text/css" rel="stylesheet" href="member.css" /> ---> CSS 링크
       <script type="text/javascript" src="member.js"></script>     ---> JavaScript 함수 링크
 </head>

 

 <body>
      <!-- #include file="header.inc" -->
 <br />
 <h3>회원 정보 입력</h3>

 <!-- 쿼리를 통해 MSSQL에 Member 이란 테이블을 만들었어. 그럼 안에 내용을 넣어보자. -->
 <table width="400" bordercolor="black" border="1" style="border-collapse:collapse;" cellpadding="5">
  <!-- form태그도 눈에 안보이지만 공간을 차지하기때문에 자료가 안들어가는 부분에 넣었어. -->
  <form method="post" action="addOK.asp"> <!-- addOK.asp 에서 회원가입처리를 할꺼야. -->
  <tr>
    <td class="LabelTD" >이름</td>
    <td class="ContentTD">
          <input type="text" placeholder="Your name" name="txtname" size="15" />
    </td>
  </tr>

<!--  <tr>
    <td class="LabelTD">전화번호</td>
    <td class="ContentTD">
          <input type="text" name="txtTel1" size="3" />
          <input type="text" name="txtTel2" size="4" />
          <input type="text" name="txtTel3" size="4" />
    </td>
  </tr>
*-->
  <tr>
    <td class="LabelTD">코멘트</td>
    <td class="ContentTD">
          <input type="text" placeholder="Your Message" name="txtEmail" size="35" />
    </td>
  </tr>
<!--
  <tr>
    <td class="LabelTD">나이</td>
    <td class="ContentTD">
          <input type="text" name="txtAge" size="3" />
    </td>
  </tr>
 </table>
-->
 <p>
       <input type="button" value=" 가입하겠습니다. " onClick="Send();" />
 </p>

 </form>


      <!-- #include file="footer.inc" -->
 </body>
</html>
