<%
      '1.  DB 접속을 위한 Connection객체 생성
      '연결문자열.. add.asp 페이지에서 데이터를 받아와서 DB에 넣는 작업을 할꺼야.
      Set conn = Server.CreateObject("ADODB.Connection")
      conn.Open Application("conStr")

 

      '2.  목록 가져오기
      query = "Select txtname, txtcomment from wed "
      '반환값이 있는 쿼리는 보통 RecordSet 객체를 생성해서 가져와.. ADO.NET의 reader 객체와 비슷해.. 
      Set rs = Server.CreateObject("ADODB.RecordSet")
        
 

      'rs.Open 실행시킬 쿼리, 실행시킬 대상.... 우리가 만든 conn이 Connection String 을 다 가지고 있어.
      rs.Open query, conn

 

      'rs(0)은 현재 커서가 가르키는 레코드의 0번째 행을 가져온거야.. 
      '근데 내가 다시 한 줄 찍어보지만... 커서는 여전히 같은 자리에 있기 때문에... 문제가 발생해
      '커서를 이동해야 다음 사람의 정보를 볼 수 있잖아... 이 때 등장하는 rs.MoveNext 야.. 고마운 녀석이내
      'Response.Write rs(0) & "-" & rs(1) & "-" & rs(2) & "<br>"
      'rs.MoveNext '--> 다음 사람의 정보를 보기위해 커서를 한 줄 내려준거야..
      'Response.Write rs(0) & "-" & rs(1) & "-" & rs(2) & "<br>"
      'rs.MoveNext '--> 현재 2명밖에 없어. 근데 다시 커서를 내려도 에러는 나지 않아.

 

      'rs.MoveNext '--> 근데 여기서는 에러가 또 나!!!! 장난하나.. -_-;;
       '--> 커서가 최상단(BOF), 최하단(EOF)야.. 커서가 최하단에서 더 내린다면 에러가 나..
%>

 

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
 <head>
  <title> new document </title>
       <link type="text/css" rel="stylesheet" href="member.css" />
       <script type="text/javascript" src="member.js"></script>
 </head>

 

 <body>
    <!-- #include file = "header.inc" -->
    <!-- member.asp -->
    <h3>회원 정보 목록</h3>
 <table width="400" bordercolor="black" border="1" style="border-collapse:collapse;" cellpadding="5">
  <tr style="background-color:#CCFF99;">
        <td width="50">번호</td>
        <td width="150">이름</td>
        <td width="200">이메일</td>
  </tr>

 

  <%  '----------------------------------------------------------------------------------------------------┐
        '커서가 파일의 끝을 만날 때까지 반복시킬꺼야.
        'Do Until 은 부정을 만날 때까지 반복하는 녀석이야.
        '시작됨과 동시에 False를 반환해. 그럼 실제 존재하는 레코드를 가리킨다는 바람직한 현상이겠지.
        Do Until rs.EOF
  %>


  <tr> <!-- rs(0)이라고 하면 직관성이 떨어져.. 그러나 속도는 rs(0)이 좀 더 빨라. -->
        <td><%= rs("seq") %></td>
        <!-- 이름을 클릭하면 view.asp로 가게 만들었어. 사람마다 다른 페이지를 보여줘야해.. 식별자필요 -->
        <!-- 우리가 seq를 만들었어. <퍼센트 rs("seq") 퍼센트> 부분이지.. 고유식별자를 위한 코딩이야 -->
        <td><a href="view.asp?seq=<%= rs("seq") %>"><%= rs("name") %></a></td>
        <td><%= rs("email") %></td>
  </tr>


  <%
        rs.MoveNext '--> 레코드를 한번 읽었다면 커서를 한줄 밑으로 내려주란 말이야~.... 기본적인 탐색
        Loop
  %> <!-------------------------------------------------------------------------------------------------------->

 </table>

 

 <br />
 <!-- 따로 자바스크립트함수를 만들지 않고 바로 '회원가입'페이지로 이동시키자. 페이지주소는 홑따음표야 -->
     <input type="button" value=" 가입하기 " onclick="location.href='add.asp'" />

 <!-- #include file = "footer.inc" -->
 </body>
</html>



