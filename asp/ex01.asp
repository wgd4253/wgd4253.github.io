<%
       'ex01.asp
       'ODBC에 연결
       conStr = "Provider=sqloledb;Data Source=db.assatech.com;Initial Catalog=dbassatech;User ID=assatech;Password=assa3095@@;"

 

       'Connection 객체 생성
       Set conn = Server.CreateObject("ADODB.Connection")

 

       'DB연결 확인
        'Response.Write conn.State '--> 화면에 "0"이 출력.. 연결이 되지 않으면 0이다.
       conn.Open conStr

 

       'DB연결 확인

        Response.Write conn.State  '--> 화면에 "1"이 출력.. 연결이 되면 1이다.


       '반환값이 없는 쿼리는 connection으로 바로 사용가능해.----------------------------------------┐
       '5개의 필드값을 가진 테이블 Member을 만들꺼야.
             query = "Create Table Member ("
             query = query & "seq int identity(1,1) not null, "
             query = query & "name varchar(15) not null, "
             query = query & "tel varchar(20) not null, "
             query = query & "email varchar(50) null, "
             query = query & "age int not null)"

 

       '--> 쿼리는 무조건 그냥 실행하지 말고 내가 만든 구문을 한번 확인해보자고..
       '--> 에러가 없다면 밑의 구문에 주석을 붙여버리자.
       '--> 간혹 봐도 에러를 모르겠다면 '쿼리분석기'에 붙여넣고 실행시켜보자
        'Response.Write query 
        'Response.End

 

        '쿼리실행 -> 테이블 생성(엔터프라이즈관리자에서 확인가능)---------------------------------┘
        conn.Execute query


        'Close
        conn.Close
        Set conn = Nothing '--> 사용객체 날려버리기
%>

 

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
 <head>
 </head>

 <body>
 </body>
</html>



