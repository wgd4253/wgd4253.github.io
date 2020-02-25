function Send()
{
     var form = document.forms[0];
     //if (form.txtName.value == "")....... DB문제때문에 권장 X
     //{
     // alert("이름을 입력하세요");
     // return;
     //}

 

     // 모든 텍스트박스가 빈값인지 아닌지를 확인해보자.
     var txts = document.getElementsByTagName("input");
     for (i=0; i<txts.length; i++)
     {
          // button도 input태그이므로 주의하자.
          // alert(txts[i].name)
          if (txts[i].name != "txtEmail")
          {
               if (txts[i].value == "")
               {
                     alert("값을 입력하세요");
                     txts[i].focus();
                     return;
               }
          }
     }

     // for문을 빠져나온다음 전송
     form.submit();
}



