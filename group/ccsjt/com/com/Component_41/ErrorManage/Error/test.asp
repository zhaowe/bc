<%
dim testDate
dim strdate
testDate = Time()
strdate = cstr(testdate)
strdate = Formatdatetime(time,3)
Response.Write strdate
%>