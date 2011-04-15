<%
function LocaleIntToStr(byval intIn)
	select case intIn
	case 1
		LocaleIntToStr = "en"
	case 2
		LocaleIntToStr = "zh"
	case 3
		LocaleIntToStr = "zh-hk"
	end select
end function

function LocaleStrToInt(byval strIn)
	Dim strInput
	strInput=lcase(strIn)
	select case strInput
	case "en"
		LocaleStrToInt = 1
	case "zh"
		LocaleStrToInt = 2
	case "zh-hk"
		LocaleStrToInt = 3
	end select
end function

function ProtocolIntToStr(byval intIn)
	select case intIn
	case 1
		ProtocolIntToStr = "http"
	case 2
		ProtocolIntToStr = "wap"
	end select
end function

function ProtocolStrToInt(byval strIn)
	Dim strInput
	strInput=lcase(strIn)
	select case strInput
	case "http"
		ProtocolStrToInt = 1
	case "wap"
		ProtocolStrToInt = 2
	end select
end function
%>
