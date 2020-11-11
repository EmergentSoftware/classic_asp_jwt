<!--#include file="external/base64.asp"-->
<!--#include file="external/aspJSON.asp"-->
<%
' The URL- and filename-safe Base64 encoding described in RFC 4648 [RFC4648], Section 5,
' with the (non URL-safe) '=' padding characters omitted, as permitted by Section 3.2.
' (See Appendix C of [JWS] for notes on implementing base64url encoding without padding.)
' http://tools.ietf.org/html/rfc4648
' http://tools.ietf.org/html/draft-ietf-jose-json-web-signature-10
Function SafeBase64Encode(sIn)
  Dim sOut
  sOut = Base64Encode(sIn)
  sOut = Base64ToSafeBase64(sOut)

  SafeBase64Encode = sOut
End Function

' Strips unsafe characters from a Base64 encoded string
Function Base64ToSafeBase64(sIn)
  Dim sOut
  sOut = Replace(sIn,"+","-")
  sOut = Replace(sOut,"/","_")
  sOut = Replace(sOut,"\r","")
  sOut = Replace(sOut,"\n","")
  sOut = Replace(sOut,"=","")

  Base64ToSafeBase64 = sOut
End Function

' Converts an ASP dictionary to a JSON string
Function DictionaryToJSONString(dDictionary)
  Dim oJSONpayload
  Set oJSONpayload = New aspJSON

  
  Dim i, aKeys
  aKeys = dDictionary.keys
  
  For i = 0 to dDictionary.Count-1
    oJSONpayload.data (aKeys(i))= dDictionary(aKeys(i))
  Next

  DictionaryToJSONString = oJSONpayload.JSONoutput()
End Function
%>

<script language='Javascript' runat='server'>
function jsGetUTCTime() {
	var d = new Date();
	return (d.getUTCMonth() + 1) + "/" + d.getUTCDate() + "/" + d.getUTCFullYear()
		+ " " + d.getUTCHours() + ":" + d.getUTCMinutes() + ":" + d.getUTCSeconds();
}
</script>
<script language='VBScript' runat='server'>
Function getUTCTime()
    ' Use JavaScript to get the current GMT time stamp
    getUTCTime = jsGetUTCTime()
End Function
</script>

<%
' Returns the number of seconds since epoch
Function SecsSinceEpoch()
  SecsSinceEpoch = DateDiff("s", "01/01/1970 00:00:00", getUTCTime())
End Function

' Returns a random string to prevent replays
Function UniqueString()
    UniqueString = Left(CStr(CreateObject("Scriptlet.TypeLib").Guid), 38)
End Function
%>
