Dim scriptName, countryCode, vatNumber, apiUrl, http, requestBody, responseText

' Obtener el nombre del script autom�ticamente
scriptName = WScript.ScriptName

' Leer argumentos de la l�nea de comandos
If WScript.Arguments.Count < 2 Then
    WScript.Echo "Uso: cscript " &  scriptName & " <CodigoPais> <CIF>"
    WScript.Quit 1
End If

countryCode = WScript.Arguments(0)
vatNumber = WScript.Arguments(1)

' URL de la API VIES
apiUrl = "https://ec.europa.eu/taxation_customs/vies/rest-api/check-vat-number"

' Crear objeto HTTP
Set http = CreateObject("MSXML2.XMLHTTP")

' Crear cuerpo de la petici�n en formato JSON (sin comillas dobles en VBScript)
requestBody = "{""countryCode"":""" & countryCode & """, ""vatNumber"":""" & vatNumber & """}"

' Enviar solicitud POST
http.Open "POST", apiUrl, False
http.setRequestHeader "Content-Type", "application/json"
http.Send requestBody

' Leer respuesta de la API
responseText = http.responseText

' Normalizar la respuesta eliminando espacios y saltos de l�nea
responseText = Replace(responseText, vbCr, "") ' Quitar retornos de carro
responseText = Replace(responseText, vbLf, "") ' Quitar saltos de l�nea

' Buscar si la respuesta contiene `"valid":true`
If InStr(responseText, """valid"" : true") > 0 Then
    WScript.Echo "true"
Else
    WScript.Echo "false"
End If

