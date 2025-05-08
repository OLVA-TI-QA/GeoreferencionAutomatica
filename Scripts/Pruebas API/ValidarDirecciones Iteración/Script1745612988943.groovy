import com.kms.katalon.core.testobject.RequestObject
import com.kms.katalon.core.testobject.TestObjectProperty
import com.kms.katalon.core.testobject.impl.HttpTextBodyContent
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.testobject.ConditionType
import groovy.json.JsonSlurper
import com.kms.katalon.core.testdata.ExcelData
import com.kms.katalon.core.testdata.TestDataFactory as TestDataFactory

// Leer el Excel
def data = TestDataFactory.findTestData('Data Files/coordenadas') // <- ruta en Katalon Test Data (no poner .xlsx)

// Construir el array de coordenadas
def coordinates = []

for (int i = 1; i <= data.getRowNumbers(); i++) {
    def lat = data.getValue('lat', i) as Double
    def lon = data.getValue('lon', i) as Double
    def radius = data.getValue('radius', i)

    def coord = [
        lat: lat,
        lon: lon
    ]
    if (radius) {
        coord['radius'] = radius as Integer
    }
    coordinates.add(coord)
}

// Convertir coordenadas a JSON para el body
def requestBodyJson = [
    coordinates: coordinates
]
def requestBody = groovy.json.JsonOutput.toJson(requestBodyJson)

// Crear objeto de petición
RequestObject request = new RequestObject()
request.setRestRequestMethod('POST')
request.setRestUrl('https://geo-api-dev.olvaexpress.pe/api/v2/geo/batch-reverse')

// Agregar encabezados
List<TestObjectProperty> headers = new ArrayList<>()
headers.add(new TestObjectProperty("Content-Type", ConditionType.EQUALS, "application/json"))
headers.add(new TestObjectProperty("x-api-key", ConditionType.EQUALS, "\$2y\$10\$MZZvFZgsT6tiRT1tfhPWvuII6y.edh8Z1XiVKEvLrfUGTPAAAE416"))
request.setHttpHeaderProperties(headers)

request.setBodyContent(new HttpTextBodyContent(requestBody, "UTF-8", "application/json"))

// Enviar petición y validar respuesta
def response = WS.sendRequest(request)
WS.verifyResponseStatusCode(response, 200)

// Parsear respuesta JSON
def json = new JsonSlurper().parseText(response.getResponseBodyContent())

// Validar número de resultados
assert json.results.size() == data.getRowNumbers() : "Cantidad de resultados no coincide con el Excel."

// Validar que cada dirección sea igual a la esperada
for (int i = 0; i < data.getRowNumbers(); i++) {
    def direccionEsperada = data.getValue('direccionEsperada', i + 1)
    def result = json.results[i]

    assert result.status == "success" : "Resultado ${i + 1} no tiene status 'success'."
    def address = result.data?.address
    assert address != null && !address.trim().isEmpty() : "Resultado ${i + 1} no contiene dirección válida"
    assert address == direccionEsperada : "Dirección ${i + 1} esperada '${direccionEsperada}' no coincide con '${address}'"
    println "✅ Dirección ${i + 1} correcta: ${address}"
}
