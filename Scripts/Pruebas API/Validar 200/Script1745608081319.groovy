import com.kms.katalon.core.testobject.RequestObject
import com.kms.katalon.core.testobject.TestObjectProperty
import com.kms.katalon.core.testobject.impl.HttpTextBodyContent
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.testobject.ConditionType
import groovy.json.JsonSlurper

// Crear objeto de petición
RequestObject request = new RequestObject()
request.setRestRequestMethod('POST')
request.setRestUrl('https://geo-api-dev.olvaexpress.pe/api/v2/geo/batch-reverse')

// Agregar encabezados
List<TestObjectProperty> headers = new ArrayList<>()
headers.add(new TestObjectProperty("Content-Type", ConditionType.EQUALS, "application/json"))
headers.add(new TestObjectProperty("x-api-key", ConditionType.EQUALS, "\$2y\$10\$MZZvFZgsT6tiRT1tfhPWvuII6y.edh8Z1XiVKEvLrfUGTPAAAE416")) // Header correcto
request.setHttpHeaderProperties(headers)

// Cuerpo de la petición
def requestBody = '''
{
  "coordinates": [
    {
      "lon": -77.0282,
      "lat": -12.0464,
      "radius": 50
    },
    {
      "lon": -77.03884,
      "lat": -12.06251
    },
    {
      "lon": -77.02943,
      "lat": -12.09994,
      "radius": 30
    }
  ]
}
'''
request.setBodyContent(new HttpTextBodyContent(requestBody, "UTF-8", "application/json"))

// Enviar petición y validar respuesta
def response = WS.sendRequest(request)
WS.verifyResponseStatusCode(response, 200)

// Parsear respuesta JSON
def json = new JsonSlurper().parseText(response.getResponseBodyContent())

// Validar resultados
assert json.results.size() == 3 : "Se esperaban 3 resultados"
json.results.eachWithIndex { result, idx ->
    assert result.status == "success" : "Resultado ${idx + 1} no tiene status 'success'"
    assert result.data?.address : "Resultado ${idx + 1} no contiene dirección válida"
    println "✅ Dirección ${idx + 1}: ${result.data.address}"
}
