import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.testobject.TestObject
import com.kms.katalon.core.testobject.TestObjectProperty
import com.kms.katalon.core.testobject.RequestObject
import com.kms.katalon.core.testobject.ConditionType
import groovy.json.JsonSlurper
import com.kms.katalon.core.util.KeywordUtil

// =============================
// Paso 1: Login para obtener el token
// =============================

RequestObject loginRequest = new RequestObject()
loginRequest.setRestUrl("https://dev-olva-corp.olvacourier.com/envioRest/webresources/usuario/login")
loginRequest.setRestRequestMethod("POST")

loginRequest.setHttpHeaderProperties([
    new TestObjectProperty("Content-Type", ConditionType.EQUALS, "application/json")
])

String loginBody = """
{
    "usuario": "olvati",
    "clave": "J&_Mv9]H^2Vx"
}
"""

loginRequest.setBodyContent(new com.kms.katalon.core.testobject.impl.HttpTextBodyContent(loginBody, "UTF-8", "application/json"))

def loginResponse = WS.sendRequest(loginRequest)
WS.verifyResponseStatusCode(loginResponse, 200)

def jsonSlurper = new JsonSlurper()
def loginResult = jsonSlurper.parseText(loginResponse.getResponseBodyContent())
def token = loginResult.token

KeywordUtil.logInfo("🔐 Token obtenido: " + token)

// =============================
// Paso 2: Crear Envío usando el token
// =============================

RequestObject crearEnvioRequest = new RequestObject()
crearEnvioRequest.setRestUrl("https://dev-olva-corp.olvacourier.com/envioRest/webresources/envio/crear")
crearEnvioRequest.setRestRequestMethod("POST")

crearEnvioRequest.setHttpHeaderProperties([
    new TestObjectProperty("Content-Type", ConditionType.EQUALS, "application/json"),
    new TestObjectProperty("Authorization", ConditionType.EQUALS, "Bearer " + token)
])

String envioBody = """
[
  {
    "parentOpti": null,
    "codigoOptitrack": 1600,
    "idRecojo": null,
    "idSede": 43,
    "idOficina": 39,
    "direccionEntrega": "LA PAZ MZ 246 LT 13, YARINACOCHA, CORONEL PORTILLO, UCAYALI",
    "decJurada": 0,
    "decJuradaMonto": 5555,
    "cargoAjuntoCant": 0,
    "idPersJurArea": "314061",
    "consignado": "TIENDA OECHSLE - CUSCO   ANDREA SILVA  DIGNA CONDORI",
    "consignadoTelf": "949078370",
    "consignadoDni": "",
    "codExterno": "950534987",
    "idUbigeo": 1648,
    "codPostal": "",
    "createUser": 111,
    "idTipoVia": 0,
    "idTipoZona": 0,
    "nombreVia": "0",
    "nombreZona": "0",
    "numero": "0",
    "manzana": "0",
    "lote": "0",
    "latitud": -77.0696622743156,
    "longitud": -12.042115838754475,
    "poligono": null,
    "idServicio": 35,
    "codOperador": null,
    "tipoGestion": null,
    "envioArticulo": {
      "pesoKgs": 10,
      "ancho": 42,
      "largo": 58,
      "alto": 21,
      "idContenedorArticulo": 19,
      "idArticulo": 0
    },
    "flgOficina": false,
    "idOfiDest": null,
    "montoBase": 16.25,
    "montoExceso": 63,
    "montoSeguro": 454,
    "montoIgv": 95.985,
    "precioVenta": 629.235,
    "montoEmbalaje": 0,
    "montoOtrosCostos": 0,
    "montoTransporte": 0,
    "entregaEnOficina": "0",
    "numDocSeller": "",
    "nombreSeller": "",
    "codigoAlmacen": "",
    "codUbigeo": "",
    "direccionSeller": "",
    "referenciaSeller": "",
    "contacto": "",
    "telefono": "",
    "observacion": "",
    "nroPiezas": 10
  }
]
"""

crearEnvioRequest.setBodyContent(new com.kms.katalon.core.testobject.impl.HttpTextBodyContent(envioBody, "UTF-8", "application/json"))

def envioResponse = WS.sendRequest(crearEnvioRequest)
WS.verifyResponseStatusCode(envioResponse, 201)

KeywordUtil.markPassed("✅ El envío fue creado correctamente")
