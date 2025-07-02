import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.testobject.*
import com.kms.katalon.core.util.KeywordUtil
import com.kms.katalon.core.testdata.TestData
import com.kms.katalon.core.testdata.TestDataFactory
import com.kms.katalon.core.testobject.impl.HttpTextBodyContent
import groovy.json.*


// ======================================
// Configuración inicial
// ======================================
int baseCodigoOptitrack = 48000
TestData datosDirecciones = TestDataFactory.findTestData("pruebasAPI/DireccionesTrackingUbigeosTodosLosCasos")


// ======================================
// Paso 1: Login
// ======================================
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
loginRequest.setBodyContent(new HttpTextBodyContent(loginBody, "UTF-8", "application/json"))
def loginResponse = WS.sendRequest(loginRequest)
WS.verifyResponseStatusCode(loginResponse, 200)
def token = new JsonSlurper().parseText(loginResponse.getResponseBodyContent()).token
KeywordUtil.logInfo("🔐 Token JWT: ${token}")

// ======================================
// Paso 2: Crear lista de envíos desde Excel
// ======================================
List<Map> listaEnvios = []

for (int i = 1; i <= datosDirecciones.getRowNumbers(); i++) {
    String direccion = datosDirecciones.getValue("DIRECCIONES", i)
    String ubigeoStr = datosDirecciones.getValue("IDUBIGEO", i)
    int idUbigeo = ubigeoStr?.isNumber() ? ubigeoStr.toInteger() : 0
    int codigoOptitrack = baseCodigoOptitrack + i - 1

    def envio = [
        "parentOpti": null,
        "codigoOptitrack": codigoOptitrack,
        "idRecojo": null,
        "idSede": 174,
        "idOficina": 342,
        "direccionEntrega": direccion,
        "decJurada": 0,
        "decJuradaMonto": 5555,
        "cargoAjuntoCant": 0,
        "idPersJurArea": "314061",
        "consignado": "Pruebas Alem origen Cusco",
        "consignadoTelf": "949078370",
        "consignadoDni": "",
        "codExterno": "950534987",
        "idUbigeo": idUbigeo,
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
        "envioArticulo": [
            "pesoKgs": 10,
            "ancho": 42,
            "largo": 58,
            "alto": 21,
            "idContenedorArticulo": 19,
            "idArticulo": 0
        ],
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
    ]
    listaEnvios.add(envio)
}

String jsonBody = JsonOutput.toJson(listaEnvios)

// ======================================
// Paso 3: Enviar los envíos
// ======================================
RequestObject envioRequest = new RequestObject()
envioRequest.setRestUrl("https://dev-olva-corp.olvacourier.com/envioRest/webresources/envio/crear")
envioRequest.setRestRequestMethod("POST")
envioRequest.setHttpHeaderProperties([
    new TestObjectProperty("Content-Type", ConditionType.EQUALS, "application/json"),
    new TestObjectProperty("Authorization", ConditionType.EQUALS, "Bearer " + token)
])
envioRequest.setBodyContent(new HttpTextBodyContent(jsonBody, "UTF-8", "application/json"))

def envioResponse = WS.sendRequest(envioRequest)
WS.verifyResponseStatusCode(envioResponse, 201)

// ======================================
// Paso 4: Mostrar emision y tracking
// ======================================
def respuestaJson = new JsonSlurper().parseText(envioResponse.getResponseBodyContent())

if (respuestaJson.envios instanceof List) {
    KeywordUtil.logInfo("📋 Resultado final:")
    respuestaJson.envios.eachWithIndex { envio, idx ->
        KeywordUtil.logInfo("🔖 Envío #${idx + 1}: emision = ${envio.emision}, remito = ${envio.remito}")
    }
} else {
    KeywordUtil.markWarning("⚠️ Estructura de respuesta inesperada. No se encontró 'envios'")
}
