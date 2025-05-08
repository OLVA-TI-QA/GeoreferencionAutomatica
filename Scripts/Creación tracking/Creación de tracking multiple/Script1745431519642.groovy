import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.testobject.RequestObject
import com.kms.katalon.core.testobject.TestObjectProperty
import com.kms.katalon.core.testobject.ConditionType
import com.kms.katalon.core.util.KeywordUtil
import groovy.json.JsonOutput
import groovy.json.JsonSlurper

// ======================================
// 👤 Login para obtener token
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
loginRequest.setBodyContent(new com.kms.katalon.core.testobject.impl.HttpTextBodyContent(loginBody, "UTF-8", "application/json"))
def loginResponse = WS.sendRequest(loginRequest)
WS.verifyResponseStatusCode(loginResponse, 200)
def loginJson = new JsonSlurper().parseText(loginResponse.getResponseBodyContent())
def token = loginJson.token
KeywordUtil.logInfo("🔐 Token JWT: ${token}")

// ======================================
// 🧮 Parámetros configurables
// ======================================
int baseCodigoOptitrack = 1660      // Cambia este valor según necesidad
int cantidad = 2                    // Cambia este valor para N envíos

// ======================================
// 🧱 Crear lista de envíos
// ======================================
List<Map> envios = []
for (int i = 0; i < cantidad; i++) {
    def envio = [
        "parentOpti" : null,
        "codigoOptitrack" : baseCodigoOptitrack + i,
        "idRecojo" : null,
        "idSede" : 43,
        "idOficina" : 39,
        "direccionEntrega" : "LA PAZ MZ 246 LT 13, YARINACOCHA, CORONEL PORTILLO, UCAYALI",
        "decJurada" : 0,
        "decJuradaMonto" : 5555,
        "cargoAjuntoCant" : 0,
        "idPersJurArea" : "314061",
        "consignado" : "TIENDA OECHSLE - CUSCO   ANDREA SILVA  DIGNA CONDORI",
        "consignadoTelf" : "949078370",
        "consignadoDni" : "",
        "codExterno" : "950534987",
        "idUbigeo" : 1648,
        "codPostal" : "",
        "createUser" : 111,
        "idTipoVia" : 0,
        "idTipoZona" : 0,
        "nombreVia" : "0",
        "nombreZona" : "0",
        "numero" : "0",
        "manzana" : "0",
        "lote" : "0",
        "latitud" : -77.0696622743156,
        "longitud" : -12.042115838754475,
        "poligono" : null,
        "idServicio" : 35,
        "codOperador" : null,
        "tipoGestion" : null,
        "envioArticulo" : [
            "pesoKgs": 10,
            "ancho": 42,
            "largo": 58,
            "alto": 21,
            "idContenedorArticulo": 19,
            "idArticulo": 0
        ],
        "flgOficina" : false,
        "idOfiDest" : null,
        "montoBase" : 16.25,
        "montoExceso" : 63,
        "montoSeguro" : 454,
        "montoIgv" : 95.985,
        "precioVenta" : 629.235,
        "montoEmbalaje" : 0,
        "montoOtrosCostos" : 0,
        "montoTransporte" : 0,
        "entregaEnOficina" : "0",
        "numDocSeller" : "",
        "nombreSeller" : "",
        "codigoAlmacen" : "",
        "codUbigeo" : "",
        "direccionSeller" : "",
        "referenciaSeller" : "",
        "contacto" : "",
        "telefono" : "",
        "observacion" : "",
        "nroPiezas" : 10
    ]
    envios.add(envio)
}

String bodyEnvio = JsonOutput.toJson(envios)
KeywordUtil.logInfo("📦 JSON enviado:\n" + bodyEnvio)

// ======================================
// 🚚 Enviar a /envio/crear
// ======================================
RequestObject envioRequest = new RequestObject()
envioRequest.setRestUrl("https://dev-olva-corp.olvacourier.com/envioRest/webresources/envio/crear")
envioRequest.setRestRequestMethod("POST")
envioRequest.setHttpHeaderProperties([
    new TestObjectProperty("Content-Type", ConditionType.EQUALS, "application/json"),
    new TestObjectProperty("Authorization", ConditionType.EQUALS, "Bearer " + token)
])
envioRequest.setBodyContent(new com.kms.katalon.core.testobject.impl.HttpTextBodyContent(bodyEnvio, "UTF-8", "application/json"))

def envioResponse = WS.sendRequest(envioRequest)
WS.verifyResponseStatusCode(envioResponse, 201)

String responseText = envioResponse.getResponseBodyContent()
KeywordUtil.logInfo("📨 Respuesta completa del servidor:\n" + responseText)

// ======================================
// 📋 Extraer remito y emisión
// ======================================
def respuestaJson = new JsonSlurper().parseText(responseText)

if (respuestaJson.envios instanceof List) {
    KeywordUtil.logInfo("🧾 Resultado de envíos:")
    respuestaJson.envios.eachWithIndex { envio, idx ->
        def emision = envio.emision
        def remito = envio.remito
        KeywordUtil.logInfo("🔖 Envío #${idx + 1}: emision = ${emision}, tracking = ${remito}")
    }
} else {
    KeywordUtil.markWarning("⚠️ No se encontró la lista 'envios' en la respuesta.")
}

KeywordUtil.markPassed("✅ Se procesaron ${cantidad} envío(s) correctamente.")
