import com.kms.katalon.core.testobject.RequestObject
import com.kms.katalon.core.testobject.TestObjectProperty
import com.kms.katalon.core.testobject.impl.HttpTextBodyContent
import com.kms.katalon.core.webservice.keyword.WSBuiltInKeywords as WS
import com.kms.katalon.core.testobject.ConditionType
import com.kms.katalon.core.configuration.RunConfiguration
import groovy.json.JsonSlurper
import groovy.json.JsonOutput
import com.kms.katalon.core.testdata.TestDataFactory as TestDataFactory
import com.kms.katalon.core.util.KeywordUtil
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.usermodel.*

import java.io.FileOutputStream
import java.nio.file.Paths
import java.time.LocalDateTime
import java.time.format.DateTimeFormatter

// Leer el Excel con direcciones y polígonos
def data = TestDataFactory.findTestData('pruebasAPI/DireccionesConPoligonos') // Archivo de datos que debe crearse

// Validar que el archivo de datos existe y tiene datos
assert data != null : "Archivo de datos 'DireccionesConPoligonos' no encontrado"
assert data.getRowNumbers() > 0 : "Archivo de datos 'DireccionesConPoligonos' no contiene registros"

// ======================================
// Lista para almacenar resultados para validación posterior
def resultadosValidacion = []
boolean coincidePoligono = false
String addressIdSuggested = ''
String apiKey = "e06bc536-47da-46d7-a795-b12bb1aa1141"

// Crear el request para el endpoint de dirección sugerida
RequestObject addressSuggestedRequest = new RequestObject()
addressSuggestedRequest.setRestRequestMethod("GET")

// Crear el request para el endpoint de obtención de infomación de la dirección
RequestObject getAddressInfoRequest = new RequestObject()
getAddressInfoRequest.setRestRequestMethod("GET")

// Agregar encabezados
List<TestObjectProperty> headers = new ArrayList<>()
headers.add(new TestObjectProperty("Content-Type", ConditionType.EQUALS, "application/json"))
headers.add(new TestObjectProperty("x-api-key", ConditionType.EQUALS, "\$2y\$10\$MZZvFZgsT6tiRT1tfhPWvuII6y.edh8Z1XiVKEvLrfUGTPAAAE416"))

// Iterar por cada fila del Excel
for (int i = 1; i <= data.getRowNumbers(); i++) {
	def nro = data.getValue('NRO', i)
	def direccion = data.getValue('DIRECCIONES', i)
	String encodedSearch = direccion.replace(" ", "%20")
	def poligonoEsperado = data.getValue('POLIGONO', i)
	
	KeywordUtil.logInfo("Procesando registro #${nro}: ${direccion}")
	
	String addressSuggestedUrl = "https://apis.geodir.co/places/autocomplete/v1/json?search=${encodedSearch}&key=${apiKey}"
	addressSuggestedRequest.setRestUrl(addressSuggestedUrl)
	
	// Enviar request y validar respuesta
	def addressSuggestedResponse = WS.sendRequest(addressSuggestedRequest)
	
	// Verificar código de respuesta
	if (WS.verifyResponseStatusCode(addressSuggestedResponse, 200)) {
		def respuestaJson = new JsonSlurper().parseText(addressSuggestedResponse.getResponseBodyContent())
		
		// Verificar si no se encontraron resultados
		if (respuestaJson.status == 'ZERO_RESULTS' || !respuestaJson.predictions || respuestaJson.predictions.isEmpty()) {
			KeywordUtil.markWarning("⚠️ No se encontraron sugerencias para la dirección #${nro}: '${direccion}'")
			
			resultadosValidacion.add([
				nro: nro,
				direccion: direccion,
				poligonoEsperado: poligonoEsperado,
				poligonoObtenido: "SIN RESULTADOS",
				coincide: false
			])
			
			continue // Saltar a la siguiente iteración
		}
		
		// Extraer place_id
		addressIdSuggested = respuestaJson.predictions[0].place_id
	} else {
		KeywordUtil.markError("❌ Error en dirección #${nro}: Status code no es 200")
		
		resultadosValidacion.add([
			nro: nro,
			direccion: direccion,
			poligonoEsperado: poligonoEsperado,
			poligonoObtenido: "ERROR API AUTOCOMPLETE",
			coincide: false
		])
		
		continue
	}
	
	// Crear el request para el segundo endpoint
	// Ruta en producción
	String getAddressInfoUrl = "https://geo-api.olvaexpress.pe/api/v2/geo/code/?addressId=${addressIdSuggested}"
	// Ruta en desarrollo
	// String getAddressInfoUrl = "https://geo-api-dev.olvaexpress.pe/api/v2/geo/code/?addressId=${addressIdSuggested}"
	
	getAddressInfoRequest.setRestUrl(getAddressInfoUrl)
	getAddressInfoRequest.setHttpHeaderProperties(headers)
	
	// Enviar request y validar respuesta
	def getAddressInfoResponse = WS.sendRequest(getAddressInfoRequest)
	
	if (WS.verifyResponseStatusCode(getAddressInfoResponse, 200)) {
		KeywordUtil.logInfo("✅ Status code 200 para registro #${nro}")
		
		def respuestaJson = new JsonSlurper().parseText(getAddressInfoResponse.getResponseBodyContent())
		
		def direccionObtenida = respuestaJson.address
		def poligonoObtenido = respuestaJson.polygon
		
		coincidePoligono = (poligonoObtenido == poligonoEsperado)
		
		resultadosValidacion.add([
			nro: nro,
			direccion: direccionObtenida,
			poligonoEsperado: poligonoEsperado,
			poligonoObtenido: poligonoObtenido,
			coincide: coincidePoligono
		])
		
		if (coincidePoligono) {
			KeywordUtil.logInfo("✅ Polígono coincide para registro #${nro}")
		} else {
			KeywordUtil.markWarning("⚠️ Polígono NO coincide para registro #${nro}. Esperado: ${poligonoEsperado}, Obtenido: ${poligonoObtenido}")
		}
	} else {
		KeywordUtil.markError("❌ Error en registro #${nro}: Status code no es 200")
		resultadosValidacion.add([
			nro: nro,
			direccion: direccion,
			poligonoEsperado: poligonoEsperado,
			poligonoObtenido: "ERROR API GEO CODE",
			coincide: false
		])
	}
}

// Resumen final
int totalRegistros = resultadosValidacion.size()
int exitosos = resultadosValidacion.count { it.coincide }
int fallidos = totalRegistros - exitosos

// Crear el archivo Excel
Workbook workbook = new XSSFWorkbook()
Sheet sheet = workbook.createSheet("Resultados Validación")

// Crear fila de encabezados
def excelHeaders = ["NRO", "DIRECCIÓN", "POLÍGONO ESPERADO", "POLÍGONO OBTENIDO", "COINCIDE"]
Row headerRow = sheet.createRow(0)
excelHeaders.eachWithIndex { header, idx ->
	Cell cell = headerRow.createCell(idx)
	cell.setCellValue(header)
}

// Llenar los datos
resultadosValidacion.eachWithIndex { resultado, index ->
	Row row = sheet.createRow(index + 1)
	row.createCell(0).setCellValue(resultado.nro.toString())
	row.createCell(1).setCellValue(resultado.direccion)
	row.createCell(2).setCellValue(resultado.poligonoEsperado)
	row.createCell(3).setCellValue(resultado.poligonoObtenido)
	row.createCell(4).setCellValue(resultado.coincide.toString())
}

// Ajustar tamaño de columnas
excelHeaders.eachWithIndex { _, idx ->
	sheet.autoSizeColumn(idx)
}

// Obtener la fecha actual en formato ddMMyyyy_HHmmss
def fechaHoraActual = LocalDateTime.now().format(DateTimeFormatter.ofPattern("ddMMyyyy_HHmmss"))
// Crear el nombre del archivo con la fecha
def nombreArchivo = "Resultados_Validacion_Poligonos_DireccionSugerida_${fechaHoraActual}.xlsx"
// Construir la ruta completa al archivo
def outputPath = Paths.get(RunConfiguration.getProjectDir(), nombreArchivo).toString()

FileOutputStream fileOut = new FileOutputStream(outputPath)
workbook.write(fileOut)
fileOut.close()
workbook.close()

KeywordUtil.logInfo("📁 Archivo Excel generado: ${outputPath}")
KeywordUtil.logInfo("📊 Resumen de validación:")
KeywordUtil.logInfo("- Total registros procesados: ${totalRegistros}")
KeywordUtil.logInfo("- Polígonos coincidentes: ${exitosos}")
KeywordUtil.logInfo("- Polígonos no coincidentes: ${fallidos}")

// Verificación final (pasa si todos los polígonos coinciden)
assert exitosos == totalRegistros : "Hay ${fallidos} polígonos que no coinciden con los esperados"
