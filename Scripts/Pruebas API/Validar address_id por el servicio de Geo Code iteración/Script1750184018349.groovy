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
import java.text.Normalizer
import java.net.URLEncoder

// Leer el Excel con direcciones no georeferenciadas | address_id == 0
def data = TestDataFactory.findTestData('pruebasAPI/DireccionesNoGeoreferenciadasAddressId') // Archivo de datos que debe crearse

// Validar que el archivo de datos existe y tiene datos
assert data != null : "Archivo de datos 'DireccionesNoGeoreferenciadasAddressId' no encontrado"
assert data.getRowNumbers() > 0 : "Archivo de datos 'DireccionesNoGeoreferenciadasAddressId' no contiene registros"

// ======================================
// Lista para almacenar resultados para validación posterior
def resultadosValidacion = []

// Crear el request para el endpoint de creación de tracking
RequestObject addressUbigeoRequest = new RequestObject()
addressUbigeoRequest.setRestRequestMethod("GET")

// Agregar encabezados
List<TestObjectProperty> headers = new ArrayList<>()
headers.add(new TestObjectProperty("Content-Type", ConditionType.EQUALS, "application/json"))
headers.add(new TestObjectProperty("x-api-key", ConditionType.EQUALS, "\$2y\$10\$MZZvFZgsT6tiRT1tfhPWvuII6y.edh8Z1XiVKEvLrfUGTPAAAE416"))

// Iterar por cada fila del Excel
for (int i = 1; i <= data.getRowNumbers(); i++) {
	def nro = data.getValue('NRO', i)
	
	// Normalizar acentos: áéíóúñ -> aeioun
	direccion = Normalizer.normalize(direccion, Normalizer.Form.NFD)
	direccion = direccion.replaceAll("\\p{InCombiningDiacriticalMarks}+", "")
	
	// Validaciones para que se las direcciones se puedan pasar por el endpoint
	def direccion = data.getValue('DIRECCIONES', i)
	direccion = direccion.replaceAll("[\\u00A0\\u2007\\u202F\\)\\(\\\"\\,\\:\\.\\;\\-º]", " ")
	direccion = direccion.replaceAll("\\u0099", "  ")
	direccion = direccion.replaceAll("[\\p{Cntrl}&&[^\r\n\t]]", "  ")
	direccion = direccion.replaceAll("[^\\x00-\\x7F]", "")
	String direccionSinEspacios = URLEncoder.encode(direccion, "UTF-8").replace("+", "%20")
	
	// Agregar 0 faltantes en Ubigeos Enviados
	def ubigeo = data.getValue('CODUBIGEO', i)
	String ubigeoCompletado = ubigeo.toString().padLeft(6, '0')
	
	def tracking = data.getValue('TRACKING', i)
	def servicioCodigo = data.getValue('SERVICIOCODIGO', i)
	def nombreCliente = data.getValue('NOMBRECLIENTE', i)

	KeywordUtil.logInfo("Procesando registro #${nro}: ${direccion}")
	
	// Ruta en desarrollo
	String addressSuggestedUrl = "https://geo-api-dev.olvaexpress.pe/api/v2/geo/code/?address=${direccionSinEspacios}&ubigeo=${ubigeoCompletado}"
	// Ruta en producción
	//String addressSuggestedUrl = "https://geo-api.olvaexpress.pe/api/v2/geo/code/?address=${direccionSinEspacios}&ubigeo=${ubigeoCompletado}"
	addressUbigeoRequest.setRestUrl(addressSuggestedUrl)
	addressUbigeoRequest.setHttpHeaderProperties(headers)
	
	// Enviar request
	def getAddressInfoResponse = WS.sendRequest(addressUbigeoRequest)
	def statusCode = getAddressInfoResponse.getStatusCode()
	def responseBody = getAddressInfoResponse.getResponseBodyContent()

	// Validar status 200 y contenido no vacío
	if (statusCode == 200 && responseBody && !responseBody.trim().isEmpty()) {
		KeywordUtil.logInfo("✅ Status code 200 para registro #${nro}")
		
		def respuestaJson = new JsonSlurper().parseText(responseBody)

		// Validar campos requeridos
		if (!respuestaJson.address || !respuestaJson.polygon) {
			KeywordUtil.markWarning("⚠️ Respuesta incompleta para registro #${nro}. No se encontró address o polygon.")

			resultadosValidacion.add([
				nro: nro,
				direccionEnviada: direccion,
				direccionObtenida: "SIN RESULTADOS",
				ubigeo: ubigeoCompletado,
				poligonoObtenido: "SIN RESULTADOS",
				tracking: tracking,
				servicioCodigo: servicioCodigo,
				nombreCliente: nombreCliente
			])
			continue
		}

		def direccionObtenida = respuestaJson.address
		def poligonoObtenido = respuestaJson.polygon
		
		resultadosValidacion.add([
			nro: nro,
			direccionEnviada: direccion,
			direccionObtenida: direccionObtenida,
			ubigeo: ubigeoCompletado,
			poligonoObtenido: poligonoObtenido,
			tracking: tracking,
			servicioCodigo: servicioCodigo,
			nombreCliente: nombreCliente
		])
	} else {
		KeywordUtil.markWarning("⚠️ No se obtuvo contenido para registro #${nro} (Status code: ${statusCode})")
		resultadosValidacion.add([
			nro: nro,
			direccionEnviada: direccion,
			direccionObtenida: "SIN RESULTADOS",
			ubigeo: ubigeoCompletado,
			poligonoObtenido: "SIN RESULTADOS",
			tracking: tracking,
			servicioCodigo: servicioCodigo,
			nombreCliente: nombreCliente
		])
	}
}

// Resumen final
int totalRegistros = resultadosValidacion.size()
int noEncontrado = resultadosValidacion.count { it.direccionObtenida == "SIN RESULTADOS" }
int exitosos = totalRegistros - noEncontrado

// Crear el archivo Excel
Workbook workbook = new XSSFWorkbook()
Sheet sheet = workbook.createSheet("Resultados Validación")

// Crear fila de encabezados
def excelHeaders = ["NRO", "DIRECCIÓN ENVIADA", "DIRECCIÓN OBTENIDA", "UBIGEO", "POLÍGONO OBTENIDO", "TRACKING", "SERVICIO CÓDIGO", "NOMBRE CLIENTE"]
Row headerRow = sheet.createRow(0)
excelHeaders.eachWithIndex { header, idx ->
	Cell cell = headerRow.createCell(idx)
	cell.setCellValue(header)
}

// Llenar los datos
resultadosValidacion.eachWithIndex { resultado, index ->
	Row row = sheet.createRow(index + 1)
	row.createCell(0).setCellValue(resultado.nro.toString())
	row.createCell(1).setCellValue(resultado.direccionEnviada)
	row.createCell(2).setCellValue(resultado.direccionObtenida)
	row.createCell(3).setCellValue(resultado.ubigeo.toString())
	row.createCell(4).setCellValue(resultado.poligonoObtenido)
	row.createCell(5).setCellValue(resultado.tracking)
	row.createCell(6).setCellValue(resultado.servicioCodigo)
	row.createCell(7).setCellValue(resultado.nombreCliente)
}

// Ajustar tamaño de columnas
excelHeaders.eachWithIndex { _, idx ->
	sheet.autoSizeColumn(idx)
}

// Obtener la fecha actual en formato ddMMyyyy_HHmmss
def fechaHoraActual = LocalDateTime.now().format(DateTimeFormatter.ofPattern("ddMMyyyy_HHmmss"))

// Crear el nombre del archivo con la fecha
def nombreArchivo = "Resultados_Validacion_Georeferenciacion_DireccionUbigeo_${fechaHoraActual}.xlsx"

// Construir la ruta completa al archivo
def outputPath = Paths.get(RunConfiguration.getProjectDir(), nombreArchivo).toString()
FileOutputStream fileOut = new FileOutputStream(outputPath)
workbook.write(fileOut)
fileOut.close()
workbook.close()

KeywordUtil.logInfo("📁 Archivo Excel generado: ${outputPath}")
KeywordUtil.logInfo("📊 Resumen de validación:")
KeywordUtil.logInfo("- Total registros procesados: ${totalRegistros}")
KeywordUtil.logInfo("- Direcciones encotradas: ${exitosos}")
KeywordUtil.logInfo("- Direcciones no encotradas: ${noEncontrado}")

// Verificación final (pasa si todos los polígonos coinciden)
assert noEncontrado == totalRegistros : "Hay ${exitosos.toString()} de ${totalRegistros.toString()} direcciones que fueron encontradas"
