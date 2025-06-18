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

// Leer el Excel con direcciones de oficinas
def data = TestDataFactory.findTestData('pruebasAPI/DataDeOficinasLima') // Archivo de datos que debe crearse

// Validar que el archivo de datos existe y tiene datos
assert data != null : "Archivo de datos 'DataDeOficinasLima' no encontrado"
assert data.getRowNumbers() > 0 : "Archivo de datos 'DataDeOficinasLima' no contiene registros"

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
	def direccion = data.getValue('DIRECCIONES', i)
	
	// Normalizar acentos: áéíóúñ -> aeioun
	direccion = Normalizer.normalize(direccion, Normalizer.Form.NFD)
	direccion = direccion.replaceAll("\\p{InCombiningDiacriticalMarks}+", "")
	
	// Validaciones para que se las direcciones se puedan pasar por el endpoint
	direccion = direccion.replaceAll("[\\u00A0\\u2007\\u202F\\)\\(\\\"\\,\\:\\.\\;\\-º]", " ")
	direccion = direccion.replaceAll("\\u0099", "  ")
	direccion = direccion.replaceAll("[\\p{Cntrl}&&[^\r\n\t]]", "  ")
	direccion = direccion.replaceAll("[^\\x00-\\x7F]", "")
	String direccionSinEspacios = URLEncoder.encode(direccion, "UTF-8").replace("+", "%20")
	
	def ubigeo = data.getValue('CODUBIGEO', i)
	String ubigeoCompletado = ubigeo.toString().padLeft(6, '0')
	
	def nombreOficina = data.getValue('NOMBREOFICINA', i)
	def idAddress = data.getValue('IDADDRESS', i)

	KeywordUtil.logInfo("Procesando registro #${nro}: ${direccion}")
	
	// Ruta en desarrollo
	String fullUrl = "https://geo-api-dev.olvaexpress.pe/api/v2/geo/code/?address=${direccionSinEspacios}&ubigeo=${ubigeoCompletado}"
	// Ruta en producción
	//String fullUrl = "https://geo-api.olvaexpress.pe/api/v2/geo/code/?address=${direccionSinEspacios}&ubigeo=${ubigeoCompletado}"
	
	addressUbigeoRequest.setRestUrl(fullUrl)
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
		if (!respuestaJson.address) {
			KeywordUtil.markWarning("⚠️ Respuesta incompleta para registro #${nro}. No se encontró address o polygon.")

			resultadosValidacion.add([
				nro: nro,
				direccionEnviada: direccion,
				direccionObtenida: "SIN RESULTADOS",
				codUbigeoEnviado: ubigeo,
				codUbigeoObtenido: "SIN RESULTADOS",
				oficina: false,
				idAddress: idAddress,
				nombreOficina: nombreOficina
			])
			continue
		}

		def codUbigeoObtenido = respuestaJson.ubigeo
		def direccionObtenida = respuestaJson.address
		def isOffice = respuestaJson.office
		
		resultadosValidacion.add([
			nro: nro,
			direccionEnviada: direccion,
			direccionObtenida: direccionObtenida,
			codUbigeoEnviado: ubigeo,
			codUbigeoObtenido: codUbigeoObtenido,
			oficina: isOffice,
			idAddress: idAddress,
			nombreOficina: nombreOficina
		])
	} else {
		KeywordUtil.markWarning("⚠️ No se obtuvo contenido para registro #${nro} (Status code: ${statusCode})")
		resultadosValidacion.add([
			nro: nro,
			direccionEnviada: direccion,
			direccionObtenida: "SIN RESULTADOS",
			codUbigeoEnviado: ubigeo,
			codUbigeoObtenido: "SIN RESULTADOS",
			oficina: false,
			idAddress: idAddress,
			nombreOficina: nombreOficina
		])
	}
}

// Resumen final
int totalRegistros = resultadosValidacion.size()
int exitosos = resultadosValidacion.count { it.oficina == true }
int fallidos = totalRegistros - exitosos

// Crear el archivo Excel
Workbook workbook = new XSSFWorkbook()
Sheet sheet = workbook.createSheet("Resultados Validación")

// Crear fila de encabezados
def excelHeaders = ["NRO", "DIRECCIÓN ENVIADA", "DIRECCIÓN OBTENIDA", "CODUBIGEO ENVIADO", "CODUBIGEO OBTENIDO", "OFICINA", "ID ADDRESS", "NOMBRE OFICINA"]
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
	row.createCell(3).setCellValue(resultado.codUbigeoEnviado.toString())
	row.createCell(4).setCellValue(resultado.codUbigeoObtenido)
	row.createCell(5).setCellValue(resultado.oficina.toString())
	row.createCell(6).setCellValue(resultado.idAddress.toString())
	row.createCell(7).setCellValue(resultado.nombreOficina)
}

// Ajustar tamaño de columnas
excelHeaders.eachWithIndex { _, idx ->
	sheet.autoSizeColumn(idx)
}

// Obtener la fecha actual en formato ddMMyyyy_HHmmss
def fechaHoraActual = LocalDateTime.now().format(DateTimeFormatter.ofPattern("ddMMyyyy_HHmmss"))

// Crear el nombre del archivo con la fecha
def nombreArchivo = "Resultados_Validacion_Oficina_DireccionUbigeo_${fechaHoraActual}.xlsx"

// Construir la ruta completa al archivo
def outputPath = Paths.get(RunConfiguration.getProjectDir(), nombreArchivo).toString()
FileOutputStream fileOut = new FileOutputStream(outputPath)
workbook.write(fileOut)
fileOut.close()
workbook.close()

KeywordUtil.logInfo("📁 Archivo Excel generado: ${outputPath}")
KeywordUtil.logInfo("📊 Resumen de validación:")
KeywordUtil.logInfo("- Total registros procesados: ${totalRegistros}")
KeywordUtil.logInfo("- Oficinas validadas correctamente: ${exitosos}")
KeywordUtil.logInfo("- Oficinas validadas incorrectas: ${fallidos}")

// Verificación final (pasa si todos los polígonos coinciden)
assert exitosos == totalRegistros : "Hay ${exitosos.toString()} de ${totalRegistros.toString()} oficinas que fueron validadas"
