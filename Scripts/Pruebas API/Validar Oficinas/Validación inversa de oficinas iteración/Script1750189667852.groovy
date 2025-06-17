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

// Leer el Excel con direcciones de oficinas
def data = TestDataFactory.findTestData('pruebasAPI/DataDeOficinasLima') // Archivo de datos que debe crearse

// Validar que el archivo de datos existe y tiene datos
assert data != null : "Archivo de datos 'DataDeOficinasLima' no encontrado"
assert data.getRowNumbers() > 0 : "Archivo de datos 'DataDeOficinasLima' no contiene registros"

// ======================================
// Lista para almacenar resultados para validación posterior
def resultadosValidacion = []

// Crear el request para el endpoint de creación de tracking
RequestObject officeRequest = new RequestObject()
officeRequest.setRestRequestMethod("GET")

// Agregar encabezados
List<TestObjectProperty> headers = new ArrayList<>()
headers.add(new TestObjectProperty("Content-Type", ConditionType.EQUALS, "application/json"))
headers.add(new TestObjectProperty("x-api-key", ConditionType.EQUALS, "\$2y\$10\$MZZvFZgsT6tiRT1tfhPWvuII6y.edh8Z1XiVKEvLrfUGTPAAAE416"))

// Iterar por cada fila del Excel
for (int i = 1; i <= data.getRowNumbers(); i++) {
	def nro = data.getValue('NRO', i)
	def lon = data.getValue('LONGITUD', i)
	def lat = data.getValue('LATITUD', i)
	def direccion = data.getValue('DIRECCIONES', i)
	def codUbigeoEnviado = data.getValue('CODUBIGEO', i)
	def nombreOficina = data.getValue('NOMBREOFICINA', i)
	def idAddress = data.getValue('IDADDRESS', i)
	
	KeywordUtil.logInfo("Procesando registro #${nro}: ${direccion}")
	
	// Ruta en desarrollo
	String fullUrl = "https://geo-api-dev.olvaexpress.pe/api/v2/geo/reverse/?lon=${lon}&lat=${lat}"
	// Ruta en producción
	//String fullUrl = "https://geo-api.olvaexpress.pe/api/v2/geo/reverse/?lon=${lon}&lat=${lat}"
	
	officeRequest.setRestUrl(fullUrl)
	officeRequest.setHttpHeaderProperties(headers)
	
	// Enviar request y validar respuesta
	def officeResponse = WS.sendRequest(officeRequest)
	def statusCode = officeResponse.getStatusCode()
	def responseBody = officeResponse.getResponseBodyContent()
	
	// Verificar código de respuesta
	if (statusCode == 200 && responseBody && !responseBody.trim().isEmpty()) {
		KeywordUtil.logInfo("✅ Status code 200 para registro #${nro}")
		
		// Parsear respuesta JSON
		def respuestaJson = new JsonSlurper().parseText(responseBody)
		
		// Validar campos requeridos
		if (!respuestaJson.address || !respuestaJson.coordinates.longitude || !respuestaJson.coordinates.latitude) {
			KeywordUtil.markWarning("⚠️ Respuesta incompleta para registro #${nro}. No se encontró address o longitude o latitude.")
			
			resultadosValidacion.add([
				//longitud: (respuestaJson?.coordinates.longitude && respuestaJson.coordinates.longitude.trim()) ? respuestaJson.coordinates.longitude : "SIN RESULTADOS",
				nro: nro,
				longitud: "SIN RESULTADOS",
				latitud: "SIN RESULTADOS",
				direccion: direccion,
				codUbigeoEnviado: codUbigeoEnviado,
				codUbigeoObtenido: "SIN RESULTADOS",
				oficina: false,
				idAddress: idAddress,
				nombreOficina: nombreOficina
			])
			continue
		}
		// Extraer polígono del response (ajustar según estructura real del response)
		def codUbigeoObtenido = respuestaJson.ubigeo
		def longitude = respuestaJson.coordinates.longitude
		def latitude = respuestaJson.coordinates.latitude
		def isOffice = respuestaJson.office
		
		// Almacenar resultados para reporte
		resultadosValidacion.add([
			nro: nro,
			longitud: longitude,
			latitud: latitude,
			direccion: direccion,
			codUbigeoEnviado: codUbigeoEnviado,
			codUbigeoObtenido: codUbigeoObtenido,
			oficina: isOffice,
			idAddress: idAddress,
			nombreOficina: nombreOficina
		])
	} else {
		KeywordUtil.markWarning("⚠️ No se obtuvo contenido para registro #${nro} (Status code: ${statusCode})")
		resultadosValidacion.add([
			nro: nro,
			longitud: "SIN RESULTADOS",
			latitud: "SIN RESULTADOS",
			direccion: direccion,
			codUbigeoEnviado: codUbigeoEnviado,
			codUbigeoObtenido: "SIN RESULTADOS",
			oficina: false,
			idAddress: idAddress,
			nombreOficina: nombreOficina
		])
	}
}

// Resumen final
int totalRegistros = resultadosValidacion.size()
int exitosos = resultadosValidacion.count { it.oficina == true}
int fallidos = totalRegistros - exitosos

// Crear el archivo Excel
Workbook workbook = new XSSFWorkbook()
Sheet sheet = workbook.createSheet("Resultados Validación")

// Crear fila de encabezados
def excelHeaders = ["NRO", "LONGITUD", "LATITUD", "DIRECCIONES", "CODUBIGEO ENVIADO", "CODUBIGEO OBTENIDO", "OFICINA", "ID ADDRESS", "NOMBRE OFICINA"]
Row headerRow = sheet.createRow(0)
excelHeaders.eachWithIndex { header, idx ->
	Cell cell = headerRow.createCell(idx)
	cell.setCellValue(header)
}

// Llenar los datos
resultadosValidacion.eachWithIndex { resultado, index ->
	Row row = sheet.createRow(index + 1)
	row.createCell(0).setCellValue(resultado.nro.toString())
	row.createCell(1).setCellValue(resultado.longitud)
	row.createCell(2).setCellValue(resultado.latitud)
	row.createCell(3).setCellValue(resultado.direccion)
	row.createCell(4).setCellValue(resultado.codUbigeoEnviado)
	row.createCell(5).setCellValue(resultado.codUbigeoObtenido)
	row.createCell(6).setCellValue(resultado.oficina.toString())
	row.createCell(7).setCellValue(resultado.idAddress.toString())
	row.createCell(8).setCellValue(resultado.nombreOficina)
}

// Ajustar tamaño de columnas
excelHeaders.eachWithIndex { _, idx ->
	sheet.autoSizeColumn(idx)
}

// Obtener la fecha actual en formato ddMMyyyy_HHmmss
def fechaHoraActual = LocalDateTime.now().format(DateTimeFormatter.ofPattern("ddMMyyyy_HHmmss"))

// Crear el nombre del archivo con la fecha
def nombreArchivo = "Resultados_Validacion_Oficina_LongLat_${fechaHoraActual}.xlsx"

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
