import java.sql.*
import java.util.regex.*
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.*
import java.io.FileOutputStream

class DireccionValidator {

	List<String> ciudadesConocidas = ["tacna", "pocollay", "lima", "arequipa", "miraflores", "san isidro"]
	List<String> manzanasLotes = ["lt", "mz", "lot", "lte", "manz", "mzn", "km", "secto"]
	List<String> caracteresEspeciales = ["!", "@", "#", "\$", "%", "&", "*", "/", "\\", "ª", "º", "©", "³", "±"]
	Map<String, String> erroresTipograficos = ["av": "avenida", "jr": "jiron", "clle": "calle", "cl": "calle", "cale": "calle", "mz": "manzana", "maz": "manzana", "lte": "lote", "urb" : "urbanizacion", "lt": "lote"]

	void validar() {
		// Para poder ejecutarse se necesita que la VPN esté activada
		def connection = DriverManager.getConnection("jdbc:postgresql://10.136.0.2:5444/address_dev", "usr_address_read", "DWHXTBUi8ZlpvjunPkme6C")
		def statement = connection.createStatement()
		// La cantidad de direcciones a validar es a criterio
		def sql = '''
		    SELECT t.tracking, t.address_id, t.address, t.address_normalized
		    FROM olva.trackings t
		    WHERE t.fecha_emision > '2025-06-08'
		      AND t.emision = '25'
		    ORDER BY t.tracking DESC
		    LIMIT 10
		'''
		
		def resultSet = statement.executeQuery(sql)

		XSSFWorkbook workbook = new XSSFWorkbook()
		XSSFSheet sheet = workbook.createSheet("ValidacionDirecciones")

		def header = ["tracking", "address_id", "address", "address_normalized", "Errores"]
		def headerRow = sheet.createRow(0)
		header.eachWithIndex { col, idx ->
			headerRow.createCell(idx).setCellValue(col)
		}

		int rowIndex = 1
		while (resultSet.next()) {
			def id = resultSet.getInt("tracking")
			def addressId = resultSet.getInt("address_id")
			def original = resultSet.getString("address")?.toLowerCase()?.trim()
			def normalizada = resultSet.getString("address_normalized")?.toLowerCase()?.trim()

			List<String> errores = []

			if (addressId == 0) errores.add("No georreferenciado")
			if (original == normalizada) errores.add("Sin normalización visible")
			errores.addAll(clasificarErrores(original, normalizada))

			def row = sheet.createRow(rowIndex++)
			row.createCell(0).setCellValue(id)
			row.createCell(1).setCellValue(addressId)
			row.createCell(2).setCellValue(original ?: "")
			row.createCell(3).setCellValue(normalizada ?: "")
			row.createCell(4).setCellValue(errores.join("; ") ?: "Correcta")
		}

		String archivo = "Validacion_Direcciones.xlsx"
		FileOutputStream out = new FileOutputStream(archivo)
		workbook.write(out)
		out.close()

		println "✅ Validación completada. Archivo generado: ${archivo}"

		resultSet.close()
		statement.close()
		connection.close()
	}

	List<String> clasificarErrores(String direccion, String direccionNormalizada) {
		List<String> errores = []
	
		if (!direccion || direccion.length() < 8) return ["Dirección vacía o demasiado corta"]
	
		// === VALIDACIÓN DE GUION Y TEXTO POSTERIOR ===
		def palabrasClave = ["avenida", "calle", "jiron", "jirón", "manzana", "lote", "urbanización", "jr", "av", "mz", "ca."]

		String direccionBase = direccion
		String direccionNormalizadaBase = direccionNormalizada
		
		if (direccion.contains("-")) {
			def partes = direccion.split(/\s*-\s*/, 2)
			def parteIzquierda = partes[0]?.trim()
			def parteDerecha = partes.size() > 1 ? partes[1]?.trim() : ""
		
			boolean contienePalabraClave = palabrasClave.any { palabra ->
				parteIzquierda.toLowerCase().contains(palabra)
			}
		
			if (contienePalabraClave && parteDerecha) {
				def normalizadaMinuscula = direccionNormalizada.toLowerCase()
				def textoPosterior = parteDerecha.toLowerCase()
		
				// 🛡️ EXCEPCIONES PERMITIDAS
				def excepcionesValidas = [
					/[a-záéíóúñ]+\-\d+/,         // ej: manzana-4, lt-1, mz-4
					/\d+\-\d+/,                 // ej: 123-125
					/[a-záéíóúñ]+\-[a-záéíóúñ]+/  // ej: san-juan
				]
		
				boolean esExcepcionValida = excepcionesValidas.any { regex ->
					(direccion =~ regex).findAll().any { match ->
						normalizadaMinuscula.contains(match.toLowerCase())
					}
				}
		
				// Solo marcar error si NO es una excepción
				boolean guionPersistente = normalizadaMinuscula.contains(" - ")
				boolean textoPersistente = textoPosterior.split(/\s+/).any { palabra ->
					normalizadaMinuscula.contains(palabra)
				}
		
				// Esta validación tiene algunas deficiencias por trabajar
				if (!esExcepcionValida && (guionPersistente || textoPersistente)) {
					errores.add("No se eliminó texto posterior al guion o el guion mismo: '- ${parteDerecha}'")
				}
		
				// Ajustamos para las validaciones siguientes
				if (!esExcepcionValida) {
					direccionBase = parteIzquierda
					direccionNormalizadaBase = direccionNormalizada.split(/\s*-\s*/)[0]?.trim()
				}
			}
		}
	
		// Usar direccionBase y direccionNormalizadaBase para todas las demás validaciones
	
		if (direccionBase.contains("@") && direccionNormalizadaBase.contains("@") ) errores.add("Dirección con correo electrónico")
		if ((direccionBase =~ /\b9\d{8}\b/ || direccionBase =~ /\b\d{6,}\b/) &&
			(direccionNormalizadaBase =~ /\b9\d{8}\b/ || direccionNormalizadaBase =~ /\b\d{6,}\b/)) {
			errores.add("Contiene número telefónico")
		}
	
		ciudadesConocidas.each { ciudad ->
			if (direccionBase.startsWith(ciudad) && direccionNormalizadaBase.startsWith(ciudad)) {
				errores.add("Inicia con nombre de ciudad: ${ciudad}")
			}
		}
	
		if ((direccionBase =~ /(aa\.?hh|asoc|a\.h\.|municipalid|ministeri|fiscali|condominio|cooperativ)/) &&
			(direccionNormalizadaBase =~ /(aa\.?hh|asoc|a\.h\.|municipalid|ministeri|fiscali|condominio|cooperativ)/)) {
			errores.add("Dirección institucional o de asentamiento")
		}
	
		manzanasLotes.each { termino ->
			def pattern = ~/(?i)\b${termino}\b/
			if (direccionBase ==~ pattern && direccionNormalizadaBase ==~ pattern) {
				errores.add("Contiene término de lote/manzana: ${termino}")
			}
		}
	
		if ((direccionBase.contains("s/n") || direccionBase.contains("nro null") || !(direccionBase =~ /\b\d{1,4}\b/)) &&
			(direccionNormalizadaBase.contains("s/n") || direccionNormalizadaBase.contains("nro null") || !(direccionNormalizadaBase =~ /\b\d{1,4}\b/))) {
			errores.add("Dirección sin número")
		}
	
		if ((direccionBase =~ /\b(\w+)\s+\1\b/) && (direccionNormalizadaBase =~ /\b(\w+)\s+\1\b/)) {
			errores.add("Palabra o número duplicado")
		}
	
		if (direccionBase.contains("referencia") && direccionNormalizadaBase.contains("referencia")) {
			errores.add("Contiene palabra 'referencia'")
		}
	
		caracteresEspeciales.each {
			if (direccionBase.contains(it) && direccionNormalizadaBase.contains(it)) {
				errores.add("Caracter especial: ${it}")
			}
		}
	
		erroresTipograficos.each { original, esperado ->
			def patternOriginal = ~/(?i)\b${original}\b/
			def patternEsperado = ~/(?i)\b${esperado}\b/
		
			if (direccionBase.toLowerCase().find(patternOriginal) && direccionNormalizadaBase.toLowerCase().find(patternOriginal)) {
				errores.add("No se normalizó '${original}' a '${esperado}'")
			}
		}
	
		return errores
	}
}

// 🟩 Llamada al final del test case:
new DireccionValidator().validar()
