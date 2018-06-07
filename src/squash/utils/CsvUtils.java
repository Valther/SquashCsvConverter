package squash.utils;

import java.io.File;
import java.io.IOException;
import java.nio.charset.Charset;
import java.util.ArrayList;
import java.util.List;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;

import squash.model.SquashReportData;

public class CsvUtils {
	File csvFile;

	public CsvUtils(File file) {
		csvFile = file;
	}

	private CSVParser parse(CSVFormat format) throws IOException {
		CSVParser parser = CSVParser.parse(csvFile, Charset.defaultCharset(), format);
		return parser;
	}

	/**
	 * Enregistre un excel qui est la traduction littérale (sans transformation) du csv fournis
	 * @throws Exception
	 */
	public void csvToXlsx() throws Exception {
		SquashReportData excelSquash = getSquashReporteData();
		String nameFile = csvFile.getName().substring(0, csvFile.getName().lastIndexOf('.'));
		File outputFile = new File(csvFile.getParent() + File.separator + nameFile +".xlsx");
		SquashReportUtils.saveTranslationOfCsv(outputFile, excelSquash);
	}

	public File csvToFormattedXlsx() throws Exception {
		SquashReportData excelSquash = getSquashReporteData();
		
		String nameFile = csvFile.getName().substring(0, csvFile.getName().lastIndexOf('.')) + "-modified";
		File outputFile = new File(csvFile.getParent() + File.separator + nameFile +".xlsx");
		SquashReportUtils.saveFormatedExcel(outputFile, excelSquash);
		return outputFile;
	}
	
	private SquashReportData getSquashReporteData() throws IOException {
		if(csvFile == null)
			throw new IOException("Le fichier CSV à traiter ne peut être null");
		CSVFormat csvFileFormat = CSVFormat.EXCEL.withDelimiter(';').withRecordSeparator("\n");
		CSVParser parser = this.parse(csvFileFormat);
		SquashReportData excelSquash = new SquashReportData();
		for(CSVRecord record : parser){
			List<String> content = new ArrayList<String>();
			for(String field : record){
				content.add(field);
			}
			excelSquash.addLineOfContent(content);
		}
		return excelSquash;
	}
	
}
