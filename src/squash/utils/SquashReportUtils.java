package squash.utils;

import java.io.File;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.concurrent.ConcurrentHashMap;
import java.util.function.Function;
import java.util.function.Predicate;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.apache.commons.codec.binary.Base64;
import org.apache.poi.ss.usermodel.ClientAnchor;
import org.apache.poi.ss.usermodel.ClientAnchor.AnchorType;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import squash.model.LineReportSquash;
import squash.model.SquashReportData;
import squash.model.TestCaseData;

import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;

/**
 * 
 * @author VHU
 *
 */
public final class SquashReportUtils {
	private static int NbImageGenerated = 0;
	private static int NbHyperlinkDone = 0;
	private static int NbHyperlinkKO = 0;

	/**
	 * @return
	 */
	public static int getNbImageGenerated() {
		return NbImageGenerated;
	}

	/**
	 * @return
	 */
	public static int getNbHyperlinkDone() {
		return NbHyperlinkDone;
	}

	/**
	 * @return
	 */
	public static int getNbHyperlinkKO() {
		return NbHyperlinkKO;
	}

	public enum CellColumn {
		LABEL(0), ACTION(1), EXPECTED_RESULT(2), ID(3), DESCRIPTION(4), PRE_REQUISITE(5), STATUS_OK_NOK(6), PRIORITY(
				7), DURATION(8), NATURE(9), COMMENT(10), MODIFICATION_DATE(11), CREATED_BY(12);

		private int _value;

		private CellColumn(int i) {
			this._value = i;
		}

		public int getValue() {
			return _value;
		}
	}

	private static CellStyle longTextStyle;

	/**
	 * @param workBook
	 * @return
	 */
	private static CellStyle getLongTextStyle(XSSFWorkbook workBook) {
		if (longTextStyle == null) {
			longTextStyle = workBook.createCellStyle();
			longTextStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			longTextStyle.setWrapText(true);
		}
		return longTextStyle;
	}

	/**
	 * @param outputFile
	 * @param excelSquash
	 * @throws Exception
	 */
	public static void saveFormatedExcel(File outputFile, SquashReportData excelSquash) throws Exception {
		NbImageGenerated = 0;
		Map<String, List<TestCaseData>> sheetsMap = getMappedSheets(excelSquash);
		XSSFWorkbook workBook = createExcelFromMap(sheetsMap);
		
		FileOutputStream fos = new FileOutputStream(outputFile);
		workBook.write(fos);
		fos.close();
		workBook.close();
		System.out.println("Images generated: " + NbImageGenerated);
		System.out.println("Nb hyperlink ok: " + NbHyperlinkDone);
		System.out.println("Nb hyperlink ko: " + NbHyperlinkKO);
	}

	/**
	 * @param sheetsMap
	 * @return
	 * @throws Exception
	 */
	private static XSSFWorkbook createExcelFromMap(Map<String, List<TestCaseData>> sheetsMap) throws Exception {
		NbImageGenerated = 0;
		NbHyperlinkDone = 0;
		NbHyperlinkKO = 0;
		XSSFWorkbook workBook = new XSSFWorkbook();
		for (Map.Entry<String, List<TestCaseData>> entry : sheetsMap.entrySet()) {
			// System.out.println(entry.getKey());
			String sheetName = entry.getKey().split("/")[2];
			XSSFSheet sheet = workBook.getSheet(sheetName);
			// System.out.println(sheet);
			int rowNum = 0;
			if (sheet == null) {
				sheet = workBook.createSheet(sheetName);
				XSSFRow firstRow = sheet.createRow(rowNum);
				firstRow.createCell(CellColumn.LABEL.getValue()).setCellValue("Nom du test / Scope");
				firstRow.createCell(CellColumn.ACTION.getValue()).setCellValue("Action");
				firstRow.createCell(CellColumn.EXPECTED_RESULT.getValue()).setCellValue("Résultat attendu");
				firstRow.createCell(CellColumn.ID.getValue()).setCellValue("N°test");
				firstRow.createCell(CellColumn.DESCRIPTION.getValue()).setCellValue("Description");
				firstRow.createCell(CellColumn.PRE_REQUISITE.getValue()).setCellValue("Pré-requis");
				firstRow.createCell(CellColumn.STATUS_OK_NOK.getValue()).setCellValue("OK / NOK");
				firstRow.createCell(CellColumn.PRIORITY.getValue()).setCellValue("Priorité");
				firstRow.createCell(CellColumn.DURATION.getValue()).setCellValue("Durée");
				firstRow.createCell(CellColumn.NATURE.getValue()).setCellValue("Nature");
				firstRow.createCell(CellColumn.COMMENT.getValue()).setCellValue("Commentaire");
				firstRow.createCell(CellColumn.MODIFICATION_DATE.getValue()).setCellValue("Date de modification");
				firstRow.createCell(CellColumn.CREATED_BY.getValue()).setCellValue("Auteur");
				rowNum++;
			} else
				rowNum = sheet.getLastRowNum() + 1;

			// boucler sur les testCase
			for (TestCaseData testCase : entry.getValue()) {
				XSSFRow currentRow = sheet.createRow(rowNum);
				// cellule label
				XSSFCell cell = currentRow.createCell(CellColumn.LABEL.getValue());

				CellStyle style = getLongTextStyle(workBook);
				cell.setCellStyle(style);
				cell.setCellValue(testCase.getLabel());
				sheet.setColumnWidth(0, 30 * 256);

				currentRow.createCell(CellColumn.ID.getValue()).setCellValue(testCase.getId());
				currentRow.createCell(CellColumn.DESCRIPTION.getValue())
						.setCellValue(replaceHtmlTags(testCase.getDescription(), workBook));
				currentRow.createCell(CellColumn.PRE_REQUISITE.getValue())
						.setCellValue(replaceHtmlTags(testCase.getPreRequisite(), workBook));
				currentRow.createCell(CellColumn.PRIORITY.getValue()).setCellValue(testCase.getWeight());
				currentRow.createCell(CellColumn.NATURE.getValue()).setCellValue(testCase.getNature());
				currentRow.createCell(CellColumn.MODIFICATION_DATE.getValue())
						.setCellValue(testCase.getDateModification());
				currentRow.createCell(CellColumn.CREATED_BY.getValue()).setCellValue(testCase.getCreatedBy());

				List<String> actions = testCase.getActions();
				List<String> expectedResults = testCase.getExpectedResults();
				int nbRowsToCreate = Math.max(actions.size(), expectedResults.size());
				for (int i = 0; i < nbRowsToCreate; i++) {
					currentRow = sheet.getRow(rowNum + i);
					if (currentRow == null)
						currentRow = sheet.createRow(rowNum + i);
					if (i < actions.size()) {
						createCellWithHandledImages(workBook, sheet, currentRow, rowNum + i,
								CellColumn.ACTION.getValue(), actions.get(i));
					}
					if (i < expectedResults.size()) {
						createCellWithHandledImages(workBook, sheet, currentRow, rowNum + i,
								CellColumn.EXPECTED_RESULT.getValue(), expectedResults.get(i));
					}
				}
				// System.out.println("test label: " + testCase.getLabel());
				if (rowNum + nbRowsToCreate - 1 > rowNum) {
					mergeRegion(sheet, rowNum, rowNum + nbRowsToCreate - 1, CellColumn.LABEL.getValue(),
							CellColumn.LABEL.getValue());
					mergeRegion(sheet, rowNum, rowNum + nbRowsToCreate - 1, CellColumn.DESCRIPTION.getValue(),
							CellColumn.DESCRIPTION.getValue());
					mergeRegion(sheet, rowNum, rowNum + nbRowsToCreate - 1, CellColumn.PRE_REQUISITE.getValue(),
							CellColumn.PRE_REQUISITE.getValue());
				}
				rowNum += nbRowsToCreate;
			}

		}
		handleHyperLinks(workBook);
		
		return workBook;
	}

	/**
	 * 
	 * @param workBook
	 */
	private static void handleHyperLinks(XSSFWorkbook workBook) {
		//un test peut avoir plusieurs liens le référençant
		// ou sheetName!cellule / label du test?
		// stock: label test / adresse colonne à lier
		Map<String, String> map = new HashMap<String, String>();
		Map<String, String> labelMap = new HashMap<String, String>();
		
		for(int i=0; i < workBook.getNumberOfSheets(); i++){
			XSSFSheet sheet = workBook.getSheetAt(i);
			//System.out.println(sheet.getSheetName());
			int nbRows = sheet.getLastRowNum();
			for(int j = 0; j < nbRows; j++){
				Cell cell = sheet.getRow(j).getCell(CellColumn.ACTION.getValue());
				if(cell.getStringCellValue().contains("Calls :")){
					CellAddress cellAddress = new CellAddress(cell);
					//System.out.println(cellAddress.formatAsString());
					String cellValue = cell.getStringCellValue();
					//System.out.println(cellValue);
					map.put("'"+sheet.getSheetName()+"'!"+cellAddress.formatAsString(), cellValue.split("Calls : ")[1]);
				}

				Cell labelCell = sheet.getRow(j).getCell(CellColumn.LABEL.getValue());
				if(labelCell != null){
					CellAddress cellAddress = new CellAddress(labelCell);
					labelMap.put(labelCell.getStringCellValue(), "'"+sheet.getSheetName()+"'!"+cellAddress.formatAsString());
				}
				
			}
		}
		//System.out.println(map);
		//System.out.println(labelMap);
		
		CellStyle hlink_style = workBook.createCellStyle();
	    Font hlink_font = workBook.createFont();
	    hlink_font.setUnderline(Font.U_SINGLE);
	    hlink_font.setColor(IndexedColors.BLUE.getIndex());
	    hlink_style.setFont(hlink_font);
		for(Map.Entry<String, String> entry : map.entrySet()){
			try {
				String labelCellAddress = labelMap.get(entry.getValue());
				if(labelCellAddress != null && !labelCellAddress.isEmpty()){
					//System.out.println("labelFound: " + entry.getValue() + ", address: " + labelCellAddress);
					CellReference targetCellRef = new CellReference(labelCellAddress);
					
					if(targetCellRef != null){
						//System.out.println("cell ref: " + targetCellRef + " row:" + targetCellRef.getRow() + ",col: " + targetCellRef.getCol() + ",sheetName: " + targetCellRef.getSheetName());
						//System.out.println(targetCellRef.formatAsString());
						Hyperlink link = workBook.getCreationHelper().createHyperlink(HyperlinkType.DOCUMENT);
						link.setAddress(targetCellRef.formatAsString());
						CellReference cellRef = new CellReference(entry.getKey());
						if(cellRef != null){
							Cell cell = workBook.getSheet(cellRef.getSheetName()).getRow(cellRef.getRow()).getCell(cellRef.getCol());
							if(cell != null) {
								cell.setHyperlink(link);
								cell.setCellStyle(hlink_style);
							}
						}
					}
				}
				NbHyperlinkDone++;
			} catch(Exception ex){
				NbHyperlinkKO++;
			}
		}
		
	}
	
	/**
	 * @param sheet
	 * @param firstRow
	 * @param lastrow
	 * @param firstCol
	 * @param lastCol
	 */
	private static void mergeRegion(XSSFSheet sheet, int firstRow, int lastrow, int firstCol, int lastCol) {
		sheet.addMergedRegion(new CellRangeAddress(firstRow, lastrow, firstCol, lastCol));
	}

	/**
	 * @param excelSquash
	 * @return
	 */
	private static Map<String, List<TestCaseData>> getMappedSheets(SquashReportData excelSquash) {
		Map<String, List<TestCaseData>> sheetsMap = new HashMap<String, List<TestCaseData>>();
		List<String> distinctPathList = excelSquash.getLinesOfContent().stream()
				.filter(distinctByKey(LineReportSquash::getPath)).filter(x -> !x.getPath().equals("PATH"))
				.map(LineReportSquash::getPath).collect(Collectors.toList());
		distinctPathList.forEach(dp -> {
			List<TestCaseData> listCases = new ArrayList<TestCaseData>();
			List<LineReportSquash> listSameHierarchy = excelSquash.getLinesOfContent().stream()
					.filter(x -> x.getPath().equals(dp)).collect(Collectors.toList());
			List<LineReportSquash> listLabels = listSameHierarchy.stream()
					.filter(distinctByKey(LineReportSquash::getLabel)).collect(Collectors.toList());
			listLabels.forEach(x -> {
				TestCaseData testCase = new TestCaseData();
				testCase.setHierarchy(x.getPath());
				testCase.setId(x.getId());
				testCase.setLabel(x.getLabel());
				testCase.setCreatedBy(x.getCreated_by());
				testCase.setDateModification(
						x.getLast_modified_on() != null ? x.getLast_modified_on() : x.getCreated_on());
				testCase.setNature(x.getNature());
				testCase.setPreRequisite(x.getPre_requisite());
				testCase.setWeight(x.getWeight());
				testCase.setDescription(x.getDescription());
				List<LineReportSquash> listTestCaseByLabel = excelSquash.getLinesOfContent().stream()
						.filter(y -> y.getLabel().equals(x.getLabel())).collect(Collectors.toList());
				listTestCaseByLabel.forEach(y -> {
					testCase.addAction(y.getAction());
					testCase.addExpectedResults(y.getExpected_result());
				});
				listCases.add(testCase);
			});
			// (key=sheetName, value=liste test case)
			sheetsMap.put(dp, listCases);
		});

		return sheetsMap;
	}

	/**
	 * @param outputFile
	 * @param excelSquash
	 * @throws Exception
	 */
	public static void saveTranslationOfCsv(File outputFile, SquashReportData excelSquash) throws Exception {
		XSSFWorkbook workBook = new XSSFWorkbook();
		XSSFSheet sheet = workBook.createSheet();

		List<LineReportSquash> lines = excelSquash.getLinesOfContent();
		int rowNum = 0;
		for (LineReportSquash line : lines) {
			XSSFRow currentRow = sheet.createRow(rowNum);
			if (rowNum > 0)
				currentRow.setHeightInPoints(150);
			currentRow.createCell(0).setCellValue(line.getPath());
			currentRow.createCell(1).setCellValue(line.getId());
			currentRow.createCell(2).setCellValue(line.getReference());
			currentRow.createCell(3).setCellValue(line.getLabel());
			currentRow.createCell(4).setCellValue(line.getWeight());
			currentRow.createCell(5).setCellValue(line.getNature());
			currentRow.createCell(6).setCellValue(line.getType());
			currentRow.createCell(7).setCellValue(line.getStatus());
			currentRow.createCell(8).setCellValue(line.getDescription());
			currentRow.createCell(9).setCellValue(line.getPre_requisite());
			currentRow.createCell(10).setCellValue(line.getCreated_on());
			currentRow.createCell(11).setCellValue(line.getCreated_by());
			currentRow.createCell(12).setCellValue(line.getLast_modified_on());
			currentRow.createCell(13).setCellValue(line.getLast_modified_by());
			createCellWithHandledImages(workBook, sheet, currentRow, rowNum, 14, line.getAction());
			createCellWithHandledImages(workBook, sheet, currentRow, rowNum, 15, line.getExpected_result());
			rowNum++;
		}

		setAutosizeToColumns(sheet);
		FileOutputStream fos = new FileOutputStream(outputFile);
		workBook.write(fos);
		fos.close();
		workBook.close();
	}

	/**
	 * @param sheet
	 */
	private static void setAutosizeToColumns(XSSFSheet sheet) {
		List<Integer> columnsToAutosize = new ArrayList<Integer>();
		columnsToAutosize.add(0);
		columnsToAutosize.add(1);
		columnsToAutosize.add(2);
		// columnsToAutosize.add(3);
		columnsToAutosize.add(4);
		columnsToAutosize.add(5);
		columnsToAutosize.add(6);
		columnsToAutosize.add(7);
		// columnsToAutosize.add(8);
		columnsToAutosize.add(9);
		columnsToAutosize.add(10);
		columnsToAutosize.add(11);
		columnsToAutosize.add(12);
		columnsToAutosize.add(13);
		columnsToAutosize.forEach(col -> sheet.autoSizeColumn(col));
	}

	/**
	 * @param workBook
	 * @param sheet
	 * @param row
	 * @param rowNum
	 * @param col
	 * @param value
	 * @throws Exception
	 */
	private static void createCellWithHandledImages(XSSFWorkbook workBook, XSSFSheet sheet, XSSFRow row, int rowNum,
			int col, String value) throws Exception {
		String encodingPrefix = "base64,";
		Pattern pattern = Pattern.compile("<img[^>]*(" + encodingPrefix + ")*>");

		String lineContent = value;
		Matcher matcher = pattern.matcher(lineContent);

		List<byte[]> listImages = new ArrayList<byte[]>();
		if (matcher.find()) {
			for (int i = 0; i < matcher.groupCount(); i++) {
				String imgUrl = matcher.group(i);
				int imageStartIndex = imgUrl.indexOf(encodingPrefix) + encodingPrefix.length();
				int imageEndIndex = imgUrl.substring(imageStartIndex).indexOf('"');
				listImages.add(Base64.decodeBase64(imgUrl.substring(imageStartIndex, imageStartIndex + imageEndIndex)));
			}
		}
		Cell cell = row.createCell(col);
		
		CellStyle style = getLongTextStyle(workBook);
		cell.setCellStyle(style);
		cell.setCellValue(replaceHtmlTags(value.replaceAll("<img[^>]*>", "<IMAGE A INSERER>"), workBook));
		sheet.setColumnWidth(col, 60 * 256);
		for (byte[] img : listImages) {
			int pictureIdx = workBook.addPicture(img, Workbook.PICTURE_TYPE_PNG);

			int left = 0;
			int top = 0;
			int width = Math.round(sheet.getColumnWidthInPixels(col) - left - left); // width
																						// in
																						// px
			int height = Math.round(200 - top - 10/* pt */); // height
																// in
																// pt
			// System.out.println("width: " + width + ", height: " + height);
			drawImageOnExcelSheet(sheet, rowNum, col, left, top, width, height, pictureIdx);
		}

	}

	/**
	 * @param sheet
	 * @param row
	 * @param col
	 * @param left
	 * @param top
	 * @param width
	 * @param height
	 * @param pictureIdx
	 * @throws Exception
	 */
	private static void drawImageOnExcelSheet(XSSFSheet sheet, int row, int col, int left/* in px */,
			int top/* in pt */, int width/* in px */, int height/* in pt */, int pictureIdx) throws Exception {

		CreationHelper helper = sheet.getWorkbook().getCreationHelper();

		@SuppressWarnings("rawtypes")
		Drawing drawing = sheet.createDrawingPatriarch();

		ClientAnchor anchor = helper.createClientAnchor();
		anchor.setAnchorType(AnchorType.MOVE_AND_RESIZE);

		anchor.setCol1(col); // first anchor determines upper left position
		anchor.setRow1(row);
		anchor.setDx1(Units.pixelToEMU(left)); // dx = left in px
		anchor.setDy1(Units.toEMU(top)); // dy = top in pt

		anchor.setCol2(col); // second anchor determines bottom right position
		anchor.setRow2(row);
		anchor.setDx2(Units.pixelToEMU(left + width)); // dx = left + wanted
														// width in px
		anchor.setDy2(Units.toEMU(top + height)); // dy= top + wanted height in
													// pt

		drawing.createPicture(anchor, pictureIdx);
		NbImageGenerated++;
		// System.out.println("img created: row=" + row + ", col="+col);
	}

	/**
	 * @param keyExtractor
	 * @return
	 */
	private static <T> Predicate<T> distinctByKey(Function<? super T, ?> keyExtractor) {
		Set<Object> seen = ConcurrentHashMap.newKeySet();
		return t -> seen.add(keyExtractor.apply(t));
	}

	/**
	 * @param value
	 * @param workBook
	 * @return
	 */
	private static XSSFRichTextString replaceHtmlTags(String value, XSSFWorkbook workBook) {
		String result = value;
		result = result.replace("<p>", "").replace("</p>", "\n").replace("<li>", "- ").replace("</li>", "")
				.replaceAll("</*ol>", "").replace("<br />", "\n");
		XSSFRichTextString richText = new XSSFRichTextString(result);

		String boldTag = "<strong>";
		if (result.contains("<u>")) {
			XSSFFont font = workBook.createFont();
			font.setUnderline(XSSFFont.U_SINGLE);

			String underlineAndBold = "<u>" + boldTag;
			if (result.contains(underlineAndBold)) {
				font.setBold(true);
			}

			while (result.contains(underlineAndBold)) {
				int indexBegin = result.indexOf(underlineAndBold);
				int indexEnd = result.indexOf("</strong></u>");
				result = result.substring(0, indexBegin)
						+ result.substring(indexBegin + underlineAndBold.length(), indexEnd)
						+ result.substring(indexEnd + underlineAndBold.length() + 2);
				richText = new XSSFRichTextString(result);
				richText.applyFont(indexBegin, indexEnd - underlineAndBold.length(), font);
			}
		}

		if (result.contains(boldTag)) {
			XSSFFont font = workBook.createFont();
			font.setBold(true);

			while (result.contains(boldTag)) {
				int indexBegin = result.indexOf(boldTag);
				int indexEnd = result.indexOf("</strong>");
				result = result.substring(0, indexBegin) + result.substring(indexBegin + boldTag.length(), indexEnd)
						+ result.substring(indexEnd + boldTag.length() + 1);
				richText = new XSSFRichTextString(result);
				richText.applyFont(indexBegin, indexEnd - boldTag.length(), font);
			}
		}

		return richText;
	}
}
