package it.movioletto.helper.excel;

import it.movioletto.helper.excel.annotation.ExcelCell;
import it.movioletto.helper.excel.annotation.ExcelSheet;
import it.movioletto.helper.excel.annotation.ExcelTable;
import org.apache.commons.text.WordUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.Date;
import java.util.List;

class Write {

	static byte[] createSheet(Object foglio) throws IllegalAccessException {
		String sheetName = getSheetName(foglio);

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet(sheetName);

		int rowNum = getSheetOffsetX(foglio);
		int colNum = getSheetOffsetY(foglio);

		for (Field field : foglio.getClass().getDeclaredFields()) {
			field.setAccessible(true);
			String classNameField = field.getType().getName();
			classNameField = classNameField.substring(classNameField.lastIndexOf(".") + 1);

			List<?> lista;

			if((lista = (List<?>) field.get(foglio)) != null) {
				if ("List".equals(classNameField)) {
					rowNum = creaTitoloTabella(field, sheet, rowNum, colNum);

					rowNum = creaIntestazione(lista, sheet, rowNum, colNum);

					rowNum = aggiungiLista(lista, workbook, sheet, rowNum, colNum);

					rowNum++;
				}
			}
		}

		return getBytesFromWorkbook(workbook);
	}

	private static byte[] getBytesFromWorkbook(XSSFWorkbook workbook) {
		ByteArrayOutputStream outputStream = null;
		try {
			outputStream = new ByteArrayOutputStream();
			workbook.write(outputStream);
			workbook.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return outputStream.toByteArray();
	}

	private static String getSheetName(Object foglio) {
		String sheetName;
		ExcelSheet excelSheet = foglio.getClass().getAnnotation(ExcelSheet.class);

		if(excelSheet != null) {
			sheetName = excelSheet.titolo();
		} else {
			sheetName = foglio.getClass().getName();
		}
		return sheetName;
	}

	private static int getSheetOffsetX(Object foglio) {
		int sheetOffsetX;
		ExcelSheet excelSheet = foglio.getClass().getAnnotation(ExcelSheet.class);

		if(excelSheet != null) {
			sheetOffsetX = excelSheet.offsetX();
		}
		else {
			sheetOffsetX = 0;
		}

		return sheetOffsetX;
	}

	private static int getSheetOffsetY(Object foglio) {
		int sheetOffsetY;
		ExcelSheet excelSheet = foglio.getClass().getAnnotation(ExcelSheet.class);

		if(excelSheet != null) {
			sheetOffsetY = excelSheet.offsetY();
		}
		else {
			sheetOffsetY = 0;
		}

		return sheetOffsetY;
	}

	static byte[] createByList(List<?> list) throws IllegalAccessException {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("Foglio 0");

		int rowNum = 0;

		rowNum = creaIntestazione(list, sheet, rowNum, 0);

		aggiungiLista(list, workbook, sheet, rowNum, 0);

		return getBytesFromWorkbook(workbook);
	}

	private static int aggiungiLista(List<?> list, XSSFWorkbook workbook, XSSFSheet sheet, int rowNum, int colNumDef) throws IllegalAccessException {
		for (Object object: list) {
			Row row = sheet.createRow(rowNum++);

			int colNum = colNumDef;
			for (Field field : object.getClass().getDeclaredFields()) {
				Cell cell = row.createCell(colNum++);
				field.setAccessible(true);
				String classNameField = field.getType().getName();
				classNameField = classNameField.substring(classNameField.lastIndexOf(".") + 1);

				if(field.get(object) != null) {
					if ("String".equals(classNameField)) {
						cell.setCellValue((String) field.get(object));
					} else if ("Integer".equals(classNameField) || "int".equals(classNameField)) {
						cell.setCellValue((Integer) field.get(object));
					} else if ("Double".equalsIgnoreCase(classNameField)) {
						cell.setCellValue((Double) field.get(object));
					} else if ("BigDecimal".equalsIgnoreCase(classNameField)) {
						cell.setCellValue(((BigDecimal) field.get(object)).doubleValue());
					} else if ("Float".equalsIgnoreCase(classNameField)) {
						cell.setCellValue((Float) field.get(object));
					} else if ("Date".equalsIgnoreCase(classNameField)) {
						CellStyle cellStyle = workbook.createCellStyle();
						CreationHelper createHelper = workbook.getCreationHelper();
						cellStyle.setDataFormat(
								createHelper.createDataFormat().getFormat("mm/dd/yyyy hh:mm"));
						cell.setCellValue((Date) field.get(object));
						cell.setCellStyle(cellStyle);
					}
				}
			}
		}

		return rowNum;
	}

	private static int creaIntestazione(List<?> list, XSSFSheet sheet, int rowNum, int colNumDef) {
		if (list.isEmpty()) {
			return rowNum;
		}
		Object object = list.get(0);

		Row row = sheet.createRow(rowNum++);

		int colNum = colNumDef;
		for (Field field : object.getClass().getDeclaredFields()) {
			Cell cell = row.createCell(colNum++);

			boolean annotationTrovata = false;
			Annotation[] annotations = field.getDeclaredAnnotations();
			for (Annotation annotation : annotations) {
				if (annotation instanceof ExcelCell) {
					ExcelCell cellAnnotation = (ExcelCell) annotation;
					annotationTrovata = true;
					cell.setCellValue(cellAnnotation.intestazione());
				}
			}

			if (!annotationTrovata) {
				field.setAccessible(true);
				cell.setCellValue(splitCamelCase(WordUtils.capitalize(field.getName())));
			}
		}

		return rowNum;
	}

	private static int creaTitoloTabella(Field field, XSSFSheet sheet, int rowNum, int colNumDef) {

		int colNum = colNumDef;

		Annotation[] annotations = field.getDeclaredAnnotations();
		for(Annotation annotation : annotations){
			if(annotation instanceof ExcelTable){
				ExcelTable cellAnnotation = (ExcelTable) annotation;
				Row row = sheet.createRow(rowNum++);
				Cell cell = row.createCell(colNum++);
				cell.setCellValue(cellAnnotation.titolo());
			}
		}

		return rowNum;
	}

	private static String splitCamelCase(String s) {
		return s.replaceAll(
				String.format("%s|%s|%s",
						"(?<=[A-Z])(?=[A-Z][a-z])",
						"(?<=[^A-Z])(?=[A-Z])",
						"(?<=[A-Za-z])(?=[^A-Za-z])"
				),
				" "
		);
	}
}