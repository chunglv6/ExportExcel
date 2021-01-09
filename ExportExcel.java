package excelExport;

import java.beans.IntrospectionException;
import java.beans.Introspector;
import java.beans.PropertyDescriptor;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.sql.Timestamp;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.BuiltinFormats;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExportExcel {
	private static CellStyle cellStyleFormatNumber = null;
	private static CellStyle cellStyleFormatDate = null;

	public static void main(String[] args) throws Exception {
		List<Student> students = new ArrayList<Student>();

		Class<Student> classExport = Student.class;
		students.add(new Student(1, "nguyen van a",new Date(),new Timestamp(new Date().getTime())));
		students.add(new Student(2, "nguyen van b",new Date(),new Timestamp(new Date().getTime())));
		students.add(new Student(3, "nguyen van c",new Date(),new Timestamp(new Date().getTime())));

		XSSFWorkbook wb = writeExcel(students);
		String path = "D:\\chunglv6\\" + classExport.getName() + ".xlsx";
		try (FileOutputStream fos = new FileOutputStream(path)) {
			wb.write(fos);
			System.out.println("done !");
		} catch (Exception e) {
			throw new Exception("co loi xay ra : " + e.getMessage());
		}

	}

	public static List<String> convertMethodName(List<String> headers, List<String> methods) {
		List<String> methodNameMap = new ArrayList<String>();
		for (String header : headers) {
			for (String method : methods) {
				if (header.toLowerCase().equals(method.substring(3, method.length()).toLowerCase())) {
					methodNameMap.add(method);
				}
			}
		}
		return methodNameMap;
	}

	public static List<String> getListHeader(Class<? extends Object> c) {
		List<String> headers = new ArrayList<String>();
		Field[] fields = c.getDeclaredFields();
		for (int i = 0; i < fields.length; i++) {
			String field = fields[i].toString();

			headers.add(field.substring(field.lastIndexOf(".") + 1, field.length()));
		}
		return headers;
	}

	public static List<String> getMethodName(Class<? extends Object> c) {
		List<String> listGetters = new ArrayList<String>();

		try {
			for (PropertyDescriptor propertyDescriptor : Introspector.getBeanInfo(c, Object.class)
					.getPropertyDescriptors()) {
				String getter = propertyDescriptor.getReadMethod().toString();
				listGetters.add(getter.substring(getter.lastIndexOf(".") + 1, getter.length() - 2));
			}
		} catch (IntrospectionException e) {
			e.printStackTrace();
		}
		return listGetters;
	}

	public static XSSFWorkbook writeExcel(List<? extends Object> objs)
			throws IOException, IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		XSSFWorkbook workbook = new XSSFWorkbook();
		// Create sheet
		XSSFSheet sheet = workbook.createSheet(); // Create sheet with sheet name
		if (objs.isEmpty()) {
			return workbook;
		}
		Class<? extends Object> classExport = objs.get(0).getClass();
		List<String> headers = getListHeader(classExport);
		List<String> methods = convertMethodName(headers, getMethodName(classExport));

		int rowIndex = 0;

		// Write header
		writeHeader(sheet, convertHeader(headers));

		// Write data
		rowIndex++;
		for (Object obj : objs) {
			// Create row
			Row row = sheet.createRow(rowIndex);
			// Write data on row
			writeBook(obj, row, methods);
			rowIndex++;
		}
		// Auto resize column witdth
		int numberOfColumn = sheet.getRow(0).getPhysicalNumberOfCells();
		autosizeColumn(sheet, numberOfColumn);
		return workbook;

	}

	// Write header with format
	private static void writeHeader(XSSFSheet sheet, List<String> headers) {
		// create CellStyle
		CellStyle cellStyle = createStyleForHeader(sheet);

		// Create row
		XSSFRow row = sheet.createRow(0);

		// Create cells
		for (int i = 0; i < headers.size(); i++) {
			XSSFCell cell = row.createCell(i);
			cell.setCellStyle(cellStyle);
			cell.setCellValue(headers.get(i));
		}

	}

	private static CellStyle getBorderStyle(Row row) {
		CellStyle cellStyle = row.getSheet().getWorkbook().createCellStyle();
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		return cellStyle;
	}

	private static CellStyle createCellStyleFormatNumber(Row row) {
		// Format number
		short format = (short) BuiltinFormats.getBuiltinFormat("#,##0");
		// Create CellStyle
		Workbook workbook = row.getSheet().getWorkbook();
		CellStyle cellStyleFormatNumber = workbook.createCellStyle();
		cellStyleFormatNumber.setDataFormat(format);
		cellStyleFormatNumber.setBorderBottom(BorderStyle.THIN);
		cellStyleFormatNumber.setBorderLeft(BorderStyle.THIN);
		cellStyleFormatNumber.setBorderRight(BorderStyle.THIN);
		cellStyleFormatNumber.setBorderTop(BorderStyle.THIN);
		cellStyleFormatNumber.setAlignment(HorizontalAlignment.CENTER);
		cellStyleFormatNumber.setVerticalAlignment(VerticalAlignment.CENTER);
		return cellStyleFormatNumber;
	}

	private static CellStyle createCellStyleFormatDate(Row row) {
		Workbook workbook = row.getSheet().getWorkbook();
		CreationHelper createHelper = workbook.getCreationHelper();

		CellStyle cellStyleFormatDate = workbook.createCellStyle();
		cellStyleFormatDate.setDataFormat(createHelper.createDataFormat().getFormat("dd/mm/yyyy"));
		cellStyleFormatDate.setBorderBottom(BorderStyle.THIN);
		cellStyleFormatDate.setBorderLeft(BorderStyle.THIN);
		cellStyleFormatDate.setBorderRight(BorderStyle.THIN);
		cellStyleFormatDate.setBorderTop(BorderStyle.THIN);
		cellStyleFormatDate.setAlignment(HorizontalAlignment.CENTER);
		cellStyleFormatDate.setVerticalAlignment(VerticalAlignment.CENTER);
		return cellStyleFormatDate;

	}

	// Write data
	private static <T> void writeBook(T obj, Row row, List<String> methods)
			throws IllegalAccessException, IllegalArgumentException, InvocationTargetException {
		CellStyle borderStyle = getBorderStyle(row);
		cellStyleFormatNumber = createCellStyleFormatNumber(row);
		cellStyleFormatDate = createCellStyleFormatDate(row);
		int index = 0;
		for (String method : methods) {
			Cell cell = row.createCell(index);
			try {
				Method methodName = obj.getClass().getMethod(method);
				Object o = methodName.invoke(obj);
				cell.setCellStyle(borderStyle);
				setCellValue(o, cell);

			} catch (NoSuchMethodException | SecurityException e) {
				e.printStackTrace();
			}
			index++;
		}
	}

	public static void setCellValue(Object obj, Cell cell) {
		if (obj instanceof Long) {
			cell.setCellValue((Long) obj);
			cell.setCellStyle(cellStyleFormatNumber);
		} else if (obj instanceof Date) {
			cell.setCellValue((Date) obj);
			cell.setCellStyle(cellStyleFormatDate);
		} else if (obj instanceof Integer) {
			cell.setCellValue((Integer) obj);
			cell.setCellStyle(cellStyleFormatNumber);
		} else if (obj instanceof Double) {
			cell.setCellValue((Double) obj);
			cell.setCellStyle(cellStyleFormatNumber);
		} else if (obj instanceof Float) {
			cell.setCellValue((Float) obj);
			cell.setCellStyle(cellStyleFormatNumber);
		} else {

			cell.setCellValue((String) obj);
		}
	}

	// Create CellStyle for header
	private static CellStyle createStyleForHeader(Sheet sheet) {
		// Create font
		Font font = sheet.getWorkbook().createFont();
		font.setFontName("Times New Roman");
		font.setBold(true);
		font.setFontHeightInPoints((short) 14); // font size
		font.setColor(IndexedColors.WHITE.getIndex()); // text color

		// Create CellStyle
		CellStyle cellStyle = sheet.getWorkbook().createCellStyle();
		cellStyle.setFont(font);
		cellStyle.setFillForegroundColor(IndexedColors.BLUE.getIndex());
		cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		cellStyle.setBorderBottom(BorderStyle.THIN);
		cellStyle.setBorderLeft(BorderStyle.THIN);
		cellStyle.setBorderRight(BorderStyle.THIN);
		cellStyle.setBorderTop(BorderStyle.THIN);
		cellStyle.setAlignment(HorizontalAlignment.CENTER);
		cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
		return cellStyle;
	}

	// Write footer
//	private static void writeFooter(XSSFSheet sheet, int rowIndex) {
//		if (cellStyleFormatNumber == null) {
//			short format = (short) BuiltinFormats.getBuiltinFormat("#,##0");
//			Workbook workbook = sheet.getWorkbook();
//			cellStyleFormatNumber = workbook.createCellStyle();
//			cellStyleFormatNumber.setDataFormat(format);
//		}
//		Row row = sheet.createRow(rowIndex);
//		Cell cell = row.createCell(5, CellType.FORMULA);
//
//		cell.setCellFormula("SUM(F2:F" + rowIndex + ")");
//		cell.setCellStyle(cellStyleFormatNumber);
//	}

	// Auto resize column width
	private static void autosizeColumn(Sheet sheet, int lastColumn) {

		for (int columnIndex = 0; columnIndex < lastColumn; columnIndex++) {
			sheet.autoSizeColumn(columnIndex);
		}

	}

	private static List<String> convertHeader(List<String> headers) {
		List<String> convertheaders = new ArrayList<String>();
		for (String header : headers) {
			convertheaders.add(converString(header));
		}
		return convertheaders;

	}
	private static String converString(String header) {
		String convertHeader  = header.substring(0, 1).toUpperCase();
		for(int i=1;i<header.length();i++) {
			
			if(header.charAt(i)>=65 && header.charAt(i) <=90) {
				convertHeader += " "+header.charAt(i);
			}else {
				convertHeader += header.charAt(i);
			}
		}
		return convertHeader;
	}

}
