package poiTest;

import java.io.FileOutputStream;
import java.util.Calendar;
import java.util.Date;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.WorkbookUtil;
import org.junit.Test;

/**
 * Dimit 2015年1月10日
 */
public class PoiTestCase {
	public static final String PATH = "F:/poiExcelWork/";
	public static final String FILENAME = "poiWorkBook.xls";
	public static final String FULLPATH = PATH + FILENAME;
	public static final String SHEETNAME = "default";

	/**
	 * 创建文件方法
	 */
	@Test
	public void createWorkBook() {
		Workbook wb = new HSSFWorkbook();
		try {
			FileOutputStream fileOut = new FileOutputStream(FULLPATH);
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
		}
	}

	/**
	 * 创建表单
	 */
	@Test
	@SuppressWarnings("unused")
	public void createSheet() {
		Workbook wb = new HSSFWorkbook(); // or new XSSFWorkbook();
		Sheet sheet1 = wb.createSheet("one sheet");
		Sheet sheet2 = wb.createSheet("second sheet");
		String safeName = WorkbookUtil
				.createSafeSheetName("[O'Brien's sales*?]");
		Sheet sheet3 = wb.createSheet(safeName);
		String safeName1 = WorkbookUtil.createSafeSheetName("[one sheet*?]");
		Sheet sheet4 = wb.createSheet(safeName1);

		try {
			FileOutputStream fileOut = new FileOutputStream(FULLPATH);
			wb.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	/**
	 * 创建单元格
	 */
	@Test
	public void createCell() {
		Workbook wb = new HSSFWorkbook();
		CreationHelper createHelper = wb.getCreationHelper();
		Sheet sheet = null;
		Row row = null;
		try {
			sheet = wb.createSheet();
			row = sheet.createRow((short) 0);
		} catch (Exception e) {
			e.printStackTrace();
		}
		Cell cell = row.createCell(0);
		cell.setCellValue(1);
		row.createCell(1).setCellValue(1.2);
		row.createCell(2).setCellValue(
				createHelper.createRichTextString("This is a string"));
		row.createCell(3).setCellValue(true);
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(FILENAME);
			wb.write(fileOut);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				fileOut.close();
			} catch (Exception e2) {
				e2.printStackTrace();
			}
		}
	}

	/**
	 * 创建时间单元格
	 */
	@Test
	public void creatDateCell() {
		Workbook wb = new HSSFWorkbook();
		CreationHelper createHelper = wb.getCreationHelper();
		Sheet sheet = wb.createSheet("new sheet");
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		cell.setCellValue(new Date());
		CellStyle cellStyle = wb.createCellStyle();
		cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(
				"m/d/yy h:mm"));
		cell = row.createCell(1);
		cell.setCellValue(new Date());
		cell.setCellStyle(cellStyle);
		cell = row.createCell(2);
		cell.setCellValue(Calendar.getInstance());
		cell.setCellStyle(cellStyle);
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(FILENAME);
			wb.write(fileOut);
		} catch (Exception e) {
		} finally {
			try {
				fileOut.close();
			} catch (Exception e2) {
			}
		}
	}

	/**
	 * 创建不同类型的单元格
	 */
	@Test
	public void creatDiffCell() {
		Workbook wb = new HSSFWorkbook();
		Sheet sheet = wb.createSheet("new sheet");
		Row row = sheet.createRow((short) 2);
		row.createCell(0).setCellValue(1.1);
		row.createCell(1).setCellValue(new Date());
		row.createCell(2).setCellValue(Calendar.getInstance());
		row.createCell(3).setCellValue("a string");
		row.createCell(4).setCellValue(true);
		row.createCell(5).setCellType(Cell.CELL_TYPE_ERROR);
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(FILENAME);
			wb.write(fileOut);
		} catch (Exception e) {
		} finally {
			try {
				fileOut.close();
			} catch (Exception e2) {
			}
		}
	}
}
