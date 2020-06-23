import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.HashMap;
// import java.util.Iterator;
// import java.util.Map;
// import java.util.Map.Entry;

import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;



public class CUS_HistoryDataInitialClass_mxJPO {

	/**
	 * check if the file is excel or empty
	 * 
	 * @param filePath
	 * @return
	 */
	private boolean validateExcel(String excelPath) throws Exception {

		// check path
//		if (UIUtil.isNullOrEmpty(excelPath)) {
//
//			System.out.println("=========ERROR LOG========= File path is empty!");
//			return Boolean.FALSE;
//		}

		// check file
		File excelFile = new File(excelPath);
		if (!excelFile.isFile()) {

			System.out.println("=========ERROR LOG========= File error in path!");
			return Boolean.FALSE;
		}

		// check file extension
		String excelName = excelFile.getName().toLowerCase();
		if ((excelName.indexOf(".xls") == -1) && (excelName.indexOf(".xlsx") == -1)) {

			System.out.println("=========ERROR LOG========= File is not excel!");
			return Boolean.FALSE;
		}

		return Boolean.TRUE;
	}

	/**
	 * get cell values
	 * 
	 * @param cell
	 * @return
	 */
	private String getCellValue(XSSFCell cell) {

		String value = "";
		switch (cell.getCellType()) {

		case 1:
			value = cell.getRichStringCellValue().getString();
			break;
		case 0:
			if (DateUtil.isCellDateFormatted(cell)) {

				value = cell.getDateCellValue().toString();
			} else {

				value = String.valueOf(cell.getNumericCellValue());
			}
			break;
		case 4:
			value = String.valueOf(cell.getBooleanCellValue());
			break;
		case 2:
			value = cell.getCellFormula();
			break;
		case 3:
			break;
		default:
			break;
		}

		return value.trim();
	}

	/**
	 * get excel content
	 * 
	 * @return
	 * @throws Exception
	 */
	private HashMap<String, String> getExcelContent(String excelPath) throws Exception {

		HashMap<String, String> result = new HashMap<String, String>();

		// read excel contents
		File excelFile = new File(excelPath);
		InputStream excelIs = new FileInputStream(excelFile);
		XSSFWorkbook workBook = new XSSFWorkbook(excelIs);
		XSSFSheet sheet = workBook.getSheetAt(0);
		int rowNum = sheet.getLastRowNum();
		for (int i = 1; i <= rowNum; i++) {

			XSSFRow row = sheet.getRow(i);
			StringBuilder cellContent = new StringBuilder();
			int cellNum = row.getLastCellNum();
			for (int j = 0; j < cellNum; j++) {

				if (j != 0) {

					cellContent.append(",");
				}

				cellContent.append(getCellValue(row.getCell(j)));
			}

			if (cellContent.length() > 0 && (cellContent.toString() != null)
					&& !"null".equalsIgnoreCase(cellContent.toString())) {

				result.put(String.valueOf(i), cellContent.toString());
			}
		}

		return result;
	}

	/**
	 * get digital serial number
	 * 
	 * @param strSerialNum
	 * @return
	 * @throws Exception
	 */
	private int getDigitalSerialNum(String strSerialNum) throws Exception {

		int digitalSerialNum = 0;
		if (strSerialNum.length() == 2) {

			if (!strSerialNum.matches("[0-9]+")) {

				// string to number
				char charFirst = strSerialNum.charAt(0);
				char charSecond = strSerialNum.charAt(1);
				int differ = charFirst - 55;
				digitalSerialNum = differ * 10;
				// add 1 create the next serial number
				digitalSerialNum = digitalSerialNum + Integer.valueOf(String.valueOf(charSecond));
			} else {

				digitalSerialNum = Integer.valueOf(strSerialNum);
			}
		} else {

			digitalSerialNum = Integer.valueOf(strSerialNum);
		}

		return digitalSerialNum;
	}

	
}
