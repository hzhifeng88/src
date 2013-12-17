import java.util.*;
import org.apache.poi.ss.usermodel.*;

public class ValidateColors extends CommonValidate {

	private Sheet colorsSheet;
	private boolean sheetCorrect = false;
	private static String hexadecimal = "0123456789abcdefABCDEF";
	private List<String> storeInvalidColorRGB = new ArrayList<String>();
	private static int[] mandatoryColumn = {0,1};

	public ValidateColors(Sheet sheet, Workbook templateWorkbook, List<String> colorList, String validateMessage) {

		super(sheet, templateWorkbook, colorList, validateMessage);
		this.colorsSheet = sheet;
	}

	public boolean isSheetCorrect() {
		return sheetCorrect;
	}

	public String getMessage() {
		return validateMessage;
	}
	
	public void validateSheet() {

		if(hasFormatErrors()) {
			return;
		}

		if(colorsSheet.getLastRowNum() == 3) {
			printNoErrorMsg();
			sheetCorrect = true;
		}else {
			for (int rowIndex = 4; rowIndex <= colorsSheet.getLastRowNum(); rowIndex++) {

				Row row = colorsSheet.getRow(rowIndex);

				checkMandatoryAttributes(row, mandatoryColumn);
				checkLineBreakInCells(row);
				checkIDAndDuplicate(row, 'C', "A", 0);
				checkRGB(rowIndex);
			}

			if(hasError) {
				printValueError();
				printRGBError();
				sheetCorrect = false;
			}else {
				printNoErrorMsg();
				sheetCorrect = true;
			}
		}
	}

	public void checkRGB(int rowIndex) {

		int tempColumnIndex = findColumnIndex("sRGB", 2);
		String columnLetter = columnIndexToLetter(tempColumnIndex);

		Row row = colorsSheet.getRow(rowIndex);

		if(tempColumnIndex != -1 && row.getCell(tempColumnIndex) != null && row.getCell(tempColumnIndex).getCellType() != Cell.CELL_TYPE_BLANK ){

			String tempString = row.getCell(tempColumnIndex).toString();

			if (!tempString.equalsIgnoreCase("")) {

				if (!checkIsRGB(tempString)) {
					hasError = true;
					storeInvalidColorRGB.add(columnLetter + Integer.toString(rowIndex + 1));
				}
			}
		}
	}

	public boolean checkIsRGB(String tempStringColor) {	

		if (tempStringColor.charAt(0) != '#'){
			return false;	
		}
		if (tempStringColor.length() != 7){
			return false;	
		}

		for (int stringIndex = 1; stringIndex < tempStringColor.length(); stringIndex ++) {

			if (hexadecimal.indexOf(tempStringColor.charAt(stringIndex)) == -1) 
				return false; 		
		}		
		return true;		
	}

	public void printRGBError(){
		
		if (!storeInvalidColorRGB.isEmpty()) {
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> sRGB is invalid (Rule 1: Begins with '#', Rule 2: 6 hexadecimal representation)<br></font color></font");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeInvalidColorRGB + "<br></font color></font>");
		}
	}
}
