import java.io.IOException;
import java.util.*;

import javax.swing.JOptionPane;
import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.*;

public class ValidateColors extends CommonValidate {

	private HTMLEditorKit kit;
	private HTMLDocument doc;
	private Sheet colorsSheet;
	private boolean sheetCorrect = false;
	private static String hexadecimal = "0123456789abcdefABCDEF";
	private List<String> storeInvalidColorRGB = new ArrayList<String>();
	private static int[] mandatoryColumn = {0,1};

	public ValidateColors(Sheet sheet, Workbook templateWorkbook, List<String> colorList ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, templateWorkbook, colorList, kit, doc);
		this.colorsSheet = sheet;
		this.kit = kit;
		this.doc = doc;
	}

	public boolean isSheetCorrect() {
		return sheetCorrect;
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

		try {

			if (!storeInvalidColorRGB.isEmpty()) {
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> sRGB is invalid (Rule 1: Begins with '#', Rule 2: 6 hexadecimal representation)</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeInvalidColorRGB + "</font color></font>", 0, 0, null);
			}
		} catch (BadLocationException | IOException e) {
			JOptionPane.showMessageDialog(null, "An error has occurred (ValidateColors-HTMLkit). Application will now terminate.");
			System.exit(0);
		}
	}
}
