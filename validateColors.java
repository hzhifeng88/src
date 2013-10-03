import java.io.IOException;
import java.util.ArrayList;
import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class validateColors extends commonValidate {

	private HTMLEditorKit kit;
	private HTMLDocument doc;
	private Sheet colorsSheet;
	private boolean sheetCorrect = false;
	private static String hexadecimal = "0123456789abcdefABCDEF";
	private ArrayList<String> storeInvalidColorRGB = new ArrayList<String>();
	private static int[] mandatoryColumn = {0,1};

	public validateColors(Sheet sheet, Workbook originalWorkbook, ArrayList<String> colorList ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, colorList, kit, doc);
		this.colorsSheet = sheet;
		this.kit = kit;
		this.doc = doc;
	}

	public boolean isSheetCorrect() {

		return sheetCorrect;
	}

	public void validateSheet() {

		// There are no formatting errors
		if(validateFormat() == false) {

			if(colorsSheet.getLastRowNum() == 3) {
				printNoErrorMsg();
				sheetCorrect = true;
			}else {
				for (int rowIndex = 4; rowIndex <= colorsSheet.getLastRowNum(); rowIndex++) {

					Row row = colorsSheet.getRow(rowIndex);

					// Check missing attribute
					checkMandatoryAttributes(row, rowIndex, mandatoryColumn);

					// Check for line break in cell
					checkLineBreakInCells(row);

					// Check valid ID and duplicate
					if(row.getCell(0) != null && row.getCell(0).getCellType() != Cell.CELL_TYPE_BLANK) {

						int tempColumnIndex = findColumnIndex("Color Id", 2);
						String columnLetter = columnIndexToLetter(tempColumnIndex);
						checkIDAndDuplicate('C', columnLetter, rowIndex, tempColumnIndex);
					}
					checkRGB(rowIndex);
				}

				if(hasError == true) {
					printValueError();
					printRGBError();
					sheetCorrect = false;
				}else {
					printNoErrorMsg();
					sheetCorrect = true;

				}
			}
		}else {
			sheetCorrect = false;
		}
	}

	public void checkRGB(int rowIndex) {

		int tempColumnIndex = findColumnIndex("sRGB", 2);
		String columnLetter = columnIndexToLetter(tempColumnIndex);

		Row row = colorsSheet.getRow(rowIndex);

		if(tempColumnIndex != -1 && row.getCell(tempColumnIndex) != null && row.getCell(tempColumnIndex).getCellType() != Cell.CELL_TYPE_BLANK ){

			String tempString = row.getCell(tempColumnIndex).toString();

			if (!tempString.equalsIgnoreCase("")) {

				if (checkIsRGB(tempString) == false) {
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

			if (storeInvalidColorRGB.isEmpty() == false) {
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> sRGB is invalid (Rule 1: Begins with '#', Rule 2: 6 hexadecimal representation)</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeInvalidColorRGB + "</font color></font>", 0, 0, null);
			}
		} catch (BadLocationException | IOException e) {
			e.printStackTrace();
		}
	}
}
