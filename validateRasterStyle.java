import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class validateRasterStyle extends commonValidate {

	private Sheet rasterSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0};
	
	public validateRasterStyle(Sheet sheet, Workbook originalWorkbook, ArrayList<String> colorList ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, colorList, kit, doc);
		this.rasterSheet = sheet;
	}
	
	public boolean isSheetCorrect() {

		return sheetCorrect;
	}
	
	public void validateSheet() {
		
		// There are no formatting errors
		if(validateFormat() == false) {
			
			if(rasterSheet.getLastRowNum() == 3) {
				printNoErrorMsg();
				sheetCorrect = true;
			}else {
				
				for (int rowIndex = 4; rowIndex <= rasterSheet.getLastRowNum(); rowIndex++) {

					Row row = rasterSheet.getRow(rowIndex);

					// Check missing attribute
					checkMandatoryAttributes(row, rowIndex, mandatoryColumn);
					
					// Check for line break in cell
					checkLineBreakInCells(row);
					
					if(hasError == false) {
						
						// Check valid ID and duplicate
						int tempColumnIndex = findColumnIndex("Style ID", 1);
						String columnLetter = columnIndexToLetter(tempColumnIndex);
						checkIDAndDuplicate('R', columnLetter, rowIndex, tempColumnIndex);
					}
				}
				if(hasError == true) {
					printValueError();
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
}
