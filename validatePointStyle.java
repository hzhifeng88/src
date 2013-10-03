import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class validatePointStyle extends commonValidate {

	private Sheet pointSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0,1,2};

	public validatePointStyle(Sheet sheet, Workbook originalWorkbook, ArrayList<String> colorList ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, colorList, kit, doc);
		this.pointSheet = sheet;
	}
	
	public boolean isSheetCorrect() {

		return sheetCorrect;
	}

	public void validateSheet() {

		// There are no formatting errors
		if(validateFormat() == false) {

			if(pointSheet.getLastRowNum() == 3) {
				printNoErrorMsg();
				sheetCorrect = true;
			}else {
				
				for (int rowIndex = 4; rowIndex <= pointSheet.getLastRowNum(); rowIndex++) {

					Row row = pointSheet.getRow(rowIndex);

					// Check missing attribute
					checkMandatoryAttributes(row, rowIndex, mandatoryColumn);

					// Check for line break in cell
					checkLineBreakInCells(row);
					
					// Check valid ID and duplicate
					if(row.getCell(0) != null && row.getCell(0).getCellType() != Cell.CELL_TYPE_BLANK) {
						
						int tempColumnIndex = findColumnIndex("Style ID", 1);
						String columnLetter = columnIndexToLetter(tempColumnIndex);
						checkIDAndDuplicate('P', columnLetter, rowIndex, tempColumnIndex);
					}
				
					// Check color valid
					if(row.getCell(7) != null && row.getCell(7).getCellType() != Cell.CELL_TYPE_BLANK) {
						matchColor(row.getCell(7).toString(), "H", rowIndex);
					}
					if(row.getCell(14) != null && row.getCell(14).getCellType() != Cell.CELL_TYPE_BLANK) {
						matchColor(row.getCell(14).toString(), "O", rowIndex);
					}

					// Check line join
					if(row.getCell(12) != null && row.getCell(12).getCellType() != Cell.CELL_TYPE_BLANK) {
						checkPencilLineJoin(row, rowIndex, 12, "M");
					}
					
					// Check line cap
					if(row.getCell(13) != null && row.getCell(13).getCellType() != Cell.CELL_TYPE_BLANK) {
						checkPencilLineCap(row, rowIndex, 13, "N");
					}
				}
				
				if(hasError == true) {
					printValueError();
					printPencilBasedError();
					printColorError();
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
