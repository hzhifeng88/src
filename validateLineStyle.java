import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class validateLineStyle extends commonValidate {

	private Workbook workbook;
	private Sheet lineSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0};

	public validateLineStyle(Sheet sheet, Workbook workbook, Workbook originalWorkbook, ArrayList<String> colorList ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, colorList, kit, doc);
		this.lineSheet = sheet;
		this.workbook = workbook;
	}
	
	public boolean isSheetCorrect() {
		return sheetCorrect;
	}

	public void validateSheet() {

		int rowIndex = 0;
		Row row = null;

		// There are no formatting errors
		if(validateFormat() == false) {

			if(lineSheet.getLastRowNum() == 3) {
				printNoErrorMsg();
				sheetCorrect = true;
			}else {
				
				for (rowIndex = 4; rowIndex <= lineSheet.getLastRowNum(); rowIndex++) {

					row = lineSheet.getRow(rowIndex);

					// Check missing attribute
					checkMandatoryAttributes(row, rowIndex, mandatoryColumn);

					// Check for line break in cell
					checkLineBreakInCells(row);
					
					// Check valid ID and duplicate
					if(row.getCell(0) != null && row.getCell(0).getCellType() != Cell.CELL_TYPE_BLANK) {
						int tempColumnIndex = findColumnIndex("Style ID", 1);
						String columnLetter = columnIndexToLetter(tempColumnIndex);
						checkIDAndDuplicate('L', columnLetter, rowIndex, tempColumnIndex);
					}

					// Check size
					checkSize(row, 5, "F");
					
					// Check opacity
					checkOpacity(row, 2, "C");
					
					// Check color valid
					if(row.getCell(1) != null && row.getCell(1).getCellType() != Cell.CELL_TYPE_BLANK) {
						matchColor(row.getCell(1).toString(), "B", rowIndex);
					}
					
					// Check line join
					if(row.getCell(6) != null && row.getCell(6).getCellType() != Cell.CELL_TYPE_BLANK) {
						checkPencilLineJoin(row, rowIndex, 6, "G");
					}
					
					// Check line cap
					if(row.getCell(7) != null && row.getCell(7).getCellType() != Cell.CELL_TYPE_BLANK) {
						checkPencilLineCap(row, rowIndex, 7, "H");
					}
					
					// Check point reference
					if(row.getCell(8) != null && row.getCell(8).getCellType() != Cell.CELL_TYPE_BLANK) {
						referenceCheck(workbook, lineSheet, row, rowIndex, 8, 'P');
					}
				}
				if(hasError == true) {
					printValueError();
					printPencilBasedError();
					printColorError();
					printReferenceError();
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
