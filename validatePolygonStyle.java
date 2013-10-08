import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class validatePolygonStyle extends commonValidate {

	private Workbook workbook;
	private Sheet polygonSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0,5};

	public validatePolygonStyle(Sheet sheet, Workbook workbook, Workbook originalWorkbook, ArrayList<String> colorList ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, colorList, kit, doc);
		this.polygonSheet = sheet;
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

			if(polygonSheet.getLastRowNum() == 3) {
				printNoErrorMsg();
				sheetCorrect = true;
			}else {

				for (rowIndex = 4; rowIndex <= polygonSheet.getLastRowNum(); rowIndex++) {

					row = polygonSheet.getRow(rowIndex);

					// Check missing attribute
					checkMandatoryAttributes(row, rowIndex, mandatoryColumn);

					// Check for line break in cell
					checkLineBreakInCells(row);
					
					// Check valid ID and duplicate
					if(row.getCell(0) != null && row.getCell(0).getCellType() != Cell.CELL_TYPE_BLANK) {

						int tempColumnIndex = findColumnIndex("Style ID", 1);
						String columnLetter = columnIndexToLetter(tempColumnIndex);
						checkIDAndDuplicate('A', columnLetter, rowIndex, tempColumnIndex);
					}

					// Check opacity
					checkOpacity(row, 3, "D");
					
					// Check color valid
					if(row.getCell(2) != null && row.getCell(2).getCellType() != Cell.CELL_TYPE_BLANK) {
						matchColor(row.getCell(2).toString(), "C", rowIndex);
					}
					
					// Check point reference
					if(row.getCell(4) != null && row.getCell(4).getCellType() != Cell.CELL_TYPE_BLANK) {
						referenceCheck(workbook, polygonSheet, row, rowIndex, 4, 'P');
					}
					
					// Check point reference
					if(row.getCell(5) != null && row.getCell(5).getCellType() != Cell.CELL_TYPE_BLANK) {
						referenceCheck(workbook, polygonSheet, row, rowIndex, 5, 'L');
					}
				}

				if(hasError == true) {
					printValueError();
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
