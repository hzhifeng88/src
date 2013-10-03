import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class validateTextStyle extends commonValidate {

	private Sheet textSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0,1,2};
	private ArrayList<String> storeErrorRow = new ArrayList<String>();
	private ArrayList<String> storeErrorMsg = new ArrayList<String>();
	
	public validateTextStyle(Sheet sheet, Workbook originalWorkbook, ArrayList<String> colorList ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, colorList, kit, doc);
		this.textSheet = sheet;
	}
	
	public boolean isSheetCorrect() {

		return sheetCorrect;
	}
	
	public void validateSheet(ArrayList<String> storeTextGeometry) {
		
		int tempColumnIndex = -1;
		
		// There are no formatting errors
		if(validateFormat() == false) {
			
			if(textSheet.getLastRowNum() == 3) {
				printNoErrorMsg();
				sheetCorrect = true;
			}else {
				
				for (int rowIndex = 4; rowIndex <= textSheet.getLastRowNum(); rowIndex++) {

					Row row = textSheet.getRow(rowIndex);

					// Check missing attribute
					checkMandatoryAttributes(row, rowIndex, mandatoryColumn);
					
					// Check for line break in cell
					checkLineBreakInCells(row);
					
					// Check valid ID and duplicate
					if(row.getCell(0) != null && row.getCell(0).getCellType() != Cell.CELL_TYPE_BLANK) {

						tempColumnIndex = findColumnIndex("Style ID", 1);
						String columnLetter = columnIndexToLetter(tempColumnIndex);
						checkIDAndDuplicate('T', columnLetter, rowIndex, tempColumnIndex);
					}
					
					// Check color valid
					if(row.getCell(3) != null && row.getCell(3).getCellType() != Cell.CELL_TYPE_BLANK) {
						String haloColorRadius = row.getCell(3).toString();
						String tempHaloArray[] = haloColorRadius.split(",");
						matchColor(tempHaloArray[0], "D", rowIndex);
					}
					
					// Check color valid
					if(row.getCell(11) != null && row.getCell(11).getCellType() != Cell.CELL_TYPE_BLANK) {
						matchColor(row.getCell(11).toString(), "L", rowIndex);
					}
					
					checkLabelPlacement(row, rowIndex, tempColumnIndex, storeTextGeometry);
				}
				if(hasError == true) {
					printColorError();
					printValueError();
					sheetCorrect = false;
				}else {
					printNoErrorMsg();
					sheetCorrect = true;
				}
			}
		}
	}
	
	public void checkLabelPlacement(Row row, int rowIndex, int tempColumnIndex, ArrayList<String> storeTextGeometry) {
		
		for(int count = 0; count < storeTextGeometry.size(); count++) {
			
			if(row.getCell(tempColumnIndex).toString().equalsIgnoreCase(storeTextGeometry.get(count).substring(1).toString())) {
				
				char labelPlacement = storeTextGeometry.get(count).toString().charAt(0);
				checkNow(row, rowIndex, labelPlacement);
			}
		}
	}
	
	public void checkNow(Row row, int rowIndex, char labelPlacement) {
		
		if(labelPlacement == 'P') {
			
			if(((row.getCell(5) == null || row.getCell(5).getCellType() == Cell.CELL_TYPE_BLANK) || (row.getCell(6) == null || row.getCell(6).getCellType() == Cell.CELL_TYPE_BLANK) || (row.getCell(7) == null || row.getCell(7).getCellType() == Cell.CELL_TYPE_BLANK)) || ((row.getCell(8) != null && row.getCell(8).getCellType() != Cell.CELL_TYPE_BLANK) || (row.getCell(9) != null && row.getCell(9).getCellType() != Cell.CELL_TYPE_BLANK) || (row.getCell(10) != null && row.getCell(10).getCellType() != Cell.CELL_TYPE_BLANK))) {
				hasError = true;
				storeErrorMsg.add("Error! Only column F, G and H must be filled in for point label placement.");
				storeErrorRow.add(Integer.toString(rowIndex + 1));
			}
		}else if(labelPlacement == 'L') {
			
			if(((row.getCell(5) != null && row.getCell(5).getCellType() != Cell.CELL_TYPE_BLANK) || (row.getCell(6) != null && row.getCell(6).getCellType() != Cell.CELL_TYPE_BLANK) || (row.getCell(7) != null && row.getCell(7).getCellType() != Cell.CELL_TYPE_BLANK)) || ((row.getCell(8) == null || row.getCell(8).getCellType() == Cell.CELL_TYPE_BLANK) || (row.getCell(9) == null || row.getCell(9).getCellType() == Cell.CELL_TYPE_BLANK) || (row.getCell(10) == null || row.getCell(10).getCellType() == Cell.CELL_TYPE_BLANK))) {
				
				hasError = true;
				storeErrorMsg.add("Error! Only column I, J and K must be filled in for line label placement.");
				storeErrorRow.add(Integer.toString(rowIndex + 1));
			}
		}
	}
}
