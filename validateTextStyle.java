import java.util.*;
import org.apache.poi.ss.usermodel.*;

public class ValidateTextStyle extends CommonValidate {

	private Sheet textSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0,1,2};
	private List<String> storeErrorRow = new ArrayList<String>();
	private List<String> storeErrorMsg = new ArrayList<String>();

	public ValidateTextStyle(Sheet sheet, Workbook templateWorkbook, List<String> colorList, String validateMessage) {

		super(sheet, templateWorkbook, colorList, validateMessage);
		this.textSheet = sheet;
	}

	public boolean isSheetCorrect() {
		return sheetCorrect;
	}

	public String getMessage() {
		return validateMessage;
	}
	
	public void validateSheet(List<String> storeTextGeometry) {

		if(hasFormatErrors()) {
			return;
		}

		if(textSheet.getLastRowNum() == 3) {
			printNoErrorMsg();
			sheetCorrect = true;
		}else {

			for (int rowIndex = 4; rowIndex <= textSheet.getLastRowNum(); rowIndex++) {

				Row row = textSheet.getRow(rowIndex);

				checkMandatoryAttributes(row, mandatoryColumn);
				checkLineBreakInCells(row);
				checkIDAndDuplicate(row, 'T', "A", 0);
				checkSizePositive(row, 4, "E");
				checkOpacity(row, 13, "M");
				matchHaloColor(row, 3, "D");
				matchColor(row, 12, "L");
				checkLabelPlacement(row, rowIndex, 0, storeTextGeometry);
			}
			if(hasError) {
				printValueError();
				sheetCorrect = false;
			}else {
				printNoErrorMsg();
				sheetCorrect = true;
			}
		}
	}

	public void checkLabelPlacement(Row row, int rowIndex, int columnIndex, List<String> storeTextGeometry) {

		for(int count = 0; count < storeTextGeometry.size(); count++) {

			if(row.getCell(columnIndex).toString().equalsIgnoreCase(storeTextGeometry.get(count).substring(1).toString())) {
				char labelPlacement = storeTextGeometry.get(count).toString().charAt(0);
				checkNow(row, rowIndex, labelPlacement);
			}
		}
	}

	public void checkNow(Row row, int rowIndex, char labelPlacement) {

		if(labelPlacement == 'P') {

			if(((row.getCell(6) == null || row.getCell(6).getCellType() == Cell.CELL_TYPE_BLANK) || (row.getCell(7) == null || row.getCell(7).getCellType() == Cell.CELL_TYPE_BLANK) || (row.getCell(8) == null || row.getCell(8).getCellType() == Cell.CELL_TYPE_BLANK)) || ((row.getCell(9) != null && row.getCell(9).getCellType() != Cell.CELL_TYPE_BLANK) || (row.getCell(10) != null && row.getCell(10).getCellType() != Cell.CELL_TYPE_BLANK) || (row.getCell(11) != null && row.getCell(11).getCellType() != Cell.CELL_TYPE_BLANK))) {
				hasError = true;
				storeErrorMsg.add("Error! Only column G, H and I must be filled in for point label placement.");
				storeErrorRow.add(Integer.toString(rowIndex + 1));
			}
		}else if(labelPlacement == 'L') {

			if(((row.getCell(6) != null && row.getCell(6).getCellType() != Cell.CELL_TYPE_BLANK) || (row.getCell(7) != null && row.getCell(7).getCellType() != Cell.CELL_TYPE_BLANK) || (row.getCell(8) != null && row.getCell(8).getCellType() != Cell.CELL_TYPE_BLANK)) || ((row.getCell(9) == null || row.getCell(9).getCellType() == Cell.CELL_TYPE_BLANK) || (row.getCell(10) == null || row.getCell(10).getCellType() == Cell.CELL_TYPE_BLANK) || (row.getCell(11) == null || row.getCell(11).getCellType() == Cell.CELL_TYPE_BLANK))) {
				hasError = true;
				storeErrorMsg.add("Error! Only column J, K, L must be filled in for line label placement.");
				storeErrorRow.add(Integer.toString(rowIndex + 1));
			}
		}
	}
}
