import java.util.*;
import org.apache.poi.ss.usermodel.*;

public class ValidateLineStyle extends CommonValidate {

	private Workbook workbook;
	private Sheet lineSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0};

	public ValidateLineStyle(Sheet sheet, Workbook workbook, Workbook templateWorkbook, List<String> colorList, String validateMessage) {

		super(sheet, templateWorkbook, colorList, validateMessage);
		this.lineSheet = sheet;
		this.workbook = workbook;
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

		if(lineSheet.getLastRowNum() == 3) {
			printNoErrorMsg();
			sheetCorrect = true;
		}else {

			for (int rowIndex = 4; rowIndex <= lineSheet.getLastRowNum(); rowIndex++) {

				Row row = lineSheet.getRow(rowIndex);

				checkMandatoryAttributes(row, mandatoryColumn);
				checkLineBreakInCells(row);
				checkIDAndDuplicate(row, 'L', "A", 0);
				checkSizePositive(row, 5, "F");
				checkOpacity(row, 2, "C");
				matchColor(row, 1, "B");
				checkPencilLineJoin(row, 6, "G");
				checkPencilLineCap(row, 7, "H");
				referenceCheck(workbook, lineSheet, row, 8, 'P');
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
}
