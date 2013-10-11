import java.util.*;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class ValidatePointStyle extends CommonValidate {

	private Sheet pointSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0,1,2};

	public ValidatePointStyle(Sheet sheet, Workbook originalWorkbook, List<String> colorList ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, colorList, kit, doc);
		this.pointSheet = sheet;
	}

	public boolean isSheetCorrect() {
		return sheetCorrect;
	}

	public void validateSheet() {

		if(hasFormatErrors()) {
			return;
		}

		if(pointSheet.getLastRowNum() == 3) {
			printNoErrorMsg();
			sheetCorrect = true;
		}else {

			for (int rowIndex = 4; rowIndex <= pointSheet.getLastRowNum(); rowIndex++) {

				Row row = pointSheet.getRow(rowIndex);

				checkMandatoryAttributes(row, mandatoryColumn);
				checkLineBreakInCells(row);
				checkIDAndDuplicate(row, 'P', "A", 0);
				checkSizePositive(row, 1, "B");
				checkSizePositive(row, 11, "L");
				checkOpacity(row, 8, "I");
				checkOpacity(row, 15, "P");
				matchColor(row, 7, "H");
				matchColor(row, 14, "O");
				checkPencilLineJoin(row, 12, "M");
				checkPencilLineCap(row, 13, "N");
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
