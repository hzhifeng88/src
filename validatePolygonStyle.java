import java.util.*;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class ValidatePolygonStyle extends CommonValidate {

	private Workbook workbook;
	private Sheet polygonSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0,5};

	public ValidatePolygonStyle(Sheet sheet, Workbook workbook, Workbook templateWorkbook, List<String> colorList ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, templateWorkbook, colorList, kit, doc);
		this.polygonSheet = sheet;
		this.workbook = workbook;
	}

	public boolean isSheetCorrect() {
		return sheetCorrect;
	}

	public void validateSheet() {

		int rowIndex = 0;
		Row row = null;

		if(hasFormatErrors()) {
			return;
		}

		if(polygonSheet.getLastRowNum() == 3) {
			printNoErrorMsg();
			sheetCorrect = true;
		}else {

			for (rowIndex = 4; rowIndex <= polygonSheet.getLastRowNum(); rowIndex++) {

				row = polygonSheet.getRow(rowIndex);

				checkMandatoryAttributes(row, mandatoryColumn);
				checkLineBreakInCells(row);
				checkIDAndDuplicate(row, 'A', "A", 0);
				checkOpacity(row, 3, "D");
				matchColor(row, 2, "C");
				referenceCheck(workbook, polygonSheet, row, 4, 'P');
				referenceCheck(workbook, polygonSheet, row, 5, 'L');
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
