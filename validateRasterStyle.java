import java.util.*;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class ValidateRasterStyle extends CommonValidate {

	private Sheet rasterSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0};

	public ValidateRasterStyle(Sheet sheet, Workbook templateWorkbook, List<String> colorList ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, templateWorkbook, colorList, kit, doc);
		this.rasterSheet = sheet;
	}

	public boolean isSheetCorrect() {
		return sheetCorrect;
	}

	public void validateSheet() {

		if(hasFormatErrors()) {
			return;
		}

		if(rasterSheet.getLastRowNum() == 3) {
			printNoErrorMsg();
			sheetCorrect = true;
		}else {

			for (int rowIndex = 4; rowIndex <= rasterSheet.getLastRowNum(); rowIndex++) {

				Row row = rasterSheet.getRow(rowIndex);

				checkMandatoryAttributes(row, mandatoryColumn);
				checkLineBreakInCells(row);
				checkIDAndDuplicate(row, 'R', "A", 0);
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
