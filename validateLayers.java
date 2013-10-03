import java.util.ArrayList;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class validateLayers extends commonValidate {

	private Workbook workbook;
	private Sheet layersSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0,2,3,4,6};

	public validateLayers(Sheet sheet, Workbook workbook, Workbook originalWorkbook, ArrayList<String> colorList ,  HTMLEditorKit kit, HTMLDocument doc) {

		super(sheet, originalWorkbook, colorList, kit, doc);
		this.layersSheet = sheet;
		this.workbook = workbook;
	}
	
	public boolean isSheetCorrect() {

		return sheetCorrect;
	}

	public ArrayList<String> validateSheet() {

		// There are no formatting errors
		if(validateFormat() == false) {

			if(layersSheet.getLastRowNum() == 3) {
				printNoErrorMsg();
				sheetCorrect = true;
			}else {

				for (int rowIndex = 4; rowIndex <= layersSheet.getLastRowNum(); rowIndex++) {

					Row row = layersSheet.getRow(rowIndex);

					// Check missing attributes
					checkMandatoryAttributes(row, rowIndex, mandatoryColumn);
					
					// Check for line break in cell
					checkLineBreakInCells(row);
				}
				
				if(hasError == true) {
					printValueError();
					sheetCorrect = false;
				}else {
					sheetCorrect = true;
				}
				
				referenceCheck();
			}
		}else {
			sheetCorrect = false;
		}
		return storeTextGeometry;
	}
	
	public void referenceCheck() {

		String cellLocation = "";
		String referenceStyleID = "";

		// Check for Reference error(s)
		for (int rowIndex = 4; rowIndex <= layersSheet.getLastRowNum(); rowIndex++) {

			Row row = layersSheet.getRow(rowIndex);
			int styleIDIndex = findColumnIndex("Style ID", 1);

			cellLocation = columnIndexToLetter(styleIDIndex) + Integer.toString(rowIndex + 1);
			referenceStyleID = row.getCell(styleIDIndex).toString();
			getReferenceStyle(workbook, layersSheet, referenceStyleID, cellLocation, row.getCell(4).toString());
		}

		if(hasError == true) {
			printReferenceError();
		}else {
			printNoErrorMsg();
		}
	}
}
