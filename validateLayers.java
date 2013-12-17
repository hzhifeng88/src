import java.util.*;
import org.apache.poi.ss.usermodel.*;

public class ValidateLayers extends CommonValidate {

	private Workbook workbook;
	private Sheet layersSheet;
	private boolean sheetCorrect = false;
	private static int[] mandatoryColumn = {0,2,4,6};

	public ValidateLayers(Sheet sheet, Workbook workbook, Workbook templateWorkbook, List<String> colorList, String validateMessage) {

		super(sheet, templateWorkbook, colorList, validateMessage);
		this.layersSheet = sheet;
		this.workbook = workbook;
	}

	public boolean isSheetCorrect() {
		return sheetCorrect;
	}

	public String getMessage() {
		return validateMessage;
	}
	
	public List<String> validateSheet() {

		if(hasFormatErrors()) {
			return storeTextGeometry;
		}
		
		if(layersSheet.getLastRowNum() == 3) {
			printNoErrorMsg();
			sheetCorrect = true;
		}else {

			for (int rowIndex = 4; rowIndex <= layersSheet.getLastRowNum(); rowIndex++) {

				Row row = layersSheet.getRow(rowIndex);

				checkMandatoryAttributes(row, mandatoryColumn);
				checkLineBreakInCells(row);
				checkClassValues(row);
			}
			
			referenceCheck();
			
			if(hasError) {
				printValueError();
				sheetCorrect = false;
			}else {
				printNoErrorMsg();
				sheetCorrect = true;
			}
		}
		return storeTextGeometry;
	}

	public void referenceCheck() {

		for (int rowIndex = 4; rowIndex <= layersSheet.getLastRowNum(); rowIndex++) {

			Row row = layersSheet.getRow(rowIndex);
			int styleIDIndex = findColumnIndex("Style ID", 1);
			String cellLocation = columnIndexToLetter(styleIDIndex) + Integer.toString(rowIndex + 1);
			String referenceStyleID = row.getCell(styleIDIndex).toString();

			getReferenceStyle(workbook, layersSheet, referenceStyleID, cellLocation, row.getCell(4).toString());
		}
	}
}
