import java.util.*;
import org.apache.poi.ss.usermodel.*;

public class CommonValidate {

	private char idAlphabet;
	public boolean hasError = false;
	public String validateMessage;
	private Sheet sheet;
	private Workbook templateWorkbook;
	private List<String> storeColorID = new ArrayList<String>();					//for cross referencing
	private List<String> storeModifiedHeaderCells = new ArrayList<String>();
	private List<String> storeMergedCells = new ArrayList<String>();
	private List<String> storeEmptyRows = new ArrayList<String>();
	private List<String> storeLineBreakCells = new ArrayList<String>();
	private List<String> storeRightID = new ArrayList<String>();					//for checking duplicates
	private List<String> storeWrongID = new ArrayList<String>();
	private List<String> storeDuplicateID = new ArrayList<String>();
	private List<String> storeMissingValueCells = new ArrayList<String>();
	private List<String> storeInvalidColorCells = new ArrayList<String>();
	private List<String> storeReferenceErrorCells = new ArrayList<String>();
	private List<String> storeErrorRow = new ArrayList<String>();
	private List<String> storeErrorMsg = new ArrayList<String>();
	private List<String> storePencilLineJoin = new ArrayList<String>();
	private List<String> storePencilLineCap = new ArrayList<String>();
	public List<String> storeTextGeometry = new ArrayList<String>();
	private List<String> storeErrorSize = new ArrayList<String>();
	private List<String> storeOpacityError = new ArrayList<String>();
	private List<String> storeClassValueError = new ArrayList<String>();

	public CommonValidate(Sheet sheet, Workbook templateWorkbook, List<String> colorList, String validateMessage){

		this.sheet = sheet;	
		this.templateWorkbook = templateWorkbook;
		this.validateMessage = validateMessage;
		storeColorID = colorList;
	}

	public boolean hasFormatErrors() {

		checkModifiedHeader();
		checkMergedCells();
		checkEmptyRows();

		return printFormatError();
	}

	public int findColumnIndex(String columnName, int rowIndex) {

		Row tempRow = sheet.getRow(rowIndex);

		for(int countColumn = 0; countColumn < tempRow.getLastCellNum(); countColumn++) {

			if(columnName.equalsIgnoreCase(tempRow.getCell(countColumn).toString())) {
				return countColumn;
			}
		}
		return -1;
	}

	public String columnIndexToLetter(int columnIndex) {

		int base = 26;   
		StringBuffer b = new StringBuffer(); 

		do {  
			int digit = columnIndex % base + 65;  
			b.append(Character.valueOf((char) digit));  
			columnIndex = (columnIndex / base) - 1; 

		} while (columnIndex >= 0);   

		return b.reverse().toString();
	}

	public void checkModifiedHeader() {

		Sheet templateSheet = templateWorkbook.getSheet(sheet.getSheetName());

		for(int rowIndex = 0; rowIndex < 4; rowIndex++){

			Row row = sheet.getRow(rowIndex);
			Row templateRow = templateSheet.getRow(rowIndex);

			for(int columnIndex = 0; columnIndex < row.getLastCellNum(); columnIndex++){

				if(row.getCell(columnIndex) == null && templateRow.getCell(columnIndex) == null){
					continue;
				}

				if(row.getCell(columnIndex) != null && templateRow.getCell(columnIndex) == null){
					storeModifiedHeaderCells.add(columnIndexToLetter(columnIndex) + Integer.toString(rowIndex + 1));
					continue;
				}

				if(row.getCell(columnIndex).toString().equalsIgnoreCase(templateRow.getCell(columnIndex).toString())){
					continue;
				}else {
					hasError = true;
					storeModifiedHeaderCells.add(columnIndexToLetter(columnIndex) + Integer.toString(rowIndex + 1));
				}
			}
		}
	}

	public void checkMergedCells() {

		String cellNumber;

		for (int count = 0; count < sheet.getNumMergedRegions(); count++) {

			cellNumber = "";

			String tempString = sheet.getMergedRegion(count).toString().substring(41);
			StringTokenizer tokenizer = new StringTokenizer(tempString, ":");

			String cell = tokenizer.nextToken();

			for (int count1 = 0; count1 < cell.length(); count1++) {

				char checkChar = cell.charAt(count1);

				if (Character.isDigit(checkChar)) {
					cellNumber = cellNumber.concat(String.valueOf(checkChar));
				}
			}

			// Begin check from row 5 onwards
			if (Integer.parseInt(cellNumber) > 4) {
				hasError = true;
				storeMergedCells.add(cell);
			}
		}
	}

	public void checkEmptyRows() {

		boolean isRowEmpty = false;

		for (int rowCount = 4; rowCount <= sheet.getLastRowNum(); rowCount++) {

			isRowEmpty = false;
			Row row = sheet.getRow(rowCount);

			if (row == null) {
				hasError = true;
				storeEmptyRows.add(Integer.toString(rowCount + 1));
				continue;
			}

			// Check if all cells are empty
			for (int cellCount = 0; cellCount < row.getLastCellNum(); cellCount++) {

				if (row.getCell(cellCount) == null || row.getCell(cellCount).toString().trim().equals("")) {
					isRowEmpty = true;
				} else {
					isRowEmpty = false;
					break;
				}
			}
			if (isRowEmpty) {
				hasError = true;
				storeEmptyRows.add(Integer.toString(rowCount + 1));
			}
		}
	}

	public void checkLineBreakInCells(Row row) {

		for(int cellCount = 0; cellCount < row.getLastCellNum(); cellCount++) {

			if(row.getCell(cellCount) != null && row.getCell(cellCount).getCellType() != Cell.CELL_TYPE_BLANK && row.getCell(cellCount).toString().contains("\n")){
				hasError = true;
				String columnLetter = columnIndexToLetter(cellCount);
				storeLineBreakCells.add(columnLetter + Integer.toString(row.getRowNum() + 1));
			}
		}
	}

	public void checkIDAndDuplicate(Row row, char idAlphabet, String columnLetter, int columnIndex){

		boolean wrongStyleID = false;
		this.idAlphabet = idAlphabet;

		if(row.getCell(columnIndex) != null && row.getCell(columnIndex).getCellType() != Cell.CELL_TYPE_BLANK) {

			String styleID = row.getCell(columnIndex).toString();
			char firstChar = styleID.charAt(0);

			if (firstChar != idAlphabet) {
				wrongStyleID = true;
				hasError = true;
				storeWrongID.add(columnIndex + Integer.toString(row.getRowNum() + 1));
			}

			if (!wrongStyleID) {

				if (storeRightID.isEmpty()) {
					storeRightID.add(styleID);
				} else {
					if(storeRightID.contains(styleID)){
						hasError = true;
						storeDuplicateID.add(columnIndex + Integer.toString(row.getRowNum() + 1));
					}else{
						storeRightID.add(styleID);
					}
				}
			}
		}
	}

	public void checkMandatoryAttributes(Row row, int[] columnArray) {

		for(int columnCount = 0; columnCount < columnArray.length; columnCount++) {

			if(row.getCell(columnArray[columnCount]) == null || row.getCell(columnArray[columnCount]).toString().equalsIgnoreCase("")){
				hasError = true;
				storeMissingValueCells.add(columnIndexToLetter(columnArray[columnCount]) + Integer.toString(row.getRowNum() + 1));
			}
		}
	}

	public void checkSizePositive(Row row, int columnIndex, String columnLetter) {

		if(row.getCell(columnIndex) != null && row.getCell(columnIndex).getCellType() != Cell.CELL_TYPE_BLANK) {
			if(Double.parseDouble(row.getCell(columnIndex).toString()) <= 0.0) {
				hasError = true;
				storeErrorSize.add(columnLetter + Integer.toString(row.getRowNum() + 1));
			}
		}
	}

	public void checkOpacity(Row row, int columnIndex, String columnLetter) {

		if(row.getCell(columnIndex) != null && row.getCell(columnIndex).getCellType() != Cell.CELL_TYPE_BLANK) {

			Double opacityDouble = Double.parseDouble(row.getCell(columnIndex).toString());
			if(opacityDouble < 0 || opacityDouble > 1.0) {
				hasError = true;
				storeOpacityError.add(columnLetter + Integer.toString(row.getRowNum() + 1));
			}
		}
	}

	public void matchColor(Row row, int columnIndex, String currentColumn) {

		if(row.getCell(columnIndex) != null && row.getCell(columnIndex).getCellType() != Cell.CELL_TYPE_BLANK) {

			if(storeColorID.contains(row.getCell(columnIndex).toString())){
				return;
			}else {
				hasError = true;
				storeInvalidColorCells.add(currentColumn + Integer.toString(row.getRowNum() + 1));
			}
		}
	}

	public void matchHaloColor(Row row, int columnIndex, String currentColumn) {

		if(row.getCell(columnIndex) != null && row.getCell(columnIndex).getCellType() != Cell.CELL_TYPE_BLANK) {

			String haloColorRadius = row.getCell(columnIndex).toString();
			String tempHaloArray[] = haloColorRadius.split(",");

			if(storeColorID.contains(tempHaloArray[0])){
				return;
			}
			hasError = true;
			storeInvalidColorCells.add(currentColumn + Integer.toString(row.getRowNum() + 1));
		}
	}

	public void checkPencilLineJoin(Row row, int columnIndex, String columnLetter) {

		final List<String> lineJoin = new ArrayList<String>() {{
			add("mitre");
			add("round");
			add("bevel");
		}};

		if(row.getCell(columnIndex) != null && row.getCell(columnIndex).getCellType() != Cell.CELL_TYPE_BLANK) {
			if(lineJoin.contains(row.getCell(columnIndex).toString())){
				return;
			}else {
				hasError = true;
				storePencilLineJoin.add(columnLetter + Integer.toString(row.getRowNum() + 1));
			}
		}
	}

	public void checkPencilLineCap(Row row, int columnIndex, String columnLetter) {

		final List<String> lineCap = new ArrayList<String>() {{
			add("butt");
			add("round");
			add("square");
		}};

		if(row.getCell(columnIndex) != null && row.getCell(columnIndex).getCellType() != Cell.CELL_TYPE_BLANK) {
			if(lineCap.contains(row.getCell(columnIndex).toString())){
				return;
			}else {
				hasError = true;
				storePencilLineCap.add(columnLetter + Integer.toString(row.getRowNum() + 1));
			}
		}
	}

	public void referenceCheck(Workbook workbook, Sheet currentSheet, Row row, int columnIndex, char referenceTo) {

		if(row.getCell(columnIndex) != null && row.getCell(columnIndex).getCellType() != Cell.CELL_TYPE_BLANK) {

			String cellLocation = columnIndexToLetter(columnIndex) + Integer.toString(row.getRowNum() + 1);
			String referenceStyleID = row.getCell(columnIndex).toString();

			char firstChar = referenceStyleID.charAt(0);

			if(firstChar == referenceTo) {
				getReferenceStyle(workbook, currentSheet, referenceStyleID, cellLocation, null);
			}else {
				hasError = true;
				storeReferenceErrorCells.add(cellLocation);
			}
		}
	}

	public void getReferenceStyle(Workbook workbook, Sheet currentSheet, String referenceStyleID, String cellLocation, String geometry) {

		char firstChar = referenceStyleID.charAt(0);

		switch(firstChar) {
		case 'P':
			checkReference(currentSheet, workbook.getSheet("PointStyle"),  referenceStyleID, cellLocation);
			break;
		case 'L':  
			checkReference(currentSheet, workbook.getSheet("LineStyle"), referenceStyleID, cellLocation);
			break;
		case 'A':  
			checkReference(currentSheet, workbook.getSheet("PolygonStyle"), referenceStyleID, cellLocation);
			break;
		case 'T':
			storeTextGeometry.add(geometry + referenceStyleID);
			checkReference(currentSheet, workbook.getSheet("TextStyle"),  referenceStyleID, cellLocation);
			break;  
		case 'R': 
			checkReference(currentSheet, workbook.getSheet("RasterStyle"),  referenceStyleID, cellLocation);
			break;
		default: 
			break;
		}
	}

	public void checkReference(Sheet currentSheet, Sheet referenceSheet, String referenceStyleID, String cell) {

		// Check all styleID for match
		for(int rowCount = 4; rowCount <= referenceSheet.getLastRowNum(); rowCount++) {

			Row getRow = referenceSheet.getRow(rowCount);

			if(getRow.getCell(0) != null && getRow.getCell(0).getCellType() != Cell.CELL_TYPE_BLANK) {
				if(getRow.getCell(0).toString().equalsIgnoreCase(referenceStyleID)) {
					return;
				}
			}
		}
		hasError = true;
		storeReferenceErrorCells.add(cell);
	}

	public void checkClassValues(Row row) {

		if(row.getCell(2) != null && row.getCell(2).getCellType() != Cell.CELL_TYPE_BLANK) {

			String getClassValues = row.getCell(2).toString();
			String allowedChar = "0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ-_";

			for(int charCount = 0; charCount < getClassValues.length(); charCount++) {

				char getEachChar = getClassValues.charAt(charCount);

				if(allowedChar.indexOf(getEachChar) == -1) {
					hasError = true;
					storeClassValueError.add("C" + Integer.toString(row.getRowNum() + 1));
					break;
				}
			}
		}
	}

	public boolean printFormatError() {


		if(sheet.getSheetName().equalsIgnoreCase("Layers")){
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Error(s) in sheet: <font color=#ED0E3F><b>" + sheet.getSheetName() + "<br></b></font color></font>");
		}else{
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4><br>Error(s) in sheet: <font color=#ED0E3F><b>" + sheet.getSheetName() + "<br></b></font color></font>");
		}
		validateMessage = validateMessage.concat("<font size = 3> <font color=#088542>---------------------------------------------- <br></font color></font>");

		if(!storeModifiedHeaderCells.isEmpty()){
			hasError = true;
			Collections.sort(storeModifiedHeaderCells);
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Header cells are modified! Please correct this and try again.<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeModifiedHeaderCells + "<br></font color></font>");
		}

		if(!storeMergedCells.isEmpty()) {
			hasError = true;
			Collections.sort(storeMergedCells);
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>Merged cells found! Please correct this and try again.<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeMergedCells + "<br></font color></font>");
		}

		if(!storeEmptyRows.isEmpty()) {
			hasError = true;
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Empty rows found! Please remove them and try again.<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Row number: <font color=#ED0E3F>" + storeEmptyRows + "<br></font color></font>");
		}		
		return hasError;
	}

	public void printValueError() {

		if(!storeWrongID.isEmpty()) {
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>ID does not begin with '" + idAlphabet + "'<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeWrongID + "<br></font color></font>");
		}

		if(!storeDuplicateID.isEmpty()) {
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Duplicate ID<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeDuplicateID + "<br></font color></font>");
		}	

		if(!storeMissingValueCells.isEmpty()) {
			Collections.sort(storeMissingValueCells);
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Missing values found (Mandatory)<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeMissingValueCells + "<br></font color></font>");
		}

		if(!storeLineBreakCells.isEmpty()) {
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Cell contains line break<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeLineBreakCells + "<br></font color></font>");
		}

		if(!storeErrorMsg.isEmpty()) {
			for(int errorCount = 0; errorCount < storeErrorMsg.size(); errorCount++) {
				validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>" + storeErrorMsg.get(errorCount) + " <font color=#ED0E3F>(Row: " + storeErrorRow.get(errorCount) + ")<br></font color></font>");
			}
		}

		if (!storeInvalidColorCells.isEmpty()) {
			Collections.sort(storeInvalidColorCells);
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Color reference not found! (Check Colors sheet)<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeInvalidColorCells + "<br></font color></font>");
		}

		if(!storeReferenceErrorCells.isEmpty()) {
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Reference style not found! Check style ID again.<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cell: <font color=#ED0E3F>" + storeReferenceErrorCells + "<br></font color></font>");
		}

		if(!storeErrorSize.isEmpty()) {
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Size must be larger than 0.</font color><br></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeErrorSize + "<br></font color></font>");
		}

		if(!storeOpacityError.isEmpty()) {
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Opacity must be from 0 to 1.<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeOpacityError + "<br></font color></font>");
		}

		if(!storeClassValueError.isEmpty()) {
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Class name(s) only accept A-Z, a-z, 0-9, dash and underscores.<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeClassValueError + "<br></font color></font>");
		}

		if(storePencilLineJoin != null && !storePencilLineJoin.isEmpty()) {
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Pencil line join is not valid<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storePencilLineJoin + "<br></font color></font>");
		}

		if(storePencilLineCap != null && !storePencilLineCap.isEmpty()) {
			validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Pencil line cap is not valid<br></font color></font>");
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storePencilLineCap + "<br></font color></font>");
		}
	}

	public void printNoErrorMsg() {

		validateMessage = validateMessage.concat("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> No error found! </font color><br></font>");
	}
}
