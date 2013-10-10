import java.io.IOException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.StringTokenizer;

import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.*;

public class commonValidate {

	private char idAlphabet;
	public boolean hasError = false;
	private HTMLEditorKit kit;
	private HTMLDocument doc;
	private Sheet sheet;
	private Workbook originalWorkbook;
	private ArrayList<String> storeColorID = new ArrayList<String>();					//for cross referencing
	private ArrayList<String> storeModifiedHeaderCells = new ArrayList<String>();
	private ArrayList<String> storeMergedCells = new ArrayList<String>();
	private ArrayList<String> storeEmptyRows = new ArrayList<String>();
	private ArrayList<String> storeLineBreakCells = new ArrayList<String>();
	private ArrayList<String> storeRightID = new ArrayList<String>();					//for checking duplicates
	private ArrayList<String> storeWrongID = new ArrayList<String>();
	private ArrayList<String> storeDuplicateID = new ArrayList<String>();
	private ArrayList<String> storeMissingValueCells = new ArrayList<String>();
	private ArrayList<String> storeInvalidColorCells = new ArrayList<String>();
	private ArrayList<String> storeReferenceErrorCells = new ArrayList<String>();
	private ArrayList<String> storeErrorRow = new ArrayList<String>();
	private ArrayList<String> storeErrorMsg = new ArrayList<String>();
	private ArrayList<String> storePencilLineJoin = new ArrayList<String>();
	private ArrayList<String> storePencilLineCap = new ArrayList<String>();
	public ArrayList<String> storeTextGeometry = new ArrayList<String>();
	private ArrayList<String> storeErrorSize = new ArrayList<String>();
	private ArrayList<String> storeOpacityError = new ArrayList<String>();
	private ArrayList<String> storeClassValueError = new ArrayList<String>();

	public commonValidate(Sheet sheet, Workbook originalWorkbook, ArrayList<String> colorList, HTMLEditorKit kit, HTMLDocument doc){

		this.sheet = sheet;	
		this.originalWorkbook = originalWorkbook;
		this.kit = kit;
		this.doc = doc;
		storeColorID = colorList;
	}

	public boolean validateFormat() {

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

		Sheet originalSheet = originalWorkbook.getSheet(sheet.getSheetName());

		for(int rowIndex = 0; rowIndex < 4; rowIndex++){

			Row row = sheet.getRow(rowIndex);
			Row originalRow = originalSheet.getRow(rowIndex);

			for(int columnIndex = 0; columnIndex < row.getLastCellNum(); columnIndex++){

				if(row.getCell(columnIndex) == null && originalRow.getCell(columnIndex) == null){
					continue;
				}

				if(row.getCell(columnIndex) != null && originalRow.getCell(columnIndex) == null){
					storeModifiedHeaderCells.add(columnIndexToLetter(columnIndex) + Integer.toString(rowIndex + 1));
					continue;
				}

				if(row.getCell(columnIndex).toString().equalsIgnoreCase(originalRow.getCell(columnIndex).toString())){
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
			if (isRowEmpty == true) {
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

	public void checkIDAndDuplicate(char idAlphabet, String column, int rowIndex, int columnIndex){

		boolean wrongStyleID = false;
		this.idAlphabet = idAlphabet;

		Row row = sheet.getRow(rowIndex);

		if(row.getCell(columnIndex) != null){

			String tempString = row.getCell(columnIndex).toString();

			if (!tempString.equalsIgnoreCase("")) {

				char firstChar = tempString.charAt(0);

				if (firstChar != idAlphabet) {
					wrongStyleID = true;
					hasError = true;
					storeWrongID.add(column + Integer.toString(rowIndex + 1));
				}

				if (wrongStyleID == false) {

					if (storeRightID.isEmpty() == true) {
						storeRightID.add(tempString);
					} else {

						if(storeRightID.contains(tempString)){
							hasError = true;
							storeDuplicateID.add(column + Integer.toString(rowIndex + 1));
						}else{
							storeRightID.add(tempString);
						}
					}
				}
			}
		}
	}

	public void checkMandatoryAttributes(Row row, int rowIndex, int[] columnArray) {

		// Checks for mandatory columns here
		for(int columnCount = 0; columnCount < columnArray.length; columnCount++) {

			if(row.getCell(columnArray[columnCount]) == null || row.getCell(columnArray[columnCount]).toString().equalsIgnoreCase("")){
				hasError = true;
				storeMissingValueCells.add(columnIndexToLetter(columnArray[columnCount]) + Integer.toString(rowIndex + 1));
			}
		}
	}

	public void checkSize(Row row, int columnIndex, String columnLetter) {

		// Check Size (positive)
		if(row.getCell(columnIndex) != null && row.getCell(columnIndex).getCellType() != Cell.CELL_TYPE_BLANK) {
			if(Double.parseDouble(row.getCell(columnIndex).toString()) <= 0.0) {
				hasError = true;
				storeErrorSize.add(columnLetter + Integer.toString(row.getRowNum() + 1));
			}
		}
	}

	public void checkOpacity(Row row, int columnIndex, String columnLetter) {

		// Opacity should be from 0 to 1
		if(row.getCell(columnIndex) != null && row.getCell(columnIndex).getCellType() != Cell.CELL_TYPE_BLANK) {

			Double opacityDouble = Double.parseDouble(row.getCell(columnIndex).toString());
			if(opacityDouble < 0 || opacityDouble > 1.0) {
				hasError = true;
				storeOpacityError.add(columnLetter + Integer.toString(row.getRowNum() + 1));
			}
		}
	}

	public void matchColor(String tempStringColor, String currentColumn, int rowIndex) {

		if(storeColorID.contains(tempStringColor)){
			return;
		}
		hasError = true;
		storeInvalidColorCells.add(currentColumn + Integer.toString(rowIndex + 1));
	}

	public void checkPencilLineJoin(Row row, int rowIndex, int column, String columnLetter) {

		final ArrayList<String> lineJoin = new ArrayList<String>() {{
			add("mitre");
			add("round");
			add("bevel");
		}};

		if(lineJoin.contains(row.getCell(column).toString())){
			return;
		}
		hasError = true;
		storePencilLineJoin.add(columnLetter + Integer.toString(rowIndex + 1));
	}

	public void checkPencilLineCap(Row row, int rowIndex, int column, String columnLetter) {

		final ArrayList<String> lineCap = new ArrayList<String>() {{
			add("butt");
			add("round");
			add("square");
		}};

		if(lineCap.contains(row.getCell(column).toString())){
			return;
		}
		hasError = true;
		storePencilLineCap.add(columnLetter + Integer.toString(rowIndex + 1));
	}

	public void referenceCheck(Workbook workbook, Sheet currentSheet, Row row, int rowIndex, int column, char referenceTo) {

		String cellLocation = "";
		String referenceStyleID = "";

		cellLocation = columnIndexToLetter(column) + Integer.toString(rowIndex + 1);
		referenceStyleID = row.getCell(column).toString();

		char firstChar = referenceStyleID.charAt(0);

		if(firstChar == referenceTo) {
			getReferenceStyle(workbook, currentSheet, referenceStyleID, cellLocation, null);
		}else {
			hasError = true;
			storeReferenceErrorCells.add(cellLocation);
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
	
	public boolean printFormatError() {

		try {
			if(sheet.getSheetName().equalsIgnoreCase("Layers")){
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Error(s) in sheet: <font color=#ED0E3F><b>" + sheet.getSheetName() + "</b></font color></font>", 0, 0, null);
			}else{
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4><br>Error(s) in sheet: <font color=#ED0E3F><b>" + sheet.getSheetName() + "</b></font color></font>", 0, 0, null);
			}
			kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#088542>---------------------------------------------- </font color></font>", 0, 0, null);

			if(storeModifiedHeaderCells.isEmpty() == false){
				hasError = true;
				Collections.sort(storeModifiedHeaderCells);
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Header cells are modified! Please correct this and try again.</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeModifiedHeaderCells + "</font color></font>", 0, 0, null);
			}

			if (storeMergedCells.isEmpty() == false) {
				hasError = true;
				Collections.sort(storeMergedCells);
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>Merged cells found! Please correct this and try again.</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeMergedCells + "</font color></font>", 0, 0, null);
			}

			if (storeEmptyRows.isEmpty() == false) {
				hasError = true;
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Empty rows found! Please remove them and try again.</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Row number: <font color=#ED0E3F>" + storeEmptyRows + "</font color></font>", 0, 0, null);
			}		
		} catch (BadLocationException | IOException e) {
			e.printStackTrace();
		}
		return hasError;
	}

	public boolean printValueError() {

		try {
			if (storeWrongID.isEmpty() == false) {
				hasError = true;
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>ID does not begin with '" + idAlphabet + "'</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeWrongID + "</font color></font>", 0, 0, null);
			}

			if(storeDuplicateID.isEmpty() == false) {
				hasError = true;
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Duplicate ID</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeDuplicateID + "</font color></font>", 0, 0, null);
			}	

			if (storeMissingValueCells.isEmpty() == false) {
				hasError = true;
				Collections.sort(storeMissingValueCells);
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Missing values found (Mandatory)</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeMissingValueCells + "</font color></font>", 0, 0, null);
			}

			if(storeLineBreakCells.isEmpty() == false) {
				hasError = true;
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Cell contains line break</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeLineBreakCells + "</font color></font>", 0, 0, null);
			}

			if(storeErrorMsg.isEmpty() == false) {
				hasError = true;
				for(int errorCount = 0; errorCount < storeErrorMsg.size(); errorCount++) {
					kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>" + storeErrorMsg.get(errorCount) + " <font color=#ED0E3F>(Row: " + storeErrorRow.get(errorCount) + ")</font color></font>", 0, 0,null);
				}
			}

			if(storeErrorSize.isEmpty() == false) {
				hasError = true;
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Size must be larger than 0.</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeErrorSize + "</font color></font>", 0, 0, null);
			}

			if(storeOpacityError.isEmpty() == false) {
				hasError = true;
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Opacity must be from 0 to 1.</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeOpacityError + "</font color></font>", 0, 0, null);
			}

			if(storeClassValueError.isEmpty() == false) {
				hasError = true;
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Class name(s) only accept A-Z, a-z, 0-9, dash and underscores.</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeClassValueError + "</font color></font>", 0, 0, null);
			}
		} catch (BadLocationException | IOException e) {
			e.printStackTrace();
		}
		return hasError;
	}

	public void printColorError() {

		try {
			if (storeInvalidColorCells.isEmpty() == false) {
				Collections.sort(storeInvalidColorCells);
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Color reference not found! (Check Colors sheet)</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeInvalidColorCells + "</font color></font>", 0, 0, null);
			}

			if(storeReferenceErrorCells.isEmpty() == false) {
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Reference style not found! Check style ID again.</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cell: <font color=#ED0E3F>" + storeReferenceErrorCells + "</font color></font>", 0, 0, null);
			}
		} catch (BadLocationException | IOException e) {
			e.printStackTrace();
		}
	}

	public void printReferenceError() {

		try {
			if(storeReferenceErrorCells.isEmpty() == false) {
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Reference style not found! Check style ID again.</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cell: <font color=#ED0E3F>" + storeReferenceErrorCells + "</font color></font>", 0, 0, null);
			}
		} catch (BadLocationException | IOException e) {
			e.printStackTrace();
		}
	}

	public void printPencilBasedError() {

		try {
			if(storePencilLineJoin != null && storePencilLineJoin.isEmpty() == false) {
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Pencil line join is not valid</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storePencilLineJoin + "</font color></font>", 0, 0, null);
			}

			if(storePencilLineCap != null && storePencilLineCap.isEmpty() == false) {
				kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> Pencil line cap is not valid</font color></font>", 0, 0,null);
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storePencilLineCap + "</font color></font>", 0, 0, null);
			}
		} catch (BadLocationException | IOException e) {
			e.printStackTrace();
		}
	}

	public void printNoErrorMsg() {

		try {
			kit.insertHTML(doc, doc.getLength(), "<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3> No error found! </font color></font>", 0, 0,null);
		} catch (BadLocationException | IOException e) {
			e.printStackTrace();
		}
	}
}
