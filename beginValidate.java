import java.io.IOException;
import java.util.*;

import javax.swing.JOptionPane;
import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.ss.usermodel.*;

public class BeginValidate {

	private Workbook workbook;
	private Workbook templateWorkbook;
	private List<String> storeColorID = new ArrayList<String>();
	private List<String> storeExtraSheets= new ArrayList<String>();
	private List<Boolean> storeIsSheetsCorrect= new ArrayList<Boolean>();
	private ValidateLayers vLayers;
	private ValidatePointStyle vPoint;
	private ValidateLineStyle vLine;
	private ValidatePolygonStyle vPolygon;
	private ValidateTextStyle vText;
	private ValidateRasterStyle vRaster;
	private ValidateColors vColors;

	public BeginValidate(Workbook workbook, Workbook templateWorkbook){

		this.workbook = workbook;	
		this.templateWorkbook = templateWorkbook;
	}

	public boolean startValidate(String validateMessage, boolean onlyValidate, HTMLEditorKit kit, HTMLDocument doc) {

		validateMessage = checkExtraSheets(validateMessage);
		readColorsSheet();

		vLayers = new ValidateLayers(workbook.getSheetAt(0), workbook, templateWorkbook, storeColorID, validateMessage);
		List<String> storeTextGeometry = vLayers.validateSheet();
		storeIsSheetsCorrect.add(vLayers.isSheetCorrect());
		validateMessage = vLayers.getMessage();

		vPoint = new ValidatePointStyle(workbook.getSheetAt(1), templateWorkbook, storeColorID, validateMessage);
		vPoint.validateSheet();
		storeIsSheetsCorrect.add(vPoint.isSheetCorrect());
		validateMessage = vPoint.getMessage();

		vLine = new ValidateLineStyle(workbook.getSheetAt(2), workbook, templateWorkbook, storeColorID, validateMessage);
		vLine.validateSheet();
		storeIsSheetsCorrect.add(vLine.isSheetCorrect());
		validateMessage = vLine.getMessage();

		vPolygon = new ValidatePolygonStyle(workbook.getSheetAt(3), workbook, templateWorkbook, storeColorID, validateMessage);
		vPolygon.validateSheet();
		storeIsSheetsCorrect.add(vPolygon.isSheetCorrect());
		validateMessage = vPolygon.getMessage();

		vText = new ValidateTextStyle(workbook.getSheetAt(4), templateWorkbook, storeColorID, validateMessage);
		vText.validateSheet(storeTextGeometry);
		storeIsSheetsCorrect.add(vText.isSheetCorrect());
		validateMessage = vText.getMessage();

		vRaster = new ValidateRasterStyle(workbook.getSheetAt(5), templateWorkbook, storeColorID, validateMessage);
		vRaster.validateSheet();
		storeIsSheetsCorrect.add(vRaster.isSheetCorrect());
		validateMessage = vRaster.getMessage();

		vColors = new ValidateColors(workbook.getSheetAt(6), templateWorkbook, storeColorID, validateMessage);
		vColors.validateSheet();
		storeIsSheetsCorrect.add(vColors.isSheetCorrect());
		validateMessage = vColors.getMessage();

		try {
			if(onlyValidate) {
				kit.insertHTML(doc, doc.getLength(), validateMessage, 0, 0,null);
			}else {
				if(storeIsSheetsCorrect.contains(false)){
					kit.insertHTML(doc, doc.getLength(), validateMessage, 0, 0,null);
					JOptionPane.showMessageDialog(null, "Please correct the error first!");
					return false;
				}
			}
		} catch (BadLocationException | IOException e) {
			JOptionPane.showMessageDialog(null, "An error has occurred (BeginValidate-startValidate). Application will now terminate.");
			System.exit(0);
		} 
		return true;
	}

	public String checkExtraSheets(String validateMessage) {

		String tempSheetName;
		storeExtraSheets.clear();

		for(int countSheet = 0; countSheet < workbook.getNumberOfSheets(); countSheet++){

			Sheet tempSheet = workbook.getSheetAt(countSheet);
			tempSheetName = tempSheet.getSheetName();

			switch(tempSheetName) {
			case "Layers":
				break;
			case "PointStyle":  
				break;
			case "LineStyle":  
				break;
			case "PolygonStyle": 
				break;
			case "TextStyle": 
				break;
			case "RasterStyle": 
				break;	  
			case "Colors":
				break;
			default: storeExtraSheets.add(tempSheetName);
			break;
			}
		}

		if(!storeExtraSheets.isEmpty()){
			validateMessage = validateMessage.concat("<font size = 3> <font color=#0A23C4>Extra sheet(s) found: <font color=#088542>" + storeExtraSheets + "<br><br></font color></font>");
		}
		return validateMessage;
	}

	public void readColorsSheet() {

		Sheet sheet = workbook.getSheetAt(6);

		for (int rowIndex = 4; rowIndex <= sheet.getLastRowNum(); rowIndex++) {

			Row row = sheet.getRow(rowIndex);

			if(row.getCell(0) != null) {
				storeColorID.add(row.getCell(0).toString());
			}
		}
	}
}
