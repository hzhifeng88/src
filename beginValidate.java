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
	private HTMLEditorKit kit;
	private HTMLDocument doc;
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

	public BeginValidate(Workbook workbook, Workbook templateWorkbook, HTMLEditorKit kit, HTMLDocument doc){

		this.workbook = workbook;	
		this.templateWorkbook = templateWorkbook;
		this.kit = kit;
		this.doc = doc;
	}

	public List<Boolean> startValidate() {

		checkExtraSheets();
		readColorsSheet();

		vLayers = new ValidateLayers(workbook.getSheetAt(0), workbook, templateWorkbook, storeColorID, kit, doc);
		List<String> storeTextGeometry = vLayers.validateSheet();
		storeIsSheetsCorrect.add(vLayers.isSheetCorrect());
		
		vPoint = new ValidatePointStyle(workbook.getSheetAt(1), templateWorkbook, storeColorID, kit, doc);
		vPoint.validateSheet();
		storeIsSheetsCorrect.add(vPoint.isSheetCorrect());

		vLine = new ValidateLineStyle(workbook.getSheetAt(2), workbook, templateWorkbook, storeColorID, kit, doc);
		vLine.validateSheet();
		storeIsSheetsCorrect.add(vLine.isSheetCorrect());

		vPolygon = new ValidatePolygonStyle(workbook.getSheetAt(3), workbook, templateWorkbook, storeColorID, kit, doc);
		vPolygon.validateSheet();
		storeIsSheetsCorrect.add(vPolygon.isSheetCorrect());

		vText = new ValidateTextStyle(workbook.getSheetAt(4), templateWorkbook, storeColorID, kit, doc);
		vText.validateSheet(storeTextGeometry);
		storeIsSheetsCorrect.add(vText.isSheetCorrect());

		vRaster = new ValidateRasterStyle(workbook.getSheetAt(5), templateWorkbook, storeColorID, kit, doc);
		vRaster.validateSheet();
		storeIsSheetsCorrect.add(vRaster.isSheetCorrect());

		vColors = new ValidateColors(workbook.getSheetAt(6), templateWorkbook, storeColorID, kit, doc);
		vColors.validateSheet();
		storeIsSheetsCorrect.add(vColors.isSheetCorrect());
		
		return storeIsSheetsCorrect;
	}

	public void checkExtraSheets() {

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

		try {

			if(!storeExtraSheets.isEmpty()){
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Extra sheet(s) found: <font color=#088542>" + storeExtraSheets + "<br><br></font color></font>", 0, 0, null);
			}
		} catch (BadLocationException | IOException e1) {
			JOptionPane.showMessageDialog(null, "An error has occurred (BeginValidate-HTMLkit). Application will now terminate.");
			System.exit(0);
		} 
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
