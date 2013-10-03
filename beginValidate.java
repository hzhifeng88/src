import java.io.IOException;
import java.util.ArrayList;
import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;
import org.apache.poi.ss.usermodel.*;

public class beginValidate {

	private Workbook workbook;
	private Workbook originalWorkbook;
	private HTMLEditorKit kit;
	private HTMLDocument doc;
	private ArrayList<String> storeColorID = new ArrayList<String>();
	private ArrayList<String> storeExtraSheets= new ArrayList<String>();
	private ArrayList<Boolean> storeIsSheetsCorrect= new ArrayList<Boolean>();

	private validateLayers vLayers;
	private validatePointStyle vPoint;
	private validateLineStyle vLine;
	private validatePolygonStyle vPolygon;
	private validateTextStyle vText;
	private validateRasterStyle vRaster;
	private validateColors vColors;

	public beginValidate(Workbook workbook, Workbook originalWorkbook, HTMLEditorKit kit, HTMLDocument doc){

		this.workbook = workbook;	
		this.originalWorkbook = originalWorkbook;
		this.kit = kit;
		this.doc = doc;
	}

	public ArrayList<Boolean> startValidate() {

		checkExtraSheets();
		readColorsSheet();

		vLayers = new validateLayers(workbook.getSheetAt(0), workbook, originalWorkbook, storeColorID, kit, doc);
		ArrayList<String> storeTextGeometry = vLayers.validateSheet();
		storeIsSheetsCorrect.add(vLayers.isSheetCorrect());
		
		vPoint = new validatePointStyle(workbook.getSheetAt(1), originalWorkbook, storeColorID, kit, doc);
		vPoint.validateSheet();
		storeIsSheetsCorrect.add(vPoint.isSheetCorrect());

		vLine = new validateLineStyle(workbook.getSheetAt(2), workbook, originalWorkbook, storeColorID, kit, doc);
		vLine.validateSheet();
		storeIsSheetsCorrect.add(vLine.isSheetCorrect());

		vPolygon = new validatePolygonStyle(workbook.getSheetAt(3), workbook, originalWorkbook, storeColorID, kit, doc);
		vPolygon.validateSheet();
		storeIsSheetsCorrect.add(vPolygon.isSheetCorrect());

		vText = new validateTextStyle(workbook.getSheetAt(4), originalWorkbook, storeColorID, kit, doc);
		vText.validateSheet(storeTextGeometry);
		storeIsSheetsCorrect.add(vText.isSheetCorrect());

		vRaster = new validateRasterStyle(workbook.getSheetAt(5), originalWorkbook, storeColorID, kit, doc);
		vRaster.validateSheet();
		storeIsSheetsCorrect.add(vRaster.isSheetCorrect());

		vColors = new validateColors(workbook.getSheetAt(6), originalWorkbook, storeColorID, kit, doc);
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

			if(storeExtraSheets.isEmpty() == false){
				kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Extra sheet(s) found: <font color=#088542>" + storeExtraSheets + "<br><br></font color></font>", 0, 0, null);
			}
		} catch (BadLocationException | IOException e1) {
			e1.printStackTrace();
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
