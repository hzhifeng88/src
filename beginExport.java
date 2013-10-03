import java.io.BufferedWriter;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.*;

public class beginExport  extends commonExport{

	private Workbook workbook;
	private Sheet layersSheet;
	private BufferedWriter writer;
	public ArrayList<String> storeInvalidFont = new ArrayList<String>();
	private String currentModel;
	private String currentClass;
	private String desktopPath = System.getProperty("user.home") + "/Desktop"; 
	private String filePath = desktopPath.replace("\\", "/"); 

	private exportPoint exPoint;
	private exportLine exLine;
	private exportPolygon exPolygon;
	private exportText exText;
	private exportRaster exRaster;
	private exportReport cartoReport;

	public beginExport(Workbook workbook, exportReport cartoReport){

		super(workbook.getSheet("Colors"));
		this.workbook = workbook;
		this.cartoReport = cartoReport;
	}

	public void startExport() {

		layersSheet = workbook.getSheet("Layers");
		String geometryType = null;
		String styleID = null;
		
		try {

			Row currentRow = layersSheet.getRow(4);
			currentModel = currentRow.getCell(0).toString();
			currentClass = currentRow.getCell(2).toString();

			writer = new BufferedWriter(new FileWriter(filePath + "/" + currentModel + " - " + currentClass + ".mss"));
			String storeCartoCSS = "";

			for (int rowIndex = 4; rowIndex <= layersSheet.getLastRowNum(); rowIndex++) {

				currentRow = layersSheet.getRow(rowIndex);
				String tempModel = currentRow.getCell(0).toString();
				String tempClass = currentRow.getCell(2).toString();

				if(tempModel.equalsIgnoreCase(currentModel) && tempClass.equalsIgnoreCase(currentClass)) {

					geometryType = currentRow.getCell(4).toString();
					styleID = currentRow.getCell(6).toString();
					
					appendLayerConditions(currentRow, storeCartoCSS, writer, tempClass);
					getRespectiveStyle(geometryType, styleID);
					
				}else {
					writer.close();
					currentModel = tempModel;
					currentClass = tempClass;
					writer = new BufferedWriter(new FileWriter(filePath + "/" + currentModel + " - " + currentClass + ".mss"));
					
					geometryType = currentRow.getCell(4).toString();
					styleID = currentRow.getCell(6).toString();
					
					appendLayerConditions(currentRow, storeCartoCSS, writer, tempClass);
					getRespectiveStyle(geometryType, styleID);
				}
				
			}
			writer.close();
			exText.printExportReport(storeInvalidFont);
		}catch (IOException e) {
			e.printStackTrace();
		}
	}
	
	public void getRespectiveStyle(String geometryType, String styleID) throws IOException {
		
		switch(geometryType) {

		case "P":
			if(styleID.charAt(0) == 'P') {
				exPoint = new exportPoint(workbook, styleID);
				exPoint.exportNow(writer);
			}else if(styleID.charAt(0) == 'T') {
				exText = new exportText(workbook, geometryType, styleID, cartoReport);
				storeInvalidFont.addAll(exText.exportNow(writer));
			}
			break;

		case "L":
			if(styleID.charAt(0) == 'L') {
				exLine = new exportLine(workbook, styleID);
				exLine.exportNow(writer);
			}else if(styleID.charAt(0) == 'T') {
				exText = new exportText(workbook, geometryType, styleID, cartoReport);
				storeInvalidFont.addAll(exText.exportNow(writer));
			}
			break;

		case "A":
			if(styleID.charAt(0) == 'A') {
				exPolygon = new exportPolygon(workbook, styleID);
				exPolygon.exportNow(writer);
			}
			break;

		case "R":
			if(styleID.charAt(0) == 'R') {
				exRaster = new exportRaster();
			}
			break;
		}
	}
}
