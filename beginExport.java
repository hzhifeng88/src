import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import org.apache.poi.ss.usermodel.*;

public class beginExport extends commonExport{

	private Workbook workbook;
	private Sheet layersSheet;
	private BufferedWriter writer;
	private ArrayList<String> storeInvalidFont = new ArrayList<String>();
	private ArrayList<layersClassObject> storeSortedClassObjects = new ArrayList<layersClassObject>();
	private String fileName;
	private String desktopPath = System.getProperty("user.home") + "/Desktop"; 
	private String filePath = desktopPath.replace("\\", "/"); 

	private exportPoint exPoint;
	private exportLine exLine;
	private exportPolygon exPolygon;
	private exportText exText;
	private exportRaster exRaster;

	public beginExport(Workbook workbook, String fileName){

		super(workbook.getSheet("Colors"));
		this.workbook = workbook;
		this.fileName = fileName;
	}

	public boolean checkFileExist() {

		File file = new File(filePath + "/" + fileName.substring(0, fileName.length()-5) +".mss");

		if(!file.exists()) {
			return false;
		}
		return true;
	}

	public ArrayList<String> exportNow(ArrayList<layersClassObject> storeClassObjects) {

		layersSheet = workbook.getSheet("Layers");
		String geometryType = null;
		String styleID = null;
		String storeCartoCSS = "";
		Row currentRow;

		sortDrawingOrder(storeClassObjects);

		try {
			
			writer = new BufferedWriter(new FileWriter(filePath + "/" + fileName.substring(0, fileName.length()-5) +".mss"));

			for (int rowIndex = 0; rowIndex < storeSortedClassObjects.size(); rowIndex++) {

				currentRow = layersSheet.getRow(Integer.parseInt(storeSortedClassObjects.get(rowIndex).getRowIndex())-1);
				String tempClass = currentRow.getCell(2).toString();

				geometryType = currentRow.getCell(4).toString();
				styleID = currentRow.getCell(6).toString();

				appendLayerConditions(currentRow, storeCartoCSS, writer, tempClass, storeClassObjects);
				getRespectiveStyle(geometryType, styleID);
				writer.append("\r\n");
			}
			writer.close();
		}catch (IOException e) {
			e.printStackTrace();
		}
		return storeInvalidFont;
	}

	public void sortDrawingOrder(ArrayList<layersClassObject> storeClassObjects) {

		boolean found = false;
		
		// Sort to ascending drawing order			
		for(int objectCount = 0; objectCount < storeClassObjects.size(); objectCount++) {

			found = false;
			
			if(storeSortedClassObjects.isEmpty() == true) {
				storeSortedClassObjects.add(storeClassObjects.get(0));
			}else {
				
				for(int objectCount1 = 0; objectCount1 < storeSortedClassObjects.size(); objectCount1++) {

					if(storeClassObjects.get(objectCount).getDrawingOrder() < storeSortedClassObjects.get(objectCount1).getDrawingOrder()) {
						storeSortedClassObjects.add(objectCount1, storeClassObjects.get(objectCount));
						found = true;
						break;
					}else if(storeClassObjects.get(objectCount).getDrawingOrder() == storeSortedClassObjects.get(objectCount1).getDrawingOrder()) {
						storeSortedClassObjects.add(objectCount1+1, storeClassObjects.get(objectCount));
						found = true;
						break;
					}
				}
				if(found == false) {
					storeSortedClassObjects.add(storeClassObjects.get(objectCount));
				}
			}
		}
	}

	public void getRespectiveStyle(String geometryType, String styleID) throws IOException {

		switch(geometryType) {

		case "P":
			if(styleID.charAt(0) == 'P') {
				exPoint = new exportPoint(workbook, styleID);
				exPoint.exportNow(writer, false);
			}else if(styleID.charAt(0) == 'T') {
				exText = new exportText(workbook, geometryType, styleID);
				storeInvalidFont.addAll(exText.exportNow(writer));
			}
			break;

		case "L":
			if(styleID.charAt(0) == 'L') {
				exLine = new exportLine(workbook, styleID);
				exLine.exportNow(writer, false);
			}else if(styleID.charAt(0) == 'T') {
				exText = new exportText(workbook, geometryType, styleID);
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
