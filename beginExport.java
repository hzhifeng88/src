import java.util.*;
import java.io.*;

import javax.swing.JOptionPane;

import org.apache.poi.ss.usermodel.*;

public class BeginExport extends CommonExport{

	private Workbook workbook;
	private Sheet layersSheet;
	private BufferedWriter writer;
	private List<String> storeInvalidFont = new ArrayList<String>();
	private List<LayersClassObject> storeSortedClassObjects = new ArrayList<LayersClassObject>();
	private final String systemFileSeparator = System.getProperty("file.separator");
	private String fileName;
	private String desktopPath; 
	private String filePath; 
	private ExportPoint exPoint;
	private ExportLine exLine;
	private ExportPolygon exPolygon;
	private ExportText exText;
	private ExportRaster exRaster;

	public BeginExport(Workbook workbook, String fileName){

		super(workbook.getSheet("Colors"));
		this.workbook = workbook;
		this.fileName = fileName;
	}

	public boolean checkFileExist() {

		desktopPath = System.getProperty("user.home") + systemFileSeparator + "Desktop"; 
		filePath = desktopPath.replace("\\", systemFileSeparator); 
		
		File file = new File(filePath + systemFileSeparator + fileName.substring(0, fileName.length()-5) +".mss");
		return file.exists();
	}

	public List<String> exportNow(List<LayersClassObject> storeClassObjects) {

		layersSheet = workbook.getSheet("Layers");
		String storeCartoCSS = "";

		sortDrawingOrderAscending(storeClassObjects);

		try {
			
			writer = new BufferedWriter(new FileWriter(filePath + systemFileSeparator + fileName.substring(0, fileName.length()-5) +".mss"));

			for (int rowIndex = 0; rowIndex < storeSortedClassObjects.size(); rowIndex++) {

				Row currentRow = layersSheet.getRow(Integer.parseInt(storeSortedClassObjects.get(rowIndex).getRowIndex())-1);
				String tempClass = currentRow.getCell(2).toString();
				String geometryType = currentRow.getCell(4).toString();
				String styleID = currentRow.getCell(6).toString();

				appendLayerConditions(currentRow, storeCartoCSS, writer, tempClass, storeClassObjects);
				exportReferencedStyle(geometryType, styleID);
				writer.append(NEWLINE);
			}
			writer.close();
		}catch (IOException e) {
			JOptionPane.showMessageDialog(null, "An error has occurred(BeginExport-IOException). Application will now terminate.");
			System.exit(0);
		}
		return storeInvalidFont;
	}

	public void sortDrawingOrderAscending(List<LayersClassObject> storeClassObjects) {

		boolean found = false;
					
		for(int objectCount = 0; objectCount < storeClassObjects.size(); objectCount++) {

			found = false;
			
			if(storeSortedClassObjects.isEmpty()) {
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
				if(!found) {
					storeSortedClassObjects.add(storeClassObjects.get(objectCount));
				}
			}
		}
	}

	public void exportReferencedStyle(String geometryType, String styleID) throws IOException {

		switch(geometryType) {

		case "P":
			if(styleID.charAt(0) == 'P') {
				exPoint = new ExportPoint(workbook, styleID);
				exPoint.exportNow(writer, false);
			}else if(styleID.charAt(0) == 'T') {
				exText = new ExportText(workbook, geometryType, styleID);
				storeInvalidFont.addAll(exText.exportNow(writer));
			}
			break;

		case "L":
			if(styleID.charAt(0) == 'L') {
				exLine = new ExportLine(workbook, styleID);
				exLine.exportNow(writer, false);
			}else if(styleID.charAt(0) == 'T') {
				exText = new ExportText(workbook, geometryType, styleID);
				storeInvalidFont.addAll(exText.exportNow(writer));
			}
			break;

		case "A":
			if(styleID.charAt(0) == 'A') {
				exPolygon = new ExportPolygon(workbook, styleID);
				exPolygon.exportNow(writer);
			}
			break;

		case "R":
			if(styleID.charAt(0) == 'R') {
				exRaster = new ExportRaster();
			}
			break;
		}
	}
}
