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

		File file = new File(filePath + "/" + fileName +".mss");

		if(!file.exists()) {
			return false;
		}
		return true;
	}

	public ArrayList<String> exportNow(ArrayList<layersClassObject> storeClassObjects) {

		layersSheet = workbook.getSheet("Layers");
		String geometryType = null;
		String styleID = null;

		try {

			Row currentRow = layersSheet.getRow(4);

			writer = new BufferedWriter(new FileWriter(filePath + "/" + fileName +".mss"));
			String storeCartoCSS = "";

			for (int rowIndex = 4; rowIndex <= layersSheet.getLastRowNum(); rowIndex++) {

				currentRow = layersSheet.getRow(rowIndex);
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

	public void getRespectiveStyle(String geometryType, String styleID) throws IOException {

		switch(geometryType) {

		case "P":
			if(styleID.charAt(0) == 'P') {
				exPoint = new exportPoint(workbook, styleID);
				exPoint.exportNow(writer);
			}else if(styleID.charAt(0) == 'T') {
				exText = new exportText(workbook, geometryType, styleID);
				storeInvalidFont.addAll(exText.exportNow(writer));
			}
			break;

		case "L":
			if(styleID.charAt(0) == 'L') {
				exLine = new exportLine(workbook, styleID);
				exLine.exportNow(writer);
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
