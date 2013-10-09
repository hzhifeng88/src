import java.io.BufferedWriter;
import java.io.IOException;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

public class commonExport {

	private Sheet colorsSheet;
	public static String newline = System.getProperty("line.separator");

	public commonExport(Sheet colorSheet) {

		this.colorsSheet = colorSheet;
	}

	public String referenceColor(String givenColor) {

		for (int rowIndex = 4; rowIndex <= colorsSheet.getLastRowNum(); rowIndex++) {

			Row tempRow = colorsSheet.getRow(rowIndex);

			if(tempRow.getCell(0).toString().equalsIgnoreCase(givenColor)) {
				return tempRow.getCell(1).toString();
			}
		}
		return "not found";
	}

	public void appendLayerConditions(Row row, String storeCartoCSS, BufferedWriter writer, String currentClass, ArrayList<layersClassObject> storeClassObjects) throws IOException {

		Double tempDouble;

		// Generates comments if classes have duplicates
		for(int classCount = 0; classCount < storeClassObjects.size(); classCount++) {

			if(storeClassObjects.get(classCount).isHaveSame() == true && storeClassObjects.get(classCount).getRowIndex().equalsIgnoreCase(String.valueOf(row.getRowNum()+1))) {
				
				if(storeClassObjects.get(classCount).getTopic() != null) {
					writer.append("//Model: " + storeClassObjects.get(classCount).getModelName() + "  Topic: " + storeClassObjects.get(classCount).getTopic());
					writer.append("\r\n");
				}else {
					writer.append("//Model: " + storeClassObjects.get(classCount).getModelName());
					writer.append("\r\n");
				}
			}
		}
		
		storeCartoCSS = storeCartoCSS.concat("#"+ currentClass);

		// Attribute dependency
		if (row.getCell(5) != null && row.getCell(5).getCellType() != Cell.CELL_TYPE_BLANK) {

			if(row.getCell(5).toString().contains("AND")) {

				String tempString = row.getCell(5).toString();
				String tempAttrDependencyArray[] = tempString.split("AND");

				for(int count = 0; count < tempAttrDependencyArray.length; count++) {
					storeCartoCSS = storeCartoCSS.concat("["+ tempAttrDependencyArray[count].trim() + "]");
				}
			}else if(row.getCell(5).toString().contains("OR")) {

				String tempString = row.getCell(5).toString();
				String tempAttrDependencyArray[] = tempString.split("OR"); 

				for(int count = 0; count < tempAttrDependencyArray.length; count++) {
					storeCartoCSS = storeCartoCSS.concat("["+ tempAttrDependencyArray[count].trim() + "]");

					if(count < tempAttrDependencyArray.length-1) {
						storeCartoCSS = storeCartoCSS.concat("," + newline + "#"+ currentClass);
					}
				}
			}else {
				storeCartoCSS = storeCartoCSS.concat("["+ row.getCell(5).toString() + "]");
			}
		}

		if (row.getCell(9) != null && row.getCell(9).getCellType() != Cell.CELL_TYPE_BLANK) {
			tempDouble = new Double(row.getCell(9).toString());
			storeCartoCSS = storeCartoCSS.concat("[zoom>"+ (int)Math.round(tempDouble-1) + "]");
		}

		if (row.getCell(10) != null && row.getCell(10).getCellType() != Cell.CELL_TYPE_BLANK) {
			tempDouble = new Double(row.getCell(10).toString());
			storeCartoCSS = storeCartoCSS.concat("[zoom<"+ (int)Math.round(tempDouble+1) + "]");
		}
		writer.append(storeCartoCSS);
	}

	public void printExportReport(exportReport cartoReport, ArrayList<String> storeInvalidFont) {

		if(storeInvalidFont != null && storeInvalidFont.isEmpty() == false) {
			cartoReport.writeHeader("TextStyle");
			cartoReport.writeTextToReport("<font size = 4> <font color=#0A23C4><b>-> </b><font size = 3>Specified font not found!</font color></font>");
			cartoReport.writeTextToReport("<font size = 4> <font color=#0A23C4><font size = 3>** Default font <font color=#ED0E3F>\"Times New Roman Regular\" <font color=#0A23C4>is used.</font color></font>");
			cartoReport.writeTextToReport("<font size = 3> <font color=#0A23C4>Cells: <font color=#ED0E3F>" + storeInvalidFont + "</font color></font>");
		}else {
			cartoReport.writeTextToReport("<font size = 4> <font color=#088542><br><b>-> </b><font size = 3>CartoCSS files has been successfully exported. </font color></font>");
		}
	}
}
