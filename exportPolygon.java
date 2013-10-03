import java.io.BufferedWriter;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;

public class exportPolygon extends commonExport {

	private String styleID;
	private Sheet polygonSheet;
//	private Workbook workbook;
	
	public exportPolygon(Workbook workbook, String styleID) throws IOException {

		super(workbook.getSheet("Colors"));

//		this.workbook = workbook;
		this.styleID = styleID;
		this.polygonSheet = workbook.getSheet("PolygonStyle");
	}

	public void exportNow(BufferedWriter writer) throws IOException {

		writer.append(" {\r\n");

		for (int rowIndex = 4; rowIndex <= polygonSheet.getLastRowNum(); rowIndex++) {

			Row currentRow = polygonSheet.getRow(rowIndex);

			if(currentRow.getCell(0).toString().equalsIgnoreCase(styleID)) {

				fillArea(currentRow, writer);
				drawOutline(currentRow, writer);
				writer.append("}\r\n");
			}
		}
	}

	public void fillArea(Row row, BufferedWriter writer) throws IOException {

		// Solid color based
		if(row.getCell(2) != null && row.getCell(2).getCellType() != Cell.CELL_TYPE_BLANK) {
			
			String foundColor = referenceColor(row.getCell(2).toString());
			writer.append("\tpolygon-fill: " + foundColor + ";");
			writer.append("\r\n");
			
		}else {
			writer.append("\tmarker-line-color: #000000;");
			writer.append("\r\n");
		}

		// Solid color opacity
		if(row.getCell(3) != null && row.getCell(3).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tpolygon-opacity: " + row.getCell(3) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tpolygon-opacity: 1;");
			writer.append("\r\n");
		}

		// Implement Pattern-based area HERE
	}
	
	public void drawOutline(Row row, BufferedWriter writer) throws IOException {
		
		// Reference to a line style (is this correct?)
//		String referenceID = row.getCell(5).toString();
//		exportLine exLine = new exportLine(workbook, writer, referenceID);
	}
}
