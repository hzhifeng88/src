import java.io.BufferedWriter;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;

public class ExportPolygon extends CommonExport {

	private String styleID;
	private Sheet polygonSheet;
	private Workbook workbook;
	
	public ExportPolygon(Workbook workbook, String styleID) throws IOException {

		super(workbook.getSheet("Colors"));
		this.workbook = workbook;
		this.styleID = styleID;
		this.polygonSheet = workbook.getSheet("PolygonStyle");
	}

	public void exportNow(BufferedWriter writer) throws IOException {

		writer.append(" {" + NEWLINE);

		for (int rowIndex = 4; rowIndex <= polygonSheet.getLastRowNum(); rowIndex++) {

			Row currentRow = polygonSheet.getRow(rowIndex);

			if(currentRow.getCell(0).toString().equalsIgnoreCase(styleID)) {

				fillArea(currentRow, writer);
				drawOutline(currentRow, writer);
				writer.append("}" + NEWLINE);
			}
		}
	}

	public void fillArea(Row row, BufferedWriter writer) throws IOException {

		// Solid color based
		if(row.getCell(2) != null && row.getCell(2).getCellType() != Cell.CELL_TYPE_BLANK) {
			String foundColor = referenceColor(row.getCell(2).toString());
			writer.append("\tpolygon-fill: " + foundColor + ";");
			writer.append(NEWLINE);
			
		}else {
			writer.append("\tpolygon-fill: #808080;");
			writer.append(NEWLINE);
		}

		// Solid color opacity
		if(row.getCell(3) != null && row.getCell(3).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tpolygon-opacity: " + row.getCell(3) + ";");
			writer.append(NEWLINE);
		}else {
			writer.append("\tpolygon-opacity: 1;");
			writer.append(NEWLINE);
		}
		
		// Reference to a point style
		if(row.getCell(4) != null && row.getCell(4).getCellType() != Cell.CELL_TYPE_BLANK) {
			String referenceID = row.getCell(4).toString();
			ExportPoint exPoint = new ExportPoint(workbook, referenceID);
			exPoint.exportNow(writer, true);
		}
	}
	
	public void drawOutline(Row row, BufferedWriter writer) throws IOException {
		
		// Reference to a line style
		String referenceID = row.getCell(5).toString();
		ExportLine exLine = new ExportLine(workbook, referenceID);
		exLine.exportNow(writer, true);
	}
}
