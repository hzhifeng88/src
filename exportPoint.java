import java.io.BufferedWriter;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;

public class exportPoint extends commonExport {

	private String styleID;
	private Sheet pointSheet;

	public exportPoint(Workbook workbook, String styleID)  throws IOException {

		super(workbook.getSheet("Colors"));

		this.styleID = styleID;
		this.pointSheet = workbook.getSheet("PointStyle");
	}

	public void exportNow(BufferedWriter writer) throws IOException {

		writer.append(" {\r\n");

		for (int rowIndex = 4; rowIndex <= pointSheet.getLastRowNum(); rowIndex++) {

			Row currentRow = pointSheet.getRow(rowIndex);

			if(currentRow.getCell(0).toString().equalsIgnoreCase(styleID)) {

				drawPointSymbols(currentRow, writer);
				drawMarkerOutlines(currentRow, writer);
				fillUpMarkerAreas(currentRow, writer);
				writer.append("}\r\n");
			}
		}	
	}

	public void drawPointSymbols(Row row, BufferedWriter writer) throws IOException {

		
		// Size
		if(row.getCell(1) != null && row.getCell(1).getCellType() != Cell.CELL_TYPE_BLANK) {

			writer.append("\tmarker-width: " + row.getCell(1) +";");
			writer.append("\r\n");
		}
		
		// Rotation
		if(row.getCell(2) != null && row.getCell(2).getCellType() != Cell.CELL_TYPE_BLANK) {

			writer.append("\tmarker-transform: rotate(" + row.getCell(2) +",0,0);");
			writer.append("\r\n");
		}
		
		// Graphic based
		if(row.getCell(6) != null && row.getCell(6).getCellType() != Cell.CELL_TYPE_BLANK) {

			writer.append("\tmarker-file: url(" + row.getCell(6) +");");
			writer.append("\r\n");
		}

	}

	public void drawMarkerOutlines(Row row, BufferedWriter writer) throws IOException {

		// Marker based color
		if(row.getCell(7) != null && row.getCell(7).getCellType() != Cell.CELL_TYPE_BLANK) {

			String foundColor = referenceColor(row.getCell(7).toString());
			writer.append("\tmarker-line-color: " + foundColor + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tmarker-line-color: #000000;");
			writer.append("\r\n");
		}

		// Marker color opacity
		if(row.getCell(8) != null && row.getCell(8).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tmarker-fill-opacity: " + row.getCell(8) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tmarker-fill-opacity: 1;");
			writer.append("\r\n");
		}

		// Marker width
		if(row.getCell(11) != null && row.getCell(11).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tmarker-line-width: " + row.getCell(11) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tmarker-line-width: 1;");
			writer.append("\r\n");
		}

	}

	public void fillUpMarkerAreas(Row row, BufferedWriter writer) throws IOException {

		// Marker color fill
		if(row.getCell(14) != null && row.getCell(14).getCellType() != Cell.CELL_TYPE_BLANK) {
			String foundColor = referenceColor(row.getCell(14).toString());
			writer.append("\tmarker-fill: " + foundColor + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tmarker-fill: #808080;");
			writer.append("\r\n");
		}

		// Marker color opacity
		if(row.getCell(15) != null && row.getCell(15).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tmarker-opacity: " + row.getCell(15) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tmarker-opacity: 1;");
			writer.append("\r\n");
		}
	}
}
