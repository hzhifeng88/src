import java.io.BufferedWriter;
import java.io.IOException;
import org.apache.poi.ss.usermodel.*;

public class exportLine extends commonExport{

	private String styleID;
	private Sheet lineSheet;

	public exportLine(Workbook workbook, String styleID) throws IOException {

		super(workbook.getSheet("Colors"));

		this.styleID = styleID;
		this.lineSheet = workbook.getSheet("LineStyle");
	}

	public void exportNow(BufferedWriter writer) throws IOException {

		writer.append(" {\r\n");

		for (int rowIndex = 4; rowIndex <= lineSheet.getLastRowNum(); rowIndex++) {

			Row currentRow = lineSheet.getRow(rowIndex);

			if(currentRow.getCell(0).toString().equalsIgnoreCase(styleID)) {

				drawGeometryLines(currentRow, writer);
				writer.append("}\r\n");
			}
		}	
	}

	public void drawGeometryLines(Row row, BufferedWriter writer) throws IOException {

		// Pencil based color
		if(row.getCell(1) != null && row.getCell(1).getCellType() != Cell.CELL_TYPE_BLANK) {

			String foundColor = referenceColor(row.getCell(1).toString());
			writer.append("\tline-color: " + foundColor + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tline-color: #000000;");
			writer.append("\r\n");
		}

		// Pencil color opacity
		if(row.getCell(2) != null && row.getCell(2).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-opacity: " + row.getCell(2) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tline-opacity: 1;");
			writer.append("\r\n");
		}

		// Pencil dash array
		if(row.getCell(3) != null && row.getCell(3).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-dasharray: " + row.getCell(3) + ";");
			writer.append("\r\n");
		}

		// Pencil dash offset
		if(row.getCell(4) != null && row.getCell(4).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-dash-offset: " + row.getCell(4) + ";");
			writer.append("\r\n");
		}

		// Pencil width
		if(row.getCell(5) != null && row.getCell(5).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-width: " + row.getCell(5) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tline-width: 1;");
			writer.append("\r\n");
		}

		// Pencil line joint
		if(row.getCell(6) != null && row.getCell(6).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-join: " + row.getCell(6) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tline-join: round;");
			writer.append("\r\n");
		}

		// Pencil line cap
		if(row.getCell(7) != null && row.getCell(7).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-cap: " + row.getCell(7) + ";");
			writer.append("\r\n");
		}else {
			writer.append("\tline-cap: round;");
			writer.append("\r\n");
		}
	}
}
