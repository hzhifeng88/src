import java.io.BufferedWriter;
import java.io.IOException;

import org.apache.poi.ss.usermodel.*;

public class ExportLine extends CommonExport{

	private String styleID;
	private Sheet lineSheet;
	private Workbook workbook;

	public ExportLine(Workbook workbook, String styleID) throws IOException {

		super(workbook.getSheet("Colors"));
		this.workbook = workbook;
		this.styleID = styleID;
		this.lineSheet = workbook.getSheet("LineStyle");
	}

	public void exportNow(BufferedWriter writer, boolean isReference) throws IOException {

		if(!isReference) {
			writer.append("}" + NEWLINE);
		}	

		for (int rowIndex = 4; rowIndex <= lineSheet.getLastRowNum(); rowIndex++) {

			Row currentRow = lineSheet.getRow(rowIndex);

			if(currentRow.getCell(0).toString().equalsIgnoreCase(styleID)) {

				drawGeometryLines(currentRow, writer);
				referenceToPointStyle(currentRow,  writer);
			
				if(!isReference) {
					writer.append("}" + NEWLINE);
				}
			}
		}	
	}

	public void drawGeometryLines(Row row, BufferedWriter writer) throws IOException {
		
		pencilColor(row, writer);
		pencilOpacity(row, writer);
		pencilDashArrayNOffset(row, writer);
		pencilWidth(row, writer);
		pencilLineJoinNCap(row, writer);
	}
	
	public void pencilColor(Row row, BufferedWriter writer) throws IOException {
		
		if(row.getCell(1) != null && row.getCell(1).getCellType() != Cell.CELL_TYPE_BLANK) {

			String foundColor = referenceColor(row.getCell(1).toString());
			writer.append("\tline-color: " + foundColor + ";");
			writer.append(NEWLINE);
		}else {
			writer.append("\tline-color: #000000;");
			writer.append(NEWLINE);
		}
	}
	
	public void pencilOpacity(Row row, BufferedWriter writer) throws IOException {
		
		if(row.getCell(2) != null && row.getCell(2).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-opacity: " + row.getCell(2) + ";");
			writer.append(NEWLINE);
		}else {
			writer.append("\tline-opacity: 1;");
			writer.append(NEWLINE);
		}
	}
	
	public void pencilDashArrayNOffset(Row row, BufferedWriter writer) throws IOException {
		
		if(row.getCell(3) != null && row.getCell(3).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-dasharray: " + row.getCell(3) + ";");
			writer.append(NEWLINE);
		}

		if(row.getCell(4) != null && row.getCell(4).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-dash-offset: " + row.getCell(4) + ";");
			writer.append(NEWLINE);
		}
	}

	public void pencilWidth(Row row, BufferedWriter writer) throws IOException {
		
		if(row.getCell(5) != null && row.getCell(5).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-width: " + row.getCell(5) + ";");
			writer.append(NEWLINE);
		}else {
			writer.append("\tline-width: 1;");
			writer.append(NEWLINE);
		}
	}
	
	public void pencilLineJoinNCap(Row row, BufferedWriter writer) throws IOException {
		
		if(row.getCell(6) != null && row.getCell(6).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-join: " + row.getCell(6) + ";");
			writer.append(NEWLINE);
		}else {
			writer.append("\tline-join: round;");
			writer.append(NEWLINE);
		}

		if(row.getCell(7) != null && row.getCell(7).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\tline-cap: " + row.getCell(7) + ";");
			writer.append(NEWLINE);
		}else {
			writer.append("\tline-cap: round;");
			writer.append(NEWLINE);
		}
	}
	
	public void referenceToPointStyle(Row row, BufferedWriter writer) throws IOException {
		
		if(row.getCell(8) != null && row.getCell(8).getCellType() != Cell.CELL_TYPE_BLANK) {
			
			String referenceID = row.getCell(8).toString();
			ExportPoint exPoint = new ExportPoint(workbook, referenceID);
			exPoint.exportNow(writer, true);
		}
	}
}
