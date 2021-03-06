import java.io.*;
import java.util.ArrayList;
import java.util.Scanner;
import org.apache.poi.ss.usermodel.*;

public class ExportText extends CommonExport {

	private String geometryType;
	private String styleID;
	private Sheet textSheet;
	private ArrayList<String> storeInvalidFont = new ArrayList<String>();

	public ExportText(Workbook workbook, String geometryType, String styleID) throws IOException {

		super(workbook.getSheet("Colors"));
		this.geometryType = geometryType;
		this.styleID = styleID;
		this.textSheet = workbook.getSheet("TextStyle");
	}

	public ArrayList<String> exportNow(BufferedWriter writer) throws IOException {

		writer.append(" {" + NEWLINE);

		for (int rowIndex = 4; rowIndex <= textSheet.getLastRowNum(); rowIndex++) {

			Row currentRow = textSheet.getRow(rowIndex);

			if(currentRow.getCell(0).toString().equalsIgnoreCase(styleID)) {

				labelingGeneral(writer, currentRow);
				
				if(geometryType.equalsIgnoreCase("P")) {
					labelingToPoint(writer, currentRow);
				} else if(geometryType.equalsIgnoreCase("L")){
					labelingToLine(writer, currentRow);
				}

				fillArea(writer, currentRow);
				writer.append("}" + NEWLINE);
			}
		}
		return storeInvalidFont;
	}

	public boolean referenceFont(String givenFont) {

		InputStream is = ExportText.class.getResourceAsStream("resource/Fonts");
		
		@SuppressWarnings("resource")
		Scanner scanFont = new Scanner(is);
		scanFont.useDelimiter(System.getProperty("line.separator"));

		while (scanFont.hasNext()) {  
			if(givenFont.trim().equalsIgnoreCase(scanFont.next())){
				return true;
			}
		}
		return false;
	}

	public void labelingGeneral(BufferedWriter writer, Row row) throws IOException {

		// Label text
		writer.append("\ttext-name: \"[" + row.getCell(1) + "]\";");
		writer.append(NEWLINE);

		// Default font: Times new Roman Regular
		String fontFamily = row.getCell(2).toString();
		String tempFontArray[] = fontFamily.split(",");
		if(referenceFont(tempFontArray[0]) == true) {
			writer.append("\ttext-face-name: \"" + tempFontArray[0] + "\";");
			writer.append(NEWLINE);

		}else {
			writer.append("\ttext-face-name: \"Times New Roman Regular\";");
			writer.append(NEWLINE);
			storeInvalidFont.add("C" + Integer.toString(row.getRowNum() + 1));
		}

		// Halo color, Radius
		String haloColorRadius = row.getCell(3).toString();
		String tempHaloArray[] = haloColorRadius.split(",");
		writer.append("\ttext-halo-radius: " + tempHaloArray[1] + ";");
		writer.append(NEWLINE);
		String foundColor = referenceColor(tempHaloArray[0].toString());
		writer.append("\ttext-halo-fill: " + foundColor + ";");
		writer.append(NEWLINE);

		// Font size
		if(row.getCell(4) != null && row.getCell(4).getCellType() != Cell.CELL_TYPE_BLANK) {
			Double tempDouble = new Double(row.getCell(4).toString());
			writer.append("\ttext-size: " + (int)Math.round(tempDouble) + ";");
			writer.append(NEWLINE);
		}else {
			writer.append("\ttext-size: 10;");
			writer.append(NEWLINE);
		}
	}

	public void labelingToPoint(BufferedWriter writer, Row row) throws IOException {

		// Placement
		writer.append("\ttext-placement: point;");
		writer.append(NEWLINE);

		// Rotation
		writer.append("\ttext-orientation: " + row.getCell(6) + ";");
		writer.append(NEWLINE);

		// X,Y displacement
		String xyDisplacement = row.getCell(8).toString();
		String tempDisArray[] = xyDisplacement.split(",");
		writer.append("\ttext-dx: " + tempDisArray[0]  + ";");
		writer.append(NEWLINE);
		writer.append("\ttext-dy: " + tempDisArray[1]  + ";");
		writer.append(NEWLINE);
	}

	public void labelingToLine(BufferedWriter writer, Row row) throws IOException {

		// Perpendicular offset (dy)
		if(row.getCell(9) != null && row.getCell(9).getCellType() != Cell.CELL_TYPE_BLANK) {
			writer.append("\ttext-dy: " + row.getCell(9)  + ";");
			writer.append(NEWLINE);
		}

		// Repeated Gap; Initial Gap not supported
		// Only works when text is aligned to geometry(line)
		if(row.getCell(10) != null && row.getCell(10).getCellType() != Cell.CELL_TYPE_BLANK) {
			String repeatedGaps = row.getCell(10).toString();
			String tempGapArray[] = repeatedGaps.split(",");
			writer.append("\ttext-spacing: " + tempGapArray[1]  + ";");
			writer.append(NEWLINE);
		}

		// Alignment (Geometry or Horizontal)
		// When line placement is specified, geometry means line,
		// if not specified, default is Point, which is Horizontal
		if(row.getCell(11) != null && row.getCell(11).getCellType() != Cell.CELL_TYPE_BLANK) {

			if(row.getCell(11).toString().equalsIgnoreCase("geometry")){
				writer.append("\ttext-placement: line;");
				writer.append(NEWLINE);
			}else if(row.getCell(11).toString().equalsIgnoreCase("horizontal")){
				writer.append("\ttext-placement: point;");
				writer.append(NEWLINE);
			}
		}else {
			writer.append("\ttext-placement: line;");
			writer.append(NEWLINE);
		}
	}

	public void fillArea(BufferedWriter writer, Row row) throws IOException {

		// Solid color based
		if(row.getCell(12) != null && row.getCell(12).getCellType() != Cell.CELL_TYPE_BLANK) {
			String foundColor = referenceColor(row.getCell(12).toString());
			writer.append("\ttext-fill: " + foundColor + ";");
			writer.append(NEWLINE);
		}else {
			writer.append("\ttext-fill: #808080;");
			writer.append(NEWLINE);
		}

		// Solid color opacity
		writer.append("\ttext-opacity: " + row.getCell(13) + ";");
		writer.append(NEWLINE);
	}
}