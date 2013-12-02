import java.io.*;
import java.text.*;
import java.util.*;
import java.util.List;
import java.awt.*;
import java.awt.Color;
import java.awt.event.*;
import javax.swing.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.text.BadLocationException;
import javax.swing.text.html.HTMLDocument;
import javax.swing.text.html.HTMLEditorKit;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.*;

public class CatalogMain extends JFrame {

	private static CatalogMain mainWindow;
	public final static String version = "1.1";
	private boolean countOne = false;
	private boolean hasValidated = false;
	private String excelFilePath;
	private String fileDirectory;
	private String fileName;
	private String oldExcelFilePath = "";
	private JFileChooser chooser;
	private JPanel northPanel;
	private JButton openFileButton;
	private JButton validateButton;
	private JButton exportCSSButton;
	private JButton aboutButton;
	private JButton exitButton;
	private JTextField pathTextField;
	private JTextPane errorPane;
	private JTextPane aboutPane;
	private JScrollPane scrollPane;
	private List<Boolean> storeIsSheetsCorrect= new ArrayList<Boolean>();
	private List<LayersClassObject> storeClassObjects = new ArrayList<LayersClassObject>();
	private HTMLEditorKit kit;
	private HTMLDocument doc;
	private HTMLEditorKit kit1;
	private HTMLDocument doc1;
	private BeginValidate beginValidate;
	private BeginExport beginExport;
	private ExportReport cartoReport;
	private LayersClassObject classObjects;
	private Workbook workbook;
	private Workbook originalWorkbook;
	private Properties prop = new Properties();

	public CatalogMain() {

		createNorthPanel();
		createSouthPanel();

		getContentPane().add(northPanel, BorderLayout.NORTH);
		getContentPane().add(scrollPane, BorderLayout.CENTER);
	}

	public static void main(String[] args) {

		if(checkArgument(args)) {
			
			setLookAndFeel();

			mainWindow = new CatalogMain();
			mainWindow.setTitle("Portrayal Catalogue Valdiator " + version);
			mainWindow.setSize(530, 560);
			mainWindow.setResizable(false);
			mainWindow.setVisible(true);

			mainWindow.setIconImage(new ImageIcon(CatalogMain.class.getClassLoader().getResource("images/Icon.png")).getImage());

			// Set to center of the screen
			Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
			int framePosX = (screenSize.width - mainWindow.getWidth()) / 2;
			int framePosY = (screenSize.height - mainWindow.getHeight()) / 2;
			mainWindow.setLocation(framePosX, framePosY);

			mainWindow.getContentPane();
			mainWindow.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		}
	}

	public static boolean checkArgument(String[] args) {

		if(args.length == 0) {
			System.out.println("Default Language: English");
		}else if(args.length == 1) {

			String argument = args[0];

			switch(argument) {

			case "-e":
			case "E":System.out.println("Chosen Language: English");
			return true;

			case "-g":
			case "G":System.out.println("Chosen Language: German");
			return true;

			case "-f":
			case "F":System.out.println("Chosen Language: French");
			return true;

			default:
				System.out.println("Invalid arguments. -e: English, -g: German, -f: French");
				return false;
			}
		}else {
			System.out.println("Invalid arguments. Only one argument is allowed.");
			return false;
		}
		return true;
	}

	public static Properties getLAFProps() {
		return new Properties();
	}

	public static void setLookAndFeel() {

		try {			
			Properties props = getLAFProps();
			com.jtattoo.plaf.graphite.GraphiteLookAndFeel.setTheme(props);
			UIManager.setLookAndFeel("com.jtattoo.plaf.graphite.GraphiteLookAndFeel");
		} catch (Exception e) {
			JOptionPane.showMessageDialog(null, "An error has occurred (CatalogMain-UIManager). Application will now terminate.");
			System.exit(0);
		}
	}

	public void createNorthPanel() {

		northPanel = new JPanel();
		northPanel.setPreferredSize(new Dimension(600, 100));
		northPanel.setBorder(BorderFactory.createTitledBorder("<html><font size = 4> <font color=#0B612D>Select an Excel File (only .xlsx extension)</font color></font></html>"));

		pathTextField = new JTextField();
		pathTextField.setEditable(false);
		pathTextField.setPreferredSize(new Dimension(350, 25));

		openFileButton = new JButton(" ... ");
		openFileButton.addActionListener(new ButtonHandler());

		validateButton = new JButton(" Validate ");
		validateButton.addActionListener(new ButtonHandler());

		exitButton = new JButton( "      Exit      ");
		exitButton.addActionListener(new ButtonHandler());

		aboutButton = new JButton( "      About      ");
		aboutButton.addActionListener(new ButtonHandler());

		exportCSSButton = new JButton(" Export to CartoCSS ");
		exportCSSButton.addActionListener(new ButtonHandler());

		northPanel.add(pathTextField);
		northPanel.add(openFileButton);
		northPanel.add(validateButton);
		northPanel.add(exitButton);
		northPanel.add(aboutButton);
		northPanel.add(exportCSSButton);

		getProperty();

		chooser.setDialogTitle("Select an Excel File (only .xlsx extension)");
		chooser.setFileFilter(new FileNameExtensionFilter("Excel Files", "xlsx"));
	}

	public void createSouthPanel() {

		errorPane = new JTextPane();
		errorPane.setOpaque(false);
		kit = new HTMLEditorKit();
		doc = new HTMLDocument();
		errorPane.setEditorKit(kit);
		errorPane.setDocument(doc);
		errorPane.setEditable(false);
		errorPane.setSize(450, 450);
		errorPane.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

		JViewport viewport = new JViewport() {
			public void paintComponent(Graphics og) {
				super.paintComponent(og);
				Graphics2D g = (Graphics2D) og;
				g.setRenderingHint(RenderingHints.KEY_ANTIALIASING, RenderingHints.VALUE_ANTIALIAS_ON);
				GradientPaint gradient = new GradientPaint(0, 0, new Color(248, 248, 248), 0, getHeight(), Color.WHITE, true);
				g.setPaint(gradient);
				g.fillRoundRect(0, 0, getWidth(), getHeight(), 10, 10);
			}
		};
		viewport.add(errorPane);
		scrollPane = new JScrollPane();
		scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
		scrollPane.setViewport(viewport);
	}

	public void aboutWindow() {

		mainWindow.setEnabled(false);
		kit1 = new HTMLEditorKit();
		doc1 = new HTMLDocument();

		JFrame aboutFrame = new JFrame("");
		aboutFrame.setTitle("About");     
		aboutFrame.setSize(370, 460);      
		aboutFrame.setVisible(true);       
		aboutFrame.setResizable(false); 

		aboutFrame.setIconImage(new ImageIcon(CatalogMain.class.getClassLoader().getResource("images/Icon.png")).getImage());

		// Set to center of the screen
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		int framePosX = (screenSize.width - aboutFrame.getWidth()) / 2;
		int framePosY = (screenSize.height - aboutFrame.getHeight()) / 2;
		aboutFrame.setLocation(framePosX, framePosY);

		aboutPane = new JTextPane();
		aboutPane.setOpaque(true);
		aboutPane.setEditable(false);
		aboutPane.setSize(370, 460);
		aboutPane.setEditorKit(kit1);
		aboutPane.setDocument(doc1);
		aboutPane.setBorder(BorderFactory.createEmptyBorder(10, 10, 10, 10));

		try {
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2><b>VERSION</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>" + version + "</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2><b><br>MORE INFORMATION</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Website: http://wiki.hsr.ch/StefanKeller/PortrayalCatalogueValidator</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Feedback: sfkeller@hsr.ch</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Report a bug: sfkeller@hsr.ch\n</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2><b><br>DEVELOPERS</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Heng Zhi Feng</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Nguyen Duc Duong</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Stefan Keller</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2><b><br>PROJECT</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Intership 2013/14</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>HSR Hochschule für Technik Rapperswil</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Lead: Prof. Stefan Keller, Geometa Lab\n</font>", 0, 0,null);
			aboutPane.insertIcon(new ImageIcon(CatalogMain.class.getClassLoader().getResource("images/Logo.png")));
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2><b>CREDITS</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Specifications: KOGIS, IKGEO, Prof. Olivier Ertz HEIGVD</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Software Developement: Michael Rüegg, IFS</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Software Libraries: ProGuard, JTattoo 1.6.9, Apache POI 3.9\n</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2><b><br>LICENSE</font>", 0, 0,null);
			kit1.insertHTML(doc1, doc1.getLength(), "<font size = 2>Simplified BSD License (http://opensource.org/licenses/BSD-3-Clause)</font>", 0, 0,null);
		} catch (BadLocationException | IOException e1) {
			e1.printStackTrace();
		}

		aboutFrame.add(aboutPane, BorderLayout.CENTER);

		aboutFrame.addWindowListener(new WindowAdapter() {            
			@Override        
			public void windowClosing(WindowEvent e) {
				mainWindow.enableWindows();         
			}
		}); 
	}

	public void enableWindows() {
		mainWindow.setEnabled(true);    
	} 

	public void initializeRead() {

		workbook = null;
		try {
			workbook = WorkbookFactory.create(new FileInputStream(excelFilePath));

			if(!countOne){
				originalWorkbook = workbook;
				countOne = true;
			}
		} catch (InvalidFormatException | IOException e) {
			JOptionPane.showMessageDialog(null, "An error has occurred (CatalogMain-initializeRead). Application will now terminate.");
			System.exit(0);
		} 
	}

	public void createClassObject() {

		storeClassObjects.clear();
		Sheet layersSheet = workbook.getSheet("Layers");

		// Adds all rows as class objects
		for (int rowIndex = 4; rowIndex <= layersSheet.getLastRowNum(); rowIndex++) {

			Row row = layersSheet.getRow(rowIndex);

			if((row.getCell(1) != null && row.getCell(1).getCellType() != Cell.CELL_TYPE_BLANK) && (row.getCell(11) != null && row.getCell(11).getCellType() != Cell.CELL_TYPE_BLANK)) {

				classObjects = new LayersClassObject(row.getCell(0).toString(), row.getCell(1).toString(), row.getCell(2).toString(), String.valueOf(rowIndex+1), Double.parseDouble(row.getCell(11).toString()));

			}else if((row.getCell(1) != null && row.getCell(1).getCellType() != Cell.CELL_TYPE_BLANK) && (row.getCell(11) == null || row.getCell(11).getCellType() == Cell.CELL_TYPE_BLANK)) {

				classObjects = new LayersClassObject(row.getCell(0).toString(), row.getCell(1).toString(), row.getCell(2).toString(), String.valueOf(rowIndex+1), 1);

			}else if((row.getCell(1) == null || row.getCell(1).getCellType() == Cell.CELL_TYPE_BLANK) && (row.getCell(11) != null && row.getCell(11).getCellType() != Cell.CELL_TYPE_BLANK)) {

				classObjects = new LayersClassObject(row.getCell(0).toString(), null, row.getCell(2).toString(), String.valueOf(rowIndex+1), Double.parseDouble(row.getCell(11).toString()));
			}
			else {
				classObjects = new LayersClassObject(row.getCell(0).toString(), null, row.getCell(2).toString(), String.valueOf(rowIndex+1), 1);
			}
			storeClassObjects.add(classObjects);
		}

		// Sets true to objects which class has duplicates
		if(storeClassObjects.isEmpty() == false) {

			for(int count = 0; count < storeClassObjects.size(); count++) {

				String tempString = storeClassObjects.get(count).getClassName();

				for(int count1 = count+1; count1 < storeClassObjects.size(); count1++) {

					if(storeClassObjects.get(count1).getClassName().equalsIgnoreCase(tempString)) {
						storeClassObjects.get(count).setHaveSame(true);
						storeClassObjects.get(count1).setHaveSame(true);
					}
				}
			}
		}
	}

	public void getProperty() {

		try {			
			File configFile = new File("config.properties");
			if(configFile.exists()) {

				prop.load(new FileInputStream("config.properties"));

				if(prop.size() > 0) {

					String lastOpenDir = prop.getProperty("lastOpenDir");
					chooser = new JFileChooser(lastOpenDir);
				}
			}else {
				chooser = new JFileChooser();
			}
		} catch (IOException e) {
			JOptionPane.showMessageDialog(null, "An error has occurred (CatalogMain-getProperty). Application will now terminate.");
			System.exit(0);
		}
	}

	public void setUserPropertise() {

		try {
			prop.setProperty("lastOpenDir", fileDirectory);
			prop.store(new FileOutputStream("config.properties"), null);
		}catch (IOException e) {
			JOptionPane.showMessageDialog(null, "An error has occurred (CatalogMain-UserPropertise). Application will now terminate.");
			System.exit(0);
		}
	}

	public void beginValidate() {

		beginValidate = new BeginValidate(workbook, originalWorkbook, kit, doc);
		storeIsSheetsCorrect = beginValidate.startValidate();
	}

	public void beginExport() {

		setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));
		createClassObject();

		beginExport = new BeginExport(workbook, fileName, fileDirectory);
		boolean isFileExist = beginExport.checkFileExist();

		if(isFileExist == false) {

			List<String> storeInvalidFont = beginExport.exportNow(storeClassObjects);
			mainWindow.setEnabled(false);  
			String getSavePath = beginExport.getSaveDirectory();
			cartoReport = new ExportReport(mainWindow, getSavePath);
			beginExport.printExportReport(cartoReport, storeInvalidFont);
		}else {

			int response = JOptionPane.showConfirmDialog(null, "File already exist. Overwrite? (Yes/No)", "Confirmation", JOptionPane.YES_NO_OPTION);

			if(response == JOptionPane.YES_OPTION)  {

				List<String> storeInvalidFont = beginExport.exportNow(storeClassObjects);
				mainWindow.setEnabled(false);  
				String getSavePath = beginExport.getSaveDirectory();
				cartoReport = new ExportReport(mainWindow, getSavePath);
				beginExport.printExportReport(cartoReport, storeInvalidFont);
			}else {
				JOptionPane.showMessageDialog(null, "Export unsuccessful, file already exist!");
			}
		}
	}

	private class ButtonHandler implements ActionListener {

		public void actionPerformed(ActionEvent e) {

			if (e.getSource() == openFileButton) {

				int feedBack = chooser.showOpenDialog(null);

				if (feedBack == JFileChooser.OPEN_DIALOG) {

					excelFilePath = chooser.getSelectedFile().toString();
					fileDirectory = chooser.getCurrentDirectory().toString();
					fileName = chooser.getSelectedFile().getName();

					if(!oldExcelFilePath.equalsIgnoreCase(excelFilePath)) {
						countOne = false;
					}
					pathTextField.setText(excelFilePath);
				}	

			} else if (e.getSource() == validateButton) {

				errorPane.setText("");

				if (excelFilePath == null) {

					JOptionPane.showMessageDialog(null,"Please select an excel file first!");

				} else {

					setUserPropertise();
					oldExcelFilePath = excelFilePath;
					hasValidated =  true;
					String tempString = excelFilePath.substring(excelFilePath.length() - 5, excelFilePath.length());

					if (tempString.equalsIgnoreCase(".xlsx")) {

						DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
						Date date = new Date();

						setCursor(Cursor.getPredefinedCursor(Cursor.WAIT_CURSOR));

						initializeRead();

						try {
							kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Last validated: <font color=#088542>" + dateFormat.format(date) + "<br></font color></font>", 0, 0, null);
						} catch (IOException | BadLocationException e1) {
							JOptionPane.showMessageDialog(null, "An error has occurred (CatalogMain-HTMLkit). Application will now terminate.");
							System.exit(0);
						} 

						beginValidate();
						setCursor(Cursor.getDefaultCursor());
					} else {
						JOptionPane.showMessageDialog(null, "Could not process selected file. Did you select the right file?");
					}
				}
			}
			else if(e.getSource() == exitButton) {
				System.exit(0);
			}else if(e.getSource() == aboutButton) {
				aboutWindow();
			}else if(e.getSource() == exportCSSButton) {

				if(!hasValidated){
					JOptionPane.showMessageDialog(null, "Please validate the catalogue first.");
				}else if(storeIsSheetsCorrect.contains(false)){
					JOptionPane.showMessageDialog(null, "Please correct all errors first.");
				}else {
					beginExport();
				}
			}
		}
	}
}
