import java.io.*;
import java.text.*;
import java.util.*;
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

public class catalogMain extends JFrame {

	private static catalogMain mainWindow;
	private boolean countOne = false;
	private boolean hasValidated = false;
	private String excelFilePath;
	private String oldExcelFilePath = "";
	private JFileChooser chooser;
	private JPanel northPanel;
	private JButton openFileButton;
	private JButton validateButton;
	private JButton exportCSSButton;
	private JTextField pathTextField;
	private JTextPane errorPane;
	private JScrollPane scrollPane;
	private ArrayList<Boolean> storeIsSheetsCorrect= new ArrayList<Boolean>();
	private HTMLEditorKit kit;
	private HTMLDocument doc;
	private beginValidate beginValidate;
	private beginExport beginExport;
	
	// ApachePOI (reading of excel)
	private Workbook workbook;
	private Workbook originalWorkbook;
	
	public catalogMain() {

		createNorthPanel();
		createSouthPanel();
		
		getContentPane().add(northPanel, BorderLayout.NORTH);
		getContentPane().add(scrollPane, BorderLayout.CENTER);
	}
	
	public static void main(String[] args) {

		try {
			UIManager.setLookAndFeel("com.jtattoo.plaf.texture.TextureLookAndFeel");
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}

		mainWindow = new catalogMain();
		mainWindow.setTitle("Portrayal Catalogue Validator");
		mainWindow.setSize(530, 560);
		mainWindow.setResizable(false);
		mainWindow.setVisible(true);

		// Set to center of the screen
		Dimension screenSize = Toolkit.getDefaultToolkit().getScreenSize();
		int framePosX = (screenSize.width - mainWindow.getWidth()) / 2;
		int framePosY = (screenSize.height - mainWindow.getHeight()) / 2;
		mainWindow.setLocation(framePosX, framePosY);

		mainWindow.getContentPane();
		mainWindow.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
	}
	
	public void createNorthPanel() {

		northPanel = new JPanel();
		northPanel.setPreferredSize(new Dimension(600, 100));
		northPanel.setBorder(BorderFactory.createTitledBorder("<html><font size = 4> <font color=#0B612D>Select an Excel File (only .xlsx extension)</font color></font></html>"));

		pathTextField = new JTextField();
		pathTextField.setEditable(false);
		pathTextField.setPreferredSize(new Dimension(400, 30));

		openFileButton = new JButton(" ... ");
		openFileButton.addActionListener(new ButtonHandler());

		validateButton = new JButton(" Validate ");
		validateButton.addActionListener(new ButtonHandler());
		
		exportCSSButton = new JButton(" Export to CartoCSS ");
		exportCSSButton.addActionListener(new ButtonHandler());

		northPanel.add(pathTextField);
		northPanel.add(openFileButton);
		northPanel.add(validateButton);
		northPanel.add(exportCSSButton);

		chooser = new JFileChooser();
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
				GradientPaint gradient = new GradientPaint(0, 0, new Color(247, 237, 204), 0, getHeight(), Color.WHITE, true);
				g.setPaint(gradient);
				g.fillRoundRect(0, 0, getWidth(), getHeight(), 50, 50);
			}
		};
		viewport.add(errorPane);
		scrollPane = new JScrollPane();
		scrollPane.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_AS_NEEDED);
		scrollPane.setViewport(viewport);
	}
	
	public void enableWindows() {
		mainWindow.setEnabled(true);    
	} 
	
	public void initializeRead() {

		workbook = null;

		try {
			workbook = WorkbookFactory.create(new FileInputStream(excelFilePath));
			
			if(countOne == false){
				originalWorkbook = workbook;
				countOne = true;
			}
		} catch (InvalidFormatException | IOException e) {
			e.printStackTrace();
		} 
	}
	
	public void beginValidate() {
		
		beginValidate = new beginValidate(workbook, originalWorkbook, kit, doc);
		storeIsSheetsCorrect = beginValidate.startValidate();
	}
	
	public void beginExport() {
		
		// Display export report
		mainWindow.setEnabled(false);   
		exportReport cartoReport = new exportReport(mainWindow);
		
		beginExport = new beginExport(workbook, cartoReport);
		beginExport.startExport();
	}
	
	private class ButtonHandler implements ActionListener {

		public void actionPerformed(ActionEvent e) {

			if (e.getSource() == openFileButton) {

				int feedBack = chooser.showOpenDialog(null);

				if (feedBack == JFileChooser.OPEN_DIALOG) {
					
					excelFilePath = chooser.getSelectedFile().toString();
					
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
					
					oldExcelFilePath = excelFilePath;
					hasValidated =  true;
					String tempString = excelFilePath.substring(excelFilePath.length() - 5, excelFilePath.length());

					if (tempString.equalsIgnoreCase(".xlsx")) {
									
						DateFormat dateFormat = new SimpleDateFormat("dd/MM/yyyy HH:mm:ss");
						Date date = new Date();

						initializeRead();

						try {
							kit.insertHTML(doc, doc.getLength(), "<font size = 3> <font color=#0A23C4>Last validated: <font color=#088542>" + dateFormat.format(date) + "<br></font color></font>", 0, 0, null);
						} catch (IOException | BadLocationException e1) {
							e1.printStackTrace();
						} 
						
						beginValidate();
					} else {
						JOptionPane.showMessageDialog(null, "Could not process selected file. Did you select the right file?");
					}
				}
			} else if(e.getSource() == exportCSSButton) {
				
				if(hasValidated == false){
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
