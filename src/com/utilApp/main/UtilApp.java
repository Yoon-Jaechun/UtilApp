package com.utilApp.main;
import java.awt.EventQueue;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.List;
import java.util.Vector;

import javax.swing.GroupLayout;
import javax.swing.GroupLayout.Alignment;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JMenu;
import javax.swing.JMenuBar;
import javax.swing.JMenuItem;
import javax.swing.JScrollPane;
import javax.swing.JTabbedPane;
import javax.swing.JTable;
import javax.swing.JTextArea;
import javax.swing.JTextField;
import javax.swing.LayoutStyle.ComponentPlacement;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.table.TableColumn;

import com.utilApp.java.ExcelController;
import javax.swing.JPanel;
import javax.swing.BoxLayout;

public class UtilApp {

	private JFrame frmMadeBy;
	private JTextField textField;
	private File selectedFile;
	ExcelController ec;
	private JTextArea textArea;
	private JTable createFolderTable;
	private JTable fetchFolderTable;
	private JTable fetchFileTable;
	private JTable renameFolderTable;
	private JTable renameFileTable;
	private JTable replaceConextTable;
	private JPanel panel;

	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UtilApp window = new UtilApp();
					window.frmMadeBy.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}

	/**
	 * Create the application.
	 */
	public UtilApp() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		frmMadeBy = new JFrame();
		frmMadeBy.setTitle("Made by Yoon, JaeChun - Ver 1.06(jre-7u10)");
		frmMadeBy.setBounds(100, 100, 752, 485);
		frmMadeBy.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		JMenuBar menuBar = new JMenuBar();
		frmMadeBy.setJMenuBar(menuBar);
		
		JMenu mnFile = new JMenu("File");
		menuBar.add(mnFile);
		
		JMenuItem mntmLoad = new JMenuItem("Load");
		mntmLoad.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				JFileChooser fileChooser = new JFileChooser();
				FileNameExtensionFilter filter = new FileNameExtensionFilter("please, select xls or xlsx file", "xls", "xlsx");
				fileChooser.setFileFilter(filter);
				
				int result = fileChooser.showOpenDialog(null);
				if(result == fileChooser.APPROVE_OPTION){
					selectedFile = fileChooser.getSelectedFile();
					textField.setText(selectedFile.getAbsolutePath());
					textArea.setText("");
					
					allLoadExcelInfo();
				}
			}
		});	
		mnFile.add(mntmLoad);
		
		JMenuItem mntmExit = new JMenuItem("Exit");
		mntmExit.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				System.exit(0);
			}
		});
		
		JMenuItem mntmReload = new JMenuItem("Reload");
		mntmReload.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if(selectedFile == null){
					textArea.setText("Again Load File");
				}else{
					allLoadExcelInfo();
				}
			}
		});
		mnFile.add(mntmReload);
		mnFile.add(mntmExit);
		
		JMenu mnNewMenu = new JMenu("Action");
		menuBar.add(mnNewMenu);
		
		JMenuItem mntmFetchFileList = new JMenuItem("Create Folder List");
		mntmFetchFileList.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if(selectedFile == null){
					textArea.setText("Again Load File");
				}else{
					try {
						List list = ec.createFolderList(selectedFile, 0);
						StringBuffer sb = new StringBuffer();
						for(int i = 0; i < list.size(); i++){
							sb.append(list.get(i));
							if(i != (list.size()-1)){
								sb.append("\n");
							}
						}
						allLoadExcelInfo();
						textArea.setText(sb.toString());
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		});
		mnNewMenu.add(mntmFetchFileList);
		
		JMenuItem mntmFatchFolderList = new JMenuItem("Fatch Folder List");
		mntmFatchFolderList.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if(selectedFile == null){
					textArea.setText("Again Load File");
				}else{
					try {
						List list = ec.fetchFolderList(selectedFile, 1);
						StringBuffer sb = new StringBuffer();
						for(int i = 0; i < list.size(); i++){
							sb.append(list.get(i));
							if(i != (list.size()-1)){
								sb.append("\n");
							}
						}
						allLoadExcelInfo();
						textArea.setText(sb.toString());
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		});
		mnNewMenu.add(mntmFatchFolderList);
		
		JMenuItem mntmFetchFileList_1 = new JMenuItem("Fetch File List");
		mntmFetchFileList_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if(selectedFile == null){
					textArea.setText("Again Load File");
				}else{
					try {
						List list = ec.fetchFileList(selectedFile, 2);
						StringBuffer sb = new StringBuffer();
						for(int i = 0; i < list.size(); i++){
							sb.append(list.get(i));
							if(i != (list.size()-1)){
								sb.append("\n");
							}
						}
						allLoadExcelInfo();
						textArea.setText(sb.toString());
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		});
		mnNewMenu.add(mntmFetchFileList_1);
		
		JMenuItem mntmRenameFile = new JMenuItem("Rename Folder List");
		mntmRenameFile.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if(selectedFile == null){
					textArea.setText("Again Load File");
				}else{
					try {
						List list = ec.renameFolderList(selectedFile, 3);
						StringBuffer sb = new StringBuffer();
						for(int i = 0; i < list.size(); i++){
							sb.append(list.get(i));
							if(i != (list.size()-1)){
								sb.append("\n");
							}
						}
						allLoadExcelInfo();
						textArea.setText(sb.toString());
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		});
		mnNewMenu.add(mntmRenameFile);
		
		JMenuItem mntmRenameFile_1 = new JMenuItem("Rename File List");
		mntmRenameFile_1.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if(selectedFile == null){
					textArea.setText("Again Load File");
				}else{
					try {
						List list = ec.renameFileList(selectedFile, 4);
						StringBuffer sb = new StringBuffer();
						for(int i = 0; i < list.size(); i++){
							sb.append(list.get(i));
							if(i != (list.size()-1)){
								sb.append("\n");
							}
						}
						allLoadExcelInfo();
						textArea.setText(sb.toString());
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		});
		mnNewMenu.add(mntmRenameFile_1);
		
		JMenuItem mntmNewMenuItem = new JMenuItem("Replace Context List");
		mntmNewMenuItem.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent arg0) {
				if(selectedFile == null){
					textArea.setText("Again Load File");
				}else{
					try {
						List list = ec.replaceConextFilList(selectedFile, 5);
						StringBuffer sb = new StringBuffer();
						for(int i = 0; i < list.size(); i++){
							sb.append(list.get(i));
							if(i != (list.size()-1)){
								sb.append("\n");
							}
						}
						allLoadExcelInfo();
						textArea.setText(sb.toString());
					} catch (Exception e) {
						// TODO Auto-generated catch block
						e.printStackTrace();
					}
				}
			}
		});
		mnNewMenu.add(mntmNewMenuItem);
		
		textField = new JTextField();
		textField.setColumns(10);
		
		JLabel lblLoadFile = new JLabel("Load File");
		
		JScrollPane scrollPane = new JScrollPane();
		
		JTabbedPane tabbedPane = new JTabbedPane(JTabbedPane.TOP);
		GroupLayout groupLayout = new GroupLayout(frmMadeBy.getContentPane());
		groupLayout.setHorizontalGroup(
			groupLayout.createParallelGroup(Alignment.LEADING)
				.addGroup(Alignment.TRAILING, groupLayout.createSequentialGroup()
					.addContainerGap()
					.addGroup(groupLayout.createParallelGroup(Alignment.TRAILING)
						.addComponent(scrollPane, Alignment.LEADING, GroupLayout.DEFAULT_SIZE, 712, Short.MAX_VALUE)
						.addComponent(tabbedPane, Alignment.LEADING, GroupLayout.DEFAULT_SIZE, 712, Short.MAX_VALUE)
						.addGroup(Alignment.LEADING, groupLayout.createSequentialGroup()
							.addComponent(lblLoadFile)
							.addGap(18)
							.addComponent(textField, GroupLayout.DEFAULT_SIZE, 642, Short.MAX_VALUE)))
					.addContainerGap())
		);
		groupLayout.setVerticalGroup(
			groupLayout.createParallelGroup(Alignment.LEADING)
				.addGroup(groupLayout.createSequentialGroup()
					.addContainerGap()
					.addGroup(groupLayout.createParallelGroup(Alignment.BASELINE)
						.addComponent(textField, GroupLayout.PREFERRED_SIZE, GroupLayout.DEFAULT_SIZE, GroupLayout.PREFERRED_SIZE)
						.addComponent(lblLoadFile))
					.addPreferredGap(ComponentPlacement.UNRELATED)
					.addComponent(tabbedPane, GroupLayout.PREFERRED_SIZE, 218, GroupLayout.PREFERRED_SIZE)
					.addPreferredGap(ComponentPlacement.UNRELATED)
					.addComponent(scrollPane, GroupLayout.DEFAULT_SIZE, 146, Short.MAX_VALUE)
					.addContainerGap())
		);
		
		panel = new JPanel();
		tabbedPane.addTab("Create Folder", null, panel, null);
		panel.setLayout(new BoxLayout(panel, BoxLayout.X_AXIS));
		
		JScrollPane scrollPane_1 = new JScrollPane();
		panel.add(scrollPane_1);
		
		createFolderTable = new JTable();
		createFolderTable.setModel(new DefaultTableModel(
			new Object[][] {
			},
			new String[] {
				"unload"
			}
		));
		createFolderTable.getColumnModel().getColumn(0).setPreferredWidth(119);
		scrollPane_1.setViewportView(createFolderTable);
		
		JPanel panel_2 = new JPanel();
		tabbedPane.addTab("Fetch Folder", null, panel_2, null);
		panel_2.setLayout(new BoxLayout(panel_2, BoxLayout.X_AXIS));
		
		JScrollPane scrollPane_2 = new JScrollPane();
		panel_2.add(scrollPane_2);
		
		fetchFolderTable = new JTable();
		fetchFolderTable.setModel(new DefaultTableModel(
			new Object[][] {
			},
			new String[] {
				"unload"
			}
		));
		fetchFolderTable.getColumnModel().getColumn(0).setPreferredWidth(144);
		scrollPane_2.setViewportView(fetchFolderTable);
		
		JPanel panel_3 = new JPanel();
		tabbedPane.addTab("Fetch File", null, panel_3, null);
		panel_3.setLayout(new BoxLayout(panel_3, BoxLayout.X_AXIS));
		
		JScrollPane scrollPane_3 = new JScrollPane();
		panel_3.add(scrollPane_3);
		
		fetchFileTable = new JTable();
		fetchFileTable.setModel(new DefaultTableModel(
			new Object[][] {
			},
			new String[] {
				"unload"
			}
		));
		scrollPane_3.setViewportView(fetchFileTable);
		
		JPanel panel_4 = new JPanel();
		tabbedPane.addTab("Rename Folder", null, panel_4, null);
		panel_4.setLayout(new BoxLayout(panel_4, BoxLayout.X_AXIS));
		
		JScrollPane scrollPane_4 = new JScrollPane();
		panel_4.add(scrollPane_4);
		
		renameFolderTable = new JTable();
		renameFolderTable.setModel(new DefaultTableModel(
			new Object[][] {
			},
			new String[] {
				"unload"
			}
		));
		scrollPane_4.setViewportView(renameFolderTable);
		
		JPanel panel_5 = new JPanel();
		tabbedPane.addTab("Rename File", null, panel_5, null);
		panel_5.setLayout(new BoxLayout(panel_5, BoxLayout.X_AXIS));
		
		JScrollPane scrollPane_5 = new JScrollPane();
		panel_5.add(scrollPane_5);
		
		renameFileTable = new JTable();
		renameFileTable.setModel(new DefaultTableModel(
			new Object[][] {
			},
			new String[] {
				"unload"
			}
		));
		scrollPane_5.setViewportView(renameFileTable);
		
		JPanel panel_6 = new JPanel();
		tabbedPane.addTab("Replace Context", null, panel_6, null);
		panel_6.setLayout(new BoxLayout(panel_6, BoxLayout.X_AXIS));
		
		JScrollPane scrollPane_6 = new JScrollPane();
		panel_6.add(scrollPane_6);
		
		replaceConextTable = new JTable();
		replaceConextTable.setModel(new DefaultTableModel(
			new Object[][] {
			},
			new String[] {
				"unload"
			}
		));
		replaceConextTable.getColumnModel().getColumn(0).setPreferredWidth(83);
		scrollPane_6.setViewportView(replaceConextTable);
		
		textArea = new JTextArea();
		scrollPane.setViewportView(textArea);
		frmMadeBy.getContentPane().setLayout(groupLayout);
		ec = new ExcelController();
	}
	
	public void loadExcelInfo(String colunm, File file, int sheetIndex, JTable jTable){
		DefaultTableModel model = new DefaultTableModel();
		model.addColumn(colunm);
		
		try {
			List<String> list = ec.getLoadExcelOneColumnList(file, sheetIndex);
			for(int i = 0; i < list.size(); i++){
				Vector<String> vector = new Vector<String>();
				vector.addElement(list.get(i));
				model.insertRow(i, vector);
			}
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
			
		jTable.setModel(model);
	}
	
	public void loadExcelInfo(String column1, String column2, File file, int sheetIndex, JTable jTable){
		DefaultTableModel model = new DefaultTableModel();
		model.addColumn(column1);
		model.addColumn(column2);
		
		try {
			List<Object[]> list = ec.getLoadExcelTwoColumnList(file, sheetIndex);
			for(int i = 0; i < list.size(); i++){
				model.insertRow(i, list.get(i));
			}
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
			
		jTable.setModel(model);
	}
	
	public void loadExcelInfo(String column1, String column2, String column3, String column4, File file, int sheetIndex, JTable jTable){
		DefaultTableModel model = new DefaultTableModel();
		model.addColumn(column1);
		model.addColumn(column2);
		model.addColumn(column3);
		model.addColumn(column4);
		
		try {
			List<Object[]> list = ec.getLoadExcelFourColumnList(file, sheetIndex);
			for(int i = 0; i < list.size(); i++){
				model.insertRow(i, list.get(i));
			}
		} catch (Exception e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
			
		jTable.setModel(model);
	}
	
	public void allLoadExcelInfo(){
		loadExcelInfo("Loaded Root Folder for Creating Folder", selectedFile, 0, createFolderTable);
		loadExcelInfo("Loaded Root Folder for Fetching Folder", selectedFile, 1, fetchFolderTable);
		loadExcelInfo("Loaded Root Folder for Fetching File", selectedFile, 2, fetchFileTable);
		loadExcelInfo("Loaded Original Folder", "Loaded Destination Folder", selectedFile, 3, renameFolderTable);
		loadExcelInfo("Loaded Original File", "Loaded Destination File", selectedFile, 4, renameFileTable);
		loadExcelInfo("Loaded File", "Before Replace Conext", "After Replace Conext", "Encoding",  selectedFile, 5, replaceConextTable);
	}
}
