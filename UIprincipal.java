import java.awt.BorderLayout;
import java.awt.EventQueue;
import java.awt.Font;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JLabel;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.JRadioButton;
import javax.swing.JScrollPane;
import javax.swing.JTable;
import javax.swing.JTextField;
import javax.swing.SwingConstants;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.table.DefaultTableModel;
import javax.swing.JTextPane;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class UIprincipal {
	

	private JFrame UIPrincipalFrame;
	
	int contador = 0; 
	int fin = 0;
	int numerador = 1; 
	int table = 1; 
	
	String[] columnas= {"#", "Descripción:","Solicitante:", "Cantidad:", "Transferido:" };
	String transferido = "";
	String descripcion = "";
	String solicitante = "";
	String cantidad = "";
	String [][] data = {};
	
	JButton BTotal = new JButton("TOTAL:");
	JButton SaveExcel = new JButton("GUARDAR EN EXCEL");
	
	private JTable UItable;
	private DefaultTableModel modelo;
	private DefaultTableModel model;
	JPanel UIpanel = new JPanel();
	
	private JTextField textDesc;
	private JTextField textSol;
	private JTextField textCant;
	/**
	 * Launch the application.
	 */
	public static void main(String[] args) {
		EventQueue.invokeLater(new Runnable() {
			public void run() {
				try {
					UIprincipal window = new UIprincipal();
					window.UIPrincipalFrame.setVisible(true);
				} catch (Exception e) {
					e.printStackTrace();
				}
			}
		});
	}


	public UIprincipal() {
		initialize();
	}

	/**
	 * Initialize the contents of the frame.
	 */
	private void initialize() {
		UIPrincipalFrame = new JFrame();
		UIPrincipalFrame.setTitle("DIVEFAC - Registro diario");
		UIPrincipalFrame.setAlwaysOnTop(true);
		UIPrincipalFrame.setBounds(100, 100, 408, 548);
		UIPrincipalFrame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
		
		UIPrincipalFrame.getContentPane().add(UIpanel, BorderLayout.CENTER);
		UIpanel.setBounds(100, 100, 407, 495);
		UIpanel.setLayout(null);
		
		JLabel ControlLabel = new JLabel("CONTROL DE TRANSFERENCIAS");
		ControlLabel.setHorizontalAlignment(SwingConstants.CENTER);
		ControlLabel.setFont(new Font("Microsoft Tai Le", Font.BOLD, 20));
		ControlLabel.setBounds(37, 10, 335, 45);
		UIpanel.add(ControlLabel);
		
		JLabel Registro = new JLabel("Por favor, registre la entrega en efectivo.");
		Registro.setHorizontalAlignment(SwingConstants.CENTER);
		Registro.setFont(new Font("Microsoft YaHei UI Light", Font.BOLD, 15));
		Registro.setBounds(26, 32, 358, 45);
		UIpanel.add(Registro);
		
		modelo = new DefaultTableModel(data, columnas);
		UItable = new JTable(modelo);
		UItable.setBounds(10, 214, 374, 169);
		
		JRadioButton RadiTransf = new JRadioButton(": Esta transferido");
		RadiTransf.setBounds(139, 159, 154, 21);
		RadiTransf.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		UIpanel.add(RadiTransf);
		
		JScrollPane UIscroll = new JScrollPane(UItable);
		UIscroll.setVerticalScrollBarPolicy(JScrollPane.VERTICAL_SCROLLBAR_ALWAYS); 
		UIpanel.add(UIscroll);
		UIscroll.setBounds(10, 217, 374, 169);
		
		JButton GuardaBT = new JButton("GUARDAR");
		GuardaBT.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				/**
				 *Verificamos los datos del JtexField.
				 */
				String texto=textCant.getText();
		        texto=texto.replaceAll(" ", "");
		        if(texto.length()==0){
		        	JOptionPane.showMessageDialog(UIpanel, "Este es un mensaje de Advertencia\nNo hay datos, por favor registra la información!.",
		        			  "Mensaje de error.", JOptionPane.WARNING_MESSAGE);
		        
		        } else {
		        	
		        	int resp=JOptionPane.showConfirmDialog(UIpanel,"¿Estás seguro de guardar los datos.?","Guardar", JOptionPane.YES_NO_OPTION);
		            if (JOptionPane.OK_OPTION == resp){

					/**
					 *Verificamos los datos del JtexField.
					 */
   
				/**
				 * Extraemos los datos del JtexField.
				 */
				if (RadiTransf.isSelected() == true) {transferido = "Si";} 
				else { transferido = "No";}
				descripcion = textDesc.getText();
				solicitante = textSol.getText();
				cantidad = textCant.getText();
				/**
				 * Guardamos los datos del JtexField.
				 */
				Object[] row = {numerador + "", descripcion +"",solicitante + "", "$ " + cantidad,transferido +  ""};
				model = (DefaultTableModel) UItable.getModel();
				model.addRow(row);
				/**
				 * Incrementamos el contador..
				 */
				numerador ++;
				table++;
				contador= Integer.parseInt(textCant.getText());
				fin= contador + fin;
				
				/**
				 * Delete text de JtexField.
				 */
				textCant.setText(null);
				textDesc.setText(null);
				textSol.setText(null);
				       }
		            else{
						/**
						 * Variable vacia..
						 */
		         }
		        }
		       }
		});
		GuardaBT.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		GuardaBT.setBounds(37, 186, 154, 27);
		UIpanel.add(GuardaBT);
		
		textDesc = new JTextField();
		textDesc.setColumns(10);
		textDesc.setBounds(171, 75, 183, 22);
		textDesc.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		UIpanel.add(textDesc);
		
		JLabel LDesc = new JLabel("Descripción:");
		LDesc.setHorizontalAlignment(SwingConstants.CENTER);
		LDesc.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		LDesc.setBounds(45, 75, 137, 19);
		UIpanel.add(LDesc);
		
		JLabel LSol = new JLabel("Solicitante:");
		LSol.setHorizontalAlignment(SwingConstants.CENTER);
		LSol.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		LSol.setBounds(45, 104, 137, 19);
		UIpanel.add(LSol);
		
		textSol = new JTextField();
		textSol.setColumns(10);
		textSol.setBounds(171, 104, 183, 22);
		textSol.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		UIpanel.add(textSol);
		
		JLabel LCant = new JLabel("Cantidad:");
		LCant.setHorizontalAlignment(SwingConstants.CENTER);
		LCant.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		LCant.setBounds(100, 134, 137, 19);
		UIpanel.add(LCant);

		textCant = new JTextField();
		textCant.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		textCant.setColumns(10);
		textCant.setBounds(210, 134, 70, 22);
		UIpanel.add(textCant);
		
		JTextPane textPane = new JTextPane();
		textPane.setEditable(false);
		textPane.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		textPane.setBounds(242, 401, 142, 27);
		UIpanel.add(textPane);
		
		JButton Limpiar = new JButton("LIMPIAR");
		Limpiar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				
                UIpanel.revalidate();
                UIpanel.repaint();
                
                model.setRowCount(0);
            	contador = 0; 
            	fin = 0;
            	numerador = 1; 
            	table = 1; 
            	data  = new String[0][0];
                


				
			}
		});
		Limpiar.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		Limpiar.setBounds(218, 186, 154, 27);
		UIpanel.add(Limpiar);
		
		BTotal.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
                
		        if(UItable.getRowCount() == 0){
		        	JOptionPane.showMessageDialog(UIpanel, "Este es un mensaje de advertencia.\nNo hay datos, por favor registrar la información!.",
		        			  "Mensaje de error.", JOptionPane.WARNING_MESSAGE);
		        
		        } else {
				textPane.setText("$ "+fin);
				
			}
			}
		});
		BTotal.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		BTotal.setBounds(37, 401, 200, 27);
		UIpanel.add(BTotal);
		SaveExcel.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
		        if(UItable.getRowCount() == 0){
		        	JOptionPane.showMessageDialog(UIpanel, "Este es un mensaje de advertencia.\nNo hay datos, por favor registrar la información!.",
		        			  "Mensaje de error.", JOptionPane.WARNING_MESSAGE);
		        
		        } else {
		        	
					Object[] f = {"\n","", "Total:" + "","$ "+fin, ""};
					model.addRow(f);
					saveTableToExcel(UItable);					
					

	                UIpanel.revalidate();
	                UIpanel.repaint();
			}
			}
		});
		SaveExcel.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		SaveExcel.setBounds(120, 440, 183, 27);
		UIpanel.add(SaveExcel);
		
		JButton Importar = new JButton("IMPORTAR DESDE EXCEL");
		Importar.addActionListener(new ActionListener() {
			public void actionPerformed(ActionEvent e) {
				model = (DefaultTableModel) UItable.getModel();
				importarDesdeExcel();
			}
		});
		Importar.setFont(new Font("Microsoft YaHei", Font.BOLD, 14));
		Importar.setBounds(100, 477, 234, 27);
		UIpanel.add(Importar);
		
		
	}
	   private void importarDesdeExcel() {
	        JFileChooser fileChooser = new JFileChooser();
	        fileChooser.setFileFilter(new FileNameExtensionFilter("Archivos Excel (*.xlsx)", "xlsx"));

	        int option = fileChooser.showSaveDialog(UIpanel);

	        if (option == JFileChooser.APPROVE_OPTION) {
	            try (Workbook workbook = WorkbookFactory.create(fileChooser.getSelectedFile())) {
	                Sheet sheet = workbook.getSheetAt(0);

	      
	                model.setRowCount(0);

	                for (Row row : sheet) {
	                    Object[] rowData = new Object[row.getLastCellNum()];

	                    for (int i = 0; i < row.getLastCellNum(); i++) {
	                        Cell cell = row.getCell(i, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
	                        rowData[i] = cell.toString();
	                    }

	                    model.addRow(rowData);
	                }
                    JOptionPane.showMessageDialog(UIpanel, "¡Datos importados orrectamente desde Excel!");


	            } catch (IOException e) {
	                e.printStackTrace();
	                JOptionPane.showMessageDialog(UIpanel, "Error al importar desde Excel", "Error", JOptionPane.ERROR_MESSAGE);
	            }
	        }
	    }

	private void saveTableToExcel(JTable UItable) {
		
        JFileChooser fileChooser = new JFileChooser();
        fileChooser.setFileFilter(new FileNameExtensionFilter("Archivos Excel (*.xlsx)", "xlsx"));
        int option =  fileChooser.showSaveDialog(UIpanel);
        if (option == JFileChooser.APPROVE_OPTION) {
            try (Workbook workbook = new XSSFWorkbook()) {
                Sheet sheet = workbook.createSheet("Datos");


                Row headerRow = sheet.createRow(0);
                for (int i = 0; i < UItable.getColumnCount(); i++) {
                    Cell cell = headerRow.createCell(i);
                    cell.setCellValue(UItable.getColumnName(i));
                }


                for (int i = 0; i < UItable.getRowCount(); i++) {
                    Row row = sheet.createRow(i + 1);
                    for (int j = 0; j < UItable.getColumnCount(); j++) {
                        Cell cell = row.createCell(j);
                        cell.setCellValue(String.valueOf(UItable.getValueAt(i, j)));
                    }
                }


                try (FileOutputStream fileOut = new FileOutputStream(fileChooser.getSelectedFile() + ".xlsx")) {
                    workbook.write(fileOut);
                    JOptionPane.showMessageDialog(UIpanel, "¡Se ha guardado correctamente en Excel!");
                } catch (IOException ex) {
                    ex.printStackTrace();
                }
            } catch (IOException ex) {
                ex.printStackTrace();
            }
            
        }
        
        
    }
}
