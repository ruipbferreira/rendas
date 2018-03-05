import static utils.PropertiesUtils.readPropertiesToExcel;
import static utils.PropertiesUtils.readPropertiesToWord;

import java.awt.GridLayout;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.Date;
import java.util.GregorianCalendar;
import java.util.List;
import java.util.Map;

import javax.swing.JButton;
import javax.swing.JFileChooser;
import javax.swing.JFrame;
import javax.swing.JOptionPane;
import javax.swing.JPanel;
import javax.swing.SwingUtilities;
import javax.swing.UIManager;
import javax.swing.UnsupportedLookAndFeelException;

import org.jdatepicker.JDatePicker;

import model.Fraction;
import utils.ExcelUtils;
import utils.WordUtils;

public class GenerateFile extends JFrame {

    private JFileChooser propertiesFile;
    private JButton properties;
    private JFileChooser templateChoser;
    private JButton template;
    private JFileChooser excelChoser;
    private JButton excel;
    private JFileChooser outputChoser;
    private JButton output;
    private JButton generate;
    private JDatePicker datePicker;
    private File proprertiesFile;
    private File templateFile;
    private File excelFile;
    private File outputFile;
    private Map<String, String> propsExcel;
    private Map<String, String> propsWord;
    private Date chooserDate;

    public GenerateFile() {
        JPanel panel = new JPanel();
        GridLayout layout = new GridLayout(0,1);
        panel.setLayout(layout);
        this.getContentPane().add(panel);

        datePicker = new JDatePicker();
        panel.add(datePicker);

        propertiesFile = new JFileChooser();
        propertiesFile.setCurrentDirectory(new java.io.File("."));
        properties = new JButton("Propriedades");
        panel.add(properties);

        templateChoser = new JFileChooser();
        templateChoser.setCurrentDirectory(new java.io.File("."));
        template = new JButton("Template");
        panel.add(template);

        excelChoser = new JFileChooser();
        excelChoser.setCurrentDirectory(new java.io.File("."));
        excel = new JButton("Excel");
        panel.add(excel);
        
        outputChoser = new JFileChooser();
        outputChoser.setCurrentDirectory(new java.io.File("."));
        outputChoser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);
        outputChoser.setAcceptAllFileFilterUsed(false);
        output = new JButton("Destino");
        panel.add(output);

        generate = new JButton("Gerar");
        panel.add(generate);
        addListeners();
    }

    private void addListeners() {
        excel.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                int returnValue = excelChoser.showOpenDialog(GenerateFile.this);
                if(returnValue == JFileChooser.APPROVE_OPTION) {
                    excelFile = excelChoser.getSelectedFile();
                }
            }
        });
        
        output.addActionListener(new ActionListener() {
			
			public void actionPerformed(ActionEvent e) {
				int returnValue = outputChoser.showOpenDialog(GenerateFile.this);
				if(returnValue == JFileChooser.APPROVE_OPTION) {
					outputFile = outputChoser.getSelectedFile();
                }
			}
		});

        template.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                int returnVal = templateChoser.showOpenDialog(GenerateFile.this);
                if(returnVal == JFileChooser.APPROVE_OPTION) {
                    templateFile = templateChoser.getSelectedFile();
                }
            }
        });

        properties.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                int returnVal = propertiesFile.showOpenDialog(GenerateFile.this);
                if(returnVal == JFileChooser.APPROVE_OPTION) {
                    proprertiesFile = propertiesFile.getSelectedFile();
                    try {
                        propsExcel = readPropertiesToExcel(proprertiesFile);
                        propsWord = readPropertiesToWord(proprertiesFile);
                    } catch (Exception e1) {
                        JOptionPane.showMessageDialog(GenerateFile.this, "Erro ao ler o ficheiro " + proprertiesFile.getName(),"Erro", JOptionPane.ERROR_MESSAGE);
                    }
                }

            }
        });

        generate.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                if(propertiesFile == null || templateFile == null || excelFile == null || outputFile == null) {
                    JOptionPane.showMessageDialog(GenerateFile.this, "Devem ser escolhidos os quatro ficheiros", "Erro", JOptionPane.ERROR_MESSAGE);
                } else {
                    chooserDate = datePicker.getModel().getValue() != null ? 
                    		((GregorianCalendar)datePicker.getModel().getValue()).getTime() : new GregorianCalendar().getTime();
                    try {
                        List<Fraction> fractionsFromFile = ExcelUtils.getFractionsFromFile(excelFile, propsExcel);
                        if(!fractionsFromFile.isEmpty()) {
                        	WordUtils.generateWord(fractionsFromFile, templateFile, outputFile.getAbsolutePath(), propsWord);
                        }
                        System.exit(EXIT_ON_CLOSE);
                    } catch (Exception e1) {
                    	e1.printStackTrace();
                        JOptionPane.showMessageDialog(GenerateFile.this,
                                e1.getMessage(), "Erro", JOptionPane.ERROR_MESSAGE);
                    }
                }

            }
        });

    }

    private static void createAndShowUI() {
        JFrame frame = new GenerateFile();
        frame.setTitle("Rendas 1.0");
        frame.setSize(200, 200);
        frame.setLocationRelativeTo(null);
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setResizable(false);
        frame.setVisible(true);
    }



    public static void main(String[] args) {
        /* Use an appropriate Look and Feel */
        try {
            //UIManager.setLookAndFeel("com.sun.java.swing.plaf.windows.WindowsLookAndFeel");
            UIManager.setLookAndFeel("javax.swing.plaf.metal.MetalLookAndFeel");
        } catch (UnsupportedLookAndFeelException ex) {
            ex.printStackTrace();
        } catch (IllegalAccessException ex) {
            ex.printStackTrace();
        } catch (InstantiationException ex) {
            ex.printStackTrace();
        } catch (ClassNotFoundException ex) {
            ex.printStackTrace();
        }
        /* Turn off metal's use of bold fonts */
        UIManager.put("swing.boldMetal", Boolean.FALSE);

        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                createAndShowUI();
            }
        });
    }
}