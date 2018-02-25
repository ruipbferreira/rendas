import model.Fraction;
import org.jdatepicker.JDatePicker;
import utils.ExcelUtils;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.util.Date;
import java.util.List;
import java.util.Properties;

import static utils.PropertiesUtils.readProperties;

public class GenerateFile extends JFrame {

    private JFileChooser propertiesFile;
    private JButton properties;
    private JFileChooser templateChoser;
    private JButton template;
    private JFileChooser excelChoser;
    private JButton excel;
    private JButton generate;
    private JDatePicker datePicker;
    private File proprertiesFile;
    private File templateFile;
    private File excelFile;
    private Properties props;
    private Date chooserDate;

    public GenerateFile() {
        JPanel panel = new JPanel();
        GridLayout layout = new GridLayout(0,1);
        panel.setLayout(layout);
        this.getContentPane().add(panel);

        datePicker = new JDatePicker();
        panel.add(datePicker);

        propertiesFile = new JFileChooser();
        properties = new JButton("Propriedades");
        panel.add(properties);

        templateChoser = new JFileChooser();
        template = new JButton("Template");
        panel.add(template);

        excelChoser = new JFileChooser();
        excel = new JButton("Excel");
        panel.add(excel);

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
                        props = readProperties(proprertiesFile);
                    } catch (Exception e1) {
                        JOptionPane.showMessageDialog(GenerateFile.this, "Erro ao ler o ficheiro " + proprertiesFile.getName(),"Erro", JOptionPane.ERROR_MESSAGE);
                    }
                }

            }
        });

        generate.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                if(propertiesFile == null || templateFile == null || excelFile == null) {
                    JOptionPane.showMessageDialog(GenerateFile.this, "Devem ser escolhidos os três ficheiros", "Erro",JOptionPane.ERROR_MESSAGE);
                } else {
                    chooserDate = (Date)datePicker.getModel().getValue();
                    try {
                        List<Fraction> fractionsFromFile = ExcelUtils.getFractionsFromFile(excelFile, props);
                        for(Fraction f : fractionsFromFile) {
                            System.out.println(f.getFractionCode());
                        }
                    } catch (Exception e1) {
                        JOptionPane.showMessageDialog(GenerateFile.this,
                                e1.getMessage(), "Erro",JOptionPane.ERROR_MESSAGE);
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