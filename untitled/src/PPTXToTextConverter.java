import org.apache.poi.xslf.usermodel.*;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;
import javax.swing.filechooser.FileNameExtensionFilter;
import static java.awt.Color.*;
import static javax.swing.JLabel.*;

public class PPTXToTextConverter {
    private JFileChooser fileChooser;
    private String selectedOutputDirectory = "";
    private JFrame frame;
    private JLabel inputFileLabel;
    private JTextField inputFileField;
    private JButton browseButton;
    private JLabel outputFileLabel;
    private JTextField outputFileField;
    private JButton outputBrowseButton;
    private JButton convertButton;
    private JButton exitButton;

    public PPTXToTextConverter() {

        frame = new JFrame("Конвертер");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(new GridLayout(6, 1));

        inputFileLabel = new JLabel("Выберите PPTX файл:");
        inputFileField = new JTextField();
        browseButton = new JButton("Обзор");

        // ActionListener для кнопки "Обзор"
        this.browseButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                fileChooser.setMultiSelectionEnabled(true);
                int result = fileChooser.showOpenDialog(PPTXToTextConverter.this.frame);
                if (result == JFileChooser.APPROVE_OPTION) {
                    File[] selectedFiles = fileChooser.getSelectedFiles();
                    StringBuilder filesText = new StringBuilder();
                    for (File file : selectedFiles) {
                        filesText.append(file.getAbsolutePath()).append("; ");
                    }
                    PPTXToTextConverter.this.inputFileField.setText(filesText.toString());
                }
            }
        });

        outputFileLabel = new JLabel("Выберите путь для сохранения результата:");
        outputFileField = new JTextField();
        outputBrowseButton = new JButton("Сохранить");

        fileChooser = new JFileChooser();
        fileChooser.setFileSelectionMode(JFileChooser.FILES_ONLY);
        fileChooser.setAcceptAllFileFilterUsed(false);
        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("Текстовые файлы (*.txt)", "txt"));
        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("Документы (*.doc)", "doc", "docx"));
        fileChooser.addChoosableFileFilter(new FileNameExtensionFilter("HTML файлы (*.html)", "html", "htm"));

        // ActionListener для кнопки "Сохранить"
        this.outputBrowseButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {

                if (!selectedOutputDirectory.isEmpty()) {
                    fileChooser.setSelectedFile(new File(selectedOutputDirectory));
                }

                int result = fileChooser.showSaveDialog(PPTXToTextConverter.this.frame);
                if (result == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    FileNameExtensionFilter chosenFilter = (FileNameExtensionFilter) fileChooser.getFileFilter();
                    selectedOutputDirectory = selectedFile.getAbsolutePath();

                    PPTXToTextConverter.this.outputFileField.setText(selectedOutputDirectory);
                }
            }
        });

        convertButton = new JButton("Конвертировать");

        // ActionListener для кнопки "Конвертировать"
        this.convertButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                String[] inputFiles = PPTXToTextConverter.this.inputFileField.getText().split(";\\s*");
                String outputDirectory = PPTXToTextConverter.this.outputFileField.getText();

                if (inputFiles.length == 0 || inputFiles[0].isEmpty()) {
                    JOptionPane.showMessageDialog(PPTXToTextConverter.this.frame, "Выберите хотя бы один входной файл", "Нет файлов для конвертации", JOptionPane.WARNING_MESSAGE);
                } else if (outputDirectory.isEmpty()) {
                    JOptionPane.showMessageDialog(PPTXToTextConverter.this.frame, "Выберите папку для сохранения результатов", "Нет директории для сохранения", JOptionPane.WARNING_MESSAGE);
                } else {
                    FileNameExtensionFilter chosenFilter = (FileNameExtensionFilter) fileChooser.getFileFilter();
                    PPTXToTextConverter.this.convertPPTXToText(inputFiles, outputDirectory, chosenFilter);
                }
            }
        });

        exitButton = new JButton("Выйти");

        // ActionListener для кнопки "Выйти"
        exitButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                System.exit(0);
            }
        });

        Font font = new Font("Arial", Font.PLAIN, 24);
        Color textColor = new Color(36, 47, 65);
        Color buttonColor = new Color(255, 179, 71);
        Color exitButtonColor = new Color(255, 99, 71);

        frame.setLayout(new GridBagLayout());
        GridBagConstraints constraints = new GridBagConstraints();
        constraints.gridx = 0;
        constraints.gridy = 0;
        constraints.anchor = GridBagConstraints.WEST;
        constraints.insets = new Insets(5, 5, 5, 5);
        constraints.weightx = 1.0;

        inputFileLabel.setFont(font);
        inputFileLabel.setForeground(textColor);
        inputFileField.setFont(font);
        inputFileField.setEditable(false);

        browseButton.setFont(font);
        browseButton.setBackground(buttonColor);
        browseButton.setForeground(Color.WHITE);

        outputFileLabel.setFont(font);
        outputFileLabel.setForeground(textColor);
        outputFileField.setFont(font);
        outputFileField.setEditable(false);

        outputBrowseButton.setFont(font);
        outputBrowseButton.setBackground(buttonColor);
        outputBrowseButton.setForeground(Color.WHITE);

        exitButton.setFont(font);
        exitButton.setBackground(exitButtonColor);
        exitButton.setForeground(Color.WHITE);

        convertButton.setFont(font);
        convertButton.setBackground(buttonColor);
        convertButton.setForeground(Color.WHITE);

        frame.add(inputFileLabel, constraints);
        constraints.gridy++;
        constraints.fill = GridBagConstraints.HORIZONTAL;
        frame.add(inputFileField, constraints);
        constraints.gridy++;
        frame.add(browseButton, constraints);

        constraints.gridx = 1;
        constraints.gridy = 0;
        constraints.fill = GridBagConstraints.NONE;
        frame.add(outputFileLabel, constraints);
        constraints.gridy++;
        constraints.fill = GridBagConstraints.HORIZONTAL;
        frame.add(outputFileField, constraints);
        constraints.gridy++;
        frame.add(outputBrowseButton, constraints);

        constraints.gridx = 0;
        constraints.gridy++;
        constraints.gridwidth = 2;
        constraints.fill = GridBagConstraints.HORIZONTAL;
        frame.add(convertButton, constraints);

        constraints.gridy++;
        constraints.gridwidth = 2;
        frame.add(exitButton, constraints);

        frame.pack();
        frame.setVisible(true);
        frame.setLocationRelativeTo(null);

    }

    private void convertPPTXToText(String[] inputFiles, String outputDirectory, FileNameExtensionFilter chosenFilter) {
        try {
            for (String inputFile : inputFiles) {
                FileInputStream fis = new FileInputStream(inputFile);
                XMLSlideShow pptx = new XMLSlideShow(fis);

                File outputFolder = new File(outputDirectory);
                if (!outputFolder.exists()) {
                    outputFolder.mkdirs();
                }

                String outputFileName = outputDirectory + File.separator + new File(inputFile).getName();
                String chosenExtension = chosenFilter.getExtensions()[0];
                if (!outputFileName.toLowerCase().endsWith("." + chosenExtension)) {
                    outputFileName += "." + chosenExtension;
                }

                FileOutputStream fos = new FileOutputStream(outputFileName);
                PrintWriter pw = new PrintWriter(fos);

                for (XSLFSlide slide : pptx.getSlides()) {
                    for (XSLFShape shape : slide.getShapes()) {
                        if (shape instanceof XSLFTextShape textShape) {
                            for (XSLFTextParagraph paragraph : textShape.getTextParagraphs()) {
                                String paragraphText = paragraph.getText();
                                pw.println(paragraphText);
                            }
                        }
                    }
                }

                pw.close();
                fos.close();
                fis.close();
            }

            System.out.println("Файлы были успешно сконвертированы");
            JOptionPane.showMessageDialog(this.frame, "Файлы были успешно сконвертированы", "Статус конвертации", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(this.frame, "Произошла ошибка при конвертации файлов", "Ошибка конвертации", JOptionPane.ERROR_MESSAGE);
        }
    }

    public static void main(String[] args) {
        SwingUtilities.invokeLater(new Runnable() {
            public void run() {
                new PPTXToTextConverter();
            }
        });
    }
}
