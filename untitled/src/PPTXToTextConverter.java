import org.apache.poi.xslf.usermodel.*;
import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.*;

import static java.awt.Color.*;
import static javax.swing.JLabel.*;

public class PPTXToTextConverter {
    private JFrame frame;
    private JLabel inputFileLabel;
    private JTextField inputFileField;
    private JButton browseButton;
    private JLabel outputFileLabel;
    private JTextField outputFileField;
    private JButton outputBrowseButton;
    private JButton convertButton;
    private JButton exitButton; // Добавленная кнопка "Выйти"

    public PPTXToTextConverter() {
        frame = new JFrame("PPTX to Text Converter");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setLayout(new GridLayout(6, 1));

        inputFileLabel = new JLabel("Выберите PPTX файл:");
        inputFileField = new JTextField();
        browseButton = new JButton("Обзор");
        // ActionListener для кнопки "Обзор" (аналогично добавьте для outputBrowseButton)
        browseButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                int result = fileChooser.showOpenDialog(frame);
                if (result == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    inputFileField.setText(selectedFile.getAbsolutePath());
                }
            }
        });

        outputFileLabel = new JLabel("Выберите куда сохранить сконвертированный файл:");
        outputFileField = new JTextField();
        outputBrowseButton = new JButton("Сохранить");
        outputBrowseButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                JFileChooser fileChooser = new JFileChooser();
                int result = fileChooser.showSaveDialog(frame);
                if (result == JFileChooser.APPROVE_OPTION) {
                    File selectedFile = fileChooser.getSelectedFile();
                    outputFileField.setText(selectedFile.getAbsolutePath());
                }
            }
        });
        // ActionListener для кнопки "Путь"

        // ActionListener для кнопки "Конвертировать"
        convertButton = new JButton("Конвертировать");
        convertButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                String inputFile = inputFileField.getText();
                String outputFile = outputFileField.getText();
                if (!inputFile.toLowerCase().endsWith(".pptx")) {
                    JOptionPane.showMessageDialog(frame, "Вы выбрали неверный формат, выберите файл в формате pptx", "Invalid Format", JOptionPane.ERROR_MESSAGE);
                    return;
                }
                convertPPTXToText(inputFile, outputFile);
            }
        });

        // ActionListener для кнопки "Выйти"
        exitButton = new JButton("Выйти");
        exitButton.addActionListener(new ActionListener() {
            public void actionPerformed(ActionEvent e) {
                System.exit(0);
            }
        });

        Font font = new Font("Arial", Font.PLAIN, 24);
        Color textColor = new Color(36, 47, 65);
        Color buttonColor = new Color(97, 212, 195);

        inputFileLabel.setFont(font);
        inputFileLabel.setForeground(textColor);
        inputFileField.setFont(font);

        browseButton.setFont(font);
        browseButton.setBackground(buttonColor);
        browseButton.setForeground(Color.WHITE);

        outputBrowseButton.setFont(font);
        outputBrowseButton.setBackground(buttonColor);
        outputBrowseButton.setForeground(Color.WHITE);

        outputFileLabel.setFont(font);
        outputFileLabel.setForeground(textColor);
        outputFileField.setFont(font);

        // Аналогично установите стили для outputFileLabel, outputFileField, outputBrowseButton

        convertButton.setFont(font);
        convertButton.setBackground(buttonColor);
        convertButton.setForeground(Color.WHITE);

        exitButton.setFont(font);
        exitButton.setBackground(buttonColor);
        exitButton.setForeground(Color.WHITE);

        frame.add(inputFileLabel);
        frame.add(inputFileField);
        frame.add(browseButton);
        frame.add(outputFileLabel);
        frame.add(outputFileField);
        frame.add(outputBrowseButton);
        frame.add(convertButton);
        frame.add(exitButton);

        frame.pack();
        frame.setVisible(true);
        frame.setLocationRelativeTo(null);
    }

    private void convertPPTXToText(String inputFile, String outputFile) {
        try {
            FileInputStream fis = new FileInputStream(inputFile);
            XMLSlideShow pptx = new XMLSlideShow(fis);

            File file = new File(outputFile);
            if (!file.getParentFile().exists()) {
                file.getParentFile().mkdirs();
            }

            FileOutputStream fos = new FileOutputStream(file);
            PrintWriter pw = new PrintWriter(fos);
            for (XSLFSlide slide : pptx.getSlides()) {
                for (XSLFShape shape : slide.getShapes()) {
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape textShape = (XSLFTextShape) shape;
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

            System.out.println("Conversion completed!");
            JOptionPane.showMessageDialog(frame, "Файл был успешно сконвертирован", "Статус конвертации", JOptionPane.INFORMATION_MESSAGE);
        } catch (IOException e) {
            e.printStackTrace();
            JOptionPane.showMessageDialog(frame, "Не удается найти указанный файл", "Ошибка конвертации", JOptionPane.ERROR_MESSAGE);
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