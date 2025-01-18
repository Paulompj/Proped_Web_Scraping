import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;

public class AtualizarExcelComFinanciamento {
    public static void main(String[] args) {
        Workbook workbook = null;
        FileInputStream fis = null;

        try {
            // Caminho do arquivo Excel
            File file = new File("C:\\Users\\PC\\Desktop\\projetoProped\\Projetos7.xlsx");
            if (file.exists()) {
                fis = new FileInputStream(file);
                workbook = WorkbookFactory.create(fis);
            } else {
                workbook = new XSSFWorkbook();
            }

            Sheet sheet = workbook.getSheet("Projetos");
            if (sheet == null) {
                sheet = workbook.createSheet("Projetos");
            }

            // Criar cabeçalho para a nova coluna de financiamento
            Row headerRow = sheet.getRow(0);
            if (headerRow == null) {
                headerRow = sheet.createRow(0);
            }
            headerRow.createCell(14).setCellValue("Financiamento");  // Nova coluna para o financiamento

            int lastRowNum = sheet.getPhysicalNumberOfRows();

            // Iterar pelos projetos no Excel e buscar o financiamento no PDF
            for (int i = 1; i < lastRowNum; i++) {  // Começar da linha 1, pois a linha 0 é o cabeçalho
                Row row = sheet.getRow(i);
                String projectCode = row.getCell(0).getStringCellValue();  // Código do projeto
                String financing = extractFinancing("C:\\path\\to\\SeuPDF.pdf", projectCode);

                if (financing != null) {
                    row.createCell(14).setCellValue(financing);  // Adicionar o financiamento à nova coluna
                }
            }

            // Ajustar o tamanho das colunas
            for (int col = 0; col <= 14; col++) {
                sheet.autoSizeColumn(col);
            }

            // Salvar as mudanças no arquivo Excel
            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            }

            System.out.println("Arquivo Excel atualizado com sucesso!");

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            try {
                if (workbook != null) {
                    workbook.close();
                }
                if (fis != null) {
                    fis.close();
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    // Método para extrair financiamento do PDF baseado no código do projeto
    public static String extractFinancing(String pdfPath, String projectCode) {
        try {
            File file = new File(pdfPath);
            PDDocument document = PDDocument.load(file);
            PDFTextStripper stripper = new PDFTextStripper();
            String text = stripper.getText(document);
            document.close();

            String codeTag = "Projeto: " + projectCode;
            int projectIndex = text.indexOf(codeTag);
            if (projectIndex != -1) {
                int financingIndex = text.indexOf("Financiamento", projectIndex);
                if (financingIndex != -1) {
                    int start = financingIndex + "Financiamento".length();
                    int end = text.indexOf("\n", start);
                    return text.substring(start, end).trim();
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }
}
