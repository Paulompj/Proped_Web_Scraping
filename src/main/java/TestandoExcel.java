import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;

public class TestandoExcel {
    public static void main(String[] args) {
        Workbook workbook = null;
        FileInputStream fis = null;

        try {
            // Tentar abrir o arquivo Excel existente (caso o arquivo não exista, ele cria um novo)
            File file = new File("Projetos.xlsx");
            if (file.exists()) {
                // Se o arquivo existe, carrega o arquivo Excel existente
                fis = new FileInputStream(file);
                workbook = WorkbookFactory.create(fis);
            } else {
                // Se o arquivo não existe, cria um novo workbook
                workbook = new XSSFWorkbook();
            }

            // Se a planilha não existir, cria uma nova planilha chamada "Projetos"
            Sheet sheet = workbook.getSheet("Projetos");
            if (sheet == null) {
                sheet = workbook.createSheet("Projetos");
                // Criar a primeira linha de cabeçalho (títulos das colunas)
                Row headerRow = sheet.createRow(0);
                headerRow.createCell(0).setCellValue("Código");
                headerRow.createCell(1).setCellValue("Título");
                headerRow.createCell(2).setCellValue("Coordenador");
                headerRow.createCell(3).setCellValue("Tipo de Projeto");
                headerRow.createCell(4).setCellValue("Status");
            }

            // Obter o número da última linha (caso já tenha dados, começaremos a adicionar a partir da próxima linha)
            int lastRowNum = sheet.getPhysicalNumberOfRows();

            ArrayList<Projeto> projetos = Conversor.stringExcel();
            for (int i = 0; i < projetos.size(); i++) {
                // Criar novas linhas com os dados do Projeto
                Row row = sheet.createRow(lastRowNum + i);
                row.createCell(0).setCellValue(projetos.get(i).getCodigo());
                row.createCell(1).setCellValue(projetos.get(i).getTitulo());
                row.createCell(2).setCellValue(projetos.get(i).getCoordenador());
                row.createCell(3).setCellValue(projetos.get(i).getTipo());
                row.createCell(4).setCellValue(projetos.get(i).getStatus());
            }

            // Ajustar automaticamente o tamanho das colunas para o conteúdo
            sheet.autoSizeColumn(0);
            sheet.autoSizeColumn(1);
            sheet.autoSizeColumn(2);
            sheet.autoSizeColumn(3);
            sheet.autoSizeColumn(4);

            // Escrever o arquivo Excel no disco
            try (FileOutputStream fos = new FileOutputStream(file)) {
                workbook.write(fos);
            } catch (IOException e) {
                e.printStackTrace();
            }

            System.out.println("Arquivo Excel atualizado com sucesso!");

        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            // Fechar os recursos
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
}
