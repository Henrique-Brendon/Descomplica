package com.henrique;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
       try {
            // Caminho para o arquivo da planilha
            String caminhoArquivo = "caminho/para/sua/planilha.xlsx";

            // Carrega o arquivo da planilha
            FileInputStream arquivo = new FileInputStream(new File(caminhoArquivo));

            // Cria o workbook (representa toda a planilha)
            XSSFWorkbook workbook = new XSSFWorkbook(arquivo);

            // Obtém a primeira planilha (sheet) no workbook
            XSSFSheet sheet = workbook.getSheetAt(0);

            // Itera sobre as linhas da planilha
            for (Row row : sheet) {
                // Itera sobre as células de cada linha
                for (Cell cell : row) {
                    // Obtém o valor da célula
                    System.out.print(cell.toString() + "\t");
                }
                System.out.println(); // Pula para a próxima linha
            }

            // Fecha o arquivo após a leitura
            arquivo.close();
            workbook.close();
        } catch (IOException e) {
           System.out.println("Planilha não encontrada!" + e.getMessage());
        }
    }
}
