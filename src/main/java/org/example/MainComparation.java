package org.example;

import java.io.FileInputStream;
import java.util.HashSet;

public class Main {
    public static void main(String[] args) {
        try {
            // Abrindo o arquivo Excel do Fusion
            FileInputStream fusionFile = new FileInputStream("caminho_para_tabelaFusion.xls");
            Workbook fusionWorkbook = new HSSFWorkbook(fusionFile);
            Sheet fusionSheet = fusionWorkbook.getSheetAt(0);

            // Abrindo o arquivo Excel do ATS
            FileInputStream atsFile = new FileInputStream("caminho_para_relatorioATS.xlsx");
            Workbook atsWorkbook = new XSSFWorkbook(atsFile);
            Sheet atsSheet = atsWorkbook.getSheetAt(0);

            // Lendo os IDs de venda do Fusion e armazenando em um HashSet
            HashSet<String> fusionSaleIds = new HashSet<>();
            for (Row row : fusionSheet) {
                Cell cell = row.getCell(row.getFirstCellNum());
                fusionSaleIds.add(cell.toString());
            }

            // Lendo os códigos do ATS e armazenando em um HashSet
            HashSet<String> atsCodes = new HashSet<>();
            for (Row row : atsSheet) {
                Cell cell = row.getCell(row.getFirstCellNum());
                atsCodes.add(cell.toString());
            }

            // Comparando os IDs e encontrando os que estão faltando no ATS
            fusionSaleIds.removeAll(atsCodes);

            // Exibindo os IDs que estão faltando no ATS
            System.out.println("IDs de venda do Fusion que estão faltando no ATS:");
            for (String saleId : fusionSaleIds) {
                System.out.println(saleId);
            }

            // Fechando os arquivos
            fusionWorkbook.close();
            atsWorkbook.close();
            fusionFile.close();
            atsFile.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}