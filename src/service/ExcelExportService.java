package service;

import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import model.MedicineFomatModel;
import model.MedicineNameModel;

public class ExcelExportService {
    // Excel フォーマットと医薬品名を受け取り、Excel ファイルを設定
    public static void setExcel(MedicineFomatModel fomatModel, List<MedicineNameModel> medicineNameModel){
        // ファイルパスの指定
        String pathFile = "C:\\Users\\juneu\\Downloads\\帳票\\format.xlsx";

        try {
            // ファイルを読み込むためのパスを作成
            Path tempPath = Paths.get(pathFile);
            InputStream inSt = Files.newInputStream(tempPath);

            // Excel ワークブックを作成
            Workbook tempbook = new XSSFWorkbook(inSt);

            // シートを取得
            Sheet sheet = tempbook.getSheetAt(0);
//            EXCELファイルの情報に挿入する
            // 現在時間
            Row rowToDay = sheet.getRow(2);// 3 行目を取得
            Cell cellToDay = rowToDay.getCell(9);// J 列目のセルを取得
            cellToDay.setCellValue(fomatModel.getToDay());// セルの値を変更
            
         // 調剤コード
            Row rowMedicineCode = sheet.getRow(4);// 5 行目を取得
            Cell cellMedicineCode = rowMedicineCode.getCell(1);// B 列目のセルを取得
            cellMedicineCode.setCellValue(fomatModel.getMedicineCode());// セルの値を変更

            // お客様の名前
            Row rowGuestName = sheet.getRow(4);// 5 行目を取得
            Cell cellGuestName = rowGuestName.getCell(6);// G 列目のセルを取得
            cellGuestName.setCellValue(fomatModel.getGuestName());// セルの値を変更

            // 担当者
            Row rowManager = sheet.getRow(4);// 5 行目を取得
            Cell cellManager = rowManager.getCell(12);// M 列目のセルを取得
            cellManager.setCellValue(fomatModel.getManager());// セルの値を変更

            // 支店名
            Row rowBranchName = sheet.getRow(5);// 6 行目を取得
            Cell cellBranchName = rowBranchName.getCell(1);// B 列目のセルを取得
            cellBranchName.setCellValue(fomatModel.getBranchName());// セルの値を変更

            // 調剤日
            Row rowMedicineTime = sheet.getRow(5);// 6 行目を取得
            Cell cellMedicineTime = rowMedicineTime.getCell(9);// J 列目のセルを取得
            cellMedicineTime.setCellValue(fomatModel.getMedicineTime());// セルの値を変更

            
//            処方薬名を挿入する
            int itemRow = 8;
            for (MedicineNameModel item : medicineNameModel) {

            	// 処方薬名
            	Row rowMedicineName = sheet.getRow(itemRow);// itemRow 行目を取得
            	Cell cellMedicineName = rowMedicineName.getCell(0);// A 列目のセルを取得
            	cellMedicineName.setCellValue(item.getMedicineName());// セルの値を変更

            	// 使用法
            	Row rowUsage = sheet.getRow(itemRow);// itemRow 行目を取得
            	Cell cellUsage = rowUsage.getCell(9);// J 列目のセルを取得
            	cellUsage.setCellValue(item.getUsage());// セルの値を変更

            	// 個数
            	Row rowQuantity = sheet.getRow(itemRow);// itemRow 行目を取得
            	Cell cellQuantity = rowQuantity.getCell(13);// N 列目のセルを取得
            	cellQuantity.setCellValue(item.getQuantity());// セルの値を変更
            	itemRow++;
			}
            
            // ファイル出力のためのパスを作成
            Path outPath = Paths.get(
                    "C:\\Users\\juneu\\Downloads\\帳票\\" + fomatModel.getMedicineCode() + "-"
                            + fomatModel.getGuestName() + ".xlsx");

            // ファイル出力ストリームを作成
            OutputStream outSt = Files.newOutputStream(outPath);

            // Excel ファイルを出力
            tempbook.write(outSt);

            // 完了メッセージを表示
            System.out.print("入出力が完成！");
        } catch (Exception e) {
            // 入出力例外が発生した場合の処理
            System.out.print("入出力例外が発生！");
        }
    }
}
