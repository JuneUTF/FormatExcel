package controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.EncryptedDocumentException;

import model.MedicineFomatModel;
import model.MedicineNameModel;
import service.ExcelExportService;

public class Main {

	public static void main(String[] args) throws EncryptedDocumentException, IOException {
		// TODO 自動生成されたメソッド・スタブ
//		tạo model ảo
//		MedicineFomatModel
		MedicineFomatModel fomatModel = new MedicineFomatModel();
		fomatModel.setToDay("2023年8月19日 - 19:00");
		fomatModel.setMedicineCode(25061996);
		fomatModel.setGuestName("フィン");
		fomatModel.setManager("June");
		fomatModel.setBranchName("朝日浅草支店");
		fomatModel.setMedicineTime("2023年8月19日");
//		MedicineNameModel
		List<MedicineNameModel> medicineNameModel = new ArrayList<>();

		MedicineNameModel medicine1 = new MedicineNameModel();
		medicine1.setMedicineName("薬01");
		medicine1.setUsage("2玉で朝・昼・晩食事号後");
		medicine1.setQuantity(12);
		medicineNameModel.add(medicine1);

		MedicineNameModel medicine2 = new MedicineNameModel();
		medicine2.setMedicineName("薬02");
		medicine2.setUsage("2玉で朝日食事号後");
		medicine2.setQuantity(15);
		medicineNameModel.add(medicine2);

		MedicineNameModel medicine3 = new MedicineNameModel();
		medicine3.setMedicineName("薬03");
		medicine3.setUsage("2玉で朝・晩食事号後");
		medicine3.setQuantity(18);
		medicineNameModel.add(medicine3);

		System.out.println(medicineNameModel);

		ExcelExportService.setExcel(fomatModel,medicineNameModel);
	}

}
