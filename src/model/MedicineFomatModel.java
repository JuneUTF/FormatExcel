package model;

import lombok.Data;

/**
 * EXCELファイルの情報に挿入するためのモデルクラスです。
 */
@Data
public class MedicineFomatModel {
	/**現在時間（****年**月**日ー時：分）
	 * J3/2.9**/
	private String toDay;
	/**調剤コード
	 * B5/4.1**/
	private int medicineCode;
	/**お客様の名前
	 * G5/4.6**/
	private String guestName;
	/**担当者
	 * M5/4.12**/
	private String manager;
	/**支店名
	 * B6/5.1**/
	private String branchName;
	/**調剤日
	 * J6/5.9**/
	private String medicineTime;
	
}
