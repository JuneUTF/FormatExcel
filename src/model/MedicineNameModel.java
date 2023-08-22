package model;

import lombok.Data;

/**EXCELファイルの処方薬名に挿入する**/
@Data
public class MedicineNameModel {
	/**処方薬名*
	 *A9-A38/8.0-37.0 */
	private String medicineName;
	/**使用法*
	 *J9-J38/8.9-37.9 */
	private String usage;
	/**個数*
	 *N9-N38/8.13-37.13 */
	private int quantity;
	
}
