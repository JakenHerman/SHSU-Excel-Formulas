public class retestChecker {

	public static void main(String[] args){

	HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(retest1232014.xlsx);
	HSSFSheet sheet = workbook.getSheetAt(0);
	FormulaEvaluator evaluator = sheet.getCreationHelper().createFormulaEvaluator();	

	HSSFRow row;

	for (i = 0; i <= 1000; i++){
		row = sheet.getRow(i);

		
	}	

}

}
