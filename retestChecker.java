public class retestChecker {

	public static void main(String[] args){
	
	FileInputStream fis = new FileInputStream(retest1232014.xlsx);
	HSSFWorkbook workbook = new HSSFWorkbook(fis);
	HSSFSheet sheet = workbook.getSheetAt(0);
	FormulaEvaluator evaluator = sheet.getCreationHelper().createFormulaEvaluator();
	
	HSSFRow rankRow;
	HSSFCell rankCell;

	for (int i = 0; i <= 1000; i++){
		CellReference rankReference = new CellReference("F%d", i); //checks rank
		rankRow = sheet.getRow(rankReference.getRow());
		rankCell = row.getCell(rankReference.getCol());
		CellReference AoD = new CellReference("I%d", i);
		
		CellReference scoreReference = new CellReference("H%d", i);
		CellReference GPAReference = new CellReference("G%d", i);
		
		if (rankCell!=null) {
			if (rankCell >= 75){
				AoD.setCellValue("Accept");//Accept or Deny Column
			}
			else if ((rankCell < 75) && (rankCell >= 50){
				if (scoreReference.length() ==2){
					if (scoreReference >= 20){
						AoD.setCellValue("Accept");
					}
					else {
						AoD.setCellValue("Deny");
					}
				}
				else if (scoreReference.length() == 4){
					if (scoreReference >= 0960){
						AoD.setCellValue("Accept");
					}
					else {
						AoD.setCellValue("Deny");
					}
				}
				else{
					AoD.setCellValue("ERROR");
				}
			}
			else if ((rankCell < 50) && (rankCell >= 25){
				if (scoreReference.length() == 2){
					if (scoreReference >= 23){
						AoD.setCellValue("Accept");
					}
					else {
						AoD.setCellValue("Deny");
					}
				}
				else if (scoreReference.length() == 4){
					if (scoreReference >= 1060){
						AoD.setCellValue("Accept");
					}
					else{
						AoD.setCellValue("Deny");
					}
				}
				else {
					AoD.setCellValue("ERROR");
				}
			}
			else {
				AoD.setCellValue("Deny");
			}
    		}
    		else if (GPAReference != null){
    			if (GPAReference >= 3.5){
    				if (scoreReference.length() == 4){
    					if(scoreReference >= 0851){
    						AoD.setCellValue("Accept");
    					}
    					else{
    						AoD.setCellValue("Deny");
    					}
    				}
    				else if (scoreReference.length() == 2){
    					if(scoreReference >= 18){
    						AoD.setCellValue("Accept");
    					}
    					else{
    						AoD.setCellValue("Deny");
    					}
    				}
    				else{
    					AoD.setCellValue("ERROR");
    				}
    			}
    			else if ((GPAReference >= 3.0) && (GPAReference <= 3.49)){
    				if (scoreReference.length() == 2) {
    					if (scoreReference >= 20){
    						AoD.setCellValue("Accept");
    					}
    					else{
    						AoD.setCellValue("Deny");
    					}
    				}
    				else if (scoreReference.length() == 4){
    					if (scoreReference >= 0931){
    						AoD.setCellValue("Accept");
    					}
    					else{
    						AoD.setCellValue("Deny");
    					}
    				}
    				else{
    					AoD.setCellValue("ERROR");
    				}
    			}
    			else if ((GPAReference >= 2.5) && (GPAReference >= 2.99)){
    				if (scoreReference.length() == 2) {
    					if (scoreReference >= 23){
    						AoD.setCellValue("Accept");
    					}
    					else{
    						AoD.setCellValue("Deny");
    					}
    				}
    				else if (scoreReference.length() == 4){
    					if (scoreReference >= 1031){
    						AoD.setCellValue("Accept");
    					}
    					else{
    						AoD.setCellValue("Deny");
    					}
    				}
    				else{
    					AoD.setCellValue("ERROR");
    				}
    			}
    			else if ((GPAReference >= 2.0) && (GPAReference <= 2.49){
    				if (scoreReference.length() == 2) {
    					if (scoreReference >= 26){
    						AoD.setCellValue("Accept");
    					}
    					else{
    						AoD.setCellValue("Deny");
    					}
    				}
    				else if (scoreReference.length() == 4){
    					if (scoreReference >= 1141){
    						AoD.setCellValue("Accept");
    					}
    					else{
    						AoD.setCellValue("Deny");
    					}
    				}
    				else{
    					AoD.setCellValue("ERROR");
    				}
    			}
    			else if ((GPAReference >= 0) && (GPAReference <= 1.99)){
    				if (scoreReference.length() == 2) {
    					if (scoreReference >= 28){
    						AoD.setCellValue("Accept");
    					}
    					else{
    						AoD.setCellValue("Deny");
    					}
    				}
    				else if (scoreReference.length() == 4){
    					if (scoreReference >= 1201){
    						AoD.setCellValue("Accept");
    					}
    					else{
    						AoD.setCellValue("Deny");
    					}
    				}
    				else{
    					AoD.setCellValue("ERROR");
    				}
    			}
    		}
		
	}	

}

}
