
public class POIP_Demo3 {

	public static void main(String[] args) {
		
    MyXLSReader reader = new MyXLSReader ("src//ExcelXLSX.xlsx");
    
    int count = reader.getRowCount("EmployeeData");
    
    System.out.println(count);
    
    String str= reader.getCellData("EmployeeData", "Occupation", 4);
    
    System.out.println(str);
    
    String str1= reader.getCellData("EmployeeData", 3, 4);
    
    System.out.println(str1);
    
   boolean b = reader.isSheetExist("EmployeeData");
   
   System.out.println(b);
   
   if(reader.isSheetExist("EmployeeData") )
		   {
	  
   int k = reader.getColumnCount("EmployeeData");
   
   System.out.println(k);
   
  int m = reader.getCellRowNum("EmployeeData", "ID", "002");
  
  System.out.println(m);
	   }
    
	}
}
