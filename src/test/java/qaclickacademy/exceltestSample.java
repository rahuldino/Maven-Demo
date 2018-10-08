// getting the data from excel file
package qaclickacademy;

import java.io.IOException;
import java.util.ArrayList;

public class exceltestSample {

	public static void main(String[] args) throws IOException {
		// TODO Auto-generated method stub
	    exceldriven d=new exceldriven();
	   ArrayList data= d.getData("Add Profile");
	   System.out.println(data.get(0));
	   System.out.println(data.get(1));
	   System.out.println(data.get(2));
	   System.out.println(data.get(3));

	}

}
