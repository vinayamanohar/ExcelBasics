import java.io.IOException;
import java.util.ArrayList;

public class testSample {

	public static void main(String[] args) throws IOException {
		
		ExcelDemo ed = new ExcelDemo();
		ArrayList<String> aa = ed.getData("Purchase");
		System.out.println(aa.get(0));
		System.out.println(aa.get(1));
		System.out.println(aa.get(2));
		System.out.println(aa.get(3));
	}
	
	
}
