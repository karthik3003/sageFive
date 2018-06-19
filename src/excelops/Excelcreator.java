package excelops;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashSet;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelcreator {

	public static void main(String args[]) throws IOException, illegalSheetIndexException {

		String x = "bonmots@outlook.com,seano@optonline.net,aardo@optonline.net,kassiesa@aol.com,mmccool@gmail.com,mleary@att.net,jesse@comcast.net,benits@optonline.net,jbearp@optonline.net,slanglois@live.com,gomor@me.com,pgolle@me.com,jbarta@sbcglobal.net,solomon@att.net,lcheng@aol.com,panolex@gmail.com,chrisj@me.com,wenzlaff@mac.com,lukka@icloud.com,djpig@mac.com,maradine@att.net,keutzer@me.com,oechslin@me.com,koudas@live.com,dwsauder@me.com,uraeus@aol.com,nasor@sbcglobal.net,dwheeler@yahoo.ca,ccohen@msn.com,violinhi@gmail.com,godeke@yahoo.ca,jbearp@outlook.com,ylchang@yahoo.com,cgarcia@yahoo.ca,marcs@mac.com,brickbat@mac.com,nelson@optonline.net,frikazoyd@yahoo.com,fatelk@live.com,themer@aol.com,zeller@yahoo.com,yfreund@comcast.net,vmalik@hotmail.com,frode@att.net,emcleod@msn.com,arebenti@hotmail.com,techie@comcast.net,biglou@comcast.net,ylchang@optonline.net,martyloo@icloud.com,matthijs@mac.com,dkrishna@verizon.net,kenja@verizon.net,monopole@me.com,blixem@verizon.net,vlefevre@mac.com,fbriere@optonline.net,eminence@optonline.net,madanm@msn.com,gospodin@yahoo.com,lushe@me.com,dialworld@aol.com,glenz@verizon.net,rsteiner@me.com,nweaver@outlook.com,amaranth@yahoo.ca,loscar@yahoo.ca,staikos@msn.com,uraeus@yahoo.ca,eidac@verizon.net,mchugh@comcast.net,dkeeler@hotmail.com,andrei@live.com,tkrotchko@msn.com,thaljef@aol.com,bwcarty@sbcglobal.net,hstiles@sbcglobal.net,teverett@att.net,webteam@msn.com,lydia@yahoo.ca,thomasj@sbcglobal.net,parrt@msn.com,rsmartin@yahoo.com,dbrobins@mac.com,ahmad@hotmail.com,epeeist@gmail.com,wildfire@optonline.net,smpeters@live.com,maradine@yahoo.ca,johnbob@aol.com,nweaver@yahoo.ca,leviathan@att.net,stern@outlook.com,richard@comcast.net,nimaclea@comcast.net,louise@mac.com,pemungkah@mac.com,andrei@att.net,yenya@sbcglobal.net,treit@yahoo.ca";
		String[] email = x.split(",");
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet("mysheet");

		for (int i = 0; i < 100; i++) {
			XSSFRow row = sheet.createRow(i);
			XSSFCell cell = row.createCell(0);
			cell.setCellValue(email[i]);
		}
		sheet.autoSizeColumn(0);
		wb.write(new FileOutputStream("exl.xlsx"));
		wb.close();

		List l = new LinkedList();
		Set s = new HashSet();
		XSSFWorkbook wb1 = new XSSFWorkbook(new FileInputStream("exl.xlsx"));
		XSSFSheet sheet1 = wb1.getSheetAt(0);
		for (int i = 0; i < 100; i++) {
			XSSFRow row1 = sheet1.getRow(i);
			l.add(row1.getCell(0).getStringCellValue());
		}
		// System.out.println(l.get(4));
		s.addAll(l);
		// System.out.println(s.size());
		
		Iterator iter = l.iterator();
		while(iter.hasNext()) {
			System.out.println(iter.next());
		}
		System.out.println("**********************************************");
		
		for(Object o : s) {
			System.out.println(o);
		}
		

		// cell = row.createCell(1);
		// DataFormat format = wb.createDataFormat();
		// CellStyle dtStyle = wb.createCellStyle();
		// dtStyle.setDataFormat(format.getFormat("dd.mm.yyyy"));
		// cell.setCellStyle(dtStyle);
		// cell.setCellValue(new Date());
		//
		// row.createCell(2).setCellValue("karthik");
		//
		// sheet.autoSizeColumn(2);
		//
		// wb.write(new FileOutputStream("exl.xlsx"));
		// wb.close();

		// XSSFWorkbook wb1 = new XSSFWorkbook(new FileInputStream("exl.xlsx"));
		// XSSFSheet sheet1 = wb1.getSheetAt(0);
		// XSSFRow row1 = sheet1.getRow(0);
		// System.out.println(row1.getCell(0).getStringCellValue());
		// System.out.println(row1.getCell(1).getDateCellValue());
		// System.out.println(row1.getCell(2).getStringCellValue());
		//
		// }
	}
}
