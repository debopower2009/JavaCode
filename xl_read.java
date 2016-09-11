package pkg_1;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class xl_read {

	public static void main(String[] args) {
		FileInputStream fis;
		fis = null;
		try {
			fis = new FileInputStream("G:\\workspace\\EXperiment\\XL_sample.xlsx");
		} catch (FileNotFoundException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		XSSFWorkbook xwb;
		xwb = null;
		XSSFRow row;
		XSSFCell col;
		col = null;
		try {
			xwb = new XSSFWorkbook(fis);
		} catch (IOException e2) {
			// TODO Auto-generated catch block
			e2.printStackTrace();
		}
		XSSFSheet xsh = xwb.getSheetAt(0);
		Iterator row_itr = xsh.rowIterator();
		while (row_itr.hasNext()) {
			row = (XSSFRow) row_itr.next();
			Iterator col_itr = row.cellIterator();
			while (col_itr.hasNext()) {
				col = (XSSFCell) col_itr.next();

				if (col.getCellType() == XSSFCell.CELL_TYPE_STRING)

				{
					System.out.print(col.getStringCellValue()+"\t");
				} else if (col.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
					System.out.print(col.getNumericCellValue()+"\t");
				}
				
			}
			System.out.print("\n");
		}

		try {
			xwb.close();
		} catch (IOException e1) {
			// TODO Auto-generated catch block
			e1.printStackTrace();
		}
		try {
			fis.close();
		} catch (IOException e) {

			e.printStackTrace();
		}

	}

}
