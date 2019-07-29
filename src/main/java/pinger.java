import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.net.InetAddress;
import java.net.UnknownHostException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class pinger {
	static int timeOuts = 0;
	static int colToCheck = 1;

	public static void setTimeOut(int timeOut) {
		timeOuts = timeOut;
	}

	public static void main(String[] args) {
		int column = 0;
		int destinationColumn = 0;
		Scanner sc = new Scanner(System.in);
		System.out.println("Provide TimeOut in ms");
		setTimeOut(sc.nextInt());
		File pinger = new File("test.txt");
		String root = pinger.getAbsolutePath().split("test.txt")[0];
		// System.out.println(root);
		File rootFolder = new File(root);
		File[] files = rootFolder.listFiles();
		System.out.println("Which file should be checked? (type number)");
		List<File> xlsFiles = new ArrayList<File>();

		for (int i = 0; i < files.length; i++) {
			if (files[i].getName().contains("xlsx")) {
				xlsFiles.add(files[i]);
				System.out.println((xlsFiles.size() - 1) + " ---  " + xlsFiles.get(xlsFiles.size() - 1).getName());
			}
		}
		boolean correctNumber = false;
		int fileIndex = 0;
		while (!correctNumber) {
			fileIndex = sc.nextInt();
			if (fileIndex < 0 || fileIndex > xlsFiles.size() - 1) {
				System.out.println("Wrong Number type number from range: 0-" + (xlsFiles.size() - 1));

			} else {
				correctNumber = true;
			}
		}

		System.out.println("Column to check  (A,B...):");
		column = CellReference.convertColStringToIndex(sc.next());
		System.out.println("Where put result (A,B...)?: ");
		destinationColumn = CellReference.convertColStringToIndex(sc.next());
		FileInputStream file = null;

		try {

			System.out.println("Scaning..  " + xlsFiles.get(fileIndex));
			file = new FileInputStream(xlsFiles.get(fileIndex));

		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
			sc.nextLine();
		}

		XSSFWorkbook workbook = null;
		try {
			workbook = new XSSFWorkbook(file);
			System.out.println("Worbook loaded");
		} catch (IOException e) {
			System.out.println("Could not load workbook");
			// TODO Auto-generated catch block
			e.printStackTrace();
			sc.nextLine();
		}

		XSSFSheet sheet = workbook.getSheetAt(0);

		for (int i = 0; i < sheet.getLastRowNum() + 1; i++) {

			XSSFRow row = sheet.getRow(i);
			if (row != null && row.getCell(column) != null
					&& row.getCell(column).getCellType() == CellType.STRING.getCode()
					&& row.getCell(column).getRichStringCellValue() != null) {
				String msg = row.getCell(column).getRichStringCellValue().getString();
			//	if (row.getCell(column).getStringCellValue().equals("Adresacja publiczna DSL")
				//		|| row.getCell(column).getStringCellValue().equals("Czekamy na ³¹cze")) {

			//	} else {
					try {
						if (sendPingRequest(row.getCell(column).getStringCellValue().trim())) {
							msg = msg + " OK";
							if (row.getCell(destinationColumn) != null) {
								row.getCell(destinationColumn).setCellValue("OK");
							} else {
								row.createCell(destinationColumn).setCellValue("OK");
							}

						} else {
							msg = msg + " NOT OK";
							
							if (row.getCell(destinationColumn) != null) {
								row.getCell(destinationColumn).setCellValue("Err");
							} else {
								row.createCell(destinationColumn).setCellValue("Err");
							}

						}
						System.out.println(msg);
					} catch (UnknownHostException e1) {
						if (row.getCell(destinationColumn) != null) {
							row.getCell(destinationColumn).setCellValue("Wrong Ip");
						} else {
							row.createCell(destinationColumn).setCellValue("Wrong Ip");
						}
						
					}

			//	}
			} else {
				System.out.println("Empty or wrong cell format ");
			}

		}
		FileOutputStream out = null;
		try {
			
			out = new FileOutputStream(
					new File(root + "//" + xlsFiles.get(fileIndex).getName().split(".xlsx")[0] + "_DONE.xlsx"));
			System.out.println("Checking finished. Output file: " + root
					+ xlsFiles.get(fileIndex).getName().split(".xlsx")[0] + "_DONE.xlsx");
		} catch (FileNotFoundException e1) {
			
			e1.printStackTrace();
			
		}
		
		try {
			workbook.write(out);
		} catch (IOException e) {
			//
			e.printStackTrace();

		}
		try {
			out.close();
		} catch (IOException e) {

			e.printStackTrace();

		}

	}

	public static boolean sendPingRequest(String ipAddress) throws UnknownHostException {
		Scanner sc = new Scanner(System.in);
		// sc.nextLine();
		InetAddress adress = null;

		adress = InetAddress.getByName(ipAddress);

		boolean pingable = false;
		try {

			if (adress.isReachable(timeOuts)) {

				pingable = true;

			} else {
				pingable = false;

			}
		} catch (IOException e) {

		}

		return pingable;
	}
}
