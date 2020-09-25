package com.excelcracker;

import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.IOException;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ExcelCracker {

	public static void doExcelCracker(String fileName, String outFile, String passList, Boolean pass_is_file)
			throws IOException {
		// SS Workbook object
		Workbook workbook = null;
		Boolean Cracked = false;
		// Handles both XSSF and HSSF automatically
		if (pass_is_file) {
			BufferedReader br = new BufferedReader(new FileReader(passList));
			String line;
			while ((line = br.readLine()) != null && Cracked == false) {
				try {
					// System.out.println(line);
					workbook = WorkbookFactory.create(new FileInputStream(fileName), line);
					Cracked = true;
					System.out.println("password found:" + line);
				} catch (Exception e) {
					// System.out.println(e);
				}
			}
			br.close();
		} else {
			try {
				// System.out.println(line);
				workbook = WorkbookFactory.create(new FileInputStream(fileName), passList);
				Cracked = true;
				System.out.println("password found:" + passList);
			} catch (Exception e) {
				// System.out.println(e);
			}
		}
		if (workbook == null) {
			System.out.println("Unable to crack file");
			System.exit(1);
		}
		try {
			FileOutputStream out = new FileOutputStream(outFile);
			workbook.write(out);
			out.close();
			System.out.println("Decrypted file written to: " + outFile);
		} catch (Exception ex) {
			System.out.println("Unable to save decyrpted file" + ex);
			System.exit(1);
		}
	}

	public static void main(String args[]) throws IOException {
		String inFilePath = null;
		String outFilePath = null;
		String passList = null;
		for (int i = 0; i < args.length; i++) {
			if (args[i].equals("-i")) {
				i++;
				inFilePath = args[i];
			} else if (args[i].equals("-o")) {
				i++;
				outFilePath = args[i];
			} else if (args[i].equals("-p")) {
				i++;
				passList = args[i];
			} else {
			}
		}
		File f = new File(inFilePath);
		Boolean pass_is_file = false;
		if (passList != null) {
			File p = new File(passList);
			if (p.exists() && p.isFile() && p.canRead()) {
				pass_is_file = true;
			}
		}
		if (f.exists() && f.isFile() && f.canRead()) {
			if (outFilePath == null) {
				outFilePath = inFilePath + ".decrypted";
			}
			doExcelCracker(inFilePath, outFilePath, passList, pass_is_file);
		} else {
			System.out.println(
					"I need to know what you want cracked -i <input file> -o <output file> -p <password or passlist>");
			System.exit(1);
		}
	}
}
