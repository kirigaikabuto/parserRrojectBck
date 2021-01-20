import org.apache.commons.io.IOUtils;

import java.io.*;
import java.net.URL;
import java.util.ArrayList;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Main {

	public static void main(String[] args) throws IOException {
		//SanctionsDownload();
		ArrayList<String> files = ReadDataFromExcelFiles("./sanctions");
		for (String s : files) {
			System.out.println(s);
		}
		MergeExcelFiles(files);
	}

	public static void MergeExcelFiles(ArrayList<String> files) throws IOException {
		String fileName = files.get(0);

		FileInputStream file = new FileInputStream("./sanctions/1020.xls");

		//Create Workbook instance holding reference to .xlsx file
		XSSFWorkbook workbook = new XSSFWorkbook(file);

		//Get first/desired sheet from the workbook
		XSSFSheet sheet = workbook.getSheetAt(0);

		//Iterate through each rows one by one
		Iterator<Row> rowIterator = sheet.iterator();
	}

	public static void SanctionsDownload() throws IOException {
		int start = 1;
		int end = 16595;
		int step = 100;
		for (int i = start; i < end; i += step) {
			System.out.println(i);
			String mainDirectory = "./sanctions";
			String url = "https://www.finreg.kz/index.cfm?docid=3227&switch=russian&organization=&organizationid=&organizationtypeid=&typeid=&kindid=&startdate=&enddate=&npatypeid=&sourceid=&start=" + String.valueOf(i) + "&excel";
			String fileName = String.valueOf(i) + ".xls";
			String directory = mainDirectory + "/" + fileName;
			int k = DownloadFileFromUrl(url, directory);
		}
	}

	public static int DownloadFileFromUrl(String url, String pathToSaveFile) throws IOException {
		InputStream inputStream = new URL(url).openStream();
		FileOutputStream fileOs = new FileOutputStream(pathToSaveFile);
		int k = IOUtils.copy(inputStream, fileOs);
		return k;
	}

	public static ArrayList<String> ReadDataFromExcelFiles(String pathToFolder) {
		String[] paths;
		ArrayList<String> absolutePath = new ArrayList<>();
		File newFiles = new File(pathToFolder);
		paths = newFiles.list();
		System.out.println(paths.length);
		for (String pathname : paths) {
			absolutePath.add(pathToFolder + "/" + pathname);
		}
		return absolutePath;
	}
}
