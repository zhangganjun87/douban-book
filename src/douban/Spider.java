package douban;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.regex.*;

import org.jsoup.*;
import org.jsoup.nodes.*;
import org.jsoup.select.*;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Spider {
	private String url = "https://book.douban.com/tag/";
	private String bookType = "";
	private String page = "?start=";
	private int pageNumber = 0;
	String title[] = { "���", "����", "����", "��������", "����", "������", "��������", "�۸�" };
	Map<String, String> map = new HashMap<String, String>();
	static List<String> excelList = new ArrayList<String>();
	public static int rowId = 0;

	private String writer = null;
	private String publition = null;
	private String date = null;
	private String price = null;

	private HSSFWorkbook workbook = null;

	public Spider(String bookType) {
		this.bookType = bookType;
	}

	public void createExcel(String fileDir, String sheetName, String titleRow[]) throws Exception {
		// ����workbook
		workbook = new HSSFWorkbook();
		// ���worksheet
		HSSFSheet sheet1 = workbook.createSheet("����ͼ��");
		FileOutputStream out = null;
		try {
			HSSFRow row = workbook.getSheet("����ͼ��").createRow(0); // ������һ��
			for (int i = 0; i < titleRow.length; i++) {
				HSSFCell cell = row.createCell(i);
				cell.setCellValue(titleRow[i]);
			}
			out = new FileOutputStream(fileDir);
			workbook.write(out);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				out.close();
			} catch (IOException e1) {
				e1.printStackTrace();
			}
		}

	}

	public void writeToExcel(String fileDir, String sheetName, Map map) throws Exception {
		File file = new File(fileDir);
		try {
			workbook = new HSSFWorkbook(new FileInputStream(file));
		} catch (FileNotFoundException e3) {
			e3.printStackTrace();
		} catch (IOException e4) {
			e4.printStackTrace();
		}
		FileOutputStream out = null;
		HSSFSheet sheet = workbook.getSheet("����ͼ��");
		// ��ȡ��ͷ����
		int columnCount = sheet.getRow(0).getLastCellNum() + 1;
		try {
			HSSFRow titleRow = sheet.getRow(0);
			if (titleRow != null) {
				// �����µ���
				HSSFRow newRow = sheet.createRow(rowId);
				for (int columnIndex = 0; columnIndex < columnCount - 1; columnIndex++) {
					String mapKey = titleRow.getCell(columnIndex).toString().trim();
					HSSFCell cell = newRow.createCell(columnIndex);
					cell.setCellValue(map.get(mapKey) == null ? null : map.get(mapKey).toString());
				}
			}
			out = new FileOutputStream(fileDir);
			workbook.write(out);
			System.out.println("���ڼ�¼��" + rowId + "�������Ϣ...");
		} catch (Exception e5) {
			throw e5;
		} finally {
			try {
				out.close();
			} catch (IOException e6) {
				e6.printStackTrace();
			}
		}
	}

	public int getWebbookInfo() {
		try {
			// �ӵ�һҳ��ʼץȡ��ÿ20����Ϊһҳ
			for (pageNumber = 0; pageNumber < 20 * 100; pageNumber += 20) {
				Document doc = Jsoup.connect(url + bookType + page + pageNumber).data("query", "Java")
						.userAgent(
								"Mozilla/5.0 (Windows NT 10.0; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/55.0.2883.87 Safari/537.36 QIHU 360SE")
						.cookie("auth", "token").timeout(30000).get();
				Elements elements = doc.select("div.info");
				if (elements != null) {
					for (Element element : elements) {
						String title = element.getElementsByTag("h2").text();
						String rate = element.getElementsByClass("rating_nums").text();
						String peoplenum = element.getElementsByClass("pl").text();
						int len = element.getElementsByClass("pub").text().split("/").length;
						if (len > 2) {
							writer = element.getElementsByClass("pub").text().split("/")[0];
							publition = element.getElementsByClass("pub").text().split("/")[len - 3];
							date = element.getElementsByClass("pub").text().split("/")[len - 2];
							price = element.getElementsByClass("pub").text().split("/")[len - 1];
						} else {
							writer = element.getElementsByClass("pub").text().split("/")[0];
							publition = "δ֪";
							date = "δ֪";
							price = "δ֪";
							System.out.println("��" + rowId + "������Ϣ��ȫ");
						}
						rowId++;

						map.put("���", rowId + "");
						map.put("����", title);
						map.put("����", rate);
						map.put("��������", peoplenum);
						map.put("����", writer);
						map.put("������", publition);
						map.put("��������", date);
						map.put("�۸�", price);
						try {
							writeToExcel("D:/test.xls", "����ͼ��", map);
						} catch (Exception e8) {
							e8.printStackTrace();
						}

					}
				} else {
					break;
				}
			}
		} catch (IOException e2) {
			e2.printStackTrace();
		}
		return rowId;
	}

	public void getFirstHunderedBooks(String fileDir, String sheetName, List<String> excelList2) {
		// ֻʵ������excel���ҵ�����������1000���飬�����뵽list�����ն˴�ӡ��������δʵ�����ֵ�����
		File file = new File(fileDir);
		try {
			workbook = new HSSFWorkbook(new FileInputStream(file));
		} catch (FileNotFoundException e3) {
			e3.printStackTrace();
		} catch (IOException e4) {
			e4.printStackTrace();
		}
		HSSFSheet sheet = workbook.getSheet("����ͼ��");
		// ��ȡ����
		int rowCount = sheet.getLastRowNum();
		try {
			String regPattern = "\\((\\d+)������\\)";
			Pattern pattern = Pattern.compile(regPattern);
			for (int rowNum = 1; rowNum < rowCount - 1; rowNum++) {
				HSSFRow row = sheet.getRow(rowNum);
				HSSFCell cell = row.getCell(3);
				Matcher matcher = pattern.matcher(cell.getStringCellValue());
				boolean rs = matcher.find();
				if (rs == true) {
					String mat = matcher.group(1);
					int number = Integer.parseInt(mat);
					if (number > 1000) {
						String str = row.getCell(1).getStringCellValue() + "/" + row.getCell(2).getStringCellValue()
								+ "/" + row.getCell(3).getStringCellValue();
						excelList2.add(str);
					}
				}
			}
			for (String s : excelList2) {
				System.out.println(s);
			}
		} catch (Exception e5) {
			e5.printStackTrace();
		}
	}

	public static void main(String[] args) {
		//��������Ӷ��߳�ͬ���������һЩ���⣬û�е��Գɹ�������û����Ӷ��߳�
		Spider sp = new Spider("������");
		Spider sp2 = new Spider("���");
		Spider sp3 = new Spider("�㷨");
		// Thread t1 = new Thread(sp);
		// Thread t2 = new Thread(sp2);
		// Thread t3 = new Thread(sp3);
		try {
			sp.createExcel("D:/test.xls", "����ͼ��", sp.title);
		} catch (Exception e9) {
			e9.printStackTrace();
		}
		rowId = sp.getWebbookInfo();
		rowId = sp2.getWebbookInfo();
		rowId = sp3.getWebbookInfo();
		System.out.println("��ȡ��Ϣ���");
		System.out.println("���ڴ�ӡ���۸���1000����");
		sp.getFirstHunderedBooks("D:/test.xls", "����ͼ��", excelList);

	}

	// @Override
	// public synchronized void run() {
	// getWebbookInfo();
	// }
}
