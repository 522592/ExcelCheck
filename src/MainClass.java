import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

public class MainClass {

	/**
	 * @param args
	 */
	public static void main(String[] args) {
		MainClass m = new MainClass();
		m.ReadExcel();
	}

	/**
	 * �ҵ�ֵ�����룩���ڵ���
	 * 
	 * @param sheet
	 *            Ҫ�ҵ�excel���
	 * @param columnIndex
	 *            ��ֵ����excel�����±꣨��0��ʼ A-0 B-1...��
	 * @param value
	 *            Ҫ�ҵ�ֵ
	 * @return ֵ���ڵ��У���0��ʼ����excel�к�С1��
	 */
	int FindIndex(Sheet sheet, int columnIndex, String value) {
		int index = 0;
		for (int i = 0; i < sheet.getRows(); i++) {
			if (sheet.getCell(columnIndex, i).getContents().equals(value)) {
				index = i;
				break;
			}
		}
		return index;
	}

	// �ַ���תС��
	float StringParseFloat(String value) {
		float result = 0;
		if (value != null && !value.equals("")) {
			result = Float.parseFloat(value.replace(",", "").replace("%", ""));
		}
		return result;
	}

	// ͨ����Ӧ��ϵ���¾�ϵͳ�����ݽ��бȽ�
	void ReadExcel() {
		try {
			String path = "F:\\����\\981";
			if (!new File(path).exists()) {
				System.out.println(String.format("Ŀ¼������ %s", path));
				return;
			}
			String fileName1 = path + "\\�ʲ���ֵ��_20160621_000981_001.xls";
			File file1 = new File(fileName1);
			if (!file1.exists()) {
				System.out.println(String.format("��ϵͳ�ļ������� %s", fileName1));
				return;
			}
			String fileName2 = path + "\\�ʲ���ֵ��20160620-000981.xls";
			File file2 = new File(fileName2);
			if (!file2.exists()) {
				System.out.println(String.format("��ϵͳ�ļ������� %s", fileName2));
				return;
			}
			Sheet sheet1 = Workbook.getWorkbook(file1).getSheet(0);
			Sheet sheet2 = Workbook.getWorkbook(file2).getSheet(0);
			String fileName = path + "\\��Ӧ��ϵ.xls";
			File file = new File(fileName);
			if (!file.exists()) {
				System.out.println(String.format("��ϵ�ļ������� %s", fileName));
				System.out.println("��ʼ���ɶ�Ӧ��ϵ");
				GenerateRelation(sheet1,sheet2,fileName);
				System.out.println("��Ӧ��ϵ������ϣ���鿴��");
				return;
			}
			Sheet sheet = Workbook.getWorkbook(file).getSheet(0);
			for (int i = 1; i < sheet.getRows(); i++) {
				String code1 = sheet.getCell(1, i).getContents();
				String code2 = sheet.getCell(4, i).getContents();
				int r1 = FindIndex(sheet1, 0, code1);
				if (r1 == 0) {
					System.out.println(String.format("��ϵͳ��δ�ҵ� %s", code1));
				}
				int r2 = FindIndex(sheet2, 1, code2);
				if (r2 == 0) {
					System.out.println(String.format("��ϵͳ��δ�ҵ� %s", code2));
				}
				if (r1 != 0 && r2 != 0) {
					boolean notEquals = false;
					for (int j = 2; j < 10; j++) {
						String nu1 = sheet1.getCell(j, r1).getContents();
						float a1 = StringParseFloat(nu1);
						String nu2 = sheet2.getCell(j + 1, r2).getContents();
						float a2 = StringParseFloat(nu2);
						if (a1 != a2) {
							notEquals = true;
							System.out.println(String.format(
									"%s%s %s<>%s %s%s", r1 + 2,
									(char) ((int) 'A' + j), nu1, nu2, r2 + 2,
									(char) ((int) 'A' + j + 1)));
						}
					}
					if (notEquals) {
						System.out.println(String.format("%s -- %s", sheet1
								.getCell(1, r1).getContents(), sheet2.getCell(
								2, r2).getContents()));
						System.out.println();
					}
				}
			}
			System.out.println("OK");
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
/**
 * ���ɶ�Ӧ��ϵ
 * @param sheet1 ��ϵͳ���
 * @param sheet2 ��ϵͳ���
 * @param fileName ��Ӧ��ϵ�ļ�ȫ��
 */
	void GenerateRelation(Sheet sheet1, Sheet sheet2, String fileName) {
		List<String> list1 = FindCodes(sheet1, 4, 0);
		List<String> list2 = FindCodes(sheet2, 5, 1);
		int count = list1.size();
		if (list2.size() < count) {
			count = list2.size();
		}
		try {
			OutputStream os = new FileOutputStream(fileName);
			// ����������
			WritableWorkbook workbook = Workbook.createWorkbook(os);
			// �����µ�һҳ
			WritableSheet sheet = workbook.createSheet("Sheet1", 0);
			// ����Ҫ��ʾ������,����һ����Ԫ�񣬵�һ������Ϊ�����꣬�ڶ�������Ϊ�����꣬����������Ϊ����
			sheet.addCell(new Label(0, 0, "�к�1"));
			sheet.addCell(new Label(1, 0, "����1"));
			sheet.addCell(new Label(2, 0, ""));
			sheet.addCell(new Label(3, 0, "�к�2"));
			sheet.addCell(new Label(4, 0, "����2"));
			for (int i = 0; i < count; i++) {
				String[] con1 = list1.get(i).split("\t");
				String[] con2 = list2.get(i).split("\t");
				int row = i + 1;
				sheet.addCell(new Label(0, row, con1[0]));
				sheet.addCell(new Label(1, row, con1[1]));
				sheet.addCell(new Label(2, row, ""));
				sheet.addCell(new Label(3, row, con2[0]));
				sheet.addCell(new Label(4, row, con2[1]));
			}
			// �Ѵ���������д�뵽������У����ر������
			workbook.write();
			workbook.close();
			os.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	List<String> FindCodes(Sheet sheet, int startRowIndex, int columnIndex) {
		String lastNum = "1";
		int lastRowIndex = 0;
		List<String> result = new ArrayList<String>();
		for (int i = startRowIndex; i < sheet.getRows(); i++) {
			String num = sheet.getCell(columnIndex, i).getContents();
			if (num != null && num.length() > 0) {
				if (!num.startsWith(lastNum)) {
					result.add(String.format("%s\t%s", lastRowIndex + 1,
							lastNum));
				}
				lastNum = num;
				lastRowIndex = i;
			}
		}
		return result;
	}
}
