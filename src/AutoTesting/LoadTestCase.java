package AutoTesting;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class LoadTestCase {
	public ArrayList<String> StepList = new ArrayList<String>();// �Ҧ����ծרҪ��ʧ@�M��
	public ArrayList<String> CaseList = new ArrayList<String>();// �Ҧ����ծרҪ��W�ٲM��
	LoadWebInfor DeviceInformation = new LoadWebInfor();

	public LoadTestCase() {
		XSSFWorkbook workbook = null;
		XSSFSheet sheet;

		try {
			workbook = new XSSFWorkbook(new FileInputStream("C:\\TUTK_QA_TestTool\\TestTool\\Web_TestScrpit.xlsm"));

			CaseList = new ArrayList<String>();
			StepList = new ArrayList<String>();
			for (int k = 0; k < DeviceInformation.ScriptList.size(); k++) {

				sheet = workbook.getSheet(DeviceInformation.ScriptList.get(k).toString());// ���w�ݴ��ո}����sheet
				int i = 0;
				try {
					do {// column Number

						// System.out.println(sheet.getRow(i).getPhysicalNumberOfCells());

						for (int j = 0; j < sheet.getRow(i).getPhysicalNumberOfCells(); j++) {

							if (sheet.getRow(i).getCell(j) != null) {// Apache
																		// POI
																		// Ū��Excel�x�s��ɡA�����v�N�ť��x�s��Ū�J�A�]���ݧP�_�x�s��O�_���ťաA��null

								if (sheet.getRow(i).getCell(j).toString().equals("CaseName")) {
									CaseList.add(sheet.getRow(i).getCell(1).toString());// �q���w�ݴ��ո}����sheet���x�s���ծרҪ��W��
									break;
								} else {

									// StepList.add(sheet.getRow(i).getCell(j).toString());//�q���w�ݴ��ո}����sheet���x�s���ծרҪ��B�J
									StepList.add(sheet.getRow(i).getCell(j).getStringCellValue());// �q���w�ݴ��ո}����sheet���x�s���ծרҪ��B�J
																								// Excel�Ʀr�n�ন�r�ꫬ�A
								}
							}
						}

						i++;
					} while (!sheet.getRow(i).getCell(0).toString().equals(""));
				} catch (Exception e) {
					;
				}
			}

		} catch (Exception e) {
			;
		}

		System.out.println( "���ըB�J�G" + StepList);
		// �إߦU�˸m��Test Report
		
		
		for (int i = 0; i < DeviceInformation.BrowserList.size(); i++) {

			if (DeviceInformation.BrowserList.get(i).toString().length() > 20) {// Excel�u�@��W�ٳ̱`31�r���]�A�G�ݧP�_UDID���׬O�_�j��31
				char[] NewUdid = new char[20];// �]�ݥ]�t_TestReport�r��(�@11�r��)�A�G�]�w20��r���}�C(31-11)
				DeviceInformation.BrowserList.get(i).toString().getChars(0, 20, NewUdid, 0);// ���XUDID�e20�r����NewUdid
				sheet = workbook.createSheet(String.valueOf(NewUdid) + "_TestReport");// �ϥ�NewUdid�R�W�s�u�@��
			} else {
				sheet = workbook.createSheet(DeviceInformation.BrowserList.get(i).toString() + "_TestReport");
			}

			sheet.createRow(0).createCell(0).setCellValue("CaseName");
			sheet.getRow(0).createCell(1).setCellValue("Result");

			for (int k = 0; k < CaseList.size(); k++) {// write case name
				sheet.createRow(k + 1).createCell(0).setCellValue(CaseList.get(k).toString());
			}
		}
		
/*
		if (DeviceInformation.Browser.toString().length() > 20) {// Excel�u�@��W�ٳ̱`31�r���]�A�G�ݧP�_UDID���׬O�_�j��31
			char[] NewUdid = new char[20];// �]�ݥ]�t_TestReport�r��(�@11�r��)�A�G�]�w20��r���}�C(31-11)
			DeviceInformation.Browser.toString().getChars(0, 20, NewUdid, 0);// ���XUDID�e20�r����NewUdid
			sheet = workbook.createSheet(String.valueOf(NewUdid) + "_TestReport");// �ϥ�NewUdid�R�W�s�u�@��
		} else {
			sheet = workbook.createSheet(DeviceInformation.Browser.toString() + "_TestReport");
		}

		sheet.createRow(0).createCell(0).setCellValue("CaseName");
		sheet.getRow(0).createCell(1).setCellValue("Result");

		for (int k = 0; k < CaseList.size(); k++) {// write case name
			sheet.createRow(k + 1).createCell(0).setCellValue(CaseList.get(k).toString());

		}
*/
		// ����g�JExcel�᪺�s�ɰʧ@
		FileOutputStream out;
		try {
			out = new FileOutputStream(new File("C:\\TUTK_QA_TestTool\\TestReport\\Web_TestReport.xlsm"));// �t�s�s��
			workbook.write(out);
			out.close();
			workbook.close();
		} catch (Exception e) {
			;
		}

	}
	
}
