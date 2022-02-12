package main;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.net.ServerSocket;
import java.net.Socket;
import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ProcessDataServer {

	public static void main(String[] args) {
		try {
			ServerSocket socketServer = new ServerSocket(4000);
			while (true) {
				Socket soc = socketServer.accept();
				new ProcessData(soc).run();
			}
		} catch (Exception e) {
			System.out.println(e.getMessage());
		}
	}

}

class ProcessData extends Thread {
	Socket socket;

	public ProcessData(Socket socket) {
		this.socket = socket;
	}
	
	public static List<List<Integer>> handle(int t, int n, int m) {
		List<List<Integer>> result = new ArrayList<>();
		List<Integer> lecturerIndexes = new ArrayList<>();
		for (int i=1; i<=n; i++) {
			lecturerIndexes.add(i);
		}
		Collections.shuffle(lecturerIndexes);
		Set<String> checkClasses = new HashSet<>();
		Set<String> checkLectureTogether = new HashSet<>();
		while (t>0) {
			t--;
			List<Integer> session = new ArrayList<>();
			boolean[] isLectureArranged = new boolean[n+1];
			for (int classIndex=1; classIndex<=m; classIndex++) {
				Integer firstLectureIndex = -1;
				for (Integer lectureIndex : lecturerIndexes) {
					String s = lectureIndex + "," + classIndex;
					if (!isLectureArranged[lectureIndex] && !checkClasses.contains(s)) {
						checkClasses.add(s);
						session.add(lectureIndex);
						isLectureArranged[lectureIndex] = true;
						firstLectureIndex = lectureIndex;
						break;
					}
				}
				for (Integer lectureIndex : lecturerIndexes) {
					String s = lectureIndex + "," + classIndex;
					String lectureTogether = firstLectureIndex + "&" + lectureIndex;
					if (!isLectureArranged[lectureIndex] && !checkClasses.contains(s) && !checkLectureTogether.contains(lectureTogether)) {
						checkClasses.add(s);
						checkLectureTogether.add(lectureTogether);
						session.add(lectureIndex);
						isLectureArranged[lectureIndex]=true;
						break;
					}
				}
			}
			result.add(session);
			Collections.shuffle(lecturerIndexes);
		}
		return result;
	}

	public void run() {
		try {
			DataInputStream dis = new DataInputStream(socket.getInputStream());
			byte[] pc, gs;
			XSSFWorkbook gv = null, pt = null;

			int n = dis.readInt();
			System.out.println(n);

			int length = dis.readInt();
			if (length > 0) {
				byte[] bytes = new byte[length];
				dis.readFully(bytes, 0, length);
				ByteArrayInputStream bais = new ByteArrayInputStream(bytes);
				gv = new XSSFWorkbook(bais);
				System.out.println(gv.getSheetAt(0).getLastRowNum());
			}

			length = dis.readInt();
			if (length > 0) {
				// System.out.println(length);
				byte[] bytes = new byte[length];
				dis.readFully(bytes, 0, length);
				ByteArrayInputStream bais = new ByteArrayInputStream(bytes);
				pt = new XSSFWorkbook(bais);
				System.out.println(pt.getSheetAt(0).getLastRowNum());
			}

			List<List<Integer>> ds = handle(n, gv.getSheetAt(0).getLastRowNum(), pt.getSheetAt(0).getLastRowNum());

			XSSFWorkbook dspc = new XSSFWorkbook();
			XSSFWorkbook dsgs = new XSSFWorkbook();
			int numsheet = 0;
			for (List<Integer> list : ds) {
				// Danh sach phan cong
				XSSFSheet sheet = dspc.createSheet("Ca thi " + (++numsheet));
				int rowCount = -1;
				XSSFRow row;
				XSSFCell cell;
				row = sheet.createRow(++rowCount);
				// set up header for sheet
				CellRangeAddress mergedRegion = new CellRangeAddress(0, 1, 0, 0);
				sheet.addMergedRegion(mergedRegion);
				cell = row.createCell(0);
				cell.setCellValue("STT");
				CellUtil.setCellStyleProperty(cell, CellUtil.VERTICAL_ALIGNMENT, VerticalAlignment.CENTER);
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				mergedRegion = new CellRangeAddress(0, 1, 1, 1);
				sheet.addMergedRegion(mergedRegion);
				cell = row.createCell(1);
				cell.setCellValue("Ma GV");
				CellUtil.setCellStyleProperty(cell, CellUtil.VERTICAL_ALIGNMENT, VerticalAlignment.CENTER);
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				mergedRegion = new CellRangeAddress(0, 1, 2, 2);
				sheet.addMergedRegion(mergedRegion);
				sheet.setColumnWidth(2, 8000);
				cell = row.createCell(2);
				cell.setCellValue("Ho va ten");
				CellUtil.setCellStyleProperty(cell, CellUtil.VERTICAL_ALIGNMENT, VerticalAlignment.CENTER);
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				mergedRegion = new CellRangeAddress(0, 0, 3, 4);
				sheet.addMergedRegion(mergedRegion);
				cell = row.createCell(3);
				cell.setCellValue("Giam thi");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				mergedRegion = new CellRangeAddress(0, 1, 5, 5);
				sheet.addMergedRegion(mergedRegion);
				sheet.setColumnWidth(5, 3000);
				cell = row.createCell(5);
				cell.setCellValue("Phong thi");
				CellUtil.setCellStyleProperty(cell, CellUtil.VERTICAL_ALIGNMENT, VerticalAlignment.CENTER);
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				row = sheet.createRow(++rowCount);
				sheet.setColumnWidth(3, 3000);
				cell = row.createCell(3);
				cell.setCellValue("Giam thi 1");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
				sheet.setColumnWidth(4, 3000);
				cell = row.createCell(4);
				cell.setCellValue("Giam thi 2");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
				//
				int stt = 0;
				for (int i : list) {
					row = sheet.createRow(++rowCount);
					cell = row.createCell(0);
					cell.setCellValue(++stt);
					CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

					cell = row.createCell(1);
					cell.setCellValue(xuat(gv.getSheetAt(0).getRow(i).getCell(1)));
					CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

					cell = row.createCell(2);
					cell.setCellValue(xuat(gv.getSheetAt(0).getRow(i).getCell(2)));
					CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

					if (stt % 2 != 0) {
						cell = row.createCell(3);
						cell.setCellValue("X");
						CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
					} else {
						cell = row.createCell(4);
						cell.setCellValue("X");
						CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
					}

					cell = row.createCell(5);
					cell.setCellValue(xuat(pt.getSheetAt(0).getRow((stt + 1) / 2).getCell(1)));
					CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
				}

				sheet = dsgs.createSheet("Ca thi " + numsheet);
				rowCount = -1;
				row = sheet.createRow(++rowCount);
				cell = row.createCell(0);
				cell.setCellValue("STT");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				cell = row.createCell(1);
				cell.setCellValue("Ma GV");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				sheet.setColumnWidth(2, 8000);
				cell = row.createCell(2);
				cell.setCellValue("Ho va ten");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				sheet.setColumnWidth(3, 8000);
				cell = row.createCell(3);
				cell.setCellValue("Phong thi");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				int a = gv.getSheetAt(0).getLastRowNum();
				int b = pt.getSheetAt(0).getLastRowNum();
				list = loc(a, b, list);
				int u=0;
				int bn=5;
				if (a-2*b<b) bn=b/(a-2*b)+1;
				stt = 0;
				for (int i : list) {
					row = sheet.createRow(++rowCount);
					cell = row.createCell(0);
					cell.setCellValue(++stt);
					CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

					cell = row.createCell(1);
					cell.setCellValue(xuat(gv.getSheetAt(0).getRow(i).getCell(1)));
					CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

					cell = row.createCell(2);
					cell.setCellValue(xuat(gv.getSheetAt(0).getRow(i).getCell(2)));
					CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
					
					u++;
					if (u>b) u=1;
					String s="Từ "+xuat(pt.getSheetAt(0).getRow(u).getCell(1))+" đến ";
					u=u+bn-1;
					if (u>b) u=b;
					s+=xuat(pt.getSheetAt(0).getRow(u).getCell(1));
					cell = row.createCell(3);
					cell.setCellValue(s);
					CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
				}
			}

			ByteArrayOutputStream baos1 = new ByteArrayOutputStream();
			dspc.write(baos1);
			pc = baos1.toByteArray();
			ByteArrayOutputStream baos2 = new ByteArrayOutputStream();
			dsgs.write(baos2);
			gs = baos2.toByteArray();

			DataOutputStream dos = new DataOutputStream(socket.getOutputStream());
			dos.writeInt(pc.length);
			dos.write(pc);
			System.out.println(pc.length);
			dos.writeInt(gs.length);
			dos.write(gs);
			System.out.println(gs.length);
		} catch (

		Exception e) {
			System.out.println(e.getMessage());
		}
	}

	private List<Integer> loc(int a, int b, List<Integer> list) {
		boolean[] kt = new boolean[a + 1];
		for (Integer i : list) {
			kt[i]=true;
		}
		List<Integer> res = new ArrayList<>();
		for (int i=1;i<=a;i++) 
			if (!kt[i]) res.add(i);
		Collections.shuffle(res);
		return res;
	}

	private String xuat(XSSFCell cell) {
		switch (cell.getCellType()) {
		case STRING:
			return cell.getStringCellValue();
		case BOOLEAN:
			return cell.getBooleanCellValue() + "";
		case NUMERIC:
			return NumberToTextConverter.toText(cell.getNumericCellValue());
		default:
			return "";
		}
	}
}