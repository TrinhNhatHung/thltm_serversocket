package Server;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.net.ServerSocket;
import java.net.Socket;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;
import java.util.List;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.ss.util.SheetUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import ToMau.TM;

public class DanhSachServer {

	public static void main(String[] args) {
		try {
			ServerSocket server = new ServerSocket(4000);
			while (true) {
				Socket soc = server.accept();
				new Xulydanhsach(soc).run();
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

}

class Xulydanhsach extends Thread {
	Socket soc;

	public Xulydanhsach(Socket s) {
		this.soc = s;
	}

	public void run() {
		try {
			System.out.println("Nhận dữ liệu:");
			DataInputStream dis = new DataInputStream(soc.getInputStream());
			byte[] pc, gs;
			XSSFWorkbook gv = null, pt = null;

			int n = dis.readInt();
			System.out.println(n);

			int length = dis.readInt();
			if (length > 0) {
				// System.out.println(length);
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

			List<List<Integer>> ds = TM.xuly(n, gv.getSheetAt(0).getLastRowNum(), pt.getSheetAt(0).getLastRowNum());
//			for (List<Integer> list : ds) {
//				System.out.println(list);
//				System.out.println(list.size());
//			}

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
				cell.setCellValue("Mã GV");
				CellUtil.setCellStyleProperty(cell, CellUtil.VERTICAL_ALIGNMENT, VerticalAlignment.CENTER);
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				mergedRegion = new CellRangeAddress(0, 1, 2, 2);
				sheet.addMergedRegion(mergedRegion);
				sheet.setColumnWidth(2, 8000);
				cell = row.createCell(2);
				cell.setCellValue("Họ và tên");
				CellUtil.setCellStyleProperty(cell, CellUtil.VERTICAL_ALIGNMENT, VerticalAlignment.CENTER);
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				mergedRegion = new CellRangeAddress(0, 0, 3, 4);
				sheet.addMergedRegion(mergedRegion);
				cell = row.createCell(3);
				cell.setCellValue("Giám thị");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				mergedRegion = new CellRangeAddress(0, 1, 5, 5);
				sheet.addMergedRegion(mergedRegion);
				sheet.setColumnWidth(5, 3000);
				cell = row.createCell(5);
				cell.setCellValue("Phòng Thi");
				CellUtil.setCellStyleProperty(cell, CellUtil.VERTICAL_ALIGNMENT, VerticalAlignment.CENTER);
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				row = sheet.createRow(++rowCount);
				sheet.setColumnWidth(3, 3000);
				cell = row.createCell(3);
				cell.setCellValue("Giám thị 1");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
				sheet.setColumnWidth(4, 3000);
				cell = row.createCell(4);
				cell.setCellValue("Giám thị 2");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);
				//
				int stt = 0;
				for (int i : list) {
					row = sheet.createRow(++rowCount);
					cell = row.createCell(0);
					cell.setCellValue(++stt);
					CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

					cell = row.createCell(1);
					cell.setCellValue(xuat(gv.getSheetAt(0).getRow(i).getCell(3)));
					CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

					cell = row.createCell(2);
					cell.setCellValue(xuat(gv.getSheetAt(0).getRow(i).getCell(1)));
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

				// Danh sach giam sat
				sheet = dsgs.createSheet("Ca thi " + numsheet);
				rowCount = -1;
				row = sheet.createRow(++rowCount);
				// set up header for sheet
				cell = row.createCell(0);
				cell.setCellValue("STT");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				cell = row.createCell(1);
				cell.setCellValue("Mã GV");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				sheet.setColumnWidth(2, 8000);
				cell = row.createCell(2);
				cell.setCellValue("Họ và tên");
				CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

				sheet.setColumnWidth(3, 8000);
				cell = row.createCell(3);
				cell.setCellValue("Phòng thi");
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
					cell.setCellValue(xuat(gv.getSheetAt(0).getRow(i).getCell(3)));
					CellUtil.setCellStyleProperty(cell, CellUtil.ALIGNMENT, HorizontalAlignment.CENTER);

					cell = row.createCell(2);
					cell.setCellValue(xuat(gv.getSheetAt(0).getRow(i).getCell(1)));
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

			DataOutputStream dos = new DataOutputStream(soc.getOutputStream());
			dos.writeInt(pc.length);
			dos.write(pc);
			System.out.println(pc.length);
			dos.writeInt(gs.length);
			dos.write(gs);
			System.out.println(gs.length);
		} catch (

		Exception e) {
			// TODO: handle exception
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