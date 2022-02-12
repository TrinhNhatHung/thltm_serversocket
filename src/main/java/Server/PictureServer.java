package Server;

import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.net.ServerSocket;
import java.net.Socket;

import javax.imageio.ImageIO;

public class PictureServer {

	public static void main(String[] args) {
		try {
			ServerSocket server = new ServerSocket(4000);
			while (true) {
				Socket soc = server.accept();
				new Xuly(soc).run();
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

}

class Xuly extends Thread {
	Socket soc;

	public Xuly(Socket s) {
		this.soc = s;
	}

	public void run() {
		try {
			DataInputStream dis = new DataInputStream(soc.getInputStream());
			int length = dis.readInt();
			if (length > 0) {
				byte[] bytes = new byte[length];
				dis.readFully(bytes, 0, length);
				ByteArrayInputStream bais = new ByteArrayInputStream(bytes);
				BufferedImage img = ImageIO.read(bais);

				// get image width and height
				int width = img.getWidth();
				int height = img.getHeight();

				// convert to grayscale
				for (int y = 0; y < height; y++) {
					for (int x = 0; x < width; x++) {
						int p = img.getRGB(x, y);

						int a = (p >> 24) & 0xff;
						int r = (p >> 16) & 0xff;
						int g = (p >> 8) & 0xff;
						int b = p & 0xff;

						// calculate average
						int avg = (r + g + b) / 3;

						// replace RGB value with avg
						p = (a << 24) | (avg << 16) | (avg << 8) | avg;

						img.setRGB(x, y, p);
					}
				}

				ByteArrayOutputStream baos = new ByteArrayOutputStream();
				ImageIO.write(img, "jpg", baos);
				bytes = baos.toByteArray();

				DataOutputStream dos = new DataOutputStream(soc.getOutputStream());
				dos.writeInt(bytes.length);
				dos.write(bytes);
			}
		} catch (Exception e) {
			// TODO: handle exception
		}
	}
}