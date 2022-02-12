package Server;

import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.net.Socket;
import java.util.Scanner;

public class Client {

	public static void main(String[] args) {
		Scanner sc = new Scanner(System.in);
		String st = sc.nextLine();
		try {
			Socket soc = new Socket("localhost",3000);
			DataOutputStream dos = new DataOutputStream(soc.getOutputStream());
			dos.writeUTF(st);
			DataInputStream dis = new DataInputStream(soc.getInputStream());
			System.out.println(dis.readUTF());
		} catch (Exception e) {
			// TODO: handle exception
		}
	}

}

