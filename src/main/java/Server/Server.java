package Server;

import java.io.DataInputStream;
import java.io.DataOutputStream;
import java.net.ServerSocket;
import java.net.Socket;
import java.util.Stack;

public class Server {

	public static void main(String[] args) {
		try {
			ServerSocket server = new ServerSocket(3000);
			while(true) {
				Socket soc = server.accept();
				DataInputStream dis =new DataInputStream(soc.getInputStream());
				String s=dis.readUTF();
				DataOutputStream dos = new DataOutputStream(soc.getOutputStream());
				try {
					dos.writeUTF(evaluatePostfix(infixToPostfix(s))+"");
				} catch (Exception e) {
					dos.writeUTF("Biểu thức không chính xác!");
				}
				
			}
		} catch (Exception e) {
			// TODO: handle exception
		}

	}
	static int Prec(char ch) 
    { 
        switch (ch) 
        { 
        case '+': 
        case '-': 
            return 1; 
       
        case '*': 
        case '/': 
            return 2; 
       
        case '^': 
            return 3; 
        } 
        return -1; 
    } 
       
    // The main method that converts given infix expression 
    // to postfix expression.  
    static String infixToPostfix(String exp) 
    { 
    	//if (!Character.isLetterOrDigit(exp.charAt(0))) exp='0'+exp;
        // initializing empty String for result 
        String result = new String(""); 
          
        // initializing empty stack 
        Stack<Character> stack = new Stack<>(); 
          
        for (int i = 0; i<exp.length(); ++i) 
        { 
            char c = exp.charAt(i); 
             // if negative
            if (c=='-' && exp.charAt(i-1)=='(') result+=" 0";
             // If the scanned character is an operand, add it to output. 
            if (Character.isLetterOrDigit(c) || c=='.') 
                result +=c; 
               
            // If the scanned character is an '(', push it to the stack. 
            else if (c == '(') 
                stack.push(c); 
              
            //  If the scanned character is an ')', pop and output from the stack  
            // until an '(' is encountered. 
            else if (c == ')') 
            { 
            	if (result.charAt(result.length()-1)!=' ') result+=" ";
                while (!stack.isEmpty() && stack.peek() != '(') 
                    result += stack.pop()+" "; 
                  
                if (!stack.isEmpty() && stack.peek() != '(') 
                    return "Invalid Expression"; // invalid expression                 
                else
                    stack.pop(); 
            } 
            else // an operator is encountered 
            { 
            	if (result.charAt(result.length()-1)!=' ') result+=" ";  
                while (!stack.isEmpty() && Prec(c) <= Prec(stack.peek())){ 
                    if(stack.peek() == '(') 
                        return "Invalid Expression"; 
                    result += stack.pop()+" "; 
             } 
                stack.push(c); 
            } 
       
        } 
       
        // pop all the operators from the stack 
        if (result.charAt(result.length()-1)!=' ') result+=" ";
        while (!stack.isEmpty()){ 
            if(stack.peek() == '(') 
                return "Invalid Expression"; 
            result +=stack.pop()+" "; 
         } 
        return result; 
    } 
    static double evaluatePostfix(String exp) 
    { 
        //create a stack 
        Stack<Double> stack = new Stack<>(); 
        System.out.println(exp);  
        // Scan all characters one by one 
        for(int i = 0; i < exp.length(); i++) 
        { 
            char c = exp.charAt(i); 
              
            if(c == ' ') 
            continue; 
              
            // If the scanned character is an operand  
            // (number here),extract the number 
            // Push it to the stack. 
            else if(Character.isDigit(c) || c=='.') 
            { 
                //int n = 0; 
                String s="";
                  
                //extract the characters and store it in num 
                while(Character.isDigit(c) || c=='.') 
                { 
                    //n = n*10 + (int)(c-'0'); 
                    s+=c;
                    i++; 
                    c = exp.charAt(i); 
                } 
                i--; 
  
                //push the number in stack 
                //System.out.println(s);
                stack.push(Double.parseDouble(s)); 
            } 
              
            // If the scanned character is an operator, pop two 
            // elements from stack apply the operator 
            else
            { 
                double val1 = stack.pop(); 
                double val2 = stack.pop(); 
                  
                switch(c) 
                { 
                    case '+': 
                    stack.push(val2+val1); 
                    break; 
                      
                    case '-': 
                    stack.push(val2- val1); 
                    break; 
                      
                    case '/': 
                    stack.push(val2/val1); 
                    break; 
                      
                    case '*': 
                    stack.push(val2*val1); 
                    break; 
            } 
            } 
        } 
        return stack.pop();  
    }
}


