import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

class CheckProtectionTest {
public static void main(String[] args) {
    String path = "C:\\Users\\Duy\\Desktop\\read no password.xlsx";
    CheckProtection cp = new CheckProtection();
    if(cp.isEncrypted(path)== true) {
        System.out.println("please enter password to open this file");
        cp.inputPassword(path);
    }
    else
        System.out.println("ko co khoa");
    }
}