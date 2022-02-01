package test;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.crypto.Cipher;
import javax.crypto.spec.SecretKeySpec;
import javax.swing.*;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.security.Key;
import java.util.Base64;

public class AESEncryption {
    private static String algorithm = "AES";
    private static byte[] keyValue = new byte[] {'0','2','3','4','5','6','7','8','9','1','2','3','4','5','6','7'};

    public static String encypt (String plainText) throws Exception{
        Key key = generateKey();
        Cipher cipher = Cipher.getInstance(algorithm);
        cipher.init(Cipher.ENCRYPT_MODE,key);
        byte [] encVal = cipher.doFinal(plainText.getBytes());
        String encryptedVal = new String(Base64.getEncoder().encode(encVal));
        return encryptedVal;
    }

    public static String decrypt (String encryptedText) throws Exception{
        Key key = generateKey();
        Cipher cipher = Cipher.getInstance(algorithm);
        cipher.init(Cipher.DECRYPT_MODE,key);
        byte [] decodedVal = Base64.getDecoder().decode(encryptedText);
        byte [] decryptedVal = cipher.doFinal(decodedVal);
        String decVal = new String(decodedVal);
        return  decVal;

    }

    public static Key generateKey() throws Exception{
        Key key = new SecretKeySpec(keyValue,algorithm);
        return key;
    }

    private String getPassword()
    {
        JPasswordField pwd = new JPasswordField(10);
        String passwordTxt = null;
        int action = JOptionPane.showConfirmDialog(null,pwd,"Enter Password",JOptionPane.OK_CANCEL_OPTION);
        if(action<0)JOptionPane.showMessageDialog(null,"Cancel, X or escape key Selected");
        else
            passwordTxt = new String(pwd.getPassword());

        return passwordTxt;
    }

    public static void main(String args[]) throws Exception {
        FileInputStream inputStream;
        Workbook workbook = null;
        Sheet sh = null;

        String fileName = null;
        String passTxt = new AESEncryption().getPassword();
        String encryptedText = AESEncryption.encypt(passTxt);
        JTextArea encryptedText1 = new JTextArea();
        encryptedText1.setText(encryptedText);
        encryptedText1.setEditable(true);
        JOptionPane.showMessageDialog(null,encryptedText1);

        fileName = System.getProperty("user.dir")+"\\Test.xlsx";

        try
        {
            inputStream = new FileInputStream(new File(fileName));

            workbook = new XSSFWorkbook(inputStream);
            sh = workbook.getSheetAt(0);

            for(int x = 1;x<sh.getPhysicalNumberOfRows();x++)
            {
                Row row = sh.getRow(x);

                String plainText = row.getCell(0).getStringCellValue();
                encryptedText = AESEncryption.encypt(plainText);
                Cell cell = row.createCell(1);
                System.out.println("Encrypted Text: "+ encryptedText);
                cell.setCellValue(encryptedText);


            }
        } catch(FileNotFoundException e){e.printStackTrace();  }
        FileOutputStream fos = new FileOutputStream(fileName);
        fos.flush();
        workbook.write(fos);
        fos.close();
        workbook.close();


    }

}
