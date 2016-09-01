package controller;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import model.User;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

public class CreateExcel {
	 private static final String fileName = "C:/apachepoi/usuarios.xls";
	   
     public static void main(String[] args) throws IOException {

           HSSFWorkbook workbook = new HSSFWorkbook();
           HSSFSheet sheetUsers = workbook.createSheet("Usuarios");
 
           List<User> lisUsers = new ArrayList<User>();
           lisUsers.add(new User(1, "Bruno", "bruno@gmail.com","senha", 9876525, "Aclimacao, 155"));
           lisUsers.add(new User(1, "Eduardo", "eduardo@gmail.com","senha", 9876525, "Santana, 520"));
           lisUsers.add(new User(1, "Carol", "carol@gmail.com","senha", 9876525, "Itaquera, 302"));
           lisUsers.add(new User(1, "Marcia", "marcia@gmail.com","senha", 9876525, "Sapopemba, 49"));
           lisUsers.add(new User(1, "Gustavo", "gustavo@gmail.com","senha", 9876525, "Vila Prudente, 155"));
           lisUsers.add(new User(1, "Rafael", "eduardo@gmail.com","senha", 9876525, "Vila Ema, 275"));
           
            
           int rownum = 0;
           for (User user : lisUsers) {
               Row row = sheetUsers.createRow(rownum++);
               int cellnum = 0;
               Cell cellId = row.createCell(cellnum++);
               cellId.setCellValue(user.getId());
               Cell cellNome = row.createCell(cellnum++);
               cellNome.setCellValue(user.getName());
               Cell cellEmail = row.createCell(cellnum++);
               cellEmail.setCellValue(user.getEmail());
               Cell cellSenha = row.createCell(cellnum++);
               cellSenha.setCellValue(user.getPassword());
               Cell cellTelefone = row.createCell(cellnum++);
               cellTelefone.setCellValue(user.getTelefone());
               Cell cellAdress = row.createCell(cellnum++);
               cellAdress.setCellValue(user.getAdress());  
           }
            
           try {
               FileOutputStream out = 
                       new FileOutputStream(new File(CreateExcel.fileName));
               workbook.write(out);
               out.close();
               System.out.println("Arquivo Excel criado com sucesso!");
                
           } catch (FileNotFoundException e) {
               e.printStackTrace();
                  System.out.println("Arquivo não encontrado!");
           } catch (IOException e) {
               e.printStackTrace();
                  System.out.println("Erro na edição do arquivo!");
           }
     }
}