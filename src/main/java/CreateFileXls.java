import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class CreateFileXls {

  private static final String fileName = "/home/rafaelbanhos/testfile.xls"; //change folder in your computer

  public static void main(String[] args) throws IOException {

    HSSFWorkbook workbook = new HSSFWorkbook();
    HSSFSheet sheetAlunos = workbook.createSheet("Alunos");

    List<Student> listaAlunos = new ArrayList<Student>();
    listaAlunos.add(new Student("Rafael", "9876525", 7, 8, 0, false));
    listaAlunos.add(new Student("Chico", "1234466", 5, 8, 0, false));
    listaAlunos.add(new Student("Jerismonio", "6545657", 7, 6, 0, false));
    listaAlunos.add(new Student("Shiroskava", "3456558", 10, 3, 0, false));
    listaAlunos.add(new Student("Raspyxerry", "6544546", 7, 8, 0, false));
    listaAlunos.add(new Student("Herbet Richards", "3234535", 6, 5, 0, false));
    listaAlunos.add(new Student("China Number one", "4234524", 7, 5, 0, false));
    listaAlunos.add(new Student("Mostramberg", "5434513", 7, 2, 0, false));
    listaAlunos.add(new Student("Meslovaldo", "6543452", 7, 8, 0, false));
    listaAlunos.add(new Student("Chesvirato", "4345651", 5, 8, 0, false));
    listaAlunos.add(new Student("Gotardo", "4332341", 7, 9, 0, false));

    int rownum = 0;
    for (Student aluno : listaAlunos) {
      HSSFRow row = sheetAlunos.createRow(rownum++);
      int cellnum = 0;
      Cell cellNome = row.createCell(cellnum++);
      cellNome.setCellValue(aluno.getNome());
      Cell cellRa = row.createCell(cellnum++);
      cellRa.setCellValue(aluno.getRa());
      Cell cellNota1 = row.createCell(cellnum++);
      cellNota1.setCellValue(aluno.getNota1());
      Cell cellNota2 = row.createCell(cellnum++);
      cellNota2.setCellValue(aluno.getNota2());
      Cell cellMedia = row.createCell(cellnum++);
      cellMedia.setCellValue((aluno.getNota1() + aluno.getNota2()) / 2);
      Cell cellAprovado =row.createCell(cellnum++);
      cellAprovado.setCellValue(cellMedia.getNumericCellValue() >= 6);
    }

    try {
      FileOutputStream out =
          new FileOutputStream(new File(CreateFileXls.fileName));
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