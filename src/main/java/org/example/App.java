package org.example;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

/**
 Класс, который считывает данные из таблицы Excel и сравнивает их с введёными данными
 */
public class App {
    /**
      * В методе Main требуется ввод данных с клавиатуры, используя класс Scanner.
      * Для более быстрой проверки метод findNumbersFromExcel протестирован
     */
    public static void main(String[] args) {
        Scanner input = new Scanner(System.in);
        System.out.println("Введите три значения (русские символы, заглавные буквы):");
        String value1 = input.next();
        String value2 = input.next();
        String value3 = input.next();
        outputResult(findNumbersFromExcel("name_java.xlsx", value1, value2, value3));
    }
    /**
     *Метод, возвращающий список строк. При его вызове открывается файл, создаётся книга, произволится проход по ячейкам
     * (игнорируя первую строку с названием), сравниваются значения. EqualsIgnoreCase() по заданию жёстко требуется, и значения
     * представляют собой просто набор символов, поэтому оставил просто equals(). Список возвращается с записями, найдены или не
     * найдены значения. Метод протестирован
     */
    public static List<String> findNumbersFromExcel(String fileName, String value1, String value2, String value3) {
          List<String> data = new ArrayList<>();
        try (FileInputStream file = new FileInputStream(fileName)) {
            String message1 = "Номер " + value1 + " не найден";
            String message2 = "Номер " + value2 + " не найден";
            String message3 = "Номер " + value3 + " не найден";
            String newMessage1 = "Номер " + value1 + " найден";
            String newMessage2 = "Номер " + value2 + " найден";
            String newMessage3 = "Номер " + value3 + " найден";
            try (Workbook workbook = new XSSFWorkbook(file)) {
                data.add(message1);
                data.add(message2);
                data.add(message3);
                if (value1.equals("Name") || (value2.equals("Name") || (value3.equals("Name")))) {
                    data.add("Name является названием столбца");
                }
                Sheet sheet = workbook.getSheetAt(0);
                for (Row row : sheet) {
                    if (row.getRowNum() == 0) {
                        continue;
                    }
                    Cell cell = row.getCell(0);
                    if (value1.equals(cell.getStringCellValue()) && data.get(0).equals(message1)) {
                        data.set(0, newMessage1);
                    }
                    if (value2.equals(cell.getStringCellValue()) && data.get(1).equals(message2)) {
                        data.set(1, newMessage2);
                    }
                    if (value3.equals(cell.getStringCellValue()) && data.get(2).equals(message3)) {
                        data.set(2, newMessage3);
                    }

                }
            } catch (EncryptedDocumentException ex) {
                ex.printStackTrace();
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        return data;
    }
    /**
     Метод вывода не списка из предыдущего метода, а а более удобном виде
     */
    public static void outputResult(List<String> list) {
        System.out.println("Результаты поиска:");
        for (String i : list) {
            System.out.println(i);
        }
    }
}
