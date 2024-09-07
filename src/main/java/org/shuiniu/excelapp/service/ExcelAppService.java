package org.shuiniu.excelapp.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.text.SimpleDateFormat;
import java.util.HashSet;
import java.util.Set;

@Service
public class ExcelAppService {
    private final static String TMP_EXCEL = "/app/tmp.xlsx";

    private String currentFile = "";

    private static final Set<Integer> dateTimeIndexes = new HashSet<>();

    static {
        dateTimeIndexes.add(7);
        dateTimeIndexes.add(17);
        dateTimeIndexes.add(22);
        dateTimeIndexes.add(19);
        dateTimeIndexes.add(24);
    }

    public String saveExcelFile(MultipartFile file) throws IOException {
        Path path = Paths.get(TMP_EXCEL);
        currentFile = TMP_EXCEL;
        // 确保目录存在
        Files.createDirectories(path.getParent());
        // 文件写入指定路径
        file.transferTo(path);

        return "'" + file.getOriginalFilename() + "' 文件上传成功";
    }

    public String getUniqueKeyFile() {
        if (currentFile == null) {
            return "请选择要处理的Excel文件";
        }

        String outputDirectory = "/app/files";
        File directory = new File(outputDirectory);
        if (!directory.exists()) {
            directory.mkdirs();
        }

        String fileName = "unique_keys_output" + System.currentTimeMillis() + ".xlsx";
        String absoluteFilePath = Paths.get(outputDirectory, fileName).toString();
        try (FileInputStream fis = new FileInputStream(currentFile);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fos = new FileOutputStream(absoluteFilePath)) {
            Sheet sheet = workbook.getSheetAt(0);
            Set<String> keySet = new HashSet<>();
            Set<String> uniqueKeySet = new HashSet<>();
            Workbook outputWorkbook = new XSSFWorkbook();
            Sheet outputSheet = outputWorkbook.createSheet();

            for (Row row : sheet) {
                Cell keyCell = row.getCell(5);
                String keyValue = keyCell.toString();
                if (!uniqueKeySet.contains(keyValue) && !keySet.contains(keyValue)) {
                    uniqueKeySet.add(keyValue);
                }
                if (keySet.contains(keyValue)) {
                    uniqueKeySet.remove(keyValue);
                }
                keySet.add(keyValue);
            }

            int outputRowNum = 0;
            for (Row row : sheet) {
                Cell keyCell = row.getCell(5);
                String keyValue = keyCell.toString();
                if (row.getRowNum() == 0 || uniqueKeySet.contains(keyValue)) {
                    Row outputRow = outputSheet.createRow(outputRowNum++);
                    copyRow(row, outputRow);
                }
            }

            outputWorkbook.write(fos);
        } catch (IOException e) {
            System.out.println("failed to handler unique key file: " + e);
            return "整理单次会诊的病人记录失败";
        }

        return "理单次会诊的病人记录成功：" + fileName;
    }

    private void copyRow(Row sourceRow, Row destinationRow) {
        for (int cn = 0; cn < sourceRow.getLastCellNum(); cn++) {
            Cell sourceCell = sourceRow.getCell(cn);
            Cell newCell = destinationRow.createCell(cn);
            if (sourceCell != null) {
                newCell.setCellValue(getCellAsString(sourceCell, cn, sourceRow.getRowNum()));
            }
        }
    }

    private String getCellAsString(Cell cell, int index, int rowNum) {
        DataFormatter formatter = new DataFormatter();
        boolean isDate = dateTimeIndexes.contains(index);
        if (isDate && rowNum != 0) {
            // 为日期类型的单元格设置特定的格式
            if (cell.getCellType() == CellType.STRING) {
                cell.setCellValue(cell.getStringCellValue().replaceAll("-", "/"));
                return cell.getStringCellValue();
            }
            if (cell.getDateCellValue() == null) {
                return "";
            }
            SimpleDateFormat sdf = new SimpleDateFormat("yyyy/MM/dd");
            return sdf.format(cell.getDateCellValue());
        } else {
            // 对于非日期类型的单元格，使用默认的格式化
            return formatter.formatCellValue(cell);
        }
    }
}
