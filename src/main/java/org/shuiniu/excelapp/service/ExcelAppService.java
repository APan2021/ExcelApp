package org.shuiniu.excelapp.service;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
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
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.temporal.ChronoUnit;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

@Service
public class ExcelAppService {
    private final static String TMP_EXCEL = "/app/tmp.xlsx";

    private String currentFile = "";

    public String saveExcelFile(MultipartFile file) throws IOException {
        Path path = Paths.get(TMP_EXCEL);
        currentFile = TMP_EXCEL;
        // 确保目录存在
        Files.createDirectories(path.getParent());
        // 文件写入指定路径
        file.transferTo(path);

        return "'" + file.getOriginalFilename() + "' 文件上传成功";
    }

    // 统一日期格式
    private static final SimpleDateFormat outputDateFormat = new SimpleDateFormat("yyyy-MM-dd");

    // 日期和时间相关列的索引，注意所有索引加1，因表头第一列是空的
    private static final Set<Integer> dateColumnIndexes = new HashSet<>();

    static {
        dateColumnIndexes.add(7);  // 报告日期
        dateColumnIndexes.add(17); // 液基时间1
        dateColumnIndexes.add(22); // 病理时间1
        dateColumnIndexes.add(19); // 病理时间2
        dateColumnIndexes.add(24); // 液基时间2
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

        String overSixMonthsFile = "over_six_months_" + System.currentTimeMillis() + ".xlsx";
        String withinSixMonthsFile = "within_six_months_" + System.currentTimeMillis() + ".xlsx";

        String overSixMonthsPath = Paths.get(outputDirectory, overSixMonthsFile).toString();
        String withinSixMonthsPath = Paths.get(outputDirectory, withinSixMonthsFile).toString();

        try (FileInputStream fis = new FileInputStream(currentFile);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fosOverSix = new FileOutputStream(overSixMonthsPath);
             FileOutputStream fosWithinSix = new FileOutputStream(withinSixMonthsPath)) {

            Sheet sheet = workbook.getSheetAt(0);

            // 创建两个输出文件的工作簿和表
            Workbook overSixMonthsWorkbook = new XSSFWorkbook();
            Sheet overSixMonthsSheet = overSixMonthsWorkbook.createSheet();
            Workbook withinSixMonthsWorkbook = new XSSFWorkbook();
            Sheet withinSixMonthsSheet = withinSixMonthsWorkbook.createSheet();

            // 复制表头
            copyRow(sheet.getRow(0), overSixMonthsSheet.createRow(0));
            copyRow(sheet.getRow(0), withinSixMonthsSheet.createRow(0));

            int overSixRowNum = 1;
            int withinSixRowNum = 1;

            // 遍历每一行，从第二行开始
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);

                // 获取报告日期（第7列，索引调整为8）和病理时间（第19列，索引调整为20）
                Cell reportDateCell = row.getCell(7);  // 调整索引为8
                Cell pathologyDateCell = row.getCell(19);  // 调整索引为20

                Date reportDate = null;
                Date pathologyDate = null;

                // 尝试解析日期
                if (reportDateCell != null) {
                    reportDate = parseDateCell(reportDateCell);
                }
                if (pathologyDateCell != null) {
                    pathologyDate = parseDateCell(pathologyDateCell);
                }

                // 如果报告日期和病理时间任一为空，则归入不超过半年的文件
                if (reportDate == null || pathologyDate == null) {
                    copyRowWithFormattedDates(row, withinSixMonthsSheet.createRow(withinSixRowNum++));
                    continue;
                }

                // 病理时间必须大于等于报告日期
                if (!pathologyDate.before(reportDate)) {
                    // 计算两个日期之间的差值，以月为单位
                    long monthsDifference = calculateMonthDifference(reportDate, pathologyDate);

                    // 如果超过半年，加入超过六个月的文件
                    if (monthsDifference > 6) {
                        copyRowWithFormattedDates(row, overSixMonthsSheet.createRow(overSixRowNum++));
                    } else {
                        // 不超过半年，加入不超过六个月的文件
                        copyRowWithFormattedDates(row, withinSixMonthsSheet.createRow(withinSixRowNum++));
                    }
                }
            }

            // 将两个工作簿写入文件
            overSixMonthsWorkbook.write(fosOverSix);
            withinSixMonthsWorkbook.write(fosWithinSix);

        } catch (IOException e) {
            System.out.println("处理Excel文件时出错: " + e.getMessage());
            return "处理文件时出错：" + e.getMessage();
        }

        return "处理完成，文件生成：" + overSixMonthsFile + " 和 " + withinSixMonthsFile;
    }

    // 辅助方法：从单元格中解析日期
    private Date parseDateCell(Cell cell) {
        if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
            return cell.getDateCellValue(); // 如果是日期格式的数字，直接返回日期
        }
        if (cell.getCellType() == CellType.STRING) {
            String dateString = cell.getStringCellValue();
            // 尝试不同的日期格式
            String[] dateFormats = {"yyyy-MM-dd", "yyyy/MM/dd"};
            for (String format : dateFormats) {
                try {
                    SimpleDateFormat dateFormat = new SimpleDateFormat(format);
                    return dateFormat.parse(dateString); // 尝试解析日期
                } catch (Exception e) {
                    // 忽略解析异常，继续尝试下一个格式
                }
            }
        }
        return null; // 如果解析失败，返回null
    }

    // 辅助方法：计算两个日期之间的月份差
    private long calculateMonthDifference(Date reportDate, Date pathologyDate) {
        LocalDate reportLocalDate = convertToLocalDate(reportDate);
        LocalDate pathologyLocalDate = convertToLocalDate(pathologyDate);
        return ChronoUnit.MONTHS.between(reportLocalDate, pathologyLocalDate);
    }

    // 辅助方法：将 Date 转换为 LocalDate
    private LocalDate convertToLocalDate(Date date) {
        return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
    }

    // 复制一行，并确保日期列的格式为 "yyyy-MM-dd"
    private void copyRowWithFormattedDates(Row sourceRow, Row destinationRow) {
        for (int cn = 0; cn < sourceRow.getLastCellNum(); cn++) {
            Cell sourceCell = sourceRow.getCell(cn);
            Cell newCell = destinationRow.createCell(cn);
            if (sourceCell != null) {
                if (dateColumnIndexes.contains(cn)) {
                    // 对于日期相关的列，统一格式化为 yyyy-MM-dd
                    newCell.setCellValue(formatDateCell(sourceCell));
                } else {
                    newCell.setCellValue(getCellAsString(sourceCell));
                }
            }
        }
    }

    // 辅助方法：格式化日期单元格为 yyyy-MM-dd
    private String formatDateCell(Cell cell) {
        Date date = parseDateCell(cell);
        if (date != null) {
            return outputDateFormat.format(date);
        }
        return ""; // 如果日期为空，返回空字符串
    }

    // 将单元格转换为字符串
    private String getCellAsString(Cell cell) {
        DataFormatter formatter = new DataFormatter();
        return formatter.formatCellValue(cell);
    }

    // 复制一行
    private void copyRow(Row sourceRow, Row destinationRow) {
        for (int cn = 0; cn < sourceRow.getLastCellNum(); cn++) {
            Cell sourceCell = sourceRow.getCell(cn);
            Cell newCell = destinationRow.createCell(cn);
            if (sourceCell != null) {
                newCell.setCellValue(getCellAsString(sourceCell));
            }
        }
    }

    public String mergeMultiRowData() {
        if (currentFile == null) {
            return "请选择要处理的Excel文件";
        }

        String outputDirectory = "/app/files";
        File directory = new File(outputDirectory);
        if (!directory.exists()) {
            directory.mkdirs();
        }

        String mergedFile = "merged_data_" + System.currentTimeMillis() + ".xlsx";
        String overSixMonthsFile = "over_six_months_group_" + System.currentTimeMillis() + ".xlsx";

        String mergedFilePath = Paths.get(outputDirectory, mergedFile).toString();
        String overSixMonthsFilePath = Paths.get(outputDirectory, overSixMonthsFile).toString();

        try (FileInputStream fis = new FileInputStream(currentFile);
             Workbook workbook = new XSSFWorkbook(fis);
             FileOutputStream fosMerged = new FileOutputStream(mergedFilePath);
             FileOutputStream fosOverSixMonths = new FileOutputStream(overSixMonthsFilePath)) {

            Sheet sheet = workbook.getSheetAt(0);

            // 创建两个输出文件的工作簿和表
            Workbook mergedWorkbook = new XSSFWorkbook();
            Sheet mergedSheet = mergedWorkbook.createSheet();
            Workbook overSixMonthsWorkbook = new XSSFWorkbook();
            Sheet overSixMonthsSheet = overSixMonthsWorkbook.createSheet();

            // 复制表头
            copyRow(sheet.getRow(0), mergedSheet.createRow(0));
            copyRow(sheet.getRow(0), overSixMonthsSheet.createRow(0));

            int mergedRowNum = 1;
            int overSixMonthsRowNum = 1;

            Map<String, List<Row>> groupedData = new HashMap<>();
            DataFormatter formatter = new DataFormatter();

            // 遍历每一行，将具有相同关键列的数据分组
            for (int i = 1; i <= sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                Cell keyCell = row.getCell(5);  // 病理号是关键列
                String key = formatter.formatCellValue(keyCell);

                groupedData.computeIfAbsent(key, k -> new ArrayList<>()).add(row);
            }

            // 处理每一组数据
            for (Map.Entry<String, List<Row>> entry : groupedData.entrySet()) {
                List<Row> rows = entry.getValue();
                if (rows.size() == 1) {
                    // 只有一条数据，直接复制
                    continue;
                }

                Row firstRow = rows.get(0);
                Row mergedRow = mergedSheet.createRow(mergedRowNum++);
                copyRowWithFormattedDates(firstRow, mergedRow);

                // 开始合并数据
                int colOffset = firstRow.getLastCellNum();
                for (int j = 1; j < rows.size(); j++) {
                    Row currentRow = rows.get(j);
                    addMergedRow(currentRow, mergedRow, colOffset, j);  // 合并数据，并设置新表头
                    colOffset += 9; // 每次增加9列
                }

                // 判断日期是否超过半年
                boolean allOverSixMonths = true;
                for (Row row : rows) {
                    Date reportDate = parseDateCell(row.getCell(7));  // 报告日期
                    Date pathologyDate = parseDateCell(row.getCell(19));  // 病理时间
                    if (reportDate == null || pathologyDate == null || calculateMonthDifference(reportDate, pathologyDate) <= 6) {
                        allOverSixMonths = false;
                        break;
                    }
                }

                // 如果全部超过半年，放入overSixMonths文件中
                if (allOverSixMonths) {
                    copyRowWithFormattedDates(firstRow, overSixMonthsSheet.createRow(overSixMonthsRowNum++));
                }
            }

            // 将两个工作簿写入文件
            mergedWorkbook.write(fosMerged);
            overSixMonthsWorkbook.write(fosOverSixMonths);

        } catch (IOException e) {
            System.out.println("处理Excel文件时出错: " + e.getMessage());
            return "处理文件时出错：" + e.getMessage();
        }

        return "处理完成，文件生成：" + mergedFile + " 和 " + overSixMonthsFile;
    }

    // 辅助方法：添加合并的数据列，生成新的表头
    private void addMergedRow(Row sourceRow, Row destinationRow, int colOffset, int index) {
        String[] headers = {"报告日期", "项目名称", "标本名称", "诊断结果", "肉眼所见", "液基时间", "液基结果", "病理时间", "病理结果", "病理分类"};
        for (int i = 0; i < headers.length; i++) {
            Cell sourceCell = sourceRow.getCell(i + 7); // 合并的列从报告日期开始
            Cell newCell = destinationRow.createCell(colOffset + i);
            newCell.setCellValue(getCellAsString(sourceCell));

            // 生成新的表头
            Cell headerCell = destinationRow.getSheet().getRow(0).createCell(colOffset + i);
            headerCell.setCellValue(headers[i] + index);
            // 根据index设置不同的颜色
            CellStyle style = destinationRow.getSheet().getWorkbook().createCellStyle();
            style.setFillForegroundColor(index % 2 == 0 ? IndexedColors.YELLOW.getIndex() : IndexedColors.LIGHT_BLUE.getIndex());
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            headerCell.setCellStyle(style);
        }
    }

}
