package org.shuiniu.excelapp.controller;

import org.shuiniu.excelapp.service.ExcelAppService;
import org.springframework.http.HttpStatus;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.annotation.Resource;

@RestController
public class ExcelAppController {
    @Resource
    private ExcelAppService excelAppService;

    @PostMapping("/upload")
    public ResponseEntity<String> uploadFile(@RequestParam("file") MultipartFile file) {
        if (file.isEmpty()) {
            return ResponseEntity.status(HttpStatus.BAD_REQUEST).body("No file uploaded.");
        }

        try {
            // 你可以在这里调用一个服务来处理文件
            String message = excelAppService.saveExcelFile(file);
            return ResponseEntity.ok(message);
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Error processing file: " + e.getMessage());
        }
    }

    @GetMapping("/unique")
    public ResponseEntity<String> handleUniqueKeys() {
        try {
            // 你可以在这里调用一个服务来处理文件
            String message = excelAppService.getUniqueKeyFile();
            return ResponseEntity.ok(message);
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Error processing file: " + e.getMessage());
        }
    }

    @GetMapping("/merge")
    public ResponseEntity<String> handleMergeData() {
        try {
            // 你可以在这里调用一个服务来处理文件
            String message = excelAppService.mergeMultiRowData();
            return ResponseEntity.ok(message);
        } catch (Exception e) {
            return ResponseEntity.status(HttpStatus.INTERNAL_SERVER_ERROR).body("Error processing file: " + e.getMessage());
        }

    }
}