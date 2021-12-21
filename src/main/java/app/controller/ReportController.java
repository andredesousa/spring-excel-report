package app.controller;

import app.model.Product;
import app.service.ReportService;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestBody;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

@RestController
public class ReportController {

    @Autowired
    private transient ReportService reportService;

    @PostMapping(value = "/write")
    public ResponseEntity<byte[]> write(@RequestBody List<Product> products) throws IOException {
        return ResponseEntity.ok()
            .contentType(MediaType.APPLICATION_OCTET_STREAM)
            .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=Report.xlsx")
            .body(mapToOutputStream(reportService.write(products)).toByteArray());
    }

    @PostMapping(value = "/read", consumes = MediaType.MULTIPART_FORM_DATA_VALUE)
    public ResponseEntity<List<Product>> read(@RequestParam("file") MultipartFile file) throws IOException {
        return ResponseEntity.ok()
            .contentType(MediaType.APPLICATION_JSON)
            .body(reportService.read(new XSSFWorkbook(file.getInputStream())));
    }

    public static ByteArrayOutputStream mapToOutputStream(Workbook workbook) throws IOException {
        ByteArrayOutputStream os = new ByteArrayOutputStream();
        workbook.write(os);
        workbook.close();

        return os;
    }
}
