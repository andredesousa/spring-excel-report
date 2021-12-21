package app.controller;

import static app.controller.ReportController.mapToOutputStream;
import static org.assertj.core.api.Assertions.assertThat;
import static org.mockito.Mockito.any;
import static org.mockito.Mockito.when;
import static org.springframework.http.HttpHeaders.CONTENT_DISPOSITION;
import static org.springframework.http.MediaType.APPLICATION_JSON;
import static org.springframework.http.MediaType.APPLICATION_OCTET_STREAM;
import static org.springframework.http.MediaType.MULTIPART_FORM_DATA_VALUE;

import app.model.Product;
import app.service.ReportService;
import java.io.IOException;
import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.api.Test;
import org.mockito.InjectMocks;
import org.mockito.junit.jupiter.MockitoExtension;
import org.mockito.Mock;
import org.springframework.http.ResponseEntity;
import org.springframework.mock.web.MockMultipartFile;

@DisplayName("ReportController")
@ExtendWith(MockitoExtension.class)
public class ReportControllerTests {

    @Mock
    transient ReportService reportService;;

    @InjectMocks
    transient ReportController reportController;

    @Test
    @DisplayName("#write returns an excel file")
    void write() throws IOException {
        when(reportService.write(List.of())).thenReturn(new XSSFWorkbook());

        ResponseEntity<byte[]> actual = reportController.write(List.of());
        ResponseEntity<byte[]> expected = ResponseEntity.ok()
            .contentType(APPLICATION_OCTET_STREAM)
            .header(CONTENT_DISPOSITION, "attachment; filename=Report.xlsx")
            .body(actual.getBody());

        assertThat(actual).usingRecursiveComparison().isEqualTo(expected);
    }

    @Test
    @DisplayName("#read returns a list of products")
    void read() throws IOException {
        Workbook workbook = new XSSFWorkbook().createSheet("Products").getWorkbook();
        byte[] byteArray = mapToOutputStream(workbook).toByteArray();
        MockMultipartFile file = new MockMultipartFile("file", "Report.xlsx", MULTIPART_FORM_DATA_VALUE, byteArray);

        when(reportService.read(any(XSSFWorkbook.class))).thenReturn(List.of(new Product()));

        ResponseEntity<List<Product>> actual = reportController.read(file);
        ResponseEntity<List<Product>> expected = ResponseEntity.ok()
            .contentType(APPLICATION_JSON)
            .body(List.of(new Product()));

        assertThat(actual).usingRecursiveComparison().isEqualTo(expected);
    }
}
