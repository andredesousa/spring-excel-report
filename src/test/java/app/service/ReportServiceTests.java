package app.service;

import static org.assertj.core.api.Assertions.assertThat;

import app.model.Product;
import java.time.Instant;
import java.util.Date;
import java.util.List;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.extension.ExtendWith;
import org.junit.jupiter.api.Test;
import org.mockito.junit.jupiter.MockitoExtension;

@DisplayName("ReportService")
@ExtendWith(MockitoExtension.class)
public class ReportServiceTests {

    transient ReportService reportService;

    transient Product product;

    @BeforeEach
    void beforeEach() {
        reportService = new ReportService();

        product = new Product();
        product.id = 1;
        product.name = "Product Name";
        product.quantity = 1.0;
        product.unit = "L";
        product.price = 5.0;
        product.currency = "€";
        product.expirationDate = Date.from(Instant.parse("2021-12-12T00:00:00.000Z"));
    }

    @Test
    @DisplayName("#write returns a workbook")
    void write() {
        List<Product> products = List.of(product);

        Workbook workbook = reportService.write(products);

        assertThat(workbook.getSheetAt(0).getRow(0).getCell(0).getStringCellValue()).isEqualTo("Id");
        assertThat(workbook.getSheetAt(0).getRow(0).getCell(1).getStringCellValue()).isEqualTo("Name");
        assertThat(workbook.getSheetAt(0).getRow(0).getCell(2).getStringCellValue()).isEqualTo("Quantity");
        assertThat(workbook.getSheetAt(0).getRow(0).getCell(3).getStringCellValue()).isEqualTo("Unit");
        assertThat(workbook.getSheetAt(0).getRow(0).getCell(4).getStringCellValue()).isEqualTo("Price");
        assertThat(workbook.getSheetAt(0).getRow(0).getCell(5).getStringCellValue()).isEqualTo("Currency");
        assertThat(workbook.getSheetAt(0).getRow(0).getCell(6).getStringCellValue()).isEqualTo("Expiration Date");

        assertThat(workbook.getSheetAt(0).getRow(1).getCell(0).getNumericCellValue()).isEqualTo(product.id);
        assertThat(workbook.getSheetAt(0).getRow(1).getCell(1).getStringCellValue()).isEqualTo(product.name);
        assertThat(workbook.getSheetAt(0).getRow(1).getCell(2).getNumericCellValue()).isEqualTo(product.quantity);
        assertThat(workbook.getSheetAt(0).getRow(1).getCell(3).getStringCellValue()).isEqualTo(product.unit);
        assertThat(workbook.getSheetAt(0).getRow(1).getCell(4).getNumericCellValue()).isEqualTo(product.price);
        assertThat(workbook.getSheetAt(0).getRow(1).getCell(5).getStringCellValue()).isEqualTo(product.currency);
        assertThat(workbook.getSheetAt(0).getRow(1).getCell(6).getDateCellValue()).isEqualTo(product.expirationDate);
    }

    @Test
    @DisplayName("#read returns a list of products")
    void read() {
        Workbook workbook = reportService.write(List.of(product));

        List<Product> products = reportService.read(workbook);

        assertThat(products).usingRecursiveComparison().isEqualTo(List.of(product));
    }
}
