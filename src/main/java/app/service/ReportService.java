package app.service;

import app.model.Product;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.StreamSupport;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;

@Service
public class ReportService {

    /**
     * Creates an excel file based on product list.
     * @param products - The list of products.
     * @return The excel file.
     */
    public Workbook write(List<Product> products) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Products");
        Row header = sheet.createRow(sheet.getPhysicalNumberOfRows());
        List<String> columns = List.of("Id", "Name", "Quantity", "Unit", "Price", "Currency", "Expiration Date");

        columns.forEach(column -> header.createCell(header.getPhysicalNumberOfCells()).setCellValue(column));
        products.forEach(product -> addCellsToRow(sheet.createRow(sheet.getPhysicalNumberOfRows()), product));

        return workbook;
    }

    /**
     * Creates a list of products based on excel file.
     * @param workbook - The excel file.
     * @return The list of products.
     */
    public List<Product> read(Workbook workbook) {
        return StreamSupport.stream(workbook.getSheetAt(0).spliterator(), false)
            .skip(1)
            .map(row -> getProductFromRow(row))
            .collect(Collectors.toList());
    }

    /**
     * Add table cells for the current table row.
     * @param row - The table row.
     * @param product - The product.
     */
    private void addCellsToRow(Row row, Product product) {
        row.createCell(0).setCellValue(product.id);
        row.createCell(1).setCellValue(product.name);
        row.createCell(2).setCellValue(product.quantity);
        row.createCell(3).setCellValue(product.unit);
        row.createCell(4).setCellValue(product.price);
        row.createCell(5).setCellValue(product.currency);
        row.createCell(6).setCellValue(product.expirationDate);
    }

    /**
     * Gests the Product object from table row.
     * @param row - The table row.
     * @return New Product.
     */
    private Product getProductFromRow(Row row) {
        Product product = new Product();

        product.id = (int) row.getCell(0).getNumericCellValue();
        product.name = row.getCell(1).getStringCellValue();
        product.quantity = row.getCell(2).getNumericCellValue();
        product.unit = row.getCell(3).getStringCellValue();
        product.price = row.getCell(4).getNumericCellValue();
        product.currency = row.getCell(5).getStringCellValue();
        product.expirationDate = row.getCell(6).getDateCellValue();

        return product;
    }
}
