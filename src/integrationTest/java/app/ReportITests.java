package app;

import static app.controller.ReportController.mapToOutputStream;
import static org.assertj.core.api.Assertions.assertThat;
import static org.springframework.http.MediaType.APPLICATION_JSON;
import static org.springframework.http.MediaType.MULTIPART_FORM_DATA_VALUE;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.multipart;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.post;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.context.ActiveProfiles;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.MvcResult;

@SpringBootTest
@AutoConfigureMockMvc
@ActiveProfiles("test")
@DisplayName("Report Resource")
public class ReportITests {

    @Autowired
    transient MockMvc mockMvc;

    @Test
    @DisplayName("/write (POST)")
    void writeReport() throws Exception {
        String products =
            "[{\"id\":1,\"name\":\"Product Name\",\"quantity\":1.0,\"unit\":\"L\"," +
            "\"price\":5.0,\"currency\":\"€\",\"expirationDate\":\"2021-12-12\"}]";

        MvcResult response = mockMvc
            .perform(post("/write").content(products).contentType(APPLICATION_JSON))
            .andReturn();

        assertThat(response.getResponse().getStatus()).isEqualTo(200);
    }

    @Test
    @DisplayName("/read (POST)")
    void readReport() throws Exception {
        Workbook workbook = new XSSFWorkbook().createSheet("Products").getWorkbook();
        byte[] byteArray = mapToOutputStream(workbook).toByteArray();
        MockMultipartFile file = new MockMultipartFile("file", "Report.xlsx", MULTIPART_FORM_DATA_VALUE, byteArray);

        MvcResult response = mockMvc
            .perform(multipart("/read").file(file))
            .andReturn();

        assertThat(response.getResponse().getStatus()).isEqualTo(200);
    }
}
