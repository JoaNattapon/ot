package Controller;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestParam;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.OutputStream;

@Controller
public class WordController {

    @GetMapping("/outputWord")
    public void generateWordDocument(@RequestParam("calculatedValue") String calculatedValue,
                                     HttpServletResponse response) throws IOException {

        // Create a Word document
        XWPFDocument document = new XWPFDocument();

        // Add content to the document
        XWPFParagraph paragraph = document.createParagraph();
        XWPFRun run = paragraph.createRun();
        run.setText("Hello Test. Document !!!" + calculatedValue);

        // Set the response headers for downloading the Word document.
        response.setContentType("application/vnd.openxmlformats-officedocument.wordprocessingl.dcoument");
        response.setHeader("Content-Disposition", "attachment; filename=calculated_output.docx");

        // Write the document content to the response stream
        OutputStream outputStream = response.getOutputStream();
        document.write(outputStream);
        outputStream.close();
    }


}
