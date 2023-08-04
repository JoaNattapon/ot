package Controller;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RequestParam;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@Controller
public class ExcelController {

    @PostMapping("/uploadExcel")
    public String uploadExcelFile(@RequestParam("file") MultipartFile file) {

        int calculatedValue = 10;
        return "redirect:/genWord?calculatedValue=" + calculatedValue;

    }

}
