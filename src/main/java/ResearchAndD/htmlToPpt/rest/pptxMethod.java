package ResearchAndD.htmlToPpt.rest;

import java.io.IOException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.http.HttpHeaders;
import org.springframework.http.HttpStatus;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.web.bind.annotation.CrossOrigin;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RestController;

import ResearchAndD.htmlToPpt.SecondSlideService;

@RestController
public class pptxMethod {
	@Autowired
    private SecondSlideService secondSlideService;
    String secondSlidPath = "second.pptx";

    @GetMapping("/DownloadPPT")
    public byte[] getPPT() {
        try {
            byte[] pptContent = secondSlideService.convertHtmlToPpt(secondSlidPath);
            HttpHeaders headers = new HttpHeaders();
            headers.setContentType(MediaType.APPLICATION_OCTET_STREAM);
            headers.setContentDispositionFormData("attachment", "presentation.pptx");

            return pptContent;
        } catch (IOException e) {
            e.printStackTrace();
            return new byte[0];
        }
    }
}