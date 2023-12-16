package ResearchAndD.htmlToPpt;

import java.io.FileNotFoundException;
import java.io.IOException;

import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.CommandLineRunner;
import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class HtmlToPptApplication implements CommandLineRunner {

	 @Autowired
	    private HtmlTextToPptService htmlTextToPptService;
	 @Autowired
	 private SecondSlideService secondSlideService;
	public static void main(String[] args) {
		SpringApplication.run(HtmlToPptApplication.class, args);
	}
	@Override
    public void run(String... args) throws FileNotFoundException, IOException {
        // Specify the output file path
//        String outputPath = "output.pptx";
//
//        // Call the service to perform the conversion
//        htmlTextToPptService.convertHtmlToPpt(outputPath);
		
		String secondSlidPath = "second.pptx";
//        secondSlideService.convertHtmlToPpt(secondSlidPath);
    }
}
