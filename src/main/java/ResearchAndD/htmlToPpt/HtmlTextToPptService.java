package ResearchAndD.htmlToPpt;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.GradientPaint;
import java.awt.Graphics2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.xslf.usermodel.*;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import javax.imageio.ImageIO;
import java.awt.*;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import javax.imageio.ImageIO;

import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFGroupShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextBox;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.springframework.stereotype.Service;
@Service
public class HtmlTextToPptService {
	public void convertHtmlToPpt(String filePath) throws FileNotFoundException, IOException {
	    XMLSlideShow ppt = new XMLSlideShow();

	    // Create a slide
	    XSLFSlide slide = ppt.createSlide();
	    BufferedImage gradientImage = createVerticalGradientImage("#A434D9", "#00D4D1", "#E94EB3", 0.17, slide.getSlideShow().getPageSize());

        // Insert the image into the slide
        XSLFPictureData pictureData = ppt.addPicture(toByteArray(gradientImage), PictureData.PictureType.PNG);
        XSLFPictureShape picture = slide.createPicture(pictureData);

        // Set the image position and size to cover the entire slide
        picture.setAnchor(new Rectangle2D.Double(0, 0, slide.getSlideShow().getPageSize().getWidth(), slide.getSlideShow().getPageSize().getHeight()));
	    
	    // Create a group shape (equivalent to a rectangle)
	    XSLFGroupShape groupShape = slide.createGroup();

	    // Set the anchor for the group shape
	    groupShape.setAnchor(new java.awt.Rectangle(50, 50, 700, 400));  // Adjust the size and position as needed

	    // Create a rectangle within the group shape
	    XSLFAutoShape rectangle = groupShape.createAutoShape();
	    rectangle.setShapeType(ShapeType.RECT);
	    rectangle.setAnchor(new java.awt.Rectangle(0, 0, 500, 300));

	    // Create a text box within the group shape for the title
	    XSLFTextBox titleTextBox = groupShape.createTextBox();

	    // Set the anchor for the title text box to cover the full width of the slide
	    int slideWidth = (int) slide.getSlideShow().getPageSize().getWidth();
	    int titleWidth = slideWidth; // Set the width of the title box to cover the full width
	    int titleHeight = 50;
	    int titleX = 0;
	    int titleY = 0;

	    titleTextBox.setAnchor(new java.awt.Rectangle(titleX, titleY, titleWidth, titleHeight));

	    // Add title to the text box
	    XSLFTextParagraph titleParagraph = titleTextBox.addNewTextParagraph();
	    titleParagraph.setTextAlign(TextAlign.LEFT); // Center align the title text
	    XSLFTextRun titleRun = titleParagraph.addNewTextRun();
	    titleRun.setText("Your Box Title");
	    titleRun.setFontSize(20.0);

	    // Create cards within the group shape
	    createCard(groupShape, "Card 1 Heading", new String[]{"Bullet 1", "Bullet 2", "Bullet 3"});
	    createCard(groupShape, "Card 2 Heading", new String[]{"Bullet A", "Bullet B", "Bullet C"});
	    createCard(groupShape, "Card 3 Heading", new String[]{"Bullet X", "Bullet Y", "Bullet Z"});
	    createCard(groupShape, "Card 1 Heading", new String[]{"Bullet 1", "Bullet 2", "Bullet 3"});
	    createCard(groupShape, "Card 2 Heading", new String[]{"Bullet A", "Bullet B", "Bullet C"});
	    createCard(groupShape, "Card 3 Heading", new String[]{"Bullet X", "Bullet Y", "Bullet Z"});

	    // Save the presentation to a file
	    try (FileOutputStream out = new FileOutputStream(filePath)) {
	        ppt.write(out);
	    }
	    System.out.println("Presentation created successfully!");
	}

	private static void createCard(XSLFGroupShape groupShape, String heading, String[] bulletPoints) {
	    // Create the card (rectangle) within the group shape
	    XSLFAutoShape card = groupShape.createAutoShape();
	    card.setShapeType(ShapeType.RECT);
	    card.setLineColor(Color.BLACK);
	    card.setLineWidth(2.0);
	    
	    // Adjust the position and size of each card
	    int cardWidth = 200;
	    int cardHeight = 100;
	    int cardSpacingX = 20;
	    int cardSpacingY = 20;

	    int cardIndex = groupShape.getShapes().size(); // Index of the last created card

	    int rowIndex = cardIndex / 3;  // 3 cards per row
	    int columnIndex = cardIndex % 3;

	    int x = columnIndex * (cardWidth + cardSpacingX);
	    int y = rowIndex * (cardHeight + cardSpacingY);

	    card.setAnchor(new java.awt.Rectangle(x, y, cardWidth, cardHeight));

	    // Set border properties for the entire card
	    card.setLineColor(Color.BLACK);
	    card.setLineWidth(2.0);

	    // Add content to the card
	    XSLFTextParagraph paragraph = card.addNewTextParagraph();
	    paragraph.setTextAlign(TextAlign.LEFT);

	    // Add heading to the text box
	    XSLFTextRun headingRun = paragraph.addNewTextRun();
	    headingRun.setText(heading);
	    headingRun.setBold(true);

	    // Add bullet points
	    for (String bulletPoint : bulletPoints) {
	        XSLFTextParagraph bulletParagraph = card.addNewTextParagraph();
	        bulletParagraph.setTextAlign(TextAlign.LEFT);

	        XSLFTextRun bulletRun = bulletParagraph.addNewTextRun();
	        bulletRun.setText("• " + bulletPoint);
	    }

	}
	private static BufferedImage createVerticalGradientImage(String startColor, String middleColor, String endColor, double opacity, Dimension size) {
        BufferedImage image = new BufferedImage(size.width, size.height, BufferedImage.TYPE_INT_ARGB);
        Graphics2D g2d = image.createGraphics();

        // Parse color codes
        Color color1 = Color.decode(startColor);
        Color color2 = Color.decode(middleColor);
        Color color3 = Color.decode(endColor);

        // Create a gradient paint
        GradientPaint gradientPaint = new GradientPaint(0, 0, color1, 0, size.height, color3);

        // Set transparency for the middle color
        Color middleColorWithOpacity = new Color(color2.getRed(), color2.getGreen(), color2.getBlue(), (int) (opacity * 255));

        // Create a gradient with three colors
        MultipleGradientPaint.ColorSpaceType colorSpace = MultipleGradientPaint.ColorSpaceType.LINEAR_RGB;
        int[] fractions = {0, 50, 100};
        Color[] colors = {color1, middleColorWithOpacity, color3};

        // Fill the entire image with the gradient paint
        g2d.setPaint(gradientPaint);
        g2d.fillRect(0, 0, size.width, size.height);

        // Dispose of the graphics context
        g2d.dispose();

        return image;
    }

    private static byte[] toByteArray(BufferedImage image) {
        try (ByteArrayOutputStream baos = new ByteArrayOutputStream()) {
            ImageIO.write(image, "png", baos);
            return baos.toByteArray();
        } catch (IOException e) {
            e.printStackTrace();
            return null;
        }
    }


	}