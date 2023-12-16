package ResearchAndD.htmlToPpt;
import java.awt.Color;
import java.awt.Dimension;
import java.awt.GradientPaint;
import java.awt.Graphics2D;
import java.awt.MultipleGradientPaint;
import java.awt.geom.Path2D;
import java.awt.geom.Rectangle2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;

import javax.imageio.ImageIO;

import java.awt.Color;
import java.awt.geom.Path2D;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.ByteArrayOutputStream;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.sl.usermodel.PictureData;
import org.apache.poi.sl.usermodel.ShapeType;
import org.apache.poi.sl.usermodel.TextParagraph.TextAlign;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.springframework.stereotype.Service;

@Service
public class SecondSlideService {
	 public static byte[] convertHtmlToPpt(String filePath) throws IOException {
		 try (XMLSlideShow ppt = new XMLSlideShow()) {
	            // Create a slide
	            XSLFSlide slide = ppt.createSlide();

	            //BACKGROUND COLOR
	    	    BufferedImage gradientImage = createVerticalGradientImage("#A434D9", "#00D4D1", "#E94EB3", 0.17, slide.getSlideShow().getPageSize());

	            // Insert the image into the slide
	            XSLFPictureData pictureData = ppt.addPicture(toByteArray(gradientImage), PictureData.PictureType.PNG);
	            XSLFPictureShape picture = slide.createPicture(pictureData);

	            // Set the image position and size to cover the entire slide
	            picture.setAnchor(new Rectangle2D.Double(0, 0, slide.getSlideShow().getPageSize().getWidth(), slide.getSlideShow().getPageSize().getHeight()));
	    	    
	            // Create the first box
	            XSLFAutoShape card1 = createCard(slide, 50, 175, 87, 150, Color.decode("#F1F1F1"));
	            XSLFAutoShape card2 = createCard(slide, 135, 160, 191, 165, Color.decode("#EDF0F4"));
	            XSLFAutoShape card3 = createCard(slide, 325, 144, 190, 180, Color.decode("#D3DBE4"));
	            XSLFAutoShape card4 = createCard(slide, 515, 128, 156, 195, Color.decode("#ABBBCC"));
	            XSLFAutoShape card5 = createCard(slide, 670, 115, 30, 210, Color.decode("#88A1B8"));
//	            
	            
	            addText(card2, "Card 2 Title", "Bullet Point 1", "Bullet Point 2", "Bullet Point 3");
	            addText(card3, "Card 3 Title", "Bullet Point 1", "Bullet Point 2", "Bullet Point 3", "Bullet Point 4");
	            addText(card4, "Card 4 Title", "Bullet Point 1", "Bullet Point 2", "Bullet Point 3", "Bullet Point 4", "Bullet Point 5");

	             
	            XSLFAutoShape box1 = createBox(slide, 50, 160, 85, 15, Color.decode("#3D556A"));

	            // Create the second box rotated -45 degrees
	            XSLFAutoShape box2 = createBox(slide, 130, 153, 50, 15, Color.decode("#3D556A"));
	            box2.setRotation(-20);

	           
	            // Create the third box connected to th e second box
	            XSLFAutoShape box3 = createBox(slide, 175, 145, 150, 15, Color.decode("#3D556A"));

	            XSLFAutoShape box4 = createBox(slide, 320, 137, 50, 15, Color.decode("#3D556A"));
	            box4.setRotation(-20);
	            
	            XSLFAutoShape box5 = createBox(slide, 365, 129, 150, 15, Color.decode("#3D556A"));
	            
	            XSLFAutoShape box6 = createBox(slide, 510, 120, 50, 15, Color.decode("#3D556A"));
	            box6.setRotation(-20);
	            
	            XSLFAutoShape box7 = createBox(slide, 555, 113, 115, 15, Color.decode("#3D556A"));
	            	           
	            XSLFAutoShape arrow = createArrow(slide, 664, 93, 55, 32, Color.decode("#3D556A"));
	            
	            arrow.setRotation(-25);
	            
	            // Save the presentation to a file
	            try (FileOutputStream out = new FileOutputStream(filePath)) {
	                ppt.write(out);
	                
	                System.out.println("Presentation created successfully!");
	            } catch (IOException e) {
	                e.printStackTrace();
	            }
	            try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
	                ppt.write(out);
	                return out.toByteArray();
	            } catch (IOException e) {
	                e.printStackTrace();
	            }
	        } catch (Exception e) {
	            e.printStackTrace();
	        }
		 return new byte[0];
	    }

	 	public static void addText(XSLFAutoShape card, String heading, String... bulletPoints) {
	        XSLFTextParagraph paragraph = card.addNewTextParagraph();
	        paragraph.setTextAlign(TextAlign.CENTER);

	        
	        
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
	    private static XSLFAutoShape createBox(XSLFSlide slide, int x, int y, int width, int height, Color fillColor) {
	        XSLFAutoShape box = slide.createAutoShape();
	        box.setAnchor(new java.awt.Rectangle(x, y, width, height));
	        box.setFillColor(fillColor);

	        return box;
	    }
	    public static XSLFAutoShape createCard(XSLFSlide slide, int x, int y, int width, int height, Color color) {
	        XSLFAutoShape shape = slide.createAutoShape();
	        shape.setShapeType(ShapeType.RECT);
	        shape.setAnchor(new Rectangle2D.Double(x, y, width, height));
	        shape.setFillColor(color);
	        return shape;
	    }

	    public static XSLFAutoShape createArrow(XSLFSlide slide, int x, int y, int width, int height, Color color) {
	        XSLFAutoShape shape = slide.createAutoShape();

	        // Set the arrow shape properties
	        shape.setShapeType(ShapeType.RIGHT_ARROW);

	        // Set the anchor (position and size)
	        shape.setAnchor(new Rectangle2D.Double(x, y, width, height));

	        // Set the fill color
	        shape.setFillColor(color);

	        // Customize other properties as needed

	        return shape;
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