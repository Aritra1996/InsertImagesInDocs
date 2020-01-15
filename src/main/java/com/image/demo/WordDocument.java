package com.image.demo;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.BreakType;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

public class WordDocument {
	
	XWPFDocument doc;
	XWPFParagraph p;
	XWPFRun r;
	
	public WordDocument() {
		doc = new XWPFDocument();
		p = doc.createParagraph();
		r = p.createRun();
	}
	
	public void closeWordDoument() throws IOException {
		FileOutputStream out = new FileOutputStream("word_images.docx");
		doc.write(out);
		out.close();
		doc.close();
	}
	
	public void addImagesToWordDocument(File imageFile1)
			throws IOException, InvalidFormatException {
	
		int width = 500;
		int height = 575;
		String imgFile1 = imageFile1.getName();
		int imgFormat1 = getImageFormat(imgFile1);
		
		String p1 = "Sample Paragraph Post. This is a sample Paragraph post. Sample Paragraph text is being cut and pasted again and again. This is a sample Paragraph post. peru-duellmans-poison-dart-frog.";
		r.setText(p1);
		r.addBreak();
		r.addBreak();
		//r.addBreak();
		r.addPicture(new FileInputStream(imageFile1), imgFormat1, imgFile1, Units.toEMU(width), Units.toEMU(height));
		// page break
		r.addBreak(BreakType.PAGE);
		// line break
	}
	public int getImageFormat(String imgFileName) {
		int format;
		if (imgFileName.endsWith(".emf"))
			format = XWPFDocument.PICTURE_TYPE_EMF;
		else if (imgFileName.endsWith(".wmf"))
			format = XWPFDocument.PICTURE_TYPE_WMF;
		else if (imgFileName.endsWith(".pict"))
			format = XWPFDocument.PICTURE_TYPE_PICT;
		else if (imgFileName.endsWith(".jpeg") || imgFileName.endsWith(".jpg"))
			format = XWPFDocument.PICTURE_TYPE_JPEG;
		else if (imgFileName.endsWith(".png"))
			format = XWPFDocument.PICTURE_TYPE_PNG;
		else if (imgFileName.endsWith(".dib"))
			format = XWPFDocument.PICTURE_TYPE_DIB;
		else if (imgFileName.endsWith(".gif"))
			format = XWPFDocument.PICTURE_TYPE_GIF;
		else if (imgFileName.endsWith(".tiff"))
			format = XWPFDocument.PICTURE_TYPE_TIFF;
		else if (imgFileName.endsWith(".eps"))
			format = XWPFDocument.PICTURE_TYPE_EPS;
		else if (imgFileName.endsWith(".bmp"))
			format = XWPFDocument.PICTURE_TYPE_BMP;
		else if (imgFileName.endsWith(".wpg"))
			format = XWPFDocument.PICTURE_TYPE_WPG;
		else {
			return 0;
		}
		return format;
	}
}
