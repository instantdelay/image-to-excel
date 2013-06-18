package com.instantdelay;

import java.awt.image.BufferedImage;
import java.awt.image.DataBufferByte;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;

import javax.imageio.ImageIO;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.extensions.XSSFCellFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTPatternFill;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.CTXf;
import org.openxmlformats.schemas.spreadsheetml.x2006.main.STPatternType;

/**
 * Dumb utility to read a grid of pixels from an image file and output them as cells in an Excel document.
 * 
 * @author Spencer Van Hoose
 */
public class ImageToExcel {
	
	// Row height and column width are in different units...
	// TODO Calculate exact square and make configurable
	private static final int PIXEL_HEIGHT = 256;
	private static final int PIXEL_WIDTH = 515;

	private static class ImageData {
		private byte[] pixels;
		private int width;
		private int height;
		private boolean hasAlphaChannel;
		private final int pixelLength;

		public ImageData(BufferedImage image) {
			pixels = ((DataBufferByte) image.getRaster().getDataBuffer()).getData();
			width = image.getWidth();
			height = image.getHeight();
			hasAlphaChannel = image.getAlphaRaster() != null;
			pixelLength = hasAlphaChannel ? 4 : 3;
		}
		
		public byte[] getPixel(int x, int y) {
			int index = (width * y + x) * pixelLength;
			
			if (hasAlphaChannel) {
				return new byte[] { pixels[index + 1], pixels[index + 2], pixels[index + 3] };
			}
			else {
				return new byte[] { pixels[index], pixels[index + 1], pixels[index + 2] };
			}
		}
	}
	
	/**
	 * Usage: imagepath offsetX offsetY step
	 * 
	 * @param args
	 * @throws IOException 
	 */
	public static void main(String[] args) throws IOException {
		BufferedImage image = ImageIO.read(new File(args[0]));
		int offsetX = Integer.parseInt(args[1]);
		int offsetY = Integer.parseInt(args[2]);
		int step = Integer.parseInt(args[3]);
		
		new ImageToExcel().createWorkbook(image, offsetX, offsetY, step, args[0] + ".xlsx");
	}

	private void createWorkbook(BufferedImage image, int offsetX, int offsetY, int step, String outputFile) throws IOException {
		XSSFWorkbook wb = new XSSFWorkbook();
		XSSFSheet sheet = wb.createSheet();
		ImageData imageData = new ImageData(image);
		
		Map<Integer, XSSFCellStyle> styleCache = new HashMap<Integer, XSSFCellStyle>();
		
		int sheetX = 0;
		int sheetY = 0;
		
		for (int y = offsetY; y < image.getHeight(); y += step) {
			XSSFRow row = sheet.createRow(sheetY++);
			sheetX = 0;
			
			row.setHeight((short) PIXEL_HEIGHT);
			
			for (int x = offsetX; x < image.getWidth(); x += step) {
				byte[] pixel = imageData.getPixel(x, y);
				
				pixel[0] = (byte) (((pixel[0] & 0xFF) / 16) * 16);
				pixel[1] = (byte) (((pixel[1] & 0xFF) / 16) * 16);
				pixel[2] = (byte) (((pixel[2] & 0xFF) / 16) * 16);
				
				int hash = Arrays.hashCode(pixel);
				XSSFCellStyle style = styleCache.get(hash);
				
				if (style == null) {
					style = makeColorStyle(wb, pixel);
					styleCache.put(hash, style);
				}
				
				XSSFCell cell = row.createCell(sheetX++);
				cell.setCellStyle(style);
			}
		}
		
		for (int x = 0; x < sheetX; x++) {
			sheet.setColumnWidth(x, PIXEL_WIDTH);
		}

		FileOutputStream fos = new FileOutputStream(outputFile);
		wb.write(fos);
		fos.close();
	}

	/**
	 * Create a new cell style object with a background color set to the value of the given RGB pixel.
	 * 
	 * @param wb
	 * @param pixel
	 * @return
	 */
	private static XSSFCellStyle makeColorStyle(XSSFWorkbook wb, byte[] pixel) {
		XSSFCellStyle style = wb.createCellStyle();
		
//		style.setFillForegroundColor(new XSSFColor(pixel));
//		style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
		
		// Short method of doing setFillForegroundColor and setFillPattern is way too slow
		// Use some POI internals directly to speed things up
		
		CTFill ctFill = CTFill.Factory.newInstance();
		CTPatternFill patternFill = ctFill.addNewPatternFill();
		patternFill.setPatternType(STPatternType.Enum.forInt(FillPatternType.SOLID_FOREGROUND.ordinal() + 1));
		patternFill.setFgColor(new XSSFColor(pixel).getCTColor());
		
		CTXf coreXf = style.getCoreXf();
		XSSFCellFill xssfCellFill = new XSSFCellFill(ctFill);
		
		// Don't use putFill since it uses indexOf internally and scales horribly
		// We know we're not adding the same fill twice
		wb.getStylesSource().getFills().add(xssfCellFill);
		coreXf.setFillId(wb.getStylesSource().getFills().size() - 1);
		coreXf.setApplyFill(true);
		
		return style;
	}
	
}
