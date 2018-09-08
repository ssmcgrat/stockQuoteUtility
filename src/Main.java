import java.io.BufferedReader;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.ArrayList;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import yahoofinance.Stock;
import yahoofinance.YahooFinance;

public class Main {

	private static final String USER_AGENT = "Mozilla/5.0";

	public static void main(String[] args) throws Exception {
		System.out.println("<--- LET'S BUY SOME FUCKIN' STOCKS MAAAAAAAAN!!!! --->\n\n");
		ArrayList<String> stocks = readStocksFromExcel();
		String stockListStr = "{ ";
		for (int i=0; i<stocks.size(); i++) {
			if (i > 0) 
				stockListStr += ", ";
			stockListStr += stocks.get(i);
		}
		stockListStr += " }";
		System.out.println("Getting prices for following stocks: " + stockListStr);

		ArrayList<String> prices = new ArrayList<String>();

		for (String stock : stocks) {
			prices.add(getStockPrice(stock));
		}		
		
		writePricesToExcel(prices);
		System.out.println("All done. I hope it's a fortune, Clark.");
	}

	private static String getStockPrice(String stock) throws Exception {
		/*try{
			return sendGet(stock).trim();			
		} catch (Exception e) {
			try {
				System.out.println("need to get from nasdaq");
				String result = getFromNasdaq(stock);
				System.out.println(result);
				if (result.equals("nbsp;")) {
					result = getFromYahoo(stock);
				}
				return result;
			} catch (Exception e1) {
				System.out.println("need to get from OTC");
				return getFromYahoo(stock);
			}
		}*/
		//return getFromYahoo(stock.toUpperCase());
		String strPrice = getFromYahoo(stock);
		Double price;
		try {
			price = Double.parseDouble(strPrice);
		} catch (Exception e) {
			strPrice = getFromNasdaq(stock);
			price = Double.parseDouble(strPrice); 
		}
		
		System.out.println(stock + ": " + price);
		return strPrice;

	}

	// HTTP GET request
	private static String sendGet(String stock) throws Exception {

		String url = "https://api.iextrading.com/1.0/stock/" + stock + "/price";

		URL obj = new URL(url);
		HttpURLConnection con = (HttpURLConnection) obj.openConnection();

		// optional default is GET
		con.setRequestMethod("GET");

		//add request header
		con.setRequestProperty("User-Agent", USER_AGENT);

		BufferedReader in = new BufferedReader(
				new InputStreamReader(con.getInputStream()));
		String inputLine;
		StringBuffer response = new StringBuffer();

		while ((inputLine = in.readLine()) != null) {
			response.append(inputLine);
		}
		in.close();

		//print result
		System.out.printf("%4s: " + response.toString() + "\n", stock);
		return response.toString();

	}

	private static String getFromNasdaq(String stock) throws IOException {
		URL url;
	    InputStream is = null;
	    BufferedReader br;
	    String line;
	    StringBuffer sb = new StringBuffer();

	    try {
	        url = new URL("https://www.nasdaq.com/symbol/" + stock);
	        is = url.openStream();  // throws an IOException
	        br = new BufferedReader(new InputStreamReader(is));

	        while ((line = br.readLine()) != null) {
	            sb.append(line);
	        }
	        
	        int index = sb.indexOf("<b>Last Net Asset Value (NAV)</b>");
	        int startIndex = sb.indexOf("<td>", index) + 5;
	        String price = sb.substring(startIndex, startIndex + 5);
	        
	        return price;
	    } catch (Exception e) {
	    	System.err.println(e.getMessage());
	    	return "";
	    } finally {
	        try {
	            if (is != null) is.close();
	        } catch (IOException ioe) {
	            // nothing to see here
	        }
	    }
	}
	
	private static String getFromYahoo(String stock) throws IOException {
		URL url;
	    InputStream is = null;
	    BufferedReader br;
	    String line;
	    StringBuffer sb = new StringBuffer();

	    try {
	    	String urlStr = "https://finance.yahoo.com/quote/<stock>?p=<stock>";
	    	urlStr = urlStr.replaceAll("<stock>", stock);
	        url = new URL(urlStr);
	        is = url.openStream();  // throws an IOException
	        //System.out.println("opened Stream: " + urlStr);
	        br = new BufferedReader(new InputStreamReader(is));

	        while ((line = br.readLine()) != null) {
	            sb.append(line);
	            //System.out.println(line);
	        }
	        
	        /*String key = "Previous Close</span>";
	        int index = sb.indexOf(key);
	        System.out.println("index: " + index);
	        int endIndex = sb.indexOf("</span></td></tr>", index);
	        System.out.println("endIndex: " + endIndex);
	        int startIndex = 0;
	        for (int i = endIndex - 1; i>0; i--) {
	        	if (sb.charAt(i) == '>') {
	        		startIndex = i + 1;
	        		break;
	        	}
	        }*/
	        
	        String key = "currentPrice\":{\"raw\":";
	        int index = sb.indexOf(key);
	        int endIndex = sb.indexOf(",", index);
	        int startIndex = index + key.length();

	        String price = sb.substring(startIndex, endIndex);
	        
	        return price;
	    } catch (Exception e) {
	    	System.err.println(e.getMessage());
	    	return "";
	    } finally {
	        try {
	            if (is != null) is.close();
	        } catch (IOException ioe) {
	            // nothing to see here
	        }
	    }
	}

	private static ArrayList<String> readStocksFromExcel() throws Exception {
		System.out.println("Getting list of stock symbols...");
		ArrayList<String> stockList = new ArrayList<String>();

		File file = new File(System.getProperty("user.dir") + File.separator + "stocks.xlsx");

		//Create an object of FileInputStream class to read excel file
		FileInputStream inStream = new FileInputStream(file);
		Workbook workbook = new XSSFWorkbook(inStream);




		//Read sheet inside the workbook by its name

		Sheet sheet = workbook.getSheet("Sheet1");

		//Find number of rows in excel file

		int rowCount = sheet.getLastRowNum()- sheet.getFirstRowNum();

		//Create a loop over all the rows of excel file to read it
		DataFormatter formatter = new DataFormatter();

		for (int i = 0; i <= rowCount; i++) {
			Row row = sheet.getRow(i);

			Cell stockCell = row.getCell(0);
			stockList.add(formatter.formatCellValue(stockCell));	        
		}	
		workbook.close(); // gracefully closes the underlying zip file

		return stockList;
	}

	private static void writePricesToExcel(ArrayList<String> prices) throws Exception {
		System.out.println("Writing prices to excel...");
		File file = new File(System.getProperty("user.dir") + File.separator + "stocks.xlsx");

		//Create an object of FileInputStream class to read excel file

		FileInputStream inStream = new FileInputStream(file);
		//

		Workbook workbook = new XSSFWorkbook(inStream);


		//Read sheet inside the workbook by its name

		Sheet sheet = workbook.getSheet("Sheet1");

		//Find number of rows in excel file

		int rowCount = sheet.getLastRowNum()- sheet.getFirstRowNum();

		//Create a loop over all the rows of excel file to read it

		for (int i = 0; i <= rowCount; i++) {
			Row row = sheet.getRow(i);

			Cell stockCell = row.createCell(1);
			stockCell.setCellValue(prices.get(i));	        
		}	
		inStream.close();
		FileOutputStream outStream = new FileOutputStream(file);
		workbook.write(outStream);
		workbook.close();
	}

}