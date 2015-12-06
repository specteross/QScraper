/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package qscraper;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.HttpStatusException;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

/**
 *
 * @author mohmishr
 * last edited on 7/12/2015 by specteross
 */
public class QScraper {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws IOException{
        // TODO code application logic here
    	//Blank workbook
        XSSFWorkbook workbook = new XSSFWorkbook();
         
        //Create a blank sheet
        XSSFSheet sheet = workbook.createSheet("Employee Data");
          
        //This data needs to be written (Object[])
        Map<String, Object[]> data = new TreeMap<String, Object[]>();
        data.put("1", new Object[] {"Question ID", "Question", "A", "B", "C", "D", "E", "Correct Answer"});
        
        int n_urls = 4;
        // URLs from which questions need to be extracted
        String URLList[] = new String[n_urls];
        URLList[0] = "http://gmatclub.com/forum/sentence-correction-question-directory-topic-and-difficulty-129750.html";
        URLList[1] = "http://gmatclub.com/forum/critical-reasoning-question-directory-topic-and-difficulty-128861.html";
        URLList[2] = "http://gmatclub.com/forum/gmat-ps-question-directory-by-topic-difficulty-127957.html";
        URLList[3] = "http://gmatclub.com/forum/ds-question-directory-by-topic-and-difficulty-128728.html";
        
        // Connect jsoup with the link
        int i = 1;
        for (int n_url_i = 0; n_url_i < n_urls; n_url_i++) {
        	String url = URLList[n_url_i];
        	Document document;
            try{
            	document = Jsoup.connect(url).get();
            } catch (HttpStatusException e){
            	continue;
            }
            Elements links = document.select("a.postlink-local[href]");
            for (Element mainLink : links){
            	url = mainLink.attr("href");
            	Document doc;
                try{
                	doc = Jsoup.connect(url).get();
                } catch (HttpStatusException e){
                	continue;
                }
	            for (Element td : doc.select("td.topicsName")) {
	                for (Element link : td.select("a[title]")){
	                    url = link.attr("href");
	                    Document subdoc;
	                    try{
	                    	subdoc = Jsoup.connect(url).get();
	                    } catch (HttpStatusException e){
	                    	continue;
	                    }
	                    Element tr = subdoc.select("tr[id*=p_]").first();                    
	                    if (tr != null) {
	                        String text = (tr.select("div[class=item text").text());
	                        int index = text.indexOf("[Reveal]");
	                        
	                        if(index != -1){
		                        int[] indices = new int[5];
		                        
		                        // this if-else havoc since the GMAT site people can't seem to decide 
		                        //whether they want to use A. or (A) etc for denoting the start of an option.
		                        boolean flag=true;
		                        if(text.indexOf("(A) ") != -1){
		                        	indices[0] = text.indexOf("(A) ");
		                        	indices[1] = text.indexOf("(B) ");
			                        indices[2] = text.indexOf("(C) ");
			                        indices[3] = text.indexOf("(D) ");
			                        indices[4] = text.indexOf("(E) ");
		                        } else if(text.indexOf("A. ") != -1){
		                        	indices[0] = text.indexOf("A. ");
		                            indices[1] = text.indexOf("B. ");
	                                indices[2] = text.indexOf("C. ");
	                                indices[3] = text.indexOf("D. ");
	                                indices[4] = text.indexOf("E. ");
		                        } else if(text.indexOf("A) ") != -1){
		                        	indices[0] = text.indexOf("A) ");
	                                indices[1] = text.indexOf("B) ");
	                                indices[2] = text.indexOf("C) ");
	                                indices[3] = text.indexOf("D) ");
	                                indices[4] = text.indexOf("E) ");
		                        } else if(text.indexOf("a) ") != -1){
		                        	indices[0] = text.indexOf("a) ");
	                                indices[1] = text.indexOf("b) ");
	                                indices[2] = text.indexOf("c) ");
	                                indices[3] = text.indexOf("d) ");
	                                indices[4] = text.indexOf("e) ");
		                        } else {
		                        	flag = false;
		                        }
		                        
		                        String question;
		                        String A[] = new String[5];
		                        if(flag){
		                        	question = getSubstring(text, 0, indices[0]);
			                        for(int it = 1; it < 5; it++){
			                        	A[it-1] = getSubstring(text, indices[it-1]+3, indices[it]);
			                        }
			                        A[4] = getSubstring(text, indices[4]+3, index);
		                        }
		                        else{
		                        	question = getSubstring(text, 0, index);
			                        for(int it = 0; it < 5; it++){
			                        	A[it] = "";
			                        }
		                        }                        		
		                        		
		                        //String A = text.substring(indices[0]+3, indices[1]);                          
		                        int correctAnswerIndex = text.indexOf("[Reveal] Spoiler: OA");
		                        String correctAnswer = getSubstring(text, correctAnswerIndex+21, correctAnswerIndex+22);
		                        System.out.println("\nQ. "+i+": ");
		                        System.out.println("\n\ntext = " + text + "\n\nquestion = " + question
		                        			+ "\nA = " + A[0]
		                                		+ "\nB = " + A[1]
		                           			+ "\nC = " + A[2]
		                                        	+ "\nD = " + A[3]
		                                        	+ "\nE = " + A[4]
		                                        	+ "\n answer = " + correctAnswer);
		                        
		                        data.put(Integer.toString(i+1), new Object[] {i, question, A[0], A[1], A[2], A[3], A[4], correctAnswer});
		                        
		                        i++;
	                        }
	                    }                   
	                }
	            }
            }
        }
        
        //Iterate over data and write to sheet
        Set<String> keyset = data.keySet();
        int rownum = 0;
        for (String key : keyset)
        {
            Row row = sheet.createRow(rownum++);
            Object [] objArr = data.get(key);
            int cellnum = 0;
            for (Object obj : objArr)
            {
               Cell cell = row.createCell(cellnum++);
               if(obj instanceof String)
                    cell.setCellValue((String)obj);
                else if(obj instanceof Integer)
                    cell.setCellValue((Integer)obj);
            }
        }
        try
        {
            //Write the workbook in file system
            FileOutputStream out = new FileOutputStream(new File("data.xlsx"));
            workbook.write(out);
            out.close();
            System.out.println("data.xlsx written successfully on disk.");
        }
        catch (Exception e)
        {
            e.printStackTrace();
        }
    }

	private static String getSubstring(String text, int startIndex, int endIndex){
    	if (startIndex > endIndex) return "";
    	if (startIndex == -1) return "";
    	else return text.substring(startIndex, endIndex);
    }
}
