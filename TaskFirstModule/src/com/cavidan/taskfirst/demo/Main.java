
package com.cavidan.taskfirst.demo;


import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

import java.io.File;
import java.io.FileOutputStream;
import java.io.FileNotFoundException;

import com.sun.org.apache.xml.internal.serialize.XMLSerializer;
import com.sun.org.apache.xml.internal.serialize.OutputFormat;
import java.io.FileInputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Iterator;

import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.HashMap;
import javax.swing.JFileChooser;
import javax.swing.filechooser.FileNameExtensionFilter;
import javax.swing.filechooser.FileSystemView;
import org.apache.commons.codec.binary.StringUtils;

public class Main {

    public Main() {
        try {
            this.writeProductsXML();
        } catch (ParserConfigurationException pce) {

        } catch (FileNotFoundException fnf) {

        } catch (IOException ioe) {

        }
    }

    private void writeProductsXML() throws ParserConfigurationException, FileNotFoundException, IOException {

        
        // Upload excel file and get the file's  url
        
        Upload up1 = new Upload();
        String path = up1.getPath();
        
        
        // Start reading the excel file
        
        File excelFile = new File(path);
        FileInputStream fis = new FileInputStream(excelFile);

        XSSFWorkbook workbook = new XSSFWorkbook(fis);

        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFSheet sheet2 = workbook.getSheetAt(1);

        Iterator rowIt = sheet.rowIterator();
        Iterator rowItBranches = sheet2.rowIterator();
        
        
        
        
       // Branch
     
       Branch branch = new Branch();
       
        HashMap<String, String> filials = branch.createHashMap(rowItBranches);

      
      // End of branches
        
        
        
        XML xml = new XML();
        
        xml.createXML(rowIt, filials, workbook, path, fis);
        

    }

    public static void main(String[] args) {
        Main productsXML = new Main();
    }
}
