
package com.cavidan.taskfirst.demo;

import java.util.HashMap;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;

public class Branch {
    
    public HashMap createHashMap(Iterator rowItBranches){
         String key, trimmedKey, value ;
        char character;
        HashMap<String, String> filials = new HashMap<String, String>();

       while (rowItBranches.hasNext()) {
           
            XSSFRow rowBranches = (XSSFRow) rowItBranches.next();

            //System.out.println("ROW:-->");
            Iterator<Cell> cellIteratorBranch = rowBranches.cellIterator();
            Iterator<Cell> cellIteratorBranch2 = rowBranches.cellIterator();

            XSSFCell cell22 = (XSSFCell) cellIteratorBranch2.next();
            if (cell22.toString() == "") {
                break;
            }
            
              int flag = 0;
              key = "";
              trimmedKey = "";
              value = "";
                while (cellIteratorBranch.hasNext()) {
                    XSSFCell cell = (XSSFCell) cellIteratorBranch.next();
                  
                    if(flag == 1){
                        key = cell.toString();
                    for (int i = 0; i < key.length(); i++) {
                        
                        
                        character = key.charAt(i);    
                        int ascii = (int) character;
                        if(ascii != 160){
                            trimmedKey += character;
                        }
                        else{
                            break;
                        }
                        
                       
                    }
                    
                    }
                    if(flag == 2){
                        value = cell.toString().trim();                                              
                    }
                    flag++;
                }
                
                    filials.put(trimmedKey, value);
                    
                    
            
       }
       
       return filials;
    }
    
}
