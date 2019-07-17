/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.cavidan.taskfirst.demo;

import com.sun.org.apache.xml.internal.serialize.OutputFormat;
import com.sun.org.apache.xml.internal.serialize.XMLSerializer;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.time.LocalDate;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.HashMap;
import java.util.Iterator;
import javax.xml.parsers.DocumentBuilder;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.parsers.ParserConfigurationException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Text;

public class XML {
    
    public void createXML(Iterator rowIt, HashMap filials,XSSFWorkbook workbook,String path,FileInputStream fis) throws ParserConfigurationException, IOException{
        
        
        
        
        //DocumentBuilderFactory
        DocumentBuilderFactory docFactory = DocumentBuilderFactory.newInstance();
        //DocumentBuilder
        DocumentBuilder docBuilder = docFactory.newDocumentBuilder();
        //Document
        Document xmlDoc = docBuilder.newDocument();
        //Build XML Elements

        Element rootElement = xmlDoc.createElement("ApplicationFile");

        //  Beginning of FileHeader
        Element mainElement1 = xmlDoc.createElement("FileHeader");
        Element mainElement2 = xmlDoc.createElement("ApplicationsList");

        Element FormatVersion = xmlDoc.createElement("FormatVersion");
        Element Sender = xmlDoc.createElement("Sender");
        Element CreationDate = xmlDoc.createElement("CreationDate");
        Element CreationTime = xmlDoc.createElement("CreationTime");
        Element Number = xmlDoc.createElement("Number");
        Element Institution = xmlDoc.createElement("Institution");

        mainElement1.appendChild(FormatVersion);
        mainElement1.appendChild(Sender);
        mainElement1.appendChild(CreationDate);
        mainElement1.appendChild(CreationTime);
        mainElement1.appendChild(Number);
        mainElement1.appendChild(Institution);

        rootElement.appendChild(mainElement1);
        rootElement.appendChild(mainElement2);

        Text formatVer = xmlDoc.createTextNode(("2.0"));
        FormatVersion.appendChild(formatVer);
        Text senderCode = xmlDoc.createTextNode(("0400"));
        Sender.appendChild(senderCode);

        Text num = xmlDoc.createTextNode(("84"));
        Number.appendChild(num);
        Text inst = xmlDoc.createTextNode(("0400"));
        Institution.appendChild(inst);
        DateTimeFormatter dtf = DateTimeFormatter.ofPattern("yyyy/MM/dd");
        LocalDate localDate = LocalDate.now();

        Text currentDate = xmlDoc.createTextNode((dtf.format(localDate)));
        CreationDate.appendChild(currentDate);

        Text currentTime = xmlDoc.createTextNode(new SimpleDateFormat("HH:mm:ss").format(Calendar.getInstance().getTime()));
        CreationTime.appendChild(currentTime);

        int iter = 0, count = 0, exit = 0;

        while (rowIt.hasNext()) {
            if (exit == 1) {
                break;
            }
            count++;
            XSSFRow row = (XSSFRow) rowIt.next();

            //System.out.println("ROW:-->");
            Iterator<Cell> cellIterator = row.cellIterator();
            Iterator<Cell> cellIterator2 = row.cellIterator();

            XSSFCell cell2 = (XSSFCell) cellIterator2.next();
            if (cell2.toString() == "") {
                break;
            }
            if (iter != 0) {

                // BUrdaaaaaaaaaaaaaaaaaaan
                Element Application = xmlDoc.createElement("Application");

                Element RegNumber = xmlDoc.createElement("RegNumber");
                Element OrderDprt = xmlDoc.createElement("OrderDprt");
                Element ObjectType = xmlDoc.createElement("ObjectType");
                Element ActionType = xmlDoc.createElement("ActionType");
                Element ObjectFor = xmlDoc.createElement("ObjectFor");
                Element Data = xmlDoc.createElement("Data");

                mainElement2.appendChild(Application);
                Application.appendChild(RegNumber);
                Application.appendChild(OrderDprt);
                Application.appendChild(ObjectType);
                Application.appendChild(ActionType);
                Application.appendChild(ObjectFor);
                Application.appendChild(Data);

                Element ContractIDT = xmlDoc.createElement("ContractIDT");
                ObjectFor.appendChild(ContractIDT);
                Element ContractNumber = xmlDoc.createElement("ContractNumber");
                Element Client = xmlDoc.createElement("Client");
                ContractIDT.appendChild(ContractNumber);
                ContractIDT.appendChild(Client);
                Element ClientInfo = xmlDoc.createElement("ClientInfo");
                Client.appendChild(ClientInfo);
                Element ShortName = xmlDoc.createElement("ShortName");
                ClientInfo.appendChild(ShortName);
                Element SetStatus = xmlDoc.createElement("SetStatus");
                Data.appendChild(SetStatus);
                Element StatusCode = xmlDoc.createElement("StatusCode");
                Element StatusComment = xmlDoc.createElement("StatusComment");
                SetStatus.appendChild(StatusCode);
                SetStatus.appendChild(StatusComment);

                //BUraaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa
                int flag = 0;
                while (cellIterator.hasNext()) {
                    XSSFCell cell = (XSSFCell) cellIterator.next();
                    if (cell.toString() == "") {
                        exit = 1;
                        break;
                    }
                    //  System.out.println("CELL:-->"+cell.toString());

                    if (flag == 0) {

                        Text ContractNumberText = xmlDoc.createTextNode((cell.toString()));
                        ContractNumber.appendChild(ContractNumberText);
                    }
                    if (flag == 1) {
                        Text ShortNameText = xmlDoc.createTextNode((cell.toString()));
                        ShortName.appendChild(ShortNameText);
                    }
                    if (flag == 2) {
                       Text OrderDprtText = xmlDoc.createTextNode((String) filials.get(cell.toString()));
                        OrderDprt.appendChild(OrderDprtText);
                    }
                    if (flag == 3) {
                        Text StatusCommentText = xmlDoc.createTextNode((cell.toString()));
                        StatusComment.appendChild(StatusCommentText);
                    }
                    if (flag == 4) {
                        Text RegNumberText = xmlDoc.createTextNode((cell.toString()));
                        RegNumber.appendChild(RegNumberText);
                    }

                    flag++;

                }
                // Statics
                //ObjectTYpe
                Text ObjectTypeText = xmlDoc.createTextNode(("Status"));
                ObjectType.appendChild(ObjectTypeText);

                //ActionType
                Text ActionTypeText = xmlDoc.createTextNode(("Update"));
                ActionType.appendChild(ActionTypeText);

                //StatusCode
                Text StatusCodeText = xmlDoc.createTextNode(("14"));
                StatusCode.appendChild(StatusCodeText);

            }

            iter++;

        }

        workbook.close();

        fis.close();

        System.out.println("Count  is " + count);

        xmlDoc.appendChild(rootElement);
        //Set OutputFormat
        OutputFormat outFormat = new OutputFormat(xmlDoc);
        outFormat.setIndenting(true);
        //Declare the file
        String path2 = path.replace(".xlsx", ".xml");
        File xmlFile = new File(path2);
        //Declare the FileOutputStream
        FileOutputStream outStream = new FileOutputStream(xmlFile);
        //XMLSerializer to serialize the XML data with
        XMLSerializer serializer = new XMLSerializer(outStream, outFormat);
        // the specified output format
        serializer.serialize(xmlDoc);
    }
}
