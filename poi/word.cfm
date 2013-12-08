
<cfscript>

path = "C:\Railo\railo-express-4.1.1.009-nojre\webapps\www\word.docx"

fis = CreateObject("java", "java.io.FileInputStream").init(path)
dump(fis)

doc = CreateObject("java", "org.apache.poi.xwpf.usermodel.XWPFDocument").init(fis)
dump(doc)

pics = doc.getAllPictures()
dump(pics)

dump(pics[1].getParent())

fis.close()

/*
	POIFSFileSystem fis = new POIFSFileSystem(new FileInputStream(wordFile);
	HWPFDocument wdDoc = new HWPFDocument(fis);
	int pagesNo = wdDoc.getSummaryInformation().getPageCount();
	pagesCount += pagesNo;
	System.out.println(files[i].getName()+":\t"+pagesNo);
*/

</cfscript>

<!---
package org.word.POI;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

/*
Romesh Soni
soni.romesh@gmail.com
*/

public class TestCustom
{

    public static void main(String []a) throws FileNotFoundException, IOException, InvalidFormatException
    {

        CustomXWPFDocument document = new CustomXWPFDocument(new FileInputStream(new File("C:\\Users\\amitabh\\Documents\\Apache POI\\Word File\\new.doc")));
        FileOutputStream fos = new FileOutputStream(new File("C:\\Users\\amitabh\\Documents\\Apache POI\\Word File\\new.doc"));

        String blipId = document.addPictureData(new FileInputStream(new File("C:\\Users\\amitabh\\Pictures\\pics\\3.jpg")), Document.PICTURE_TYPE_JPEG);

        System.out.println(document.getNextPicNameNumber(Document.PICTURE_TYPE_JPEG));

        //System.out.println(document.getNextPicNameNumber(Document.PICTURE_TYPE_JPEG));
        document.createPicture(blipId,document.getNextPicNameNumber(Document.PICTURE_TYPE_JPEG), 500, 500);


        document.write(fos);
        fos.flush();
        fos.close();

    }

}
--->