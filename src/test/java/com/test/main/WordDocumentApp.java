package com.test.main;

import java.io.File;
import java.io.IOException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import com.image.demo.WordDocument;

public class WordDocumentApp {
	public static void main(String[] args) throws InvalidFormatException, IOException {
		WordDocument wordDocument = new WordDocument();
		
		File folder = new File("src/test/resources/Screenshots");
		File[] listOfFiles = folder.listFiles();

		String[] fileNames = new String[listOfFiles.length];
		for (int i = 0; i < listOfFiles.length; i++) {
			if (listOfFiles[i].isFile()) {
				fileNames[i] = listOfFiles[i].getName();
			} else if (listOfFiles[i].isDirectory()) {
				fileNames[i] = listOfFiles[i].getName();
			}
		}
		
		sortStringExchange(fileNames);
		
		for (int i = 0; i < fileNames.length; i++) {
			File image = new File("src/test/resources/Screenshots/"+fileNames[i]);
			wordDocument.addImagesToWordDocument(image);
			System.out.println(fileNames[i]+" added");
		}
		wordDocument.closeWordDoument();
		
		System.out.println("Finished");
	}
	
	public static void sortStringExchange( String  x [ ] )
    {
          int i, j;
          String temp;

          for ( i = 0;  i < x.length - 1;  i++ )
          {
              for ( j = i + 1;  j < x.length;  j++ )
              { 
                       if ( x [ i ].compareToIgnoreCase( x [ j ] ) > 0 )
                        {                                             // ascending sort
                                    temp = x [ i ];
                                    x [ i ] = x [ j ];    // swapping
                                    x [ j ] = temp;
                                   
                         }
                 }
           }
    }
	
	
}
