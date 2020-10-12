package com.company;

import org.apache.commons.io.FileUtils;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Scanner;

public class BoardSeatsFiles {
    private static final String allFilesLocation = "C:\\Users\\Purrrr\\Downloads\\pbpeople\\";
    private static final String copiedFilesLocation = "C:\\Users\\Purrrr\\Desktop\\BoardSeats\\";
    private static final String requiredFilesList = "C:\\Users\\Purrrr\\Downloads\\dirlist (2).txt";

    private static ArrayList<String> fileNames(){

        ArrayList<String> fileNames = new ArrayList<String>();
        String line;

        try
        {
            FileInputStream fis=new FileInputStream(requiredFilesList);
            Scanner sc=new Scanner(fis);

            while(sc.hasNextLine())
            {
                line = sc.nextLine();
                System.out.println(line);
                fileNames.add(line);
            }
            sc.close();
        }
        catch(IOException e)
        {
            e.printStackTrace();
        }

        return fileNames;
    }

    public static void copyRequiredFiles() throws IOException {
       for(String filename : fileNames()) {
           if(filename.contains("docx")) {
               File original = new File(allFilesLocation + filename);
               File copied = new File(
                       copiedFilesLocation + filename);
               FileUtils.copyFile(original, copied);
           }
       }

    }

    public static ArrayList<String> filesToProcess(){
        ArrayList<String> results = new ArrayList<String>();
        File[] files = new File(copiedFilesLocation).listFiles();

        for (File file : files) {
            if (file.isFile()) {
                results.add(file.getPath());
            }
        }

        return results;
    }





}
