package com.company;

import java.io.IOException;

public class Main {

    public static void main(String[] args) throws IOException {

        //BoardSeatsFiles.copyRequiredFiles();
        TableIndexes tableIndexes = new TableIndexes();
/*
        for(String file : BoardSeatsFiles.filesToProcess()) {
            if(!file.contains("$"))
            tableIndexes.getTableIndexes(file);
        }*/

        tableIndexes.writeTableData();
    }
}
