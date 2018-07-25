package de.haw.hamburg.haeckel.workinghoursedit;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import java.io.FileNotFoundException;

public class Main {



    public static void main(String[] args) {
	    System.out.println("Starting Processor...");

	    String inFile = "src/main/resources/Juni.xlsx";
	    String outFile = "output/Juni_edit.xlsx";

	    if(args.length > 0){
            inFile = findArgs(args, "input");
            outFile = findArgs(args, "output");
        }

        try {
            Processor instance = new Processor(inFile, outFile);
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
    }

    private static String findArgs(String[] args, String match) {
        String ret = null;
        for (String arg : args){
            String [] split = arg.split("=");
            if(split.length == 2) {
                String key = split[0];
                String value = split[1];
                if(key.toUpperCase().equals(match.toUpperCase())){
                    ret = value;
                }
            }

        }
        return ret;
    }
}
