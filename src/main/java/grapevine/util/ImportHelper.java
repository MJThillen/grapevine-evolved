package grapevine.util;

import grapevine.constants.Constants;
import grapevine.constants.FileFormat;

import java.io.IOException;
import java.io.ObjectInputStream;

public class ImportHelper {

    public static FileFormat detectFileFormat(ObjectInputStream inputStream) throws IOException {
        FileFormat fileFormat = FileFormat.INVALID;
        byte[] bytes = (new char[Constants.BINARY_HEADER_LENGTH]).toString().getBytes();
        byte[] longest = Constants.GAME_FILE_2_0V2.getBytes();
        inputStream.mark(longest.length + 5);
        String header = new String(inputStream.readNBytes(bytes.length));
        switch(header) {
            case Constants.BINARY_HEADER_GAME:
                return FileFormat.BINARY_GAME;
            case Constants.BINARY_HEADER_MENU:
                return FileFormat.BINARY_MENU;
            case Constants.BINARY_HEADER_EXCHANGE:
                return FileFormat.BINARY_EXCHANGE;
            default:
                inputStream.reset();
                String longHeader = new String(inputStream.readNBytes(longest.length));
                if (longHeader.substring(0, 5).equals("<?xml")) {
                    return FileFormat.XML;
                }
        }
        return fileFormat;
    }
}
