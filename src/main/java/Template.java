import javax.xml.bind.annotations;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import javax.xml.stream.XMLStreamWriter;

@XMLRootElement
public class Template {
    private String name;
    private boolean isCharacterSheet;
    private String[] fileName;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public boolean isCharacterSheet() {
        return isCharacterSheet;
    }

    public void setCharacterSheet(boolean characterSheet) {
        isCharacterSheet = characterSheet;
    }

    public String getFileName(String extension) {
        String lowExtension = extension.toLowerCase().trim();
        switch(lowExtension) {
            case "text":
                return fileName[0];
            case "rtf":
                return fileName[1];
            case "html":
                return fileName[2];
            default:
                return "";
        }
    }

    public void setFileName(String fileName, String extension) {
        String lowExtension = extension.toLowerCase().trim();
        switch(lowExtension) {
            case "text":
                this.fileName[0] = fileName;
            case "rtf":
                this.fileName[1] = fileName;
            case "html":
                this.fileName[2] = fileName;
        }
    }

    //TODO: Check XML writing best practices.

    /**
     * Write the object to an XML file.
     * @param xml The XML Stream Writer (that just wrote the opening tag)
     */
    public void outputToFile(XMLStreamWriter xml) throws XMLStreamException {
    }

    /**
     * Read the object from an XML file
     * @param xml The XML Stream Reader (That just read the opening tag)
     * @param version Version of the file Format
     */
    public void inputFromFile(XMLStreamReader xml, double version) {
        if (xml.hasNext() && xml.n)
    }
}
