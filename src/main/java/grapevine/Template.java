package grapevine;

import javax.xml.bind.annotation.XmlElement;
import javax.xml.bind.annotation.XmlRootElement;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.Serializable;

@XmlRootElement
public class Template implements Serializable {
    @XmlElement
    private String name;
    @XmlElement
    private boolean isCharacterSheet;

    public Template() {
        //This constructor intentionally left blank.
    }

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

    //Output to File and From file will be handled by jaxb in a calling method. No need for methods here.

    /**
     * ToDo: Test this for backwards compatibility. A lot.
     * Writes the object to binary depending on software version
     * @param outputStream The in-progress binary output stream
     * @param version The version of the file to be written
     * @throws IOException on error
     */
    public void outputToBinary(ObjectOutputStream outputStream, double version) throws IOException {
        if (version >= 4.0) {
            outputStream.writeObject(this);
        } else {
            byte[] nameBytes = name.getBytes();
            byte[] text = ("Templates/Text/" + name + ".txt").getBytes();
            byte[] rtf = ("Templates/RTF/" + name + ".rtf").getBytes();
            byte[] html = ("Templates/HTML/" + name + ".html").getBytes();

            outputStream.writeInt(nameBytes.length);
            outputStream.write(nameBytes);
            outputStream.writeInt(isCharacterSheet ? -1 : 0); //VB6 writes -1 for true, 0 for false.
            outputStream.writeInt(text.length);
            outputStream.write(text);
            outputStream.writeInt(rtf.length);
            outputStream.write(rtf);
            outputStream.writeInt(html.length);
            outputStream.write(html);
        }
    }

    /**
     * Reads the object from binary depending on software version
     * @param inputStream The in-progress binary input stream
     * @param version the version of the file to be read
     * @return The reconstructed object
     * @throws IOException on error
     * @throws ClassNotFoundException on error
     */
    public static Template inputFromBinary(ObjectInputStream inputStream, double version) throws IOException, ClassNotFoundException {
        Template template = new Template();
        int length = 0;
        if (version >= 4.0) {
            template = (Template) inputStream.readObject();
        } else {
            length = inputStream.readInt();
            template.setName(new String(inputStream.readNBytes(length)));
            template.setCharacterSheet(inputStream.readInt() != 0);
            //Purging the old information as the new format doesn't need the duplicates by file type.
            length = inputStream.readInt();
            inputStream.readNBytes(length);
            length = inputStream.readInt();
            inputStream.readNBytes(length);
            length = inputStream.readInt();
            inputStream.readNBytes(length);
    }
        return template;
    }
}

